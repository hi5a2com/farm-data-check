"""牧场繁育记录自动核对。

读入一份繁育记录 Excel（工作表「繁育记录」和「牛群名册」），找出录入错误和需要人工复核的记录，
输出一份核对结果 Excel：问题清单 + 原表里把问题行标黄。

用法：
    uv run check.py 繁育记录.xlsx
    uv run check.py 繁育记录.xlsx -o 核对结果.xlsx --today 2026-10-06
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

# 阈值按牧场自己的标准改
VOLUNTARY_WAITING_DAYS = 45  # 产后至少多少天才配种
MIN_PREG_CHECK_DAYS = 28  # 配种后至少多少天才能做孕检
HEAT_INTERVAL_RANGE = (18, 24)  # 正常发情间隔（天）

EVENTS = {"发情", "配种", "孕检", "产犊", "用药"}
RECORD_COLUMNS = ["序号", "日期", "耳号", "事件", "详情", "操作人"]

RULES = {
    "E01": ("错误", "耳号不在牛群名册里"),
    "E02": ("错误", "日期为空、格式不对或晚于今天"),
    "E03": ("错误", "和前面的记录完全重复"),
    "E04": ("错误", "操作人没填"),
    "E05": ("错误", "事件名称不在规定列表里"),
    "W01": ("复核", f"产后不到 {VOLUNTARY_WAITING_DAYS} 天就配种"),
    "W02": ("复核", f"配种后不到 {MIN_PREG_CHECK_DAYS} 天就做孕检"),
    "W03": (
        "复核",
        f"两次发情间隔不在 {HEAT_INTERVAL_RANGE[0]}–{HEAT_INTERVAL_RANGE[1]} 天",
    ),
}


@dataclass
class Record:
    row: int  # Excel 里的行号，用来标黄
    no: str
    raw_date: object
    tag: str
    event: str
    detail: str
    operator: str
    day: date | None = None


@dataclass(frozen=True)
class Issue:
    no: str
    rule: str
    note: str = ""


def parse_date(value: object) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        text = value.strip()
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
    return None


def cell_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


def load_records(path: Path) -> tuple[list[Record], set[str]]:
    wb = load_workbook(path, data_only=True)
    herd = {
        cell_text(row[0])
        for row in wb["牛群名册"].iter_rows(min_row=2, values_only=True)
        if row and row[0] is not None
    }
    sheet = wb["繁育记录"]
    header = [cell_text(c.value) for c in sheet[1]]
    idx = {name: header.index(name) for name in RECORD_COLUMNS}
    records = []
    for row_no, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        if not any(v is not None for v in row):
            continue
        records.append(
            Record(
                row=row_no,
                no=cell_text(row[idx["序号"]]),
                raw_date=row[idx["日期"]],
                tag=cell_text(row[idx["耳号"]]),
                event=cell_text(row[idx["事件"]]),
                detail=cell_text(row[idx["详情"]]),
                operator=cell_text(row[idx["操作人"]]),
            )
        )
    return records, herd


def check(records: list[Record], herd: set[str], today: date) -> list[Issue]:
    issues: list[Issue] = []
    seen: dict[tuple, str] = {}
    usable: list[Record] = []  # 通过基础检查、参与逻辑检查的记录

    for r in records:
        ok = True
        if r.tag not in herd:
            issues.append(Issue(r.no, "E01", f"耳号「{r.tag}」"))
            ok = False
        r.day = parse_date(r.raw_date)
        if r.day is None or r.day > today:
            issues.append(Issue(r.no, "E02", f"日期「{cell_text(r.raw_date)}」"))
            ok = False
        if not r.operator:
            issues.append(Issue(r.no, "E04"))
        if r.event not in EVENTS:
            issues.append(Issue(r.no, "E05", f"事件「{r.event}」"))
            ok = False
        key = (cell_text(r.raw_date), r.tag, r.event, r.detail, r.operator)
        if key in seen:
            issues.append(Issue(r.no, "E03", f"与序号 {seen[key]} 重复"))
            continue  # 重复行不参与逻辑检查，避免同一天两次发情算出 0 天间隔
        seen[key] = r.no
        if ok:
            usable.append(r)

    by_cow: dict[str, list[Record]] = {}
    for r in usable:
        by_cow.setdefault(r.tag, []).append(r)

    for cow_records in by_cow.values():
        cow_records.sort(key=lambda r: r.day)
        last_calving = last_breeding = last_heat = None
        for r in cow_records:
            if r.event == "产犊":
                last_calving, last_heat = r.day, None
            elif r.event == "配种":
                if last_calving and (r.day - last_calving).days < VOLUNTARY_WAITING_DAYS:
                    days = (r.day - last_calving).days
                    issues.append(Issue(r.no, "W01", f"产后第 {days} 天"))
                last_breeding = r.day
            elif r.event == "孕检":
                if last_breeding and (r.day - last_breeding).days < MIN_PREG_CHECK_DAYS:
                    days = (r.day - last_breeding).days
                    issues.append(Issue(r.no, "W02", f"配种后第 {days} 天"))
                if "阳性" in r.detail:
                    last_heat = None  # 怀上了，发情间隔重新算
            elif r.event == "发情":
                if last_heat:
                    gap = (r.day - last_heat).days
                    low, high = HEAT_INTERVAL_RANGE
                    if not low <= gap <= high:
                        issues.append(Issue(r.no, "W03", f"距上次发情 {gap} 天"))
                last_heat = r.day
    return issues


def write_report(src: Path, out: Path, records: list[Record], issues: list[Issue]) -> None:
    by_no = {r.no: r for r in records}
    wb = Workbook()
    summary = wb.active
    summary.title = "汇总"
    summary.append(["规则", "级别", "说明", "数量"])
    for rule, (level, desc) in RULES.items():
        summary.append([rule, level, desc, sum(1 for i in issues if i.rule == rule)])
    summary.append([])
    summary.append(["记录总数", len(records)])
    summary.append(["有问题的记录", len({i.no for i in issues})])

    detail = wb.create_sheet("问题清单")
    detail.append(["序号", "Excel行", "耳号", "日期", "事件", "级别", "规则", "说明", "具体情况"])
    for i in sorted(issues, key=lambda i: (by_no[i.no].row, i.rule)):
        r = by_no[i.no]
        level, desc = RULES[i.rule]
        detail.append([i.no, r.row, r.tag, cell_text(r.raw_date), r.event, level, i.rule, desc, i.note])
    for ws in (summary, detail):
        for cell in ws[1]:
            cell.font = Font(bold=True)

    # 在原表副本里标出问题行：错误标红，只需复核的标黄
    src_wb = load_workbook(src)
    marked = wb.create_sheet("原表标注")
    red = PatternFill("solid", fgColor="F8CBAD")
    yellow = PatternFill("solid", fgColor="FFF2CC")
    levels: dict[int, str] = {}
    for i in issues:
        row = by_no[i.no].row
        if RULES[i.rule][0] == "错误" or row not in levels:
            levels[row] = RULES[i.rule][0]
    for row_no, row in enumerate(src_wb["繁育记录"].iter_rows(values_only=True), start=1):
        marked.append(list(row))
        if row_no in levels:
            fill = red if levels[row_no] == "错误" else yellow
            for cell in marked[row_no]:
                cell.fill = fill
    wb.save(out)


def main() -> None:
    parser = argparse.ArgumentParser(description="核对牧场繁育记录")
    parser.add_argument("input", type=Path, help="繁育记录 Excel")
    parser.add_argument("-o", "--output", type=Path, help="核对结果 Excel（默认在输入文件旁边）")
    parser.add_argument("--today", type=date.fromisoformat, default=date.today(), help="晚于这天的日期算错误")
    args = parser.parse_args()

    out = args.output or args.input.with_name(f"{args.input.stem}-核对结果.xlsx")
    records, herd = load_records(args.input)
    issues = check(records, herd, args.today)
    write_report(args.input, out, records, issues)

    errors = sum(1 for i in issues if RULES[i.rule][0] == "错误")
    print(f"共 {len(records)} 条记录：{errors} 处错误，{len(issues) - errors} 处需要复核。")
    print(f"结果已写入：{out}")


if __name__ == "__main__":
    main()
