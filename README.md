# Automatic check of farm breeding records (farm-data-check)

[English](#english) | [中文](#中文)

---

## English

A demo project. It checks the breeding records that a cattle farm enters by hand each day (heat, breeding, pregnancy checks, calving, and medication) in one step. It finds records entered wrong and records that need a manual review.

**All data is simulated.** `make_demo_data.py` generates it at random. None of it comes from a real farm.

### What it finds

| Rule | Level | What it checks |
|---|---|---|
| E01 | 错误 (error) | The ear tag is not in the herd register (a missing 0, an extra digit, the letter O used as 0, a space in the middle) |
| E02 | 错误 (error) | The date is empty, has a bad format (for example month 13, or 「9月5日」), or is later than today |
| E03 | 错误 (error) | The same record was entered twice |
| E04 | 错误 (error) | The operator is empty |
| E05 | 错误 (error) | The event name is misspelled (for example 「用要」 or 「用 药」) |
| W01 | 复核 (review) | Breeding less than 45 days after calving |
| W02 | 复核 (review) | A pregnancy check less than 28 days after breeding |
| W03 | 复核 (review) | An interval between two heats outside 18–24 days |

The thresholds are at the top of `check.py`. Change them to match your farm's own standards.

### Usage

```bash
uv run make_demo_data.py                                   # 生成 demo/繁育记录-演示.xlsx
uv run check.py demo/繁育记录-演示.xlsx --today 2026-10-06  # 输出 …-核对结果.xlsx
uv run pytest -q                                           # 跑测试
```

The check result has three sheets. **汇总** shows how many issues each rule found. **问题清单** lists every issue. **原表标注** marks errors in red and records that need review in yellow.

The image below shows the **问题清单** sheet that the commands above produce (default seed 7): 403 records, 20 errors, and 8 records that need review.

![The 问题清单 sheet: one row per issue, with the record number, Excel row, ear tag, date, event, level, rule, and details](docs/check-result.png)

### Input format

The input is one Excel file with two sheets:

- `繁育记录`: 序号, 日期, 耳号, 事件, 详情, 操作人
- `牛群名册`: 耳号, 出生日期, 胎次

### How the tests prove nothing is missed or wrong

When the generator creates data, it records every problem it plants on purpose (the "answers"). The tests require the check result to match the answers **exactly**. No planted problem may be missed, and no normal record may be flagged. The tests must pass with 5 different random seeds.

---

## 中文

牧场繁育记录自动化核对项目（Demo 展示）。一步自动核对牛场每日手工录入的繁育流水记录（发情、配种、妊检、产犊、用药），自动识别录入错误与需要人工复核的异常逻辑。

**所有数据均为模拟数据。** 由 `make_demo_data.py` 随机生成，不包含任何真实牛场数据。

### 核对规则

| 规则编码 | 级别 | 核对内容 |
|---|---|---|
| E01 | 错误 | 耳号不在牛群名册中（如漏输0、多输一位、字母O误输为0、中间带空格） |
| E02 | 错误 | 日期为空、格式非法（如13月、「9月5日」）或晚于当日 |
| E03 | 错误 | 完全相同的记录重复录入 |
| E04 | 错误 | 操作人为空 |
| E05 | 错误 | 事件名称错别字（如「用要」或「用 药」） |
| W01 | 复核 | 产犊后 45 天内配种 |
| W02 | 复核 | 配种后 28 天内妊检 |
| W03 | 复核 | 两次发情间隔不在 18–24 天正常周期内 |

各项判断阈值位于 `check.py` 头部，可根据牧场自身生产管理规范调整。

### 使用方法

```bash
uv run make_demo_data.py                                   # 生成 demo/繁育记录-演示.xlsx
uv run check.py demo/繁育记录-演示.xlsx --today 2026-10-06  # 输出 …-核对结果.xlsx
uv run pytest -q                                           # 运行自动化测试
```

核验结果包含三个工作表：**汇总** 展示各规则命中统计；**问题清单** 逐条列出异常明细；**原表标注** 保持原表格格式，将错误标红、待复核标黄。

下方截图展示了上方命令生成的 **问题清单** 表格（默认随机种子 7）：403 条记录中检出 20 处错误与 8 处待复核记录。

![问题清单工作表截图](docs/check-result.png)

### 输入格式

输入为包含两个工作表的 Excel 文件：

- `繁育记录`：序号、日期、耳号、事件、详情、操作人
- `牛群名册`：耳号、出生日期、胎次

### 自动化测试说明

数据生成器在生成测试数据时，会记录所有故意植入的异常（标准答案）。自动化测试要求核对算法的检出结果与标准答案**完全一致**：既不能遗漏任何已知问题，也不能误报任何正常记录。测试套件在 5 组不同随机数种子下均能 100% 验证通过。
