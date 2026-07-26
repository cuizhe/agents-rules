---
name: agriculture-stats
description: 汇总统计农业项目群当月及当年累计运营指标数据。根据用户提供的 Excel 输入模板和输出报告模板进行定制化统计。支持模糊查找文件（忽略全角/半角、空格）。当用户询问农业项目统计、运营指标、月度报表、年度累计数据时自动应用此技能。
---

# 农业项目群运营指标统计

## 功能说明

本技能从 `2026年电信与AIoT业务线FP项目毛利管控表` 中提取农业行业项目数据，生成 4 个输出工作表：

| 输出工作表       | 内容                 |
| ----------- | ------------------ |
| 当月项目明细数据    | 当月在行农业项目明细（10行/项目） |
| 2026年项目明细数据 | 所有在行农业项目明细         |
| 当月汇总毛利率     | 当月在行项目按核算科目汇总      |
| 2026年汇总毛利率  | 所有在行项目按核算科目汇总      |

**数据写入规则**：数据写入目标表必须严格遵守 rules.md 中excel 表数据操作规约

**农业项目识别**：通过项目名称匹配关键字（可配置）。

**项目分类规则**：

- 当月在行：项目结束时间 > 上月末
- 当年在行：项目结束时间 > 当年1月1日
- 已结项：项目结束时间 ≤ 当年1月1日

---

## 快速开始

### 一行命令完成统计

```bash
cd C:/Users/Administrator/WorkBuddy/20260323145210

python .workbuddy/skills/agriculture-stats/scripts/agriculture_stats.py \
  -i "references/2026年电信与AIoT业务线FP项目毛利管控表20260310.xlsx" \
  -o "references/农业项目群预算毛利表汇总表-26.2.28.xlsx" \
  -m 3 -y 2026
```

### Python API

```python
from pathlib import Path
from scripts.agriculture_stats import AgricultureStats

stats = AgricultureStats(
    input_file=Path('references/2026年电信与AIoT业务线FP项目毛利管控表20260310.xlsx'),
    output_file=Path('references/农业项目群预算毛利表汇总表-26.2.28.xlsx'),
    current_month=3,
    year=2026,
    sheet_name='实际运营数据-崔哲',  # 默认工作表，可切换
)

result = stats.run()
# {'total_projects': 20, 'agri_projects': 2, 'current_month_projects': 1, ...}
```

---

## 输入数据结构

**输入文件**：`2026年电信与AIoT业务线FP项目毛利管控表20260310.xlsx`

**数据工作表**：默认读取 **`实际运营数据-崔哲`**（可配置，变更 `sheet_name` 参数或 `DEFAULT_DATA_SHEET` 常量即可切换）。

**项目结构**：每个项目占 10 行（10 个核算科目）

| 列号     | 字段               | 说明                 |
| ------ | ---------------- | ------------------ |
| 1      | 项目编码             | 项目唯一标识             |
| 2      | 项目名称             | 项目全称（关键字匹配依据）      |
| 3      | 项目开始时间           | datetime           |
| 4      | 项目结束时间           | datetime（项目分类依据）   |
| 17     | 结算方式             | 工作量结算/总价进度结算       |
| **18** | **类别**           | **核算科目名称（每项目10行）** |
| 19     | 项目整体毛利目标预算       | 预算值                |
| 20     | 项目开始-当前实际累计      | 实际累计值              |
| 22-33  | 26年YTD实际累计12月-1月 | 2026年各月YTD（累计）     |
| 41-52  | 202612-202601    | 2026年各月实际值（当月）     |
| 53-55  | 2025/2024/2023年  | 历史年度数据             |

### 10 个核算科目

1. 累计收入（NR）
2. 人工成本（成本１）
3. **闲置成本（成本2）** ← 新增
4. 差旅费（成本３）
5. 其他成本（成本３）
6. 累计成本
7. GP1毛利
8. GP1毛利率(%)
9. GP3毛利
10. GP3毛利率(%)

---

## 输出数据结构

### 汇总毛利率表（当月 / 2026年）

按核算科目汇总所有农业项目：

| 列   | 内容                                 |
| --- | ---------------------------------- |
| A   | 项目状态说明（如"13个农业项目"）                 |
| B   | 核算科目                               |
| C   | 项目整体毛利目标预算（汇总）                     |
| D   | 项目开始-当前实际累计（汇总）                    |
| E   | **2026年YTD**（公式，自动计算）              |
| F   | 偏差（实际-预算）（公式，自动计算）                 |
| G   | 数据分析与说明                            |
| H-S | 2026年12月至1月月度数据（col8=12月，col19=1月） |

**累计成本行，GP行（GP1毛利/毛利率、GP3毛利/毛利率）**：所有列均为公式，不受数据刷新影响。

**2026年YTD 和 偏差**：由参考文件中的公式自动计算，脚本不覆盖。

### 项目明细表（当月 / 2026年）

每项目 10 行，每行一个核算科目。关键列：项目编码、项目名称、核算科目、预算、实际、YTD、偏差。

---

## 农业项目关键字配置

编辑 `references/agriculture_keywords.md`，或在 Python 中直接设置：

```python
stats.include_kw = ['农业', '农村', '种植', '农机', '乡村振兴']
stats.exclude_kw = ['银行', '保险']
```

---

## 工具函数

### file_utils（文件模糊查找）

```python
from scripts.file_utils import (
    normalize_text,           # 全角转半角，去空格，转小写
    find_file_fuzzy,           # 模糊查找单个文件
    find_files_fuzzy,          # 模糊查找所有匹配
    find_excel_sheet_fuzzy,    # 模糊查找 Excel 工作表
    find_input_file,           # 查找毛利管控表
    find_output_template_file,  # 查找预算毛利汇总表
)
```

---

## 参考文档

- [详细 API 文档](reference.md)
- [使用示例](examples.md)
- [文件查找工具](scripts/file_utils.py)
- [核心统计模块](scripts/agriculture_stats.py)