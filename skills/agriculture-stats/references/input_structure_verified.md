# 输入表结构说明（已验证）

**文件**: `2026 年电信与 AIoT 业务线 FP 项目毛利管控表 20260310.xlsx`  
**工作表**: `实际运营数据 - 崔哲`  
**更新时间**: 2026-03-20

---

## 表结构关键信息

### 1. 表头位置
- **表头在第 1 行**
- 表头包含的列（根据截图）：
  - 项目编码
  - 项目名称
  - 项目开始时间
  - 项目结束时间
  - 2 月项目人数（黄色高亮）
  - 结算方式
  - ...（其他列）

### 2. 数据行结构
- **每个项目占用 9 行数据**
- 前 6 列（项目编码、项目名称、项目开始时间、项目结束时间、项目人数、结算方式）存在**行合并**
  - 即这 6 列在 9 行中只显示一次（合并单元格）
- **"类别"列**（或"核算"列）：每行一个指标，9 行分别为：
  1. 累计收入 (NR)
  2. 人工成本_直接成本_...
  3. 差旅费_直接成本_...
  4. 采购成本_直接成本_...
  5. GP1 毛利
  6. GP1 毛利率 (%)
  7. GP3 毛利
  8. GP3 毛利率 (%)
  9. （可能还有其他指标）

### 3. 数据示例结构

```
行 1:  [表头] 项目编码 | 项目名称 | 项目开始时间 | 项目结束时间 | 2 月项目人数 | 结算方式 | 核算方式 | 类别 | 预算 | 实际 | ...
行 2:  [项目 A 数据 - 第 1 行] P001(合并) | 项目 A(合并) | 2025-01-01(合并) | 2026-05-01(合并) | 10(合并) | 总额法 (合并) | ... | 累计收入 | 数值 | 数值 | ...
行 3:  [项目 A 数据 - 第 2 行] (合并)    | (合并)      | (合并)         | (合并)         | (合并)      | (合并)      | ... | 人工成本 | 数值 | 数值 | ...
行 4:  [项目 A 数据 - 第 3 行] (合并)    | (合并)      | (合并)         | (合并)         | (合并)      | (合并)      | ... | 差旅费 | 数值 | 数值 | ...
...
行 10: [项目 A 数据 - 第 9 行] (合并)    | (合并)      | (合并)         | (合并)         | (合并)      | (合并)      | ... | GP3 毛利率 | 数值 | 数值 | ...
行 11: [项目 B 数据 - 第 1 行] P002      | 项目 B      | ...
```

---

## 项目归属规则

### 当月在行项目
**规则**: 项目结束时间 > 上月最后一天

```python
# 2026 年 3 月为例
上月最后一天 = 2026-02-28
当月在行 = 项目结束时间 > 2026-02-28
```

### 当年在行项目
**规则**: 项目结束时间 > 2026-01-01

```python
当年在行 = 项目结束时间 > 2026-01-01
```

### 关系说明
- **当月在行项目** 一定是 **当年在行项目** 的子集
- 当月在行：项目结束时间在 2026-03-01 及之后
- 当年在行但非当月在行：项目结束时间在 2026-01-02 至 2026-02-28 之间（即 1-2 月已结束的项目）

---

## 数据筛选与分发流程

### 步骤 1: 读取输入表
```python
import pandas as pd
from openpyxl import load_workbook

# 读取 Excel（保留合并单元格信息）
wb = load_workbook('输入文件.xlsx')
ws = wb['实际运营数据 - 崔哲']
```

### 步骤 2: 解析项目数据
```python
projects = []
current_project = None

for row_idx in range(2, ws.max_row + 1):  # 从第 2 行开始（跳过表头）
    # 检查是否是新的项目（第 1 列有值，表示新项目的开始）
    project_code = ws.cell(row=row_idx, column=1).value
    
    if project_code and project_code != '':  # 新项目开始
        if current_project:  # 保存上一个项目
            projects.append(current_project)
        
        # 读取项目基本信息（合并单元格的值在第 1 行）
        current_project = {
            '项目编码': project_code,
            '项目名称': ws.cell(row=row_idx, column=2).value,
            '项目开始时间': ws.cell(row=row_idx, column=3).value,
            '项目结束时间': ws.cell(row=row_idx, column=4).value,
            '2 月项目人数': ws.cell(row=row_idx, column=5).value,
            '结算方式': ws.cell(row=row_idx, column=6).value,
            'metrics': []  # 存储 9 个指标的数据
        }
    
    # 读取当前行的指标数据
    if current_project:
        metric_name = ws.cell(row=row_idx, column=8).value  # 假设类别在第 8 列
        metric_data = {
            '类别': metric_name,
            # 读取其他列的数据...
        }
        current_project['metrics'].append(metric_data)

# 添加最后一个项目
if current_project:
    projects.append(current_project)
```

### 步骤 3: 筛选农业项目
```python
from pathlib import Path

def load_agriculture_keywords():
    """加载农业关键字"""
    include_keywords = []
    exclude_keywords = []
    
    # 读取 agriculture_keywords.md 文件
    # ...
    
    return include_keywords, exclude_keywords

def is_agriculture_project(project_name, include_kw, exclude_kw):
    """判断是否为农业项目"""
    # 检查排除关键字
    for kw in exclude_kw:
        if kw in project_name:
            return False
    
    # 检查农业关键字
    for kw in include_kw:
        if kw in project_name:
            return True
    
    return False

# 筛选
include_kw, exclude_kw = load_agriculture_keywords()
agri_projects = [
    p for p in projects 
    if is_agriculture_project(p['项目名称'], include_kw, exclude_kw)
]
```

### 步骤 4: 按项目结束时间分类
```python
from datetime import datetime

# 上月最后一天（2026 年 2 月）
last_month_end = datetime(2026, 2, 28)
year_start = datetime(2026, 1, 1)

current_month_projects = []  # 当月在行
year_projects = []  # 当年在行

for project in agri_projects:
    end_date = project['项目结束时间']
    
    if end_date > last_month_end:
        current_month_projects.append(project)
    
    if end_date > year_start:
        year_projects.append(project)
```

### 步骤 5: 写入输出表
```python
# 创建输出 Excel
from openpyxl import Workbook

output_wb = Workbook()

# 写入当月项目明细
ws_current = output_wb.create_sheet('当月项目明细数据')
write_project_data(ws_current, current_month_projects)

# 写入当年项目明细
ws_year = output_wb.create_sheet('2026 年项目明细数据')
write_project_data(ws_year, year_projects)

output_wb.save('农业项目群预算毛利表汇总表 -26.2.28.xlsx')
```

---

## 关键注意事项

### 1. 合并单元格处理
- 使用 `openpyxl` 读取时可以获取合并单元格的值
- 每个项目的第 1 行包含完整的项目基本信息
- 后续 8 行只包含指标相关的数据

### 2. 保持原格式
- **不改变输入表数据单元格数值**
- **不改变格式**（如黄色高亮、数字格式等）
- **不改变公式**（如果有）

### 3. 包含表头
- 输出表必须包含完整的表头（第 1 行）
- 表头结构与输入表一致

### 4. 项目归属判断
- 使用 `项目结束时间` 列进行判断
- 注意日期格式转换和比较

---

## 待确认信息

### 1. 列位置确认
请确认以下列的位置（列号）：
- [ ] 项目编码：第 1 列？
- [ ] 项目名称：第 2 列？
- [ ] 项目开始时间：第 3 列？
- [ ] 项目结束时间：第 4 列？
- [ ] 2 月项目人数：第 5 列？
- [ ] 结算方式：第 6 列？
- [ ] 核算方式/类别：第几列？

### 2. 9 个指标的具体内容
请确认每个项目的 9 行分别是什么指标：
1. 累计收入 (NR)
2. ?
3. ?
4. ?
5. ?
6. ?
7. ?
8. ?
9. ?

### 3. 输出表工作表名称
确认输出表的工作表名称：
- [ ] `当月项目明细数据`（对应 3 月在行项目）
- [ ] `2026 年项目明细数据`（对应 2026 年在行项目）

### 4. 其他输出表
是否需要同时更新：
- [ ] `2026 年项目毛利表`（汇总表）？
- [ ] `当月项目毛利表`（月汇总表）？

---

**下一步**: 请确认上述待确认信息，然后我将创建完整的数据转换脚本。
