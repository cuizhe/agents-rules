# 农业项目统计技能使用示例

---

## 示例 1：基本统计流程

### 场景
对 2026 年 3 月的农业项目进行当月统计，输出完整的 4 个工作表。

### 代码

```python
from pathlib import Path
from scripts.agriculture_stats import AgricultureStats

stats = AgricultureStats(
    input_file=Path('references/2026年电信与AIoT业务线FP项目毛利管控表20260310.xlsx'),
    output_file=Path('references/农业项目群预算毛利表汇总表-26.2.28.xlsx'),
    current_month=3,
    year=2026,
    sheet_name='实际运营数据-崔哲',  # 默认工作表
)

result = stats.run()
print(result)
```

### 输出

```
======================================================================
农业项目统计
======================================================================
读取输入文件：2026年电信与AIoT业务线FP项目毛利管控表20260310.xlsx
找到 1 个数据工作表：['实际运营数据-崔哲']
  处理工作表：实际运营数据-崔哲

共解析 34 个项目
农业项目：13 个
  - [P202212000496] 东部地区畜禽遗传资源基因库建设项目仪器设备采购-软件...  结束:2026-10-31
  - [P202601001368] 全国畜牧总站-2026年畜牧核心数据查询APP（牧数查）项目-FP项目...  结束:2026-12-31
  ...

项目分类（截止 2026-02-28）：
  当月在行：10 个
  当年在行：11 个
  已结项：2 个

生成输出文件：农业项目群预算毛利表汇总表-26.2.28.xlsx
  已有工作表：['当月汇总毛利率', '2026年汇总毛利率', '当月项目明细数据', '2026年项目明细数据']
  当月项目明细数据：10 个项目 × 9科目 = 90 行（重建 220 个合并单元格）
  2026年项目明细数据：11 个项目 × 9科目 = 99 行（重建 242 个合并单元格）
  当月汇总毛利率：9 个科目
  2026年汇总毛利率：9 个科目
  年月列填充完成（2026年）
  保存成功：references\农业项目群预算毛利表汇总表-26.2.28.xlsx

======================================================================
统计完成
======================================================================
```

---

## 示例 2：自定义关键字统计

### 场景
项目命名方式特殊，默认关键字无法识别，需要添加特定关键字。

### 方法 1：编辑配置文件

编辑 `references/agriculture_keywords.md`：

```markdown
## 关键字列表

农业
农村
农田
种植
养殖
畜牧
渔业
农机
种业
农资
农产品
畜禽
林业
农场
审计署
兽医
中监所
全国畜牧总站
乡村振兴
农垦    ← 新增
粮食局  ← 新增
```

### 方法 2：传入自定义关键字列表

```python
from scripts.agriculture_stats import AgricultureStats

# 直接修改类属性
stats = AgricultureStats(input_file=..., output_file=...)
stats.include_kw = ['农业', '农村', '种植', '农垦', '粮食局']
stats.exclude_kw = ['银行', '保险']

stats.load_data()
```

---

## 示例 3：只获取汇总数据，不写文件

### 场景
只需要看汇总指标，不需要生成 Excel 文件。

```python
from scripts.agriculture_stats import AgricultureStats

stats = AgricultureStats(input_file='input.xlsx', output_file='dummy.xlsx')
stats.load_data()
current_month_list, year_list, closed_list = stats.classify_projects()

# 获取年度汇总
agg_year = stats.aggregate_by_category(year_list)
print("年度汇总 GP1 毛利：", agg_year['GP1毛利']['current_actual'])
print("年度汇总 GP3 毛利：", agg_year['GP3毛利']['current_actual'])

# 获取毛利率
revenue = agg_year['累计收入（NR）']['current_actual']
gp1 = agg_year['GP1毛利']['current_actual']
gp3 = agg_year['GP3毛利']['current_actual']
print(f"GP1 毛利率：{gp1/revenue*100:.2f}%" if revenue else "无收入")
print(f"GP3 毛利率：{gp3/revenue*100:.2f}%" if revenue else "无收入")
```

---

## 示例 4：指定月份进行历史统计

### 场景
查看 2026 年 2 月的统计数据（截止到 2 月底的在行项目）。

```python
from scripts.agriculture_stats import AgricultureStats

stats = AgricultureStats(
    input_file='input.xlsx',
    output_file='output_202602.xlsx',
    current_month=2,  # 指定2月
    year=2026,
)

result = stats.run()
```

---

## 示例 5：分析特定项目的各核算科目明细

### 场景
查看某个农业项目的完整 9 行核算数据。

```python
from scripts.agriculture_stats import AgricultureStats
from scripts.agriculture_stats import CATEGORY_ORDER

stats = AgricultureStats(input_file='input.xlsx', output_file='dummy.xlsx')
stats.load_data()

# 找到目标项目
target_name = '农业农村部信息中心'
target = next((p for p in stats.agri_projects if target_name in p.name), None)

if target:
    print(f"项目：{target.name}")
    print(f"结束时间：{target.end_date}")
    print()
    for cat in CATEGORY_ORDER:
        if cat in target.metrics:
            m = target.metrics[cat]
            print(f"{cat}:")
            print(f"  预算={m['budget']:,.2f}, 实际={m['current_actual']:,.2f}, 偏差={m['偏差']:,.2f}")
```

---

## 示例 6：命令行运行

### 基本用法

```bash
cd C:/Users/Administrator/WorkBuddy/20260323145210

python .workbuddy/skills/agriculture-stats/scripts/agriculture_stats.py \
  -i "references/2026年电信与AIoT业务线FP项目毛利管控表20260310.xlsx" \
  -o "references/农业项目群预算毛利表汇总表-26.2.28.xlsx"
```

### 指定月份

```bash
python scripts/agriculture_stats.py \
  -i input.xlsx \
  -o output.xlsx \
  -m 2 \
  -y 2026
```

### 使用自定义关键字文件

```bash
python scripts/agriculture_stats.py \
  -i input.xlsx \
  -o output.xlsx \
  -k references/agriculture_keywords.md
```

### 切换数据工作表

```bash
# 切换到李陶的工作表
python scripts/agriculture_stats.py \
  -i input.xlsx \
  -o output.xlsx \
  -s "实际运营数据-李陶"
```

---

## 示例 7：解读汇总毛利率表

### 场景
生成汇总表后，如何理解各行数据。

```python
# 汇总表结构
agg = stats.aggregate_by_category(year_list)

# 关键指标解读
revenue = agg['累计收入（NR）']['current_actual']  # 累计收入
labor = agg['人工成本（成本１）']['current_actual']  # 人工成本
travel = agg['差旅费（成本３）']['current_actual']   # 差旅费
other_cost = agg['其他成本（成本３）']['current_actual']  # 其他成本
total_cost = agg['累计成本']['current_actual']  # 累计成本
gp1 = agg['GP1毛利']['current_actual']          # GP1毛利
gp3 = agg['GP3毛利']['current_actual']          # GP3毛利

print(f"项目整体毛利目标预算：{agg['GP1毛利']['budget']:,.2f}")
print(f"当前实际 GP1 毛利：{gp1:,.2f}")
print(f"GP1 毛利偏差：{agg['GP1毛利']['偏差']:,.2f}")
print(f"GP1 毛利率：{gp1/revenue*100:.2f}%" if revenue else "N/A")
print(f"GP3 毛利率：{gp3/revenue*100:.2f}%" if revenue else "N/A")
```

---

## 最佳实践

### 1. 关键字维护
- 定期检查 `agriculture_keywords.md`，确保覆盖所有农业项目命名规则
- 项目名称中出现新型农业业务时，及时补充关键字

### 2. 月度统计流程
```
每月初执行：
1. 更新输入文件（新增项目数据）
2. 运行统计脚本
3. 检查"当月在行"和"当月汇总毛利率"工作表
4. 如有偏差过大的项目，检查原因分析列
```

### 3. 数据验证
- 汇总表的预算合计应与输入表各项目预算之和一致
- GP1毛利 = 累计收入 - 直接成本合计（人工+差旅+其他）
- GP3毛利与GP1毛利的差异反映间接成本