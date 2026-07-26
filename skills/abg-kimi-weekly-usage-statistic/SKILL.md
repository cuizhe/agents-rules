---
name: abg-kimi-weekly-usage-statistic
description: Generate weekly ABG department Kimi token usage summary reports by department and project from CSV data. Use when the user provides a CSV file exported from the Kimi admin console and needs to (1) match employees against a whitelist Excel to get department/project info, (2) handle merged cells in the whitelist '所在项目' column, (3) generate formatted Excel reports with 5 sheets (按项目汇总, 按部门汇总, 白名单有消耗明细, 白名单未使用明细, 被剔除人员明细).
---

# ABG Kimi Weekly Usage Statistic

Generate formatted Excel reports summarizing ABG department employees' weekly Kimi token usage and costs, grouped by department hierarchy and project. This skill processes CSV files exported from the Kimi admin console.

## Workflow

### Step 1: Prepare Files

Ensure the following files are available in the working directory:
- **Whitelist Excel**: The employee whitelist file (e.g. `ABG FP项目申请中软网关Kimi账号白名单-6.18.xlsx`). Must contain: 工号, 姓名, 一级部门, 二级部门, 三级部门, 四级部门, 五级部门, 岗位, 所在项目. The "所在项目" column may contain merged cells — the script handles forward-fill automatically via openpyxl.
- **Usage CSV**: The Kimi admin console export (e.g. `token_usage_*.csv` or `week*.csv`). Must contain: 用户名, 总Tokens, 花费金额, 调用次数.

### Step 2: Run the CSV Report Generator

```bash
python .agents/skills/abg-kimi-weekly-usage-statistic/scripts/generate_csv_report.py \
  --whitelist "白名单.xlsx" \
  --usage-csv "token_usage_*.csv" \
  --output "ABG_Kimi_Usage_白名单汇总_按项目按部门.xlsx"
```

**Parameters:**
- `--whitelist` (required): Path to the whitelist Excel file
- `--usage-csv` (required): Path to the token usage CSV file
- `--output` (optional): Output Excel path, defaults to `ABG_Kimi_Usage_白名单汇总_按项目按部门.xlsx`

**What the script does:**
1. Loads the whitelist Excel, automatically forward-filling merged cells in the `所在项目` column using openpyxl
2. Parses the CSV, aggregating usage data per user by extracting name and employee number from the `用户名` field (e.g. `丁召海0000081658`)
3. Matches users against the whitelist by employee number (stripping leading zeros), falling back to name matching if needed
4. Classifies users into: 白名单有消耗 / 白名单未使用 / 被剔除人员
5. Generates a formatted Excel with 6 sheets
6. **Validates data consistency** by comparing detail records against summary totals

### Step 3: Verify the Output

The generated Excel contains six sheets:

**Sheet1 "按项目汇总":**
- Columns: 三级部门, 四级部门, 项目, 人数, 总Tokens, 金额
- Sorted by 总Tokens descending

**Sheet2 "按部门汇总":**
- Columns: 部门路径, 人数, 总Tokens, 金额
- Sorted by 总Tokens descending

**Sheet3 "白名单有消耗明细":**
- One row per matched employee with usage data
- Columns: 工号, 姓名, 部门路径, 所在项目, 总Tokens, 金额, 数据来源

**Sheet4 "白名单未使用明细":**
- Employees in whitelist but with zero usage in the CSV period
- Columns: 工号, 姓名, 二级部门, 三级部门, 四级部门, 所在项目, 岗位

**Sheet5 "被剔除人员明细":**
- Employees with usage in CSV but not found in the whitelist
- Columns: 工号, 姓名, 总Tokens, 金额, 备注

**Sheet6 "人员排名":**
- All matched employees sorted by 总Tokens descending
- Columns: 排名, 部门路径, 工号, 姓名, 所在项目, 总Tokens, 金额

**Formatting (all sheets):**
- Font: Microsoft YaHei (微软雅黑), size 10
- Header: Bold, white text on blue background (#4472C4)
- All cells: Black thin borders
- Number columns: Right-aligned, formatted with `#,##0` for Tokens and `0.00` for amounts
- Other columns: Left-aligned
- Frozen header row
- Auto-adjusted column widths

### Step 4: Data Consistency Validation

After generating the Excel, the script performs automatic validation:

| 验证项 | 验证内容 |
|---|---|
| 总人数 | 明细人数 = 白名单人数 + 被剔除人数 |
| 总Tokens | 明细Tokens = 按项目汇总Tokens = 按部门汇总Tokens |
| 总金额 | 明细金额 = 按项目汇总金额 = 按部门汇总金额 |
| 人员排名 | 排名sheet Tokens = 明细Tokens |
| 项目汇总人数 | 项目汇总人数 = 明细人数 |
| 部门汇总人数 | 部门汇总人数 = 明细人数 |

**验证通过**: 所有数据一致，报告可用
**验证失败**: 输出具体偏差信息，脚本抛出异常，需检查数据后重新生成

## Special Handling

### Employee ID Matching
- Strip leading zeros from `emp_id` before matching against whitelist
- Special case: If CSV shows "0000081658" but whitelist has "81658", the script handles this via `lstrip('0')`
- CSV usernames may contain leading non-Chinese characters (e.g. `·丁肖倩0000421541`) — these are cleaned before parsing

### Merged Cells in Whitelist
- The "所在项目" column in the whitelist Excel often contains merged cells
- The script uses openpyxl to detect merged cell ranges and fills only those rows within each merged range with the value from the first cell
- **Important**: Only merged cells are filled. Rows with genuinely empty project values remain empty — the script does NOT perform global forward-fill (`ffill`) across the entire column, to avoid propagating project names to unrelated employees

### Unmatched Employees
- Unmatched employees appear in the "被剔除人员明细" sheet with a note "该员工不在白名单中"
- Unmatched employees are excluded from the "按项目汇总" and "按部门汇总" sheets

## Typical Weekly Run

```bash
source .venv/Scripts/activate
python .agents/skills/abg-kimi-weekly-usage-statistic/scripts/generate_csv_report.py \
  --whitelist "ABG FP项目申请中软网关Kimi账号白名单-6.18.xlsx" \
  --usage-csv "week6.26_token_usage_1782469203911.csv" \
  --output "ABG_Kimi_Usage_618白名单汇总_按项目按部门_6.26周.xlsx"
```
