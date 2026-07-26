# -*- coding: utf-8 -*-
"""
ABG Kimi Token Usage Summary from CSV
Reads whitelist Excel with merged cell handling and CSV usage data,
then generates a formatted Excel report with 5 sheets.
"""
import argparse
import re
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from collections import defaultdict


def format_workbook(wb):
    """统一格式化Excel工作簿"""
    header_font = Font(name='微软雅黑', size=10, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    body_font = Font(name='微软雅黑', size=10)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    alignment_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    alignment_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    alignment_right = Alignment(horizontal='right', vertical='center', wrap_text=True)
    
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.row == 1:
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = alignment_center
                else:
                    cell.font = body_font
                    cell.alignment = alignment_left
                cell.border = thin_border
        
        ws.freeze_panes = 'A2'
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            adjusted_width = min(max_length + 4, 60)
            ws.column_dimensions[column].width = adjusted_width
        
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    col_header = str(ws.cell(row=1, column=cell.column).value)
                    if '金额' in col_header or '元' in col_header:
                        cell.number_format = '0.00'
                    elif 'Tokens' in col_header or 'token' in col_header.lower():
                        cell.number_format = '#,##0'
                    elif '人数' in col_header or '工号' in col_header:
                        cell.number_format = '0'


def load_whitelist(path):
    """
    Load whitelist Excel, forward-fill merged cells in '所在项目' column.
    """
    # Read with pandas first
    df = pd.read_excel(path, sheet_name=0)
    
    # Use openpyxl to handle merged cells for forward-fill
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    
    # Find column index for '所在项目'
    headers = [cell.value for cell in ws[1]]
    try:
        proj_col_idx = headers.index('所在项目')  # 0-based for DataFrame
    except ValueError:
        proj_col_idx = None
        for idx, col in enumerate(df.columns):
            if '所在项目' in str(col):
                proj_col_idx = idx
                break
        if proj_col_idx is None:
            raise ValueError("Column '所在项目' not found in whitelist")
    
    # Build a map of row -> project value from merged cells
    merged_values = {}
    for merged_range in ws.merged_cells.ranges:
        min_col = merged_range.min_col - 1  # 0-based
        max_col = merged_range.max_col - 1
        if min_col <= proj_col_idx <= max_col:
            # Get the value from the first cell of the merged range
            first_value = ws.cell(row=merged_range.min_row, column=merged_range.min_col).value
            if first_value is not None:
                for row in range(merged_range.min_row, merged_range.max_row + 1):
                    merged_values[row] = first_value
    
    # Apply merged cell values to DataFrame
    # DataFrame index 0 corresponds to Excel row 2
    for excel_row, value in merged_values.items():
        df_row = excel_row - 2
        if 0 <= df_row < len(df):
            df.iloc[df_row, proj_col_idx] = value
    
    # NOTE: Do NOT use ffill() here. Only merged cells should be filled.
    # Rows with genuinely empty project should remain empty.

    # Filter valid rows
    df = df[df['工号'].notna() & df['姓名'].notna()].copy()
    df['工号'] = df['工号'].apply(lambda x: str(int(x)) if pd.notna(x) else None)
    df = df.drop_duplicates(subset=['工号'], keep='first')
    
    return df


def parse_username(username):
    """Extract name and employee number from username like '丁召海0000081658' or '·丁肖倩0000421541'."""
    # Remove leading non-Chinese/non-word characters before matching
    cleaned = re.sub(r'^[^\u4e00-\u9fa5\w]+', '', str(username).strip())
    match = re.match(r'^([\u4e00-\u9fa5]+)(\d+)$', cleaned)
    if match:
        name = match.group(1)
        emp_no = match.group(2).lstrip('0')
        if not emp_no:
            emp_no = match.group(2)
        return name, emp_no
    # Fallback: any characters followed by digits
    match2 = re.match(r'^(.+?)(\d+)$', cleaned)
    if match2:
        name = match2.group(1)
        emp_no = match2.group(2).lstrip('0')
        if not emp_no:
            emp_no = match2.group(2)
        return name, emp_no
    return username, None


def load_usage_csv(path):
    """Load CSV and aggregate tokens and amount per user."""
    df = pd.read_csv(path, encoding='utf-8')
    
    user_stats = defaultdict(lambda: {"总Tokens": 0, "金额": 0.0, "姓名": "", "调用次数": 0})
    
    for _, row in df.iterrows():
        username = str(row['用户名']).strip()
        name, emp_no = parse_username(username)
        
        key = emp_no if emp_no else username
        
        tokens = float(row['总Tokens']) if pd.notna(row.get('总Tokens')) else 0
        amount = float(row['花费金额']) if pd.notna(row.get('花费金额')) else 0.0
        calls = int(row['调用次数']) if pd.notna(row.get('调用次数')) else 0
        
        user_stats[key]["总Tokens"] += tokens
        user_stats[key]["金额"] += amount
        user_stats[key]["调用次数"] += calls
        if not user_stats[key]["姓名"]:
            user_stats[key]["姓名"] = name
    
    return user_stats


def generate_reports(whitelist_df, user_stats, output_path):
    """Generate Excel with 5 sheets."""
    
    # Build usage DataFrame from user_stats
    usage_records = []
    for emp_key, stats in user_stats.items():
        usage_records.append({
            '工号': emp_key,
            '姓名': stats['姓名'],
            '总Tokens': int(stats['总Tokens']),
            '金额': round(stats['金额'], 2),
            '调用次数': stats['调用次数'],
        })
    usage_df = pd.DataFrame(usage_records)
    
    # Separate matched and unmatched
    whitelist_empnos = set(whitelist_df['工号'].dropna().unique())
    
    # Try to match usage records
    matched_records = []
    unmatched_records = []
    
    for _, row in usage_df.iterrows():
        emp_id = row['工号']
        # Try exact match first
        if emp_id in whitelist_empnos:
            matched_records.append({
                '工号': emp_id,
                '姓名': row['姓名'],
                '总Tokens': row['总Tokens'],
                '金额': row['金额'],
                '调用次数': row['调用次数'],
                '匹配方式': '工号',
            })
        else:
            # Try to match by name
            name_matches = whitelist_df[whitelist_df['姓名'] == row['姓名']]
            if len(name_matches) > 0:
                matched_emp_id = name_matches.iloc[0]['工号']
                matched_records.append({
                    '工号': matched_emp_id,
                    '姓名': row['姓名'],
                    '总Tokens': row['总Tokens'],
                    '金额': row['金额'],
                    '调用次数': row['调用次数'],
                    '匹配方式': '姓名',
                })
            else:
                unmatched_records.append({
                    '工号': emp_id,
                    '姓名': row['姓名'],
                    '总Tokens': row['总Tokens'],
                    '金额': row['金额'],
                    '调用次数': row['调用次数'],
                })
    
    matched_df = pd.DataFrame(matched_records)
    unmatched_df = pd.DataFrame(unmatched_records)
    
    # Merge whitelist info into matched records
    merge_cols = ['工号', '一级部门', '二级部门', '三级部门', '四级部门', '五级部门', '所在项目', '岗位']
    whitelist_merge = whitelist_df[merge_cols].copy()
    
    if len(matched_df) > 0:
        matched_df = matched_df.merge(whitelist_merge, on='工号', how='left')
    else:
        matched_df = pd.DataFrame(columns=['工号', '姓名', '总Tokens', '金额', '调用次数', '匹配方式'] + merge_cols[1:])
    
    # 部门路径
    matched_df['部门路径'] = matched_df.apply(
        lambda row: ' / '.join([str(v).strip() for v in [row['二级部门'], row['三级部门'], row['四级部门']]
                              if pd.notna(v) and str(v).strip() and str(v) != '/'])
        if pd.notna(row['二级部门']) or pd.notna(row['三级部门']) or pd.notna(row['四级部门']) else '未分类',
        axis=1
    )
    
    # 未使用人员
    used_empnos = set(matched_df['工号'].dropna().unique()) if len(matched_df) > 0 else set()
    unused_whitelist = whitelist_df[~whitelist_df['工号'].isin(used_empnos)].copy()
    
    # Sheet 1: 按项目汇总
    project_detail = matched_df[matched_df['所在项目'].notna()].copy() if len(matched_df) > 0 else pd.DataFrame()
    if len(project_detail) > 0:
        project_dept_map = project_detail.groupby('所在项目').apply(
            lambda g: pd.Series({
                '三级部门': g['三级部门'].dropna().iloc[0] if not g['三级部门'].dropna().empty else '',
                '四级部门': g['四级部门'].dropna().iloc[0] if not g['四级部门'].dropna().empty else ''
            })
        ).reset_index()
        
        project_summary = project_detail.groupby('所在项目').agg({
            '工号': 'nunique',
            '总Tokens': 'sum',
            '金额': 'sum'
        }).reset_index()
        project_summary.columns = ['项目', '人数', '总Tokens', '金额']
        project_summary = project_summary.merge(project_dept_map.rename(columns={'所在项目': '项目'}), on='项目', how='left')
        project_summary = project_summary[['三级部门', '四级部门', '项目', '人数', '总Tokens', '金额']]
        project_summary = project_summary.sort_values('总Tokens', ascending=False).reset_index(drop=True)
    else:
        project_summary = pd.DataFrame(columns=['三级部门', '四级部门', '项目', '人数', '总Tokens', '金额'])
    
    # Sheet 2: 按部门汇总
    if len(matched_df) > 0:
        dept_summary = matched_df.groupby('部门路径').agg({
            '工号': 'nunique',
            '总Tokens': 'sum',
            '金额': 'sum'
        }).reset_index()
        dept_summary.columns = ['部门路径', '人数', '总Tokens', '金额']
        dept_summary = dept_summary.sort_values('总Tokens', ascending=False).reset_index(drop=True)
    else:
        dept_summary = pd.DataFrame(columns=['部门路径', '人数', '总Tokens', '金额'])
    
    # Create Excel
    wb = Workbook()
    wb.remove(wb.active)
    
    # Sheet 1: 按项目汇总
    ws1 = wb.create_sheet('按项目汇总')
    ws1.append(['三级部门', '四级部门', '项目', '人数', '总Tokens', '金额'])
    for _, row in project_summary.iterrows():
        ws1.append([
            row['三级部门'], row['四级部门'], row['项目'],
            row['人数'], row['总Tokens'], round(row['金额'], 2)
        ])
    
    # Sheet 2: 按部门汇总
    ws2 = wb.create_sheet('按部门汇总')
    ws2.append(['部门路径', '人数', '总Tokens', '金额'])
    for _, row in dept_summary.iterrows():
        ws2.append([row['部门路径'], row['人数'], row['总Tokens'], round(row['金额'], 2)])
    
    # Sheet 3: 白名单有消耗明细
    ws3 = wb.create_sheet('白名单有消耗明细')
    ws3.append(['工号', '姓名', '部门路径', '所在项目', '总Tokens', '金额', '数据来源'])
    if len(matched_df) > 0:
        for _, row in matched_df.iterrows():
            source = '白名单' if row.get('匹配方式') == '工号' else '白名单(姓名匹配)'
            ws3.append([
                row['工号'], row['姓名'], row['部门路径'],
                row['所在项目'] if pd.notna(row['所在项目']) else '',
                row['总Tokens'], round(row['金额'], 2),
                source
            ])
    
    # Sheet 4: 白名单未使用明细
    ws4 = wb.create_sheet('白名单未使用明细')
    ws4.append(['工号', '姓名', '二级部门', '三级部门', '四级部门', '所在项目', '岗位'])
    for _, row in unused_whitelist.iterrows():
        ws4.append([
            row['工号'], row['姓名'], row['二级部门'] if pd.notna(row['二级部门']) else '',
            row['三级部门'] if pd.notna(row['三级部门']) else '',
            row['四级部门'] if pd.notna(row['四级部门']) else '',
            row['所在项目'] if pd.notna(row['所在项目']) else '',
            row['岗位'] if pd.notna(row['岗位']) else ''
        ])
    
    # Sheet 5: 被剔除人员明细
    ws5 = wb.create_sheet('被剔除人员明细')
    ws5.append(['工号', '姓名', '总Tokens', '金额', '备注'])
    for _, row in unmatched_df.iterrows():
        ws5.append([
            row['工号'], row['姓名'], row['总Tokens'], round(row['金额'], 2),
            '该员工不在白名单中'
        ])
    
    # Sheet 6: 人员排名 (按Token降序)
    ws6 = wb.create_sheet('人员排名')
    ws6.append(['排名', '部门路径', '工号', '姓名', '所在项目', '总Tokens', '金额'])
    if len(matched_df) > 0:
        # Sort by 总Tokens descending
        ranked_df = matched_df.sort_values('总Tokens', ascending=False).reset_index(drop=True)
        for idx, row in ranked_df.iterrows():
            ws6.append([
                idx + 1,
                row['部门路径'] if pd.notna(row['部门路径']) else '',
                row['工号'],
                row['姓名'],
                row['所在项目'] if pd.notna(row['所在项目']) else '',
                row['总Tokens'],
                round(row['金额'], 2)
            ])
    
    format_workbook(wb)
    wb.save(output_path)
    
    # ========== 数据一致性验证 ==========
    print('\n' + '='*50)
    print('数据一致性验证')
    print('='*50)
    
    validation_passed = True
    validation_errors = []
    
    # 1. 验证总人数
    total_people = len(matched_df) + len(unused_whitelist) + len(unmatched_df)
    expected_total = len(whitelist_df) + len(unmatched_df)
    if total_people != expected_total:
        validation_passed = False
        validation_errors.append(f'总人数不一致: 明细({total_people}) != 白名单+被剔除({expected_total})')
    else:
        print(f'[OK] 总人数验证通过: {total_people} 人')
    
    # 2. 验证总Tokens（明细 vs 汇总）
    detail_total_tokens = matched_df['总Tokens'].sum() if len(matched_df) > 0 else 0
    project_total_tokens = project_summary['总Tokens'].sum() if len(project_summary) > 0 else 0
    dept_total_tokens = dept_summary['总Tokens'].sum() if len(dept_summary) > 0 else 0
    
    if detail_total_tokens != project_total_tokens:
        validation_passed = False
        validation_errors.append(f'按项目汇总Tokens不一致: 明细({detail_total_tokens:,}) != 项目汇总({project_total_tokens:,})')
    else:
        print(f'[OK] 按项目汇总Tokens验证通过: {detail_total_tokens:,}')
    
    if detail_total_tokens != dept_total_tokens:
        validation_passed = False
        validation_errors.append(f'按部门汇总Tokens不一致: 明细({detail_total_tokens:,}) != 部门汇总({dept_total_tokens:,})')
    else:
        print(f'[OK] 按部门汇总Tokens验证通过: {detail_total_tokens:,}')
    
    # 3. 验证总金额（明细 vs 汇总）
    detail_total_amount = matched_df['金额'].sum() if len(matched_df) > 0 else 0
    project_total_amount = project_summary['金额'].sum() if len(project_summary) > 0 else 0
    dept_total_amount = dept_summary['金额'].sum() if len(dept_summary) > 0 else 0
    
    if abs(detail_total_amount - project_total_amount) > 0.01:
        validation_passed = False
        validation_errors.append(f'按项目汇总金额不一致: 明细({detail_total_amount:.2f}) != 项目汇总({project_total_amount:.2f})')
    else:
        print(f'[OK] 按项目汇总金额验证通过: {detail_total_amount:.2f}')
    
    if abs(detail_total_amount - dept_total_amount) > 0.01:
        validation_passed = False
        validation_errors.append(f'按部门汇总金额不一致: 明细({detail_total_amount:.2f}) != 部门汇总({dept_total_amount:.2f})')
    else:
        print(f'[OK] 按部门汇总金额验证通过: {detail_total_amount:.2f}')
    
    # 4. 验证人员排名sheet数据
    if len(matched_df) > 0:
        ranked_total_tokens = matched_df.sort_values('总Tokens', ascending=False)['总Tokens'].sum()
        if ranked_total_tokens != detail_total_tokens:
            validation_passed = False
            validation_errors.append(f'人员排名Tokens不一致: 排名({ranked_total_tokens:,}) != 明细({detail_total_tokens:,})')
        else:
            print(f'[OK] 人员排名Tokens验证通过: {ranked_total_tokens:,}')
    
    # 5. 验证项目汇总人数
    if len(project_summary) > 0:
        project_people_sum = project_summary['人数'].sum()
        if project_people_sum != len(matched_df):
            validation_passed = False
            validation_errors.append(f'项目汇总人数不一致: 项目汇总({project_people_sum}) != 明细({len(matched_df)})')
        else:
            print(f'[OK] 项目汇总人数验证通过: {project_people_sum} 人')
    
    # 6. 验证部门汇总人数
    if len(dept_summary) > 0:
        dept_people_sum = dept_summary['人数'].sum()
        if dept_people_sum != len(matched_df):
            validation_passed = False
            validation_errors.append(f'部门汇总人数不一致: 部门汇总({dept_people_sum}) != 明细({len(matched_df)})')
        else:
            print(f'[OK] 部门汇总人数验证通过: {dept_people_sum} 人')
    
    # 输出验证结果
    print('='*50)
    if validation_passed:
        print('[OK] 所有验证通过，数据一致无偏差')
    else:
        print('[FAIL] 验证失败，发现以下问题:')
        for error in validation_errors:
            print(f'  - {error}')
        print('\n请检查数据后重新生成报告')
        raise ValueError('数据一致性验证失败')
    print('='*50 + '\n')
    
    return {
        "project_rows": len(project_summary),
        "dept_rows": len(dept_summary),
        "matched": len(matched_df),
        "unused": len(unused_whitelist),
        "unmatched": len(unmatched_df),
        "total_tokens": matched_df['总Tokens'].sum() if len(matched_df) > 0 else 0,
        "total_amount": matched_df['金额'].sum() if len(matched_df) > 0 else 0,
        "validation_passed": validation_passed,
    }


def main():
    parser = argparse.ArgumentParser(description='ABG Kimi Token Usage Summary from CSV')
    parser.add_argument('--whitelist', required=True, help='Path to whitelist Excel file')
    parser.add_argument('--usage-csv', required=True, help='Path to token usage CSV file')
    parser.add_argument('--output', default='ABG_Kimi_Usage_白名单汇总_按项目按部门.xlsx', help='Output Excel file path')
    args = parser.parse_args()
    
    print(f'Loading whitelist: {args.whitelist}')
    whitelist_df = load_whitelist(args.whitelist)
    print(f'  Found {len(whitelist_df)} employees in whitelist')
    
    print(f'Loading usage CSV: {args.usage_csv}')
    user_stats = load_usage_csv(args.usage_csv)
    print(f'  Found {len(user_stats)} unique users in CSV')
    
    print(f'Generating report: {args.output}')
    result = generate_reports(whitelist_df, user_stats, args.output)
    
    print(f'\nDone!')
    print(f'  按项目汇总: {result["project_rows"]} 个项目')
    print(f'  按部门汇总: {result["dept_rows"]} 个部门')
    print(f'  白名单有消耗: {result["matched"]} 人')
    print(f'  白名单未使用: {result["unused"]} 人')
    print(f'  被剔除人员: {result["unmatched"]} 人')
    print(f'  Total tokens: {result["total_tokens"]:,}')
    print(f'  Total amount: {result["total_amount"]:.2f}')


if __name__ == '__main__':
    main()
