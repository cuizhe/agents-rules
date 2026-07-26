"""
汇总"当月项目明细数据"到"当月汇总毛利率"工作表

统计规则：
1. 将"当月项目明细数据"的所有项目按照以下核算方式汇总求和：
   - 累计收入（NR）
   - 人工成本（成本 1）
   - 差旅费（成本 3）
   - 其他成本（成本 3）
2. 对应填入到"当月汇总毛利率"工作表中的对应行
3. 列匹配规则：
   - "项目整体毛利目标预算"、"项目开始 - 当前实际累计"：直接按列名匹配
   - "2026 年 YTD"：使用用户指定的列（26 年 YTD 实际累计 2 月）
   - 年月列（如 2026-03-30）：按年月匹配（如 202603）
4. 如果目标单元格已有数据则跳过
5. 如果源数据为空也跳过该列
6. 只填充数值，不改变格式
"""

import openpyxl
from pathlib import Path
from datetime import datetime

def find_file_fuzzy(keyword, search_dir):
    """模糊查找文件"""
    search_dir = Path(search_dir)
    for file in search_dir.glob('*.xlsx'):
        if keyword in file.name:
            return file
    return None

def sum_by_accounting_method(ws_detail, accounting_method, month_col_header):
    """
    按核算方式汇总指定月份/列的数据
    
    Args:
        ws_detail: 当月项目明细数据工作表
        accounting_method: 核算方式（如"累计收入（NR）"）
        month_col_header: 月份列头（如"202603"）或列名
    
    Returns:
        汇总值
    """
    total = 0
    has_data = False
    
    # 找到月份列索引
    month_col_idx = None
    
    # 首先尝试在列 41-52 中查找月份列
    for col in range(41, 53):
        header = ws_detail.cell(row=1, column=col).value
        if header and str(header) == str(month_col_header):
            month_col_idx = col
            break
    
    # 如果没找到，尝试在整个第 1 行查找匹配的列名
    if not month_col_idx:
        for col in range(1, 60):
            header = ws_detail.cell(row=1, column=col).value
            if header and str(header) == str(month_col_header):
                month_col_idx = col
                break
    
    if not month_col_idx:
        return None
    
    # 遍历所有行，查找匹配的核算方式（使用完全匹配）
    for row in range(2, ws_detail.max_row + 1):
        # 检查第 18 列（核算方式列）
        method = ws_detail.cell(row=row, column=18).value
        
        if method and accounting_method == str(method):  # 完全匹配
            # 获取该行的月份数据
            val = ws_detail.cell(row=row, column=month_col_idx).value
            
            # 如果是公式，尝试获取计算后的值
            if val is None:
                continue
            
            try:
                # 尝试转换为数字
                if isinstance(val, (int, float)):
                    total += val
                    has_data = True
                elif isinstance(val, str) and val.strip():
                    # 如果是字符串，尝试转换
                    try:
                        total += float(val)
                        has_data = True
                    except ValueError:
                        pass
            except (TypeError, ValueError):
                pass
    
    return total if has_data else None

def find_target_row(ws_summary, indicator_name):
    """
    在汇总表中查找指标对应的行（使用完全匹配）
    
    Args:
        ws_summary: 当月汇总毛利率工作表
        indicator_name: 指标名称
    
    Returns:
        行号，如果未找到返回 None
    """
    for row in range(1, ws_summary.max_row + 1):
        val = ws_summary.cell(row=row, column=2).value  # B 列是指标列
        if val and indicator_name == str(val):  # 使用完全匹配
            return row
    return None

def find_target_column(ws_summary, year_month):
    """
    在汇总表中查找年月对应的列
    
    Args:
        ws_summary: 当月汇总毛利率工作表
        year_month: 月份列头（如"202603"）或日期字符串
    
    Returns:
        列号，如果未找到返回 None
    """
    # 将 202603 格式转换为 2026-03
    ym_str = str(year_month)
    if len(ym_str) == 6 and ym_str.isdigit():
        target_year = int(ym_str[:4])
        target_month = int(ym_str[4:6])
    else:
        try:
            parts = ym_str.split('-')
            target_year = int(parts[0])
            target_month = int(parts[1])
        except (ValueError, IndexError):
            return None
    
    for col in range(1, ws_summary.max_column + 1):
        header_val = ws_summary.cell(row=3, column=col).value  # 第 3 行是列名
        if header_val:
            header_str = str(header_val)
            
            # 尝试解析汇总表的日期列名（格式：2026-03-30）
            if '-' in header_str:
                try:
                    date_parts = header_str.split('-')
                    header_year = int(date_parts[0])
                    header_month = int(date_parts[1])
                    
                    if header_year == target_year and header_month == target_month:
                        return col
                except (ValueError, IndexError):
                    pass
    
    return None

def find_column_by_name(ws_summary, col_name):
    """
    在汇总表中按列名查找列
    
    Args:
        ws_summary: 当月汇总毛利率工作表
        col_name: 列名
    
    Returns:
        列号，如果未找到返回 None
    """
    for col in range(1, ws_summary.max_column + 1):
        header_val = ws_summary.cell(row=3, column=col).value
        if header_val and col_name == str(header_val):
            return col
    return None

def main():
    # 查找输出文件
    ref_dir = Path(__file__).parent.parent / 'references'
    output_file = find_file_fuzzy('汇总', ref_dir)
    
    if not output_file:
        print("错误：未找到输出文件")
        return
    
    print(f"处理文件：{output_file.name}")
    
    # 加载工作簿
    wb = openpyxl.load_workbook(output_file)
    
    # 检查工作表是否存在
    if '当月项目明细数据' not in wb.sheetnames:
        print("错误：未找到'当月项目明细数据'工作表")
        return
    
    if '当月汇总毛利率' not in wb.sheetnames:
        print("错误：未找到'当月汇总毛利率'工作表")
        return
    
    ws_detail = wb['当月项目明细数据']
    ws_summary = wb['当月汇总毛利率']
    
    # 直接从 Excel 读取要汇总的核算方式（第 4-7 行 B 列）
    accounting_methods = []
    for row in range(4, 8):
        val = ws_summary.cell(row=row, column=2).value
        if val:
            accounting_methods.append(val)
    
    print(f"\n要汇总的核算方式：{accounting_methods}")
    
    # 定义汇总表的列名与明细表列名的映射关系
    # 直接从汇总表读取列名，确保精确匹配
    summary_col_names = {}
    for col in range(1, 10):
        val = ws_summary.cell(row=3, column=col).value
        if val:
            summary_col_names[col] = str(val)
    
    print(f"\n汇总表列名：{summary_col_names}")
    
    # 固定映射：汇总表列号 -> 明细表列名
    fixed_mappings = {
        3: '项目立项毛利目标预算',  # 汇总表 C 列 -> 明细表列 19
        4: '项目开始 - 当前实际累计',  # 汇总表 D 列 -> 明细表列 20
    }
    
    # 用户指定的 YTD 列映射
    ytd_source_column = '26 年 YTD 实际累计 2 月'  # 用户指定的列
    print(f"\n2026 年 YTD 使用明细表列：{ytd_source_column}")
    
    # 统计结果
    stats = {
        'filled': 0,
        'skipped_existing': 0,
        'skipped_empty': 0,
        'skipped_no_column': 0
    }
    
    # 对每个核算方式进行汇总
    for method in accounting_methods:
        print(f"\n处理核算方式：{method}")
        
        # 查找汇总表中对应的行
        target_row = find_target_row(ws_summary, method)
        
        if not target_row:
            print(f"  警告：在汇总表中未找到指标'{method}'对应的行")
            continue
        
        print(f"  目标行：{target_row}")
        
        # 1. 处理固定列（项目立项毛利目标预算、项目开始 - 当前实际累计）
        for target_col, detail_col_name in fixed_mappings.items():
            col_name = summary_col_names.get(target_col, f'列{target_col}')
            
            # 检查目标单元格是否已有数据
            existing_val = ws_summary.cell(row=target_row, column=target_col).value
            
            if existing_val is not None and existing_val != '':
                print(f"  列'{col_name}': 单元格已有数据 ({existing_val})，跳过")
                stats['skipped_existing'] += 1
                continue
            
            # 从明细表汇总数据
            total = sum_by_accounting_method(ws_detail, method, detail_col_name)
            
            if total is None:
                print(f"  列'{col_name}': 无数据，跳过")
                stats['skipped_empty'] += 1
                continue
            
            # 填充数据
            ws_summary.cell(row=target_row, column=target_col).value = total
            print(f"  列'{col_name}': 填充 {total}")
            stats['filled'] += 1
        
        # 2. 处理 2026 年 YTD 列（列 5）
        target_col = 5
        col_name = summary_col_names.get(target_col, f'列{target_col}')
        
        # 检查目标单元格是否已有数据
        existing_val = ws_summary.cell(row=target_row, column=target_col).value
        
        if existing_val is not None and existing_val != '':
            print(f"  列'{col_name}': 单元格已有数据 ({existing_val})，跳过")
            stats['skipped_existing'] += 1
        else:
            # 从明细表汇总数据（使用用户指定的列）
            total = sum_by_accounting_method(ws_detail, method, ytd_source_column)
            
            if total is None:
                print(f"  列'{col_name}': 无数据，跳过")
                stats['skipped_empty'] += 1
            else:
                # 填充数据
                ws_summary.cell(row=target_row, column=target_col).value = total
                print(f"  列'{col_name}': 填充 {total}")
                stats['filled'] += 1
        
        # 3. 处理年月列（H 列及之后）
        # 获取汇总表中所有年月列
        month_columns = []
        for col in range(8, ws_summary.max_column + 1):  # 从 H 列开始
            header_val = ws_summary.cell(row=3, column=col).value
            if header_val:
                header_str = str(header_val)
                # 检查是否为日期格式
                if '-' in header_str:
                    try:
                        date_parts = header_str.split('-')
                        year = int(date_parts[0])
                        month = int(date_parts[1])
                        month_columns.append({
                            'col': col,
                            'year': year,
                            'month': month,
                            'header': header_str
                        })
                    except (ValueError, IndexError):
                        pass
        
        # 对每个月份列进行填充
        for month_info in month_columns:
            # 生成月份标识（如 202603）
            month_id = f"{month_info['year']}{month_info['month']:02d}"
            
            # 检查目标单元格是否已有数据
            existing_val = ws_summary.cell(row=target_row, column=month_info['col']).value
            
            if existing_val is not None and existing_val != '':
                print(f"  月份 {month_id}: 单元格已有数据 ({existing_val})，跳过")
                stats['skipped_existing'] += 1
                continue
            
            # 从明细表汇总数据
            total = sum_by_accounting_method(ws_detail, method, month_id)
            
            if total is None:
                print(f"  月份 {month_id}: 无数据，跳过")
                stats['skipped_empty'] += 1
                continue
            
            # 填充数据
            ws_summary.cell(row=target_row, column=month_info['col']).value = total
            print(f"  月份 {month_id}: 填充 {total}")
            stats['filled'] += 1
    
    # 保存文件
    wb.save(output_file)
    print(f"\n=== 统计结果 ===")
    print(f"成功填充：{stats['filled']} 个单元格")
    print(f"跳过（已有数据）: {stats['skipped_existing']} 个单元格")
    print(f"跳过（源数据为空）: {stats['skipped_empty']} 个单元格")
    print(f"跳过（未找到列）: {stats['skipped_no_column']} 个单元格")
    print(f"\n文件已保存：{output_file}")

if __name__ == '__main__':
    main()
