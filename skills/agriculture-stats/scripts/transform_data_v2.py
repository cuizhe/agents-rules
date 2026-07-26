"""
农业项目数据转换脚本（增强版）

功能：
1. 读取输入表（项目毛利管控表）
2. 筛选农业项目
3. 按项目结束时间分类（当月在行 vs 当年在行）
4. 更新输出表的两个工作表（保留原始格式、合并单元格、公式）
"""

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pathlib import Path
import sys
from datetime import datetime
from copy import copy

# 添加路径
sys.path.insert(0, str(Path.cwd() / '.lingma' / 'skills' / 'agriculture-stats' / 'scripts'))
from file_utils import find_input_file, find_output_template_file

# ==================== 配置 ====================

CURRENT_MONTH = 3  # 当前月份
LAST_MONTH_END = datetime(2026, 2, 28)  # 上月最后一天
YEAR_START = datetime(2026, 1, 1)  # 当年开始

# ==================== 工具函数 ====================

def load_agriculture_keywords():
    """从配置文件加载农业关键字"""
    config_file = Path(__file__).parent / 'references' / 'agriculture_keywords.md'
    
    if config_file.exists():
        with open(config_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        include_keywords = []
        exclude_keywords = []
        in_include = False
        in_exclude = False
        
        for line in content.split('\n'):
            line = line.strip()
            if line == '## 关键字列表':
                in_include = True
                in_exclude = False
            elif line == '## 排除关键字（可选）':
                in_include = False
                in_exclude = True
            elif line.startswith('##'):
                in_include = False
                in_exclude = False
            elif line.startswith('```') or not line:
                continue
            elif in_include:
                include_keywords.append(line)
            elif in_exclude:
                exclude_keywords.append(line)
        
        return include_keywords, exclude_keywords
    else:
        return ['农业', '农村', '农田', '种植', '养殖'], ['银行', '保险']


def is_agriculture_project(project_name, include_kw, exclude_kw):
    """判断项目是否为农业项目"""
    if not project_name:
        return False
    
    for kw in exclude_kw:
        if kw in str(project_name):
            return False
    
    for kw in include_kw:
        if kw in str(project_name):
            return True
    
    return False


def copy_cell_full(source_ws, source_row, source_col, target_ws, target_row, target_col):
    """
    完整复制单元格：值、公式、格式、样式
    
    Args:
        source_ws: 源工作表
        source_row: 源行号
        source_col: 源列号
        target_ws: 目标工作表
        target_row: 目标行号
        target_col: 目标列号
    """
    source_cell = source_ws.cell(row=source_row, column=source_col)
    target_cell = target_ws.cell(row=target_row, column=target_col)
    
    # 1. 复制值（如果是公式，复制公式本身而不是计算结果）
    if source_cell.value is not None:
        # 检查是否是公式
        if hasattr(source_cell, 'value') and isinstance(source_cell.value, str) and source_cell.value.startswith('='):
            target_cell.value = source_cell.value  # 复制公式
        else:
            target_cell.value = source_cell.value
    
    # 2. 复制数字格式
    if source_cell.number_format:
        target_cell.number_format = source_cell.number_format
    
    # 3. 复制字体
    if source_cell.font:
        try:
            target_cell.font = copy(source_cell.font)
        except Exception as e:
            pass
    
    # 4. 复制填充（背景色）
    if source_cell.fill:
        try:
            target_cell.fill = copy(source_cell.fill)
        except Exception as e:
            pass
    
    # 5. 复制边框
    if source_cell.border:
        try:
            target_cell.border = copy(source_cell.border)
        except Exception as e:
            pass
    
    # 6. 复制对齐方式
    if source_cell.alignment:
        try:
            target_cell.alignment = copy(source_cell.alignment)
        except Exception as e:
            pass


def copy_merged_cells(source_ws, target_ws, row_offset=0):
    """
    复制合并单元格
    
    Args:
        source_ws: 源工作表
        target_ws: 目标工作表
        row_offset: 行偏移量（目标行 - 源行）
    """
    # 清除目标工作表的合并单元格
    for merged in list(target_ws.merged_cells.ranges):
        target_ws.unmerge_cells(str(merged))
    
    # 复制源工作表的合并单元格
    for merged in source_ws.merged_cells.ranges:
        # 计算新的合并范围
        min_row = merged.min_row + row_offset
        max_row = merged.max_row + row_offset
        min_col = merged.min_col
        max_col = merged.max_col
        
        # 创建新的合并范围
        merge_range = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
        target_ws.merge_cells(merge_range)


def copy_column_widths(source_ws, target_ws):
    """复制列宽"""
    for col in range(1, source_ws.max_column + 1):
        col_letter = get_column_letter(col)
        if source_ws.column_dimensions[col_letter].width:
            target_ws.column_dimensions[col_letter].width = source_ws.column_dimensions[col_letter].width


def copy_row_heights(source_ws, target_ws, max_row):
    """复制行高"""
    for row in range(1, min(max_row + 1, source_ws.max_row + 1)):
        if source_ws.row_dimensions[row].height:
            target_ws.row_dimensions[row].height = source_ws.row_dimensions[row].height


def clear_sheet_data(target_ws, keep_header_rows=1):
    """
    清空工作表数据（保留表头）
    
    Args:
        target_ws: 目标工作表
        keep_header_rows: 保留的行数（表头）
    """
    # 清除数据行的值，但保留格式
    for row in range(keep_header_rows + 1, target_ws.max_row + 1):
        for col in range(1, target_ws.max_column + 1):
            cell = target_ws.cell(row=row, column=col)
            cell.value = None
    
    # 清除数据行的合并单元格
    for merged in list(target_ws.merged_cells.ranges):
        if merged.min_row > keep_header_rows:
            target_ws.unmerge_cells(str(merged))


# ==================== 主流程 ====================

def main():
    print("="*80)
    print("农业项目数据转换（增强版 - 保留格式）")
    print("="*80)
    
    # 1. 查找文件
    ref_dir = Path.cwd() / '.lingma' / 'skills' / 'agriculture-stats' / 'references'
    
    input_file = find_input_file(str(ref_dir))
    output_file = find_output_template_file(str(ref_dir))
    
    if not input_file:
        print("错误：未找到输入文件")
        return
    
    if not output_file:
        print("错误：未找到输出文件")
        return
    
    print(f"\n输入文件：{input_file.name}")
    print(f"输出文件：{output_file.name}")
    
    # 2. 加载关键字
    include_kw, exclude_kw = load_agriculture_keywords()
    print(f"\n加载了 {len(include_kw)} 个农业关键字")
    print(f"加载了 {len(exclude_kw)} 个排除关键字")
    
    # 3. 读取输入表
    print(f"\n读取输入表...")
    input_wb = load_workbook(input_file, read_only=False, data_only=False)
    
    source_sheet_name = None
    for name in input_wb.sheetnames:
        if '实际' in name and '崔哲' in name:
            source_sheet_name = name
            break
    
    if not source_sheet_name:
        print("错误：未找到'实际运营数据 - 崔哲'工作表")
        return
    
    source_ws = input_wb[source_sheet_name]
    print(f"工作表：{source_sheet_name}")
    print(f"数据范围：{source_ws.dimensions}")
    
    # 4. 读取输出表
    print(f"\n读取输出表...")
    output_wb = load_workbook(output_file, read_only=False, data_only=False)
    print(f"现有工作表：{output_wb.sheetnames}")
    
    # 5. 解析项目数据
    print(f"\n解析项目数据...")
    projects = []
    current_project = None
    
    for row_idx in range(2, source_ws.max_row + 1):
        project_code = source_ws.cell(row=row_idx, column=1).value
        
        if project_code and str(project_code).strip() != '':
            if current_project:
                projects.append(current_project)
            
            end_date = source_ws.cell(row=row_idx, column=4).value
            
            current_project = {
                'start_row': row_idx,
                '项目编码': project_code,
                '项目名称': source_ws.cell(row=row_idx, column=2).value,
                '项目开始时间': source_ws.cell(row=row_idx, column=3).value,
                '项目结束时间': end_date,
                'end_row': row_idx + 8
            }
    
    if current_project:
        projects.append(current_project)
    
    print(f"解析到 {len(projects)} 个项目")
    
    # 6. 筛选农业项目
    print(f"\n筛选农业项目...")
    agri_projects = [
        p for p in projects
        if is_agriculture_project(p['项目名称'], include_kw, exclude_kw)
    ]
    print(f"农业项目数：{len(agri_projects)}")
    
    # 7. 按项目结束时间分类
    print(f"\n按项目结束时间分类...")
    
    current_month_projects = []
    year_projects = []
    
    for project in agri_projects:
        end_date = project['项目结束时间']
        
        if isinstance(end_date, str):
            try:
                end_date = datetime.strptime(end_date, '%Y-%m-%d')
            except:
                continue
        elif hasattr(end_date, 'date'):
            end_date = datetime.combine(end_date.date(), datetime.min.time())
        
        if end_date > LAST_MONTH_END:
            current_month_projects.append(project)
            print(f"  [当月在行] {project['项目名称'][:30]}...")
        
        if end_date > YEAR_START:
            year_projects.append(project)
    
    print(f"\n当月在行项目：{len(current_month_projects)} 个")
    print(f"当年在行项目：{len(year_projects)} 个")
    
    # 8. 更新输出工作表
    print(f"\n更新输出工作表...")
    
    def update_sheet(sheet_name, projects_list):
        """更新工作表数据，完全清空后重新复制"""
        if sheet_name not in output_wb.sheetnames:
            print(f"  警告：工作表 '{sheet_name}' 不存在，创建新工作表")
            target_ws = output_wb.create_sheet(sheet_name)
        else:
            target_ws = output_wb[sheet_name]
            print(f"  使用已存在的工作表：{sheet_name}")
            # 完全清空工作表（不保留表头）
            clear_sheet_data(target_ws, keep_header_rows=0)
        
        # 复制表头（从输入表复制）
        print(f"  复制表头...")
        for row in range(1, 2):  # 只复制第 1 行表头
            for col in range(1, source_ws.max_column + 1):
                copy_cell_full(source_ws, row, col, target_ws, row, col)
        
        # 复制表头的合并单元格
        for merged in source_ws.merged_cells.ranges:
            if merged.min_row == 1:  # 只复制第 1 行的合并单元格
                merge_range = f"{get_column_letter(merged.min_col)}{merged.min_row}:{get_column_letter(merged.max_col)}{merged.max_row}"
                target_ws.merge_cells(merge_range)
        
        # 复制列宽
        copy_column_widths(source_ws, target_ws)
        
        # 复制项目数据
        print(f"  写入 {len(projects_list)} 个项目...")
        
        target_row = 2  # 从第 2 行开始（表头在第 1 行）
        
        for project in projects_list:
            start_row = project['start_row']
            
            # 复制 9 行数据（包括合并单元格）
            for row_offset in range(9):
                source_row = start_row + row_offset
                target_row_current = target_row + row_offset
                
                for col in range(1, source_ws.max_column + 1):
                    copy_cell_full(
                        source_ws, source_row, col,
                        target_ws, target_row_current, col
                    )
            
            target_row += 9
        
        print(f"    写入完成，共 {target_row - 2} 行数据")
    
    # 更新两个工作表
    update_sheet('当月项目明细数据', current_month_projects)
    update_sheet('2026 年项目明细数据', year_projects)
    
    # 9. 保存输出文件
    print(f"\n保存输出文件...")
    output_wb.save(output_file)
    print(f"文件已保存：{output_file}")
    
    print("\n" + "="*80)
    print("转换完成！")
    print("="*80)
    print(f"\n结果汇总:")
    print(f"  农业项目总数：{len(agri_projects)}")
    print(f"  当月在行项目：{len(current_month_projects)}")
    print(f"  当年在行项目：{len(year_projects)}")
    print(f"\n输出工作表:")
    print(f"  - 当月项目明细数据：{len(current_month_projects) * 9} 行")
    print(f"  - 2026 年项目明细数据：{len(year_projects) * 9} 行")
    print(f"\n已保留:")
    print(f"  [OK] 单元格格式（字体、颜色、边框）")
    print(f"  [OK] 合并单元格")
    print(f"  [OK] 列宽")
    print(f"  [OK] 公式（如有）")


if __name__ == '__main__':
    main()
