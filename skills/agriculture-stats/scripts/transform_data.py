"""
农业项目数据转换脚本

功能：
1. 读取输入表（项目毛利管控表）
2. 筛选农业项目
3. 按项目结束时间分类（当月在行 vs 当年在行）
4. 写入输出表的两个工作表（保持原格式、数值和公式）
"""

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pathlib import Path
import sys
from datetime import datetime

# ==================== 配置 ====================

# 筛选基准日期
CURRENT_MONTH = 3  # 当前月份（3 月）
LAST_MONTH_END = datetime(2026, 2, 28)  # 上月最后一天
YEAR_START = datetime(2026, 1, 1)  # 当年开始

# ==================== 工具函数 ====================

def load_agriculture_keywords():
    """从配置文件加载农业关键字
    
    配置文件路径：references/agriculture_keywords.md
    
    Returns:
        tuple: (include_keywords, exclude_keywords)
    """
    config_file = Path(__file__).parent.parent / 'references' / 'agriculture_keywords.md'
    
    if not config_file.exists():
        print(f"警告：未找到关键词配置文件 {config_file}")
        print("将使用空列表，请确保配置文件存在")
        return [], []
    
    with open(config_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    include_keywords = []
    exclude_keywords = []
    in_include = False
    in_exclude = False
    
    for line in content.split('\n'):
        line = line.strip()
        
        # 检测章节标题
        if line == '## 关键字列表':
            in_include = True
            in_exclude = False
            continue
        elif line == '## 排除关键字（可选）':
            in_include = False
            in_exclude = True
            continue
        elif line.startswith('##'):
            in_include = False
            in_exclude = False
            continue
        
        # 跳过代码块标记和空行
        if line.startswith('```') or not line:
            continue
        
        # 添加关键字
        if in_include:
            include_keywords.append(line)
        elif in_exclude:
            exclude_keywords.append(line)
    
    return include_keywords, exclude_keywords


def is_agriculture_project(project_name, include_kw, exclude_kw):
    """判断项目是否为农业项目
    
    规则：只要项目名称中包含任意一个农业关键词，就是农业项目
    """
    if not project_name:
        return False
    
    # 检查排除关键字
    for kw in exclude_kw:
        if kw in str(project_name):
            return False
    
    # 检查农业关键字（包含任意一个即可）
    for kw in include_kw:
        if kw in str(project_name):
            return True
    
    return False


def copy_cell_complete(source_ws, source_row, source_col, target_ws, target_row, target_col):
    """完整复制单元格：值、公式、格式
    
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
    
    # 1. 复制值（包括公式字符串）
    if source_cell.value is not None:
        target_cell.value = source_cell.value
    
    # 2. 复制样式（字体、填充、边框、对齐、数字格式）
    if source_cell.has_style:
        # 字体
        if source_cell.font:
            target_cell.font = source_cell.font.copy()
        # 填充
        if source_cell.fill:
            target_cell.fill = source_cell.fill.copy()
        # 边框
        if source_cell.border:
            target_cell.border = source_cell.border.copy()
        # 对齐
        if source_cell.alignment:
            target_cell.alignment = source_cell.alignment.copy()
        # 数字格式
        if source_cell.number_format:
            target_cell.number_format = source_cell.number_format
        # 保护
        if source_cell.protection:
            target_cell.protection = source_cell.protection.copy()
    
    # 3. 复制超链接（如果有）
    if source_cell.hyperlink:
        target_cell.hyperlink = source_cell.hyperlink.copy()
        if source_cell.hyperlink.tooltip:
            target_cell.hyperlink.tooltip = source_cell.hyperlink.tooltip


def copy_column_widths(source_ws, target_ws, max_col):
    """复制列宽
    
    Args:
        source_ws: 源工作表
        target_ws: 目标工作表
        max_col: 最大列数
    """
    for col in range(1, max_col + 1):
        col_letter = get_column_letter(col)
        source_dim = source_ws.column_dimensions[col_letter]
        target_dim = target_ws.column_dimensions[col_letter]
        
        # 复制列宽
        if source_dim.width:
            target_dim.width = source_dim.width
        
        # 复制其他列属性
        target_dim.hidden = source_dim.hidden
        target_dim.outlineLevel = source_dim.outlineLevel
        target_dim.collapsed = source_dim.collapsed


def copy_row_heights(source_ws, target_ws, start_source_row, end_source_row, start_target_row):
    """复制行高
    
    Args:
        source_ws: 源工作表
        target_ws: 目标工作表
        start_source_row: 源起始行
        end_source_row: 源结束行
        start_target_row: 目标起始行
    """
    row_offset = start_target_row - start_source_row
    for row in range(start_source_row, end_source_row + 1):
        source_dim = source_ws.row_dimensions[row]
        target_dim = target_ws.row_dimensions[row + row_offset]
        
        # 复制行高
        if source_dim.height:
            target_dim.height = source_dim.height
        
        # 复制其他行属性
        target_dim.hidden = source_dim.hidden
        target_dim.outlineLevel = source_dim.outlineLevel
        target_dim.collapsed = source_dim.collapsed


def copy_merged_cells(source_ws, target_ws, start_source_row, end_source_row, start_target_row):
    """复制合并单元格
    
    Args:
        source_ws: 源工作表
        target_ws: 目标工作表
        start_source_row: 源起始行
        end_source_row: 源结束行
        start_target_row: 目标起始行
    """
    row_offset = start_target_row - start_source_row
    
    for merged_range in source_ws.merged_cells.ranges:
        # 检查合并区域是否在要复制的范围内
        if (merged_range.min_row >= start_source_row and 
            merged_range.max_row <= end_source_row):
            
            # 计算目标区域的坐标
            new_min_row = merged_range.min_row + row_offset
            new_max_row = merged_range.max_row + row_offset
            
            # 创建新的合并区域
            new_range = f"{get_column_letter(merged_range.min_col)}{new_min_row}:{get_column_letter(merged_range.max_col)}{new_max_row}"
            target_ws.merge_cells(new_range)


# ==================== 主流程 ====================

def find_file_fuzzy(pattern, directory):
    """模糊查找文件（忽略全角/半角、空格）"""
    import os
    import glob
    
    # 尝试直接查找
    files = glob.glob(os.path.join(directory, pattern))
    if files:
        return Path(files[0])
    
    # 尝试不带空格的版本
    pattern_no_space = pattern.replace(' ', '').replace(' ', '')
    files = glob.glob(os.path.join(directory, '*' + pattern_no_space.split('*')[-1] if '*' in pattern_no_space else pattern_no_space))
    if files:
        return Path(files[0])
    
    # 列出目录下所有 xlsx 文件
    xlsx_files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
    
    # 简单匹配
    for f in xlsx_files:
        if pattern.split('.')[0] in f or pattern.split('.')[0] in f.replace(' ', ''):
            return Path(directory) / f
    
    return None


def main():
    print("="*80)
    print("农业项目数据转换")
    print("="*80)
    
    # 1. 查找文件
    ref_dir = Path.cwd() / '.lingma' / 'skills' / 'agriculture-stats' / 'references'
    
    # 使用模糊查找
    input_wb_path = find_file_fuzzy('*毛利管控表*.xlsx', str(ref_dir))
    output_wb_path = find_file_fuzzy('*农业项目群*.xlsx', str(ref_dir))
    
    if not input_wb_path or not isinstance(input_wb_path, Path) or not input_wb_path.exists():
        print(f"错误：未找到输入文件")
        return
    
    if not output_wb_path or not isinstance(output_wb_path, Path) or not output_wb_path.exists():
        print(f"错误：未找到输出文件")
        return
    
    print(f"\n输入文件：{input_wb_path.name}")
    print(f"输出文件：{output_wb_path.name}")
    
    # 2. 加载关键字
    include_kw, exclude_kw = load_agriculture_keywords()
    print(f"\n加载了 {len(include_kw)} 个农业关键字")
    print(f"加载了 {len(exclude_kw)} 个排除关键字")
    
    # 3. 读取输入表
    print(f"\n读取输入表...")
    input_wb = load_workbook(input_wb_path, read_only=False, data_only=False)
    
    # 查找实际运营数据 - 崔哲工作表
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
    print(f"行数：{source_ws.max_row}, 列数：{source_ws.max_column}")
    
    # 4. 解析项目数据
    print(f"\n解析项目数据...")
    projects = []
    current_project = None
    
    for row_idx in range(2, source_ws.max_row + 1):
        project_code = source_ws.cell(row=row_idx, column=1).value
        
        if project_code and str(project_code).strip() != '':
            # 新项目开始
            if current_project:
                projects.append(current_project)
            
            # 读取项目结束时间（第 4 列）
            end_date = source_ws.cell(row=row_idx, column=4).value
            
            current_project = {
                'start_row': row_idx,
                '项目编码': project_code,
                '项目名称': source_ws.cell(row=row_idx, column=2).value,
                '项目开始时间': source_ws.cell(row=row_idx, column=3).value,
                '项目结束时间': end_date,
                'end_row': row_idx + 8  # 每个项目 9 行
            }
    
    # 添加最后一个项目
    if current_project:
        projects.append(current_project)
    
    print(f"解析到 {len(projects)} 个项目")
    
    # 5. 筛选农业项目
    print(f"\n筛选农业项目...")
    agri_projects = [
        p for p in projects
        if is_agriculture_project(p['项目名称'], include_kw, exclude_kw)
    ]
    print(f"农业项目数：{len(agri_projects)}")
    
    for p in agri_projects:
        name = p['项目名称'][:50] if len(p['项目名称']) > 50 else p['项目名称']
        print(f"  - {name}... (结束：{p['项目结束时间']})")
    
    # 6. 按项目结束时间分类
    print(f"\n按项目结束时间分类...")
    
    current_month_projects = []  # 当月在行
    year_projects = []  # 当年在行
    
    for project in agri_projects:
        end_date = project['项目结束时间']
        
        # 处理日期格式
        if isinstance(end_date, str):
            try:
                end_date = datetime.strptime(end_date, '%Y-%m-%d')
            except:
                continue
        elif hasattr(end_date, 'date'):
            end_date = datetime.combine(end_date.date(), datetime.min.time())
        
        if end_date > LAST_MONTH_END:
            current_month_projects.append(project)
            name = project['项目名称'][:30] if len(project['项目名称']) > 30 else project['项目名称']
            print(f"  [当月在行] {name}...")
        
        if end_date > YEAR_START:
            year_projects.append(project)
    
    print(f"\n当月在行项目：{len(current_month_projects)} 个")
    print(f"当年在行项目：{len(year_projects)} 个")
    
    # 7. 加载输出表
    print(f"\n加载输出表...")
    output_wb = load_workbook(output_wb_path, read_only=False, data_only=False)
    
    # 删除旧的明细工作表（如果存在）
    for sheet_name in ['当月项目明细数据', '2026 年项目明细数据']:
        if sheet_name in output_wb.sheetnames:
            del output_wb[sheet_name]
    
    # 创建新的工作表
    ws_current = output_wb.create_sheet('当月项目明细数据')
    ws_year = output_wb.create_sheet('2026 年项目明细数据')
    
    # 8. 复制表头（第 1 行）
    print(f"\n复制表头...")
    max_col = source_ws.max_column
    
    # 复制表头行的所有单元格（值和格式）
    for col in range(1, max_col + 1):
        copy_cell_complete(source_ws, 1, col, ws_current, 1, col)
        copy_cell_complete(source_ws, 1, col, ws_year, 1, col)
    
    # 复制表头行的行高
    source_row_dim = source_ws.row_dimensions[1]
    if source_row_dim.height:
        ws_current.row_dimensions[1].height = source_row_dim.height
        ws_year.row_dimensions[1].height = source_row_dim.height
    
    # 复制列宽
    print(f"复制列宽...")
    copy_column_widths(source_ws, ws_current, max_col)
    copy_column_widths(source_ws, ws_year, max_col)
    
    # 9. 复制项目数据
    def copy_projects_to_sheet(projects_list, target_ws, sheet_name):
        """复制项目数据到工作表（包含完整格式）"""
        print(f"\n写入 {sheet_name} ({len(projects_list)} 个项目)...")
        
        target_row = 2  # 从第 2 行开始（表头在第 1 行）
        
        for idx, project in enumerate(projects_list):
            start_row = project['start_row']
            end_row = project['end_row']
            
            # 复制 9 行数据（包含完整格式）
            for row_offset in range(9):
                source_row = start_row + row_offset
                target_row_current = target_row + row_offset
                
                # 复制每个单元格
                for col in range(1, max_col + 1):
                    copy_cell_complete(
                        source_ws, source_row, col,
                        target_ws, target_row_current, col
                    )
            
            # 复制行高
            for row_offset in range(9):
                source_row = start_row + row_offset
                target_row_current = target_row + row_offset
                source_row_dim = source_ws.row_dimensions[source_row]
                if source_row_dim.height:
                    target_ws.row_dimensions[target_row_current].height = source_row_dim.height
            
            # 复制合并单元格
            copy_merged_cells(source_ws, target_ws, start_row, end_row, target_row)
            
            target_row += 9
        
        print(f"  写入完成，共 {target_row - 2} 行数据")
    
    # 写入当月项目
    copy_projects_to_sheet(current_month_projects, ws_current, '当月项目明细数据')
    
    # 写入当年项目
    copy_projects_to_sheet(year_projects, ws_year, '2026 年项目明细数据')
    
    # 10. 保存输出文件
    print(f"\n保存输出文件...")
    output_wb.save(output_wb_path)
    print(f"文件已保存：{output_wb_path}")
    
    print("\n" + "="*80)
    print("转换完成！")
    print("="*80)
    print(f"\n结果汇总:")
    print(f"  农业项目总数：{len(agri_projects)}")
    print(f"  当月在行项目：{len(current_month_projects)}")
    print(f"  当年在行项目：{len(year_projects)}")
    print(f"\n输出工作表:")
    print(f"  - 当月项目明细数据：{len(current_month_projects) * 9} 行数据 + 1 行表头")
    print(f"  - 2026 年项目明细数据：{len(year_projects) * 9} 行数据 + 1 行表头")


if __name__ == '__main__':
    main()
