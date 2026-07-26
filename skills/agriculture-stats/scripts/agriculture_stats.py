"""
农业项目统计核心模块

功能：
1. 读取多个"实际运营数据"工作表（李陶、崔哲、赵云平、刘科源）
2. 按关键字筛选农业项目
3. 按项目结束时间分类（当月在行 / 当年在行 / 已结项）
4. 生成4个输出工作表：
   - 当月项目明细数据
   - 2026年项目明细数据
   - 当月汇总毛利率（按核算科目汇总）
   - 2026年汇总毛利率（按核算科目汇总）
"""

import sys
import re
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Dict, Tuple
from copy import copy

try:
    import openpyxl
    from openpyxl import load_workbook
    from openpyxl.cell.cell import MergedCell
    from openpyxl.utils import get_column_letter
except ImportError:
    print("错误：需要安装 openpyxl")
    print("pip install openpyxl")
    sys.exit(1)

# ============ 配置 =============
# 农业项目关键字（包含即匹配）- 与 agriculture_keywords.md 保持同步
AGRI_INCLUDE_KEYWORDS = [
    '农业科学院', '农科院', '农业农村部', '农村', '农田', '种植', '养殖',
    '畜牧', '渔业', '农机', '种业', '农资', '农产品', '畜禽', '林业',
    '农场', '全国畜牧总站', '审计署', '兽医', '中监所'
]

# 排除关键字（优先级高于包含关键字）
AGRI_EXCLUDE_KEYWORDS = ['农业项目群', '银行', '保险', '证券', '贷款', '融资']

# 默认数据工作表名称（可配置，变更时修改此常量即可）
DEFAULT_DATA_SHEET = '实际运营数据-崔哲'

# 核算科目名称标准化映射
CATEGORY_MAPPING = {
    '累计收入（NR）': '累计收入（NR）',
    '人工成本（成本１）': '人工成本（成本１）',
    '闲置成本（成本2）': '闲置成本（成本2）',
    '差旅费（成本３）': '差旅费（成本３）',
    '其他成本（成本３）': '其他成本（成本３）',
    '累计成本': '累计成本',
    'GP1毛利': 'GP1毛利',
    'GP1毛利率(%)': 'GP1毛利率(%)',
    'GP3毛利': 'GP3毛利',
    'GP3毛利率(%)': 'GP3毛利率(%)',
}

# 核算科目展示顺序（10行/项目）
CATEGORY_ORDER = [
    '累计收入（NR）',
    '人工成本（成本１）',
    '闲置成本（成本2）',
    '差旅费（成本３）',
    '其他成本（成本３）',
    '累计成本',
    'GP1毛利',
    'GP1毛利率(%)',
    'GP3毛利',
    'GP3毛利率(%)',
]

# 每个项目块行数（与 CATEGORY_ORDER 长度一致）
PROJECT_ROW_COUNT = len(CATEGORY_ORDER)  # 10

# 列位置映射（1-based）
COL_PROJECT_CODE = 1       # 项目编码
COL_PROJECT_NAME = 2      # 项目名称
COL_START_DATE = 3        # 项目开始时间
COL_END_DATE = 4          # 项目结束时间
COL_PEOPLE_MONTH = 5      # 1月项目人数
COL_SETTLE_MODE = 17      # 结算方式
COL_CATEGORY = 18         # 类别（核算科目）
COL_BUDGET = 19           # 项目整体毛利目标预算
COL_CURRENT_ACTUAL = 20   # 项目开始-当前实际累计
COL_25_ACTUAL = 21        # 项目开始-25年实际累计
COL_26_YTD_BASE = 22      # 26年YTD实际累计12月（列22）
# YTD各月：22=12月, 23=11月, ..., 33=1月
# 月度数据：41=202612, 42=202611, ..., 52=202601
COL_MONTHLY_BASE = 41     # 2026年12月数据（列41）
COL_YEAR_2025 = 53        # 2025年
COL_YEAR_2024 = 54        # 2024年
COL_YEAR_2023 = 55        # 2023年
COL_REMARK = 56           # 备注
# ================================


def normalize_text(text: str) -> str:
    """标准化文本：全角转半角，移除空格，转小写"""
    if not isinstance(text, str):
        return ''
    fullwidth_to_halfwidth = {chr(0xFF01 + i): chr(0x21 + i) for i in range(94)}
    fullwidth_to_halfwidth.update({chr(0xFF10 + i): chr(0x30 + i) for i in range(10)})
    fullwidth_to_halfwidth.update({chr(0xFF21 + i): chr(0x41 + i) for i in range(26)})
    fullwidth_to_halfwidth.update({chr(0xFF41 + i): chr(0x61 + i) for i in range(26)})
    fullwidth_to_halfwidth['　'] = ' '
    result = ''.join(fullwidth_to_halfwidth.get(c, c) for c in text)
    result = re.sub(r'\s+', '', result)
    return result.lower()


def is_agriculture_project(name: str) -> bool:
    """判断是否为农业项目"""
    if not name:
        return False
    name_str = str(name)
    # 先排除
    for kw in AGRI_EXCLUDE_KEYWORDS:
        if kw in name_str:
            return False
    # 再匹配包含
    for kw in AGRI_INCLUDE_KEYWORDS:
        if kw in name_str:
            return True
    return False


def copy_cell_style(source_cell, target_cell):
    """复制单元格格式（字体、填充、边框、对齐、数字格式）"""
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.fill = copy(source_cell.fill)
        target_cell.border = copy(source_cell.border)
        target_cell.alignment = copy(source_cell.alignment)
        target_cell.number_format = source_cell.number_format
        target_cell.protection = copy(source_cell.protection)


def find_project_merged_cells(worksheet, project_start_row: int, project_row_count: int = PROJECT_ROW_COUNT):
    """
    查找指定项目块的合并单元格信息

    Args:
        worksheet: 源工作表
        project_start_row: 项目块起始行（1-based）
        project_row_count: 项目块行数（默认9）

    Returns:
        List of (col_start, col_end) tuples，表示在项目块列范围内纵向合并的列
    """
    project_end_row = project_start_row + project_row_count - 1
    merged_cols = []
    for merge_range in worksheet.merged_cells.ranges:
        # 检查是否是纵向合并（多行，单列，或多列但行跨度与项目块一致）
        min_row = merge_range.min_row
        max_row = merge_range.max_row
        min_col = merge_range.min_col
        max_col = merge_range.max_col

        # 只处理在项目块行范围内完全包含的合并
        if min_row >= project_start_row and max_row <= project_end_row:
            # 记录（只处理纵向合并，即列跨度为1的情况）
            if min_col == max_col:
                merged_cols.append((min_col, max_col))
    return merged_cols


def rewrite_row_references_in_formula(formula: str, source_start_row: int, target_start_row: int) -> str:
    """
    重写公式中的行号引用（用于将源表公式适配到输出表行号）

    场景：从源表读取一个公式并写入输出表时，公式中的行号需要
    根据源起始行和目标起始行的偏移进行调整。

    块内引用逻辑：
        源表项目块从 source_start_row 开始（通常 PROJECT_ROW_COUNT 行：source_start 到 source_start+PROJECT_ROW_COUNT-1）
        输出表项目块从 target_start_row 开始（同样 PROJECT_ROW_COUNT 行）

        公式中 >= source_start_row 的行号都是块内引用（或块延伸），
        需要根据偏移量调整到目标行号。

    例如：
        源表第2-10行（项目块）被写入输出表第8-16行
        source_start=2, target_start=8, 偏移量=6
        源公式 '=S7/S2' 引用块内行7和行2
        转换: 行7→13, 行2→8 (都加6)
        结果: '=S13/S8'

    Args:
        formula: 原始公式字符串，如 '=S7/S2' 或 '=SUM(AO2:BC2)'
        source_start_row: 源数据块起始行号
        target_start_row: 目标位置起始行号

    Returns:
        行号已更新的新公式字符串
    """
    import re

    row_offset = target_start_row - source_start_row

    def replace_row_num(match):
        col_letters = match.group(1)  # 如 'S', 'AO', 'BC'
        row_num = int(match.group(2))  # 行号

        # 只更新 >= source_start_row 的行号
        # 这些通常是块内引用，需要映射到目标行号
        if row_num >= source_start_row:
            new_row = row_num + row_offset
            return f'{col_letters}{new_row}'
        else:
            # 行号在块之前（表头等固定引用），保持不变
            return match.group(0)

    # 匹配列字母+数字的组合（如 S7, AO12, BC2）
    pattern = r'([A-Za-z]+)(\d+)'
    result = re.sub(pattern, replace_row_num, formula)
    return result


def parse_date(val) -> Optional[datetime]:
    """解析日期值"""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val
    if isinstance(val, (int, float)):
        # Excel 日期序列号
        try:
            return datetime(1899, 12, 30) + __import__('datetime').timedelta(days=int(val))
        except:
            return None
    if isinstance(val, str):
        for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y年%m月%d日'):
            try:
                return datetime.strptime(val.strip(), fmt)
            except:
                continue
    return None


def load_keywords_from_file(keywords_path: Path) -> Tuple[List[str], List[str]]:
    """从 agriculture_keywords.md 加载关键字"""
    if not keywords_path.exists():
        return AGRI_INCLUDE_KEYWORDS, AGRI_EXCLUDE_KEYWORDS

    with open(keywords_path, 'r', encoding='utf-8') as f:
        content = f.read()

    include_kw, exclude_kw = [], []
    in_include, in_exclude = False, False

    for line in content.split('\n'):
        line = line.strip()
        if line == '## 关键字列表':
            in_include, in_exclude = True, False
        elif line == '## 排除关键字（可选）':
            in_include, in_exclude = False, True
        elif line.startswith('##'):
            in_include = in_exclude = False
        elif line.startswith('```') or not line:
            continue
        elif in_include:
            include_kw.append(line)
        elif in_exclude:
            exclude_kw.append(line)

    return include_kw, exclude_kw


def find_data_sheets(workbook, sheet_name: str = None) -> List[str]:
    """
    查找数据工作表

    Args:
        workbook: Excel 工作簿
        sheet_name: 指定工作表名称，默认为 DEFAULT_DATA_SHEET（实际运营数据-崔哲）
                   支持模糊匹配（包含即匹配）

    Returns:
        匹配的工作表名称列表
    """
    target = sheet_name or DEFAULT_DATA_SHEET
    target_norm = normalize_text(target)
    result = []
    for name in workbook.sheetnames:
        if normalize_text(name) == target_norm:
            result.append(name)
            break
        elif target in name:
            result.append(name)
            break
    return result


class ProjectData:
    """单个项目数据"""
    def __init__(self, code: str, name: str, start_row: int, source_worksheet=None):
        self.code = code
        self.name = name
        self.start_row = start_row
        self.end_row = start_row + PROJECT_ROW_COUNT - 1  # 每个项目 PROJECT_ROW_COUNT 行
        self.source_worksheet = source_worksheet  # 保留公式时需要引用源工作表
        self.start_date: Optional[datetime] = None
        self.end_date: Optional[datetime] = None
        self.settle_mode = ''
        self.metrics: Dict[str, Dict] = {}  # category -> {budget, actual, ytd, monthly...}

    def load_from_sheet(self, ws):
        """从工作表（公式模式）加载项目数据，用于保留公式"""
        r = self.start_row
        self.source_worksheet = ws
        # 基本信息（只在第1行有值，后续行要跳过）
        self.start_date = parse_date(ws.cell(row=r, column=COL_START_DATE).value)
        self.end_date = parse_date(ws.cell(row=r, column=COL_END_DATE).value)
        self.settle_mode = ws.cell(row=r, column=COL_SETTLE_MODE).value or ''

        # PROJECT_ROW_COUNT 行数据，每行一个核算科目
        for offset in range(PROJECT_ROW_COUNT):
            row = r + offset
            category = ws.cell(row=row, column=COL_CATEGORY).value
            if not category:
                continue
            category = str(category).strip()

            self.metrics[category] = {
                'budget': self._safe_num(ws.cell(row=row, column=COL_BUDGET).value),
                'current_actual': 0.0,  # 由 load_values_from_sheet 填充
                '2025_actual': 0.0,     # 由 load_values_from_sheet 填充
                # YTD各月（22=12月, 23=11月, ..., 33=1月）
                'ytd_12': 0.0, 'ytd_11': 0.0, 'ytd_10': 0.0, 'ytd_09': 0.0,
                'ytd_08': 0.0, 'ytd_07': 0.0, 'ytd_06': 0.0, 'ytd_05': 0.0,
                'ytd_04': 0.0, 'ytd_03': 0.0, 'ytd_02': 0.0, 'ytd_01': 0.0,
                # 月度数据（41=202612, 42=202611, ..., 52=202601）
                'monthly_12': 0.0, 'monthly_11': 0.0, 'monthly_10': 0.0,
                'monthly_09': 0.0, 'monthly_08': 0.0, 'monthly_07': 0.0,
                'monthly_06': 0.0, 'monthly_05': 0.0, 'monthly_04': 0.0,
                'monthly_03': 0.0, 'monthly_02': 0.0, 'monthly_01': 0.0,
                # 年度数据
                'year_2025': 0.0, 'year_2024': 0.0, 'year_2023': 0.0,
                '偏差': 0.0, 'ytd_偏差': 0.0,
                '偏差原因': str(ws.cell(row=row, column=36).value or ''),
                'ytd_偏差原因': str(ws.cell(row=row, column=37).value or ''),
                '改进措施': str(ws.cell(row=row, column=38).value or ''),
                '进展': str(ws.cell(row=row, column=39).value or ''),
                'remark': ws.cell(row=row, column=COL_REMARK).value or '',
            }

    def load_values_from_sheet(self, ws):
        """
        用 data_only=True 的工作表更新 metrics 中的计算值（用于汇总聚合）。
        明细表写入仍使用源表公式，但汇总聚合需要真实数值。
        """
        r = self.start_row
        for offset in range(PROJECT_ROW_COUNT):
            row = r + offset
            category = ws.cell(row=row, column=COL_CATEGORY).value
            if not category:
                continue
            category = str(category).strip()
            if category not in self.metrics:
                continue

            self.metrics[category].update({
                'current_actual': self._safe_num(ws.cell(row=row, column=COL_CURRENT_ACTUAL).value),
                '2025_actual': self._safe_num(ws.cell(row=row, column=COL_25_ACTUAL).value),
                'ytd_12': self._safe_num(ws.cell(row=row, column=22).value),
                'ytd_11': self._safe_num(ws.cell(row=row, column=23).value),
                'ytd_10': self._safe_num(ws.cell(row=row, column=24).value),
                'ytd_09': self._safe_num(ws.cell(row=row, column=25).value),
                'ytd_08': self._safe_num(ws.cell(row=row, column=26).value),
                'ytd_07': self._safe_num(ws.cell(row=row, column=27).value),
                'ytd_06': self._safe_num(ws.cell(row=row, column=28).value),
                'ytd_05': self._safe_num(ws.cell(row=row, column=29).value),
                'ytd_04': self._safe_num(ws.cell(row=row, column=30).value),
                'ytd_03': self._safe_num(ws.cell(row=row, column=31).value),
                'ytd_02': self._safe_num(ws.cell(row=row, column=32).value),
                'ytd_01': self._safe_num(ws.cell(row=row, column=33).value),
                'monthly_12': self._safe_num(ws.cell(row=row, column=41).value),
                'monthly_11': self._safe_num(ws.cell(row=row, column=42).value),
                'monthly_10': self._safe_num(ws.cell(row=row, column=43).value),
                'monthly_09': self._safe_num(ws.cell(row=row, column=44).value),
                'monthly_08': self._safe_num(ws.cell(row=row, column=45).value),
                'monthly_07': self._safe_num(ws.cell(row=row, column=46).value),
                'monthly_06': self._safe_num(ws.cell(row=row, column=47).value),
                'monthly_05': self._safe_num(ws.cell(row=row, column=48).value),
                'monthly_04': self._safe_num(ws.cell(row=row, column=49).value),
                'monthly_03': self._safe_num(ws.cell(row=row, column=50).value),
                'monthly_02': self._safe_num(ws.cell(row=row, column=51).value),
                'monthly_01': self._safe_num(ws.cell(row=row, column=52).value),
                'year_2025': self._safe_num(ws.cell(row=row, column=COL_YEAR_2025).value),
                'year_2024': self._safe_num(ws.cell(row=row, column=COL_YEAR_2024).value),
                'year_2023': self._safe_num(ws.cell(row=row, column=COL_YEAR_2023).value),
                '偏差': self._safe_num(ws.cell(row=row, column=34).value),
                'ytd_偏差': self._safe_num(ws.cell(row=row, column=35).value),
            })

    def get_source_formula(self, row_offset: int, col_idx: int) -> Optional[str]:
        """
        获取源单元格公式

        Args:
            row_offset: 相对项目起始行的行偏移（0-8）
            col_idx: 输出表列索引（1-based）

        Returns:
            公式字符串（如 '=S7/S2'），无公式则返回 None
        """
        if self.source_worksheet is None:
            return None
        source_row = self.start_row + row_offset
        src_col = self._output_col_to_source_col(col_idx)
        if src_col is None:
            return None
        cell = self.source_worksheet.cell(row=source_row, column=src_col)
        if cell.data_type == 'f' and cell.value and str(cell.value).startswith('='):
            return str(cell.value)
        return None

    @staticmethod
    def _output_col_to_source_col(out_col: int) -> Optional[int]:
        """
        将输出表列索引映射为源表列索引

        输出明细表列顺序（1-based）：
        1=项目编码, 2=项目名称, 3=开始时间, 4=结束时间, 5=结算方式,
        6=类别, 7=预算, 8=实际, 9=25年实际,
        10-21=YTD12月~YTD1月, 22=偏差, 23=2025年, 24=2024年, 25=2023年, 26=备注

        对应源表列号：
        """
        col_map = {
            1: COL_PROJECT_CODE,   # 项目编码
            2: COL_PROJECT_NAME,   # 项目名称
            3: COL_START_DATE,     # 项目开始时间
            4: COL_END_DATE,       # 项目结束时间
            5: 5, 6: 6, 7: 7, 8: 8, 9: 9, 10: 10,  # 1-6月项目人数
            11: 11, 12: 12, 13: 13, 14: 14, 15: 15, 16: 16,  # 7-12月项目人数
            17: COL_SETTLE_MODE,   # 结算方式
            18: COL_CATEGORY,      # 类别
            19: COL_BUDGET,        # 项目整体毛利目标预算
            20: COL_CURRENT_ACTUAL, # 项目开始-当前实际累计
            21: COL_25_ACTUAL,     # 项目开始-25年实际累计
            22: 22, 23: 23, 24: 24, 25: 25, 26: 26, 27: 27,  # YTD 12月~7月
            28: 28, 29: 29, 30: 30, 31: 31, 32: 32, 33: 33,  # YTD 6月~1月
            34: 34,  # 项目周期偏差
            35: 35,  # YTD当年偏差
            36: 36, 37: 37, 38: 38, 39: 39, 40: 40,  # 偏差原因/改进措施/进展/真实经营数据
            41: 41, 42: 42, 43: 43, 44: 44, 45: 45, 46: 46,  # 月度 202612~202607
            47: 47, 48: 48, 49: 49, 50: 50, 51: 51, 52: 52,  # 月度 202606~202601
            53: COL_YEAR_2025,  # 2025年
            54: COL_YEAR_2024,   # 2024年
            55: COL_YEAR_2023,   # 2023年
            56: COL_REMARK,      # 备注
        }
        return col_map.get(out_col)

    @staticmethod
    def _safe_num(val):
        """安全转换为数值"""
        if val is None:
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)
        if isinstance(val, str):
            val = val.strip()
            if val in ('', '#DIV/0!', '#N/A', '#REF!'):
                return 0.0
            try:
                return float(val.replace(',', ''))
            except:
                return 0.0
        return 0.0


class AgricultureStats:
    """农业项目统计主类"""

    def __init__(self,
                 input_file: Path,
                 output_file: Path,
                 current_month: int = None,
                 year: int = None,
                 keywords_path: Path = None,
                 sheet_name: str = None):
        """
        Args:
            input_file: 输入Excel文件路径
            output_file: 输出Excel文件路径
            current_month: 当前月份（用于判断当月在行），默认自动检测
            year: 统计年份，默认当前年份
            keywords_path: 关键字配置文件路径
            sheet_name: 数据工作表名称，默认使用 DEFAULT_DATA_SHEET（实际运营数据-崔哲）
        """
        self.input_file = Path(input_file)
        self.output_file = Path(output_file)
        self.year = year or datetime.now().year
        self.current_month = current_month or datetime.now().month
        self.sheet_name = sheet_name or DEFAULT_DATA_SHEET
        self._wb_source = None  # 保留源工作簿引用直至写入完成
        self._wb_source_values = None  # data_only=True 的值版本
        # 上月最后一天
        if self.current_month == 1:
            self.last_month_end = datetime(self.year - 1, 12, 31)
        else:
            self.last_month_end = datetime(self.year, self.current_month - 1, 28)  # 简化
        self.year_start = datetime(self.year, 1, 1)

        # 加载关键字
        if keywords_path:
            self.include_kw, self.exclude_kw = load_keywords_from_file(keywords_path)
        else:
            self.include_kw, self.exclude_kw = AGRI_INCLUDE_KEYWORDS, AGRI_EXCLUDE_KEYWORDS

        self.projects: List[ProjectData] = []
        self.agri_projects: List[ProjectData] = []

    def load_data(self) -> int:
        """加载所有数据表，返回农业项目数量"""
        print(f"读取输入文件：{self.input_file.name}")
        # 使用 data_only=False 以保留公式（用于后续重写行号）
        # 注意：不能用 read_only=True，因为需要访问 merged_cells.ranges
        wb = load_workbook(self.input_file, data_only=False)
        # 额外加载 data_only=True 以获取公式的计算值（用于汇总聚合）
        wb_values = load_workbook(self.input_file, data_only=True)

        # 查找数据工作表
        sheet_names = find_data_sheets(wb, self.sheet_name)
        if not sheet_names:
            print(f"错误：未找到工作表 '{self.sheet_name}'")
            print(f"可用工作表：{wb.sheetnames}")
            wb.close()
            wb_values.close()
            return 0
        print(f"使用工作表：{sheet_names[0]}")

        all_projects = []

        for sname in sheet_names:
            ws = wb[sname]
            ws_vals = wb_values[sname]
            print(f"  处理工作表：{sname}")

            # 遍历行，识别项目（每 PROJECT_ROW_COUNT 行为一个项目）
            row = 2  # 跳过表头
            while row <= ws.max_row:
                code = ws.cell(row=row, column=COL_PROJECT_CODE).value
                name = ws.cell(row=row, column=COL_PROJECT_NAME).value

                if code and name and str(code).strip() not in ('', 'None'):
                    code_str = str(code).strip()
                    name_str = str(name).strip()

                    project = ProjectData(code_str, name_str, row)
                    project.load_from_sheet(ws)
                    # 用 data_only=True 的工作表更新 metrics 中的计算值（用于汇总聚合）
                    project.load_values_from_sheet(ws_vals)
                    all_projects.append(project)
                    row += PROJECT_ROW_COUNT  # 跳到下个项目的开始
                else:
                    row += 1

        self._wb_source = wb  # 保留引用，写入完成后再关闭
        self._wb_source_values = wb_values  # 保留值版本引用

        self.projects = all_projects

        # 筛选农业项目
        self.agri_projects = [
            p for p in self.projects
            if is_agriculture_project(p.name)
        ]

        print(f"\n共解析 {len(self.projects)} 个项目")
        print(f"农业项目：{len(self.agri_projects)} 个")
        for p in self.agri_projects:
            end_str = p.end_date.strftime('%Y-%m-%d') if p.end_date else '未知'
            print(f"  - [{p.code}] {p.name[:50]}... 结束:{end_str}")

        return len(self.agri_projects)

    def classify_projects(self) -> Tuple[List[ProjectData], List[ProjectData], List[ProjectData]]:
        """按项目结束时间分类项目"""
        current_month_list = []  # 当月在行（结束时间 > 上月末）
        year_list = []            # 当年在行（结束时间 > 当年1月1日）
        closed_list = []          # 已结项（结束时间 <= 当年1月1日）

        for p in self.agri_projects:
            end = p.end_date
            if end is None:
                continue

            if end > self.year_start:
                year_list.append(p)
                if end > self.last_month_end:
                    current_month_list.append(p)
            else:
                closed_list.append(p)

        print(f"\n项目分类（截止 {self.last_month_end.strftime('%Y-%m-%d')}）：")
        print(f"  当月在行：{len(current_month_list)} 个")
        print(f"  当年在行：{len(year_list)} 个")
        print(f"  已结项：{len(closed_list)} 个")

        return current_month_list, year_list, closed_list

    def aggregate_by_category(self, projects: List[ProjectData]) -> Dict[str, Dict]:
        """按核算科目汇总多个项目的指标"""
        agg: Dict[str, Dict] = {}

        for cat in CATEGORY_ORDER:
            agg[cat] = {
                'budget': 0.0,
                'current_actual': 0.0,
                '2025_actual': 0.0,
                'ytd_01': 0.0, 'ytd_02': 0.0, 'ytd_03': 0.0,
                'ytd_04': 0.0, 'ytd_05': 0.0, 'ytd_06': 0.0,
                'ytd_07': 0.0, 'ytd_08': 0.0, 'ytd_09': 0.0,
                'ytd_10': 0.0, 'ytd_11': 0.0, 'ytd_12': 0.0,
                'monthly_01': 0.0, 'monthly_02': 0.0, 'monthly_03': 0.0,
                'monthly_04': 0.0, 'monthly_05': 0.0, 'monthly_06': 0.0,
                'monthly_07': 0.0, 'monthly_08': 0.0, 'monthly_09': 0.0,
                'monthly_10': 0.0, 'monthly_11': 0.0, 'monthly_12': 0.0,
                'year_2025': 0.0,
            }

        for p in projects:
            for cat, metrics in p.metrics.items():
                if cat not in agg:
                    continue
                for key in agg[cat]:
                    if key in metrics:
                        agg[cat][key] += metrics[key]

        # 计算偏差
        for cat in agg:
            agg[cat]['偏差'] = agg[cat]['current_actual'] - agg[cat]['budget']

        return agg

    def write_output(self,
                     current_month_list: List[ProjectData],
                     year_list: List[ProjectData]):
        """生成输出Excel文件（4个工作表）"""
        print(f"\n生成输出文件：{self.output_file.name}")

        # 加载输出模板（或创建新文件）
        if self.output_file.exists():
            wb = load_workbook(self.output_file)
            print(f"  已有工作表：{wb.sheetnames}")
        else:
            from openpyxl import Workbook
            wb = Workbook()
            # 删除默认sheet
            if 'Sheet' in wb.sheetnames:
                del wb['Sheet']
            print("  创建新文件")

        # 写入明细数据
        self._write_detail_sheet(wb, '当月项目明细数据', current_month_list)
        self._write_detail_sheet(wb, '2026年项目明细数据', year_list)

        # 写入汇总表
        self._write_summary_sheet(wb, '当月汇总毛利率',
                                  self.aggregate_by_category(current_month_list),
                                  f"{self.year}年{self.current_month}月在行农业项目")
        year_agg = self.aggregate_by_category(year_list)
        self._write_summary_sheet(wb, '2026年汇总毛利率',
                                  year_agg,
                                  f"{self.year}年在行农业项目")

        # 从2026年项目明细数据填充汇总表的年月列（用 year_agg 中的月度数据）
        self._fill_year_monthly_from_agg(wb, year_agg)

        wb.save(self.output_file)
        # 关闭源工作簿（不再需要读公式/值）
        if self._wb_source is not None:
            self._wb_source.close()
            self._wb_source = None
        if self._wb_source_values is not None:
            self._wb_source_values.close()
            self._wb_source_values = None
        print(f"  保存成功：{self.output_file}")
        return wb.sheetnames

    def _write_detail_sheet(self, wb, sheet_name: str, projects: List[ProjectData]):
        """写入明细数据工作表"""
        # 找到或创建工作表
        existing = [s for s in wb.sheetnames if normalize_text(s) == normalize_text(sheet_name)]
        if existing:
            del wb[existing[0]]
            ws = wb.create_sheet(sheet_name)
        else:
            ws = wb.create_sheet(sheet_name)

        # 获取源表第一行（用于复制表头格式）
        source_header_ws = None
        if projects:
            source_header_ws = projects[0].source_worksheet

        # 写入表头（同时复制源表第1行格式）
        header = self._build_detail_header()
        for col_idx, header_text in enumerate(header, 1):
            cell = ws.cell(row=1, column=col_idx, value=header_text)
            if source_header_ws is not None:
                src_cell = source_header_ws.cell(row=1, column=col_idx)
                copy_cell_style(src_cell, cell)

        # 需要在每个项目块（PROJECT_ROW_COUNT 行）中纵向合并的列（固定列，不依赖源表）
        # 列1-17：项目编码/名称/时间/月人数/结算方式 — 每个项目只有一个值，需跨 PROJECT_ROW_COUNT 行合并
        FIXED_MERGE_COLS = list(range(1, 18))  # 列1到列17

        # 收集所有需要重建的合并单元格（output_col范围的列表）
        all_output_merges = []

        # 写入项目数据
        target_row = 2
        for proj in projects:
            source_start = proj.start_row
            source_ws = proj.source_worksheet
            target_proj_start = target_row  # 记录该项目在输出表的起始行

            # 固定合并列：列1-17 每个项目块强制合并 PROJECT_ROW_COUNT 行
            for col in FIXED_MERGE_COLS:
                all_output_merges.append((
                    target_proj_start, target_proj_start + PROJECT_ROW_COUNT - 1,
                    col, col
                ))

            # 从源表获取其他合并单元格（列18以后，排除已固定的列）
            project_merges = find_project_merged_cells(source_ws, source_start, PROJECT_ROW_COUNT)
            for merge_col_start, merge_col_end in project_merges:
                out_col_start = self._source_col_to_output_col(merge_col_start)
                out_col_end = self._source_col_to_output_col(merge_col_end)
                # 跳过列1-17（已由固定合并覆盖），只处理列18以后的合并
                if (out_col_start is not None and out_col_end is not None
                        and out_col_start > 17):
                    all_output_merges.append((
                        target_proj_start, target_proj_start + PROJECT_ROW_COUNT - 1,  # 行范围（整个项目块）
                        out_col_start, out_col_end                 # 列范围
                    ))

            # 写入9个核算科目行
            for offset, cat in enumerate(CATEGORY_ORDER):
                metrics = proj.metrics.get(cat, {})
                source_row = source_start + offset
                # 项目基本信息只在第一行写入，后续行由合并单元格覆盖
                row_data = [
                    proj.code if offset == 0 else None,   # 1: 项目编码
                    proj.name if offset == 0 else None,   # 2: 项目名称
                    proj.start_date if offset == 0 else None,  # 3: 项目开始时间
                    proj.end_date if offset == 0 else None,    # 4: 项目结束时间
                    # 5-16: 月项目人数（从源表按行读取）
                    source_ws.cell(row=source_row, column=5).value,
                    source_ws.cell(row=source_row, column=6).value,
                    source_ws.cell(row=source_row, column=7).value,
                    source_ws.cell(row=source_row, column=8).value,
                    source_ws.cell(row=source_row, column=9).value,
                    source_ws.cell(row=source_row, column=10).value,
                    source_ws.cell(row=source_row, column=11).value,
                    source_ws.cell(row=source_row, column=12).value,
                    source_ws.cell(row=source_row, column=13).value,
                    source_ws.cell(row=source_row, column=14).value,
                    source_ws.cell(row=source_row, column=15).value,
                    source_ws.cell(row=source_row, column=16).value,
                    proj.settle_mode if offset == 0 else None,  # 17: 结算方式
                    cat,   # 18: 类别（核算科目名称，每行都要写）
                    metrics.get('budget', None),           # 19: 项目整体毛利目标预算
                    metrics.get('current_actual', None),   # 20: 项目开始-当前实际累计
                    metrics.get('2025_actual', None),      # 21: 项目开始-25年实际累计
                    # 22-33: YTD各月
                    metrics.get('ytd_12', None),  # 22: 26年YTD12月
                    metrics.get('ytd_11', None),  # 23: 26年YTD11月
                    metrics.get('ytd_10', None),  # 24: 26年YTD10月
                    metrics.get('ytd_09', None),  # 25: 26年YTD9月
                    metrics.get('ytd_08', None),  # 26: 26年YTD8月
                    metrics.get('ytd_07', None),  # 27: 26年YTD7月
                    metrics.get('ytd_06', None),  # 28: 26年YTD6月
                    metrics.get('ytd_05', None),  # 29: 26年YTD5月
                    metrics.get('ytd_04', None),  # 30: 26年YTD4月
                    metrics.get('ytd_03', None),  # 31: 26年YTD3月
                    metrics.get('ytd_02', None),  # 32: 26年YTD2月
                    metrics.get('ytd_01', None),  # 33: 26年YTD1月
                    metrics.get('偏差', None),      # 34: 项目周期偏差（源文件列34）
                    metrics.get('ytd_偏差', None), # 35: YTD当年偏差（源文件列35）
                    metrics.get('偏差原因', '') if offset == 0 else None,      # 36
                    metrics.get('ytd_偏差原因', '') if offset == 0 else None,  # 37
                    metrics.get('改进措施', '') if offset == 0 else None,      # 38
                    metrics.get('进展', '') if offset == 0 else None,          # 39
                    None,                                                    # 40: 真实经营数据
                    metrics.get('monthly_12', None),  # 41: 202612
                    metrics.get('monthly_11', None),  # 42: 202611
                    metrics.get('monthly_10', None),  # 43: 202610
                    metrics.get('monthly_09', None),  # 44: 202609
                    metrics.get('monthly_08', None),  # 45: 202608
                    metrics.get('monthly_07', None),  # 46: 202607
                    metrics.get('monthly_06', None),  # 47: 202606
                    metrics.get('monthly_05', None),  # 48: 202605
                    metrics.get('monthly_04', None),  # 49: 202604
                    metrics.get('monthly_03', None),  # 50: 202603
                    metrics.get('monthly_02', None),  # 51: 202602
                    metrics.get('monthly_01', None),  # 52: 202601
                    metrics.get('year_2025', None),   # 53: 2025年
                    metrics.get('year_2024', None),  # 54: 2024年
                    metrics.get('year_2023', None),  # 55: 2023年
                    metrics.get('remark', ''),        # 56: 备注
                ]
                assert len(row_data) == 56, f"row_data长度应为56，实际为{len(row_data)}"
                for col_idx, val in enumerate(row_data, 1):
                    target_cell = ws.cell(row=target_row, column=col_idx)

                    # 获取源单元格用于复制格式（跳过列5-16，由后续直接复制）
                    src_col = ProjectData._output_col_to_source_col(col_idx)
                    if src_col is not None:
                        src_cell = source_ws.cell(row=source_row, column=src_col)
                        copy_cell_style(src_cell, target_cell)

                    # 尝试从源工作表获取公式，有则重写行号后写入，否则写值
                    src_formula = proj.get_source_formula(offset, col_idx)
                    if src_formula is not None:
                        new_formula = rewrite_row_references_in_formula(
                            src_formula, source_start, target_proj_start
                        )
                        target_cell.value = new_formula
                    else:
                        target_cell.value = val

                target_row += 1

        # 重建合并单元格
        for min_row, max_row, min_col, max_col in all_output_merges:
            if min_col == max_col:
                # 单列合并（纵向合并）
                ws.merge_cells(
                    start_row=min_row, start_column=min_col,
                    end_row=max_row, end_column=max_col
                )

        print(f"  {sheet_name}：{len(projects)} 个项目 × {PROJECT_ROW_COUNT}科目 = {target_row - 2} 行（重建 {len(all_output_merges)} 个合并单元格）")

    def _source_col_to_output_col(self, source_col: int) -> Optional[int]:
        """将源表列号映射到输出表列号（反向映射）"""
        # 基于 col_map 构建反向映射：col_map[output_col] = source_col
        _reverse_col_map = {
            1: 1,    # 1→1 项目编码
            2: 2,    # 2→2 项目名称
            3: 3,    # 3→3 开始时间
            4: 4,    # 4→4 结束时间
            5: 5,    # 5→17 结算方式 (col_map: 5=17)
            17: 6,   # 6→18 类别 (col_map: 6=18) — wait, this is wrong in our col_map!
        }
        # 实际上 col_map = {1:1, 2:2, 3:3, 4:4, 5:17, 6:18, 7:19, 8:20, 9:21, ...}
        # 反向应该是 source_col → output_col
        # 所以 1→1, 2→2, 3→3, 4→4, 17→5, 18→6, 19→7, 20→8, 21→9, 22→10...
        _reverse_col_map = {
            1: 1, 2: 2, 3: 3, 4: 4,
            5: 5, 6: 6, 7: 7, 8: 8, 9: 9, 10: 10,  # 1-6月项目人数
            11: 11, 12: 12, 13: 13, 14: 14, 15: 15, 16: 16,  # 7-12月项目人数
            17: 17, 18: 18, 19: 19, 20: 20, 21: 21,  # 结算/类别/预算/实际
            22: 22, 23: 23, 24: 24, 25: 25, 26: 26, 27: 27, 28: 28,  # YTD 12月~7月
            29: 29, 30: 30, 31: 31, 32: 32, 33: 33,  # YTD 6月~1月
            34: 34, 35: 35,  # 项目周期偏差 / YTD当年偏差
            36: 36, 37: 37, 38: 38, 39: 39, 40: 40,  # 偏差原因/改进措施/进展
            41: 41, 42: 42, 43: 43, 44: 44, 45: 45, 46: 46,  # 月度 202612~202607
            47: 47, 48: 48, 49: 49, 50: 50, 51: 51, 52: 52,  # 月度 202606~202601
            53: 53, 54: 54, 55: 55, 56: 56,  # 2025年/2024年/2023年/备注
        }
        return _reverse_col_map.get(source_col)

    def _write_summary_sheet(self, wb, sheet_name: str, agg: Dict,
                             title: str):
        """
        写入汇总毛利率工作表。
        优先保留用户已调整过的格式和公式，仅更新需要刷新的数据。
        新建表时会写入标准公式（与参考文件一致）。
        """
        existing = [s for s in wb.sheetnames if normalize_text(s) == normalize_text(sheet_name)]
        is_new_sheet = not existing
        if existing:
            ws = wb[existing[0]]
        else:
            ws = wb.create_sheet(sheet_name)

        # 从参考文件复制表头格式
        ref_wb = None
        try:
            ref_wb = load_workbook(
                'C:/Users/Administrator/.workbuddy/skills/agriculture-stats/references/农业项目群预算毛利表汇总表-26.4.17.xlsx',
                data_only=False
            )
            # 优先找同名工作表，否则找第一个
            ref_sheet_name = next(
                (n for n in ref_wb.sheetnames if normalize_text(n) == normalize_text(sheet_name)),
                ref_wb.sheetnames[0] if ref_wb.sheetnames else None
            )
            ref_ws = ref_wb[ref_sheet_name] if ref_sheet_name else None
        except Exception:
            ref_ws = None

        # 行1：标题
        c = ws.cell(row=1, column=1, value=title)
        if ref_ws and ref_ws.cell(row=1, column=1).has_style:
            copy_cell_style(ref_ws.cell(row=1, column=1), c)

        # 行3：列标签
        headers = [
            '项目名称', '类别', '项目整体毛利目标预算',
            '项目开始-当前实际累计', '2026年YTD', '偏差（实际-预算）',
            '数据分析与说明',
        ] + [202612 - i for i in range(12)]  # 年月用整数 202612, 202611, ...

        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=h)
            if ref_ws:
                ref_cell = ref_ws.cell(row=3, column=col_idx)
                copy_cell_style(ref_cell, cell)

        # 行4+：数据
        project_count = len(self.agri_projects)
        data_start_row = 4
        row = data_start_row

        # 确定月度数据顺序
        monthly_keys = [
            'monthly_12', 'monthly_11', 'monthly_10', 'monthly_09',
            'monthly_08', 'monthly_07', 'monthly_06', 'monthly_05',
            'monthly_04', 'monthly_03', 'monthly_02', 'monthly_01',
        ]

        # 汇总表行号映射（10行：4-13）
        # 行4=累计收入, 5=人工成本, 6=闲置成本, 7=差旅费, 8=其他成本, 9=累计成本
        # 行10=GP1毛利, 11=GP1毛利率, 12=GP3毛利, 13=GP3毛利率
        R_NR = 4          # 累计收入
        R_LABOR = 5       # 人工成本
        R_IDLE = 6        # 闲置成本
        R_TRAVEL = 7      # 差旅费
        R_OTHER = 8       # 其他成本
        R_TOTAL_COST = 9  # 累计成本
        R_GP1 = 10        # GP1毛利
        R_GP1_RATE = 11   # GP1毛利率
        R_GP3 = 12        # GP3毛利
        R_GP3_RATE = 13   # GP3毛利率

        # GP行公式模板（按列号生成）
        # YTD列(E)=SUM(H行:S行), 偏差列(F)=D行-C行
        # GP1毛利: C=NR-人工, GP3毛利: C=NR-人工-差旅-其他
        def _gp_formula(cat_name: str, col_idx: int, col_letter: str, row: int) -> Optional[str]:
            """根据类别和列生成GP行公式。col_idx=2(类别列)不生成公式"""
            if col_idx == 2:  # 类别列保持文本
                return None
            col_l = col_letter
            if cat_name == '累计成本':
                # =SUM(人工:其他) 即 E5:E8（YTD列）或 列字母+行号
                if col_l == 'E':
                    return f'=SUM(E{R_LABOR}:E{R_OTHER})'
                else:
                    return f'=SUM({col_l}{R_LABOR}:{col_l}{R_OTHER})'
            elif cat_name == 'GP1毛利':
                # =NR-人工成本（跳过闲置成本）
                return f'={col_l}{R_NR}-{col_l}{R_LABOR}'
            elif cat_name == 'GP1毛利率(%)':
                return f'={col_l}{R_GP1}/{col_l}{R_NR}'
            elif cat_name == 'GP3毛利':
                # =NR-人工-差旅-其他（跳过闲置成本）
                return f'={col_l}{R_NR}-{col_l}{R_LABOR}-{col_l}{R_TRAVEL}-{col_l}{R_OTHER}'
            elif cat_name == 'GP3毛利率(%)':
                return f'={col_l}{R_GP3}/{col_l}{R_NR}'
            return None

        for cat in CATEGORY_ORDER:
            if cat not in agg:
                continue
            m = agg[cat]
            row_data = [
                f"{project_count}个农业项目",
                cat,
                m['budget'],
                m['current_actual'],
                '',   # 2026年YTD（公式列）
                '',   # 偏差（公式列）
                '',   # 数据分析与说明
            ] + [m.get(k, 0.0) for k in monthly_keys]

            for col_idx, val in enumerate(row_data, 1):
                target_cell = ws.cell(row=row, column=col_idx)
                # 合并单元格的次级单元格（MergedCell）不能写入，跳过
                if isinstance(target_cell, MergedCell):
                    continue
                # 已有公式的单元格 → 保留公式（不覆盖用户手动调整的公式）
                if target_cell.data_type == 'f' and target_cell.value:
                    if ref_ws:
                        ref_cell = ref_ws.cell(row=row, column=col_idx)
                        copy_cell_style(ref_cell, target_cell)
                    continue

                # 确定写入内容
                col_letter = get_column_letter(col_idx)
                write_val = val

                # 公式列处理：YTD(E)、偏差(F)、GP行所有列
                if col_idx in (5, 6) or cat in ('累计成本', 'GP1毛利', 'GP1毛利率(%)', 'GP3毛利', 'GP3毛利率(%)'):
                    formula = _gp_formula(cat, col_idx, col_letter, row)
                    if formula:
                        write_val = formula
                    elif col_idx == 5 and cat not in ('累计成本', 'GP1毛利', 'GP1毛利率(%)', 'GP3毛利', 'GP3毛利率(%)'):
                        # YTD列：数据行用 SUM(H行:S行)
                        write_val = f'=SUM(H{row}:S{row})'
                    elif col_idx == 6:
                        # 偏差列：=D行-C行
                        write_val = f'=D{row}-C{row}'

                if write_val != '':
                    target_cell.value = write_val if write_val != '' else None
                # 从参考文件复制格式
                if ref_ws:
                    ref_cell = ref_ws.cell(row=row, column=col_idx)
                    copy_cell_style(ref_cell, target_cell)
            row += 1

        # 合并单元格：项目名称列(A) 和 数据分析与说明列(G)
        data_end_row = data_start_row + PROJECT_ROW_COUNT - 1
        merge_ranges = [
            (data_start_row, data_end_row, 1, 1),  # A4:A13
            (data_start_row, data_end_row, 7, 7),  # G4:G13
        ]
        for min_row, max_row, min_col, max_col in merge_ranges:
            merge_str = f'{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}'
            # 先解除已有合并（避免重复）
            existing_merges = [m for m in ws.merged_cells.ranges
                               if m.min_row == min_row and m.max_row == max_row
                               and m.min_col == min_col and m.max_col == max_col]
            if not existing_merges:
                ws.merge_cells(merge_str)

        if ref_wb:
            ref_wb.close()

        print(f"  {sheet_name}：{row - data_start_row} 个科目")

    def _fill_year_monthly_from_agg(self, wb, agg: Dict):
        """
        用聚合数据 agg 中的月度字段填充2026年汇总毛利率的年月列。
        汇总表年月列：col8=2026/12, col9=2026/11, ..., col19=2026/01
        agg 中月度 key：monthly_12, monthly_11, ..., monthly_01（对应2026/12→2026/01）

        规则：
        - 仅填充数据行（累计收入/人工成本/闲置成本/差旅费/其他成本/累计成本），跳过GP行
        - 已有公式的单元格 → 保留公式不覆盖
        - 已有数值数据（且非0）的单元格 → 跳过不覆盖
        - 源数据为0或None → 跳过该列
        - 不填充项目名称列
        """
        try:
            summary_ws = wb['2026年汇总毛利率']
        except KeyError:
            return

        # 汇总表数据行（跳过GP行）
        # 行4=累计收入, 行5=人工成本, 行6=闲置成本, 行7=差旅费, 行8=其他成本, 行9=累计成本
        category_summary_row = {
            '累计收入（NR）': 4,
            '人工成本（成本１）': 5,
            '闲置成本（成本2）': 6,
            '差旅费（成本３）': 7,
            '其他成本（成本３）': 8,
            '累计成本': 9,
        }

        # 汇总表列号 → agg 月度 key（col8=monthly_12, ..., col19=monthly_01）
        month_keys = [f'monthly_{i:02d}' for i in range(12, 0, -1)]  # monthly_12→monthly_01
        summary_col_start = 8

        for cat, summary_row in category_summary_row.items():
            if cat not in agg:
                continue
            m = agg[cat]
            for offset, month_key in enumerate(month_keys):
                summary_col = summary_col_start + offset
                val = m.get(month_key, 0.0)
                if val is None or val == 0.0:
                    continue  # 源数据为空或0，跳过

                target_cell = summary_ws.cell(row=summary_row, column=summary_col)
                # 已有公式的单元格 → 保留
                if target_cell.data_type == 'f' and target_cell.value:
                    continue
                # 已有数据（且非0）的单元格 → 跳过
                if target_cell.value not in (None, 0):
                    continue
                target_cell.value = val

        print(f"  年月列填充完成（2026年）")

    def _build_detail_header(self) -> List[str]:
        """构建明细表表头（共56列）"""
        return [
            '项目编码', '项目名称', '项目开始时间', '项目结束时间',
            '1月项目人数', '2月项目人数', '3月项目人数', '4月项目人数',
            '5月项目人数', '6月项目人数', '7月项目人数', '8月项目人数',
            '9月项目人数', '10月项目人数', '11月项目人数', '12月项目人数',
            '结算方式', '类别',
            '项目整体毛利目标预算', '项目开始-当前实际累计', '项目开始-25年实际累计',
            '26年YTD实际累计12月', '26年YTD实际累计11月', '26年YTD实际累计10月',
            '26年YTD实际累计9月', '26年YTD实际累计8月', '26年YTD实际累计7月',
            '26年YTD实际累计6月', '26年YTD实际累计5月', '26年YTD实际累计4月',
            '26年YTD实际累计3月', '26年YTD实际累计2月', '26年YTD实际累计1月',
            '项目周期偏差（实际-预估）', 'YTD当年偏差（实际-预估）',
            '项目周期偏差原因分析（截止25年底毛利偏差原因）',
            'YTD当年偏差原因分析', '改进措施', '进展', '真实经营数据',
            202612, 202611, 202610, 202609, 202608, 202607,
            202606, 202605, 202604, 202603, 202602, 202601,
            '2025年', '2024年', '2023年', '备注'
        ]

    def run(self) -> dict:
        """执行完整统计流程"""
        print("=" * 70)
        print("农业项目统计")
        print("=" * 70)

        count = self.load_data()
        if count == 0:
            print("\n警告：未找到农业项目！")
            print(f"请检查关键字配置，或确认输入表中是否有包含以下关键字的项目：")
            print(f"  {self.include_kw}")

        current_month_list, year_list, closed_list = self.classify_projects()

        sheets = self.write_output(current_month_list, year_list)

        print("\n" + "=" * 70)
        print("统计完成")
        print("=" * 70)

        return {
            'total_projects': len(self.projects),
            'agri_projects': count,
            'current_month_projects': len(current_month_list),
            'year_projects': len(year_list),
            'closed_projects': len(closed_list),
            'output_sheets': sheets,
        }


def main():
    """命令行入口"""
    import argparse

    parser = argparse.ArgumentParser(description='农业项目统计')
    parser.add_argument('--input', '-i', required=True, help='输入Excel文件')
    parser.add_argument('--output', '-o', required=True, help='输出Excel文件')
    parser.add_argument('--month', '-m', type=int, default=None, help='当前月份')
    parser.add_argument('--year', '-y', type=int, default=None, help='统计年份')
    parser.add_argument('--keywords', '-k', type=str, default=None,
                        help='关键字配置文件路径')
    parser.add_argument('--sheet', '-s', type=str, default=None,
                        help=f'数据工作表名称（默认：{DEFAULT_DATA_SHEET}）')

    args = parser.parse_args()

    stats = AgricultureStats(
        input_file=Path(args.input),
        output_file=Path(args.output),
        current_month=args.month,
        year=args.year,
        keywords_path=Path(args.keywords) if args.keywords else None,
        sheet_name=args.sheet,
    )
    result = stats.run()
    print(f"\n结果：{result}")


if __name__ == '__main__':
    main()