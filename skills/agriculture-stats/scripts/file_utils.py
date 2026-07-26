"""
文件查找工具模块
支持模糊查找文件名，忽略：
- 中文/英文标点（全角/半角）
- 空格
- 大小写
"""

import os
import re
from pathlib import Path
from typing import List, Optional, Tuple


# ============================================================
# 文本标准化
# ============================================================

def normalize_text(text: str) -> str:
    """
    标准化文本：转换全角字符为半角，移除空格，转小写

    Args:
        text: 原始文本

    Returns:
        标准化后的文本
    """
    if not isinstance(text, str):
        return ''
    # 全角转半角映射
    fullwidth_to_halfwidth = {chr(0xFF01 + i): chr(0x21 + i) for i in range(94)}
    fullwidth_to_halfwidth.update({chr(0xFF10 + i): chr(0x30 + i) for i in range(10)})
    fullwidth_to_halfwidth.update({chr(0xFF21 + i): chr(0x41 + i) for i in range(26)})
    fullwidth_to_halfwidth.update({chr(0xFF41 + i): chr(0x61 + i) for i in range(26)})
    fullwidth_to_halfwidth['　'] = ' '
    result = ''.join(fullwidth_to_halfwidth.get(c, c) for c in text)
    # 移除所有空白字符
    result = re.sub(r'\s+', '', result)
    result = result.lower()
    return result


# ============================================================
# 文件模糊查找
# ============================================================

def find_file_fuzzy(
    filename_pattern: str,
    search_dir: str = '.',
    recursive: bool = True
) -> Optional[Path]:
    """
    模糊查找文件，返回第一个匹配的文件路径

    Args:
        filename_pattern: 文件名模式（支持部分匹配）
        search_dir: 搜索目录
        recursive: 是否递归搜索子目录

    Returns:
        匹配的文件路径，未找到返回 None
    """
    search_path = Path(search_dir)
    pattern_normalized = normalize_text(filename_pattern)

    if '*' in filename_pattern or '?' in filename_pattern:
        matches = list(search_path.rglob(filename_pattern) if recursive
                      else search_path.glob(filename_pattern))
    else:
        matches = list(search_path.rglob('*') if recursive
                      else search_path.glob('*'))

    files = [m for m in matches if m.is_file()]
    best_match = None
    best_score = 0

    for file_path in files:
        filename = file_path.name
        filename_norm = normalize_text(filename)
        if pattern_normalized in filename_norm:
            score = len(pattern_normalized) / len(filename_norm)
            if score > best_score:
                best_score = score
                best_match = file_path

    return best_match


def find_files_fuzzy(
    filename_pattern: str,
    search_dir: str = '.',
    recursive: bool = True
) -> List[Path]:
    """
    模糊查找所有匹配的文件

    Returns:
        匹配的文件路径列表
    """
    search_path = Path(search_dir)
    pattern_normalized = normalize_text(filename_pattern)

    if '*' in filename_pattern or '?' in filename_pattern:
        matches = list(search_path.rglob(filename_pattern) if recursive
                      else search_path.glob(filename_pattern))
    else:
        matches = list(search_path.rglob('*') if recursive
                      else search_path.glob('*'))

    files = [m for m in matches if m.is_file()]
    matched_files = []
    for file_path in files:
        filename_norm = normalize_text(file_path.name)
        if pattern_normalized in filename_norm:
            matched_files.append(file_path)
    return sorted(matched_files)


def find_excel_sheet_fuzzy(
    excel_path: str,
    sheet_pattern: str
) -> Optional[str]:
    """
    模糊查找 Excel 工作表名称

    Args:
        excel_path: Excel 文件路径
        sheet_pattern: 工作表名称模式

    Returns:
        匹配的工作表名称，未找到返回 None
    """
    import pandas as pd
    xl = pd.ExcelFile(excel_path)
    pattern_normalized = normalize_text(sheet_pattern)

    best_match = None
    best_score = 0

    for sheet_name in xl.sheet_names:
        sheet_norm = normalize_text(sheet_name)
        if pattern_normalized in sheet_norm:
            score = len(pattern_normalized) / len(sheet_norm)
            if score > best_score:
                best_score = score
                best_match = sheet_name

    return best_match


# ============================================================
# 目录工具
# ============================================================

def get_references_dir() -> Path:
    """
    获取 references 目录路径

    搜索顺序：
    1. scripts/references（相对于当前脚本）
    2. .workbuddy/skills/agriculture-stats/references
    3. 当前工作目录/references
    """
    current_dir = Path(__file__).parent
    # 1. scripts/references
    ref_dir = current_dir / 'references'
    if ref_dir.exists():
        return ref_dir
    # 2. .workbuddy/skills/agriculture-stats/references
    ref_dir = current_dir.parent.parent / 'references'
    if ref_dir.exists():
        return ref_dir
    # 3. cwd/references
    ref_dir = Path.cwd() / 'references'
    if ref_dir.exists():
        return ref_dir
    raise FileNotFoundError("未找到 references 目录")


# ============================================================
# 便捷函数：农业项目专用
# ============================================================

def find_input_file(search_dir: str = '.') -> Optional[Path]:
    """
    查找输入数据文件（毛利管控表）

    支持的文件名模式：
    - 毛利管控表
    - FP项目毛利
    - 电信AIoT毛利
    - 2026毛利管控
    """
    patterns = [
        "毛利管控表",
        "FP项目毛利",
        "电信AIoT毛利",
        "2026毛利管控",
    ]
    for pattern in patterns:
        result = find_file_fuzzy(pattern, search_dir)
        if result:
            return result
    return None


def find_output_template_file(search_dir: str = '.') -> Optional[Path]:
    """
    查找输出模板文件（预算毛利汇总表）

    支持的文件名模式：
    - 预算毛利表汇总
    - 农业项目预算毛利
    - 毛利汇总表
    """
    patterns = [
        "预算毛利表汇总",
        "农业项目预算毛利",
        "毛利汇总表",
    ]
    for pattern in patterns:
        result = find_file_fuzzy(pattern, search_dir)
        if result:
            return result
    return None


# ============================================================
# 测试入口
# ============================================================

if __name__ == '__main__':
    print("=== 文件查找工具测试 ===\n")

    test_strings = [
        "2026年电信与AIoT业务线FP项目毛利管控表20260310.xlsx",
        "毛　利　管　控　表",
        "农业项目群预算毛利表汇总表-26.2.28.xlsx",
    ]
    print("文本标准化测试:")
    for s in test_strings:
        print(f"  原始：{s}")
        print(f"  标准化：{normalize_text(s)}")
        print()

    try:
        ref_dir = get_references_dir()
        print(f"References 目录：{ref_dir}\n")

        print("查找输入文件（毛利管控表）:")
        input_file = find_input_file(str(ref_dir))
        print(f"  找到：{input_file.name if input_file else '未找到'}\n")

        print("查找输出模板文件（预算毛利汇总表）:")
        output_file = find_output_template_file(str(ref_dir))
        print(f"  找到：{output_file.name if output_file else '未找到'}\n")

        print("模糊查找测试:")
        for pattern in ["毛利管控", "预算毛利", "实际运营数据"]:
            matches = find_files_fuzzy(pattern, str(ref_dir))
            print(f"  模式 '{pattern}': 找到 {len(matches)} 个文件")
            for m in matches[:3]:
                print(f"    - {m.name}")
    except FileNotFoundError as e:
        print(f"警告：{e}")