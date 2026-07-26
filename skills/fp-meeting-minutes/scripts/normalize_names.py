#!/usr/bin/env python3
"""
会议口水稿项目名称规范化工具。

读取 FP项目名称清单模板.xlsx，构建关键词索引，对口水稿文本执行
语音识别错误修正、口语化名称规范化、项目经理关联替换。

用法:
    python normalize_names.py <excel_path> <transcript_path> [output_path]

输出:
    规范化后的文本文件（默认覆盖原文件，或写入 output_path）
"""

import sys
import re

try:
    import openpyxl
except ImportError:
    print("错误：需要 openpyxl。请执行 pip install openpyxl")
    sys.exit(1)


def load_project_index(excel_path: str) -> dict:
    """
    从Excel加载项目名称索引。
    返回结构:
        {
            "pm_to_projects": { "向锐": ["成都鼎桥-2026年ICP框架...", ...], ... },
            "keyword_to_project": { "鼎桥": "成都鼎桥-2026年ICP框架...", "ICP": "...", ... },
            "full_names": ["完整项目名1", ...],
        }
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    pm_to_projects = {}
    keyword_to_project = {}
    full_names = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]

        # 尝试定位关键列
        col_pm = None
        col_project = None
        for idx, h in enumerate(headers):
            if h and "PM" in str(h).upper() and "姓名" in str(h):
                col_pm = idx
            if h and "项目名称" in str(h):
                col_project = idx

        # 如果按表头没找到，尝试按常见位置兜底
        if col_pm is None:
            col_pm = 5
        if col_project is None:
            col_project = 2

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) <= max(col_pm, col_project):
                continue
            pm = str(row[col_pm]).strip() if row[col_pm] else None
            project = str(row[col_project]).strip() if row[col_project] else None

            if not project or project == "nan" or project == "None":
                continue

            full_names.append(project)

            if pm and pm != "nan" and pm != "None":
                pm_to_projects.setdefault(pm, []).append(project)

            # 提取关键词：取项目名中不含数字和符号的连续中文字段
            keywords = re.findall(r'[一-鿿]+', project)
            for kw in keywords:
                if len(kw) >= 2:
                    keyword_to_project[kw] = project
            # 同时提取常见的英文/数字简称
            tokens = re.findall(r'[A-Za-z0-9]+', project)
            for tok in tokens:
                if len(tok) >= 2:
                    keyword_to_project[tok.upper()] = project

    # 去重并保持最简映射：长关键词优先（避免短词覆盖长词）
    sorted_keywords = sorted(keyword_to_project.keys(), key=len, reverse=True)
    deduped = {}
    for k in sorted_keywords:
        if k not in deduped:
            deduped[k] = keyword_to_project[k]

    return {
        "pm_to_projects": pm_to_projects,
        "keyword_to_project": deduped,
        "full_names": list(set(full_names)),
    }


def normalize_transcript(text: str, index: dict) -> str:
    """
    对口水稿文本执行名称规范化。
    策略：按关键词长度降序全文替换，避免短词误伤长词。
    """
    # 1. 关键词替换（长词优先）
    keywords = sorted(index["keyword_to_project"].keys(), key=len, reverse=True)
    for kw in keywords:
        if kw in text:
            project = index["keyword_to_project"][kw]
            short = _extract_short_name(project)
            text = text.replace(kw, short)

    return text


def _extract_short_name(full_name: str) -> str:
    """从完整项目名中提取标准简称。"""
    keywords = re.findall(r'[一-鿿]+', full_name)
    if not keywords:
        return full_name
    stopwords = {"项目", "合同", "服务", "采购", "开发", "年度", "公司", "移动", "年", "FP"}
    filtered = [k for k in keywords if k not in stopwords and len(k) >= 2]
    if len(filtered) >= 2:
        return "".join(filtered[-3:])
    if len(filtered) == 1:
        return filtered[0]
    return full_name[:30]


def main():
    if len(sys.argv) < 3:
        print("用法: python normalize_names.py <excel_path> <transcript_path> [output_path]")
        sys.exit(1)

    excel_path = sys.argv[1]
    transcript_path = sys.argv[2]
    output_path = sys.argv[3] if len(sys.argv) > 3 else transcript_path

    print(f"[1/3] 加载项目清单: {excel_path}")
    index = load_project_index(excel_path)
    print(f"      发现 {len(index['full_names'])} 个项目, {len(index['keyword_to_project'])} 个关键词, {len(index['pm_to_projects'])} 位PM")

    print(f"[2/3] 读取口水稿: {transcript_path}")
    with open(transcript_path, "r", encoding="utf-8") as f:
        text = f.read()

    print(f"[3/3] 执行规范化...")
    normalized = normalize_transcript(text, index)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(normalized)

    print(f"      完成，已写入: {output_path}")


if __name__ == "__main__":
    main()
