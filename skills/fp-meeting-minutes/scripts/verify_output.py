#!/usr/bin/env python3
"""
会议纪要 docx 输出验证工具。

检查生成的 docx 是否符合格式规范，输出验证报告。

用法:
    python verify_output.py <input.docx> [--type fp|solution]

检查项:
    - 标题格式（居中、加粗、22pt、微软雅黑）
    - 遗留事项表格列数和表头（3列=排产，7列=解决方案）
    - 同一项目是否合并到 1 行（仅排产评审）
    - 表格单元格是否有意外缩进
    - 是否包含"重点事项与管理要求说明"章节（解决方案业务会议）
"""

import argparse
import sys

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt


def _check_font(run, expected_name="微软雅黑"):
    """检查字体是否为微软雅黑（三个属性任一匹配即可）。"""
    rFonts = run._element.rPr.rFonts if run._element.rPr is not None else None
    if rFonts is not None:
        east_asia = rFonts.get(qn("w:eastAsia"))
        if east_asia == expected_name:
            return True
    return run.font.name == expected_name


def check_title_format(doc, meeting_type="fp"):
    """检查主标题格式。"""
    issues = []
    title_paras = []
    for p in doc.paragraphs:
        if p.text.strip():
            title_paras.append(p)
            if meeting_type == "solution" and len(title_paras) >= 2:
                break
            if meeting_type == "fp" and len(title_paras) >= 1:
                break

    if not title_paras:
        issues.append("未找到主标题段落")
        return issues

    if meeting_type == "solution":
        expected_titles = ["电信与AIoT业务线", "解决方案项目月度管理会议纪要"]
        for i, expected in enumerate(expected_titles):
            if i >= len(title_paras):
                issues.append(f"缺少第 {i+1} 行标题，期望: '{expected}'")
                continue
            para = title_paras[i]
            if expected not in para.text:
                issues.append(f"第 {i+1} 行标题内容不匹配。期望包含: '{expected}', 实际: '{para.text.strip()}'")
            if para.alignment != 1:
                issues.append(f"第 {i+1} 行标题未居中")
            for run in para.runs:
                if not run.font.bold:
                    issues.append(f"第 {i+1} 行标题未加粗")
                break
    else:
        para = title_paras[0]
        if para.alignment != 1:
            issues.append("主标题未居中")
        for run in para.runs:
            if not run.font.bold:
                issues.append("主标题未加粗")
            if run.font.size != Pt(22):
                issues.append(f"主标题字号不是22pt (当前: {run.font.size})")
            if not _check_font(run):
                issues.append("主标题字体不是微软雅黑")
            break
    return issues


def check_todo_table(doc, meeting_type="fp"):
    """检查遗留事项表格。"""
    issues = []
    found = False

    if meeting_type == "fp":
        expected_cols = 3
        expected_headers = ["项目", "遗留问题/待办事项", "责任人"]
        for table in doc.tables:
            col_count = len(table.columns)
            if col_count == expected_cols:
                found = True
                headers = [cell.text.strip() for cell in table.rows[0].cells]
                if headers != expected_headers:
                    issues.append(f"表格表头不匹配。期望: {expected_headers}, 实际: {headers}")
                # 检查同一项目是否合并到1行
                projects_seen = []
                for row in table.rows[1:]:
                    cells = row.cells
                    if len(cells) < 3:
                        continue
                    proj = cells[0].text.strip()
                    if proj in projects_seen:
                        issues.append(f"项目 '{proj}' 在遗留事项表格中重复出现，未合并到1行")
                    projects_seen.append(proj)
                break
    else:
        expected_cols = 7
        expected_headers = ["序号", "问题/事项", "责任人", "预计闭环\n时间", "实际闭环\n时间", "当前\n状态", "当前进展"]
        for table in doc.tables:
            col_count = len(table.columns)
            if col_count == expected_cols:
                found = True
                headers = [cell.text.strip() for cell in table.rows[0].cells]
                if headers != expected_headers:
                    issues.append(f"表格表头不匹配。期望: {expected_headers}, 实际: {headers}")
                break

    if not found:
        issues.append(f"未找到{expected_cols}列的遗留事项跟踪清单表格")
    return issues


def check_cell_indent(doc):
    """检查表格单元格是否有意外缩进。"""
    issues = []
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    pPr = p._p.find(qn("w:pPr"))
                    if pPr is not None:
                        ind = pPr.find(qn("w:ind"))
                        if ind is not None:
                            flc = ind.get(qn("w:firstLineChars"))
                            fl = ind.get(qn("w:firstLine"))
                            if flc and flc != "0":
                                issues.append(f"发现单元格缩进: firstLineChars={flc} (单元格文本: '{p.text[:30]}...')")
                            if fl and fl != "0":
                                issues.append(f"发现单元格缩进: firstLine={fl} (单元格文本: '{p.text[:30]}...')")
    return issues


def check_leadership_section(doc):
    """检查是否包含重点事项与管理要求说明章节。"""
    issues = []
    found = False
    for p in doc.paragraphs:
        text = p.text.strip()
        if "重点事项与管理要求说明" in text:
            found = True
            break
    if not found:
        issues.append("未找到'重点事项与管理要求说明'章节（解决方案业务会议必须包含）")
    return issues


def check_meeting_info_table(doc, meeting_type="fp"):
    """检查会议信息表格结构。"""
    issues = []
    for table in doc.tables:
        col_count = len(table.columns)
        if meeting_type == "fp" and col_count == 4:
            headers = [cell.text.strip() for cell in table.rows[0].cells]
            expected_labels = {"会议名称", "会议类别", "会议地点", "会议时间", "参会人员", "纪要整理", "纪要审核"}
            found_labels = set()
            for row in table.rows:
                cells = row.cells
                if cells:
                    found_labels.add(cells[0].text.strip())
            if not expected_labels.issubset(found_labels):
                missing = expected_labels - found_labels
                issues.append(f"FP会议信息表缺少标签: {missing}")
            return issues
        elif meeting_type == "solution" and col_count == 2:
            headers = [cell.text.strip() for cell in table.rows[0].cells]
            expected_labels = {"主题", "时间", "会议方式", "召集人", "纪要整理", "与会人", "议题"}
            found_labels = set()
            for row in table.rows:
                cells = row.cells
                if cells:
                    found_labels.add(cells[0].text.strip())
            if not expected_labels.issubset(found_labels):
                missing = expected_labels - found_labels
                issues.append(f"解决方案会议信息表缺少标签: {missing}")
            return issues

    expected = "4列" if meeting_type == "fp" else "2列"
    issues.append(f"未找到{expected}的会议信息表格")
    return issues


def main():
    parser = argparse.ArgumentParser(description="验证会议纪要 docx 输出")
    parser.add_argument("input", help="输入 docx 文件路径")
    parser.add_argument("--type", default="fp", choices=["fp", "solution"],
                        help="会议类型: fp=排产评审, solution=解决方案业务月度管理")
    args = parser.parse_args()

    input_path = args.input
    doc = Document(input_path)

    all_issues = []
    all_issues.extend(check_title_format(doc, args.type))
    all_issues.extend(check_meeting_info_table(doc, args.type))
    all_issues.extend(check_todo_table(doc, args.type))
    all_issues.extend(check_cell_indent(doc))
    if args.type == "solution":
        all_issues.extend(check_leadership_section(doc))

    print(f"验证文件: {input_path}")
    print(f"会议类型: {args.type}")
    print(f"检查项数: 5")
    if all_issues:
        print(f"发现问题: {len(all_issues)} 项")
        for i, issue in enumerate(all_issues, 1):
            print(f"  {i}. {issue}")
        sys.exit(1)
    else:
        print("验证通过，未发现格式问题。")
        sys.exit(0)


if __name__ == "__main__":
    main()
