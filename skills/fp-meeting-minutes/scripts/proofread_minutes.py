#!/usr/bin/env python3
"""
会议纪要精校脚本。

用法:
    python proofread_minutes.py <docx_path> [--output report.md]

功能：
    1. 读取生成的 docx 文件，提取所有段落和表格文本
    2. 应用多种校对规则，检查语句不通顺、错别字、语音转译错误
    3. 输出 Markdown 格式的校对报告，列出明显错误、不同地方、疑问点
    4. 报告结构清晰，便于用户逐项确认后统一修正

检查规则覆盖：
    - 重复词/句（如"签完后马上能签"）
    - 常见错别字（如"前项"→"前期"）
    - 口语化表达（如"搞"、"甩"）
    - 缺少主语（如"需进一步提供新数据"）
    - 语音转译错误（如"起作业"→"启动作业"）
    - 金额/日期不一致与逻辑矛盾
    - 中英文标点混用
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from docx import Document


# ── 数据模型 ───────────────────────────────────────────────────────

@dataclass
class Issue:
    """单个校对问题。"""
    level: str           # "error" | "warning" | "question"
    category: str        # 类别标签
    location: str        # 位置描述（如"正文-黄海博-国投项目"）
    original: str        # 原文
    suggestion: str      # 建议修改
    reason: str          # 原因说明


# ── 校对规则库 ───────────────────────────────────────────────────────

# 常见错别字/用词不当（错误模式 → 建议）
COMMON_ERRORS = {
    r"签完后马上能签": ("签完后马上推进立项", "重复表述，语义不通"),
    r"前项[未未]": ("前期未", "'前项'应为'前期'或'前序'"),
    r"前向[未未]": ("前期未", "'前向'应为'前期'或'前序'"),
    r"跟客户起作业": ("跟客户启动作业", "语音转译错误，'起作业'应为'启动作业'"),
    r"8月跟客户起": ("8月跟客户启动", "语音转译错误"),
    r"要数据也[不没]": ("索要数据对方也不提供", "缺少主语，表述不完整"),
    r"搞鸿蒙": ("暂停鸿蒙开发", "口语化表达，建议改为正式表述"),
    r"销售甩过来": ("销售转交的", "口语化，'甩'过于随意"),
    r"销售甩项目": ("销售转交的项目", "口语化，'甩'过于随意"),
    r"至少开到下一个": ("至少覆盖到下一个", "'开到'表述不正式，建议改为'覆盖到'或'延长至'"),
    r"深度上面提升": ("深度上提升", "'上面'多余，应为'上'或'方面'"),
    r"深度上面": ("深度上", "'上面'多余"),
    r"给到": ("提交给", "口语化，建议改为'提交给'或'提供给'"),
    r"拉通": ("打通/协调", "'拉通'是行业黑话，对外文档建议改为'打通'或'协调'"),
    r"对齐": ("协调/确认", "'对齐'是行业黑话，建议改为'协调'或'确认'"),
    r"冲量": ("提升业务量", "口语化"),
    r"过年拉低": ("春节假期拉低", "'过年'表述不够正式"),
    r"王岚那边": ("王岚处", "口语化，'那边'建议改为'处'或'部门'"),
    r"江总邮件未回": ("江总邮件未回复", "'未回'应为'未回复'"),
    r"nopo": ("NoPO", "大小写不统一，应统一为 NoPO"),
    r"NoPO立项流程.*必须.*7月内": ("NoPO立项流程，必须在7月内完成", "缺少'完成'等动词，语义不完整"),
    r"等待排进议程": ("等待排上议程", "'排进'应为'排上'或'列入'"),
    r"给过后": ("交付后", "口语化，'给过后'应为'交付后'或'提交后'"),
    r"小样本模型给过": ("小样本模型提交后", "口语化"),
    r"被舒畅拿去": ("由舒畅承接", "口语化，'拿去'过于随意"),
    r"等陶颖沟通确认后找客户": ("等陶颖沟通确认后，再找客户", "缺少逗号，语句不通顺"),
    r"与陶颖沟通确认后找客户": ("与陶颖沟通确认后，再找客户", "缺少逗号，语句不通顺"),
    r"预计8月初可签合同，如果7月底": ("预计8月初可签合同；但如7月底", "逻辑转折需明确，避免误解"),
    r"蜀回合同": ("签回合同", "错别字，'蜀回'应为'签回'"),
    r"蜀单": ("签单", "错别字，'蜀单'应为'签单'"),
    r"收入报到": ("收入达到", "'报到'应为'达到'或'达'"),
    r"无信息导致盲报": ("无信息导致盲目报价", "'盲报'建议补全为'盲目报价'"),
    r"结婚团队": ("建峰团队", "语音转译错误：'建峰'被误译为'结婚'"),
}

# 重复词检测模式（两字词重复）
REPEAT_PATTERNS = [
    r"([\u4e00-\u9fff]{2,})\1",  # 中文字重复
]

# 缺少主语的句首模式（口语中常见）
MISSING_SUBJECT_PATTERNS = [
    (r"^[▸\s]*需进一步", "缺少主语，建议补充'需要客户/项目组进一步'"),
    (r"^[▸\s]*需去对齐", "缺少主语，建议补充'需项目经理去对齐'或'需对齐'"),
    (r"^[▸\s]*需.*?提供[\s，。]", "缺少主语，建议明确'需谁提供'"),
    (r"^[▸\s]*至少开到", "缺少主语"),
    (r"^[▸\s]*尽快签", "缺少主语，建议明确'需尽快签订'"),
]

# 日期格式检查
DATE_PATTERN = re.compile(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日|\d{4}\.\d{1,2}\.\d{1,2}")

# 金额格式检查（万元、元）
AMOUNT_PATTERN = re.compile(r"(\d+(?:\.\d+)?)\s*(万?元)")


# ── 文本提取 ───────────────────────────────────────────────────────

def extract_text_blocks(docx_path: str) -> list[dict]:
    """从 docx 提取结构化文本块，保留位置信息。"""
    doc = Document(docx_path)
    blocks = []
    section = "正文"
    subsection = ""

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # 识别章节/小节
        if text.startswith("【"):
            subsection = text
        elif "项目" in text and "汇报" in text and len(text) < 20:
            section = text
            subsection = ""
        elif "遗留" in text and "跟踪" in text:
            section = "遗留表"
            subsection = ""

        blocks.append({
            "type": "paragraph",
            "section": section,
            "subsection": subsection,
            "text": text,
        })

    # 提取表格
    for t_idx, table in enumerate(doc.tables):
        for r_idx, row in enumerate(table.rows):
            cells = [cell.text.strip() for cell in row.cells]
            if not any(cells):
                continue
            blocks.append({
                "type": "table",
                "section": "遗留表",
                "subsection": f"表格-{t_idx+1}-行{r_idx+1}",
                "text": " | ".join(cells),
                "cells": cells,
            })

    return blocks


# ── 校对引擎 ───────────────────────────────────────────────────────

class Proofreader:
    def __init__(self, blocks: list[dict]):
        self.blocks = blocks
        self.issues: list[Issue] = []
        self._seen_amounts: dict[str, list[str]] = {}  # 项目 -> 金额列表
        self._seen_dates: dict[str, list[str]] = {}    # 项目 -> 日期列表

    def _location(self, block: dict) -> str:
        parts = [block["section"]]
        if block.get("subsection"):
            parts.append(block["subsection"])
        return "-".join(parts)

    def _add(self, level: str, category: str, block: dict, original: str,
             suggestion: str, reason: str):
        self.issues.append(Issue(
            level=level,
            category=category,
            location=self._location(block),
            original=original,
            suggestion=suggestion,
            reason=reason,
        ))

    def check_common_errors(self):
        """检查常见错别字和用词错误。"""
        for block in self.blocks:
            text = block["text"]
            for pattern, (suggestion, reason) in COMMON_ERRORS.items():
                if re.search(pattern, text):
                    self._add("error", "错别字/用词", block,
                              text[:80] + "..." if len(text) > 80 else text,
                              suggestion, reason)

    def check_repetition(self):
        """检查重复词/句。"""
        for block in self.blocks:
            text = block["text"]
            # 检查相邻重复词（如"签完后马上能签"中的重复结构）
            match = re.search(r"([\u4e00-\u9fff]{2,})[^\u4e00-\u9fff]*\1", text)
            if match and len(match.group(1)) >= 2:
                # 过滤一些合法重复（如"会议纪要纪要"）
                REPEAT_WHITELIST = {"核减", "验收", "项目", "合同", "会议", "工作", "需求", "客户", "开发", "确认", "沟通", "汇报", "测试", "部署"}
                if match.group(1) not in REPEAT_WHITELIST:
                    self._add("error", "重复表述", block,
                              text[:80] + "..." if len(text) > 80 else text,
                              "删除重复词或重新组织语句",
                              f"检测到重复：'{match.group(1)}'")

    def check_missing_subject(self):
        """检查缺少主语的句子。"""
        for block in self.blocks:
            text = block["text"]
            for pattern, reason in MISSING_SUBJECT_PATTERNS:
                if re.search(pattern, text):
                    self._add("warning", "缺少主语", block,
                              text[:80] + "..." if len(text) > 80 else text,
                              "补充主语或调整句式", reason)

    def check_colloquial(self):
        """检查口语化表达。"""
        colloquial_words = {
            "搞": "进行/开展/处理",
            "甩": "转交/移交",
            "那边": "处/部门",
            "给到": "提交给/提供给",
            "拉通": "打通/协调",
            "对齐": "协调/确认",
            "冲量": "提升业务量",
        }
        for block in self.blocks:
            text = block["text"]
            for word, formal in colloquial_words.items():
                if word in text and word != text:
                    # 避免误报（如"对齐"在"对齐工作"中是行业术语，可以接受）
                    if word in ("搞", "甩", "给到", "冲量"):
                        self._add("warning", "口语化表达", block,
                                  text[:80] + "..." if len(text) > 80 else text,
                                  formal,
                                  f"'{word}'过于口语化，建议改为'{formal}'")

    def check_amount_consistency(self):
        """检查金额一致性。"""
        for block in self.blocks:
            text = block["text"]
            # 尝试提取项目名称（简化处理）
            project_name = block.get("subsection", "")
            matches = AMOUNT_PATTERN.findall(text)
            for amount, unit in matches:
                key = f"{project_name}-{amount}{unit}"
                if project_name:
                    self._seen_amounts.setdefault(project_name, []).append(f"{amount}{unit}")

        # 检查同一项目是否有矛盾金额
        for project, amounts in self._seen_amounts.items():
            if len(amounts) > 1:
                unique = list(dict.fromkeys(amounts))
                if len(unique) > 1:
                    self._add("question", "金额一致性", {"section": "正文", "subsection": project},
                              f"项目'{project}'中出现多个金额：{', '.join(unique)}",
                              "请确认各金额所指是否一致",
                              "同一项目出现多个不同金额，可能存在转译错误或混淆")

    def check_date_logic(self):
        """检查日期逻辑矛盾。"""
        # 收集所有日期
        all_dates = []
        for block in self.blocks:
            text = block["text"]
            matches = DATE_PATTERN.findall(text)
            for m in matches:
                if isinstance(m, tuple) and len(m) == 3:
                    try:
                        dt = datetime(int(m[0]), int(m[1]), int(m[2]))
                        all_dates.append((dt, text[:60]))
                    except (ValueError, TypeError):
                        pass

        # 检查会议日期与闭环日期矛盾
        # 简化：检查是否有"已闭环"但日期在"预计闭环"之后的情况
        for block in self.blocks:
            text = block["text"]
            if "已闭环" in text and "预计闭环" in text:
                expected = re.search(r"预计闭环[:\s]*(\d{4}\.\d{1,2}\.\d{1,2})", text)
                actual = re.search(r"实际闭环[:\s]*(\d{4}\.\d{1,2}\.\d{1,2})", text)
                if expected and actual:
                    try:
                        e = datetime.strptime(expected.group(1), "%Y.%m.%d")
                        a = datetime.strptime(actual.group(1), "%Y.%m.%d")
                        if a > e:
                            self._add("warning", "日期逻辑", block,
                                      text[:80] + "...",
                                      "请确认实际闭环时间是否确实晚于预计",
                                      "实际闭环时间晚于预计闭环时间")
                    except ValueError:
                        pass

    def check_punctuation(self):
        """检查中英文标点混用。"""
        for block in self.blocks:
            text = block["text"]
            # 检测中文内容中使用英文逗号/句号
            if re.search(r"[\u4e00-\u9fff][,]", text) or re.search(r"[,][\u4e00-\u9fff]", text):
                self._add("warning", "标点符号", block,
                          text[:80] + "..." if len(text) > 80 else text,
                          "将英文逗号改为中文逗号'，'",
                          "中文文档中应使用中文标点")
            if re.search(r"[\u4e00-\u9fff][.]", text) or re.search(r"[.][\u4e00-\u9fff]", text):
                self._add("warning", "标点符号", block,
                          text[:80] + "..." if len(text) > 80 else text,
                          "将英文句号改为中文句号'。'",
                          "中文文档中应使用中文标点")

    def check_voice_transcription(self):
        """检查语音转译常见错误。"""
        # 同音/近音字错误
        homophone_errors = {
            r"起作业": ("启动作业", "'起'与'启'同音"),
            r"起做业务": ("启动作业", "'起做'与'启动'近音"),
            r"蜀回": ("签回", "'蜀'与'签'近音"),
            r"蜀单": ("签单", "'蜀'与'签'近音"),
            r"前项[未未]": ("前期未", "'项'与'期'近音"),
            r"前向[未未]": ("前期未", "'向'与'期'近音"),
            r"开到": ("覆盖到/延长至", "'开'与'覆盖/延长'语义不同"),
            r"排进": ("排上/列入", "'进'与'上'近音"),
            r"未回": ("未回复", "口语省略"),
            r"结婚团队": ("建峰团队", "'结婚'与'建峰'近音，人名误译"),
        }
        for block in self.blocks:
            text = block["text"]
            for pattern, (suggestion, reason) in homophone_errors.items():
                if re.search(pattern, text):
                    self._add("error", "语音转译错误", block,
                              text[:80] + "..." if len(text) > 80 else text,
                              suggestion, reason)

    def run_all(self) -> list[Issue]:
        """执行所有检查。"""
        self.check_common_errors()
        self.check_repetition()
        self.check_missing_subject()
        self.check_colloquial()
        self.check_amount_consistency()
        self.check_date_logic()
        self.check_punctuation()
        self.check_voice_transcription()
        return self.issues


# ── 报告生成 ───────────────────────────────────────────────────────

def generate_report(issues: list[Issue], docx_path: str) -> str:
    """生成 Markdown 校对报告。"""
    lines = [
        "# 会议纪要精校报告",
        "",
        f"**文档**：`{docx_path}`",
        f"**检查时间**：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"**发现问题数**：{len(issues)} 条",
        "",
        "---",
        "",
    ]

    # 按级别分组
    errors = [i for i in issues if i.level == "error"]
    warnings = [i for i in issues if i.level == "warning"]
    questions = [i for i in issues if i.level == "question"]

    # 明显错误（必须修正）
    if errors:
        lines.extend([
            "## 一、明显错误（建议必须修正）",
            "",
        ])
        for idx, issue in enumerate(errors, 1):
            lines.extend([
                f"{idx}. **[{issue.category}]** `{issue.location}`",
                f"   - **原文**：{issue.original}",
                f"   - **建议**：{issue.suggestion}",
                f"   - **原因**：{issue.reason}",
                "",
            ])
    else:
        lines.extend([
            "## 一、明显错误",
            "",
            "✅ 未发现明显错误。",
            "",
        ])

    # 警告（建议优化）
    if warnings:
        lines.extend([
            "## 二、建议优化（正式场合建议调整）",
            "",
        ])
        for idx, issue in enumerate(warnings, 1):
            lines.extend([
                f"{idx}. **[{issue.category}]** `{issue.location}`",
                f"   - **原文**：{issue.original}",
                f"   - **建议**：{issue.suggestion}",
                f"   - **原因**：{issue.reason}",
                "",
            ])
    else:
        lines.extend([
            "## 二、建议优化",
            "",
            "✅ 无优化建议。",
            "",
        ])

    # 疑问（需确认）
    if questions:
        lines.extend([
            "## 三、存在疑问（需人工确认）",
            "",
        ])
        for idx, issue in enumerate(questions, 1):
            lines.extend([
                f"{idx}. **[{issue.category}]** `{issue.location}`",
                f"   - **原文**：{issue.original}",
                f"   - **建议**：{issue.suggestion}",
                f"   - **原因**：{issue.reason}",
                "",
            ])
    else:
        lines.extend([
            "## 三、存在疑问",
            "",
            "✅ 无待确认项。",
            "",
        ])

    lines.extend([
        "---",
        "",
        "## 精校说明",
        "",
        "1. **明显错误**：错别字、重复表述、语音转译错误等，建议必须修正。",
        "2. **建议优化**：口语化表达、缺少主语、标点符号等，根据正式程度决定是否调整。",
        "3. **存在疑问**：金额/日期不一致、逻辑矛盾等，需结合原始会议记录确认。",
        "",
        "> 请逐项确认后，统一修正文档。精校报告仅作为辅助，最终内容以会议原始记录为准。",
        "",
    ])

    return "\n".join(lines)


# ── 主入口 ───────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="会议纪要精校工具")
    parser.add_argument("docx_path", help="待校对的 docx 文件路径")
    parser.add_argument("-o", "--output", default="proofread_report.md",
                        help="输出报告路径（默认：proofread_report.md）")
    args = parser.parse_args()

    docx_path = Path(args.docx_path)
    if not docx_path.exists():
        print(f"错误：文件不存在 {docx_path}", file=sys.stderr)
        sys.exit(1)

    print(f"正在读取：{docx_path}")
    blocks = extract_text_blocks(str(docx_path))
    print(f"  提取到 {len(blocks)} 个文本块")

    print("正在执行校对...")
    proofreader = Proofreader(blocks)
    issues = proofreader.run_all()
    print(f"  发现 {len(issues)} 个问题")

    report = generate_report(issues, str(docx_path))
    output_path = Path(args.output)
    output_path.write_text(report, encoding="utf-8")
    print(f"校对报告已保存：{output_path}")

    # 按级别汇总
    errors = sum(1 for i in issues if i.level == "error")
    warnings = sum(1 for i in issues if i.level == "warning")
    questions = sum(1 for i in issues if i.level == "question")
    print(f"  错误：{errors}，警告：{warnings}，疑问：{questions}")

    return 0 if errors == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
