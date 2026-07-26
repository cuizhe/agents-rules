#!/usr/bin/env python3
"""月度活跃项目清单更新脚本。

用法：
    python update_active_projects.py --excel "D:\\path\\to\\清单.xlsx"

首次运行时，若 active_projects.json 不存在，会自动生成并提示确认。
无 --excel 参数时，若 active_projects.json 已存在则沿用现有数据，否则报错。
"""

import argparse
import os
import sys
from datetime import datetime

from project_index import (
    build_active_projects_from_excel,
    load_active_projects,
    save_active_projects,
)


def prompt_confirm(added_projects: list[dict]) -> bool:
    """打印新增项目索引并等待用户确认。"""
    if not added_projects:
        return True

    print("\n" + "=" * 60)
    print("新增项目索引（请确认简称和关键字是否准确）")
    print("=" * 60)
    print(f"{'序号':<4} | {'完整名称':<30} | {'简称':<20} | {'关键字'}")
    print("-" * 60)
    for idx, proj in enumerate(added_projects, 1):
        full_name = proj["full_name"]
        short_names = "、".join(proj.get("short_names", [])[:3])
        keywords = "、".join(proj.get("keywords", [])[:5])
        # 截断显示
        full_disp = full_name[:28] + "..." if len(full_name) > 30 else full_name
        print(f"{idx:<4} | {full_disp:<30} | {short_names:<20} | {keywords}")
    print("-" * 60)
    print("以上新增项目的简称和关键字由系统自动生成。")
    print("请确认是否准确。回复 '确认' 后写入文件，或回复项目序号+修改建议。")
    print("=" * 60)

    while True:
        try:
            user_input = input("\n请输入 '确认' 或项目序号+修改建议: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n[取消] 用户取消操作，未写入文件。")
            return False

        if user_input == "确认":
            return True
        elif user_input.isdigit():
            idx = int(user_input)
            if 1 <= idx <= len(added_projects):
                print(f"项目 {idx}: {added_projects[idx - 1]['full_name']}")
                print("请输入修改后的 short_names（用逗号分隔）:")
                try:
                    sn = input().strip()
                    if sn:
                        added_projects[idx - 1]["short_names"] = [s.strip() for s in sn.split(",") if s.strip()]
                except (EOFError, KeyboardInterrupt):
                    pass
                print("请输入修改后的 keywords（用逗号分隔）:")
                try:
                    kw = input().strip()
                    if kw:
                        added_projects[idx - 1]["keywords"] = [k.strip() for k in kw.split(",") if k.strip()]
                except (EOFError, KeyboardInterrupt):
                    pass
                print("已更新，继续确认其他项目或输入 '确认' 完成。")
            else:
                print(f"序号 {idx} 超出范围，请重新输入。")
        else:
            print("输入格式不正确，请直接输入 '确认' 或项目序号。")


def main():
    parser = argparse.ArgumentParser(description="更新月度活跃项目清单")
    parser.add_argument(
        "--excel",
        type=str,
        default=None,
        help="Excel 排产清单文件路径（如 D:\\path\\to\\清单.xlsx）",
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="自动确认，不进入交互式确认环节（适合自动化场景）",
    )
    parser.add_argument(
        "--output",
        type=str,
        default="active_projects.json",
        help="输出 JSON 文件路径（默认: active_projects.json）",
    )
    args = parser.parse_args()

    output_path = args.output
    existing_data = load_active_projects(output_path)

    if not args.excel:
        if existing_data.get("last_updated"):
            print("[信息] 未提供 --excel 参数，沿用现有活跃项目数据。")
            print(f"[信息] 当前活跃项目清单最后更新于: {existing_data['last_updated']}")
            print(f"[信息] 项目总数: {len(existing_data.get('projects', []))}")
            active_count = sum(1 for p in existing_data.get("projects", []) if p.get("status") == "active")
            print(f"[信息] 其中 active 状态: {active_count}")
            return
        else:
            print("[错误] 未检测到现有活跃项目清单，且未提供 --excel 参数。")
            print("[提示] 请提供 Excel 文件路径进行初始化:")
            print("  python update_active_projects.py --excel \"D:\\\\path\\\\to\\\\清单.xlsx\"")
            sys.exit(1)

    if not os.path.exists(args.excel):
        print(f"[错误] 文件不存在: {args.excel}")
        sys.exit(1)

    print(f"[读取] 正在读取 Excel: {args.excel}")
    merged_data, added_projects = build_active_projects_from_excel(args.excel, existing_data)

    today_str = datetime.now().strftime("%Y-%m-%d")
    merged_data["last_updated"] = today_str

    # 为新增项目设置 added_date
    for proj in added_projects:
        if proj.get("added_date") is None:
            proj["added_date"] = today_str

    if added_projects:
        print(f"[信息] 本月新增项目: {len(added_projects)} 个")
        if not args.yes and not prompt_confirm(added_projects):
            print("[取消] 用户未确认，未写入文件。")
            sys.exit(1)
    else:
        print("[信息] 本月无新增项目，仅更新现有项目状态。")

    save_active_projects(merged_data, output_path)
    print(f"[成功] 活跃项目清单已更新: {output_path}")
    active_count = sum(1 for p in merged_data.get("projects", []) if p.get("status") == "active")
    print(f"[信息] 项目总数: {len(merged_data.get('projects', []))}, active: {active_count}")


if __name__ == "__main__":
    main()
