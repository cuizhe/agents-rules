"""项目索引管理模块：活跃项目清单加载、简称解析、元数据生成。"""

import json
import os
import re
from pathlib import Path


# 通用停用词（用于 keywords 提取）
STOP_WORDS = {
    "项目", "合同", "维保", "服务", "实施", "外包", "开发", "平台", "系统",
    "年度", "年", "月", "日", "公司", "集团", "有限", "股份", "科技",
    "进行", "完成", "开展", "工作", "相关", "以及", "及其", "等",
    "一期", "二期", "三期", "四期", "五期", "六期", "七期", "八期", "九期", "十期",
    "第", "个", "次", "项", "各", "类", "中", "的", "了", "与", "及",
    "", " ", "\t", "\n",
}

# 通用后缀（用于 short_names 生成时去除）
GENERIC_SUFFIXES = [
    "项目", "FP项目", "TM项目", "合同", "维保", "服务", "实施", "外包",
    "开发", "平台", "系统", "定制", "交付", "运维",
]


def load_active_projects(path: str = "active_projects.json") -> dict:
    """加载活跃项目清单。返回完整字典（含 version、last_updated、projects）。"""
    if not os.path.exists(path):
        return {"version": "1.0", "last_updated": None, "projects": []}
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)


def save_active_projects(data: dict, path: str = "active_projects.json") -> None:
    """保存活跃项目清单。"""
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)


def get_active_project_names(active_projects: dict) -> set[str]:
    """获取所有 status == 'active' 的 full_name 集合。"""
    return {
        p["full_name"] for p in active_projects.get("projects", [])
        if p.get("status") == "active"
    }


def get_active_project_records(task_db: list[dict], active_projects: dict) -> list[dict]:
    """从历史库中筛选出属于活跃项目的记录。"""
    active_names = get_active_project_names(active_projects)
    return [r for r in task_db if r.get("项目名称") in active_names]


def resolve_project_name(user_input: str, active_projects: dict) -> str | None:
    """
    将用户输入的项目简称/关键字映射到 full_name。

    匹配优先级（从高到低）：
    1. full_name 精确匹配（去除首尾空格后）
    2. short_names 列表中包含 user_input
    3. keywords 列表中包含 user_input（且 user_input 长度 >= 2）
    4. full_name 子串包含 user_input（且 user_input 长度 >= 2）

    若多个项目匹配，返回匹配度最高的。
    """
    user_input = user_input.strip()
    if not user_input or not active_projects:
        return None

    best_match = None
    best_score = -1

    for p in active_projects.get("projects", []):
        if p.get("status") != "active":
            continue

        full_name = p.get("full_name", "")
        short_names = p.get("short_names", [])
        keywords = p.get("keywords", [])

        score = -1
        if user_input == full_name.strip():
            score = 4
        elif user_input in short_names:
            score = 3
        elif len(user_input) >= 2 and user_input in keywords:
            score = 2
        elif len(user_input) >= 2 and user_input in full_name:
            score = 1

        if score > best_score:
            best_score = score
            best_match = full_name

    return best_match


def _remove_year_prefix(name: str) -> str:
    """去除年份前缀，如 '2026年'、'2024-2025年'。"""
    s = re.sub(r"\d{4}[-~]\d{4}年?", "", name)
    s = re.sub(r"\d{4}年?", "", s)
    return s.strip(" -")


def _remove_generic_suffixes(name: str) -> str:
    """去除通用后缀，如 '项目'、'FP项目'、'合同'。"""
    # 按长度降序，避免先去除短后缀导致长后缀残留
    for suffix in sorted(GENERIC_SUFFIXES, key=len, reverse=True):
        if name.endswith(suffix):
            name = name[: -len(suffix)].strip(" -")
    return name


def _extract_tokens(name: str) -> list[str]:
    """按连接符拆分名称，提取有效 token（长度 >= 2，非纯数字）。"""
    parts = re.split(r"[-_/\s]", name)
    tokens = []
    for p in parts:
        p = p.strip()
        if p and len(p) >= 2 and not re.fullmatch(r"\d+", p):
            tokens.append(p)
    return tokens


def _is_location(token: str) -> bool:
    """简单启发式判断 token 是否为地名。"""
    location_suffixes = ["省", "市", "县", "区", "州", "盟", "旗", "镇", "乡"]
    return any(token.endswith(s) for s in location_suffixes)


def _is_generic_token(token: str) -> bool:
    """判断 token 是否为无意义通用词（如 '公司'、'合同'）。"""
    generic = {"公司", "集团", "有限", "股份", "科技", "信息", "技术", "软件",
               "合同", "维保", "服务", "实施", "外包", "开发", "平台", "系统",
               "项目", "FP", "TM", "年度", "进行", "完成", "开展", "工作"}
    return token in generic


def generate_project_metadata(full_name: str) -> dict:
    """
    根据项目全称自动生成 short_names 和 keywords。

    返回 {"short_names": [...], "keywords": [...]}
    """
    if not full_name:
        return {"short_names": [], "keywords": []}

    # 1. 去除年份和通用后缀
    core = _remove_year_prefix(full_name)
    core = _remove_generic_suffixes(core)
    core = core.strip(" -_/\\s")

    # 2. 提取 token
    tokens = _extract_tokens(core)
    if not tokens:
        # 若去除后缀后无 token，至少保留原始名称（去除年份后）
        tokens = [full_name.strip()]

    # 3. 识别地名、机构名、业务领域
    locations = [t for t in tokens if _is_location(t)]
    institutions = [t for t in tokens if not _is_location(t) and not _is_generic_token(t)]
    # 业务领域：所有非地名、非通用词、非纯数字的 token
    business = [t for t in tokens if not _is_location(t) and not _is_generic_token(t)]

    short_names = set()

    # 3.1 完整核心名称（去除年份和通用后缀）
    if core and len(core) >= 2:
        short_names.add(core)

    # 3.2 地名 + 机构名
    for loc in locations:
        for inst in institutions:
            combined = loc + inst
            if len(combined) >= 2:
                short_names.add(combined)

    # 3.3 机构名（去地名）
    for inst in institutions:
        if len(inst) >= 2:
            short_names.add(inst)

    # 3.4 地名 + 业务领域（取前两个业务 token）
    for loc in locations:
        for biz in business[:2]:
            combined = loc + biz
            if len(combined) >= 2:
                short_names.add(combined)

    # 3.5 原始名称本身（如果去除年份后不同）
    no_year = _remove_year_prefix(full_name).strip()
    if no_year and no_year != full_name and len(no_year) >= 2:
        short_names.add(no_year)

    # 3.6 取前两个 token 组合（如果长度 >= 4）
    if len(tokens) >= 2:
        combo = tokens[0] + tokens[1]
        if len(combo) >= 4:
            short_names.add(combo)

    # 去重并过滤
    short_names = sorted([s for s in short_names if len(s) >= 2 and s != full_name])
    # 若 short_names 为空，兜底为 no_year
    if not short_names and no_year:
        short_names = [no_year]

    # 4. 生成 keywords
    keywords = set()
    for t in tokens:
        if t and len(t) >= 2 and not _is_generic_token(t) and t not in STOP_WORDS:
            keywords.add(t)
    # 加入 short_names 中的词作为 keyword
    for sn in short_names:
        if len(sn) >= 2:
            keywords.add(sn)
    # 加入地名（即使地名只有 2 字，也作为 keyword）
    for loc in locations:
        keywords.add(loc)

    keywords = sorted([k for k in keywords if len(k) >= 2])

    return {
        "short_names": short_names,
        "keywords": keywords,
    }


def build_active_projects_from_excel(excel_path: str, existing_data: dict | None = None) -> tuple[dict, list[dict]]:
    """
    读取 Excel 排产清单，构建/更新 active_projects 数据。
    返回 (merged_data, added_projects_list)。
    需要 openpyxl。
    """
    try:
        import openpyxl
    except ImportError:
        raise ImportError("需要安装 openpyxl: pip install openpyxl")

    wb = openpyxl.load_workbook(excel_path, read_only=True)

    # 过滤匹配 "X月排产项目清单" 的 sheet
    month_sheets = []
    for name in wb.sheetnames:
        m = re.match(r"(\d+)月.*清单", name.strip())
        if m:
            month_sheets.append((int(m.group(1)), name))

    if not month_sheets:
        raise ValueError(f"未找到符合 'X月排产项目清单' 格式的 sheet，现有 sheet: {wb.sheetnames}")

    # 取月份数字最大的 sheet
    month_sheets.sort(key=lambda x: x[0], reverse=True)
    target_sheet = month_sheets[0][1]
    ws = wb[target_sheet]

    # 查找表头行（前5行中，包含"项目名称"的行）
    header_row_idx = None
    headers = []
    for row_idx in range(1, 6):
        row = [cell.value for cell in ws[row_idx]]
        if any("项目名称" in str(v) for v in row if v):
            header_row_idx = row_idx
            headers = [str(v).strip() if v else None for v in row]
            break

    if header_row_idx is None:
        raise ValueError(f"在 sheet '{target_sheet}' 的前5行中未找到 '项目名称' 列")

    # 定位关键列索引
    col_idx = {}
    for needed in ["项目名称", "项目状态", "是否排产"]:
        for i, h in enumerate(headers):
            if h and needed in h:
                col_idx[needed] = i
                break
    if "项目名称" not in col_idx:
        raise ValueError("未找到 '项目名称' 列")

    # 读取数据行
    new_projects = []
    for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
        if not row or len(row) <= col_idx["项目名称"]:
            continue
        full_name = str(row[col_idx["项目名称"]]).strip() if row[col_idx["项目名称"]] else ""
        if not full_name:
            continue

        # 状态判断
        status = "inactive"
        proj_status = ""
        is_scheduled = ""
        if "项目状态" in col_idx and row[col_idx["项目状态"]]:
            proj_status = str(row[col_idx["项目状态"]]).strip()
        if "是否排产" in col_idx and row[col_idx["是否排产"]]:
            is_scheduled = str(row[col_idx["是否排产"]]).strip()

        # 仅当 "项目状态" == "进行中" 且 "是否排产" == "是" 时标记 active
        if proj_status == "进行中" and is_scheduled == "是":
            status = "active"

        new_projects.append({
            "full_name": full_name,
            "status": status,
            "proj_status": proj_status,
            "is_scheduled": is_scheduled,
        })

    # 合并现有数据
    if existing_data is None:
        existing_data = {"version": "1.0", "last_updated": None, "projects": []}

    existing_map = {p["full_name"]: p for p in existing_data.get("projects", [])}
    merged_projects = []
    added_projects = []  # 用于首次确认

    for np in new_projects:
        full_name = np["full_name"]
        if full_name in existing_map:
            # 更新状态
            existing_map[full_name]["status"] = np["status"]
            merged_projects.append(existing_map[full_name])
        else:
            # 新增项目，生成 metadata
            metadata = generate_project_metadata(full_name)
            proj_entry = {
                "full_name": full_name,
                "short_names": metadata["short_names"],
                "keywords": metadata["keywords"],
                "status": np["status"],
                "added_date": None,  # 由调用方填入
            }
            merged_projects.append(proj_entry)
            added_projects.append(proj_entry)

    # 保留 existing 中但不在 new 中的项目（标记为 inactive）
    new_names = {p["full_name"] for p in new_projects}
    for full_name, ep in existing_map.items():
        if full_name not in new_names:
            ep["status"] = "inactive"
            if ep not in merged_projects:
                merged_projects.append(ep)

    existing_data["projects"] = merged_projects
    return existing_data, added_projects
