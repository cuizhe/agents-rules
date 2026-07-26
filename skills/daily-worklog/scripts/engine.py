import json
import math
import re
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path


ASSETS_DIR = Path(__file__).parent.parent / "assets"


def load_task_db(path: str = None) -> list[dict]:
    if path is None:
        path = str(ASSETS_DIR / "task_db.json")
    with open(path, "r", encoding="utf-8") as fh:
        raw = json.load(fh)
    # 仅保留填报流程所需的字段，剔除审批状态等系统生成字段
    needed = {
        "申报类型", "申报日期", "项目名称", "工作项1",
        "工作项2", "工作项3", "投入时间(H)", "工作日志", "search_text",
    }
    cleaned = []
    for r in raw:
        cleaned.append({k: v for k, v in r.items() if k in needed})
    return cleaned


def match_tasks(keywords: list[str], task_db: list[dict], limit: int = 15) -> list[dict]:
    """
    多关键词 OR 匹配：任一关键词在 search_text 中出现即入选。
    返回匹配记录列表，按匹配关键词数量降序排列，最多 limit 条。
    """
    results = []
    for record in task_db:
        search_text = record.get("search_text", "")
        if not search_text:
            continue
        match_count = 0
        for kw in keywords:
            kw = kw.strip().lower()
            if kw and kw in search_text.lower():
                match_count += 1
        if match_count > 0:
            record_copy = dict(record)
            record_copy["_match_score"] = match_count
            results.append(record_copy)
    # 按匹配关键词数量降序，限制上限
    results.sort(key=lambda x: x["_match_score"], reverse=True)
    if len(results) > limit:
        results = results[:limit]
    return results


def round_to_half(value: float) -> float:
    return max(0.5, math.ceil(value * 2) / 2)


def aggregate_tasks(records: list[dict]) -> list[dict]:
    """
    按 申报类型 + 项目名称 + 工作项1 分组聚合。
    - 日志：去重后拼接，截断至500字
    - 工时：组内平均值，先向上取整到0.5倍数
    - 总工和超8小时：按比例压缩，再按0.5取整
    """
    groups = defaultdict(list)
    for r in records:
        # 按叶子工作项聚合，避免不同叶子节点被合并到同一行
        工作项3 = r.get("工作项3", "")
        工作项2 = r.get("工作项2", "")
        工作项1 = r.get("工作项1", "")
        工作项叶子 = 工作项3 or 工作项2 or 工作项1
        key = (
            r.get("申报类型", ""),
            r.get("项目名称", ""),
            工作项叶子,
        )
        groups[key].append(r)

    items = []
    for (申报类型, 项目名称, 工作项叶子), group in groups.items():
        # 日志去重拼接
        logs = []
        for g in group:
            log = g.get("工作日志", "")
            if log and log not in logs:
                logs.append(str(log))
        merged_log = "；".join(logs)
        if len(merged_log) > 500:
            merged_log = merged_log[:500]

        # 工时取平均
        hours = [g.get("投入时间(H)", 0) for g in group]
        avg_hour = sum(hours) / len(hours) if hours else 0
        avg_hour = round_to_half(avg_hour)

        # 取组内第一条的层级信息
        工作项3 = group[0].get("工作项3", "")
        工作项2 = group[0].get("工作项2", "")
        工作项1 = group[0].get("工作项1", "")

        items.append({
            "申报类型": 申报类型,
            "项目名称": 项目名称,
            "工作项": 工作项叶子,
            "工作项1": 工作项1,
            "工作项2": 工作项2,
            "工作项3": 工作项3,
            "工时": avg_hour,
            "日志": merged_log,
        })

    return items


def adjust_to_eight(items: list[dict], reserved_hours: float = 0.0) -> list[dict]:
    """
    将 items 的总工时严格调整到 8.0 小时（扣除预留工时后）。
    支持双向修正：超过则压缩，不足则放大。
    reserved_hours: 固定预留的工时（如每日必出的公共工时），这部分不参与调整。
    """
    target = 8.0 - reserved_hours
    total = sum(item["工时"] for item in items)
    if total > 0 and target > 0 and total != target:
        # 受保护项（用户指定、固定）不参与缩放
        protected_hours = sum(item["工时"] for item in items if item.get("_user_specified") or item.get("_fixed"))
        adjustable_total = total - protected_hours
        adjustable_target = target - protected_hours

        if adjustable_total > 0 and adjustable_target > 0:
            factor = adjustable_target / adjustable_total
            for item in items:
                if not item.get("_user_specified") and not item.get("_fixed"):
                    item["工时"] = round_to_half(item["工时"] * factor)
        else:
            # 无可调整项，直接按比例
            factor = target / total
            for item in items:
                item["工时"] = round_to_half(item["工时"] * factor)
        # 二次修正：严格对齐 target
        while sum(item["工时"] for item in items) > target:
            # 优先压缩非用户指定、非固定项
            adjustable = [x for x in items if x["工时"] >= 0.5 and not x.get("_user_specified") and not x.get("_fixed")]
            if adjustable:
                max_item = max(adjustable, key=lambda x: x["工时"])
            else:
                max_item = max(items, key=lambda x: x["工时"])
            if max_item["工时"] >= 0.5:
                max_item["工时"] -= 0.5
        while sum(item["工时"] for item in items) < target:
            max_item = max(items, key=lambda x: x["工时"])
            max_item["工时"] += 0.5
    return items


def merge_project_rows(items: list[dict], project_log_map: dict | None = None, project_hour_map: dict | None = None) -> list[dict]:
    """
    按项目名称合并项目工时：同一项目只保留一行。
    - 工作项：选择出现次数最多的叶子工作项（用于下拉框回填）
    - 工时：优先使用用户指定值（project_hour_map），否则组内相加后按0.5取整
    - 日志：若用户在 project_log_map 中指定了该项目，直接用指定内容；
            否则拼接去重，多个工作项时前缀注明
    """
    from collections import Counter

    project_groups: dict[str, list[dict]] = {}
    other_items = []

    for item in items:
        if item.get("申报类型") == "项目工时" and item.get("项目名称"):
            proj = item["项目名称"]
            project_groups.setdefault(proj, []).append(item)
        else:
            other_items.append(item)

    merged = []
    for proj, group in project_groups.items():
        # 选择频率最高的工作项作为下拉框代表
        wi_counter = Counter(g.get("工作项", "") for g in group if g.get("工作项"))
        representative_wi = wi_counter.most_common(1)[0][0] if wi_counter else group[0].get("工作项", "")
        rep_item = next((g for g in group if g.get("工作项") == representative_wi), group[0])

        # 工时：优先使用用户指定值（project_hour_map），否则组内相加后按0.5取整
        user_hour = None
        if project_hour_map:
            for proj_kw, hour in project_hour_map.items():
                if proj_kw in proj or _has_signature_overlap(proj_kw, proj):
                    user_hour = hour
                    break
        if user_hour is not None:
            total_hours = round_to_half(user_hour)
        else:
            total_hours = sum(g.get("工时", 0) for g in group)
            total_hours = round_to_half(total_hours)

        # 判断该项目是否在用户指定日志中（支持子串匹配及特征重叠）
        user_log = None
        if project_log_map:
            for proj_kw, log_content in project_log_map.items():
                if proj_kw in proj or _has_signature_overlap(proj_kw, proj):
                    user_log = log_content
                    break

        if user_log is not None:
            merged_log = user_log
        else:
            # 收集所有去重工作项
            all_workitems = []
            for g in group:
                wi = g.get("工作项", "")
                if wi and wi not in all_workitems:
                    all_workitems.append(wi)

            # 日志合并去重（不再添加内部工作项前缀，该前缀仅供系统内部聚合使用）
            all_logs = []
            for g in group:
                log = g.get("日志", "")
                if log and log not in all_logs:
                    all_logs.append(log)
            merged_log = "；".join(all_logs)

            if len(merged_log) > 500:
                merged_log = merged_log[:500]

        merged.append({
            "申报类型": "项目工时",
            "项目名称": proj,
            "工作项": representative_wi,
            "工作项1": rep_item.get("工作项1", representative_wi),
            "工作项2": rep_item.get("工作项2", ""),
            "工作项3": rep_item.get("工作项3", ""),
            "工时": total_hours,
            "日志": merged_log,
        })

    return other_items + merged


def is_project_specific_keyword(keywords: list[str], task_db: list[dict]) -> bool:
    """
    判断关键词是否为具体项目名称（在 task_db 的 '项目名称' 字段中出现过）。
    用于过滤掉公共工时记录，避免项目关键词误匹配合共工时的工作日志。
    """
    for record in task_db:
        project_name = _safe_str(record.get("项目名称", ""))
        if not project_name:
            continue
        for kw in keywords:
            kw = kw.strip().lower()
            if kw and kw in project_name.lower():
                return True
    return False


def is_quality_planning_related(record: dict) -> bool:
    """判断记录的工作项（1/2/3）中是否包含'质量策划'。"""
    for key in ("工作项1", "工作项2", "工作项3"):
        text = _safe_str(record.get(key, ""))
        if "质量策划" in text:
            return True
    return False


def should_allow_quality_planning(keywords: list[str], task_db: list[dict]) -> bool:
    """
    质量策划相关记录仅在关键词同时包含'质量策划'和具体项目名称时才允许出现。
    """
    has_qp = any("质量策划" in kw for kw in keywords)
    if not has_qp:
        return False
    return is_project_specific_keyword(keywords, task_db)


def get_projects_with_quality_planning(task_db: list[dict]) -> set[str]:
    """
    扫描历史任务库，返回已经填报过'制定与辅导项目质量策划'的项目名称集合。
    用于单项目质量策划终身去重。
    """
    projects = set()
    for r in task_db:
        if is_quality_planning_related(r) and r.get("申报类型") == "项目工时":
            pname = r.get("项目名称", "")
            if pname:
                projects.add(pname)
    return projects


def enforce_daily_quality_planning_limit(items: list[dict], limit: int = 1) -> list[dict]:
    """
    硬性约束：同一天内，最多只允许 `limit` 个项目出现质量策划类工作项。
    超出部分的工作项会被替换为'项目成本监控'，并打印调整日志。
    """
    qp_count = 0
    logs = []
    for item in items:
        if item.get("申报类型") == "项目工时" and is_quality_planning_related(item):
            qp_count += 1
            if qp_count > limit:
                old_wi = item.get("工作项", "")
                item["工作项"] = "项目成本监控"
                item["工作项1"] = "项目成本监控"
                item["工作项2"] = ""
                item["工作项3"] = ""
                logs.append(
                    f"[质量策划限制] 项目 '{item.get('项目名称', '')}' 超出单日质量策划上限({limit})，"
                    f"工作项从 '{old_wi}' 调整为 '项目成本监控'"
                )
    if logs:
        print("\n[质量策划单日限制]")
        for log in logs:
            print(log)
    return items


def replace_month_in_logs(records: list[dict], date_str: str) -> list[str]:
    """
    将日志中明显偏离合理时间范围的月份数字替换为当前月数字。
    允许范围：
    - 默认: 上月、当月
    - 月末(>=20号): 额外允许下月（月末常做下月计划）
    返回替换操作日志列表。
    """
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    current_month = dt.month
    prev_month = 12 if current_month == 1 else current_month - 1
    next_month = 1 if current_month == 12 else current_month + 1

    allowed_months = {current_month, prev_month}
    if dt.day >= 20:
        allowed_months.add(next_month)

    month_pattern = re.compile(r"(\d{1,2})(月|月份)")

    logs = []
    for r in records:
        log_text = r.get("日志", "")

        def _replace(match):
            month_num = int(match.group(1))
            suffix = match.group(2)
            if 1 <= month_num <= 12 and month_num not in allowed_months:
                return f"{current_month}{suffix}"
            return match.group(0)

        new_log = month_pattern.sub(_replace, log_text)

        # 清洗内部定制规则术语，避免写入正式提交日志
        if "活跃项目" in new_log:
            new_log = re.sub(r"活跃项目\s*", "", new_log)
            # 清理可能出现的多余空格或前缀顿号
            new_log = re.sub(r"^[、，,\s]+", "", new_log)
            new_log = re.sub(r"\s+", " ", new_log).strip()

        # 清洗内部工作项聚合前缀，避免写入正式提交日志
        if "【工作项：" in new_log:
            new_log = re.sub(r"【工作项：[^】]+】\s*", "", new_log)
            new_log = re.sub(r"\s+", " ", new_log).strip()

        if new_log != log_text:
            r["日志"] = new_log
            logs.append(f"[替换] 日志月份更新为{current_month}月: {r.get('工作项', '')}")

    return logs


def filter_by_date_rules(records: list[dict], date_str: str, holidays: list[str] | None = None) -> tuple[list[dict], list[str]]:
    """
    按日期业务规则过滤记录：
    - 361评估：仅周五或节假日前一天可填报
    - 审计（流程遵从度审计）：仅每月20日及以后（下旬）可填报
    - 排产：仅每月1-10日（上旬）可填报
    返回 (过滤后记录列表, 过滤原因日志列表)
    """
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    is_friday = dt.weekday() == 4
    holiday_set = set(holidays) if holidays else set()
    next_day = (dt + timedelta(days=1)).strftime("%Y-%m-%d")
    is_day_before_holiday = next_day in holiday_set
    is_late_month = dt.day >= 20
    is_early_month = dt.day <= 10

    kept, logs = [], []
    for r in records:
        workitem = r.get("工作项", "")
        log_text = r.get("日志", "")
        combined = f"{workitem} {log_text}"

        if "361评估" in combined:
            if not (is_friday or is_day_before_holiday):
                logs.append(f"[过滤] 361评估非周五/节假日前一天: {workitem}")
                continue

        # 精确匹配"审计"，避免"评审计划"等词中的子串误匹配
        if re.search(r'(?<![一-龥])审计(?![一-龥])', combined):
            if not is_late_month:
                logs.append(f"[过滤] 审计非每月下旬(>=20日): {workitem}")
                continue

        if "排产" in combined:
            if not is_early_month:
                logs.append(f"[过滤] 排产非每月上旬(<=10日): {workitem}")
                continue

        kept.append(r)

    return kept, logs


# ---------------------------------------------------------------------------
# 加权随机与手动任务支持
# ---------------------------------------------------------------------------

import random  # noqa: E402


def _safe_str(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        import math
        if math.isnan(value):
            return ""
    s = str(value).strip()
    if s.lower() in ("nan", "none", "null"):
        return ""
    return s


def _extract_feature_signatures(name: str) -> set[str]:
    """提取项目名中的特征子串，用于去重匹配（如电渠↔电子渠道、H5↔H5）。"""
    if not name:
        return set()
    no_year = re.sub(r"\d{4}[-~]\d{4}年?", "", name)
    no_year = re.sub(r"\d{4}年?", "", no_year)
    common_words = {"中国移动", "中国电信", "中国联通", "公司", "项目", "合同",
                    "服务", "平台", "系统", "实施", "开发", "外包", "运维",
                    "交付", "定制", "FP", "采购", "支撑", "属地化", "营销",
                    "深圳", "广州", "广东", "年", "月"}
    for w in common_words:
        no_year = no_year.replace(w, "")
    parts = re.split(r"[-_/\s]", no_year)
    sigs = set()
    for p in parts:
        p = p.strip()
        if not p or p.isdigit():
            continue
        # 提取英文/数字段和连续中文字段（如电子渠道H5 → 电子渠道, H5）
        segments = re.findall(r"[a-zA-Z0-9]+|[一-龥]{2,}", p)
        for seg in segments:
            if len(seg) >= 2:
                sigs.add(seg)
    return sigs


def _has_signature_overlap(name_a: str, name_b: str) -> bool:
    """检查两个项目名是否有特征子串重叠（含子串包含关系）。"""
    sigs_a = _extract_feature_signatures(name_a)
    sigs_b = _extract_feature_signatures(name_b)
    if not sigs_a or not sigs_b:
        return False
    if sigs_a & sigs_b:
        return True
    for sa in sigs_a:
        for sb in sigs_b:
            if sa in sb or sb in sa:
                return True
    return False


def get_max_task_id(path: str = None) -> int:
    if path is None:
        path = str(ASSETS_DIR / "task_db.json")
    with open(path, "r", encoding="utf-8") as fh:
        raw = json.load(fh)
    max_id = 0
    for r in raw:
        tid = r.get("工时ID", 0)
        try:
            tid = int(tid)
        except Exception:
            tid = 0
        if tid > max_id:
            max_id = tid
    return max_id


def infer_workitem_levels(task_db: list[dict], leaf_name: str) -> tuple[str, str, str]:
    """
    根据叶子工作项名称，从历史记录中反推工作项1/2/3层级。
    找不到则返回 (leaf_name, "", "")。
    """
    leaf_name = leaf_name.strip()
    for r in task_db:
        w3 = _safe_str(r.get("工作项3", ""))
        w2 = _safe_str(r.get("工作项2", ""))
        w1 = _safe_str(r.get("工作项1", ""))
        if w3 and w3 == leaf_name:
            return w1, w2, w3
        if not w3 and w2 and w2 == leaf_name:
            return w1, w2, ""
        if not w3 and not w2 and w1 and w1 == leaf_name:
            return w1, "", ""
    return leaf_name, "", ""


def _get_leaf_workitem(record: dict) -> str:
    return _safe_str(record.get("工作项3", "") or record.get("工作项2", "") or record.get("工作项1", "") or record.get("工作项", ""))


def _is_month_end(date_str: str) -> bool:
    """判断是否为每月25日及以后"""
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    return dt.day >= 25


def _has_month_end_audit_or_metric(task_db: list[dict], project_name: str, date_str: str) -> bool:
    """检查某项目在本月是否已有审计或度量分析记录"""
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    current_month = dt.strftime("%Y-%m")
    for r in task_db:
        record_date = r.get("申报日期", "")
        if not record_date.startswith(current_month):
            continue
        if r.get("项目名称") != project_name:
            continue
        leaf = _get_leaf_workitem(r)
        if leaf in ("审计", "度量分析"):
            return True
    return False


def _generate_month_end_tasks(
    existing_items: list[dict],
    active_projects: list[dict],
    date_str: str,
    task_db: list[dict],
) -> list[dict]:
    """
    每月25日及以后，从活跃项目中随机选1-2个（本月未执行过审计/度量的），
    生成审计（1.5H）或度量分析（1.0H）项目工时。
    """
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    if dt.day < 25:
        return []

    # 收集已有项目名（避免当天重复）
    existing_projects = set()
    for item in existing_items:
        if item.get("申报类型") == "项目工时":
            pname = item.get("项目名称", "")
            if pname:
                existing_projects.add(pname)

    # 获取活跃项目
    from project_index import get_active_project_names
    active_names = get_active_project_names({"projects": active_projects}) if active_projects else set()

    # 过滤掉已有项目，以及本月已执行过审计/度量的项目
    available = []
    for name in active_names:
        if name in existing_projects:
            continue
        if _has_month_end_audit_or_metric(task_db, name, date_str):
            continue
        available.append(name)

    if not available:
        return []

    # 随机选1-2个
    count = random.choice([1, 2])
    selected = random.sample(available, min(count, len(available)))

    # 工作项模板
    templates = [
        ("审计", 1.5, "流程遵从度审计，输出审计报告及问题跟踪清单"),
        ("度量分析", 1.0, "月度过程度量数据统计分析，输出度量分析报告"),
    ]

    result = []
    for proj in selected:
        wi_name, hours, log = random.choice(templates)
        w1, w2, w3 = infer_workitem_levels(task_db, wi_name)
        result.append({
            "申报类型": "项目工时",
            "项目名称": proj,
            "工作项": w3 or w2 or w1,
            "工作项1": w1,
            "工作项2": w2,
            "工作项3": w3,
            "工时": hours,
            "日志": log,
        })
        print(f"[月末例行] 已追加项目工时：{proj} | {wi_name} | {hours}H")

    return result


def _time_weight(record_date_str: str, ref_date: datetime) -> float:
    try:
        record_date = datetime.strptime(record_date_str, "%Y-%m-%d")
        days_diff = (ref_date - record_date).days
    except Exception:
        return 0.5
    if days_diff <= 30:
        return 2.0
    elif days_diff <= 60:
        return 1.5
    elif days_diff <= 90:
        return 1.0
    else:
        return 0.5


def weighted_random_tasks(
    task_db: list[dict],
    total_count: int,
    project_count: int,
    public_count: int,
    ref_date_str: str | None = None,
    active_projects: list[dict] | None = None,
) -> list[dict]:
    """
    按权重随机抽取指定数量的记录，并按项目/公共比例约束分组。
    返回标准 item 格式列表（与 aggregate_tasks 输出一致）。
    """
    ref_date = datetime.strptime(ref_date_str, "%Y-%m-%d") if ref_date_str else datetime.now()

    # 频率权重：统计每个工作项叶子的出现次数
    leaf_counts = defaultdict(int)
    for r in task_db:
        leaf = _get_leaf_workitem(r)
        if leaf:
            leaf_counts[leaf] += 1

    def _calc_weight(r: dict) -> float:
        leaf = _get_leaf_workitem(r)
        tw = _time_weight(r.get("申报日期", ""), ref_date)
        fw = 1.0 + math.log1p(leaf_counts.get(leaf, 1))
        return tw * fw

    # 分池
    project_pool = [r for r in task_db if r.get("申报类型") == "项目工时"]
    public_pool = [r for r in task_db if r.get("申报类型") != "项目工时"]

    # 若提供了活跃项目清单，优先从活跃项目池中抽取项目工时
    active_project_pool = None
    if active_projects is not None:
        from project_index import get_active_project_names
        active_names = get_active_project_names({"projects": active_projects})
        active_project_pool = [r for r in project_pool if r.get("项目名称") in active_names]

    def _weighted_sample(pool: list[dict], n: int, mark_inactive: bool = False) -> list[dict]:
        if not pool or n <= 0:
            return []
        pool = list(pool)
        result = []
        for _ in range(n):
            if not pool:
                break
            weights = [_calc_weight(r) for r in pool]
            total_w = sum(weights)
            if total_w <= 0:
                pick = random.choice(pool)
            else:
                pick = random.choices(pool, weights=weights, k=1)[0]
            if mark_inactive:
                pick["_from_inactive"] = True
            result.append(pick)
            pool.remove(pick)
        return result

    # 项目工时抽取：优先活跃池，不足降级
    if active_project_pool is not None:
        project_samples = _weighted_sample(active_project_pool, project_count)
        if len(project_samples) < project_count:
            # 从剩余项目池中补充，标记降级
            remaining_project_pool = [r for r in project_pool if r not in project_samples]
            extra = _weighted_sample(remaining_project_pool, project_count - len(project_samples), mark_inactive=True)
            project_samples.extend(extra)
    else:
        project_samples = _weighted_sample(project_pool, project_count)

    public_samples = _weighted_sample(public_pool, public_count)

    # 若某类不足，从另一类补充
    shortfall = total_count - len(project_samples) - len(public_samples)
    if shortfall > 0:
        remaining = [r for r in task_db if r not in project_samples and r not in public_samples]
        extra = _weighted_sample(remaining, shortfall)
        # 优先补充到数量不足的那一类，否则均分
        for e in extra:
            if len(project_samples) < project_count:
                project_samples.append(e)
            else:
                public_samples.append(e)

    samples = project_samples + public_samples
    random.shuffle(samples)

    # 转成标准 item 格式
    items = []
    for r in samples:
        leaf = _get_leaf_workitem(r)
        w1, w2, w3 = infer_workitem_levels(task_db, leaf)
        item = {
            "申报类型": r.get("申报类型", ""),
            "项目名称": _safe_str(r.get("项目名称", "")),
            "工作项": leaf,
            "工作项1": w1,
            "工作项2": w2,
            "工作项3": w3,
            "工时": float(r.get("投入时间(H)", 0) or 0),
            "日志": _safe_str(r.get("工作日志", "")),
            "_supplemented": True,
        }
        if r.get("_from_inactive"):
            item["_from_inactive"] = True
        items.append(item)

    return items


def _get_project_core_keyword(name: str) -> str:
    """从完整项目名中提取核心关键字（去掉年份前缀，取第一个有意义的词）。
    例如 '鼎捷-2024-2025年铁塔AIOT平台开发外包项目-FP项目' -> '鼎捷'
    """
    if not name:
        return ""
    # 去掉年份前缀如 2024-2025年、2026年
    no_year = re.sub(r"\d{4}[-~]\d{4}年?", "", name)
    no_year = re.sub(r"\d{4}年?", "", no_year)
    # 按常见分隔符分割，取第一个非空且非通用的词
    parts = re.split(r"[-_/\s]", no_year)
    skip_words = {"项目", "FP", "期", "一期", "二期", "三期", "四期", "五期", ""
                  "开发", "外包", "服务", "平台", "实施", "定制", "运维", "交付"}
    for p in parts:
        p = p.strip()
        if p and p not in skip_words and len(p) >= 2:
            return p
    return ""


def supplement_with_random(
    existing_items: list[dict],
    task_db: list[dict],
    target_count: int = 6,
    ref_date_str: str | None = None,
    active_projects: list[dict] | None = None,
) -> list[dict]:
    """
    在已有 items 基础上，随机补充记录直到总条数达到 target_count。
    补充时优先避开已存在的工作项叶子，且**严格避开已存在的项目核心关键字**
    （如已有'鼎捷-2024年项目'则不再补充任何带'鼎捷'关键字的项目）。
    若严格过滤后 pool 不足，逐步放宽条件重新抽取，最后不得已才允许项目重复。

    新增：若提供 active_projects，优先从活跃项目池中抽取；
         活跃项目池不足时，降级到全历史库并标记 _from_inactive。
    """
    existing_leafs = {i["工作项"] for i in existing_items}
    existing_project_names = {
        _safe_str(i.get("项目名称", ""))
        for i in existing_items
        if _safe_str(i.get("项目名称", ""))
    }
    existing_project_keywords = {
        _get_project_core_keyword(name)
        for name in existing_project_names
    }
    ref_date = datetime.strptime(ref_date_str, "%Y-%m-%d") if ref_date_str else datetime.now()

    # 构建活跃项目池（若提供了 active_projects）
    active_pool = None
    if active_projects is not None:
        from project_index import get_active_project_names
        active_names = get_active_project_names({"projects": active_projects})
        active_pool = [r for r in task_db if r.get("项目名称") in active_names]

    leaf_counts = defaultdict(int)
    for r in task_db:
        leaf = _get_leaf_workitem(r)
        if leaf:
            leaf_counts[leaf] += 1

    def _calc_weight(r: dict) -> float:
        leaf = _get_leaf_workitem(r)
        tw = _time_weight(r.get("申报日期", ""), ref_date)
        fw = 1.0 + math.log1p(leaf_counts.get(leaf, 1))
        if leaf in existing_leafs:
            fw *= 0.1
        return tw * fw

    def _draw_samples(pool: list[dict], n: int) -> list[dict]:
        """不放回加权抽取 n 条记录"""
        pool = list(pool)
        drawn = []
        for _ in range(n):
            if not pool:
                break
            weights = [_calc_weight(r) for r in pool]
            total_w = sum(weights)
            if total_w <= 0:
                pick = random.choice(pool)
            else:
                pick = random.choices(pool, weights=weights, k=1)[0]
            drawn.append(pick)
            pool.remove(pick)
        return drawn

    def _is_project_duplicate(r: dict) -> bool:
        """判断历史记录的项目是否与已有项目核心关键字重复。"""
        proj_name = _safe_str(r.get("项目名称", ""))
        if not proj_name:
            return False  # 公共工时无项目名，不视为重复
        core_kw = _get_project_core_keyword(proj_name)
        if not core_kw:
            return False
        return core_kw in existing_project_keywords

    needed = target_count - len(existing_items)

    # 确定主池：若提供了活跃项目清单，优先使用活跃池
    primary_pool = active_pool if active_pool is not None else task_db

    # 阶段1：严格过滤（排除已存在工作项叶子 + 已存在项目核心关键字）
    strict_pool = [
        r for r in primary_pool
        if _get_leaf_workitem(r) not in existing_leafs
        and not _is_project_duplicate(r)
    ]
    supplements = _draw_samples(strict_pool, needed)

    # 阶段2：若不足，放宽工作项叶子限制，但仍排除已存在项目核心关键字
    # 对公共工时（无项目名）严格排除已有工作项，防止重复；项目工时允许同工作项不同项目
    if len(supplements) < needed:
        relax_pool = [
            r for r in primary_pool
            if r not in supplements
            and not _is_project_duplicate(r)
            and (_safe_str(r.get("项目名称")) or _get_leaf_workitem(r) not in existing_leafs)
        ]
        extra = _draw_samples(relax_pool, needed - len(supplements))
        supplements.extend(extra)

    # 阶段3：若活跃池仍不足，降级到全历史库补充
    if len(supplements) < needed:
        last_pool = [r for r in task_db if r not in supplements]
        last_pool_filtered = [
            r for r in last_pool
            if not _is_project_duplicate(r)
            and (_safe_str(r.get("项目名称")) or _get_leaf_workitem(r) not in existing_leafs)
        ]
        if len(last_pool_filtered) >= needed - len(supplements):
            extra = _draw_samples(last_pool_filtered, needed - len(supplements))
        else:
            # 实在没有更多新项目了，只能允许项目重复
            extra = _draw_samples(last_pool, needed - len(supplements))
        # 若使用了活跃池优先策略，从全历史库补充的记录标记降级来源
        if active_pool is not None:
            for e in extra:
                e["_from_inactive"] = True
        supplements.extend(extra)

    # 工时截断
    capped_supplements = []
    for pick in supplements:
        pick_copy = dict(pick)
        original_hour = pick_copy.get("工时") or pick_copy.get("投入时间(H)", 0)
        if original_hour > 2.0:
            pick_copy["工时"] = 2.0
            pick_copy["投入时间(H)"] = 2.0
        elif original_hour <= 0:
            pick_copy["工时"] = 1.0
            pick_copy["投入时间(H)"] = 1.0
        capped_supplements.append(pick_copy)

    for r in capped_supplements:
        leaf = _get_leaf_workitem(r)
        w1, w2, w3 = infer_workitem_levels(task_db, leaf)
        hour = float(r.get("工时") or r.get("投入时间(H)", 0) or 0)
        item = {
            "申报类型": r.get("申报类型", ""),
            "项目名称": _safe_str(r.get("项目名称", "")),
            "工作项": leaf,
            "工作项1": w1,
            "工作项2": w2,
            "工作项3": w3,
            "工时": hour,
            "日志": _safe_str(r.get("工作日志", "")),
            "_supplemented": True,
        }
        if r.get("_from_inactive"):
            item["_from_inactive"] = True
        existing_items.append(item)

    return existing_items


def parse_manual_tasks(add_task_strings: list[str], date_str: str, task_db: list[dict]) -> list[dict]:
    """
    解析 --add-task 参数字符串列表。
    格式: "申报类型|项目名称|工作项|工时|日志"
    申报类型支持 "项目报工"/"通用报工"，内部映射为 "项目工时"/"公共工时"。
    """
    items = []
    for s in add_task_strings:
        parts = [p.strip() for p in s.split("|")]
        if len(parts) != 5:
            print(f"[警告] --add-task 格式错误，跳过: {s}")
            continue
        raw_type, project, leaf, hours_str, log = parts
        if raw_type == "项目报工":
            decl_type = "项目工时"
        elif raw_type == "通用报工":
            decl_type = "公共工时"
        else:
            decl_type = raw_type
        try:
            hours = float(hours_str)
        except ValueError:
            hours = 1.0
        w1, w2, w3 = infer_workitem_levels(task_db, leaf)
        items.append({
            "申报类型": decl_type,
            "项目名称": project if project else "",
            "工作项": leaf,
            "工作项1": w1,
            "工作项2": w2,
            "工作项3": w3,
            "工时": hours,
            "日志": log,
        })
    return items


def append_to_task_db(records: list[dict], date_str: str, path: str = None):
    if path is None:
        path = str(ASSETS_DIR / "task_db.json")
    """
    将新增记录追加到 task_db.json。
    自动填充工时ID、申报日期、search_text、业务集团/业务线等字段。
    """
    with open(path, "r", encoding="utf-8") as fh:
        raw = json.load(fh)

    max_id = 0
    for r in raw:
        try:
            tid = int(r.get("工时ID", 0))
            if tid > max_id:
                max_id = tid
        except Exception:
            continue

    for idx, item in enumerate(records, start=1):
        max_id += 1
        project = item.get("项目名称", "")
        leaf = item.get("工作项", "")
        w1 = item.get("工作项1", leaf)
        w2 = item.get("工作项2", "")
        w3 = item.get("工作项3", "")
        log = item.get("日志", "")
        hours = item.get("工时", 0)
        decl_type = item.get("申报类型", "")

        # 业务线默认值
        biz_line = "电信与AIoT业务线"
        biz_group = "ABG"

        new_record = {
            "工时ID": max_id,
            "申报类型": decl_type,
            "申报日期": date_str,
            "业务集团": biz_group,
            "业务线": biz_line,
            "工作项1": w1,
            "工作项2": w2,
            "工作项3": w3,
            "项目名称": project if decl_type == "项目工时" else None,
            "工作日志": log,
            "投入时间(H)": hours,
            "search_text": f"{w1} {w2} {w3} {log} {project} {biz_line}".strip(),
        }
        raw.append(new_record)

    with open(path, "w", encoding="utf-8") as fh:
        json.dump(raw, fh, ensure_ascii=False, indent=2)

    print(f"[成功] 已追加 {len(records)} 条新记录到 {path}")

    # 同步到 skill 目录，保持运行时数据与 skill 副本一致
    try:
        skill_dir = Path.home() / ".skills-manager" / "skills" / "daily-worklog"
        skill_db_path = skill_dir / "assets" / "task_db.json"
        if skill_dir.exists() and Path(path).resolve() != skill_db_path.resolve():
            with open(skill_db_path, "w", encoding="utf-8") as fh:
                json.dump(raw, fh, ensure_ascii=False, indent=2)
            print(f"[成功] 已同步更新 skill 目录的 task_db.json")
    except Exception as e:
        print(f"[警告] 同步到 skill 目录失败: {e}")
