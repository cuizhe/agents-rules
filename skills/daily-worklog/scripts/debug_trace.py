import json
from fill_worklog import _run_fill_impl
import sys

# Monkey-patch to trace items
original_impl = _run_fill_impl.__globals__['_run_fill_impl']

def traced_impl(*args, **kwargs):
    # Call the actual function but we'll just call it normally
    # Instead, let's directly test the problematic path
    pass

# Direct test: simulate what happens with manual tasks only
from engine import (
    ASSETS_DIR,
    load_task_db, filter_by_date_rules, supplement_with_random,
    adjust_to_eight, enforce_daily_quality_planning_limit, merge_project_rows,
    replace_month_in_logs, infer_workitem_levels
)
import yaml

try:
    with open(str(ASSETS_DIR / "config.yaml"), "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
except Exception:
    config = {}

db = load_task_db()
holidays = config.get("preferences", {}).get("holidays", [])

date_str = "2026-06-05"

# Step 1: parse manual tasks
from fill_worklog import parse_manual_tasks
add_tasks = [
    "项目工时|明生恒卓-2026年信息服务驻点外包框架-自动化管控工具FP项目|项目立项与排产计划管理|0.5|集成管理系统项目关联、项目关键角色任命，质量策划、CP评审计划沟通",
    "项目工时|鼎捷MES-2026-2027年服务外包合同-TM项目|项目立项与排产计划管理|0.5|集成管理系统项目关联、项目关键角色任命，质量策划、CP评审计划沟通",
    "项目工时|深圳移动-2026-2027年血液中心综合楼配套信息化建设实施-IOC定制开发FP项目|项目立项与排产计划管理|0.5|集成管理系统项目关联、项目关键角色任命，质量策划、CP评审计划沟通",
    "通用报工||关键节点评审与过程审视|2|6月业务线FP项目排产计划评审会议",
    "通用报工||AI工具引入与推广策划|2|集团转型AI组织运作与AI工具应用推广相关工作"
]

manual_items = parse_manual_tasks(add_tasks, date_str, db)
print("=== Manual items ===")
for i, m in enumerate(manual_items):
    print(f"{i}: type={m['申报类型']}, proj={m['项目名称']}, wi={m['工作项']}, hour={m['工时']}")

items = []
items.extend(manual_items)
print(f"\n=== After extend: {len(items)} items ===")
for i, it in enumerate(items):
    print(f"{i}: type={it['申报类型']}, proj={it['项目名称']}, wi={it['工作项']}, hour={it['工时']}")

# Simulate the fix: populate maps
project_hour_map = {}
project_workitem_map = {}
for m in manual_items:
    if m.get("申报类型") == "项目工时" and m.get("项目名称"):
        project_hour_map[m["项目名称"]] = m["工时"]
        project_workitem_map[m["项目名称"]] = m["工作项"]

print(f"\n=== project_hour_map ===")
print(project_hour_map)
print(f"=== project_workitem_map ===")
print(project_workitem_map)

# Step: supplement
from datetime import datetime
total_hours = sum(i["工时"] for i in items)
print(f"\n=== Before supplement: {len(items)} items, {total_hours}H ===")

if len(items) < 5 or total_hours < 8.0:
    import random
    target = random.choice([5, 6])
    if len(items) < target:
        print(f"[补充随机] 当前 {len(items)} 条/{total_hours}H，补充至 {target} 条")
        items = supplement_with_random(items, db, target, date_str)
        print(f"\n=== After supplement: {len(items)} items ===")
        for i, it in enumerate(items):
            print(f"{i}: type={it['申报类型']}, proj={it['项目名称']}, wi={it['工作项']}, hour={it['工时']}")

        items, extra_filter_logs = filter_by_date_rules(items, date_str, holidays)
        print(f"\n=== After filter: {len(items)} items ===")
        if extra_filter_logs:
            print("Filter logs:")
            for log in extra_filter_logs:
                print(f"  {log}")
        for i, it in enumerate(items):
            print(f"{i}: type={it['申报类型']}, proj={it['项目名称']}, wi={it['工作项']}, hour={it['工时']}")

# Step: adjust to eight
print(f"\n=== Before adjust_to_eight ===")
for i, it in enumerate(items):
    print(f"{i}: type={it['申报类型']}, proj={it['项目名称']}, wi={it['工作项']}, hour={it['工时']}")

dynamic_items = [i for i in items if i.get("申报类型") == "项目工时"]
public_items = [i for i in items if i.get("申报类型") != "项目工时"]
user_specified_items = []
other_dynamic_items = []
for item in dynamic_items:
    pname = item.get("项目名称", "")
    is_specified = False
    for kw in project_hour_map.keys():
        if kw in pname:
            item["工时"] = project_hour_map[kw]
            is_specified = True
            break
    if is_specified:
        user_specified_items.append(item)
    else:
        other_dynamic_items.append(item)

print(f"\ndynamic_items={len(dynamic_items)}, public={len(public_items)}")
print(f"user_specified={len(user_specified_items)}, other_dynamic={len(other_dynamic_items)}")
print(f"project_hour_map keys={list(project_hour_map.keys())}")

public_hours = sum(i["工时"] for i in public_items)
reserved_hours = public_hours + sum(i["工时"] for i in user_specified_items)
print(f"public_hours={public_hours}, reserved_hours={reserved_hours}")

if other_dynamic_items:
    other_dynamic_items = adjust_to_eight(other_dynamic_items, reserved_hours=reserved_hours)
elif reserved_hours < 8.0 and user_specified_items:
    print(f"[工时调整] 指定项目已满但合计仅 {reserved_hours}H，按比例放大项目工时至 8H")
    user_specified_items = adjust_to_eight(user_specified_items, reserved_hours=public_hours)
items = user_specified_items + other_dynamic_items + public_items

print(f"\n=== After adjust: {len(items)} items ===")
for i, it in enumerate(items):
    print(f"{i}: type={it['申报类型']}, proj={it['项目名称']}, wi={it['工作项']}, hour={it['工时']}")

# Step: enforce quality planning limit
items = enforce_daily_quality_planning_limit(items, limit=1)
print(f"\n=== After enforce_daily_qp: {len(items)} items ===")
for i, it in enumerate(items):
    print(f"{i}: type={it['申报类型']}, proj={it['项目名称']}, wi={it['工作项']}, hour={it['工时']}")
