import sys
sys.path.insert(0, r"D:\Document\study\claude-code-projects\daily-report")

from engine import adjust_to_eight, supplement_with_random, load_task_db, filter_by_date_rules
from fill_worklog import parse_manual_tasks

db = load_task_db()

# 场景：用户精确指定 3项目×0.5h + 1公共×2h = 3.5h
# 系统会补充随机记录，验证方案一兜底逻辑
add_tasks = [
    "项目工时|明生恒卓-2026年信息服务驻点外包框架-自动化管控工具FP项目|项目立项与排产计划管理|0.5|测试日志",
    "项目工时|鼎捷MES-2026-2027年服务外包合同-TM项目|项目立项与排产计划管理|0.5|测试日志",
    "项目工时|深圳移动-2026-2027年血液中心综合楼配套信息化建设实施-IOC定制开发FP项目|项目立项与排产计划管理|0.5|测试日志",
    "通用报工||排产与PO运营|2|6月业务线FP项目排产计划评审会议",
]

date_str = "2026-06-09"

manual_items = parse_manual_tasks(add_tasks, date_str, db)
print("=== Manual items ===")
for m in manual_items:
    tag = "(固定)"
    print(f"  {tag} {m['申报类型']} | {m.get('项目名称','')} | {m['工作项']} | {m['工时']}H")

items = []
items.extend(manual_items)

# 模拟 fill_worklog.py 中的 project_hour_map 构建
project_hour_map = {}
for m in manual_items:
    if m.get("申报类型") == "项目工时" and m.get("项目名称"):
        project_hour_map[m["项目名称"]] = m["工时"]

# 补充随机记录到 5~6 条
import random
target = random.choice([5, 6])
if len(items) < target:
    print(f"\n[补充随机] 当前 {len(items)} 条，补充至 {target} 条")
    items = supplement_with_random(items, db, target, date_str)

print("\n=== After supplement ===")
for i, it in enumerate(items):
    tag = "(补充)" if it.get("_supplemented") else "(固定)"
    print(f"  {tag} {it['申报类型']} | {it.get('项目名称','')} | {it['工作项']} | {it['工时']}H")

# 日期过滤
items, logs = filter_by_date_rules(items, date_str, [])
if logs:
    print("\n[过滤日志]")
    for log in logs:
        print(f"  {log}")

# 模拟 fill_worklog.py 第10步逻辑
dynamic_items = [i for i in items if i.get("申报类型") == "项目工时"]
public_items = [i for i in items if i.get("申报类型") != "项目工时"]
user_specified_items = []
other_dynamic_items = []
for item in dynamic_items:
    pname = item.get("项目名称", "")
    is_specified = any(kw in pname for kw in project_hour_map.keys())
    if is_specified:
        user_specified_items.append(item)
    else:
        other_dynamic_items.append(item)

public_hours = sum(i["工时"] for i in public_items)
reserved_hours = public_hours + sum(i["工时"] for i in user_specified_items)

print(f"\n[诊断] public_hours={public_hours}, reserved_hours={reserved_hours}")
print(f"[诊断] other_dynamic_items={len(other_dynamic_items)}, user_specified_items={len(user_specified_items)}")

if other_dynamic_items:
    print("[逻辑] 存在可调项目工时，对其缩放")
    other_dynamic_items = adjust_to_eight(other_dynamic_items, reserved_hours=reserved_hours)
elif reserved_hours < 8.0:
    print("[逻辑] 无可调项目且不足8H，进入方案一兜底")
    supplemented_items = [i for i in items if i.get("_supplemented")]
    if supplemented_items:
        fixed_items = [i for i in items if not i.get("_supplemented")]
        fixed_hours = sum(i["工时"] for i in fixed_items)
        print(f"  fixed_hours={fixed_hours}, supplemented_count={len(supplemented_items)}")
        print(f"  缩放前 supplemented 工时: {[i['工时'] for i in supplemented_items]}")
        supplemented_items = adjust_to_eight(supplemented_items, reserved_hours=fixed_hours)
        print(f"  缩放后 supplemented 工时: {[i['工时'] for i in supplemented_items]}")
    else:
        print("  [警告] 无补充记录可缩放")

items = user_specified_items + other_dynamic_items + public_items
total = sum(i["工时"] for i in items)
print(f"\n=== 最终草稿（合计 {total}H）===")
for it in items:
    tag = "(补充)" if it.get("_supplemented") else "(固定)"
    print(f"  {tag} {it['申报类型']} | {it.get('项目名称','')} | {it['工作项']} | {it['工时']}H")

# 断言验证
assert total == 8.0, f"总工时必须为 8.0，实际为 {total}"
print("\n[验证通过] 总工时严格等于 8.0H")


print("\n\n========== 测试用例2：无可调项目工时，仅补充公共工时 ==========")
items2 = []
items2.extend(manual_items)
# 手动添加一条补充的公共工时，模拟 supplement_with_random 的行为
items2.append({
    "申报类型": "通用报工",
    "项目名称": None,
    "工作项": "内部专项工作",
    "工时": 1.0,
    "日志": "测试补充公共工时",
    "_supplemented": True,
})

# 模拟 fill_worklog.py 第10步逻辑
dynamic_items2 = [i for i in items2 if i.get("申报类型") == "项目工时"]
public_items2 = [i for i in items2 if i.get("申报类型") != "项目工时"]
user_specified_items2 = []
other_dynamic_items2 = []
for item in dynamic_items2:
    pname = item.get("项目名称", "")
    is_specified = any(kw in pname for kw in project_hour_map.keys())
    if is_specified:
        user_specified_items2.append(item)
    else:
        other_dynamic_items2.append(item)

public_hours2 = sum(i["工时"] for i in public_items2)
reserved_hours2 = public_hours2 + sum(i["工时"] for i in user_specified_items2)
print(f"[诊断] public_hours={public_hours2}, reserved_hours={reserved_hours2}")
print(f"[诊断] other_dynamic_items={len(other_dynamic_items2)}, user_specified_items={len(user_specified_items2)}")

if other_dynamic_items2:
    print("[逻辑] 存在可调项目工时，对其缩放")
    other_dynamic_items2 = adjust_to_eight(other_dynamic_items2, reserved_hours=reserved_hours2)
elif reserved_hours2 < 8.0:
    print("[逻辑] 无可调项目且不足8H，进入方案一兜底")
    supplemented_items2 = [i for i in items2 if i.get("_supplemented")]
    if supplemented_items2:
        fixed_items2 = [i for i in items2 if not i.get("_supplemented")]
        fixed_hours2 = sum(i["工时"] for i in fixed_items2)
        print(f"  fixed_hours={fixed_hours2}, supplemented_count={len(supplemented_items2)}")
        print(f"  缩放前 supplemented 工时: {[i['工时'] for i in supplemented_items2]}")
        supplemented_items2 = adjust_to_eight(supplemented_items2, reserved_hours=fixed_hours2)
        print(f"  缩放后 supplemented 工时: {[i['工时'] for i in supplemented_items2]}")
    else:
        print("  [警告] 无补充记录可缩放")

items2 = user_specified_items2 + other_dynamic_items2 + public_items2
total2 = sum(i["工时"] for i in items2)
print(f"\n=== 最终草稿（合计 {total2}H）===")
for it in items2:
    tag = "(补充)" if it.get("_supplemented") else "(固定)"
    print(f"  {tag} {it['申报类型']} | {it.get('项目名称','')} | {it['工作项']} | {it['工时']}H")

assert total2 == 8.0, f"总工时必须为 8.0，实际为 {total2}"
print("\n[验证通过] 总工时严格等于 8.0H")
