import openpyxl
from pathlib import Path

ref_dir = Path(__file__).parent.parent / 'references'
output_file = list(ref_dir.glob('*汇总*.xlsx'))[0]

wb = openpyxl.load_workbook(output_file)
ws_summary = wb['当月汇总毛利率']

# 检查列 4 的值
val = ws_summary.cell(row=3, column=4).value
print(f"汇总表列 4: '{val}'")
print(f"长度：{len(val)}")

# 测试字符串
test_val = '项目开始 - 当前实际累计'
print(f"\n测试字符串：'{test_val}'")
print(f"长度：{len(test_val)}")

# 逐字符比较
print(f"\n逐字符比较:")
for i in range(min(len(val), len(test_val))):
    match = "✓" if val[i] == test_val[i] else "✗"
    print(f"  [{i}] '{val[i]}' (U+{ord(val[i]):04X}) vs '{test_val[i]}' (U+{ord(test_val[i]):04X}) {match}")

# 测试相等
print(f"\n相等测试：{val == test_val}")
