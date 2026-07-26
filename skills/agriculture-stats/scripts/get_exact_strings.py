import openpyxl
from pathlib import Path

ref_dir = Path(__file__).parent.parent / 'references'
output_file = list(ref_dir.glob('*汇总*.xlsx'))[0]

wb = openpyxl.load_workbook(output_file)
ws_summary = wb['当月汇总毛利率']

# 直接从 Excel 读取指标名称
accounting_methods = []
for row in range(4, 8):
    val = ws_summary.cell(row=row, column=2).value
    accounting_methods.append(val)
    print(f"行{row}: '{val}'")
    print(f"  字节：{val.encode('utf-8')}")
    print(f"  长度：{len(val)}")
    print(f"  字符:")
    for i, c in enumerate(val):
        print(f"    [{i}] '{c}' U+{ord(c):04X}")
    print()

# 保存为 Python 代码格式
print("\n=== Python 代码格式 ===")
print("accounting_methods = [")
for method in accounting_methods:
    # 使用 Unicode 转义序列显示不可见字符
    escaped = ''.join(f'\\u{ord(c):04X}' if ord(c) > 127 else c for c in method)
    print(f"    '{escaped}',  # {method}")
print("]")
