# -*- coding: utf-8 -*-
from openpyxl import load_workbook

wb = load_workbook('C:/Users/Administrator/.workbuddy/skills/agriculture-stats/references/2026年电信与AIoT业务线FP项目毛利管控表20260413.xlsx', data_only=False)
ws = wb['实际运营数据-崔哲']

# 找到 P202504000428 项目
row = 2
while row <= ws.max_row:
    code = ws.cell(row=row, column=1).value
    name = ws.cell(row=row, column=2).value
    if str(code) == 'P202504000428':
        print(f'=== {name} (起始行{row}) ===')
        # 打印10行的关键列
        for offset in range(10):
            r = row + offset
            cat = ws.cell(row=r, column=18).value
            
            print(f'\n  行{r}: {cat}')
            # col19-33: 预算、实际、YTD各月
            print('    col19-25:', end=' ')
            for c in range(19, 26):
                v = ws.cell(row=r, column=c).value
                dt = ws.cell(row=r, column=c).data_type
                val_str = repr(v)[:15] if v else '-'
                print(f'[{dt}]{val_str}', end=' | ')
            
            print()
            print('    col35-52 (月度):', end=' ')
            for c in [35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46]:
                v = ws.cell(row=r, column=c).value
                dt = ws.cell(row=r, column=c).data_type
                val_str = f'{v:.0f}' if isinstance(v, (int, float)) else (repr(v)[:10] if v else '-')
                print(f'[{dt}]{val_str}', end=' | ')
        break
    row += 10  # 跳到下一个项目

wb.close()
