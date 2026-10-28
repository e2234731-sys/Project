# -*- coding: utf-8 -*-
import os, sys, re, openpyxl

files_to_check = [
    r"2026年年度计划汇总.xlsx",
    r"2026年各实验室仪器设备校准清单.xlsx",
    r"D:\工作\01.实验室法定资质与17025体系\1.定量实验室17025体系维护\2.设备管理\1.设备检定校准\2026年年度计划汇总.xlsx",
    r"D:\工作\03.快检质量网络与驻点管理\8.快检工作\7.快检设备检定校准\2026年各实验室仪器设备校准清单.xlsx"
]

search_codes = ['WJ136', 'WJ137', 'WJ138', 'WJ139', 'WJ140', 'WJ141', 'WJ142', 'H0537', 'H0538', 'H0541', 'H0545', 'YYQ150', 'YYQ211']

for fpath in files_to_check:
    if not os.path.exists(fpath): continue
    print("="*50)
    print("Searching in:", fpath)
    wb = openpyxl.load_workbook(fpath, data_only=True)
    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        for r in range(1, min(ws.max_row+1, 500)):
            for c in range(1, min(ws.max_column+1, 25)):
                val = str(ws.cell(r, c).value or '').strip()
                for sc in search_codes:
                    if sc.lower() in val.lower():
                        row_vals = [str(ws.cell(r, col).value or '').strip() for col in range(1, 15)]
                        print(f"  [{sheetname}] Row {r}: Col {c}={val} | Row data: {row_vals}")
