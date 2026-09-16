# -*- coding: utf-8 -*-
import os, sys, fitz, pandas as pd, re

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

excel_p = r'D:\下载报告\本次扫描_校准信息识别与台账核对表_V2.0.xlsx'
df = pd.read_excel(excel_p, sheet_name='校准证书识别明细')
unknown_rows = df[df['所属组别/实验室'] == '未知组别']
print(f"Total unknown group rows: {len(unknown_rows)}")

def extract_address_station(doc):
    full_text = doc[0].get_text('text')
    stations = [
        ('安庆', '安庆市场组'),
        ('蚌埠', '蚌埠市场组'),
        ('西安', '西安项目组'),
        ('西北农副', '西安项目组'),
        ('果菜', '果菜配送中心'),
        ('坂田', '坂田街道'),
        ('龙田', '龙田街道'),
        ('横岗', '横岗街道'),
        ('南湾', '南湾街道'),
        ('碧岭', '碧岭街道'),
        ('布吉高中', '布吉高级中学'),
        ('布吉高级中学', '布吉高级中学'),
        ('布吉', '布吉市场组'),
        ('龙华', '龙华区委'),
        ('罗芳', '罗芳市场组'),
        ('大东', '大东市场'),
        ('大学城', '大学城驻点'),
        ('惠阳', '惠阳项目组'),
        ('快筛一组', '快筛一组'),
        ('快筛二组', '快筛二组'),
        ('快筛三组', '快筛三组'),
    ]
    for kw, st in stations:
        if kw in full_text:
            return st
    return '未知组别'

resolved = 0
for idx, r in unknown_rows.iterrows():
    fn = r['原始文件名']
    fp = os.path.join(r'D:\下载报告', fn)
    if os.path.exists(fp):
        with fitz.open(fp) as doc:
            st = extract_address_station(doc)
            if st != '未知组别':
                resolved += 1
            print(f"{fn:42s} -> 驻点: {st}")

print(f"\nResolved: {resolved}/{len(unknown_rows)}")
