# -*- coding: utf-8 -*-
import os, sys, fitz

dir_path = r"D:\wechatfile\xwechat_files\e2234731_f65d\msg\file\2026-10\26AX04192深圳凯吉星农产品检测认证有限公司"
files = [f for f in os.listdir(dir_path) if "Description" in f or "WJ" in f]

for f in sorted(files):
    p = os.path.join(dir_path, f)
    doc = fitz.open(p)
    txt = doc[0].get_text()
    print("="*40)
    print("FILE:", f)
    for line in txt.splitlines()[:35]:
        print("  ", line.strip())
