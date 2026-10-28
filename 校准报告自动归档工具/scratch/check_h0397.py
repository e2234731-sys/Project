# -*- coding: utf-8 -*-
import os, sys, fitz

dir_path = r"D:\wechatfile\xwechat_files\e2234731_f65d\msg\file\2026-10\26AX04192深圳凯吉星农产品检测认证有限公司"

for name in ["H0397_二级生物安全柜_未知校准_微生物组.pdf", "H0398_二级生物安全柜_未知校准_微生物组.pdf"]:
    p = os.path.join(dir_path, name)
    print("="*50)
    print("FILE:", name)
    doc = fitz.open(p)
    for pno in range(min(2, len(doc))):
        print(f"--- Page {pno+1} ---")
        print(doc[pno].get_text())
