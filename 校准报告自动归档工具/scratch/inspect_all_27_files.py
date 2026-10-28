# -*- coding: utf-8 -*-
import os, sys, re, fitz

dir_path = r"D:\wechatfile\xwechat_files\e2234731_f65d\msg\file\2026-10\26AX04192深圳凯吉星农产品检测认证有限公司"
pdf_files = [f for f in os.listdir(dir_path) if f.lower().endswith('.pdf')]

print(f"Total PDFs: {len(pdf_files)}")

for idx, f in enumerate(sorted(pdf_files)):
    p = os.path.join(dir_path, f)
    doc = fitz.open(p)
    text = ""
    for page in doc:
        text += page.get_text() + "\n"
    
    # Identify supplier
    supplier = "未知"
    if "力赛" in text or "L7127" in text or "lisaitest" in text:
        supplier = "力赛"
    elif "中恒" in text or "L7054" in text or "zh-lab" in text or "ZHLH" in text or "ZHLX" in text or "ZHRG" in text:
        supplier = "中恒"
        
    p1 = doc[0].get_text()
    p2 = doc[1].get_text() if len(doc) > 1 else ""
    
    print(f"\n=================================================================")
    print(f"[{idx+1}/{len(pdf_files)}] 文件名: {f} | 供应商: {supplier} | 总页数: {len(doc)}")
    print("--- 第一页 (Page 1) ---")
    for l in p1.splitlines():
        if l.strip():
            print("  ", l.strip())
    if "中恒" in supplier:
        # also check page 2 if needed
        print("--- 第一页/第二页关键信息 ---")
