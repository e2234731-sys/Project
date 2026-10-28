# -*- coding: utf-8 -*-
import os, sys, fitz

dir_path = r"D:\wechatfile\xwechat_files\e2234731_f65d\msg\file\2026-10\26AX04192深圳凯吉星农产品检测认证有限公司"
pdf_files = [f for f in os.listdir(dir_path) if f.lower().endswith('.pdf') and ('未知' in f or '未查找到' in f or 'YYQ211' in f)]

print(f"Total problematic PDFs: {len(pdf_files)}")

for f in sorted(pdf_files):
    p = os.path.join(dir_path, f)
    doc = fitz.open(p)
    txt = doc[0].get_text()
    print("="*50)
    print("FILE:", f)
    for line in txt.splitlines():
        if any(k in line for k in ["证 书 编 号", "证书编号", "Report  No", "Certificate", "客 户 名 称", "器 具 名 称", "型 号 规 格", "制 造 单 位", "仪 器 编 号", "出厂编号", "管 理 编 号", "管理编号", "管 理 号", "校 准 日 期", "校准日期", "检 测 日 期", "检定日期", "发 布 日 期", "发布日期"]):
            print("  ", line.strip())
