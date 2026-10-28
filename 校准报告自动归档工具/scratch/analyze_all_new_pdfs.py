# -*- coding: utf-8 -*-
import os, sys, re
import fitz

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

dir_path = r"D:\wechatfile\xwechat_files\e2234731_f65d\msg\file\2026-10\26AX04192深圳凯吉星农产品检测认证有限公司"
pdf_files = [f for f in os.listdir(dir_path) if f.lower().endswith('.pdf')]

print(f"Total PDFs: {len(pdf_files)}")

for f in sorted(pdf_files):
    path = os.path.join(dir_path, f)
    try:
        doc = fitz.open(path)
        full_text = ""
        for i in range(len(doc)):
            full_text += f"\n--- Page {i+1} ---\n" + doc[i].get_text()
        
        # Check机构
        org = []
        if "广州力赛" in full_text or "力赛" in full_text:
            org.append("广州力赛")
        if "中恒" in full_text or "ZH" in full_text:
            org.append("深圳中恒")
        if "达丰" in full_text:
            org.append("达丰")
        if "盟翊" in full_text:
            org.append("盟翊")
        
        print(f"\n=======================================================")
        print(f"File: {f}")
        print(f"Possible Org: {org}")
        print(f"Total Pages: {len(doc)}")
        
        # Let's search for key labels:
        # 器具名称 / 仪器名称 / 设备名称 / 设备编号 / 管理编号 / 出厂编号 / 制造单位 / 校准日期 / 检定日期 / 有效期 / 建议再校准日期
        lines = [line.strip() for line in full_text.splitlines() if line.strip()]
        for line in lines:
            if any(k in line for k in ["力赛", "中恒", "证书编号", "Certificate", "器具名称", "仪器名称", "型号规格", "出厂编号", "管理编号", "设备编号", "校准日期", "检定日期", "核准日期", "有效日期", "建议下", "建议再"]):
                print(f"  > {line}")
                
    except Exception as e:
        print(f"Error {f}: {e}")
