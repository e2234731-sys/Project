# -*- coding: utf-8 -*-
import os, sys, re, fitz

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

import main_gui

dir_path = r"D:\wechatfile\xwechat_files\e2234731_f65d\msg\file\2026-10\26AX04192深圳凯吉星农产品检测认证有限公司"
ledger = main_gui.LedgerDatabase()
ledger.load()

pdf_files = [f for f in os.listdir(dir_path) if f.lower().endswith('.pdf')]

print(f"Total files: {len(pdf_files)}")

for f in sorted(pdf_files):
    p = os.path.join(dir_path, f)
    doc = fitz.open(p)
    text = ""
    for page in doc:
        text += page.get_text() + "\n"
    text_clean = re.sub(r'\s+', '', text)
    
    current_issuer = main_gui.identify_issuer(text_clean)
    
    detected_target = "UNKNOWN"
    if "力赛" in text_clean or "L7127" in text_clean or "lisaitest" in text_clean:
        detected_target = "LISAI"
    elif "中恒" in text_clean or "L7054" in text_clean or "zh-lab" in text_clean or "ZHLH" in text_clean or "ZHLX" in text_clean or "ZHRG" in text_clean:
        detected_target = "ZH"
        
    print(f"\n---------------------------------------------")
    print(f"FILE: {f}")
    print(f"Current Issuer: {current_issuer}  -->  Target Issuer: {detected_target}")
    
    res = main_gui.parse_certificate_data(p, ledger, ocr_engine=None)
    fields = res.get('fields', {})
    print(f"Fields extracted by current code:")
    print(f"  asset_no: {fields.get('asset_no')}")
    print(f"  serial_no: {fields.get('serial_no')}")
    print(f"  inst_name: {fields.get('inst_name')}")
    print(f"  cal_date: {fields.get('cal_date')}")
    print(f"  group: {fields.get('group')}")
    print(f"  device_type: {fields.get('device_type')}")
