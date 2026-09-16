# -*- coding: utf-8 -*-
import os, sys, fitz, pandas as pd
from main_gui import LedgerDatabase, identify_issuer, extract_dn, extract_daf, extract_nem, extract_smq, extract_ccic, extract_generic

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

folder = r"D:\工作\01.实验室法定资质与17025体系\1.定量实验室17025体系维护\2.设备管理\新建文件夹"
ledger = LedgerDatabase()
ledger.load(base_dir=folder, log_callback=lambda m: None)

for root, dirs, files in os.walk(folder):
    for f in sorted(files):
        if f.lower().endswith(".pdf"):
            p = os.path.join(root, f)
            text = ""
            with fitz.open(p) as doc:
                for i in range(min(3, len(doc))):
                    text += doc[i].get_text("text") + "\n"
            text_clean = text.replace(" ", "").replace("\n", "")
            issuer = identify_issuer(text_clean)
            if issuer == "DAF":
                res = extract_daf(text, text_clean, ledger, f)
                print(f"{f:40s} -> date: {res.get('cal_date')} asset: {res.get('asset_no')} sn: {res.get('serial_no')} inst: {res.get('inst_name')}")
