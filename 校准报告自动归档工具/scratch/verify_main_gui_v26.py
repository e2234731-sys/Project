# -*- coding: utf-8 -*-
import os, sys, re, fitz

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import main_gui

print("main_gui VERSION:", main_gui.VERSION)
print("main_gui ISSUER_LABEL:", main_gui.ISSUER_LABEL)

dir_path = r"D:\wechatfile\xwechat_files\e2234731_f65d\msg\file\2026-10\26AX04192深圳凯吉星农产品检测认证有限公司"
ledger = main_gui.LedgerDatabase()
ledger.load()

pdf_files = [f for f in os.listdir(dir_path) if f.lower().endswith('.pdf')]
print(f"Total PDFs in directory: {len(pdf_files)}")

success_count = 0
for idx, f in enumerate(sorted(pdf_files)):
    pdf_path = os.path.join(dir_path, f)
    doc = fitz.open(pdf_path)
    text = ""
    for p_i in range(min(3, len(doc))):
        text += (doc[p_i].get_text("text") or '') + '\n'
    text_clean = text.replace(' ', '').replace('\n', '')

    p1_clean = (doc[0].get_text("text") or '').replace(' ', '').replace('\n', '') if len(doc) > 0 else ''
    issuer = main_gui.identify_issuer(text_clean, text_p1_clean=p1_clean)
    if issuer == 'DN': fields = main_gui.extract_dn(text, text_clean, ledger, filename=f, doc=doc)
    elif issuer == 'DAF': fields = main_gui.extract_daf(text, text_clean, ledger, filename=f, doc=doc)
    elif issuer == 'NEM': fields = main_gui.extract_nem(text, text_clean, ledger, filename=f, doc=doc)
    elif issuer == 'SMQ': fields = main_gui.extract_smq(text, text_clean, ledger, filename=f, doc=doc)
    elif issuer == 'CCIC': fields = main_gui.extract_ccic(text, text_clean, ledger, filename=f, doc=doc)
    elif issuer == 'LISAI': fields = main_gui.extract_lisai(text, text_clean, ledger, filename=f, doc=doc)
    elif issuer == 'ZH': fields = main_gui.extract_zh(text, text_clean, ledger, filename=f, doc=doc)
    else: fields = main_gui.extract_generic(text, text_clean, ledger, filename=f, doc=doc)

    doc.close()

    new_fn = f"{main_gui.build_new_filename(issuer, fields)}.pdf"
    
    # Validation checks
    is_asset_ok = fields['asset_no'] != '未知编号'
    is_inst_ok = fields['inst_name'] != '未查找到' and fields['inst_name'] != 'Description'
    is_date_ok = fields['cal_date'] != '未知校准'
    is_issuer_ok = issuer in ['LISAI', 'ZH']
    
    status = "OK" if (is_asset_ok and is_inst_ok and is_date_ok and is_issuer_ok) else "WARN"
    if status == "OK": success_count += 1
    
    print(f"[{idx+1:02d}] [{status}] 机构:{issuer:5s} | 编号:{fields['asset_no']:8s} | 名称:{fields['inst_name']:18s} | 日期:{fields['cal_date']} | 组别:{fields['group']:10s} -> {new_fn}")

print(f"\n==========================================")
print(f"验证完成: 成功识别 {success_count}/{len(pdf_files)} 个文件！")
