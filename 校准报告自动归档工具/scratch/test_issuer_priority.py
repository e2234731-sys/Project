# -*- coding: utf-8 -*-
import os, sys, re, fitz

dir_path = r"D:\wechatfile\xwechat_files\e2234731_f65d\msg\file\2026-10\26AX04192深圳凯吉星农产品检测认证有限公司"
pdf_files = [f for f in os.listdir(dir_path) if f.lower().endswith('.pdf')]

def identify_issuer_enhanced(p1_text, full_text):
    p1_clean = re.sub(r'\s+', '', p1_text)
    full_clean = re.sub(r'\s+', '', full_text)
    
    # 优先第一页（签发机构抬头或页脚）
    for target_text in [p1_clean, full_clean]:
        if re.search(r'深圳中恒|中恒检测|zh-lab\.com|CNASL7054|ZHLH|ZHLX|ZHRG|ZHJS', target_text, re.I):
            return 'ZH'
        if re.search(r'广州力赛|力赛计量|力赛|lisaitest\.com|CNASL7127', target_text, re.I):
            return 'LISAI'
        if re.search(r'帝恩|DNTesting|CNASL6483', target_text, re.I):
            return 'DN'
        if re.search(r'达丰|DAF|DAFTJAX|DAFCJAX|DAFGJAX|CNASL6592', target_text, re.I):
            return 'DAF'
        if re.search(r'中广测|广州分析测试中心|广东省测试分析研究所|NEM|TiC2600|CNASL7613', target_text, re.I):
            return 'NEM'
        if re.search(r'中检|CCIC|ccic-mts\.com|CNASL3103', target_text, re.I):
            return 'CCIC'
        if re.search(r'深圳市计量质量检测研究院|深圳计量院|smq\.com\.cn|CNASL0579|SZJLY', target_text, re.I):
            return 'SMQ'
    return 'UNKNOWN'

for f in sorted(pdf_files):
    doc = fitz.open(os.path.join(dir_path, f))
    p1 = doc[0].get_text()
    full = ""
    for page in doc: full += page.get_text() + "\n"
    iss = identify_issuer_enhanced(p1, full)
    print(f"{f} -> {iss}")
