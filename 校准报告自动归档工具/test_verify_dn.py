# -*- coding: utf-8 -*-
import os, sys, fitz, re, pandas as pd
from main_gui import LedgerDatabase, identify_issuer, normalize_date

if sys.platform.startswith('win'):
    try: sys.stdout.reconfigure(encoding='utf-8')
    except: pass

fitz.TOOLS.mupdf_display_errors(False)

DN_STATIONS = [
    ('安庆', '安庆市场组'),
    ('蚌埠', '蚌埠市场组'),
    ('西安', '西安项目组'),
    ('西北农副', '西安项目组'),
    ('成都', '成都项目组'),
    ('南昌', '南昌项目组'),
    ('青云谱', '南昌项目组'),
    ('昌南', '南昌项目组'),
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
    ('快检一组', '快检一组'),
    ('快检二组', '快检二组'),
    ('快检三组', '快检三组'),
    ('平湖', '快检一组'),
]

def new_extract_dn(text, text_clean, ledger, filename="", doc=None):
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    asset_no = None
    serial_no = ""
    blocks = doc[0].get_text('blocks') if (doc and len(doc) > 0) else []

    # 1. 优先台账已知设备编号
    for asset in ledger.sorted_asset_list:
        if re.search(r'\b' + re.escape(asset) + r'\b', text, re.I) or (filename and asset in filename):
            asset_no = asset
            break

    # 2. 空间定位抽取管理号与出厂编号
    if (not asset_no or asset_no == '未知编号') and blocks:
        for b in blocks:
            if 400 <= b[1] <= 475 and 320 <= b[0] <= 500:
                t = b[4].strip().replace('\n', '')
                if t and t != '/' and not re.match(r'^(Asset|Serial|Model|出厂|管理|器\s*具)', t, re.I):
                    asset_no = t
                    break
    if not serial_no and blocks:
        for b in blocks:
            if 400 <= b[1] <= 475 and 130 <= b[0] <= 310:
                t = b[4].strip().replace('\n', '')
                if t and t != '/' and not re.match(r'^(Asset|Serial|Model|出厂|管理)', t, re.I):
                    serial_no = t
                    break

    # 3. 文件名兜底
    if not asset_no and filename:
        parts = os.path.splitext(filename)[0].split('_')
        for p in parts:
            p_clean = re.sub(r'^HO(\d+)$', r'H0\1', p.strip())
            if p_clean in ledger.all_known_assets or (re.match(r'^(H\d+|YYQ\d+|WJ\d+|MDP\d+|KJ\d+|\d{6,12})$', p_clean) and not re.match(r'^20\d{6}$', p_clean)):
                asset_no = p_clean
                break
        if not asset_no:
            m_paren = re.search(r'[(（]([A-Za-z0-9\-]+)[)）]', filename)
            if m_paren:
                asset_no = m_paren.group(1).strip()

    if not asset_no and serial_no:
        asset_no = serial_no

    asset_no = re.sub(r'^HO(\d+)$', r'H0\1', asset_no or '')

    # 4. 校准日期抽取 (空间坐标极速精准定位 + 正则双保险)
    cal_raw = ""
    if blocks:
        for b in blocks:
            if 480 <= b[1] <= 565 and 130 <= b[0] <= 310:
                m = re.search(r'(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日|\d{4}[-/.]\d{1,2}[-/.]\d{1,2})', b[4])
                if m:
                    cal_raw = m.group(1)
                    break
    if not cal_raw:
        for i, l in enumerate(lines):
            if re.search(r'(?:Date of Calibr|校\s*准\s*日\s*期|检\s*定\s*日\s*期)', l, re.I):
                for j in range(i+1, min(i+4, len(lines))):
                    m_d = re.search(r'(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)', lines[j])
                    if m_d:
                        cal_raw = m_d.group(1)
                        break
                if cal_raw: break
    if not cal_raw:
        m_cal = re.search(r'(?:校\s*准\s*日\s*期|检\s*定\s*日\s*期|Date\s*of\s*Calibration)[^\n\r\d]*(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)', text, re.I)
        if m_cal and m_cal.group(1):
            cal_raw = m_cal.group(1)
    cal_date = normalize_date(cal_raw, filename=filename)

    # 5. 签发日期抽取
    issue_raw = ""
    if blocks:
        for b in blocks:
            if 480 <= b[1] <= 565 and 310 <= b[0] <= 500:
                m = re.search(r'(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日|\d{4}[-/.]\d{1,2}[-/.]\d{1,2})', b[4])
                if m:
                    issue_raw = m.group(1)
                    break
    if not issue_raw:
        for i, l in enumerate(lines):
            if re.search(r'(?:Date of Issue|签\s*发\s*日\s*期|批\s*准\s*日\s*期|发布日期|Issued Date)', l, re.I):
                for j in range(i+1, min(i+4, len(lines))):
                    m_d = re.search(r'(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)', lines[j])
                    if m_d:
                        issue_raw = m_d.group(1)
                        break
                if issue_raw: break
    if not issue_raw and filename:
        dates_in_fn = re.findall(r'\b(20\d{2}[01]\d[0-3]\d)\b', os.path.basename(filename))
        if len(dates_in_fn) >= 2:
            issue_raw = dates_in_fn[1]
    issue_date = normalize_date(issue_raw, filename=filename)

    # 6. 仪器名称抽取 (避开 'Description' 标签)
    inst_name = ''
    if blocks:
        for b in blocks:
            if 290 <= b[1] <= 370 and 130 <= b[0] <= 340:
                t = b[4].strip().replace('\n', '')
                if t and t != 'Description' and not re.match(r'^(Model|Serial|Asset|Manufacturer|Type|型号|出厂|管理)', t, re.I):
                    inst_name = t
                    break
    if not inst_name or inst_name == 'Description':
        if filename and '_DN' in filename:
            cand = filename.split('_DN')[0].strip()
            if cand and not cand.startswith('DN'):
                inst_name = cand

    info = ledger.query(asset_no=asset_no, serial_no=serial_no, inst_name=inst_name)
    if (not asset_no or asset_no == '未知编号') and info.get('asset_no') != '未知编号':
        asset_no = info.get('asset_no')

    if not inst_name or inst_name in {'未查找到', 'Description'}:
        inst_name = info.get('inst') if info.get('inst') not in {'未查找到', 'nan', 'None', 'Description'} else ''
    if not inst_name and filename:
        parts = os.path.splitext(filename)[0].split('_')
        if len(parts) >= 3 and not re.match(r'^\d+$', parts[-1]):
            inst_name = parts[-1]
    inst_name = inst_name or '未查找到'

    # 7. 组别与驻点智能解析 (台账 + 地址全景匹配)
    lab_name = info.get('lab') or '未知实验室'
    group_name = info.get('group') or '未知组别'

    doc_full_text = text
    if doc and len(doc) > 1:
        try:
            doc_full_text += '\n' + doc[1].get_text('text')
        except Exception:
            pass

    if group_name in {'未知组别', '未知', '未知实验室', '定量实验室'}:
        for kw, st in DN_STATIONS:
            if kw in doc_full_text:
                group_name = st
                lab_name = st
                break

    if (lab_name == '未知实验室' or lab_name == '定量实验室') and filename:
        parts = os.path.splitext(filename)[0].split('_')
        if len(parts) >= 2 and not re.match(r'^(20\d{2}|H\d+|YYQ|KJ|WJ|\d+$)', parts[0]):
            lab_name = parts[0]
            if group_name == '未知组别':
                group_name = parts[0]

    return {
        'asset_no': asset_no or '未知编号',
        'serial_no': serial_no or info.get('serial_no', ''),
        'inst_name': inst_name,
        'lab_name': lab_name,
        'cal_date': cal_date,
        'issue_date': issue_date,
        'group': group_name,
        'device_type': '快检' if info.get('type') == '快检' or group_name != '定量实验室' else '定量'
    }

# Run test on D:\下载报告
folder = r'D:\下载报告'
pdf_files = [f for f in os.listdir(folder) if f.lower().endswith('.pdf')]
ledger = LedgerDatabase()
ledger.load(base_dir=folder, log_callback=lambda m: None)

cal_ok = 0
inst_ok = 0
grp_ok = 0
for fn in pdf_files:
    fp = os.path.join(folder, fn)
    with fitz.open(fp) as doc:
        text = ''
        for i in range(min(3, len(doc))):
            text += doc[i].get_text('text') + '\n'
        text_clean = text.replace(' ', '').replace('\n', '')
        res = new_extract_dn(text, text_clean, ledger, filename=fn, doc=doc)
        if res['cal_date'] != '未知校准': cal_ok += 1
        if res['inst_name'] != '未查找到' and res['inst_name'] != 'Description': inst_ok += 1
        if res['group'] != '未知组别': grp_ok += 1

print(f"Total PDFs: {len(pdf_files)}")
print(f"Calibration dates valid: {cal_ok}/{len(pdf_files)} ({cal_ok/len(pdf_files)*100:.1f}%)")
print(f"Instrument names valid: {inst_ok}/{len(pdf_files)} ({inst_ok/len(pdf_files)*100:.1f}%)")
print(f"Groups resolved: {grp_ok}/{len(pdf_files)} ({grp_ok/len(pdf_files)*100:.1f}%)")
