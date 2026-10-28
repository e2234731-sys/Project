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

def test_extract_lisai(text, text_clean, ledger, filename="", doc=None):
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    asset_no = ''
    serial_no = ''
    inst_name = ''
    cal_date = ''
    issue_date = ''
    
    # 1. 管理号 / 设备编号
    for i, l in enumerate(lines):
        if re.search(r'管\s*理\s*号|Asset\s*No', l, re.I):
            for j in range(i+1, min(i+4, len(lines))):
                cand = lines[j].strip()
                if cand and not any(k in cand for k in ['Date', 'Receipt', 'Calibration', '接收', '校准', 'Asset', 'No']):
                    asset_no = cand
                    break
            if asset_no: break

    # 2. 出厂编号
    for i, l in enumerate(lines):
        if re.search(r'出\s*厂\s*编\s*号|Serial\s*No', l, re.I):
            for j in range(i+1, min(i+4, len(lines))):
                cand = lines[j].strip()
                if cand and not any(k in cand for k in ['管', 'Asset', 'Date', 'Receipt', 'Calibration', '接收', '校准', 'Serial', 'No']):
                    serial_no = cand
                    break
            if serial_no: break

    # 3. 仪器/样品名称
    for i, l in enumerate(lines):
        if re.search(r'仪\s*器\s*/?\s*样\s*品\s*名\s*称|样\s*品\s*名\s*称', l, re.I):
            for j in range(i+1, min(i+4, len(lines))):
                cand = lines[j].strip()
                if cand and cand.lower() != 'description' and not any(k in cand for k in ['Model', 'Type', '型号', '规格', 'Manufacturer', '制造']):
                    inst_name = cand
                    break
            if inst_name: break

    # 4. 校准日期
    for i, l in enumerate(lines):
        if re.search(r'校\s*准\s*日\s*期|Date\s*of\s*Calibration', l, re.I):
            for j in range(i+1, min(i+4, len(lines))):
                m_d = re.search(r'(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)', lines[j])
                if m_d:
                    cal_date = m_d.group(1)
                    break
            if cal_date: break

    # 5. 发布/签发日期
    for i, l in enumerate(lines):
        if re.search(r'发\s*布\s*日\s*期|Issued\s*Date|批\s*准\s*日\s*期', l, re.I):
            for j in range(i+1, min(i+4, len(lines))):
                m_d = re.search(r'(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)', lines[j])
                if m_d:
                    issue_date = m_d.group(1)
                    break
            if issue_date: break

    asset_no = re.sub(r'^HO(\d+)$', r'H0\1', asset_no or '')
    cal_date_norm = main_gui.normalize_date(cal_date, filename=filename)
    issue_date_norm = main_gui.normalize_date(issue_date, filename=filename)
    
    info = ledger.query(asset_no=asset_no, serial_no=serial_no, inst_name=inst_name)
    if not asset_no and info.get('asset_no') != '未知编号':
        asset_no = info.get('asset_no')
    if not inst_name or inst_name == '未查找到':
        inst_name = info.get('inst') if info.get('inst') not in {'未查找到', 'nan', 'None'} else '未查找到'

    return {
        'asset_no': asset_no or '未知编号',
        'serial_no': serial_no or '',
        'inst_name': inst_name,
        'cal_date': cal_date_norm,
        'issue_date': issue_date_norm,
        'group': info.get('group', '未知组别'),
        'lab_name': info.get('lab', '定量实验室'),
        'device_type': info.get('type', '定量')
    }

def test_extract_zh(text, text_clean, ledger, filename="", doc=None):
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    asset_no = ''
    serial_no = ''
    inst_name = ''
    cal_date = ''
    issue_date = ''

    for l in lines:
        # 管理编号
        m_a = re.search(r'管\s*理\s*编\s*号\s*[：:]\s*([A-Za-z0-9\-]+)', l)
        if m_a and not asset_no:
            asset_no = m_a.group(1).strip()
            
        # 仪器编号 / 出厂编号
        m_s = re.search(r'(?:仪\s*器\s*编\s*号|出\s*厂\s*编\s*号)\s*[：:]\s*([^/\s]+)', l)
        if m_s and not serial_no:
            sn_val = m_s.group(1).strip()
            if sn_val != '/':
                serial_no = sn_val
                
        # 器具名称 / 仪器名称
        m_i = re.search(r'(?:器\s*具\s*名\s*称|仪\s*器\s*名\s*称)\s*[：:]\s*(.+)', l)
        if m_i and not inst_name:
            cand = m_i.group(1).strip()
            if cand and not any(k in cand for k in ['Name of', 'Type', 'Model', '规格', '型号']):
                inst_name = cand
                
        # 校准日期 / 检测日期
        m_d = re.search(r'(?:校\s*准\s*日\s*期|检\s*测\s*日\s*期|检\s*定\s*日\s*期)\s*[：:]\s*(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)', l)
        if m_d and not cal_date:
            cal_date = m_d.group(1).strip()

        # 发布日期
        m_iss = re.search(r'发\s*布\s*日\s*期\s*[：:]\s*(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)', l)
        if m_iss and not issue_date:
            issue_date = m_iss.group(1).strip()

    # 如果有换行格式（标签与值不在同一行）
    if not asset_no:
        for i, l in enumerate(lines):
            if re.search(r'^管\s*理\s*编\s*号\s*[：:]?$', l) and i+1 < len(lines):
                cand = lines[i+1].strip()
                if re.match(r'^[A-Za-z0-9\-]+$', cand):
                    asset_no = cand
                    break

    if not serial_no:
        for i, l in enumerate(lines):
            if re.search(r'^(?:仪\s*器\s*编\s*号|出\s*厂\s*编\s*号)\s*[：:]?$', l) and i+1 < len(lines):
                cand = lines[i+1].strip()
                if cand != '/' and not any(k in cand for k in ['Serial', 'Management', 'Date']):
                    serial_no = cand
                    break

    if not inst_name:
        for i, l in enumerate(lines):
            if re.search(r'^(?:器\s*具\s*名\s*称|仪\s*器\s*名\s*称)\s*[：:]?$', l) and i+1 < len(lines):
                cand = lines[i+1].strip()
                if not any(k in cand for k in ['Name of', 'Type', 'Model']):
                    inst_name = cand
                    break

    if not cal_date:
        for i, l in enumerate(lines):
            if re.search(r'^(?:校\s*准\s*日\s*期|检\s*测\s*日\s*期|Test\s*Date|Calibration\s*Date)\s*[：:]?$', l) and i+1 < len(lines):
                m = re.search(r'(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)', lines[i+1])
                if m:
                    cal_date = m.group(1)
                    break

    # 规范化
    asset_no = re.sub(r'^HO(\d+)$', r'H0\1', asset_no or '')
    cal_date_norm = main_gui.normalize_date(cal_date, filename=filename)
    issue_date_norm = main_gui.normalize_date(issue_date, filename=filename)

    info = ledger.query(asset_no=asset_no, serial_no=serial_no, inst_name=inst_name)
    if not asset_no and info.get('asset_no') != '未知编号':
        asset_no = info.get('asset_no')
    if not inst_name or inst_name == '未查找到':
        inst_name = info.get('inst') if info.get('inst') not in {'未查找到', 'nan', 'None'} else '未查找到'

    return {
        'asset_no': asset_no or '未知编号',
        'serial_no': serial_no or '',
        'inst_name': inst_name,
        'cal_date': cal_date_norm,
        'issue_date': issue_date_norm,
        'group': info.get('group', '未知组别'),
        'lab_name': info.get('lab', '定量实验室'),
        'device_type': info.get('type', '定量')
    }

print("=== 测试所有 27 个文件提取 ===")
pdf_files = [f for f in os.listdir(dir_path) if f.lower().endswith('.pdf')]
for idx, f in enumerate(sorted(pdf_files)):
    p = os.path.join(dir_path, f)
    doc = fitz.open(p)
    text = ""
    for page in doc:
        text += page.get_text() + "\n"
    text_clean = re.sub(r'\s+', '', text)
    
    if "力赛" in text_clean or "L7127" in text_clean or "lisaitest" in text_clean:
        issuer = "LISAI"
        fields = test_extract_lisai(text, text_clean, ledger, filename=f, doc=doc)
    elif "中恒" in text_clean or "L7054" in text_clean or "zh-lab" in text_clean or "ZHLH" in text_clean or "ZHLX" in text_clean or "ZHRG" in text_clean or "ZHJS" in text_clean:
        issuer = "ZH"
        fields = test_extract_zh(text, text_clean, ledger, filename=f, doc=doc)
    else:
        issuer = "UNKNOWN"
        fields = main_gui.extract_generic(text, text_clean, ledger, filename=f, doc=doc)

    fn = main_gui.build_new_filename(issuer, fields)
    print(f"[{idx+1:02d}] {issuer:5s} | 编号: {fields['asset_no']:8s} | 名称: {fields['inst_name']:18s} | 日期: {fields['cal_date']} | 组别: {fields['group']:10s} -> 新名: {fn}.pdf")
