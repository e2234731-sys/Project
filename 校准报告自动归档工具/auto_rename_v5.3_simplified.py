import os
import re
import shutil
import pandas as pd
import pdfplumber
from pathlib import Path
from collections import defaultdict

# ============================================================
# 校准证书智能重命名工具 v5.3 简化版
# 快检按项目组自动分类，定量按多层级结构
#
# 【v5.3 新增功能】
#   ✅ 快检设备：按"年月项目组"自动生成文件夹
#   ✅ 定量设备：按"年→月份机构→组别"多层级结构
#   ✅ 完全自动，无需指定路径
#   ✅ 同名文件跳过处理
#   ✅ 快检和定量分类统计
#
# ============================================================

print("=" * 70)
print("  校准证书智能重命名工具 v5.3 简化版")
print("  快检按项目组自动分类，定量按多层级结构")
print("=" * 70)

# ==========================================
# 第一步：加载快检和定量设备台账
# ==========================================

# ── 快检设备台账（帝恩专用，包含项目组信息）
quick_check_mapping = {}  # {编号: 项目组名}
quick_check_devices = set()
QUICK_CHECK_EXCEL = '2026年各实验室仪器设备校准清单.xlsx'
QUICK_CHECK_SHEET = '  2026年度仪器检定校准计划总表'

if os.path.exists(QUICK_CHECK_EXCEL):
    print(f"\n✅ 检测到快检设备台账：{QUICK_CHECK_EXCEL}")
    try:
        df_quick = pd.read_excel(QUICK_CHECK_EXCEL, sheet_name=QUICK_CHECK_SHEET, header=1)
        df_quick.columns = df_quick.columns.str.strip().str.replace('\n','').str.replace('\r','')
        
        if '设备编号' in df_quick.columns and '实验室名称' in df_quick.columns:
            for _, row in df_quick.iterrows():
                no = str(row['设备编号']).strip()
                lab = str(row['实验室名称']).strip()
                if no and lab and no not in {'nan','None',''}:
                    quick_check_mapping[no] = lab
                    quick_check_devices.add(no)
            print(f"   快检设备台账加载完成，共 {len(quick_check_devices)} 台设备。")
    except Exception as e:
        print(f"⚠️  快检设备台账加载失败：{e}")
else:
    print(f"\n⚠️  未找到快检设备台账【{QUICK_CHECK_EXCEL}】")

# ── 定量设备台账（量值溯源）
quantitative_devices = set()
QUANTITATIVE_EXCEL = '2026年年度计划汇总.xlsx'
QUANTITATIVE_SHEET = '2026年量值溯源总表'

if os.path.exists(QUANTITATIVE_EXCEL):
    print(f"\n✅ 检测到定量设备台账：{QUANTITATIVE_EXCEL}")
    try:
        df_raw = pd.read_excel(QUANTITATIVE_EXCEL, sheet_name=QUANTITATIVE_SHEET, header=None)
        header_row = 0
        for idx, row in df_raw.iterrows():
            row_str = ' '.join(str(v) for v in row.values)
            if ('设备编号' in row_str or '器具编号' in row_str or '资产编号' in row_str) \
               and ('组' in row_str):
                header_row = idx
                break
        df_quant = pd.read_excel(QUANTITATIVE_EXCEL, sheet_name=QUANTITATIVE_SHEET, header=header_row)
        df_quant.columns = df_quant.columns.str.strip().str.replace('\n','').str.replace('\r','')
        
        no_col = next(
            (c for c in df_quant.columns
             if re.search(r'设备编号|器具编号|资产编号|管理号', c)), None)
        
        if no_col:
            df_quant[no_col] = df_quant[no_col].astype(str).str.strip()
            quantitative_devices = set(df_quant[no_col].unique())
            invalid_q = {'nan','None','','/','—','无'}
            quantitative_devices = {d for d in quantitative_devices if d not in invalid_q}
            print(f"   定量设备台账加载完成，共 {len(quantitative_devices)} 台设备。")
    except Exception as e:
        print(f"⚠️  定量设备台账加载失败：{e}")
else:
    print(f"\n⚠️  未找到定量设备台账【{QUANTITATIVE_EXCEL}】")

# ── 台账C：帝恩设备详细信息
dn_mapping = {}
dn_valid_assets = []
if os.path.exists(QUICK_CHECK_EXCEL):
    try:
        df_dn = pd.read_excel(QUICK_CHECK_EXCEL, sheet_name=QUICK_CHECK_SHEET, header=1)
        df_dn.columns = df_dn.columns.str.strip().str.replace('\n','').str.replace('\r','')
        for col in ['设备编号','实验室名称','设备名称']:
            if col in df_dn.columns:
                df_dn[col] = df_dn[col].astype(str).str.strip()
        for _, row in df_dn.iterrows():
            dn_mapping[row['设备编号']] = {
                'lab': row.get('实验室名称', '未查找到'),
                'inst': row.get('设备名称', '未查找到')
            }
        invalid = {'nan','None','','/','\\','-','无','无编号'}
        dn_valid_assets = sorted(
            [k for k in dn_mapping if k not in invalid], key=len, reverse=True)
    except Exception as e:
        pass

# ── 台账D：量值溯源总表（组别查询）
group_mapping = {}
if os.path.exists(QUANTITATIVE_EXCEL):
    try:
        df_raw = pd.read_excel(QUANTITATIVE_EXCEL, sheet_name=QUANTITATIVE_SHEET, header=None)
        header_row = 0
        for idx, row in df_raw.iterrows():
            row_str = ' '.join(str(v) for v in row.values)
            if ('设备编号' in row_str or '器具编号' in row_str or '资产编号' in row_str) \
               and ('组' in row_str):
                header_row = idx
                break
        df_trace = pd.read_excel(QUANTITATIVE_EXCEL, sheet_name=QUANTITATIVE_SHEET, header=header_row)
        df_trace.columns = df_trace.columns.str.strip().str.replace('\n','').str.replace('\r','')

        no_col = next(
            (c for c in df_trace.columns
             if re.search(r'设备编号|器具编号|资产编号|管理号', c)), None)
        grp_col = next(
            (c for c in df_trace.columns
             if re.search(r'所属组|组别|实验室组', c)), None)

        if no_col and grp_col:
            df_trace[no_col]  = df_trace[no_col].astype(str).str.strip()
            df_trace[grp_col] = df_trace[grp_col].astype(str).str.strip()
            invalid_g = {'nan','None','','/','—','无'}
            for _, row in df_trace.iterrows():
                code = row[no_col]; grp = row[grp_col]
                if code not in invalid_g and grp not in invalid_g:
                    group_mapping[code] = grp
    except Exception as e:
        pass

extracted_info_list = []

# ==========================================
# 第二步：通用工具函数
# ==========================================

def sanitize_filename(text):
    if not text:
        return "未查找到"
    cleaned = re.sub(r'[\\/*?:"<>|]', '-', str(text)).strip()
    return cleaned if cleaned else "未查找到"

def clean_date_chinese(date_str):
    if not date_str:
        return None
    return re.sub(r'[年月日\s]', '', date_str)

def clean_date_hyphen(date_str):
    if not date_str:
        return None
    return date_str.replace('-', '').strip()

def clean_date_slash(date_str):
    if not date_str:
        return None
    return re.sub(r'[\s/]', '', date_str)

def parse_yyyymm(date8):
    if date8 and re.match(r'^\d{8}$', date8):
        return date8[:4], date8[4:6]
    return None, None

def find_value_by_keyword(text, keyword_pattern, value_pattern,
                          search_range=4,
                          col_tolerance_left=15, col_tolerance_right=60):
    lines = text.split('\n')
    for i, line in enumerate(lines):
        kw_match = re.search(keyword_pattern, line)
        if not kw_match:
            continue
        col_idx = kw_match.start()
        end_idx = kw_match.end()
        for j in range(i, min(i + search_range, len(lines))):
            target_line = lines[j]
            for val_match in re.finditer(value_pattern, target_line):
                v_start = val_match.start()
                val = val_match.group(0).strip()
                if not val:
                    continue
                if j == i:
                    if v_start >= end_idx:
                        return val
                else:
                    if (col_idx - col_tolerance_left) <= v_start <= (col_idx + col_tolerance_right):
                        return val
    return None

def lookup_group(code):
    if not code or code in ('未知编号', '未查找到'):
        return '未知组别'
    if code in group_mapping:
        return group_mapping[code]
    code_upper = code.upper()
    for k, v in group_mapping.items():
        if k.upper() == code_upper:
            return v
    return '未知组别'

def determine_device_type(asset_no):
    """判断设备是快检还是定量"""
    if asset_no in quick_check_devices:
        return '快检'
    elif asset_no in quantitative_devices:
        return '定量'
    else:
        return '未知'

# ==========================================
# 第三步：机构识别
# ==========================================

def identify_issuer(text_clean):
    if re.search(r'帝恩|DNTesting|CNASL6483', text_clean):
        return 'DN'
    if re.search(r'达丰|DAF|DAFDJAX|CNASL6592', text_clean):
        return 'DAF'
    if re.search(r'深圳市计量质量检测研究院|smq\.com\.cn|CNASL0579', text_clean):
        return 'SMQ'
    if re.search(r'中检|CCIC|ccic-mts\.com|CNASL3103', text_clean):
        return 'CCIC'
    if re.search(r'中广测|NEM|TiC2600|CNASL7613', text_clean):
        return 'NEM'
    return 'UNKNOWN'

# ==========================================
# 第四步：各机构提取器
# ==========================================

def extract_dn(text, text_clean):
    asset_no = None
    for asset in dn_valid_assets:
        if asset in text or asset in text_clean:
            asset_no = asset
            break
    if not asset_no:
        asset_no = find_value_by_keyword(
            text, r'管\s*理\s*号|Asset\s*No\.?',
            r'[A-Z0-9\-]{3,30}',
            col_tolerance_left=10, col_tolerance_right=40)
    asset_no = asset_no or '未知编号'

    info      = dn_mapping.get(asset_no, {})
    lab_name  = info.get('lab', '未查找到')
    inst_name = info.get('inst', '未查找到')

    cal_raw   = find_value_by_keyword(
        text, r'校\s*准\s*日\s*期|Date\s*of\s*Calibration',
        r'\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日',
        col_tolerance_left=15, col_tolerance_right=40)
    cal_date  = clean_date_chinese(cal_raw) or '未知校准'

    return {'asset_no': asset_no, 'inst_name': inst_name,
            'lab_name': lab_name, 'cal_date': cal_date,
            'group': lookup_group(asset_no)}

def extract_daf(text, text_clean):
    asset_no = find_value_by_keyword(
        text, r'管\s*理\s*号|Asset\s*No\.?',
        r'[A-Z0-9][A-Z0-9\-/]{2,29}',
        col_tolerance_left=0, col_tolerance_right=80)
    
    if not asset_no:
        fallback = re.search(r'管\s*理\s*号[:：\s]*([A-Z0-9][A-Z0-9\-/]{2,20})', text)
        if fallback:
            asset_no = fallback.group(1)
    
    if asset_no:
        if re.match(r'^(JJF|JJG|GB|YY|PALL|SCIEX|Waters|Thermo|Agilent|Shimadzu|testo)\b', 
                    asset_no, re.IGNORECASE):
            asset_no = None
        elif not re.search(r'\d', asset_no):
            asset_no = None
    asset_no = asset_no or '未知编号'

    inst_name = find_value_by_keyword(
        text, r'仪\s*器\s*名\s*称|Description',
        r'[\u4e00-\u9fa5][A-Za-z0-9\-\u4e00-\u9fa5]{2,30}',
        col_tolerance_right=100) or '未查找到'
    
    blacklist = {'型号规格','型号','规格','制造商','制造厂商','出厂编号',
                 'Serial','管理号','Asset','联络信息','Information',
                 'Model','Type','Manufacturer'}
    if any(b in inst_name for b in blacklist):
        inst_name = '未查找到'

    cal_raw  = find_value_by_keyword(
        text, r'校\s*准\s*日\s*期|Date\s*of\s*Calibration',
        r'\d{4}-\d{2}-\d{2}',
        col_tolerance_left=15, col_tolerance_right=80)
    cal_date = clean_date_hyphen(cal_raw) or '未知校准'

    return {'asset_no': asset_no, 'inst_name': inst_name,
            'cal_date': cal_date, 'group': lookup_group(asset_no)}

def extract_nem(text, text_clean):
    serial_no = find_value_by_keyword(
        text, r'器\s*具\s*编\s*号|Serial\s*[N№]',
        r'[A-Z0-9][A-Z0-9\-/]{2,29}',
        col_tolerance_left=5, col_tolerance_right=80)
    
    if not serial_no:
        fallback = re.search(r'器\s*具\s*编\s*号[:：\s]*([A-Z0-9][A-Z0-9\-/]{2,20})', text)
        if fallback:
            serial_no = fallback.group(1)
    
    if serial_no and serial_no.startswith('TiC'):
        serial_no = None
    serial_no = serial_no or '未知编号'

    inst_name = find_value_by_keyword(
        text, r'器\s*具\s*名\s*称|Description',
        r'[\u4e00-\u9fa5][A-Za-z0-9\-\u4e00-\u9fa5]{2,30}',
        col_tolerance_right=100) or '未查找到'
    
    blacklist = {'型号规格','型号','制造商','器具编号','联络信息','委托方'}
    if any(b in inst_name for b in blacklist):
        inst_name = '未查找到'

    cal_raw  = find_value_by_keyword(
        text, r'校\s*准\s*日\s*期|Date\s*of\s*Calibrat(?:e|ion)',
        r'\d{4}-\d{2}-\d{2}',
        col_tolerance_left=15, col_tolerance_right=80)
    cal_date = clean_date_hyphen(cal_raw) or '未知校准'

    return {'serial_no': serial_no, 'inst_name': inst_name,
            'cal_date': cal_date, 'group': lookup_group(serial_no)}

def extract_smq(text, text_clean):
    asset_no = find_value_by_keyword(
        text, r'资\s*产\s*编\s*号|Asset\s*No',
        r'[A-Z][A-Z0-9]{2,29}',
        col_tolerance_left=5, col_tolerance_right=80)
    
    if not asset_no:
        fallback = re.search(r'资\s*产\s*编\s*号[:：\s]*([A-Z][A-Z0-9]{2,20})', text)
        if fallback:
            asset_no = fallback.group(1)
    
    if asset_no and not re.search(r'\d', asset_no):
        asset_no = None
    asset_no = asset_no or '未知编号'

    inst_name = find_value_by_keyword(
        text, r'计\s*量\s*器\s*具\s*名\s*称|Name\s*of\s*Instrument',
        r'[\u4e00-\u9fa5][A-Za-z0-9\-\u4e00-\u9fa5]{2,30}',
        col_tolerance_right=100) or '未查找到'
    
    blacklist = {'型号','规格','制造单位','出厂编号','资产编号','校准依据'}
    if any(b in inst_name for b in blacklist):
        inst_name = '未查找到'

    cal_raw   = find_value_by_keyword(
        text, r'校\s*准\s*日\s*期|Operation\s*Date',
        r'\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日',
        col_tolerance_left=15, col_tolerance_right=80)
    cal_date  = clean_date_chinese(cal_raw) or '未知校准'

    return {'asset_no': asset_no, 'inst_name': inst_name,
            'cal_date': cal_date, 'group': lookup_group(asset_no)}

def extract_ccic(text, text_clean):
    asset_no = find_value_by_keyword(
        text, r'管\s*理\s*编\s*号|Asset\s*No\.?',
        r'[A-Z][A-Z0-9]{2,29}',
        col_tolerance_left=5, col_tolerance_right=80)
    
    if not asset_no:
        fallback = re.search(r'管\s*理\s*编\s*号[:：\s]*([A-Z][A-Z0-9]{2,20})', text)
        if fallback:
            asset_no = fallback.group(1)
    
    if asset_no and not re.search(r'\d', asset_no):
        asset_no = None
    asset_no = asset_no or '未知编号'

    inst_name = find_value_by_keyword(
        text, r'仪\s*器\s*名\s*称|Description',
        r'[\u4e00-\u9fa5][A-Za-z0-9\-\u4e00-\u9fa5]{2,30}',
        col_tolerance_right=100) or '未查找到'
    
    blacklist = {'型号规格','型号','制造厂商','出厂编号','管理编号',
                 '接收日期','接收状态','结论','Serial'}
    if any(b in inst_name for b in blacklist):
        inst_name = '未查找到'

    cal_raw  = find_value_by_keyword(
        text, r'校\s*准\s*日\s*期|Cal\.?\s*Date',
        r'\d{4}\s*/\s*\d{1,2}\s*/\s*\d{1,2}',
        col_tolerance_left=15, col_tolerance_right=80)
    cal_date = clean_date_slash(cal_raw) or '未知校准'

    return {'asset_no': asset_no, 'inst_name': inst_name,
            'cal_date': cal_date, 'group': lookup_group(asset_no)}

# ==========================================
# 第五步：生成文件名和路径
# ==========================================

def build_new_name(issuer, fields):
    if issuer == 'DN':
        parts = [fields['lab_name'], fields['asset_no'],
                 fields['cal_date'], fields['inst_name']]
    elif issuer == 'DAF':
        parts = [fields['asset_no'], fields['inst_name'],
                 fields['cal_date'], fields['group']]
    elif issuer == 'NEM':
        parts = [fields['serial_no'], fields['inst_name'],
                 fields['cal_date'], fields['group']]
    elif issuer == 'SMQ':
        parts = [fields['asset_no'], fields['inst_name'],
                 fields['cal_date'], fields['group']]
    elif issuer == 'CCIC':
        parts = [fields['asset_no'], fields['inst_name'],
                 fields['cal_date'], fields['group']]
    else:
        parts = ['未识别机构']
    return '_'.join(sanitize_filename(p) for p in parts)

def build_archive_path_quick_check(issuer, fields):
    """快检设备：按年月项目组生成路径"""
    cal_date = fields.get('cal_date', '')
    year, month = parse_yyyymm(cal_date)
    
    if not year:
        return None
    
    project_group = fields.get('lab_name', '未知项目组')
    
    # 快检文件夹：YYYY年MM月{项目组}设备校准证书
    folder_name = f"{year}年{month}月{project_group}设备校准证书"
    
    return folder_name

def build_archive_path_quantitative(issuer, fields):
    """定量设备：按年→月份机构→组别多层级生成路径"""
    cal_date = fields.get('cal_date', '')
    year, month = parse_yyyymm(cal_date)
    
    if not year:
        return None
    
    issuer_name = {
        'DN':   '帝恩',
        'DAF':  '达丰',
        'NEM':  '中广测',
        'SMQ':  '深圳计量院',
        'CCIC': '中检',
    }.get(issuer, '未知机构')
    
    group = fields.get('group', '未知组别')
    
    year_folder = f"{year}年检定校准证书"
    month_issuer_folder = f"{year}年{month}月{issuer_name}校准证书"
    detail_folder = f"{year}年{month}月{group}校准证书-{issuer_name}"
    
    full_path = os.path.join(year_folder, month_issuer_folder, detail_folder)
    
    return full_path

# ==========================================
# 第六步：主循环处理
# ==========================================

ISSUER_LABEL = {
    'DN':      '东莞帝恩',
    'DAF':     '深圳达丰',
    'NEM':     '广州中广测',
    'SMQ':     '深圳计量院',
    'CCIC':    '中检深圳',
    'UNKNOWN': '未识别机构',
}

print("\n开始扫描 PDF 证书...\n")

ARCHIVE_ROOT = './【归档完成】校准证书库'
os.makedirs(ARCHIVE_ROOT, exist_ok=True)

skipped_count = 0

for root, dirs, files in os.walk('.'):
    if '【归档完成】' in root or '年检定校准证书' in root or '设备校准证书' in root:
        continue

    for filename in files:
        if not filename.lower().endswith('.pdf'):
            continue

        file_path = os.path.join(root, filename)

        try:
            with pdfplumber.open(file_path) as pdf:
                first_page = pdf.pages[0]
                text = first_page.extract_text(layout=True) or first_page.extract_text() or ''

            text_clean = text.replace(' ', '')

            # ① 识别机构
            issuer = identify_issuer(text_clean)
            label = ISSUER_LABEL[issuer]

            # ② 提取字段
            if issuer == 'DN':
                fields = extract_dn(text, text_clean)
            elif issuer == 'DAF':
                fields = extract_daf(text, text_clean)
            elif issuer == 'NEM':
                fields = extract_nem(text, text_clean)
            elif issuer == 'SMQ':
                fields = extract_smq(text, text_clean)
            elif issuer == 'CCIC':
                fields = extract_ccic(text, text_clean)
            else:
                print(f"⚠️  未识别机构：{filename}，跳过。")
                continue

            # ③ 判断设备类型
            code_key = 'serial_no' if issuer == 'NEM' else 'asset_no'
            device_code = fields.get(code_key, '')
            device_type = determine_device_type(device_code)

            # ④ 组装新文件名
            new_stem = build_new_name(issuer, fields)
            new_name = f"{new_stem}.pdf"

            # ⑤ 确定目标路径
            if device_type == '快检':
                folder_name = build_archive_path_quick_check(issuer, fields)
                dst_dir = os.path.join(ARCHIVE_ROOT, '快检设备', folder_name or '未分类')
            elif device_type == '定量':
                folder_path = build_archive_path_quantitative(issuer, fields)
                dst_dir = os.path.join(ARCHIVE_ROOT, '定量设备', folder_path or '未分类')
            else:
                dst_dir = os.path.join(ARCHIVE_ROOT, '未分类设备')

            os.makedirs(dst_dir, exist_ok=True)
            dst_path = os.path.join(dst_dir, new_name)

            # ⑥ 检查同名冲突
            if os.path.exists(dst_path):
                print(f"⏭️  跳过：{filename}")
                print(f"     原因：目标路径已存在同名文件")
                skipped_count += 1
                continue

            # ⑦ 搬运文件
            shutil.move(file_path, dst_path)
            print(f"✅ [{label}] [{device_type}]")
            print(f"     原文件：{filename}")
            print(f"     新文件：{new_name}")
            print(f"     归档至：{os.path.relpath(dst_dir, ARCHIVE_ROOT)}/")

            # ⑧ 汇总记录
            folder_name = os.path.basename(root) or '当前主文件夹'
            group_val = fields.get('group', '—')

            extracted_info_list.append({
                '所在文件夹': folder_name,
                '原始文件名': filename,
                '最终重命名': new_name,
                '机构': label,
                '编号': fields.get(code_key, '—'),
                '仪器名称': fields['inst_name'],
                '校准日期': fields['cal_date'],
                '所属组别': group_val,
                '设备类型': device_type,
                '归档路径': os.path.relpath(dst_dir, ARCHIVE_ROOT),
            })

        except Exception as e:
            print(f"❌ 处理 [{filename}] 时发生错误：{e}")

# ==========================================
# 第七步：生成统计和汇总
# ==========================================

print("\n" + "=" * 70)
print("【处理完成统计】")
print("=" * 70)

if extracted_info_list:
    # 按设备类型分类
    quick_check_list = [x for x in extracted_info_list if x['设备类型'] == '快检']
    quantitative_list = [x for x in extracted_info_list if x['设备类型'] == '定量']
    unknown_list = [x for x in extracted_info_list if x['设备类型'] == '未知']
    
    print(f"\n✅ 处理总数：{len(extracted_info_list)} 份")
    print(f"   快检设备：{len(quick_check_list)} 份")
    print(f"   定量设备：{len(quantitative_list)} 份")
    print(f"   未分类：{len(unknown_list)} 份")
    print(f"⏭️  跳过（同名冲突）：{skipped_count} 份")
    
    # 按机构统计
    print(f"\n📊 按机构分布：")
    issuer_count = {}
    for item in extracted_info_list:
        issuer = item['机构']
        issuer_count[issuer] = issuer_count.get(issuer, 0) + 1
    
    for issuer, count in sorted(issuer_count.items(), key=lambda x: -x[1]):
        print(f"   {issuer}: {count} 份")
    
    # 生成Excel汇总
    df_result = pd.DataFrame(extracted_info_list)
    col_order = ['所在文件夹', '原始文件名', '最终重命名',
                 '机构', '编号', '仪器名称', '校准日期',
                 '所属组别', '设备类型', '归档路径']
    df_result = df_result[col_order]
    output_excel = '本次扫描_校准信息提取汇总表.xlsx'
    df_result.to_excel(output_excel, index=False, engine='openpyxl')
    
    print(f"\n🎉 处理完成！")
    print(f"📋 汇总清单：{output_excel}")
    print(f"📁 归档结构：")
    print(f"   ./【归档完成】校准证书库/")
    print(f"   ├── 快检设备/")
    print(f"   │   ├── 2026年4月龙岗街道设备校准证书/")
    print(f"   │   ├── 2026年4月南山街道设备校准证书/")
    print(f"   │   └── ...")
    print(f"   ├── 定量设备/")
    print(f"   │   ├── 2026年检定校准证书/")
    print(f"   │   │   ├── 2026年4月中广测校准证书/")
    print(f"   │   │   │   └── 2026年4月理化微生物组校准证书-中广测/")
    print(f"   │   │   └── ...")
    print(f"   │   └── ...")
    print(f"   └── 未分类设备/")

else:
    print("\n⚠️  未找到任何 PDF 文件。")

print("\n" + "=" * 70)
