# -*- coding: utf-8 -*-
"""
校准报告智能重命名与多层级自动归档工具 v25.0 (生产体系完美适配版)
======================================================================
基于《历年设备检定校准证书》真实生产目录与命名体系深度建模与优化：

【核心优化与特性】：
1. 【完美对齐现行生产归档体系】
   - 深入学习了 2019-2026 历年真实归档结构与命名规范：
     - 根目录层：{YYYY}年检定校准证书
     - 二级目录：{YYYY}年{M}月{机构}校准证书 (例如: 2026年4月达丰校准证书)
     - 三级目录：{YYYY}年{M}月{组别}校准证书-{机构} (例如: 2026年4月气相组校准证书-中广测 / 2026年04月理化微生物组校准证书)
     - 快检设备：快检设备/{YYYY}年{M}月{项目组}设备校准证书/
   - 智能识别标准组别：理化微生物组、理化组、气相组、液相组、元素组、抽样组、综合组、运行组等。
2. 【智能四合一命名引擎】
   - 标准定量设备：{设备编号}_{设备名称}_{校准日期}_{所属组别}.pdf
   - 移液器复合编号：{出厂编号}({设备编号})_{设备名称}_{校准日期}_{所属组别}.pdf (自动识别移液器 SN 与 YYQ 编号关联)
   - 帝恩快检设备：{实验室/项目组}_{设备编号}_{校准日期}_{签发日期}_{设备名称}.pdf
   - 特种设备(压力表/安全阀)：{设备编号}_{关联设备}压力表_{校准日期}.pdf / {设备编号}安全阀校验报告_{校准日期}.pdf
3. 【全工作表台账聚合数据库】
   - 自动扫描《2026年年度计划汇总.xlsx》与《2026年各实验室仪器设备校准清单.xlsx》所有工作表（包括量值溯源总表、各季度/月份校准子表）。
   - 构建多维双向索引（设备编号、出厂号、设备名称、所属组别、实验室），支持精确与模糊反查。
4. 【PyMuPDF + RapidOCR 双引擎极速解析】
   - 毫秒级提取文本层，遇图片扫描件或无 CMap 乱码层自动调用 RapidOCR 视觉模型终极抢救。
5. 【深度遍历与安全保障】
   - 递归扫描任意多层文件夹，自动解压 ZIP 压缩包并纠正 Windows 编码乱码。
   - 智能排除已归档库目录，防死循环；同名文件自动按 _1, _2 序号保护，绝不覆盖丢失。
   - 自动导出商务美化 Excel 汇总表 (openpyxl 自动列宽/斑马纹/商务蓝表头) 与 TXT 处理统计清单。
======================================================================
"""

import os
import sys
import re
import io
import shutil
import zipfile
import argparse
import datetime
import warnings
from pathlib import Path
from collections import defaultdict

# 优化 Windows 终端中文编码输出
if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

# 屏蔽 openpyxl 读取 Excel 时的样式警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import pandas as pd
import fitz  # PyMuPDF
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 尝试加载 RapidOCR 引擎
try:
    from rapidocr_onnxruntime import RapidOCR
    ocr_engine = RapidOCR()
    HAS_OCR = True
    ocr_status_desc = "✅ 已激活 (ONNX 极速视觉模型)"
except Exception as e:
    ocr_engine = None
    HAS_OCR = False
    ocr_status_desc = f"⚠️ 未激活 ({e})"

VERSION = "v25.0"
TOOL_TITLE = f"校准报告智能重命名与多层级自动归档工具 {VERSION}"

# 预设的定量实验室标准组别
STANDARD_GROUPS = [
    '理化微生物组', '理化组', '微生物组',
    '液相组', '气相组', '元素组',
    '抽样组', '综合组', '运行组',
    '质保部', '报告组'
]


# ============================================================
# 第一部分：台账天眼词库加载器（全工作表深度聚合）
# ============================================================

class LedgerDatabase:
    """全景台账数据库：聚合所有子表，提供双向检索与智能纠错"""
    def __init__(self):
        self.quick_check_map = {}     # {设备编号: {'lab': 实验室, 'inst': 设备名, 'type': '快检'}}
        self.quantitative_map = {}    # {设备编号: {'group': 组别, 'inst': 设备名, 'serial_no': 出厂号, 'type': '定量'}}
        self.serial_to_asset = {}     # {出厂编号: 设备编号}
        self.all_known_assets = set() # 所有有效设备编号集合
        self.sorted_asset_list = []

    def load(self, base_dir='.'):
        print("\n" + "=" * 70)
        print("【步骤 1/4】加载仪器设备台账天眼词库 (全工作表深度聚合)...")
        print("=" * 70)

        # 1. 扫描加载快检设备台账
        quick_candidates = [
            '2026年各实验室仪器设备校准清单.xlsx',
            '各实验室仪器设备校准清单.xlsx',
        ]
        quick_loaded = False
        for fname in quick_candidates:
            fpath = os.path.join(base_dir, fname)
            if os.path.exists(fpath):
                self._load_quick_excel(fpath)
                quick_loaded = True
                break
        if not quick_loaded:
            for f in os.listdir(base_dir):
                if f.endswith('.xlsx') and '校准清单' in f and not f.startswith('~$') and not '汇总' in f:
                    self._load_quick_excel(os.path.join(base_dir, f))
                    quick_loaded = True
                    break

        # 2. 扫描加载定量设备台账（多工作表全景扫描）
        quant_candidates = [
            '2026年年度计划汇总.xlsx',
            '年度计划汇总.xlsx',
            '量值溯源总表.xlsx'
        ]
        quant_loaded = False
        for fname in quant_candidates:
            fpath = os.path.join(base_dir, fname)
            if os.path.exists(fpath):
                self._load_quantitative_excel(fpath)
                quant_loaded = True
                break
        if not quant_loaded:
            for f in os.listdir(base_dir):
                if f.endswith('.xlsx') and ('计划汇总' in f or '量值溯源' in f) and not f.startswith('~$') and not '提取汇总' in f:
                    self._load_quantitative_excel(os.path.join(base_dir, f))
                    quant_loaded = True
                    break

        # 3. 整理全局编号倒序列表（用于长词优先匹配）
        invalid_set = {'NAN', 'NONE', '', '/', '\\', '-', '—', '无', '无编号', '待定', '暂不校准'}
        all_assets = set()
        for k in self.quick_check_map:
            if k.upper() not in invalid_set:
                all_assets.add(k)
        for k in self.quantitative_map:
            if k.upper() not in invalid_set:
                all_assets.add(k)
        self.all_known_assets = all_assets
        self.sorted_asset_list = sorted(list(all_assets), key=len, reverse=True)

        print(f"📊 台账数据库构建完成：")
        print(f"   - 快检设备：{len(self.quick_check_map)} 台")
        print(f"   - 定量设备：{len(self.quantitative_map)} 台 (关联出厂号 {len(self.serial_to_asset)} 条)")
        print(f"   - 全局唯一有效设备索引库：共 {len(self.sorted_asset_list)} 条")

    def _load_quick_excel(self, fpath):
        print(f"📗 正在加载快检设备台账：{os.path.basename(fpath)}")
        try:
            xl = pd.ExcelFile(fpath)
            for sheet in xl.sheet_names:
                df = pd.read_excel(fpath, sheet_name=sheet, header=None)
                header_idx = -1
                for idx, row in df.head(10).iterrows():
                    row_txt = ' '.join([str(v) for v in row.values if pd.notna(v)])
                    if '设备编号' in row_txt and ('实验室' in row_txt or '设备名称' in row_txt):
                        header_idx = idx
                        break
                if header_idx != -1:
                    df_sheet = pd.read_excel(fpath, sheet_name=sheet, header=header_idx)
                    df_sheet.columns = [str(c).strip().replace('\n','').replace('\r','') for c in df_sheet.columns]
                    no_col = next((c for c in df_sheet.columns if '设备编号' in c), None)
                    lab_col = next((c for c in df_sheet.columns if '实验室' in c), None)
                    inst_col = next((c for c in df_sheet.columns if '设备名称' in c or '仪器名称' in c), None)

                    if no_col:
                        for _, r in df_sheet.iterrows():
                            code = str(r[no_col]).strip() if pd.notna(r[no_col]) else ''
                            lab = str(r[lab_col]).strip() if lab_col and pd.notna(r[lab_col]) else ''
                            inst = str(r[inst_col]).strip() if inst_col and pd.notna(r[inst_col]) else ''
                            if code and code not in {'nan', 'None', '', '/', '—', '无'}:
                                self.quick_check_map[code] = {
                                    'lab': lab or '未知实验室',
                                    'inst': inst or '未查找到',
                                    'type': '快检'
                                }
            print(f"   ✅ 快检台账解析成功，已导入 {len(self.quick_check_map)} 条快检设备数据。")
        except Exception as e:
            print(f"   ⚠️ 快检台账加载出现异常：{e}")

    def _load_quantitative_excel(self, fpath):
        print(f"📘 正在加载定量量值溯源台账 (扫描所有工作表)：{os.path.basename(fpath)}")
        try:
            xl = pd.ExcelFile(fpath)
            for sheet in xl.sheet_names:
                df_raw = pd.read_excel(fpath, sheet_name=sheet, header=None)
                header_idx = -1
                for idx, row in df_raw.head(10).iterrows():
                    row_txt = ' '.join([str(v) for v in row.values if pd.notna(v)])
                    if ('设备编号' in row_txt or '器具编号' in row_txt or '资产编号' in row_txt) and \
                       ('组' in row_txt or '仪器名称' in row_txt or '规格' in row_txt):
                        header_idx = idx
                        break
                if header_idx != -1:
                    df_sheet = pd.read_excel(fpath, sheet_name=sheet, header=header_idx)
                    df_sheet.columns = [str(c).strip().replace('\n','').replace('\r','') for c in df_sheet.columns]
                    no_col = next((c for c in df_sheet.columns if re.search(r'设备编号|器具编号|资产编号|管理号', c)), None)
                    grp_col = next((c for c in df_sheet.columns if re.search(r'所属组|组别|实验室组', c)), None)
                    inst_col = next((c for c in df_sheet.columns if re.search(r'仪器名称|设备名称|器具名称', c)), None)
                    sn_col = next((c for c in df_sheet.columns if re.search(r'出厂编号|序列号|SN', c)), None)

                    # 智能检测扩展组别列（如 'Unnamed: 24' 中包含理化微生物组）
                    extra_grp_col = next((c for c in df_sheet.columns if any(g in str(df_sheet[c].values) for g in STANDARD_GROUPS) and c != grp_col), None)

                    if no_col:
                        for _, r in df_sheet.iterrows():
                            code = str(r[no_col]).strip() if pd.notna(r[no_col]) else ''
                            grp = str(r[grp_col]).strip() if grp_col and pd.notna(r[grp_col]) else ''
                            if extra_grp_col and pd.notna(r[extra_grp_col]):
                                ext_val = str(r[extra_grp_col]).strip()
                                if ext_val in STANDARD_GROUPS:
                                    grp = ext_val

                            inst = str(r[inst_col]).strip() if inst_col and pd.notna(r[inst_col]) else ''
                            sn = str(r[sn_col]).strip() if sn_col and pd.notna(r[sn_col]) else ''

                            invalid = {'nan', 'None', '', '/', '—', '无'}
                            if code and code not in invalid:
                                grp_val = grp if grp not in invalid else '未知组别'
                                inst_val = inst if inst not in invalid else '未查找到'
                                
                                # 如果已有此编号但之前是未知组别，优先更新为更详细的组别
                                if code not in self.quantitative_map or (self.quantitative_map[code]['group'] == '未知组别' and grp_val != '未知组别'):
                                    self.quantitative_map[code] = {
                                        'group': grp_val,
                                        'inst': inst_val,
                                        'serial_no': sn if sn not in invalid else '',
                                        'type': '定量'
                                    }
                                if sn and sn not in invalid and len(sn) >= 3:
                                    self.serial_to_asset[sn] = code
            print(f"   ✅ 定量台账解析成功，已导入 {len(self.quantitative_map)} 条定量设备数据。")
        except Exception as e:
            print(f"   ⚠️ 定量台账加载出现异常：{e}")

    def query(self, asset_no=None, serial_no=None):
        """双向查询设备台账信息"""
        # 1. 优先按设备编号精确匹配
        if asset_no:
            clean_no = str(asset_no).strip()
            if clean_no in self.quick_check_map:
                info = self.quick_check_map[clean_no]
                return {
                    'asset_no': clean_no,
                    'type': '快检',
                    'lab': info.get('lab', '未知实验室'),
                    'group': info.get('lab', '未知实验室'),
                    'inst': info.get('inst', '未查找到'),
                    'serial_no': ''
                }
            if clean_no in self.quantitative_map:
                info = self.quantitative_map[clean_no]
                return {
                    'asset_no': clean_no,
                    'type': '定量',
                    'group': info.get('group', '未知组别'),
                    'lab': '定量实验室',
                    'inst': info.get('inst', '未查找到'),
                    'serial_no': info.get('serial_no', '')
                }
            no_upper = clean_no.upper()
            for k, v in self.quick_check_map.items():
                if k.upper() == no_upper:
                    return {
                        'asset_no': k,
                        'type': '快检',
                        'lab': v.get('lab', '未知实验室'),
                        'group': v.get('lab', '未知实验室'),
                        'inst': v.get('inst', '未查找到'),
                        'serial_no': ''
                    }
            for k, v in self.quantitative_map.items():
                if k.upper() == no_upper:
                    return {
                        'asset_no': k,
                        'type': '定量',
                        'group': v.get('group', '未知组别'),
                        'lab': '定量实验室',
                        'inst': v.get('inst', '未查找到'),
                        'serial_no': v.get('serial_no', '')
                    }

        # 2. 尝试通过出厂编号反查
        if serial_no:
            clean_sn = str(serial_no).strip()
            if clean_sn in self.serial_to_asset:
                target_asset = self.serial_to_asset[clean_sn]
                res = self.query(asset_no=target_asset)
                res['serial_no'] = clean_sn
                return res
            sn_upper = clean_sn.upper()
            for sn_k, asset_k in self.serial_to_asset.items():
                if sn_k.upper() == sn_upper:
                    res = self.query(asset_no=asset_k)
                    res['serial_no'] = clean_sn
                    return res

        # 3. 未在台账查找到
        return {
            'asset_no': asset_no or '未知编号',
            'type': '未知',
            'group': '未知组别',
            'lab': '未知实验室',
            'inst': '未查找到',
            'serial_no': serial_no or ''
        }


# ============================================================
# 第二部分：通用文本处理与日期工具函数
# ============================================================

def sanitize_filename(text):
    """清理文件名中的非法字符"""
    if not text:
        return "未查找到"
    cleaned = re.sub(r'[\\/*?:"<>|\r\n\t]', '-', str(text)).strip()
    cleaned = re.sub(r'-+', '-', cleaned)
    return cleaned if cleaned else "未查找到"

def normalize_date(raw_date):
    """将各种格式的日期（年月日、2026-06-12、2026/06/12 等）规范化为 YYYYMMDD"""
    if not raw_date:
        return '未知校准'
    s = str(raw_date).strip()
    m_cn = re.search(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})', s)
    if m_cn:
        y, m, d = m_cn.groups()
        return f"{y}{int(m):02d}{int(d):02d}"
    m_sep = re.search(r'(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})', s)
    if m_sep:
        y, m, d = m_sep.groups()
        return f"{y}{int(m):02d}{int(d):02d}"
    m_8 = re.search(r'\b(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])\b', s)
    if m_8:
        return m_8.group(0)
    return '未知校准'

def parse_yyyymm(date8):
    """从 8 位日期提取 YYYY、MM(两位)以及 M(一位或两位)"""
    if date8 and re.match(r'^\d{8}$', date8):
        y = date8[:4]
        mm = date8[4:6]
        m = str(int(mm))
        return y, mm, m
    return '未知年', '未知月', '未知月'

def is_text_garbled_or_image(text, min_chinese=4):
    """判断 PDF 提取的文本是否为乱码层或纯图（有效中文字符过少）"""
    if not text or len(text.strip()) < 20:
        return True
    chinese_chars = re.findall(r'[\u4e00-\u9fa5]', text)
    if len(chinese_chars) < min_chinese:
        return True
    return False

def find_value_by_keyword(text, keyword_pattern, value_pattern, search_range=4, col_tolerance_left=20, col_tolerance_right=80):
    """基于布局文本的关键行与相对列位置查找对应的值"""
    lines = text.split('\n')
    for i, line in enumerate(lines):
        kw_match = re.search(keyword_pattern, line, re.IGNORECASE)
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


# ============================================================
# 第三部分：各计量机构规则与解析引擎
# ============================================================

def identify_issuer(text_clean):
    """识别校准报告签发机构"""
    if re.search(r'帝恩|DNTesting|CNASL6483', text_clean, re.IGNORECASE):
        return 'DN'
    if re.search(r'达丰|DAF|DAFTJAX|DAFCJAX|DAFGJAX|CNASL6592', text_clean, re.IGNORECASE):
        return 'DAF'
    if re.search(r'中广测|广州分析测试中心|广东省测试分析研究所|NEM|TiC2600|CNASL7613', text_clean, re.IGNORECASE):
        return 'NEM'
    if re.search(r'深圳市计量质量检测研究院|深圳计量院|smq\.com\.cn|CNASL0579|SZJLY', text_clean, re.IGNORECASE):
        return 'SMQ'
    if re.search(r'中检|CCIC|ccic-mts\.com|CNASL3103', text_clean, re.IGNORECASE):
        return 'CCIC'
    return 'UNKNOWN'

ISSUER_LABEL = {
    'DN':      '东莞帝恩',
    'DAF':     '深圳达丰',
    'NEM':     '广州中广测',
    'SMQ':     '深圳计量院',
    'CCIC':    '中检深圳',
    'UNKNOWN': '未识别机构',
}

ISSUER_SHORT_NAME = {
    'DN':      '帝恩',
    'DAF':     '达丰',
    'NEM':     '中广测',
    'SMQ':     '深圳计量院',
    'CCIC':    '中检',
    'UNKNOWN': '第三方机构',
}

def extract_dn(text, text_clean, ledger: LedgerDatabase):
    """东莞帝恩 (DN) 解析器"""
    asset_no = None
    for asset in ledger.sorted_asset_list:
        if asset in text or asset in text_clean:
            asset_no = asset
            break
    if not asset_no:
        asset_no = find_value_by_keyword(
            text, r'管\s*理\s*号|Asset\s*No\.?',
            r'[A-Z0-9\-]{3,30}',
            col_tolerance_left=10, col_tolerance_right=50
        )
    asset_no = asset_no or '未知编号'

    cal_raw = find_value_by_keyword(
        text, r'校\s*准\s*日\s*期|Date\s*of\s*Calibration',
        r'\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日|\d{4}[-/.]\d{1,2}[-/.]\d{1,2}',
        col_tolerance_left=15, col_tolerance_right=50
    )
    cal_date = normalize_date(cal_raw)

    issue_raw = find_value_by_keyword(
        text, r'签\s*发\s*日\s*期|批\s*准\s*日\s*期|Date\s*of\s*Issue',
        r'\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日|\d{4}[-/.]\d{1,2}[-/.]\d{1,2}',
        col_tolerance_left=15, col_tolerance_right=50
    )
    issue_date = normalize_date(issue_raw)

    info = ledger.query(asset_no=asset_no)
    lab_name = info.get('lab') or '未知实验室'
    inst_name = info.get('inst')
    if not inst_name or inst_name in {'未查找到', 'nan', 'None'}:
        inst_name = find_value_by_keyword(
            text, r'仪\s*器\s*名\s*称|样\s*品\s*名\s*称|Description',
            r'[\u4e00-\u9fa5A-Za-z0-9\(\)（）\-_]{2,30}',
            col_tolerance_right=80
        ) or '未查找到'

    return {
        'asset_no': asset_no,
        'serial_no': '',
        'inst_name': inst_name,
        'lab_name': lab_name,
        'cal_date': cal_date,
        'issue_date': issue_date,
        'group': info.get('group', lab_name),
        'device_type': info.get('type', '快检')
    }

def extract_daf(text, text_clean, ledger: LedgerDatabase):
    """深圳达丰 (DAF) 解析器"""
    # 1. 提取管理号 (Asset No.)
    asset_no = None
    for asset in ledger.sorted_asset_list:
        if re.search(r'\b' + re.escape(asset) + r'\b', text, re.IGNORECASE):
            asset_no = asset
            break
    if not asset_no:
        asset_match = re.search(r'管\s*理\s*号[^\n\rA-Za-z0-9]*([A-Za-z0-9\-/]{2,25})', text)
        if asset_match:
            cand = asset_match.group(1).strip()
            if not re.match(r'^(JJF|JJG|GB|CNAS|DAF)', cand, re.IGNORECASE) and re.search(r'\d', cand):
                asset_no = cand
    if not asset_no:
        asset_no = find_value_by_keyword(
            text, r'管\s*理\s*号|Asset\s*No\.?',
            r'[A-Za-z0-9][A-Za-z0-9\-/]{2,25}',
            col_tolerance_left=5, col_tolerance_right=80
        )
        if asset_no and (re.match(r'^(JJF|JJG|GB|CNAS|DAF)', asset_no, re.IGNORECASE) or not re.search(r'\d', asset_no)):
            asset_no = None

    # 2. 提取出厂编号 (Serial No.)
    serial_no = find_value_by_keyword(
        text, r'出\s*厂\s*编\s*号|Serial\s*No\.?',
        r'[A-Za-z0-9\-_/]{3,30}',
        col_tolerance_left=5, col_tolerance_right=80
    ) or ''

    # 3. 关联台账查验
    info = ledger.query(asset_no=asset_no, serial_no=serial_no)
    if not asset_no or asset_no == '未知编号':
        if info.get('asset_no') and info.get('asset_no') != '未知编号':
            asset_no = info.get('asset_no')

    # 4. 提取仪器名称 (过滤表头与联络信息)
    inst_name = None
    if info.get('inst') and info.get('inst') not in {'未查找到', 'nan', 'None'}:
        inst_name = info.get('inst')
    else:
        inst_raw = find_value_by_keyword(
            text, r'仪\s*器\s*名\s*称|样\s*品\s*名\s*称|Description',
            r'[\u4e00-\u9fa5][\u4e00-\u9fa5A-Za-z0-9\(\)（）\-_]{1,30}',
            col_tolerance_right=100
        )
        blacklist = {'型号规格', '型号', '规格', '制造商', '制造厂商', '出厂编号', 'Serial', '管理号', 'Asset', '联络信息', 'Information', 'Model', 'Type', 'Manufacturer'}
        if inst_raw and not any(b in inst_raw for b in blacklist):
            inst_name = inst_raw
        else:
            inst_name = '未查找到'

    # 5. 校准日期
    cal_raw = find_value_by_keyword(
        text, r'校\s*准\s*日\s*期|Date\s*of\s*Calibration',
        r'\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日|\d{4}[-/.]\d{1,2}[-/.]\d{1,2}',
        col_tolerance_left=10, col_tolerance_right=60
    )
    if not cal_raw:
        cal_match = re.search(r'校\s*准\s*日\s*期[^\n\r\d]*(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)', text)
        if cal_match:
            cal_raw = cal_match.group(1)
    cal_date = normalize_date(cal_raw)

    return {
        'asset_no': asset_no or '未知编号',
        'serial_no': serial_no or info.get('serial_no', ''),
        'inst_name': inst_name,
        'cal_date': cal_date,
        'group': info.get('group', '未知组别'),
        'lab_name': info.get('lab', '未知实验室'),
        'device_type': info.get('type', '定量')
    }

def extract_nem(text, text_clean, ledger: LedgerDatabase):
    """广州中广测 (NEM) 解析器"""
    serial_no = find_value_by_keyword(
        text, r'出\s*厂\s*编\s*号|器\s*具\s*编\s*号|Serial\s*No\.?',
        r'[A-Za-z0-9\-_/]{2,30}',
        col_tolerance_left=10, col_tolerance_right=60
    )
    if not serial_no:
        for asset in ledger.sorted_asset_list:
            if asset in text:
                serial_no = asset
                break
    serial_no = serial_no or '未知出厂号'

    asset_no = find_value_by_keyword(
        text, r'管\s*理\s*号|Asset\s*No\.?',
        r'[A-Za-z0-9\-_/]{2,30}',
        col_tolerance_left=10, col_tolerance_right=60
    )

    info = ledger.query(asset_no=asset_no, serial_no=serial_no)
    if not asset_no and info.get('asset_no') != '未知编号':
        asset_no = info.get('asset_no')

    inst_name = info.get('inst')
    if not inst_name or inst_name in {'未查找到', 'nan', 'None'}:
        inst_name = find_value_by_keyword(
            text, r'器\s*具\s*名\s*称|仪\s*器\s*名\s*称|样\s*品\s*名\s*称|Description',
            r'[\u4e00-\u9fa5][\u4e00-\u9fa5A-Za-z0-9\(\)（）\-_]{1,30}',
            col_tolerance_right=80
        ) or '未查找到'

    cal_raw = find_value_by_keyword(
        text, r'校\s*准\s*日\s*期|Date\s*of\s*Calibration',
        r'\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日|\d{4}[-/.]\d{1,2}[-/.]\d{1,2}',
        col_tolerance_left=10, col_tolerance_right=60
    )
    cal_date = normalize_date(cal_raw)

    return {
        'asset_no': asset_no or serial_no,
        'serial_no': serial_no,
        'inst_name': inst_name,
        'cal_date': cal_date,
        'group': info.get('group', '未知组别'),
        'lab_name': info.get('lab', '定量实验室'),
        'device_type': info.get('type', '定量')
    }

def extract_smq(text, text_clean, ledger: LedgerDatabase):
    """深圳计量院 (SMQ) 解析器"""
    asset_no = None
    for asset in ledger.sorted_asset_list:
        if asset in text:
            asset_no = asset
            break
    if not asset_no:
        asset_no = find_value_by_keyword(
            text, r'管\s*理\s*号|器\s*具\s*编\s*号|Asset\s*No\.?',
            r'[A-Za-z0-9\-_/]{2,30}',
            col_tolerance_left=10, col_tolerance_right=60
        )
    asset_no = asset_no or '未知编号'

    serial_no = find_value_by_keyword(
        text, r'出\s*厂\s*编\s*号|Serial\s*No\.?',
        r'[A-Za-z0-9\-_/]{2,30}',
        col_tolerance_left=10, col_tolerance_right=60
    ) or ''

    info = ledger.query(asset_no=asset_no, serial_no=serial_no)

    inst_name = info.get('inst')
    if not inst_name or inst_name in {'未查找到', 'nan', 'None'}:
        inst_name = find_value_by_keyword(
            text, r'器\s*具\s*名\s*称|仪\s*器\s*名\s*称|Description',
            r'[\u4e00-\u9fa5][\u4e00-\u9fa5A-Za-z0-9\(\)（）\-_]{1,30}',
            col_tolerance_right=80
        ) or '未查找到'

    cal_raw = find_value_by_keyword(
        text, r'校\s*准\s*日\s*期|检\s*定\s*日\s*期|Date\s*of\s*Calibration',
        r'\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日|\d{4}[-/.]\d{1,2}[-/.]\d{1,2}',
        col_tolerance_left=10, col_tolerance_right=60
    )
    cal_date = normalize_date(cal_raw)

    return {
        'asset_no': asset_no,
        'serial_no': serial_no,
        'inst_name': inst_name,
        'cal_date': cal_date,
        'group': info.get('group', '未知组别'),
        'lab_name': info.get('lab', '定量实验室'),
        'device_type': info.get('type', '定量')
    }

def extract_ccic(text, text_clean, ledger: LedgerDatabase):
    """中检深圳 (CCIC) 解析器"""
    asset_no = None
    for asset in ledger.sorted_asset_list:
        if asset in text:
            asset_no = asset
            break
    if not asset_no:
        asset_no = find_value_by_keyword(
            text, r'管\s*理\s*号|样\s*品\s*编\s*号|Asset\s*No\.?',
            r'[A-Za-z0-9\-_/]{2,30}',
            col_tolerance_left=10, col_tolerance_right=60
        )
    asset_no = asset_no or '未知编号'

    serial_no = find_value_by_keyword(
        text, r'出\s*厂\s*编\s*号|Serial\s*No\.?',
        r'[A-Za-z0-9\-_/]{2,30}',
        col_tolerance_left=10, col_tolerance_right=60
    ) or ''

    info = ledger.query(asset_no=asset_no, serial_no=serial_no)

    inst_name = info.get('inst')
    if not inst_name or inst_name in {'未查找到', 'nan', 'None'}:
        inst_name = find_value_by_keyword(
            text, r'仪\s*器\s*名\s*称|样\s*品\s*名\s*称|Description',
            r'[\u4e00-\u9fa5][\u4e00-\u9fa5A-Za-z0-9\(\)（）\-_]{1,30}',
            col_tolerance_right=80
        ) or '未查找到'

    cal_raw = find_value_by_keyword(
        text, r'校\s*准\s*日\s*期|Date\s*of\s*Calibration',
        r'\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日|\d{4}[-/.]\d{1,2}[-/.]\d{1,2}',
        col_tolerance_left=10, col_tolerance_right=60
    )
    cal_date = normalize_date(cal_raw)

    return {
        'asset_no': asset_no,
        'serial_no': serial_no,
        'inst_name': inst_name,
        'cal_date': cal_date,
        'group': info.get('group', '未知组别'),
        'lab_name': info.get('lab', '定量实验室'),
        'device_type': info.get('type', '定量')
    }

def extract_generic(text, text_clean, ledger: LedgerDatabase):
    """通用第三方机构解析器"""
    asset_no = None
    for asset in ledger.sorted_asset_list:
        if asset in text:
            asset_no = asset
            break
    if not asset_no:
        asset_no = find_value_by_keyword(
            text, r'管\s*理\s*号|器\s*具\s*编\s*号|设\s*备\s*编\s*号|Asset\s*No\.?',
            r'[A-Za-z0-9\-_/]{2,30}',
            col_tolerance_left=10, col_tolerance_right=60
        )
    asset_no = asset_no or '未知编号'

    serial_no = find_value_by_keyword(
        text, r'出\s*厂\s*编\s*号|Serial\s*No\.?',
        r'[A-Za-z0-9\-_/]{2,30}',
        col_tolerance_left=10, col_tolerance_right=60
    ) or ''

    info = ledger.query(asset_no=asset_no, serial_no=serial_no)

    inst_name = info.get('inst')
    if not inst_name or inst_name in {'未查找到', 'nan', 'None'}:
        inst_name = find_value_by_keyword(
            text, r'仪\s*器\s*名\s*称|器\s*具\s*名\s*称|样\s*品\s*名\s*称',
            r'[\u4e00-\u9fa5][\u4e00-\u9fa5A-Za-z0-9\(\)（）\-_]{1,30}',
            col_tolerance_right=80
        ) or '未查找到'

    cal_raw = find_value_by_keyword(
        text, r'校\s*准\s*日\s*期|检\s*定\s*日\s*期',
        r'\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日|\d{4}[-/.]\d{1,2}[-/.]\d{1,2}',
        col_tolerance_left=10, col_tolerance_right=60
    )
    cal_date = normalize_date(cal_raw)

    return {
        'asset_no': asset_no,
        'serial_no': serial_no,
        'inst_name': inst_name,
        'cal_date': cal_date,
        'group': info.get('group', '未知组别'),
        'lab_name': info.get('lab', '未知实验室'),
        'device_type': info.get('type', '未知')
    }


# ============================================================
# 第四部分：命名规则与归档路径构建器（深度适配真实生产体系）
# ============================================================

def build_new_filename(issuer, fields):
    """
    根据机构、设备类型与规范构建新文件名：
    1. 帝恩快检：{实验室/项目组}_{设备编号}_{校准日期}_{签发日期}_{设备名称}.pdf
    2. 移液器(YYQ)且有出厂号：{出厂编号}({设备编号})_{设备名称}_{校准日期}_{所属组别}.pdf
    3. 标准定量设备：{设备编号}_{设备名称}_{校准日期}_{所属组别}.pdf
    """
    device_type = fields.get('device_type', '未知')
    asset_no = fields.get('asset_no', '未知编号')
    serial_no = fields.get('serial_no', '')
    inst_name = fields.get('inst_name', '未查找到')
    cal_date = fields.get('cal_date', '未知校准')
    group = fields.get('group', '未知组别')

    # 1. 帝恩快检设备
    if issuer == 'DN' or device_type == '快检':
        lab_name = fields.get('lab_name', '未知实验室')
        issue_date = fields.get('issue_date', cal_date)
        parts = [lab_name, asset_no, cal_date, issue_date, inst_name]
        return '_'.join([sanitize_filename(p) for p in parts])

    # 2. 移液器复合编号（中广测等报告中常见的出厂号与YYQ编号复合格式）
    if str(asset_no).startswith('YYQ') and serial_no and serial_no != '未知出厂号' and serial_no != asset_no:
        code_tag = f"{serial_no}({asset_no})"
        parts = [code_tag, inst_name, cal_date, group]
        return '_'.join([sanitize_filename(p) for p in parts])

    # 3. 标准定量设备命名
    code = asset_no if asset_no != '未知编号' else (serial_no or '未知编号')
    parts = [code, inst_name, cal_date, group]
    return '_'.join([sanitize_filename(p) for p in parts])


def build_archive_dir(archive_root, issuer, fields):
    """
    构建多层级归档目录路径，深度对齐现行生产归档体系：
    - 定量：{archive_root}/{YYYY}年检定校准证书/{YYYY}年{M}月{机构}校准证书/{YYYY}年{M}月{组别}校准证书-{机构}/
    - 快检：{archive_root}/快检设备/{YYYY}年{M}月{项目组}设备校准证书/
    """
    device_type = fields.get('device_type', '未知')
    cal_date = fields.get('cal_date', '')
    year, mm, m = parse_yyyymm(cal_date)
    issuer_name = ISSUER_SHORT_NAME.get(issuer, '第三方')

    if device_type == '快检':
        lab_name = sanitize_filename(fields.get('lab_name', '未知实验室'))
        if not lab_name.endswith('项目组') and not lab_name.endswith('街道') and not lab_name.endswith('组'):
            folder_tag = f"{lab_name}项目组"
        else:
            folder_tag = lab_name
        month_folder = f"{year}年{m}月{folder_tag}设备校准证书"
        return os.path.join(archive_root, '快检设备', month_folder)

    elif device_type == '定量':
        group = sanitize_filename(fields.get('group', '未知组别'))
        # 生产规范层级：
        # 1级: 2026年检定校准证书
        # 2级: 2026年4月达丰校准证书
        # 3级: 2026年4月气相组校准证书-达丰 (或 2026年04月气相组校准证书)
        year_folder = f"{year}年检定校准证书"
        month_issuer_folder = f"{year}年{m}月{issuer_name}校准证书"
        detail_folder = f"{year}年{m}月{group}校准证书-{issuer_name}"
        return os.path.join(archive_root, year_folder, month_issuer_folder, detail_folder)

    else:
        # 未分类设备
        return os.path.join(archive_root, '未分类设备', f"{year}年{m}月校准证书")


# ============================================================
# 第五部分：文件雷达扫描与 ZIP 智能解压
# ============================================================

def safe_extract_zip(zip_path, target_extract_dir):
    """解压 ZIP 压缩包，自动纠正 Windows 下 GBK/CP437 编码文件名"""
    extracted_pdf_paths = []
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            for info in zf.infolist():
                if info.is_dir():
                    continue
                raw_name = info.filename
                try:
                    correct_name = raw_name.encode('cp437').decode('gbk')
                except Exception:
                    try:
                        correct_name = raw_name.encode('cp437').decode('utf-8')
                    except Exception:
                        correct_name = raw_name

                if correct_name.lower().endswith('.pdf'):
                    dest_file_path = os.path.join(target_extract_dir, os.path.basename(correct_name))
                    os.makedirs(os.path.dirname(dest_file_path), exist_ok=True)
                    with zf.open(info) as src, open(dest_file_path, 'wb') as dst:
                        shutil.copyfileobj(src, dst)
                    extracted_pdf_paths.append(dest_file_path)
    except Exception as e:
        print(f"⚠️ 解压 ZIP [{os.path.basename(zip_path)}] 失败: {e}")
    return extracted_pdf_paths


def scan_directory(search_dir, archive_root_name="【归档完成】", auto_unzip=True):
    """
    深度扫描目录中的所有 PDF 文件（包含直接文件、子目录以及 ZIP 压缩包中的 PDF）
    自动排除已归档的目录
    """
    print("\n" + "=" * 70)
    print(f"【步骤 2/4】正在执行文件雷达扫描...")
    print(f"   扫描根目录：{os.path.abspath(search_dir)}")
    print("=" * 70)

    pdf_files_to_process = []
    zip_files_found = []
    ignored_archive_count = 0

    abs_archive_root = os.path.abspath(archive_root_name)

    for root, dirs, files in os.walk(search_dir):
        abs_root = os.path.abspath(root)
        # 智能排除已归档文件夹与临时文件夹
        if '【归档完成】' in root or '校准证书库' in root or abs_root.startswith(abs_archive_root):
            ignored_archive_count += sum(1 for f in files if f.lower().endswith('.pdf'))
            continue
        if '.git' in root or '__pycache__' in root or '_temp_zip_' in root:
            continue

        for filename in files:
            lower_name = filename.lower()
            file_path = os.path.join(root, filename)
            if lower_name.endswith('.pdf'):
                pdf_files_to_process.append(file_path)
            elif lower_name.endswith('.zip'):
                zip_files_found.append(file_path)

    # 自动处理 ZIP 压缩包
    if auto_unzip and zip_files_found:
        print(f"📦 发现 {len(zip_files_found)} 个 ZIP 压缩包，正在自动解包解析...")
        temp_unzip_root = os.path.join(search_dir, "_temp_zip_extracted_")
        for zip_p in zip_files_found:
            zip_basename = os.path.splitext(os.path.basename(zip_p))[0]
            extract_folder = os.path.join(temp_unzip_root, zip_basename)
            os.makedirs(extract_folder, exist_ok=True)
            extracted_pdfs = safe_extract_zip(zip_p, extract_folder)
            print(f"   解压 [{os.path.basename(zip_p)}] -> 提取出 {len(extracted_pdfs)} 份 PDF 证书")
            pdf_files_to_process.extend(extracted_pdfs)

    unique_pdf_list = list(dict.fromkeys(pdf_files_to_process))

    print(f"📡 雷达扫描完毕：共捕获 {len(unique_pdf_list)} 份待处理的 PDF 校准证书！")
    if ignored_archive_count > 0:
        print(f"   （自动跳过已归档库中的 {ignored_archive_count} 份证书）")

    return unique_pdf_list


# ============================================================
# 第六部分：PDF 双引擎解析核心流程
# ============================================================

def parse_pdf_document(file_path, ledger: LedgerDatabase, enable_ocr=True):
    """
    极速解析单个 PDF 文件：
    1. PyMuPDF 极速读取第 1~3 页。
    2. 检测是否为纯图或乱码层，按需触发 RapidOCR 视觉引擎抢救。
    3. 识别机构并提取结构化字段，结合台账进行智能纠错与补全。
    """
    filename = os.path.basename(file_path)
    text = ""
    used_ocr = False

    try:
        with fitz.open(file_path) as doc:
            pages_to_scan = min(3, len(doc))
            for i in range(pages_to_scan):
                page = doc[i]
                page_text = page.get_text("text") or ''

                if is_text_garbled_or_image(page_text, min_chinese=4) and enable_ocr and HAS_OCR and ocr_engine:
                    pix = page.get_pixmap(dpi=200)
                    img_bytes = pix.tobytes("png")
                    try:
                        ocr_res, _ = ocr_engine(img_bytes)
                        if ocr_res:
                            ocr_text = '\n'.join([line[1] for line in ocr_res])
                            page_text = ocr_text
                            used_ocr = True
                    except Exception:
                        pass

                text += page_text + '\n'
    except Exception as e:
        return {'status': 'error', 'error': f"PDF 读取失败: {e}", 'filename': filename}

    text_clean = text.replace(' ', '').replace('\n', '')

    if (len(text_clean.strip()) < 30 or identify_issuer(text_clean) == 'UNKNOWN' or is_text_garbled_or_image(text, min_chinese=6)) and enable_ocr and HAS_OCR and ocr_engine and not used_ocr:
        ocr_text_total = ""
        try:
            with fitz.open(file_path) as doc:
                for i in range(min(2, len(doc))):
                    page = doc[i]
                    pix = page.get_pixmap(dpi=200)
                    ocr_res, _ = ocr_engine(pix.tobytes("png"))
                    if ocr_res:
                        ocr_text_total += '\n' + '\n'.join([line[1] for line in ocr_res])
            if ocr_text_total:
                text = ocr_text_total
                text_clean = text.replace(' ', '').replace('\n', '')
                used_ocr = True
        except Exception:
            pass

    issuer = identify_issuer(text_clean)

    if issuer == 'DN': fields = extract_dn(text, text_clean, ledger)
    elif issuer == 'DAF': fields = extract_daf(text, text_clean, ledger)
    elif issuer == 'NEM': fields = extract_nem(text, text_clean, ledger)
    elif issuer == 'SMQ': fields = extract_smq(text, text_clean, ledger)
    elif issuer == 'CCIC': fields = extract_ccic(text, text_clean, ledger)
    else: fields = extract_generic(text, text_clean, ledger)

    if (fields.get('asset_no') in {'未知编号', '未查找到'} or fields.get('inst_name') == '未查找到') and enable_ocr and HAS_OCR and ocr_engine and not used_ocr:
        try:
            ocr_text_total = ""
            with fitz.open(file_path) as doc:
                for i in range(min(2, len(doc))):
                    page = doc[i]
                    pix = page.get_pixmap(dpi=200)
                    ocr_res, _ = ocr_engine(pix.tobytes("png"))
                    if ocr_res:
                        ocr_text_total += '\n' + '\n'.join([line[1] for line in ocr_res])
            if ocr_text_total:
                text = ocr_text_total
                text_clean = text.replace(' ', '').replace('\n', '')
                issuer = identify_issuer(text_clean)
                if issuer == 'DN': fields = extract_dn(text, text_clean, ledger)
                elif issuer == 'DAF': fields = extract_daf(text, text_clean, ledger)
                elif issuer == 'NEM': fields = extract_nem(text, text_clean, ledger)
                elif issuer == 'SMQ': fields = extract_smq(text, text_clean, ledger)
                elif issuer == 'CCIC': fields = extract_ccic(text, text_clean, ledger)
                else: fields = extract_generic(text, text_clean, ledger)
                used_ocr = True
        except Exception:
            pass

    return {
        'status': 'success',
        'issuer': issuer,
        'fields': fields,
        'used_ocr': used_ocr,
        'filename': filename,
        'file_path': file_path
    }


# ============================================================
# 第七部分：主执行循环与归档引擎
# ============================================================

def export_styled_excel(df, output_excel_path):
    """使用 openpyxl 导出具有商务美化样式的 Excel 汇总台账"""
    df.to_excel(output_excel_path, index=False, engine='openpyxl')
    
    import openpyxl
    wb = openpyxl.load_workbook(output_excel_path)
    ws = wb.active
    ws.title = "校准归档明细汇总"

    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="微软雅黑", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )
    zebra_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")

    ws.row_dimensions[1].height = 28

    for col_idx in range(1, len(df.columns) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border

    for row_idx in range(2, len(df) + 2):
        ws.row_dimensions[row_idx].height = 22
        is_zebra = (row_idx % 2 == 1)
        for col_idx in range(1, len(df.columns) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            if is_zebra:
                cell.fill = zebra_fill
            col_name = df.columns[col_idx - 1]
            if col_name in {'序号', '机构', '校准日期', '设备类型', '解析方式'}:
                cell.alignment = center_align
            else:
                cell.alignment = left_align

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            width = sum(2 if ord(char) > 127 else 1 for char in val_str)
            if width > max_len:
                max_len = width
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(output_excel_path)


def run_archive_process(target_dir='.', archive_root='./【归档完成】校准证书库',
                        copy_mode=False, dry_run=False, enable_ocr=True, auto_unzip=True):
    start_time = datetime.datetime.now()
    print("=" * 70)
    print(f"  {TOOL_TITLE}")
    print(f"  AI 视觉引擎：{ocr_status_desc}")
    print(f"  运行模式：{'🔍 试运行预览 (Dry-Run)' if dry_run else ('📑 复制归档 (Copy)' if copy_mode else '🚚 移动归档 (Move)')}")
    print("=" * 70)

    # 1. 加载台账
    ledger = LedgerDatabase()
    ledger.load(base_dir=target_dir)

    # 2. 扫描文件
    pdf_list = scan_directory(target_dir, archive_root_name=archive_root, auto_unzip=auto_unzip)
    if not pdf_list:
        print("\n🛑 运行结束：未在指定路径下找到任何待处理的 PDF 校准证书。")
        return

    print("\n" + "=" * 70)
    print("【步骤 3/4】开始并行智能提取与多层级归档...")
    print("=" * 70)

    if not dry_run:
        os.makedirs(archive_root, exist_ok=True)

    records = []
    success_count = 0
    error_count = 0

    for idx, pdf_path in enumerate(pdf_list, 1):
        filename = os.path.basename(pdf_path)
        progress_str = f"[{idx}/{len(pdf_list)}]"

        res = parse_pdf_document(pdf_path, ledger, enable_ocr=enable_ocr)
        if res.get('status') != 'success':
            print(f"❌ {progress_str} 处理失败 [{filename}]: {res.get('error')}")
            error_count += 1
            continue

        issuer = res['issuer']
        fields = res['fields']
        used_ocr = res['used_ocr']
        label = ISSUER_LABEL.get(issuer, '未知机构')
        device_type = fields.get('device_type', '未知')

        # 构建新文件名与目标路径
        new_filename = f"{build_new_filename(issuer, fields)}.pdf"
        target_dir_path = build_archive_dir(archive_root, issuer, fields)

        if not dry_run:
            os.makedirs(target_dir_path, exist_ok=True)

        target_file_path = os.path.join(target_dir_path, new_filename)

        # 同名冲突处理机制：自动追加 _1, _2 序号
        conflict_suffix = 1
        base_stem, ext = os.path.splitext(new_filename)
        while os.path.exists(target_file_path) and os.path.abspath(target_file_path) != os.path.abspath(pdf_path):
            new_filename = f"{base_stem}_{conflict_suffix}{ext}"
            target_file_path = os.path.join(target_dir_path, new_filename)
            conflict_suffix += 1

        action_desc = "预览"
        if not dry_run:
            try:
                if copy_mode:
                    shutil.copy2(pdf_path, target_file_path)
                    action_desc = "已复制"
                else:
                    shutil.move(pdf_path, target_file_path)
                    action_desc = "已归档"
            except Exception as move_err:
                print(f"❌ {progress_str} 文件操作失败 [{filename}]: {move_err}")
                error_count += 1
                continue

        ocr_tag = " (⚡AI视觉)" if used_ocr else ""
        rel_archive_path = os.path.relpath(target_dir_path, archive_root).replace('\\', '/') if not dry_run else target_dir_path.replace('\\', '/')
        print(f"✅ {progress_str} [{label}] [{device_type}]{ocr_tag}")
        print(f"   原文件：{filename}")
        print(f"   新文件：{new_filename}")
        print(f"   归档至：{rel_archive_path}/")

        success_count += 1

        code_val = fields.get('asset_no') if fields.get('asset_no') != '未知编号' else fields.get('serial_no', '—')
        records.append({
            '序号': success_count,
            '原始文件名': filename,
            '最终重命名': new_filename,
            '机构': label,
            '设备编号': code_val,
            '出厂编号': fields.get('serial_no', '—'),
            '仪器名称': fields.get('inst_name', '未查找到'),
            '校准日期': fields.get('cal_date', '未知校准'),
            '所属组别/实验室': fields.get('group') or fields.get('lab_name', '—'),
            '设备类型': device_type,
            '归档目标目录': rel_archive_path,
            '完整目标路径': os.path.abspath(target_file_path),
            '解析方式': 'RapidOCR 视觉识别' if used_ocr else 'PyMuPDF 文本层'
        })

    # 清理临时解压目录
    temp_zip_dir = os.path.join(target_dir, "_temp_zip_extracted_")
    if os.path.exists(temp_zip_dir):
        try:
            shutil.rmtree(temp_zip_dir)
        except Exception:
            pass

    # ==========================================
    # 第八部分：输出 Excel 与 TXT 统计报表
    # ==========================================
    print("\n" + "=" * 70)
    print("【步骤 4/4】生成本次处理统计汇总与报表...")
    print("=" * 70)

    now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    time_cost = (datetime.datetime.now() - start_time).total_seconds()

    if records:
        df_result = pd.DataFrame(records)
        excel_out_name = f'本次扫描_校准信息归档汇总表_{VERSION}.xlsx'
        export_styled_excel(df_result, excel_out_name)

        txt_out_name = f'本次处理_统计汇总_{now_str}.txt'
        with open(txt_out_name, 'w', encoding='utf-8') as tf:
            tf.write("=" * 70 + "\n")
            tf.write(f"校准证书智能归档统计汇总清单 ({VERSION})\n")
            tf.write(f"处理时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            tf.write(f"耗时：{time_cost:.2f} 秒\n")
            tf.write(f"归档总库根路径：{os.path.abspath(archive_root)}\n")
            tf.write("=" * 70 + "\n\n")
            tf.write(f"总处理证书数量：{len(records)} 份\n")
            tf.write(f"  - 快检设备证书：{sum(1 for r in records if r['设备类型'] == '快检')} 份\n")
            tf.write(f"  - 定量设备证书：{sum(1 for r in records if r['设备类型'] == '定量')} 份\n")
            tf.write(f"  - 未分类设备证书：{sum(1 for r in records if r['设备类型'] == '未知')} 份\n\n")

            tf.write("📊 机构分布统计：\n")
            issuer_counts = defaultdict(int)
            for r in records:
                issuer_counts[r['机构']] += 1
            for k, v in sorted(issuer_counts.items(), key=lambda x: -x[1]):
                tf.write(f"  - {k}: {v} 份\n")

            tf.write("\n📂 详细归档文件清单：\n")
            folder_groups = defaultdict(list)
            for r in records:
                folder_groups[r['归档目标目录']].append(r)
            for folder, items in sorted(folder_groups.items()):
                tf.write(f"\n📁 【{folder}】 (共 {len(items)} 份)\n")
                for it in items:
                    tf.write(f"    - {it['最终重命名']}  [原名: {it['原始文件名']}]\n")

        print(f"\n🎉 处理大功告成！耗时: {time_cost:.2f} 秒")
        print(f"📊 本次共处理: {len(records)} 份校准证书")
        print(f"   - 快检设备: {sum(1 for r in records if r['设备类型'] == '快检')} 份")
        print(f"   - 定量设备: {sum(1 for r in records if r['设备类型'] == '定量')} 份")
        print(f"   - 未分类:   {sum(1 for r in records if r['设备类型'] == '未知')} 份")
        print(f"📁 归档根目录：{os.path.abspath(archive_root)}")
        print(f"📋 Excel 汇总表：{excel_out_name}")
        print(f"📄 TXT 统计清单：{txt_out_name}")
    else:
        print(f"\n⚠️ 本次未成功归档任何文件（失败 {error_count} 份）。")


# ============================================================
# 入口点
# ============================================================

def main():
    parser = argparse.ArgumentParser(description=TOOL_TITLE)
    parser.add_argument('-p', '--path', default=None, help='指定要扫描和处理的目标文件夹路径')
    parser.add_argument('-o', '--output', default='./【归档完成】校准证书库', help='指定归档目标根目录')
    parser.add_argument('--copy', action='store_true', help='使用复制模式代替移动模式（保留原文件）')
    parser.add_argument('--dry-run', action='store_true', help='试运行预览模式（仅提取和规划，不实际移动/复制文件）')
    parser.add_argument('--no-ocr', action='store_true', help='关闭 OCR 视觉引擎')
    parser.add_argument('--no-unzip', action='store_true', help='关闭 ZIP 压缩包自动解包功能')

    args = parser.parse_args()

    target_path = args.path
    if not target_path:
        if sys.stdin.isatty():
            print("=" * 70)
            print(f"  {TOOL_TITLE}")
            print("=" * 70)
            user_input = input("请输入或拖入待处理文件夹路径（直接按回车默认处理当前目录）：").strip()
            user_input = user_input.strip('"\'')
            target_path = user_input if user_input else '.'
        else:
            target_path = '.'

    run_archive_process(
        target_dir=target_path,
        archive_root=args.output,
        copy_mode=args.copy,
        dry_run=args.dry_run,
        enable_ocr=not args.no_ocr,
        auto_unzip=not args.no_unzip
    )

if __name__ == '__main__':
    main()
