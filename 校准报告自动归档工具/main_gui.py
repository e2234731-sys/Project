# -*- coding: utf-8 -*-
"""
FQT 实验室校准证书智能归档与量值溯源工作台 V2.1
======================================================================
1. 【现代化双轨工作台架构】
   - 🏢 17025定量实验室设备校准工作台
   - 🧪 快检质量网络驻点设备校准工作台
   - 📊 计划台账闭环核对与审计看板 (已校准 / 缺漏待收回 / 计划外新增)
   - 📜 系统运行实时日志
   - ⚙️ 常用预设与环境配置
2. 【高精解析与全机构覆盖】
   - 深度识别东莞帝恩 (DN)、广州中广测 (NEM)、深圳达丰 (DAF)、深圳计量院 (SMQ)、中检深圳 (CCIC)
   - 自动解析复合出厂编号（如 2013C004-31(H0406) -> 编号: H0406, 出厂号: 2013C004-31）
   - 自动解析移液枪复合标识（如 18F47863(YYQ132) -> 设备号: YYQ132）
   - 智能提取签发日期与校准日期，规范重命名快检与定量报告
3. 【全工作表台账联动与闭环比对】
   - 自动扫描并载入《2026年年度计划汇总.xlsx》与《2026年各实验室仪器设备校准清单.xlsx》
   - 覆盖 1060+ 台仪器资产，支持台账计划 vs. 实际证书的差异比对与一键导出
4. 【底层稳定性与纯单文件打包】
   - 彻底修复 OpenBLAS 线程限制 (OPENBLAS_NUM_THREADS=1)
   - PyMuPDF 文本层优先解析 + RapidOCR 离线模型自愈备份
======================================================================
"""

import os
import sys

# ── 1. 核心底层环境初始化 (必须在任何科学计算库导入前执行)
os.environ["OPENBLAS_NUM_THREADS"] = "1"
os.environ["MKL_NUM_THREADS"] = "1"
os.environ["OMP_NUM_THREADS"] = "1"
os.environ["NUMEXPR_NUM_THREADS"] = "1"
os.environ["VECLIB_MAXIMUM_THREADS"] = "1"

import re
import json
import shutil
import zipfile
import datetime
import importlib
import traceback
import ctypes
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

# ── 2. 动态注册 DLL 搜索目录并预加载底层 C/C++ 动态链接库
def setup_dll_directories():
    search_dirs = []
    if hasattr(sys, '_MEIPASS'):
        base = sys._MEIPASS
        search_dirs.extend([
            base,
            os.path.join(base, 'onnxruntime', 'capi'),
            os.path.join(base, 'pymupdf'),
            os.path.join(base, 'fitz'),
            os.path.join(base, 'cv2')
        ])
    else:
        app_dir = os.path.dirname(os.path.abspath(__file__))
        search_dirs.append(app_dir)

    for p in search_dirs:
        if os.path.exists(p):
            try:
                os.add_dll_directory(p)
            except Exception:
                pass

    if search_dirs:
        os.environ['PATH'] = ';'.join([d for d in search_dirs if os.path.exists(d)]) + ';' + os.environ.get('PATH', '')

    if hasattr(sys, '_MEIPASS'):
        capi_dir = os.path.join(sys._MEIPASS, 'onnxruntime', 'capi')
        if os.path.exists(capi_dir):
            for dll_name in ['onnxruntime.dll', 'onnxruntime_providers_shared.dll', 'onnxruntime_pybind11_state.pyd']:
                fp = os.path.join(capi_dir, dll_name)
                if os.path.exists(fp):
                    try:
                        ctypes.windll.kernel32.LoadLibraryExW(fp, 0, 8)
                    except Exception:
                        pass

setup_dll_directories()

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import pandas as pd

# 稳健导入 PyMuPDF / fitz
try:
    import fitz
    try:
        fitz.TOOLS.mupdf_display_errors(False)
    except Exception:
        pass
except Exception:
    try:
        import pymupdf as fitz
        try:
            fitz.TOOLS.mupdf_display_errors(False)
        except Exception:
            pass
    except Exception:
        fitz = None

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QLineEdit, QPushButton, QCheckBox,
    QRadioButton, QButtonGroup, QProgressBar, QTextEdit, QTableWidget,
    QTableWidgetItem, QHeaderView, QTabWidget, QGroupBox, QFileDialog,
    QMessageBox, QFrame, QDialog, QListWidget, QListWidgetItem,
    QInputDialog, QMenu, QAction, QSizePolicy, QScrollArea, QStackedWidget,
    QComboBox, QSplitter
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize, QUrl
from PyQt5.QtGui import QFont, QColor, QIcon, QCursor, QDesktopServices

# ── 定位真实运行根目录 (支持独立可执行文件与源码运行环境)
def get_app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))

APP_ROOT = get_app_dir()
PRESETS_FILE = os.path.join(APP_ROOT, "config_presets.json")
UI_SETTINGS_FILE = os.path.join(APP_ROOT, "ui_settings.json")

# ── 全自愈加载 RapidOCR 视觉引擎
def init_rapidocr():
    try:
        import rapidocr_onnxruntime
        import rapidocr_onnxruntime.ch_ppocr_v2_cls as _cls_mod
        import rapidocr_onnxruntime.ch_ppocr_v3_det as _det_mod
        import rapidocr_onnxruntime.ch_ppocr_v3_rec as _rec_mod
        from rapidocr_onnxruntime.rapid_ocr_api import RapidOCR

        def _safe_init_module(module_name, class_name):
            if 'det' in module_name.lower():
                return getattr(_det_mod, class_name, getattr(_det_mod, 'TextDetector', None))
            elif 'cls' in module_name.lower():
                return getattr(_cls_mod, class_name, getattr(_cls_mod, 'TextClassifier', None))
            elif 'rec' in module_name.lower():
                return getattr(_rec_mod, class_name, getattr(_rec_mod, 'TextRecognizer', None))
            try:
                mod = importlib.import_module(module_name)
                return getattr(mod, class_name)
            except Exception:
                mod = importlib.import_module(f'rapidocr_onnxruntime.{module_name}')
                return getattr(mod, class_name)

        RapidOCR.init_module = staticmethod(_safe_init_module)

        sys.modules['ch_ppocr_v2_cls'] = _cls_mod
        sys.modules['ch_ppocr_v3_det'] = _det_mod
        sys.modules['ch_ppocr_v3_rec'] = _rec_mod

        base_dir = getattr(sys, '_MEIPASS', None) or os.path.dirname(rapidocr_onnxruntime.__file__)
        model_candidates = [
            os.path.join(base_dir, 'rapidocr_onnxruntime', 'models'),
            os.path.join(base_dir, 'models'),
            os.path.join(os.path.dirname(rapidocr_onnxruntime.__file__), 'models')
        ]
        det_p, rec_p, cls_p = None, None, None
        for mdir in model_candidates:
            if os.path.exists(mdir):
                dp = os.path.join(mdir, 'ch_PP-OCRv3_det_infer.onnx')
                rp = os.path.join(mdir, 'ch_PP-OCRv3_rec_infer.onnx')
                cp = os.path.join(mdir, 'ch_ppocr_mobile_v2.0_cls_infer.onnx')
                if os.path.exists(dp) and os.path.exists(rp) and os.path.exists(cp):
                    det_p, rec_p, cls_p = dp, rp, cp
                    break

        if det_p:
            engine = RapidOCR(det_model_path=det_p, rec_model_path=rec_p, cls_model_path=cls_p)
            return engine, True
        else:
            engine = RapidOCR()
            return engine, True
    except Exception:
        return None, False

ocr_engine, HAS_OCR = init_rapidocr()

VERSION = "V2.1"
APP_NAME = "FQT 实验室校准证书智能归档与管理工作台"

STANDARD_GROUPS = [
    '理化微生物组', '理化组', '微生物组',
    '液相组', '气相组', '元素组',
    '抽样组', '综合组', '运行组',
    '质保部', '报告组'
]

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

ISSUER_LABEL = {
    'DN':      '东莞帝恩',
    'DAF':     '深圳达丰',
    'NEM':     '广州中广测',
    'SMQ':     '深圳计量院',
    'CCIC':    '中检深圳',
    'UNKNOWN': '第三方机构',
}

ISSUER_SHORT_NAME = {
    'DN':      '帝恩',
    'DAF':     '达丰',
    'NEM':     '中广测',
    'SMQ':     '深圳计量院',
    'CCIC':    '中检',
    'UNKNOWN': '第三方',
}

DEFAULT_PRESETS = [
    {"name": "🏢 2026定量校准库", "path": r"D:\工作\01.实验室法定资质与17025体系\1.定量实验室17025体系维护\2.设备管理\1.设备检定校准\1.历年设备检定校准证书\2026年检定校准证书", "type": "quant"},
    {"name": "🧪 2026快检校准库", "path": r"D:\工作\03.快检质量网络与驻点管理\8.快检工作\7.快检设备检定校准\1.设备校准证书\2026年校准证书", "type": "quick"},
    {"name": "🏠 本地归档完成库", "path": os.path.join(APP_ROOT, "【归档完成】校准证书库"), "type": "archive"}
]

def load_presets():
    if os.path.exists(PRESETS_FILE):
        try:
            with open(PRESETS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, list) and data:
                    return data
        except Exception:
            pass
    return list(DEFAULT_PRESETS)

def save_presets(presets):
    try:
        with open(PRESETS_FILE, 'w', encoding='utf-8') as f:
            json.dump(presets, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ============================================================
# 第二部分：核心解析引擎与台账全景词库
# ============================================================

def sanitize_filename(text):
    if not text:
        return "未查找到"
    cleaned = re.sub(r'[\\/*?:"<>|\r\n\t]', '-', str(text)).strip()
    cleaned = re.sub(r'-+', '-', cleaned)
    return cleaned if cleaned else "未查找到"

def normalize_date(raw_date, filename=""):
    if raw_date:
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

    if filename:
        m_fn = re.search(r'[-_](20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])[-_.]', filename)
        if m_fn:
            return m_fn.group(1) + m_fn.group(2) + m_fn.group(3)
        m_fn2 = re.search(r'\b(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])\b', filename)
        if m_fn2:
            return m_fn2.group(0)

    return '未知校准'

def parse_yyyymm(date8):
    if date8 and re.match(r'^\d{8}$', date8):
        y = date8[:4]
        mm = date8[4:6]
        m = str(int(mm))
        return f"{y}年", f"{mm}月", f"{m}月"
    return '未知年份', '未知月份', '未知月份'

def is_text_garbled_or_image(text, min_chinese=4):
    if not text or len(text.strip()) < 20:
        return True
    chinese_chars = re.findall(r'[\u4e00-\u9fa5]', text)
    return len(chinese_chars) < min_chinese

def identify_issuer(text_clean):
    if re.search(r'帝恩|DNTesting|CNASL6483', text_clean, re.IGNORECASE):
        return 'DN'
    if re.search(r'达丰|DAF|DAFTJAX|DAFCJAX|DAFGJAX|CNASL6592', text_clean, re.IGNORECASE):
        return 'DAF'
    if re.search(r'中广测|广州分析测试中心|广东省测试分析研究所|NEM|TiC2600|CNASL7613|gznem\.com', text_clean, re.IGNORECASE):
        return 'NEM'
    if re.search(r'深圳市计量质量检测研究院|深圳计量院|smq\.com\.cn|CNASL0579|SZJLY', text_clean, re.IGNORECASE):
        return 'SMQ'
    if re.search(r'中检|CCIC|ccic-mts\.com|CNASL3103', text_clean, re.IGNORECASE):
        return 'CCIC'
    return 'UNKNOWN'


class LedgerDatabase:
    def __init__(self):
        self.quick_check_map = {}
        self.quantitative_map = {}
        self.serial_to_asset = {}
        self.all_known_assets = set()
        self.sorted_asset_list = []
        self.loaded_sources = []

    def load(self, base_dir=APP_ROOT, log_callback=None):
        def log(msg):
            if log_callback:
                log_callback(msg)

        log("📗 正在检索并加载快检与定量仪器设备台账...")
        
        search_dirs = [
            base_dir,
            APP_ROOT,
            os.path.join(APP_ROOT, ".."),
            r"D:\工作\01.实验室法定资质与17025体系\1.定量实验室17025体系维护\2.设备管理\1.设备检定校准\1.历年设备检定校准证书",
            r"D:\工作\01.实验室法定资质与17025体系\1.定量实验室17025体系维护\2.设备管理\1.设备检定校准",
            r"D:\工作\01.实验室法定资质与17025体系\1.定量实验室17025体系维护\2.设备管理",
            r"D:\工作\03.快检质量网络与驻点管理\8.快检工作\7.快检设备检定校准",
            r"D:\工作\03.快检质量网络与驻点管理\8.快检工作\7.快检设备检定校准\1.设备校准证书",
            r"D:\工作\9.行政工作\3.财务\1.历年付款及报销\2026年付款及报销\2026年校准费用",
            r"D:\工作\9.行政工作\3.财务\1.历年付款及报销\2026年付款及报销\2026年校准费用\2026年第三季度技术中心达丰校准设备",
            r"D:\工作\9.行政工作\3.财务\1.历年付款及报销\2026年付款及报销\2026年校准费用\2026年第二季度技术中心设备校准-达丰",
            r"D:\工作\9.行政工作\3.财务\1.历年付款及报销\2026年付款及报销\2026年校准费用\2026年第一季度技术中心设备校准（除移液枪和玻璃器皿）-达丰",
            r"D:\工作\9.行政工作\3.财务\1.历年付款及报销\2026年付款及报销\2026年校准费用\2026年第三季度技术中心中广测校准设备",
            r"D:\工作\9.行政工作\3.财务\1.历年付款及报销\2026年付款及报销\2026年校准费用\2026年第二季度技术中心移液枪及扩项设备校准-中广测",
            r"D:\工作\9.行政工作\3.财务\1.历年付款及报销\2026年付款及报销\2026年校准费用\2026年第一季度技术中心玻璃器皿及移液枪校准-中广测",
            os.getcwd()
        ]
        search_dirs = list(dict.fromkeys([os.path.abspath(d) for d in search_dirs if os.path.exists(d)]))
        
        # 1. 检索快检台账
        quick_candidates = ['2026年各实验室仪器设备校准清单.xlsx', '各实验室仪器设备校准清单.xlsx']
        found_quick = False
        for sdir in search_dirs:
            for fname in quick_candidates:
                fpath = os.path.join(sdir, fname)
                if os.path.exists(fpath):
                    self._load_quick_excel(fpath, log)
                    found_quick = True
                    self.loaded_sources.append(fpath)
                    break
            if found_quick:
                break
        if not found_quick:
            for sdir in search_dirs:
                for f in os.listdir(sdir):
                    if f.endswith('.xlsx') and '校准清单' in f and not f.startswith('~$') and not '汇总' in f:
                        fpath = os.path.join(sdir, f)
                        self._load_quick_excel(fpath, log)
                        found_quick = True
                        self.loaded_sources.append(fpath)
                        break
                if found_quick: break

        # 2. 检索定量计划汇总与台账
        quant_candidates = ['2026年年度计划汇总.xlsx', '年度计划汇总.xlsx', '量值溯源总表.xlsx']
        for sdir in search_dirs:
            for fname in quant_candidates:
                fpath = os.path.join(sdir, fname)
                if os.path.exists(fpath) and fpath not in self.loaded_sources:
                    self._load_quantitative_excel(fpath, log)
                    self.loaded_sources.append(fpath)

        # 3. 递归加载财务/季度批次送检台账中的关键设备
        for sdir in search_dirs:
            try:
                for f in os.listdir(sdir):
                    if f.endswith('.xlsx') and not f.startswith('~$') and not '本次扫描' in f and not '统计汇总' in f and not '提取汇总' in f:
                        fpath = os.path.join(sdir, f)
                        if fpath not in self.loaded_sources:
                            self._load_quantitative_excel(fpath, log)
                            self.loaded_sources.append(fpath)
            except Exception:
                pass

        invalid_set = {'NAN', 'NONE', '', '/', '\\', '-', '—', '无', '无编号', '待定', '暂不校准', 'SERIAL'}
        all_assets = set()
        for k in self.quick_check_map:
            if k.upper() not in invalid_set:
                all_assets.add(k)
        for k in self.quantitative_map:
            if k.upper() not in invalid_set:
                all_assets.add(k)
        self.all_known_assets = all_assets
        self.sorted_asset_list = sorted(list(all_assets), key=len, reverse=True)
        log(f"✅ 台账构建就绪：快检 {len(self.quick_check_map)} 台，定量 {len(self.quantitative_map)} 台 (全局有效资产编号 {len(self.sorted_asset_list)} 条)")

    def _load_quick_excel(self, fpath, log):
        try:
            xl = pd.ExcelFile(fpath)
            for sheet in xl.sheet_names:
                df = pd.read_excel(fpath, sheet_name=sheet, header=None)
                header_idx = -1
                for idx, row in df.head(10).iterrows():
                    row_txt = ' '.join([str(v) for v in row.values if pd.notna(v)])
                    if '设备编号' in row_txt and ('实验室' in row_txt or '设备名称' in row_txt or '仪器名称' in row_txt):
                        header_idx = idx
                        break
                if header_idx != -1:
                    df_sheet = pd.read_excel(fpath, sheet_name=sheet, header=header_idx)
                    df_sheet.columns = [str(c).strip().replace('\n','').replace('\r','') for c in df_sheet.columns]
                    no_col = next((c for c in df_sheet.columns if '设备编号' in c), None)
                    lab_col = next((c for c in df_sheet.columns if '实验室' in c or '项目' in c), None)
                    inst_col = next((c for c in df_sheet.columns if '设备名称' in c or '仪器名称' in c), None)
                    model_col = next((c for c in df_sheet.columns if '型号' in c or '规格' in c), None)

                    if no_col:
                        for _, r in df_sheet.iterrows():
                            code = str(r[no_col]).strip() if pd.notna(r[no_col]) else ''
                            lab = str(r[lab_col]).strip() if lab_col and pd.notna(r[lab_col]) else ''
                            inst = str(r[inst_col]).strip() if inst_col and pd.notna(r[inst_col]) else ''
                            model = str(r[model_col]).strip() if model_col and pd.notna(r[model_col]) else ''
                            if code and code not in {'nan', 'None', '', '/', '—', '无'}:
                                self.quick_check_map[code] = {
                                    'code': code,
                                    'lab': lab or '未知实验室',
                                    'inst': inst or '未查找到',
                                    'model': model if model not in {'nan', 'None'} else '',
                                    'type': '快检'
                                }
        except Exception as e:
            log(f"⚠️ 加载快检台账异常: {e}")

    def _load_quantitative_excel(self, fpath, log):
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
                    model_col = next((c for c in df_sheet.columns if re.search(r'规格型号|型号规格|型号|规格', c)), None)
                    sn_col = next((c for c in df_sheet.columns if re.search(r'出厂编号|序列号|SN', c)), None)
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
                            model = str(r[model_col]).strip() if model_col and pd.notna(r[model_col]) else ''
                            sn = str(r[sn_col]).strip() if sn_col and pd.notna(r[sn_col]) else ''

                            invalid = {'nan', 'None', '', '/', '—', '无'}
                            if code and code not in invalid:
                                grp_val = grp if grp not in invalid else '未知组别'
                                inst_val = inst if inst not in invalid else '未查找到'
                                
                                if code not in self.quantitative_map or (self.quantitative_map[code]['group'] == '未知组别' and grp_val != '未知组别'):
                                    self.quantitative_map[code] = {
                                        'code': code,
                                        'group': grp_val,
                                        'inst': inst_val,
                                        'model': model if model not in invalid else '',
                                        'serial_no': sn if sn not in invalid else '',
                                        'type': '定量'
                                    }
                                if sn and sn not in invalid and len(sn) >= 3:
                                    self.serial_to_asset[sn] = code
        except Exception as e:
            log(f"⚠️ 加载定量台账异常: {e}")

    def _lookup_direct(self, clean_no):
        if clean_no in self.quick_check_map:
            info = self.quick_check_map[clean_no]
            return {
                'asset_no': clean_no, 'type': '快检',
                'lab': info.get('lab', '未知实验室'),
                'group': info.get('lab', '未知实验室'),
                'inst': info.get('inst', '未查找到'), 'serial_no': ''
            }
        if clean_no in self.quantitative_map:
            info = self.quantitative_map[clean_no]
            return {
                'asset_no': clean_no, 'type': '定量',
                'group': info.get('group', '未知组别'),
                'lab': '定量实验室',
                'inst': info.get('inst', '未查找到'),
                'serial_no': info.get('serial_no', '')
            }
        no_upper = clean_no.upper()
        for k, v in self.quick_check_map.items():
            if k.upper() == no_upper:
                return {'asset_no': k, 'type': '快检', 'lab': v.get('lab', '未知实验室'), 'group': v.get('lab', '未知实验室'), 'inst': v.get('inst', '未查找到'), 'serial_no': ''}
        for k, v in self.quantitative_map.items():
            if k.upper() == no_upper:
                return {'asset_no': k, 'type': '定量', 'group': v.get('group', '未知组别'), 'lab': '定量实验室', 'inst': v.get('inst', '未查找到'), 'serial_no': v.get('serial_no', '')}
        return None

    def query(self, asset_no=None, serial_no=None, inst_name=""):
        if asset_no:
            clean_no = str(asset_no).strip()
            clean_no = re.sub(r'^HO(\d+)$', r'H0\1', clean_no)
            
            res = self._lookup_direct(clean_no)
            if res:
                if serial_no: res['serial_no'] = serial_no
                return res

            # 零填充兼容模糊检索 (如 FYS047 <-> FYS47, WJ051 <-> WJ51, WSJ035 <-> WSJ35)
            m_pad = re.match(r'^([A-Za-z]+)0+(\d+)$', clean_no)
            if m_pad:
                alt_no = f"{m_pad.group(1)}{m_pad.group(2)}"
                res = self._lookup_direct(alt_no)
                if res:
                    res['asset_no'] = clean_no
                    if serial_no: res['serial_no'] = serial_no
                    return res
            else:
                m_unpad = re.match(r'^([A-Za-z]+)(\d{1,2})$', clean_no)
                if m_unpad:
                    alt_no = f"{m_unpad.group(1)}{int(m_unpad.group(2)):03d}"
                    res = self._lookup_direct(alt_no)
                    if res:
                        res['asset_no'] = clean_no
                        if serial_no: res['serial_no'] = serial_no
                        return res

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

        # 启发式仪器与组别智能映射
        grp_fallback = '未知组别'
        dev_type = '未知'
        if inst_name:
            if any(k in inst_name for k in ['液相', 'HPLC', '荧光检测器', '二极管阵列', 'U3000']):
                grp_fallback = '液相组'
                dev_type = '定量'
            elif any(k in inst_name for k in ['气相', 'GC', '顶空', 'ECD', 'FID', 'FPD']):
                grp_fallback = '气相组'
                dev_type = '定量'
            elif any(k in inst_name for k in ['离子计', '折光仪', '折射仪', '酸度计', 'pH计', '旋光仪', '电导率', '水分仪', '天平', '分析天平']):
                grp_fallback = '理化组'
                dev_type = '定量'
            elif any(k in inst_name for k in ['培养箱', '生化培养箱', '灭菌器', '蒸汽灭菌', '生物安全柜', '超净工作台', '菌落计数']):
                grp_fallback = '微生物组'
                dev_type = '定量'
            elif any(k in inst_name for k in ['ICP', '原子吸收', 'AAS', '原子荧光', 'AFS', '测汞仪', '重金属']):
                grp_fallback = '元素组'
                dev_type = '定量'
            elif any(k in inst_name for k in ['标准筛', '分样筛', '药典筛', '试验筛', '振筛机', '制样']):
                grp_fallback = '运行组'
                dev_type = '定量'
            elif any(k in inst_name for k in ['温湿度', '温湿度计', '温度计', '温湿度记录仪']):
                grp_fallback = '综合组'
                dev_type = '定量'

        return {
            'asset_no': asset_no or '未知编号', 'type': dev_type,
            'group': grp_fallback, 'lab': '未知实验室', 'inst': inst_name or '未查找到', 'serial_no': serial_no or ''
        }


# ============================================================
# 各大计量校准机构高精字段解析器
# ============================================================

def extract_dn(text, text_clean, ledger, filename="", doc=None):
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    asset_no = None
    serial_no = ""
    blocks = doc[0].get_text('blocks') if (doc and len(doc) > 0) else []

    # 1. 优先从台账检索已知的设备资产编号
    for asset in ledger.sorted_asset_list:
        if re.search(r'\b' + re.escape(asset) + r'\b', text, re.I) or (filename and asset in filename):
            asset_no = asset
            break

    # 2. 空间结构抽取管理号与出厂编号
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

    # 3. 文本行后备检索
    if not asset_no:
        for i, l in enumerate(lines):
            if re.search(r'管\s*理\s*号|Asset\s*No|器\s*具\s*编\s*号', l, re.I):
                for j in range(i+1, min(i+4, len(lines))):
                    cand = lines[j].strip()
                    if re.search(r'[A-Za-z0-9]', cand) and not any(k in cand for k in ['Date', 'Model', 'Manufacturer', 'Description', 'Serial', 'Asset', 'JJG', 'JJF']):
                        asset_no = cand
                        break
                if asset_no: break

    if not serial_no:
        for i, l in enumerate(lines):
            if re.search(r'出\s*厂\s*编\s*号|Serial\s*No\.?', l, re.I):
                for j in range(i+1, min(i+4, len(lines))):
                    cand = lines[j].strip()
                    if re.search(r'[A-Za-z0-9]', cand) and not any(k in cand for k in ['Date', 'Model', 'Manufacturer', 'Description', 'Serial', 'Asset', 'JJG', 'JJF']):
                        serial_no = cand
                        break
                if serial_no: break

    # 4. 文件名兜底
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

    # 5. 校准日期抽取 (空间坐标精确定位 + 正则双保险)
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

    # 6. 签发日期抽取
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

    # 7. 仪器名称抽取 (避开 'Description' 标签)
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

    # 8. 组别与驻点智能解析 (台账 + 地址全景匹配)
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

def extract_daf(text, text_clean, ledger, filename="", doc=None):
    asset_no = None
    serial_no = ""
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    blocks = doc[0].get_text('blocks') if (doc and len(doc) > 0) else []

    for i, l in enumerate(lines):
        if re.search(r'出\s*厂\s*编\s*号|Serial\s*No\.?', l, re.I) and not serial_no:
            for j in range(i+1, min(i+4, len(lines))):
                cand = lines[j].strip()
                if re.search(r'[A-Za-z0-9]', cand) and not any(k in cand for k in ['Date', 'Model', 'Manufacturer', 'Description', 'Serial', 'Asset', 'JJG', 'JJF']):
                    serial_no = cand
                    break
        elif re.search(r'管\s*理\s*号|Asset\s*No\.?', l, re.I) and not asset_no:
            for j in range(i+1, min(i+4, len(lines))):
                cand = lines[j].strip()
                if re.search(r'[A-Za-z0-9]', cand) and not re.match(r'^(JJF|JJG|GB|CNAS|DAF)', cand, re.I) and not any(k in cand for k in ['Date', 'Model', 'Manufacturer', 'Description', 'Serial', 'Asset']):
                    asset_no = cand
                    break

    if blocks:
        if not asset_no or asset_no == '未知编号':
            for b in blocks:
                if 380 <= b[1] <= 475 and 320 <= b[0] <= 520:
                    t = b[4].strip().replace('\n', '')
                    if t and t != '/' and not re.match(r'^(Asset|Serial|Model|出厂|管理)', t, re.I):
                        asset_no = t
                        break
        if not serial_no:
            for b in blocks:
                if 380 <= b[1] <= 475 and 130 <= b[0] <= 320:
                    t = b[4].strip().replace('\n', '')
                    if t and t != '/' and not re.match(r'^(Asset|Serial|Model|出厂|管理)', t, re.I):
                        serial_no = t
                        break

    if serial_no:
        m_paren = re.search(r'^(.*?)[(（]([A-Za-z0-9\-]+)[)）]$', serial_no)
        if m_paren:
            serial_no = m_paren.group(1).strip()
            if not asset_no:
                asset_no = m_paren.group(2).strip()

    if not asset_no and filename:
        m_fn = re.search(r'[(（]([A-Za-z0-9\-]+)[)）]', filename)
        if m_fn:
            asset_no = m_fn.group(1).strip()

    if not asset_no:
        for asset in ledger.sorted_asset_list:
            if re.search(r'\b' + re.escape(asset) + r'\b', text, re.I) or (filename and asset in filename):
                asset_no = asset
                break

    asset_no = re.sub(r'^HO(\d+)$', r'H0\1', asset_no or '')

    inst_name = ''
    if blocks:
        for b in blocks:
            if 250 <= b[1] <= 360 and 130 <= b[0] <= 350:
                t = b[4].strip().replace('\n', '')
                if t and t != 'Description' and not re.match(r'^(Model|Serial|Asset|Manufacturer|Type|型号|出厂|管理)', t, re.I):
                    inst_name = t
                    break
    if not inst_name or inst_name == 'Description':
        for i, l in enumerate(lines):
            if re.search(r'仪\s*器\s*名\s*称|样\s*品\s*名\s*称|Description', l, re.I) and i+1 < len(lines):
                cand = lines[i+1].strip()
                if cand and not any(b in cand for b in ['Model', 'Type', 'Serial', 'Asset', 'Manufacturer', '型号', '出厂']):
                    inst_name = cand
                    break

    cal_raw = ""
    for i, l in enumerate(lines):
        if re.search(r'(?:Date of Calibr|校\s*准\s*日\s*期|检\s*定\s*日\s*期)', l, re.I):
            for j in range(i+1, min(i+4, len(lines))):
                m_d = re.search(r'(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)', lines[j])
                if m_d:
                    cal_raw = m_d.group(1)
                    break
            if cal_raw: break

    if not cal_raw and blocks:
        for b in blocks:
            if 420 <= b[1] <= 560:
                m = re.search(r'(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)', b[4])
                if m:
                    cal_raw = m.group(1)
                    break

    if not cal_raw:
        m_cal = re.search(r'(?:校\s*准\s*日\s*期|检\s*定\s*日\s*期|Date\s*of\s*Calibration)[^\n\r\d]*(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)', text, re.I)
        if m_cal and m_cal.group(1):
            cal_raw = m_cal.group(1)

    cal_date = normalize_date(cal_raw, filename=filename)

    info = ledger.query(asset_no=asset_no, serial_no=serial_no, inst_name=inst_name)
    if not asset_no or asset_no == '未知编号':
        if info.get('asset_no') and info.get('asset_no') != '未知编号':
            asset_no = info.get('asset_no')

    if not inst_name or inst_name == '未查找到':
        inst_name = info.get('inst') if info.get('inst') not in {'未查找到', 'nan', 'None'} else '未查找到'

    return {
        'asset_no': asset_no or serial_no or '未知编号',
        'serial_no': serial_no or info.get('serial_no', ''),
        'inst_name': inst_name,
        'cal_date': cal_date,
        'group': info.get('group', '未知组别'),
        'lab_name': info.get('lab', '定量实验室'),
        'device_type': info.get('type', '定量')
    }

def extract_nem(text, text_clean, ledger, filename="", doc=None):
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    desc, model, raw_sn, cal_date, issue_date = '', '', '', '', ''
    for i, l in enumerate(lines):
        if l == 'Description' and i+1 < len(lines):
            desc = lines[i+1]
        elif l == 'Model/Type' and i+1 < len(lines):
            model = lines[i+1]
        elif (re.search(r'^(Serial|出\s*厂\s*编\s*号|器\s*具\s*编\s*号|管\s*理\s*号)', l, re.I)) and not raw_sn:
            for j in range(i+1, min(i+4, len(lines))):
                if re.search(r'[A-Za-z0-9]', lines[j]) and not any(k in lines[j] for k in ['Date', 'Model', 'Manufacturer', 'Description', 'Serial']):
                    raw_sn = lines[j]
                    break
        elif (re.search(r'(?:Date of Calibr|校\s*准\s*日\s*期|检\s*定\s*日\s*期)', l, re.I)) and not cal_date:
            for j in range(i+1, min(i+4, len(lines))):
                m_d = re.search(r'(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)', lines[j])
                if m_d:
                    cal_date = m_d.group(1)
                    break
        elif (re.search(r'(?:Date of Issue|签\s*发\s*日\s*期|批\s*准\s*日\s*期|发布日期|Issued Date)', l, re.I)) and not issue_date:
            for j in range(i+1, min(i+4, len(lines))):
                m_d = re.search(r'(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)', lines[j])
                if m_d:
                    issue_date = m_d.group(1)
                    break

    asset_no = None
    serial_no = raw_sn
    
    m_paren = re.search(r'^(.*?)[(（]([A-Za-z0-9\-]+)[)）]$', raw_sn)
    if m_paren:
        serial_no = m_paren.group(1).strip()
        asset_cand = m_paren.group(2).strip()
        asset_cand = re.sub(r'^HO(\d+)$', r'H0\1', asset_cand)
        asset_no = asset_cand

    if not asset_no and filename:
        m_fn = re.search(r'[(（]([A-Za-z0-9\-]+)[)）]', filename)
        if m_fn:
            asset_cand = m_fn.group(1).strip()
            asset_cand = re.sub(r'^HO(\d+)$', r'H0\1', asset_cand)
            if asset_cand in ledger.all_known_assets or re.match(r'^(H\d{4}|YYQ\d+|WJ\d+|MDP\d+|KJ\d+)', asset_cand, re.I):
                asset_no = asset_cand

    if not asset_no:
        for a in ledger.sorted_asset_list:
            if a in text or (filename and a in filename):
                asset_no = a
                break

    asset_no = re.sub(r'^HO(\d+)$', r'H0\1', asset_no or '')

    info = ledger.query(asset_no=asset_no, serial_no=serial_no, inst_name=desc)
    if not asset_no and info.get('asset_no') != '未知编号':
        asset_no = info.get('asset_no')

    cal_date_norm = normalize_date(cal_date, filename=filename)
    inst_name = info.get('inst') if info.get('inst') not in {'未查找到', 'nan', 'None'} else desc
    inst_name = re.sub(r'[\uFFFD\?]', '', str(inst_name)).strip() or '未查找到'

    return {
        'asset_no': asset_no or serial_no or '未知编号',
        'serial_no': serial_no or '',
        'inst_name': inst_name,
        'cal_date': cal_date_norm,
        'group': info.get('group', '未知组别'),
        'lab_name': info.get('lab', '定量实验室'),
        'device_type': info.get('type', '定量')
    }

def extract_smq(text, text_clean, ledger, filename="", doc=None):
    asset_no = None
    serial_no = ""
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    for asset in ledger.sorted_asset_list:
        if asset in text or (filename and asset in filename):
            asset_no = asset
            break
    if not asset_no and filename:
        m_fn = re.search(r'[(（]([A-Za-z0-9\-]+)[)）]', filename)
        if m_fn:
            asset_no = m_fn.group(1).strip()

    for i, l in enumerate(lines):
        if re.search(r'出\s*厂\s*编\s*号|Serial\s*No\.?', l, re.I) and not serial_no:
            for j in range(i+1, min(i+4, len(lines))):
                if re.search(r'[A-Za-z0-9]', lines[j]) and not any(k in lines[j] for k in ['Date', 'Model', 'Manufacturer', 'Description', 'Serial', 'Asset']):
                    serial_no = lines[j]
                    break

    asset_no = re.sub(r'^HO(\d+)$', r'H0\1', asset_no or '')

    inst_name = ''
    for i, l in enumerate(lines):
        if re.search(r'仪\s*器\s*名\s*称|样\s*品\s*名\s*称|Description', l, re.I) and i+1 < len(lines):
            cand = lines[i+1].strip()
            if cand and not any(b in cand for b in ['Model', 'Type', 'Serial', 'Asset', 'Manufacturer', '型号', '出厂']):
                inst_name = cand
                break

    info = ledger.query(asset_no=asset_no, serial_no=serial_no, inst_name=inst_name)
    if not inst_name or inst_name == '未查找到':
        inst_name = info.get('inst') if info.get('inst') not in {'未查找到', 'nan', 'None'} else '未查找到'

    cal_raw = ""
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

    return {
        'asset_no': asset_no or serial_no or '未知编号',
        'serial_no': serial_no,
        'inst_name': inst_name,
        'cal_date': cal_date,
        'group': info.get('group', '未知组别'),
        'lab_name': info.get('lab', '定量实验室'),
        'device_type': info.get('type', '定量')
    }

def extract_ccic(text, text_clean, ledger, filename="", doc=None):
    asset_no = None
    serial_no = ""
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    for asset in ledger.sorted_asset_list:
        if asset in text or (filename and asset in filename):
            asset_no = asset
            break
    if not asset_no and filename:
        m_fn = re.search(r'[(（]([A-Za-z0-9\-]+)[)）]', filename)
        if m_fn:
            asset_no = m_fn.group(1).strip()

    asset_no = re.sub(r'^HO(\d+)$', r'H0\1', asset_no or '')

    inst_name = ''
    for i, l in enumerate(lines):
        if re.search(r'仪\s*器\s*名\s*称|样\s*品\s*名\s*称|Description', l, re.I) and i+1 < len(lines):
            cand = lines[i+1].strip()
            if cand and not any(b in cand for b in ['Model', 'Type', 'Serial', 'Asset', 'Manufacturer', '型号', '出厂']):
                inst_name = cand
                break

    info = ledger.query(asset_no=asset_no, serial_no=serial_no, inst_name=inst_name)
    if not inst_name or inst_name == '未查找到':
        inst_name = info.get('inst') if info.get('inst') not in {'未查找到', 'nan', 'None'} else '未查找到'

    cal_raw = ""
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

    return {
        'asset_no': asset_no or serial_no or '未知编号',
        'serial_no': serial_no,
        'inst_name': inst_name,
        'cal_date': cal_date,
        'group': info.get('group', '未知组别'),
        'lab_name': info.get('lab', '定量实验室'),
        'device_type': info.get('type', '定量')
    }

def extract_generic(text, text_clean, ledger, filename="", doc=None):
    asset_no = None
    serial_no = ""
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    
    if filename:
        m_fn = re.search(r'[(（]([A-Za-z0-9\-]+)[)）]', filename)
        if m_fn:
            asset_cand = m_fn.group(1).strip()
            asset_cand = re.sub(r'^HO(\d+)$', r'H0\1', asset_cand)
            if asset_cand in ledger.all_known_assets or re.match(r'^(H\d{4}|YYQ\d+|WJ\d+|MDP\d+|KJ\d+)', asset_cand, re.I):
                asset_no = asset_cand

    if not asset_no:
        for asset in ledger.sorted_asset_list:
            if asset in text or (filename and asset in filename):
                asset_no = asset
                break

    asset_no = re.sub(r'^HO(\d+)$', r'H0\1', asset_no or '')

    inst_name = ''
    for i, l in enumerate(lines):
        if re.search(r'仪\s*器\s*名\s*称|样\s*品\s*名\s*称|Description', l, re.I) and i+1 < len(lines):
            cand = lines[i+1].strip()
            if cand and not any(b in cand for b in ['Model', 'Type', 'Serial', 'Asset', 'Manufacturer', '型号', '出厂']):
                inst_name = cand
                break

    info = ledger.query(asset_no=asset_no, serial_no=serial_no, inst_name=inst_name)
    if not inst_name or inst_name == '未查找到':
        inst_name = info.get('inst') if info.get('inst') not in {'未查找到', 'nan', 'None'} else '未查找到'

    cal_raw = ""
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

    return {
        'asset_no': asset_no or '未知编号',
        'serial_no': serial_no,
        'inst_name': inst_name,
        'cal_date': cal_date,
        'group': info.get('group', '未知组别'),
        'lab_name': info.get('lab', '未知实验室'),
        'device_type': info.get('type', '未知')
    }


def build_new_filename(issuer, fields):
    device_type = fields.get('device_type', '未知')
    asset_no = fields.get('asset_no', '未知编号')
    serial_no = fields.get('serial_no', '')
    inst_name = fields.get('inst_name', '未查找到')
    cal_date = fields.get('cal_date', '未知校准')
    group = fields.get('group', '未知组别')

    # 快检设备或帝恩报告
    if issuer == 'DN' or device_type == '快检':
        lab_name = fields.get('lab_name', '未知实验室')
        issue_date = fields.get('issue_date', '')
        if issue_date and issue_date != '未知校准' and issue_date != cal_date:
            parts = [lab_name, asset_no, cal_date, issue_date, inst_name]
        else:
            parts = [lab_name, asset_no, cal_date, inst_name]
        return '_'.join([sanitize_filename(p) for p in parts])

    # 移液枪复合前缀
    if str(asset_no).startswith('YYQ') and serial_no and serial_no != '未知出厂号' and serial_no != asset_no:
        code_tag = f"{serial_no}({asset_no})"
        parts = [code_tag, inst_name, cal_date, group]
        return '_'.join([sanitize_filename(p) for p in parts])

    # 定量设备规范：{设备编号}_{仪器名称}_{校准日期}_{所属组别}
    code = asset_no if asset_no != '未知编号' else (serial_no or '未知编号')
    parts = [code, inst_name, cal_date, group]
    return '_'.join([sanitize_filename(p) for p in parts])


def build_archive_dir(archive_root, issuer, fields):
    device_type = fields.get('device_type', '未知')
    cal_date = fields.get('cal_date', '')
    year_tag, mm_tag, m_tag = parse_yyyymm(cal_date)
    issuer_name = ISSUER_SHORT_NAME.get(issuer, '第三方')

    if device_type == '快检':
        lab_name = sanitize_filename(fields.get('lab_name', '未知实验室'))
        if not lab_name.endswith('项目组') and not lab_name.endswith('街道') and not lab_name.endswith('组'):
            folder_tag = f"{lab_name}项目组"
        else:
            folder_tag = lab_name
        month_folder = f"{year_tag}{m_tag}{folder_tag}设备校准证书"
        return os.path.join(archive_root, '快检设备', month_folder)

    elif device_type == '定量':
        group = sanitize_filename(fields.get('group', '未知组别'))
        year_folder = f"{year_tag}检定校准证书"
        month_issuer_folder = f"{year_tag}{m_tag}{issuer_name}校准证书"
        detail_folder = f"{year_tag}{m_tag}{group}校准证书-{issuer_name}"
        return os.path.join(archive_root, year_folder, month_issuer_folder, detail_folder)

    else:
        return os.path.join(archive_root, '未分类设备', f"{year_tag}{m_tag}校准证书")


def safe_extract_zip(zip_path, target_extract_dir):
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
    except Exception:
        pass
    return extracted_pdf_paths


def export_styled_excel(df_records, df_missing, output_excel_path):
    import openpyxl
    wb = openpyxl.Workbook()
    
    # Sheet 1: 扫描识别与归档明细
    ws1 = wb.active
    ws1.title = "校准证书识别明细"
    
    # Write headers
    cols1 = list(df_records.columns)
    ws1.append(cols1)
    for row in df_records.itertuples(index=False):
        ws1.append(list(row))
        
    # Sheet 2: 计划对比与待收回清单
    if df_missing is not None and not df_missing.empty:
        ws2 = wb.create_sheet(title="计划待收回与缺漏清单")
        cols2 = list(df_missing.columns)
        ws2.append(cols2)
        for row in df_missing.itertuples(index=False):
            ws2.append(list(row))
    else:
        ws2 = None

    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    header_font = Font(name="Microsoft YaHei UI", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Microsoft YaHei UI", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"), bottom=Side(style="thin", color="CBD5E1")
    )
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    for ws in [ws1, ws2]:
        if ws is None: continue
        ws.row_dimensions[1].height = 30
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border

        for row_idx in range(2, ws.max_row + 1):
            ws.row_dimensions[row_idx].height = 24
            is_zebra = (row_idx % 2 == 1)
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = data_font
                cell.border = thin_border
                if is_zebra:
                    cell.fill = zebra_fill
                val_str = str(cell.value or '')
                if any(k in str(ws.cell(row=1, column=col_idx).value or '') for k in ['序号', '机构', '日期', '类型', '方式', '状态']):
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


# ============================================================
# 第三部分：多线程后台处理 Worker (QThread)
# ============================================================

class ArchiveWorker(QThread):
    progress_signal = pyqtSignal(int, int, str)
    log_signal = pyqtSignal(str, str)
    record_signal = pyqtSignal(dict)
    finished_signal = pyqtSignal(dict)

    def __init__(self, target_dir, archive_root, mode='dry_run', enable_ocr=True, auto_unzip=True):
        super().__init__()
        self.target_dir = target_dir
        self.archive_root = archive_root
        self.mode = mode
        self.enable_ocr = enable_ocr
        self.auto_unzip = auto_unzip
        self._is_stopped = False

    def stop(self):
        self._is_stopped = True

    def run(self):
        start_time = datetime.datetime.now()
        self.log_signal.emit("INFO", f"🚀 启动工作台识别引擎 (模式: {self.mode.upper()})...")

        ledger = LedgerDatabase()
        ledger.load(base_dir=self.target_dir, log_callback=lambda m: self.log_signal.emit("INFO", m))

        if self._is_stopped:
            return

        self.log_signal.emit("INFO", f"🔍 正在执行深度文件雷达扫描: {self.target_dir}")
        pdf_list = []
        zip_files = []
        ignored_count = 0
        abs_archive_root = os.path.abspath(self.archive_root)

        for root, dirs, files in os.walk(self.target_dir):
            if self._is_stopped:
                return
            abs_root = os.path.abspath(root)
            if abs_root == abs_archive_root:
                pass
            elif abs_root.startswith(abs_archive_root) and '未分类设备' not in abs_root and self.target_dir != self.archive_root:
                ignored_count += sum(1 for f in files if f.lower().endswith('.pdf'))
                continue
            if '.git' in root or '__pycache__' in root or '_temp_zip_' in root:
                continue

            for filename in files:
                lower = filename.lower()
                fp = os.path.join(root, filename)
                if lower.endswith('.pdf'):
                    pdf_list.append(fp)
                elif lower.endswith('.zip'):
                    zip_files.append(fp)

        if self.auto_unzip and zip_files:
            self.log_signal.emit("INFO", f"📦 扫描到 {len(zip_files)} 个 ZIP 压缩包，正在自动解包...")
            temp_zip_dir = os.path.join(self.target_dir, "_temp_zip_extracted_")
            for zp in zip_files:
                if self._is_stopped:
                    return
                z_base = os.path.splitext(os.path.basename(zp))[0]
                ex_folder = os.path.join(temp_zip_dir, z_base)
                os.makedirs(ex_folder, exist_ok=True)
                extracted = safe_extract_zip(zp, ex_folder)
                self.log_signal.emit("INFO", f"   已解包 [{os.path.basename(zp)}] -> 提取 {len(extracted)} 份 PDF")
                pdf_list.extend(extracted)

        unique_pdfs = list(dict.fromkeys(pdf_list))
        total_files = len(unique_pdfs)

        if total_files == 0:
            self.log_signal.emit("WARN", "⚠️ 未在指定目录扫描到任何待处理的 PDF 校准证书！")
            self.finished_signal.emit({'records': [], 'total': 0, 'time_cost': 0, 'missing': []})
            return

        self.log_signal.emit("SUCCESS", f"📡 雷达扫描完毕：共捕获 {total_files} 份待处理证书 (跳过已归档 {ignored_count} 份)")

        if self.mode != 'dry_run':
            os.makedirs(self.archive_root, exist_ok=True)

        records = []
        success_count = 0
        error_count = 0
        scanned_asset_codes = set()

        for idx, pdf_path in enumerate(unique_pdfs, 1):
            if self._is_stopped:
                self.log_signal.emit("WARN", "⏹️ 任务已被用户中止。")
                break

            filename = os.path.basename(pdf_path)
            self.progress_signal.emit(idx, total_files, filename)

            text = ""
            used_ocr = False
            doc = None
            try:
                if fitz is not None:
                    doc = fitz.open(pdf_path)
                    for p_i in range(min(3, len(doc))):
                        page = doc[p_i]
                        p_text = page.get_text("text") or ''
                        if is_text_garbled_or_image(p_text, min_chinese=4) and self.enable_ocr and HAS_OCR and ocr_engine:
                            pix = page.get_pixmap(dpi=200)
                            ocr_res, _ = ocr_engine(pix.tobytes("png"))
                            if ocr_res:
                                p_text = '\n'.join([line[1] for line in ocr_res])
                                used_ocr = True
                        text += p_text + '\n'
            except Exception as e:
                self.log_signal.emit("ERROR", f"❌ 读取失败 [{filename}]: {e}")
                error_count += 1
                if doc is not None:
                    try: doc.close()
                    except Exception: pass
                continue

            text_clean = text.replace(' ', '').replace('\n', '')

            if (len(text_clean.strip()) < 30 or identify_issuer(text_clean) == 'UNKNOWN' or is_text_garbled_or_image(text, min_chinese=6)) and self.enable_ocr and HAS_OCR and ocr_engine and not used_ocr:
                try:
                    ocr_total = ""
                    if doc is not None:
                        for p_i in range(min(2, len(doc))):
                            pix = doc[p_i].get_pixmap(dpi=200)
                            ocr_res, _ = ocr_engine(pix.tobytes("png"))
                            if ocr_res:
                                ocr_total += '\n' + '\n'.join([line[1] for line in ocr_res])
                    if ocr_total:
                        text = ocr_total
                        text_clean = text.replace(' ', '').replace('\n', '')
                        used_ocr = True
                except Exception:
                    pass

            issuer = identify_issuer(text_clean)
            if issuer == 'DN': fields = extract_dn(text, text_clean, ledger, filename=filename, doc=doc)
            elif issuer == 'DAF': fields = extract_daf(text, text_clean, ledger, filename=filename, doc=doc)
            elif issuer == 'NEM': fields = extract_nem(text, text_clean, ledger, filename=filename, doc=doc)
            elif issuer == 'SMQ': fields = extract_smq(text, text_clean, ledger, filename=filename, doc=doc)
            elif issuer == 'CCIC': fields = extract_ccic(text, text_clean, ledger, filename=filename, doc=doc)
            else: fields = extract_generic(text, text_clean, ledger, filename=filename, doc=doc)

            if (not fields.get('group') or fields.get('group') in {'未知组别', '未知'}) and filename:
                for g in STANDARD_GROUPS:
                    if g in filename:
                        fields['group'] = g
                        if fields.get('device_type') == '未知':
                            fields['device_type'] = '定量'
                        break

            if (not fields.get('lab_name') or fields.get('lab_name') in {'未知实验室', '未知'}) and filename:
                parts = os.path.splitext(os.path.basename(filename))[0].split('_')
                if len(parts) >= 2 and not re.match(r'^(20\d{2}|H\d+|YYQ|KJ|WJ|\d+$)', parts[0]):
                    fields['lab_name'] = parts[0]

            if (fields.get('asset_no') in {'未知编号', '未查找到'} or fields.get('inst_name') == '未查找到') and self.enable_ocr and HAS_OCR and ocr_engine and not used_ocr:
                try:
                    ocr_total = ""
                    if doc is not None:
                        for p_i in range(min(2, len(doc))):
                            pix = doc[p_i].get_pixmap(dpi=200)
                            ocr_res, _ = ocr_engine(pix.tobytes("png"))
                            if ocr_res:
                                ocr_total += '\n' + '\n'.join([line[1] for line in ocr_res])
                    if ocr_total:
                        text = ocr_total
                        text_clean = text.replace(' ', '').replace('\n', '')
                        issuer = identify_issuer(text_clean)
                        if issuer == 'DN': fields = extract_dn(text, text_clean, ledger, filename=filename, doc=doc)
                        elif issuer == 'DAF': fields = extract_daf(text, text_clean, ledger, filename=filename, doc=doc)
                        elif issuer == 'NEM': fields = extract_nem(text, text_clean, ledger, filename=filename, doc=doc)
                        elif issuer == 'SMQ': fields = extract_smq(text, text_clean, ledger, filename=filename, doc=doc)
                        elif issuer == 'CCIC': fields = extract_ccic(text, text_clean, ledger, filename=filename, doc=doc)
                        else: fields = extract_generic(text, text_clean, ledger, filename=filename, doc=doc)
                        used_ocr = True
                except Exception:
                    pass

            if doc is not None:
                try: doc.close()
                except Exception: pass
                doc = None

            label = ISSUER_LABEL.get(issuer, '未知机构')
            device_type = fields.get('device_type', '未知')

            new_filename = f"{build_new_filename(issuer, fields)}.pdf"
            target_dir_path = build_archive_dir(self.archive_root, issuer, fields)

            if self.mode != 'dry_run':
                os.makedirs(target_dir_path, exist_ok=True)

            target_file_path = os.path.join(target_dir_path, new_filename)

            conflict_suffix = 1
            base_stem, ext = os.path.splitext(new_filename)
            while os.path.exists(target_file_path) and os.path.abspath(target_file_path) != os.path.abspath(pdf_path):
                new_filename = f"{base_stem}_{conflict_suffix}{ext}"
                target_file_path = os.path.join(target_dir_path, new_filename)
                conflict_suffix += 1

            if self.mode == 'copy':
                shutil.copy2(pdf_path, target_file_path)
            elif self.mode == 'move':
                shutil.move(pdf_path, target_file_path)

            success_count += 1
            rel_archive_path = os.path.relpath(target_dir_path, self.archive_root).replace('\\', '/') if self.mode != 'dry_run' else target_dir_path.replace('\\', '/')
            
            code_val = fields.get('asset_no') if fields.get('asset_no') != '未知编号' else fields.get('serial_no', '—')
            if code_val and code_val != '—':
                scanned_asset_codes.add(code_val)

            record = {
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
                '源文件路径': os.path.abspath(pdf_path),
                '归档目标目录': rel_archive_path,
                '完整目标路径': os.path.abspath(target_file_path),
                '解析方式': 'RapidOCR 视觉识别' if used_ocr else 'PyMuPDF 文本层'
            }
            records.append(record)
            self.record_signal.emit(record)

            ocr_tag = " (⚡AI视觉)" if used_ocr else ""
            self.log_signal.emit("SUCCESS", f"✅ [{label}] [{device_type}]{ocr_tag} {filename} -> {new_filename}")

        temp_zip_dir = os.path.join(self.target_dir, "_temp_zip_extracted_")
        if os.path.exists(temp_zip_dir):
            try:
                shutil.rmtree(temp_zip_dir)
            except Exception:
                pass

        # ── 闭环比对台账中的待收回/缺漏设备清单
        missing_records = []
        is_quant_run = sum(1 for r in records if r['设备类型'] == '定量') > sum(1 for r in records if r['设备类型'] == '快检')
        
        target_dict = ledger.quantitative_map if is_quant_run else ledger.quick_check_map
        m_idx = 1
        for code, info in target_dict.items():
            if code not in scanned_asset_codes:
                missing_records.append({
                    '序号': m_idx,
                    '设备编号': code,
                    '仪器名称': info.get('inst', '—'),
                    '规格型号': info.get('model', '—'),
                    '所属组别/实验室': info.get('group') or info.get('lab', '—'),
                    '设备类型': info.get('type', '—'),
                    '溯源状态': '⚠️ 计划内未见 2026 证书 (待送检/待收回)'
                })
                m_idx += 1

        time_cost = (datetime.datetime.now() - start_time).total_seconds()
        now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

        excel_path = ""
        txt_path = ""
        if records:
            out_base = self.target_dir if os.path.exists(self.target_dir) else APP_ROOT
            excel_path = os.path.join(out_base, f'本次扫描_校准信息识别与台账核对表_{VERSION}.xlsx')
            df_rec = pd.DataFrame(records)
            df_mis = pd.DataFrame(missing_records) if missing_records else None
            export_styled_excel(df_rec, df_mis, excel_path)

            txt_path = os.path.join(out_base, f'本次处理_统计汇总_{now_str}.txt')
            with open(txt_path, 'w', encoding='utf-8') as tf:
                tf.write("=" * 70 + "\n")
                tf.write(f"校准证书智能识别与质量台账核对清单 ({VERSION})\n")
                tf.write(f"处理时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                tf.write(f"耗时：{time_cost:.2f} 秒\n")
                tf.write(f"待处理扫描路径：{os.path.abspath(self.target_dir)}\n")
                tf.write(f"归档目标总库：{os.path.abspath(self.archive_root)}\n")
                tf.write("=" * 70 + "\n\n")
                tf.write(f"总处理证书数量：{len(records)} 份\n")
                tf.write(f"  - 定量设备证书：{sum(1 for r in records if r['设备类型'] == '定量')} 份\n")
                tf.write(f"  - 快检设备证书：{sum(1 for r in records if r['设备类型'] == '快检')} 份\n")
                tf.write(f"  - 未分类证书：{sum(1 for r in records if r['设备类型'] == '未知')} 份\n\n")
                tf.write(f"📊 计划比对结果：台账中待收回/未扫描到证书设备共 {len(missing_records)} 台\n\n")

        self.finished_signal.emit({
            'records': records,
            'missing': missing_records,
            'total': len(records),
            'time_cost': time_cost,
            'excel_path': excel_path,
            'txt_path': txt_path,
            'error_count': error_count
        })


# ============================================================
# 第四部分：现代化主界面窗口 (QMainWindow)
# ============================================================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} {VERSION}")
        self.resize(1280, 880)
        self.setMinimumSize(1080, 750)
        self.setAcceptDrops(True)

        self.last_excel_path = ""
        self.worker = None
        self.presets = load_presets()
        self.records_cache = []
        self.missing_cache = []

        self.init_ui()
        self.apply_modern_stylesheet()

    def init_ui(self):
        central_widget = QWidget()
        central_widget.setObjectName("centralWidget")
        self.setCentralWidget(central_widget)

        root_layout = QHBoxLayout(central_widget)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # ── 1. 左侧现代化导航侧边栏 (Sidebar)
        sidebar = QFrame()
        sidebar.setObjectName("navSidebar")
        sidebar.setFixedWidth(240)
        sb_layout = QVBoxLayout(sidebar)
        sb_layout.setContentsMargins(16, 22, 16, 20)
        sb_layout.setSpacing(8)

        # 侧边栏 Logo 与标题
        logo_box = QVBoxLayout()
        logo_box.setSpacing(4)
        app_title = QLabel("FQT 质量工作台")
        app_title.setObjectName("sidebarAppTitle")
        app_subtitle = QLabel(f"校准管理与自动归档 {VERSION} · Antigravity 暗夜模式")
        app_subtitle.setObjectName("sidebarAppSubtitle")
        logo_box.addWidget(app_title)
        logo_box.addWidget(app_subtitle)
        sb_layout.addLayout(logo_box)
        sb_layout.addSpacing(16)

        # 导航按钮组
        self.nav_btn_group = QButtonGroup(self)
        self.btn_nav_quant = self._create_nav_btn("🏢 17025 定量校准", 0, True)
        self.btn_nav_quick = self._create_nav_btn("🧪 快检驻点校准", 1, False)
        self.btn_nav_audit = self._create_nav_btn("📊 计划比对审计", 2, False)
        self.btn_nav_log = self._create_nav_btn("📜 系统运行日志", 3, False)
        self.btn_nav_settings = self._create_nav_btn("⚙️ 预设与配置", 4, False)

        sb_layout.addWidget(self.btn_nav_quant)
        sb_layout.addWidget(self.btn_nav_quick)
        sb_layout.addWidget(self.btn_nav_audit)
        sb_layout.addWidget(self.btn_nav_log)
        sb_layout.addWidget(self.btn_nav_settings)
        sb_layout.addStretch()

        # 侧边栏底部引擎状态 Card
        engine_card = QFrame()
        engine_card.setObjectName("sidebarStatusCard")
        ec_layout = QVBoxLayout(engine_card)
        ec_layout.setContentsMargins(12, 12, 12, 12)
        ec_layout.setSpacing(6)

        ocr_status_txt = "🟢 RapidOCR 视觉就绪" if HAS_OCR else "🟡 基础文本层模式"
        lbl_ocr_stat = QLabel(ocr_status_txt)
        lbl_ocr_stat.setObjectName("sidebarOcrLabel")
        lbl_mupdf_stat = QLabel("🟢 PyMuPDF 核心已载入")
        lbl_mupdf_stat.setObjectName("sidebarPdfLabel")

        ec_layout.addWidget(lbl_ocr_stat)
        ec_layout.addWidget(lbl_mupdf_stat)
        sb_layout.addWidget(engine_card)

        root_layout.addWidget(sidebar)

        # ── 2. 右侧主工作区 (QStackedWidget)
        self.main_stack = QStackedWidget()
        self.main_stack.setObjectName("mainStackArea")

        # Stack Page 0 & 1: 主工作台页面 (定量 / 快检)
        self.page_quant = self._create_workbench_page(workbench_type='quant')
        self.page_quick = self._create_workbench_page(workbench_type='quick')
        
        # Stack Page 2: 计划比对与审计页面
        self.page_audit = self._create_audit_page()

        # Stack Page 3: 实时日志页面
        self.page_log = self._create_log_page()

        # Stack Page 4: 预设管理页面
        self.page_settings = self._create_settings_page()

        self.main_stack.addWidget(self.page_quant)
        self.main_stack.addWidget(self.page_quick)
        self.main_stack.addWidget(self.page_audit)
        self.main_stack.addWidget(self.page_log)
        self.main_stack.addWidget(self.page_settings)

        root_layout.addWidget(self.main_stack, stretch=1)

    def _create_nav_btn(self, text, page_index, is_checked=False):
        btn = QPushButton(text)
        btn.setCheckable(True)
        btn.setChecked(is_checked)
        btn.setObjectName("navButton")
        btn.setFixedHeight(42)
        btn.setCursor(QCursor(Qt.PointingHandCursor))
        self.nav_btn_group.addButton(btn, page_index)
        btn.clicked.connect(lambda: self.switch_page(page_index))
        return btn

    def switch_page(self, index):
        self.main_stack.setCurrentIndex(index)

    def _create_workbench_page(self, workbench_type='quant'):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        
        page_widget = QWidget()
        layout = QVBoxLayout(page_widget)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)

        is_quant = (workbench_type == 'quant')
        default_in = DEFAULT_PRESETS[0]['path'] if is_quant else DEFAULT_PRESETS[1]['path']

        # 1. 顶部 Header 标题卡片
        header = QFrame()
        header.setObjectName("bannerCard")
        h_box = QHBoxLayout(header)
        h_box.setContentsMargins(20, 16, 20, 16)

        v_title = QVBoxLayout()
        v_title.setSpacing(4)
        title_text = "🏢 17025 定量实验室设备校准工作台" if is_quant else "🧪 快检质量网络驻点校准工作台"
        desc_text = "自动关联《2026年年度计划汇总.xlsx》量值溯源总表 · 支持理化/微生物/液相/气相/元素/抽样/综合组" if is_quant else "自动关联《2026年各实验室仪器设备校准清单.xlsx》· 支持各驻点项目组、街道与快检室多维度核对"

        t_lbl = QLabel(title_text)
        t_lbl.setObjectName("bannerTitle")
        d_lbl = QLabel(desc_text)
        d_lbl.setObjectName("bannerSubtitle")
        d_lbl.setWordWrap(True)

        v_title.addWidget(t_lbl)
        v_title.addWidget(d_lbl)
        h_box.addLayout(v_title)
        h_box.addStretch()

        type_badge = QLabel("定量 17025 体系" if is_quant else "快检质量网络")
        type_badge.setObjectName("badgeType")
        h_box.addWidget(type_badge)
        layout.addWidget(header)

        # 2. 指标看板 (Metrics)
        metrics_grid = QGridLayout()
        metrics_grid.setSpacing(12)

        lbl_total = QLabel("0 份")
        lbl_matched = QLabel("0 份")
        lbl_groups = QLabel("0 组" if is_quant else "0 站")
        lbl_time = QLabel("0.00 秒")

        if is_quant:
            self.lbl_q_total = lbl_total
            self.lbl_q_matched = lbl_matched
            self.lbl_q_groups = lbl_groups
            self.lbl_q_time = lbl_time
        else:
            self.lbl_k_total = lbl_total
            self.lbl_k_matched = lbl_matched
            self.lbl_k_groups = lbl_groups
            self.lbl_k_time = lbl_time

        def make_metric(title, val_lbl, badge_tag, text_color):
            f = QFrame()
            f.setStyleSheet("QFrame { background: #18191E; border-radius: 8px; border: 1px solid #282A33; padding: 12px; }")
            vb = QVBoxLayout(f)
            vb.setSpacing(4)
            top_h = QHBoxLayout()
            t = QLabel(title)
            t.setStyleSheet("font-size: 12px; font-weight: 600; color: #9CA3AF;")
            top_h.addWidget(t)
            top_h.addStretch()
            badge = QLabel(badge_tag)
            badge.setStyleSheet(f"font-size: 11px; font-weight: bold; color: {text_color}; background-color: rgba(59, 130, 246, 0.12); padding: 2px 6px; border-radius: 4px; border: 1px solid {text_color}44;")
            top_h.addWidget(badge)
            val_lbl.setStyleSheet(f"font-size: 24px; font-weight: bold; color: {text_color};")
            vb.addLayout(top_h)
            vb.addWidget(val_lbl)
            return f

        metrics_grid.addWidget(make_metric("📦 扫描证书总数", lbl_total, "[总量]", "#3B82F6"), 0, 0)
        metrics_grid.addWidget(make_metric("✅ 台账精准匹配", lbl_matched, "[已核对]", "#34D399"), 0, 1)
        metrics_grid.addWidget(make_metric("🔬 覆盖组别/驻点", lbl_groups, "[组别/站]", "#FBBF24"), 0, 2)
        metrics_grid.addWidget(make_metric("⏱️ 处理任务耗时", lbl_time, "[耗时]", "#E4E4E7"), 0, 3)

        layout.addLayout(metrics_grid)

        # 3. 路径选择与交互投放卡片
        path_card = QFrame()
        path_card.setObjectName("modernCard")
        pc_layout = QVBoxLayout(path_card)
        pc_layout.setContentsMargins(18, 16, 18, 16)
        pc_layout.setSpacing(12)

        card_title = QLabel("📂 扫描路径与运行控制")
        card_title.setObjectName("cardTitle")
        pc_layout.addWidget(card_title)

        p_grid = QGridLayout()
        p_grid.setSpacing(10)
        
        lbl_in = QLabel("待识别目录:")
        lbl_in.setObjectName("formLabel")
        p_grid.addWidget(lbl_in, 0, 0)

        in_edit = QLineEdit(default_in)
        in_edit.setPlaceholderText("选择或直接拖拽待处理校准证书文件夹...")
        p_grid.addWidget(in_edit, 0, 1)

        btn_browse = QPushButton("📁 浏览目录")
        btn_browse.setObjectName("btnOutline")
        btn_browse.clicked.connect(lambda: self._browse_dir(in_edit))
        p_grid.addWidget(btn_browse, 0, 2)

        pc_layout.addLayout(p_grid)

        # 运行模式与功能勾选
        mode_row = QHBoxLayout()
        mode_row.setSpacing(18)
        
        rb_preview = QRadioButton("🔍 仅智能识别与结果返回 (安全预览/不改动文件)")
        rb_preview.setChecked(True)
        rb_copy = QRadioButton("📑 复制并归档")
        rb_move = QRadioButton("🚚 移动并归档")
        
        btn_grp = QButtonGroup(self)
        btn_grp.addButton(rb_preview)
        btn_grp.addButton(rb_copy)
        btn_grp.addButton(rb_move)

        mode_row.addWidget(rb_preview)
        mode_row.addWidget(rb_copy)
        mode_row.addWidget(rb_move)
        mode_row.addStretch()

        cb_ocr = QCheckBox("⚡ AI OCR 视觉抢救")
        cb_ocr.setChecked(HAS_OCR)
        cb_ocr.setEnabled(HAS_OCR)
        cb_unzip = QCheckBox("📦 自动解压 ZIP")
        cb_unzip.setChecked(True)

        mode_row.addWidget(cb_ocr)
        mode_row.addWidget(cb_unzip)
        pc_layout.addLayout(mode_row)

        # 核心按钮栏
        ctrl_bar = QHBoxLayout()
        ctrl_bar.setSpacing(12)

        btn_start = QPushButton("🚀 一键智能识别与核对")
        btn_start.setObjectName("btnPrimary")
        btn_start.setFixedHeight(42)
        btn_start.setCursor(QCursor(Qt.PointingHandCursor))

        btn_stop = QPushButton("⏹️ 终止")
        btn_stop.setObjectName("btnDanger")
        btn_stop.setFixedHeight(42)
        btn_stop.setEnabled(False)

        btn_export = QPushButton("📊 导出 Excel 审计报表")
        btn_export.setObjectName("btnAction")
        btn_export.setFixedHeight(42)
        btn_export.setEnabled(False)

        btn_open = QPushButton("📂 打开所在目录")
        btn_open.setObjectName("btnAction")
        btn_open.setFixedHeight(42)
        btn_open.clicked.connect(lambda: self._open_dir(in_edit.text()))

        ctrl_bar.addWidget(btn_start, stretch=3)
        ctrl_bar.addWidget(btn_stop, stretch=1)
        ctrl_bar.addWidget(btn_export, stretch=2)
        ctrl_bar.addWidget(btn_open, stretch=2)
        pc_layout.addLayout(ctrl_bar)

        layout.addWidget(path_card)

        # 4. 进度条
        prog_bar = QProgressBar()
        prog_bar.setValue(0)
        prog_bar.setFixedHeight(14)
        layout.addWidget(prog_bar)

        # 5. 数据表格与搜索过滤栏
        table_card = QFrame()
        table_card.setObjectName("modernCard")
        tc_layout = QVBoxLayout(table_card)
        tc_layout.setContentsMargins(14, 14, 14, 14)
        tc_layout.setSpacing(10)

        filter_bar = QHBoxLayout()
        filter_bar.setSpacing(10)

        search_edit = QLineEdit()
        search_edit.setPlaceholderText("🔍 全局即时搜索：输入设备编号、仪器名称、出厂编号或组别...")
        search_edit.setFixedHeight(34)
        filter_bar.addWidget(search_edit, stretch=2)

        group_combo = QComboBox()
        group_combo.setFixedHeight(34)
        group_combo.addItem("📂 所有组别 / 驻点")
        if is_quant:
            for g in STANDARD_GROUPS:
                group_combo.addItem(g)
        filter_bar.addWidget(group_combo, stretch=1)

        tc_layout.addLayout(filter_bar)

        table = QTableWidget(0, 9)
        table.setHorizontalHeaderLabels([
            "序号", "原文件名", "规范命名/建议重命名", "机构",
            "设备编号", "出厂编号", "仪器名称", "校准日期", "所属组别/驻点实验室"
        ])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        table.horizontalHeader().resizeSection(0, 50)
        table.horizontalHeader().resizeSection(1, 240)
        table.horizontalHeader().resizeSection(2, 260)
        table.horizontalHeader().resizeSection(3, 100)
        table.horizontalHeader().resizeSection(4, 100)
        table.horizontalHeader().resizeSection(5, 100)
        table.horizontalHeader().resizeSection(6, 140)
        table.horizontalHeader().resizeSection(7, 90)
        table.horizontalHeader().resizeSection(8, 140)
        table.setAlternatingRowColors(True)
        table.setContextMenuPolicy(Qt.CustomContextMenu)
        table.customContextMenuRequested.connect(lambda pos, t=table: self.show_table_menu(pos, t))
        table.itemDoubleClicked.connect(self.on_table_double_clicked)
        tc_layout.addWidget(table)

        layout.addWidget(table_card)

        # 绑定引用与事件
        if is_quant:
            self.table_quant = table
            self.search_quant = search_edit
            self.combo_quant = group_combo
            self.btn_q_start = btn_start
            self.btn_q_stop = btn_stop
            self.btn_q_export = btn_export
            self.prog_quant = prog_bar
            self.in_edit_quant = in_edit
            self.rb_q_preview = rb_preview
            self.rb_q_copy = rb_copy
            self.rb_q_move = rb_move
            self.cb_q_ocr = cb_ocr
            self.cb_q_unzip = cb_unzip
        else:
            self.table_quick = table
            self.search_quick = search_edit
            self.combo_quick = group_combo
            self.btn_k_start = btn_start
            self.btn_k_stop = btn_stop
            self.btn_k_export = btn_export
            self.prog_quick = prog_bar
            self.in_edit_quick = in_edit
            self.rb_k_preview = rb_preview
            self.rb_k_copy = rb_copy
            self.rb_k_move = rb_move
            self.cb_k_ocr = cb_ocr
            self.cb_k_unzip = cb_unzip

        btn_start.clicked.connect(lambda: self.start_worker(is_quant=is_quant))
        btn_stop.clicked.connect(self.stop_worker)
        btn_export.clicked.connect(self.open_excel_report)
        search_edit.textChanged.connect(lambda text, t=table, c=group_combo: self._filter_table(t, text, c.currentText()))
        group_combo.currentTextChanged.connect(lambda grp, t=table, s=search_edit: self._filter_table(t, s.text(), grp))

        scroll.setWidget(page_widget)
        return scroll

    def _create_audit_page(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)

        page_widget = QWidget()
        layout = QVBoxLayout(page_widget)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)

        header = QFrame()
        header.setObjectName("bannerCard")
        h_box = QHBoxLayout(header)
        h_box.setContentsMargins(20, 16, 20, 16)

        v_title = QVBoxLayout()
        v_title.setSpacing(4)
        t_lbl = QLabel("📊 计划台账闭环核对与量值溯源审计")
        t_lbl.setObjectName("bannerTitle")
        d_lbl = QLabel("对比年度校准计划 vs. 实际识别证书 · 精准定位待送检、漏检及未收回证书清单")
        d_lbl.setObjectName("bannerSubtitle")
        v_title.addWidget(t_lbl)
        v_title.addWidget(d_lbl)
        h_box.addLayout(v_title)
        h_box.addStretch()

        btn_exp_audit = QPushButton("📊 导出审计明细 Excel")
        btn_exp_audit.setObjectName("btnAction")
        btn_exp_audit.setFixedHeight(38)
        btn_exp_audit.clicked.connect(self.open_excel_report)
        h_box.addWidget(btn_exp_audit)
        layout.addWidget(header)

        # 缺漏表格卡片
        audit_card = QFrame()
        audit_card.setObjectName("modernCard")
        ac_layout = QVBoxLayout(audit_card)
        ac_layout.setContentsMargins(16, 16, 16, 16)
        ac_layout.setSpacing(10)

        card_title = QLabel("⚠️ 计划内待收回 / 未见 2026 校准证书设备清单")
        card_title.setObjectName("cardTitle")
        ac_layout.addWidget(card_title)

        self.table_audit = QTableWidget(0, 7)
        self.table_audit.setHorizontalHeaderLabels([
            "序号", "设备编号", "仪器名称", "规格型号", "所属组别/实验室", "设备类型", "核对审计状态"
        ])
        self.table_audit.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_audit.horizontalHeader().resizeSection(0, 50)
        self.table_audit.horizontalHeader().resizeSection(1, 120)
        self.table_audit.horizontalHeader().resizeSection(2, 200)
        self.table_audit.horizontalHeader().resizeSection(3, 160)
        self.table_audit.horizontalHeader().resizeSection(4, 160)
        self.table_audit.horizontalHeader().resizeSection(5, 100)
        self.table_audit.horizontalHeader().resizeSection(6, 260)
        self.table_audit.setAlternatingRowColors(True)
        ac_layout.addWidget(self.table_audit)

        layout.addWidget(audit_card)
        scroll.setWidget(page_widget)
        return scroll

    def _create_log_page(self):
        page_widget = QWidget()
        layout = QVBoxLayout(page_widget)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        header = QFrame()
        header.setObjectName("bannerCard")
        h_box = QHBoxLayout(header)
        h_box.setContentsMargins(20, 16, 20, 16)
        t_lbl = QLabel("📜 系统运行实时日志流")
        t_lbl.setObjectName("bannerTitle")
        h_box.addWidget(t_lbl)
        h_box.addStretch()

        btn_clear = QPushButton("🗑️ 清空日志")
        btn_clear.setObjectName("btnOutline")
        btn_clear.clicked.connect(lambda: self.log_text.clear())
        h_box.addWidget(btn_clear)
        layout.addWidget(header)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setObjectName("modernLogView")
        layout.addWidget(self.log_text)

        return page_widget

    def _create_settings_page(self):
        page_widget = QWidget()
        layout = QVBoxLayout(page_widget)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)

        header = QFrame()
        header.setObjectName("bannerCard")
        h_box = QHBoxLayout(header)
        h_box.setContentsMargins(20, 16, 20, 16)
        t_lbl = QLabel("⚙️ 常用路径预设与系统环境")
        t_lbl.setObjectName("bannerTitle")
        h_box.addWidget(t_lbl)
        layout.addWidget(header)

        preset_card = QFrame()
        preset_card.setObjectName("modernCard")
        pc_layout = QVBoxLayout(preset_card)
        pc_layout.setContentsMargins(18, 16, 18, 16)
        pc_layout.setSpacing(12)

        pc_title = QLabel("🏷️ 常用归档与证书库快捷方式")
        pc_title.setObjectName("cardTitle")
        pc_layout.addWidget(pc_title)

        self.preset_list_widget = QListWidget()
        self.refresh_settings_preset_list()
        pc_layout.addWidget(self.preset_list_widget)

        btn_bar = QHBoxLayout()
        btn_add = QPushButton("➕ 添加新预设")
        btn_add.clicked.connect(self.add_preset_item)
        btn_del = QPushButton("🗑️ 删除选中")
        btn_del.clicked.connect(self.del_preset_item)
        btn_bar.addWidget(btn_add)
        btn_bar.addWidget(btn_del)
        btn_bar.addStretch()
        pc_layout.addLayout(btn_bar)

        layout.addWidget(preset_card)
        layout.addStretch()
        return page_widget

    def refresh_settings_preset_list(self):
        self.preset_list_widget.clear()
        for p in self.presets:
            item = QListWidgetItem(f"📁 {p['name']} \n    {p['path']}")
            self.preset_list_widget.addItem(item)

    def add_preset_item(self):
        dir_p = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if not dir_p: return
        name, ok = QInputDialog.getText(self, "预设名称", "请输入快捷方式名称:", text=os.path.basename(dir_p))
        if ok and name.strip():
            self.presets.append({"name": name.strip(), "path": os.path.abspath(dir_p)})
            save_presets(self.presets)
            self.refresh_settings_preset_list()

    def del_preset_item(self):
        row = self.preset_list_widget.currentRow()
        if row >= 0:
            del self.presets[row]
            save_presets(self.presets)
            self.refresh_settings_preset_list()

    def _browse_dir(self, line_edit):
        p = QFileDialog.getExistingDirectory(self, "选择文件夹", line_edit.text() or APP_ROOT)
        if p: line_edit.setText(os.path.abspath(p))

    def _open_dir(self, path):
        if os.path.exists(path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.abspath(path)))
        else:
            QMessageBox.information(self, "提示", f"目录不存在：\n{path}")

    def _filter_table(self, table, query, group_filter):
        q = query.strip().lower()
        grp = group_filter.strip()
        is_all_grp = ("所有" in grp)
        
        for r in range(table.rowCount()):
            match_q = True
            if q:
                row_txt = ' '.join([table.item(r, c).text().lower() for c in range(table.columnCount()) if table.item(r, c)])
                match_q = (q in row_txt)
            match_grp = True
            if not is_all_grp:
                grp_val = table.item(r, 8).text() if table.item(r, 8) else ''
                match_grp = (grp in grp_val)
            table.setRowHidden(r, not (match_q and match_grp))

    def append_log(self, level, message):
        color_map = {
            "INFO": "#E4E4E7",
            "SUCCESS": "#60A5FA",
            "WARN": "#FBBF24",
            "ERROR": "#F87171"
        }
        tag_map = {
            "INFO": "[信息]",
            "SUCCESS": "[成功]",
            "WARN": "[警告]",
            "ERROR": "[错误]"
        }
        color = color_map.get(level, "#E4E4E7")
        tag = tag_map.get(level, f"[{level}]")
        now_time = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f'<span style="color:#71717A;">[{now_time}]</span> <span style="color:{color}; font-weight:bold;">{tag}</span> <span style="color:{color};">{message}</span>')

    def start_worker(self, is_quant=True):
        if is_quant:
            target_dir = self.in_edit_quant.text().strip()
            table = self.table_quant
            prog = self.prog_quant
            btn_start = self.btn_q_start
            btn_stop = self.btn_q_stop
            btn_exp = self.btn_q_export
            mode = 'dry_run'
            if self.rb_q_copy.isChecked(): mode = 'copy'
            elif self.rb_q_move.isChecked(): mode = 'move'
            enable_ocr = self.cb_q_ocr.isChecked()
            auto_unzip = self.cb_q_unzip.isChecked()
        else:
            target_dir = self.in_edit_quick.text().strip()
            table = self.table_quick
            prog = self.prog_quick
            btn_start = self.btn_k_start
            btn_stop = self.btn_k_stop
            btn_exp = self.btn_k_export
            mode = 'dry_run'
            if self.rb_k_copy.isChecked(): mode = 'copy'
            elif self.rb_k_move.isChecked(): mode = 'move'
            enable_ocr = self.cb_k_ocr.isChecked()
            auto_unzip = self.cb_k_unzip.isChecked()

        if not target_dir or not os.path.exists(target_dir):
            QMessageBox.warning(self, "路径错误", "待处理路径不存在，请先选择有效的文件夹！")
            return

        archive_root = os.path.join(APP_ROOT, "【归档完成】校准证书库")

        table.setRowCount(0)
        prog.setValue(0)
        btn_start.setEnabled(False)
        btn_stop.setEnabled(True)
        btn_exp.setEnabled(False)
        self.records_cache = []
        self.missing_cache = []

        self.append_log("INFO", "=" * 60)
        self.append_log("INFO", f"🏁 任务启动：{'17025 定量校准' if is_quant else '快检驻点校准'}")
        self.append_log("INFO", f"   扫描路径：{target_dir}")
        self.append_log("INFO", f"   运行模式：{mode.upper()}")
        self.append_log("INFO", "=" * 60)

        self.worker = ArchiveWorker(
            target_dir=target_dir,
            archive_root=archive_root,
            mode=mode,
            enable_ocr=enable_ocr,
            auto_unzip=auto_unzip
        )
        self.worker.progress_signal.connect(lambda cur, tot, fn, p=prog: self.on_progress(cur, tot, fn, p))
        self.worker.log_signal.connect(self.append_log)
        self.worker.record_signal.connect(lambda rec, t=table: self.on_record_extracted(rec, t))
        self.worker.finished_signal.connect(lambda sm, q=is_quant: self.on_worker_finished(sm, q))
        self.worker.start()

    def stop_worker(self):
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.append_log("WARN", "🛑 正在请求中止任务...")

    def on_progress(self, current, total, filename, prog_bar):
        val = int((current / total) * 100)
        prog_bar.setValue(val)

    def on_record_extracted(self, rec, table):
        self.records_cache.append(rec)
        r = table.rowCount()
        table.insertRow(r)
        
        items = [
            str(rec['序号']),
            rec['原始文件名'],
            rec['最终重命名'],
            rec['机构'],
            rec['设备编号'],
            rec['出厂编号'],
            rec['仪器名称'],
            rec['校准日期'],
            rec['所属组别/实验室']
        ]
        
        for c, text_val in enumerate(items):
            item = QTableWidgetItem(str(text_val))
            item.setToolTip(str(text_val))
            if c in (0, 3, 4, 7):
                item.setTextAlignment(Qt.AlignCenter)
            else:
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            table.setItem(r, c, item)
            
        table.scrollToBottom()

    def on_worker_finished(self, summary, is_quant):
        records = summary.get('records', [])
        missing = summary.get('missing', [])
        total = len(records)
        time_cost = summary.get('time_cost', 0)
        self.last_excel_path = summary.get('excel_path', '')
        self.missing_cache = missing

        if is_quant:
            self.btn_q_start.setEnabled(True)
            self.btn_q_stop.setEnabled(False)
            self.btn_q_export.setEnabled(bool(self.last_excel_path))
            self.prog_quant.setValue(100)
            self.lbl_q_total.setText(f"{total} 份")
            self.lbl_q_matched.setText(f"{sum(1 for r in records if r['设备编号'] not in {'未知编号', '—'})} 份")
            unique_grps = set(r['所属组别/实验室'] for r in records if r['所属组别/实验室'] not in {'—', '未知组别'})
            self.lbl_q_groups.setText(f"{len(unique_grps)} 组")
            self.lbl_q_time.setText(f"{time_cost:.2f} 秒")
        else:
            self.btn_k_start.setEnabled(True)
            self.btn_k_stop.setEnabled(False)
            self.btn_k_export.setEnabled(bool(self.last_excel_path))
            self.prog_quick.setValue(100)
            self.lbl_k_total.setText(f"{total} 份")
            self.lbl_k_matched.setText(f"{sum(1 for r in records if r['设备编号'] not in {'未知编号', '—'})} 份")
            unique_labs = set(r['所属组别/实验室'] for r in records if r['所属组别/实验室'] not in {'—', '未知实验室'})
            self.lbl_k_groups.setText(f"{len(unique_labs)} 站")
            self.lbl_k_time.setText(f"{time_cost:.2f} 秒")

        # 填充审计表格
        self.table_audit.setRowCount(0)
        for idx, m in enumerate(missing, 1):
            r = self.table_audit.rowCount()
            self.table_audit.insertRow(r)
            self.table_audit.setItem(r, 0, QTableWidgetItem(str(idx)))
            self.table_audit.setItem(r, 1, QTableWidgetItem(m['设备编号']))
            self.table_audit.setItem(r, 2, QTableWidgetItem(m['仪器名称']))
            self.table_audit.setItem(r, 3, QTableWidgetItem(m['规格型号']))
            self.table_audit.setItem(r, 4, QTableWidgetItem(m['所属组别/实验室']))
            self.table_audit.setItem(r, 5, QTableWidgetItem(m['设备类型']))
            self.table_audit.setItem(r, 6, QTableWidgetItem(m['溯源状态']))

        self.append_log("SUCCESS", f"🎉 处理完成！解析 {total} 份，计划比对待收回 {len(missing)} 台，耗时 {time_cost:.2f} 秒")

        if total > 0:
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("识别与核对完成")
            msg_box.setIcon(QMessageBox.Information)
            msg_box.setText(f"🎉 <b>校准证书识别与计划比对完成！</b><br><br>"
                            f"📊 <b>解析总数：</b>{total} 份 (耗时 {time_cost:.2f} 秒)<br>"
                            f"🔍 <b>计划待收回：</b>{len(missing)} 台设备尚未扫描到 2026 证书<br><br>"
                            f"📋 <b>审计报表：</b>{os.path.basename(self.last_excel_path)}<br>")
            btn_exp = msg_box.addButton("📊 查看汇总 Excel", QMessageBox.ActionRole)
            btn_audit = msg_box.addButton("🔍 查看待收回清单", QMessageBox.ActionRole)
            btn_ok = msg_box.addButton("确定", QMessageBox.AcceptRole)
            msg_box.exec_()

            if msg_box.clickedButton() == btn_exp:
                self.open_excel_report()
            elif msg_box.clickedButton() == btn_audit:
                self.btn_nav_audit.click()

    def open_excel_report(self):
        if self.last_excel_path and os.path.exists(self.last_excel_path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.abspath(self.last_excel_path)))
        else:
            QMessageBox.information(self, "提示", "暂未找到生成的 Excel 报表！")

    def show_table_menu(self, pos, table):
        item = table.itemAt(pos)
        if not item: return
        menu = QMenu(self)
        row = item.row()
        new_name = table.item(row, 2).text()
        
        act_copy = QAction(f"📋 复制规范文件名 ({new_name})", self)
        act_copy.triggered.connect(lambda: QApplication.clipboard().setText(new_name))
        menu.addAction(act_copy)

        if row < len(self.records_cache):
            rec = self.records_cache[row]
            src_p = rec.get('源文件路径', '')
            if src_p and os.path.exists(src_p):
                act_locate = QAction("📂 在资源管理器中定位源文件", self)
                act_locate.triggered.connect(lambda: os.system(f'explorer /select,"{src_p}"'))
                menu.addAction(act_locate)

        menu.exec_(QCursor.pos())

    def on_table_double_clicked(self, item):
        table = self.sender()
        row = item.row()
        if row < len(self.records_cache):
            rec = self.records_cache[row]
            src_p = rec.get('源文件路径', '')
            if src_p and os.path.exists(src_p):
                QDesktopServices.openUrl(QUrl.fromLocalFile(src_p))

    def apply_modern_stylesheet(self):
        self.setStyleSheet("""
            QWidget {
                font-family: "Microsoft YaHei UI", "Segoe UI", "PingFang SC", sans-serif;
                font-size: 13px;
                color: #F4F4F5;
            }
            QWidget#centralWidget {
                background-color: #121316;
            }
            #navSidebar {
                background-color: #16171B;
                border-right: 1px solid #27282D;
            }
            #sidebarAppTitle {
                color: #F4F4F5;
                font-size: 16px;
                font-weight: bold;
                letter-spacing: 0.5px;
            }
            #sidebarAppSubtitle {
                color: #9CA3AF;
                font-size: 11px;
            }
            #navButton {
                text-align: left;
                padding-left: 14px;
                color: #9CA3AF;
                font-size: 13px;
                font-weight: 600;
                background-color: transparent;
                border-radius: 6px;
                border: none;
            }
            #navButton:hover {
                background-color: #22242B;
                color: #F4F4F5;
            }
            #navButton:checked {
                background-color: #3B82F6;
                color: #FFFFFF;
                font-weight: bold;
            }
            #sidebarStatusCard {
                background-color: #1A1B20;
                border: 1px solid #282A33;
                border-radius: 8px;
            }
            #sidebarOcrLabel {
                color: #38BDF8;
                font-size: 11px;
                font-weight: bold;
            }
            #sidebarPdfLabel {
                color: #34D399;
                font-size: 11px;
            }
            #bannerCard {
                background-color: #18191E;
                border: 1px solid #2A2C34;
                border-radius: 10px;
            }
            #bannerTitle {
                color: #F4F4F5;
                font-size: 18px;
                font-weight: bold;
            }
            #bannerSubtitle {
                color: #9CA3AF;
                font-size: 12px;
            }
            #badgeType {
                background-color: rgba(59, 130, 246, 0.18);
                color: #60A5FA;
                padding: 5px 14px;
                border-radius: 6px;
                font-size: 12px;
                font-weight: bold;
                border: 1px solid #3B82F6;
            }
            #modernCard {
                background-color: #18191E;
                border: 1px solid #282A33;
                border-radius: 10px;
            }
            #cardTitle {
                font-size: 14px;
                font-weight: bold;
                color: #F4F4F5;
            }
            #formLabel {
                font-weight: bold;
                color: #D4D4D8;
            }
            QLineEdit, QComboBox {
                padding: 7px 12px;
                border: 1px solid #32353E;
                border-radius: 6px;
                background-color: #1A1B20;
                color: #F4F4F5;
                font-size: 13px;
                selection-background-color: #3B82F6;
            }
            QLineEdit:focus, QComboBox:focus {
                border: 2px solid #3B82F6;
                background-color: #22242B;
            }
            QComboBox QAbstractItemView {
                background-color: #1E1F26;
                color: #F4F4F5;
                selection-background-color: #3B82F6;
                selection-color: #FFFFFF;
                border: 1px solid #32353E;
            }
            QRadioButton, QCheckBox {
                color: #E4E4E7;
                font-size: 13px;
                spacing: 6px;
            }
            QRadioButton::indicator, QCheckBox::indicator {
                width: 16px;
                height: 16px;
                border: 1px solid #4B5563;
                border-radius: 3px;
                background: #1A1B20;
            }
            QRadioButton::indicator:checked, QCheckBox::indicator:checked {
                background-color: #3B82F6;
                border-color: #3B82F6;
            }
            #btnOutline {
                padding: 7px 16px;
                border: 1px solid #383A44;
                border-radius: 6px;
                background-color: #202228;
                font-weight: 600;
                color: #E4E4E7;
            }
            #btnOutline:hover {
                background-color: #2A2C34;
                border-color: #3B82F6;
                color: #FFFFFF;
            }
            #btnPrimary {
                background-color: #3B82F6;
                color: #FFFFFF;
                font-size: 14px;
                font-weight: bold;
                border-radius: 6px;
                border: none;
            }
            #btnPrimary:hover {
                background-color: #2563EB;
            }
            #btnPrimary:pressed {
                background-color: #1D4ED8;
            }
            #btnPrimary:disabled {
                background-color: #2A3548;
                color: #64748B;
            }
            #btnDanger {
                background-color: #DC2626;
                color: #FFFFFF;
                font-size: 13px;
                font-weight: bold;
                border-radius: 6px;
                border: none;
            }
            #btnDanger:hover {
                background-color: #B91C1C;
            }
            #btnDanger:disabled {
                background-color: #3D1C1C;
                color: #64748B;
            }
            #btnAction {
                background-color: #202228;
                border: 1px solid #383A44;
                border-radius: 6px;
                color: #E4E4E7;
                font-weight: 600;
                font-size: 13px;
            }
            #btnAction:hover {
                background-color: #2A2C34;
                border-color: #3B82F6;
                color: #FFFFFF;
            }
            #btnAction:disabled {
                background-color: #1A1B20;
                border-color: #282A33;
                color: #52525B;
            }
            QProgressBar {
                border: 1px solid #282A33;
                border-radius: 4px;
                text-align: center;
                background-color: #1A1B20;
                color: #F4F4F5;
                font-weight: bold;
            }
            QProgressBar::chunk {
                background-color: #3B82F6;
                border-radius: 3px;
            }
            QTableWidget {
                border: 1px solid #282A33;
                border-radius: 6px;
                background-color: #16171B;
                gridline-color: #26282E;
                color: #F4F4F5;
                font-size: 12px;
                selection-background-color: #2563EB;
                selection-color: #FFFFFF;
            }
            QHeaderView::section {
                background-color: #202228;
                color: #E4E4E7;
                padding: 8px;
                font-weight: bold;
                border: none;
                border-bottom: 2px solid #3B82F6;
            }
            #modernLogView {
                border: 1px solid #26282E;
                border-radius: 8px;
                background-color: #0E0F12;
                color: #F4F4F5;
                font-family: "Consolas", "Courier New", monospace;
                font-size: 12px;
                line-height: 1.5;
                padding: 12px;
            }
            QScrollBar:vertical {
                background-color: #16171B;
                width: 10px;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background-color: #32353E;
                min-height: 20px;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #4B5563;
            }
            QScrollBar:horizontal {
                background-color: #16171B;
                height: 10px;
                margin: 0px;
            }
            QScrollBar::handle:horizontal {
                background-color: #32353E;
                min-width: 20px;
                border-radius: 5px;
            }
            QScrollBar::handle:horizontal:hover {
                background-color: #4B5563;
            }
            QScrollBar::add-line, QScrollBar::sub-line {
                width: 0px;
                height: 0px;
            }
            QListWidget {
                background-color: #1A1B20;
                border: 1px solid #2E3038;
                border-radius: 6px;
                color: #F4F4F5;
            }
            QListWidget::item {
                padding: 8px;
                border-bottom: 1px solid #26282E;
            }
            QListWidget::item:selected {
                background-color: #2563EB;
                color: #FFFFFF;
            }
            QMessageBox, QDialog {
                background-color: #1A1B20;
                color: #F4F4F5;
            }
            QMenu {
                background-color: #1E1F26;
                color: #F4F4F5;
                border: 1px solid #32353E;
            }
            QMenu::item:selected {
                background-color: #2563EB;
                color: #FFFFFF;
            }
        """)


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
