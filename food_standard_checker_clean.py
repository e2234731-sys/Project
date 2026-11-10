# -*- coding: utf-8 -*-
"""
食品安全国家标准判定核查工具 V1.4
标准覆盖：GB 2760-2024, GB 2761-2017, GB 2762-2025, GB 2763-2026, GB 29921-2021, GB 31650-2019 及 2026年国家食品安全监督抽检实施细则
Powered by 陈耿忠
"""
import multiprocessing
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import json
import os
import sys
import re
import csv
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    import openpyxl.styles.fills
    _orig_fill_from_tree = openpyxl.styles.fills.Fill.from_tree
    openpyxl.styles.fills.Fill.from_tree = lambda el: _orig_fill_from_tree(el) or openpyxl.styles.fills.PatternFill()
except ImportError:
    openpyxl = None
from datetime import datetime
def highlight_keywords_text(text, keywords):
    """在文本中为命中的关键词添加【】突出标记"""
    if not text or not keywords:
        return text if text is not None else ""
    valid_kws = [str(kw).strip() for kw in keywords if kw and str(kw).strip()]
    if not valid_kws:
        return str(text)
    valid_kws.sort(key=len, reverse=True)
    pattern = re.compile('(' + '|'.join(re.escape(k) for k in valid_kws) + ')', re.IGNORECASE)
    return pattern.sub(r'【\1】', str(text))

def strip_highlights(text):
    """去除【】高亮标记"""
    if not text:
        return ""
    return str(text).replace('【', '').replace('】', '')

def get_config_path():
    try:
        app_dir = os.path.join(os.path.expanduser("~"), ".food_standard_checker")
        os.makedirs(app_dir, exist_ok=True)
        return os.path.join(app_dir, "ui_config.json")
    except Exception:
        return "ui_config.json"

def load_ui_config():
    try:
        path = get_config_path()
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception as e:
        print(f"Failed to load UI config: {e}")
    return {}

def save_ui_config(config):
    try:
        path = get_config_path()
        with open(path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"Failed to save UI config: {e}")


# ============================================================
# 数据加载（支持PyInstaller打包）
# ============================================================
def get_data_path(filename):
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, filename)
    p1 = os.path.join(os.path.dirname(__file__), '..', 'output', filename)
    if os.path.exists(p1):
        return p1
    p2 = os.path.join(os.path.dirname(__file__), filename)
    if os.path.exists(p2):
        return p2
    return p1

def load_json_safe(filename):
    try:
        with open(get_data_path(filename), 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Warning: Cannot load {filename}: {e}")
        return None

def load_standards():
    """加载所有标准数据，并为每个条目注入官方要求/推荐的检验方法标准"""
    standards = {}
    
    # 加载GB 2761-2017
    gb2761_raw = load_json_safe('gb2761_data.json')
    if gb2761_raw:
        items = []
        for rec in gb2761_raw:
            cat = rec.get('food_category', '')
            t_item = rec.get('test_item_cn', '')
            method_str = resolve_item_test_method('', t_item, cat, 'GB 2761')
            items.append({
                'food_category': cat,
                'test_item': t_item,
                'limit_value': rec.get('limit_value', ''),
                'unit': '',
                'method': method_str,
                'notes': rec.get('test_item_en', '')
            })
        standards["GB 2761-2017"] = {
            'name': '食品中真菌毒素限量',
            'items': items
        }
    
    # 加载GB 2762-2025
    gb2762_raw = load_json_safe('gb2762_2025_data.json')
    if gb2762_raw:
        items = []
        for rec in gb2762_raw:
            cat = rec.get('food_category', '')
            t_item = rec.get('test_item', '')
            method_str = resolve_item_test_method(rec.get('method', ''), t_item, cat, 'GB 2762')
            items.append({
                'food_category': cat,
                'test_item': t_item,
                'limit_value': rec.get('limit_value', ''),
                'unit': '',
                'method': method_str,
                'notes': rec.get('notes', '')
            })
        standards["GB 2762-2025"] = {
            'name': '食品中污染物限量',
            'items': items
        }
    
    # 加载GB 29921-2021
    gb29921_raw = load_json_safe('gb29921_data.json')
    if gb29921_raw:
        items = []
        for rec in gb29921_raw:
            cat = rec.get('category_name', '')
            t_item = rec.get('test_item', '')
            method_str = resolve_item_test_method(rec.get('method', ''), t_item, cat, 'GB 29921')
            items.append({
                'food_category': cat,
                'test_item': t_item,
                'limit_value': rec.get('limit_value', ''),
                'unit': '',
                'method': method_str,
                'notes': rec.get('notes', '')
            })
        standards["GB 29921-2021"] = {
            'name': '预包装食品中致病菌限量',
            'items': items
        }
    
    # 加载GB 31650-2019
    gb31650_raw = load_json_safe('gb31650_data.json')
    if gb31650_raw:
        items = []
        for rec in gb31650_raw:
            cat = f"{rec.get('species', '')}-{rec.get('target_tissue', '')}"
            t_item = rec.get('drug_name', '')
            method_str = resolve_item_test_method('', t_item, cat, 'GB 31650')
            items.append({
                'food_category': cat,
                'test_item': t_item,
                'limit_value': rec.get('limit_value', ''),
                'unit': 'μg/kg',
                'method': method_str,
                'notes': f"{rec.get('drug_category', '')} | 残留标志物:{rec.get('mark_residue', '')}"
            })
        standards["GB 31650-2019"] = {
            'name': '食品中兽药最大残留限量',
            'items': items
        }
    
    # 加载2763数据
    gb2763_raw = load_json_safe('gb2763_data.json')
    cat2763_map = load_json_safe('category_map_2763.json')
    if gb2763_raw and cat2763_map:
        items = []
        for rec in gb2763_raw:
            cid = rec.get('category_id', 0)
            cat_name = cat2763_map.get(str(cid), f"分类#{cid}")
            t_item = rec['pesticide_cn']
            method_str = resolve_item_test_method(rec.get('method', ''), t_item, cat_name, 'GB 2763')
            items.append({
                'food_category': cat_name,
                'test_item': t_item,
                'limit_value': rec['mrl'],
                'unit': 'mg/kg',
                'method': method_str,
                'notes': f"{rec.get('function','')} | ADI:{rec.get('adi','')} | 残留物:{rec.get('residue','')}"
            })
        standards["GB 2763-2026"] = {
            'name': '食品中农药最大残留限量',
            'items': items
        }
    
    # 加载2760数据
    gb2760_raw = load_json_safe('gb2760_data.json')
    cat2760_map = load_json_safe('category_map_2760.json')
    if gb2760_raw and cat2760_map:
        items = []
        for rec in gb2760_raw:
            cid = rec.get('category_id', 0)
            cat_name = cat2760_map.get(str(cid), f"分类#{cid}")
            t_item = rec['additive_cn']
            method_str = resolve_item_test_method('', t_item, cat_name, 'GB 2760')
            items.append({
                'food_category': cat_name,
                'test_item': t_item,
                'limit_value': rec['max_use'],
                'unit': 'g/kg',
                'method': method_str,
                'notes': f"{rec.get('additive_en','')} | {rec.get('notes','')}"
            })
        standards["GB 2760-2024"] = {
            'name': '食品添加剂使用标准',
            'items': items
        }
    
    return standards

# ============================================================
# 检验方法标准映射数据库
# ============================================================
TEST_METHOD_DB = {
    # 重金属与理化元素
    "铅": "GB 5009.12-2023 食品中铅的测定（石墨炉原子吸收光谱法 / ICP-MS）",
    "镉": "GB 5009.15-2023 食品中镉的测定（石墨炉原子吸收光谱法 / ICP-MS）",
    "总汞": "GB 5009.17-2021 食品中总汞及有机汞的测定（原子荧光光谱法）",
    "甲基汞": "GB 5009.17-2021 食品中总汞及有机汞的测定（液相色谱-原子荧光联用法）",
    "汞": "GB 5009.17-2021 食品中总汞及有机汞的测定（原子荧光光谱法）",
    "总砷": "GB 5009.11-2024 食品中总砷及无机砷的测定（氢化物原子荧光法 / ICP-MS）",
    "无机砷": "GB 5009.11-2024 食品中总砷及无机砷的测定（液相色谱-原子荧光联用法）",
    "砷": "GB 5009.11-2024 食品中总砷及无机砷的测定（氢化物发生原子荧光法）",
    "铬": "GB 5009.123-2023 食品中铬的测定（石墨炉原子吸收光谱法 / ICP-MS）",
    "锡": "GB 5009.16-2023 食品中锡的测定（氢化物发生原子荧光光谱法）",
    "镍": "GB 5009.138-2017 食品中镍的测定（石墨炉原子吸收光谱法）",
    "铝": "GB 5009.182-2017 食品中铝的测定（电感耦合等离子体质谱法 / 分光光度法）",
    "铜": "GB 5009.13-2017 食品中铜的测定（火焰原子吸收法）",
    "锌": "GB 5009.14-2017 食品中锌的测定（火焰原子吸收法）",
    "铁": "GB 5009.90-2016 食品中铁的测定",
    # 污染物
    "苯并[a]芘": "GB 5009.27-2016 食品中苯并(a)芘的测定（同位素稀释GC-MS / HPLC-FLD）",
    "N-二甲基亚硝胺": "GB 5009.26-2023 食品中N-亚硝基化合物的测定（GC-MS/MS法）",
    "多氯联苯": "GB 5009.190-2014 食品中指示性多氯联苯含量的测定（气相色谱-质谱法）",
    "丙烯酰胺": "GB 5009.204-2024 食品中丙烯酰胺的测定（同位素稀释LC-MS/MS）",
    "3-氯-1,2-丙二醇": "GB 5009.191-2024 食品中氯丙醇及其脂肪酸酯的测定",
    # 真菌毒素
    "黄曲霉毒素B1": "GB 5009.22-2016 食品中黄曲霉毒素B族和G族的测定（同位素稀释LC-MS/MS）",
    "黄曲霉毒素M1": "GB 5009.24-2016 食品中黄曲霉毒素M族的测定（同位素稀释LC-MS/MS）",
    "脱氧雪腐镰刀菌烯醇": "GB 5009.111-2016 食品中脱氧雪腐镰刀菌烯醇的测定（LC-MS/MS）",
    "赭曲霉毒素A": "GB 5009.96-2016 食品中赭曲霉毒素A的测定（LC-MS/MS / HPLC法）",
    "展青霉素": "GB 5009.185-2016 食品中展青霉素的测定（同位素稀释LC-MS/MS / HPLC法）",
    "玉米赤霉烯酮": "GB 5009.209-2016 食品中玉米赤霉烯酮的测定（同位素稀释LC-MS/MS）",
    "伏马毒素": "GB 5009.240-2016 食品中伏马毒素的测定",
    "杂色曲霉素": "GB 5009.244-2016 食品中杂色曲霉素的测定",
    "T-2毒素": "GB 5009.118-2016 食品中T-2毒素的测定",
    # 食品添加剂
    "苯甲酸": "GB 5009.28-2024 食品中苯甲酸、山梨酸和糖精钠的测定（高效液相色谱法）",
    "山梨酸": "GB 5009.28-2024 食品中苯甲酸、山梨酸和糖精钠的测定（高效液相色谱法）",
    "糖精钠": "GB 5009.28-2024 食品中苯甲酸、山梨酸和糖精钠的测定（高效液相色谱法）",
    "脱氢乙酸": "GB 5009.121-2016 食品中脱氢乙酸的测定（气相色谱法 / 高效液相色谱法）",
    "甜蜜素": "GB 5009.97-2023 食品中环己基氨基磺酸盐的测定（LC-MS/MS / 气相色谱法）",
    "环己基氨基磺酸": "GB 5009.97-2023 食品中环己基氨基磺酸盐的测定",
    "安赛蜜": "GB 5009.140-2023 食品中乙酰磺胺酸钾的测定（高效液相色谱法）",
    "乙酰磺胺酸钾": "GB 5009.140-2023 食品中乙酰磺胺酸钾的测定（高效液相色谱法）",
    "二氧化硫": "GB 5009.34-2022 食品中二氧化硫的测定（酸碱滴定法 / 气相色谱法 / 离子色谱法）",
    "亚硝酸盐": "GB 5009.33-2024 食品中亚硝酸盐与硝酸盐的测定（离子色谱法 / 分光光度法）",
    "硝酸盐": "GB 5009.33-2024 食品中亚硝酸盐与硝酸盐的测定（离子色谱法 / 分光光度法）",
    "纳他霉素": "GB 5009.286-2022 食品中纳他霉素的测定（高效液相色谱法）",
    "三氯蔗糖": "GB 5009.141-2024 食品中三氯蔗糖的测定（HPLC-ELSD / 液质联用法）",
    "阿斯巴甜": "GB 5009.263-2016 食品中阿斯巴甜和阿力甜的测定",
    "纽甜": "GB 5009.247-2016 食品中纽甜的测定",
    "柠檬黄": "GB 5009.35-2023 食品中合成着色剂的测定（高效液相色谱法）",
    "日落黄": "GB 5009.35-2023 食品中合成着色剂的测定（高效液相色谱法）",
    "胭脂红": "GB 5009.35-2023 食品中合成着色剂的测定（高效液相色谱法）",
    "苋菜红": "GB 5009.35-2023 食品中合成着色剂的测定（高效液相色谱法）",
    "诱惑红": "GB 5009.35-2023 食品中合成着色剂的测定（高效液相色谱法）",
    "亮蓝": "GB 5009.35-2023 食品中合成着色剂的测定（高效液相色谱法）",
    "TBHQ": "GB 5009.32-2016 食品中9种酚类抗氧化剂的测定（气相色谱法）",
    "BHA": "GB 5009.32-2016 食品中9种酚类抗氧化剂的测定（气相色谱法）",
    "BHT": "GB 5009.32-2016 食品中9种酚类抗氧化剂的测定（气相色谱法）",
    # 微生物指标
    "菌落总数": "GB 4789.2-2022 食品微生物学检验 菌落总数测定（平板计数法）",
    "大肠菌群": "GB 4789.3-2016 食品微生物学检验 大肠菌群计数（MPN法 / 平板计数法）",
    "沙门氏菌": "GB 4789.4-2024 食品微生物学检验 沙门氏菌检验",
    "金黄色葡萄球菌": "GB 4789.10-2016 食品微生物学检验 金黄色葡萄球菌检验（平板计数法 / MPN法）",
    "单核细胞增生李斯特氏菌": "GB 4789.30-2016 食品微生物学检验 单核细胞增生李斯特氏菌检验",
    "大肠埃希氏菌": "GB 4789.36-2016 食品微生物学检验 大肠埃希氏菌O157:H7/NM检验",
    "副溶血性弧菌": "GB 4789.7-2013 食品微生物学检验 副溶血性弧菌检验",
    "蜡样芽胞杆菌": "GB 4789.14-2014 食品微生物学检验 蜡样芽胞杆菌检验",
    "霉菌": "GB 4789.15-2016 食品微生物学检验 霉菌和酵母计数",
    "酵母": "GB 4789.15-2016 食品微生物学检验 霉菌和酵母计数",
    "商业无菌": "GB 4789.26-2023 食品微生物学检验 商业无菌检验",
    # 兽药残留
    "恩诺沙星": "GB 31656.3-2021 水产品中喹诺酮类多残留测定 / 农业部公告",
    "环丙沙星": "GB 31656.3-2021 水产品中喹诺酮类多残留测定",
    "诺氟沙星": "GB 31656.3-2021 水产品中喹诺酮类多残留测定",
    "磺胺": "GB 31658.1-2021 动物性食品中磺胺类药物残留量的测定（LC-MS/MS）",
    "氯霉素": "GB 31656.2-2021 水产品中氯霉素类药物残留量的测定 / GB 31658.20-2022",
    "氟苯尼考": "GB 31656.2-2021 水产品中氯霉素类残留测定 / GB 31658.20-2022",
    "甲砜霉素": "GB 31656.2-2021 水产品中氯霉素类残留测定",
    "克伦特罗": "GB 31658.22-2022 动物性食品中β-受体激动剂残留量的测定（液质联用法）",
    "莱克多巴胺": "GB 31658.22-2022 动物性食品中β-受体激动剂残留量的测定（液质联用法）",
    "沙丁胺醇": "GB 31658.22-2022 动物性食品中β-受体激动剂残留量的测定（液质联用法）",
    "土霉素": "GB 31658.17-2021 动物性食品中四环素类药物残留量的测定",
    "金霉素": "GB 31658.17-2021 动物性食品中四环素类药物残留量的测定",
    "四环素": "GB 31658.17-2021 动物性食品中四环素类药物残留量的测定",
    "多西环素": "GB 31658.17-2021 动物性食品中四环素类药物残留量的测定",
    "甲硝唑": "GB 31658.3-2021 动物性食品中硝基咪唑类药物残留量的测定",
    "地美硝唑": "GB 31658.3-2021 动物性食品中硝基咪唑类药物残留量的测定",
    "呋喃唑酮": "GB 31656.13-2021 水产品中硝基呋喃类代谢物残留量的测定（LC-MS/MS）",
    "呋喃它酮": "GB 31656.13-2021 水产品中硝基呋喃类代谢物残留量的测定（LC-MS/MS）",
    "呋喃西林": "GB 31656.13-2021 水产品中硝基呋喃类代谢物残留量的测定（LC-MS/MS）",
    "呋喃妥因": "GB 31656.13-2021 水产品中硝基呋喃类代谢物残留量的测定（LC-MS/MS）"
}

# ============================================================
# ============================================================
# 2026年国家食品安全监督抽检实施细则 四级分类与检验方法引擎
# ============================================================
SAMPLING_RULES_FULL_DB = {}
SAMPLING_RULES_CAT_INDEX = {}
SAMPLING_RULES_METHOD_INDEX = {}

def _norm_item_key(name):
    if not name:
        return ""
    s = re.sub(r'\s+', '', str(name)).strip()
    s = s.replace('(', '（').replace(')', '）')
    return s.lower()

def _core_item_key(name):
    if not name:
        return ""
    s = re.sub(r'[\(（].*?[\)）]', '', str(name))
    s = re.sub(r'\s+', '', s).strip()
    return s.lower()

def init_sampling_rules_method_index():
    """初始化2026抽检细则四级分类库与项目检验方法索引"""
    global SAMPLING_RULES_FULL_DB, SAMPLING_RULES_CAT_INDEX, SAMPLING_RULES_METHOD_INDEX
    try:
        s_db = load_json_safe('sampling_rules_2026_db.json')
        if s_db and isinstance(s_db, dict):
            SAMPLING_RULES_FULL_DB = s_db
            for cat_key, c_data in s_db.items():
                c1 = str(c_data.get('major_category', '')).strip()
                c2 = str(c_data.get('sub_category', '')).strip()
                c3 = str(c_data.get('variety_category', '')).strip()
                c4 = str(c_data.get('fine_category', '') or c_data.get('product_name', '')).strip()
                p_name = str(c_data.get('product_name', '')).strip()
                items_list = c_data.get('items', [])
                for k_cand in {cat_key, p_name, c4, c3, c2, f"{c1}-{c4}", f"{c1}-{p_name}", f"{c2}-{c4}"}:
                    k_clean = re.sub(r'\s+', '', k_cand).strip()
                    if k_clean:
                        if k_clean not in SAMPLING_RULES_CAT_INDEX:
                            SAMPLING_RULES_CAT_INDEX[k_clean] = []
                        SAMPLING_RULES_CAT_INDEX[k_clean].extend(items_list)
                for it in items_list:
                    raw_name = it.get('item', '').strip()
                    m = it.get('method', '').strip()
                    if raw_name and m:
                        clean_name = re.sub(r'[\(（].*?[\)）]', '', raw_name).strip()
                        if clean_name and clean_name not in SAMPLING_RULES_METHOD_INDEX:
                            SAMPLING_RULES_METHOD_INDEX[clean_name] = m
                        if raw_name not in SAMPLING_RULES_METHOD_INDEX:
                            SAMPLING_RULES_METHOD_INDEX[raw_name] = m
    except Exception as e:
        print(f"Warning: Failed to build sampling rules method index: {e}")

init_sampling_rules_method_index()
def clean_method_text(m_text):
    """清洗标准中附带的方法文本，提取纯净标准编号"""
    if not m_text:
        return ""
    m = str(m_text).strip()
    m = re.sub(r'^(按照|按)\s*', '', m)
    m = re.sub(r'规定的方法测定[。.]?$', '', m)
    m = re.sub(r'测定[。.]?$', '', m)
    return m.strip()

def resolve_item_test_method(raw_method="", test_item="", food_cat="", std_name=""):
    """多层级解析并获取官方要求或推荐的检验方法标准"""
    # 第1层：数据本身自带的权威方法
    if raw_method and str(raw_method).strip() not in ("", "/", "None", "--"):
        cleaned = clean_method_text(raw_method)
        if cleaned:
            return cleaned

    clean_item = strip_highlights(test_item).strip() if test_item else ""
    if not clean_item:
        return "/"

    # 简化项目名，如 "脱氢乙酸及其钠盐(以脱氢乙酸计)" -> "脱氢乙酸"
    core_name = re.sub(r'[\(（].*?[\)）]', '', clean_item).strip()

    # 第2层：从《2026年国家食品安全抽检实施细则》精准匹配
    if SAMPLING_RULES_METHOD_INDEX:
        if clean_item in SAMPLING_RULES_METHOD_INDEX:
            return SAMPLING_RULES_METHOD_INDEX[clean_item]
        if core_name in SAMPLING_RULES_METHOD_INDEX:
            return SAMPLING_RULES_METHOD_INDEX[core_name]
        for k, v in SAMPLING_RULES_METHOD_INDEX.items():
            if core_name and (core_name in k or k in core_name):
                return v

    # 第3层：从 TEST_METHOD_DB 权威库匹配
    for k, v in TEST_METHOD_DB.items():
        if k == clean_item or (core_name and k == core_name):
            return v
    for k, v in TEST_METHOD_DB.items():
        if k in clean_item or clean_item in k:
            return v

    # 第4层：依据所属国标大类给出官方通用系列方法
    std_upper = str(std_name).upper()
    if "2763" in std_upper or "农药" in std_upper:
        return "GB 23200.121 (液质) / GB 23200.113 (气质)"
    if "31650" in std_upper or "兽药" in std_upper:
        return "GB 31658 系列 / 农业农村部公告方法"
    if "2760" in std_upper or "添加剂" in std_upper:
        return "GB 5009 系列添加剂检验方法"
    if "2761" in std_upper or "真菌毒素" in std_upper:
        return "GB 5009 系列真菌毒素检验方法"
    if "29921" in std_upper or "致病菌" in std_upper:
        return "GB 4789 系列微生物检验方法"

    return "参照 GB 5009 相应检验方法"

def get_test_method_info(test_item, food_cat="", std_name=""):
    """匹配权威官方检验方法标准"""
    return resolve_item_test_method("", test_item, food_cat, std_name)


def get_chinese_pinyin_initials(text):
    """提取汉字字符串的拼音首字母，支持大小写字母与数字提取"""
    try:
        import pypinyin
        letters = pypinyin.lazy_pinyin(str(text), style=pypinyin.Style.FIRST_LETTER)
        return ''.join(l.lower() for l in letters if l)
    except Exception:
        pass
        
    res = []
    for char in str(text):
        if '\u4e00' <= char <= '\u9fff':
            try:
                gb_bytes = char.encode('gbk')
                if len(gb_bytes) == 2:
                    code = gb_bytes[0] * 256 + gb_bytes[1]
                    if 0xB0A1 <= code <= 0xB0C4: res.append('a')
                    elif 0xB0C5 <= code <= 0xB2C0: res.append('b')
                    elif 0xB2C1 <= code <= 0xB4ED: res.append('c')
                    elif 0xB4EE <= code <= 0xB6E9: res.append('d')
                    elif 0xB6EA <= code <= 0xB7A1: res.append('e')
                    elif 0xB7A2 <= code <= 0xB8C0: res.append('f')
                    elif 0xB8C1 <= code <= 0xB9FD: res.append('g')
                    elif 0xB9FE <= code <= 0xBBF6: res.append('h')
                    elif 0xBBF7 <= code <= 0xBFA5: res.append('j')
                    elif 0xBFA6 <= code <= 0xC0AB: res.append('k')
                    elif 0xC0AC <= code <= 0xC2E7: res.append('l')
                    elif 0xC2E8 <= code <= 0xC4C2: res.append('m')
                    elif 0xC4C3 <= code <= 0xC5B5: res.append('n')
                    elif 0xC5B6 <= code <= 0xC5BD: res.append('o')
                    elif 0xC5BE <= code <= 0xC6D9: res.append('p')
                    elif 0xC6DA <= code <= 0xC8BA: res.append('q')
                    elif 0xC8BB <= code <= 0xC8F5: res.append('r')
                    elif 0xC8F6 <= code <= 0xCBF9: res.append('s')
                    elif 0xCBFA <= code <= 0xCDD9: res.append('t')
                    elif 0xCDDA <= code <= 0xCEF3: res.append('w')
                    elif 0xCEF4 <= code <= 0xD1B8: res.append('x')
                    elif 0xD1B9 <= code <= 0xD4D0: res.append('y')
                    elif 0xD4D1 <= code <= 0xD7F9: res.append('z')
                    else: res.append(char.lower())
                else:
                    res.append(char.lower())
            except Exception:
                res.append(char.lower())
        elif char.isalnum():
            res.append(char.lower())
    return ''.join(res)

# ============================================================
# GB 29921 / GB 4789.1 微生物三级采样方案 (n, c, m, M) 知识库
def universal_tree_sort_column(tree, col_name, tree_key, sort_state_dict, columns, on_reload_callback=None):
    """
    通用 Treeview 列排序函数（支持三态循环：升序 ▲ -> 降序 ▼ -> 恢复默认 ▾）
    智能识别浮点数值进行数学大小排序，判定结论超标优先置顶，文本进行拼音/字符排序
    """
    state = sort_state_dict.get(tree_key, {"col": None, "reverse": None})
    curr_col = state.get("col")
    curr_rev = state.get("reverse")
    
    if curr_col != col_name or curr_rev is None:
        new_rev = False  # 升序
    elif curr_rev is False:
        new_rev = True   # 降序
    else:
        new_rev = None   # 恢复默认
        
    sort_state_dict[tree_key] = {"col": col_name if new_rev is not None else None, "reverse": new_rev}
    
    # 更新所有表头指示图标
    for c in columns:
        base_title = c.replace(' ▲', '').replace(' ▼', '').replace(' ▾', '').replace(' (已筛选)', '')
        if c == col_name and new_rev is False:
            tree.heading(c, text=f"{base_title} ▲")
        elif c == col_name and new_rev is True:
            tree.heading(c, text=f"{base_title} ▼")
        else:
            tree.heading(c, text=f"{base_title} ▾")
            
    if new_rev is None:
        # 恢复默认排列
        if on_reload_callback:
            on_reload_callback()
        return

    items = tree.get_children('')
    if not items:
        return
        
    col_idx = list(columns).index(col_name) if col_name in columns else 0
    
    def sort_key(item_id):
        vals = tree.item(item_id, 'values')
        if col_idx >= len(vals):
            return (1, "")
        raw_val = strip_highlights(str(vals[col_idx])).strip()
        
        # 特殊处理判定结论（超标置顶优先）
        if "判定" in col_name or "结论" in col_name or "状态" in col_name:
            if "超标" in raw_val or "不合格" in raw_val or "禁用" in raw_val or "超限" in raw_val:
                return (0, 0, raw_val)
            elif "合格" in raw_val or "未检出" in raw_val:
                return (0, 1, raw_val)
            else:
                return (0, 2, raw_val)
                
        # 提取浮点数进行数学大小排序
        num = extract_first_number(raw_val)
        if num is not None:
            return (0, num)
        else:
            return (1, raw_val.lower())
            
    sorted_items = sorted(items, key=sort_key, reverse=new_rev)
    for index, item_id in enumerate(sorted_items):
        tree.move(item_id, '', index)


def extract_first_number(s):
    """从文本中提取第一个浮点数（支持正负小数、科学计数法）"""
    if not s:
        return None
    m = re.search(r'[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?', str(s))
    if m:
        try:
            return float(m.group(0))
        except ValueError:
            return None
    return None

def judge_measured_value(limit_val_str, measured_val_str):
    """
    智能比对实测值与限量标准要求
    返回: (status, result_text)
    status: 'pass' (合格), 'fail' (超标), 'unknown' (无法比对), '' (无实测值)
    """
    if not measured_val_str or str(measured_val_str).strip() in ("", "-", "--", "None"):
        return ("", "--")
        
    meas_s = str(measured_val_str).strip()
    lim_s = str(limit_val_str).strip() if limit_val_str else ""
    
    # 1. 限量要求为“不得检出”或“不得使用”
    if "不得检出" in lim_s or "不得使用" in lim_s or "0/25" in lim_s:
        meas_lower = meas_s.lower()
        if meas_lower in ["0", "0.0", "0.00", "nd", "未检出", "阴性", "-"] or meas_s.startswith("<") or meas_s.startswith("＜"):
            return ("pass", "✔ 合格 (未检出)")
        else:
            return ("fail", f"✘ 不合格 (检出: {meas_s})")
            
    # 2. 实测值为“未检出”或“<方法检出限”
    meas_lower = meas_s.lower()
    if meas_lower in ["nd", "未检出", "阴性"] or meas_s.startswith("<") or meas_s.startswith("＜"):
        return ("pass", f"✔ 合格 ({meas_s})")
        
    # 3. 数值化比对
    lim_num = extract_first_number(lim_s)
    meas_num = extract_first_number(meas_s)
    
    if lim_num is not None and meas_num is not None:
        if meas_num <= lim_num:
            return ("pass", f"✔ 合格 ({meas_num} ≤ {lim_num})")
        else:
            return ("fail", f"✘ 超标 ({meas_num} > {lim_num})")
            
    return ("unknown", f"待核对 ({meas_s})")

def split_cell_test_items(cell_text):
    """
    智能拆分单元格内的多个检测项目（支持顿号、逗号、分号、换行）：
    1. 保护括号 () 与 （） 内部内容不被拆分（例如 合成着色剂（柠檬黄、日落黄...）、土霉素/金霉素/四环素（组合含量））。
    2. 支持的分隔符：顿号 、、中文逗号 ，、英文逗号 ,、分号 ； ;、换行符 \n \r。
    3. 自动清洗前导数字序号（如 '1.'、'①'）与多余空格。
    """
    if not cell_text:
        return []
    text = str(cell_text).strip()
    if not text:
        return []

    text = text.replace('\r\n', '\n').replace('\r', '\n')

    tokens = []
    cur = []
    paren_depth = 0
    bracket_depth = 0

    for ch in text:
        if ch in ('(', '（'):
            paren_depth += 1
            cur.append(ch)
        elif ch in (')', '）'):
            if paren_depth > 0:
                paren_depth -= 1
            cur.append(ch)
        elif ch in ('[', '【'):
            bracket_depth += 1
            cur.append(ch)
        elif ch in (']', '】'):
            if bracket_depth > 0:
                bracket_depth -= 1
            cur.append(ch)
        elif paren_depth == 0 and bracket_depth == 0 and ch in ('、', '，', ',', '；', ';', '\n', '\t'):
            tok = ''.join(cur).strip()
            if tok:
                tokens.append(tok)
            cur = []
        else:
            cur.append(ch)

    if cur:
        tok = ''.join(cur).strip()
        if tok:
            tokens.append(tok)

    items = []
    for t in tokens:
        clean_t = re.sub(r'^[0-9]+[\.、\s]*', '', t).strip()
        clean_t = re.sub(r'^[（\(][0-9]+[）\)]\s*', '', clean_t).strip()
        clean_t = re.sub(r'^[①②③④⑤⑥⑦⑧⑨⑩]\s*', '', clean_t).strip()
        clean_t = re.sub(r'\s+', '', clean_t).strip()
        if clean_t and clean_t not in ('/', '-', '--', '无', 'None'):
            items.append(clean_t)
    return items

def parse_batch_input_line(line):
    """
    解析多项目批量输入的一行，支持纯项目名或带实测值的格式
    例如:
      - "铅" -> ("铅", None)
      - "铅, 0.08" -> ("铅", "0.08")
      - "脱氢乙酸 0.5" -> ("脱氢乙酸", "0.5")
      - "镉\t0.03" -> ("镉", "0.03")
      - "黄曲霉毒素B1: 15" -> ("黄曲霉毒素B1", "15")
    """
    line = line.strip()
    if not line:
        return None, None
        
    # 优先根据常见分隔符拆分：制表符、中文/英文逗号、冒号、等号、分号
    for sep in ['\t', '，', ',', '：', ':', '=', '；', ';']:
        if sep in line:
            parts = line.split(sep, 1)
            item_name = parts[0].strip()
            val_str = parts[1].strip()
            if item_name:
                return item_name, val_str if val_str else None
                
    # 空格分隔（当最后一段是数字或ND等检测值时）
    parts = line.split()
    if len(parts) >= 2:
        last = parts[-1].strip()
        if re.match(r'^[<＜>＞]?\s*\d+(\.\d+)?$|^nd$|^未检出$|^阴性$|^阳性$', last, re.IGNORECASE):
            item_name = " ".join(parts[:-1]).strip()
            return item_name, last
            
    return line, None


# ============================================================
# GB 2760-2024 新旧国标对比数据库
class ColumnVisibilityDialog(tb.Toplevel):
    def __init__(self, parent, tree, tree_key, all_columns, default_widths, on_change_callback):
        super().__init__(parent)
        self.title("自定义表格列显示/隐藏")
        self.geometry("320x400")
        self.minsize(280, 320)
        self.transient(parent)
        self.grab_set()
        self.tree = tree
        self.tree_key = tree_key
        self.all_columns = all_columns
        self.default_widths = default_widths
        self.on_change_callback = on_change_callback
        
        try:
            self.update_idletasks()
            x = parent.winfo_rootx() + (parent.winfo_width() - 320) // 2
            y = parent.winfo_rooty() + (parent.winfo_height() - 400) // 2
            self.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

        container = tb.Frame(self, padding=16)
        container.pack(fill=tk.BOTH, expand=True)

        tb.Label(container, text="请勾选需要显示的表格列：", font=("Microsoft YaHei", 10, "bold")).pack(anchor="w", pady=(0, 10))
        
        self.vars = {}
        for col in self.all_columns:
            current_w = self.tree.column(col, 'width')
            is_visible = current_w > 0
            var = tk.BooleanVar(value=is_visible)
            self.vars[col] = var
            cb = tb.Checkbutton(container, text=f"  {col}", variable=var, bootstyle="round-toggle")
            cb.pack(anchor="w", pady=4)
            
        btn_box = tb.Frame(container)
        btn_box.pack(fill=tk.X, side=tk.BOTTOM, pady=(15, 0))
        tb.Button(btn_box, text="全选", command=self.select_all, bootstyle="secondary-outline").pack(side=tk.LEFT)
        tb.Button(btn_box, text="应用保存", command=self.apply_changes, bootstyle=SUCCESS).pack(side=tk.RIGHT)
        tb.Button(btn_box, text="取消", command=self.destroy, bootstyle=SECONDARY).pack(side=tk.RIGHT, padx=5)

    def select_all(self):
        for v in self.vars.values():
            v.set(True)

    def apply_changes(self):
        visible_cols = {col: self.vars[col].get() for col in self.all_columns}
        self.on_change_callback(self.tree, self.tree_key, visible_cols, self.default_widths)
        self.destroy()


class ColumnFilterDialog(tb.Toplevel):
    def __init__(self, parent, col_name, unique_values, current_filters, callback):
        super().__init__(parent)
        self.title(f"筛选: {col_name}")
        self.geometry("340x480")
        self.minsize(300, 380)
        self.transient(parent)
        self.grab_set()
        self.callback = callback
        self.unique_values = [str(v) for v in unique_values]
        
        # 维护一个独立的已选集合
        if current_filters is None:
            self.selected_set = set(self.unique_values)
        else:
            self.selected_set = set(str(v) for v in current_filters)
            
        # 窗口居中
        try:
            self.update_idletasks()
            x = parent.winfo_rootx() + (parent.winfo_width() - 340) // 2
            y = parent.winfo_rooty() + (parent.winfo_height() - 480) // 2
            self.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

        container = tb.Frame(self, padding=12)
        container.pack(fill=tk.BOTH, expand=True)

        # 搜索框
        tb.Label(container, text="快速查找选项:", font=("Microsoft YaHei", 9, "bold")).pack(anchor="w", pady=(0, 4))
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.update_list)
        tb.Entry(container, textvariable=self.search_var, font=("Microsoft YaHei", 10)).pack(fill=tk.X, pady=(0, 8))
        
        # 操作按钮区
        btn_frame = tb.Frame(container)
        btn_frame.pack(fill=tk.X, pady=(0, 6))
        tb.Button(btn_frame, text="全选", command=self.select_all_displayed, bootstyle="success-outline").pack(side=tk.LEFT, padx=(0, 4))
        tb.Button(btn_frame, text="全不选", command=self.unselect_all_displayed, bootstyle="secondary-outline").pack(side=tk.LEFT, padx=4)
        tb.Button(btn_frame, text="反选", command=self.invert_select_displayed, bootstyle="warning-outline").pack(side=tk.LEFT, padx=4)
        tb.Button(btn_frame, text="清除筛选", command=self.clear_filter, bootstyle="danger-outline").pack(side=tk.RIGHT)
        
        # 数量状态标签
        self.count_label = tb.Label(container, text="", font=("Microsoft YaHei", 8), bootstyle=SECONDARY)
        self.count_label.pack(anchor="w", pady=(0, 4))

        # 带复选框前缀的列表框
        frame = tb.Frame(container)
        frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        
        self.listbox = tk.Listbox(frame, selectmode=tk.SINGLE, font=("Microsoft YaHei", 9), activestyle='none')
        sb = ttk.Scrollbar(frame, command=self.listbox.yview)
        self.listbox.config(yscrollcommand=sb.set)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 单击/空格切换勾选
        self.listbox.bind('<Button-1>', self.on_item_click)
        self.listbox.bind('<space>', self.on_item_space)
        
        # 底部确定/取消
        bottom_box = tb.Frame(container)
        bottom_box.pack(fill=tk.X)
        tb.Button(bottom_box, text="确定", command=self.on_ok, bootstyle=PRIMARY).pack(side=tk.RIGHT, padx=(5, 0))
        tb.Button(bottom_box, text="取消", command=self.destroy, bootstyle=SECONDARY).pack(side=tk.RIGHT)
        
        self.update_list()
        
    def update_list(self, *args):
        search_kw = self.search_var.get().strip().lower()
        self.listbox.delete(0, tk.END)
        self.display_items = []
        for val in self.unique_values:
            if not search_kw or search_kw in val.lower():
                self.display_items.append(val)
                checked = val in self.selected_set
                prefix = "☑  " if checked else "☐  "
                self.listbox.insert(tk.END, prefix + val)
        self.update_count_label()
        
    def update_count_label(self):
        sel_count = len(self.selected_set)
        tot_count = len(self.unique_values)
        self.count_label.config(text=f"已选 {sel_count} / 共 {tot_count} 项（点击任意行切换勾选）")

    def on_item_click(self, event):
        idx = self.listbox.nearest(event.y)
        if 0 <= idx < len(self.display_items):
            val = self.display_items[idx]
            if val in self.selected_set:
                self.selected_set.remove(val)
            else:
                self.selected_set.add(val)
            checked = val in self.selected_set
            prefix = "☑  " if checked else "☐  "
            self.listbox.delete(idx)
            self.listbox.insert(idx, prefix + val)
            self.update_count_label()

    def on_item_space(self, event):
        sel = self.listbox.curselection()
        if sel:
            idx = sel[0]
            val = self.display_items[idx]
            if val in self.selected_set:
                self.selected_set.remove(val)
            else:
                self.selected_set.add(val)
            checked = val in self.selected_set
            prefix = "☑  " if checked else "☐  "
            self.listbox.delete(idx)
            self.listbox.insert(idx, prefix + val)
            self.listbox.selection_set(idx)
            self.update_count_label()

    def select_all_displayed(self):
        for val in self.display_items:
            self.selected_set.add(val)
        self.update_list()

    def unselect_all_displayed(self):
        for val in self.display_items:
            self.selected_set.discard(val)
        self.update_list()

    def invert_select_displayed(self):
        for val in self.display_items:
            if val in self.selected_set:
                self.selected_set.remove(val)
            else:
                self.selected_set.add(val)
        self.update_list()

    def clear_filter(self):
        self.callback(None)
        self.destroy()

    def on_ok(self):
        if len(self.selected_set) == len(self.unique_values) or len(self.selected_set) == 0:
            self.callback(None)
        else:
            self.callback(set(self.selected_set))
        self.destroy()



class TreeviewCellTooltip:
    """高级 Treeview 单元格智能悬浮气泡，支持完整展开、鼠标划词拉选与一键复制"""
    def __init__(self, tree):
        self.tree = tree
        self.tip_window = None
        self.after_id = None
        self.last_cell = None  # (item_id, col_id)
        
        self.tree.bind("<Motion>", self.on_motion, add="+")
        self.tree.bind("<Leave>", self.on_leave, add="+")

    def on_motion(self, event):
        item_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        
        if not item_id or not col_id:
            self.cancel_timer()
            self.hide_tooltip()
            self.last_cell = None
            return

        current_cell = (item_id, col_id)
        if current_cell != self.last_cell:
            self.cancel_timer()
            self.hide_tooltip()
            self.last_cell = current_cell
            # 悬停 280ms 自动触发浮动气泡
            self.after_id = self.tree.after(280, lambda: self.show_tooltip(event.x_root, event.y_root, item_id, col_id))

    def cancel_timer(self):
        if self.after_id:
            try:
                self.tree.after_cancel(self.after_id)
            except Exception:
                pass
            self.after_id = None

    def on_leave(self, event=None):
        self.cancel_timer()
        if self.tip_window:
            self.tree.after(250, self.check_leave)

    def check_leave(self):
        if not self.tip_window:
            return
        try:
            mx = self.tree.winfo_pointerx()
            my = self.tree.winfo_pointery()
            wx = self.tip_window.winfo_rootx()
            wy = self.tip_window.winfo_rooty()
            ww = self.tip_window.winfo_width()
            wh = self.tip_window.winfo_height()
            # 鼠标如果移动进气泡窗口内，保持显示以便用户拉选复制
            if wx - 2 <= mx <= wx + ww + 2 and wy - 2 <= my <= wy + wh + 2:
                return
        except Exception:
            pass
        self.hide_tooltip()

    def hide_tooltip(self):
        if self.tip_window:
            try:
                self.tip_window.destroy()
            except Exception:
                pass
            self.tip_window = None

    def show_tooltip(self, x, y, item_id, col_id):
        self.hide_tooltip()
        try:
            col_idx = int(col_id.replace("#", "")) - 1
            columns = self.tree["columns"]
            if col_idx < 0 or col_idx >= len(columns):
                return
            col_name = columns[col_idx]
            vals = self.tree.item(item_id, "values")
            if col_idx >= len(vals):
                return
            text = strip_highlights(str(vals[col_idx])).strip()
            if not text or text == "/":
                return

            # 当文字长度大于10，或者属于检验方法、备注、分类等较长列时弹出
            if len(text) < 12 and "方法" not in col_name and "备注" not in col_name:
                return

            self.tip_window = tw = tk.Toplevel(self.tree)
            tw.wm_overrideredirect(True)
            tw.attributes("-topmost", True)
            
            outer_frame = tk.Frame(tw, bg="#0F172A", bd=1)
            outer_frame.pack(fill=tk.BOTH, expand=True)
            
            inner_frame = tk.Frame(outer_frame, bg="#FFFFFF", padx=8, pady=6)
            inner_frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

            # 气泡标题行
            hdr = tk.Frame(inner_frame, bg="#FFFFFF")
            hdr.pack(fill=tk.X, pady=(0, 4))
            tk.Label(hdr, text=f"🔍 {col_name} (完整展开 / 鼠标可划选复制)", font=("Microsoft YaHei", 8, "bold"), fg="#0284C7", bg="#FFFFFF").pack(side=tk.LEFT)
            
            def copy_text():
                self.tree.clipboard_clear()
                self.tree.clipboard_append(text)
                self.tree.update()
                btn_copy.config(text="✔ 已复制")

            btn_copy = tk.Button(hdr, text="📋 复制全部", font=("Microsoft YaHei", 8), command=copy_text, relief=tk.FLAT, bg="#E2E8F0", cursor="hand2")
            btn_copy.pack(side=tk.RIGHT)

            # 多行可选中文本框
            calc_h = min(6, max(2, len(text) // 25 + 1))
            calc_w = min(55, max(28, len(text) + 2))
            txt_box = tk.Text(inner_frame, wrap=tk.WORD, height=calc_h, width=calc_w,
                              font=("Consolas", 9), bg="#F8FAFC", fg="#0F172A", bd=1, relief=tk.SOLID, padx=4, pady=4)
            txt_box.insert("1.0", text)
            txt_box.pack(fill=tk.BOTH, expand=True)
            # 允许拉选光标选中复制，禁止修改
            txt_box.bind("<Key>", lambda e: "break" if (e.state != 4 and e.keysym not in ('c', 'C', 'a', 'A')) else None)

            tw.bind("<Leave>", lambda e: self.on_leave())

            # 避开屏幕边界防溢出
            screen_w = self.tree.winfo_screenwidth()
            screen_h = self.tree.winfo_screenheight()
            pos_x = min(x + 10, screen_w - 450)
            pos_y = min(y + 12, screen_h - 180)
            tw.geometry(f"+{pos_x}+{pos_y}")
        except Exception as e:
            pass


class ItemDetailDialog(tb.Toplevel):
    """双击行时弹出的标准项目详情卡片"""
    def __init__(self, parent, std_name, food_cat, test_item, limit_val, unit="", notes="", measured_val=None, judge_res=None, method_str=""):
        super().__init__(parent)
        self.title("食品安全标准限量详情")
        self.geometry("560x540")
        self.minsize(480, 440)
        self.transient(parent)
        self.grab_set()
        
        # 窗口居中
        try:
            self.update_idletasks()
            x = parent.winfo_rootx() + (parent.winfo_width() - 560) // 2
            y = parent.winfo_rooty() + (parent.winfo_height() - 540) // 2
            self.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass
            
        container = tb.Frame(self, padding=16)
        container.pack(fill=tk.BOTH, expand=True)
        
        # 顶部标准名称标签
        hdr_frame = tb.Frame(container)
        hdr_frame.pack(fill=tk.X, pady=(0, 10))
        tb.Label(hdr_frame, text=f"📌 标准依据：{std_name}", font=("Microsoft YaHei", 12, "bold"), bootstyle=PRIMARY).pack(side=tk.LEFT)
        
        # 信息卡片区
        info_frame = tb.LabelFrame(container, text=" 指标详情与检验依据 ", padding=10)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 行1: 检测项目
        r1 = tb.Frame(info_frame)
        r1.pack(fill=tk.X, pady=3)
        tb.Label(r1, text="检测项目：", font=("Microsoft YaHei", 10, "bold"), width=12, anchor="w").pack(side=tk.LEFT)
        tb.Label(r1, text=test_item, font=("Microsoft YaHei", 10, "bold"), bootstyle=INFO).pack(side=tk.LEFT)
        
        # 行2: 限量值
        r2 = tb.Frame(info_frame)
        r2.pack(fill=tk.X, pady=3)
        tb.Label(r2, text="限量要求：", font=("Microsoft YaHei", 10, "bold"), width=12, anchor="w").pack(side=tk.LEFT)
        limit_str = f"{limit_val} {unit}".strip() if limit_val else "详见备注"
        tb.Label(r2, text=limit_str, font=("Microsoft YaHei", 11, "bold"), bootstyle=DANGER).pack(side=tk.LEFT)
        
        # 可选行: 实测值与判定结论 (若有)
        has_measured = measured_val and str(measured_val).strip() not in ("", "--", "None")
        if has_measured:
            r_meas = tb.Frame(info_frame)
            r_meas.pack(fill=tk.X, pady=3)
            tb.Label(r_meas, text="实测与判定：", font=("Microsoft YaHei", 10, "bold"), width=12, anchor="w").pack(side=tk.LEFT)
            j_style = DANGER if ("✘" in str(judge_res) or "不合格" in str(judge_res) or "超标" in str(judge_res)) else SUCCESS
            tb.Label(r_meas, text=f"实测值: {measured_val}  |  判定结论: {judge_res}", font=("Microsoft YaHei", 10, "bold"), bootstyle=j_style).pack(side=tk.LEFT)
        
        # 行3: 食品类别与完整层级
        r3 = tb.Frame(info_frame)
        r3.pack(fill=tk.X, pady=3)
        tb.Label(r3, text="食品类别：", font=("Microsoft YaHei", 10, "bold"), width=12, anchor="w").pack(side=tk.LEFT)
        clean_cat = strip_highlights(food_cat)
        tb.Label(r3, text=clean_cat, font=("Microsoft YaHei", 9), wraplength=380, justify=tk.LEFT).pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 行4: 官方要求/推荐检验方法标准 (支持一键复制方法)
        r4 = tb.Frame(info_frame)
        r4.pack(fill=tk.X, pady=3)
        tb.Label(r4, text="要求检验方法：", font=("Microsoft YaHei", 10, "bold"), width=12, anchor="w").pack(side=tk.LEFT)
        method_info = method_str if method_str and method_str != "/" else get_test_method_info(test_item, food_cat, std_name)
        tb.Label(r4, text=method_info, font=("Microsoft YaHei", 9, "bold"), wraplength=310, justify=tk.LEFT, bootstyle=SUCCESS).pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        def copy_method():
            self.clipboard_clear()
            self.clipboard_append(method_info)
            self.update()
            messagebox.showinfo("复制成功", f"已将推荐检验方法复制到剪贴板：\n\n{method_info}")
            
        tb.Button(r4, text="📋 复制方法", command=copy_method, bootstyle="info-outline").pack(side=tk.RIGHT, padx=(4,0))
        
        # 备注说明文本框
        notes_frame = tb.LabelFrame(container, text=" 备注说明与判定注释 ", padding=8)
        notes_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        clean_notes = strip_highlights(notes).strip() if notes else "（该项目无特殊备注说明）"
        notes_text = tk.Text(notes_frame, font=("Microsoft YaHei", 9), wrap=tk.WORD, height=3, relief=tk.FLAT)
        notes_text.insert("1.0", clean_notes)
        notes_text.config(state="disabled")
        notes_scroll = ttk.Scrollbar(notes_frame, orient=tk.VERTICAL, command=notes_text.yview)
        notes_text.config(yscrollcommand=notes_scroll.set)
        notes_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        notes_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 底部按钮区 (支持一键复制整卡摘要)
        btn_box = tb.Frame(container)
        btn_box.pack(fill=tk.X)
        
        def copy_card_summary():
            meas_line = f"【实测与判定】实测: {measured_val} | 结论: {judge_res}\n" if has_measured else ""
            card_text = (
                f"【标准依据】{std_name}\n"
                f"【检测项目】{test_item}\n"
                f"【限量要求】{limit_str}\n"
                f"{meas_line}"
                f"【食品类别】{clean_cat}\n"
                f"【推荐检验方法】{method_info}\n"
                f"【备注说明】{clean_notes}"
            )
            self.clipboard_clear()
            self.clipboard_append(card_text)
            self.update()
            messagebox.showinfo("复制成功", "已将该指标完整详情卡片复制到剪贴板！可直接粘贴至检验报告或原始记录中。")
            
        tb.Button(btn_box, text="📋 复制完整详情卡片", command=copy_card_summary, bootstyle="info-outline").pack(side=tk.LEFT)
        tb.Button(btn_box, text="关闭窗口 (ESC)", command=self.destroy, bootstyle=SECONDARY).pack(side=tk.RIGHT)
        self.bind("<Escape>", lambda e: self.destroy())


# ============================================================
# V1.2 全局智能别名库、检出限(LOD)数据库与倒排索引加速引擎
# ============================================================
FOOD_ALIAS_MAP = {
    "大米": ["大米", "糙米", "稻谷", "谷物及其制品", "谷物"],
    "小米": ["谷物及其制品", "谷物", "谷物碾磨加工品"],
    "小麦粉": ["小麦粉", "面粉", "通用小麦粉", "专用小麦粉", "谷物碾磨加工品", "谷物及其制品"],
    "面粉": ["小麦粉", "面粉", "谷物碾磨加工品", "谷物及其制品"],
    "挂面": ["挂面", "花色挂面", "手工面", "谷物制品", "带馅(料)面米制品", "谷物碾磨加工品", "谷物及其制品", "小麦粉"],
    "生湿面制品": ["生湿面制品", "谷物制品", "发酵面制品", "谷物碾磨加工品", "谷物及其制品", "小麦粉"],
    "发酵面制品": ["发酵面制品", "谷物制品", "谷物碾磨加工品", "谷物及其制品", "小麦粉"],
    "米粉制品": ["米粉制品", "河粉", "米粉", "大米粉", "谷物制品", "谷物碾磨加工品", "谷物及其制品", "大米"],
    "大米粉": ["大米粉", "米粉", "大米", "谷物碾磨加工品", "谷物及其制品"],
    "玉米粉": ["玉米粉", "玉米片", "玉米渣", "玉米", "谷物碾磨加工品", "谷物及其制品"],
    "谷物加工品": ["谷物加工品", "其他谷物加工品", "谷物及其制品", "谷物碾磨加工品"],
    "其他谷物碾磨加工品": ["谷物碾磨加工品", "谷物及其制品", "小麦粉"],
    "粉丝粉条": ["粉丝", "粉条", "淀粉制品", "淀粉及淀粉制品"],
    "方便面": ["方便面", "方便米面制品", "方便食品", "谷物及其制品"],
    "调味面制品": ["调味面制品", "方便食品", "糕点", "谷物及其制品"],
    "糕点": ["糕点", "烘烤类糕点", "月饼", "焙烤食品", "面包"],
    "面包": ["面包", "糕点", "焙烤食品"],
    "熟肉制品": ["熟肉制品", "酱卤肉", "熏烧烤肉", "熟肉"],
    "酱卤肉": ["熟肉制品", "酱卤肉", "熟肉"],
    "猪肉": ["肉", "猪肉", "畜肉", "生鲜肉", "肌肉"],
    "牛肉": ["肉", "牛肉", "畜肉", "生鲜肉", "肌肉"],
    "羊肉": ["肉", "羊肉", "畜肉", "生鲜肉", "肌肉"],
    "鸡肉": ["肉", "鸡肉", "禽肉", "生鲜肉", "肌肉"],
    "鸭肉": ["肉", "鸭肉", "禽肉", "生鲜肉", "肌肉"],
    "鲜牛奶": ["生乳", "生鲜乳", "纯牛奶", "巴氏杀菌乳", "灭菌乳", "液体乳", "乳及乳制品"],
    "牛奶": ["生乳", "生鲜乳", "纯牛奶", "巴氏杀菌乳", "灭菌乳", "液体乳", "乳及乳制品"],
    "纯牛奶": ["纯牛奶", "生乳", "灭菌乳", "液体乳", "乳及乳制品"],
    "生乳": ["生乳", "生鲜牛乳", "液体乳", "乳及乳制品"],
    "花生油": ["花生油", "植物油脂", "食用植物油"],
    "菜籽油": ["菜籽油", "植物油脂", "食用植物油"],
    "大豆油": ["大豆油", "植物油脂", "食用植物油"],
    "食用植物油": ["食用植物油", "植物油脂"],
    "韭菜": ["韭菜", "鳞茎类蔬菜", "蔬菜"],
    "普通白菜": ["普通白菜", "小油菜", "上海青", "叶菜类蔬菜", "芸薹类蔬菜", "蔬菜"],
    "小白菜": ["普通白菜", "小油菜", "上海青", "叶菜类蔬菜", "芸薹类蔬菜", "蔬菜"],
    "上海青": ["普通白菜", "小油菜", "上海青", "叶菜类蔬菜", "芸薹类蔬菜", "蔬菜"],
    "芹菜": ["芹菜", "叶菜类蔬菜", "蔬菜"],
    "豇豆": ["豇豆", "豆角", "豆类蔬菜", "蔬菜"],
    "苹果": ["苹果", "仁果类水果", "新鲜水果", "水果"],
    "草莓": ["草莓", "浆果和其他小型水果", "新鲜水果", "水果"],
    "葡萄": ["葡萄", "浆果和其他小型水果", "新鲜水果", "水果"],
    "茶叶": ["茶叶", "绿茶", "红茶", "乌龙茶"],
    "淡水鱼": ["鱼类", "淡水鱼", "水产动物及其制品", "水产品"],
    "草鱼": ["鱼类", "淡水鱼", "草鱼", "水产动物及其制品", "水产品"],
    "罗非鱼": ["鱼类", "淡水鱼", "罗非鱼", "水产动物及其制品", "水产品"],
    "鲈鱼": ["鱼类", "淡水鱼", "海水鱼", "鲈鱼", "水产品"],
    "鸡蛋": ["蛋制品", "鲜蛋", "鸡蛋", "禽蛋"],
    "鲜蛋": ["蛋制品", "鲜蛋", "鸡蛋", "禽蛋"],
}

ITEM_ALIAS_MAP = {
    "铅": "铅",
    "pb": "铅",
    "铅(以pb计)": "铅",
    "铅（以 pb 计）": "铅",
    "镉": "镉",
    "cd": "镉",
    "镉(以cd计)": "镉",
    "汞": "总汞",
    "总汞": "总汞",
    "甲基汞": "甲基汞",
    "砷": "总砷",
    "总砷": "总砷",
    "无机砷": "无机砷",
    "脱氢乙酸": "脱氢乙酸",
    "脱氢乙酸及其钠盐": "脱氢乙酸",
    "山梨酸": "山梨酸",
    "山梨酸及其钾盐": "山梨酸",
    "苯甲酸": "苯甲酸",
    "苯甲酸及其钠盐": "苯甲酸",
    "糖精钠": "糖精钠",
    "甜蜜素": "甜蜜素",
    "环己基氨基磺酸钠": "甜蜜素",
    "安赛蜜": "乙酰磺胺酸钾",
    "乙酰磺胺酸钾": "乙酰磺胺酸钾",
    "二氧化硫": "二氧化硫",
    "二氧化硫残留量": "二氧化硫",
    "亚硝酸盐": "亚硝酸盐",
    "黄曲霉毒素b1": "黄曲霉毒素B1",
    "黄曲霉毒素m1": "黄曲霉毒素M1",
    "呕吐毒素": "脱氧雪腐镰刀菌烯醇",
    "don": "脱氧雪腐镰刀菌烯醇",
    "脱氧雪腐镰刀菌烯醇": "脱氧雪腐镰刀菌烯醇",
    "沙门氏菌": "沙门氏菌",
    "金黄色葡萄球菌": "金黄色葡萄球菌",
    "单核细胞增生李斯特氏菌": "单核细胞增生李斯特氏菌",
    "单增李斯特菌": "单核细胞增生李斯特氏菌",
    "副溶血性弧菌": "副溶血性弧菌",
    "大肠埃希氏菌o157": "大肠埃希氏菌O157:H7/NM",
    "克伦特罗": "克伦特罗",
    "莱克多巴胺": "莱克多巴胺",
    "沙丁胺醇": "沙丁胺醇",
    "氯霉素": "氯霉素",
    "孔雀石绿": "孔雀石绿",
    "恩诺沙星": "恩诺沙星",
    "氟苯尼考": "氟苯尼考",
    "腐霉利": "腐霉利",
    "啶虫脒": "啶虫脒",
    "毒死蜱": "毒死蜱",
    "吡虫啉": "吡虫啉",
    "噻虫胺": "噻虫胺",
    "克百威": "克百威",
    "甲拌磷": "甲拌磷",
    "倍硫磷": "倍硫磷",
}

LOD_DATABASE = {
    "铅": "0.005 mg/kg (ICP-MS) / 0.02 mg/kg (石墨炉)",
    "镉": "0.002 mg/kg (ICP-MS) / 0.003 mg/kg (石墨炉)",
    "总汞": "0.003 mg/kg (原子荧光法)",
    "甲基汞": "0.005 mg/kg (液相-原子荧光)",
    "总砷": "0.01 mg/kg (原子荧光法)",
    "无机砷": "0.003 mg/kg (液相-原子荧光)",
    "铬": "0.01 mg/kg (石墨炉/ICP-MS)",
    "铝": "1.0 mg/kg (分光法) / 0.1 mg/kg (ICP-MS)",
    "锡": "2.0 mg/kg (原子荧光)",
    "镍": "0.01 mg/kg (石墨炉)",
    "铜": "0.5 mg/kg (火焰原子吸收)",
    "锌": "0.5 mg/kg (火焰原子吸收)",
    "铁": "0.5 mg/kg (火焰原子吸收)",
    "苯并[a]芘": "0.2 μg/kg (GC-MS / HPLC)",
    "N-二甲基亚硝胺": "0.3 μg/kg (GC-MS/MS)",
    "多氯联苯": "0.5 μg/kg (GC-MS)",
    "丙烯酰胺": "5.0 μg/kg (LC-MS/MS)",
    "脱氢乙酸": "0.05 mg/kg (液质) / 2.0 mg/kg (气相/液相)",
    "苯甲酸": "1.0 mg/kg (高效液相色谱法)",
    "山梨酸": "1.0 mg/kg (高效液相色谱法)",
    "糖精钠": "1.0 mg/kg (高效液相色谱法)",
    "甜蜜素": "2.0 mg/kg (气相/液质色谱法)",
    "安赛蜜": "2.0 mg/kg (高效液相色谱法)",
    "乙酰磺胺酸钾": "2.0 mg/kg (高效液相色谱法)",
    "二氧化硫": "1.0 mg/kg (离子色谱) / 3.0 mg/kg (滴定法)",
    "亚硝酸盐": "1.0 mg/kg (离子色谱/分光光度法)",
    "硝酸盐": "2.0 mg/kg (离子色谱法)",
    "纳他霉素": "0.5 mg/kg (高效液相色谱法)",
    "三氯蔗糖": "2.0 mg/kg (HPLC-ELSD / LC-MS)",
    "阿斯巴甜": "2.0 mg/kg (高效液相色谱法)",
    "纽甜": "1.0 mg/kg (高效液相色谱法)",
    "黄曲霉毒素B1": "0.1 μg/kg (LC-MS/MS)",
    "黄曲霉毒素M1": "0.02 μg/kg (LC-MS/MS)",
    "脱氧雪腐镰刀菌烯醇": "20 μg/kg (LC-MS/MS)",
    "展青霉素": "5.0 μg/kg (LC-MS/MS / HPLC)",
    "赭曲霉毒素A": "0.5 μg/kg (LC-MS/MS)",
    "玉米赤霉烯酮": "2.0 μg/kg (LC-MS/MS)",
    "伏马毒素": "10 μg/kg (LC-MS/MS)",
    "T-2毒素": "5.0 μg/kg (LC-MS/MS)",
    "克伦特罗": "0.2 μg/kg (LC-MS/MS)",
    "莱克多巴胺": "0.2 μg/kg (LC-MS/MS)",
    "沙丁胺醇": "0.2 μg/kg (LC-MS/MS)",
    "氯霉素": "0.1 μg/kg (LC-MS/MS)",
    "孔雀石绿": "0.5 μg/kg (LC-MS/MS)",
    "恩诺沙星": "1.0 μg/kg (LC-MS/MS)",
    "环丙沙星": "1.0 μg/kg (LC-MS/MS)",
    "氟苯尼考": "0.5 μg/kg (LC-MS/MS)",
    "沙门氏菌": "0/25g (定性检出限)",
    "单核细胞增生李斯特氏菌": "0/25g (定性检出限)",
    "大肠埃希氏菌O157:H7/NM": "0/25g (定性检出限)",
    "金黄色葡萄球菌": "10 CFU/g (平板计数)",
    "副溶血性弧菌": "3 MPN/g (MPN法)",
}

def resolve_item_lod(item_name, method_str="", limit_val="", unit=""):
    """权威检出限 (LOD) 解析引擎"""
    if not item_name:
        return "/"
    clean_k = re.sub(r'[\(（].*?[\)）]', '', str(item_name)).strip()
    if clean_k in LOD_DATABASE:
        return LOD_DATABASE[clean_k]
    for k, v in LOD_DATABASE.items():
        if k in clean_k or clean_k in k:
            return v
    if "23200" in str(method_str) or "农药" in str(item_name):
        return "0.01 mg/kg (GC-MS/MS / LC-MS/MS)"
    if "31658" in str(method_str) or "31650" in str(method_str):
        return "1.0 μg/kg (LC-MS/MS)"
    u = str(unit).lower()
    if "μg" in u or "ug" in u:
        return "0.1 μg/kg (依据推荐检验方法)"
    if "mg" in u:
        return "0.01 mg/kg (依据推荐检验方法)"
    if "g" in u:
        return "0.001 g/kg (1.0 mg/kg)"
    return "0.01 mg/kg (依据推荐检验方法)"

def calc_match_rank(item, search_text):
    """
    计算搜索条目精准匹配优先级 (越小越排在最上面):
      0: 项目名称完全精准匹配
      1: 食品分类完全精准匹配
      2: 括号子项或分词完全命中
      3: 项目名称以搜索词开头
      4: 食品分类以搜索词开头
      5: 项目名称包含搜索词
      6: 食品分类包含搜索词
      7: 检验方法或备注包含搜索词
      99: 未命中
    """
    if not search_text:
        return (99, 999, 999)
    s_clean = re.sub(r'[\(（].*?[\)）]', '', search_text).strip().lower()
    s_raw = search_text.strip().lower()
    t_item = str(item.get('test_item', '')).strip()
    f_cat = str(item.get('food_category', '')).strip()
    m_str = str(item.get('method', '')).strip()
    n_str = str(item.get('notes', '')).strip()
    t_clean = re.sub(r'[\(（].*?[\)）]', '', t_item).strip().lower()
    f_clean = re.sub(r'[\(（].*?[\)）]', '', f_cat).strip().lower()
    t_raw = t_item.lower()
    f_raw = f_cat.lower()
    
    if s_clean and (s_clean == t_clean or s_raw == t_raw):
        return (0, len(t_item), len(f_cat))
    if s_clean and (s_clean == f_clean or s_raw == f_raw):
        return (1, len(t_item), len(f_cat))
        
    subs_t = re.split(r'[,，、/()（）\s]+', t_raw)
    subs_f = re.split(r'[,，、/()（）\s]+', f_raw)
    if s_clean in subs_t or s_raw in subs_t:
        return (2, len(t_item), len(f_cat))
    if s_clean in subs_f or s_raw in subs_f:
        return (2, len(t_item), len(f_cat))
        
    if t_raw.startswith(s_raw) or (s_clean and t_clean.startswith(s_clean)):
        return (3, len(t_item), len(f_cat))
    if f_raw.startswith(s_raw) or (s_clean and f_clean.startswith(s_clean)):
        return (4, len(t_item), len(f_cat))
        
    if s_raw in t_raw or (s_clean and s_clean in t_raw):
        return (5, len(t_item), len(f_cat))
    if s_raw in f_raw or (s_clean and s_clean in f_raw):
        return (6, len(t_item), len(f_cat))
    if s_raw in m_str.lower() or s_raw in n_str.lower():
        return (7, len(t_item), len(f_cat))
    return (99, len(t_item), len(f_cat))

def build_global_item_index(standards):
    """构建全局跨标准倒排索引，加速批量自动匹配引擎"""
    index = {}
    for std_code, s_info in standards.items():
        items = s_info.get('items', [])
        for it in items:
            raw_t = it.get('test_item', '')
            c_t = re.sub(r'[\(（].*?[\)）]', '', raw_t).strip().lower()
            if c_t:
                if c_t not in index:
                    index[c_t] = []
                index[c_t].append((std_code, it))
            std_alias = ITEM_ALIAS_MAP.get(c_t)
            if std_alias and std_alias.lower() != c_t:
                ak = std_alias.lower()
                if ak not in index:
                    index[ak] = []
                index[ak].append((std_code, it))
    return index

def match_2026_sampling_rule(cat_hierarchy, item_query):
    """
    在《2026年国家食品安全监督抽检实施细则》四级分类库中匹配 (依据法律法规或标准, 检测方法, 匹配细类)
    cat_hierarchy: [四级细类/样品名, 三级品种, 二级亚类, 一级大类]
    """
    if not SAMPLING_RULES_CAT_INDEX or not item_query:
        return None
    raw_i = str(item_query).strip()
    norm_i = _norm_item_key(raw_i)
    core_i = _core_item_key(raw_i)
    alias_i = ITEM_ALIAS_MAP.get(core_i, core_i).lower()

    # 构建候选分类键（从最精细的四级分类到一级大类）
    cat_keys = []
    valid_cats = [re.sub(r'\s+', '', str(c)).strip() for c in cat_hierarchy if c and str(c).strip() not in ('', '/', 'None')]
    if len(valid_cats) >= 2:
        cat_keys.append(f"{valid_cats[-1]}-{valid_cats[0]}")
    for c in valid_cats:
        cat_keys.append(c)
        c_no_paren = re.sub(r'[\(（].*?[\)）]', '', c).strip()
        if c_no_paren and c_no_paren not in cat_keys:
            cat_keys.append(c_no_paren)
        if c_no_paren in FOOD_ALIAS_MAP:
            for al in FOOD_ALIAS_MAP[c_no_paren]:
                if al not in cat_keys:
                    cat_keys.append(al)

    # 1. 先在匹配的四级/三级/二级分类桶中精准查找该项目
    for ck in cat_keys:
        bucket = SAMPLING_RULES_CAT_INDEX.get(ck)
        if not bucket:
            # 尝试子串匹配分类键
            for k_idx, b_list in SAMPLING_RULES_CAT_INDEX.items():
                if ck and (ck == k_idx or ck in k_idx or k_idx in ck):
                    bucket = b_list
                    break
        if bucket:
            # 精确匹配项目名
            for it in bucket:
                it_norm = _norm_item_key(it.get('item', ''))
                it_core = _core_item_key(it.get('item', ''))
                if norm_i == it_norm or (core_i and (core_i == it_core or alias_i == it_core)):
                    return (it.get('std', '').strip(), it.get('method', '').strip(), ck, it.get('remark', ''))
            # 包含匹配项目名（如 合成着色剂）
            for it in bucket:
                it_norm = _norm_item_key(it.get('item', ''))
                it_core = _core_item_key(it.get('item', ''))
                if core_i and (core_i in it_norm or it_core in norm_i or (alias_i and alias_i in it_norm)):
                    return (it.get('std', '').strip(), it.get('method', '').strip(), ck, it.get('remark', ''))

    # 2. 若具体细类未命中，跨全细则库检索同名检验项目的通用判定依据与检测方法
    for _, c_data in SAMPLING_RULES_FULL_DB.items():
        for it in c_data.get('items', []):
            it_norm = _norm_item_key(it.get('item', ''))
            it_core = _core_item_key(it.get('item', ''))
            if norm_i == it_norm or (core_i and (core_i == it_core or alias_i == it_core)):
                return (it.get('std', '').strip(), it.get('method', '').strip(), c_data.get('product_name', '通用细类'), it.get('remark', ''))

    return None


def auto_match_record(standards, global_index, food_query, item_query, measured_val=None, cat1="", cat2="", cat3="", cat4=""):
    """
    四级分类 + 6大国标全量库 双引擎智能匹配与反写：
    支持传入 (四级细类/样品名称, 检验项目, 实测值) 以及可选的 (一级大类, 二级亚类, 三级品种, 四级细类)
    返回: (依据法律法规或标准, 标准全称, 标准限量值, 检测方法, 检出限LOD, 实测值, 判定结论, 匹配状态, 匹配分类路径)
    """
    raw_f = str(cat4 or food_query or cat3 or cat2 or cat1 or "").strip()
    raw_i = str(item_query or "").strip()
    if not raw_f and not raw_i:
        return ("/", "/", "/", "/", "/", "/", "/", "空白行", "/")

    # 整理四级分类层级列表（从细到粗：四级 -> 三级 -> 二级 -> 一级）
    cat_hierarchy = []
    for c in [cat4, food_query, cat3, cat2, cat1]:
        cs = str(c or "").strip()
        if cs and cs not in ("", "/", "None") and cs not in cat_hierarchy:
            cat_hierarchy.append(cs)

    # 第一步：查询《2026年国家食品安全监督抽检实施细则》四级分类库
    rule_hit = match_2026_sampling_rule(cat_hierarchy, raw_i)
    rule_std = rule_hit[0] if rule_hit else ""
    rule_method = rule_hit[1] if rule_hit else ""
    rule_cat = rule_hit[2] if rule_hit else ""

    # 第二步：在 6 大国标 (83,973条限量库) 中按四级分类层级检索具体限量值
    clean_i = re.sub(r'[\(（].*?[\)）]', '', raw_i).strip().lower()
    std_item_alias = ITEM_ALIAS_MAP.get(clean_i, clean_i)

    search_keys = []
    for sk in [clean_i, std_item_alias.lower(), raw_i.lower()]:
        if sk and sk not in search_keys:
            search_keys.append(sk)

    candidates = []
    if global_index:
        for sk in search_keys:
            if sk in global_index:
                candidates.extend(global_index[sk])
        if not candidates:
            for k, v in global_index.items():
                if clean_i and (clean_i in k or k in clean_i):
                    candidates.extend(v)

    # 构建带权重的食品分类候选词（四级细类权重最高，其次三级品种、二级亚类、一级大类及别名）
    food_candidates_weighted = []
    for level_idx, base_c in enumerate(cat_hierarchy):
        level_bonus = max(0, 12 - level_idx * 3)
        c_clean = re.sub(r'[\(（].*?[\)）]', '', base_c).strip()
        c_core = re.sub(r'^(新鲜|散装|有机|绿色|无公害|即食|预包装|冷冻|速冻|市售)\s*', '', c_clean).strip()
        for fc in [base_c, c_clean, c_core]:
            if fc:
                food_candidates_weighted.append((fc, level_bonus))
        for k_al in [c_core, c_clean, base_c]:
            if k_al in FOOD_ALIAS_MAP:
                for al in FOOD_ALIAS_MAP[k_al]:
                    food_candidates_weighted.append((al, max(0, level_bonus - 2)))

    best_cand = None
    best_score = -1

    # 若2026抽检细则已指明所属国标编号（如 GB 2762 / GB 2761 / GB 2760 / GB 2763 / GB 29921 / GB 31650），优先加分匹配该标准
    rule_std_num = ""
    if rule_std:
        m_num = re.search(r'(2760|2761|2762|2763|29921|31650)', rule_std)
        if m_num:
            rule_std_num = m_num.group(1)

    for std_code, it in candidates:
        cand_cat = it.get('food_category', '').lower()
        cand_item = it.get('test_item', '').lower()
        score = 0
        if rule_std_num and rule_std_num in std_code:
            score += 15

        cand_item_core = re.sub(r'[\(（].*?[\)）]', '', cand_item).strip()
        if clean_i and (clean_i == cand_item_core or std_item_alias.lower() == cand_item_core):
            score += 50
        elif clean_i in cand_item or cand_item_core in clean_i:
            score += 40

        food_hit = False
        best_fc_score = 0
        for fc, bonus in food_candidates_weighted:
            fc_lower = fc.lower()
            if fc_lower == cand_cat:
                best_fc_score = max(best_fc_score, 50 + bonus)
                food_hit = True
            else:
                subs = re.split(r'[,，、/()（）\s]+', cand_cat)
                if fc_lower in subs:
                    best_fc_score = max(best_fc_score, 45 + bonus)
                    food_hit = True
                elif fc_lower in cand_cat or cand_cat in fc_lower:
                    best_fc_score = max(best_fc_score, 35 + bonus)
                    food_hit = True

        score += best_fc_score
        if food_hit and score > best_score:
            best_score = score
            best_cand = (std_code, it)

    # 情况1：6大国标限量库 + 2026抽检细则库 双重命中或高分命中
    if best_cand and best_score >= 70:
        std_code, it = best_cand
        std_name = standards[std_code]['name'] if std_code in standards else "食品安全国家标准"
        basis_display = rule_std if rule_std else std_code
        limit_val = it.get('limit_value', '')
        unit = it.get('unit', '')
        limit_str = f"{limit_val} {unit}".strip() if limit_val else "详见标准要求"
        method_str = rule_method or it.get('method', '') or resolve_item_test_method("", raw_i, raw_f, std_code)
        lod_str = resolve_item_lod(raw_i, method_str, limit_val, unit)

        meas_display = "/"
        judge_display = "/"
        if measured_val is not None and str(measured_val).strip() not in ("", "/", "None", "--"):
            meas_display = str(measured_val).strip()
            _, j_text = judge_measured_value(limit_val, meas_display)
            judge_display = j_text

        status_tag = "✔ 四级分类精准匹配" if (rule_hit or best_score >= 88) else "⚠️ 品类层级推导"
        return (basis_display, f"{std_code} {std_name}", limit_str, method_str, lod_str, meas_display, judge_display, status_tag, it.get('food_category'))

    # 情况2：2026抽检细则四级分类库精准命中（如 GB 2760 范围外不得使用、部委禁用公告、产品明示标准等）
    if rule_hit:
        basis_display = rule_std if rule_std else "《2026年食品安全监督抽检实施细则》"
        method_str = rule_method or resolve_item_test_method("", raw_i, raw_f, basis_display)
        if "2760" in basis_display:
            std_full = "GB 2760-2024 食品添加剂使用标准"
            limit_str = "不得使用 (或按GB 2760规定)"
            limit_for_judge = "不得使用"
        elif any(k in basis_display for k in ["公告", "整顿办", "250号", "235号", "193号", "560号"]):
            std_full = "国家部委禁用/限用物质公告"
            limit_str = "不得检出 (禁止使用)"
            limit_for_judge = "不得检出"
        elif "2762" in basis_display:
            std_full = "GB 2762-2025 食品中污染物限量"
            limit_str = "详见GB 2762对应类别"
            limit_for_judge = ""
        elif "2761" in basis_display:
            std_full = "GB 2761-2017 食品中真菌毒素限量"
            limit_str = "详见GB 2761对应类别"
            limit_for_judge = ""
        elif "2763" in basis_display:
            std_full = "GB 2763-2026 食品中农药最大残留限量"
            limit_str = "详见GB 2763对应类别"
            limit_for_judge = ""
        elif "29921" in basis_display:
            std_full = "GB 29921-2021 预包装食品中致病菌限量"
            limit_str = "按GB 29921三级采样方案"
            limit_for_judge = ""
        elif "31650" in basis_display:
            std_full = "GB 31650-2019 食品中兽药最大残留限量"
            limit_str = "详见GB 31650对应类别"
            limit_for_judge = ""
        else:
            std_full = f"抽检细则指定依据 ({basis_display})"
            limit_str = "按产品明示标准及细则要求"
            limit_for_judge = ""

        lod_str = resolve_item_lod(raw_i, method_str, "", "")
        meas_display = "/"
        judge_display = "/"
        if measured_val is not None and str(measured_val).strip() not in ("", "/", "None", "--"):
            meas_display = str(measured_val).strip()
            if limit_for_judge:
                _, j_text = judge_measured_value(limit_for_judge, meas_display)
                judge_display = j_text
        return (basis_display, std_full, limit_str, method_str, lod_str, meas_display, judge_display, "✔ 2026细则精准匹配", rule_cat)

    # 情况3：GB 2760 常见禁用添加剂兜底规则
    if any(k in raw_i for k in ["脱氢乙酸", "合成着色剂", "柠檬黄", "日落黄", "胭脂红", "苋菜红", "亮蓝", "偶氮甲酰胺"]):
        basis_display = "GB 2760"
        std_full = "GB 2760-2024 食品添加剂使用标准"
        limit_str = "不得使用 (该食品分类未允许使用)"
        method_str = resolve_item_test_method("", raw_i, raw_f, "GB 2760")
        lod_str = resolve_item_lod(raw_i, method_str, "", "")
        meas_display = "/"
        judge_display = "/"
        if measured_val is not None and str(measured_val).strip() not in ("", "/", "None", "--"):
            meas_display = str(measured_val).strip()
            _, j_text = judge_measured_value("不得使用", meas_display)
            judge_display = j_text
        return (basis_display, std_full, limit_str, method_str, lod_str, meas_display, judge_display, "✔ GB 2760禁用匹配", raw_f)

    method_str = resolve_item_test_method("", raw_i, raw_f, "")
    lod_str = resolve_item_lod(raw_i, method_str, "", "")
    meas_display = str(measured_val).strip() if (measured_val and str(measured_val).strip() not in ("", "/", "None", "--")) else "/"
    return ("未查见对应标准", "/", "/", method_str if method_str != "/" else "参照推荐检验方法", lod_str, meas_display, "/", "❓ 待人工核实分类", "/")

class FoodStandardChecker:
    def __init__(self, root):
        self.root = root
        self.root.title("食品安全国家标准判定核查工具 V1.4")
        self.root.geometry("1260x800")
        self.root.minsize(1020, 660)

        self.ui_config = load_ui_config()
        if 'window_geometry' in self.ui_config:
            try:
                self.root.geometry(self.ui_config['window_geometry'])
            except Exception:
                pass

        self.root.protocol("WM_DELETE_WINDOW", self.on_window_close)

        self.standards = load_standards()
        self.global_index = build_global_item_index(self.standards)

        total_items = sum(len(v['items']) for v in self.standards.values())
        rules_count = len(SAMPLING_RULES_FULL_DB)
        self.root.title(f"食品安全国家标准判定核查工具 V1.4  —  6大国标 {total_items:,} 条限量 | 2026抽检细则 {rules_count} 类  |  Powered by 陈耿忠")

        self.judge_history = []
        self.selected_category_filter = ""
        self.tree_sort_states = {}
        self.standard_colors = {
            "GB 2760-2024": {"bg": "#d4e6f1", "fg": "#2471a3"},
            "GB 2761-2017": {"bg": "#d5f5e3", "fg": "#1e8449"},
            "GB 2762-2025": {"bg": "#fdebd0", "fg": "#d35400"},
            "GB 2763-2026": {"bg": "#e8daef", "fg": "#7d3c98"},
            "GB 29921-2021": {"bg": "#fadbd8", "fg": "#c0392b"},
            "GB 31650-2019": {"bg": "#d1f2eb", "fg": "#148f77"},
        }

        self.apply_global_styles()
        self.setup_ui()

    def apply_global_styles(self):
        """配置深邃午夜蓝商务主题、高可读性表格行高与无标签页头的容器样式"""
        style = ttk.Style()
        try:
            style.configure("Treeview", font=("Microsoft YaHei", 9), rowheight=30)
            style.configure("Treeview.Heading", font=("Microsoft YaHei", 9, "bold"))
            style.layout("Hidden.TNotebook.Tab", [])
            style.configure("Hidden.TNotebook", borderwidth=0, tabmargins=0)
        except Exception:
            pass

    def setup_ui(self):
        total_items = sum(len(v['items']) for v in self.standards.values())
        rules_count = len(SAMPLING_RULES_FULL_DB)

        # ==================== 1. 顶部深邃午夜蓝商务导航栏 (#0F172A) ====================
        header_bar = tk.Frame(self.root, bg="#0F172A", height=58)
        header_bar.pack(fill=tk.X, side=tk.TOP)
        header_bar.pack_propagate(False)

        # 左侧品牌与数据量徽章
        brand_frame = tk.Frame(header_bar, bg="#0F172A")
        brand_frame.pack(side=tk.LEFT, padx=(16, 12), pady=8)

        tk.Label(
            brand_frame,
            text="🛡️ 食品安全国家标准判定核查工具 V1.4",
            font=("Microsoft YaHei", 13, "bold"),
            fg="#F8FAFC",
            bg="#0F172A"
        ).pack(side=tk.LEFT)

        tk.Label(
            brand_frame,
            text=f" {len(self.standards)}大国标 {total_items:,}条 | 2026细则 {rules_count}类 ",
            font=("Microsoft YaHei", 8, "bold"),
            fg="#38BDF8",
            bg="#1E293B",
            padx=8,
            pady=3
        ).pack(side=tk.LEFT, padx=(10, 0))

        # 中部：核心功能下拉菜单 (Combobox) + 快捷切换胶囊按钮
        nav_frame = tk.Frame(header_bar, bg="#0F172A")
        nav_frame.pack(side=tk.LEFT, padx=10, pady=10)

        tk.Label(
            nav_frame,
            text="🧭 功能模块选择:",
            font=("Microsoft YaHei", 10, "bold"),
            fg="#E2E8F0",
            bg="#0F172A"
        ).pack(side=tk.LEFT, padx=(0, 6))

        self.module_options = [
            "🔍 1. 标准限量查询与单项判定 (6大国标全量库)",
            "📑 2. 四级分类与台账智能反写 (Excel自动匹配标准与方法)",
            "🧪 3. 多项目批量判定与表头筛选 (支持实测值批量核查)",
            "🔗 4. 样品跨标全景匹配 (一键穿透6大国标)"
        ]
        self.module_short_labels = [
            "🔍 标准限量查询",
            "📑 四级分类反写",
            "🧪 多项目批量判定",
            "🔗 样品跨标匹配"
        ]

        self.nav_combo_var = tk.StringVar(value=self.module_options[0])
        self.nav_combo = ttk.Combobox(
            nav_frame,
            textvariable=self.nav_combo_var,
            values=self.module_options,
            state="readonly",
            width=38,
            font=("Microsoft YaHei", 10, "bold")
        )
        self.nav_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.nav_combo.bind("<<ComboboxSelected>>", self.on_nav_combo_changed)

        # 快捷胶囊按钮组（与下拉菜单实时双向联动）
        self.nav_pill_btns = []
        for idx, short_lbl in enumerate(self.module_short_labels):
            btn = tk.Button(
                nav_frame,
                text=short_lbl,
                font=("Microsoft YaHei", 9, "bold"),
                bd=0,
                padx=10,
                pady=4,
                cursor="hand2",
                command=lambda i=idx: self.switch_module_by_index(i)
            )
            btn.pack(side=tk.LEFT, padx=3)
            self.nav_pill_btns.append(btn)

        # 右侧专家署名
        tk.Label(
            header_bar,
            text="⚡ Powered by 陈耿忠",
            font=("Microsoft YaHei", 9, "bold"),
            fg="#38BDF8",
            bg="#0F172A"
        ).pack(side=tk.RIGHT, padx=16)

        # ==================== 2. 底部商务状态栏 ====================
        status_bar = tk.Frame(self.root, bg="#1E293B", height=28)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)
        status_bar.pack_propagate(False)
        tk.Label(
            status_bar,
            text="🛡️ 核心标准库：GB 2760-2024 / GB 2761-2017 / GB 2762-2025 / GB 2763-2026 / GB 29921-2021 / GB 31650-2019 + 《2026年国家食品安全监督抽检实施细则》四级分类库",
            font=("Microsoft YaHei", 8),
            fg="#94A3B8",
            bg="#1E293B"
        ).pack(side=tk.LEFT, padx=12)
        tk.Label(
            status_bar,
            text="⚡ Powered by 陈耿忠 (CCAA 国家注册审核员 / 实验室质控专家)",
            font=("Microsoft YaHei", 8, "bold"),
            fg="#38BDF8",
            bg="#1E293B"
        ).pack(side=tk.RIGHT, padx=12)

        # ==================== 3. 主工作区容器 (隐藏原生简陋Tab条，由顶部下拉菜单控制) ====================
        self.notebook = ttk.Notebook(self.root, style="Hidden.TNotebook")
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)

        # 模块 0: 标准限量查询与单项判定
        self.tab_query = tb.Frame(self.notebook)
        self.notebook.add(self.tab_query, text=self.module_options[0])
        self.setup_query_tab()

        # 模块 1: 四级分类与台账智能反写 (Excel自动匹配)
        self.tab_excel_backfill = tb.Frame(self.notebook)
        self.notebook.add(self.tab_excel_backfill, text=self.module_options[1])
        self.setup_excel_backfill_tab()

        # 模块 2: 多项目批量查询与判定
        self.tab_batch = tb.Frame(self.notebook)
        self.notebook.add(self.tab_batch, text=self.module_options[2])
        self.setup_batch_tab()

        # 模块 3: 样品跨标全景匹配
        self.tab_match = tb.Frame(self.notebook)
        self.notebook.add(self.tab_match, text=self.module_options[3])
        self.setup_match_tab()

        # 初始激活第一个模块并刷新
        self.switch_module_by_index(0)
        self.on_standard_selected()

    def on_nav_combo_changed(self, event=None):
        """响应顶部下拉菜单切换功能模块"""
        sel = self.nav_combo_var.get()
        if sel in self.module_options:
            idx = self.module_options.index(sel)
            self.switch_module_by_index(idx)

    def switch_module_by_index(self, idx):
        """切换至指定序号的功能模块并同步下拉菜单与胶囊按钮高亮"""
        if 0 <= idx < len(self.module_options):
            self.nav_combo_var.set(self.module_options[idx])
            self.notebook.select(idx)
            for i, btn in enumerate(self.nav_pill_btns):
                if i == idx:
                    btn.configure(bg="#0284C7", fg="#FFFFFF", activebackground="#0369A1", activeforeground="#FFFFFF")
                else:
                    btn.configure(bg="#1E293B", fg="#CBD5E1", activebackground="#334155", activeforeground="#FFFFFF")

    def setup_query_tab(self):
        """模块1：标准限量查询与单项判定 UI (商务分栏卡片布局)"""
        main_frame = tb.Frame(self.tab_query)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 左侧：查询与单项判定控制面板
        left_frame = tb.Frame(main_frame, width=395)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 8))
        left_frame.pack_propagate(False)

        # 卡片1：标准与分类筛选
        filter_card = tb.Labelframe(left_frame, text=" 📚 1. 国家标准与食品分类检索 ", padding=10)
        filter_card.pack(fill=tk.X, pady=(0, 8))

        tb.Label(filter_card, text="选择食品安全国家标准:", font=("Microsoft YaHei", 9, "bold"), anchor="w").pack(fill=tk.X, pady=(0, 3))
        self.standard_var = tk.StringVar()
        standard_codes = sorted(self.standards.keys())
        self.standard_combo = ttk.Combobox(filter_card, textvariable=self.standard_var,
                                           values=standard_codes, state="readonly", font=("Microsoft YaHei", 10, "bold"))
        self.standard_combo.pack(fill=tk.X, pady=(0, 3))
        if standard_codes:
            self.standard_combo.current(0)
        self.standard_combo.bind('<<ComboboxSelected>>', self.on_standard_selected)

        self.standard_info_label = tb.Label(filter_card, text="", font=("Microsoft YaHei", 9), bootstyle=PRIMARY, anchor="w", wraplength=360, justify=tk.LEFT)
        self.standard_info_label.pack(fill=tk.X, pady=(0, 6))

        tb.Label(filter_card, text="🍽️ 样品/食品类别下拉筛选:", font=("Microsoft YaHei", 9, "bold"), anchor="w").pack(fill=tk.X, pady=(0, 3))
        self.selected_category_filter = ""
        self.cat_combo_var = tk.StringVar(value="全部食品分类 (显示全部)")
        self.cat_combo = ttk.Combobox(filter_card, textvariable=self.cat_combo_var, state="readonly", font=("Microsoft YaHei", 9))
        self.cat_combo.pack(fill=tk.X, pady=(0, 4))
        self.cat_combo.bind("<<ComboboxSelected>>", self.on_cat_combo_selected)

        cat_search_row = tb.Frame(filter_card)
        cat_search_row.pack(fill=tk.X, pady=(0, 6))
        tb.Label(cat_search_row, text="类别快筛:", font=("Microsoft YaHei", 8), bootstyle=SECONDARY).pack(side=tk.LEFT, padx=(0, 4))
        self.cat_search_var = tk.StringVar()
        self.cat_search_var.trace('w', lambda *args: self.filter_category_dropdown())
        tb.Entry(cat_search_row, textvariable=self.cat_search_var, font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        tb.Button(cat_search_row, text="🔄 重置", command=self.clear_category_filter, bootstyle="secondary-outline").pack(side=tk.RIGHT)

        tb.Label(filter_card, text="🔍 关键词搜索 (食品名称 / 检测项目 / 方法号):", font=("Microsoft YaHei", 9, "bold"), anchor="w").pack(fill=tk.X, pady=(2, 3))
        search_frame = tb.Frame(filter_card)
        search_frame.pack(fill=tk.X)
        self.search_var = tk.StringVar()
        self.search_var.trace('w', lambda *args: self.on_search())
        tb.Entry(search_frame, textvariable=self.search_var, font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tb.Button(search_frame, text="清除", command=self.clear_search, bootstyle=SECONDARY).pack(side=tk.RIGHT, padx=(4, 0))

        self.result_count_label = tb.Label(filter_card, text="", font=("Microsoft YaHei", 9, "bold"), bootstyle=INFO, anchor="w")
        self.result_count_label.pack(fill=tk.X, pady=(4, 0))

        # 卡片2：单项实测值合规判定
        judge_frame = tb.Labelframe(left_frame, text=" ⚖️ 2. 单项限量核查与检验方法复制 ", padding=10)
        judge_frame.pack(fill=tk.X, pady=(0, 8))

        self.selected_info_label = tb.Label(judge_frame, text="请在右侧表格点击选择一条标准数据",
                                           font=("Microsoft YaHei", 9), anchor="w", wraplength=355, justify=tk.LEFT)
        self.selected_info_label.pack(fill=tk.X, pady=(0, 4))

        self.limit_display_label = tb.Label(judge_frame, text="限量值: --",
                                           font=("Microsoft YaHei", 10, "bold"), bootstyle=DANGER, anchor="w")
        self.limit_display_label.pack(fill=tk.X)

        method_container = tb.Frame(judge_frame)
        method_container.pack(fill=tk.X, pady=(4, 2))
        m_header = tb.Frame(method_container)
        m_header.pack(fill=tk.X, pady=(0, 2))
        tb.Label(m_header, text="🔬 要求检验方法标准 (可划选复制):", font=("Microsoft YaHei", 9, "bold"), bootstyle=PRIMARY).pack(side=tk.LEFT)
        self.btn_copy_method = tb.Button(m_header, text="📋 一键复制", command=self.copy_current_selected_method, bootstyle="info-outline", padding=(4, 1), cursor="hand2")
        self.btn_copy_method.pack(side=tk.RIGHT)

        self.method_text = tk.Text(method_container, height=2, font=("Consolas", 9, "bold"), wrap=tk.WORD,
                                   bg="#F0F9FF", fg="#0369A1", bd=1, relief=tk.SOLID, padx=4, pady=3)
        self.method_text.pack(fill=tk.X)
        self.method_text.insert("1.0", "请在右侧表格选择一条数据查看检验方法")
        self.method_text.bind("<Key>", lambda e: "break" if (e.state != 4 and e.keysym not in ('c', 'C', 'a', 'A')) else None)

        iv_frame = tb.Frame(judge_frame)
        iv_frame.pack(fill=tk.X, pady=(6, 0))
        tb.Label(iv_frame, text="实测值:", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT)
        self.test_value_var = tk.StringVar()
        tb.Entry(iv_frame, textvariable=self.test_value_var, font=("Microsoft YaHei", 10), width=11).pack(side=tk.LEFT, padx=5)
        tb.Button(iv_frame, text="⚡ 立即判定", command=self.do_judge, bootstyle=SUCCESS).pack(side=tk.LEFT)
        self.judge_result_label = tb.Label(iv_frame, text="", font=("Microsoft YaHei", 11, "bold"))
        self.judge_result_label.pack(side=tk.LEFT, padx=(8, 0))

        # 卡片3：判定历史与列宽工具
        hist_card = tb.Labelframe(left_frame, text=" 📝 3. 单项判定记录 ", padding=8)
        hist_card.pack(fill=tk.BOTH, expand=True)

        btn_frame = tb.Frame(hist_card)
        btn_frame.pack(fill=tk.X, pady=(0, 4))
        tb.Button(btn_frame, text="💾 导出记录", command=self.export_history, bootstyle=INFO, padding=(6, 2)).pack(side=tk.LEFT)
        tb.Button(btn_frame, text="↔ 自适应列宽", command=lambda: self.autofit_tree_columns(self.tree), bootstyle="info-outline", padding=(6, 2)).pack(side=tk.LEFT, padx=4)
        tb.Button(btn_frame, text="🗑️ 清空", command=self.clear_history, bootstyle="danger-outline", padding=(6, 2)).pack(side=tk.RIGHT)

        self.history_text = tk.Text(hist_card, height=4, font=("Consolas", 8), wrap=tk.WORD)
        hist_scroll = ttk.Scrollbar(hist_card, orient=tk.VERTICAL, command=self.history_text.yview)
        self.history_text.configure(yscrollcommand=hist_scroll.set)
        self.history_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        hist_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 右侧：数据列表卡片
        right_frame = tb.Labelframe(main_frame, text=" 📋 标准限量与检验方法明细表 (精准匹配优先置顶 | 双击行查看完整指标卡 | 右键复制或调整列宽) ", padding=6)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        columns = ("食品分类", "检测项目", "限量值", "单位", "检验方法", "备注")
        self.tree = ttk.Treeview(right_frame, columns=columns, show="headings", selectmode="browse")

        for col in columns:
            self.tree.heading(col, text=col + " ▾", command=lambda c=col: self._sort_query_tree(c))

        self.default_query_widths = {"食品分类": 180, "检测项目": 140, "限量值": 80, "单位": 60, "检验方法": 320, "备注": 220}
        self.apply_tree_column_widths(self.tree, "query_tree", self.default_query_widths)

        self.tree.tag_configure('highlight_match', background='#FEF9C3', foreground='#991B1B')
        self.tree.tag_configure('even_row', background='#F8FAFC')
        self.tree.tag_configure('odd_row', background='#FFFFFF')

        v_scroll = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scroll = ttk.Scrollbar(right_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")

        right_frame.grid_rowconfigure(0, weight=1)
        right_frame.grid_columnconfigure(0, weight=1)

        self.tree.bind("<<TreeviewSelect>>", self.on_item_selected)
        self.tree.bind("<Double-Button-1>", self.on_query_tree_double_click)
        self.setup_tree_context_menu(self.tree, "query_tree", self.default_query_widths, self.on_query_tree_double_click)
    def setup_match_tab(self):
        """样品匹配 Tab UI"""
        main_frame = tb.Frame(self.tab_match)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标准颜色映射
        self.standard_colors = {
            "GB 2760-2024": {"bg": "#d4e6f1", "fg": "#2471a3"},
            "GB 2761-2017": {"bg": "#d5f5e3", "fg": "#1e8449"},
            "GB 2762-2025": {"bg": "#fdebd0", "fg": "#d35400"},
            "GB 2763-2026": {"bg": "#e8daef", "fg": "#7d3c98"},
            "GB 29921-2021": {"bg": "#fadbd8", "fg": "#c0392b"},
            "GB 31650-2019": {"bg": "#d1f2eb", "fg": "#148f77"},
        }
        
        # 顶部搜索区
        search_bar = tb.Frame(main_frame, height=50)
        search_bar.pack(fill=tk.X)
        search_bar.pack_propagate(False)
        
        tb.Label(search_bar, text="样品匹配 - 输入食品名称，跨所有标准查询限量值", 
                font=("Microsoft YaHei", 11, "bold")).pack(side=tk.LEFT, padx=(15,15), pady=12)
        
        tb.Label(search_bar, text="食品名称:", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, pady=12)
        self.match_search_var = tk.StringVar()
        match_entry = tb.Entry(search_bar, textvariable=self.match_search_var, 
                               font=("Microsoft YaHei", 10), width=30)
        match_entry.pack(side=tk.LEFT, padx=(5,10), pady=12)
        match_entry.bind('<Return>', lambda e: self.do_sample_match())
        
        tb.Button(search_bar, text="匹配", command=self.do_sample_match, bootstyle=PRIMARY).pack(side=tk.LEFT, pady=10)
        
        # 联合查询复选框（默认勾选）
        self.match_unified_var = tk.BooleanVar(value=True)
        self.match_unified_cb = tb.Checkbutton(search_bar, text="联合查询",
                                               variable=self.match_unified_var,
                                               command=self._on_unified_toggle)
        self.match_unified_cb.pack(side=tk.LEFT, padx=(5,5), pady=10)
        
        tb.Button(search_bar, text="导出结果...", command=self.export_match_results, bootstyle=SUCCESS).pack(side=tk.LEFT, padx=(5,0), pady=10)
        tb.Button(search_bar, text="↔ 自适应列宽", command=lambda: self.autofit_tree_columns(self.match_unified_tree), bootstyle="info-outline").pack(side=tk.LEFT, padx=(5,0), pady=10)
        tb.Button(search_bar, text="清除", command=self.clear_match, bootstyle=SECONDARY).pack(side=tk.LEFT, padx=(5,0), pady=10)
        
        self.match_count_label = tb.Label(search_bar, text="", font=("Microsoft YaHei", 9))
        self.match_count_label.pack(side=tk.RIGHT, padx=15, pady=12)
        
        # 提示标签
        self.match_hint_label = tb.Label(main_frame, text="支持多关键词（空格分隔），关键词之间为 AND 关系  |  💡 提示：双击行查看详情，右键可自适应/重置列宽", 
                                         font=("Microsoft YaHei", 9), anchor="w")
        self.match_hint_label.pack(fill=tk.X, padx=5, pady=2)
        
        # ===== 联合查询统一列表 =====
        self.match_unified_frame = tb.Frame(main_frame)
        self.match_unified_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ("标准名称", "检测项目", "食品类别", "限量值", "检验方法", "备注")
        self.match_unified_tree = ttk.Treeview(self.match_unified_frame, columns=columns, 
                                               show="headings", selectmode="browse")
        
        self.match_unified_tree.heading("标准名称", text="标准名称",
                                        command=lambda: self._sort_match_tree("标准名称"))
        self.match_unified_tree.heading("检测项目", text="检测项目",
                                        command=lambda: self._sort_match_tree("检测项目"))
        self.match_unified_tree.heading("食品类别", text="食品类别",
                                        command=lambda: self._sort_match_tree("食品类别"))
        self.match_unified_tree.heading("限量值", text="限量值",
                                        command=lambda: self._sort_match_tree("限量值"))
        self.match_unified_tree.heading("检验方法", text="检验方法",
                                        command=lambda: self._sort_match_tree("检验方法"))
        self.match_unified_tree.heading("备注", text="备注",
                                        command=lambda: self._sort_match_tree("备注"))
        
        self.default_match_widths = {"标准名称": 120, "检测项目": 130, "食品类别": 180, "限量值": 90, "检验方法": 280, "备注": 200}
        self.apply_tree_column_widths(self.match_unified_tree, "match_unified_tree", self.default_match_widths)
        
        # 为每个标准配置彩色标签
        for std_code, colors in self.standard_colors.items():
            tag_name = std_code.replace(" ", "_").replace("-", "_")
            self.match_unified_tree.tag_configure(tag_name,
                                                  background=colors["bg"],
                                                  foreground=colors["fg"])
        
        self.match_unified_tree.bind("<Double-Button-1>", self.on_match_unified_tree_double_click)
        self.setup_tree_context_menu(self.match_unified_tree, "match_unified_tree", self.default_match_widths, self.on_match_unified_tree_double_click)
        
        v_scroll_u = ttk.Scrollbar(self.match_unified_frame, orient=tk.VERTICAL,
                                   command=self.match_unified_tree.yview)
        h_scroll_u = ttk.Scrollbar(self.match_unified_frame, orient=tk.HORIZONTAL,
                                   command=self.match_unified_tree.xview)
        self.match_unified_tree.configure(yscrollcommand=v_scroll_u.set, xscrollcommand=h_scroll_u.set)
        
        self.match_unified_tree.grid(row=0, column=0, sticky="nsew")
        v_scroll_u.grid(row=0, column=1, sticky="ns")
        h_scroll_u.grid(row=1, column=0, sticky="ew")
        
        self.match_unified_frame.grid_rowconfigure(0, weight=1)
        self.match_unified_frame.grid_columnconfigure(0, weight=1)
        
        # 排序状态
        self._sort_col = None
        self._sort_reverse = False
        self._match_data_cache = []
        
        # ===== 分Tab结果区（联合查询关闭时显示，初始隐藏） =====
        self.match_tab_frame = tb.Frame(main_frame)
        
        self.match_notebook = ttk.Notebook(self.match_tab_frame)
        self.match_notebook.pack(fill=tk.BOTH, expand=True)
        
        self.match_trees = {}
        sorted_standards = sorted(self.standards.keys())
        for std_code in sorted_standards:
            tab = tb.Frame(self.match_notebook)
            self.match_notebook.add(tab, text=std_code)
            
            columns_tab = ("检测项目", "食品类别", "限量值", "备注")
            tree = ttk.Treeview(tab, columns=columns_tab, show="headings", selectmode="browse")
            
            tree.heading("检测项目", text="检测项目")
            tree.heading("食品类别", text="食品类别")
            tree.heading("限量值", text="限量值")
            tree.heading("备注", text="备注")
            
            tree.column("检测项目", width=160)
            tree.column("食品类别", width=220)
            tree.column("限量值", width=100, anchor="center")
            tree.column("备注", width=350)
            
            tree.bind("<Double-Button-1>", lambda e, s=std_code, t=tree: self.on_match_tab_tree_double_click(s, t))
            
            v_scroll = ttk.Scrollbar(tab, orient=tk.VERTICAL, command=tree.yview)
            h_scroll = ttk.Scrollbar(tab, orient=tk.HORIZONTAL, command=tree.xview)
            tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
            
            tree.grid(row=0, column=0, sticky="nsew")
            v_scroll.grid(row=0, column=1, sticky="ns")
            h_scroll.grid(row=1, column=0, sticky="ew")
            
            tab.grid_rowconfigure(0, weight=1)
            tab.grid_columnconfigure(0, weight=1)
            
            self.match_trees[std_code] = tree
    
    def _on_unified_toggle(self):
        """联合查询复选框切换"""
        if self.match_unified_var.get():
            self.match_tab_frame.pack_forget()
            self.match_unified_frame.pack(fill=tk.BOTH, expand=True)
        else:
            self.match_unified_frame.pack_forget()
            self.match_tab_frame.pack(fill=tk.BOTH, expand=True)
        # 如果已有搜索关键词，重新执行匹配
        if self.match_search_var.get().strip():
            self.do_sample_match()
    
    def do_sample_match(self):
        """执行样品跨标准匹配"""
        keyword_text = self.match_search_var.get().strip()
        if not keyword_text:
            self.match_count_label.config(text="请输入食品名称")
            return
        
        raw_keywords = [kw.strip() for kw in keyword_text.split() if kw.strip()]
        keywords = [kw.lower() for kw in raw_keywords]
        unified = self.match_unified_var.get()
        
        if unified:
            # 联合查询模式：统一列表
            self.match_unified_tree.delete(*self.match_unified_tree.get_children())
            self._match_data_cache = []
            total_hits = 0
            
            for std_code in sorted(self.standards.keys()):
                items = self.standards[std_code]['items']
                for item in items:
                    food_cat = item.get('food_category', '')
                    if all(kw in food_cat.lower() for kw in keywords):
                        m_str = item.get('method', '')
                        raw_row = (
                            std_code,
                            item.get('test_item', ''),
                            item.get('food_category', ''),
                            item.get('limit_value', ''),
                            m_str,
                            item.get('notes', '')
                        )
                        disp_row = (
                            std_code,
                            item.get('test_item', ''),
                            highlight_keywords_text(food_cat, raw_keywords),
                            item.get('limit_value', ''),
                            m_str,
                            item.get('notes', '')
                        )
                        self._match_data_cache.append((raw_row, disp_row))
                        total_hits += 1
            
            # 核心改进：精准匹配优先排序 (优先显示完全精准匹配的食品分类或项目，其次按标准名称)
            self._match_data_cache.sort(key=lambda r: (calc_match_rank({'test_item': r[0][1], 'food_category': r[0][2], 'method': r[0][4]}, keyword_text), r[0][0]))
            for raw_row, disp_row in self._match_data_cache:
                tag_name = raw_row[0].replace(" ", "_").replace("-", "_")
                if tag_name not in self.standard_colors:
                    tag_name = ""
                self.match_unified_tree.insert('', 'end', values=disp_row, tags=(tag_name,))
            
            self.match_count_label.config(text=f"共匹配 {total_hits} 条记录（联合查询）")
        else:
            # 分Tab模式
            total_hits = 0
            for std_code in sorted(self.standards.keys()):
                tree = self.match_trees[std_code]
                tree.delete(*tree.get_children())
                
                items = self.standards[std_code]['items']
                std_hits = 0
                
                for item in items:
                    food_cat = item.get('food_category', '')
                    if all(kw in food_cat.lower() for kw in keywords):
                        disp_food = highlight_keywords_text(food_cat, raw_keywords)
                        tree.insert('', 'end', values=(
                            item.get('test_item', ''),
                            disp_food,
                            item.get('limit_value', ''),
                            item.get('notes', '')
                        ))
                        std_hits += 1
                
                total_hits += std_hits
                
                for i in range(self.match_notebook.index('end')):
                    if self.match_notebook.tab(i, 'text') == std_code:
                        title = f"{std_code} ({std_hits})" if std_hits > 0 else std_code
                        self.match_notebook.tab(i, text=title)
                        break
            
            self.match_count_label.config(text=f"共匹配 {total_hits} 条记录（分标准）")
    
    def _sort_match_tree(self, col):
        """点击列头排序样品匹配统一列表"""
        cols = ("标准名称", "检测项目", "食品类别", "限量值", "备注")
        universal_tree_sort_column(self.match_unified_tree, col, "match_unified_tree", self.tree_sort_states, cols, on_reload_callback=self.do_sample_match)

    def _sort_query_tree(self, col):
        """点击列头排序标准查询列表"""
        cols = ("食品分类", "检测项目", "限量值", "单位", "备注")
        universal_tree_sort_column(self.tree, col, "query_tree", self.tree_sort_states, cols, on_reload_callback=self.on_search)

    def _sort_batch_tree(self, col):
        """点击列头排序批量判定列表"""
        universal_tree_sort_column(self.batch_tree, col, "batch_tree", self.tree_sort_states, self.batch_columns, on_reload_callback=self.filter_batch_results)

    def clear_match(self):
        """清除样品匹配结果"""
        self.match_search_var.set("")
        # 清除统一列表
        self.match_unified_tree.delete(*self.match_unified_tree.get_children())
        self._match_data_cache = []
        # 清除分Tab
        for std_code, tree in self.match_trees.items():
            tree.delete(*tree.get_children())
            for i in range(self.match_notebook.index('end')):
                if self.match_notebook.tab(i, 'text').startswith(std_code):
                    self.match_notebook.tab(i, text=std_code)
                    break
        self.match_count_label.config(text="")
    
    def setup_batch_tab(self):
        """多项目批量查询 Tab UI"""
        main_frame = tb.Frame(self.tab_batch)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 顶部：输入区
        input_frame = tb.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 左侧：食品名称
        left_input = tb.Frame(input_frame)
        left_input.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        tb.Label(left_input, text="食品名称 (可选):", font=("Microsoft YaHei", 10)).pack(anchor="w")
        self.batch_food_var = tk.StringVar()
        tb.Entry(left_input, textvariable=self.batch_food_var, font=("Microsoft YaHei", 10), width=20).pack(pady=(4, 0))
        
        tb.Label(left_input, text="批量检测项目 (实测值可选):", font=("Microsoft YaHei", 10, "bold")).pack(anchor="w", pady=(8, 0))
        tb.Label(left_input, text="每行支持格式 (实测值选填):\n• 铅\n• 脱氢乙酸, 0.05\n• 黄曲霉毒素B1: 10\n• 菌落总数 50000\n(带实测值自动判定合规/超标)", font=("Microsoft YaHei", 8), bootstyle=SECONDARY, justify=tk.LEFT).pack(anchor="w")
        
        # 右侧：检测项目多行输入
        right_input = tb.Frame(input_frame)
        right_input.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.batch_items_text = tk.Text(right_input, height=6, font=("Microsoft YaHei", 10))
        text_scroll = ttk.Scrollbar(right_input, orient=tk.VERTICAL, command=self.batch_items_text.yview)
        self.batch_items_text.configure(yscrollcommand=text_scroll.set)
        self.batch_items_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        text_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 按钮区
        btn_frame = tb.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        tb.Button(btn_frame, text="批量查询与判定", command=self.do_batch_query, bootstyle=PRIMARY).pack(side=tk.LEFT)
        tb.Button(btn_frame, text="导入项目及实测值...", command=self.import_batch_items, bootstyle=INFO).pack(side=tk.LEFT, padx=(10, 0))
        tb.Button(btn_frame, text="导出当前结果...", command=self.export_batch_results, bootstyle=SUCCESS).pack(side=tk.LEFT, padx=(10, 0))
        tb.Button(btn_frame, text="↔ 自适应列宽", command=lambda: self.autofit_tree_columns(self.batch_tree), bootstyle="info-outline").pack(side=tk.LEFT, padx=(10, 0))
        tb.Button(btn_frame, text="清空", command=self.clear_batch, bootstyle=SECONDARY).pack(side=tk.LEFT, padx=10)
        
        self.batch_count_label = tb.Label(btn_frame, text="", font=("Microsoft YaHei", 9))
        self.batch_count_label.pack(side=tk.LEFT, padx=10)
        
        # 结果多维智能筛选工具栏
        filter_frame = tb.LabelFrame(main_frame, text=" 🔍 结果多维过滤与快捷筛选 ", padding=6)
        filter_frame.pack(fill=tk.X, pady=(0, 6))

        filter_row = tb.Frame(filter_frame)
        filter_row.pack(fill=tk.X)

        # 1. 判定结论过滤
        tb.Label(filter_row, text="判定结论:", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT, padx=(4, 2))
        self.batch_status_filter_var = tk.StringVar(value="全部结论")
        status_opts = ["全部结论", "❌ 超标/不合格", "✅ 合格", "⚪ 未判定 (/)"]
        self.status_cb = ttk.Combobox(filter_row, textvariable=self.batch_status_filter_var, values=status_opts, state="readonly", width=12, font=("Microsoft YaHei", 9))
        self.status_cb.pack(side=tk.LEFT, padx=(0, 10))
        self.status_cb.bind("<<ComboboxSelected>>", lambda e: self.filter_batch_results())

        # 2. 适用标准过滤
        tb.Label(filter_row, text="适用标准:", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT, padx=(4, 2))
        self.batch_std_filter_var = tk.StringVar(value="全部标准")
        std_opts = ["全部标准", "GB 2760", "GB 2761", "GB 2762", "GB 2763", "GB 29921", "GB 31650"]
        self.std_filter_cb = ttk.Combobox(filter_row, textvariable=self.batch_std_filter_var, values=std_opts, state="readonly", width=12, font=("Microsoft YaHei", 9))
        self.std_filter_cb.pack(side=tk.LEFT, padx=(0, 10))
        self.std_filter_cb.bind("<<ComboboxSelected>>", lambda e: self.filter_batch_results())

        # 3. 关键词过滤
        tb.Label(filter_row, text="关键词搜索:", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT, padx=(4, 2))
        self.batch_filter_var = tk.StringVar()
        self.batch_filter_var.trace('w', lambda *args: self.filter_batch_results())
        tb.Entry(filter_row, textvariable=self.batch_filter_var, font=("Microsoft YaHei", 9), width=20).pack(side=tk.LEFT, padx=(0, 8))

        # 4. 重置筛选按钮
        tb.Button(filter_row, text="🔄 重置全部筛选", command=self.reset_batch_filters, bootstyle="secondary-outline").pack(side=tk.LEFT, padx=4)

        tb.Label(filter_row, text="💡 提示：按填入顺序展示 | 结果支持列头点击筛选与双击查看详情", font=("Microsoft YaHei", 8), bootstyle=SECONDARY).pack(side=tk.RIGHT, padx=6)
        
        self._batch_data_cache = []
        self.batch_columns = ("标准名称", "食品类别", "检测项目", "限量值", "实测值", "判定结论", "检验方法", "备注")
        self.batch_col_filters = {col: None for col in self.batch_columns}
        
        # 结果显示区 (Treeview)
        result_frame = tb.Frame(main_frame)
        result_frame.pack(fill=tk.BOTH, expand=True)
        
        self.batch_tree = ttk.Treeview(result_frame, columns=self.batch_columns, show="headings", selectmode="browse")
        
        for col in self.batch_columns:
            self.batch_tree.heading(col, text=col + " ▾", command=lambda c=col: self._sort_batch_tree(c))
            
        self.default_batch_widths = {"标准名称": 110, "食品类别": 140, "检测项目": 120, "限量值": 80, "实测值": 80, "判定结论": 110, "检验方法": 280, "备注": 180}
        self.apply_tree_column_widths(self.batch_tree, "batch_tree", self.default_batch_widths)
        
        # 为每个标准配置彩色标签
        for std_code, colors in self.standard_colors.items():
            tag_name = std_code.replace(" ", "_").replace("-", "_")
            self.batch_tree.tag_configure(tag_name, background=colors["bg"], foreground=colors["fg"])
            
        # 判定结果专属行标签
        self.batch_tree.tag_configure('judge_fail', background='#ffebe9', foreground='#cf222e')
        self.batch_tree.tag_configure('judge_pass', background='#dafbe1', foreground='#1a7f37')
        self.batch_tree.tag_configure('judge_unknown', background='#fff8c5', foreground='#9a6700')
            
        self.batch_tree.bind("<Double-Button-1>", self.on_batch_tree_double_click)
        self.setup_tree_context_menu(self.batch_tree, "batch_tree", self.default_batch_widths, self.on_batch_tree_double_click)
            
        v_scroll_b = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.batch_tree.yview)
        h_scroll_b = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=self.batch_tree.xview)
        self.batch_tree.configure(yscrollcommand=v_scroll_b.set, xscrollcommand=h_scroll_b.set)
        
        self.batch_tree.grid(row=0, column=0, sticky="nsew")
        v_scroll_b.grid(row=0, column=1, sticky="ns")
        h_scroll_b.grid(row=1, column=0, sticky="ew")
        
        result_frame.grid_rowconfigure(0, weight=1)
        result_frame.grid_columnconfigure(0, weight=1)

    def reset_batch_filters(self):
        """重置批量判定所有筛选条件"""
        if hasattr(self, 'batch_status_filter_var'):
            self.batch_status_filter_var.set("全部结论")
        if hasattr(self, 'batch_std_filter_var'):
            self.batch_std_filter_var.set("全部标准")
        if hasattr(self, 'batch_filter_var'):
            self.batch_filter_var.set("")
        for col in self.batch_columns:
            self.batch_col_filters[col] = None
            self.batch_tree.heading(col, text=col + " ▾")
        self.filter_batch_results()

    def clear_batch(self):
        self.batch_food_var.set("")
        self.batch_items_text.delete("1.0", tk.END)
        self.batch_tree.delete(*self.batch_tree.get_children())
        self.batch_count_label.config(text="")
        self.reset_batch_filters()
        self._batch_data_cache = []

    def on_batch_header_click(self, col_name):
        if not self._batch_data_cache:
            return
            
        col_idx = self.batch_columns.index(col_name)
        unique_values = set()
        for item_entry in self._batch_data_cache:
            raw_row = item_entry[0]
            unique_values.add(str(raw_row[col_idx]))
            
        unique_values = sorted(list(unique_values))
        current_filters = self.batch_col_filters[col_name]
        
        def on_filter_applied(selected_values):
            self.batch_col_filters[col_name] = selected_values
            # 更新表头文字
            if selected_values is None:
                self.batch_tree.heading(col_name, text=col_name + " ▾")
            else:
                self.batch_tree.heading(col_name, text=col_name + " (已筛选) ▾")
            self.filter_batch_results()
            
        dialog = ColumnFilterDialog(self.root, col_name, unique_values, current_filters, on_filter_applied)

    def import_batch_items(self):
        file_path = filedialog.askopenfilename(
            title="导入批量项目及实测值",
            filetypes=[("文本或CSV文件", "*.txt;*.csv"), ("文本文件", "*.txt"), ("CSV表格文件", "*.csv"), ("所有文件", "*.*")]
        )
        if not file_path:
            return
        try:
            with open(file_path, "r", encoding="utf-8-sig", errors="ignore") as f:
                lines = f.readlines()
            
            current_text = self.batch_items_text.get("1.0", tk.END).strip()
            new_items = []
            for line in lines:
                l = line.strip()
                if not l:
                    continue
                if l.startswith("检测项目") or l.startswith("项目名称") or l.startswith("标准名称"):
                    continue
                if l and l not in new_items:
                    new_items.append(l)
            
            if current_text:
                combined = current_text + "\n" + "\n".join(new_items)
            else:
                combined = "\n".join(new_items)
            
            self.batch_items_text.delete("1.0", tk.END)
            self.batch_items_text.insert(tk.END, combined)
            messagebox.showinfo("导入成功", f"成功导入 {len(new_items)} 行项目数据！\n支持纯项目名或带实测值自动比对。")
        except Exception as e:
            messagebox.showerror("导入失败", f"无法读取文件: {e}")

    def filter_batch_results(self):
        filter_text = self.batch_filter_var.get().strip() if hasattr(self, 'batch_filter_var') else ""
        filter_lower = filter_text.lower()
        
        status_filter = self.batch_status_filter_var.get() if hasattr(self, 'batch_status_filter_var') else "全部结论"
        std_filter = self.batch_std_filter_var.get() if hasattr(self, 'batch_std_filter_var') else "全部标准"

        self.batch_tree.delete(*self.batch_tree.get_children())
        
        display_count = 0
        for item_entry in self._batch_data_cache:
            if len(item_entry) == 3:
                raw_row, tag_name, matched_kws = item_entry
            else:
                raw_row, tag_name = item_entry
                matched_kws = []
                
            std_code = str(raw_row[0])
            food_cat = str(raw_row[1])
            test_item = str(raw_row[2])
            limit_val = str(raw_row[3])
            meas_val = str(raw_row[4])
            judge_res = str(raw_row[5])
            method_val = str(raw_row[6]) if len(raw_row) > 6 else "/"
            notes = str(raw_row[7]) if len(raw_row) > 7 else (raw_row[6] if len(raw_row) > 6 else "")

            # 1. 判定状态筛选
            if status_filter == "❌ 超标/不合格":
                if tag_name != 'judge_fail' and "超标" not in judge_res and "不合格" not in judge_res:
                    continue
            elif status_filter == "✅ 合格":
                if tag_name != 'judge_pass' and "合格" not in judge_res:
                    continue
            elif status_filter == "⚪ 未判定 (/)":
                if judge_res != "/" and meas_val != "/":
                    continue

            # 2. 适用标准筛选
            if std_filter != "全部标准":
                if std_filter not in std_code:
                    continue

            # 3. 关键词文本模糊筛选（包含检验方法搜索）
            if filter_text:
                if (filter_lower not in food_cat.lower() and 
                    filter_lower not in test_item.lower() and 
                    filter_lower not in method_val.lower() and 
                    filter_lower not in notes.lower() and 
                    filter_lower not in std_code.lower()):
                    continue
            
            # 4. 表头点击交互式精确筛选
            col_match = True
            for col_idx, col_name in enumerate(self.batch_columns):
                filter_vals = self.batch_col_filters.get(col_name)
                if filter_vals is not None:
                    if str(raw_row[col_idx]) not in filter_vals:
                        col_match = False
                        break
                        
            if not col_match:
                continue
            
            all_highlight_kws = list(matched_kws)
            if filter_text and filter_text not in all_highlight_kws:
                all_highlight_kws.append(filter_text)
                
            disp_row = (
                raw_row[0],
                highlight_keywords_text(raw_row[1], all_highlight_kws),
                highlight_keywords_text(raw_row[2], all_highlight_kws),
                raw_row[3],
                raw_row[4],
                raw_row[5],
                highlight_keywords_text(method_val, [filter_text]) if filter_text else method_val,
                highlight_keywords_text(notes, [filter_text]) if filter_text else notes
            )
                
            self.batch_tree.insert('', 'end', values=disp_row, tags=(tag_name,))
            display_count += 1
            
        summary_text = f"共查询到 {len(self._batch_data_cache)} 条结果，当前展示 {display_count} 条"
        if getattr(self, '_batch_has_measured', False):
            p_cnt = getattr(self, '_batch_pass_count', 0)
            f_cnt = getattr(self, '_batch_fail_count', 0)
            summary_text += f"  |  📊 判定结果: ✔ 合格 {p_cnt} 项，✘ 超标 {f_cnt} 项"
        self.batch_count_label.config(text=summary_text)

    def do_batch_query(self):
        """按用户在文本框填写的项目顺序逐个检索并呈现结果，没有判定的以 / 显示"""
        food_kw = self.batch_food_var.get().strip()
        food_kw_lower = food_kw.lower()
        items_text = self.batch_items_text.get("1.0", tk.END).strip()
        
        if not items_text:
            messagebox.showwarning("提示", "请输入至少一个检测项目！")
            return
            
        # 解析每行：支持纯项目名、项目+实测值，以及单行用顿号/逗号拼接的多项目列表
        parsed_queries = []
        for line in items_text.split('\n'):
            line_str = line.strip()
            if not line_str:
                continue
            # 若本行包含多个顿号/分号且无纯数字判定值，自动按多项目展开
            if ('、' in line_str or '；' in line_str) and not re.search(r'[:=：]\s*\d+', line_str):
                sub_items = split_cell_test_items(line_str)
                for s_it in sub_items:
                    parsed_queries.append((s_it, s_it.lower(), None))
            else:
                item_name, meas_val = parse_batch_input_line(line_str)
                if item_name:
                    parsed_queries.append((item_name, item_name.lower(), meas_val))
                
        if not parsed_queries:
            messagebox.showwarning("提示", "未能解析到有效的检测项目！")
            return
        
        self._batch_data_cache = []
        pass_count = 0
        fail_count = 0
        has_any_measured = False
        
        # 核心改进：以用户填入的检测项目列表作为最外层遍历，严格按用户输入顺序输出
        for q_name, q_lower, q_meas in parsed_queries:
            item_matched_rows = []
            
            for std_code in sorted(self.standards.keys()):
                items = self.standards[std_code]['items']
                for item in items:
                    food_cat = item.get('food_category', '')
                    test_item = item.get('test_item', '')
                    food_cat_lower = food_cat.lower()
                    test_item_lower = test_item.lower()
                    
                    # 食品名称过滤
                    if food_kw and food_kw_lower not in food_cat_lower:
                        continue
                        
                    # 项目名称匹配
                    if q_lower in test_item_lower or test_item_lower in q_lower:
                        std_tag = std_code.replace(" ", "_").replace("-", "_")
                        if std_tag not in self.standard_colors:
                            std_tag = ""
                        
                        limit_raw = item.get('limit_value', '')
                        unit_str = item.get('unit', '')
                        disp_limit = limit_raw + (" " + unit_str if unit_str else "")
                        
                        # 判定逻辑与 / 符号表示
                        if q_meas is not None:
                            has_any_measured = True
                            j_status, j_text = judge_measured_value(limit_raw, q_meas)
                            if j_status == 'pass':
                                tag_name = 'judge_pass'
                                pass_count += 1
                            elif j_status == 'fail':
                                tag_name = 'judge_fail'
                                fail_count += 1
                            elif j_status == 'unknown':
                                tag_name = 'judge_unknown'
                            else:
                                tag_name = std_tag
                            meas_display = str(q_meas)
                            judge_display = j_text
                        else:
                            # 用户需求：没有判定的以 / 表示
                            tag_name = std_tag
                            meas_display = "/"
                            judge_display = "/"
                            
                        test_method = item.get('method', '')
                        if not test_method:
                            test_method = resolve_item_test_method('', item.get('test_item', ''), item.get('food_category', ''), std_code)
                        
                        raw_row = (
                            std_code,
                            item.get('food_category', ''),
                            item.get('test_item', ''),
                            disp_limit,
                            meas_display,
                            judge_display,
                            test_method,
                            item.get('notes', '')
                        )
                        matched_kws = [food_kw, q_name] if food_kw else [q_name]
                        item_matched_rows.append((raw_row, tag_name, matched_kws))
                        
            if item_matched_rows:
                # 核心改进：每个检测项目匹配到的条目按精准匹配优先排序
                item_matched_rows.sort(key=lambda r: calc_match_rank({'test_item': r[0][2], 'food_category': r[0][1], 'method': r[0][6]}, q_name))
                self._batch_data_cache.extend(item_matched_rows)
            else:
                # 若未查见限量指标，也按顺序保留一行记录，并自动匹配其推荐检验方法
                meas_display = str(q_meas) if q_meas is not None else "/"
                guessed_method = resolve_item_test_method('', q_name, food_kw, '')
                raw_row = (
                    "未查见对应限量",
                    food_kw if food_kw else "/",
                    q_name,
                    "/",
                    meas_display,
                    "/",
                    guessed_method if guessed_method != "/" else "/",
                    "在选定食品分类及现行6大标准中未查见对应限量指标"
                )
                self._batch_data_cache.append((raw_row, "", [q_name]))
                    
        self._batch_has_measured = has_any_measured
        self._batch_pass_count = pass_count
        self._batch_fail_count = fail_count
        self.filter_batch_results()

    

    # ============================================================
    # 模块 2: 四级分类与台账智能反写 (Excel自动匹配检测方法与判定标准)
    # ============================================================
    def setup_excel_backfill_tab(self):
        """配置四级分类与台账智能反写工作台 UI (完美对齐2026年监督抽检项目表结构)"""
        self.default_backfill_widths = {
            "序号": 48,
            "食品大类(一级)": 115,
            "食品亚类(二级)": 115,
            "食品品种(三级)": 115,
            "食品细类(四级)/样品": 135,
            "检验项目": 140,
            "依据法律法规或标准": 145,
            "检测方法": 240,
            "标准限量值": 105,
            "检出限(LOD)": 155,
            "实测值": 70,
            "判定结论": 85,
            "匹配状态": 130
        }

        self.backfill_file_var = tk.StringVar()
        self.backfill_sheet_var = tk.StringVar()
        self.backfill_cat1_col_var = tk.StringVar(value="(自动/无)")
        self.backfill_cat2_col_var = tk.StringVar(value="(自动/无)")
        self.backfill_cat3_col_var = tk.StringVar(value="(自动/无)")
        self.backfill_food_col_var = tk.StringVar()
        self.backfill_item_col_var = tk.StringVar()
        self.backfill_meas_col_var = tk.StringVar(value="(无/不导入)")

        # 2026细则四级分类速查变量
        self.quick_major_var = tk.StringVar()
        self.quick_fine_var = tk.StringVar()

        self._backfill_raw_headers = []
        self._backfill_raw_rows = []
        self._backfill_processed_rows = []
        self._backfill_current_wb = None

        main_frame = tb.Frame(self.tab_excel_backfill, padding=6)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. 顶部导入与四级分类字段映射卡片
        ctl_card = tb.Labelframe(
            main_frame,
            text=" 📂 四级分类抽检计划表 / Excel委托单导入 与 智能字段映射 (支持合并单元格自动填充、标题行自动识别) ",
            padding=10
        )
        ctl_card.pack(fill=tk.X, pady=(0, 6))

        # 第一行：文件选择 + 标准四级分类模板下载 + 2026细则四级分类速查载入
        row1 = tb.Frame(ctl_card)
        row1.pack(fill=tk.X, pady=(0, 6))
        tb.Label(row1, text="Excel/CSV文件:", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.backfill_file_entry = tb.Entry(row1, textvariable=self.backfill_file_var, font=("Microsoft YaHei", 9), state="readonly", width=34)
        self.backfill_file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        tb.Button(row1, text="📂 选择抽检表/台账...", command=self.browse_backfill_file, bootstyle=PRIMARY).pack(side=tk.LEFT, padx=(0, 6))
        tb.Button(row1, text="📄 下载四级分类标准导入模板 (.xlsx)", command=self.download_backfill_template, bootstyle="info-outline").pack(side=tk.LEFT, padx=(0, 12))

        # 右侧：直接按2026抽检细则四级分类一键调出项目表
        tb.Label(row1, text="| 📌 2026细则速查:", font=("Microsoft YaHei", 9, "bold"), bootstyle=PRIMARY).pack(side=tk.LEFT, padx=(4, 4))
        majors = sorted(list({v.get("major_category", "") for v in SAMPLING_RULES_FULL_DB.values() if v.get("major_category")}))
        self.quick_major_combo = ttk.Combobox(row1, textvariable=self.quick_major_var, values=majors, state="readonly", width=14, font=("Microsoft YaHei", 9))
        self.quick_major_combo.pack(side=tk.LEFT, padx=(0, 4))
        self.quick_major_combo.bind("<<ComboboxSelected>>", self.on_quick_major_changed)

        self.quick_fine_combo = ttk.Combobox(row1, textvariable=self.quick_fine_var, state="readonly", width=16, font=("Microsoft YaHei", 9))
        self.quick_fine_combo.pack(side=tk.LEFT, padx=(0, 4))
        if majors:
            self.quick_major_var.set(majors[0])
            self.on_quick_major_changed()

        tb.Button(row1, text="📥 载入该细类抽检项目", command=self.load_from_2026_rules_selector, bootstyle="warning").pack(side=tk.LEFT)

        # 第二行：工作表与四级分类列映射
        row2 = tb.Frame(ctl_card)
        row2.pack(fill=tk.X, pady=(2, 0))

        tb.Label(row2, text="工作表:", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT, padx=(0, 2))
        self.backfill_sheet_combo = ttk.Combobox(row2, textvariable=self.backfill_sheet_var, state="readonly", width=11, font=("Microsoft YaHei", 9))
        self.backfill_sheet_combo.pack(side=tk.LEFT, padx=(0, 8))
        self.backfill_sheet_combo.bind("<<ComboboxSelected>>", self.on_backfill_sheet_changed)

        tb.Label(row2, text="一级(大类):", font=("Microsoft YaHei", 8)).pack(side=tk.LEFT, padx=(0, 2))
        self.backfill_cat1_combo = ttk.Combobox(row2, textvariable=self.backfill_cat1_col_var, state="readonly", width=11, font=("Microsoft YaHei", 8))
        self.backfill_cat1_combo.pack(side=tk.LEFT, padx=(0, 6))

        tb.Label(row2, text="二级(亚类):", font=("Microsoft YaHei", 8)).pack(side=tk.LEFT, padx=(0, 2))
        self.backfill_cat2_combo = ttk.Combobox(row2, textvariable=self.backfill_cat2_col_var, state="readonly", width=11, font=("Microsoft YaHei", 8))
        self.backfill_cat2_combo.pack(side=tk.LEFT, padx=(0, 6))

        tb.Label(row2, text="三级(品种):", font=("Microsoft YaHei", 8)).pack(side=tk.LEFT, padx=(0, 2))
        self.backfill_cat3_combo = ttk.Combobox(row2, textvariable=self.backfill_cat3_col_var, state="readonly", width=11, font=("Microsoft YaHei", 8))
        self.backfill_cat3_combo.pack(side=tk.LEFT, padx=(0, 6))

        tb.Label(row2, text="四级(细类)/样品*:", font=("Microsoft YaHei", 9, "bold"), bootstyle=PRIMARY).pack(side=tk.LEFT, padx=(0, 2))
        self.backfill_food_combo = ttk.Combobox(row2, textvariable=self.backfill_food_col_var, state="readonly", width=13, font=("Microsoft YaHei", 9))
        self.backfill_food_combo.pack(side=tk.LEFT, padx=(0, 8))

        tb.Label(row2, text="检验项目列*:", font=("Microsoft YaHei", 9, "bold"), bootstyle=PRIMARY).pack(side=tk.LEFT, padx=(0, 2))
        self.backfill_item_combo = ttk.Combobox(row2, textvariable=self.backfill_item_col_var, state="readonly", width=12, font=("Microsoft YaHei", 9))
        self.backfill_item_combo.pack(side=tk.LEFT, padx=(0, 8))

        tb.Label(row2, text="实测值(选填):", font=("Microsoft YaHei", 8)).pack(side=tk.LEFT, padx=(0, 2))
        self.backfill_meas_combo = ttk.Combobox(row2, textvariable=self.backfill_meas_col_var, state="readonly", width=10, font=("Microsoft YaHei", 8))
        self.backfill_meas_combo.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_run_backfill = tb.Button(row2, text="⚡ 开始四级分类智能匹配反写", command=self.run_excel_backfill, bootstyle=SUCCESS, cursor="hand2")
        self.btn_run_backfill.pack(side=tk.RIGHT)

        # 第三行：单单元格多项目自动拆分展开选项 (针对《必检品种项目表》等“、/，/换行”多项目合一的场景)
        row3 = tb.Frame(ctl_card)
        row3.pack(fill=tk.X, pady=(4, 0))

        self.backfill_auto_split_var = tk.BooleanVar(value=True)
        tb.Checkbutton(
            row3,
            text="☑️ 单元格多项目自动拆行展开 (针对“必检项目/可选项目”中顿号、逗号、分号或换行合并在同一格的抽检表，自动拆分并继承四级分类)",
            variable=self.backfill_auto_split_var,
            bootstyle="success-roundtoggle"
        ).pack(side=tk.LEFT, padx=(0, 15))

        tb.Label(
            row3,
            text="💡 提示：系统采用深度括号保护状态机，如【合成着色剂（柠檬黄、日落黄...）】等复合指标不会被误拆，精准炸裂为多行检测任务！",
            font=("Microsoft YaHei", 8),
            bootstyle=SECONDARY
        ).pack(side=tk.LEFT)

        # 2. 质控统计看板
        stats_frame = tb.Frame(main_frame)
        stats_frame.pack(fill=tk.X, pady=(0, 4))
        self.backfill_stats_label = tb.Label(
            stats_frame,
            text="📊 质控看板: 支持导入含【食品大类(一级)~食品细类(四级)+检验项目】或【样品名称+检测项目】的Excel表，自动反写【依据法律法规或标准】、【检测方法】、【标准限量值】与【检出限】",
            font=("Microsoft YaHei", 9, "bold"),
            bootstyle=INFO
        )
        self.backfill_stats_label.pack(side=tk.LEFT)

        # 3. 结果预览表格
        grid_card = tb.Labelframe(
            main_frame,
            text=" 📋 四级分类与台账智能反写实时预览 (双击行查看完整指标卡片 | 鼠标悬停气泡展示长文本 | 右键可复制或自定义列) ",
            padding=6
        )
        grid_card.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        columns = (
            "序号", "食品大类(一级)", "食品亚类(二级)", "食品品种(三级)", "食品细类(四级)/样品",
            "检验项目", "依据法律法规或标准", "检测方法", "标准限量值", "检出限(LOD)", "实测值", "判定结论", "匹配状态"
        )
        self.backfill_tree = ttk.Treeview(grid_card, columns=columns, show='headings', selectmode='browse')

        headings_map = {
            "序号": "序号",
            "食品大类(一级)": "食品大类(一级)",
            "食品亚类(二级)": "食品亚类(二级)",
            "食品品种(三级)": "食品品种(三级)",
            "食品细类(四级)/样品": "食品细类(四级)/样品",
            "检验项目": "检验项目",
            "依据法律法规或标准": "【反写】依据法律法规或标准",
            "检测方法": "【反写】检测方法",
            "标准限量值": "【反写】标准限量值",
            "检出限(LOD)": "【反写】检出限(LOD)",
            "实测值": "实测值",
            "判定结论": "判定结论",
            "匹配状态": "匹配状态"
        }
        for col in columns:
            self.backfill_tree.heading(col, text=headings_map.get(col, col))
            anchor = "center" if col in ("序号", "标准限量值", "实测值", "判定结论", "匹配状态") else "w"
            self.backfill_tree.column(col, width=self.default_backfill_widths.get(col, 120), anchor=anchor)

        self.apply_tree_column_widths(self.backfill_tree, "backfill_tree", self.default_backfill_widths)
        self.setup_tree_context_menu(self.backfill_tree, "backfill_tree", self.default_backfill_widths, self.on_backfill_tree_double_click)

        self.backfill_tree.tag_configure("tag_pass", background="#ECFDF5", foreground="#065F46")
        self.backfill_tree.tag_configure("tag_fail", background="#FEF2F2", foreground="#991B1B")
        self.backfill_tree.tag_configure("tag_infer", background="#FFFBEB", foreground="#92400E")
        self.backfill_tree.tag_configure("tag_miss", background="#F1F5F9", foreground="#64748B")

        v_scroll = ttk.Scrollbar(grid_card, orient=tk.VERTICAL, command=self.backfill_tree.yview)
        h_scroll = ttk.Scrollbar(grid_card, orient=tk.HORIZONTAL, command=self.backfill_tree.xview)
        self.backfill_tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        self.backfill_tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        grid_card.grid_rowconfigure(0, weight=1)
        grid_card.grid_columnconfigure(0, weight=1)

        # 4. 底部导出与操作区
        bot_bar = tb.Frame(main_frame)
        bot_bar.pack(fill=tk.X)
        tb.Button(bot_bar, text="💾 导出已反写Excel表格 (.xlsx)", command=self.export_backfilled_excel, bootstyle=PRIMARY, cursor="hand2").pack(side=tk.LEFT, padx=(0, 8))
        tb.Button(bot_bar, text="📋 复制全表到剪贴板 (可直接粘贴Excel)", command=self.copy_backfill_preview, bootstyle="secondary-outline").pack(side=tk.LEFT, padx=(0, 8))
        tb.Button(bot_bar, text="↔ 自适应列宽", command=lambda: self.autofit_tree_columns(self.backfill_tree), bootstyle="info-outline").pack(side=tk.LEFT, padx=(0, 8))
        tb.Button(bot_bar, text="🔄 清空列表", command=self.clear_backfill_data, bootstyle="danger-outline").pack(side=tk.LEFT)

    def on_quick_major_changed(self, event=None):
        """2026抽检细则一级大类切换时联动更新四级细类下拉框"""
        maj = self.quick_major_var.get().strip()
        fines = []
        for k, v in SAMPLING_RULES_FULL_DB.items():
            if v.get("major_category", "").strip() == maj:
                fn = v.get("fine_category") or v.get("product_name") or k
                if fn not in fines:
                    fines.append(fn)
        self.quick_fine_combo['values'] = fines
        if fines:
            self.quick_fine_var.set(fines[0])
        else:
            self.quick_fine_var.set("")

    def load_from_2026_rules_selector(self):
        """直接从2026抽检细则选择器载入指定四级细类的全套检验项目并自动匹配反写"""
        maj = self.quick_major_var.get().strip()
        fin = self.quick_fine_var.get().strip()
        if not maj or not fin:
            messagebox.showinfo("提示", "请先选择【一级大类】与【四级细类】！")
            return

        target_entry = None
        for k, v in SAMPLING_RULES_FULL_DB.items():
            if v.get("major_category", "").strip() == maj and (v.get("fine_category") == fin or v.get("product_name") == fin):
                target_entry = v
                break
        if not target_entry:
            messagebox.showwarning("提示", f"未找到【{maj} - {fin}】的细则数据！")
            return

        c1 = target_entry.get("major_category", maj)
        c2 = target_entry.get("sub_category", fin)
        c3 = target_entry.get("variety_category", fin)
        c4 = target_entry.get("fine_category", fin)
        items = target_entry.get("items", [])

        self._backfill_raw_headers = ["序号", "食品大类(一级)", "食品亚类(二级)", "食品品种(三级)", "食品细类(四级)", "检验项目", "依据法律法规或标准", "检测方法", "备注"]
        self._backfill_raw_rows = []
        for idx, it in enumerate(items, 1):
            self._backfill_raw_rows.append([
                idx, c1, c2, c3, c4,
                it.get("item", ""),
                it.get("std", ""),
                it.get("method", ""),
                it.get("remark", "")
            ])
        self.auto_detect_backfill_columns()
        self.run_excel_backfill(silent_success=True)

    def browse_backfill_file(self):
        """选择Excel或CSV抽检项目表/委托单文件"""
        fn = filedialog.askopenfilename(
            title="选择四级分类抽检项目表或委托单文件",
            filetypes=[("Excel工作簿 (*.xlsx;*.xls)", "*.xlsx;*.xls"), ("CSV表格 (*.csv)", "*.csv"), ("所有文件 (*.*)", "*.*")]
        )
        if not fn:
            return
        self.backfill_file_var.set(fn)
        self.load_backfill_file_structure(fn)

    def load_backfill_file_structure(self, filepath):
        """加载文件表头并智能识别四级分类列与检验项目列（兼容WPS格式与合并单元格）"""
        ext = os.path.splitext(filepath)[1].lower()
        self._backfill_raw_headers = []
        self._backfill_raw_rows = []

        try:
            if ext in ('.xlsx', '.xls'):
                if openpyxl is None:
                    messagebox.showerror("错误", "当前运行环境缺少 openpyxl 库，无法直接读取 Excel 文件！")
                    return
                wb = openpyxl.load_workbook(filepath, data_only=True)
                self._backfill_current_wb = wb
                sheet_names = wb.sheetnames
                self.backfill_sheet_combo['values'] = sheet_names
                if sheet_names:
                    self.backfill_sheet_combo.set(sheet_names[0])
                    self.read_active_sheet_data(wb[sheet_names[0]])
            elif ext == '.csv':
                self._backfill_current_wb = None
                self.backfill_sheet_combo['values'] = ["CSV数据表"]
                self.backfill_sheet_combo.set("CSV数据表")
                encodings = ['utf-8-sig', 'utf-8', 'gbk', 'gb18030']
                lines_data = None
                for enc in encodings:
                    try:
                        with open(filepath, 'r', encoding=enc) as f:
                            reader = csv.reader(f)
                            all_rows = list(reader)
                            if all_rows:
                                lines_data = all_rows
                                break
                    except Exception:
                        continue
                if not lines_data:
                    messagebox.showerror("错误", "无法以 UTF-8 或 GBK 编码解析该 CSV 文件！")
                    return
                hdr_idx = self._find_best_header_row_index(lines_data)
                self._backfill_raw_headers = [re.sub(r'\s+', '', str(c)).strip() if c is not None and str(c).strip() else f"列{i+1}" for i, c in enumerate(lines_data[hdr_idx])]
                self._backfill_raw_rows = lines_data[hdr_idx+1:]
                self.auto_detect_backfill_columns()
        except Exception as e:
            messagebox.showerror("文件加载失败", f"无法读取该文件：\n{e}")

    def _find_best_header_row_index(self, rows):
        """智能扫描前10行，跳过《2026年...抽检项目表》等横幅标题行，精准定位真正的表头行"""
        header_keywords = ["检验项目", "检测项目", "食品细类", "食品大类", "食品亚类", "食品品种", "样品名称", "依据法律法规", "检测方法", "产品名称"]
        best_idx = 0
        best_hits = -1
        for idx in range(min(10, len(rows))):
            r = rows[idx]
            non_empty = [str(c).strip() for c in r if c is not None and str(c).strip()]
            if not non_empty:
                continue
            hits = sum(1 for cell_str in non_empty if any(kw in cell_str for kw in header_keywords))
            if hits > best_hits or (hits == best_hits and hits > 0 and len(non_empty) > 2):
                best_hits = hits
                best_idx = idx
        if best_hits <= 0:
            for idx, r in enumerate(rows):
                if sum(1 for c in r if c is not None and str(c).strip()) >= 2:
                    return idx
        return best_idx

    def on_backfill_sheet_changed(self, event=None):
        """切换 Excel Sheet"""
        if not self._backfill_current_wb:
            return
        s_name = self.backfill_sheet_var.get()
        if s_name in self._backfill_current_wb.sheetnames:
            self.read_active_sheet_data(self._backfill_current_wb[s_name])

    def read_active_sheet_data(self, sheet):
        """读取选定工作表（自动拆分并向下填充合并单元格、跳过横幅标题行）"""
        # 1. 自动拆分合并单元格并填充值（解决一级~四级分类纵向合并导致后续行为空的问题）
        try:
            if sheet.merged_cells and sheet.merged_cells.ranges:
                for rng in list(sheet.merged_cells.ranges):
                    top_val = sheet.cell(rng.min_row, rng.min_col).value
                    sheet.unmerge_cells(str(rng))
                    for r in range(rng.min_row, rng.max_row + 1):
                        for c in range(rng.min_col, rng.max_col + 1):
                            sheet.cell(r, c).value = top_val
        except Exception:
            pass

        rows = [list(r) for r in sheet.iter_rows(values_only=True)]
        if not rows:
            messagebox.showwarning("提示", "该工作表为空！")
            return

        hdr_idx = self._find_best_header_row_index(rows)
        raw_hdr = rows[hdr_idx]
        clean_headers = []
        for i, c in enumerate(raw_hdr):
            h_str = re.sub(r'\s+', '', str(c)).strip() if (c is not None and str(c).strip()) else f"列{i+1}"
            if h_str in clean_headers:
                h_str = f"{h_str}_{i+1}"
            clean_headers.append(h_str)

        self._backfill_raw_headers = clean_headers
        self._backfill_raw_rows = rows[hdr_idx+1:]
        self.auto_detect_backfill_columns()

    def auto_detect_backfill_columns(self):
        """智能识别并绑定：一级(大类)、二级(亚类)、三级(品种)、四级(细类)/样品名、检验项目、实测值"""
        hdrs = [h for h in self._backfill_raw_headers if h]
        if not hdrs:
            return

        opt_hdrs = ["(自动/无)"] + hdrs
        self.backfill_cat1_combo['values'] = opt_hdrs
        self.backfill_cat2_combo['values'] = opt_hdrs
        self.backfill_cat3_combo['values'] = opt_hdrs
        self.backfill_food_combo['values'] = hdrs
        self.backfill_item_combo['values'] = hdrs
        self.backfill_meas_combo['values'] = ["(无/不导入)"] + hdrs

        # 1. 四级分类列推导
        c1_col = "(自动/无)"
        c2_col = "(自动/无)"
        c3_col = "(自动/无)"
        c4_col = ""

        for h in hdrs:
            if any(k in h for k in ["大类", "一级"]):
                c1_col = h
            elif any(k in h for k in ["亚类", "二级"]):
                c2_col = h
            elif any(k in h for k in ["品种", "三级"]):
                c3_col = h
            elif any(k in h for k in ["细类", "四级"]):
                c4_col = h

        if not c4_col:
            for h in hdrs:
                if any(k in h for k in ["样品名称", "食品名称", "产品名称", "样品", "品名", "种类", "类别", "食品", "标的物"]):
                    if h not in (c1_col, c2_col, c3_col):
                        c4_col = h
                        break
        if not c4_col:
            c4_col = c3_col if c3_col != "(自动/无)" else (c2_col if c2_col != "(自动/无)" else (c1_col if c1_col != "(自动/无)" else hdrs[0]))

        self.backfill_cat1_col_var.set(c1_col)
        self.backfill_cat2_col_var.set(c2_col)
        self.backfill_cat3_col_var.set(c3_col)
        self.backfill_food_col_var.set(c4_col)

        # 2. 检验项目列推导
        item_col = ""
        for h in hdrs:
            if any(k in h for k in ["必检项目", "检验项目", "检测项目", "抽检项目", "不合格项目", "必检", "项目名称", "可选项目", "项目", "参数", "指标"]):
                if h not in (c1_col, c2_col, c3_col, c4_col):
                    item_col = h
                    break
        if not item_col and len(hdrs) > 1:
            for h in hdrs:
                if h not in (c1_col, c2_col, c3_col, c4_col, "序号"):
                    item_col = h
                    break
            if not item_col:
                item_col = hdrs[1]
        self.backfill_item_col_var.set(item_col)

        # 3. 实测值列推导
        meas_col = "(无/不导入)"
        for h in hdrs:
            if any(k in h for k in ["实测值", "检验结果", "检测结果", "测定值", "实测结果"]):
                meas_col = h
                break
        self.backfill_meas_col_var.set(meas_col)

        count = len(self._backfill_raw_rows)
        mode_desc = "四级分类抽检表模式" if (c1_col != "(自动/无)" or c2_col != "(自动/无)" or "细类" in c4_col) else "常规样品台账模式"
        self.backfill_stats_label.config(
            text=f"📊 文件解析就绪 [{mode_desc}]: 共载入 {count} 行 | 已自动识别【{c4_col}】与【{item_col}】，点击右侧【开始四级分类智能匹配反写】即可！",
            bootstyle=PRIMARY
        )

    def download_backfill_template(self):
        """生成并下载完全贴合实际业务的《四级分类监督抽检项目表》标准导入模板 (.xlsx)"""
        save_path = filedialog.asksaveasfilename(
            title="保存四级分类标准导入模板",
            defaultextension=".xlsx",
            initialfile="2026年食品安全监督抽检项目表_四级分类智能反写标准模板.xlsx",
            filetypes=[("Excel工作簿 (*.xlsx)", "*.xlsx")]
        )
        if not save_path:
            return

        try:
            if openpyxl is None:
                messagebox.showerror("错误", "缺少 openpyxl 库！")
                return
            wb = openpyxl.Workbook()

            # Sheet 1: 四级分类监督抽检项目表模板（与用户截图3完全一致）
            ws1 = wb.active
            ws1.title = "四级分类抽检项目反写模板"

            headers1 = [
                "序号", "食品大类（一级）", "食品亚类（二级）", "食品品种（三级）", "食品细类（四级）",
                "检验项目", "依据法律法规或标准", "检测方法", "标准限量值", "方法检出限(LOD)", "实测值(可选)", "备注"
            ]
            ws1.append(headers1)

            samples1 = [
                (1, "粮食加工品", "小麦粉", "小麦粉", "小麦粉", "镉（以Cd计）", "", "", "", "", "", "导入后自动反写标准与方法"),
                (2, "粮食加工品", "小麦粉", "小麦粉", "小麦粉", "铅（以Pb计）", "", "", "", "", "", ""),
                (3, "粮食加工品", "小麦粉", "小麦粉", "小麦粉", "苯并[a]芘", "", "", "", "", "", ""),
                (4, "粮食加工品", "小麦粉", "小麦粉", "小麦粉", "玉米赤霉烯酮", "", "", "", "", "", ""),
                (5, "粮食加工品", "小麦粉", "小麦粉", "小麦粉", "脱氧雪腐镰刀菌烯醇", "", "", "", "", "", ""),
                (6, "粮食加工品", "小麦粉", "小麦粉", "小麦粉", "赭曲霉毒素A", "", "", "", "", "", ""),
                (7, "粮食加工品", "小麦粉", "小麦粉", "小麦粉", "黄曲霉毒素B1", "", "", "", "", "", ""),
                (8, "粮食加工品", "小麦粉", "小麦粉", "小麦粉", "偶氮甲酰胺", "", "", "", "", "", ""),
                (9, "粮食加工品", "小麦粉", "小麦粉", "小麦粉", "过氧化苯甲酰", "", "", "", "", "", ""),
                (10, "粮食加工品", "大米", "大米", "大米", "铅（以Pb计）", "", "", "", "", "0.05", ""),
                (11, "粮食加工品", "大米", "大米", "大米", "镉（以Cd计）", "", "", "", "", "0.02", ""),
                (12, "粮食加工品", "大米", "大米", "大米", "苯并[a]芘", "", "", "", "", "", ""),
                (13, "粮食加工品", "大米", "大米", "大米", "无机砷（以As计）", "", "", "", "", "", ""),
                (14, "粮食加工品", "大米", "大米", "大米", "赭曲霉毒素A", "", "", "", "", "", ""),
                (15, "粮食加工品", "大米", "大米", "大米", "黄曲霉毒素B1", "", "", "", "", "", ""),
                (16, "粮食加工品", "挂面", "挂面", "挂面", "铅（以Pb计）", "", "", "", "", "", ""),
                (17, "粮食加工品", "挂面", "挂面", "挂面", "黄曲霉毒素B1", "", "", "", "", "", ""),
                (18, "粮食加工品", "挂面", "挂面", "挂面", "合成着色剂（柠檬黄、日落黄、胭脂红、苋菜红、诱惑红、亮蓝）", "", "", "", "", "", ""),
                (19, "粮食加工品", "挂面", "挂面", "挂面", "脱氢乙酸及其钠盐（以脱氢乙酸计）", "", "", "", "", "", ""),
                (20, "食用农产品", "蔬菜", "叶菜类蔬菜", "普通白菜(小白菜)", "啶虫脒", "", "", "", "", "0.12", ""),
                (21, "食用农产品", "蔬菜", "鳞茎类蔬菜", "韭菜", "腐霉利", "", "", "", "", "0.05", ""),
                (22, "食用农产品", "畜禽肉及副产品", "畜肉", "猪肉", "恩诺沙星", "", "", "", "", "15", "")
            ]
            for s in samples1:
                ws1.append(list(s))

            header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
            header_font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
            thin_border = Border(
                left=Side(style='thin', color='CBD5E1'),
                right=Side(style='thin', color='CBD5E1'),
                top=Side(style='thin', color='CBD5E1'),
                bottom=Side(style='thin', color='CBD5E1')
            )

            for col in range(1, len(headers1) + 1):
                cell = ws1.cell(row=1, column=col)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for row in range(2, len(samples1) + 2):
                for col in range(1, len(headers1) + 1):
                    c = ws1.cell(row=row, column=col)
                    c.border = thin_border
                    c.alignment = Alignment(horizontal="center" if col in (1, 7, 9, 11) else "left", vertical="center")

            col_widths = [8, 18, 16, 16, 20, 32, 22, 22, 16, 20, 14, 22]
            for i, w in enumerate(col_widths, 1):
                ws1.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

            wb.save(save_path)
            messagebox.showinfo(
                "模板生成成功",
                f"《四级分类监督抽检项目表》标准导入模板已生成！\n\n保存路径：{save_path}\n\n"
                f"✔ 支持填写【食品大类(一级) ~ 食品细类(四级) + 检验项目】\n"
                f"✔ 支持单元格纵向合并（导入时自动拆分向下填充）\n"
                f"✔ 导入后一键反写【依据法律法规或标准】、【检测方法】、【标准限量值】与【检出限(LOD)】！"
            )
        except Exception as e:
            messagebox.showerror("生成失败", f"无法创建模板文件：\n{e}")

    def run_excel_backfill(self, silent_success=False):
        """执行四级分类 + 检验项目全量智能匹配与信息反写"""
        if not self._backfill_raw_rows:
            messagebox.showwarning("提示", "请先选择抽检表/台账文件，或使用上方【2026细则速查】载入项目！")
            return

        food_col_name = self.backfill_food_col_var.get()
        item_col_name = self.backfill_item_col_var.get()
        cat1_col_name = self.backfill_cat1_col_var.get()
        cat2_col_name = self.backfill_cat2_col_var.get()
        cat3_col_name = self.backfill_cat3_col_var.get()
        meas_col_name = self.backfill_meas_col_var.get()

        if not food_col_name or not item_col_name:
            messagebox.showwarning("提示", "请先选择【四级(细类)/样品列】与【检验项目列】！")
            return

        def _get_col_idx(c_name):
            if c_name and c_name not in ("(自动/无)", "(无/不导入)") and c_name in self._backfill_raw_headers:
                return self._backfill_raw_headers.index(c_name)
            return -1

        food_idx = _get_col_idx(food_col_name)
        item_idx = _get_col_idx(item_col_name)
        c1_idx = _get_col_idx(cat1_col_name)
        c2_idx = _get_col_idx(cat2_col_name)
        c3_idx = _get_col_idx(cat3_col_name)
        meas_idx = _get_col_idx(meas_col_name)

        if food_idx < 0 or item_idx < 0:
            messagebox.showerror("错误", "所选【四级(细类)/样品列】或【检验项目列】在表头中未找到！")
            return

        self.backfill_tree.delete(*self.backfill_tree.get_children())
        self._backfill_processed_rows = []

        exact_hits = 0
        infer_hits = 0
        miss_hits = 0

        # 支持合并单元格/省略填写行的向下自动延续填充 (Forward-Fill)
        last_c1 = ""
        last_c2 = ""
        last_c3 = ""
        last_f = ""

        seq_counter = 0
        for row in self._backfill_raw_rows:
            def _cell_str(idx_col):
                if 0 <= idx_col < len(row) and row[idx_col] is not None:
                    s = re.sub(r'[\r\n]+', '', str(row[idx_col])).strip()
                    return "" if s == "None" else s
                return ""

            raw_i = _cell_str(item_idx)
            raw_c1 = _cell_str(c1_idx)
            raw_c2 = _cell_str(c2_idx)
            raw_c3 = _cell_str(c3_idx)
            raw_f = _cell_str(food_idx)
            m_val = _cell_str(meas_idx) if meas_idx >= 0 else None

            # 跳过纯空行或重复表头行
            if not raw_i and not raw_f:
                continue
            if raw_i in ("检验项目", "检测项目", "抽检项目"):
                continue

            # 若当前行有检验项目，自动向下继承上方合并格的四级分类
            if raw_c1: last_c1 = raw_c1
            if raw_c2: last_c2 = raw_c2
            if raw_c3: last_c3 = raw_c3
            if raw_f: last_f = raw_f

            c1_val = raw_c1 or last_c1
            c2_val = raw_c2 or last_c2
            c3_val = raw_c3 or last_c3
            f_val = raw_f or last_f or c3_val or c2_val or c1_val

            if not raw_i:
                continue

            # 智能拆分多检测项目单元格（如 '恩诺沙星、氟苯尼考、磺胺类（总量）'）
            if getattr(self, 'backfill_auto_split_var', None) and self.backfill_auto_split_var.get():
                items_to_process = split_cell_test_items(raw_i)
                if not items_to_process:
                    items_to_process = [raw_i]
            else:
                items_to_process = [raw_i]

            for sub_item in items_to_process:
                seq_counter += 1
                res = auto_match_record(
                    self.standards, self.global_index,
                    f_val, sub_item, m_val,
                    cat1=c1_val, cat2=c2_val, cat3=c3_val, cat4=f_val
                )
                std_code, std_name, limit_str, method_str, lod_str, meas_display, judge_display, status_tag, matched_cat = res

                if len(items_to_process) > 1 and "精准" in status_tag:
                    status_tag = status_tag.replace("精准匹配", "拆分精准匹配")

                if "精准" in status_tag or "2026细则" in status_tag or "禁用" in status_tag:
                    exact_hits += 1
                    row_tag = "tag_fail" if "不合格" in judge_display else "tag_pass"
                elif "推导" in status_tag:
                    infer_hits += 1
                    row_tag = "tag_fail" if "不合格" in judge_display else "tag_infer"
                else:
                    miss_hits += 1
                    row_tag = "tag_miss"

                tree_vals = (
                    seq_counter,
                    c1_val if c1_val else "/",
                    c2_val if c2_val else "/",
                    c3_val if c3_val else "/",
                    f_val,
                    sub_item,
                    std_code,
                    method_str,
                    limit_str,
                    lod_str,
                    meas_display,
                    judge_display,
                    status_tag
                )
                self.backfill_tree.insert('', 'end', values=tree_vals, tags=(row_tag,))

                self._backfill_processed_rows.append({
                    'raw_row': list(row),
                    'seq': seq_counter,
                    'cat1': c1_val,
                    'cat2': c2_val,
                    'cat3': c3_val,
                    'food': f_val,
                    'item': sub_item,
                    'std_code': std_code,
                    'std_name': std_name,
                    'limit': limit_str,
                    'method': method_str,
                    'lod': lod_str,
                    'meas': meas_display,
                    'judge': judge_display,
                    'status': status_tag,
                    'category': matched_cat
                })

        valid_total = exact_hits + infer_hits + miss_hits
        pct = ((exact_hits + infer_hits) / valid_total * 100) if valid_total > 0 else 0
        self.backfill_stats_label.config(
            text=f"📊 质控看板: 共匹配 {valid_total} 项 | ✔ 四级分类/国标精准匹配: {exact_hits} 项 | ⚠️ 层级推导: {infer_hits} 项 | ❓ 待核实: {miss_hits} 项 (总覆盖率 {pct:.1f}%)",
            bootstyle=SUCCESS if pct >= 85 else INFO
        )
        if not silent_success:
            messagebox.showinfo(
                "四级分类智能匹配反写完成",
                f"已融合《2026年国家食品安全监督抽检实施细则》四级分类库与 6 大国标全量库完成匹配反写！\n\n"
                f"• 总处理行数：{valid_total} 行\n"
                f"• 精准匹配：{exact_hits} 行\n"
                f"• 品类层级推导：{infer_hits} 行\n"
                f"• 综合匹配覆盖率：{pct:.1f}%\n\n"
                f"双击任意行可查看完整指标卡，点击左下角【导出已反写Excel表格】即可保存！"
            )

    def export_backfilled_excel(self):
        """导出反写完成的高保真 Excel 文件（若原表已有【依据法律法规或标准】/【检测方法】列，自动原地填入并追加完整列）"""
        if not self._backfill_processed_rows:
            messagebox.showwarning("提示", "当前没有可导出的反写结果，请先执行智能匹配！")
            return

        save_path = filedialog.asksaveasfilename(
            title="导出四级分类反写完成的Excel工作簿",
            defaultextension=".xlsx",
            initialfile=f"食品安全四级分类标准与方法反写台账_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            filetypes=[("Excel工作簿 (*.xlsx)", "*.xlsx")]
        )
        if not save_path:
            return

        try:
            if openpyxl is None:
                messagebox.showerror("错误", "缺少 openpyxl 库！")
                return
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "四级分类标准与检测方法反写表"

            orig_headers = [str(h) for h in self._backfill_raw_headers]

            # 检查原表是否已包含 依据法律法规或标准 / 检测方法 / 分类列，支持原地自动反写！
            inplace_std_idx = -1
            inplace_mth_idx = -1
            inplace_lim_idx = -1
            inplace_lod_idx = -1
            c1_idx = -1
            c2_idx = -1
            c3_idx = -1
            c4_idx = -1

            for idx, h in enumerate(orig_headers):
                if any(k in h for k in ["依据法律法规", "判定标准", "判定依据", "执行标准"]):
                    inplace_std_idx = idx
                elif any(k in h for k in ["检测方法", "检验方法", "方法标准"]):
                    inplace_mth_idx = idx
                elif "限量" in h:
                    inplace_lim_idx = idx
                elif "检出限" in h or "LOD" in h.upper():
                    inplace_lod_idx = idx
                elif any(k in h for k in ["大类", "一级"]):
                    c1_idx = idx
                elif any(k in h for k in ["亚类", "二级"]):
                    c2_idx = idx
                elif any(k in h for k in ["品种", "三级"]):
                    c3_idx = idx
                elif any(k in h for k in ["细类", "四级"]):
                    c4_idx = idx

            backfill_headers = [
                "【反写】依据法律法规或标准",
                "【反写】检测方法",
                "【反写】标准限量值",
                "【反写】方法检出限(LOD)",
                "【反写】标准全称及归属",
                "【反写】实测值",
                "【反写】判定结论",
                "【反写】匹配状态"
            ]
            full_headers = orig_headers + backfill_headers
            ws.append(full_headers)

            for item in self._backfill_processed_rows:
                orig_vals = [cell if cell is not None else "" for cell in item['raw_row']]
                while len(orig_vals) < len(orig_headers):
                    orig_vals.append("")

                # 补齐合并单元格拆开后的四级分类列及原地反写列
                if c1_idx >= 0 and not str(orig_vals[c1_idx]).strip():
                    orig_vals[c1_idx] = item['cat1']
                if c2_idx >= 0 and not str(orig_vals[c2_idx]).strip():
                    orig_vals[c2_idx] = item['cat2']
                if c3_idx >= 0 and not str(orig_vals[c3_idx]).strip():
                    orig_vals[c3_idx] = item['cat3']
                if c4_idx >= 0 and not str(orig_vals[c4_idx]).strip():
                    orig_vals[c4_idx] = item['food']

                if inplace_std_idx >= 0 and not str(orig_vals[inplace_std_idx]).strip():
                    orig_vals[inplace_std_idx] = item['std_code']
                if inplace_mth_idx >= 0 and not str(orig_vals[inplace_mth_idx]).strip():
                    orig_vals[inplace_mth_idx] = item['method']
                if inplace_lim_idx >= 0 and not str(orig_vals[inplace_lim_idx]).strip():
                    orig_vals[inplace_lim_idx] = item['limit']
                if inplace_lod_idx >= 0 and not str(orig_vals[inplace_lod_idx]).strip():
                    orig_vals[inplace_lod_idx] = item['lod']

                bf_vals = [
                    item['std_code'],
                    item['method'],
                    item['limit'],
                    item['lod'],
                    item['std_name'],
                    item['meas'],
                    item['judge'],
                    item['status']
                ]
                ws.append(orig_vals + bf_vals)

            navy_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
            tech_fill = PatternFill(start_color="0284C7", end_color="0284C7", fill_type="solid")
            hdr_font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")

            pass_fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
            pass_font = Font(name="Microsoft YaHei", size=9, bold=True, color="065F46")
            fail_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
            fail_font = Font(name="Microsoft YaHei", size=9, bold=True, color="991B1B")

            thin_border = Border(
                left=Side(style='thin', color='CBD5E1'),
                right=Side(style='thin', color='CBD5E1'),
                top=Side(style='thin', color='CBD5E1'),
                bottom=Side(style='thin', color='CBD5E1')
            )

            orig_len = len(orig_headers)
            total_cols = len(full_headers)

            for col in range(1, total_cols + 1):
                cell = ws.cell(row=1, column=col)
                cell.fill = navy_fill if col <= orig_len else tech_fill
                cell.font = hdr_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            judge_col_idx = orig_len + 7

            for r_idx in range(2, len(self._backfill_processed_rows) + 2):
                for c_idx in range(1, total_cols + 1):
                    cell = ws.cell(row=r_idx, column=c_idx)
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal="left", vertical="center")

                    if c_idx == judge_col_idx:
                        val_str = str(cell.value or "")
                        if "不合格" in val_str:
                            cell.fill = fail_fill
                            cell.font = fail_font
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                        elif "合格" in val_str:
                            cell.fill = pass_fill
                            cell.font = pass_font
                            cell.alignment = Alignment(horizontal="center", vertical="center")

            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col[:60]:
                    v = str(cell.value or '')
                    length = sum(2 if ord(ch) > 127 else 1 for ch in v)
                    if length > max_len:
                        max_len = length
                ws.column_dimensions[col_letter].width = min(45, max(12, max_len + 3))

            wb.save(save_path)
            messagebox.showinfo(
                "导出成功",
                f"四级分类智能反写完成的 Excel 工作簿已保存：\n\n{save_path}\n\n"
                f"✔ 原表中的【依据法律法规或标准】与【检测方法】空列已原地自动填齐\n"
                f"✔ 右侧已同步附上完整的标准限量值、检出限(LOD)与判定结论！"
            )
        except Exception as e:
            messagebox.showerror("导出失败", f"无法保存 Excel 文件：\n{e}")

    def copy_backfill_preview(self):
        """复制当前预览表格到剪贴板"""
        if not self._backfill_processed_rows:
            messagebox.showinfo("提示", "当前无预览数据！")
            return
        lines_out = []
        headers = [
            "序号", "食品大类(一级)", "食品亚类(二级)", "食品品种(三级)", "食品细类(四级)/样品",
            "检验项目", "依据法律法规或标准", "检测方法", "标准限量值", "检出限(LOD)", "实测值", "判定结论", "匹配状态"
        ]
        lines_out.append("\t".join(headers))
        for it in self._backfill_processed_rows:
            row_str = "\t".join([
                str(it['seq']),
                it.get('cat1') or '/',
                it.get('cat2') or '/',
                it.get('cat3') or '/',
                it['food'],
                it['item'],
                it['std_code'],
                it['method'],
                it['limit'],
                it['lod'],
                it['meas'],
                it['judge'],
                it['status']
            ])
            lines_out.append(row_str)
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(lines_out))
        self.root.update()
        messagebox.showinfo("复制成功", f"已将全部 {len(self._backfill_processed_rows)} 条四级分类反写数据复制到剪贴板！可直接粘贴至 Excel。")

    def clear_backfill_data(self):
        """清空数据"""
        self.backfill_tree.delete(*self.backfill_tree.get_children())
        self._backfill_processed_rows = []
        self._backfill_raw_rows = []
        self.backfill_stats_label.config(text="📊 质控看板: 已清空台账数据", bootstyle=INFO)

    def on_backfill_tree_double_click(self, event=None):
        """双击查看反写行详情卡片"""
        sel = self.backfill_tree.selection()
        if not sel:
            return
        vals = self.backfill_tree.item(sel[0], 'values')
        if not vals or len(vals) < 13:
            return
        seq, c1, c2, c3, food, item, std_code, method_str, limit_val, lod_str, meas, judge, status = vals
        cat_path = " > ".join([x for x in (c1, c2, c3, food) if x and x != "/"])
        extra_note = f"【四级分类路径】{cat_path}  |  【匹配状态】{status}  |  【方法检出限(LOD)】{lod_str}"
        ItemDetailDialog(
            self.root,
            std_code,
            cat_path or food,
            item,
            limit_val,
            unit="",
            notes=extra_note,
            measured_val=meas if meas != "/" else None,
            judge_res=judge if judge != "/" else None,
            method_str=method_str
        )
    def on_standard_selected(self, event=None):
        """标准选择变更"""
        standard_code = self.standard_var.get()
        if standard_code not in self.standards:
            return
        
        std = self.standards[standard_code]
        self.standard_info_label.config(text=f"{std['name']}")
        
        self.selected_category_filter = ""
        if hasattr(self, 'cat_combo'):
            self.populate_category_dropdown()
            
        self.refresh_tree()

    def populate_category_dropdown(self, filter_kw=""):
        """提取当前标准下所有样品/食品类别并填充下拉菜单"""
        if not hasattr(self, 'cat_combo'):
            return
        standard_code = self.standard_var.get()
        if standard_code not in self.standards:
            self.cat_combo['values'] = ["全部食品分类 (显示全部)"]
            return
        
        items = self.standards[standard_code]['items']
        unique_cats = sorted(list(set(item.get('food_category', '').strip() for item in items if item.get('food_category'))))
        
        filter_kw_lower = filter_kw.strip().lower()
        if filter_kw_lower:
            filtered_cats = [c for c in unique_cats if filter_kw_lower in c.lower()]
        else:
            filtered_cats = unique_cats
            
        combo_values = ["全部食品分类 (显示全部)"] + filtered_cats
        self.cat_combo['values'] = combo_values
        
        # 若当前选中分类不在过滤列表中，重置为全部
        cur_val = self.cat_combo_var.get()
        if cur_val not in combo_values:
            self.cat_combo_var.set("全部食品分类 (显示全部)")
            self.selected_category_filter = ""

    def on_cat_combo_selected(self, event=None):
        """响应样品分类下拉菜单选择"""
        sel = self.cat_combo_var.get().strip()
        if sel == "全部食品分类 (显示全部)" or not sel:
            self.selected_category_filter = ""
        else:
            self.selected_category_filter = sel
        self.refresh_tree()

    def filter_category_dropdown(self):
        """当类别搜索框输入时，动态过滤下拉菜单选项"""
        kw = self.cat_search_var.get().strip()
        self.populate_category_dropdown(kw)
        if kw:
            vals = self.cat_combo['values']
            if len(vals) > 1:
                # 自动选中第一个匹配分类
                first_match = vals[1]
                self.cat_combo_var.set(first_match)
                self.selected_category_filter = first_match
            else:
                self.selected_category_filter = kw
        else:
            self.cat_combo_var.set("全部食品分类 (显示全部)")
            self.selected_category_filter = ""
        self.refresh_tree()

    def clear_category_filter(self):
        """重置分类"""
        if hasattr(self, 'cat_search_var'):
            self.cat_search_var.set("")
        if hasattr(self, 'cat_combo_var'):
            self.cat_combo_var.set("全部食品分类 (显示全部)")
        self.selected_category_filter = ""
        self.populate_category_dropdown()
        self.refresh_tree()
    
    def refresh_tree(self):
        """刷新数据表"""
        self.tree.delete(*self.tree.get_children())
        standard_code = self.standard_var.get()
        if standard_code not in self.standards:
            return
        
        search_text = self.search_var.get().strip()
        search_lower = search_text.lower()
        cat_filter_lower = self.selected_category_filter.lower() if self.selected_category_filter else ""
        items = self.standards[standard_code]['items']
        
        count = 0
        matching_items = []
        for item in items:
            food_cat = item.get('food_category', '')
            test_item = item.get('test_item', '')
            food_cat_lower = food_cat.lower()
            test_item_lower = test_item.lower()
            
            # 分类树筛选
            if cat_filter_lower and cat_filter_lower not in food_cat_lower:
                continue
            
            method_str = item.get('method', '')
            method_lower = method_str.lower()
            
            # 关键词搜索筛选（支持食品、项目名、检验方法标准号）
            if search_text:
                if (search_lower not in food_cat_lower and 
                    search_lower not in test_item_lower and 
                    search_lower not in method_lower):
                    continue
                matching_items.append(item)
            else:
                matching_items.append(item)
                
        # 核心优化：搜索词精准匹配优先置顶排序
        if search_text:
            matching_items.sort(key=lambda it: calc_match_rank(it, search_text))
            
        for item in matching_items:
            food_cat = item.get('food_category', '')
            test_item = item.get('test_item', '')
            method_str = item.get('method', '')
            if search_text:
                disp_food = highlight_keywords_text(food_cat, [search_text])
                disp_test = highlight_keywords_text(test_item, [search_text])
                disp_method = highlight_keywords_text(method_str, [search_text])
                row_tags = ('highlight_match',)
            else:
                disp_food = food_cat
                disp_test = test_item
                disp_method = method_str
                row_tags = ()
                
            self.tree.insert('', 'end', values=(
                disp_food,
                disp_test,
                item.get('limit_value', ''),
                item.get('unit', ''),
                disp_method,
                item.get('notes', '')
            ), tags=row_tags)
            count += 1
        
        self.result_count_label.config(text=f"显示 {count} / {len(items)} 条")
    
    def on_search(self):
        """搜索过滤"""
        self.refresh_tree()
    
    def clear_search(self):
        """清除搜索"""
        self.search_var.set("")
    
    def copy_current_selected_method(self):
        """一键复制当前选中的检验方法标准"""
        if hasattr(self, 'selected_item') and self.selected_item:
            m = self.selected_item.get('method', '')
            if not m or m == "/":
                m = "参照附录或推荐方法"
            self.root.clipboard_clear()
            self.root.clipboard_append(m)
            self.root.update()
            if hasattr(self, 'btn_copy_method'):
                self.btn_copy_method.config(text="✔ 已复制到剪贴板", bootstyle=SUCCESS)
                self.root.after(2000, lambda: self.btn_copy_method.config(text="📋 一键复制", bootstyle="info-outline") if hasattr(self, 'btn_copy_method') else None)
        else:
            messagebox.showinfo("提示", "请先在右侧表格中选择一条数据！")

    def on_item_selected(self, event):
        """项目选中"""
        selection = self.tree.selection()
        if not selection:
            return
        
        values = self.tree.item(selection[0], 'values')
        food = strip_highlights(values[0])
        item_name = strip_highlights(values[1])
        limit_value = values[2]
        unit = values[3]
        method_str = strip_highlights(values[4]) if len(values) > 5 else ""
        notes = values[5] if len(values) > 5 else (values[4] if len(values) > 4 else "")
        
        self.selected_item = {
            'food_category': food,
            'test_item': item_name,
            'limit_value': limit_value,
            'unit': unit,
            'method': method_str,
            'notes': notes
        }
        
        info = f"食品: {food}\n项目: {item_name}"
        self.selected_info_label.config(text=info)
        self.limit_display_label.config(text=f"限量值: {limit_value} {unit}")
        m_display = method_str if method_str else '参照附录或推荐方法'
        if hasattr(self, 'method_text'):
            self.method_text.delete("1.0", tk.END)
            self.method_text.insert("1.0", m_display)
        if hasattr(self, 'btn_copy_method'):
            self.btn_copy_method.config(text="📋 一键复制", bootstyle="info-outline")
    
    def do_judge(self):
        """执行判定"""
        if not hasattr(self, 'selected_item') or not self.selected_item:
            messagebox.showwarning("提示", "请先在右侧表格中选择一条限量数据")
            return
        
        test_value_str = self.test_value_var.get().strip()
        if not test_value_str:
            messagebox.showwarning("提示", "请输入检测值")
            return
        
        sel = self.selected_item
        limit_value_str = sel['limit_value']
        food = sel['food_category']
        item_name = sel['test_item']
        unit = sel['unit']
        standard_code = self.standard_var.get()
        
        # 判定逻辑
        try:
            limit_val = float(limit_value_str)
            test_val = float(test_value_str)
            if test_val <= limit_val:
                result = "合格"
                result_style = SUCCESS
            else:
                result = "不合格"
                result_style = DANGER
        except ValueError:
            result = "无法判定"
            result_style = SECONDARY
        
        self.judge_result_label.config(text=result, bootstyle=result_style)
        
        # 记录历史
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        record = f"[{timestamp}] {standard_code} | {food} | {item_name} | 检测值:{test_value_str} | 限量:{limit_value_str}{unit} | 判定:{result}"
        self.judge_history.append(record)
        self.history_text.insert(tk.END, record + "\n")
        self.history_text.see(tk.END)
    
    def export_history(self):
        if not self.judge_history:
            messagebox.showinfo("提示", "暂无判定记录可导出")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("CSV文件", "*.csv"), ("所有文件", "*.*")],
            initialfile=f"判定记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            title="导出判定记录"
        )
        if not file_path:
            return
        
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write("食品安全国家标准判定记录\n")
                f.write(f"导出时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("=" * 80 + "\n\n")
                for record in self.judge_history:
                    f.write(record + "\n")
            messagebox.showinfo("成功", f"判定记录已导出到：\n{file_path}")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{str(e)}")
    
    def clear_history(self):
        if not self.judge_history:
            return
        if messagebox.askyesno("确认", "确定要清空所有判定记录吗？"):
            self.judge_history.clear()
            self.history_text.delete("1.0", tk.END)
            self.judge_result_label.config(text="")
            self.test_value_var.set("")

    def on_query_tree_double_click(self, event):
        """双击标准查询表格查看详情"""
        selection = self.tree.selection()
        if not selection:
            return
        vals = self.tree.item(selection[0], 'values')
        std_code = self.standard_var.get()
        food = strip_highlights(vals[0])
        item_name = strip_highlights(vals[1])
        limit_val = vals[2]
        unit = vals[3]
        method_val = strip_highlights(vals[4]) if len(vals) > 5 else ""
        notes = vals[5] if len(vals) > 5 else (vals[4] if len(vals) > 4 else "")
        ItemDetailDialog(self.root, std_code, food, item_name, limit_val, unit, notes, method_str=method_val)

    def on_match_unified_tree_double_click(self, event):
        """双击样品匹配(联合查询)表格查看详情"""
        selection = self.match_unified_tree.selection()
        if not selection:
            return
        vals = self.match_unified_tree.item(selection[0], 'values')
        std_code = vals[0]
        item_name = strip_highlights(vals[1])
        food = strip_highlights(vals[2])
        limit_val = vals[3]
        method_str = strip_highlights(vals[4]) if len(vals) > 5 else ""
        notes = vals[5] if len(vals) > 5 else (vals[4] if len(vals) > 4 else "")
        ItemDetailDialog(self.root, std_code, food, item_name, limit_val, "", notes, method_str=method_str)

    def on_match_tab_tree_double_click(self, std_code, tree):
        """双击样品匹配(分Tab)表格查看详情"""
        selection = tree.selection()
        if not selection:
            return
        vals = tree.item(selection[0], 'values')
        item_name = strip_highlights(vals[0])
        food = strip_highlights(vals[1])
        limit_val = vals[2]
        notes = vals[3] if len(vals) > 3 else ""
        ItemDetailDialog(self.root, std_code, food, item_name, limit_val, "", notes)

    def on_batch_tree_double_click(self, event):
        """双击批量查询表格查看详情"""
        selection = self.batch_tree.selection()
        if not selection:
            return
        vals = self.batch_tree.item(selection[0], 'values')
        std_code = vals[0]
        food = strip_highlights(vals[1])
        item_name = strip_highlights(vals[2])
        limit_val = vals[3]
        meas_val = vals[4] if len(vals) > 4 else None
        judge_res = vals[5] if len(vals) > 5 else None
        method_str = strip_highlights(vals[6]) if len(vals) > 6 else ""
        notes = vals[7] if len(vals) > 7 else (vals[6] if len(vals) > 6 else "")
        ItemDetailDialog(self.root, std_code, food, item_name, limit_val, "", notes, meas_val, judge_res, method_str=method_str)

    def export_batch_results(self):
        """导出批量查询与判定当前表格中的所有结果"""
        children = self.batch_tree.get_children()
        if not children:
            messagebox.showinfo("提示", "当前表格中没有可导出的数据！")
            return
            
        file_path = filedialog.asksaveasfilename(
            title="导出批量查询与判定结果",
            defaultextension=".csv",
            filetypes=[("CSV表格文件 (Excel兼容)", "*.csv"), ("文本文件", "*.txt"), ("所有文件", "*.*")],
            initialfile=f"食品标准批量判定核查结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        if not file_path:
            return
            
        try:
            with open(file_path, 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(["标准名称", "食品类别", "检测项目", "限量值", "实测值", "判定结论", "检验方法", "备注"])
                for item_id in children:
                    vals = self.batch_tree.item(item_id, 'values')
                    clean_vals = [strip_highlights(v) for v in vals]
                    writer.writerow(clean_vals)
                    
            messagebox.showinfo("导出成功", f"已成功导出 {len(children)} 条批量查询与判定记录到：\n{file_path}")
        except Exception as e:
            messagebox.showerror("导出失败", f"保存文件时出错: {e}")

    def export_match_results(self):
        """导出样品匹配当前表格中的所有结果"""
        if self.match_unified_var.get():
            children = self.match_unified_tree.get_children()
            if not children:
                messagebox.showinfo("提示", "当前表格中没有可导出的数据！")
                return
                
            file_path = filedialog.asksaveasfilename(
                title="导出样品匹配结果",
                defaultextension=".csv",
                filetypes=[("CSV表格文件 (Excel兼容)", "*.csv"), ("文本文件", "*.txt"), ("所有文件", "*.*")],
                initialfile=f"样品匹配结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            )
            if not file_path:
                return
                
            try:
                with open(file_path, 'w', encoding='utf-8-sig', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(["标准名称", "检测项目", "食品类别", "限量值", "检验方法", "备注"])
                    for item_id in children:
                        vals = self.match_unified_tree.item(item_id, 'values')
                        clean_vals = [strip_highlights(v) for v in vals]
                        writer.writerow(clean_vals)
                        
                messagebox.showinfo("导出成功", f"已成功导出 {len(children)} 条匹配记录到：\n{file_path}")
            except Exception as e:
                messagebox.showerror("导出失败", f"保存文件时出错: {e}")
        else:
            current_tab_idx = self.match_notebook.index(self.match_notebook.select())
            std_code = sorted(self.standards.keys())[current_tab_idx]
            tree = self.match_trees[std_code]
            children = tree.get_children()
            if not children:
                messagebox.showinfo("提示", f"【{std_code}】分类下暂无可导出数据！")
                return
                
            file_path = filedialog.asksaveasfilename(
                title=f"导出【{std_code}】匹配结果",
                defaultextension=".csv",
                filetypes=[("CSV表格文件 (Excel兼容)", "*.csv"), ("文本文件", "*.txt"), ("所有文件", "*.*")],
                initialfile=f"{std_code}_匹配结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            )
            if not file_path:
                return
                
            try:
                with open(file_path, 'w', encoding='utf-8-sig', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(["检测项目", "食品类别", "限量值", "备注"])
                    for item_id in children:
                        vals = tree.item(item_id, 'values')
                        clean_vals = [strip_highlights(v) for v in vals]
                        writer.writerow(clean_vals)
                        
                messagebox.showinfo("导出成功", f"已成功导出 {len(children)} 条【{std_code}】记录到：\n{file_path}")
            except Exception as e:
                messagebox.showerror("导出失败", f"保存文件时出错: {e}")

    def clear_all_batch_col_filters(self):
        """清除批量判定表中所有列的自定义下拉筛选"""
        if hasattr(self, 'batch_col_filters'):
            for col in self.batch_columns:
                self.batch_col_filters[col] = None
                self.batch_tree.heading(col, text=col + " ▾")
            self.filter_batch_results()
            messagebox.showinfo("提示", "已清除所有列的自定义筛选条件！")

    def on_window_close(self):
        """窗口关闭时自动保存所有表格的列宽与窗口尺寸"""
        try:
            self.save_column_widths()
        except Exception as e:
            print(f"Error saving config on close: {e}")
        self.root.destroy()

    def save_column_widths(self):
        """保存当前各表格的列宽与窗口位置尺寸"""
        config = load_ui_config()
        config['window_geometry'] = self.root.geometry()
        try:
            if hasattr(self, 'tree') and self.tree.winfo_exists():
                config['query_tree'] = {col: self.tree.column(col, 'width') for col in ("食品分类", "检测项目", "限量值", "单位", "备注")}
        except Exception:
            pass
        try:
            if hasattr(self, 'match_unified_tree') and self.match_unified_tree.winfo_exists():
                config['match_unified_tree'] = {col: self.match_unified_tree.column(col, 'width') for col in ("标准名称", "检测项目", "食品类别", "限量值", "备注")}
        except Exception:
            pass
        try:
            if hasattr(self, 'batch_tree') and self.batch_tree.winfo_exists():
                config['batch_tree'] = {col: self.batch_tree.column(col, 'width') for col in self.batch_columns}
        except Exception:
            pass
        try:
            if hasattr(self, 'backfill_tree') and self.backfill_tree.winfo_exists():
                config['backfill_tree'] = {col: self.backfill_tree.column(col, 'width') for col in self.backfill_tree['columns']}
        except Exception:
            pass
        save_ui_config(config)

    def open_column_visibility_dialog(self, tree, tree_key, default_widths):
        """弹出列显示/隐藏自定义窗口"""
        all_cols = list(tree['columns'])
        ColumnVisibilityDialog(self.root, tree, tree_key, all_cols, default_widths, self.apply_column_visibility)

    def apply_column_visibility(self, tree, tree_key, visible_map, default_widths):
        """应用列可见性更改并保存"""
        config = load_ui_config()
        if 'hidden_columns' not in config:
            config['hidden_columns'] = {}
            
        hidden_list = []
        for col, is_visible in visible_map.items():
            if not is_visible:
                hidden_list.append(col)
                tree.column(col, width=0, minwidth=0, stretch=False)
            else:
                curr_w = tree.column(col, 'width')
                target_w = curr_w if curr_w > 0 else default_widths.get(col, 100)
                tree.column(col, width=target_w, minwidth=30, stretch=True)
                
        config['hidden_columns'][tree_key] = hidden_list
        save_ui_config(config)
        messagebox.showinfo("设置已保存", "已更新表格列的显示/隐藏设置！\n此偏好将在退出时自动记忆保存。")

    def apply_tree_column_widths(self, tree, tree_key, default_widths):
        """从本地持久化配置中恢复记忆的列宽与隐藏状态"""
        saved_widths = self.ui_config.get(tree_key, {})
        hidden_cols = self.ui_config.get('hidden_columns', {}).get(tree_key, [])
        valid_cols = set(tree['columns'])
        
        for col, default_w in default_widths.items():
            if col not in valid_cols:
                continue
            try:
                if col in hidden_cols:
                    tree.column(col, width=0, minwidth=0, stretch=False)
                else:
                    w = saved_widths.get(col, default_w)
                    try:
                        w_int = int(w)
                        if "方法" in col and w_int < 260:
                            w_int = max(default_w, 280)
                        tree.column(col, width=w_int, minwidth=30, stretch=True)
                    except Exception:
                        tree.column(col, width=default_w, minwidth=30, stretch=True)
            except Exception:
                pass

    def autofit_tree_columns(self, tree, columns=None, min_w=80, max_w=500):
        """根据当前表格中的内容自动计算并调整最佳列宽"""
        if not columns:
            columns = tree['columns']
        children = tree.get_children()
        if not children:
            messagebox.showinfo("提示", "表格中暂无数据可供计算自适应列宽")
            return
            
        sample_ids = children[:150]
        
        for col_idx, col in enumerate(columns):
            # 如果是隐藏列则跳过自适应
            if tree.column(col, 'width') == 0:
                continue
                
            col_text = str(tree.heading(col, 'text')).replace(' ▾', '').replace(' (已筛选)', '')
            max_len = len(col_text.encode('gbk', errors='ignore'))
            
            for item_id in sample_ids:
                vals = tree.item(item_id, 'values')
                if col_idx < len(vals):
                    val_str = strip_highlights(str(vals[col_idx]))
                    val_len = len(val_str.encode('gbk', errors='ignore'))
                    if val_len > max_len:
                        max_len = val_len
                        
            calc_w = max(min_w, min(max_w, max_len * 8 + 30))
            tree.column(col, width=calc_w, minwidth=30, stretch=True)
            
        messagebox.showinfo("自适应调整完成", "已根据当前数据内容自适应优化了各列宽度！\n调整后的列宽将在退出时自动记忆保存。")

    def reset_tree_column_widths(self, tree, tree_key, default_widths):
        """重置表格为初始默认列宽并显示所有列"""
        for col, width in default_widths.items():
            tree.column(col, width=width, minwidth=30, stretch=True)
            
        config = load_ui_config()
        if 'hidden_columns' in config and tree_key in config['hidden_columns']:
            config['hidden_columns'][tree_key] = []
            save_ui_config(config)
        messagebox.showinfo("提示", "已恢复该表格的默认列宽与全部列显示设置。")

    def setup_tree_context_menu(self, tree, tree_key, default_widths, double_click_callback=None):
        """为表格配置右键快捷菜单：自适应列宽、自定义列显示/隐藏、恢复默认、查看详情、智能气泡与快捷复制"""
        # 1. 挂载高级单元格智能悬浮气泡
        TreeviewCellTooltip(tree)

        # 2. 绑定 Ctrl+C 快捷键一键复制检验方法或整行
        def on_tree_ctrl_c(event):
            sel = tree.selection()
            if not sel:
                return
            vals = tree.item(sel[0], 'values')
            columns = tree['columns']
            method_val = ""
            for c_idx, c_name in enumerate(columns):
                if "方法" in c_name and c_idx < len(vals):
                    method_val = strip_highlights(str(vals[c_idx])).strip()
            copy_text = method_val if method_val and method_val != "/" else "\t".join(strip_highlights(str(v)) for v in vals)
            self.root.clipboard_clear()
            self.root.clipboard_append(copy_text)
            self.root.update()

        tree.bind("<Control-c>", on_tree_ctrl_c)
        tree.bind("<Control-C>", on_tree_ctrl_c)

        # 3. 动态右键菜单
        def show_context_menu(event):
            item = tree.identify_row(event.y)
            col = tree.identify_column(event.x)
            if item:
                tree.selection_set(item)
                
            ctx_menu = tk.Menu(self.root, tearoff=0)
            
            clicked_cell_text = ""
            row_method_text = ""
            row_all_text = ""
            
            if item:
                vals = tree.item(item, 'values')
                columns = tree['columns']
                if col:
                    try:
                        col_idx = int(col.replace("#", "")) - 1
                        if 0 <= col_idx < len(vals):
                            clicked_cell_text = strip_highlights(str(vals[col_idx])).strip()
                    except Exception:
                        pass
                for c_idx, c_name in enumerate(columns):
                    if "方法" in c_name and c_idx < len(vals):
                        row_method_text = strip_highlights(str(vals[c_idx])).strip()
                row_all_text = "\t".join(strip_highlights(str(v)) for v in vals)

            if clicked_cell_text and clicked_cell_text != "/":
                def copy_cell():
                    self.root.clipboard_clear()
                    self.root.clipboard_append(clicked_cell_text)
                    self.root.update()
                lbl = clicked_cell_text[:22] + "..." if len(clicked_cell_text) > 22 else clicked_cell_text
                ctx_menu.add_command(label=f"📋  复制单元格内容: 【{lbl}】", command=copy_cell)

            if row_method_text and row_method_text != "/":
                def copy_method():
                    self.root.clipboard_clear()
                    self.root.clipboard_append(row_method_text)
                    self.root.update()
                m_lbl = row_method_text[:22] + "..." if len(row_method_text) > 22 else row_method_text
                ctx_menu.add_command(label=f"🔬  复制本项检验方法: 【{m_lbl}】", command=copy_method)

            if row_all_text:
                def copy_row():
                    self.root.clipboard_clear()
                    self.root.clipboard_append(row_all_text)
                    self.root.update()
                ctx_menu.add_command(label="📄  复制本行全部数据 (Tab分隔制表符)", command=copy_row)
                
            ctx_menu.add_separator()
            ctx_menu.add_command(label="↔  智能自适应此表所有列宽", command=lambda: self.autofit_tree_columns(tree))
            ctx_menu.add_command(label="👁️  自定义各列显示/隐藏...", command=lambda: self.open_column_visibility_dialog(tree, tree_key, default_widths))
            ctx_menu.add_command(label="↺  恢复此表默认列宽与列显示", command=lambda: self.reset_tree_column_widths(tree, tree_key, default_widths))
            if double_click_callback:
                ctx_menu.add_separator()
                ctx_menu.add_command(label="🔍  查看选中项详细指标卡片 (双击)", command=lambda: double_click_callback(None))
                
            ctx_menu.post(event.x_root, event.y_root)
            
        tree.bind("<Button-3>", show_context_menu)



def acquire_single_instance_mutex():
    """获取进程级单实例互斥锁，防止用户连续双击启动多个进程"""
    if sys.platform == "win32":
        try:
            import ctypes
            mutex_name = "Global\\FQT_FoodStandardChecker_Clean_Mutex_V13"
            kernel32 = ctypes.windll.kernel32
            mutex = kernel32.CreateMutexW(None, False, mutex_name)
            if kernel32.GetLastError() == 183:
                os._exit(0)
            return mutex
        except Exception:
            return None
    return None


def main():
    multiprocessing.freeze_support()
    _mutex_handle = acquire_single_instance_mutex()

    root = tb.Window(themename="cosmo")
    root.withdraw()

    splash = tb.Toplevel(root)
    splash.title("加载中")
    splash.geometry("420x160")
    splash.overrideredirect(True)

    sw = splash.winfo_screenwidth()
    sh = splash.winfo_screenheight()
    splash.geometry(f"+{sw//2-210}+{sh//2-80}")

    splash_frame = tk.Frame(splash, bg="#0F172A", bd=2, relief=tk.RIDGE)
    splash_frame.pack(fill=tk.BOTH, expand=True)

    tk.Label(
        splash_frame,
        text="🛡️ 食品安全国家标准判定核查工具 V1.4",
        font=("Microsoft YaHei", 13, "bold"),
        fg="#F8FAFC",
        bg="#0F172A"
    ).pack(pady=(32, 8))
    tk.Label(
        splash_frame,
        text="正在加载 6 大国标 (83,973条) 与 2026 抽检细则四级分类库...",
        font=("Microsoft YaHei", 9),
        fg="#38BDF8",
        bg="#0F172A"
    ).pack()
    tk.Label(
        splash_frame,
        text="⚡ Powered by 陈耿忠",
        font=("Microsoft YaHei", 8, "bold"),
        fg="#94A3B8",
        bg="#0F172A"
    ).pack(side=tk.BOTTOM, pady=10)

    splash.update()

    app = FoodStandardChecker(root)

    splash.destroy()
    root.deiconify()
    root.lift()
    root.focus_force()
    root.mainloop()


if __name__ == "__main__":
    main()
