# -*- coding: utf-8 -*-
"""
食品安全国家标准判定核查工具（完整版）
标准覆盖：GB 2760-2024, GB 2761-2017, GB 2762-2025, GB 2763-2026, GB 29921-2021, GB 31650-2019
"""
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
# 检验方法标准智能匹配与抽取引擎
# ============================================================
SAMPLING_RULES_METHOD_INDEX = {}

def init_sampling_rules_method_index():
    """初始化抽检细则项目检验方法索引"""
    global SAMPLING_RULES_METHOD_INDEX
    try:
        s_db = load_json_safe('sampling_rules_2026_db.json')
        if s_db and isinstance(s_db, dict):
            for cat_name, c_data in s_db.items():
                for it in c_data.get('items', []):
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


# ============================================================
# GB 2760 添加剂功能分组数据库（用于比例之和计算）
# ============================================================
ADDITIVE_FUNCTION_GROUPS = {
    "防腐剂": {
        "items": ["苯甲酸", "苯甲酸钠", "山梨酸", "山梨酸钾", "脱氢乙酸", "脱氢乙酸钠", "对羟基苯甲酸甲酯钠", "对羟基苯甲酸乙酯", "对羟基苯甲酸丙酯", "双乙酸钠", "丙酸", "丙酸钠", "丙酸钙", "二氧化硫", "焦亚硫酸钠", "亚硫酸钠", "低亚硫酸钠", "纳他霉素", "乳酸链球菌素", "单辛酸甘油酯"],
        "color": "#0d6efd"
    },
    "相同色泽着色剂(黄色系)": {
        "items": ["柠檬黄", "柠檬黄铝色淀", "日落黄", "日落黄铝色淀", "姜黄", "栀子黄", "核黄素", "核黄素5'-磷酸钠", "β-胡萝卜素", "叶黄素"],
        "color": "#fd7e14"
    },
    "相同色泽着色剂(红色系)": {
        "items": ["胭脂红", "胭脂红铝色淀", "苋菜红", "苋菜红铝色淀", "诱惑红", "诱惑红铝色淀", "赤藓红", "赤藓红铝色淀", "红曲红", "高粱红", "萝卜红", "甜菜红", "辣椒红"],
        "color": "#dc3545"
    },
    "相同色泽着色剂(蓝色系)": {
        "items": ["亮蓝", "亮蓝铝色淀", "栀子蓝", "靛蓝", "靛蓝铝色淀"],
        "color": "#0dcaf0"
    },
    "相同色泽着色剂(绿色系)": {
        "items": ["叶绿素铜钠盐", "叶绿素铜钾盐", "叶绿素铜"],
        "color": "#198754"
    },
    "相同色泽着色剂(黑色/棕色系)": {
        "items": ["焦糖色", "植物炭黑", "可可壳色素"],
        "color": "#6c757d"
    },
    "抗氧化剂": {
        "items": ["TBHQ", "特丁基对苯二酚", "BHA", "丁基羟基茴香醚", "BHT", "二丁基羟基甲苯", "没食子酸丙酯", "PG", "植酸", "茶多酚", "抗坏血酸棕榈酸酯", "异抗坏血酸钠", "竹叶黄酮"],
        "color": "#6f42c1"
    },
    "甜味剂": {
        "items": ["糖精钠", "甜蜜素", "环己基氨基磺酸钠", "环己基氨基磺酸钙", "安赛蜜", "乙酰磺胺酸钾", "阿斯巴甜", "三氯蔗糖", "纽甜", "甜菊糖苷", "罗汉果甜苷", "索马甜", "爱德万甜"],
        "color": "#20c997"
    }
}

# ============================================================
# 常见市售商品名 / 执行标准 ➔ 国标分类智能反向映射知识库
# ============================================================
FOOD_CATEGORY_INFERENCE_DB = [
    {
        "names": ["辣条", "大面筋", "素大头", "大辣棒", "亲嘴烧", "面筋零食", "香辣豆皮面筋"],
        "standards": ["GB/T 22106 调味面制品", "GB 2714 酱腌菜(部分地区旧标)", "Q/企业标准 调味面制品"],
        "gb2760_cat": "16.07 调味面制品",
        "gb2762_cat": "粮食及其制品 > 面制品",
        "gb29921_cat": "即食预包装面制品",
        "desc": "以小麦粉或其他谷物粉等为主要原料，添加调味料、食用植物油等，经挤压熟化制成。",
        "key_risk_items": ["脱氢乙酸(2024新标禁用)", "山梨酸", "甜蜜素", "糖精钠", "菌落总数", "酸价", "过氧化值"]
    },
    {
        "names": ["面包", "老式大面包", "吐司", "全麦面包", "切片面包", "牛角面包", "法棍", "小餐包", "手撕面包"],
        "standards": ["GB 7099 糕点、面包", "GB/T 20981 面包"],
        "gb2760_cat": "07.01 面包",
        "gb2762_cat": "粮食及其制品 > 糕点、面包",
        "gb29921_cat": "糕点类、面包类",
        "desc": "以小麦粉为主原料，酵母发酵烘烤制成。",
        "key_risk_items": ["脱氢乙酸(2024新标禁用)", "丙酸钙", "山梨酸钾", "安赛蜜", "霉菌", "菌落总数"]
    },
    {
        "names": ["糕点", "月饼", "蛋黄酥", "绿豆糕", "凤梨酥", "麻薯", "泡芙", "提拉米苏", "虎皮蛋糕", "戚风蛋糕", "桃酥", "沙琪玛"],
        "standards": ["GB 7099 糕点、面包", "GB/T 20977 糕点通则", "GB/T 19855 月饼"],
        "gb2760_cat": "07.02 糕点 (07.02.01 中式糕点 / 07.02.02 西式糕点)",
        "gb2762_cat": "粮食及其制品 > 糕点",
        "gb29921_cat": "糕点类、面包类",
        "desc": "以谷物、豆类、薯类或油脂、糖、蛋等为主要原料，经成型、熟化制成。",
        "key_risk_items": ["脱氢乙酸(2024新标禁用)", "山梨酸", "铝的残留量", "防腐剂比例之和", "菌落总数", "过氧化值"]
    },
    {
        "names": ["饼干", "苏打饼干", "曲奇", "威化饼干", "韧性饼干", "酥性饼干", "夹心饼干", "手指饼干"],
        "standards": ["GB 7100 饼干", "GB/T 20980 饼干"],
        "gb2760_cat": "07.03 饼干",
        "gb2762_cat": "粮食及其制品 > 饼干",
        "gb29921_cat": "饼干类",
        "desc": "以小麦粉、糖、油脂为主要原料经烘烤制成。",
        "key_risk_items": ["二氧化硫", "铝的残留量", "甜蜜素", "菌落总数", "过氧化值"]
    },
    {
        "names": ["火腿肠", "午餐肉罐头", "红肠", "烤肠", "热狗肠", "熏煮香肠", "台式烤肠", "玉米热狗肠"],
        "standards": ["GB/T 20712 火腿肠", "GB 2726 熟肉制品", "GB/T 20711 熏煮香肠"],
        "gb2760_cat": "08.03.01 熟肉制品 > 灌肠类",
        "gb2762_cat": "肉及肉制品 > 熟肉制品",
        "gb29921_cat": "熟肉制品",
        "desc": "以畜禽肉为主要原料，添加淀粉、调味品等灌肠后经高温灭菌或熏煮制成。",
        "key_risk_items": ["亚硝酸盐", "山梨酸", "诱惑红/胭脂红", "氯霉素", "单核细胞增生李斯特氏菌"]
    },
    {
        "names": ["酱卤肉", "卤猪蹄", "卤牛肉", "绝味鸭脖", "周黑鸭", "酱鸭", "烧鸡", "扒鸡", "盐水鸭", "卤肉饭浇头"],
        "standards": ["GB 2726 熟肉制品", "GB/T 23586 酱卤肉制品"],
        "gb2760_cat": "08.03.01.01 酱卤肉制品类",
        "gb2762_cat": "肉及肉制品 > 熟肉制品",
        "gb29921_cat": "熟肉制品",
        "desc": "肉类加入调味料和香辛料，以煮制或卤制熟化制成的肉制品。",
        "key_risk_items": ["亚硝酸钠", "山梨酸钾", "苯甲酸(熟肉禁用)", "胭脂红", "防腐剂比例之和", "大肠菌群"]
    },
    {
        "names": ["肉干", "牛肉干", "猪肉脯", "肉松", "肉条", "风干牛肉", "肉脆片", "猪肉干"],
        "standards": ["GB 2726 熟肉制品", "GB/T 23969 肉干", "GB/T 31406 肉脯"],
        "gb2760_cat": "08.03.02 熟肉制品 > 肉干类",
        "gb2762_cat": "肉及肉制品 > 熟肉制品",
        "gb29921_cat": "熟肉制品",
        "desc": "肉类经切块或切片、调味、熟制、干燥制成的熟肉产品。",
        "key_risk_items": ["山梨酸", "苯甲酸(禁用)", "防腐剂比例之和", "铅", "菌落总数"]
    },
    {
        "names": ["泡椒凤爪", "卤鸡爪", "虎皮凤爪", "去骨鸭掌", "糟禽肉", "麻辣鸡翅"],
        "standards": ["GB 2726 熟肉制品", "DBS50/ 028 泡椒凤爪(地方标准)"],
        "gb2760_cat": "08.03.01.01 酱卤肉制品 或 08.03.09 糟醉肉制品",
        "gb2762_cat": "肉及肉制品",
        "gb29921_cat": "熟肉制品",
        "desc": "禽爪经水煮后，在泡椒液中浸泡腌制发酵或卤制熟化而成。",
        "key_risk_items": ["山梨酸钾", "苯甲酸(禁用)", "二氧化硫", "脱氢乙酸(禁用)", "菌落总数"]
    },
    {
        "names": ["发酵乳", "风味发酵乳", "酸奶", "酸牛奶", "希腊酸奶", "发酵酸乳", "果粒酸奶"],
        "standards": ["GB 19302 发酵乳"],
        "gb2760_cat": "01.02.02 风味发酵乳 (纯发酵乳01.02.01不得加任何添加剂)",
        "gb2762_cat": "乳及乳制品",
        "gb29921_cat": "乳与乳制品",
        "desc": "以生牛(羊)乳为原料，接种乳酸菌发酵，风味发酵乳允许添加果蔬、谷物及食品添加剂。",
        "key_risk_items": ["山梨酸(禁用)", "安赛蜜(禁用)", "三聚氰胺", "黄曲霉毒素M1", "酵母", "霉菌"]
    },
    {
        "names": ["果汁", "橙汁", "苹果汁", "浓缩果汁", "果蔬汁饮料", "NFC果汁", "纯果汁"],
        "standards": ["GB 7101 饮料", "GB/T 31121 果蔬汁类及其饮料"],
        "gb2760_cat": "14.02.01 果蔬汁(肉) / 14.02.03 果蔬汁(肉)饮料",
        "gb2762_cat": "饮料类",
        "gb29921_cat": "饮料类",
        "desc": "以水果或蔬菜为原料制成的汁液或饮料。",
        "key_risk_items": ["苯甲酸", "山梨酸", "甜蜜素", "安赛蜜", "合成着色剂", "展青霉素(苹果汁)"]
    },
    {
        "names": ["碳酸饮料", "汽水", "可乐", "雪碧", "苏打水汽水", "果味汽水"],
        "standards": ["GB 7101 饮料", "GB/T 10792 碳酸饮料(汽水)"],
        "gb2760_cat": "14.04 碳酸饮料",
        "gb2762_cat": "饮料类",
        "gb29921_cat": "饮料类",
        "desc": "充入二氧化碳气的清凉饮料。",
        "key_risk_items": ["苯甲酸", "山梨酸", "甜蜜素", "阿斯巴甜", "焦糖色", "咖啡因"]
    },
    {
        "names": ["包装饮用水", "矿泉水", "纯净水", "饮用天然水", "蒸馏水", "山泉水", "桶装水"],
        "standards": ["GB 19298 包装饮用水", "GB 8537 饮用天然矿泉水"],
        "gb2760_cat": "14.01 饮用水 (除允许加工助剂外不得添加任何食品添加剂)",
        "gb2762_cat": "包装饮用水",
        "gb29921_cat": "包装饮用水",
        "desc": "密封于符合标准的容器中，可直接饮用的水。",
        "key_risk_items": ["溴酸盐", "亚硝酸盐", "铜绿假单胞菌", "大肠菌群", "三氯甲烷"]
    },
    {
        "names": ["果冻", "吸吸果冻", "果肉果冻", "布丁果冻", "可吸果冻", "果冻布丁"],
        "standards": ["GB 19299 果冻", "GB/T 19883 果冻"],
        "gb2760_cat": "16.01 果冻",
        "gb2762_cat": "其他类食品",
        "gb29921_cat": "果冻",
        "desc": "以水、白砂糖、卡拉胶、魔芋粉等为原料制成的胶冻食品。",
        "key_risk_items": ["苯甲酸(果冻禁用)", "山梨酸", "甜蜜素", "柠檬黄/日落黄", "防腐剂比例之和"]
    },
    {
        "names": ["蜜饯", "话梅", "果脯", "葡萄干", "芒果干", "无花果干", "陈皮梅", "杨梅干", "加应子", "九制陈皮"],
        "standards": ["GB 14884 蜜饯凉果", "GB/T 10782 蜜饯通则"],
        "gb2760_cat": "04.01.02.08 蜜饯凉果",
        "gb2762_cat": "水果及其制品 > 水果干类、蜜饯凉果",
        "gb29921_cat": "蜜饯凉果",
        "desc": "以水果或果脯为原料，经糖或蜂蜜、盐腌制浸渍、干燥等工艺制成。",
        "key_risk_items": ["二氧化硫残留量", "甜蜜素", "糖精钠", "苯甲酸", "山梨酸", "合成着色剂", "铅"]
    },
    {
        "names": ["酱腌菜", "酸菜", "泡菜", "榨菜", "泡豇豆", "红油金针菇", "腌萝卜", "糖蒜", "野山椒", "辣白菜"],
        "standards": ["GB 2714 酱腌菜", "SB/T 10439 酱腌菜"],
        "gb2760_cat": "04.02.02.03 腌渍的蔬菜",
        "gb2762_cat": "蔬菜及其制品 > 腌渍蔬菜",
        "gb29921_cat": "酱腌菜",
        "desc": "以新鲜蔬菜为原料，经盐腌、酱渍或发酵制成的蔬菜制品。",
        "key_risk_items": ["二氧化硫", "苯甲酸", "山梨酸", "脱氢乙酸(2024新标禁用)", "糖精钠", "亚硝酸盐", "大肠菌群"]
    },
    {
        "names": ["豆腐干", "豆干", "辣豆干", "豆腐皮", "腐竹", "素肉", "素大豆蛋白制品", "麻辣豆干"],
        "standards": ["GB 2712 豆制品", "GB/T 22108 非发酵性豆制品"],
        "gb2760_cat": "04.04.01.02 豆干类 / 04.04.01.03 再制豆制品",
        "gb2762_cat": "大豆及大豆制品",
        "gb29921_cat": "熟制豆制品",
        "desc": "以大豆或大豆蛋白为主要原料制成的非发酵豆制品。",
        "key_risk_items": ["苯甲酸(禁用)", "山梨酸", "脱氢乙酸(禁用)", "铝的残留量", "防腐剂比例之和", "菌落总数"]
    },
    {
        "names": ["坚果炒货", "瓜子", "花生", "巴旦木", "开心果", "核桃", "夏威夷果", "每日坚果", "碧根果", "腰果"],
        "standards": ["GB 19300 坚果与籽类食品", "GB/T 22165 坚果与籽类食品通则"],
        "gb2760_cat": "04.05.02 熟制坚果与籽类 (04.05.02.01 带壳 / 04.05.02.02 脱壳)",
        "gb2762_cat": "坚果及籽类",
        "gb29921_cat": "熟制坚果与籽类食品",
        "desc": "以坚果、籽类为主要原料，经烘炒、油炸、蒸煮或其他熟制工艺加工而成。",
        "key_risk_items": ["黄曲霉毒素B1", "二氧化硫", "糖精钠", "甜蜜素", "酸价", "过氧化值"]
    },
    {
        "names": ["食用植物油", "大豆油", "花生油", "玉米油", "菜籽油", "葵花籽油", "芝麻油", "食用调和油", "橄榄油"],
        "standards": ["GB 2716 植物油", "GB/T 1535 大豆油", "GB/T 1536 花生油"],
        "gb2760_cat": "02.01 植物油",
        "gb2762_cat": "油脂及其制品",
        "gb29921_cat": "油脂类(不适用)",
        "desc": "以食用植物油料为原料制取的油脂。",
        "key_risk_items": ["酸价", "过氧化值", "TBHQ (抗氧化剂)", "BHA", "BHT", "抗氧化剂比例之和", "苯并[a]芘", "溶剂残留量"]
    },
    {
        "names": ["火锅底料", "复合调味料", "烧烤料", "排骨汤料", "麻辣烫底料", "炸鸡裹粉", "调味粉", "浓汤宝"],
        "standards": ["GB 31644 复合调味料", "DBS50/ 022 火锅底料"],
        "gb2760_cat": "12.10 复合调味料 (12.10.01 固态 / 12.10.02 半固态)",
        "gb2762_cat": "调味品",
        "gb29921_cat": "复合调味料",
        "desc": "由两种或两种以上调味原料复合加工而成的调味料。",
        "key_risk_items": ["罂粟碱/吗啡(非法添加)", "苯甲酸", "山梨酸", "甜蜜素", "脱氢乙酸(2024新标禁用)", "防腐剂比例之和"]
    },
    {
        "names": ["酱油", "生抽", "老抽", "酿造酱油", "配制酱油", "海鲜酱油", "蒸鱼豉油"],
        "standards": ["GB 2717 酱油", "GB/T 18186 酿造酱油"],
        "gb2760_cat": "12.03 酱油",
        "gb2762_cat": "调味品 > 酱油",
        "gb29921_cat": "酱油",
        "desc": "以大豆和(或)脱脂大豆、小麦和(或)小麦粉为原料，经微生物发酵制成的调味品。",
        "key_risk_items": ["氨基酸态氮", "苯甲酸", "山梨酸", "糖精钠", "三氯蔗糖", "3-氯-1,2-丙二醇", "防腐剂比例之和"]
    },
    {
        "names": ["食醋", "陈醋", "香醋", "米醋", "白醋", "酿造食醋", "饺子醋"],
        "standards": ["GB 2719 食醋", "GB/T 18187 酿造食醋"],
        "gb2760_cat": "12.04 食醋",
        "gb2762_cat": "调味品 > 食醋",
        "gb29921_cat": "食醋",
        "desc": "以谷物、果类等为原料，经酒精发酵和醋酸发酵制成的酸性调味品。",
        "key_risk_items": ["总酸", "苯甲酸", "山梨酸", "游离矿酸(不得检出)", "菌落总数"]
    }
]

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
# ============================================================
MICROBIAL_STANDARDS_DB = {
    # 1. 熟肉制品类 (GB 2726 / GB 29921 / 2026细则)
    "熟肉制品 - 酱卤肉制品 (GB 2726/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 10000, "M": 100000, "unit": "CFU/g", "std": "GB 2726-2016", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 2726-2016", "method": "GB 4789.3 平板法"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10 第二法"},
        {"item": "单核细胞增生李斯特氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.30"},
        {"item": "大肠埃希氏菌 O157:H7 (牛肉类)", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.36"}
    ],
    "熟肉制品 - 熏烧烤肉制品 (GB 2726/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 10000, "M": 100000, "unit": "CFU/g", "std": "GB 2726-2016", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 2726-2016", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"},
        {"item": "单核细胞增生李斯特氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.30"}
    ],
    "熟肉制品 - 熟肉干制品 (肉干/肉松/肉脯) (GB 2726/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 10000, "M": 100000, "unit": "CFU/g", "std": "GB 2726-2016", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 2726-2016", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    "熟肉制品 - 发酵肉制品 (发酵火腿/发酵香肠) (GB 2726/29921)": [
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 2726-2016", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"},
        {"item": "单核细胞增生李斯特氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.30"}
    ],
    "熟肉制品 - 灌肠类与西式火腿 (火腿肠/熏煮香肠) (GB 2726/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 10000, "M": 100000, "unit": "CFU/g", "std": "GB 2726-2016", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 2726-2016", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    "熟肉制品 - 油炸肉制品 (GB 2726/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 10000, "M": 100000, "unit": "CFU/g", "std": "GB 2726-2016", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 2726-2016", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    # 2. 乳制品类 (GB 19645 / GB 19302 / GB 19644 / GB 5420 / GB 29921)
    "乳制品 - 巴氏杀菌乳 (GB 19645/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 50000, "M": 100000, "unit": "CFU/mL", "std": "GB 19645-2010", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 1, "M": 5, "unit": "CFU/mL", "std": "GB 19645-2010", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g(mL)", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g(mL)", "std": "GB 29921-2021", "method": "GB 4789.10"},
        {"item": "单核细胞增生李斯特氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g(mL)", "std": "GB 29921-2021", "method": "GB 4789.30"}
    ],
    "乳制品 - 灭菌乳 (GB 25190 商业无菌)": [
        {"item": "商业无菌", "n": 1, "c": 0, "m": 0, "M": 0, "unit": "项", "std": "GB 25190-2010", "method": "GB 4789.26"}
    ],
    "乳制品 - 调制乳 (GB 25191/29921)": [
        {"item": "菌落总数 (限巴氏杀菌工艺)", "n": 5, "c": 2, "m": 50000, "M": 100000, "unit": "CFU/mL", "std": "GB 25191-2010", "method": "GB 4789.2"},
        {"item": "大肠菌群 (限巴氏杀菌工艺)", "n": 5, "c": 2, "m": 1, "M": 5, "unit": "CFU/mL", "std": "GB 25191-2010", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25mL", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25mL", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    "乳制品 - 发酵乳 / 酸奶 (GB 19302/29921)": [
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 1, "M": 5, "unit": "CFU/g(mL)", "std": "GB 19302-2010", "method": "GB 4789.3"},
        {"item": "酵母", "n": 1, "c": 0, "m": 100, "M": 100, "unit": "CFU/g(mL)", "std": "GB 19302-2010", "method": "GB 4789.15"},
        {"item": "霉菌", "n": 1, "c": 0, "m": 30, "M": 30, "unit": "CFU/g(mL)", "std": "GB 19302-2010", "method": "GB 4789.15"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g(mL)", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g(mL)", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    "乳制品 - 乳粉 / 奶粉 (GB 19644/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 50000, "M": 200000, "unit": "CFU/g", "std": "GB 19644-2010", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 90, "unit": "CFU/g", "std": "GB 19644-2010", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    "乳制品 - 干酪 / 再制干酪及干酪制品 (GB 5420/25192/29921)": [
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 5420-2021", "method": "GB 4789.3"},
        {"item": "酵母 (限原制干酪)", "n": 5, "c": 2, "m": 50, "M": 500, "unit": "CFU/g", "std": "GB 5420-2021", "method": "GB 4789.15"},
        {"item": "霉菌 (限原制干酪)", "n": 5, "c": 2, "m": 50, "M": 500, "unit": "CFU/g", "std": "GB 5420-2021", "method": "GB 4789.15"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 2, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"},
        {"item": "单核细胞增生李斯特氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.30"}
    ],
    "乳制品 - 固态成型乳制品 (奶片/奶条) (GB 29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 10000, "M": 50000, "unit": "CFU/g", "std": "行业明示标准", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "行业明示标准", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    # 3. 婴幼儿配方食品与辅食类 (GB 10765 / GB 10766 / GB 10767 / GB 10769)
    "婴幼儿配方食品 - 婴儿配方食品 (0-6月龄) (GB 10765)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 1000, "M": 10000, "unit": "CFU/g", "std": "GB 10765-2021", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 10765-2021", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 10765-2021", "method": "GB 4789.4"},
        {"item": "阪崎肠杆菌 (粉状乳粉)", "n": 3, "c": 0, "m": 0, "M": 0, "unit": "/100g", "std": "GB 10765-2021", "method": "GB 4789.40"}
    ],
    "婴幼儿配方食品 - 较大婴儿与幼儿配方食品 (GB 10766/10767)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 1000, "M": 10000, "unit": "CFU/g", "std": "GB 10766/10767", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 10766/10767", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 10766/10767", "method": "GB 4789.4"}
    ],
    "婴幼儿辅助食品 - 婴幼儿谷类辅助食品 (米粉/面条) (GB 10769/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 10000, "M": 50000, "unit": "CFU/g", "std": "GB 10769-2010", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 10769-2010", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    "婴幼儿辅助食品 - 婴幼儿罐装辅助食品 (泥糊状/颗粒片状/汁类) (GB 10770)": [
        {"item": "商业无菌", "n": 1, "c": 0, "m": 0, "M": 0, "unit": "项", "std": "GB 10770-2010", "method": "GB 4789.26"}
    ],
    # 4. 饮料与包装饮用水类 (GB 19298 / GB 7101 / GB 29921)
    "包装饮用水 - 包装饮用水 (纯净水/山泉水/其他饮用水) (GB 19298)": [
        {"item": "大肠菌群", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "CFU/mL", "std": "GB 19298-2014", "method": "GB 4789.3 平板法"},
        {"item": "铜绿假单胞菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "CFU/250mL", "std": "GB 19298-2014", "method": "GB 8538"}
    ],
    "包装饮用水 - 饮用天然矿泉水 (GB 8537)": [
        {"item": "大肠菌群", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "CFU/mL", "std": "GB 8537-2018", "method": "GB 8538"},
        {"item": "粪链球菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "CFU/250mL", "std": "GB 8537-2018", "method": "GB 8538"},
        {"item": "铜绿假单胞菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "CFU/250mL", "std": "GB 8537-2018", "method": "GB 8538"},
        {"item": "产气荚膜梭菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "CFU/50mL", "std": "GB 8537-2018", "method": "GB 8538"}
    ],
    "饮料类 - 碳酸饮料 (汽水) (GB 7101/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 100, "M": 10000, "unit": "CFU/mL", "std": "GB 7101-2022", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 1, "M": 10, "unit": "CFU/mL", "std": "GB 7101-2022", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25mL", "std": "GB 29921-2021", "method": "GB 4789.4"}
    ],
    "饮料类 - 茶饮料与果蔬汁饮料 (GB 7101/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 100, "M": 10000, "unit": "CFU/mL", "std": "GB 7101-2022", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 1, "M": 10, "unit": "CFU/mL", "std": "GB 7101-2022", "method": "GB 4789.3"},
        {"item": "霉菌与酵母", "n": 1, "c": 0, "m": 20, "M": 20, "unit": "CFU/mL", "std": "GB 7101-2022", "method": "GB 4789.15"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25mL", "std": "GB 29921-2021", "method": "GB 4789.4"}
    ],
    "饮料类 - 蛋白饮料 (植物蛋白/含乳饮料) (GB 7101/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 100, "M": 10000, "unit": "CFU/mL", "std": "GB 7101-2022", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 1, "M": 10, "unit": "CFU/mL", "std": "GB 7101-2022", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25mL", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/mL", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    "饮料类 - 固体饮料 (速溶咖啡/奶茶粉/果粉) (GB 7101/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 1000, "M": 50000, "unit": "CFU/g", "std": "GB 7101-2022", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 90, "unit": "CFU/g", "std": "GB 7101-2022", "method": "GB 4789.3"},
        {"item": "霉菌与酵母", "n": 1, "c": 0, "m": 50, "M": 50, "unit": "CFU/g", "std": "GB 7101-2022", "method": "GB 4789.15"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"}
    ],
    # 5. 方便食品与调味面制品类 (GB 17400 / GB 2714 / GB 19640 / GB 29921)
    "方便食品 - 油炸方便面 / 方便米粉 (GB 17400/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 10000, "M": 100000, "unit": "CFU/g", "std": "GB 17400-2015", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 17400-2015", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    "方便食品 - 调味面制品 (辣条) (GB 2714/29921/2026细则)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 10000, "M": 100000, "unit": "CFU/g", "std": "产品明示标准", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "产品明示标准", "method": "GB 4789.3"},
        {"item": "霉菌", "n": 1, "c": 0, "m": 50, "M": 50, "unit": "CFU/g", "std": "产品明示标准", "method": "GB 4789.15"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    "方便食品 - 冲调类方便食品 (麦片/藕粉/黑芝麻糊) (GB 19640/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 10000, "M": 100000, "unit": "CFU/g", "std": "GB 19640-2016", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 19640-2016", "method": "GB 4789.3"},
        {"item": "霉菌", "n": 1, "c": 0, "m": 50, "M": 50, "unit": "CFU/g", "std": "GB 19640-2016", "method": "GB 4789.15"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"}
    ],
    # 6. 糕点、面包与饼干类 (GB 7099 / GB 7100 / GB 29921)
    "糕点与面包 - 热加工糕点/面包 (烘烤类/油炸类/水蒸类) (GB 7099/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 10000, "M": 100000, "unit": "CFU/g", "std": "GB 7099-2015", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 7099-2015", "method": "GB 4789.3"},
        {"item": "霉菌", "n": 1, "c": 0, "m": 150, "M": 150, "unit": "CFU/g", "std": "GB 7099-2015", "method": "GB 4789.15"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    "糕点与面包 - 冷加工糕点 (裱花蛋糕/慕斯) (GB 7099/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 10000, "M": 100000, "unit": "CFU/g", "std": "GB 7099-2015", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 7099-2015", "method": "GB 4789.3"},
        {"item": "霉菌", "n": 1, "c": 0, "m": 100, "M": 100, "unit": "CFU/g", "std": "GB 7099-2015", "method": "GB 4789.15"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"},
        {"item": "单核细胞增生李斯特氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.30"}
    ],
    "饼干 - 饼干 (酥性/韧性/发酵/曲奇/夹心) (GB 7100/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 1000, "M": 10000, "unit": "CFU/g", "std": "GB 7100-2015", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 7100-2015", "method": "GB 4789.3"},
        {"item": "霉菌", "n": 1, "c": 0, "m": 50, "M": 50, "unit": "CFU/g", "std": "GB 7100-2015", "method": "GB 4789.15"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    # 7. 冷冻饮品与速冻食品类 (GB 2759 / GB 19295 / GB 29921)
    "冷冻饮品 - 冰淇淋 / 雪糕 / 甜味冰 (GB 2759/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 25000, "M": 100000, "unit": "CFU/g", "std": "GB 2759-2015", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 2759-2015", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"},
        {"item": "单核细胞增生李斯特氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.30"}
    ],
    "速冻食品 - 速冻面米熟制品 (水饺/包子/烧麦) (GB 19295/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 10000, "M": 100000, "unit": "CFU/g", "std": "GB 19295-2021", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 19295-2021", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    "速冻食品 - 速冻调制水产制品 (鱼丸/虾丸) (GB 19295/29921)": [
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 19295-2021", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"},
        {"item": "副溶血性弧菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "MPN/g (或 CFU/g)", "std": "GB 29921-2021", "method": "GB 4789.7"}
    ],
    "速冻食品 - 速冻调制肉制品 (肉丸/肉串) (GB 19295/29921)": [
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 19295-2021", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    # 8. 水产制品与生食水产类 (GB 10136 / GB 31607 / GB 29921)
    "水产制品 - 熟制动物性水产制品 (熟虾/熟蟹/即食鱼干) (GB 10136/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 30000, "M": 100000, "unit": "CFU/g", "std": "GB 10136-2015", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 10136-2015", "method": "GB 4789.3"},
        {"item": "副溶血性弧菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "MPN/g (或 CFU/g)", "std": "GB 29921-2021", "method": "GB 4789.7"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "单核细胞增生李斯特氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.30"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    "水产制品 - 生食动物性水产品 (生鱼片/三文鱼刺身/生蚝) (GB 10136/29921/31607)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 50000, "M": 100000, "unit": "CFU/g", "std": "GB 10136-2015", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 10136-2015", "method": "GB 4789.3"},
        {"item": "副溶血性弧菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "MPN/g (或 CFU/g)", "std": "GB 29921-2021", "method": "GB 4789.7"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "单核细胞增生李斯特氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.30"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    # 9. 蔬菜制品、水果制品与酱腌菜类 (GB 2714 / GB 14884 / GB 29921)
    "蔬菜制品 - 酱腌菜 / 泡菜 / 榨菜 (GB 2714/29921)": [
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 1000, "unit": "CFU/g", "std": "GB 2714-2015", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"},
        {"item": "大肠埃希氏菌 O157:H7", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.36"}
    ],
    "水果制品 - 蜜饯凉果 (果脯/话梅/蜜枣) (GB 14884/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 1000, "M": 10000, "unit": "CFU/g", "std": "GB 14884-2016", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 14884-2016", "method": "GB 4789.3"},
        {"item": "霉菌", "n": 1, "c": 0, "m": 50, "M": 50, "unit": "CFU/g", "std": "GB 14884-2016", "method": "GB 4789.15"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    # 10. 豆制品类 (GB 2712 / GB 29921)
    "豆制品 - 发酵性豆制品 (腐乳/豆豉/纳豆) (GB 2712/29921)": [
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 2712-2014", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    "豆制品 - 非发酵性豆制品 (豆腐干/素鸡/腐竹/豆皮) (GB 2712/29921)": [
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 2712-2014", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    # 11. 蛋制品类 (GB 2749 / GB 29921)
    "蛋制品 - 再制蛋与熟制蛋 (皮蛋/咸鸭蛋/卤蛋) (GB 2749/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 10000, "M": 100000, "unit": "CFU/g", "std": "GB 2749-2015", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 2749-2015", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"}
    ],
    # 12. 糖果与巧克力制品类 (GB 17399 / GB 9678.2 / GB 29921)
    "糖果与巧克力 - 糖果 (硬糖/软糖/胶基糖果) (GB 17399/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 1000, "M": 10000, "unit": "CFU/g", "std": "GB 17399-2016", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 17399-2016", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    "糖果与巧克力 - 巧克力及代可可脂巧克力 (GB 9678.2/29921)": [
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"}
    ],
    # 13. 调味品类 (GB 2717 / GB 2718 / GB 31644 / GB 29921)
    "调味品 - 酱油 (GB 2717/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 5000, "M": 50000, "unit": "CFU/mL", "std": "GB 2717-2018", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/mL", "std": "GB 2717-2018", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25mL", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/mL", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    "调味品 - 酿造酱 / 复合调味料 / 火锅底料 (GB 2718/31644/29921)": [
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g(mL)", "std": "GB 2718/31644", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g(mL)", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g(mL)", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    # 14. 薯类和膨化食品 (GB 17401 / GB 29921)
    "薯类和膨化食品 - 膨化食品 (油炸/非油炸型) (GB 17401/29921)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 10000, "M": 100000, "unit": "CFU/g", "std": "GB 17401-2014", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 17401-2014", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    # 15. 坚果与籽类食品 (GB 19300 / GB 29921)
    "坚果与籽类食品 - 熟制坚果与籽类 (瓜子/花生/核桃) (GB 19300/29921)": [
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 19300-2014", "method": "GB 4789.3"},
        {"item": "霉菌", "n": 1, "c": 0, "m": 25, "M": 25, "unit": "CFU/g", "std": "GB 19300-2014", "method": "GB 4789.15"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"}
    ],
    # 16. 蜂产品 (GB 14963)
    "蜂产品 - 蜂蜜 (GB 14963)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 1000, "M": 5000, "unit": "CFU/g", "std": "GB 14963-2011", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 90, "unit": "CFU/g", "std": "GB 14963-2011", "method": "GB 4789.3"},
        {"item": "霉菌计数", "n": 1, "c": 0, "m": 200, "M": 200, "unit": "CFU/g", "std": "GB 14963-2011", "method": "GB 4789.15"},
        {"item": "嗜渗酵母计数", "n": 1, "c": 0, "m": 200, "M": 200, "unit": "CFU/g", "std": "GB 14963-2011", "method": "GB 4789.15"}
    ],
    # 17. 淀粉及淀粉制品 (GB 31637 / GB 29921)
    "淀粉制品 - 粉丝 / 粉条 / 粉皮 (GB 29921/2026细则)": [
        {"item": "菌落总数", "n": 5, "c": 2, "m": 10000, "M": 100000, "unit": "CFU/g", "std": "产品明示标准", "method": "GB 4789.2"},
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "产品明示标准", "method": "GB 4789.3"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 29921-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 29921-2021", "method": "GB 4789.10"}
    ],
    # 18. 食用油脂制品 (GB 15196)
    "食用油脂制品 - 人造奶油 / 起酥油 (GB 15196)": [
        {"item": "大肠菌群", "n": 5, "c": 2, "m": 10, "M": 100, "unit": "CFU/g", "std": "GB 15196-2018", "method": "GB 4789.3"},
        {"item": "霉菌", "n": 1, "c": 0, "m": 10, "M": 10, "unit": "CFU/g", "std": "GB 15196-2018", "method": "GB 4789.15"}
    ],
    # 19. 餐饮食品与餐饮具类 (GB 31607 / GB 14934)
    "餐饮食品 - 餐饮自制熟肉 (散装即食) (GB 31607)": [
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 31607-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 31607-2021", "method": "GB 4789.10"},
        {"item": "单核细胞增生李斯特氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 31607-2021", "method": "GB 4789.30"}
    ],
    "餐饮食品 - 餐饮自制生食水产品 (刺身/生蚝) (GB 31607)": [
        {"item": "副溶血性弧菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "MPN/g (或 CFU/g)", "std": "GB 31607-2021", "method": "GB 4789.7"},
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 31607-2021", "method": "GB 4789.4"},
        {"item": "单核细胞增生李斯特氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g", "std": "GB 31607-2021", "method": "GB 4789.30"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g", "std": "GB 31607-2021", "method": "GB 4789.10"}
    ],
    "餐饮食品 - 餐饮自制糕点与果蔬汁冷饮 (GB 31607)": [
        {"item": "沙门氏菌", "n": 5, "c": 0, "m": 0, "M": 0, "unit": "/25g(mL)", "std": "GB 31607-2021", "method": "GB 4789.4"},
        {"item": "金黄色葡萄球菌", "n": 5, "c": 1, "m": 100, "M": 1000, "unit": "CFU/g(mL)", "std": "GB 31607-2021", "method": "GB 4789.10"}
    ],
    "餐饮具 - 复用餐饮具 (餐馆自行/集中清洗消毒) (GB 14934)": [
        {"item": "大肠菌群", "n": 1, "c": 0, "m": 0, "M": 0, "unit": "/50cm²", "std": "GB 14934-2016", "method": "GB 14934 纸片法/发酵法"},
        {"item": "沙门氏菌", "n": 1, "c": 0, "m": 0, "M": 0, "unit": "/50cm²", "std": "GB 14934-2016", "method": "GB 4789.4"}
    ]
}

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
# ============================================================
GB2760_DIFF_DATABASE = {
    "脱氢乙酸及其钠盐 (脱氢乙酸钠)": {
        "function": "防腐剂",
        "ins": "266",
        "status": "🚨 重大禁用红线",
        "summary": "GB 2760-2024全面删除其在黄油/浓缩黄油、面包、糕点、焙烤食品馅料及表面用挂浆、熟肉制品中的使用规定！",
        "diff_items": [
            {
                "food_cat": "糕点 (烘烤糕点、月饼、蛋糕等)",
                "gb2014": "允许使用，最大使用量 0.5 g/kg",
                "gb2024": "❌ 【全面删除禁用】不再允许在糕点中使用！",
                "risk_level": "RED",
                "advice": "烘焙企业必须全面切换为山梨酸钾、丙酸钙、双乙酸钠等合规防腐剂！"
            },
            {
                "food_cat": "面包",
                "gb2014": "允许使用，最大使用量 0.5 g/kg",
                "gb2024": "❌ 【全面删除禁用】不再允许在面包中使用！",
                "risk_level": "RED",
                "advice": "建议改用丙酸钙 (最大2.5g/kg) 或山梨酸钾。"
            },
            {
                "food_cat": "熟肉制品 (酱卤肉、熏烧烤肉、火腿)",
                "gb2014": "允许使用，最大使用量 0.5 g/kg",
                "gb2024": "❌ 【全面删除禁用】不再允许在熟肉制品中使用！",
                "risk_level": "RED",
                "advice": "熟肉制品严禁添加，应使用山梨酸钾、乳酸链球菌素(Nisin)、双乙酸钠。"
            },
            {
                "food_cat": "腌渍的蔬菜 (酱菜)",
                "gb2014": "允许使用，最大使用量 1.0 g/kg",
                "gb2024": "✔ 允许使用，最大使用量调整为 0.3 g/kg (限量大幅下调)",
                "risk_level": "YELLOW",
                "advice": "注意下调配方添加量，确保实测残留 ≤ 0.3 g/kg。"
            },
            {
                "food_cat": "发酵豆制品 (腐乳)",
                "gb2014": "允许使用，最大使用量 0.3 g/kg",
                "gb2024": "✔ 允许使用，最大使用量 0.3 g/kg (维持不变)",
                "risk_level": "GREEN",
                "advice": "按原配方控制添加量。"
            }
        ]
    },
    "落葵红 (天然色素)": {
        "function": "着色剂",
        "ins": "-",
        "status": "⚠️ 使用范围收窄",
        "summary": "GB 2760-2024删除了落葵红在调味品、糕点等食品中的使用规定。",
        "diff_items": [
            {
                "food_cat": "酱油、复合调味料",
                "gb2014": "允许适量使用",
                "gb2024": "❌ 【删除禁用】不再允许在调味品中使用！",
                "risk_level": "RED",
                "advice": "改用红曲红、焦糖色等允许添加剂。"
            },
            {
                "food_cat": "配制酒",
                "gb2014": "允许使用，最大限量 0.05 g/kg",
                "gb2024": "✔ 允许使用，最大限量 0.05 g/kg",
                "risk_level": "GREEN",
                "advice": "配制酒中可继续使用。"
            }
        ]
    },
    "甜蜜素 (环己基氨基磺酸钠)": {
        "function": "甜味剂",
        "ins": "952",
        "status": "⚠️ 限量与范围严控",
        "summary": "明确严禁在白酒、面包、糕点中超范围使用，严格控制与糖精钠等比例之和。",
        "diff_items": [
            {
                "food_cat": "白酒",
                "gb2014": "❌ 严禁添加 (不得检出)",
                "gb2024": "❌ 严禁添加 (不得检出) - 严查超范围添加",
                "risk_level": "RED",
                "advice": "白酒中严禁添加任何甜味剂。"
            },
            {
                "food_cat": "冷冻饮品 (冰淇淋、雪糕)",
                "gb2014": "允许使用，最大 0.65 g/kg",
                "gb2024": "✔ 允许使用，最大 0.65 g/kg",
                "risk_level": "GREEN",
                "advice": "需注意与三氯蔗糖、安赛蜜等同类甜味剂计算比例之和！"
            },
            {
                "food_cat": "调味面制品 (辣条)",
                "gb2014": "允许使用，最大 1.6 g/kg",
                "gb2024": "✔ 允许使用，最大 1.6 g/kg",
                "risk_level": "GREEN",
                "advice": "与阿斯巴甜、纽甜等复配时注意限量。"
            }
        ]
    },
    "阿斯巴甜 (天门冬酰苯丙氨酸甲酯)": {
        "function": "甜味剂",
        "ins": "951",
        "status": "ℹ️ 强制标示警示语",
        "summary": "依据新规与GB 7718，添加阿斯巴甜的产品必须在配料表中醒目标注'阿斯巴甜(含苯丙氨酸)'。",
        "diff_items": [
            {
                "food_cat": "饮料类、风味发酵乳",
                "gb2014": "允许使用，按标准限量添加",
                "gb2024": "✔ 允许使用，且标签配料表必须强制醒目标示'阿斯巴甜(含苯丙氨酸)'！",
                "risk_level": "YELLOW",
                "advice": "包装更新时必须核对警示语完整性。"
            }
        ]
    },
    "三氯蔗糖 (蔗糖素)": {
        "function": "甜味剂",
        "ins": "955",
        "status": "✔ 部分品类范围微调",
        "summary": "高倍天然转化甜味剂，在减糖健康食品中广泛使用。",
        "diff_items": [
            {
                "food_cat": "风味发酵乳",
                "gb2014": "允许使用，最大 0.3 g/kg",
                "gb2024": "✔ 允许使用，最大 0.3 g/kg",
                "risk_level": "GREEN",
                "advice": "无糖酸奶中重点使用。"
            },
            {
                "food_cat": "烘焙食品 / 糕点",
                "gb2014": "允许使用，最大 0.25 g/kg",
                "gb2024": "✔ 允许使用，最大 0.25 g/kg",
                "risk_level": "GREEN",
                "advice": "耐高温特性优良，适合烘焙减糖。"
            }
        ]
    }
}

# ============================================================
# GB 2762 / GB 2761 浓缩物与干制品脱水折算知识库
# ============================================================
DEHYDRATION_STANDARDS_DB = {
    "干香菇 / 干食用菌": {
        "fresh_name": "鲜食用菌及其制品",
        "std": "GB 2762-2022",
        "default_dehydration_rate": 88.0,
        "default_conc_factor": 8.33,
        "items": [
            {"item": "铅 (Pb)", "fresh_limit": 0.1, "unit": "mg/kg", "method": "GB 5009.12"},
            {"item": "镉 (Cd)", "fresh_limit": 0.2, "unit": "mg/kg", "method": "GB 5009.15"},
            {"item": "总汞 (Hg)", "fresh_limit": 0.1, "unit": "mg/kg", "method": "GB 5009.17"},
            {"item": "总砷 (As)", "fresh_limit": 0.5, "unit": "mg/kg", "method": "GB 5009.11"}
        ]
    },
    "脱水蔬菜 / 干辣椒": {
        "fresh_name": "新鲜蔬菜 (叶菜/茄果类)",
        "std": "GB 2762-2022",
        "default_dehydration_rate": 85.0,
        "default_conc_factor": 6.67,
        "items": [
            {"item": "铅 (Pb)", "fresh_limit": 0.1, "unit": "mg/kg", "method": "GB 5009.12"},
            {"item": "镉 (Cd)", "fresh_limit": 0.05, "unit": "mg/kg", "method": "GB 5009.15"},
            {"item": "总汞 (Hg)", "fresh_limit": 0.01, "unit": "mg/kg", "method": "GB 5009.17"},
            {"item": "总砷 (As)", "fresh_limit": 0.5, "unit": "mg/kg", "method": "GB 5009.11"}
        ]
    },
    "牛肉干 / 猪肉干 (干制肉制品)": {
        "fresh_name": "生鲜畜肉类",
        "std": "GB 2762-2022",
        "default_dehydration_rate": 65.0,
        "default_conc_factor": 2.85,
        "items": [
            {"item": "铅 (Pb)", "fresh_limit": 0.2, "unit": "mg/kg", "method": "GB 5009.12"},
            {"item": "镉 (Cd)", "fresh_limit": 0.1, "unit": "mg/kg", "method": "GB 5009.15"},
            {"item": "总汞 (Hg)", "fresh_limit": 0.05, "unit": "mg/kg", "method": "GB 5009.17"},
            {"item": "总砷 (As)", "fresh_limit": 0.5, "unit": "mg/kg", "method": "GB 5009.11"}
        ]
    },
    "浓缩苹果汁 / 浓缩果汁": {
        "fresh_name": "新鲜水果 / 原榨果汁",
        "std": "GB 2762-2022 / GB 2761-2022",
        "default_dehydration_rate": 83.3,
        "default_conc_factor": 6.0,
        "items": [
            {"item": "铅 (Pb)", "fresh_limit": 0.05, "unit": "mg/kg", "method": "GB 5009.12"},
            {"item": "镉 (Cd)", "fresh_limit": 0.05, "unit": "mg/kg", "method": "GB 5009.15"},
            {"item": "展青霉素", "fresh_limit": 50.0, "unit": "μg/kg", "method": "GB 5009.185"}
        ]
    },
    "全脂乳粉 / 脱脂乳粉": {
        "fresh_name": "生乳 (生鲜牛奶)",
        "std": "GB 2762-2022 / GB 2761-2022",
        "default_dehydration_rate": 87.5,
        "default_conc_factor": 8.0,
        "items": [
            {"item": "铅 (Pb)", "fresh_limit": 0.02, "unit": "mg/kg", "method": "GB 5009.12"},
            {"item": "黄曲霉毒素 M1", "fresh_limit": 0.5, "unit": "μg/kg", "method": "GB 5009.24"}
        ]
    },
    "干海参 / 水产干制品": {
        "fresh_name": "鲜活水产品 (鱼虾贝类)",
        "std": "GB 2762-2022",
        "default_dehydration_rate": 80.0,
        "default_conc_factor": 5.0,
        "items": [
            {"item": "铅 (Pb)", "fresh_limit": 0.5, "unit": "mg/kg", "method": "GB 5009.12"},
            {"item": "镉 (Cd)", "fresh_limit": 0.5, "unit": "mg/kg", "method": "GB 5009.15"},
            {"item": "无机砷 (As)", "fresh_limit": 0.5, "unit": "mg/kg", "method": "GB 5009.11"},
            {"item": "甲基汞 (Hg)", "fresh_limit": 0.5, "unit": "mg/kg", "method": "GB 5009.17"}
        ]
    }
}

# ============================================================
# GB 2763 农药残留全景矩阵知识库
# ============================================================
PESTICIDE_MATRIX_DB = {
    "芹菜": [
        {"item": "毒死蜱", "type": "杀虫剂", "limit": "0.05", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.113", "risk": "重点监控高风险项"},
        {"item": "噻虫胺", "type": "杀虫剂", "limit": "1.0", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.121", "risk": "国抽高发超标项"},
        {"item": "甲拌磷", "type": "禁用高毒农药", "limit": "不得检出 (0.01)", "unit": "mg/kg", "is_banned": True, "method": "GB 23200.113", "risk": "🚨 国家明令禁用高毒农药"},
        {"item": "氟虫腈", "type": "禁用高毒农药", "limit": "不得检出 (0.02)", "unit": "mg/kg", "is_banned": True, "method": "GB 23200.113", "risk": "🚨 国家明令禁用高毒农药"},
        {"item": "辛硫磷", "type": "杀虫剂", "limit": "0.05", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.113", "risk": "常用有机磷"},
        {"item": "啶虫脒", "type": "杀虫剂", "limit": "5.0", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.113", "risk": "常见烟碱类"},
        {"item": "百菌清", "type": "杀菌剂", "limit": "5.0", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.113", "risk": "广谱保护性杀菌剂"}
    ],
    "韭菜": [
        {"item": "腐霉利", "type": "杀菌剂", "limit": "5.0", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.113", "risk": "韭菜灰霉病重点防控项(新标5.0)"},
        {"item": "毒死蜱", "type": "杀虫剂", "limit": "0.02", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.113", "risk": "韭蛆防治严查项"},
        {"item": "甲拌磷", "type": "禁用高毒农药", "limit": "不得检出 (0.01)", "unit": "mg/kg", "is_banned": True, "method": "GB 23200.113", "risk": "🚨 国家明令禁用高毒农药"},
        {"item": "氧乐果", "type": "禁用高毒农药", "limit": "不得检出 (0.02)", "unit": "mg/kg", "is_banned": True, "method": "GB 23200.113", "risk": "🚨 国家明令禁用高毒农药"},
        {"item": "敌敌畏", "type": "杀虫剂", "limit": "0.2", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.113", "risk": "有机磷杀虫剂"},
        {"item": "啶虫脒", "type": "杀虫剂", "limit": "1.0", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.113", "risk": "蚜虫防治项"}
    ],
    "草莓": [
        {"item": "烯酰吗啉", "type": "杀菌剂", "limit": "3.0", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.113", "risk": "霜霉病防治"},
        {"item": "吡虫啉", "type": "杀虫剂", "limit": "0.5", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.113", "risk": "烟碱类"},
        {"item": "嘧霉胺", "type": "杀菌剂", "limit": "5.0", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.113", "risk": "灰霉病常用药"},
        {"item": "肟菌酯", "type": "杀菌剂", "limit": "3.0", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.113", "risk": "白粉病"},
        {"item": "联苯肼酯", "type": "杀螨剂", "limit": "7.0", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.113", "risk": "红蜘蛛防治"}
    ],
    "豇豆 / 菜豆": [
        {"item": "克百威", "type": "禁用高毒农药", "limit": "不得检出 (0.02)", "unit": "mg/kg", "is_banned": True, "method": "GB 23200.113", "risk": "🚨 蔬菜禁用高毒农药"},
        {"item": "灭蝇胺", "type": "杀虫剂", "limit": "0.5", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.113", "risk": "斑潜蝇防治重点监控"},
        {"item": "噻虫胺", "type": "杀虫剂", "limit": "0.01", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.121", "risk": "极低限量红线(0.01)"},
        {"item": "乙酰甲胺磷", "type": "杀虫剂", "limit": "0.02", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.113", "risk": "严控残留"},
        {"item": "吡虫啉", "type": "杀虫剂", "limit": "2.0", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.113", "risk": "蓟马与蚜虫防治"}
    ],
    "柑橘": [
        {"item": "水胺硫磷", "type": "禁用高毒农药", "limit": "不得检出 (0.05)", "unit": "mg/kg", "is_banned": True, "method": "GB 23200.113", "risk": "🚨 果树禁用高毒农药"},
        {"item": "联苯菊酯", "type": "杀虫剂", "limit": "0.05", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.113", "risk": "木虱防治"},
        {"item": "苯醚甲环唑", "type": "杀菌剂", "limit": "0.5", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.113", "risk": "炭疽病与疮痂病"},
        {"item": "丙溴磷", "type": "杀虫剂", "limit": "0.2", "unit": "mg/kg", "is_banned": False, "method": "GB 23200.113", "risk": "红蜘蛛防治"}
    ]
}

# ============================================================
# GUI主程序与交互对话框
# ============================================================

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
    "小米": ["谷物及其制品", "谷物"],
    "小麦粉": ["小麦粉", "面粉", "通用小麦粉", "专用小麦粉", "谷物及其制品"],
    "面粉": ["小麦粉", "面粉", "谷物及其制品"],
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

def auto_match_record(standards, global_index, food_query, item_query, measured_val=None):
    """
    输入单行样品名称、检测项目（及可选实测值），跨 6 大国标全量库自动匹配并反写：
    返回 (标准代号, 标准全称, 限量要求, 检验方法, 检出限LOD, 实测值, 判定结论, 匹配状态, 原始匹配分类)
    """
    raw_f = str(food_query).strip() if food_query is not None else ""
    raw_i = str(item_query).strip() if item_query is not None else ""
    if not raw_f and not raw_i:
        return ("/", "/", "/", "/", "/", "/", "/", "空白行", "/")
        
    clean_i = re.sub(r'[\(（].*?[\)）]', '', raw_i).strip().lower()
    std_item_alias = ITEM_ALIAS_MAP.get(clean_i, clean_i)
    
    search_keys = [clean_i, std_item_alias.lower(), raw_i.lower()]
    candidates = []
    if global_index:
        for sk in search_keys:
            if sk in global_index:
                candidates.extend(global_index[sk])
        if not candidates:
            for k, v in global_index.items():
                if clean_i and (clean_i in k or k in clean_i):
                    candidates.extend(v)
                    
    clean_f = re.sub(r'[\(（].*?[\)）]', '', raw_f).strip()
    clean_f_core = re.sub(r'^(新鲜|散装|有机|绿色|无公害|即食|预包装|冷冻|速冻|市售)\s*', '', clean_f).strip()
    food_candidates = [raw_f, clean_f, clean_f_core]
    if clean_f_core in FOOD_ALIAS_MAP:
        food_candidates.extend(FOOD_ALIAS_MAP[clean_f_core])
    elif clean_f in FOOD_ALIAS_MAP:
        food_candidates.extend(FOOD_ALIAS_MAP[clean_f])
    elif raw_f in FOOD_ALIAS_MAP:
        food_candidates.extend(FOOD_ALIAS_MAP[raw_f])
        
    best_cand = None
    best_score = -1
    
    for std_code, it in candidates:
        cand_cat = it.get('food_category', '').lower()
        cand_item = it.get('test_item', '').lower()
        score = 0
        if clean_i and clean_i == re.sub(r'[\(（].*?[\)）]', '', cand_item).strip():
            score += 50
        elif clean_i in cand_item:
            score += 40
            
        food_hit = False
        for fc in food_candidates:
            if not fc:
                continue
            fc_lower = fc.lower()
            if fc_lower == cand_cat:
                score += 50
                food_hit = True
                break
            elif fc_lower in cand_cat or cand_cat in fc_lower:
                score += 35
                food_hit = True
                break
            subs = re.split(r'[,，、/()（）\s]+', cand_cat)
            if fc_lower in subs:
                score += 45
                food_hit = True
                break
                
        if food_hit and score > best_score:
            best_score = score
            best_cand = (std_code, it)
            
    if best_cand and best_score >= 70:
        std_code, it = best_cand
        std_name = standards[std_code]['name'] if std_code in standards else "食品安全标准"
        limit_val = it.get('limit_value', '')
        unit = it.get('unit', '')
        limit_str = f"{limit_val} {unit}".strip() if limit_val else "详见标准备注"
        method_str = it.get('method', '') or resolve_item_test_method("", raw_i, raw_f, std_code)
        lod_str = resolve_item_lod(raw_i, method_str, limit_val, unit)
        
        meas_display = "/"
        judge_display = "/"
        if measured_val is not None and str(measured_val).strip() not in ("", "/", "None", "--"):
            meas_display = str(measured_val).strip()
            j_status, j_text = judge_measured_value(limit_val, meas_display)
            judge_display = j_text
            
        status_tag = "✔ 精准匹配" if best_score >= 85 else "⚠️ 品类推导匹配"
        return (std_code, std_name, limit_str, method_str, lod_str, meas_display, judge_display, status_tag, it.get('food_category'))
    else:
        # 特殊情况处理：GB 2760 禁用添加剂判定 (如脱氢乙酸在糕点、面包、熟肉中禁用)
        if "脱氢乙酸" in raw_i and any(w in raw_f for w in ["糕点", "面包", "熟肉", "面条", "馒头"]):
            std_code = "GB 2760-2024"
            std_name = "食品添加剂使用标准 (2024新版)"
            limit_str = "不得使用 (0 mg/kg)"
            method_str = "GB 5009.121-2016"
            lod_str = "0.05 mg/kg (液质) / 2.0 mg/kg (气相/液相)"
            meas_display = "/"
            judge_display = "/"
            if measured_val is not None and str(measured_val).strip() not in ("", "/", "None", "--"):
                meas_display = str(measured_val).strip()
                j_status, j_text = judge_measured_value("不得使用", meas_display)
                judge_display = j_text
            return (std_code, std_name, limit_str, method_str, lod_str, meas_display, judge_display, "✔ 精准匹配(2024禁用)", "GB 2760 禁用食品类别")
            
        method_str = resolve_item_test_method("", raw_i, raw_f, "")
        lod_str = resolve_item_lod(raw_i, method_str, "", "")
        meas_display = str(measured_val).strip() if (measured_val and str(measured_val).strip() not in ("", "/", "None", "--")) else "/"
        return ("未查见对应标准", "/", "/", method_str if method_str != "/" else "参照推荐检验方法", lod_str, meas_display, "/", "❓ 未查见对应标准限量", "/")


class FoodStandardChecker:
    def __init__(self, root):
        self.root = root
        self.root.title("食品安全国家标准判定核查工具（完整版）")
        self.root.geometry("1100x750")
        self.root.minsize(900, 600)
        
        # 加载持久化配置并恢复窗口大小与位置
        self.ui_config = load_ui_config()
        if 'window_geometry' in self.ui_config:
            try:
                self.root.geometry(self.ui_config['window_geometry'])
            except Exception:
                pass
        
        # 窗口关闭协议：自动保存列宽记忆
        self.root.protocol("WM_DELETE_WINDOW", self.on_window_close)
        
        # 加载数据
        self.loading_label = tb.Label(root, text="正在加载数据...", font=("Microsoft YaHei", 14))
        self.loading_label.pack(expand=True)
        self.root.update()
        
        self.standards = load_standards()
        self.global_index = build_global_item_index(self.standards)
        
        self.loading_label.destroy()
        
        total_items = sum(len(v['items']) for v in self.standards.values())
        
        # 统计信息
        self.root.title(f"食品安全国家标准判定核查工具 V1.2 - 共 {total_items} 条限量数据  |  Powered by 陈耿忠")
        
        # 判定历史与分类筛选
        self.judge_history = []
        self.selected_category_filter = ""
        self.tree_sort_states = {}
        
        # 构建UI
        self.setup_ui()
    
    def setup_ui(self):
        # 顶部信息栏
        top_frame = tb.Frame(self.root, height=40)
        top_frame.pack(fill=tk.X)
        total_items = sum(len(v['items']) for v in self.standards.values())
        info_text = f"  |  标准: {len(self.standards)}个  |  限量数据: {total_items}条"
        tb.Label(top_frame, text=info_text, 
                font=("Microsoft YaHei", 10), anchor="w").pack(side=tk.LEFT, padx=10, pady=8)
        tb.Label(top_frame, text="⚡ Powered by 陈耿忠", 
                font=("Microsoft YaHei", 9, "bold"), bootstyle="primary").pack(side=tk.RIGHT, padx=14, pady=8)
        
        # 底部状态栏与版权署名
        status_bar = tb.Frame(self.root, padding=(10, 2))
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)
        tb.Label(status_bar, text="🛡️ 食品安全国家标准判定核查系统 V1.2 (GB 2760 / GB 2761 / GB 2762 / GB 2763 / GB 29921)", 
                 font=("Microsoft YaHei", 8), bootstyle="secondary").pack(side=tk.LEFT)
        tb.Label(status_bar, text="⚡ Powered by 陈耿忠", 
                 font=("Microsoft YaHei", 8, "bold"), bootstyle="primary").pack(side=tk.RIGHT)
        
        # ========== Notebook 分页 ==========
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # ---- Tab 1: 标准查询 ----
        self.tab_query = tb.Frame(self.notebook)
        self.notebook.add(self.tab_query, text="标准查询")
        self.setup_query_tab()
        
        # ---- Tab 2: 样品匹配 ----
        self.tab_match = tb.Frame(self.notebook)
        self.notebook.add(self.tab_match, text="样品匹配")
        self.setup_match_tab()
        
        # ---- Tab 3: 多项目批量查询与判定 ----
        self.tab_batch = tb.Frame(self.notebook)
        self.notebook.add(self.tab_batch, text="多项目批量判定")
        self.setup_batch_tab()
        
        # ---- Tab 4: 台账智能反写 (Excel自动匹配) ----
        self.tab_excel_backfill = tb.Frame(self.notebook)
        self.notebook.add(self.tab_excel_backfill, text="台账智能反写(Excel自动匹配)")
        self.setup_excel_backfill_tab()
        
        # ---- Tab 4: 添加剂比例之和计算器 ----
        self.tab_ratio = tb.Frame(self.notebook)
        self.notebook.add(self.tab_ratio, text="添加剂比例之和计算")
        self.setup_ratio_tab()
        
        # ---- Tab 5: 商品名与标准反查分类 ----
        self.tab_inference = tb.Frame(self.notebook)
        self.notebook.add(self.tab_inference, text="商品名反查国标分类")
        self.setup_inference_tab()
        
        # ---- Tab 6: 微生物三级采样判定 ----
        self.tab_microbiology = tb.Frame(self.notebook)
        self.notebook.add(self.tab_microbiology, text="微生物三级采样判定")
        self.setup_microbiology_tab()

        # ---- Tab 7: GB 2760-2024 新旧国标对比 ----
        self.tab_gb2760_diff = tb.Frame(self.notebook)
        self.notebook.add(self.tab_gb2760_diff, text="GB 2760新旧国标对比")
        self.setup_gb2760_diff_tab()

        # ---- Tab 8: GB 2762 浓缩与干制品折算换算 ----
        self.tab_dehydration = tb.Frame(self.notebook)
        self.notebook.add(self.tab_dehydration, text="GB 2762脱水浓缩折算")
        self.setup_dehydration_ratio_tab()

        # ---- Tab 9: GB 2763 农药残留全景矩阵 ----
        self.tab_pesticide = tb.Frame(self.notebook)
        self.notebook.add(self.tab_pesticide, text="GB 2763农残全景矩阵")
        self.setup_pesticide_matrix_tab()
        
        # 底部：判定记录（共用）
        bottom_frame = tb.LabelFrame(self.root, text="判定历史")
        bottom_frame.pack(fill=tk.X, padx=10, pady=(5,10))
        
        self.history_text = tk.Text(bottom_frame, height=6, font=("Consolas", 9), wrap=tk.WORD)
        hist_scroll = ttk.Scrollbar(bottom_frame, orient=tk.VERTICAL, command=self.history_text.yview)
        self.history_text.configure(yscrollcommand=hist_scroll.set)
        self.history_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        hist_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 初始加载数据
        self.on_standard_selected()
    
    def setup_query_tab(self):
        """标准查询 Tab UI"""
        main_frame = tb.Frame(self.tab_query)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左侧：查询面板
        left_frame = tb.Frame(main_frame, width=390)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0,5))
        left_frame.pack_propagate(False)
        
        # 标准选择
        tb.Label(left_frame, text="标准选择", font=("Microsoft YaHei", 10, "bold"), anchor="w").pack(fill=tk.X, pady=(2,2))
        
        standard_frame = tb.Frame(left_frame)
        standard_frame.pack(fill=tk.X, pady=(0,4))
        
        self.standard_var = tk.StringVar()
        standard_codes = sorted(self.standards.keys())
        self.standard_combo = ttk.Combobox(standard_frame, textvariable=self.standard_var,
                                           values=standard_codes, state="readonly", font=("Microsoft YaHei", 10))
        self.standard_combo.pack(fill=tk.X)
        if standard_codes:
            self.standard_combo.current(0)
        self.standard_combo.bind('<<ComboboxSelected>>', self.on_standard_selected)
        
        # 标准信息
        self.standard_info_label = tb.Label(left_frame, text="", font=("Microsoft YaHei", 9), anchor="w", wraplength=370, justify=tk.LEFT)
        self.standard_info_label.pack(fill=tk.X, pady=(0,4))
        
        # 样品/食品类别下拉选择菜单 (宽幅下拉 + 智能过滤)
        cat_box = tb.LabelFrame(left_frame, text=" 🍽️ 样品/食品类别选择 (下拉菜单) ", padding=8)
        cat_box.pack(fill=tk.X, pady=(0, 6))
        
        self.selected_category_filter = ""
        self.cat_combo_var = tk.StringVar(value="全部食品分类 (显示全部)")
        self.cat_combo = ttk.Combobox(
            cat_box, 
            textvariable=self.cat_combo_var,
            state="readonly",
            font=("Microsoft YaHei", 10)
        )
        self.cat_combo.pack(fill=tk.X, pady=(0, 4))
        self.cat_combo.bind("<<ComboboxSelected>>", self.on_cat_combo_selected)
        
        cat_search_row = tb.Frame(cat_box)
        cat_search_row.pack(fill=tk.X, pady=(2, 0))
        tb.Label(cat_search_row, text="类别过滤:", font=("Microsoft YaHei", 8), bootstyle=SECONDARY).pack(side=tk.LEFT, padx=(0, 4))
        self.cat_search_var = tk.StringVar()
        self.cat_search_var.trace('w', lambda *args: self.filter_category_dropdown())
        cat_search_entry = tb.Entry(cat_search_row, textvariable=self.cat_search_var, font=("Microsoft YaHei", 9))
        cat_search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        tb.Button(cat_search_row, text="🔄 全部/重置", command=self.clear_category_filter, bootstyle="secondary-outline").pack(side=tk.RIGHT)
        
        # 搜索框
        tb.Label(left_frame, text="搜索（食品/检测项目）", font=("Microsoft YaHei", 10, "bold"), anchor="w").pack(fill=tk.X, pady=(2,2))
        
        search_frame = tb.Frame(left_frame)
        search_frame.pack(fill=tk.X, pady=(0,2))
        
        self.search_var = tk.StringVar()
        self.search_var.trace('w', lambda *args: self.on_search())
        search_entry = tb.Entry(search_frame, textvariable=self.search_var, font=("Microsoft YaHei", 10))
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tb.Button(search_frame, text="清除", command=self.clear_search, bootstyle=SECONDARY).pack(side=tk.RIGHT, padx=(3,0))
        
        # 结果统计
        self.result_count_label = tb.Label(left_frame, text="", font=("Microsoft YaHei", 9), anchor="w")
        self.result_count_label.pack(fill=tk.X, pady=(0,3))
        
        # 判定区
        tb.Label(left_frame, text="限量判定", font=("Microsoft YaHei", 10, "bold"), anchor="w").pack(fill=tk.X, pady=(4,2))
        
        judge_frame = tb.LabelFrame(left_frame, text="")
        judge_frame.pack(fill=tk.X)
        
        self.selected_info_label = tb.Label(judge_frame, text="请在右侧点击选择一条数据", 
                                           font=("Microsoft YaHei", 9), anchor="w", wraplength=350, justify=tk.LEFT)
        self.selected_info_label.pack(fill=tk.X, pady=(0,5))
        
        self.limit_display_label = tb.Label(judge_frame, text="限量值: --", 
                                           font=("Microsoft YaHei", 10, "bold"), bootstyle=DANGER, anchor="w")
        self.limit_display_label.pack(fill=tk.X)
        
        # 检验方法展示与自由选择复制区
        method_container = tb.Frame(judge_frame)
        method_container.pack(fill=tk.X, pady=(4, 2))
        
        m_header = tb.Frame(method_container)
        m_header.pack(fill=tk.X, pady=(0, 2))
        tb.Label(m_header, text="🔬 要求使用的检验方法标准 (可自由划选):", font=("Microsoft YaHei", 9, "bold"), bootstyle=PRIMARY).pack(side=tk.LEFT)
        
        self.btn_copy_method = tb.Button(m_header, text="📋 一键复制", command=self.copy_current_selected_method, bootstyle="info-outline", padding=(4, 1), cursor="hand2")
        self.btn_copy_method.pack(side=tk.RIGHT)
        
        # 可自由拉选复制的 Text 文本框
        self.method_text = tk.Text(method_container, height=3, font=("Consolas", 9, "bold"), wrap=tk.WORD,
                                   bg="#f0f7ff", fg="#0366d6", bd=1, relief=tk.SOLID, padx=4, pady=3)
        self.method_text.pack(fill=tk.X)
        self.method_text.insert("1.0", "请在右侧表格选择一条数据查看检验方法")
        self.method_text.bind("<Key>", lambda e: "break" if (e.state != 4 and e.keysym not in ('c', 'C', 'a', 'A')) else None)
        
        iv_frame = tb.Frame(judge_frame)
        iv_frame.pack(fill=tk.X, pady=(5,0))
        tb.Label(iv_frame, text="检测值:", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        
        self.test_value_var = tk.StringVar()
        test_entry = tb.Entry(iv_frame, textvariable=self.test_value_var, font=("Microsoft YaHei", 10), width=12)
        test_entry.pack(side=tk.LEFT, padx=(5,5))
        tb.Button(iv_frame, text="判定", command=self.do_judge, bootstyle=SUCCESS).pack(side=tk.LEFT)
        
        self.judge_result_label = tb.Label(judge_frame, text="", font=("Microsoft YaHei", 14, "bold"))
        self.judge_result_label.pack(pady=(8,0))
        
        btn_frame = tb.Frame(left_frame)
        btn_frame.pack(fill=tk.X, pady=(15,0))
        tb.Button(btn_frame, text="导出判定记录", command=self.export_history, bootstyle=INFO).pack(side=tk.LEFT)
        tb.Button(btn_frame, text="↔ 自适应列宽", command=lambda: self.autofit_tree_columns(self.tree), bootstyle="info-outline").pack(side=tk.LEFT, padx=(5, 0))
        tb.Button(btn_frame, text="清空记录", command=self.clear_history, bootstyle=DANGER).pack(side=tk.RIGHT)
        
        # 右侧：数据列表
        right_frame = tb.Frame(main_frame)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        columns = ("食品分类", "检测项目", "限量值", "单位", "检验方法", "备注")
        self.tree = ttk.Treeview(right_frame, columns=columns, show="headings", selectmode="browse")
        
        for col in columns:
            self.tree.heading(col, text=col + " ▾", command=lambda c=col: self._sort_query_tree(c))
        
        self.default_query_widths = {"食品分类": 180, "检测项目": 140, "限量值": 80, "单位": 60, "检验方法": 320, "备注": 220}
        self.apply_tree_column_widths(self.tree, "query_tree", self.default_query_widths)
        
        self.tree.tag_configure('highlight_match', background='#fff9db', foreground='#b71540')
        
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

    def _sort_ratio_tree(self, col):
        """点击列头排序比例之和列表"""
        universal_tree_sort_column(self.ratio_tree, col, "ratio_tree", self.tree_sort_states, self.ratio_columns, on_reload_callback=self.recalculate_ratios)
    
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
            
        # 解析每行：支持纯项目名 或 项目+实测值，保持行输入先后顺序
        parsed_queries = []
        for line in items_text.split('\n'):
            line_str = line.strip()
            if not line_str:
                continue
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
    # Tab 4: 台账智能反写 (Excel自动匹配) 核心功能
    # ============================================================
    def setup_excel_backfill_tab(self):
        """配置台账智能反写工作台 UI"""
        self.default_backfill_widths = {
            "序号": 50, "样品名称": 130, "检测项目": 120, "判定标准": 130, 
            "标准全称": 160, "限量要求": 110, "检验方法": 280, "检出限(LOD)": 170, 
            "实测值": 75, "判定结论": 95, "匹配状态": 120
        }
        
        self.backfill_file_var = tk.StringVar()
        self.backfill_sheet_var = tk.StringVar()
        self.backfill_food_col_var = tk.StringVar()
        self.backfill_item_col_var = tk.StringVar()
        self.backfill_meas_col_var = tk.StringVar()
        
        self._backfill_raw_headers = []
        self._backfill_raw_rows = []
        self._backfill_processed_rows = []
        self._backfill_current_wb = None

        main_frame = tb.Frame(self.tab_excel_backfill, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. 顶部导入与字段映射卡片
        ctl_card = tb.Labelframe(main_frame, text="📂 导入Excel/CSV台账与智能字段映射", padding=10)
        ctl_card.pack(fill=tk.X, pady=(0, 8))

        row1 = tb.Frame(ctl_card)
        row1.pack(fill=tk.X, pady=3)
        tb.Label(row1, text="台账文件:", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT, padx=(0, 5))
        self.backfill_file_entry = tb.Entry(row1, textvariable=self.backfill_file_var, font=("Microsoft YaHei", 9), state="readonly")
        self.backfill_file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        tb.Button(row1, text="📂 选择Excel/CSV台账...", command=self.browse_backfill_file, bootstyle=PRIMARY).pack(side=tk.LEFT, padx=(0, 6))
        tb.Button(row1, text="📄 下载标准台账模板 (.xlsx)", command=self.download_backfill_template, bootstyle="info-outline").pack(side=tk.LEFT)

        row2 = tb.Frame(ctl_card)
        row2.pack(fill=tk.X, pady=(6, 2))
        
        tb.Label(row2, text="工作表:", font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(0, 4))
        self.backfill_sheet_combo = tb.Combobox(row2, textvariable=self.backfill_sheet_var, state="readonly", width=12)
        self.backfill_sheet_combo.pack(side=tk.LEFT, padx=(0, 15))
        self.backfill_sheet_combo.bind("<<ComboboxSelected>>", self.on_backfill_sheet_changed)

        tb.Label(row2, text="样品名称列:", font=("Microsoft YaHei", 9, "bold"), bootstyle=PRIMARY).pack(side=tk.LEFT, padx=(0, 4))
        self.backfill_food_combo = tb.Combobox(row2, textvariable=self.backfill_food_col_var, state="readonly", width=14)
        self.backfill_food_combo.pack(side=tk.LEFT, padx=(0, 15))

        tb.Label(row2, text="检测项目列:", font=("Microsoft YaHei", 9, "bold"), bootstyle=PRIMARY).pack(side=tk.LEFT, padx=(0, 4))
        self.backfill_item_combo = tb.Combobox(row2, textvariable=self.backfill_item_col_var, state="readonly", width=14)
        self.backfill_item_combo.pack(side=tk.LEFT, padx=(0, 15))

        tb.Label(row2, text="实测值列(可选):", font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(0, 4))
        self.backfill_meas_combo = tb.Combobox(row2, textvariable=self.backfill_meas_col_var, state="readonly", width=14)
        self.backfill_meas_combo.pack(side=tk.LEFT, padx=(0, 15))

        self.btn_run_backfill = tb.Button(row2, text="⚡ 开始全量智能匹配反写", command=self.run_excel_backfill, bootstyle=SUCCESS, cursor="hand2")
        self.btn_run_backfill.pack(side=tk.RIGHT)

        # 2. 质控统计与进度条看板
        stats_frame = tb.Frame(main_frame)
        stats_frame.pack(fill=tk.X, pady=(0, 6))
        self.backfill_stats_label = tb.Label(
            stats_frame,
            text="📊 质控看板: 请先选择包含【样品名称】与【检测项目】的Excel工作簿或CSV委托单",
            font=("Microsoft YaHei", 9, "bold"),
            bootstyle=INFO
        )
        self.backfill_stats_label.pack(side=tk.LEFT)

        # 3. 结果预览表格 (微边框卡片与滚动容器)
        grid_card = tb.Labelframe(main_frame, text="📋 智能反写结果实时全景预览 (双击查看完整指标卡片，悬停展开完整检验方法)", padding=6)
        grid_card.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        columns = ("seq", "food", "item", "std_code", "std_name", "limit", "method", "lod", "meas", "judge", "status")
        self.backfill_tree = ttk.Treeview(grid_card, columns=columns, show='headings', selectmode='browse')
        
        self.backfill_tree.heading("seq", text="序号")
        self.backfill_tree.heading("food", text="样品名称")
        self.backfill_tree.heading("item", text="检测项目")
        self.backfill_tree.heading("std_code", text="判定标准代号")
        self.backfill_tree.heading("std_name", text="判定标准全称")
        self.backfill_tree.heading("limit", text="【反写】标准限量")
        self.backfill_tree.heading("method", text="【反写】推荐检验方法")
        self.backfill_tree.heading("lod", text="【反写】检出限(LOD)")
        self.backfill_tree.heading("meas", text="实测值")
        self.backfill_tree.heading("judge", text="判定结论")
        self.backfill_tree.heading("status", text="匹配状态")

        self.backfill_tree.column("seq", width=50, anchor="center")
        self.backfill_tree.column("food", width=130, anchor="w")
        self.backfill_tree.column("item", width=120, anchor="w")
        self.backfill_tree.column("std_code", width=130, anchor="w")
        self.backfill_tree.column("std_name", width=160, anchor="w")
        self.backfill_tree.column("limit", width=110, anchor="center")
        self.backfill_tree.column("method", width=280, anchor="w")
        self.backfill_tree.column("lod", width=170, anchor="w")
        self.backfill_tree.column("meas", width=75, anchor="center")
        self.backfill_tree.column("judge", width=95, anchor="center")
        self.backfill_tree.column("status", width=120, anchor="center")

        self.apply_tree_column_widths(self.backfill_tree, "backfill_tree", self.default_backfill_widths)
        self.setup_tree_context_menu(self.backfill_tree, "backfill_tree", self.default_backfill_widths, self.on_backfill_tree_double_click)

        # 标签颜色配置
        self.backfill_tree.tag_configure("tag_pass", background="#E8F8F5", foreground="#117A65")
        self.backfill_tree.tag_configure("tag_fail", background="#FDEDEC", foreground="#C0392B")
        self.backfill_tree.tag_configure("tag_infer", background="#FEF9E7", foreground="#B7950B")
        self.backfill_tree.tag_configure("tag_miss", background="#F2F4F4", foreground="#7F8C8D")

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
        tb.Button(bot_bar, text="📋 复制预览表格到剪贴板", command=self.copy_backfill_preview, bootstyle="secondary-outline").pack(side=tk.LEFT, padx=(0, 8))
        tb.Button(bot_bar, text="🔄 清空列表", command=self.clear_backfill_data, bootstyle="danger-outline").pack(side=tk.LEFT)
        tb.Label(bot_bar, text="⚡ Powered by 陈耿忠 (CCAA 国家注册审核员 / 实验室质控专家)", font=("Microsoft YaHei", 8), bootstyle=SECONDARY).pack(side=tk.RIGHT)

    def browse_backfill_file(self):
        """选择Excel或CSV台账文件"""
        fn = filedialog.askopenfilename(
            title="选择台账或委托单文件",
            filetypes=[("Excel工作簿 (*.xlsx;*.xls)", "*.xlsx;*.xls"), ("CSV表格 (*.csv)", "*.csv"), ("所有文件 (*.*)", "*.*")]
        )
        if not fn:
            return
        self.backfill_file_var.set(fn)
        self.load_backfill_file_structure(fn)

    def load_backfill_file_structure(self, filepath):
        """加载文件表头并智能推导列映射"""
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
                # 尝试以常见编码读取
                encodings = ['utf-8-sig', 'utf-8', 'gbk', 'gb18030']
                lines = None
                for enc in encodings:
                    try:
                        with open(filepath, 'r', encoding=enc) as f:
                            reader = csv.reader(f)
                            all_rows = list(reader)
                            if all_rows:
                                lines = all_rows
                                break
                    except Exception:
                        continue
                if not lines:
                    messagebox.showerror("错误", "无法以 UTF-8 或 GBK 编码解析该 CSV 文件！")
                    return
                self._backfill_raw_headers = [str(c).strip() for c in lines[0]]
                self._backfill_raw_rows = lines[1:]
                self.auto_detect_backfill_columns()
        except Exception as e:
            messagebox.showerror("文件加载失败", f"无法读取该文件：\n{e}")

    def on_backfill_sheet_changed(self, event=None):
        """切换 Excel Sheet"""
        if not self._backfill_current_wb:
            return
        s_name = self.backfill_sheet_var.get()
        if s_name in self._backfill_current_wb.sheetnames:
            self.read_active_sheet_data(self._backfill_current_wb[s_name])

    def read_active_sheet_data(self, sheet):
        """读取选定工作表的表头与数据"""
        rows = list(sheet.iter_rows(values_only=True))
        if not rows:
            messagebox.showwarning("提示", "该工作表为空！")
            return
        
        # 寻找首个非空行作为表头
        hdr_idx = 0
        for idx, r in enumerate(rows):
            if any(cell is not None and str(cell).strip() for cell in r):
                hdr_idx = idx
                break
                
        self._backfill_raw_headers = [str(c).strip() if c is not None else f"列{i+1}" for i, c in enumerate(rows[hdr_idx])]
        self._backfill_raw_rows = rows[hdr_idx+1:]
        self.auto_detect_backfill_columns()

    def auto_detect_backfill_columns(self):
        """智能推导映射字段"""
        hdrs = [h for h in self._backfill_raw_headers if h]
        if not hdrs:
            return
            
        self.backfill_food_combo['values'] = hdrs
        self.backfill_item_combo['values'] = hdrs
        self.backfill_meas_combo['values'] = ["(无/不导入)"] + hdrs
        
        # 1. 样品列推导
        food_col = ""
        for h in hdrs:
            if any(k in h for k in ["样品名称", "食品名称", "产品名称", "样品", "食品", "品名", "标的物", "Sample"]):
                food_col = h
                break
        if not food_col and len(hdrs) > 0:
            food_col = hdrs[0]
        self.backfill_food_col_var.set(food_col)
        
        # 2. 项目列推导
        item_col = ""
        for h in hdrs:
            if any(k in h for k in ["检测项目", "检验项目", "项目", "参数", "指标", "检测指标", "Item", "Parameter"]):
                item_col = h
                break
        if not item_col and len(hdrs) > 1:
            item_col = hdrs[1]
        self.backfill_item_col_var.set(item_col)
        
        # 3. 实测值列推导
        meas_col = "(无/不导入)"
        for h in hdrs:
            if any(k in h for k in ["实测值", "检验结果", "检测结果", "测试结果", "测定值", "结果", "Result", "Value"]):
                meas_col = h
                break
        self.backfill_meas_col_var.set(meas_col)
        
        count = len(self._backfill_raw_rows)
        self.backfill_stats_label.config(
            text=f"📊 文件解析就绪: 共载入 {count} 行数据 | 请核对上方【样品名称列】与【检测项目列】映射，随后点击开始匹配！",
            bootstyle=PRIMARY
        )

    def download_backfill_template(self):
        """生成并下载标准台账导入模板"""
        save_path = filedialog.asksaveasfilename(
            title="保存台账导入模板",
            defaultextension=".xlsx",
            initialfile="食品安全标准判定_导入模板.xlsx",
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
            ws.title = "检测任务委托单"
            
            headers = ["序号", "样品编号", "样品名称", "检测项目", "检验实测值(可选)", "备注"]
            ws.append(headers)
            
            samples = [
                (1, "SP2026001", "散装大米", "铅", "0.08", "市场抽样"),
                (2, "SP2026002", "散装大米", "镉", "0.02", "常规监测"),
                (3, "SP2026003", "小白菜", "啶虫脒", "0.5", "基地采样"),
                (4, "SP2026004", "韭菜", "腐霉利", "0.05", "快检筛查"),
                (5, "SP2026005", "熟肉制品", "沙门氏菌", "未检出", "出厂检验"),
                (6, "SP2026006", "纯牛奶", "铅", "0.01", "商超自检"),
                (7, "SP2026007", "散装糕点", "脱氢乙酸", "0.05", "专项排查"),
                (8, "SP2026008", "猪肉", "恩诺沙星", "15.0", "屠宰场监测"),
                (9, "SP2026009", "淡水鱼(草鱼)", "孔雀石绿", "未检出", "水产监测"),
                (10, "SP2026010", "压榨花生油", "黄曲霉毒素B1", "8.0", "生产入库")
            ]
            for s in samples:
                ws.append(list(s))
                
            # 美化样式
            header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
            header_font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
            thin_border = Border(left=Side(style='thin', color='CBD5E1'),
                                 right=Side(style='thin', color='CBD5E1'),
                                 top=Side(style='thin', color='CBD5E1'),
                                 bottom=Side(style='thin', color='CBD5E1'))
            
            for col in range(1, len(headers) + 1):
                cell = ws.cell(row=1, column=col)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
            for row in range(2, len(samples) + 2):
                for col in range(1, len(headers) + 1):
                    c = ws.cell(row=row, column=col)
                    c.border = thin_border
                    c.alignment = Alignment(horizontal="center" if col in (1, 2, 5) else "left", vertical="center")
                    
            ws.column_dimensions['A'].width = 8
            ws.column_dimensions['B'].width = 16
            ws.column_dimensions['C'].width = 18
            ws.column_dimensions['D'].width = 16
            ws.column_dimensions['E'].width = 18
            ws.column_dimensions['F'].width = 16
            
            wb.save(save_path)
            messagebox.showinfo("生成成功", f"标准台账导入模板已成功生成！\n\n保存路径：{save_path}\n您可直接打开该模板填写样品及检测项目，随后导入反写。")
        except Exception as e:
            messagebox.showerror("生成失败", f"无法创建模板文件：\n{e}")

    def run_excel_backfill(self):
        """执行全量智能匹配与信息反写"""
        if not self._backfill_raw_rows:
            messagebox.showwarning("提示", "请先选择有效的台账或委托单文件！")
            return
            
        food_col_name = self.backfill_food_col_var.get()
        item_col_name = self.backfill_item_col_var.get()
        meas_col_name = self.backfill_meas_col_var.get()
        
        if not food_col_name or not item_col_name:
            messagebox.showwarning("提示", "请先正确选择【样品名称列】与【检测项目列】！")
            return
            
        try:
            food_idx = self._backfill_raw_headers.index(food_col_name)
            item_idx = self._backfill_raw_headers.index(item_col_name)
        except ValueError:
            messagebox.showerror("错误", "所选列名称在原始表头中未找到！")
            return
            
        meas_idx = -1
        if meas_col_name and meas_col_name != "(无/不导入)":
            try:
                meas_idx = self._backfill_raw_headers.index(meas_col_name)
            except ValueError:
                meas_idx = -1

        self.backfill_tree.delete(*self.backfill_tree.get_children())
        self._backfill_processed_rows = []
        
        total = len(self._backfill_raw_rows)
        exact_hits = 0
        infer_hits = 0
        miss_hits = 0
        
        for idx, row in enumerate(self._backfill_raw_rows):
            f_val = str(row[food_idx]).strip() if food_idx < len(row) and row[food_idx] is not None else ""
            i_val = str(row[item_idx]).strip() if item_idx < len(row) and row[item_idx] is not None else ""
            m_val = str(row[meas_idx]).strip() if (meas_idx >= 0 and meas_idx < len(row) and row[meas_idx] is not None) else None
            
            if not f_val and not i_val:
                continue
                
            res = auto_match_record(self.standards, self.global_index, f_val, i_val, m_val)
            std_code, std_name, limit_str, method_str, lod_str, meas_display, judge_display, status_tag, matched_cat = res
            
            if "精准" in status_tag:
                exact_hits += 1
                row_tag = "tag_fail" if "不合格" in judge_display else "tag_pass"
            elif "推导" in status_tag:
                infer_hits += 1
                row_tag = "tag_fail" if "不合格" in judge_display else "tag_infer"
            else:
                miss_hits += 1
                row_tag = "tag_miss"
                
            seq = idx + 1
            tree_vals = (seq, f_val, i_val, std_code, std_name, limit_str, method_str, lod_str, meas_display, judge_display, status_tag)
            self.backfill_tree.insert('', 'end', values=tree_vals, tags=(row_tag,))
            
            # 记录完整反写数据行
            self._backfill_processed_rows.append({
                'raw_row': row,
                'seq': seq,
                'food': f_val,
                'item': i_val,
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
        pct = (exact_hits / valid_total * 100) if valid_total > 0 else 0
        self.backfill_stats_label.config(
            text=f"📊 质控看板: 共匹配 {valid_total} 行 | ✔ 精准匹配: {exact_hits} 行 ({pct:.1f}%) | ⚠️ 品类推导匹配: {infer_hits} 行 | ❓ 未查见: {miss_hits} 行",
            bootstyle=SUCCESS if pct >= 80 else INFO
        )
        messagebox.showinfo("反写匹配完成", f"跨 6 大食品安全国家标准全量匹配完成！\n\n处理数据：{valid_total} 行\n精准匹配率：{pct:.1f}% ({exact_hits}行)\n品类推导：{infer_hits}行\n未查见：{miss_hits}行\n\n可在预览表中双击查看详情或点击下方【导出已反写Excel表格】保存输出！")

    def export_backfilled_excel(self):
        """导出反写完成的高保真 Excel 文件"""
        if not self._backfill_processed_rows:
            messagebox.showwarning("提示", "当前没有可导出的反写结果，请先执行智能匹配！")
            return
            
        save_path = filedialog.asksaveasfilename(
            title="导出反写完成的Excel工作簿",
            defaultextension=".xlsx",
            initialfile=f"食品安全标准智能反写台账_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
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
            ws.title = "标准判定与方法反写台账"
            
            # 拼接表头：保留原始列 + 追加反写列
            orig_headers = [str(h) for h in self._backfill_raw_headers]
            backfill_headers = [
                "【反写】判定标准代号",
                "【反写】判定标准全称",
                "【反写】标准限量要求",
                "【反写】推荐检验方法",
                "【反写】方法检出限(LOD)",
                "【反写】实测值",
                "【反写】判定结论",
                "【反写】匹配状态",
                "【反写】匹配品类"
            ]
            full_headers = orig_headers + backfill_headers
            ws.append(full_headers)
            
            # 写入各数据行
            for item in self._backfill_processed_rows:
                orig_vals = [cell if cell is not None else "" for cell in item['raw_row']]
                # 补足长度
                while len(orig_vals) < len(orig_headers):
                    orig_vals.append("")
                bf_vals = [
                    item['std_code'],
                    item['std_name'],
                    item['limit'],
                    item['method'],
                    item['lod'],
                    item['meas'],
                    item['judge'],
                    item['status'],
                    item['category']
                ]
                ws.append(orig_vals + bf_vals)
                
            # 商务严谨与现代化视觉美化
            navy_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
            tech_fill = PatternFill(start_color="0284C7", end_color="0284C7", fill_type="solid")
            hdr_font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
            
            pass_fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
            pass_font = Font(name="Microsoft YaHei", size=9, bold=True, color="065F46")
            fail_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
            fail_font = Font(name="Microsoft YaHei", size=9, bold=True, color="991B1B")
            
            thin_border = Border(left=Side(style='thin', color='CBD5E1'),
                                 right=Side(style='thin', color='CBD5E1'),
                                 top=Side(style='thin', color='CBD5E1'),
                                 bottom=Side(style='thin', color='CBD5E1'))
            
            orig_len = len(orig_headers)
            total_cols = len(full_headers)
            
            for col in range(1, total_cols + 1):
                cell = ws.cell(row=1, column=col)
                cell.fill = navy_fill if col <= orig_len else tech_fill
                cell.font = hdr_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                
            judge_col_idx = orig_len + 7  # 判定结论所在列序号 (1-based)
            
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
                            
            # 自适应列宽 (防截断)
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col[:50]:
                    v = str(cell.value or '')
                    # 中文按2字节估算
                    length = sum(2 if ord(ch) > 127 else 1 for ch in v)
                    if length > max_len:
                        max_len = length
                ws.column_dimensions[col_letter].width = min(42, max(12, max_len + 3))
                
            wb.save(save_path)
            messagebox.showinfo("导出成功", f"反写完成的 Excel 台账已保存：\n\n{save_path}\n\n已自动保留所有原始列并追加判定标准、推荐方法、检出限与判定结论！")
        except Exception as e:
            messagebox.showerror("导出失败", f"无法保存 Excel 文件：\n{e}")

    def copy_backfill_preview(self):
        """复制当前预览表格到剪贴板"""
        if not self._backfill_processed_rows:
            messagebox.showinfo("提示", "当前无预览数据！")
            return
        lines = []
        headers = ["序号", "样品名称", "检测项目", "判定标准代号", "标准全称", "标准限量要求", "推荐检验方法", "检出限(LOD)", "实测值", "判定结论", "匹配状态"]
        lines.append("\t".join(headers))
        for it in self._backfill_processed_rows:
            row_str = "\t".join([
                str(it['seq']), it['food'], it['item'], it['std_code'], it['std_name'],
                it['limit'], it['method'], it['lod'], it['meas'], it['judge'], it['status']
            ])
            lines.append(row_str)
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(lines))
        self.root.update()
        messagebox.showinfo("复制成功", f"已将全部 {len(self._backfill_processed_rows)} 条反写预览数据复制到剪贴板！可直接粘贴至 Excel。")

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
        if not vals or len(vals) < 11:
            return
        seq, food, item, std_code, std_name, limit_val, method_str, lod_str, meas, judge, status = vals
        full_std = f"{std_code} {std_name}".strip()
        extra_note = f"【匹配状态】{status}  |  【方法检出限(LOD)】{lod_str}"
        ItemDetailDialog(
            self.root,
            full_std,
            food,
            item,
            limit_val,
            unit="",
            notes=extra_note,
            measured_val=meas if meas != "/" else None,
            judge_res=judge if judge != "/" else None,
            method_str=method_str
        )

    # ============================================================
    # GB 2760 添加剂复配比例之和智能计算器相关方法
    # ============================================================
    def setup_ratio_tab(self):
        """GB 2760 添加剂复配/混合使用比例之和智能计算器 UI"""
        main_frame = tb.Frame(self.tab_ratio, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 顶部：食品分类选择与说明
        top_frame = tb.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 8))

        header_frame = tb.Frame(top_frame)
        header_frame.pack(fill=tk.X, pady=(0, 4))
        tb.Label(header_frame, text="🧪 GB 2760 添加剂复配比例之和智能计算器", font=("Microsoft YaHei", 12, "bold"), bootstyle=PRIMARY).pack(side=tk.LEFT)
        tb.Label(header_frame, text="  依据 GB 2760-2024 第 4.4 条：同一功能添加剂(防腐剂/着色剂/抗氧化剂)各自用量占最大限量比例之和 ≤ 1.0", font=("Microsoft YaHei", 9), bootstyle=SECONDARY).pack(side=tk.LEFT, padx=10)

        # 食品分类选择行
        cat_select_frame = tb.Frame(top_frame)
        cat_select_frame.pack(fill=tk.X, pady=4)
        tb.Label(cat_select_frame, text="目标食品分类 (GB 2760):", font=("Microsoft YaHei", 10, "bold")).pack(side=tk.LEFT)
        
        gb2760_items = self.standards.get('GB 2760-2024', {}).get('items', [])
        gb2760_cats = sorted(list(set(item.get('food_category', '') for item in gb2760_items if item.get('food_category'))))
        
        self.ratio_food_cat_var = tk.StringVar()
        self.ratio_cat_combo = ttk.Combobox(cat_select_frame, textvariable=self.ratio_food_cat_var, values=gb2760_cats, font=("Microsoft YaHei", 10), width=45)
        self.ratio_cat_combo.pack(side=tk.LEFT, padx=(8, 10))
        if gb2760_cats:
            default_cat = next((c for c in gb2760_cats if "07.02" in c or "糕点" in c), gb2760_cats[0])
            self.ratio_food_cat_var.set(default_cat)
            
        self.ratio_cat_combo.bind("<<ComboboxSelected>>", lambda e: self.recalculate_ratios())
        self.ratio_cat_combo.bind("<KeyRelease>", lambda e: self.recalculate_ratios())

        # 快速模板按钮
        tmpl_frame = tb.Frame(top_frame)
        tmpl_frame.pack(fill=tk.X, pady=(2, 0))
        tb.Label(tmpl_frame, text="常用复配模板:", font=("Microsoft YaHei", 9, "bold"), bootstyle=SECONDARY).pack(side=tk.LEFT)
        tb.Button(tmpl_frame, text="➕ 常用防腐剂组 (脱氢乙酸+山梨酸+苯甲酸)", command=self.load_preservative_template, bootstyle="info-outline").pack(side=tk.LEFT, padx=4)
        tb.Button(tmpl_frame, text="➕ 黄色着色剂组 (柠檬黄+日落黄)", command=self.load_yellow_colorant_template, bootstyle="warning-outline").pack(side=tk.LEFT, padx=4)
        tb.Button(tmpl_frame, text="➕ 红色着色剂组 (胭脂红+诱惑红)", command=self.load_red_colorant_template, bootstyle="danger-outline").pack(side=tk.LEFT, padx=4)
        tb.Button(tmpl_frame, text="➕ 抗氧化剂组 (TBHQ+BHA+BHT)", command=self.load_antioxidant_template, bootstyle="secondary-outline").pack(side=tk.LEFT, padx=4)

        # 中部：PanedWindow 分割左侧待算添加剂表格与右侧计算结论
        paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, pady=5)

        # 左侧：添加剂录入与列表
        left_box = tb.LabelFrame(paned, text="待核查添加剂列表 (实测检出量)", padding=8)
        paned.add(left_box, weight=3)

        # 添加行输入工具栏
        input_bar = tb.Frame(left_box)
        input_bar.pack(fill=tk.X, pady=(0, 6))

        tb.Label(input_bar, text="添加剂:").pack(side=tk.LEFT)
        gb2760_item_names = sorted(list(set(item.get('test_item', '') for item in gb2760_items if item.get('test_item'))))
        self.ratio_item_input_var = tk.StringVar()
        self.ratio_item_combo = ttk.Combobox(input_bar, textvariable=self.ratio_item_input_var, values=gb2760_item_names, width=16, font=("Microsoft YaHei", 9))
        self.ratio_item_combo.pack(side=tk.LEFT, padx=4)
        if gb2760_item_names:
            self.ratio_item_input_var.set(gb2760_item_names[0])

        tb.Label(input_bar, text="实测:").pack(side=tk.LEFT, padx=(4, 0))
        self.ratio_meas_input_var = tk.StringVar(value="0.1")
        tb.Entry(input_bar, textvariable=self.ratio_meas_input_var, width=7, font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=2)
        
        self.ratio_unit_var = tk.StringVar(value="g/kg")
        self.ratio_unit_combo = ttk.Combobox(input_bar, textvariable=self.ratio_unit_var, values=["g/kg", "mg/kg", "ppm", "mg/100g", "%"], width=8, state="readonly", font=("Microsoft YaHei", 9))
        self.ratio_unit_combo.pack(side=tk.LEFT, padx=2)

        tb.Button(input_bar, text="➕ 添加", command=self.add_ratio_item_from_input, bootstyle=PRIMARY).pack(side=tk.LEFT, padx=(5, 2))
        tb.Button(input_bar, text="➖ 删除", command=self.remove_selected_ratio_item, bootstyle=DANGER).pack(side=tk.LEFT, padx=2)
        tb.Button(input_bar, text="🗑️ 清空", command=self.clear_ratio_items, bootstyle=SECONDARY).pack(side=tk.LEFT, padx=2)

        # 录入表格
        self.ratio_columns = ("添加剂名称", "功能分组", "录入实测值", "换算后(g/kg)", "最大限量(g/kg)", "单项比例(C/M)", "合规状态")
        self.ratio_tree = ttk.Treeview(left_box, columns=self.ratio_columns, show="headings", selectmode="browse", height=9)
        for col in self.ratio_columns:
            self.ratio_tree.heading(col, text=col + " ▾", command=lambda c=col: self._sort_ratio_tree(c))
            
        self.ratio_default_widths = {"添加剂名称": 105, "功能分组": 110, "录入实测值": 95, "换算后(g/kg)": 95, "最大限量(g/kg)": 95, "单项比例(C/M)": 95, "合规状态": 90}
        for col, w in self.ratio_default_widths.items():
            self.ratio_tree.column(col, width=w, anchor="center")
            
        self.ratio_tree.tag_configure('ratio_pass', background='#dafbe1', foreground='#1a7f37')
        self.ratio_tree.tag_configure('ratio_fail', background='#ffebe9', foreground='#cf222e')
        self.ratio_tree.tag_configure('ratio_prohibited', background='#ffebe9', foreground='#cf222e')

        r_v_scroll = ttk.Scrollbar(left_box, orient=tk.VERTICAL, command=self.ratio_tree.yview)
        self.ratio_tree.configure(yscrollcommand=r_v_scroll.set)
        self.ratio_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        r_v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 右侧：计算结果展示与分步解析
        right_box = tb.LabelFrame(paned, text="📊 比例之和动态贡献堆叠图与合规看板", padding=8)
        paned.add(right_box, weight=3)

        # 顶部：动态堆叠柱状图与安全红线 Canvas
        self.ratio_canvas = tk.Canvas(right_box, height=135, bg="#ffffff", highlightthickness=1, highlightbackground="#dee2e6")
        self.ratio_canvas.pack(fill=tk.X, pady=(0, 6))
        self.ratio_canvas.bind("<Configure>", lambda e: self.draw_ratio_stacked_bar(getattr(self, '_last_grouped_items', {})))

        # 下部：分步计算公式与综合报告
        self.ratio_calc_text = tk.Text(right_box, font=("Consolas", 9), wrap=tk.WORD, bg="#fbfcfe")
        c_v_scroll = ttk.Scrollbar(right_box, orient=tk.VERTICAL, command=self.ratio_calc_text.yview)
        self.ratio_calc_text.configure(yscrollcommand=c_v_scroll.set)
        self.ratio_calc_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        c_v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 底部操作栏
        bottom_bar = tb.Frame(main_frame)
        bottom_bar.pack(fill=tk.X, pady=(8, 0))
        tb.Button(bottom_bar, text="⚡ 重新核算", command=self.recalculate_ratios, bootstyle=SUCCESS).pack(side=tk.LEFT)
        tb.Button(bottom_bar, text="📋 复制计算全过程与结论", command=self.copy_ratio_calculation, bootstyle="info-outline").pack(side=tk.LEFT, padx=10)
        tb.Button(bottom_bar, text="📄 导出合规核查单(CSV)...", command=self.export_ratio_calculation, bootstyle="secondary-outline").pack(side=tk.LEFT)

        # 初始数据列表
        self.ratio_item_data = []
        self.load_preservative_template()

    def _get_additive_group(self, item_name):
        """判断添加剂所属的功能分组"""
        for group_name, info in ADDITIVE_FUNCTION_GROUPS.items():
            for g_item in info["items"]:
                if g_item in item_name or item_name in g_item:
                    return group_name
        return "其他添加剂"

    def _lookup_gb2760_limit(self, food_cat, item_name):
        """在 GB 2760-2024 中查询指定分类和添加剂的最大允许使用量（折算为 g/kg）"""
        if not food_cat or not item_name:
            return None, "未指定分类/项目"
            
        gb2760_items = self.standards.get('GB 2760-2024', {}).get('items', [])
        food_cat_clean = food_cat.strip()
        item_clean = item_name.strip()
        
        # 查找匹配项
        best_match = None
        for item in gb2760_items:
            f_cat = item.get('food_category', '')
            t_item = item.get('test_item', '')
            
            # 精确或包含匹配
            if (food_cat_clean in f_cat or f_cat in food_cat_clean) and (item_clean in t_item or t_item in item_clean):
                best_match = item
                break
                
        if not best_match:
            # 降级：仅按食品类别大类前缀查找
            cat_prefix = food_cat_clean.split(' ')[0] if ' ' in food_cat_clean else food_cat_clean[:5]
            for item in gb2760_items:
                f_cat = item.get('food_category', '')
                t_item = item.get('test_item', '')
                if cat_prefix in f_cat and (item_clean in t_item or t_item in item_clean):
                    best_match = item
                    break
                    
        if not best_match:
            return 0.0, "未允许使用/无限量"
            
        limit_str = best_match.get('limit_value', '')
        unit_str = best_match.get('unit', 'g/kg').lower()
        
        if "不得检出" in limit_str or "不得使用" in limit_str:
            return 0.0, "不得使用"
            
        num = extract_first_number(limit_str)
        if num is None:
            return 0.0, limit_str
            
        # 单位换算：如果是 mg/kg，折算为 g/kg
        if "mg/kg" in unit_str or "mg" in unit_str:
            num = num / 1000.0
            
        return num, f"{num} g/kg"

    def load_preservative_template(self):
        """加载常用防腐剂模板"""
        self.ratio_item_data = [
            {"name": "脱氢乙酸", "raw_val": 0.2, "unit": "g/kg", "measured": 0.2},
            {"name": "山梨酸", "raw_val": 0.3, "unit": "g/kg", "measured": 0.3},
            {"name": "苯甲酸", "raw_val": 100, "unit": "mg/kg", "measured": 0.1}
        ]
        self.recalculate_ratios()

    def load_yellow_colorant_template(self):
        """加载黄色着色剂模板"""
        self.ratio_item_data = [
            {"name": "柠檬黄", "raw_val": 50, "unit": "mg/kg", "measured": 0.05},
            {"name": "日落黄", "raw_val": 50, "unit": "mg/kg", "measured": 0.05}
        ]
        self.recalculate_ratios()

    def load_red_colorant_template(self):
        """加载红色着色剂模板"""
        self.ratio_item_data = [
            {"name": "胭脂红", "raw_val": 20, "unit": "mg/kg", "measured": 0.02},
            {"name": "诱惑红", "raw_val": 30, "unit": "mg/kg", "measured": 0.03}
        ]
        self.recalculate_ratios()

    def load_antioxidant_template(self):
        """加载抗氧化剂模板"""
        self.ratio_item_data = [
            {"name": "TBHQ", "raw_val": 0.1, "unit": "g/kg", "measured": 0.1},
            {"name": "BHA", "raw_val": 50, "unit": "mg/kg", "measured": 0.05},
            {"name": "BHT", "raw_val": 50, "unit": "ppm", "measured": 0.05}
        ]
        self.recalculate_ratios()

    def add_ratio_item_from_input(self):
        """从输入框添加一行添加剂指标（支持多单位智能自动折算为 g/kg）"""
        name = self.ratio_item_input_var.get().strip()
        meas_str = self.ratio_meas_input_var.get().strip()
        unit = self.ratio_unit_var.get().strip() if hasattr(self, 'ratio_unit_var') else "g/kg"
        
        if not name:
            messagebox.showwarning("提示", "请输入或选择添加剂名称！")
            return
            
        val = extract_first_number(meas_str)
        if val is None or val < 0:
            messagebox.showwarning("提示", "请输入有效的非负实测值数值！")
            return
            
        # 单位自动换算为法定基准单位 g/kg
        factor_map = {
            "g/kg": 1.0,
            "mg/kg": 0.001,
            "ppm": 0.001,
            "mg/100g": 0.01,
            "%": 10.0
        }
        factor = factor_map.get(unit, 1.0)
        val_g_per_kg = val * factor
            
        # 检查是否已存在同名项目
        for item in self.ratio_item_data:
            if item["name"] == name:
                item["raw_val"] = val
                item["unit"] = unit
                item["measured"] = val_g_per_kg
                self.recalculate_ratios()
                return
                
        self.ratio_item_data.append({
            "name": name,
            "raw_val": val,
            "unit": unit,
            "measured": val_g_per_kg
        })
        self.recalculate_ratios()

    def remove_selected_ratio_item(self):
        """删除表格中选中的添加剂行"""
        selection = self.ratio_tree.selection()
        if not selection:
            messagebox.showinfo("提示", "请先在左侧表格中选中要删除的添加剂！")
            return
        vals = self.ratio_tree.item(selection[0], 'values')
        name_to_del = vals[0]
        self.ratio_item_data = [item for item in self.ratio_item_data if item["name"] != name_to_del]
        self.recalculate_ratios()

    def clear_ratio_items(self):
        """清空待核查添加剂列表"""
        self.ratio_item_data = []
        self.recalculate_ratios()

    def recalculate_ratios(self):
        """执行比例之和计算与渲染（支持原单位展示与换算公式追溯）"""
        food_cat = self.ratio_food_cat_var.get().strip()
        self.ratio_tree.delete(*self.ratio_tree.get_children())
        self.ratio_calc_text.delete("1.0", tk.END)

        if not self.ratio_item_data:
            self.ratio_calc_text.insert(tk.END, "💡 当前未录入任何添加剂指标。\n请点击上方快速模板，或使用输入栏添加待核查添加剂及实测值。")
            if hasattr(self, 'ratio_canvas') and self.ratio_canvas.winfo_exists():
                self.ratio_canvas.delete("all")
            return

        # 整理与归类
        grouped_items = {}
        rendered_rows = []

        for item in self.ratio_item_data:
            name = item["name"]
            meas = item["measured"]
            raw_v = item.get("raw_val", meas)
            unit_str = item.get("unit", "g/kg")
            raw_disp = f"{raw_v:g} {unit_str}"
            
            group = self._get_additive_group(name)
            limit_val, limit_disp = self._lookup_gb2760_limit(food_cat, name)

            if limit_val is not None and limit_val > 0:
                ratio = meas / limit_val
                status = "✔ 单项合格" if meas <= limit_val else "✘ 单项超标"
                tag = 'ratio_pass' if meas <= limit_val else 'ratio_fail'
                ratio_disp = f"{ratio:.4f}"
                limit_disp_num = f"{limit_val:.4f}"
            else:
                ratio = None
                status = "✘ 未允许使用"
                tag = 'ratio_prohibited'
                ratio_disp = "--"
                limit_disp_num = "0.0000 (禁用)"

            rendered_rows.append((name, group, raw_disp, f"{meas:.4f}", limit_disp_num, ratio_disp, status, tag))

            if group not in grouped_items:
                grouped_items[group] = []
            grouped_items[group].append({
                "name": name,
                "raw_val": raw_v,
                "unit": unit_str,
                "measured": meas,
                "limit": limit_val,
                "ratio": ratio,
                "status": status
            })

        # 填充左侧表格
        for row in rendered_rows:
            self.ratio_tree.insert('', 'end', values=row[:7], tags=(row[7],))

        # 构建右侧计算解析与合规报告看板
        out = []
        out.append("=" * 55)
        out.append(" 📋 GB 2760 食品添加剂复配比例之和智能核查看板")
        out.append("=" * 55)
        out.append(f"【目标食品类别】: {food_cat if food_cat else '未指定'}")
        out.append(f"【依据标准条款】: GB 2760-2024 第 4.4 条 (相同功能添加剂比例之和 ≤ 1.0)")
        out.append("-" * 55)

        overall_all_pass = True

        for group_name, items_in_grp in grouped_items.items():
            out.append(f"\n📂 功能分组: 【{group_name}】 (共 {len(items_in_grp)} 项)")
            
            # 检查是否有禁用项目
            has_prohibited = any(it["limit"] is None or it["limit"] <= 0 for it in items_in_grp)
            if has_prohibited:
                prohib_names = [it["name"] for it in items_in_grp if it["limit"] is None or it["limit"] <= 0]
                out.append(f"  ❌ 违规警告: [{', '.join(prohib_names)}] 在当前食品类别中属于【禁用/未允许使用】添加剂！")
                overall_all_pass = False

            # 计算比例之和
            valid_items = [it for it in items_in_grp if it["ratio"] is not None]
            if not valid_items:
                out.append("  ⚠️ 该分组内无允许使用的有效限量，无法进行比例之和相加。")
                continue

            sum_ratio = sum(it["ratio"] for it in valid_items)
            terms = [f"({it['name']}: {it['measured']:.3f}/{it['limit']:.3f}={it['ratio']:.3f})" for it in valid_items]
            formula_str = " + ".join(terms)
            
            out.append(f"  🧮 计算公式: Σ (实测值 / 最大限量)")
            out.append(f"            = {formula_str}")
            out.append(f"            = {sum_ratio:.4f}")

            if not has_prohibited and sum_ratio <= 1.0001:
                out.append(f"  ✅ 判定结论: ✔ 合规！比例之和 = {sum_ratio:.4f} ≤ 1.0")
            else:
                overall_all_pass = False
                if sum_ratio > 1.0001:
                    over_pct = (sum_ratio - 1.0) * 100
                    out.append(f"  ❌ 判定结论: ✘ 超标！比例之和 = {sum_ratio:.4f} > 1.0 (超出安全限量 {over_pct:.1f}%)")

        out.append("\n" + "=" * 55)
        if overall_all_pass:
            out.append(" 🏆 综合合规结论: ✔ 全部功能分组均符合 GB 2760 比例之和要求！")
        else:
            out.append(" 🚨 综合合规结论: ✘ 存在比例之和超标或违规添加项，产品初审不合格！")
        out.append("=" * 55)

        self.ratio_calc_text.insert(tk.END, "\n".join(out))
        
        # 缓存分组数据并动态绘制堆叠柱状图
        self._last_grouped_items = grouped_items
        self.draw_ratio_stacked_bar(grouped_items)

    def draw_ratio_stacked_bar(self, grouped_items):
        """在 Canvas 上动态绘制各功能分组的比例之和堆叠柱状图与安全红线"""
        if not hasattr(self, 'ratio_canvas') or not self.ratio_canvas.winfo_exists():
            return
            
        self.ratio_canvas.delete("all")
        
        c_w = self.ratio_canvas.winfo_width()
        if c_w < 100:
            c_w = 480
        c_h = self.ratio_canvas.winfo_height()
        if c_h < 50:
            c_h = 135

        # 调色盘
        palette = [
            "#4361ee", "#7209b7", "#f72585", "#06d6a0", 
            "#fd7e14", "#20c997", "#3a0ca3", "#ef476f"
        ]

        # 绘制柔和卡片背景
        self.ratio_canvas.create_rectangle(0, 0, c_w, c_h, fill="#fdfdfe", outline="")

        if not grouped_items:
            self.ratio_canvas.create_text(c_w / 2, c_h / 2, text="💡 暂无添加剂数据，请先录入或加载模板", fill="#8c959f", font=("Microsoft YaHei", 9))
            return

        y_cursor = 10
        color_idx = 0

        for group_name, items in grouped_items.items():
            valid_items = [it for it in items if it.get("ratio") is not None]
            sum_ratio = sum(it["ratio"] for it in valid_items) if valid_items else 0.0

            # 绘制组名与总状态徽章
            self.ratio_canvas.create_text(12, y_cursor + 7, anchor="w", text=f"【{group_name}】", font=("Microsoft YaHei", 9, "bold"), fill="#212529")
            
            # 状态标签
            has_prohib = any(it.get("limit") is None or it.get("limit") <= 0 for it in items)
            if has_prohib:
                badge_text = "✘ 存在禁用项"
                badge_bg = "#ffebe9"
                badge_fg = "#cf222e"
            elif sum_ratio <= 1.0001:
                pct = sum_ratio * 100
                badge_text = f"✔ 合规 ({pct:.1f}% / 100%)"
                badge_bg = "#dafbe1"
                badge_fg = "#1a7f37"
            else:
                over_pct = (sum_ratio - 1.0) * 100
                badge_text = f"✘ 超标 ({sum_ratio * 100:.1f}%, 超出 {over_pct:.1f}%)"
                badge_bg = "#ffebe9"
                badge_fg = "#cf222e"

            # 绘制 Badge
            b_w = len(badge_text) * 7.5 + 14
            bx1 = c_w - b_w - 12
            by1 = y_cursor - 2
            bx2 = c_w - 12
            by2 = y_cursor + 16
            self.ratio_canvas.create_rectangle(bx1, by1, bx2, by2, fill=badge_bg, outline=badge_fg, width=1)
            self.ratio_canvas.create_text((bx1 + bx2) / 2, (by1 + by2) / 2, text=badge_text, fill=badge_fg, font=("Microsoft YaHei", 8, "bold"))

            # 绘制柱状图轨道
            bar_y1 = y_cursor + 20
            bar_y2 = bar_y1 + 20
            bar_x1 = 12
            bar_x2 = c_w - 12
            total_bar_w = bar_x2 - bar_x1

            max_scale_val = max(1.25, sum_ratio * 1.15)
            pixels_per_unit = total_bar_w / max_scale_val
            red_line_x = bar_x1 + 1.0 * pixels_per_unit

            # 绘制背景槽
            self.ratio_canvas.create_rectangle(bar_x1, bar_y1, bar_x2, bar_y2, fill="#f1f3f5", outline="#dee2e6", width=1)

            # 绘制各添加剂堆叠分段
            cur_x = bar_x1
            legend_items = []

            for it in valid_items:
                seg_ratio = it["ratio"]
                seg_w = seg_ratio * pixels_per_unit
                if seg_w > 0:
                    seg_color = palette[color_idx % len(palette)]
                    color_idx += 1
                    
                    next_x = cur_x + seg_w
                    self.ratio_canvas.create_rectangle(cur_x, bar_y1 + 1, min(next_x, bar_x2 - 1), bar_y2 - 1, fill=seg_color, outline="#ffffff", width=1)
                    
                    if seg_w >= 28:
                        self.ratio_canvas.create_text((cur_x + next_x) / 2, (bar_y1 + bar_y2) / 2, text=f"{seg_ratio*100:.0f}%", fill="#ffffff", font=("Consolas", 8, "bold"))
                        
                    legend_items.append((it["name"], seg_color, f"{seg_ratio*100:.1f}%"))
                    cur_x = next_x

            # 绘制 1.0 安全红线 (100% 限值线)
            self.ratio_canvas.create_line(red_line_x, bar_y1 - 3, red_line_x, bar_y2 + 3, fill="#e63946", width=2, dash=(4, 2))
            self.ratio_canvas.create_text(red_line_x, bar_y1 - 4, anchor="s", text="100%安全红线", fill="#e63946", font=("Microsoft YaHei", 7, "bold"))

            # 绘制图例 (Legend)
            leg_y = bar_y2 + 8
            leg_x = bar_x1
            for l_name, l_col, l_pct in legend_items:
                self.ratio_canvas.create_rectangle(leg_x, leg_y + 1, leg_x + 7, leg_y + 8, fill=l_col, outline="")
                txt = f"{l_name} ({l_pct})"
                self.ratio_canvas.create_text(leg_x + 10, leg_y + 4, anchor="w", text=txt, font=("Microsoft YaHei", 7), fill="#495057")
                leg_x += len(txt) * 6.5 + 16
                if leg_x > c_w - 40:
                    leg_x = bar_x1
                    leg_y += 12

            y_cursor = leg_y + 20

    def copy_ratio_calculation(self):
        """复制计算过程与合规结论到剪贴板"""
        txt = self.ratio_calc_text.get("1.0", tk.END).strip()
        if not txt:
            messagebox.showinfo("提示", "当前无计算结果可复制！")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(txt)
        messagebox.showinfo("复制成功", "已将完整比例之和计算过程与合规结论复制到剪贴板！\n可直接粘贴至检测报告或质检原始记录中。")

    def export_ratio_calculation(self):
        """导出比例之和计算结果为 CSV"""
        children = self.ratio_tree.get_children()
        if not children:
            messagebox.showinfo("提示", "当前表格中没有可导出的数据！")
            return
            
        file_path = filedialog.asksaveasfilename(
            title="导出添加剂复配比例之和核查单",
            defaultextension=".csv",
            filetypes=[("CSV表格文件 (Excel兼容)", "*.csv"), ("文本文件", "*.txt"), ("所有文件", "*.*")],
            initialfile=f"GB2760_添加剂比例之和核查单_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        if not file_path:
            return
            
        try:
            food_cat = self.ratio_food_cat_var.get().strip()
            with open(file_path, 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(["目标食品类别", food_cat])
                writer.writerow(["依据标准", "GB 2760-2024 食品安全国家标准 食品添加剂使用标准"])
                writer.writerow([])
                writer.writerow(["添加剂名称", "功能分组", "录入实测值", "换算后(g/kg)", "最大限量(g/kg)", "单项比例(C/M)", "合规状态"])
                for item_id in children:
                    vals = self.ratio_tree.item(item_id, 'values')
                    writer.writerow(vals)
                writer.writerow([])
                writer.writerow(["--- 详细计算过程与合规结论 ---"])
                for line in self.ratio_calc_text.get("1.0", tk.END).strip().split('\n'):
                    writer.writerow([line])
                    
            messagebox.showinfo("导出成功", f"已成功导出比例之和核查单到：\n{file_path}")
        except Exception as e:
            messagebox.showerror("导出失败", f"保存文件时出错: {e}")

    # ============================================================
    # 常见市售商品名 / 执行标准 ➔ 国标分类智能反向映射相关方法
    # ============================================================
    def setup_inference_tab(self):
        """常见市售商品名 / 执行标准 ➔ 国标分类智能反向映射 UI"""
        main_frame = tb.Frame(self.tab_inference, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 顶部：搜索框与引导
        top_frame = tb.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 8))

        header_frame = tb.Frame(top_frame)
        header_frame.pack(fill=tk.X, pady=(0, 4))
        tb.Label(header_frame, text="🔍 常见市售商品名 / 执行标准 ➔ 国标分类智能反查导航", font=("Microsoft YaHei", 12, "bold"), bootstyle=PRIMARY).pack(side=tk.LEFT)
        tb.Label(header_frame, text="  输入商品俗名或执行标准号，秒级定位 GB 2760 / GB 2762 / GB 29921 准确分类层级与高风险指标", font=("Microsoft YaHei", 9), bootstyle=SECONDARY).pack(side=tk.LEFT, padx=10)

        # 搜索输入行
        search_box = tb.Frame(top_frame)
        search_box.pack(fill=tk.X, pady=4)
        tb.Label(search_box, text="商品名或执行标准:", font=("Microsoft YaHei", 10, "bold")).pack(side=tk.LEFT)
        
        self.inf_search_var = tk.StringVar()
        self.inf_search_var.trace('w', lambda *args: self.filter_inference_db())
        inf_entry = tb.Entry(search_box, textvariable=self.inf_search_var, font=("Microsoft YaHei", 10), width=35)
        inf_entry.pack(side=tk.LEFT, padx=(8, 10))
        
        tb.Button(search_box, text="清空输入", command=lambda: self.inf_search_var.set(""), bootstyle=SECONDARY).pack(side=tk.LEFT)
        
        self.inf_count_label = tb.Label(search_box, text="", font=("Microsoft YaHei", 9))
        self.inf_count_label.pack(side=tk.LEFT, padx=15)

        # 热门高频标签
        tags_frame = tb.Frame(top_frame)
        tags_frame.pack(fill=tk.X, pady=(3, 0))
        tb.Label(tags_frame, text="🔥 热门高频检索:", font=("Microsoft YaHei", 9, "bold"), bootstyle=INFO).pack(side=tk.LEFT)
        
        hot_tags = ["辣条", "老式大面包", "月饼", "卤猪蹄", "泡椒凤爪", "火腿肠", "风味发酵乳", "果汁", "果冻", "蜜饯", "酱腌菜", "豆腐干", "坚果炒货", "食用植物油", "火锅底料", "GB/T 20977", "GB 7099", "GB 2726"]
        for tag in hot_tags:
            tb.Button(tags_frame, text=tag, command=lambda t=tag: self.inf_search_var.set(t), bootstyle="secondary-outline").pack(side=tk.LEFT, padx=2)

        # 中部：PanedWindow 分割左侧匹配结果列表与右侧详情卡片
        paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, pady=5)

        # 左侧列表
        left_frame = tb.LabelFrame(paned, text="匹配到的食品商品与标准类别", padding=6)
        paned.add(left_frame, weight=3)

        self.inf_columns = ("主要商品俗名", "推荐产品标准依据", "GB 2760 分类代码", "GB 2762 污染物分类")
        self.inf_tree = ttk.Treeview(left_frame, columns=self.inf_columns, show="headings", selectmode="browse")
        for col in self.inf_columns:
            self.inf_tree.heading(col, text=col)
            
        self.inf_tree.column("主要商品俗名", width=120)
        self.inf_tree.column("推荐产品标准依据", width=140)
        self.inf_tree.column("GB 2760 分类代码", width=160)
        self.inf_tree.column("GB 2762 污染物分类", width=130)

        inf_v_scroll = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.inf_tree.yview)
        self.inf_tree.configure(yscrollcommand=inf_v_scroll.set)
        self.inf_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        inf_v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.inf_tree.bind("<<TreeviewSelect>>", self.on_inference_select)

        # 右侧：分类详情卡片与一键带入快捷操作
        right_frame = tb.LabelFrame(paned, text="📋 分类映射详情与全链路直达操作", padding=10)
        paned.add(right_frame, weight=3)

        self.inf_detail_text = tk.Text(right_frame, font=("Consolas", 10), wrap=tk.WORD, bg="#fbfcfe", height=15)
        d_v_scroll = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.inf_detail_text.yview)
        self.inf_detail_text.configure(yscrollcommand=d_v_scroll.set)
        self.inf_detail_text.pack(fill=tk.BOTH, expand=True)
        d_v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        action_bar = tb.Frame(right_frame)
        action_bar.pack(fill=tk.X, pady=(10, 0))
        
        tb.Button(action_bar, text="👉 带入标准查询 (GB 2760)", command=self.jump_to_query_tab_with_category, bootstyle=PRIMARY).pack(side=tk.LEFT, padx=(0, 4))
        tb.Button(action_bar, text="👉 带入样品匹配", command=self.jump_to_match_tab_with_category, bootstyle=INFO).pack(side=tk.LEFT, padx=4)
        tb.Button(action_bar, text="👉 带入比例之和计算器", command=self.jump_to_ratio_tab_with_category, bootstyle=SUCCESS).pack(side=tk.LEFT, padx=4)

        self._filtered_inf_data = []
        self.filter_inference_db()

    def filter_inference_db(self):
        """根据搜索框关键字过滤商品名与执行标准反查库"""
        kw = self.inf_search_var.get().strip().lower()
        self.inf_tree.delete(*self.inf_tree.get_children())
        self._filtered_inf_data = []

        for item in FOOD_CATEGORY_INFERENCE_DB:
            names_match = any(kw in n.lower() for n in item["names"])
            std_match = any(kw in s.lower() for s in item["standards"])
            cat_match = kw in item["gb2760_cat"].lower() or kw in item["gb2762_cat"].lower()

            if not kw or names_match or std_match or cat_match:
                self._filtered_inf_data.append(item)
                main_name = item["names"][0] if item["names"] else ""
                main_std = item["standards"][0] if item["standards"] else ""
                row = (main_name, main_std, item["gb2760_cat"], item["gb2762_cat"])
                self.inf_tree.insert('', 'end', values=row)

        self.inf_count_label.config(text=f"匹配到 {len(self._filtered_inf_data)} 条分类映射数据")

        # 默认选中第一条
        children = self.inf_tree.get_children()
        if children:
            self.inf_tree.selection_set(children[0])
            self.on_inference_select()
        else:
            self.inf_detail_text.delete("1.0", tk.END)
            self.inf_detail_text.insert(tk.END, "⚠️ 未匹配到与该关键词对应的商品分类，请尝试更换关键词。")

    def on_inference_select(self, event=None):
        """选中某个推荐商品时渲染右侧详情卡片"""
        selection = self.inf_tree.selection()
        if not selection:
            return
        idx = self.inf_tree.index(selection[0])
        if idx >= len(self._filtered_inf_data):
            return

        item = self._filtered_inf_data[idx]
        self._current_selected_inf_item = item

        out = []
        out.append("=" * 55)
        out.append(f" 🍱 【目标商品/品类】: {' / '.join(item['names'])}")
        out.append("=" * 55)
        out.append(f"📜 推荐产品执行标准: \n  • " + "\n  • ".join(item['standards']))
        out.append("-" * 55)
        out.append(f"🏛️ 【GB 2760 添加剂分类】: {item['gb2760_cat']}")
        out.append(f"☣️ 【GB 2762 污染物分类】: {item['gb2762_cat']}")
        out.append(f"🦠 【GB 29921 微生物分类】: {item['gb29921_cat']}")
        out.append("-" * 55)
        out.append(f"📝 品类定义与工艺描述: \n  {item['desc']}")
        out.append("-" * 55)
        out.append(f"⚠️ 核心抽检风险指标 & 禁用提醒: \n  • " + "\n  • ".join(item['key_risk_items']))
        out.append("=" * 55)
        out.append("💡 点击下方快捷按钮，可一键将该分类带入【标准查询】、【样品匹配】或【比例之和计算器】！")

        self.inf_detail_text.delete("1.0", tk.END)
        self.inf_detail_text.insert(tk.END, "\n".join(out))

    def jump_to_query_tab_with_category(self):
        """将选中的分类一键带入标准查询 Tab"""
        if not hasattr(self, '_current_selected_inf_item') or not self._current_selected_inf_item:
            messagebox.showinfo("提示", "请先在列表中选中一个商品分类！")
            return
        cat_str = self._current_selected_inf_item["gb2760_cat"].split(' ')[0]
        # 切换到 Tab 0 (标准查询)
        self.notebook.select(0)
        self.standard_var.set("GB 2760-2024")
        self.on_standard_selected()
        self.cat_search_var.set(cat_str)
        self.filter_category_tree()

    def jump_to_match_tab_with_category(self):
        """将选中的分类一键带入样品匹配 Tab"""
        if not hasattr(self, '_current_selected_inf_item') or not self._current_selected_inf_item:
            messagebox.showinfo("提示", "请先在列表中选中一个商品分类！")
            return
        food_name = self._current_selected_inf_item["names"][0]
        # 切换到 Tab 1 (样品匹配)
        self.notebook.select(1)
        self.sample_food_var.set(food_name)
        self.match_sample()

    def jump_to_ratio_tab_with_category(self):
        """将选中的分类一键带入添加剂比例之和计算器 Tab"""
        if not hasattr(self, '_current_selected_inf_item') or not self._current_selected_inf_item:
            messagebox.showinfo("提示", "请先在列表中选中一个商品分类！")
            return
        cat_full = self._current_selected_inf_item["gb2760_cat"]
        # 切换到 Tab 3 (比例之和计算器)
        self.notebook.select(3)
        # 在下拉框中寻找最接近的分类
        found = False
        for val in self.ratio_cat_combo['values']:
            if cat_full in val or val in cat_full or cat_full.split(' ')[0] in val:
                self.ratio_food_cat_var.set(val)
                found = True
                break
        if not found:
            self.ratio_food_cat_var.set(cat_full)
        self.recalculate_ratios()

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

    # ============================================================
    # GB 29921 / GB 4789.1 微生物三级采样方案 (n, c, m, M) 判定方法
    # ============================================================
    def setup_microbiology_tab(self):
        """GB 29921 / GB 4789.1 微生物三级采样方案 (n, c, m, M) 智能判定 Tab"""
        main_frame = tb.Frame(self.tab_microbiology, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 顶部提示与说明栏
        banner = tb.Frame(main_frame, bootstyle=SECONDARY)
        banner.pack(fill=tk.X, pady=(0, 8))
        tb.Label(
            banner, 
            text="🔬 GB 29921 / GB 4789.1 微生物三级采样方案 (n, c, m, M) 批量智能判定引擎", 
            font=("Microsoft YaHei", 11, "bold"), 
            bootstyle="inverse-secondary"
        ).pack(side=tk.LEFT, padx=10, pady=6)
        tb.Label(
            banner, 
            text="判定规则：全部样品 ≤ m，或 (m < X ≤ M) 的超标样品数 ≤ c 且无任意样品 > M 时判定为合格", 
            font=("Microsoft YaHei", 9), 
            bootstyle="inverse-secondary"
        ).pack(side=tk.RIGHT, padx=10, pady=6)

        # 左右分栏
        paned = ttk.Panedwindow(main_frame, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # 左侧：配置与录入区
        left_box = tb.LabelFrame(paned, text="⚙️ 采样参数配置与实测值录入", padding=10)
        paned.add(left_box, weight=4)

        # 1. 食品类别与检测项目选择
        # 品类即时模糊/首字母搜索过滤行
        filter_row = tb.Frame(left_box)
        filter_row.pack(fill=tk.X, pady=(0, 4))
        tb.Label(filter_row, text="🔍 品类速查:", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT)
        self.micro_filter_var = tk.StringVar()
        self.micro_filter_entry = tb.Entry(filter_row, textvariable=self.micro_filter_var, font=("Microsoft YaHei", 9))
        self.micro_filter_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        self.micro_filter_entry.bind("<KeyRelease>", lambda e: self._filter_micro_categories())
        tb.Button(filter_row, text="✕ 清除", command=self._clear_micro_filter, bootstyle="secondary-outline").pack(side=tk.LEFT)

        sel_row1 = tb.Frame(left_box)
        sel_row1.pack(fill=tk.X, pady=(0, 6))
        tb.Label(sel_row1, text="食品品类:", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT)
        
        self.micro_cat_var = tk.StringVar()
        cats = list(MICROBIAL_STANDARDS_DB.keys())
        self.micro_cat_combo = ttk.Combobox(sel_row1, textvariable=self.micro_cat_var, values=cats, width=36, state="readonly", font=("Microsoft YaHei", 9))
        self.micro_cat_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.micro_cat_combo.bind("<<ComboboxSelected>>", lambda e: self._on_micro_cat_changed())

        sel_row2 = tb.Frame(left_box)
        sel_row2.pack(fill=tk.X, pady=(0, 8))
        tb.Label(sel_row2, text="检测项目:", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT)
        self.micro_item_var = tk.StringVar()
        self.micro_item_combo = ttk.Combobox(sel_row2, textvariable=self.micro_item_var, values=[], width=36, state="readonly", font=("Microsoft YaHei", 9))
        self.micro_item_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.micro_item_combo.bind("<<ComboboxSelected>>", lambda e: self._on_micro_item_changed())

        # 2. 法定三级采样参数显示与微调
        param_frame = tb.LabelFrame(left_box, text="📋 法定采样与限量参数", padding=6)
        param_frame.pack(fill=tk.X, pady=(0, 8))

        p_grid = tb.Frame(param_frame)
        p_grid.pack(fill=tk.X)

        tb.Label(p_grid, text="抽样数 (n):").grid(row=0, column=0, padx=3, pady=3, sticky="e")
        self.micro_n_var = tk.StringVar(value="5")
        tb.Entry(p_grid, textvariable=self.micro_n_var, width=5, font=("Microsoft YaHei", 9)).grid(row=0, column=1, padx=3, pady=3)

        tb.Label(p_grid, text="允许超标数 (c):").grid(row=0, column=2, padx=3, pady=3, sticky="e")
        self.micro_c_var = tk.StringVar(value="1")
        tb.Entry(p_grid, textvariable=self.micro_c_var, width=5, font=("Microsoft YaHei", 9)).grid(row=0, column=3, padx=3, pady=3)

        tb.Label(p_grid, text="限量下限 (m):").grid(row=1, column=0, padx=3, pady=3, sticky="e")
        self.micro_m_var = tk.StringVar(value="100")
        tb.Entry(p_grid, textvariable=self.micro_m_var, width=8, font=("Microsoft YaHei", 9)).grid(row=1, column=1, padx=3, pady=3)

        tb.Label(p_grid, text="限量上限 (M):").grid(row=1, column=2, padx=3, pady=3, sticky="e")
        self.micro_M_var = tk.StringVar(value="1000")
        tb.Entry(p_grid, textvariable=self.micro_M_var, width=8, font=("Microsoft YaHei", 9)).grid(row=1, column=3, padx=3, pady=3)

        tb.Label(p_grid, text="单位:").grid(row=0, column=4, padx=3, pady=3, sticky="e")
        self.micro_unit_var = tk.StringVar(value="CFU/g")
        tb.Entry(p_grid, textvariable=self.micro_unit_var, width=10, font=("Microsoft YaHei", 9)).grid(row=0, column=5, padx=3, pady=3)

        # 底部标准信息与 2026 抽检细则速查按钮
        info_row = tb.Frame(param_frame)
        info_row.pack(fill=tk.X, pady=(4, 0))
        
        self.micro_std_info_label = tb.Label(info_row, text="依据标准: GB 29921-2021 | 推荐方法: GB 4789.10", font=("Microsoft YaHei", 8), bootstyle=INFO)
        self.micro_std_info_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        tb.Button(
            info_row, 
            text="📖 2026抽检细则规范速查...", 
            command=self.open_micro_sampling_guide_dialog, 
            bootstyle="info-outline"
        ).pack(side=tk.RIGHT)

        # 3. 5件平行样实测值录入区
        samples_box = tb.LabelFrame(left_box, text="🧪 平行样实测数据录入 (支持数字、0、未检出/ND)", padding=8)
        samples_box.pack(fill=tk.X, pady=(0, 8))

        self.micro_sample_vars = [tk.StringVar(value="") for _ in range(5)]
        
        s_row = tb.Frame(samples_box)
        s_row.pack(fill=tk.X, pady=(0, 6))
        for idx in range(5):
            s_f = tb.Frame(s_row)
            s_f.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
            tb.Label(s_f, text=f"样{idx+1}:", font=("Microsoft YaHei", 8, "bold")).pack(anchor="w")
            tb.Entry(s_f, textvariable=self.micro_sample_vars[idx], font=("Microsoft YaHei", 9), width=7).pack(fill=tk.X)

        # 单行快速批量粘贴框
        paste_row = tb.Frame(samples_box)
        paste_row.pack(fill=tk.X, pady=(2, 0))
        tb.Label(paste_row, text="快速粘贴:").pack(side=tk.LEFT)
        self.micro_paste_var = tk.StringVar()
        p_entry = tb.Entry(paste_row, textvariable=self.micro_paste_var, font=("Microsoft YaHei", 9))
        p_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        p_entry.bind("<Return>", lambda e: self._apply_micro_paste())
        tb.Button(paste_row, text="解析填入", command=self._apply_micro_paste, bootstyle="secondary-outline").pack(side=tk.LEFT)

        # 4. 快捷模拟测试用例
        preset_frame = tb.LabelFrame(left_box, text="💡 快捷测试与实战案例模拟", padding=6)
        preset_frame.pack(fill=tk.X, pady=(0, 6))

        p_btns = tb.Frame(preset_frame)
        p_btns.pack(fill=tk.X)
        tb.Button(p_btns, text="✔ 全合格 (≤m)", command=self._load_micro_case_all_pass, bootstyle="success-outline").pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)
        tb.Button(p_btns, text="⚠️ 边界合格 (k≤c)", command=self._load_micro_case_boundary_pass, bootstyle="warning-outline").pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)
        tb.Button(p_btns, text="❌ 严重超标 (>M)", command=self._load_micro_case_over_M, bootstyle="danger-outline").pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)
        tb.Button(p_btns, text="❌ 超件违规 (k>c)", command=self._load_micro_case_over_c, bootstyle="danger-outline").pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)

        # 5. 立即判定按钮
        tb.Button(left_box, text="⚡ 立即执行三级采样判定", command=self.do_microbiology_judge, bootstyle=SUCCESS).pack(fill=tk.X, pady=(4, 0))

        # 右侧：判定结果卡片与详细报告
        right_box = tb.LabelFrame(paned, text="📊 微生物三级采样判定看板与合规报告", padding=8)
        paned.add(right_box, weight=5)

        # 顶部：平行样指示卡片 Canvas
        self.micro_cards_canvas = tk.Canvas(right_box, height=80, bg="#ffffff", highlightthickness=1, highlightbackground="#dee2e6")
        self.micro_cards_canvas.pack(fill=tk.X, pady=(0, 6))

        # 下部：分步计算公式与报告
        self.micro_calc_text = tk.Text(right_box, font=("Consolas", 9), wrap=tk.WORD, bg="#fbfcfe")
        m_v_scroll = ttk.Scrollbar(right_box, orient=tk.VERTICAL, command=self.micro_calc_text.yview)
        self.micro_calc_text.configure(yscrollcommand=m_v_scroll.set)
        self.micro_calc_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        m_v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 底部操作栏
        bottom_bar = tb.Frame(main_frame)
        bottom_bar.pack(fill=tk.X, pady=(8, 0))
        tb.Button(bottom_bar, text="📋 复制完整判定报告", command=self.copy_micro_report, bootstyle="info-outline").pack(side=tk.LEFT)
        tb.Button(bottom_bar, text="📄 导出判定记录(CSV)...", command=self.export_micro_report, bootstyle="secondary-outline").pack(side=tk.LEFT, padx=10)

        # 默认选中第一个
        if cats:
            self.micro_cat_combo.set(cats[0])
            self._on_micro_cat_changed()

    def _filter_micro_categories(self):
        """支持中文关键词、英文与拼音首字母模糊筛选 53+ 微生物品类"""
        kw = self.micro_filter_var.get().strip().lower()
        all_cats = list(MICROBIAL_STANDARDS_DB.keys())
        if not kw:
            self.micro_cat_combo['values'] = all_cats
            return
            
        matched = []
        for cat in all_cats:
            cat_lower = cat.lower()
            cat_pinyin = get_chinese_pinyin_initials(cat)
            if kw in cat_lower or kw in cat_pinyin:
                matched.append(cat)
                
        self.micro_cat_combo['values'] = matched
        if matched:
            if self.micro_cat_var.get() not in matched:
                self.micro_cat_combo.set(matched[0])
                self._on_micro_cat_changed()

    def _clear_micro_filter(self):
        """清空筛选框并恢复全部品类"""
        self.micro_filter_var.set("")
        all_cats = list(MICROBIAL_STANDARDS_DB.keys())
        self.micro_cat_combo['values'] = all_cats
        if all_cats:
            self.micro_cat_combo.set(all_cats[0])
            self._on_micro_cat_changed()

    def open_micro_sampling_guide_dialog(self):
        """弹出《食品安全监督抽检实施细则（2026年版）》微生物抽样规范与法定依据速查对话框"""
        cat = self.micro_cat_var.get()
        item_name = self.micro_item_var.get()
        n_val = self.micro_n_var.get()
        c_val = self.micro_c_var.get()
        m_val = self.micro_m_var.get()
        M_val = self.micro_M_var.get()
        unit_val = self.micro_unit_var.get()
        
        dialog = tb.Toplevel(self.root)
        dialog.title(f"📖 2026抽检实施细则 - 【{cat}】抽样与判定规范速查")
        dialog.geometry("780x620")
        dialog.minsize(650, 480)
        dialog.transient(self.root)
        dialog.grab_set()

        dialog.update_idletasks()
        rw = self.root.winfo_width()
        rh = self.root.winfo_height()
        rx = self.root.winfo_x()
        ry = self.root.winfo_y()
        dialog.geometry(f"+{rx + max(0, (rw - 780)//2)}+{ry + max(0, (rh - 620)//2)}")

        main_f = tb.Frame(dialog, padding=12)
        main_f.pack(fill=tk.BOTH, expand=True)

        top_f = tb.Frame(main_f)
        top_f.pack(fill=tk.X, pady=(0, 8))
        tb.Label(top_f, text=f"📋 食品安全监督抽检实施细则（2026年版）与国标判定依据", font=("Microsoft YaHei", 11, "bold"), bootstyle=PRIMARY).pack(side=tk.LEFT)

        text_box = tk.Text(main_f, font=("Microsoft YaHei", 9), wrap=tk.WORD, bg="#fbfcfe", padx=10, pady=10)
        v_scroll = ttk.Scrollbar(main_f, orient=tk.VERTICAL, command=text_box.yview)
        text_box.configure(yscrollcommand=v_scroll.set)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        text_box.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        report_lines = []
        report_lines.append("=" * 68)
        report_lines.append(f" 🔬 《食品安全监督抽检实施细则（2026年版）》微生物抽样检验全景规范")
        report_lines.append("=" * 68)
        report_lines.append(f"【当前食品品类】: {cat}")
        report_lines.append(f"【当前重点项目】: {item_name}")
        report_lines.append(f"【法定采样方案】: n = {n_val}, c = {c_val}, m = {m_val} {unit_val}, M = {M_val} {unit_val}")
        report_lines.append("-" * 68)
        
        report_lines.append("\n一、 适用标准与国抽依据 (2026版细则)")
        report_lines.append("  1. 通用致病菌安全标准: GB 29921-2021《预包装食品中致病菌限量》 / GB 31607-2021《散装即食食品中致病菌限量》")
        report_lines.append("  2. 产品专用安全标准: 参见当前品类所属国标（如 GB 2726熟肉、GB 19645巴氏奶、GB 10765婴配粉、GB 7101饮料、GB 7099糕点、GB 10136水产、GB 2714酱腌菜、GB 14934餐饮具等）")
        report_lines.append("  3. 检验方法依据: GB 4789 系列《食品安全国家标准 食品微生物学检验》（GB 4789.2 菌落总数、GB 4789.3 大肠菌群、GB 4789.4 沙门氏菌、GB 4789.10 金葡菌、GB 4789.30 单增李斯特、GB 4789.7 副溶血弧菌等）")

        report_lines.append("\n二、 GB 29921 / GB 4789.1 三级采样方案 (n, c, m, M) 判定准则")
        report_lines.append("  • 抽样件数 (n): 同一批次产品必须独立采集 n 件（通常为 5 件，粉状婴配为 3 件）样品分别进行定量检测。")
        report_lines.append("  • 容限件数 (c): 最大允许超出限量下限 m 但不得超过上限 M 的样品件数。")
        report_lines.append("  • 限量阈值 (m 与 M):")
        report_lines.append("    - m: 微生物指标的可接受安全水平限量值；")
        report_lines.append("    - M: 微生物指标的最高安全控制红线值（任何单样实测值绝对不允许突破 M）；")
        report_lines.append("  • 合格判定准则: 全部样品 ≤ m，或处于 (m, M] 之间的超标样品数 k ≤ c 且无任意样品 > M 时，判定整批产品微生物合格；")
        report_lines.append("  • 不合格判定准则: 任意样品实测值 > M，或处于 (m, M] 区间样品数 k > c 时，判定整批产品微生物不合格！")
        report_lines.append("  • 零容忍指标: 致病菌 (如沙门氏菌、单增李斯特菌、O157:H7) 规定不得检出 (m=0)，5件平行样中只要有 1 件检出即判定为严重不合格。")

        report_lines.append("\n三、 2026年细则现场抽样与封样关键质控要点")
        report_lines.append("  1. 生产环节抽样: 抽取同一批次包装完好产品，抽取件数必须严格满足 n=5 件独立包装，每件包装量需满足检验及留样要求。")
        report_lines.append("  2. 流通与餐饮环节散装抽样: 必须执行无菌操作！抽样人员须佩戴无菌手套，使用经75%酒精擦拭消毒或高压灭菌的取样工具，分装入无菌采样袋，封口并加贴封条。")
        report_lines.append("  3. 冷链运输温控红线: 冷藏样品运输温度必须维持在 0℃～4℃，冷冻样品必须 ≤ -18℃，全程配备车载温度记录仪与冷媒保温箱，严防运输升温导致微生物异常增殖。")
        report_lines.append("  4. 样品流转时效: 微生物样品送达实验室后应立即检验；如不能立即检验，冷藏样品应在 24 小时内开始检验。")

        report_lines.append("\n四、 ⚠️ 法定异议与复检特别规定 (极为重要)")
        report_lines.append("  • 《食品安全抽样检验管理办法》第四十条明确规定:")
        report_lines.append("    【微生物检验项目一律不予复检】！")
        report_lines.append("  • 法律与科学依据: 微生物是具有活性的生物体，随时间、温度和贮存环境发生不可逆的生长代谢或死亡，复检无法还原初检时的真实微生物状态。因此初检机构出具的判定结论即为法定的最终结论。")
        report_lines.append("=" * 68)

        full_txt = "\n".join(report_lines)
        text_box.insert(tk.END, full_txt)
        text_box.configure(state="disabled")

        bot_f = tb.Frame(main_f)
        bot_f.pack(fill=tk.X, pady=(8, 0))

        def copy_guide():
            self.root.clipboard_clear()
            self.root.clipboard_append(full_txt)
            messagebox.showinfo("复制成功", "已将完整 2026 抽检细则抽样规范与判定依据复制到剪贴板！")

        tb.Button(bot_f, text="📋 一键复制抽检规范与依据", command=copy_guide, bootstyle="info-outline").pack(side=tk.LEFT)
        tb.Button(bot_f, text="✔ 我已知晓并关闭", command=dialog.destroy, bootstyle=SUCCESS).pack(side=tk.RIGHT)

    def _on_micro_cat_changed(self):
        """食品类别切换时刷新对应检测项目列表"""
        cat = self.micro_cat_var.get()
        items_db = MICROBIAL_STANDARDS_DB.get(cat, [])
        item_names = [it["item"] for it in items_db]
        self.micro_item_combo['values'] = item_names
        if item_names:
            self.micro_item_combo.set(item_names[1] if len(item_names) > 1 else item_names[0])
            self._on_micro_item_changed()

    def _on_micro_item_changed(self):
        """检测项目切换时自动预填法定参数"""
        cat = self.micro_cat_var.get()
        item_name = self.micro_item_var.get()
        items_db = MICROBIAL_STANDARDS_DB.get(cat, [])
        matched = None
        for it in items_db:
            if it["item"] == item_name:
                matched = it
                break
        if matched:
            self.micro_n_var.set(str(matched.get("n", 5)))
            self.micro_c_var.set(str(matched.get("c", 0)))
            self.micro_m_var.set(str(matched.get("m", 0)))
            self.micro_M_var.set(str(matched.get("M", 0)))
            self.micro_unit_var.set(str(matched.get("unit", "CFU/g")))
            std_str = matched.get("std", "GB 29921-2021")
            method_str = matched.get("method", "")
            self.micro_std_info_label.config(text=f"依据标准: {std_str} | 推荐检验方法: {method_str}")
            
        self._load_micro_case_all_pass()

    def _apply_micro_paste(self):
        """解析单行快速粘贴数据并分发至各样品输入框"""
        raw = self.micro_paste_var.get().strip()
        if not raw:
            return
        # 按逗号、空格、分号、制表符分割
        tokens = re.split(r'[,; \t]+', raw)
        tokens = [t.strip() for t in tokens if t.strip()]
        for idx in range(5):
            if idx < len(tokens):
                self.micro_sample_vars[idx].set(tokens[idx])
            else:
                self.micro_sample_vars[idx].set("")
        self.do_microbiology_judge()

    def _load_micro_case_all_pass(self):
        """加载全部样品 <= m 的完全合规用例"""
        m_val = extract_first_number(self.micro_m_var.get()) or 0
        if m_val <= 0:
            for var in self.micro_sample_vars:
                var.set("未检出")
        else:
            vals = [f"{int(m_val * 0.2)}", f"{int(m_val * 0.5)}", f"{int(m_val * 0.3)}", f"{int(m_val * 0.8)}", f"{int(m_val * 0.1)}"]
            for idx, v in enumerate(vals):
                self.micro_sample_vars[idx].set(v)
        self.do_microbiology_judge()

    def _load_micro_case_boundary_pass(self):
        """加载有 <= c 件样品处于 (m, M] 区间的边界合规用例"""
        m_val = extract_first_number(self.micro_m_var.get()) or 0
        M_val = extract_first_number(self.micro_M_var.get()) or 0
        c_val = int(extract_first_number(self.micro_c_var.get()) or 1)
        
        if M_val <= 0:
            for idx, var in enumerate(self.micro_sample_vars):
                var.set("未检出")
        else:
            mid_val = int((m_val + M_val) / 2)
            for idx in range(5):
                if idx < c_val:
                    self.micro_sample_vars[idx].set(str(mid_val))
                else:
                    self.micro_sample_vars[idx].set(str(int(m_val * 0.4)))
        self.do_microbiology_judge()

    def _load_micro_case_over_M(self):
        """加载存在 1 件样品 > M 的严重超标违规用例"""
        M_val = extract_first_number(self.micro_M_var.get()) or 0
        m_val = extract_first_number(self.micro_m_var.get()) or 0
        if M_val <= 0:
            self.micro_sample_vars[0].set("检出")
            for idx in range(1, 5):
                self.micro_sample_vars[idx].set("未检出")
        else:
            self.micro_sample_vars[0].set(str(int(M_val * 1.5)))
            for idx in range(1, 5):
                self.micro_sample_vars[idx].set(str(int(m_val * 0.5)))
        self.do_microbiology_judge()

    def _load_micro_case_over_c(self):
        """加载处于 (m, M] 区间件数 k > c 的超件违规用例"""
        m_val = extract_first_number(self.micro_m_var.get()) or 0
        M_val = extract_first_number(self.micro_M_var.get()) or 0
        c_val = int(extract_first_number(self.micro_c_var.get()) or 1)
        if M_val <= 0:
            for idx, var in enumerate(self.micro_sample_vars):
                var.set("检出")
        else:
            mid_val = int((m_val + M_val) * 0.7)
            for idx in range(5):
                if idx <= c_val:
                    self.micro_sample_vars[idx].set(str(mid_val))
                else:
                    self.micro_sample_vars[idx].set(str(int(m_val * 0.3)))
        self.do_microbiology_judge()

    def do_microbiology_judge(self):
        """执行微生物三级采样核心算法判定并渲染看板"""
        cat = self.micro_cat_var.get()
        item_name = self.micro_item_var.get()
        unit_str = self.micro_unit_var.get().strip()
        
        try:
            n_target = int(extract_first_number(self.micro_n_var.get()) or 5)
            c_target = int(extract_first_number(self.micro_c_var.get()) or 0)
            m_val = extract_first_number(self.micro_m_var.get())
            M_val = extract_first_number(self.micro_M_var.get())
        except Exception:
            messagebox.showerror("错误", "采样参数格式有误，请确保 n, c, m, M 为有效数字！")
            return

        if m_val is None:
            m_val = 0.0
        if M_val is None:
            M_val = 0.0

        is_zero_tolerance = (m_val == 0 and M_val == 0)

        # 收集有效录入的样品数据
        sample_results = []
        for idx in range(n_target):
            if idx < len(self.micro_sample_vars):
                raw_str = self.micro_sample_vars[idx].get().strip()
            else:
                raw_str = ""
                
            if not raw_str:
                num = 0.0 if is_zero_tolerance else 0.0
                disp = "未填 (视作0)"
            elif raw_str in ("未检出", "ND", "nd", "阴性", "-", "0"):
                num = 0.0
                disp = raw_str
            elif raw_str in ("检出", "阳性", "+"):
                num = 1.0 if is_zero_tolerance else 999999.0
                disp = raw_str
            else:
                num_val = extract_first_number(raw_str)
                if num_val is not None:
                    num = num_val
                    disp = f"{num_val:g}"
                else:
                    num = 0.0
                    disp = raw_str

            # 单样判定状态: pass_m (<=m), pass_M (m < X <= M), fail (>M 或 检出)
            if is_zero_tolerance:
                if num > 0 or raw_str in ("检出", "阳性", "+"):
                    s_status = "fail"
                    s_tag = "✘ 检出(违规)"
                else:
                    s_status = "pass_m"
                    s_tag = "✔ 未检出(合格)"
            else:
                if M_val > 0 and num > M_val:
                    s_status = "fail"
                    s_tag = f"✘ >M ({num:g}>{M_val:g})"
                elif num > m_val:
                    s_status = "pass_M"
                    s_tag = f"⚠️ (m, M] ({num:g})"
                else:
                    s_status = "pass_m"
                    s_tag = f"✔ ≤m ({num:g})"

            sample_results.append({
                "sample_no": f"样{idx+1}",
                "raw_str": raw_str,
                "num": num,
                "disp": disp,
                "status": s_status,
                "tag": s_tag
            })

        # 绘制顶部 5 件平行样指示卡片
        self._draw_micro_sample_cards(sample_results)

        # 汇总统计
        count_over_M = sum(1 for s in sample_results if s["status"] == "fail")
        count_in_m_M = sum(1 for s in sample_results if s["status"] == "pass_M")
        count_le_m = sum(1 for s in sample_results if s["status"] == "pass_m")

        # 综合判定逻辑
        if is_zero_tolerance:
            is_overall_pass = (count_over_M == 0)
        else:
            is_overall_pass = (count_over_M == 0) and (count_in_m_M <= c_target)

        # 构建文本报告
        self.micro_calc_text.delete("1.0", tk.END)
        out = []
        out.append("=" * 60)
        out.append(" 🔬 GB 29921 / GB 4789.1 食品微生物三级采样智能判定报告")
        out.append("=" * 60)
        out.append(f"【食品品类】: {cat}")
        out.append(f"【检测项目】: {item_name}")
        out.append(f"【采样方案】: n={n_target}, c={c_target}, m={m_val:g} {unit_str}, M={M_val:g} {unit_str}")
        out.append("-" * 60)
        out.append("🧪 【各平行样品实测检出明细】:")
        for s in sample_results:
            out.append(f"  • {s['sample_no']}: 实测值 = {s['disp']} {unit_str:<8} ➔ {s['tag']}")

        out.append("\n📊 【三级采样方案统计指标】:")
        out.append(f"  • 抽样检验件数 (n): {n_target} 件 (实收 {len(sample_results)} 件)")
        out.append(f"  • 最大允许超标件数 (c): {c_target} 件")
        out.append(f"  • 实测处于允许超标区间 (m < X ≤ M) 件数 (k): {count_in_m_M} 件")
        out.append(f"  • 实测严重超出上限 (> M) 件数: {count_over_M} 件")

        out.append("\n🧮 【判定推导过程】:")
        if is_zero_tolerance:
            out.append(f"  • 规则依据: 该致病菌为零容忍指标（不得检出，m=0）。")
            if is_overall_pass:
                out.append(f"  • 比对推导: 全部 {n_target} 件平行样均未检出致病菌。")
            else:
                out.append(f"  • 比对推导: 存在 {count_over_M} 件样品检出致病菌！")
        else:
            out.append(f"  • 规则① (上限红线): 任意单样实测值不得超过 M ({M_val:g} {unit_str})。")
            if count_over_M == 0:
                out.append(f"    ✔ 符合！所有平行样实测值均 ≤ M。")
            else:
                out.append(f"    ✘ 违规！发现 {count_over_M} 件样品实测值 > M ({M_val:g})！")

            out.append(f"  • 规则② (件数容限): 位于 (m, M] 区间样品数 k ({count_in_m_M}) 必须 ≤ c ({c_target})。")
            if count_in_m_M <= c_target:
                out.append(f"    ✔ 符合！区间超标件数 k={count_in_m_M} ≤ 允许值 c={c_target}。")
            else:
                out.append(f"    ✘ 违规！区间超标件数 k={count_in_m_M} 超过允许超标件数 c={c_target}！")

        out.append("\n" + "=" * 60)
        if is_overall_pass:
            out.append(" 🏆 综合判定结论: 【✔ 该批次产品微生物检验合格】")
            out.append("    (符合 GB 29921-2021 / 相关产品标准三级采样限量要求)")
        else:
            out.append(" 🚨 综合判定结论: 【✘ 该批次产品微生物检验不合格！】")
            if count_over_M > 0:
                out.append(f"    (原因: 存在 {count_over_M} 件样品突破 M 限量最高安全红线)")
            elif count_in_m_M > c_target:
                out.append(f"    (原因: 处于(m,M]区间的样品数 k={count_in_m_M} 超出允许件数 c={c_target})")
        out.append("=" * 60)

        self.micro_calc_text.insert(tk.END, "\n".join(out))

    def _draw_micro_sample_cards(self, sample_results):
        """在 Canvas 上动态绘制平行样卡片与红黄绿状态灯"""
        if not hasattr(self, 'micro_cards_canvas') or not self.micro_cards_canvas.winfo_exists():
            return
            
        self.micro_cards_canvas.delete("all")
        c_w = self.micro_cards_canvas.winfo_width()
        if c_w < 100:
            c_w = 480
        c_h = self.micro_cards_canvas.winfo_height()
        if c_h < 40:
            c_h = 80

        self.micro_cards_canvas.create_rectangle(0, 0, c_w, c_h, fill="#fdfdfe", outline="")

        n = len(sample_results)
        if n == 0:
            return

        card_gap = 8
        margin = 12
        card_w = (c_w - margin * 2 - (n - 1) * card_gap) / n
        card_h = c_h - 16

        for idx, s in enumerate(sample_results):
            x1 = margin + idx * (card_w + card_gap)
            y1 = 8
            x2 = x1 + card_w
            y2 = y1 + card_h

            # 背景色与边框
            if s["status"] == "pass_m":
                bg = "#dafbe1"
                fg = "#1a7f37"
                lamp = "#2ea44f"
                status_title = "✔ ≤m 合格"
            elif s["status"] == "pass_M":
                bg = "#fff8c5"
                fg = "#9a6700"
                lamp = "#d4a72c"
                status_title = "⚠️ (m, M]"
            else:
                bg = "#ffebe9"
                fg = "#cf222e"
                lamp = "#cf222e"
                status_title = "✘ >M 超标"

            # 绘制圆角矩形卡片
            self.micro_cards_canvas.create_rectangle(x1, y1, x2, y2, fill=bg, outline=fg, width=1)

            # 样品名称与状态指示灯
            self.micro_cards_canvas.create_oval(x1 + 6, y1 + 8, x1 + 14, y1 + 16, fill=lamp, outline="")
            self.micro_cards_canvas.create_text(x1 + 18, y1 + 12, anchor="w", text=s["sample_no"], font=("Microsoft YaHei", 9, "bold"), fill="#212529")

            # 实测读数
            self.micro_cards_canvas.create_text((x1 + x2) / 2, y1 + 32, text=s["disp"], font=("Consolas", 10, "bold"), fill=fg)

            # 状态标签
            self.micro_cards_canvas.create_text((x1 + x2) / 2, y1 + 48, text=status_title, font=("Microsoft YaHei", 8), fill=fg)

    def copy_micro_report(self):
        """复制微生物判定报告到剪贴板"""
        txt = self.micro_calc_text.get("1.0", tk.END).strip()
        if not txt:
            messagebox.showinfo("提示", "当前无判定报告可复制！")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(txt)
        messagebox.showinfo("复制成功", "已将完整微生物三级采样判定报告复制到剪贴板！\n可直接粘贴至检验底稿或报告中。")

    def export_micro_report(self):
        """导出微生物判定记录为 CSV"""
        txt = self.micro_calc_text.get("1.0", tk.END).strip()
        if not txt:
            messagebox.showinfo("提示", "当前无判定记录可导出！")
            return
            
        file_path = filedialog.asksaveasfilename(
            title="导出微生物三级采样判定单",
            defaultextension=".csv",
            filetypes=[("CSV表格文件 (Excel兼容)", "*.csv"), ("文本文件", "*.txt"), ("所有文件", "*.*")],
            initialfile=f"GB29921_微生物三级采样判定单_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        if not file_path:
            return
            
        try:
            with open(file_path, 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(["项目类别", "食品微生物三级采样(n, c, m, M)检验判定单"])
                writer.writerow([])
                for line in txt.split('\n'):
                    writer.writerow([line])
            messagebox.showinfo("导出成功", f"已成功导出微生物判定单到：\n{file_path}")
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
        
        for col, default_w in default_widths.items():
            if col in hidden_cols:
                tree.column(col, width=0, minwidth=0, stretch=False)
            else:
                w = saved_widths.get(col, default_w)
                try:
                    w_int = int(w)
                    # 针对检验方法列，若历史配置过窄则自动纠偏至 280px 防截断
                    if "方法" in col and w_int < 260:
                        w_int = max(default_w, 280)
                    tree.column(col, width=w_int, minwidth=30, stretch=True)
                except Exception:
                    tree.column(col, width=default_w, minwidth=30, stretch=True)

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



    # ============================================================
    # Tab 7: GB 2760-2024 新旧国标对比引擎
    # ============================================================
    def setup_gb2760_diff_tab(self):
        """GB 2760-2024 新旧国标添加剂重大变更与禁用红线对比 Tab UI"""
        main_frame = tb.Frame(self.tab_gb2760_diff, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 顶部提示栏
        banner = tb.Frame(main_frame, bootstyle=WARNING)
        banner.pack(fill=tk.X, pady=(0, 8))
        tb.Label(banner, text="🚨 GB 2760-2024 新版食品安全国家标准重大变更与禁用红线穿透对比引擎", font=("Microsoft YaHei", 10, "bold"), bootstyle="inverse-warning").pack(side=tk.LEFT, padx=10, pady=6)
        tb.Label(banner, text="重点监控脱氢乙酸钠全面禁用、落葵红收窄、甜味剂限量下调等高风险项", font=("Microsoft YaHei", 8), bootstyle="inverse-warning").pack(side=tk.RIGHT, padx=10, pady=6)

        paned = ttk.Panedwindow(main_frame, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # 左侧：添加剂选择与搜索
        left_box = tb.LabelFrame(paned, text="🔍 添加剂选择与变更概要", padding=8)
        paned.add(left_box, weight=3)

        tb.Label(left_box, text="快速检索添加剂 (支持拼音/汉字):", font=("Microsoft YaHei", 9, "bold")).pack(anchor="w", pady=(0, 4))
        self.gb2760_diff_filter_var = tk.StringVar()
        f_ent = tb.Entry(left_box, textvariable=self.gb2760_diff_filter_var, font=("Microsoft YaHei", 9))
        f_ent.pack(fill=tk.X, pady=(0, 6))
        f_ent.bind("<KeyRelease>", lambda e: self._filter_gb2760_diff_list())

        self.gb2760_diff_listbox = tk.Listbox(left_box, font=("Microsoft YaHei", 9), activestyle="none", height=8)
        self.gb2760_diff_listbox.pack(fill=tk.X, pady=(0, 6))
        self.gb2760_diff_listbox.bind("<<ListboxSelect>>", lambda e: self._on_gb2760_diff_selected())

        # 变更概要卡片
        self.gb2760_diff_summary_box = tb.LabelFrame(left_box, text="📋 2024版核心修订要点", padding=6)
        self.gb2760_diff_summary_box.pack(fill=tk.BOTH, expand=True)
        self.gb2760_diff_summary_label = tb.Label(self.gb2760_diff_summary_box, text="", font=("Microsoft YaHei", 9), wraplength=350, justify=tk.LEFT)
        self.gb2760_diff_summary_label.pack(anchor="w", fill=tk.BOTH, expand=True)

        # 右侧：并排两版标准限量对比表格
        right_box = tb.LabelFrame(paned, text="📊 新旧国标并排对比明细表 (GB 2760-2014 vs GB 2760-2024)", padding=8)
        paned.add(right_box, weight=7)

        # 表格配置
        cols = ("food_cat", "gb2014", "gb2024", "risk", "advice")
        self.gb2760_diff_tree = ttk.Treeview(right_box, columns=cols, show="headings", height=12)
        
        self.gb2760_diff_tree.heading("food_cat", text="适用食品品类 ↕", command=lambda: self.sort_tree_column(self.gb2760_diff_tree, "food_cat", False))
        self.gb2760_diff_tree.heading("gb2014", text="GB 2760-2014 (旧国标) ↕", command=lambda: self.sort_tree_column(self.gb2760_diff_tree, "gb2014", False))
        self.gb2760_diff_tree.heading("gb2024", text="GB 2760-2024 (新国标) ↕", command=lambda: self.sort_tree_column(self.gb2760_diff_tree, "gb2024", False))
        self.gb2760_diff_tree.heading("risk", text="变更风险等级 ↕", command=lambda: self.sort_tree_column(self.gb2760_diff_tree, "risk", False))
        self.gb2760_diff_tree.heading("advice", text="企业合规与配方替代建议 ↕", command=lambda: self.sort_tree_column(self.gb2760_diff_tree, "advice", False))

        self.gb2760_diff_tree.column("food_cat", width=180, minwidth=120)
        self.gb2760_diff_tree.column("gb2014", width=160, minwidth=110)
        self.gb2760_diff_tree.column("gb2024", width=200, minwidth=140)
        self.gb2760_diff_tree.column("risk", width=100, minwidth=80, anchor="center")
        self.gb2760_diff_tree.column("advice", width=240, minwidth=160)

        d_v_scroll = ttk.Scrollbar(right_box, orient=tk.VERTICAL, command=self.gb2760_diff_tree.yview)
        self.gb2760_diff_tree.configure(yscrollcommand=d_v_scroll.set)
        self.gb2760_diff_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        d_v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 底部操作
        bot_bar = tb.Frame(main_frame)
        bot_bar.pack(fill=tk.X, pady=(6, 0))
        tb.Button(bot_bar, text="📋 一键复制当前添加剂新旧对比报告", command=self.copy_gb2760_diff_report, bootstyle="info-outline").pack(side=tk.LEFT)
        tb.Button(bot_bar, text="📄 导出新旧国标对比记录(CSV)...", command=self.export_gb2760_diff_report, bootstyle="secondary-outline").pack(side=tk.LEFT, padx=10)

        # 初始填充列表
        self._filter_gb2760_diff_list()
        if list(GB2760_DIFF_DATABASE.keys()):
            self.gb2760_diff_listbox.selection_set(0)
            self._on_gb2760_diff_selected()

    def _filter_gb2760_diff_list(self):
        kw = self.gb2760_diff_filter_var.get().strip().lower()
        self.gb2760_diff_listbox.delete(0, tk.END)
        all_keys = list(GB2760_DIFF_DATABASE.keys())
        for k in all_keys:
            py = get_chinese_pinyin_initials(k)
            if not kw or kw in k.lower() or kw in py:
                self.gb2760_diff_listbox.insert(tk.END, k)

    def _on_gb2760_diff_selected(self):
        sel = self.gb2760_diff_listbox.curselection()
        if not sel:
            return
        add_name = self.gb2760_diff_listbox.get(sel[0])
        data = GB2760_DIFF_DATABASE.get(add_name)
        if not data:
            return

        # 刷新概要
        status = data.get("status", "")
        fn = data.get("function", "")
        summary = data.get("summary", "")
        self.gb2760_diff_summary_label.config(text=f"【功能分类】: {fn}\n【变更定性】: {status}\n\n【核心修订】:\n{summary}")

        # 刷新对比表格
        for item in self.gb2760_diff_tree.get_children():
            self.gb2760_diff_tree.delete(item)

        for row in data.get("diff_items", []):
            risk = row.get("risk_level", "GREEN")
            risk_label = "🚨 严重禁用" if risk == "RED" else ("⚠️ 限量下调" if risk == "YELLOW" else "✔ 维持/微调")
            self.gb2760_diff_tree.insert("", tk.END, values=(
                row.get("food_cat", ""),
                row.get("gb2014", ""),
                row.get("gb2024", ""),
                risk_label,
                row.get("advice", "")
            ))

    def copy_gb2760_diff_report(self):
        sel = self.gb2760_diff_listbox.curselection()
        if not sel:
            return
        add_name = self.gb2760_diff_listbox.get(sel[0])
        data = GB2760_DIFF_DATABASE.get(add_name)
        if not data:
            return
        lines = [f"=== GB 2760-2024 新旧国标对比报告: {add_name} ==="]
        lines.append(f"功能: {data.get('function')} | 状态: {data.get('status')}")
        lines.append(f"概要: {data.get('summary')}\n")
        lines.append(f"{'食品品类':<22} {'2014旧标':<25} {'2024新标':<28} {'建议'}")
        lines.append("-" * 75)
        for r in data.get("diff_items", []):
            lines.append(f"{r.get('food_cat'):<22} {r.get('gb2014'):<25} {r.get('gb2024'):<28} {r.get('advice')}")
        full_t = "\n".join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(full_t)
        messagebox.showinfo("复制成功", "已将新旧国标对比报告复制到剪贴板！")

    def export_gb2760_diff_report(self):
        sel = self.gb2760_diff_listbox.curselection()
        if not sel:
            return
        add_name = self.gb2760_diff_listbox.get(sel[0])
        data = GB2760_DIFF_DATABASE.get(add_name)
        if not data:
            return
        file_path = filedialog.asksaveasfilename(
            title="导出GB 2760新旧国标对比记录",
            defaultextension=".csv",
            filetypes=[("CSV表格文件", "*.csv")],
            initialfile=f"GB2760新旧对比_{add_name[:8]}_{datetime.now().strftime('%Y%m%d')}.csv"
        )
        if not file_path:
            return
        try:
            with open(file_path, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.writer(f)
                w.writerow(["GB 2760-2024 新旧国标添加剂变更对比报告"])
                w.writerow(["添加剂名称", add_name, "功能类别", data.get("function"), "变更状态", data.get("status")])
                w.writerow(["核心修订概要", data.get("summary")])
                w.writerow([])
                w.writerow(["适用食品品类", "GB 2760-2014 旧国标规定", "GB 2760-2024 新国标规定", "风险等级", "企业合规与配方建议"])
                for r in data.get("diff_items", []):
                    w.writerow([r.get("food_cat"), r.get("gb2014"), r.get("gb2024"), r.get("risk_level"), r.get("advice")])
            messagebox.showinfo("导出成功", f"文件已成功保存至:\n{file_path}")
        except Exception as e:
            messagebox.showerror("导出失败", f"保存出错: {e}")

    # ============================================================
    # Tab 8: GB 2762 浓缩物与干制品脱水折算换算
    # ============================================================
    def setup_dehydration_ratio_tab(self):
        """GB 2762 / GB 2761 浓缩物与干制品脱水折算换算 Tab UI"""
        main_frame = tb.Frame(self.tab_dehydration, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        banner = tb.Frame(main_frame, bootstyle=INFO)
        banner.pack(fill=tk.X, pady=(0, 8))
        tb.Label(banner, text="🧮 GB 2762 / GB 2761 浓缩物与干制品'脱水率/折算系数'智能折算判定器", font=("Microsoft YaHei", 10, "bold"), bootstyle="inverse-info").pack(side=tk.LEFT, padx=10, pady=6)
        tb.Label(banner, text="依据 GB 2762-2022 第3.5条 / GB 2761-2022 第3.3条 法定折算公式推导", font=("Microsoft YaHei", 8), bootstyle="inverse-info").pack(side=tk.RIGHT, padx=10, pady=6)

        paned = ttk.Panedwindow(main_frame, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # 左侧：配置与实测值录入
        left_box = tb.LabelFrame(paned, text="⚙️ 加工品参数与实测值录入", padding=10)
        paned.add(left_box, weight=4)

        # 1. 加工品类选择
        r1 = tb.Frame(left_box)
        r1.pack(fill=tk.X, pady=(0, 6))
        tb.Label(r1, text="加工食品品类:", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT)
        self.dehydration_cat_var = tk.StringVar()
        cats = list(DEHYDRATION_STANDARDS_DB.keys())
        self.dehydration_cat_combo = ttk.Combobox(r1, textvariable=self.dehydration_cat_var, values=cats, state="readonly", font=("Microsoft YaHei", 9))
        self.dehydration_cat_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        self.dehydration_cat_combo.bind("<<ComboboxSelected>>", lambda e: self._on_dehydration_cat_changed())

        # 2. 检测项目选择
        r2 = tb.Frame(left_box)
        r2.pack(fill=tk.X, pady=(0, 6))
        tb.Label(r2, text="检测污染指标:", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT)
        self.dehydration_item_var = tk.StringVar()
        self.dehydration_item_combo = ttk.Combobox(r2, textvariable=self.dehydration_item_var, values=[], state="readonly", font=("Microsoft YaHei", 9))
        self.dehydration_item_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        self.dehydration_item_combo.bind("<<ComboboxSelected>>", lambda e: self.do_dehydration_calculate())

        # 3. 对应生鲜原料基准与限量显示
        self.dehydration_fresh_info_label = tb.Label(left_box, text="生鲜原料基准: 鲜食用菌及其制品 | 鲜品铅限量 ≤ 0.1 mg/kg", font=("Microsoft YaHei", 8), bootstyle=PRIMARY)
        self.dehydration_fresh_info_label.pack(anchor="w", pady=(0, 6))

        # 4. 实测值与折算模式
        p_frame = tb.LabelFrame(left_box, text="📋 折算系数与实测数据", padding=8)
        p_frame.pack(fill=tk.X, pady=(0, 8))

        m_row = tb.Frame(p_frame)
        m_row.pack(fill=tk.X, pady=(0, 6))
        tb.Label(m_row, text="干品/浓缩物实测值:").pack(side=tk.LEFT)
        self.dehydration_meas_var = tk.StringVar(value="0.65")
        tb.Entry(m_row, textvariable=self.dehydration_meas_var, width=10, font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=6)
        self.dehydration_unit_label = tb.Label(m_row, text="mg/kg", font=("Microsoft YaHei", 8))
        self.dehydration_unit_label.pack(side=tk.LEFT)

        # 模式切换
        mode_row = tb.Frame(p_frame)
        mode_row.pack(fill=tk.X, pady=(0, 6))
        self.dehydration_mode_var = tk.StringVar(value="rate")
        tb.Radiobutton(mode_row, text="按脱水率折算 (%)", variable=self.dehydration_mode_var, value="rate", command=self.do_dehydration_calculate).pack(side=tk.LEFT)
        tb.Radiobutton(mode_row, text="按浓缩倍数折算", variable=self.dehydration_mode_var, value="factor", command=self.do_dehydration_calculate).pack(side=tk.LEFT, padx=10)

        param_row = tb.Frame(p_frame)
        param_row.pack(fill=tk.X)
        tb.Label(param_row, text="折算参数值:").pack(side=tk.LEFT)
        self.dehydration_param_var = tk.StringVar(value="88.0")
        tb.Entry(param_row, textvariable=self.dehydration_param_var, width=10, font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=6)
        self.dehydration_param_unit_label = tb.Label(param_row, text="%", font=("Microsoft YaHei", 8))
        self.dehydration_param_unit_label.pack(side=tk.LEFT)

        tb.Button(left_box, text="⚡ 立即执行脱水/浓缩智能折算判定", command=self.do_dehydration_calculate, bootstyle=SUCCESS).pack(fill=tk.X, pady=(6, 0))

        # 右侧：计算推导与合规判定看板
        right_box = tb.LabelFrame(paned, text="📊 折算判定看板与分步推导报告", padding=8)
        paned.add(right_box, weight=6)

        self.dehydration_status_banner = tb.Label(right_box, text="✔ 判定结果计算完成", font=("Microsoft YaHei", 10, "bold"), bootstyle="success", anchor="center", padding=6)
        self.dehydration_status_banner.pack(fill=tk.X, pady=(0, 6))

        self.dehydration_report_text = tk.Text(right_box, font=("Consolas", 9), wrap=tk.WORD, bg="#ffffff", padx=8, pady=8)
        dh_scroll = ttk.Scrollbar(right_box, orient=tk.VERTICAL, command=self.dehydration_report_text.yview)
        self.dehydration_report_text.configure(yscrollcommand=dh_scroll.set)
        self.dehydration_report_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        dh_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 底部操作
        bot_bar2 = tb.Frame(main_frame)
        bot_bar2.pack(fill=tk.X, pady=(6, 0))
        tb.Button(bot_bar2, text="📋 复制完整折算报告", command=self.copy_dehydration_report, bootstyle="info-outline").pack(side=tk.LEFT)

        if cats:
            self.dehydration_cat_combo.set(cats[0])
            self._on_dehydration_cat_changed()

    def _on_dehydration_cat_changed(self):
        cat = self.dehydration_cat_var.get()
        data = DEHYDRATION_STANDARDS_DB.get(cat, {})
        items = [it["item"] for it in data.get("items", [])]
        self.dehydration_item_combo["values"] = items
        if items:
            self.dehydration_item_combo.set(items[0])
        
        # 预设参数
        if self.dehydration_mode_var.get() == "rate":
            self.dehydration_param_var.set(str(data.get("default_dehydration_rate", 85.0)))
            self.dehydration_param_unit_label.config(text="%")
        else:
            self.dehydration_param_var.set(str(data.get("default_conc_factor", 6.0)))
            self.dehydration_param_unit_label.config(text="倍")

        self.do_dehydration_calculate()

    def do_dehydration_calculate(self):
        cat = self.dehydration_cat_var.get()
        data = DEHYDRATION_STANDARDS_DB.get(cat, {})
        item_name = self.dehydration_item_var.get()
        mode = self.dehydration_mode_var.get()

        matched_item = None
        for it in data.get("items", []):
            if it["item"] == item_name:
                matched_item = it
                break

        if not matched_item:
            return

        try:
            meas_val = float(self.dehydration_meas_var.get().strip())
            param_val = float(self.dehydration_param_var.get().strip())
        except Exception:
            return

        fresh_limit = matched_item.get("fresh_limit", 0.1)
        unit = matched_item.get("unit", "mg/kg")
        std_name = data.get("std", "GB 2762-2022")
        fresh_name = data.get("fresh_name", "")

        self.dehydration_fresh_info_label.config(text=f"生鲜原料基准: {fresh_name} | {item_name} 限量 ≤ {fresh_limit} {unit}")
        self.dehydration_unit_label.config(text=unit)

        # 折算计算
        if mode == "rate":
            self.dehydration_param_unit_label.config(text="%")
            # C_fresh = C_dry * (1 - rate / 100)
            fresh_equiv = meas_val * (1.0 - param_val / 100.0)
            formula_str = f"折算等效值 C_fresh = 实测干品值 ({meas_val} {unit}) × (1 - 脱水率 {param_val}%) = {fresh_equiv:.4f} {unit}"
        else:
            self.dehydration_param_unit_label.config(text="倍")
            # C_fresh = C_dry / factor
            fresh_equiv = meas_val / max(0.0001, param_val)
            formula_str = f"折算等效值 C_fresh = 实测浓缩值 ({meas_val} {unit}) ÷ 浓缩倍数 ({param_val}倍) = {fresh_equiv:.4f} {unit}"

        is_pass = (fresh_equiv <= fresh_limit)

        if is_pass:
            self.dehydration_status_banner.config(text=f"✔ 判定合格！折算后生鲜等效值 ({fresh_equiv:.4f} {unit}) ≤ 生鲜限量 ({fresh_limit} {unit})", bootstyle="success")
        else:
            self.dehydration_status_banner.config(text=f"❌ 判定超标！折算后生鲜等效值 ({fresh_equiv:.4f} {unit}) > 生鲜限量 ({fresh_limit} {unit})", bootstyle="danger")

        # 报告生成
        lines = []
        lines.append("=" * 66)
        lines.append(" 🧮 GB 2762 / GB 2761 浓缩物与干制品脱水折算法定检验报告")
        lines.append("=" * 66)
        lines.append(f"【加工食品名称】: {cat}")
        lines.append(f"【对应生鲜基准】: {fresh_name}")
        lines.append(f"【检测项目指标】: {item_name}")
        lines.append(f"【适用国家标准】: {std_name}")
        lines.append(f"【推荐检验方法】: {matched_item.get('method')}")
        lines.append("-" * 66)
        lines.append(f"【实测干品/浓缩值】: {meas_val} {unit}")
        lines.append(f"【折算参数设定】: {'脱水率 = ' + str(param_val) + '%' if mode == 'rate' else '浓缩倍数 = ' + str(param_val) + '倍'}")
        lines.append(f"【生鲜原料限量值】: ≤ {fresh_limit} {unit}")
        lines.append("\n【分步数学推导公式】:")
        lines.append(f"  {formula_str}")
        lines.append("-" * 66)
        if is_pass:
            lines.append(f"🏆 最终判定结论: 【✔ 合格】 (折算生鲜等效值 {fresh_equiv:.4f} {unit} 未超过生鲜限量 {fresh_limit} {unit})")
        else:
            lines.append(f"🚨 最终判定结论: 【❌ 不合格/超标】 (折算生鲜等效值 {fresh_equiv:.4f} {unit} 超过生鲜限量 {fresh_limit} {unit})")
        lines.append("=" * 66)

        self.dehydration_report_text.delete("1.0", tk.END)
        self.dehydration_report_text.insert(tk.END, "\n".join(lines))

    def copy_dehydration_report(self):
        txt = self.dehydration_report_text.get("1.0", tk.END).strip()
        if not txt:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(txt)
        messagebox.showinfo("复制成功", "已将脱水折算报告复制到剪贴板！")

    # ============================================================
    # Tab 9: GB 2763 农药残留全景矩阵速查与快检初筛
    # ============================================================
    def setup_pesticide_matrix_tab(self):
        """GB 2763 农药残留多靶标全景矩阵与快检初筛 Tab UI"""
        main_frame = tb.Frame(self.tab_pesticide, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        banner = tb.Frame(main_frame, bootstyle=SUCCESS)
        banner.pack(fill=tk.X, pady=(0, 8))
        tb.Label(banner, text="🌿 GB 2763 食品安全国家标准 农药最大残留限量(MRL) 全景矩阵与快检初筛助手", font=("Microsoft YaHei", 10, "bold"), bootstyle="inverse-success").pack(side=tk.LEFT, padx=10, pady=6)
        tb.Label(banner, text="支持一品多标多靶标聚合、禁用农药红线警示、酶抑制率/胶体金快检初筛与法定方法直链", font=("Microsoft YaHei", 8), bootstyle="inverse-success").pack(side=tk.RIGHT, padx=10, pady=6)

        paned = ttk.Panedwindow(main_frame, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # 左侧：品类选择与快检初筛输入
        left_box = tb.LabelFrame(paned, text="🔍 果蔬品种选择与快检初筛", padding=10)
        paned.add(left_box, weight=3)

        tb.Label(left_box, text="选择果蔬品类 (支持模糊速查):", font=("Microsoft YaHei", 9, "bold")).pack(anchor="w", pady=(0, 4))
        self.pesticide_filter_var = tk.StringVar()
        p_ent = tb.Entry(left_box, textvariable=self.pesticide_filter_var, font=("Microsoft YaHei", 9))
        p_ent.pack(fill=tk.X, pady=(0, 6))
        p_ent.bind("<KeyRelease>", lambda e: self._filter_pesticide_table())

        self.pesticide_cat_var = tk.StringVar()
        cats = list(PESTICIDE_MATRIX_DB.keys())
        self.pesticide_cat_combo = ttk.Combobox(left_box, textvariable=self.pesticide_cat_var, values=cats, state="readonly", font=("Microsoft YaHei", 9))
        self.pesticide_cat_combo.pack(fill=tk.X, pady=(0, 8))
        self.pesticide_cat_combo.bind("<<ComboboxSelected>>", lambda e: self._on_pesticide_cat_changed())

        # 快检初筛面板 (酶抑制率法 / 胶体金卡)
        quick_box = tb.LabelFrame(left_box, text="⚡ 农药残留快速检测初筛 (快检法)", padding=8)
        quick_box.pack(fill=tk.X, pady=(0, 8))

        tb.Label(quick_box, text="酶抑制率实测值 (%):", font=("Microsoft YaHei", 8)).pack(anchor="w")
        self.pesticide_inhibit_var = tk.StringVar(value="28.5")
        tb.Entry(quick_box, textvariable=self.pesticide_inhibit_var, font=("Microsoft YaHei", 9)).pack(fill=tk.X, pady=(0, 4))
        tb.Label(quick_box, text="注: 国家标准快检阈值为 50.0% (≥50% 判定为阳性疑似超标)", font=("Microsoft YaHei", 7), bootstyle=SECONDARY).pack(anchor="w", pady=(0, 6))

        tb.Button(quick_box, text="执行快检初筛判定", command=self.do_rapid_pesticide_judge, bootstyle="success-outline").pack(fill=tk.X)
        self.pesticide_rapid_result_label = tb.Label(quick_box, text="初筛结论: 待判定", font=("Microsoft YaHei", 9, "bold"), bootstyle=INFO)
        self.pesticide_rapid_result_label.pack(anchor="w", pady=(4, 0))

        # 右侧：全套农残限量标准矩阵表格
        right_box = tb.LabelFrame(paned, text="📊 GB 2763 农药最大残留限量(MRL) 全景矩阵", padding=8)
        paned.add(right_box, weight=7)

        cols = ("item", "type", "limit", "unit", "method", "risk")
        self.pesticide_tree = ttk.Treeview(right_box, columns=cols, show="headings", height=12)
        
        self.pesticide_tree.heading("item", text="农药名称 ↕", command=lambda: self.sort_tree_column(self.pesticide_tree, "item", False))
        self.pesticide_tree.heading("type", text="农药类别 ↕", command=lambda: self.sort_tree_column(self.pesticide_tree, "type", False))
        self.pesticide_tree.heading("limit", text="最大残留限量(MRL) ↕", command=lambda: self.sort_tree_column(self.pesticide_tree, "limit", False))
        self.pesticide_tree.heading("unit", text="单位", command=lambda: self.sort_tree_column(self.pesticide_tree, "unit", False))
        self.pesticide_tree.heading("method", text="推荐法定检验方法 ↕", command=lambda: self.sort_tree_column(self.pesticide_tree, "method", False))
        self.pesticide_tree.heading("risk", text="风险特征与国抽重点 ↕", command=lambda: self.sort_tree_column(self.pesticide_tree, "risk", False))

        self.pesticide_tree.column("item", width=120, minwidth=90)
        self.pesticide_tree.column("type", width=110, minwidth=80, anchor="center")
        self.pesticide_tree.column("limit", width=140, minwidth=100, anchor="center")
        self.pesticide_tree.column("unit", width=70, minwidth=50, anchor="center")
        self.pesticide_tree.column("method", width=120, minwidth=90)
        self.pesticide_tree.column("risk", width=220, minwidth=140)

        p_v_scroll = ttk.Scrollbar(right_box, orient=tk.VERTICAL, command=self.pesticide_tree.yview)
        self.pesticide_tree.configure(yscrollcommand=p_v_scroll.set)
        self.pesticide_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        p_v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 底部操作
        bot_bar3 = tb.Frame(main_frame)
        bot_bar3.pack(fill=tk.X, pady=(6, 0))
        tb.Button(bot_bar3, text="📋 复制当前品类全套农残清单", command=self.copy_pesticide_report, bootstyle="info-outline").pack(side=tk.LEFT)

        if cats:
            self.pesticide_cat_combo.set(cats[0])
            self._on_pesticide_cat_changed()

    def _on_pesticide_cat_changed(self):
        cat = self.pesticide_cat_var.get()
        items = PESTICIDE_MATRIX_DB.get(cat, [])
        for it in self.pesticide_tree.get_children():
            self.pesticide_tree.delete(it)

        for row in items:
            self.pesticide_tree.insert("", tk.END, values=(
                row.get("item", ""),
                row.get("type", ""),
                row.get("limit", ""),
                row.get("unit", "mg/kg"),
                row.get("method", "GB 23200.113"),
                row.get("risk", "")
            ))
        self.do_rapid_pesticide_judge()

    def _filter_pesticide_table(self):
        kw = self.pesticide_filter_var.get().strip().lower()
        all_cats = list(PESTICIDE_MATRIX_DB.keys())
        matched = []
        for c in all_cats:
            py = get_chinese_pinyin_initials(c)
            if not kw or kw in c.lower() or kw in py:
                matched.append(c)
        self.pesticide_cat_combo["values"] = matched
        if matched and self.pesticide_cat_var.get() not in matched:
            self.pesticide_cat_combo.set(matched[0])
            self._on_pesticide_cat_changed()

    def do_rapid_pesticide_judge(self):
        try:
            val = float(self.pesticide_inhibit_var.get().strip())
        except Exception:
            return
        if val >= 50.0:
            self.pesticide_rapid_result_label.config(
                text=f"初筛结论: 🚨 阳性 (抑制率 {val:.1f}% ≥ 50%) 疑似有机磷/氨基甲酸酯超标！建议送法定仪器复检！", 
                bootstyle="danger"
            )
        else:
            self.pesticide_rapid_result_label.config(
                text=f"初筛结论: ✔ 阴性 (抑制率 {val:.1f}% < 50%) 快检初筛合格", 
                bootstyle="success"
            )

    def copy_pesticide_report(self):
        cat = self.pesticide_cat_var.get()
        items = PESTICIDE_MATRIX_DB.get(cat, [])
        if not items:
            return
        lines = [f"=== GB 2763 农药残留全景限量矩阵: {cat} ==="]
        lines.append(f"{'农药名称':<12} {'类别':<10} {'限量(MRL)':<14} {'单位':<8} {'推荐检验方法':<14} {'风险特征'}")
        lines.append("-" * 75)
        for r in items:
            lines.append(f"{r.get('item'):<12} {r.get('type'):<10} {r.get('limit'):<14} {r.get('unit'):<8} {r.get('method'):<14} {r.get('risk')}")
        full_t = "\n".join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(full_t)
        messagebox.showinfo("复制成功", f"已将【{cat}】全套农残限量标准复制到剪贴板！")


def main():
    # 1. 全局一机一码离线授权校验与拦截
    try:
        from license_manager import check_and_require_license
        if not check_and_require_license(None, app_name="食品安全国家标准判定核查工具 V1.2", app_id="FSC"):
            return
    except Exception as e:
        print("License check warning:", e)

    # 2. 授权通过后进入系统主界面
    root = tb.Window(themename="cosmo")
    root.withdraw()
    
    splash = tb.Toplevel()
    splash.title("加载中")
    splash.geometry("400x150")
    splash.overrideredirect(True)
    
    sw = splash.winfo_screenwidth()
    sh = splash.winfo_screenheight()
    splash.geometry(f"+{sw//2-200}+{sh//2-75}")
    
    splash_frame = tb.Frame(splash)
    splash_frame.pack(fill=tk.BOTH, expand=True)
    
    tb.Label(splash_frame, text="食品安全国家标准\n判定核查工具 V1.2", 
            font=("Microsoft YaHei", 14, "bold")).pack(pady=(25,5))
    tb.Label(splash_frame, text="正在加载数据，请稍候...", 
            font=("Microsoft YaHei", 10)).pack()
    
    splash.update()
    
    app = FoodStandardChecker(root)
    
    splash.destroy()
    root.deiconify()
    root.lift()
    root.focus_force()
    root.mainloop()


if __name__ == "__main__":
    main()

