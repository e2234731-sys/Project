# -*- coding: utf-8 -*-
"""
食品安全国家标准判定核查工具 (纯净专业版) V1.2
专注文控、标准限量检索与合规判定，去除繁复多余模块
标准覆盖：GB 2760-2024, GB 2761-2017, GB 2762-2025, GB 2763-2026, GB 29921-2021, GB 31650-2019
Author: 陈耿忠 (国家注册审核员 / 实验室质控专家)
"""

import os
import sys
import json
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ttkbootstrap as tb
from ttkbootstrap.constants import *

# 引用同一一机一码授权核心
from license_manager import check_and_require_license, get_hardware_fingerprint

# ============================================================
# 纯净内置标准数据库 (GB 2760 / 2761 / 2762 / 2763 / 29921 / 31650)
# ============================================================
STANDARD_DATABASE = {
    "GB 2760-2024": {
        "name": "食品添加剂使用标准 (2024新版)",
        "items": [
            {"food_category": "发酵面制品(包子、馒头、花卷)", "test_item": "脱氢乙酸及其钠盐", "limit_value": "不得使用", "unit": "mg/kg", "notes": "2024版明确全面禁用，原2014版允许0.5g/kg"},
            {"food_category": "面包、糕点", "test_item": "脱氢乙酸及其钠盐", "limit_value": "不得使用", "unit": "mg/kg", "notes": "2024版明确全面禁用，严禁检出"},
            {"food_category": "熟肉制品(酱卤肉、熏烧烤肉)", "test_item": "脱氢乙酸及其钠盐", "limit_value": "不得使用", "unit": "mg/kg", "notes": "2024版明确全面禁用"},
            {"food_category": "预包装糕点", "test_item": "山梨酸及其钾盐", "limit_value": "1000", "unit": "mg/kg", "notes": "以山梨酸计，比例之和≤1"},
            {"food_category": "预包装糕点", "test_item": "丙酸及其钠盐、钙盐", "limit_value": "2500", "unit": "mg/kg", "notes": "以丙酸计"},
            {"food_category": "调味面制品(辣条)", "test_item": "脱氢乙酸及其钠盐", "limit_value": "不得使用", "unit": "mg/kg", "notes": "严禁添加"},
            {"food_category": "调味面制品(辣条)", "test_item": "山梨酸及其钾盐", "limit_value": "500", "unit": "mg/kg", "notes": "以山梨酸计"},
            {"food_category": "调味面制品(辣条)", "test_item": "甜蜜素(环己基氨基磺酸钠)", "limit_value": "1600", "unit": "mg/kg", "notes": "以环己基氨基磺酸计"},
            {"food_category": "冷冻饮品(雪糕、冰淇淋)", "test_item": "甜蜜素(环己基氨基磺酸钠)", "limit_value": "650", "unit": "mg/kg", "notes": "2024版限量由0.65g/kg严格规范"},
            {"food_category": "果冻", "test_item": "甜蜜素(环己基氨基磺酸钠)", "limit_value": "不得使用", "unit": "mg/kg", "notes": "2024版果冻中删除甜蜜素"},
            {"food_category": "酱油、酿造食醋", "test_item": "苯甲酸及其钠盐", "limit_value": "1000", "unit": "mg/kg", "notes": "以苯甲酸计"},
            {"food_category": "碳酸饮料", "test_item": "苯甲酸及其钠盐", "limit_value": "200", "unit": "mg/kg", "notes": "以苯甲酸计"},
            {"food_category": "熟肉制品", "test_item": "亚硝酸钠、亚硝酸钾", "limit_value": "30", "unit": "mg/kg", "notes": "以NaNO₂残留量计，严禁超标"},
            {"food_category": "果脯蜜饯类", "test_item": "二氧化硫残留量", "limit_value": "350", "unit": "mg/kg", "notes": "以SO₂计"},
            {"food_category": "生湿面条、饺子皮", "test_item": "二氧化硫残留量", "limit_value": "不得使用", "unit": "mg/kg", "notes": "严禁非法添加"},
            {"food_category": "生鲜肉、鲜蛋", "test_item": "任何防腐剂/色素", "limit_value": "不得使用", "unit": "mg/kg", "notes": "生鲜农产品严禁添加任何防腐保鲜色素"}
        ]
    },
    "GB 2761-2017": {
        "name": "食品中真菌毒素限量",
        "items": [
            {"food_category": "玉米、玉米制品", "test_item": "黄曲霉毒素B1", "limit_value": "20", "unit": "μg/kg", "notes": "谷物毒素监控重点"},
            {"food_category": "大米、糙米、稻谷", "test_item": "黄曲霉毒素B1", "limit_value": "10", "unit": "μg/kg", "notes": "口粮主食"},
            {"food_category": "花生、花生油、玉米油", "test_item": "黄曲霉毒素B1", "limit_value": "20", "unit": "μg/kg", "notes": "油脂与原料"},
            {"food_category": "其他植物油脂(大豆油/菜籽油)", "test_item": "黄曲霉毒素B1", "limit_value": "10", "unit": "μg/kg", "notes": ""},
            {"food_category": "发酵豆制品(腐乳、豆瓣酱)", "test_item": "黄曲霉毒素B1", "limit_value": "5.0", "unit": "μg/kg", "notes": ""},
            {"food_category": "婴幼儿配方食品", "test_item": "黄曲霉毒素B1", "limit_value": "0.5", "unit": "μg/kg", "notes": "以粉状产品计，极其严苛"},
            {"food_category": "生乳、纯牛奶、巴氏杀菌乳", "test_item": "黄曲霉毒素M1", "limit_value": "0.5", "unit": "μg/kg", "notes": "以生乳计"},
            {"food_category": "小麦、大麦、小麦粉", "test_item": "脱氧雪腐镰刀菌烯醇(DON呕吐毒素)", "limit_value": "1000", "unit": "μg/kg", "notes": "小麦赤霉病监控"},
            {"food_category": "玉米、玉米面", "test_item": "脱氧雪腐镰刀菌烯醇(DON呕吐毒素)", "limit_value": "1000", "unit": "μg/kg", "notes": ""},
            {"food_category": "苹果、山楂原汁及制品", "test_item": "展青霉素", "limit_value": "50", "unit": "μg/kg", "notes": "腐烂原料鉴别指标"},
            {"food_category": "谷物加工品、豆类、咖啡粉", "test_item": "赭曲霉毒素A", "limit_value": "5.0", "unit": "μg/kg", "notes": ""},
            {"food_category": "葡萄酒、果酒", "test_item": "赭曲霉毒素A", "limit_value": "2.0", "unit": "μg/kg", "notes": ""},
            {"food_category": "小麦、玉米及其制品", "test_item": "玉米赤霉烯酮(ZEN)", "limit_value": "60", "unit": "μg/kg", "notes": ""}
        ]
    },
    "GB 2762-2025": {
        "name": "食品中污染物限量 (2025新版)",
        "items": [
            {"food_category": "大米、稻谷、小麦粉", "test_item": "铅(Pb)", "limit_value": "0.2", "unit": "mg/kg", "notes": ""},
            {"food_category": "叶菜类蔬菜、芸薹类蔬菜", "test_item": "铅(Pb)", "limit_value": "0.3", "unit": "mg/kg", "notes": ""},
            {"food_category": "根茎类、瓜类、茄果类蔬菜", "test_item": "铅(Pb)", "limit_value": "0.1", "unit": "mg/kg", "notes": ""},
            {"food_category": "水产动物(鱼类、甲壳类)", "test_item": "铅(Pb)", "limit_value": "0.5", "unit": "mg/kg", "notes": ""},
            {"food_category": "双壳贝类水产(牡蛎、花蛤)", "test_item": "铅(Pb)", "limit_value": "1.5", "unit": "mg/kg", "notes": "贝类富集较高"},
            {"food_category": "生乳、纯牛奶", "test_item": "铅(Pb)", "limit_value": "0.05", "unit": "mg/kg", "notes": "以生乳计"},
            {"food_category": "大米、糙米", "test_item": "镉(Cd)", "limit_value": "0.2", "unit": "mg/kg", "notes": "大米镉安全线"},
            {"food_category": "小麦、小麦粉", "test_item": "镉(Cd)", "limit_value": "0.1", "unit": "mg/kg", "notes": ""},
            {"food_category": "叶菜类蔬菜", "test_item": "镉(Cd)", "limit_value": "0.2", "unit": "mg/kg", "notes": ""},
            {"food_category": "双壳贝类、腹足类", "test_item": "镉(Cd)", "limit_value": "2.0", "unit": "mg/kg", "notes": "去除内脏"},
            {"food_category": "甲壳类(虾、蟹)", "test_item": "镉(Cd)", "limit_value": "0.5", "unit": "mg/kg", "notes": ""},
            {"food_category": "水产动物及其制品(肉食性鱼除外)", "test_item": "总汞(Hg)", "limit_value": "0.5", "unit": "mg/kg", "notes": "以甲基汞或总汞计"},
            {"food_category": "肉食性鱼类(金枪鱼、旗鱼)", "test_item": "甲基汞", "limit_value": "1.0", "unit": "mg/kg", "notes": ""},
            {"food_category": "稻谷、大米", "test_item": "无机砷(As)", "limit_value": "0.2", "unit": "mg/kg", "notes": "以无机砷计"},
            {"food_category": "水产动物及其制品", "test_item": "无机砷(As)", "limit_value": "0.5", "unit": "mg/kg", "notes": "鱼类0.1，其他0.5"},
            {"food_category": "植物油脂", "test_item": "苯并[a]芘", "limit_value": "10", "unit": "μg/kg", "notes": "压榨/浸出油致癌物指标"}
        ]
    },
    "GB 2763-2026": {
        "name": "食品中农药最大残留限量 (GB 2763-2026/2021)",
        "items": [
            {"food_category": "普通白菜(小油菜、上海青)", "test_item": "啶虫脒", "limit_value": "1.0", "unit": "mg/kg", "notes": "GB 2763 推荐GB 23200.113/121"},
            {"food_category": "韭菜", "test_item": "腐霉利", "limit_value": "5.0", "unit": "mg/kg", "notes": "重点监控项"},
            {"food_category": "韭菜", "test_item": "毒死蜱", "limit_value": "0.02", "unit": "mg/kg", "notes": "蔬菜禁用/极低限量"},
            {"food_category": "豇豆(豆角)", "test_item": "克百威", "limit_value": "0.02", "unit": "mg/kg", "notes": "高毒农药禁用"},
            {"food_category": "豇豆(豆角)", "test_item": "噻虫胺", "limit_value": "0.01", "unit": "mg/kg", "notes": "豇豆重点治理项"},
            {"food_category": "豇豆(豆角)", "test_item": "倍硫磷", "limit_value": "0.05", "unit": "mg/kg", "notes": ""},
            {"food_category": "芹菜", "test_item": "毒死蜱", "limit_value": "0.05", "unit": "mg/kg", "notes": ""},
            {"food_category": "芹菜", "test_item": "甲拌磷", "limit_value": "0.01", "unit": "mg/kg", "notes": "高毒农药禁用"},
            {"food_category": "芹菜", "test_item": "晴菌唑", "limit_value": "0.05", "unit": "mg/kg", "notes": ""},
            {"food_category": "结球甘蓝(包菜)", "test_item": "甲胺磷", "limit_value": "0.05", "unit": "mg/kg", "notes": "禁用农药不得使用"},
            {"food_category": "苹果、梨", "test_item": "敌敌畏", "limit_value": "0.1", "unit": "mg/kg", "notes": ""},
            {"food_category": "柑、橘、橙", "test_item": "联苯菊酯", "limit_value": "0.05", "unit": "mg/kg", "notes": ""},
            {"food_category": "草莓", "test_item": "烯酰吗啉", "limit_value": "3.0", "unit": "mg/kg", "notes": ""},
            {"food_category": "茶叶", "test_item": "草甘膦", "limit_value": "1.0", "unit": "mg/kg", "notes": "除草剂指标"},
            {"food_category": "茶叶", "test_item": "三氯杀螨醇", "limit_value": "0.01", "unit": "mg/kg", "notes": "茶叶禁用"}
        ]
    },
    "GB 29921-2021": {
        "name": "预包装食品中致病菌限量",
        "items": [
            {"food_category": "肉制品(熟肉、酱卤肉)", "test_item": "沙门氏菌", "limit_value": "n=5,c=0,m=0/25g", "unit": "/25g", "notes": "5个样品均不得检出"},
            {"food_category": "肉制品(熟肉、酱卤肉)", "test_item": "单核细胞增生李斯特氏菌", "limit_value": "n=5,c=0,m=0/25g", "unit": "/25g", "notes": "冷藏即食重点"},
            {"food_category": "肉制品(熟肉)", "test_item": "金黄色葡萄球菌", "limit_value": "n=5,c=1,m=100,M=1000", "unit": "CFU/g", "notes": "三级采样方案"},
            {"food_category": "水产制品(熟制水产)", "test_item": "副溶血性弧菌", "limit_value": "n=5,c=1,m=100,M=1000", "unit": "MPN/g", "notes": "沿海水产致病菌"},
            {"food_category": "蛋制品(再制蛋、皮蛋)", "test_item": "沙门氏菌", "limit_value": "n=5,c=0,m=0/25g", "unit": "/25g", "notes": "严禁检出"},
            {"food_category": "粮食制品(糕点、面包)", "test_item": "沙门氏菌", "limit_value": "n=5,c=0,m=0/25g", "unit": "/25g", "notes": "严禁检出"},
            {"food_category": "粮食制品(糕点、面包)", "test_item": "金黄色葡萄球菌", "limit_value": "n=5,c=1,m=100,M=1000", "unit": "CFU/g", "notes": "三级采样方案"},
            {"food_category": "即食果蔬制品", "test_item": "大肠埃希氏菌O157:H7/NM", "limit_value": "n=5,c=0,m=0/25g", "unit": "/25g", "notes": "出血性大肠杆菌"}
        ]
    },
    "GB 31650-2019": {
        "name": "食品中兽药最大残留限量",
        "items": [
            {"food_category": "猪肉、牛肉、羊肉", "test_item": "克伦特罗(瘦肉精)", "limit_value": "不得检出", "unit": "μg/kg", "notes": "农业农村部公告第250号禁用"},
            {"food_category": "猪肉、牛肉、羊肉", "test_item": "莱克多巴胺(瘦肉精)", "limit_value": "不得检出", "unit": "μg/kg", "notes": "禁用兽药"},
            {"food_category": "猪肉、牛肉、羊肉", "test_item": "沙丁胺醇(瘦肉精)", "limit_value": "不得检出", "unit": "μg/kg", "notes": "禁用兽药"},
            {"food_category": "禽肉(鸡肉、鸭肉)", "test_item": "氯霉素", "limit_value": "不得检出", "unit": "μg/kg", "notes": "产蛋期/食用动物禁用"},
            {"food_category": "淡水鱼(草鱼、罗非鱼)", "test_item": "孔雀石绿(含隐色孔雀石绿)", "limit_value": "不得检出", "unit": "μg/kg", "notes": "水产养殖违禁药物"},
            {"food_category": "淡水鱼(黑鱼、黄颡鱼)", "test_item": "恩诺沙星", "limit_value": "100", "unit": "μg/kg", "notes": "以恩诺沙星与环丙沙星之和计"},
            {"food_category": "鸡蛋、鲜蛋", "test_item": "氟苯尼考", "limit_value": "不得检出", "unit": "μg/kg", "notes": "产蛋期禁用，蛋中不得检出"},
            {"food_category": "鸡蛋、鲜蛋", "test_item": "恩诺沙星", "limit_value": "不得检出", "unit": "μg/kg", "notes": "产蛋期禁用"},
            {"food_category": "生鲜牛乳", "test_item": "β-内酰胺类抗生素", "limit_value": "符合限量", "unit": "μg/kg", "notes": "抑菌圈法或仪器定量"}
        ]
    }
}



class TreeviewCellTooltip:
    """高级 Treeview 单元格智能悬浮气泡，支持完整展开、鼠标划词拉选与一键复制"""
    def __init__(self, tree):
        self.tree = tree
        self.tip_window = None
        self.after_id = None
        self.last_cell = None

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
            text = str(vals[col_idx]).strip()
            if not text or text == "/":
                return

            if len(text) < 12 and "method" not in col_name and "notes" not in col_name:
                return

            self.tip_window = tw = tk.Toplevel(self.tree)
            tw.wm_overrideredirect(True)
            tw.attributes("-topmost", True)
            
            outer_frame = tk.Frame(tw, bg="#0F172A", bd=1)
            outer_frame.pack(fill=tk.BOTH, expand=True)
            
            inner_frame = tk.Frame(outer_frame, bg="#FFFFFF", padx=8, pady=6)
            inner_frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

            hdr = tk.Frame(inner_frame, bg="#FFFFFF")
            hdr.pack(fill=tk.X, pady=(0, 4))
            tk.Label(hdr, text="🔍 完整检验方法 (鼠标可划选复制)", font=("Microsoft YaHei", 8, "bold"), fg="#0284C7", bg="#FFFFFF").pack(side=tk.LEFT)
            
            def copy_text():
                self.tree.clipboard_clear()
                self.tree.clipboard_append(text)
                self.tree.update()
                btn_copy.config(text="✔ 已复制")

            btn_copy = tk.Button(hdr, text="📋 复制全部", font=("Microsoft YaHei", 8), command=copy_text, relief=tk.FLAT, bg="#E2E8F0", cursor="hand2")
            btn_copy.pack(side=tk.RIGHT)

            calc_h = min(6, max(2, len(text) // 25 + 1))
            calc_w = min(55, max(28, len(text) + 2))
            txt_box = tk.Text(inner_frame, wrap=tk.WORD, height=calc_h, width=calc_w,
                              font=("Consolas", 9), bg="#F8FAFC", fg="#0F172A", bd=1, relief=tk.SOLID, padx=4, pady=4)
            txt_box.insert("1.0", text)
            txt_box.pack(fill=tk.BOTH, expand=True)
            txt_box.bind("<Key>", lambda e: "break" if (e.state != 4 and e.keysym not in ('c', 'C', 'a', 'A')) else None)

            tw.bind("<Leave>", lambda e: self.on_leave())

            screen_w = self.tree.winfo_screenwidth()
            screen_h = self.tree.winfo_screenheight()
            pos_x = min(x + 10, screen_w - 450)
            pos_y = min(y + 12, screen_h - 180)
            tw.geometry(f"+{pos_x}+{pos_y}")
        except Exception:
            pass


class PureFoodStandardCheckerApp:
    """食品安全标准判定核查工具 (纯净专业版)"""
    def __init__(self, root):
        self.root = root
        total_items = sum(len(v["items"]) for v in STANDARD_DATABASE.values())
        self.root.title(f"食品安全标准判定核查工具 V1.2 - 共 {total_items} 条权威限量 (Powered by 陈耿忠)")
        self.root.geometry("1100x720")
        self.root.minsize(960, 620)
        self.judge_history = []
        self.setup_ui()

    def setup_ui(self):
        main_container = tb.Frame(self.root, padding=12)
        main_container.pack(fill=tk.BOTH, expand=True)

        # 顶部 Banner 栏
        banner = tb.Frame(main_container, bootstyle=PRIMARY)
        banner.pack(fill=tk.X, pady=(0, 10))
        tb.Label(
            banner, 
            text="🛡️ 食品安全国家标准判定核查工作台", 
            font=("Microsoft YaHei", 12, "bold"), 
            bootstyle="inverse-primary"
        ).pack(side=tk.LEFT, padx=14, pady=8)
        
        tb.Label(
            banner, 
            text="⚡ Powered by 陈耿忠 (国家注册审核员)  |  GB 2760 / 2761 / 2762 / 2763 / 29921 / 31650 现行数据库", 
            font=("Microsoft YaHei", 9), 
            bootstyle="inverse-primary"
        ).pack(side=tk.RIGHT, padx=14, pady=8)

        # 检索与筛选过滤卡片
        filter_card = tb.LabelFrame(main_container, text="🔍 标准与限量智能检索", padding=10)
        filter_card.pack(fill=tk.X, pady=(0, 10))

        row1 = tb.Frame(filter_card)
        row1.pack(fill=tk.X, pady=2)

        tb.Label(row1, text="适用标准:", font=("Microsoft YaHei", 9, "bold"), width=9).pack(side=tk.LEFT)
        self.std_choice_var = tk.StringVar(value="全部标准")
        std_list = ["全部标准"] + [f"{k} ({v['name']})" for k, v in STANDARD_DATABASE.items()]
        self.std_cb = ttk.Combobox(row1, textvariable=self.std_choice_var, values=std_list, state="readonly", width=38, font=("Microsoft YaHei", 9))
        self.std_cb.pack(side=tk.LEFT, padx=(0, 14))
        self.std_cb.bind("<<ComboboxSelected>>", lambda e: self.do_search())

        tb.Label(row1, text="食品名称/类别:", font=("Microsoft YaHei", 9, "bold"), width=12).pack(side=tk.LEFT)
        self.food_search_var = tk.StringVar()
        self.food_entry = tb.Entry(row1, textvariable=self.food_search_var, width=18, font=("Microsoft YaHei", 9))
        self.food_entry.pack(side=tk.LEFT, padx=(0, 14))
        self.food_entry.bind("<Return>", lambda e: self.do_search())

        tb.Label(row1, text="检测项目/添加剂/农兽残:", font=("Microsoft YaHei", 9, "bold"), width=20).pack(side=tk.LEFT)
        self.item_search_var = tk.StringVar()
        self.item_entry = tb.Entry(row1, textvariable=self.item_search_var, width=18, font=("Microsoft YaHei", 9))
        self.item_entry.pack(side=tk.LEFT, padx=(0, 10))
        self.item_entry.bind("<Return>", lambda e: self.do_search())

        tb.Button(row1, text="🔍 立即查询", command=self.do_search, bootstyle=INFO, width=10).pack(side=tk.LEFT, padx=4)
        tb.Button(row1, text="🔄 重置", command=self.reset_search, bootstyle="secondary-outline", width=8).pack(side=tk.LEFT, padx=4)

        # 中部主体：左右切分 (左侧结果表格，右侧实时判定核查操作台)
        paned = ttk.Panedwindow(main_container, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # 左侧：标准限量数据表格
        table_frame = tb.LabelFrame(paned, text="📋 标准限量数据库检索结果", padding=8)
        paned.add(table_frame, weight=7)

        columns = ("std_code", "food_cat", "test_item", "limit_val", "unit", "method", "notes")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        
        self.tree.heading("std_code", text="标准代号")
        self.tree.heading("food_cat", text="食品类别 / 原料")
        self.tree.heading("test_item", text="检测项目 / 物质名称")
        self.tree.heading("limit_val", text="标准限量值")
        self.tree.heading("unit", text="单位")
        self.tree.heading("method", text="要求检验方法")
        self.tree.heading("notes", text="法规备注 / 合规判定指引")

        self.tree.column("std_code", width=105, anchor="center")
        self.tree.column("food_cat", width=160, anchor="w")
        self.tree.column("test_item", width=150, anchor="w")
        self.tree.column("limit_val", width=110, anchor="center")
        self.tree.column("unit", width=60, anchor="center")
        self.tree.column("method", width=170, anchor="w")
        self.tree.column("notes", width=200, anchor="w")

        v_scroll = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=v_scroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        TreeviewCellTooltip(self.tree)

        def on_clean_tree_ctrl_c(event):
            sel = self.tree.selection()
            if not sel:
                return
            vals = self.tree.item(sel[0], 'values')
            m_val = vals[5] if len(vals) > 5 else "\t".join(str(v) for v in vals)
            self.root.clipboard_clear()
            self.root.clipboard_append(m_val)
            self.root.update()

        self.tree.bind("<Control-c>", on_clean_tree_ctrl_c)
        self.tree.bind("<Control-C>", on_clean_tree_ctrl_c)

        # 右侧：合规判定执行面板
        judge_frame = tb.LabelFrame(paned, text="⚖️ 合规判定执行台", padding=12)
        paned.add(judge_frame, weight=4)

        tb.Label(judge_frame, text="当前选中标准:", font=("Microsoft YaHei", 9)).pack(anchor="w", pady=(0, 2))
        self.lbl_cur_std = tb.Label(judge_frame, text="--", font=("Microsoft YaHei", 10, "bold"), bootstyle=PRIMARY)
        self.lbl_cur_std.pack(anchor="w", pady=(0, 6))

        tb.Label(judge_frame, text="食品类别:", font=("Microsoft YaHei", 9)).pack(anchor="w", pady=(0, 2))
        self.lbl_cur_food = tb.Label(judge_frame, text="--", font=("Microsoft YaHei", 9, "bold"))
        self.lbl_cur_food.pack(anchor="w", pady=(0, 6))

        tb.Label(judge_frame, text="检测项目:", font=("Microsoft YaHei", 9)).pack(anchor="w", pady=(0, 2))
        self.lbl_cur_item = tb.Label(judge_frame, text="--", font=("Microsoft YaHei", 9, "bold"))
        self.lbl_cur_item.pack(anchor="w", pady=(0, 6))

        tb.Label(judge_frame, text="标准限量要求:", font=("Microsoft YaHei", 9)).pack(anchor="w", pady=(0, 2))
        self.lbl_cur_limit = tb.Label(judge_frame, text="--", font=("Microsoft YaHei", 11, "bold"), bootstyle=DANGER)
        self.lbl_cur_limit.pack(anchor="w", pady=(0, 6))

        tb.Label(judge_frame, text="要求检验方法:", font=("Microsoft YaHei", 9)).pack(anchor="w", pady=(0, 2))
        self.lbl_cur_method = tb.Label(judge_frame, text="--", font=("Microsoft YaHei", 9, "bold"), bootstyle=INFO, wraplength=260, justify=tk.LEFT)
        self.lbl_cur_method.pack(anchor="w", pady=(0, 10))

        # 实测值输入
        tb.Label(judge_frame, text="✍️ 请输入实验室实测检测值:", font=("Microsoft YaHei", 9, "bold")).pack(anchor="w", pady=(4, 2))
        row_input = tb.Frame(judge_frame)
        row_input.pack(fill=tk.X, pady=(0, 10))
        self.test_val_var = tk.StringVar()
        self.val_entry = tb.Entry(row_input, textvariable=self.test_val_var, font=("Consolas", 11, "bold"))
        self.val_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        self.lbl_unit_hint = tb.Label(row_input, text="--", font=("Microsoft YaHei", 9))
        self.lbl_unit_hint.pack(side=tk.RIGHT)
        self.val_entry.bind("<Return>", lambda e: self.do_judge())

        # 快捷按钮
        row_quick = tb.Frame(judge_frame)
        row_quick.pack(fill=tk.X, pady=(0, 10))
        tb.Button(row_quick, text="填入 ND (未检出)", command=lambda: self.set_test_val("ND"), bootstyle="secondary-outline", width=14).pack(side=tk.LEFT, padx=(0, 4))
        tb.Button(row_quick, text="填入 0", command=lambda: self.set_test_val("0"), bootstyle="secondary-outline", width=8).pack(side=tk.LEFT)

        # 判定按钮
        tb.Button(judge_frame, text="⚡ 立即执行合规判定", command=self.do_judge, bootstyle=SUCCESS).pack(fill=tk.X, pady=(0, 12))

        # 判定结果卡片
        self.res_card = tb.Frame(judge_frame, padding=10, bootstyle="secondary")
        self.res_card.pack(fill=tk.X, pady=(0, 10))
        self.lbl_res_text = tb.Label(self.res_card, text="待判定", font=("Microsoft YaHei", 14, "bold"), bootstyle="inverse-secondary", anchor="center")
        self.lbl_res_text.pack(fill=tk.X)
        self.lbl_res_detail = tb.Label(self.res_card, text="请先选择表格数据并输入检测值", font=("Microsoft YaHei", 8), bootstyle="inverse-secondary", anchor="center")
        self.lbl_res_detail.pack(fill=tk.X, pady=(4, 0))

        # 下方：判定流水日志与导出栏
        bottom_box = tb.LabelFrame(main_container, text="📜 判定记录与导出", padding=8)
        bottom_box.pack(fill=tk.X, pady=(10, 0))

        self.history_text = tk.Text(bottom_box, height=4, font=("Consolas", 9), wrap=tk.WORD, bg="#f8fafc")
        h_scroll = ttk.Scrollbar(bottom_box, orient=tk.VERTICAL, command=self.history_text.yview)
        self.history_text.configure(yscrollcommand=h_scroll.set)
        self.history_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        h_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 底部状态与按钮
        btn_box = tb.Frame(main_container)
        btn_box.pack(fill=tk.X, pady=(6, 0))
        tb.Label(btn_box, text="🛡️ 食品安全国家标准判定核查纯净工作台 V1.2 (Powered by 陈耿忠)", font=("Microsoft YaHei", 8), bootstyle=PRIMARY).pack(side=tk.LEFT)
        
        tb.Button(btn_box, text="💾 导出判定记录 (TXT/CSV)", command=self.export_history, bootstyle=INFO).pack(side=tk.RIGHT, padx=4)
        tb.Button(btn_box, text="🗑️ 清空记录", command=self.clear_history, bootstyle="secondary-outline").pack(side=tk.RIGHT, padx=4)

        # 初始填充表格
        self.do_search()

    def set_test_val(self, val):
        self.test_val_var.set(val)

    def do_search(self):
        """根据标准、食品类别、检测项目执行智能搜索 (精准匹配优先置顶)"""
        self.tree.delete(*self.tree.get_children())
        
        std_filter = self.std_choice_var.get()
        food_kw = self.food_search_var.get().strip().lower()
        item_kw = self.item_search_var.get().strip().lower()

        matched_records = []
        for std_code, std_data in STANDARD_DATABASE.items():
            if std_filter != "全部标准" and std_code not in std_filter:
                continue
            for r in std_data["items"]:
                f_cat = r.get("food_category", "")
                t_item = r.get("test_item", "")
                l_val = r.get("limit_value", "")
                unit = r.get("unit", "")
                notes = r.get("notes", "")
                m_str = r.get("method", "")

                if food_kw and food_kw not in f_cat.lower():
                    continue
                if item_kw and item_kw not in t_item.lower():
                    continue
                matched_records.append((std_code, f_cat, t_item, l_val, unit, m_str, notes, r))

        # 精准匹配排序
        query_key = item_kw or food_kw
        if query_key:
            matched_records.sort(key=lambda x: calc_match_rank(x[7], query_key))

        match_count = 0
        for std_code, f_cat, t_item, l_val, unit, m_str, notes, r in matched_records:
            tag_name = std_code.replace(" ", "_").replace("-", "_")
            self.tree.insert("", "end", values=(std_code, f_cat, t_item, l_val, unit, m_str, notes), tags=(tag_name,))
            match_count += 1

        self.status_label.config(text=f"🔍 检索到 {match_count} 条相关限量指标 (精准匹配优先置顶)")

    def reset_search(self):
        self.std_choice_var.set("全部标准")
        self.food_search_var.set("")
        self.item_search_var.set("")
        self.do_search()

    def on_tree_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        if vals:
            self.lbl_cur_std.config(text=vals[0])
            self.lbl_cur_food.config(text=vals[1])
            self.lbl_cur_item.config(text=vals[2])
            unit = vals[4] if vals[4] else ""
            self.lbl_cur_limit.config(text=f"{vals[3]} {unit}")
            self.lbl_unit_hint.config(text=unit)
            self.current_selected_row = vals

    def do_judge(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("提示", "请先在左侧表格中选中一条标准限量！")
            return

        raw_test = self.test_val_var.get().strip()
        if not raw_test:
            messagebox.showwarning("提示", "请输入实验室实测检测值！")
            return

        vals = self.tree.item(sel[0], "values")
        std_code, food_cat, test_item, limit_val, unit, notes = vals
        
        is_pass = False
        reason = ""

        # 判定逻辑
        if "不得使用" in limit_val or "不得检出" in limit_val or "严禁" in limit_val:
            if raw_test.upper() in ["0", "ND", "未检出", "不得检出", "0.0"]:
                is_pass = True
                reason = "未检出，符合标准要求不得使用的规定"
            else:
                try:
                    num = float(raw_test)
                    if num <= 0:
                        is_pass = True
                        reason = "实测值未检出(0)，符合要求"
                    else:
                        is_pass = False
                        reason = f"实测值 {raw_test} {unit}，违反【不得使用/不得检出】规定"
                except Exception:
                    is_pass = False
                    reason = f"检出 {raw_test}，违反【不得使用/不得检出】规定"
        elif "n=" in limit_val:
            # 致病菌三级采样规则简易判断
            if raw_test.upper() in ["0", "ND", "未检出", "0/25G"]:
                is_pass = True
                reason = "致病菌未检出，判定合格"
            else:
                is_pass = False
                reason = f"检测值 {raw_test}，请结合三级采样 n,c,m,M 严格核验"
        else:
            # 数值型比对
            try:
                limit_num = float(limit_val)
                test_num = float(raw_test)
                if test_num <= limit_num:
                    is_pass = True
                    reason = f"实测值 {test_num} {unit} ≤ 限量标准 {limit_num} {unit}"
                else:
                    is_pass = False
                    reason = f"实测值 {test_num} {unit} 超过限量标准 {limit_num} {unit} (超标)"
            except Exception:
                # 兼容文本型
                if raw_test == limit_val or raw_test.upper() in ["0", "ND"]:
                    is_pass = True
                    reason = f"实测符合要求"
                else:
                    is_pass = False
                    reason = f"实测值 {raw_test} 与标准 {limit_val} 不符"

        # 更新 UI
        if is_pass:
            self.res_card.configure(bootstyle=SUCCESS)
            self.lbl_res_text.configure(text="✅ 合格 (Pass)", bootstyle="inverse-success")
            self.lbl_res_detail.configure(text=reason, bootstyle="inverse-success")
            result_str = "【合格】"
        else:
            self.res_card.configure(bootstyle=DANGER)
            self.lbl_res_text.configure(text="❌ 不合格 (Fail)", bootstyle="inverse-danger")
            self.lbl_res_detail.configure(text=reason, bootstyle="inverse-danger")
            result_str = "【不合格】"

        # 记录日志
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_line = f"[{ts}] {result_str} 标准:{std_code} | 食品:{food_cat} | 项目:{test_item} | 实测:{raw_test} {unit} | 限量:{limit_val} {unit} ({reason})\n"
        self.judge_history.append(log_line)
        self.history_text.insert(tk.END, log_line)
        self.history_text.see(tk.END)

    def export_history(self):
        if not self.judge_history:
            messagebox.showinfo("提示", "当前暂无判定流水记录！")
            return
        
        fpath = filedialog.asksaveasfilename(
            title="导出判定记录",
            defaultextension=".txt",
            filetypes=[("文本文件 (*.txt)", "*.txt"), ("CSV 文件 (*.csv)", "*.csv")],
            initialfile=f"食品安全标准判定记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        )
        if fpath:
            with open(fpath, "w", encoding="utf-8") as f:
                f.write("=" * 75 + "\n")
                f.write("食品安全国家标准合规判定流水记录\n")
                f.write(f"生成系统：食品安全标准判定核查工具 (纯净版) V1.0 (Powered by 陈耿忠)\n")
                f.write(f"导出时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("=" * 75 + "\n\n")
                for line in self.judge_history:
                    f.write(line)
            messagebox.showinfo("成功", f"判定记录已成功导出至：\n{fpath}")

    def clear_history(self):
        if self.judge_history:
            if messagebox.askyesno("确认", "确定要清空所有判定流水记录吗？"):
                self.judge_history.clear()
                self.history_text.delete("1.0", tk.END)
                self.res_card.configure(bootstyle="secondary")
                self.lbl_res_text.configure(text="待判定", bootstyle="inverse-secondary")
                self.lbl_res_detail.configure(text="请先选择表格数据并输入检测值", bootstyle="inverse-secondary")


def main():
    # 1. 纯净一机一码离线授权验证
    try:
        if not check_and_require_license(None, app_name="食品安全标准判定核查工具 (纯净专业版) V1.2", app_id="FSC"):
            return
    except Exception as e:
        print("License check warning:", e)

    # 2. 授权通过后进入纯净版主窗口
    root = tb.Window(themename="cosmo")
    app = PureFoodStandardCheckerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
def calc_match_rank(item, search_text):
    """计算搜索条目精准匹配优先级 (越小越排在最前面)"""
    if not search_text:
        return (99, 999, 999)
    s_clean = search_text.strip().lower()
    t_item = str(item.get('test_item', '')).strip().lower()
    f_cat = str(item.get('food_category', '')).strip().lower()
    m_str = str(item.get('method', '')).strip().lower()
    if s_clean == t_item:
        return (0, len(t_item), len(f_cat))
    if s_clean == f_cat:
        return (1, len(t_item), len(f_cat))
    if s_clean in t_item.split() or s_clean in f_cat.split():
        return (2, len(t_item), len(f_cat))
    if t_item.startswith(s_clean):
        return (3, len(t_item), len(f_cat))
    if f_cat.startswith(s_clean):
        return (4, len(t_item), len(f_cat))
    if s_clean in t_item:
        return (5, len(t_item), len(f_cat))
    if s_clean in f_cat:
        return (6, len(t_item), len(f_cat))
    if s_clean in m_str:
        return (7, len(t_item), len(f_cat))
    return (99, len(t_item), len(f_cat))


