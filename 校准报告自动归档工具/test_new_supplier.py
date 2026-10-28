# -*- coding: utf-8 -*-
import os, sys
import fitz # PyMuPDF
import pandas as pd

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

dir_path = r"D:\wechatfile\xwechat_files\e2234731_f65d\msg\file\2026-10\26AX04192深圳凯吉星农产品检测认证有限公司"
excel_path = os.path.join(dir_path, "本次扫描_校准信息识别与台账核对表_V2.5.xlsx")

print("=== 读取本次扫描 Excel ===")
if os.path.exists(excel_path):
    df = pd.read_excel(excel_path)
    for col in df.columns:
        print("Column:", col)
    print("\nData summary:")
    # print all rows with key columns
    cols_to_show = [c for c in df.columns if any(k in str(c) for k in ['原始', '编号', '设备', '仪器', '校准', '状态', '机构', '结论', '目标'])]
    print(df[cols_to_show].to_string())

print("\n=== 扫描目录中的 PDF 提取文本前 2 页与机构特征 ===")
pdf_files = [f for f in os.listdir(dir_path) if f.lower().endswith('.pdf')]
print(f"共有 {len(pdf_files)} 个 PDF 文件")

for f in pdf_files:
    file_path = os.path.join(dir_path, f)
    try:
        doc = fitz.open(file_path)
        first_pages_text = ""
        for i in range(min(2, len(doc))):
            first_pages_text += f"\n--- Page {i+1} ---\n" + doc[i].get_text()
        
        # Check institution mentions
        inst = "未知"
        if "广州力赛" in first_pages_text or "力赛计量" in first_pages_text:
            inst = "广州力赛"
        elif "深圳中恒" in first_pages_text or "中恒检测" in first_pages_text:
            inst = "深圳中恒"
        elif "达丰" in first_pages_text:
            inst = "达丰"
        elif "华南国家计量测试中心" in first_pages_text or "广东省计量科学研究院" in first_pages_text or "SCM" in first_pages_text:
            inst = "广计院/华南"
        elif "深圳市计量质量检测研究院" in first_pages_text or "SMQ" in first_pages_text:
            inst = "SMQ"
        elif "广州广电" in first_pages_text or "GRGT" in first_pages_text:
            inst = "广电计量"

        print(f"\n==========================================")
        print(f"文件: {f} -> 识别机构: {inst}")
        # print first 500 chars of page 1
        lines = [line.strip() for line in first_pages_text.splitlines() if line.strip()]
        print("前 20 行文本:")
        for line in lines[:20]:
            print("  ", line)
    except Exception as e:
        print(f"读取文件 {f} 失败: {e}")
