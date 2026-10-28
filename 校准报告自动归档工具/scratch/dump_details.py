# -*- coding: utf-8 -*-
import os, sys, fitz

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

dir_path = r"D:\wechatfile\xwechat_files\e2234731_f65d\msg\file\2026-10\26AX04192深圳凯吉星农产品检测认证有限公司"

def dump_file(filename):
    print(f"\n=======================================================")
    print(f"FILE: {filename}")
    path = os.path.join(dir_path, filename)
    if not os.path.exists(path):
        print("File not found")
        return
    doc = fitz.open(path)
    for pno in range(min(3, len(doc))):
        print(f"--- Page {pno+1} ---")
        text = doc[pno].get_text()
        print(text)

files_to_check = [
    "WJ118_Description_20260917_抽样组.pdf",
    "未知编号_Description_20260917_未知组别.pdf",
    "未知编号_Description_20260917_未知组别_1.pdf",
    "管 理 编 号  ：YYQ211_未查找到_20260918_未知组别.pdf",
    "H0397_二级生物安全柜_未知校准_微生物组.pdf",
    "未知编号_未查找到_20260920_未知组别_1.pdf",
    "未知编号_未查找到_20260920_未知组别_3.pdf",
    "未知编号_未查找到_20260921_未知组别.pdf",
    "未知编号_未查找到_20261008_未知组别.pdf",
]

for f in files_to_check:
    dump_file(f)
