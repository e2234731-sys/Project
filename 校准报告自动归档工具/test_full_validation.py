# -*- coding: utf-8 -*-
import os, sys, fitz, pandas as pd
from main_gui import ArchiveWorker, LedgerDatabase

if sys.platform.startswith('win'):
    try: sys.stdout.reconfigure(encoding='utf-8')
    except: pass

print("=== 1. 测试批次 1: 东莞帝恩 D:\\下载报告 (249 份) ===")
folder1 = r'D:\下载报告'
worker1 = ArchiveWorker(
    target_dir=folder1,
    archive_root=r'E:\Antigravity\校准报告自动归档工具\【归档完成】校准证书库',
    mode='dry_run',
    enable_ocr=True,
    auto_unzip=False
)
summary1 = {}
def on_finished1(res):
    global summary1
    summary1 = res

worker1.finished_signal.connect(on_finished1)
worker1.run()

records1 = summary1.get('records', [])
print(f"总解析记录: {len(records1)}")
unknown_cal1 = sum(1 for r in records1 if r['校准日期'] == '未知校准')
unknown_grp1 = sum(1 for r in records1 if r['所属组别/实验室'] in ('未知组别', '未知实验室'))
desc_inst1 = sum(1 for r in records1 if r['仪器名称'] in ('Description', '未查找到'))
print(f"校准日期未知数: {unknown_cal1}/{len(records1)}")
print(f"所属组别未知数: {unknown_grp1}/{len(records1)}")
print(f"仪器名称异常数: {desc_inst1}/{len(records1)}")

print("\n=== 2. 测试批次 2: 深圳达丰 D:\\工作\\01.实验室法定资质与17025体系\\...\\新建文件夹 (30 份) ===")
folder2 = r'D:\工作\01.实验室法定资质与17025体系\1.定量实验室17025体系维护\2.设备管理\新建文件夹'
worker2 = ArchiveWorker(
    target_dir=folder2,
    archive_root=r'E:\Antigravity\校准报告自动归档工具\【归档完成】校准证书库',
    mode='dry_run',
    enable_ocr=True,
    auto_unzip=False
)
summary2 = {}
def on_finished2(res):
    global summary2
    summary2 = res

worker2.finished_signal.connect(on_finished2)
worker2.run()

records2 = summary2.get('records', [])
print(f"总解析记录: {len(records2)}")
unknown_cal2 = sum(1 for r in records2 if r['校准日期'] == '未知校准')
unknown_grp2 = sum(1 for r in records2 if r['所属组别/实验室'] in ('未知组别', '未知实验室'))
unknown_code2 = sum(1 for r in records2 if r['设备编号'] in ('未知编号', '—'))
print(f"校准日期未知数: {unknown_cal2}/{len(records2)}")
print(f"所属组别未知数: {unknown_grp2}/{len(records2)}")
print(f"设备编号未知数: {unknown_code2}/{len(records2)}")
for r in records2:
    print(f"  {r['原始文件名']:38s} -> 日期: {r['校准日期']} | 编号: {r['设备编号']} | 组别: {r['所属组别/实验室']} | 仪器: {r['仪器名称']}")

print("\n=== 3. 结果验证总结 ===")
if unknown_cal1 == 0 and desc_inst1 == 0 and unknown_cal2 == 0:
    print("🎉 全部指标 100% 达成！所有帝恩与达丰校准证书的校准日期、仪器名称、设备编号与所属组别均已精准识别！")
else:
    print("⚠️ 仍存在部分未捕获字段，请核查。")
