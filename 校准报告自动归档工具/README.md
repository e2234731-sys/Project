# 📄 FQT 实验室校准证书智能归档与量值溯源工作台 V2.0
### Smart Calibration Certificate Parser, Auto-Archiver & Metrology Traceability Suite (ISO/IEC 17025)

---

## 📖 项目概述 (Overview)

面向第三方检测机构（如 FQT / 深圳凯吉星）定量实验室（17025 体系）与快检质量网络驻点打造的专业级**校准证书智能识别、台账核对与多层级自动归档工作台**。

### 核心亮点
1. **双轨工作台架构**：
   - 🏢 **17025 定量实验室工作台**：自动联动《2026年年度计划汇总.xlsx》（量值溯源总表），覆盖理化、微生物、液相、气相、元素、抽样、综合组等 750+ 台定量仪器。
   - 🧪 **快检驻点网络工作台**：自动联动《2026年各实验室仪器设备校准清单.xlsx》，覆盖 30+ 驻点项目组、街道与快检室等 300+ 台快检设备。
2. **计划台账闭环审计**：
   - 自动对比年度校准计划 vs. 实际扫描证书，一键生成「计划内待收回/未见 2026 证书清单」，实现计量溯源闭环管理。
3. **全机构高精度智能解析**：
   - 深度识别东莞帝恩 (DN)、广州中广测 (NEM)、深圳达丰 (DAF)、深圳计量院 (SMQ)、中检深圳 (CCIC)；
   - 自动拆解复合出厂编号（如 2013C004-31(H0406)）、移液枪复合编码（如 18F47863(YYQ132)）与签发日期。
4. **极速双引擎与单文件打包**：
   - PyMuPDF 文本层优先毫秒级解析 + RapidOCR 离线视觉自愈兜底；
   - 支持一键打包为独立免安装单文件 .exe。

---

## 🛠️ 运行环境与安装 (Installation)

`ash
# 安装依赖
pip install -r requirements.txt

# 启动 GUI 工作台
python main_gui.py

# 一键打包单文件可执行程序
build_exe.bat
`

---

## 📁 目录结构

`
校准报告自动归档工具/
├── main_gui.py              # 主工作台源码 (PyQt5 + Fluent UI)
├── build_exe.bat            # PyInstaller 一键打包脚本
├── config_presets.json      # 常用路径快捷预设配置
├── requirements.txt         # 项目依赖列表
└── README.md                # 项目技术说明文档
`
