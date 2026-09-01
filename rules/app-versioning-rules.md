---
name: app-versioning-rules
description: Standards for app versioning, iteration numbers, and naming constraints without sensational suffixes.
trigger: always_on
---

# APP 迭代版本号规范与命名约束

## 核心指令 (Critical Instruction)
1. **规范版本号递增**：
   - 每次开发、迭代更新、修复 Bug 或打包发布 APP 时，必须严格规范标注版本号。
   - 初始发布版本从 **V1.0** 开始；日常迭代、功能优化或 Bug 修复按 **V1.1, V1.2, V1.3...** 递增；重大重构升级按 **V2.0** 递增。
2. **严禁使用夸张/浮夸后缀**：
   - 应用名称、窗口标题、文件命名及文档中，**严禁使用任何夸张、花哨或过度营销性质的后缀**（如严禁使用“最终版”、“终极至尊版”、“Pro Max Plus”、“豪华重构版”、“超级版”等）。
   - 统一使用简明、商务、专业的命名格式：`[应用名称] Vx.x`（例如：`FQT 17025 体系文控管理工作台 V1.0`）。
3. **全局版本一致性**：
   - 应用代码配置（如 `config.py`、`__version__`）、主窗口标题栏、界面底部状态栏、说明文档及可执行文件命名均须保持版本号统一同步。
