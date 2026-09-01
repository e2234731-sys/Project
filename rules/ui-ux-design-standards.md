---
name: ui-ux-design-standards
description: Global standards and rules for professional laboratory desktop UI/UX design, layout responsiveness, and user interaction.
trigger: always_on
---

# 实验室专业桌面应用 UI/UX 设计与交互规范

## 核心设计与交互原则 (Core Principles)

### 1. 全局自适应滚动与防截断规范 (Adaptive Scroll & Anti-Truncation)
- **强制滚动容器**：所有主工作台页面、子 Tab、卡片视图均必须使用 `QScrollArea` 封装，并强制配置 `setWidgetResizable(True)`、`setFrameShape(QFrame.NoFrame)` 及 `ScrollBarAsNeeded`，确保在笔记本小屏、高 DPI 缩放（125%/150%）及窗口缩放时平滑滚动，严禁出现内容溢出或被遮挡且无法拖动的现象。
- **文本自动折行与气泡悬浮 (WordWrap & Tooltips)**：所有 `QLabel` 描述性文本、标题、说明、警告及表格单元格必须设置 `setWordWrap(True)` 或绑定鼠标悬浮提示（`setToolTip`），保障长文本完整可读。

### 2. 表格全交互拖拽与列宽持久化 (Table Interactivity & Persistence)
- **表头交互模式**：数据表格的列宽调整模式必须统一采用 `QHeaderView.Interactive`，严禁在中间列设置 `Stretch` 导致相邻列分割线锁死。
- **列宽记忆功能**：表格各列宽度在用户手动拖拽调整后，应自动持久化存储至本地配置文件（如 `ui_settings.json`），并在下次启动时无缝恢复，提供个性化的排版体验。
- **右键上下文菜单**：核心数据表格应提供鼠标右键快捷菜单（包含文件定位、编辑、状态流转、删除等快捷动作）。

### 3. 商务严谨与现代化视觉规范 (Corporate Business Styling)
- **专业色系搭配**：
  - 主导航/基底：深邃午夜蓝（Midnight Navy `#0F172A`）与深灰蓝（`#1E293B`），体现质量体系的严肃严谨。
  - 核心激活/聚焦：科技蓝（`#0284C7` / `#0EA5E9`）。
  - 状态指示：翡翠绿（`#10B981` 现行有效）、琥珀黄（`#F59E0B` 待查新/修订中）、朱红（`#EF4444` 作废销毁）。
- **微边框圆角卡片**：内容区域使用白色底色搭配 `#E2E8F0` 浅色细边框与 `border-radius: 10px~12px`，避免纯扁平化导致的层级混乱。

### 4. 一键源文件定位与双击交互 (One-Click File Locator)
- **资源管理器联动**：提供 `FileLocator` 引擎，支持选中数据行后在 Windows Explorer 中自动高亮选中物理文件（`explorer /select,"<filepath>"`），或者精准回溯至其所属的体系类别文件夹。
