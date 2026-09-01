# GitHub Release 自动发布与资产交付规范 (GitHub Release Automation Rules)

## 核心指令 (Critical Instruction)
在完成任何桌面应用（EXE）、Web 应用、脚本工具的**版本迭代与打包编译**后，Antigravity 必须自动执行 Git 提交、打 Tag，并通过 GitHub API 自动将构建产物上传至 GitHub Release，形成完整的版本归档闭环。

---

## 1. 触发条件与自动化流程 (Trigger & Automation Flow)
1. **全自动静默触发**：
   - 触发时机：当 APP 代码完成修改、版本号递增（如从 V1.0 递增至 V1.1），且成功完成打包编译（例如 PyInstaller 生成 .exe、前端生成 dist 压缩包）后，**自动静默触发 Release 发布工作流**。
2. **自动化标准流转步骤**：
   - **Step 1: 本地提交与 Tag 标记**：执行 `git add .`、`git commit -m "release: [应用名称] Vx.x 迭代发布"`，并创建标准版本 Tag（`git tag -a vX.X -m "..."`）。
   - **Step 2: 自动同步远端**：推送代码与 Tag 至 GitHub 远端仓库（`git push && git push --tags`）。
   - **Step 3: 调用 API 创建 Release**：使用 Python 脚本通过 GitHub REST API 在目标仓库创建正式 Release。
   - **Step 4: 上传 Release 资产**：将编译生成的 `.exe` / `.zip` 便携包作为 Release Assets 自动流式上传。
   - **Step 5: 结果汇报**：发布完成后，在控制台输出 GitHub Release 访问链接与资产清单。

---

## 2. 版本号、Tag 与 Release 命名约束 (Naming Conventions)
1. **版本号统一性**：
   - 严格继承 `app-versioning-rules.md`，初始版本从 V1.0 开始，日常迭代按 V1.1, V1.2... 递增，重大重构按 V2.0 递增。
2. **Git Tag 命名**：
   - 统一格式：`vX.X`（例如：`v1.0`、`v1.1`、`v2.0`），全部小写 `v` 开头。
3. **Release 标题命名 (Release Title)**：
   - 统一格式：`[应用名称] VX.X`（例如：`FQT 17025 体系文控管理工作台 V1.1`、`食品安全标准判定核查系统 V2.0`）。
   - **严禁使用任何夸张后缀**（严禁使用“最终版”、“超级版”、“Pro Max”、“至尊版”等）。

---

## 3. Release 说明文档结构 (Release Notes / Changelog)
每次自动生成的 Release Body 必须采用统一严谨的商务 Markdown 模板，**严禁夹带 LaTeX 符号（如 $\rightarrow$ 等）**，格式如下：

```markdown
## [应用名称] VX.X

### 🚀 新增功能 (Features)
- 列出本版本新增的核心功能模块与特性说明

### 🔧 优化与改进 (Improvements)
- 列出性能提效、UI/UX 改进、算法优化等

### 🐛 问题修复 (Bug Fixes)
- 列出已修复的异常与缺陷

### 📦 资产清单 (Release Assets)
- `[应用名称].exe`：独立便携版可执行文件（无需配置 Python 环境）
- `[应用名称]_vX.X_完整运行包.zip`（若有）
```

---

## 4. 鉴权机制与大文件上传保障 (Authentication & Resilience)
1. **鉴权机制**：
   - 采用 GitHub Personal Access Token (PAT) / 环境变量 `GITHUB_TOKEN` 进行 API 鉴权。
2. **大文件与断点自愈**：
   - 上传大体积 `.exe`（如含 OCR/ONNX 深度学习模型的百兆应用）时，必须采用分块流式上传，遇到网络抖动自动重试（至少 3 次），确保 100% 成功交付。

---

## 5. 目标仓库路由策略 (Repository Routing)
1. **自动识别项目远端**：优先检查项目根目录的 `git remote get-url origin`，直接对准当前绑定的 GitHub 仓库；
2. **新项目自适应绑定**：若当前目录未初始化 Git，自动初始化并关联至 `https://github.com/e2234731-sys/<project-name>.git`。