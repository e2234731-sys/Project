# -*- coding: utf-8 -*-
"""
食品安全国家标准判定核查工具 V1.4 发布脚本
自动化创建 GitHub Release V1.4 并上传 Windows 独立打包二进制文件
"""
import os
import sys
import json
import urllib.request
import urllib.parse
import subprocess

sys.stdout.reconfigure(encoding='utf-8')


def get_github_token():
    proc = subprocess.Popen(["git", "credential", "fill"], stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True)
    out, _ = proc.communicate(input="protocol=https\nhost=github.com\n\n")
    for line in out.splitlines():
        if line.startswith("password="):
            return line.split("=", 1)[1].strip()
    return ""


def main():
    token = get_github_token()
    if not token:
        print("Error: Could not retrieve GitHub token from git credential helper.")
        sys.exit(1)

    repo = "e2234731-sys/Project"
    tag_name = "V1.4"

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "User-Agent": "Antigravity-Release-Publisher",
        "X-GitHub-Api-Version": "2022-11-28"
    }

    print(f"=== Checking Release {tag_name} for {repo} ===")
    url_tag = f"https://api.github.com/repos/{repo}/releases/tags/{tag_name}"
    release_data = None
    try:
        req = urllib.request.Request(url_tag, headers=headers)
        with urllib.request.urlopen(req) as resp:
            release_data = json.loads(resp.read().decode("utf-8"))
            print(f"Existing release {tag_name} found.")
    except Exception:
        pass

    release_body = """### 🚀 食品安全国家标准判定核查工具 V1.4 正式发布

#### 🌟 核心升级与重构亮点：

1. **💥 单元格多项目智能拆分与自动拆行展开 (炸裂拆行引擎)**：
   - **痛点攻关**：针对各级市场监管局抽检方案（如《2026年广东省食用农产品必检品种项目表》及多地市局抽检计划）中，在**同一个单元格**内以顿号（`、`）、中文/英文逗号（`，,`）、分号（`；;`）或换行符合并填报多个检测项目的实际业务难题；
   - **括号深度保护状态机**：系统内置自研括号深度跟踪算法，对包含内部枚举说明的复合项目（如 `合成着色剂（柠檬黄、日落黄、胭脂红、苋菜红、诱惑红、亮蓝）`、`土霉素/金霉素/四环素（组合含量）`、`磺胺类（总量）` 等）进行原生括号保护，绝不误拆分，完整保留指标语义；
   - **四级分类与属性无损继承**：拆分后的每个检验项目自动炸裂为独立的一行检测任务，完整继承所属的 `食品大类（一级）`、`食品亚类（二级）`、`食品品种（三级）`、`食品细类（四级）`、原始序号与抽检备注，逐行智能反写判定标准、检测方法、限量值及检出限；
   - **表头自适应匹配与界面开关**：自动识别 `必检项目`、`可选项目`、`检验项目`，操作卡片增设 `☑️ 单元格多项目自动拆行展开` 选项，一键切换。

2. **📑 四级分类 + 检验项目 双引擎智能匹配反写 (6大国标 83,973 条全量数据)**：
   - 深度融合 **《2026年国家食品安全监督抽检实施细则》全量库（477 个四级细类、4,310 项指标）** 与 GB 2763-2026 (4.8万条)、GB 2760-2024 (3.2万条)、GB 31650-2019 等现行标准；
   - 原地反写或扩展反写 **【依据法律法规或标准】**、**【检测方法】**、**【标准限量值】**、**【方法检出限(LOD)】** 与 **【判定结论】**；
   - 纵向合并单元格自动拆分向下填充，横幅标题行智能跳过，原生兼容 WPS `.xlsx` 格式。

3. **🧭 下拉式功能导航菜单 + 午夜蓝现代商务 UI**：
   - 顶部整合下拉菜单（`Combobox`）与胶囊快捷切换按钮，页面全容器支持高 DPI 自适应平滑滚动；
   - 表格各列支持自由拖拽调节宽度并自动记忆持久化（`ui_settings.json`），数据行支持悬浮气泡查看完整方法及划词复制。

4. **🔐 离线“一机一码”授权防漂移与单实例锁**：
   - 采用注册表 `MachineGuid` 毫秒级校验、三路径冗余存储与 Windows 单实例互斥锁 (`Mutex`)，已激活电脑秒级静默进入主界面。

---
⚡ *Powered by 陈耿忠 (CCAA 国家注册审核员 / 实验室质控专家)*
"""

    if not release_data:
        url_create = f"https://api.github.com/repos/{repo}/releases"
        body = {
            "tag_name": tag_name,
            "target_commitish": "main",
            "name": "食品安全国家标准判定核查工具 V1.4",
            "body": release_body,
            "draft": False,
            "prerelease": False
        }
        req_post = urllib.request.Request(url_create, data=json.dumps(body).encode("utf-8"), headers=headers, method="POST")
        with urllib.request.urlopen(req_post) as resp:
            release_data = json.loads(resp.read().decode("utf-8"))
            print(f"Created release {tag_name} successfully!")

    upload_base = release_data.get("upload_url").split("{")[0]
    existing_assets = {a.get("name"): a.get("id") for a in release_data.get("assets", [])}

    output_dir = os.path.abspath(r"C:\Users\e2234\AppData\Roaming\Tencent\Marvis\User\oAN1i2Wmo3r9X5fl-BSOd67YTfFo\workspace\conv_19f31bcee46_868597138dd2\output")

    uploads = [
        ("食品安全标准判定核查工具(完整版)_批量表头筛选.exe", "FoodSafetyStandardChecker_Full_BatchFilter_V1.4.exe"),
        ("食品安全国家标准判定核查工具V1.4.exe", "FoodSafetyStandardChecker_V1.4.exe"),
        ("食品安全标准判定核查工具(纯净版)V1.4.exe", "FoodSafetyStandardChecker_Clean_V1.4.exe"),
        ("食品安全国家标准判定核查工具V1.4.exe", "食品安全国家标准判定核查工具V1.4.exe"),
        ("食品安全标准判定核查工具(纯净版)V1.4.exe", "食品安全标准判定核查工具(纯净版)V1.4.exe")
    ]

    for local_name, remote_name in uploads:
        local_path = os.path.join(output_dir, local_name)
        if not os.path.exists(local_path):
            print(f"Warning: {local_path} does not exist, skipping upload.")
            continue

        file_size = os.path.getsize(local_path)
        print(f"\nUploading {remote_name} ({file_size/1024/1024:.2f} MB)...")

        if remote_name in existing_assets:
            asset_id = existing_assets[remote_name]
            del_url = f"https://api.github.com/repos/{repo}/releases/assets/{asset_id}"
            del_req = urllib.request.Request(del_url, headers=headers, method="DELETE")
            try:
                urllib.request.urlopen(del_req)
                print(f"Deleted previous asset {remote_name} (id: {asset_id})")
            except Exception as e:
                print(f"Could not delete old asset: {e}")

        query = urllib.parse.urlencode({"name": remote_name})
        upload_url = f"{upload_base}?{query}"

        upload_headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream",
            "Content-Length": str(file_size),
            "User-Agent": "Antigravity-Release-Publisher",
            "X-GitHub-Api-Version": "2022-11-28"
        }

        with open(local_path, "rb") as f:
            data = f.read()

        up_req = urllib.request.Request(upload_url, data=data, headers=upload_headers, method="POST")
        try:
            with urllib.request.urlopen(up_req) as up_resp:
                result = json.loads(up_resp.read().decode("utf-8"))
                download_url = result.get("browser_download_url")
                print(f"Uploaded successfully! Download URL: {download_url}")
        except Exception as e:
            print(f"Upload failed for {remote_name}: {e}")

    print(f"\nAll operations completed! View release at: https://github.com/{repo}/releases/tag/{tag_name}")


if __name__ == "__main__":
    main()
