# -*- coding: utf-8 -*-
"""
一机一码离线授权验证与注册管理核心组件 (License Manager V1.3)
Author: 陈耿忠
支持激活码激活、.lic 授权证书文件一键导入激活、同目录自动识别激活、防硬件漂移与毫秒级极速校验
"""

import os
import sys
import glob
import json
import base64
import hashlib
import hmac
import subprocess
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ttkbootstrap as tb
from ttkbootstrap.constants import *

# 核心安全加盐密钥 (仅陈耿忠老师专属)
MASTER_SALT = "CHEN_GENG_ZHONG_FQT_MASTER_SALT_2026_@#$!"
LICENSE_FILE_NAME = ".software_license.dat"
CREATE_NO_WINDOW = 0x08000000 if sys.platform == "win32" else 0


def get_host_guid():
    """毫秒级获取 Windows 注册表唯一 MachineGuid + 计算机名 (不受U盘插拔或WMI超时影响)"""
    guid = ""
    if sys.platform == "win32":
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Cryptography",
                0,
                winreg.KEY_READ | getattr(winreg, "KEY_WOW64_64KEY", 0)
            )
            val, _ = winreg.QueryValueEx(key, "MachineGuid")
            winreg.CloseKey(key)
            if val:
                guid = str(val).strip()
        except Exception:
            pass
    comp = os.environ.get("COMPUTERNAME", "PC-DEFAULT")
    user = os.environ.get("USERNAME", "USER-DEFAULT")
    return f"{guid}_{comp}_{user}"


def _format_mc_from_parts(cpu_id, board_sn, disk_sn):
    raw_combined = f"{cpu_id}_{board_sn}_{disk_sn}_FQT_CGZ"
    h = hashlib.sha256(raw_combined.encode('utf-8')).hexdigest().upper()
    return f"FQT-{h[0:4]}-{h[4:8]}-{h[8:12]}-{h[12:16]}"


def get_all_candidate_fingerprints():
    """
    获取当前电脑所有可能的硬件指纹集合（兼容U盘/移动硬盘插拔、WMI单组件超时及回退特征）
    返回: (primary_mc, candidate_set)
    """
    cpu_list = [""]
    board_list = [""]
    disk_list = [""]

    # 1. CPU ProcessorId
    try:
        out = subprocess.check_output(
            "wmic cpu get processorid",
            shell=True, stderr=subprocess.DEVNULL, timeout=2, creationflags=CREATE_NO_WINDOW
        ).decode('utf-8', errors='ignore')
        lines = [l.strip() for l in out.split('\n') if l.strip() and 'ProcessorId' not in l]
        if lines:
            cpu_list = [lines[0], ""]
    except Exception:
        pass

    # 2. 主板序列号
    try:
        out = subprocess.check_output(
            "wmic baseboard get serialnumber",
            shell=True, stderr=subprocess.DEVNULL, timeout=2, creationflags=CREATE_NO_WINDOW
        ).decode('utf-8', errors='ignore')
        lines = [l.strip() for l in out.split('\n') if l.strip() and 'SerialNumber' not in l]
        if lines:
            board_list = [lines[0], ""]
    except Exception:
        pass

    # 3. 硬盘物理序列号 (遍历全部磁盘，优先固定内置硬盘，防止U盘改变lines[0])
    fixed_disks = []
    all_disks = []
    try:
        out = subprocess.check_output(
            "wmic diskdrive get mediatype,serialnumber",
            shell=True, stderr=subprocess.DEVNULL, timeout=2, creationflags=CREATE_NO_WINDOW
        ).decode('utf-8', errors='ignore')
        for l in out.split('\n'):
            s = l.strip()
            if not s or 'SerialNumber' in s:
                continue
            parts = s.split()
            sn = parts[-1].strip()
            all_disks.append(sn)
            if 'Fixed' in s or 'fixed' in s:
                fixed_disks.append(sn)
    except Exception:
        pass

    if not all_disks:
        try:
            out = subprocess.check_output(
                "wmic diskdrive get serialnumber",
                shell=True, stderr=subprocess.DEVNULL, timeout=2, creationflags=CREATE_NO_WINDOW
            ).decode('utf-8', errors='ignore')
            all_disks = [l.strip() for l in out.split('\n') if l.strip() and 'SerialNumber' not in l]
        except Exception:
            pass

    ordered_disks = []
    for d in (fixed_disks + all_disks + [""]):
        if d not in ordered_disks:
            ordered_disks.append(d)
    disk_list = ordered_disks

    candidates = set()
    primary_mc = None

    for c in cpu_list:
        for b in board_list:
            for d in disk_list:
                if not (c or b or d):
                    continue
                mc = _format_mc_from_parts(c, b, d)
                if primary_mc is None:
                    primary_mc = mc
                candidates.add(mc)

    # 备用环境变量特征
    fb_c = os.environ.get("COMPUTERNAME", "PC-DEFAULT")
    fb_b = os.environ.get("USERNAME", "USER-DEFAULT")
    fb_d = str(os.environ.get("PROCESSOR_IDENTIFIER", "CPU-GENERIC"))
    fb_mc = _format_mc_from_parts(fb_c, fb_b, fb_d)
    candidates.add(fb_mc)

    # 注册表 MachineGuid 特征
    guid_mc = _format_mc_from_parts(get_host_guid(), fb_c, fb_b)
    candidates.add(guid_mc)

    if primary_mc is None:
        primary_mc = fb_mc

    return primary_mc, candidates


def get_hardware_fingerprint():
    """提取本机主硬件特征码 (若本地已有合法激活记录，优先保持一致)"""
    lic_data = load_local_license()
    if lic_data and lic_data.get("mc"):
        return lic_data.get("mc")
    primary_mc, _ = get_all_candidate_fingerprints()
    return primary_mc


def generate_license_signature(machine_code, app_id, lic_type, expiry_date, client_name):
    """根据机器码和授权信息生成防伪签名"""
    payload = f"{machine_code}|{app_id}|{lic_type}|{expiry_date}|{client_name}|{MASTER_SALT}"
    sig = hmac.new(MASTER_SALT.encode('utf-8'), payload.encode('utf-8'), hashlib.sha256).hexdigest().upper()
    return sig[:16]


def create_activation_key(machine_code, app_id="ALL", lic_type="PERM", expiry_date="20991231", client_name="FQT用户"):
    """生成激活码 (作者注册机调用)"""
    clean_mc = machine_code.replace("FQT-", "").replace("-", "").upper()
    if len(clean_mc) < 16:
        clean_mc = clean_mc.ljust(16, '0')

    sig = generate_license_signature(machine_code, app_id, lic_type, expiry_date, client_name)
    data_dict = {
        "mc": machine_code,
        "app": app_id,
        "type": lic_type,
        "exp": expiry_date,
        "client": client_name,
        "sig": sig
    }
    json_bytes = json.dumps(data_dict).encode('utf-8')
    b64_code = base64.b32encode(json_bytes).decode('utf-8').rstrip('=')

    chunks = [b64_code[i:i+6] for i in range(0, len(b64_code), 6)]
    formatted_key = "ACT-" + "-".join(chunks)
    return formatted_key


def generate_license_file_content(machine_code, app_id="ALL", lic_type="PERM", expiry_date="20991231", client_name="FQT用户"):
    """生成 .lic 授权证书文件内容 (JSON格式)"""
    key = create_activation_key(machine_code, app_id, lic_type, expiry_date, client_name)
    sig = generate_license_signature(machine_code, app_id, lic_type, expiry_date, client_name)

    lic_obj = {
        "format_version": "1.0",
        "machine_code": machine_code,
        "application_id": app_id,
        "license_type": lic_type,
        "expiry_date": expiry_date,
        "client_name": client_name,
        "activation_key": key,
        "digital_signature": sig,
        "issued_by": "陈耿忠 (国家注册审核员 / 实验室质控专家)",
        "issued_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    return json.dumps(lic_obj, ensure_ascii=False, indent=2)


def verify_activation_key(activation_key, current_machine_code=None, target_app_id="ALL"):
    """
    校验激活码是否合法且未过期 (支持多候选硬件特征容错匹配)
    返回: (is_valid, message, license_info_dict)
    """
    try:
        clean_key = activation_key.strip().replace("ACT-", "").replace("-", "").replace("\n", "").replace("\r", "").replace(" ", "")
        pad_len = (8 - len(clean_key) % 8) % 8
        clean_key += "=" * pad_len

        json_bytes = base64.b32decode(clean_key.encode('utf-8'))
        info = json.loads(json_bytes.decode('utf-8'))

        mc = info.get("mc")
        app = info.get("app")
        lic_type = info.get("type")
        exp = info.get("exp")
        client = info.get("client")
        sig = info.get("sig")

        # 1. 校验防伪签名
        expected_sig = generate_license_signature(mc, app, lic_type, exp, client)
        if sig != expected_sig:
            return False, "激活码签名验证失败，可能已被篡改！", None

        # 2. 校验硬件指纹绑定 (兼容主机器码及全部候选硬件特征)
        if current_machine_code and mc != current_machine_code:
            _, candidates = get_all_candidate_fingerprints()
            if mc not in candidates:
                return False, f"激活码与当前电脑机器码不匹配！\n(该授权绑定机器码为: {mc})", None

        # 3. 校验应用适用范围
        if app != "ALL" and target_app_id != "ALL" and app != target_app_id:
            return False, f"该激活码仅适用于产品【{app}】，不能用于当前系统！", None

        # 4. 校验有效到期日
        today_str = datetime.now().strftime("%Y%m%d")
        if exp != "20991231" and today_str > exp:
            exp_formatted = f"{exp[:4]}-{exp[4:6]}-{exp[6:]}"
            return False, f"该授权已于 {exp_formatted} 到期失效，请联系陈耿忠老师续订授权！", None

        # 记录当前主机唯一标识，后续启动毫秒级放行，彻底杜绝因U盘或WMI超时重复弹窗
        info["host_guid"] = get_host_guid()
        return True, "授权激活成功！", info

    except Exception as e:
        return False, f"激活码格式不正确或解析失败 ({e})", None


def get_all_license_paths():
    """获取多重冗余本地加密授权证书路径，确保任何权限环境下均可持久化读写"""
    paths = []
    # 路径1: APPDATA/FQT_LicenseCenter
    app_data = os.environ.get("APPDATA", "")
    if app_data:
        try:
            d1 = os.path.join(app_data, "FQT_LicenseCenter")
            os.makedirs(d1, exist_ok=True)
            paths.append(os.path.join(d1, LICENSE_FILE_NAME))
        except Exception:
            pass
    # 路径2: 用户主目录 ~/.food_standard_checker
    try:
        d2 = os.path.join(os.path.expanduser("~"), ".food_standard_checker")
        os.makedirs(d2, exist_ok=True)
        p2 = os.path.join(d2, LICENSE_FILE_NAME)
        if p2 not in paths:
            paths.append(p2)
    except Exception:
        pass
    # 路径3: 程序所在目录
    try:
        exe_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
        p3 = os.path.join(exe_dir, LICENSE_FILE_NAME)
        if p3 not in paths:
            paths.append(p3)
    except Exception:
        pass
    return paths


def get_license_file_path():
    paths = get_all_license_paths()
    return paths[0] if paths else LICENSE_FILE_NAME


def save_license(license_info):
    """将合法授权信息同步加密保存到所有冗余路径"""
    if "host_guid" not in license_info:
        license_info["host_guid"] = get_host_guid()
    raw_str = json.dumps(license_info, ensure_ascii=False)
    enc_bytes = base64.b64encode((raw_str + "||" + MASTER_SALT).encode('utf-8'))
    for lic_path in get_all_license_paths():
        try:
            with open(lic_path, 'wb') as f:
                f.write(enc_bytes)
        except Exception:
            pass


def load_local_license():
    """从任一持久化路径读取本地授权证书"""
    for lic_path in get_all_license_paths():
        if not os.path.exists(lic_path):
            continue
        try:
            with open(lic_path, 'rb') as f:
                enc_bytes = f.read()
            raw_full = base64.b64decode(enc_bytes).decode('utf-8')
            raw_str = raw_full.split("||")[0]
            data = json.loads(raw_str)
            if isinstance(data, dict) and data.get("mc") and data.get("sig"):
                return data
        except Exception:
            continue
    return None


def _try_auto_import_adjacent_lic(target_app_id="ALL"):
    """若程序同级目录或上级目录放置了 *.lic 授权证书文件，自动静默识别并激活"""
    search_dirs = []
    try:
        exe_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
        search_dirs.append(exe_dir)
    except Exception:
        pass
    try:
        cwd = os.getcwd()
        if cwd not in search_dirs:
            search_dirs.append(cwd)
    except Exception:
        pass

    for d in search_dirs:
        for lic_file in glob.glob(os.path.join(d, "*.lic")):
            try:
                with open(lic_file, "r", encoding="utf-8") as f:
                    content = f.read().strip()
                key_to_test = ""
                if content.startswith("{"):
                    lic_obj = json.loads(content)
                    key_to_test = lic_obj.get("activation_key", "")
                else:
                    key_to_test = content
                if key_to_test:
                    is_valid, _, lic_info = verify_activation_key(key_to_test, None, target_app_id)
                    if is_valid and lic_info:
                        save_license(lic_info)
                        return True
            except Exception:
                continue
    return False


def check_and_require_license(root_window=None, app_name="食品安全国家标准判定核查工具", app_id="ALL"):
    """
    全局授权拦截入口：
    - 优先毫秒级读取本地授权签名与主机标识 (< 0.001s)，彻底消除因WMI超时或U盘插拔引起的重复弹窗
    - 若同目录存在有效 .lic 证书文件，自动静默导入激活
    - 若确实未激活，弹出激活窗口拦截
    """
    lic_data = load_local_license()

    if lic_data:
        mc = lic_data.get("mc", "")
        app = lic_data.get("app", "ALL")
        lic_type = lic_data.get("type", "PERM")
        exp = lic_data.get("exp", "20991231")
        client = lic_data.get("client", "用户")
        sig = lic_data.get("sig", "")
        saved_guid = lic_data.get("host_guid", "")

        expected_sig = generate_license_signature(mc, app, lic_type, exp, client)
        today_str = datetime.now().strftime("%Y%m%d")

        if sig == expected_sig and (exp == "20991231" or today_str <= exp):
            cur_guid = get_host_guid()
            # 若 saved_guid 匹配当前主机，或旧版本(V1.1/V1.2)已激活本机但尚未写入 host_guid，立即补齐并静默放行
            if not saved_guid or saved_guid == cur_guid:
                if saved_guid != cur_guid:
                    lic_data["host_guid"] = cur_guid
                    save_license(lic_data)
                return True
            # 若 host_guid 变化，再核对候选硬件指纹
            _, candidates = get_all_candidate_fingerprints()
            if mc in candidates:
                lic_data["host_guid"] = cur_guid
                save_license(lic_data)
                return True

    # 尝试自动检测程序同级目录下的 .lic 文件
    if _try_auto_import_adjacent_lic(app_id):
        return True

    # 确属未授权电脑，计算主机器码并弹出激活对话框
    current_mc, _ = get_all_candidate_fingerprints()
    dlg = ActivationDialog(root_window, app_name, app_id, current_mc)
    if root_window:
        root_window.wait_window(dlg.dialog)
    else:
        dlg.dialog.mainloop()
    return dlg.is_activated


class ActivationDialog:
    """高质感一机一码注册激活弹窗 (支持密匙粘贴 & 授权文件导入)"""
    def __init__(self, parent, app_name, app_id, machine_code):
        self.parent = parent
        self.app_name = app_name
        self.app_id = app_id
        self.machine_code = machine_code
        self.is_activated = False

        if parent is None:
            self.dialog = tb.Window(themename="cosmo")
        else:
            self.dialog = tb.Toplevel(parent)

        self.dialog.title("🔐 软件离线授权激活中心 (一机一码)")
        self.dialog.geometry("680x520")
        self.dialog.resizable(False, False)

        # 居中显示
        self.dialog.update_idletasks()
        sw = self.dialog.winfo_screenwidth()
        sh = self.dialog.winfo_screenheight()
        x = max(0, (sw - 680) // 2)
        y = max(0, (sh - 520) // 2)
        self.dialog.geometry(f"680x520+{x}+{y}")

        # 确保窗口立即可见并置顶获得焦点
        self.dialog.deiconify()
        self.dialog.lift()
        self.dialog.attributes('-topmost', True)
        self.dialog.after(150, lambda: self.dialog.attributes('-topmost', False))
        self.dialog.focus_force()

        self.setup_ui()
        self.dialog.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_ui(self):
        container = tb.Frame(self.dialog, padding=20)
        container.pack(fill=tk.BOTH, expand=True)

        tb.Label(
            container,
            text=f"🛡️ {self.app_name}",
            font=("Microsoft YaHei", 13, "bold"),
            bootstyle=PRIMARY
        ).pack(anchor="w", pady=(0, 4))

        tb.Label(
            container,
            text="本软件采用纯离线'一机一码'硬件绑定授权保护，首次使用请粘贴激活码或导入授权文件激活（激活后永久免弹窗）。",
            font=("Microsoft YaHei", 9),
            bootstyle=SECONDARY
        ).pack(anchor="w", pady=(0, 10))

        mc_card = tb.LabelFrame(container, text="💻 本机专属机器码 (Hardware Fingerprint)", padding=10)
        mc_card.pack(fill=tk.X, pady=(0, 10))

        row_mc = tb.Frame(mc_card)
        row_mc.pack(fill=tk.X)
        self.mc_entry = tb.Entry(row_mc, font=("Consolas", 12, "bold"), state="readonly")
        self.mc_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        self.mc_entry.config(state="normal")
        self.mc_entry.insert(0, self.machine_code)
        self.mc_entry.config(state="readonly")

        tb.Button(row_mc, text="📋 一键复制机器码", command=self.copy_machine_code, bootstyle=INFO).pack(side=tk.RIGHT)

        tb.Label(
            mc_card,
            text="💡 操作指引：点击上方按钮复制机器码，通过微信发送给软件作者【陈耿忠老师】获取专属激活码或授权文件。",
            font=("Microsoft YaHei", 8),
            bootstyle=SECONDARY
        ).pack(anchor="w", pady=(4, 0))

        key_card = tb.LabelFrame(container, text="🔑 方式一：录入专属激活码 (Activation Key)", padding=10)
        key_card.pack(fill=tk.X, pady=(0, 10))

        self.key_text = tk.Text(key_card, height=3, font=("Consolas", 9), wrap=tk.WORD)
        self.key_text.pack(fill=tk.X, pady=(0, 6))

        btn_row = tb.Frame(key_card)
        btn_row.pack(fill=tk.X)
        tb.Button(btn_row, text="⚡ 立即验证并激活软件", command=self.do_activate, bootstyle=SUCCESS).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        tb.Button(btn_row, text="📋 粘贴剪贴板", command=self.paste_from_clipboard, bootstyle="secondary-outline").pack(side=tk.RIGHT)

        file_card = tb.LabelFrame(container, text="📁 方式二：导入离线授权证书文件 (.lic)", padding=10)
        file_card.pack(fill=tk.X, pady=(0, 12))

        row_f = tb.Frame(file_card)
        row_f.pack(fill=tk.X)
        tb.Label(row_f, text="如果您已收到陈耿忠老师发送的授权证书文件 (*.lic)，可直接一键导入激活：", font=("Microsoft YaHei", 8), bootstyle=SECONDARY).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tb.Button(row_f, text="📁 选择并导入授权文件 (.lic)", command=self.import_license_file, bootstyle="primary-outline").pack(side=tk.RIGHT)

        footer = tb.Frame(container)
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        tb.Label(footer, text="⚡ Powered by 陈耿忠  |  国家注册审核员 / 实验室质控专家", font=("Microsoft YaHei", 8), bootstyle=PRIMARY).pack(side=tk.LEFT)
        tb.Button(footer, text="退出软件", command=self.on_close, bootstyle="danger-link").pack(side=tk.RIGHT)

    def copy_machine_code(self):
        self.dialog.clipboard_clear()
        self.dialog.clipboard_append(self.machine_code)
        messagebox.showinfo("复制成功", f"本机机器码已成功复制到剪贴板：\n\n{self.machine_code}\n\n请发送给陈耿忠老师获取激活码或授权文件！", parent=self.dialog)

    def paste_from_clipboard(self):
        try:
            txt = self.dialog.clipboard_get().strip()
            if txt:
                self.key_text.delete("1.0", tk.END)
                self.key_text.insert(tk.END, txt)
        except Exception:
            pass

    def do_activate(self):
        raw_key = self.key_text.get("1.0", tk.END).strip()
        if not raw_key:
            messagebox.showwarning("提示", "请输入或粘贴激活码！", parent=self.dialog)
            return

        is_valid, msg, lic_info = verify_activation_key(raw_key, self.machine_code, self.app_id)
        if is_valid:
            save_license(lic_info)
            self.is_activated = True
            exp_str = lic_info.get("exp")
            exp_desc = "永久终身授权" if exp_str == "20991231" else f"有效期至 {exp_str[:4]}-{exp_str[4:6]}-{exp_str[6:]}"
            messagebox.showinfo(
                "🎉 激活成功",
                f"恭喜！软件已成功激活！\n\n"
                f"【授权对象】: {lic_info.get('client')}\n"
                f"【授权期限】: {exp_desc}\n"
                f"【适用软件】: {lic_info.get('app')}\n\n"
                f"点击确定立即进入系统。",
                parent=self.dialog
            )
            self.dialog.destroy()
        else:
            messagebox.showerror("激活失败", f"激活码校验未通过：\n\n{msg}", parent=self.dialog)

    def import_license_file(self):
        file_path = filedialog.askopenfilename(
            title="选择授权证书文件",
            filetypes=[("授权证书文件 (*.lic)", "*.lic"), ("JSON 文件 (*.json)", "*.json"), ("所有文件 (*.*)", "*.*")],
            parent=self.dialog
        )
        if not file_path:
            return

        try:
            with open(file_path, "r", encoding="utf-8") as f:
                content = f.read().strip()

            key_to_test = ""
            if content.startswith("{"):
                lic_data = json.loads(content)
                key_to_test = lic_data.get("activation_key", "")
            else:
                key_to_test = content

            if not key_to_test:
                messagebox.showerror("导入失败", "授权证书文件内容为空或格式损坏！", parent=self.dialog)
                return

            is_valid, msg, lic_info = verify_activation_key(key_to_test, self.machine_code, self.app_id)
            if is_valid:
                save_license(lic_info)
                self.is_activated = True
                exp_str = lic_info.get("exp")
                exp_desc = "永久终身授权" if exp_str == "20991231" else f"有效期至 {exp_str[:4]}-{exp_str[4:6]}-{exp_str[6:]}"
                messagebox.showinfo(
                    "🎉 授权文件导入激活成功",
                    f"恭喜！授权文件校验通过，软件已成功激活！\n\n"
                    f"【证书文件】: {os.path.basename(file_path)}\n"
                    f"【授权对象】: {lic_info.get('client')}\n"
                    f"【授权期限】: {exp_desc}\n"
                    f"【适用软件】: {lic_info.get('app')}\n\n"
                    f"点击确定立即进入系统。",
                    parent=self.dialog
                )
                self.dialog.destroy()
            else:
                messagebox.showerror("授权文件校验未通过", f"该授权证书无法激活当前电脑：\n\n{msg}", parent=self.dialog)

        except Exception as e:
            messagebox.showerror("读取错误", f"读取授权文件失败：{e}", parent=self.dialog)

    def on_close(self):
        if not self.is_activated:
            try:
                self.dialog.destroy()
            except Exception:
                pass
            os._exit(0)
