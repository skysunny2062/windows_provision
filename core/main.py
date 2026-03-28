import argparse
import os
import subprocess
import sys
import importlib.util
from bootstrap import ensure_modules, detect_plugins, parse_winget_txt

_REQUIRED_MODULES = [
    ("pywin32", "win32gui"),
]
SC_DISABLE_OK_CODES = (0, 1060)
SCHTASKS_DELETE_OK_CODES = (0, 1)
XCOPY_OK_CODES = (0, 1)
WINGET_INSTALL_OK_CODES = (0, -1978335189, -1978335140, 2316632107)

ensure_modules(_REQUIRED_MODULES)

_UTILS_PATH = os.path.dirname(os.path.abspath(__file__))
if _UTILS_PATH not in sys.path:
    sys.path.insert(0, _UTILS_PATH)
from utils import (
    info, error, state,
    run_quiet,
    fix_acl, robocopy,
    robocopy_folder, xcopy_folder, sync_programfiles, safe_listdir,
)
import ctypes
import ctypes.wintypes
import datetime
import socket
import ssl
import struct
import time
import urllib.request
import winreg
import win32gui
import win32con
from concurrent.futures import ThreadPoolExecutor
from functools import cache
from phases import run_install

# ── 視窗與系統常數 ────────────────────────────────────
_HWND_TOPMOST   = ctypes.wintypes.HWND(-1)
_HWND_NOTOPMOST = ctypes.wintypes.HWND(-2)
_SWP_FLAGS      = 0x0001 | 0x0002  # SWP_NOSIZE | SWP_NOMOVE

PROGRAM_NAME = "Zack Provision Ver260329"

# ── 路徑定義 ──────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))  # core 目錄
ROOT_DIR   = os.path.dirname(SCRIPT_DIR)                 # 根目錄

CORE_REG_DIR = os.path.join(SCRIPT_DIR, "reg")
CORE_SETUP   = os.path.join(SCRIPT_DIR, "setup")
CORE_THEMES  = os.path.join(SCRIPT_DIR, "themes")
CORE_WINDIR  = os.path.join(SCRIPT_DIR, "windir")
CORE_WINGET  = os.path.join(SCRIPT_DIR, "winget.txt")
CORE_FONTS   = os.path.join(SCRIPT_DIR, "font")

WINDIR        = os.environ["windir"]
SYSTEM_FONTS  = os.path.join(WINDIR, "Fonts")
FONTS_REG    = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"

PROGRAMFILES  = os.environ["ProgramFiles"]
USER_DIR      = os.environ["USERPROFILE"]
DESKTOP       = os.path.join(USER_DIR, "Desktop")
LOCAL_APPDATA = os.environ["LOCALAPPDATA"]
TEMP          = os.environ["TEMP"]
PROGRAM_DATA  = os.environ["ProgramData"]
COMPUTERNAME  = socket.gethostname()

# ── 使用者資料夾定義（用於備份與還原功能）────────────
_USER_FOLDERS = [
    ("桌面", "Desktop"),
    ("下載", "Downloads"),
    ("文件", "Documents"),
    ("圖片", "Pictures"),
    ("音樂", "Music"),
    ("影片", "Videos"),
]
_BACKUP_LABELS = {label for label, _ in _USER_FOLDERS} | {"AnyDesk"}

# ── 外掛模組偵測 ──────────────────────────────────────
PLUGINS = detect_plugins(ROOT_DIR)

# ── 狀態與日誌追蹤 ────────────────────────────────────
FAILURES = state.failures

# ── UI 與視窗工具 ─────────────────────────────────────
def _print_header():
    os.system("cls")
    print(PROGRAM_NAME)
    if COMPUTERNAME:
        print(f"PC-NAME: {COMPUTERNAME}")
    print()

@cache
def _get_console_hwnd():
    """取得當前主控台的視窗代碼 (HWND)"""
    hwnd = win32gui.FindWindow("CASCADIA_HOSTING_WINDOW_CLASS", None)
    if not hwnd:
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    return hwnd

def _maximize_console():
    """將主控台視窗最大化並設為最上層顯示"""
    try:
        hwnd = _get_console_hwnd()
        if hwnd:
            user32 = ctypes.windll.user32
            user32.ShowWindow(hwnd, 3)
            user32.SetWindowPos(hwnd, _HWND_TOPMOST, 0, 0, 0, 0, _SWP_FLAGS)
    except Exception as e:
        error("Warning", "maximize_console", str(e))

def _release_topmost():
    """解除主控台最上層顯示"""
    try:
        hwnd = _get_console_hwnd()
        if hwnd:
            ctypes.windll.user32.SetWindowPos(hwnd, _HWND_NOTOPMOST, 0, 0, 0, 0, _SWP_FLAGS)
    except Exception as e:
        error("Warning", "release_topmost", str(e))

# ── 系統核心工具函式 ──────────────────────────────────
def import_reg(path):
    """匯入單一登錄檔 (.reg)"""
    r = run_quiet(
        ["reg", "import", path],
        check_return=True,
        label=f"reg import {os.path.basename(path)}",
    )
    if r and r.returncode != 0:
        print(f"新增REG失敗: {os.path.basename(path)} (code {r.returncode})")
        error("Registry", os.path.basename(path), f"returncode={r.returncode}")

def import_reg_dir(dir_path):
    """匯入指定目錄下的所有登錄檔 (.reg)"""
    for f in safe_listdir(dir_path):
        if f.lower().endswith(".reg"):
            import_reg(os.path.join(dir_path, f))

# 保留手動快取機制，允許在階段安裝 (Phase Git) 後強制清除重查路徑
_GIT_EXE_CACHE = None

def _get_git_exe():
    global _GIT_EXE_CACHE
    if _GIT_EXE_CACHE:
        return _GIT_EXE_CACHE
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\GitForWindows") as k:
            install_path, _ = winreg.QueryValueEx(k, "InstallPath")
            _GIT_EXE_CACHE = os.path.join(install_path, "bin", "git.exe")
    except Exception as e:
        info(f"_get_git_exe: registry 讀取失敗，fallback 到 PATH ({e})")
        _GIT_EXE_CACHE = "git"
    return _GIT_EXE_CACHE

def _inject_git_path():
    """動態將 Git 路徑注入至當前環境變數，確保後續指令可直接調用"""
    try:
        git_exe = _get_git_exe()
        if git_exe == "git":
            return
        install_path = os.path.dirname(os.path.dirname(git_exe))
        git_bin = os.path.join(install_path, "bin")
        git_cmd = os.path.join(install_path, "cmd")
        current = os.environ.get("PATH", "")
        paths = {p.strip().lower() for p in current.split(";") if p.strip()}
        need_bin = git_bin.lower() not in paths
        need_cmd = git_cmd.lower() not in paths
        if need_bin or need_cmd:
            to_inject = [p for p, needed in [(git_bin, need_bin), (git_cmd, need_cmd)] if needed]
            os.environ["PATH"] = ";".join(to_inject) + ";" + current
    except Exception as e:
        info(f"_inject_git_path: 無法注入 PATH ({e})")

def download_file(url, dest, retries=3):
    """下載檔案，預設關閉 SSL 憑證驗證以避免內網或憑證過期問題"""
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode    = ssl.CERT_NONE
    for attempt in range(retries):
        try:
            with urllib.request.urlopen(url, context=ctx, timeout=15) as r, open(dest, "wb") as f:
                f.write(r.read())
            return True
        except Exception as e:
            info(f"下載失敗 (嘗試 {attempt + 1}/{retries}): {e}")
            time.sleep(3)
    return False

def reg_set(hive, key, name, reg_type, value):
    try:
        with winreg.CreateKeyEx(hive, key, 0, winreg.KEY_SET_VALUE) as k:
            winreg.SetValueEx(k, name, 0, reg_type, value)
    except Exception as e:
        print(f"REG 失敗 [{key}] {name}: {e}")

def sc_disable(service):
    r = run_quiet(["sc", "config", service, "start=", "disabled"], check_return=True, ok_codes=SC_DISABLE_OK_CODES, label=f"disable service {service}")
    return r is not None and r.returncode in SC_DISABLE_OK_CODES

def sc_disable_retry(service):
    for _ in range(2):
        if sc_disable(service):
            return
        time.sleep(2)
    error("Service", service, "sc config failed")

def schtasks_delete(task):
    run_quiet(["schtasks", "/delete", "/tn", task, "/f"], check_return=True, ok_codes=SCHTASKS_DELETE_OK_CODES, label=f"delete task {task}")

# ── 通用等待機制 (Polling Helper) ─────────────────────
def _wait_until(cond_fn, timeout=10, interval=0.1) -> bool:
    """定期檢查給定條件 (cond_fn) 是否成立，用於等待視窗或特定進程出現"""
    for _ in range(int(timeout / interval)):
        if cond_fn():
            return True
        time.sleep(interval)
    return False

# ── 重試與安裝機制 ────────────────────────────────────
def winget_install_pkg(pkg, exact=False, source="winget", _retry_queue=None):
    cmd = ["winget", "install", "--id", pkg,
           "--source", source, "--silent",
           "--accept-package-agreements", "--accept-source-agreements",
           "--disable-interactivity"]
    if exact:
        cmd.append("-e")
    for attempt in range(3):
        if attempt > 0:
            info(f"  winget retry {attempt}/2: {pkg}")
            time.sleep(5)
        try:
            r = subprocess.run(cmd, timeout=600)
            # 處理已知的 winget 成功與警告回傳碼
            if r.returncode in WINGET_INSTALL_OK_CODES:
                return True
        except subprocess.TimeoutExpired:
            info(f"  timeout {pkg}，加入 final retry 佇列")
        except Exception as e:
            info(f"  例外 {pkg}: {e}")
    info(f"  → {pkg} 加入 final retry 佇列")
    if _retry_queue is not None:
        _retry_queue.append((pkg, cmd))
    return False

def _final_retry_installers(pending):
    """在部署尾聲統一對先前失敗的安裝項目進行最終重試"""
    if not pending:
        return
    print(f"\n── Error Retry ({len(pending)} 個) ──")
    time.sleep(15)
    for label, cmd in pending:
        info(f"  error retry: {label}")
        try:
            r = subprocess.run(cmd)
            if r.returncode == 0:
                info(f"  ✓ {label} (final retry OK)")
            else:
                error("Installer", label, f"returncode={r.returncode}")
        except Exception as e:
            error("Installer", label, str(e))

def _get_font_reg_name(font_path):
    """解析字型檔 (TTF/OTF) 的二進位標頭，取得系統登錄檔中需註冊的字型名稱"""
    ext = os.path.splitext(font_path)[1].lower()
    type_suffix = "(OpenType)" if ext == ".otf" else "(TrueType)"
    try:
        with open(font_path, "rb") as f:
            header = f.read(12)
            num_tables = struct.unpack_from(">H", header, 4)[0]
            table_dir  = f.read(num_tables * 16)
            offset = None
            for i in range(num_tables):
                tag = table_dir[i*16: i*16 + 4]
                if tag == b"name":
                    offset = struct.unpack_from(">I", table_dir, i*16 + 8)[0]
                    break
            if offset is None:
                raise ValueError("name table not found")
            f.seek(offset)
            name_data  = f.read(16384)
            count_n    = struct.unpack_from(">H", name_data, 2)[0]
            str_offset = struct.unpack_from(">H", name_data, 4)[0]
            family_name = ""
            full_name   = ""
            for j in range(count_n):
                base = 6 + j * 12
                pid, _, _, nid, length, noffset = struct.unpack_from(">HHHHHH", name_data, base)
                raw = name_data[str_offset + noffset: str_offset + noffset + length]
                if pid == 3:
                    try:
                        s = raw.decode("utf-16-be")
                    except Exception:
                        continue
                    if nid == 4 and not full_name:
                        full_name = s
                    elif nid == 1 and not family_name:
                        family_name = s
            name = full_name or family_name
            if name:
                return f"{name} {type_suffix}"
    except Exception:
        pass
    return f"{os.path.splitext(os.path.basename(font_path))[0]} {type_suffix}"

def install_fonts(font_dir):
    """將字型複製至系統字型庫，跳過已存在的檔案，並寫入對應登錄檔"""
    print()
    if not os.path.isdir(font_dir):
        return
    font_exts = (".ttf", ".otf", ".ttc", ".fon")
    fonts = [f for f in os.listdir(font_dir) if f.lower().endswith(font_exts)]
    if not fonts:
        return
    print("安裝字型中 請稍後...")
    installed = 0
    skipped   = 0
    for fname in fonts:
        dst = os.path.join(SYSTEM_FONTS, fname)
        if os.path.exists(dst):
            skipped += 1
            continue
        src = os.path.join(font_dir, fname)
        try:
            run_quiet(
                ["xcopy", "/y", src, SYSTEM_FONTS + "\\"],
                check_return=True,
                ok_codes=XCOPY_OK_CODES,
                label=f"install font {fname}",
            )
            reg_name = _get_font_reg_name(src)
            reg_set(winreg.HKEY_LOCAL_MACHINE, FONTS_REG, reg_name, winreg.REG_SZ, fname)
            installed += 1
        except Exception as e:
            print(f"字型安裝失敗 {fname}: {e}")
    print(f"安裝字型：{installed} 個，{skipped} 個已存在跳過\n")

def _find_button(root, targets):
    btn = win32gui.FindWindowEx(root, None, "Button", None)
    while btn:
        text = win32gui.GetWindowText(btn)
        if any(t.lower() in text.lower() for t in targets):
            return btn
        btn = win32gui.FindWindowEx(root, btn, "Button", None)
    return None

def set_best_appearance(timeout=10):
    """自動開啟系統效能選項，並點選「調整成最佳外觀」"""
    subprocess.Popen("SystemPropertiesPerformance.exe")
    hwnd = None
    def _find():
        nonlocal hwnd
        hwnd = win32gui.FindWindow(None, "效能選項")
        return bool(hwnd)
    
    if not _wait_until(_find, timeout=timeout):
        info("set_best_appearance: 視窗未出現，跳過")
        return
    
    tab_hwnd = win32gui.FindWindowEx(hwnd, None, "#32770", None)
    search_root = tab_hwnd if tab_hwnd else hwnd
    btn = _find_button(search_root, ["最佳外觀", "best appearance"])
    
    if not btn:
        info("set_best_appearance: 找不到最佳外觀按鈕，跳過")
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        return
    
    win32gui.PostMessage(btn, win32con.BM_CLICK, 0, 0)
    time.sleep(0.2)
    btn = _find_button(hwnd, ["確定", "OK"])
    if btn:
        win32gui.PostMessage(btn, win32con.BM_CLICK, 0, 0)

def wait_window_and_close(title, timeout=10):
    candidates = [title] + [t for t in ["設定", "Windows 設定", "Settings"] if t != title]
    def _find():
        return any(win32gui.FindWindow(None, t) for t in candidates)
    if _wait_until(_find, timeout=timeout):
        run_quiet(
            ["taskkill", "/im", "systemsettings.exe", "/f"],
            check_return=True,
            label="taskkill systemsettings",
        )

# ── 系統設定最佳化 ────────────────────────────────────
def apply_system_settings():
    """執行系統設定優化，包含電源管理、防火牆關閉及清除預設排程任務"""
    system_cmds = [
        ["powercfg", "/x", "monitor-timeout-ac",  "15"],
        ["powercfg", "/x", "-disk-timeout-ac",    "0"],
        ["bcdedit",  "/set", "{bootmgr}", "timeout", "0"],
        ["NetSh", "advfirewall", "set", "allprofiles", "state", "off"],
        ["schtasks", "/change", "/tn", r"\Microsoft\Windows\Defrag\ScheduledDefrag", "/disable"],
        ["powershell", "-Command", "Disable-ComputerRestore -Drive 'C:\\'"],
        ["winget", "uninstall", "--id", "Microsoft.OneDrive", "--silent", "--accept-source-agreements"],
    ]
    for cmd in system_cmds:
        run_quiet(cmd, check_return=True, label=" ".join(cmd))

    for svc in ["CscService", "DusmSvc", "Spooler"]:
        sc_disable_retry(svc)

    for task in [
        r"\Microsoft\Windows\Customer Experience Improvement Program\Consolidator",
        r"\Microsoft\Windows\Customer Experience Improvement Program\UsbCeip",
        r"Microsoft\Windows\Windows Defender\Windows Defender Cache Maintenance",
        r"Microsoft\Windows\Windows Defender\Windows Defender Cleanup",
        r"Microsoft\Windows\Windows Defender\Windows Defender Scheduled Scan",
        r"Microsoft\Windows\Windows Defender\Windows Defender Verification",
    ]:
        schtasks_delete(task)

def _write_failure_log(log_path, install_mode, plugin_dir, bg_names=None):
    """
    執行安裝最終的健康度與狀態驗證，並匯出執行 Log。
    平行檢查 winget 套件安裝狀態，比對服務狀態與必要檔案。
    """
    check_pkgs = parse_winget_txt(CORE_WINGET)
    if plugin_dir:
        check_pkgs += parse_winget_txt(os.path.join(plugin_dir, "winget.txt"))

    def _verify_pkg(pkg_entry):
        pkg_id, _exact, source, display_name = pkg_entry
        name_tail = pkg_id.rsplit(".", 1)[-1] if "." in pkg_id else pkg_id
        queries = [
            ["winget", "list", "--id", pkg_id, "--exact"],
            ["winget", "list", "--id", pkg_id],
        ]
        if source == "msstore":
            queries.extend([
                ["winget", "list", "--id", pkg_id, "--exact", "--source", "msstore"],
                ["winget", "list", "--id", pkg_id, "--source", "msstore"],
            ])
        if display_name:
            queries.extend([
                ["winget", "list", "--name", display_name, "--exact"],
                ["winget", "list", "--name", display_name],
            ])
        if name_tail and name_tail != pkg_id:
            queries.extend([
                ["winget", "list", "--name", name_tail, "--exact"],
                ["winget", "list", "--name", name_tail],
            ])
        try:
            for cmd in queries:
                r = subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                if r.returncode == 0:
                    return None
            return ("--:--:--", "Verify-Package", pkg_id, "winget list 未偵測到（id/name/source fallback 皆失敗）")
        except Exception as e:
            return ("--:--:--", "Verify-Package", pkg_id, str(e))

    with ThreadPoolExecutor(max_workers=8) as ex:
        verify_results = list(ex.map(_verify_pkg, check_pkgs))

    verifyFAILURES = [r for r in verify_results if r is not None]

    mpv_path = os.path.join(PROGRAMFILES, "mpv_PlayKit", "mpv.exe")
    if not os.path.exists(mpv_path):
        verifyFAILURES.append(("--:--:--", "Verify-Package", "mpv_PlayKit", "mpv.exe 不存在"))

    # 反向比對：服務狀態
    for svc in ["CscService", "DusmSvc", "Spooler"]:
        try:
            r = subprocess.run(["sc", "qc", svc],
                               stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            out = r.stdout.decode("cp950", errors="replace")
            start_lines = [l for l in out.splitlines() if "START_TYPE" in l]
            if start_lines and "DISABLED" not in start_lines[0].upper():
                verifyFAILURES.append(("--:--:--", "Verify-Service", svc, start_lines[0].strip()))
        except Exception as e:
            verifyFAILURES.append(("--:--:--", "Verify-Service", svc, str(e)))

    runtime_labels = {f.label for f in FAILURES}
    finalFAILURES = [
        (ts, cat, label, detail)
        for ts, cat, label, detail in verifyFAILURES
        if label not in runtime_labels
    ]

    ts_now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [
        f"{PROGRAM_NAME}  {ts_now}",
        f"PC: {COMPUTERNAME}   Mode: {install_mode}",
        "=" * 50,
    ]

    def fmt_line(ts, cat, label, detail):
        return f"[{ts}] [{cat}] {label}" + (f"  # {detail}" if detail else "")

    if not finalFAILURES:
        lines.append("All OK")
    else:
        lines.append(f"FAILED ITEMS ({len(finalFAILURES)})")
        lines.append("")
        for ts, cat, label, detail in finalFAILURES:
            lines.append(fmt_line(ts, cat, label, detail))

    if FAILURES:
        lines.append("")
        lines.append("[Runtime Warnings - retry 後可能已成功，僅供參考]")
        for item in FAILURES:
            lines.append(fmt_line(item.ts, item.category, item.label, item.detail))

    try:
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
        fail_str = f" ({len(finalFAILURES)} 個失敗)" if finalFAILURES else ""
        print(f"Log：{log_path}\n\n系統部署完成{fail_str}")
        if bg_names:
            names_str = " / ".join(bg_names)
            print(f"{names_str} 於背景執行中，請接手\n")
    except Exception as e:
        print(f"Log 寫入失敗: {e}\n")

# ── 介面與流程控制 ────────────────────────────────────
def _print_confirm_base(pc_input, pc_name, install_mode, crack_label):
    _print_header()
    if pc_input:
        print(f"PC-NAME(NEW): {pc_input}")
    print(install_mode)
    print(crack_label)

def menu():
    options = {
        "1": ("系統部署", install),
        "2": ("資料備份", data_backup),
        "3": ("結束", sys.exit),
    }
    while True:
        _print_header()
        for k, (label, _) in options.items():
            print(f"{k}.{label}")
        choice = input("\n請輸入選項 [1-3]: ")
        if choice not in options:
            continue
        options[choice][1]()

def _collect_install_options():
    """收集使用者部署資訊"""
    _print_header()
    pc_input = input("請輸入電腦名稱: ").strip()

    install_mode = "Normal Mode"
    plugin_mod   = None
    plugin_dir   = None

    if PLUGINS:
        # 僅取第一個偵測到的外掛（目前不支援多外掛防呆）
        mode_name, (mod, p_dir) = next(iter(PLUGINS.items()))
        ans = input(f"{mode_name}? (輸入Y/N繼續): ").upper()
        if ans == "Y":
            install_mode = mode_name
            plugin_mod   = mod
            plugin_dir   = p_dir

    crack      = input("Crack Activation? (輸入Y/N繼續): ").upper() == "Y"
    do_restore = input("DATA Restore? (輸入Y/N繼續): ").upper() == "Y"
    pc_name    = COMPUTERNAME or ""

    return pc_input, pc_name, install_mode, crack, do_restore, plugin_mod, plugin_dir

def _resolve_restore_dir() -> str | None:
    """自動偵測還原備份資料夾（D槽為主），允許使用者手動輸入，回傳目錄路徑或取消(None)"""
    auto_found = None
    try:
        for entry in os.listdir("D:\\"):
            if entry.lower().startswith("backup_") and os.path.isdir(f"D:\\{entry}"):
                candidate_path = f"D:\\{entry}"
                if any(os.path.isdir(os.path.join(candidate_path, l)) for l in _BACKUP_LABELS):
                    auto_found = candidate_path
                    break
    except Exception:
        pass

    candidate = auto_found
    while True:
        if candidate and os.path.isdir(candidate):
            sub_items = [l for l in _BACKUP_LABELS if os.path.isdir(os.path.join(candidate, l))]
            sub_label = "/".join(sub_items) if sub_items else ""
            print(f"\n還原來源：{candidate}  [{sub_label}]")
            if input("備份檔案所在資料夾正確嗎? (輸入Y/N繼續): ").upper() == "Y":
                return candidate
            candidate = None
        else:
            path_input = input("\n請輸入備份檔案所在資料夾(輸入N回主選單):").strip()
            if path_input.upper() == "N":
                return None
            if not os.path.isdir(path_input):
                print("路徑不存在，請重新輸入")
                continue
            if not any(os.path.isdir(os.path.join(path_input, l)) for l in _BACKUP_LABELS):
                print("資料夾無備份檔案，請重新輸入")
                continue
            candidate = path_input

def install():
    pc_input, pc_name, install_mode, crack, do_restore, plugin_mod, plugin_dir = _collect_install_options()
    crack_label = "Crack Activation" if crack else "Do not Crack Activation"

    def _show_confirm(restore_label):
        _print_confirm_base(pc_input, pc_name, install_mode, crack_label)
        print(restore_label)

    restore_dir = None

    if not do_restore:
        _show_confirm("Do not DATA Restore\n")
        if input("Provision Are you sure? (輸入Y/N繼續): ").upper() != "Y":
            return
        _show_confirm("Do not DATA Restore\n")
    else:
        restore_dir = _resolve_restore_dir()
        if restore_dir is None:
            return
        sub_items = [l for l in _BACKUP_LABELS if os.path.isdir(os.path.join(restore_dir, l))]
        restore_label = f"還原資料:{'/'.join(sub_items)}\n備份檔案所在資料夾: {restore_dir}\n"
        _show_confirm(restore_label)
        if input("Provision Are you sure? (輸入Y/N繼續): ").upper() != "Y":
            return
        os.system("cls")
        _show_confirm(restore_label)

    _run_install(pc_input, install_mode, crack, restore_dir, plugin_mod, plugin_dir)

def _reset_git_cache():
    global _GIT_EXE_CACHE
    _GIT_EXE_CACHE = None


def _run_install(pc_input, install_mode, crack, restore_dir, plugin_mod, plugin_dir):
    ctx = {
        "PROGRAM_NAME": PROGRAM_NAME,
        "SCRIPT_DIR": SCRIPT_DIR,
        "CORE_REG_DIR": CORE_REG_DIR,
        "CORE_SETUP": CORE_SETUP,
        "CORE_THEMES": CORE_THEMES,
        "CORE_WINDIR": CORE_WINDIR,
        "CORE_WINGET": CORE_WINGET,
        "CORE_FONTS": CORE_FONTS,
        "WINDIR": WINDIR,
        "PROGRAMFILES": PROGRAMFILES,
        "USER_DIR": USER_DIR,
        "DESKTOP": DESKTOP,
        "LOCAL_APPDATA": LOCAL_APPDATA,
        "TEMP": TEMP,
        "PROGRAM_DATA": PROGRAM_DATA,
        "USER_FOLDERS": _USER_FOLDERS,
        "FAILURES": FAILURES,
        "safe_listdir": safe_listdir,
        "wait_until": _wait_until,
        "run_quiet": run_quiet,
        "winget_install_pkg": winget_install_pkg,
        "reset_git_cache": _reset_git_cache,
        "get_git_exe": _get_git_exe,
        "inject_git_path": _inject_git_path,
        "install_fonts": install_fonts,
        "xcopy_folder": xcopy_folder,
        "fix_acl": fix_acl,
        "parse_winget_txt": parse_winget_txt,
        "import_reg_dir": import_reg_dir,
        "apply_system_settings": apply_system_settings,
        "reg_set": reg_set,
        "download_file": download_file,
        "wait_window_and_close": wait_window_and_close,
        "set_best_appearance": set_best_appearance,
        "sc_disable_retry": sc_disable_retry,
        "final_retry_installers": _final_retry_installers,
        "write_failure_log": _write_failure_log,
        "maximize_console": _maximize_console,
        "release_topmost": _release_topmost,
        "info": info,
        "error": error,
    }
    run_install(pc_input, install_mode, crack, restore_dir, plugin_mod, plugin_dir, ctx)

# ── 資料備份作業 ──────────────────────────────────────
def data_backup():
    anydesk_src = os.path.join(PROGRAM_DATA, "AnyDesk")

    backup_options = {
        str(i + 1): (label, os.path.join(USER_DIR, eng))
        for i, (label, eng) in enumerate(_USER_FOLDERS)
    }
    backup_options[str(len(_USER_FOLDERS) + 1)] = ("AnyDesk", anydesk_src if os.path.isdir(anydesk_src) else None)
    quit_key = str(len(backup_options) + 1)
    backup_options[quit_key] = ("回主選單", None)

    while True:
        _print_header()
        for k, (label, _) in backup_options.items():
            print(f"{k}.{label}")

        choice = input("\n請輸入選項 (多選): ").strip()
        if quit_key in choice:
            return

        # 過濾空白字元避免輸入誤判，並濾除重複的選擇
        seen = set()
        selected = []
        for ch in choice.replace(" ", ""):
            if ch not in seen and ch in backup_options and ch != quit_key:
                seen.add(ch)
                selected.append((ch, backup_options[ch]))
        if not selected:
            continue

        dest_input = input("請輸入備份目錄(預設D槽): ").strip()
        if dest_input == "":
            backup_root = f"D:\\backup_{COMPUTERNAME}"
        else:
            backup_root = dest_input.rstrip("\\") + f"\\backup_{COMPUTERNAME}"

        _print_header()
        labels = "/".join(label for _, (label, _) in selected)
        print(f"備份:{labels}")
        print(f"備份至: {backup_root}")
        confirm = input("\nBackup Are you sure? (輸入Y/N繼續): ").upper()
        if confirm != "Y":
            continue

        for _, (label, src) in selected:
            if src is None:
                print(f"\n跳過 {label}：來源資料夾不存在")
                continue
            dst = os.path.join(backup_root, label)
            print(f"\n備份{label}...")
            robocopy(src, dst)

        print(f"\n{labels} 備份完成")
        os.system("pause")
        return

def _print_dry_run_preview(auto_reason=False):
    _print_header()
    print("[DRY-RUN] 僅列印流程，不執行任何安裝、複製、寫入或登錄變更。\n")
    if auto_reason:
        print("偵測到非 bat 啟動，已自動進入 dry-run。")
    print("環境偵測：")
    print(f"- Python: {sys.executable}")
    print(f"- pywin32: {'OK' if importlib.util.find_spec('win32gui') else 'MISSING'}")
    print(f"- Core winget: {'FOUND' if os.path.isfile(CORE_WINGET) else 'MISSING'}")
    print(f"- Core reg count: {len([f for f in safe_listdir(CORE_REG_DIR) if f.lower().endswith('.reg')])}")
    print(f"- Core theme count: {len([f for f in safe_listdir(CORE_THEMES) if f.lower().endswith('.theme')])}")
    print(f"- Plugins: {', '.join(PLUGINS) if PLUGINS else 'None'}")
    print("\n可用模式：")
    print("1. Normal Mode")
    for idx, mode_name in enumerate(PLUGINS, start=2):
        print(f"{idx}. {mode_name}")
    print("\n預計流程：")
    phases = [
        "Restore (若選擇還原)",
        "Office (若有 .img)",
        "Git",
        "Files",
        "mpv_PlayKit",
        "Winget",
        "System",
        "VisualCppRedistAIO",
        "Theme",
        "Setup",
        "Final Retry + Verify Log",
    ]
    for idx, phase in enumerate(phases, start=1):
        print(f"{idx:02d}. {phase}")
    print("\n(按任意鍵後結束)")
    print()
    os.system("pause")


def _parse_args():
    parser = argparse.ArgumentParser(add_help=True)
    parser.add_argument("--dry-run", action="store_true", help="僅列印流程，不執行部署")
    parser.add_argument("--run", action="store_true", help="強制進入正式部署選單")
    return parser.parse_args()


if __name__ == "__main__":
    ctypes.windll.kernel32.SetConsoleTitleW(PROGRAM_NAME)
    args = _parse_args()
    launched_from_bat = os.environ.get("WP_LAUNCH_MODE") == "full"
    auto_dry_run = (not launched_from_bat) and (not args.run) and (not args.dry_run)
    if args.dry_run or auto_dry_run:
        os.system("COLOR 04")
        _print_dry_run_preview(auto_reason=auto_dry_run)
    else:
        os.system("COLOR 0B")
        menu()
