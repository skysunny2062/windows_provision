import importlib.util
import os
import subprocess
import sys
_REQUIRED_MODULES = [
    ("pywin32", "win32gui"),
]

def ensure_modules(mods, max_restart=2):
    """自動安裝缺少的 pip 套件，失敗超過 max_restart 次則中止"""
    missing = [pip for pip, imp in mods if importlib.util.find_spec(imp) is None]
    if not missing:
        return
    restart_count = int(os.environ.get("AUTO_INSTALL_RESTART_COUNT", "0"))
    if restart_count >= max_restart:
        print(f"錯誤：已重啟 {max_restart} 次，套件仍無法載入，請手動安裝：")
        print("  pip install " + " ".join(missing))
        os.system("pause")
        sys.exit(1)
    print(f"正在安裝pip: {', '.join(missing)}")
    try:
        subprocess.check_call([
            sys.executable, "-m", "pip", "install",
            *missing,
            "--upgrade",
            "--quiet",
            "--disable-pip-version-check",
            "--no-warn-script-location",
            "--no-cache-dir"
        ])
    except subprocess.CalledProcessError as e:
        print(f"pip 安裝失敗: {e}")
        print("請手動執行: pip install " + " ".join(missing))
        os.system("pause")
        sys.exit(1)
    print("正在重新啟動腳本...")
    new_env = os.environ.copy()
    new_env["AUTO_INSTALL_RESTART_COUNT"] = str(restart_count + 1)
    os.execve(sys.executable, [sys.executable, *sys.argv], new_env)

ensure_modules(_REQUIRED_MODULES)

_UTILS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)))
if _UTILS_PATH not in sys.path:
    sys.path.insert(0, _UTILS_PATH)
from utils import (
    info, error, _FAILURES,
    run_silent, _fix_acl, _robocopy,
    robocopy_folder, xcopy_folder, sync_programfiles, safe_listdir,
)
import base64
import ctypes
import ctypes.wintypes
import datetime
import json
import socket
import ssl
import stat
import struct
import time
import urllib.request
import winreg
import win32gui
import win32con

PROGRAM_NAME = "Zack Provision Ver260327"
# ── 路徑定義 ──────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))  # core\
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

# ── 使用者資料夾定義（備份/還原共用）────────────────
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
def _detect_plugins():
    """
    掃描 ROOT_DIR 下的子資料夾，若 xxx/xxx.py 存在則載入為外掛。
    回傳 dict：{ "Xxx Mode": (module, xxx_dir) }
    """
    plugins = {}
    try:
        for entry in os.listdir(ROOT_DIR):
            plugin_dir = os.path.join(ROOT_DIR, entry)
            plugin_py  = os.path.join(plugin_dir, f"{entry}.py")
            if os.path.isdir(plugin_dir) and os.path.isfile(plugin_py):
                mode_name = f"{entry.capitalize()} Mode"
                try:
                    spec   = importlib.util.spec_from_file_location(entry, plugin_py)
                    module = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(module)
                    plugins[mode_name] = (module, plugin_dir)
                except Exception as e:
                    print(f"外掛載入失敗 [{entry}]: {e}")
    except Exception:
        pass
    return plugins

PLUGINS = _detect_plugins()   # { "xxx Mode": (module, path), ... }

# ── winget.txt 解析 ───────────────────────────────────
def _parse_winget_txt(path):
    """
    解析 winget.txt，每行格式：
      pkg_id[,exact][,msstore][,name=顯示名稱]
    回傳 list of (pkg_id, exact:bool, source:str, display_name:str|None)
    忽略空行與 # 開頭的註解行。
    """
    pkgs = []
    if not os.path.isfile(path):
        return pkgs
    with open(path, encoding="utf-8") as f:
        for raw in f:
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            parts       = [p.strip() for p in line.split(",")]
            pkg_id      = parts[0]
            exact       = "exact" in parts[1:]
            source      = "msstore" if "msstore" in parts[1:] else "winget"
            name_part   = next((p for p in parts[1:] if p.lower().startswith("name=")), None)
            display_name = name_part.split("=", 1)[1] if name_part else None
            pkgs.append((pkg_id, exact, source, display_name))
    return pkgs

# ── 進度 Log ──────────────────────────────────────────
LOG_PATH = None
FAILURES = _FAILURES   # utils 持有的同一份 list

# ── UI 工具 ───────────────────────────────────────────
def _print_header():
    os.system("cls")
    print(PROGRAM_NAME)
    if COMPUTERNAME:
        print(f"PC-NAME: {COMPUTERNAME}")
    print()

# ── 工具函式 ──────────────────────────────────────────
_CONSOLE_HWND_CACHE = None

def _get_console_hwnd():
    global _CONSOLE_HWND_CACHE
    if _CONSOLE_HWND_CACHE:
        return _CONSOLE_HWND_CACHE
    hwnd = win32gui.FindWindow("CASCADIA_HOSTING_WINDOW_CLASS", None)
    if not hwnd:
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    _CONSOLE_HWND_CACHE = hwnd
    return hwnd

def _maximize_console():
    try:
        hwnd = _get_console_hwnd()
        if hwnd:
            user32 = ctypes.windll.user32
            user32.ShowWindow(hwnd, 3)
            HWND_TOPMOST = ctypes.wintypes.HWND(-1)
            user32.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, 0x0001 | 0x0002)
    except Exception as e:
        error("Warning", "maximize_console", str(e))

def _release_topmost():
    try:
        hwnd = _get_console_hwnd()
        if hwnd:
            HWND_NOTOPMOST = ctypes.wintypes.HWND(-2)
            ctypes.windll.user32.SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, 0x0001 | 0x0002)
    except Exception as e:
        error("Warning", "release_topmost", str(e))

def import_reg(path):
    r = run_silent(["reg", "import", path])
    if r and r.returncode != 0:
        print(f"新增REG失敗: {os.path.basename(path)} (code {r.returncode})")
        error("Registry", os.path.basename(path), f"returncode={r.returncode}")

def import_reg_dir(dir_path):
    for f in safe_listdir(dir_path):
        if f.lower().endswith(".reg"):
            import_reg(os.path.join(dir_path, f))

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
    r = run_silent(["sc", "config", service, "start=", "disabled"])
    return r is not None and r.returncode in (0, 1060)

def sc_disable_retry(service):
    for _ in range(2):
        if sc_disable(service):
            return
        time.sleep(2)
    error("Service", service, "sc config failed")

def schtasks_delete(task):
    run_silent(["schtasks", "/delete", "/tn", task, "/f"])

# ── Retry 機制 ────────────────────────────────────────
pending_final_retry = []

def winget_install_pkg(pkg, exact=False, source="winget"):
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
            if r.returncode in (0, -1978335189, -1978335140, 2316632107):
                return True
        except subprocess.TimeoutExpired:
            info(f"  timeout {pkg}，加入 final retry 佇列")
        except Exception as e:
            info(f"  例外 {pkg}: {e}")
    info(f"  → {pkg} 加入 final retry 佇列")
    pending_final_retry.append((pkg, cmd))
    return False

def _final_retry_installers():
    if not pending_final_retry:
        return
    print(f"\n── Error Retry ({len(pending_final_retry)} 個) ──")
    time.sleep(15)
    for label, cmd in pending_final_retry:
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
    """複製字型到 C:\\Windows\\Fonts，已存在就跳過，並寫入 Registry"""
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
            run_silent(["xcopy", "/y", src, SYSTEM_FONTS + "\\"])
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
    subprocess.Popen("SystemPropertiesPerformance.exe")
    hwnd = None
    for _ in range(timeout * 10):
        hwnd = win32gui.FindWindow(None, "效能選項")
        if hwnd:
            break
        time.sleep(0.1)
    if not hwnd:
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
    for _ in range(timeout * 10):
        for t in candidates:
            hwnd = win32gui.FindWindow(None, t)
            if hwnd:
                run_silent(["taskkill", "/im", "systemsettings.exe", "/f"])
                return
        time.sleep(0.1)

# ── 系統設定 ──────────────────────────────────────────
def apply_system_settings():
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
        run_silent(cmd)

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

def _write_failure_log(install_mode, plugin_dir, bg_names=None):
    """
    最終反向比對驗證。
    winget 驗證清單從 core/winget.txt 讀取，若有外掛也讀 xxx/winget.txt。
    """
    if not LOG_PATH:
        return

    verifyFAILURES = []

    # ── 反向比對：winget 套件（從 txt 動態讀取）────────
    check_pkgs = _parse_winget_txt(CORE_WINGET)
    if plugin_dir:
        check_pkgs += _parse_winget_txt(os.path.join(plugin_dir, "winget.txt"))

    for pkg_id, *_ in check_pkgs:
        try:
            r = subprocess.run(
                ["winget", "list", "--id", pkg_id, "--exact"],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )
            if r.returncode != 0:
                verifyFAILURES.append(("Verify-Package", pkg_id, "winget list --id 未偵測到"))
        except Exception as e:
            verifyFAILURES.append(("Verify-Package", pkg_id, str(e)))

    mpv_path = os.path.join(PROGRAMFILES, "mpv_PlayKit", "mpv.exe")
    if not os.path.exists(mpv_path):
        verifyFAILURES.append(("Verify-Package", "mpv_PlayKit", "mpv.exe 不存在"))

    # ── 反向比對：服務狀態 ────────────────────────────
    for svc in ["CscService", "DusmSvc", "Spooler"]:
        try:
            r = subprocess.run(["sc", "qc", svc],
                               stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            out = r.stdout.decode("cp950", errors="replace")
            start_lines = [l for l in out.splitlines() if "START_TYPE" in l]
            if start_lines and "DISABLED" not in start_lines[0].upper():
                verifyFAILURES.append(("Verify-Service", svc, start_lines[0].strip()))
        except Exception as e:
            verifyFAILURES.append(("Verify-Service", svc, str(e)))

    runtime_labels = {f[2] for f in FAILURES}
    finalFAILURES = [
        ("--:--:--", cat, label, detail)
        for cat, label, detail in verifyFAILURES
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
        for ts, cat, label, detail in FAILURES:
            lines.append(fmt_line(ts, cat, label, detail))

    try:
        with open(LOG_PATH, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
        fail_str = f" ({len(finalFAILURES)} 個失敗)" if finalFAILURES else ""
        print(f"Log：{LOG_PATH}\n\n系統部署完成{fail_str}")
        if bg_names:
            names_str = " / ".join(bg_names)
            print(f"{names_str} 於背景執行中，請接手\n")
    except Exception as e:
        print(f"Log 寫入失敗: {e}\n")

# ── 主選單 ────────────────────────────────────────────
def _print_confirm_base(pc_input, pc_name, install_mode, crack_label):
    os.system("cls")
    print(PROGRAM_NAME)
    print(f"PC-NAME: {pc_name}")
    print()
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

# ── install 問卷 ──────────────────────────────────────
def _collect_install_options():
    """
    問卷順序：
      電腦名稱
      → [若有外掛] xxx Mode? Y/N
      → Crack Activation? Y/N
      → DATA Restore? Y/N
    """
    _print_header()
    pc_input = input("請輸入電腦名稱: ").strip()

    install_mode = "Normal Mode"
    plugin_mod   = None
    plugin_dir   = None

    if PLUGINS:
        # 只取第一個偵測到的外掛（多外掛不做防呆）
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

# ── install ───────────────────────────────────────────
def install():
    pc_input, pc_name, install_mode, crack, do_restore, plugin_mod, plugin_dir = \
        _collect_install_options()
    crack_label = "Crack Activation" if crack else "Do not Crack Activation"

    def _show_confirm(restore_label):
        _print_confirm_base(pc_input, pc_name, install_mode, crack_label)
        print(restore_label)

    restore_dir = None

    if not do_restore:
        _show_confirm("Do not DATA Restore\n")
        confirm = input("Provision Are you sure? (輸入Y/N繼續): ").upper()
        if confirm != "Y":
            return
        _show_confirm("Do not DATA Restore\n")
    else:
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
                sub_items = [l for l in _BACKUP_LABELS
                             if os.path.isdir(os.path.join(candidate, l))]
                sub_label = "/".join(sub_items) if sub_items else ""
                restore_label = f"還原資料:{sub_label}\n備份檔案所在資料夾: {candidate}\n"
                _show_confirm(restore_label)
                ok = input("備份檔案所在資料夾正確嗎? (輸入Y/N繼續): ").upper()
                if ok == "Y":
                    restore_dir = candidate
                    confirm = input("Provision Are you sure? (輸入Y/N繼續): ").upper()
                    if confirm != "Y":
                        return
                    os.system("cls")
                    _show_confirm(restore_label)
                    break
                else:
                    candidate = None
            else:
                path_input = input("\n請輸入備份檔案所在資料夾(輸入N回主選單):").strip()
                if path_input.upper() == "N":
                    return
                if not os.path.isdir(path_input):
                    print("路徑不存在，請重新輸入")
                    continue
                has_backup = any(
                    os.path.isdir(os.path.join(path_input, label))
                    for label in _BACKUP_LABELS
                )
                if not has_backup:
                    print("資料夾無備份檔案，請重新輸入")
                    continue
                candidate = path_input

    _run_install(pc_input, pc_name, install_mode, crack, restore_dir, plugin_mod, plugin_dir)

# ── Install Phases ────────────────────────────────────
def _phase_restore(restore_dir):
    if not restore_dir:
        return
    anydesk_dst = os.path.join(PROGRAM_DATA, "AnyDesk")
    for label, eng in _USER_FOLDERS:
        src = os.path.join(restore_dir, label)
        dst = os.path.join(USER_DIR, eng)
        if os.path.isdir(src):
            cmd = (f'start "還原{label}" cmd /k '
                   f'"COLOR 0B & robocopy "{src}" "{dst}" /s /mt:8 /r:10 /w:3 /xf desktop.ini"')
            subprocess.Popen(cmd, shell=True)
    anydesk_src = os.path.join(restore_dir, "AnyDesk")
    if os.path.isdir(anydesk_src):
        cmd = (f'start "還原AnyDesk" cmd /k '
               f'"COLOR 0B & robocopy "{anydesk_src}" "{anydesk_dst}" /mir /mt:8 /r:10 /w:3"')
        subprocess.Popen(cmd, shell=True)
    print("開始還原資料...")

def _phase_office():
    office_dir = os.path.join(PROGRAMFILES, "Microsoft Office")
    if os.path.isdir(office_dir):
        return None, None
    img_files = [f for f in safe_listdir(SCRIPT_DIR) if f.lower().endswith(".img")]
    if not img_files:
        print("Core資料夾內無Office的img 此次將略過Office")
        return None, None
    img_path = os.path.join(SCRIPT_DIR, img_files[0])
    r = subprocess.run(
        ["powershell", "-Command",
         f"(Mount-DiskImage -ImagePath '{img_path}' -PassThru | Get-Volume).DriveLetter"],
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True
    )
    drive = r.stdout.strip()
    if not drive:
        print("Office 掛載失敗")
        return None, None
    setup_exe = f"{drive}:\\Setup.exe"
    for _ in range(50):
        if os.path.exists(setup_exe):
            break
        time.sleep(0.1)
    if not os.path.exists(setup_exe):
        print("Office 找不到 Setup.exe")
        return None, None
    print("安裝 Office...")
    office_proc = subprocess.Popen([setup_exe], cwd=os.path.dirname(setup_exe))
    return img_path, office_proc

def _phase_office_cleanup(img_path, office_proc):
    if office_proc is not None or img_path:
        run_silent(["taskkill", "/f", "/im", "OfficeC2RClient.exe"])
    if img_path:
        run_silent(["powershell", "-Command",
                    f"Dismount-DiskImage -ImagePath '{img_path}'"])

def _phase_git():
    global _GIT_EXE_CACHE
    print("安裝 Git...")
    winget_install_pkg("Git.Git", exact=False, source="winget")
    os.system("COLOR 0B")
    for _ in range(5):
        _GIT_EXE_CACHE = None
        if _get_git_exe() != "git":
            break
        time.sleep(1)
    _inject_git_path()

def _phase_files(plugin_mod):
    """
    [共同] 字型 / themes / windir
    [外掛] 呼叫 plugin_mod.custom_files()  ← api1
    """
    themes_dst = os.path.join(WINDIR, "Resources", "Themes")
    install_fonts(CORE_FONTS)
    if os.path.isdir(CORE_THEMES):
        xcopy_folder(CORE_THEMES, themes_dst)
    if os.path.isdir(CORE_WINDIR):
        xcopy_folder(CORE_WINDIR, WINDIR)
    if plugin_mod and hasattr(plugin_mod, "custom_files"):
        plugin_mod.custom_files()
    return themes_dst

def _phase_mpv():
    mpv_dst = os.path.join(PROGRAMFILES, "mpv_PlayKit")
    print("安裝 mpv_PlayKit...")
    try:
        git_exe = _get_git_exe()
        if os.path.isdir(mpv_dst):
            r_git = subprocess.run([git_exe, "pull"], cwd=mpv_dst)
        else:
            r_git = subprocess.run(
                [git_exe, "clone", "--depth=1", "https://github.com/skysunny2062/mpv_PlayKit.git"],
                cwd=PROGRAMFILES
            )
        if r_git.returncode != 0:
            raise RuntimeError(f"git 失敗 returncode={r_git.returncode}")
        _fix_acl(mpv_dst)
        mpv_com = os.path.join(mpv_dst, "mpv.com")
        subprocess.run([mpv_com, "--config=no", "--register"])
        mpv_exe  = os.path.join(mpv_dst, "mpv.exe")
        lnk_path = os.path.join(DESKTOP, "mpv_PlayKit.lnk")
        ps = (
            f'$s=(New-Object -COM WScript.Shell).CreateShortcut("{lnk_path}");'
            f'$s.TargetPath="{mpv_exe}";'
            f'$s.WorkingDirectory="{mpv_dst}";'
            f'$s.Save()'
        )
        run_silent(["powershell", "-Command", ps])
        os.system("COLOR 0B")
    except Exception as e:
        print(f"mpv_PlayKit 失敗: {e}")
        error("Installer", "mpv_PlayKit", str(e))

def _install_winget_list(path):
    for pkg_id, exact, source, display_name in _parse_winget_txt(path):
        _name = display_name or (pkg_id.split(".", 1)[1] if "." in pkg_id else pkg_id)
        print(f"\n安裝 {_name}...")
        winget_install_pkg(pkg_id, exact=exact, source=source)
        os.system("COLOR 0B")

def _phase_winget(plugin_dir):
    """
    [共同] 讀 core/winget.txt 安裝
    [外掛] 讀 xxx/winget.txt 安裝  ← api2
    """
    _install_winget_list(CORE_WINGET)
    if plugin_dir:
        _install_winget_list(os.path.join(plugin_dir, "winget.txt"))

def _phase_system(install_mode, plugin_dir, pc_input):
    """
    [共同疊加] core/reg/*.reg
    [互斥]     Normal → core/*.reg  /  xxx Mode → xxx/*.reg
    [共同]     apply_system_settings / 電腦名稱
    """
    import_reg_dir(CORE_REG_DIR)
    if plugin_dir:
        import_reg_dir(plugin_dir)
    else:
        import_reg_dir(SCRIPT_DIR)
    apply_system_settings()

    if pc_input:
        HKLM = winreg.HKEY_LOCAL_MACHINE
        SZ   = winreg.REG_SZ
        for key, val in [
            (r"SYSTEM\CurrentControlSet\Services\lanmanserver\Parameters",        "srvcomment"),
            (r"SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName", "ComputerName"),
            (r"SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName",       "ComputerName"),
            (r"SYSTEM\CurrentControlSet\Services\Tcpip\Parameters",               "NV Hostname"),
            (r"System\CurrentControlSet\Services\Tcpip\Parameters",               "Hostname"),
        ]:
            reg_set(HKLM, key, val, SZ, pc_input)
    YYMMDD = datetime.datetime.now().strftime("%y%m%d")
    reg_set(winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation",
            "Model", winreg.REG_SZ, f"{PROGRAM_NAME} build:{YYMMDD}")

def _phase_vcredist():
    print("\n安裝 VisualCppRedistAIO...")
    try:
        _ssl_ctx = ssl.create_default_context()
        _ssl_ctx.check_hostname = False
        _ssl_ctx.verify_mode    = ssl.CERT_NONE
        req = urllib.request.Request(
            "https://api.github.com/repos/abbodi1406/vcredist/releases/latest",
            headers={"User-Agent": "Mozilla/5.0"}
        )
        with urllib.request.urlopen(req, context=_ssl_ctx) as r:
            data = json.loads(r.read())
        for asset in data["assets"]:
            if asset["name"].lower() == "visualcppredist_aio_x86_x64.exe":
                dest = os.path.join(TEMP, asset["name"])
                if download_file(asset["browser_download_url"], dest):
                    subprocess.run([dest, "/y"])
                    os.remove(dest)
                else:
                    error("Installer", "VisualCppRedistAIO", "下載重試耗盡")
                break
    except Exception as e:
        print(f"VisualCppRedistAIO 下載失敗: {e}")
        error("Installer", "VisualCppRedistAIO", str(e))

def _phase_crack(crack):
    if not crack:
        return
    print("CrackActivation...")
    winrar_key = os.path.join(PROGRAMFILES, "WinRAR", "rarreg.key")
    os.makedirs(os.path.dirname(winrar_key), exist_ok=True)
    with open(winrar_key, "w") as f:
        f.write(
            "RAR registration data\n"
            "SeVeN\n"
            "Unlimited Company License\n"
            "UID=000de082d4cb7aeb3e71\n"
            "64122122503e71057c0ffe5fed5dcbb0032f2d3c5fd42bb05edfe0\n"
            "6501b129e6e067e3819160fce6cb5ffde62890079861be57638717\n"
            "7131ced835ed65cc743d9777f2ea71a8e32c7e593cf66794343565\n"
            "b41bcf56929486b8bcdac33d50ecf77399603fc518f2701b607304\n"
            "9b712761e333304a99485f38f292bc89b78036ec4c0faa35b6c3d6\n"
            "df05d84217aef4abd0b675d3309b94f4be9c9cae1734784060e0c7\n"
            "ef6e1deace43f1671ef3ef6d863944c13f3fc13a20f21488793187\n"
        )
    ps_script = "& ([ScriptBlock]::Create((irm https://get.activated.win))) /HWID /Ohook"
    encoded   = base64.b64encode(ps_script.encode("utf-16-le")).decode()
    try:
        subprocess.run(["powershell", "-ExecutionPolicy", "Bypass",
                        "-EncodedCommand", encoded], timeout=120)
    except subprocess.TimeoutExpired:
        error("Activation", "Windows Activation", "PowerShell timeout (120s)")
    except Exception as e:
        error("Activation", "Windows Activation", str(e))

def _phase_theme(themes_dst):
    theme_files = [f for f in safe_listdir(CORE_THEMES) if f.lower().endswith(".theme")]
    if theme_files:
        theme_path = os.path.join(themes_dst, theme_files[0])
        os.startfile(theme_path)
        wait_window_and_close("設定")
    subprocess.run(["taskkill", "/f", "/im", "explorer.exe"])
    iconcache = os.path.join(LOCAL_APPDATA, "iconcache.db")
    if os.path.exists(iconcache):
        try:
            run_silent(["attrib", "-s", "-r", "-h", iconcache])
            os.chmod(iconcache, stat.S_IWRITE)
            os.remove(iconcache)
        except Exception as e:
            print(f"iconcache 刪除失敗: {e}")
    time.sleep(2)
    subprocess.Popen(["explorer.exe"])

def _phase_setup(plugin_mod):
    """
    [共同] core/setup/*.exe 背景執行
    [外掛] 呼叫 plugin_mod.custom_setup()  ← api3
    """
    _bg_names = []
    for f in safe_listdir(CORE_SETUP):
        fp = os.path.join(CORE_SETUP, f)
        if os.path.isfile(fp) and fp.lower().endswith(".exe"):
            subprocess.Popen(f'cmd /c start "" /min "{fp}"', shell=True)
            _bg_names.append(os.path.splitext(f)[0])
    if plugin_mod and hasattr(plugin_mod, "custom_setup"):
        plugin_bg = plugin_mod.custom_setup()
        if plugin_bg:
            _bg_names.extend(plugin_bg)
    return _bg_names

# ── _run_install ──────────────────────────────────────
def _run_install(pc_input, pc_name, install_mode, crack, restore_dir, plugin_mod, plugin_dir):
    global LOG_PATH, pending_final_retry
    FAILURES.clear()
    pending_final_retry = []
    ts = datetime.datetime.now().strftime("%y%m%d_%H%M%S")
    LOG_PATH = os.path.join(DESKTOP, f"{PROGRAM_NAME}_Log_{ts}.txt")
    _maximize_console()

    _phase_restore(restore_dir)
    img_path, office_proc = _phase_office()
    _phase_git()
    themes_dst = _phase_files(plugin_mod)
    _phase_mpv()
    _phase_office_cleanup(img_path, office_proc)
    _phase_winget(plugin_dir)
    _phase_system(install_mode, plugin_dir, pc_input)
    set_best_appearance()
    _release_topmost()
    _phase_vcredist()
    _phase_crack(crack)
    _phase_theme(themes_dst)
    _bg_names = _phase_setup(plugin_mod)

    _final_retry_installers()
    _write_failure_log(install_mode, plugin_dir, _bg_names)
    sc_disable_retry("CryptSvc")
    os.system("pause")

# ── data_backup ───────────────────────────────────────
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

        selected = []
        seen = set()
        for ch in choice:
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
            _robocopy(src, dst)

        print(f"\n{labels} 備份完成")
        os.system("pause")
        return

if __name__ == "__main__":
    ctypes.windll.kernel32.SetConsoleTitleW(PROGRAM_NAME)
    menu()