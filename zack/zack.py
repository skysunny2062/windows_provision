"""
zack/zack.py ─ Zack Mode 擴充外掛模組

此模組會在 main.py 執行時被自動偵測並載入，作為部署流程的擴充掛鉤 (Hook)。
支援的自定義擴充點包含：
1. custom_files(): 於系統檔案複製階段 (Phase Files) 尾聲被呼叫。
2. custom_setup(): 於背景安裝階段 (Phase Setup) 尾聲被呼叫。

備註：
- Winget 套件請直接寫入 zack/winget.txt，主程式會自動掃描並安裝。
- 系統登錄檔請直接放置於 zack/*.reg，主程式會自動掃描並匯入。
"""

import os
import subprocess
import sys

# ── 載入核心共用工具 ──────────────────────────────────
_CORE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "core")
if _CORE_DIR not in sys.path:
    sys.path.insert(0, _CORE_DIR)
from utils import xcopy_folder, robocopy_folder, sync_programfiles, safe_listdir

# ── 環境變數與路徑定義 (相對於 zack 目錄) ─────────────
_DIR = os.path.dirname(os.path.abspath(__file__))
_PROGRAMFILES = os.environ["ProgramFiles"]
_APPDATA = os.environ["APPDATA"]
_LOCAL_APPDATA = os.environ["LOCALAPPDATA"]
_DESKTOP = os.path.join(os.environ["USERPROFILE"], "Desktop")

ZACK_APPDATA = os.path.join(_DIR, "appdata")
ZACK_LOCAL_APPDATA = os.path.join(_DIR, "localappdata")
ZACK_C = os.path.join(_DIR, "c")
ZACK_PROGRAMFILES = os.path.join(_DIR, "programfiles")
ZACK_DESKTOP = os.path.join(_DIR, "desktop")


# ── 掛鉤實作 (Hook Implementations) ───────────────────
def _spawn_minimized_exe(path, cwd=None):
    cmd = ["cmd", "/c", "start", "", "/min"]
    if cwd:
        cmd.extend(["/D", cwd])
    cmd.append(path)
    subprocess.Popen(cmd)


def _iter_installers(root_dir, exts):
    for name in safe_listdir(root_dir):
        path = os.path.join(root_dir, name)
        if os.path.isfile(path) and name.lower().endswith(exts):
            yield name, path


def _append_bg(bg_names, name):
    if name not in bg_names:
        bg_names.append(name)


def custom_files():
    """
    自訂檔案部署流程。
    負責將外掛目錄下的專屬設定檔與資料夾，同步到對應的系統或使用者路徑下。
    此函式會在核心系統字型與主題複製完成後自動觸發。
    """
    folder_sync_jobs = [
        (ZACK_APPDATA, _APPDATA, xcopy_folder),
        (ZACK_LOCAL_APPDATA, _LOCAL_APPDATA, xcopy_folder),
        (ZACK_DESKTOP, _DESKTOP, xcopy_folder),
    ]
    for src, dst, sync_fn in folder_sync_jobs:
        if os.path.isdir(src):
            sync_fn(src, dst)

    for folder in safe_listdir(ZACK_C):
        src = os.path.join(ZACK_C, folder)
        if os.path.isdir(src):
            robocopy_folder(src, os.path.join("C:\\", folder))

    sync_programfiles(ZACK_PROGRAMFILES, _PROGRAMFILES)
    print()


def custom_setup() -> list[str]:
    """
    自訂安裝檔執行流程。
    負責處理外掛目錄下的特定安裝程式。此函式會在 core/setup 執行完成後自動觸發。

    回傳值：
        list[str]: 正在背景執行中的安裝程式名稱清單，供主程式記錄與提示使用者接手。
    """
    setup_dir = os.path.join(_DIR, "setup")
    if not os.path.isdir(setup_dir):
        return []

    install4j = os.path.join(setup_dir, "Install4j")
    bg_names: list[str] = []

    for name, path in _iter_installers(install4j, (".exe", ".msi")):
        app_name = os.path.splitext(name)[0]
        print(f"安裝 {app_name}...")
        try:
            subprocess.run([path, "-q"], timeout=300)
        except subprocess.TimeoutExpired:
            print(f"安裝 {app_name} 逾時 (300s)，繼續執行")
        except Exception as e:
            print(f"安裝 {app_name} 失敗：{e}")

    for name, path in _iter_installers(setup_dir, (".exe",)):
        _spawn_minimized_exe(path)
        _append_bg(bg_names, os.path.splitext(name)[0])

    le_dir = os.path.join(_PROGRAMFILES, "Locale Emulator")
    le_installer = os.path.join(le_dir, "LEInstaller.exe")
    if os.path.exists(le_installer):
        _spawn_minimized_exe(le_installer, cwd=le_dir)
        _append_bg(bg_names, "Locale Emulator")

    return bg_names
