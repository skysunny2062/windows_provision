"""
zack/zack.py  ─ Zack Mode 外掛

main.py 在偵測到 zack/zack.py 時自動載入，並在對應階段呼叫：
  custom_files()   ← api1：_phase_files 尾段
  custom_setup()   ← api3：_phase_setup 尾段

winget 套件請寫在 zack/winget.txt（api2，main.py 直接掃描，不需寫 code）
登錄檔請放在 zack/*.reg（main.py 直接掃描，不需寫 code）
"""

import os
import subprocess
import sys

# ── 載入共用工具 ──────────────────────────────────────
_CORE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "core")
if _CORE_DIR not in sys.path:
    sys.path.insert(0, _CORE_DIR)
from utils import xcopy_folder, robocopy_folder, sync_programfiles, safe_listdir

# ── 路徑（相對於本檔案所在的 zack\ 資料夾）────────────
_DIR          = os.path.dirname(os.path.abspath(__file__))
_PROGRAMFILES = os.environ["ProgramFiles"]
_APPDATA      = os.environ["APPDATA"]
_LOCAL        = os.environ["LOCALAPPDATA"]
_DESKTOP      = os.path.join(os.environ["USERPROFILE"], "Desktop")

ZACK_APPDATA      = os.path.join(_DIR, "appdata")
ZACK_C            = os.path.join(_DIR, "c")
ZACK_PROGRAMFILES = os.path.join(_DIR, "programfiles")
ZACK_DESKTOP      = os.path.join(_DIR, "desktop")


# ── api1 ──────────────────────────────────────────────
def custom_files():
    """Zack 專屬資料夾複製（在 core 字型/主題複製完後呼叫）"""
    if os.path.isdir(ZACK_APPDATA):
        xcopy_folder(ZACK_APPDATA, _APPDATA)

    for folder in safe_listdir(ZACK_C):
        src = os.path.join(ZACK_C, folder)
        if os.path.isdir(src):
            robocopy_folder(src, os.path.join("C:\\", folder))

    sync_programfiles(ZACK_PROGRAMFILES, _PROGRAMFILES)

    if os.path.isdir(ZACK_DESKTOP):
        xcopy_folder(ZACK_DESKTOP, _DESKTOP)

    print()


# ── api3 ──────────────────────────────────────────────
def custom_setup():
    """Zack 專屬安裝（在 core/setup 執行完後呼叫）"""
    setup_dir = os.path.join(_DIR, "setup")
    install4j = os.path.join(setup_dir, "Install4j")
    bg_names  = []

    # Install4j 安裝檔（同步等待）
    for f in safe_listdir(install4j):
        if f.lower().endswith((".exe", ".msi")):
            print(f"安裝 {os.path.splitext(f)[0]}...")
            subprocess.run([os.path.join(install4j, f), "-q"])

    # zack/setup 根目錄的 exe（背景）
    for f in safe_listdir(setup_dir):
        fp = os.path.join(setup_dir, f)
        if os.path.isfile(fp) and f.lower().endswith(".exe"):
            subprocess.Popen(f'cmd /c start "" /min "{fp}"', shell=True)
            bg_names.append(os.path.splitext(f)[0])

    # Locale Emulator
    le_dir       = os.path.join(_PROGRAMFILES, "Locale Emulator")
    le_installer = os.path.join(le_dir, "LEInstaller.exe")
    if os.path.exists(le_installer):
        subprocess.Popen(
            f'cmd /c start "" /min /D "{le_dir}" "{le_installer}"',
            shell=True
        )
        bg_names.append("Locale Emulator")

    return bg_names