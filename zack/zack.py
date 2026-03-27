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
_DIR          = os.path.dirname(os.path.abspath(__file__))
_PROGRAMFILES = os.environ["ProgramFiles"]
_APPDATA      = os.environ["APPDATA"]
_LOCAL        = os.environ["LOCALAPPDATA"]
_DESKTOP      = os.path.join(os.environ["USERPROFILE"], "Desktop")

ZACK_APPDATA      = os.path.join(_DIR, "appdata")
ZACK_C            = os.path.join(_DIR, "c")
ZACK_PROGRAMFILES = os.path.join(_DIR, "programfiles")
ZACK_DESKTOP      = os.path.join(_DIR, "desktop")


# ── 掛鉤實作 (Hook Implementations) ───────────────────

def custom_files():
    """
    自訂檔案部署流程。
    負責將外掛目錄下的專屬設定檔與資料夾，同步到對應的系統或使用者路徑下。
    此函式會在核心系統字型與主題複製完成後自動觸發。
    """
    # 同步 AppData 設定
    if os.path.isdir(ZACK_APPDATA):
        xcopy_folder(ZACK_APPDATA, _APPDATA)

    # 複製 C 槽根目錄資料夾
    for folder in safe_listdir(ZACK_C):
        src = os.path.join(ZACK_C, folder)
        if os.path.isdir(src):
            robocopy_folder(src, os.path.join("C:\\", folder))

    # 同步 Program Files 軟體資料
    sync_programfiles(ZACK_PROGRAMFILES, _PROGRAMFILES)

    # 複製桌面捷徑或檔案
    if os.path.isdir(ZACK_DESKTOP):
        xcopy_folder(ZACK_DESKTOP, _DESKTOP)

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

    # 處理 Install4j 安裝檔（同步執行，會阻塞主流程直到安裝完成）
    for f in safe_listdir(install4j):
        if f.lower().endswith((".exe", ".msi")):
            print(f"安裝 {os.path.splitext(f)[0]}...")
            subprocess.run([os.path.join(install4j, f), "-q"])

    # 處理 setup 根目錄的安裝檔（背景靜默執行，不阻塞主流程）
    for f in safe_listdir(setup_dir):
        fp = os.path.join(setup_dir, f)
        if os.path.isfile(fp) and f.lower().endswith(".exe"):
            subprocess.Popen(f'cmd /c start "" /min "{fp}"', shell=True)
            bg_names.append(os.path.splitext(f)[0])

    # 針對特定軟體 (Locale Emulator) 啟動背景安裝器
    le_dir       = os.path.join(_PROGRAMFILES, "Locale Emulator")
    le_installer = os.path.join(le_dir, "LEInstaller.exe")
    if os.path.exists(le_installer):
        subprocess.Popen(
            f'cmd /c start "" /min /D "{le_dir}" "{le_installer}"',
            shell=True
        )
        bg_names.append("Locale Emulator")

    return bg_names