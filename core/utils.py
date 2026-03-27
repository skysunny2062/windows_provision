"""
core/utils.py ─ 共用工具函式

提供 main.py 與其他擴充外掛（如 zack/zack.py 等）匯入使用的通用工具。
本模組不包含任何絕對路徑常數，路徑狀態由呼叫端自行管理與傳遞。
"""

import datetime
import os
import subprocess
from dataclasses import dataclass, field

# ── 全域狀態管理 ──────────────────────────────────────
@dataclass
class Failure:
    """紀錄執行失敗的詳細資訊"""
    ts:       str
    category: str
    label:    str
    detail:   str = ""

@dataclass
class _State:
    """全域狀態容器，負責收集執行過程中的錯誤"""
    failures: list[Failure] = field(default_factory=list)

state = _State()

def _now() -> str:
    """取得當前時間字串 (HH:MM:SS)"""
    return datetime.datetime.now().strftime("%H:%M:%S")

def info(msg: str):
    """印出帶有時間戳記的一般資訊"""
    print(f"[{_now()}] {msg}")

def error(category: str, label: str, detail: str = ""):
    """
    記錄錯誤至全域狀態，並印出錯誤訊息。
    這將用於最終的 Log 報告產生。
    """
    ts = _now()
    state.failures.append(Failure(ts, category, label, detail))
    print(f"[{ts}] FAIL [{category}] {label}{' - ' + detail if detail else ''}")

# ── 子行程執行工具 ────────────────────────────────────
def run_quiet(cmd, *, cwd=None, capture=False, **kwargs):
    """
    靜默執行系統指令，不拋出例外。
    錯誤處理與 Policy 由各個呼叫端自行負責。
    
    :param capture: 若為 True，則將 stdout/stderr 導向 PIPE；否則導向 DEVNULL。
    :return: CompletedProcess 實例，若執行發生例外則回傳 None。
    """
    pipe = subprocess.PIPE if capture else subprocess.DEVNULL
    try:
        return subprocess.run(cmd, stdout=pipe, stderr=pipe,
                              cwd=cwd, **kwargs)
    except Exception as e:
        error("Process", " ".join(map(str, cmd)), str(e))
        return None

# ── 檔案系統與 IO 工具 ────────────────────────────────
def fix_acl(path: str):
    """重設目標路徑的存取控制清單 (ACL)，並賦予 Administrators 群組完整控制權"""
    run_quiet(["icacls", path, "/reset", "/t", "/c"])
    run_quiet(["icacls", path, "/grant:r", "*S-1-5-32-544:(OI)(CI)F", "/t", "/c"])

def robocopy(src: str, dst: str):
    """使用 robocopy 進行多執行緒鏡像複製 (支援重試機制)"""
    r = subprocess.run(["robocopy", src, dst, "/mir", "/mt:8", "/r:10", "/w:3"])
    # robocopy 結束代碼 >= 8 代表發生嚴重錯誤
    if r.returncode >= 8:
        error("IO", f"robocopy {src}", f"returncode={r.returncode}")

def robocopy_folder(src: str, dst: str):
    """執行資料夾鏡像複製，並於複製完成後重設目標資料夾的 ACL 權限"""
    robocopy(src, dst)
    fix_acl(dst)

def xcopy_folder(src: str, dst: str):
    """使用 xcopy 複製資料夾內容，並於失敗時記錄錯誤"""
    r = run_quiet(["xcopy", "/s", "/y", src, dst])
    if r and r.returncode not in (0, 1):
        error("IO", f"xcopy {src}", f"returncode={r.returncode}")

def sync_programfiles(source_dir: str, target_root: str):
    """將來源目錄下的所有子資料夾，逐一透過 robocopy 同步至目標根目錄"""
    if not os.path.isdir(source_dir):
        return
    for folder in os.listdir(source_dir):
        src = os.path.join(source_dir, folder)
        dst = os.path.join(target_root, folder)
        if os.path.isdir(src):
            robocopy_folder(src, dst)

def safe_listdir(path: str) -> list[str]:
    """安全讀取目錄內容，若路徑不存在則回傳空列表，避免拋出 FileNotFoundError"""
    if not os.path.isdir(path):
        return []
    return os.listdir(path)