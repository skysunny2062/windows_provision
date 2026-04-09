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

def _decode_output(raw) -> str:
    if raw is None:
        return ""
    if isinstance(raw, bytes):
        for enc in ("utf-8", "cp950", "cp1252"):
            try:
                return raw.decode(enc).strip()
            except Exception:
                continue
        return raw.decode("utf-8", errors="replace").strip()
    return str(raw).strip()

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
def run_quiet(cmd, *, cwd=None, capture=False, check_return=False, ok_codes=(0,), label=None, **kwargs):
    """
    靜默執行系統指令，不拋出例外。
    錯誤處理與 Policy 由各個呼叫端自行負責。
    
    :param capture: 若為 True，則將 stdout/stderr 導向 PIPE；否則導向 DEVNULL。
    :return: CompletedProcess 實例，若執行發生例外則回傳 None。
    """
    need_capture = capture or check_return
    pipe = subprocess.PIPE if need_capture else subprocess.DEVNULL
    try:
        result = subprocess.run(cmd, stdout=pipe, stderr=pipe, cwd=cwd, **kwargs)
        if check_return and result.returncode not in ok_codes:
            stderr_text = _decode_output(result.stderr)
            stdout_text = _decode_output(result.stdout)
            detail = f"returncode={result.returncode}"
            if stderr_text:
                detail += f"; stderr={stderr_text[:240]}"
            elif stdout_text:
                detail += f"; stdout={stdout_text[:240]}"
            error("Process", label or " ".join(map(str, cmd)), detail)
        return result
    except Exception as e:
        error("Process", " ".join(map(str, cmd)), str(e))
        return None

# ── 檔案系統與 IO 工具 ────────────────────────────────
def fix_acl(path: str, timeout: int = 120) -> bool:
    """重設目標路徑的存取控制清單 (ACL)，並賦予 Administrators 群組完整控制權，回傳是否成功"""
    r1 = run_quiet(["icacls", path, "/reset", "/t", "/c"], timeout=timeout)
    r2 = run_quiet(["icacls", path, "/grant:r", "*S-1-5-32-544:(OI)(CI)F", "/t", "/c"], timeout=timeout)
    return r1 is not None and r2 is not None

def robocopy(src: str, dst: str, timeout: int = 600) -> bool:
    """使用 robocopy 進行多執行緒鏡像複製 (支援重試機制)，回傳是否成功"""
    try:
        r = subprocess.run(
            ["robocopy", src, dst, "/mir", "/mt:8", "/r:10", "/w:3"],
            timeout=timeout
        )
        # 由使用者手動監控 robocopy 視窗，不再將 returncode 寫入 FAILURES。
        # robocopy 結束代碼 >= 8 仍視為失敗，僅回傳 False。
        if r.returncode >= 8:
            return False
        return True
    except subprocess.TimeoutExpired:
        error("IO", f"robocopy {src}", f"timeout after {timeout}s")
        return False
    except Exception as e:
        error("IO", f"robocopy {src}", f"unexpected error: {str(e)}")
        return False

def robocopy_folder(src: str, dst: str) -> bool:
    """執行資料夾鏡像複製，並於複製完成後重設目標資料夾的 ACL 權限，回傳是否成功"""
    ok = robocopy(src, dst)
    fix_acl(dst)
    return ok

def xcopy_folder(src: str, dst: str) -> bool:
    """使用 xcopy 複製資料夾內容，並於失敗時記錄錯誤，回傳是否成功"""
    r = run_quiet(["xcopy", "/s", "/y", src, dst])
    if r and r.returncode not in (0, 1):
        error("IO", f"xcopy {src}", f"returncode={r.returncode}")
        return False
    return r is not None

def sync_programfiles(source_dir: str, target_root: str) -> bool:
    """將來源目錄下的所有子資料夾，逐一透過 robocopy 同步至目標根目錄"""
    if not os.path.isdir(source_dir):
        return True  # 路徑不存在視為「無需執行=成功」
    all_ok = True
    for folder in os.listdir(source_dir):
        src = os.path.join(source_dir, folder)
        dst = os.path.join(target_root, folder)
        if os.path.isdir(src):
            if not robocopy_folder(src, dst):
                all_ok = False
    return all_ok

def safe_listdir(path: str) -> list[str]:
    """安全讀取目錄內容，若路徑不存在或無權限則回傳空列表"""
    try:
        if not os.path.isdir(path):
            return []
        return os.listdir(path)
    except PermissionError:
        error("IO", f"safe_listdir {path}", "PermissionError")
        return []
