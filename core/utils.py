"""
core/utils.py  ─ 共用工具函式

main.py 與外掛（zack/zack.py 等）都可 import 使用。
不包含任何路徑常數，呼叫端自行管理路徑。
"""

import datetime
import os
import subprocess

# ── Log（全域狀態由 main.py 持有，這裡只提供寫入介面）──
_FAILURES: list = []        # main.py 啟動時會以 main.FAILURES 取代參照
_LOG_PATH: list = [None]    # 用 list 包裝，方便外部修改同一物件


def _now():
    return datetime.datetime.now().strftime("%H:%M:%S")


def info(msg: str):
    print(f"[{_now()}] {msg}")


def error(category: str, label: str, detail: str = ""):
    ts = _now()
    _FAILURES.append((ts, category, label, detail))
    print(f"[{ts}] FAIL [{category}] {label}{' - ' + detail if detail else ''}")


# ── 檔案 IO ───────────────────────────────────────────
def run_silent(cmd, *, cwd=None, check=False, capture=False, **kwargs):
    pipe = subprocess.PIPE if capture else subprocess.DEVNULL
    try:
        return subprocess.run(cmd, stdout=pipe, stderr=pipe,
                              cwd=cwd, check=check, **kwargs)
    except subprocess.CalledProcessError as e:
        error("Process", " ".join(map(str, cmd)), f"returncode={e.returncode}")
    except Exception as e:
        error("Process", " ".join(map(str, cmd)), str(e))


def _fix_acl(path: str):
    run_silent(["icacls", path, "/reset", "/t", "/c"])
    run_silent(["icacls", path, "/grant:r", "*S-1-5-32-544:(OI)(CI)F", "/t", "/c"])


def _robocopy(src: str, dst: str):
    subprocess.run(["robocopy", src, dst, "/mir", "/mt:8", "/r:10", "/w:3"])


def robocopy_folder(src: str, dst: str):
    """robocopy /mir + ACL 重設"""
    _robocopy(src, dst)
    _fix_acl(dst)


def xcopy_folder(src: str, dst: str):
    """xcopy /s /y，失敗時記錄 error"""
    r = run_silent(["xcopy", "/s", "/y", src, dst])
    if r and r.returncode not in (0, 1):
        error("IO", f"xcopy {src}", f"returncode={r.returncode}")


def sync_programfiles(source_dir: str, target_root: str):
    """將 source_dir 下的各子資料夾 robocopy 到 target_root"""
    if not os.path.isdir(source_dir):
        return
    for folder in os.listdir(source_dir):
        src = os.path.join(source_dir, folder)
        dst = os.path.join(target_root, folder)
        if os.path.isdir(src):
            robocopy_folder(src, dst)


def safe_listdir(path: str) -> list:
    """路徑不存在時回傳空列表，不拋例外"""
    if not os.path.isdir(path):
        return []
    return os.listdir(path)