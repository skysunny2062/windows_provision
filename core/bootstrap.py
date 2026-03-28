import importlib.util
import os
import subprocess
import sys
from functools import lru_cache


def _clean_env_for_python():
    env = os.environ.copy()
    env.pop("PYTHONHOME", None)
    env.pop("PYTHONPATH", None)
    env.pop("PIP_TARGET", None)
    env.pop("PIP_PREFIX", None)
    env.pop("PYTHONUSERBASE", None)
    env.pop("VIRTUAL_ENV", None)
    env.pop("PIP_REQUIRE_VIRTUALENV", None)
    env["PYTHONNOUSERSITE"] = "1"
    return env


def _ensure_pip(env):
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "--version"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            env=env,
        )
        return True
    except Exception:
        pass

    print("偵測到 pip 不可用，嘗試自動修復中...")
    try:
        subprocess.check_call([sys.executable, "-m", "ensurepip", "--upgrade"], env=env)
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--upgrade", "pip", "--disable-pip-version-check"],
            env=env,
        )
        return True
    except Exception as e:
        print(f"pip 自動修復失敗: {e}")
        return False


def ensure_modules(mods, max_restart=2):
    missing = [pip for pip, imp in mods if importlib.util.find_spec(imp) is None]
    if not missing:
        return

    restart_count = int(os.environ.get("AUTO_INSTALL_RESTART_COUNT", "0"))
    if restart_count >= max_restart:
        print(f"錯誤：已重啟 {max_restart} 次，套件仍無法載入，請手動安裝：")
        print("  pip install " + " ".join(missing))
        os.system("pause")
        sys.exit(1)

    install_env = _clean_env_for_python()
    if not _ensure_pip(install_env):
        print("請重新安裝 Python（建議透過 winget 安裝正式版 Python）後再執行。")
        os.system("pause")
        sys.exit(1)

    print(f"正在安裝pip: {', '.join(missing)}")
    try:
        subprocess.check_call(
            [
                sys.executable,
                "-m",
                "pip",
                "install",
                *missing,
                "--upgrade",
                "--no-user",
                "--quiet",
                "--disable-pip-version-check",
                "--no-warn-script-location",
                "--no-cache-dir",
            ],
            env=install_env,
        )
    except subprocess.CalledProcessError as e:
        print(f"pip 安裝失敗: {e}")
        print("請手動執行: pip install " + " ".join(missing))
        os.system("pause")
        sys.exit(1)

    print("正在重新啟動腳本...")
    new_env = os.environ.copy()
    new_env["AUTO_INSTALL_RESTART_COUNT"] = str(restart_count + 1)
    os.execve(sys.executable, [sys.executable, *sys.argv], new_env)


def detect_plugins(root_dir):
    plugins = {}
    try:
        for entry in os.listdir(root_dir):
            plugin_dir = os.path.join(root_dir, entry)
            plugin_py = os.path.join(plugin_dir, f"{entry}.py")
            if os.path.isdir(plugin_dir) and os.path.isfile(plugin_py):
                mode_name = f"{entry.capitalize()} Mode"
                try:
                    spec = importlib.util.spec_from_file_location(entry, plugin_py)
                    module = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(module)
                    plugins[mode_name] = (module, plugin_dir)
                except Exception as e:
                    print(f"外掛載入失敗 [{entry}]: {e}")
    except Exception:
        pass
    return plugins


@lru_cache(maxsize=32)
def parse_winget_txt(path):
    pkgs = []
    if not os.path.isfile(path):
        return pkgs
    with open(path, encoding="utf-8") as f:
        for raw in f:
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            parts = [p.strip() for p in line.split(",")]
            pkg_id = parts[0]
            flags = set(parts[1:])
            exact = "exact" in flags
            source = "msstore" if "msstore" in flags else "winget"
            name_kv = next((p for p in flags if p.lower().startswith("name=")), None)
            display_name = name_kv.split("=", 1)[1] if name_kv else None
            pkgs.append((pkg_id, exact, source, display_name))
    return pkgs
