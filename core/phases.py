import base64
import datetime
import json
import os
import ssl
import stat
import subprocess
import time
import urllib.request
import winreg

INSTALLER_TIMEOUT_SECONDS = 300
POWERSHELL_TIMEOUT_SECONDS = 120
WINDOW_REFRESH_DELAY_SECONDS = 2
GIT_OK_CODES = (0,)


def _spawn_minimized_exe(path, cwd=None):
    cmd = ["cmd", "/c", "start", "", "/min"]
    if cwd:
        cmd.extend(["/D", cwd])
    cmd.append(path)
    subprocess.Popen(cmd)


def _phase_restore(restore_dir, ctx):
    if not restore_dir:
        return
    anydesk_dst = os.path.join(ctx["PROGRAM_DATA"], "AnyDesk")
    for label, eng in ctx["USER_FOLDERS"]:
        src = os.path.join(restore_dir, label)
        dst = os.path.join(ctx["USER_DIR"], eng)
        if os.path.isdir(src):
            cmd = (
                f'start "還原{label}" cmd /k '
                f'"COLOR 0B & robocopy "{src}" "{dst}" /s /mt:8 /r:10 /w:3 /xf desktop.ini"'
            )
            subprocess.Popen(cmd, shell=True)
    anydesk_src = os.path.join(restore_dir, "AnyDesk")
    if os.path.isdir(anydesk_src):
        cmd = (
            f'start "還原AnyDesk" cmd /k '
            f'"COLOR 0B & robocopy "{anydesk_src}" "{anydesk_dst}" /mir /mt:8 /r:10 /w:3"'
        )
        subprocess.Popen(cmd, shell=True)
    print("開始還原資料...")


def _phase_office(ctx):
    office_dir = os.path.join(ctx["PROGRAMFILES"], "Microsoft Office")
    if os.path.isdir(office_dir):
        return None, None
    img_files = [f for f in ctx["safe_listdir"](ctx["SCRIPT_DIR"]) if f.lower().endswith(".img")]
    if not img_files:
        print("Core資料夾內無Office的img 此次將略過Office")
        return None, None
    img_path = os.path.join(ctx["SCRIPT_DIR"], img_files[0])
    r = subprocess.run(
        ["powershell", "-Command", f"(Mount-DiskImage -ImagePath '{img_path}' -PassThru | Get-Volume).DriveLetter"],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    drive = r.stdout.strip()
    if not drive:
        print("Office 掛載失敗")
        return None, None
    setup_exe = f"{drive}:\\Setup.exe"
    if not ctx["wait_until"](lambda: os.path.exists(setup_exe), timeout=5):
        print("Office 找不到 Setup.exe")
        return None, None
    print("安裝 Office...")
    office_proc = subprocess.Popen([setup_exe], cwd=os.path.dirname(setup_exe))
    return img_path, office_proc


def _phase_office_cleanup(img_path, office_proc, ctx):
    if office_proc is not None:
        ctx["run_quiet"](["taskkill", "/f", "/im", "OfficeC2RClient.exe"], check_return=True, label="taskkill OfficeC2RClient")
    if img_path:
        ctx["run_quiet"](["powershell", "-Command", f"Dismount-DiskImage -ImagePath '{img_path}'"], check_return=True, label="Dismount-DiskImage")


def _phase_git(retry_queue, ctx):
    print("安裝 Git...")
    ctx["winget_install_pkg"]("Git.Git", exact=False, source="winget", _retry_queue=retry_queue)
    os.system("COLOR 0B")
    for _ in range(5):
        ctx["reset_git_cache"]()
        if ctx["get_git_exe"]() != "git":
            break
        time.sleep(1)
    ctx["inject_git_path"]()


def _phase_files(plugin_mod, ctx):
    themes_dst = os.path.join(ctx["WINDIR"], "Resources", "Themes")
    ctx["install_fonts"](ctx["CORE_FONTS"])
    if os.path.isdir(ctx["CORE_THEMES"]):
        ctx["xcopy_folder"](ctx["CORE_THEMES"], themes_dst)
    if os.path.isdir(ctx["CORE_WINDIR"]):
        ctx["xcopy_folder"](ctx["CORE_WINDIR"], ctx["WINDIR"])
    if plugin_mod and hasattr(plugin_mod, "custom_files"):
        plugin_mod.custom_files()
    return themes_dst


def _phase_mpv(ctx):
    mpv_dst = os.path.join(ctx["PROGRAMFILES"], "mpv_PlayKit")
    print("安裝 mpv_PlayKit...")
    try:
        git_exe = ctx["get_git_exe"]()
        if os.path.isdir(mpv_dst):
            r_git = subprocess.run([git_exe, "pull"], cwd=mpv_dst)
        else:
            r_git = subprocess.run(
                [git_exe, "clone", "--depth=1", "https://github.com/skysunny2062/mpv_PlayKit.git"],
                cwd=ctx["PROGRAMFILES"],
            )
        if r_git.returncode not in GIT_OK_CODES:
            raise RuntimeError(f"git 失敗 returncode={r_git.returncode}")
        ctx["fix_acl"](mpv_dst)
        mpv_com = os.path.join(mpv_dst, "mpv.com")
        subprocess.run([mpv_com, "--config=no", "--register"])
        mpv_exe = os.path.join(mpv_dst, "mpv.exe")
        lnk_path = os.path.join(ctx["DESKTOP"], "mpv_PlayKit.lnk")
        ps = (
            f'$s=(New-Object -COM WScript.Shell).CreateShortcut("{lnk_path}");'
            f'$s.TargetPath="{mpv_exe}";'
            f'$s.WorkingDirectory="{mpv_dst}";'
            f"$s.Save()"
        )
        ctx["run_quiet"](["powershell", "-Command", ps], check_return=True, label="create mpv shortcut")
        os.system("COLOR 0B")
    except Exception as e:
        print(f"mpv_PlayKit 失敗: {e}")
        ctx["error"]("Installer", "mpv_PlayKit", str(e))


def _install_winget_list(path, retry_queue, ctx):
    for pkg_id, exact, source, display_name in ctx["parse_winget_txt"](path):
        name = display_name or (pkg_id.split(".", 1)[1] if "." in pkg_id else pkg_id)
        print(f"\n安裝 {name}...")
        ctx["winget_install_pkg"](pkg_id, exact=exact, source=source, _retry_queue=retry_queue)
        os.system("COLOR 0B")


def _phase_winget(plugin_dir, retry_queue, ctx):
    _install_winget_list(ctx["CORE_WINGET"], retry_queue, ctx)
    if plugin_dir:
        _install_winget_list(os.path.join(plugin_dir, "winget.txt"), retry_queue, ctx)


def _phase_system(plugin_dir, pc_input, ctx):
    ctx["import_reg_dir"](ctx["CORE_REG_DIR"])
    if plugin_dir:
        ctx["import_reg_dir"](plugin_dir)
    else:
        ctx["import_reg_dir"](ctx["SCRIPT_DIR"])
    ctx["apply_system_settings"]()

    if pc_input:
        hklm = winreg.HKEY_LOCAL_MACHINE
        reg_sz = winreg.REG_SZ
        for key, val in [
            (r"SYSTEM\CurrentControlSet\Services\lanmanserver\Parameters", "srvcomment"),
            (r"SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName", "ComputerName"),
            (r"SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName", "ComputerName"),
            (r"SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", "NV Hostname"),
            (r"System\CurrentControlSet\Services\Tcpip\Parameters", "Hostname"),
        ]:
            ctx["reg_set"](hklm, key, val, reg_sz, pc_input)
    yymmdd = datetime.datetime.now().strftime("%y%m%d")
    ctx["reg_set"](
        winreg.HKEY_LOCAL_MACHINE,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation",
        "Model",
        winreg.REG_SZ,
        f'{ctx["PROGRAM_NAME"]} build:{yymmdd}',
    )


def _phase_vcredist(ctx):
    print("\n安裝 VisualCppRedistAIO...")
    try:
        ssl_ctx = ssl.create_default_context()
        ssl_ctx.check_hostname = False
        ssl_ctx.verify_mode = ssl.CERT_NONE
        req = urllib.request.Request(
            "https://api.github.com/repos/abbodi1406/vcredist/releases/latest",
            headers={"User-Agent": "Mozilla/5.0"},
        )
        with urllib.request.urlopen(req, context=ssl_ctx) as r:
            data = json.loads(r.read())
        for asset in data["assets"]:
            if asset["name"].lower() == "visualcppredist_aio_x86_x64.exe":
                dest = os.path.join(ctx["TEMP"], asset["name"])
                if ctx["download_file"](asset["browser_download_url"], dest):
                    try:
                        subprocess.run([dest, "/y"], timeout=INSTALLER_TIMEOUT_SECONDS)
                    except subprocess.TimeoutExpired:
                        ctx["error"]("Installer", "VisualCppRedistAIO", f"安裝程式 timeout ({INSTALLER_TIMEOUT_SECONDS}s)")
                    os.remove(dest)
                else:
                    ctx["error"]("Installer", "VisualCppRedistAIO", "下載重試耗盡")
                break
    except Exception as e:
        print(f"VisualCppRedistAIO 下載失敗: {e}")
        ctx["error"]("Installer", "VisualCppRedistAIO", str(e))


def _phase_crack(crack, ctx):
    if not crack:
        return
    print("CrackActivation...")
    winrar_key = os.path.join(ctx["PROGRAMFILES"], "WinRAR", "rarreg.key")
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
    encoded = base64.b64encode(ps_script.encode("utf-16-le")).decode()
    try:
        subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded],
            timeout=POWERSHELL_TIMEOUT_SECONDS,
        )
    except subprocess.TimeoutExpired:
        ctx["error"]("Activation", "Windows Activation", f"PowerShell timeout ({POWERSHELL_TIMEOUT_SECONDS}s)")
    except Exception as e:
        ctx["error"]("Activation", "Windows Activation", str(e))


def _phase_theme(themes_dst, ctx):
    theme_files = [f for f in ctx["safe_listdir"](ctx["CORE_THEMES"]) if f.lower().endswith(".theme")]
    if theme_files:
        theme_path = os.path.join(themes_dst, theme_files[0])
        os.startfile(theme_path)
        ctx["wait_window_and_close"]("設定")
    subprocess.run(["taskkill", "/f", "/im", "explorer.exe"])
    iconcache = os.path.join(ctx["LOCAL_APPDATA"], "iconcache.db")
    if os.path.exists(iconcache):
        try:
            ctx["run_quiet"](["attrib", "-s", "-r", "-h", iconcache], check_return=True, label="attrib iconcache")
            os.chmod(iconcache, stat.S_IWRITE)
            os.remove(iconcache)
        except Exception as e:
            print(f"iconcache 刪除失敗: {e}")
    time.sleep(WINDOW_REFRESH_DELAY_SECONDS)
    subprocess.Popen(["explorer.exe"])


def _phase_setup(plugin_mod, ctx):
    bg_names = []
    for f in ctx["safe_listdir"](ctx["CORE_SETUP"]):
        fp = os.path.join(ctx["CORE_SETUP"], f)
        if os.path.isfile(fp) and fp.lower().endswith(".exe"):
            _spawn_minimized_exe(fp)
            bg_names.append(os.path.splitext(f)[0])
    if plugin_mod and hasattr(plugin_mod, "custom_setup"):
        plugin_bg = plugin_mod.custom_setup()
        if plugin_bg:
            bg_names.extend(plugin_bg)
    return bg_names


def run_install(pc_input, install_mode, crack, restore_dir, plugin_mod, plugin_dir, ctx):
    ctx["FAILURES"].clear()
    retry_queue = []
    ts = datetime.datetime.now().strftime("%y%m%d_%H%M%S")
    log_path = os.path.join(ctx["DESKTOP"], f'{ctx["PROGRAM_NAME"]}_Log_{ts}.txt')
    ctx["maximize_console"]()

    _phase_restore(restore_dir, ctx)
    img_path, office_proc = _phase_office(ctx)
    _phase_git(retry_queue, ctx)
    themes_dst = _phase_files(plugin_mod, ctx)
    _phase_mpv(ctx)
    _phase_office_cleanup(img_path, office_proc, ctx)
    _phase_winget(plugin_dir, retry_queue, ctx)
    _phase_system(plugin_dir, pc_input, ctx)
    ctx["set_best_appearance"]()
    ctx["release_topmost"]()
    _phase_vcredist(ctx)
    _phase_crack(crack, ctx)
    _phase_theme(themes_dst, ctx)
    bg_names = _phase_setup(plugin_mod, ctx)

    ctx["final_retry_installers"](retry_queue)
    ctx["write_failure_log"](log_path, install_mode, plugin_dir, bg_names)
    ctx["sc_disable_retry"]("CryptSvc")
    os.system("pause")
