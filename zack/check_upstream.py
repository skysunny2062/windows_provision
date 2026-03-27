"""
check_upstream.py 
監控上游負責的「所有」設定項目，以及 Zack Provision 所做的系統設定值
用途：上游更新後跑一次，確認這些項目是否仍由上游正確處理
    若出現 ✗ 代表上游不再負責該項，考慮加回自己的 code

監控邏輯：
  【上游環境設定值】上游有做的項目  →  確保上游持續處理
  【Zack 設定值】   Zack Provision 做的系統設定  →  確認部署結果正確
"""
import datetime, os, socket, subprocess, winreg

OUT_DIR  = os.path.join(os.environ["USERPROFILE"], "Desktop")
TS       = datetime.datetime.now().strftime("%y%m%d_%H%M%S")
OUT_FILE = os.path.join(OUT_DIR, f"check_upstream_{TS}.txt")
LINES    = []

def w(*args):
    LINES.append(" ".join(str(a) for a in args))

def section(title):
    w("")
    w("=" * 60)
    w(f"  {title}")
    w("=" * 60)

def run_out(cmd):
    try:
        r = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        for enc in ("utf-8", "cp950", "cp1252"):
            try:
                return r.stdout.decode(enc)
            except Exception:
                continue
        return r.stdout.decode("utf-8", errors="replace")
    except Exception as e:
        return f"ERROR({e})"

def reg_get(hive, key, name):
    try:
        with winreg.OpenKey(hive, key, 0, winreg.KEY_READ) as k:
            val, _ = winreg.QueryValueEx(k, name)
            return val
    except FileNotFoundError:
        return "KEY_NOT_FOUND"
    except Exception as e:
        return f"ERROR({e})"

def reg_key_deleted(hive, key):
    """確認某個 key 已被刪除"""
    try:
        winreg.OpenKey(hive, key)
        return False  # 還存在
    except FileNotFoundError:
        return True   # 已刪除
    except Exception:
        return False

HKLM = winreg.HKEY_LOCAL_MACHINE
HKCU = winreg.HKEY_CURRENT_USER

def show(label, hive, key, name, expect):
    val = reg_get(hive, key, name)
    ok  = str(val) == str(expect)
    mark = "✓" if ok else "✗"
    w(f"  {label:<55} = {str(val):<20} [={expect}] {mark}")

def show_deleted(label, hive, key):
    deleted = reg_key_deleted(hive, key)
    mark = "✓" if deleted else "✗"
    w(f"  {label:<55}   {'DELETED' if deleted else 'EXISTS'} {mark}")

def show_svc(name, expect="disabled"):
    val = svc_start(name)
    val_up = val.upper()
    if expect == "disabled":
        ok = "DISABLED" in val_up or "NOT_INSTALLED" in val_up
    elif expect == "removed":
        ok = "NOT_INSTALLED" in val_up
    elif expect == "auto":
        ok = "AUTO" in val_up
    else:
        ok = False
    mark = "✓" if ok else "✗"
    w(f"  {name:<35} {val:<25} [={expect}] {mark}")

def svc_start(name):
    out = run_out(["sc", "qc", name])
    if "1060" in out or "does not exist" in out.lower() or "找不到" in out:
        return "NOT_INSTALLED"
    for line in out.splitlines():
        if "START_TYPE" in line:
            parts = line.split()
            return f"{parts[2]} {parts[3]}" if len(parts) >= 4 else "UNKNOWN"
    return "UNKNOWN"

def powercfg_standby_ac():
    """從 Registry 直接讀取當前電源計劃的 AC 待機逾時（秒），避免 powercfg 指令在 Administrator 環境讀取失敗"""
    try:
        # 先取得當前啟用的電源計劃 GUID
        with winreg.OpenKey(HKLM,
                r"SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes") as k:
            active_guid, _ = winreg.QueryValueEx(k, "ActivePowerScheme")
        # 待機 AC 設定路徑
        sleep_key = (
            f"SYSTEM\\CurrentControlSet\\Control\\Power\\User\\PowerSchemes\\"
            f"{active_guid}\\238c9fa8-0aad-41ed-83f4-97be242c8f20\\"
            f"29f6c1db-86da-48c5-9fdb-f2b67b1f44da"
        )
        with winreg.OpenKey(HKLM, sleep_key) as k:
            val, _ = winreg.QueryValueEx(k, "ACSettingIndex")
            return int(val)
    except Exception:
        return None

def show_svchost_mitigation():
    """EnableSvchostMitigationPolicy 是 REG_QWORD，值為 0 即可"""
    val = reg_get(HKLM, r"SYSTEM\CurrentControlSet\Control\SCMConfig", "EnableSvchostMitigationPolicy")
    if val == "KEY_NOT_FOUND":
        ok = True   # key 不存在視為已停用
    elif isinstance(val, int):
        ok = val == 0
    else:
        ok = False
    mark = "✓" if ok else "✗"
    w(f"  {'SvcHostMitigationPolicy':<55} = {str(val):<20} [=0] {mark}")

# ════════════════════════════════════════════════════════════
# OEM 環境判斷
def _get_oem_model():
    try:
        with winreg.OpenKey(HKLM,
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation") as k:
            val, _ = winreg.QueryValueEx(k, "Model")
            return str(val)
    except Exception:
        return ""

_oem = _get_oem_model()
if "Zack" in _oem:
    _env_label = f"環境:{_oem}"
else:
    _env_label = "環境:未使用過Zack-Provision"

w(f"check_upstream v9  {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
w(f"PC: {socket.gethostname()}   User: {os.environ.get('USERNAME','')}   {_env_label}")
w("")
w("  【上游環境設定值】監控上游負責的所有設定項目")
w("  【Zack 設定值】  監控 Zack Provision 所做的系統設定")

# ══════════════════════════════════════════════════════════════════
# ▌上游環境設定值
# ══════════════════════════════════════════════════════════════════

# ──────────────────────────────────────────────────────────
section("【上游】服務停用（SetupComplete.cmd）")
# ──────────────────────────────────────────────────────────
for svc in ["DPS","iphlpsvc","ShellHWDetection","stisvc","SysMain",
            "TrkWks","WdiServiceHost","WdiSystemHost","WerSvc"]:
    show_svc(svc, "disabled")

# ──────────────────────────────────────────────────────────
section("【上游】服務額外停用（我從未做過）")
# ──────────────────────────────────────────────────────────
for svc in ["XboxGipSvc","RemoteRegistry","WaaSMedicSvc","DiagTrack",
            "WSearch","CscService","DusmSvc","WpcMonSvc",
            "wisvc","Wecsvc","WebClient","UevAgentService","udfs",
            "tzautoupdate","SNMPTrap","SmsRouter","SharedAccess",
            "SCPolicySvc","PimIndexMaintenanceSvc","PeerDistSvc",
            "NetTcpPortSharing","MsKeyboardFilter","MSDTC","InventorySvc",
            "hvcrash","GraphicsPerfSvc","fhsvc","Fax",
            "DialogBlockingService","diagnosticshub.standardcollector.service",
            "cnghwassist","cdfs","AxInstSV","AppVClient","AppMgmt","ALG",
            "dmwappushservice","dcpsvc","RetailDemo"]:
    show_svc(svc, "disabled")
w("  ※ ws2ifsl / WinHttpAutoProxySvc：Windows 系統保護，上游無法停用，不列入監控")

# ──────────────────────────────────────────────────────────
section("【上游】服務額外停用（SetupComplete.cmd，之前未涵蓋）")
# ──────────────────────────────────────────────────────────
for svc in ["SgrmBroker","SgrmAgent","PolicyAgent","WPDBusEnum",
            "PcaSvc","wscsvc","SENS","shpamsvc","RemoteAccess",
            "UevAgentDriver","ssh-agent"]:
    show_svc(svc, "disabled")
w("  ※ WinHttpAutoProxySvc：Windows 系統保護維持 DEMAND，上游無法停用，不列入監控")

# ──────────────────────────────────────────────────────────
section("【上游】服務移除（NOT_INSTALLED = ✓）")
# ──────────────────────────────────────────────────────────
for svc in ["webthreatdefsvc","webthreatdefusersvc",
            "SecurityHealthService","MsSecFlt","MsSecWfp"]:
    show_svc(svc, "removed")
w("  ※ WdNisSvc：驅動實體已刪但服務鍵殘留（DEMAND），Windows 保護無法移除，屬已知現象，不列入監控")
w("  ※ PlutonHsp2/PlutonHeci/Hsp：ROG 9800X3D 等有 Pluton 晶片的硬體服務會存在（DEMAND），屬硬體差異非問題")
for svc in ["PlutonHsp2","PlutonHeci","Hsp"]:
    show_svc(svc, "removed")

# ──────────────────────────────────────────────────────────
section("【上游】電源（SetupComplete.cmd）")
# ──────────────────────────────────────────────────────────
ac_secs = powercfg_standby_ac()
if ac_secs is None:
    w("  待機 AC 逾時                                               讀取失敗")
else:
    ok = ac_secs == 0
    w(f"  {'待機 AC 逾時':<55} = {ac_secs}秒 [=0] {'✓' if ok else '✗'}")

# ──────────────────────────────────────────────────────────
section("【上游】安全性 UAC / VBS / LSA")
# ──────────────────────────────────────────────────────────
_sys = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
show("UAC EnableLUA",                  HKLM, _sys, "EnableLUA",                    0)
show("UAC ConsentPromptBehaviorAdmin", HKLM, _sys, "ConsentPromptBehaviorAdmin",    0)
show("UAC PromptOnSecureDesktop",      HKLM, _sys, "PromptOnSecureDesktop",         0)
show("UAC EnableInstallerDetection",   HKLM, _sys, "EnableInstallerDetection",      0)
show("UAC EnableCursorSuppression",    HKLM, _sys, "EnableCursorSuppression",       0)
show("UAC FilterAdministratorToken",   HKLM, _sys, "FilterAdministratorToken",      1)
show("UAC LocalAccountTokenFilter",    HKLM, _sys, "LocalAccountTokenFilterPolicy", 1)
show("UAC EnableUIADesktopToggle",     HKLM, _sys, "EnableUIADesktopToggle",        0)
show("UAC EnableSecureUIAPaths",       HKLM, _sys, "EnableSecureUIAPaths",          0)
_dg = r"SYSTEM\CurrentControlSet\Control\DeviceGuard"
show("VBS EnableVBS",                  HKLM, _dg,  "EnableVirtualizationBasedSecurity", 0)
show("VBS HVCI",                       HKLM, _dg + r"\Scenarios\HypervisorEnforcedCodeIntegrity", "Enabled", 0)
show("VBS CredentialGuard",            HKLM, _dg + r"\Scenarios\CredentialGuard",  "Enabled", 0)
show("VBS HypervisorEnforcedCI",       HKLM, _dg,  "HypervisorEnforcedCodeIntegrity", 0)
show("VBS LsaCfgFlags",                HKLM, _dg,  "LsaCfgFlags",                  0)
show("VBS ConfigureSystemGuard",       HKLM, _dg,  "ConfigureSystemGuardLaunch",   2)
show("VBS DeployConfigCIPolicy",       HKLM, _dg,  "DeployConfigCIPolicy",         0)
show("VBS RequirePlatformSecurity",    HKLM, _dg,  "RequirePlatformSecurityFeature", 0)
show("VBS HVCIMATRequired",            HKLM, _dg,  "HVCIMATRequired",              0)
_pm_dg = r"SOFTWARE\Microsoft\PolicyManager\default\DeviceGuard"
_pm_vb = r"SOFTWARE\Microsoft\PolicyManager\default\VirtualizationBasedTechnology"
show("PM DeviceGuard EnableVBS",       HKLM, _pm_dg + r"\EnableVirtualizationBasedSecurity", "value", 0)
show("PM DeviceGuard LsaCfgFlags",     HKLM, _pm_dg + r"\LsaCfgFlags",            "value", 0)
show("PM DeviceGuard ReqPlatformSec",  HKLM, _pm_dg + r"\RequirePlatformSecurityFeatures", "value", 0)
show("PM VBT HVCI",                    HKLM, _pm_vb + r"\HypervisorEnforcedCodeIntegrity", "value", 0)
_lsa = r"SYSTEM\CurrentControlSet\Control\Lsa"
show("LSA RunAsPPL",                   HKLM, _lsa, "RunAsPPL",                     0)
show("LSA RunAsPPLBoot",               HKLM, _lsa, "RunAsPPLBoot",                 0)
show("LSA restrictanonymous",          HKLM, _lsa, "restrictanonymous",             0)
show("LSA LsaConfigFlags",             HKLM, _lsa, "LsaConfigFlags",               0)
show("LSA Policies RunAsPPL",          HKLM, r"SOFTWARE\Policies\Microsoft\Windows\System", "RunAsPPL", 0)
show("VulnDriverBlocklist",            HKLM, r"SYSTEM\CurrentControlSet\Control\CI\Config", "VulnerableDriverBlocklistEnable", 0)

# ──────────────────────────────────────────────────────────
section("【上游】Defender Policy")
# ──────────────────────────────────────────────────────────
_wd  = r"SOFTWARE\Policies\Microsoft\Windows Defender"
_rtp = r"SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection"
_mam = r"SOFTWARE\Policies\Microsoft\Microsoft Antimalware"
_pm  = r"SOFTWARE\Microsoft\PolicyManager\default\Defender"

show("DisableAntiSpyware",             HKLM, _wd,  "DisableAntiSpyware",           1)
show("DisableRoutinelyTakingAction",   HKLM, _wd,  "DisableRoutinelyTakingAction", 1)
show("ServiceKeepAlive",               HKLM, _wd,  "ServiceKeepAlive",             0)
show("PUAProtection",                  HKLM, _wd,  "PUAProtection",                0)
show("DisableLocalAdminMerge",         HKLM, _wd,  "DisableLocalAdminMerge",       1)
show("RTP DisableRealtimeMonitoring",  HKLM, _rtp, "DisableRealtimeMonitoring",    1)
show("RTP DisableBehaviorMonitoring",  HKLM, _rtp, "DisableBehaviorMonitoring",    1)
show("RTP DisableOnAccessProtection",  HKLM, _rtp, "DisableOnAccessProtection",    1)
show("RTP DisableIOAVProtection",      HKLM, _rtp, "DisableIOAVProtection",        1)
show("RTP DisableIntrusionPrevention", HKLM, _rtp, "DisableIntrusionPreventionSystem", 1)
show("RTP DisableRawWriteNotification",HKLM, _rtp, "DisableRawWriteNotification",  1)
show("Scan DisableHeuristics",         HKLM, r"SOFTWARE\Policies\Microsoft\Windows Defender\Scan", "DisableHeuristics", 1)
show("Scan DisableRestorePoint",       HKLM, r"SOFTWARE\Policies\Microsoft\Windows Defender\Scan", "DisableRestorePoint", 1)
show("NIS DisableProtocolRecognition", HKLM, r"SOFTWARE\Policies\Microsoft\Windows Defender\NIS\Consumers\IPS", "DisableProtocolRecognition", 1)
show("ExploitGuard ASR Rules",         HKLM, r"SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR", "ExploitGuard_ASR_Rules", 0)
show("ExploitGuard NetworkProtection", HKLM, r"SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\Network Protection", "EnableNetworkProtection", "KEY_NOT_FOUND")
show("ExploitGuard ControlledFolder",  HKLM, r"SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\Controlled Folder Access", "EnableControlledFolderAccess", 0)
show("Antimalware DisableAntiSpyware", HKLM, _mam, "DisableAntiSpyware",           1)
show("Antimalware DisableAntiVirus",   HKLM, _mam, "DisableAntiVirus",             1)
show("PolicyManager AllowRealtimeMon", HKLM, _pm + r"\AllowRealtimeMonitoring",    "value", 0)
show("PolicyManager AllowBehaviorMon", HKLM, _pm + r"\AllowBehaviorMonitoring",    "value", 0)
show("PolicyManager AllowCloudProt",   HKLM, _pm + r"\AllowCloudProtection",       "value", 0)
show("PolicyManager AllowIOAVProt",    HKLM, _pm + r"\AllowIOAVProtection",        "value", 0)
show("PolicyManager AllowOnAccess",    HKLM, _pm + r"\AllowOnAccessProtection",    "value", 0)
show("PolicyManager EnableNetProt",    HKLM, _pm + r"\EnableNetworkProtection",    "value", 0)

# ──────────────────────────────────────────────────────────
section("【上游】Defender SpyNet / 簽章更新")
# ──────────────────────────────────────────────────────────
_spy = r"SOFTWARE\Policies\Microsoft\Windows Defender\Spynet"
_sig = r"SOFTWARE\Policies\Microsoft\Windows Defender\Signature Updates"
show("SpyNet SpynetReporting",         HKLM, _spy, "SpynetReporting",              0)
show("SpyNet DisableBlockAtFirstSeen", HKLM, _spy, "DisableBlockAtFirstSeen",      1)
show("SpyNet SubmitSamplesConsent",    HKLM, _spy, "SubmitSamplesConsent",         2)
show("Sig DisableScanOnUpdate",        HKLM, _sig, "DisableScanOnUpdate",          1)
show("Sig UpdateOnStartUp",            HKLM, _sig, "UpdateOnStartUp",              0)
show("Sig RealtimeSignatureDelivery",  HKLM, _sig, "RealtimeSignatureDelivery",    0)
show("Sig SignatureDisableNotification",HKLM, _sig, "SignatureDisableNotification", 1)

# ──────────────────────────────────────────────────────────
section("【上游】Defender Tamper / WMI Logger")
# ──────────────────────────────────────────────────────────
# TamperProtection：WinDefend 服務鍵殘留時 Windows 會自動寫回 1，屬已知現象
_tp = reg_get(HKLM, r"SOFTWARE\Microsoft\Windows Defender\Features", "TamperProtection")
_tp_ok = _tp in (0, "KEY_NOT_FOUND")
w(f"  {'TamperProtection':<55} = {str(_tp):<20} [=0/KEY_NOT_FOUND] {'✓' if _tp_ok else '✗'}")
w("  ※ TamperProtection=1：WinDefend 服務鍵殘留時 Windows 自動寫回，驅動已刪無實際效力")

# DefenderApiLogger / DefenderAuditLogger：上游 RemovalofWindowsDefenderAntivirus.reg 整個刪 key
# → 正常部署後 key 不存在（KEY_NOT_FOUND），始祖級機器因舊 code 未執行故 key 仍存在
for _logger, _key in [
    ("DefenderApiLogger Start",  r"SYSTEM\CurrentControlSet\Control\WMI\Autologger\DefenderApiLogger"),
    ("DefenderAuditLogger Start", r"SYSTEM\CurrentControlSet\Control\WMI\Autologger\DefenderAuditLogger"),
]:
    _v = reg_get(HKLM, _key, "Start")
    _ok = _v == "KEY_NOT_FOUND" or _v == 0
    w(f"  {_logger:<55} = {str(_v):<20} [=0/KEY_NOT_FOUND] {'✓' if _ok else '✗'}")

# WdBoot/WdFilter/WdNisDrv/WdNisSvc/WinDefend：服務鍵殘留（Windows 保護無法刪），
# Start 值殘留非 4；驅動實體已刪，無實際效力。期望 key 被刪（KEY_NOT_FOUND）或 Start=4
w("  ※ WdXxx/WinDefend Start：服務鍵 Windows 保護無法刪，驅動實體已刪，Start 殘留值無實際效力")
for _svcname, _key in [
    ("WdBoot Start",    r"SYSTEM\CurrentControlSet\Services\WdBoot"),
    ("WdFilter Start",  r"SYSTEM\CurrentControlSet\Services\WdFilter"),
    ("WdNisDrv Start",  r"SYSTEM\CurrentControlSet\Services\WdNisDrv"),
    ("WdNisSvc Start",  r"SYSTEM\CurrentControlSet\Services\WdNisSvc"),
    ("WinDefend Start", r"SYSTEM\CurrentControlSet\Services\WinDefend"),
]:
    _v = reg_get(HKLM, _key, "Start")
    _ok = _v == "KEY_NOT_FOUND" or _v == 4
    w(f"  {_svcname:<55} = {str(_v):<20} [=4/KEY_NOT_FOUND] {'✓' if _ok else '✗'}")

# ──────────────────────────────────────────────────────────
section("【上游】SmartScreen")
# ──────────────────────────────────────────────────────────
show("SmartScreen Shell",              HKLM, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer", "SmartScreenEnabled", "off")
show("SmartScreen EnableSmartScreen",  HKLM, r"SOFTWARE\Policies\Microsoft\Windows\System", "EnableSmartScreen", 0)
show("SmartScreen PolicyMgr Shell",    HKLM, r"SOFTWARE\Microsoft\PolicyManager\default\SmartScreen\EnableSmartScreenInShell", "value", 0)
show("SmartScreen PolicyMgr AppInst",  HKLM, r"SOFTWARE\Microsoft\PolicyManager\default\SmartScreen\EnableAppInstallControl", "value", 0)
show("SmartScreen AllowSmartScreen",   HKLM, r"SOFTWARE\Microsoft\PolicyManager\default\Browser\AllowSmartScreen", "value", 0)
show("SmartScreen PreventOverride",    HKLM, r"SOFTWARE\Microsoft\PolicyManager\default\SmartScreen\PreventOverrideForFilesInShell", "value", 0)
show("SmartScreen ConfigureAppCtrl",   HKLM, r"SOFTWARE\Policies\Microsoft\Windows Defender\SmartScreen", "ConfigureAppInstallControl", "Anywhere")
w("  ※ SmartScreen Edge/AppHost（HKCU）：Edge 純安裝無優化，key 由 Edge 首次啟動建立，不列入監控")

# ──────────────────────────────────────────────────────────
section("【上游】Spectre/Meltdown + SEHOP + FTH + SvcHost Mitigation")
# ──────────────────────────────────────────────────────────
_mm = r"SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management"
_k  = r"SYSTEM\CurrentControlSet\Control\Session Manager\kernel"
show("Spectre FeatureSettingsOverride",     HKLM, _mm, "FeatureSettingsOverride",     3)
show("Spectre FeatureSettingsOverrideMask", HKLM, _mm, "FeatureSettingsOverrideMask", 3)
show("KernelSEHOP",                         HKLM, _k,  "KernelSEHOPEnabled",          0)
show("FaultTolerantHeap",                   HKLM, r"SOFTWARE\Microsoft\FTH", "Enabled", 0)
show("WindowsMitigation UserPreference",    HKLM, r"SOFTWARE\Microsoft\WindowsMitigation", "UserPreference", 2)
show_svchost_mitigation()

# ──────────────────────────────────────────────────────────
section("【上游】關機加速 HKLM")
# ──────────────────────────────────────────────────────────
_ctrl    = r"SYSTEM\CurrentControlSet\Control"
_ctrl001 = r"SYSTEM\ControlSet001\Control"
show("VerboseStatus",                  HKLM, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "VerboseStatus", 0)
show("ShutdownReasonOn (Reliability)", HKLM, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Reliability",    "ShutdownReasonOn", 0)
show("ShutdownReasonOn (WinNT)",       HKLM, r"Software\Policies\Microsoft\Windows NT\Reliability",       "ShutdownReasonOn", 0)
show("ShutdownWarningDialogTimeout",   HKLM, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows",     "ShutdownWarningDialogTimeout", 1)
show("WaitToKillServiceTimeout (cur)", HKLM, _ctrl,    "WaitToKillServiceTimeout", "1")
show("WaitToKillServiceTimeout (001)", HKLM, _ctrl001, "WaitToKillServiceTimeout", "1")
show("HandlerTimeout (cur)",           HKLM, _ctrl,    "HandlerTimeout",  0x7FFFFFFF)
show("HandlerTimeout (001)",           HKLM, _ctrl001, "HandlerTimeout",  0x7FFFFFFF)
show("ServicesPipeTimeout (cur)",      HKLM, _ctrl,    "ServicesPipeTimeout", 0x240000)
show("ServicesPipeTimeout (001)",      HKLM, _ctrl001, "ServicesPipeTimeout", 0x240000)
show("PollBootPartitionTimeout",       HKLM, r"SYSTEM\CurrentControlSet\Control\PnP", "PollBootPartitionTimeout", 1)

# ──────────────────────────────────────────────────────────
section("【上游】VisualEffects DefaultValue（HKLM）")
# ──────────────────────────────────────────────────────────
_ve = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects"
VE_ITEMS = [
    ("AnimateMinMax",          0),
    ("ComboBoxAnimation",      0),
    ("ControlAnimations",      0),
    ("CursorShadow",           0),
    ("DragFullWindows",        1),
    ("DropShadow",             1),
    ("DWMAeroPeekEnabled",     1),
    ("DWMEnabled",             1),
    ("DWMSaveThumbnailEnabled",0),
    ("FontSmoothing",          1),
    ("ListBoxSmoothScrolling", 0),
    ("ListviewAlphaSelect",    1),
    ("ListviewShadow",         1),
    ("MenuAnimation",          0),
    ("SelectionFade",          0),
    ("TaskbarAnimations",      0),
    ("Themes",                 1),
    ("ThumbnailsOrIcon",       1),
    ("TooltipAnimation",       0),
    ("TransparentGlass",       1),
]
for name, expect in VE_ITEMS:
    show(f"VisualEffects {name}", HKLM, f"{_ve}\\{name}", "DefaultValue", expect)

# ──────────────────────────────────────────────────────────
section("【上游】Defender Security Center 通知關閉")
# ──────────────────────────────────────────────────────────
_wdsc = r"SOFTWARE\Policies\Microsoft\Windows Defender Security Center\Notifications"
show("WDSC DisableNotifications",      HKLM, _wdsc, "DisableNotifications",        1)
show("WDSC DisableEnhancedNotif",      HKLM, _wdsc, "DisableEnhancedNotifications",1)
show("SecCenter AntiVirusOverride",    HKLM, r"SOFTWARE\Microsoft\Security Center", "AntiVirusOverride",  1)
show("SecCenter FirewallOverride",     HKLM, r"SOFTWARE\Microsoft\Security Center", "FirewallOverride",   1)
show("SecCenter FirstRunDisabled",     HKLM, r"SOFTWARE\Microsoft\Security Center", "FirstRunDisabled",   1)

# ──────────────────────────────────────────────────────────
section("【上游】Security Health UI 維護報告停用")
# ──────────────────────────────────────────────────────────
_wsh = r"SOFTWARE\Microsoft\Windows Security Health"
show("SecurityHealth Platform Registered",  HKLM, _wsh + r"\Platform",       "Registered",        0)
show("SecurityHealth UIReportingDisabled",  HKLM, _wsh + r"\Health Advisor",  "UIReportingDisabled",1)
show("SecurityHealth Battery",              HKLM, _wsh + r"\Health Advisor\Battery",        "UIReportingDisabled", 1)
show("SecurityHealth DeviceDriver",         HKLM, _wsh + r"\Health Advisor\Device Driver",  "UIReportingDisabled", 1)
show("SecurityHealth Reliability",          HKLM, _wsh + r"\Health Advisor\Reliability",    "UIReportingDisabled", 1)
show("SecurityHealth StorageHealth",        HKLM, _wsh + r"\Health Advisor\Storage Health", "UIReportingDisabled", 1)

# ──────────────────────────────────────────────────────────
section("【上游】WebThreat 反釣魚移除")
# ──────────────────────────────────────────────────────────
show("WebThreat ServiceEnabled",       HKLM, r"SOFTWARE\Microsoft\PolicyManager\default\WebThreatDefense\ServiceEnabled", "value", 0)
show("WebThreat AuditMode",            HKLM, r"SOFTWARE\Microsoft\PolicyManager\default\WebThreatDefense\AuditMode", "value", 0)
show("WebThreat NotifyUnsafePassword", HKLM, r"SOFTWARE\Microsoft\PolicyManager\default\WebThreatDefense\NotifyUnsafeOrReusedPassword", "value", 0)
show("WebThreat NotifyPasswordReuse",  HKLM, r"SOFTWARE\Policies\Microsoft\Windows\WTDS\Components", "NotifyPasswordReuse", 0)
show("WebThreat NotifyMalicious",      HKLM, r"SOFTWARE\Policies\Microsoft\Windows\WTDS\Components", "NotifyMalicious",     0)

# ──────────────────────────────────────────────────────────
section("【上游】設定頁可見性")
# ──────────────────────────────────────────────────────────
show("SettingsPageVisibility hide:windowsdefender", HKLM,
     r"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer",
     "SettingsPageVisibility", "hide:windowsdefender;")

# ──────────────────────────────────────────────────────────
section("【上游】必要服務保留 AUTO")
# ──────────────────────────────────────────────────────────
for svc in ["EventSystem","wuauserv","BFE","MpsSvc",
            "RpcSs","SamSs","nsi","LanmanWorkstation","lmhosts",
            "hidserv"]:
    show_svc(svc, "auto")
w("  ※ BITS：部分環境為 DEMAND，auto 或 demand 均屬正常")
_bits_val = svc_start("BITS")
_bits_ok  = "AUTO" in _bits_val.upper() or "DEMAND" in _bits_val.upper()
w(f"  {'BITS':<35} {_bits_val:<25} [=auto/demand] {'✓' if _bits_ok else '✗'}")
w("  ※ Spooler/PrintNotify：Zack main.py 停用 Spooler；PrintNotify 不受 Spooler 影響，上游維持設定")
show_svc("Spooler", "disabled")
_pn_val = svc_start("PrintNotify")
_pn_ok  = "DEMAND" in _pn_val.upper() or "AUTO" in _pn_val.upper() or "NOT_INSTALLED" in _pn_val.upper()
w(f"  {'PrintNotify':<35} {_pn_val:<25} [=demand] {'✓' if _pn_ok else '✗'}")
# Browser 是 demand start
_browser_val = svc_start("Browser")
_browser_ok  = "DEMAND" in _browser_val.upper() or "NOT_INSTALLED" in _browser_val.upper()
w(f"  {'Browser':<35} {_browser_val:<25} [=demand] {'✓' if _browser_ok else '✗'}")

# ──────────────────────────────────────────────────────────
section("【上游】Ubpm 維護排程值刪除")
# ──────────────────────────────────────────────────────────
_ubpm0 = r"SYSTEM\ControlSet001\Control\Ubpm"
_ubpm1 = r"SYSTEM\CurrentControlSet\Control\Ubpm"
show("Ubpm DefenderCleanup (001)",      HKLM, _ubpm0, "CriticalMaintenance_DefenderCleanup",     "KEY_NOT_FOUND")
show("Ubpm DefenderVerification (001)", HKLM, _ubpm0, "CriticalMaintenance_DefenderVerification", "KEY_NOT_FOUND")
show("Ubpm DefenderCleanup (cur)",      HKLM, _ubpm1, "CriticalMaintenance_DefenderCleanup",     "KEY_NOT_FOUND")
show("Ubpm DefenderVerification (cur)", HKLM, _ubpm1, "CriticalMaintenance_DefenderVerification", "KEY_NOT_FOUND")

# ──────────────────────────────────────────────────────────
section("【上游】已刪除的 Key")
# ──────────────────────────────────────────────────────────
show_deleted("EPP ContextMenu *",             HKLM, r"SOFTWARE\Classes\*\shellex\ContextMenuHandlers\EPP")
show_deleted("EPP ContextMenu Dir",           HKLM, r"SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\EPP")
show_deleted("WindowsDefender Class",         HKLM, r"SOFTWARE\Classes\WindowsDefender")
show_deleted("WindowsDefender AppUID",        HKLM, r"SOFTWARE\Classes\AppUserModelId\Windows.Defender")
show_deleted("SecurityAndMaintenance CLSID",  HKLM, r"SOFTWARE\Classes\CLSID\{BB64F8A7-BEE7-4E1A-AB8D-7D8273F7FDB6}")
w("  ※ WinDefend/WdBoot/WdFilter/WdNisSvc 服務鍵：Windows 保護機制，上游的 RemoveServices.reg 刪不掉，兩台實測均 EXISTS，屬已知現象，不列入監控")
show_deleted("SecurityHealthService Key",     HKLM, r"SYSTEM\CurrentControlSet\Services\SecurityHealthService")
show_deleted("webthreatdefsvc Key",           HKLM, r"SYSTEM\CurrentControlSet\Services\webthreatdefsvc")
show_deleted("webthreatdefusersvc Key",       HKLM, r"SYSTEM\CurrentControlSet\Services\webthreatdefusersvc")
show_deleted("DefenderApiLogger Autologger",  HKLM, r"SYSTEM\CurrentControlSet\Control\WMI\Autologger\DefenderApiLogger")
show_deleted("DefenderAuditLogger Autologger",HKLM, r"SYSTEM\CurrentControlSet\Control\WMI\Autologger\DefenderAuditLogger")

# ════════════════════════════════════════════════════════════════════
# ▌Zack 設定值
# ════════════════════════════════════════════════════════════════════

# ──────────────────────────────────────────────────────────
section("【Zack】工作列")
# ──────────────────────────────────────────────────────────
_adv  = r"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
_srch = r"Software\Microsoft\Windows\CurrentVersion\Search"
show("SearchboxTaskbarMode（停用搜尋框）",   HKCU, _srch, "SearchboxTaskbarMode",    0)
show("TaskbarAl（工具列置中）",              HKCU, _adv,  "TaskbarAl",               1)
show("TaskbarDa（隱藏天氣小工具）",          HKCU, _adv,  "TaskbarDa",               0)
show("ShowTaskViewButton（隱藏工作檢視）",   HKCU, _adv,  "ShowTaskViewButton",       0)
show("ShowSecondsInSystemClock（時鐘秒數）", HKCU, _adv,  "ShowSecondsInSystemClock", 1)
show("ColorPrevalence（色彩強調工作列）",    HKCU, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "ColorPrevalence", 1)
show("TaskbarGlomLevel（不群組）",           HKCU, _adv,  "TaskbarGlomLevel",         2)

# ──────────────────────────────────────────────────────────
section("【Zack】檔案總管")
# ──────────────────────────────────────────────────────────
_ctx = r"Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"
_ctx_val = reg_get(HKCU, _ctx, "")
_ctx_ok  = _ctx_val == "" or _ctx_val == "KEY_NOT_FOUND"
w(f"  {'Win10 舊式右鍵選單（InprocServer32 預設值）':<55} = {str(_ctx_val):<20} [=] {'✓' if _ctx_ok else '✗'}")

import os as _os
_windir = _os.environ.get("windir", r"C:\Windows")
_shell_icons_key = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Icons"
_expected_icon = f"{_windir}\\system32\\imageres.dll,197"
show("Shell Icons 29（資料夾圖示）", HKLM, _shell_icons_key, "29", _expected_icon)
show("Shell Icons 77（資料夾圖示）", HKLM, _shell_icons_key, "77", _expected_icon)

# ──────────────────────────────────────────────────────────
section("【Zack】電源選單")
# ──────────────────────────────────────────────────────────
_fms = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FlyoutMenuSettings"
show("ShowHibernateOption（顯示休眠）", HKLM, _fms, "ShowHibernateOption", 1)
show("ShowSleepOption（隱藏睡眠）",     HKLM, _fms, "ShowSleepOption",     0)
show("Power ACSettingIndex（僅螢幕）",  HKLM,
     r"SOFTWARE\Policies\Microsoft\Power\PowerSettings\96996BC0-AD50-47EC-923B-6F41874DD9EB",
     "ACSettingIndex", 2)

# ──────────────────────────────────────────────────────────
section("【Zack】電源逾時（powercfg）")
# ──────────────────────────────────────────────────────────
def _powercfg_ac(sub_guid, setting_guid, label, expect):
    try:
        with winreg.OpenKey(HKLM,
                r"SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes") as k:
            active, _ = winreg.QueryValueEx(k, "ActivePowerScheme")
        key = (f"SYSTEM\\CurrentControlSet\\Control\\Power\\User\\PowerSchemes\\"
               f"{active}\\{sub_guid}\\{setting_guid}")
        with winreg.OpenKey(HKLM, key) as k:
            val, _ = winreg.QueryValueEx(k, "ACSettingIndex")
            ok = int(val) == expect
            w(f"  {label:<55} = {val:<20} [={expect}] {'✓' if ok else '✗'}")
    except Exception as e:
        w(f"  {label:<55} 讀取失敗: {e}")

# 螢幕逾時 AC = 900 秒（15 分鐘）
_powercfg_ac("7516b95f-f776-4464-8c53-06167f40cc99",
             "3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e",
             "螢幕逾時 AC（15 分 = 900 秒）", 900)
# 磁碟逾時 AC = 0
_powercfg_ac("0012ee47-9041-4b5d-9b77-535fba8b1442",
             "6738e2c4-e8a5-4a42-b16a-e040e769756e",
             "磁碟逾時 AC（停用 = 0 秒）", 0)

# ──────────────────────────────────────────────────────────
section("【Zack】網路")
# ──────────────────────────────────────────────────────────
_zm = r"Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range1"
show("ZoneMap Range1 *（信任區域）",    HKCU, _zm, "*",       1)
show("ZoneMap Range1 :Range",           HKCU, _zm, ":Range",  "192.168.*.*")
show("DisableUNCCheck（停用 UNC 檢查）",HKCU, r"Software\Microsoft\Command Processor", "DisableUNCCheck", 1)

# ──────────────────────────────────────────────────────────
section("【Zack】相片檢視器")
# ──────────────────────────────────────────────────────────
show("PhotoViewer BackgroundColor", HKCU,
     r"Software\Microsoft\Windows Photo Viewer\Viewer", "BackgroundColor", 4278190080)

# ──────────────────────────────────────────────────────────
section("【Zack】系統穩定性 / 記憶體")
# ──────────────────────────────────────────────────────────
show("Reliability TimeStampInterval（關閉時間戳）", HKLM,
     r"SOFTWARE\Microsoft\Windows\CurrentVersion\Reliability", "TimeStampInterval", 0)
show("DisablePagingExecutive（停用記憶體分頁執行）", HKLM,
     r"SYSTEM\ControlSet001\Control\Session Manager\Memory Management", "DisablePagingExecutive", 1)

# ──────────────────────────────────────────────────────────
section("【Zack】Win11 原生 NVMe 支援（FeatureManagement）")
# ──────────────────────────────────────────────────────────
_fm = r"SYSTEM\CurrentControlSet\Policies\Microsoft\FeatureManagement\Overrides"
show("FeatureManagement 735209102",  HKLM, _fm, "735209102",  1)
show("FeatureManagement 1853569164", HKLM, _fm, "1853569164", 1)
show("FeatureManagement 156965516",  HKLM, _fm, "156965516",  1)

# ──────────────────────────────────────────────────────────
section("【Zack】Office Protected View")
# ──────────────────────────────────────────────────────────
for _app in ["Excel", "Word", "PowerPoint"]:
    _pv = f"Software\\Microsoft\\Office\\16.0\\{_app}\\Security\\ProtectedView"
    show(f"{_app} DisableAttachmentsInPV",     HKCU, _pv, "DisableAttachmentsInPV",     1)
    show(f"{_app} DisableInternetFilesInPV",   HKCU, _pv, "DisableInternetFilesInPV",   1)
    show(f"{_app} DisableUnsafeLocationsInPV", HKCU, _pv, "DisableUnsafeLocationsInPV", 1)

# ──────────────────────────────────────────────────────────
section("【Zack】服務停用")
# ──────────────────────────────────────────────────────────
for svc in ["CscService", "DusmSvc", "Spooler"]:
    show_svc(svc, "disabled")

# ──────────────────────────────────────────────────────────
section("【Zack】系統指令結果")
# ──────────────────────────────────────────────────────────
# 防火牆
def _fw_state():
    out = run_out(["netsh", "advfirewall", "show", "allprofiles", "state"])
    lines = [l.strip() for l in out.splitlines() if "State" in l or "狀態" in l]
    on_count = sum(1 for l in lines if "ON" in l.upper() or "開啟" in l)
    return on_count == 0, out.strip().replace("\n", " | ")

_fw_ok, _fw_detail = _fw_state()
w(f"  {'防火牆（全部關閉）':<55} {'✓' if _fw_ok else '✗  ' + _fw_detail}")

# 磁碟重組排程
_defrag_out = run_out(["schtasks", "/query", "/tn",
                       r"\Microsoft\Windows\Defrag\ScheduledDefrag", "/fo", "LIST"])
_defrag_disabled = "Disabled" in _defrag_out or "停用" in _defrag_out
w(f"  {'磁碟重組排程（Disabled）':<55} {'✓' if _defrag_disabled else '✗'}")

# 系統還原
_sr_out = run_out(["powershell", "-Command",
                   "Get-ComputerRestorePoint -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count"])
_sr_val = reg_get(HKLM,
    r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore", "RPSessionInterval")
_sr_disabled = str(_sr_val) in ("0", "KEY_NOT_FOUND")
w(f"  {'系統還原停用（RPSessionInterval=0）':<55} = {str(_sr_val):<20} [=0/KEY_NOT_FOUND] {'✓' if _sr_disabled else '✗'}")

# OEM Model
_oem_model = reg_get(HKLM,
    r"SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation", "Model")
_oem_ok = isinstance(_oem_model, str) and "Zack Provision" in _oem_model
w(f"  {'OEM Model（含 Zack Provision）':<55} = {str(_oem_model):<20} {'✓' if _oem_ok else '✗'}")

# BCD 開機逾時
_bcd_out = run_out(["bcdedit", "/enum", "{bootmgr}"])
_bcd_timeout_ok = "timeout" in _bcd_out.lower() and any(
    ("timeout" in l.lower() and "0" in l)
    for l in _bcd_out.splitlines()
)
w(f"  {'BCD bootmgr timeout（=0）':<55} {'✓' if _bcd_timeout_ok else '✗'}")

# ════════════════════════════════════════════════════════════════════
w("")
w("=" * 60)
fail_count = sum(1 for l in LINES if "✗" in l)

# 判斷屬於哪個大項
_in_zack = False
_zack_fail = 0
_upstream_fail = 0
for l in LINES:
    if "▌Zack 設定值" in l:
        _in_zack = True
    if "✗" in l:
        if _in_zack:
            _zack_fail += 1
        else:
            _upstream_fail += 1

if fail_count == 0:
    w("  結果：全部正常 ✓")
else:
    parts = []
    if _upstream_fail:
        parts.append(f"上游 {_upstream_fail} 個")
    if _zack_fail:
        parts.append(f"Zack {_zack_fail} 個")
    w(f"  結果：{' / '.join(parts)} 需要確認 ✗（共 {fail_count} 個）")
w("=" * 60)

os.makedirs(OUT_DIR, exist_ok=True)
with open(OUT_FILE, "w", encoding="utf-8") as f:
    f.write("\n".join(LINES) + "\n")

print("\n".join(LINES))
print(f"\n已輸出至：{OUT_FILE}")
input("\n按 Enter 關閉...")