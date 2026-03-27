# windows_provision
Windows 自動化佈署工具，以雙擊 `.bat` 的方式完成全機設定。

## 免責聲明
- 本 README 由 AI（Claude）依Code原始碼自動生成，**不保證內容完全正確**。
- 本專案包含之系統調校與測試模組，僅限個人環境建置之學習與研究用途。
- 如有疑義，請以實際程式碼為準。

## 環境需求

- 本工具設計運行於 **CandyNext 24H2 Build 26100.1000** 客製化 Windows 映像之上
- 上游（CandyNext）已預先處理大量服務停用、系統清理等底層設定
- windows_provision在此基礎上疊加個人化設定與軟體安裝
- 若於其他 Windows 映像或版本上執行，部分功能可能無法正常運作或產生非預期結果。

## 功能概覽

- **winget 套件安裝**：自動安裝 `winget.txt` 中列出的應用程式，支援精確匹配與 Microsoft Store 來源。失敗時自動重試最多 3 次，仍失敗者進入 final retry 佇列，於全部流程結束後再嘗試一次。
- **字型安裝**：掃描 `core/font/` 並安裝至 `C:\Windows\Fonts`，同步寫入登錄檔；目的端已存在的字型自動略過。
- **登錄檔匯入**：套用 `core/reg/*.reg`，以及（選擇外掛模式時）外掛目錄下的 `*.reg`。
- **系統設定**：工作列、效能外觀、電源計畫（螢幕 15 分鐘、磁碟永不休眠）、BCD 開機倒數歸零、Office Protected View、防火牆關閉、排程工作停用等，一鍵套用。
- **主題套用**：複製 `core/themes/` 下的第一個 `.theme` 到系統主題目錄並套用，套用後刪除 `iconcache.db` 並重啟 `explorer.exe`。
- **檔案同步**：透過 `robocopy` / `xcopy` 部署字型、主題、`core/windir/` 至 `C:\Windows`，以及外掛指定的其他目錄。
- **Git 安裝與注入**：透過 winget 安裝 Git，安裝後自動從登錄檔讀取路徑並注入 `PATH`，供後續流程使用。
- **mpv_PlayKit**：從 GitHub（`skysunny2062/mpv_PlayKit`）以 `git clone --depth=1` 下載，或已存在則 `git pull` 更新，部署至 `%ProgramFiles%\mpv_PlayKit`，並在桌面建立捷徑。
- **Microsoft Office**：偵測 `core/` 下的 `.img` 檔，以 PowerShell `Mount-DiskImage` 掛載後執行 `Setup.exe`；若 `%ProgramFiles%\Microsoft Office` 已存在則跳過。安裝完成後自動卸載映像。
- **VisualCppRedistAIO**：從 GitHub Releases API 取得最新版下載連結，下載後靜默安裝。
- **資料還原**：自動偵測 `D:\backup_<電腦名稱>` 資料夾，或手動輸入路徑；在安裝流程一開始即以多個背景視窗平行 `robocopy` 還原各使用者資料夾與 AnyDesk 設定。
- **資料備份**：獨立於安裝流程的互動式選單，可備份桌面、下載、文件、圖片、音樂、影片、AnyDesk 設定至指定目錄（預設 `D:\backup_<電腦名稱>`）。
- **失敗日誌**：安裝結束後自動於桌面產生 `<程式名稱>_Log_<時間戳>.txt`，分兩層記錄：
  - **verifyFAILURES**（最終確認失敗）：以 `winget list` / `sc qc` 反向比對，確認安裝或服務停用實際未生效者。
  - **Runtime Warnings**：執行期間捕捉到的例外，retry 後可能已成功，僅供參考。
- **外掛（Plugin）系統**：偵測根目錄下的子資料夾，若 `<名稱>/<名稱>.py` 存在即載入為外掛。**目前僅支援同時啟用一個外掛**（程式碼取第一個偵測到的外掛）。

---

## 目錄結構

```
<根目錄>\
│  windows_provision.bat     ← 入口，以系統管理員身份執行
│
├─ core\
│   ├─ main.py               ← 主程式
│   ├─ utils.py              ← 共用工具函式
│   ├─ winget.txt            ← （選擇性）核心 winget 套件清單
│   ├─ *.reg                 ← （選擇性）.reg 登錄檔(only Normal Mode)
│   ├─ *.img                 ← （選擇性）Office 安裝映像
│   ├─ font\                 ← （選擇性）字型檔（.ttf / .otf / .ttc / .fon）
│   ├─ themes\               ← （選擇性）Windows 主題檔（.theme）
│   ├─ reg\                  ← （選擇性）共用 .reg 登錄檔（必套用）
│   ├─ setup\                ← （選擇性）安裝檔（.exe，背景執行）
│   └─ windir\               ← （選擇性）同步至 C:\Windows 的內容
│
└─ zack\                     ← 外掛範例（Zack Mode）
    ├─ zack.py
    ├─ winget.txt
    ├─ *.reg
    ├─ appdata\
    ├─ c\
    ├─ programfiles\
    ├─ desktop\
    └─ setup\
        └─ Install4j\
```

---

## 快速開始

1. 以**系統管理員**身份執行 `windows_provision.bat`。
2. 若未安裝 Python，腳本會自動透過 winget 安裝 **Python 3.14**，安裝完成後自動重啟（最多重啟 2 次）。
3. 主選單出現後選擇模式：
   - **1. 系統部署**：進入安裝問卷
   - **2. 資料備份**：備份使用者資料夾
   - **3. 結束**

> **注意**：腳本需要管理員權限，未具權限時會自動透過 UAC 提權重啟。

---

## winget.txt 格式

```
# 這是註解
pkg_id[,exact][,msstore][,name=顯示名稱]
```

| 旗標 | 說明 |
|------|------|
| `exact` | 加上 `-e` 旗標，精確比對套件 ID |
| `msstore` | 從 Microsoft Store 安裝（預設為 winget 來源） |
| `name=xxx` | 安裝時顯示的名稱（不影響安裝指令） |

**範例（`core/winget.txt`）：**
```
XPFCC4CD725961,msstore,name=LINE
CrystalDewWorld.CrystalDiskInfo.AoiEdition,exact
Daum.PotPlayer
Discord.Discord
```

---

## 外掛系統（Plugin）

在根目錄下建立 `<名稱>/` 資料夾並放入同名 `.py` 檔，主程式即自動載入為「`<名稱> Mode`」。

> **限制**：目前程式僅取第一個偵測到的外掛，多個外掛目錄並存時，非第一個者不會被載入。

### 外掛 API

| 函式 | 呼叫時機 | 說明 |
|------|----------|------|
| `custom_files()` | 檔案複製階段結尾（api1） | 部署外掛專屬的 AppData、C 槽、Program Files、桌面檔案 |
| `custom_setup()` | 安裝程式階段結尾（api3） | 執行外掛專屬的安裝程式；需回傳背景執行的程式名稱清單（`list[str]`） |

> **winget**：外掛的 `winget.txt` 由主程式直接掃描（api2），不需在 `.py` 中撰寫程式碼。  
> **登錄檔**：外掛目錄下的 `*.reg` 亦由主程式直接匯入，不需在 `.py` 中撰寫程式碼。  
> **一般 `.reg`**（非 `core/reg/`）：Normal Mode 時匯入 `core/*.reg`；外掛模式時改匯入外掛目錄下的 `*.reg`（兩者互斥，`core/reg/` 則無論何種模式都會套用）。

### Zack Mode 範例

`zack/zack.py` 實作了完整的外掛範例，展示如何：

- 透過 `custom_files()` 複製 `appdata/`、`c/`、`programfiles/`、`desktop/` 到對應系統目錄。
- 透過 `custom_setup()` 執行 `setup/Install4j/` 下的安裝檔（同步等待）、`setup/` 根目錄 `.exe`（背景），以及觸發 Locale Emulator 安裝。

## 共用工具（`core/utils.py`）

外掛或其他腳本可直接 import 使用：

| 函式 | 說明 |
|------|------|
| `info(msg)` | 輸出帶時間戳的資訊訊息 |
| `error(category, label, detail)` | 記錄失敗並輸出至終端（寫入全域 `_FAILURES` 清單） |
| `run_silent(cmd, ...)` | 靜默執行子程序，可選擇是否擷取輸出 |
| `robocopy_folder(src, dst)` | robocopy /mir + ACL 重設（icacls） |
| `xcopy_folder(src, dst)` | xcopy /s /y，失敗時記錄 error |
| `sync_programfiles(source_dir, target_root)` | 批次 robocopy 來源目錄下的各子資料夾 |
| `safe_listdir(path)` | 路徑不存在時回傳空列表，不拋例外 |

---

## 系統需求

- **上游環境**：CandyNext 24H2（Windows 11 24H2，Build 26100.1000）
- 系統管理員權限
- 網路連線（winget 安裝、mpv_PlayKit git clone、VisualCppRedistAIO 下載等需要）
- Python 3.x（未安裝時腳本自動透過 winget 安裝 Python 3.14）
- 部分 UI 自動化操作依賴中文視窗標題（如「效能選項」「設定」），非中文環境可能失效。