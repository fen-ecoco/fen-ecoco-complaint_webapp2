# 部署到公司內部主機

Render 免費方案已無額度，**正式執行環境改為公司內部主機**。
GitHub 只作為版本控管與派送程式碼之用，不再由 Render 執行。

執行內容分兩塊，可以只裝其中一塊：

| | 用途 | 由誰跑 |
|---|---|---|
| 排程分析 | 每天自動把新客訴分類、產出 Excel 與待複核清單 | Windows 工作排程器 → `scripts\run_analysis.bat` |
| 網頁介面 | 人工複核、看圖表、查歷史 | `streamlit run complaint_webapp.py` |

---

## 0. 目標主機是 Linux（192.168.0.108 的情況）

先確認過連通性，結果如下：

| 檢查 | 結果 |
|---|---|
| Ping | 通 |
| 22 SSH | **通** |
| 445 SMB 檔案共享 | 不通 |
| 3389 遠端桌面 | 不通 |
| 5985 WinRM | 不通 |
| SSH 識別字串 | `SSH-2.0-OpenSSH_9.6p1 Ubuntu-3ubuntu13.18`（Ubuntu 24.04） |

**那台是 Ubuntu，不是 Windows。** 所以下面幾件事在那台都用不了：
Windows 可攜版（第 0a 節）、`.bat`、工作排程器、遠端桌面、`robocopy` 到 `\\IP\D$`。
唯一的通路是 SSH，常駐機制要用 systemd。

### 步驟

在**這台 Windows** 開終端機，把程式碼送過去（`<帳號>` 換成那台的登入帳號）：

```bash
scp -r complaint_webapp.py automation scripts requirements.txt packages.txt <帳號>@192.168.0.108:~/ecoco/
```

那台若本身能連得到 GitHub，直接在那台 clone 更省事：

```bash
ssh <帳號>@192.168.0.108 "git clone https://github.com/fen-ecoco/fen-ecoco-complaint_webapp2.git ~/ecoco"
```

接著登入那台執行部署腳本：

```bash
ssh <帳號>@192.168.0.108
```

```bash
cd ~/ecoco && bash scripts/deploy_linux.sh
```

腳本會做四件事：檢查 Python 3.11+、建 venv 裝套件、裝中文字型
（沒有的話 PDF 與圖表的中文會變成空白方框）、註冊 systemd 服務。

沒有 sudo 權限時：

```bash
bash scripts/deploy_linux.sh --user-service
```

使用者服務在登出後會停止，要一直跑得請管理者執行
`sudo loginctl enable-linger <帳號>`。

### 完成後

```bash
sudo systemctl status ecoco-webapp
```

```bash
sudo journalctl -u ecoco-webapp -f
```

網址 `http://192.168.0.108:8501`。同網段連不上多半是防火牆：

```bash
sudo ufw allow 8501/tcp
```

### 憑證

把 `.streamlit/secrets.toml` 用 `scp` 送過去，或設成環境變數寫進
`~/ecoco/.env`（systemd unit 已經有 `EnvironmentFile=-.env`）：

```
GOOGLE_CREDENTIALS_JSON={"type":"service_account", ...}
HISTORY_SHEET_ID=...
```

驗證：

```bash
~/ecoco/.venv/bin/python -m automation.cli doctor
```

---

## 0a. 目標主機是 Windows 且什麼都沒裝：用可攜版

目標主機沒有 Python、沒有 Git，也不方便裝東西時，**不要在那台裝任何東西**。
在這台（已經跑得起來的機器）打包，複製過去直接執行即可。

```bash
powershell -ExecutionPolicy Bypass -File scripts\make_portable.ps1
```

會在專案外層產生 `ECOCO_可攜版\`（約 800 MB）：

```
ECOCO_可攜版\
  python\          自帶的 Python 與全部套件
  app\             專案程式碼
  啟動.bat         雙擊即啟動網頁介面
  註冊常駐.bat     註冊成排程工作，關掉視窗也不停
  README.txt       使用說明
```

**原理**：這台用的 Python 是可搬移安裝（python-build-standalone 版面配置），
換路徑、換機器都能執行。已實測從完全不同的目錄啟動，
`streamlit` / `pandas` / `gspread` 皆正常，`automation.cli doctor` 也跑得起來，
打包後的 `啟動.bat` 用的是自己那份 Python（不是系統的），服務回應 HTTP 200。

### 複製到目標主機

擇一：

```bash
robocopy "ECOCO_可攜版" "\\<公司主機IP>\D$\ECOCO_可攜版" /E /MT:8
```

或壓成一個檔再用任何方式傳過去：

```bash
powershell -ExecutionPolicy Bypass -File scripts\make_portable.ps1 -Zip
```

到目標主機上雙擊 `啟動.bat` 就會啟動；要常駐再雙擊 `註冊常駐.bat`。

### 憑證

打包預設**不含** `.streamlit\secrets.toml`（Google 服務帳戶金鑰）——
把金鑰複製到另一台機器應該是明確的決定，不該預設發生。需要時：

```bash
powershell -ExecutionPolicy Bypass -File scripts\make_portable.ps1 -IncludeSecrets
```

或事後手動放到 `app\.streamlit\` 底下。
**沒有憑證也能用**「上傳檔案 → 分析 → 下載」這條主線；
只有歷史紀錄與趨勢儀表板需要連 Google Sheets。

### 日後更新

在這台 `git pull` 之後重新打包、重新複製即可。目標主機不需要 Git。

---

## 0b. 遠端桌面部署（目標主機已有 Python 與 Git 時）

沒辦法直接坐在那台機器前面時，用**遠端桌面**是最省事的路：

```bash
mstsc /v:<公司主機IP>
```

帳號密碼在 Windows 自己的登入視窗輸入。連進去之後，開一個 PowerShell
貼下面這一段，第 1～3 節與第 7 節的常駐設定會一次做完：

```powershell
$dest = "D:\ecoco"
New-Item -ItemType Directory -Force -Path $dest | Out-Null
Set-Location $dest
if (Test-Path "$dest\fen-ecoco-complaint_webapp2\.git") {
    Set-Location "$dest\fen-ecoco-complaint_webapp2"; git pull
} else {
    git clone https://github.com/fen-ecoco/fen-ecoco-complaint_webapp2.git
    Set-Location "$dest\fen-ecoco-complaint_webapp2"
}
python -m pip install -r requirements.txt
powershell -ExecutionPolicy Bypass -File scripts\register_webapp_task.ps1
python -m automation.cli doctor
```

最後那行 `doctor` 會列出憑證與試算表設定是否就緒；缺什麼照第 3 節補。

> **為什麼不用 PowerShell 遠端（Invoke-Command）**：需要兩端都啟用 WinRM，
> 非網域環境還要設定 TrustedHosts，這兩件事都要系統管理員權限。
> 若貴公司已有網域與 WinRM，改用遠端指令當然更快。

### 前置需求（那台主機上要先有）

| 項目 | 檢查指令 |
|---|---|
| Git | `git --version` |
| Python 3.11+ | `python --version` |
| 網路可連 GitHub 與 Google Sheets | `python -m automation.cli doctor` |

---

## 1. 取得程式碼

```bash
git clone https://github.com/fen-ecoco/fen-ecoco-complaint_webapp2.git
```

日後更新：

```bash
git pull
```

## 2. 安裝 Python 與套件

Python 3.11 以上（本機已驗證 3.11 與 3.14）。

```bash
pip install -r requirements.txt
```

`requirements.txt` 內的 `anthropic` 只有在要啟用 L2 LLM 分類時才需要；
沒有 API key 時整層自動停用，缺這個套件也不影響排程與介面運作。

中文字型：Windows 內建微軟正黑體即可，圖表不會出現豆腐字。

## 3. 設定憑證

擇一，**不要兩種混用**：

* **建議**：設成 Windows 系統環境變數，磁碟上不留明文憑證。
  需要的變數：`GOOGLE_CREDENTIALS_JSON`（service account 整份 JSON）、
  選配的 `SOURCE_SHEET_ID`、`ANTHROPIC_API_KEY`。
* 或複製 `scripts\env.example.bat` 成 `scripts\env.local.bat` 填入實際值。
  `run_analysis.bat` 執行時會自動載入，該檔已列入 .gitignore。

介面另可用 `.streamlit\secrets.toml`（同樣不進版控）。
`automation/config.py` 會依「環境變數 → Streamlit secrets → 檔案」的順序找憑證，
排程與介面共用同一套解析，設一次兩邊都吃得到。

`HISTORY_SHEET_ID` 留空會用程式內建預設值，通常不必設。

檢查設定是否就緒：

```bash
python -m automation.cli doctor
```

## 4. 決定客訴來源

`run_analysis.bat` 會自動判斷，不必改指令：

| 條件 | 實際執行 |
|---|---|
| `SOURCE_SHEET_ID` 有值 | `automation.cli run --from-sheet --only-new` 讀該試算表 |
| `SOURCE_SHEET_ID` 空白 | `automation.cli watch --dir %ECOCO_INBOX%` 監看資料夾（預設 `inbox`） |

要監看共用資料夾就設 `ECOCO_INBOX`：

```bash
set ECOCO_INBOX=\server\share\客訴收件
```

處理進度記在 `.automation_state.json`，同一個檔案不會重複處理。

> 目前歷史試算表裡沒有可用的客訴來源分頁（`LIST` 分頁整欄是 `#REF!`），
> 所以**預設走資料夾模式**。要改回讀 Google Sheet，請先確認來源試算表
> 第一列是欄位名稱、且含「問題主旨」與「用戶內容」兩欄。

## 5. 註冊排程工作

```bash
powershell -ExecutionPolicy Bypass -File scripts\register_task.ps1 -Time "08:30"
```

移除：

```bash
powershell -ExecutionPolicy Bypass -File scripts\register_task.ps1 -Remove
```

排程設定包含：錯過時間會補跑（StartWhenAvailable）、失敗自動重試 2 次
（間隔 10 分鐘）、執行上限 2 小時。

**執行帳號**：預設註冊成「目前登入的使用者、登入時才執行」。
內部主機若不會長時間保持登入，請由管理者在工作排程器介面把該工作改成
「不論使用者是否登入均執行」並輸入服務帳號密碼——密碼必須由人工輸入，
不要寫進任何腳本或設定檔。

## 6. 驗證

```bash
powershell -Command "Start-ScheduledTask -TaskName 'ECOCO客訴分析'"
```

```bash
powershell -Command "Get-ScheduledTaskInfo -TaskName 'ECOCO客訴分析' | Select-Object LastRunTime,LastTaskResult"
```

`LastTaskResult` 為 0 即成功。執行過程寫在 `logs\analysis_YYYYMMDD.log`。
`run_analysis.bat` 的退出碼：0 完成、2 沒有新資料（批次檔會轉成 0，
排程器不會誤判為失敗）、其他為失敗。

第一次上線建議先設 `HISTORY_READONLY=true` 跑一輪，確認產出正確後再拿掉，
避免測試資料寫進正式歷史試算表。

## 7. 啟動網頁介面

### 手動啟動（終端機）

```bash
scripts\start_webapp.bat
```

會印出實際使用的 Python、本機與內網網址。腳本會自己處理三件事：

* **找得到能用的直譯器** —— 依序試 `ECOCO_PYTHON`、專案內建路徑、`python`、`py`，
  並且是用「跑得出結果」判斷而不是只看退出碼（`cmd.exe` 這種也會回傳 0）。
* **埠被佔用會自動換** —— 從 8501 往上找 20 個埠，換了會明講。
  這是最常見的失敗原因：埠被佔住時 Streamlit 立刻結束，雙擊的視窗一閃就沒了。
* **出錯不會讓視窗消失** —— 任何失敗都會 `pause`，訊息讀得到。

換埠號：`set ECOCO_PORT=8600` 後再執行。`Ctrl+C` 或關掉視窗即停止。

### 常駐（重點：行程不能掛在啟動它的終端機底下）

```bash
powershell -ExecutionPolicy Bypass -File scripts\register_webapp_task.ps1
```

**為什麼一定要用排程器**：直接從終端機（或任何工具工作階段）啟動的 streamlit
行程，會隨著啟動它的 shell 結束而被回收——當下驗證 HTTP 200 都是真的，
過一陣子再開就已經沒了。交給工作排程器之後，行程由排程服務持有，
跟啟動它的視窗無關。

註冊起來的工作設定：

| 設定 | 值 | 為什麼 |
|---|---|---|
| 觸發器 | 每 5 分鐘重複，持續 3650 天 | 一般使用者就能註冊（AtLogOn / AtStartup 需要管理員） |
| MultipleInstances | IgnoreNew | 還活著時後續觸發略過；掛了下次觸發拉回來 |
| ExecutionTimeLimit | 0（不限） | 否則排程器預設 3 天後殺掉長駐行程 |
| 動作環境變數 | ECOCO_NO_PORT_HUNT=1 | 埠已被自己佔用時乾淨結束，不會在 8502、8503… 一路開下去 |

等於一個會自我修復的 keep-alive：服務掛掉最多 5 分鐘內自動回來。

查狀態與手動控制：

```bash
powershell -Command "Get-ScheduledTaskInfo -TaskName 'ECOCO客訴分析網頁'"
```

```bash
powershell -Command "Stop-ScheduledTask -TaskName 'ECOCO客訴分析網頁'"
```

要「開機即啟動、不必登入」（一定要用**系統管理員身分**執行）：

```bash
powershell -ExecutionPolicy Bypass -File scripts\register_webapp_task.ps1 -AtStartup
```

### 讓它在沒人登入時也執行（操作步驟）

`-AtStartup` 只解決「開機就啟動」；若要連**登出後也繼續跑**，必須讓工作以
儲存的帳號密碼執行。密碼由 Windows 自己收，不經過任何腳本或設定檔。

1. 開始功能表搜尋「工作排程器」，以**系統管理員身分**開啟
2. 左側「工作排程器程式庫」找到 **ECOCO客訴分析網頁**
3. 右鍵 →「內容」
4. 「一般」頁籤 → 安全性選項：
   - 勾選 **「不論使用者登入與否均執行」**
   - 勾選 **「不儲存密碼…」請不要勾**（要儲存密碼才能在登出後執行）
   - 需要的話勾選「以最高權限執行」
5. 按「確定」→ Windows 跳出視窗要求輸入該帳號的密碼 → 在**那個視窗**輸入
6. 完成後用下面指令確認 `LogonType` 變成 `Password`：

```bash
powershell -Command "(Get-ScheduledTask -TaskName 'ECOCO客訴分析網頁').Principal | Format-List UserId,LogonType,RunLevel"
```

> 建議用**服務專用帳號**而不是個人帳號。個人帳號日後改密碼，這個工作就會
> 開始失敗（排程器不會自動更新儲存的密碼）。

### 只在上班時段執行（省記憶體）

網頁服務常駐約 260 MB。若希望下班後把它釋放掉：

```bash
powershell -ExecutionPolicy Bypass -File scripts\register_webapp_task.ps1 -Daily -StartTime "08:00" -StopTime "19:00"
```

移除：

```bash
powershell -ExecutionPolicy Bypass -File scripts\register_webapp_task.ps1 -Remove
```

### 更穩的做法：註冊成 Windows 服務

要「當機自動重啟」，用 NSSM 把 `start_webapp.bat` 註冊成服務。
需要另外安裝 nssm，且需管理員權限。

## 8. 備份

`history_reports/` 是知識庫最穩定的來源，納入公司既有備份範圍。
`output/` 與 `logs/` 是產出，可依保存政策定期清理。

## 注意：批次檔請保持純 ASCII

Windows 以 OEM 編碼（中文版是 cp950）解析 `.bat`，
UTF-8 中文註解會被拆成無效指令導致排程失敗。
`run_analysis.bat` 與 `env.example.bat` 因此只用英文註解，
並在開頭 `chcp 65001` 讓日誌裡的中文正常。
