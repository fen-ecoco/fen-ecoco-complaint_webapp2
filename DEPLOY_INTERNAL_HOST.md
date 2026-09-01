# 部署到公司內部主機

Render 免費方案已無額度，**正式執行環境改為公司內部主機**。
GitHub 只作為版本控管與派送程式碼之用，不再由 Render 執行。

執行內容分兩塊，可以只裝其中一塊：

| | 用途 | 由誰跑 |
|---|---|---|
| 排程分析 | 每天自動把新客訴分類、產出 Excel 與待複核清單 | Windows 工作排程器 → `scripts\run_analysis.bat` |
| 網頁介面 | 人工複核、看圖表、查歷史 | `streamlit run complaint_webapp.py` |

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

### 常駐（終端機關掉也不會停）

```bash
powershell -ExecutionPolicy Bypass -File scripts\register_webapp_task.ps1
```

會依序嘗試兩種方式，用得成哪個就用哪個：

| 方式 | 觸發時機 | 權限 |
|---|---|---|
| Windows 排程工作 | 登入時 | 需系統管理員 |
| 啟動資料夾捷徑（自動退回） | 登入時 | 不需要任何權限 |

一般使用者跑會落在第二種，訊息會明講用了哪一種。

要「開機即啟動、不必登入」：

```bash
powershell -ExecutionPolicy Bypass -File scripts\register_webapp_task.ps1 -AtStartup
```

這個一定要用**系統管理員身分**執行；若還要在沒人登入時執行，需由管理者在
工作排程器介面補上服務帳號密碼——密碼必須人工輸入，不要寫進腳本。

移除：

```bash
powershell -ExecutionPolicy Bypass -File scripts\register_webapp_task.ps1 -Remove
```

> 排程工作用了 `-ExecutionTimeLimit 0`（不限執行時間）。
> 不加這個的話排程器預設 3 天後會把長駐的服務行程殺掉。

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
