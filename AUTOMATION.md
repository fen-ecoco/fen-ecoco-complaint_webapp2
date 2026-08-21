# 客訴分析自動化說明

原本流程需要人工：上傳 → 選 3 個欄位 → 按開始分析 → **逐列覆核** → 按儲存 → 貼 API Key → 上傳憑證 → 貼 Sheet 網址 → 按產生報告 → 下載。
現在除了「複核系統沒把握的少數幾列」之外，其餘都自動化，並且可以完全無人值守排程執行。

## 架構

```
automation/
  config.py      設定與憑證解析（環境變數 → Streamlit secrets → JSON 檔）
  taxonomy.py    問題分類法（TOPIC_DETAIL_MAP / DEPT_MAP）唯一來源
  text.py        個資遮蔽（電話 / email）與標籤正規化
  rules.py       內建人工關鍵字規則（最高優先的覆寫層）
  knowledge.py   從歷史紀錄自動建立知識庫（指紋快取 / 挖掘規則 / few-shot 檢索池）
  llm.py         L2 LLM 分類與文字生成（Anthropic 優先，退回 OpenAI）
  classifier.py  L0 / L1 / L2 三層瀑布 + 信心分數
  columns.py     欄位自動偵測
  core.py        分析核心（原始表 → 已標記結果 + 稽核欄位）
  pipeline.py    端到端流程（介面與排程共用）
  sheets.py      Google Sheets 讀寫（不依賴 streamlit）
  cli.py         無人值守入口
complaint_webapp.py   Streamlit 介面（只負責顯示與人工複核）
```

`automation/` 刻意不 import streamlit，所以排程模式不需要跑網頁伺服器。

## 分類瀑布

| 層 | 內容 | 成本 | 信心來源 |
|---|---|---|---|
| L0 指紋快取 | 正規化後文字與歷史完全相同 → 沿用當時標記 | 免費 | 0.95–0.99 |
| L1a 內建規則 | 人工維護的關鍵字規則；政策細項（POLICY_DETAILS）以此為準 | 免費 | 歷史實測精確率 |
| L1b 歷史規則 | 從歷史自動挖掘的關鍵字規則（log-odds 區辨詞） | 免費 | 歷史實測精確率 |
| L1c 相似案例 | 歷史標註的 kNN 投票（倒排索引 + IDF，約 1ms/列） | 免費 | 留一法實測準確率 |
| L2 LLM | 選配：只處理前面信心不足的列，帶 few-shot；**沒有 API key 時整層停用** | 每列一次 API（有上限） | 模型自評（夾在 0.3–0.95） |

L1a/L1b/L1c 的信心都是實測準確率，可直接比大小，取最可靠的那一個。
沒有 API key 也能運作，L2 只是額外選項。

## 人工作業分三級（類型與細項分開判斷）

問題類型只有 6 種、細項有 45 種，所以類型的把握度通常遠高於細項。
兩者的信心分開計算（`_confidence` 與 `_topic_confidence`），
就能把人工作業從「複核 / 不複核」的二分法拆成三級：

| 級別 | 條件 | 人工要做的事 | 實測占比 | 類型正確率 |
|---|---|---|---|---|
| A 完全自動採用 | 細項信心足夠且各層不分歧 | 不用看 | 53.6% | 92.3% |
| B 僅需確認細項 | 類型信心足夠、細項沒把握 | 只從該類型的少數細項挑一個；**類型與部門照系統的走** | 7.7% | 86.4% |
| C 需完整判斷 | 類型也沒把握，或各層分歧 | 完整看過 | 38.7% | 64.5% |

部門是由類型決定的，所以 A+B（**61.3%**）的部門派工可以完全自動。
需要完整人工判斷的只剩 38.7%（拆分前是 46.4%）。

待複核清單的排序是「C → B → 稽核抽樣」，同組內最沒把握的優先，
人工從上往下做，最有價值的判斷先完成。
介面另有「✔ 全部接受系統判斷」，可把目前顯示的列一次標記為已複核並存檔
（適合快速掠過沒問題的列），省去逐格點選。

## 分類品質診斷（決定要不要合併細項）

細項定義重疊是準確率的天花板：連人工自己標同一句話都會標到不同格子時，
系統不可能分得出來。診斷指令會量化這件事：

```bash
python -m automation.cli taxonomy-report --cut 2026-07-01
```

輸出（`output/taxonomy_report/`）：各細項可辨識度、最常混淆的配對、
自動建議的合併群組，以及「合併後準確率與自動化率會變成多少」的模擬。
用 `--merge-map 方案.json` 可以改成模擬自己擬的合併方案。

合併建議有三個必要限制，少了任何一個結果就會荒謬：
只合併同一個問題類型底下的細項、不碰 POLICY_DETAILS、
不用連通分量串連（A↔B、B↔C 不代表 A 與 C 該併）。

確認方案後用 `relabel --merge-map 方案.json` 套用到歷史資料，
再把分類法（`automation/taxonomy.py`）改成合併後的細項。

## 自動化的審核機制

自動採用不是「分數夠高就放行」，而是要過三道關：

**1. 交叉驗證** — 各層是互相獨立的判斷依據，會互相對答案：

| 交叉驗證結果 | 處理 | 實測細項正確率 |
|---|---|---|
| 多層一致（兩層以上同一答案） | 信心 +`AGREEMENT_BOOST` | 72.1% |
| 單一依據（只有一層有結果） | 不調整 | 81.8% |
| **各層分歧** | 信心 −`DISAGREEMENT_PENALTY`，**一律進人工** | **32.5%** |

分歧是極強的錯誤訊號：這類的細項正確率只有三成，就算最高分很高也必須讓人看。
各層分別給了什麼答案會記在 `_candidates` 欄，複核時看得到判斷分歧在哪。

**2. 信心門檻** — 低於 `REVIEW_CONFIDENCE_THRESHOLD` 進人工。

**3. 抽樣稽核** — 從**已自動採用**的列隨機抽 `AUDIT_SAMPLE_RATE`（預設 3%）
進人工，標記為「稽核抽樣」。這是唯一能持續量測「自動採用到底準不準」的方法：
人工複核這批的修正率，就是自動判斷的實際錯誤率。抽樣用固定種子，重跑抽到同一批。

待複核清單按信心由低到高排序（介面與排程產出的 Excel 都是），人工時間先花在最可能出錯的列。
排程模式會另外產一份「稽核抽樣.xlsx」，跟真正沒把握的列分開。

加上交叉驗證後（同一份 holdout、同一個門檻 0.75）：

| | 只看信心門檻 | 加上交叉驗證 |
|---|---|---|
| 自動採用 | 54.7% | **53.6%** |
| ↳ 類型正確 | – | **92.3%** |
| ↳ 細項正確 | 81.2% | **81.7%** |
| 進人工的那批若直接採用 | – | 只會對 38.6%（證明攔對了） |

> **看指標時的注意事項**：如果分析的資料本來就在歷史紀錄裡（例如重跑舊檔案），
> L0 指紋會直接命中自己，自動採用率會顯示得非常高（實測可達 96%）。
> 那是自我比對，不是真實效能。要評估效能請用「時間切分」：
> 用較早期間的資料建知識庫，測較晚期間的資料。

實測（5165 筆歷史，1–6 月學、7/1–8/16 測 1049 筆，門檻 0.75）：

| | 只有內建規則 | 加上歷史規則 | 再加相似案例投票 |
|---|---|---|---|
| 全體類型正確 | 53.9% | 73.5% | **81.1%** |
| 全體細項正確 | 37.1% | 50.4% | **61.7%** |
| 落入保底 | 39.0% | 1.7% | **0.0%** |
| 自動採用 | 61.0%（細項僅 57.8%） | 29.1% | **49.0%** |
| ↳ 其中類型正確 | 77.5% | 88.5% | **90.5%** |
| ↳ 其中細項正確 | 57.8% | 81.0% | **81.9%** |
| 需人工複核 | 39.0% | 70.9% | **51.0%** |

門檻可調（`REVIEW_CONFIDENCE_THRESHOLD`）：0.60 → 自動採用 67.8%／細項 74.4%；
0.85 → 自動採用 30.9%／細項 86.4%。要多快還是要多準，由這個值決定。

信心低於 `REVIEW_CONFIDENCE_THRESHOLD`（預設 0.75）→ 標記 `_needs_review`，介面只顯示這些列。
人工修正並儲存後，該列記為「人工確認」，下次重建知識庫時成為黃金標註 → 規則越用越準、待複核越來越少。

分析結果會多出稽核欄位（介面隱藏、儲存時保留）：
`_confidence` 信心、`_source_layer` 判斷來源、`_needs_review` 待複核、`_reason` 判斷依據、`_ai_filled` 是否由系統填入。

## 分類法與政策規則

分類法只在 `automation/taxonomy.py` 定義一份，介面與排程共用。

* `POLICY_DETAILS`：由公司明訂、**必須依客訴內容判斷**的細項。目前是三種滿艙：
  * 提到瓶蓋、蓋子 → `瓶蓋桶已滿`
  * 提到塑膠類容器 → `寶特瓶滿艙`
    （寶特瓶／保特瓶／塑膠瓶／塑膠／塑料／牛奶瓶／鮮奶瓶／優酪乳／養樂多／
    塑膠杯／飲料杯，以及 PET／PVC／PP —— 英文簡寫用邊界比對，
    否則 `pp` 會誤中 `app`）
  * 只說滿了、要清運，沒指明回收物（例如只提鋁罐、鐵罐）→ `回收箱滿艙`
  這些細項的判斷**不會被歷史挖出的規則或指紋快取覆蓋**，因為舊資料在政策訂立前
  把三種滿艙混用同一個標記（`automation/rules.py` 的 `full_bin_detail()`）。
* `RETIRED_TOPICS`：已廢除的類型自動歸位（`APP帳密登入問題` → `APP帳號設定問題類型`）。
* `DETAIL_ALIASES`：更名／異體寫法對應（例：`瓶蓋箱已滿` → `瓶蓋桶已滿`）。
* 只差分隔符的寫法（`機台需維護-故障提醒` 對 `機台需維護/故障提醒`）會自動視為同一項。

調整分類法後，舊歷史資料仍可被知識庫學習，不需要改檔。
但若是**政策改變**（例如三種滿艙改為分開），歷史標記會與新政策衝突，
必須重新標記歷史，否則知識庫會繼續教系統舊行為：

```bash
python -m automation.cli relabel --input history_reports/xxx.csv --out output/relabel_review
```

產出「重新標記後的檔案」與「變更清單」供人工檢視，確認後再取代原檔。

## 環境變數

| 變數 | 用途 | 預設 |
|---|---|---|
| `GOOGLE_CREDENTIALS_JSON` | service account 整份 JSON（字串） | 無 |
| `HISTORY_SHEET_ID` | 歷史紀錄試算表 ID（知識庫的學習來源） | 無 |
| `SOURCE_SHEET_ID` | 排程模式的原始客訴來源試算表 | 無 |
| `SOURCE_WORKSHEET` | 來源工作表名稱（留空取第一個分頁） | 空 |
| `ANTHROPIC_API_KEY` / `OPENAI_API_KEY` | LLM 分類與報告；都沒設就只用規則 + 統計摘要 | 無 |
| `AUTO_ANALYZE_ON_UPLOAD` | 上傳後直接分析，不必按按鈕 | true |
| `AUTO_SAVE_HISTORY` | 分析完自動存雲端歷史 | true |
| `USE_LLM_CLASSIFIER` | 是否啟用 L2 | true（無 key 時自動關閉） |
| `LLM_CLASSIFIER_MODEL` | 分類用模型 | claude-haiku-4-5-20251001 |
| `REVIEW_CONFIDENCE_THRESHOLD` | 待複核門檻 | 0.75 |
| `AGREEMENT_BOOST` | 多層判斷一致時的信心加分 | 0.12 |
| `DISAGREEMENT_PENALTY` | 各層判斷分歧時的信心扣分（分歧一律進人工） | 0.15 |
| `AUDIT_SAMPLE_RATE` | 自動採用中抽多少比例做品質抽驗（0 = 關閉） | 0.03 |
| `LLM_MAX_ROWS` | 單次分析送 LLM 的列數上限（成本天花板） | 400 |
| `LLM_BATCH_SIZE` | 每次 API 呼叫打包幾列 | 20 |
| `KNOWLEDGE_MIN_SUPPORT` | 挖掘規則所需最少歷史筆數 | 3 |
| `LOCAL_HISTORY_DIR` | 本機歷史紀錄資料夾（內部主機的主要學習來源） | history_reports |
| `PREFER_LEARNED_DEPT` | 部門以歷史實際填寫的名稱為準，而非內建 DEPT_MAP | false |

| `HISTORY_READONLY` | 唯讀模式：**所有**歷史寫入都略過（含人工按儲存） | false |

> 在正式歷史試算表上測試或示範時，請設 `HISTORY_READONLY=true`。
> `AUTO_SAVE_HISTORY=false` 只擋自動存檔，人工按下「儲存修改」仍會寫入。

## 無人值守排程

檢查設定是否就緒：

```bash
python -m automation.cli doctor
```

分析單一檔案（產出 Excel、待複核清單、文字報告）：

```bash
python -m automation.cli run --input 客訴清單.xlsx --out output --report
```

讀 Google Sheet 來源，只處理上次之後新增的列：

```bash
python -m automation.cli run --from-sheet --only-new --report --out output
```

監看資料夾內所有還沒處理過的檔案：

```bash
python -m automation.cli watch --dir inbox --out output --report
```

處理進度記在 `.automation_state.json`（已列入 .gitignore）。
退出碼：0 成功、1 失敗、2 沒有資料可處理。

Render 已在 `render.yaml` 加上 cron service（每天 01:00 UTC 執行 `--from-sheet --only-new --report`）。
cron job 需付費方案；若維持免費方案，改用 Windows 工作排程器執行同一道指令即可。

## 個資

送往 LLM 之前一律先過 `automation/text.py` 的遮蔽（手機、市話、email 只留頭尾），
分析結果存檔時也已是遮蔽後的內容。

## 部署到公司內部主機

內部主機有持久化磁碟，所以 `history_reports/` 會成為知識庫最穩定的學習來源
（`build_knowledge()` 會同時讀本機 xlsx 與 Google Sheets）。

1. **環境變數**：不要用 `.streamlit/secrets.toml` 存憑證（那是明文檔）。
   在服務啟動環境設定 `GOOGLE_CREDENTIALS_JSON`、`HISTORY_SHEET_ID`、`ANTHROPIC_API_KEY`。
   Windows 可用系統環境變數，或在啟動批次檔中 `set`。

2. **安裝**

   ```bash
   pip install -r requirements.txt
   ```

3. **啟動網頁介面**（供人工複核與查看報表）

   ```bash
   python -m streamlit run complaint_webapp.py --server.port 8501 --server.address 0.0.0.0
   ```

   要開機自動啟動，把上面這行包成批次檔，用「工作排程器」設為「開機時執行」，
   或用 NSSM 之類的工具註冊成 Windows 服務。

4. **排程分析**（真正的無人值守）：工作排程器新增每日任務，動作為

   ```bash
   python -m automation.cli run --from-sheet --only-new --report --out output
   ```

   起始位置要設成專案目錄（`.automation_state.json` 會記在那裡）。
   若客訴來源是共用資料夾裡的檔案，改用 `watch --dir "\\server\share\客訴收件"`。

5. **上線前先跑一次自我檢查**

   ```bash
   python -m automation.cli doctor
   ```

6. **備份**：`history_reports/` 是知識庫的本機來源，納入公司既有備份範圍。
