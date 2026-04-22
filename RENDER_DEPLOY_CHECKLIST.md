# Render 部署檢查清單

## 1) 必要檔案確認
- `render.yaml` 存在，且 `buildCommand` 使用 `pip install -r requirements.txt`
- `requirements.txt` 包含：
  - `streamlit`
  - `pandas`
  - `plotly`
  - `openpyxl`
  - `python-pptx`
  - `pdfplumber`
  - `openai`
  - `gspread`
  - `google-auth`
  - `matplotlib`
- 主程式檔：`complaint_webapp.py`

## 2) Render 啟動指令
- `startCommand`：
  - `streamlit run complaint_webapp.py --server.port $PORT --server.address 0.0.0.0 --server.headless true --browser.gatherUsageStats false`

## 3) 環境變數（可選）
- 若要啟用 LLM 摘要，可在 Render 設定：
  - `OPENAI_API_KEY`

## 4) 部署後手動驗證（功能二 / 功能三）
1. 進入 `https://ecoco-complaint-analyzer.onrender.com/`
2. 上傳測試檔，完成分析
3. 功能二確認：
   - 可下載 `AI 分析文字檔`
   - 可下載 `圖表圖檔（ZIP）`
   - 可下載 `一鍵下載分析簡報 PPT`
4. 功能三確認：
   - 開啟任一歷史紀錄 → `圖表分析` 分頁
   - 可下載 `一鍵下載PPT`
   - 可下載 `下載圖檔（ZIP）`

## 5) 已完成本機 smoke test
- 功能二 PPT 產出：`tmp_verify/feature2_test.pptx`
- 功能三圖檔壓縮：`tmp_verify/feature3_charts_test.zip`
- 驗證報告：`tmp_verify/verify.json`
