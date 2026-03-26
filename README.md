# ECOCO 客訴智能分析平台

基於 Streamlit 開發的客製化客訴分析系統，支援本機部署與雲端分析功能。

## 主要功能
- **檔案上傳分析：** 支援 Excel/CSV/PDF 讀取，並能透過批次系統清理與修正資料標籤（問題類型、問題細項、負責部門）。
- **圖表化視覺分析：** 自動產生問題類型分佈長條圖、機台類型比例圓餅圖（方舟站 / 電池機 / 收瓶機）及十大客訴細項熱點長條圖。
- **AI 重點摘要：** 結合 OpenAI GPT，不僅分析根因，還能一鍵識別各大客訴熱區與城市站點異常。
- **自動化匯出與歸檔：** 支援下載 Excel / CSV / PDF 格式、產出 PPT 自動化報告簡報，並自動歸檔至歷史紀錄區塊。
- **Google Sheet 連動：** 可匯出結果直接寫入指定的雲端試算表。

## 安裝與執行
1. 請確認已安裝 Python 環境。
2. 執行以下指令安裝依賴套件：
   ```bash
   pip install -r requirements.txt
   ```
3. 啟動本機伺服器：
   ```bash
   streamlit run complaint_webapp.py --server.address 0.0.0.0 --server.port 8501
   ```
4. 在瀏覽器輸入您的 IP (如 `http://localhost:8501`) 開始使用平台。

## 目錄結構
- `complaint_webapp.py` - 主程式進入點與介面邏輯
- `history_reports/` - 儲存本機分析資料的歷史資料夾 (已加入 gitignore)
- `requirements.txt` - Python 套件清單
