"""ECOCO 客訴分析自動化核心。

這個套件刻意不依賴 streamlit，讓同一套邏輯能同時服務：
  * complaint_webapp.py（互動介面）
  * automation/cli.py（無人值守排程）

模組分工：
  config      設定與憑證（環境變數 / secrets）
  taxonomy    問題分類法（唯一來源）
  text        個資遮蔽與文字正規化
  rules       內建人工關鍵字規則
  knowledge   從歷史紀錄自動建立的知識庫（指紋快取、挖掘規則、few-shot 檢索）
  llm         L2 LLM 分類
  classifier  L0/L1/L2 三層瀑布
  columns     欄位自動偵測
  core        分析核心（原始表 → 已標記結果）
  pipeline    端到端流程
"""

__all__ = [
    "classifier",
    "columns",
    "config",
    "core",
    "knowledge",
    "llm",
    "pipeline",
    "rules",
    "taxonomy",
    "text",
]
