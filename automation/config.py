"""集中解析設定與憑證。

優先順序：環境變數 > Streamlit secrets > 預設值。
所有憑證都不再需要人工在介面貼上或上傳檔案；
無介面排程（automation/cli.py）也走同一套解析，行為一致。
"""

from __future__ import annotations

import json
import os
from typing import Any, Optional


def _from_secrets(key: str) -> Optional[Any]:
    """讀取 Streamlit secrets；在無 Streamlit 的環境（排程 CLI）安靜略過。"""
    try:
        import streamlit as st  # noqa: PLC0415
    except Exception:
        return None
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return None


def get_setting(key: str, default: Any = None) -> Any:
    val = os.environ.get(key)
    if val not in (None, ""):
        return val
    val = _from_secrets(key)
    if val not in (None, ""):
        return val
    return default


def get_bool(key: str, default: bool = False) -> bool:
    raw = get_setting(key)
    if raw is None:
        return default
    return str(raw).strip().lower() in ("1", "true", "yes", "y", "on")


def get_int(key: str, default: int) -> int:
    raw = get_setting(key)
    try:
        return int(str(raw))
    except (TypeError, ValueError):
        return default


def get_float(key: str, default: float) -> float:
    raw = get_setting(key)
    try:
        return float(str(raw))
    except (TypeError, ValueError):
        return default


# ── 憑證 ────────────────────────────────────────────────


def get_google_credentials() -> Optional[dict]:
    """取得 service account JSON（dict）。

    支援四種來源，依序嘗試：
      1. GOOGLE_CREDENTIALS_JSON / GCP_SERVICE_ACCOUNT_JSON（整份 JSON 字串）
      2. GOOGLE_APPLICATION_CREDENTIALS（JSON 檔路徑）
      3. st.secrets["google_credentials"]（table 形式）
      4. 專案目錄下既有的 *.json service account 檔（本機開發用）
    """
    for key in ("GOOGLE_CREDENTIALS_JSON", "GCP_SERVICE_ACCOUNT_JSON", "GOOGLE_CREDENTIALS"):
        raw = os.environ.get(key) or ""
        if raw.strip().startswith("{"):
            try:
                return json.loads(raw)
            except Exception:
                pass

    path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "")
    if path and os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as fh:
                return json.load(fh)
        except Exception:
            pass

    sec = _from_secrets("google_credentials")
    if sec:
        try:
            return dict(sec)
        except Exception:
            pass

    local = get_setting("GOOGLE_CREDENTIALS_FILE")
    if local and os.path.exists(str(local)):
        try:
            with open(str(local), encoding="utf-8") as fh:
                return json.load(fh)
        except Exception:
            pass
    return None


def get_openai_key() -> str:
    return str(get_setting("OPENAI_API_KEY", "") or "")


def get_anthropic_key() -> str:
    return str(get_setting("ANTHROPIC_API_KEY", "") or "")


# 與 complaint_webapp.py 的 DEFAULT_HISTORY_SHEET_ID 相同，
# 讓排程模式在沒有設環境變數時也能連到同一份歷史紀錄。
DEFAULT_HISTORY_SHEET_ID = "1Sqh_8bXtFw7jvmCPufTpStKxfIafDzwYJRlgc0HFBSs"


def get_history_sheet_id() -> str:
    return str(get_setting("HISTORY_SHEET_ID", "") or DEFAULT_HISTORY_SHEET_ID)


def get_source_sheet_id() -> str:
    """原始客訴來源試算表（Phase 4 排程讀取用）。"""
    return str(get_setting("SOURCE_SHEET_ID", "") or "")


def get_local_history_dir() -> str:
    """本機歷史紀錄資料夾（公司內部主機有持久化磁碟時的主要學習來源）。"""
    return str(get_setting("LOCAL_HISTORY_DIR", "history_reports"))


def get_source_worksheet() -> str:
    return str(get_setting("SOURCE_WORKSHEET", "") or "")


# ── 自動化行為開關 ───────────────────────────────────────

def auto_analyze_on_upload() -> bool:
    """上傳後是否直接跑分析，不必按「開始分析」。"""
    return get_bool("AUTO_ANALYZE_ON_UPLOAD", True)


def history_readonly() -> bool:
    """唯讀模式：所有歷史紀錄寫入一律略過（含人工按下儲存的動作）。

    用來在正式的歷史試算表上做測試或示範而不污染資料。
    AUTO_SAVE_HISTORY 只管「自動」存檔，這個開關連明確的存檔動作也一起擋。
    """
    return get_bool("HISTORY_READONLY", False)


def auto_save_history() -> bool:
    """分析完是否自動存入雲端歷史紀錄。"""
    return get_bool("AUTO_SAVE_HISTORY", True)


def prefer_learned_dept() -> bool:
    """部門以「歷史實際填寫的名稱」為準，而非內建 DEPT_MAP。

    公司若已改過部門編制（歷史資料顯示為客服關係部／維運工程部等），
    把這個開關設為 true 即可讓分類直接沿用歷史的部門名稱。
    """
    return get_bool("PREFER_LEARNED_DEPT", False)


def use_llm_classifier() -> bool:
    """是否啟用 L2 LLM 分類（Phase 3）。無 API key 時自動視為關閉。"""
    if not get_bool("USE_LLM_CLASSIFIER", True):
        return False
    return bool(get_anthropic_key() or get_openai_key())


def llm_classifier_model() -> str:
    return str(get_setting("LLM_CLASSIFIER_MODEL", "claude-haiku-4-5-20251001"))


def llm_report_model() -> str:
    return str(get_setting("LLM_REPORT_MODEL", "claude-haiku-4-5-20251001"))


def agreement_boost() -> float:
    """兩個獨立判斷層給出相同答案時，信心可加多少（交叉驗證通過）。"""
    return get_float("AGREEMENT_BOOST", 0.12)


def disagreement_penalty() -> float:
    """兩層答案不一致時，信心要扣多少（把分歧的列推去人工複核）。"""
    return get_float("DISAGREEMENT_PENALTY", 0.15)


def audit_sample_rate() -> float:
    """自動採用的列中，隨機抽多少比例進人工稽核（監控實際準確率）。

    設 0 代表不抽樣。抽樣是唯一能持續驗證「自動採用到底準不準」的方法，
    否則自動化的品質只能靠上線前那一次評估。
    """
    return get_float("AUDIT_SAMPLE_RATE", 0.03)


def review_confidence_threshold() -> float:
    """信心低於此值的列進入待複核佇列。"""
    return get_float("REVIEW_CONFIDENCE_THRESHOLD", 0.75)


def llm_max_rows() -> int:
    """單次分析最多送幾列給 LLM，控制成本上限。"""
    return get_int("LLM_MAX_ROWS", 400)


def llm_batch_size() -> int:
    """每次 LLM 呼叫打包幾列。"""
    return get_int("LLM_BATCH_SIZE", 20)


def knowledge_min_support() -> int:
    """自動挖掘關鍵字規則所需的最少出現次數。"""
    return get_int("KNOWLEDGE_MIN_SUPPORT", 3)
