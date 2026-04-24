import io
import json
import re
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

import matplotlib.pyplot as plt
import pandas as pd
import plotly.express as px
import streamlit as st
from pptx.dml.color import RGBColor
from pptx import Presentation
from pptx.util import Inches, Pt

try:
    import pdfplumber
except Exception:
    pdfplumber = None

try:
    import gspread
    from google.oauth2.service_account import Credentials
except Exception:
    gspread = None
    Credentials = None

try:
    from openai import OpenAI
except Exception:
    OpenAI = None


st.set_page_config(page_title="ECOCO 客訴分析平台", page_icon="📊", layout="wide")

TOPIC_DETAIL_MAP = {
    "APP使用問題類型": [
        "APP畫面顯示與機台狀態不符",
        "APP商家頁面空白",
        "APP點數顯示異常",
        "APP多重異常狀況",
        "app畫面顯示與機台狀態不符",
        "app多重異常狀況",
        "app點數顯示異常",
        "app商家頁面空白",
    ],
    "APP帳號設定問題類型": [
        "忘記密碼/無法重設密碼",
        "帳號資訊修改/設定",
        "無法接收簡訊驗證碼",
        "APP無法登入",
        "app無法登入",
    ],
    "APP帳密登入問題": [
        "APP無法登入",
        "app無法登入",
        "忘記密碼/無法重設密碼",
    ],
    "優惠券問題類型": [
        "兌換失敗/顯示錯誤",
        "無法進行兌換操作",
        "使用規則/限制條件說明",
        "查詢優惠券序號紀錄",
    ],
    "回收點數問題類型": [
        "點數重複入點",
        "點數未入帳號",
        "投入後未獲點數/點數未記錄",
    ],
    "機台問題類型": [
        "機台運作中斷/重啟",
        "黑色分選門異常或卡瓶堵塞",
        "重量偵測異常",
        "操作流程異常/無法正常操作",
        "螢幕異常顯示/畫面異常",
        "履帶未作動或異常抖動",
        "機台當機/無回應",
        "機台需維護/故障提醒",
        "機台網路連線失敗",
        "機台髒污/需要清潔",
        "網路中斷或不穩定",
        "機台關閉/無法啟動",
        "投口綠燈拒收容器",
        "投入物卡住_瓶罐/電池",
        "辨識失敗異常或錯誤",
        "機台操作畫面無法登入",
        "投入後未獲點數/點數未記錄",
        "螢幕西曬導致黑屏或反光",
        "瓶蓋桶已滿",
        "回收艙門開啟",
    ],
    "顧客關係類型": [
        "許願新增站點/設站建議",
        "申請刪除帳號",
        "更換帳號",
        "其他建議",
        "回收物使用規則",
        "相關活動規則疑問",
    ],
}

TYPE_OPTIONS = list(TOPIC_DETAIL_MAP.keys())
DETAIL_OPTIONS = [d for lst in TOPIC_DETAIL_MAP.values() for d in lst]

DEPT_OPTIONS = [
    "營運部", "研發部", "廠務部", "人資部", "行銷部", 
    "資訊部", "企劃部", "財務部", "開發部", "總經理室"
]

DEPT_MAP = {
    "機台問題類型": "營運部",
    "機台相關問題": "營運部",
    "APP帳號設定問題類型": "資訊部",
    "APP使用問題類型": "資訊部",
    "APP帳密登入問題": "資訊部",
    "回收點數問題類型": "",
    "優惠券問題類型": "行銷部",
    "顧客關係類型": "營運部",
}

HISTORY_DIR = Path("history_reports")
HISTORY_DIR.mkdir(exist_ok=True)
META_FILE = HISTORY_DIR / "history.json"


@dataclass
class AnalysisConfig:
    subject_col: str
    content_col: str
    date_col: Optional[str]


def apply_brand_theme() -> None:
    st.markdown(
        """
        <style>
          html, body, [data-testid="stAppViewContainer"] {
            font-size: 18px !important;
          }
          @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@500;700;900&display=swap');
          
          /* Noto Sans TC Medium (500) - scoped to app content only, not Streamlit portals */
          [data-testid="stAppViewContainer"] *:not(.stIconMaterial):not(.material-symbols-rounded):not([data-testid="stIconMaterial"]),
          [data-testid="stHeader"] *:not(.stIconMaterial):not(.material-symbols-rounded):not([data-testid="stIconMaterial"]),
          [data-testid="stMain"] *:not(.stIconMaterial):not(.material-symbols-rounded):not([data-testid="stIconMaterial"]) {
            font-family: 'Noto Sans TC', 'Microsoft JhengHei', sans-serif !important;
          }
          [data-testid="stAppViewContainer"] p,
          [data-testid="stAppViewContainer"] span,
          [data-testid="stAppViewContainer"] label,
          [data-testid="stAppViewContainer"] div {
            font-weight: 500;
            font-size: 18px !important;
          }
          
          /* Use Noto Sans TC Medium (500) for everything — no bold allowed */
          h1, h2, h3, h4, h5, h6, .ecoco-banner, strong, b, .side-title, section[data-testid="stSidebar"] .stButton > button {
            font-family: 'Noto Sans TC', 'Microsoft JhengHei', sans-serif !important;
            font-weight: 500 !important;
          }

          :root{
            --ecoco-orange:#FF5000;
            --ecoco-blue:#060E9F;
            --ecoco-yellow:#FFCE00;
            --ecoco-lightblue:#8EB9C9;
            --ecoco-beige:#FAE0B8;
            --ecoco-deepteal:#0076A9;
          }
          .stApp {background: linear-gradient(135deg, #fff 0%, #f8fbff 40%, #fff8f1 100%);}
          .ecoco-banner {
            padding: 14px 18px; border-radius: 12px;
            background: linear-gradient(90deg, var(--ecoco-orange), var(--ecoco-blue));
            color:white; font-weight:500; margin-bottom: 12px;
            font-size: 20px !important;
          }
          .ecoco-card{
            border:1px solid #e7e7e7; border-left:6px solid var(--ecoco-orange);
            border-radius:12px; padding:10px 14px; background:white; margin-bottom:10px;
            color: #555555 !important;
          }
          [data-testid="stAppViewContainer"] .ecoco-card,
          [data-testid="stAppViewContainer"] .ecoco-card * {
            font-size: 16px !important;
          }
          .ecoco-card b {
            color: #333333 !important;
          }
          .small-muted { color:#666 !important; font-size: 0.9rem; }
          
          /* Sidebar background */
          section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #0b3f78 0%, #083668 100%);
          }
          
          /* Sidebar Text Overrides */
          .side-title {
            color: #ffffff !important;
            font-weight: 500; font-size: 1.05rem; margin-bottom: 8px;
          }
          .side-sub {
            color: #ffffff !important;
            font-size: 0.78rem; opacity: 0.85; margin-bottom: 14px;
          }
          
          /* Sidebar Buttons — default = lightblue */
          section[data-testid="stSidebar"] .stButton > button {
            background-color: var(--ecoco-lightblue) !important;
            border-color: var(--ecoco-lightblue) !important;
            color: #333333 !important;
            border-radius: 12px;
            min-height: 46px;
            font-weight: 500;
            text-align: left;
            transition: background-color 0.12s ease, border-color 0.12s ease !important;
          }
          section[data-testid="stSidebar"] .stButton > button * {
            color: #333333 !important;
          }
          /* Hover = white immediately */
          section[data-testid="stSidebar"] .stButton > button:hover,
          section[data-testid="stSidebar"] .stButton > button:focus,
          section[data-testid="stSidebar"] .stButton > button:active,
          section[data-testid="stSidebar"] .stButton > button[kind="primary"],
          section[data-testid="stSidebar"] .stButton > button[data-testid="baseButton-primary"] {
            background-color: #FFFFFF !important;
            border-color: #FFFFFF !important;
            color: #333333 !important;
          }
          
          /* Thicker scrollbar */
          ::-webkit-scrollbar { width: 10px; height: 10px; }
          ::-webkit-scrollbar-track { background: #f1f1f1; border-radius: 6px; }
          ::-webkit-scrollbar-thumb { background: #8EB9C9; border-radius: 6px; }
          ::-webkit-scrollbar-thumb:hover { background: #060E9F; }

          /* File badge */
          .file-badge {
            display:inline-block; max-width:100%; padding:3px 10px;
            background:#eaf4fb; border:1px solid #8EB9C9; border-radius:20px;
            font-size:0.82rem; color:#333; white-space:nowrap;
            overflow:hidden; text-overflow:ellipsis; vertical-align:middle;
          }
          
          /* 移除 arrow_down 及內建圖示，避免異常顯示純文字 */
          [data-testid="stExpanderToggleIcon"], .material-symbols-rounded {
              display: none !important;
          }
          
        </style>
        """,
        unsafe_allow_html=True,
    )


def analyze_complaint(subject: str, content: str) -> tuple[str, str]:
    s = subject if isinstance(subject, str) else ""
    c = content if isinstance(content, str) else ""
    t = (s + " " + c).lower()

    # 顧客關係類型
    if "註冊" in t and "無法" in t:
        return "顧客關係類型", "其他建議"
    if any(k in t for k in ["不處理", "態度", "搞什麼", "不願意"]):
        return "顧客關係類型", "其他建議"
    if any(k in t for k in ["刪除帳號", "註銷"]):
        return "顧客關係類型", "申請刪除帳號"
    if any(k in t for k in ["手機號碼", "原帳號"]) and any(k in t for k in ["變更", "更改", "修改"]):
        return "顧客關係類型", "更換帳號"
    if any(k in t for k in ["更換帳號", "換帳號"]):
        return "顧客關係類型", "更換帳號"
    if any(k in t for k in ["新增站點", "設站", "建議", "許願"]):
        return "顧客關係類型", "許願新增站點/設站建議"
    if any(k in t for k in ["回收規則", "材質", "可回收"]):
        return "顧客關係類型", "回收物使用規則"

    # APP帳號設定
    if any(k in t for k in ["驗證碼", "認證碼", "otp", "簡訊"]):
        if "忘記密碼" in t:
            return "APP帳號設定問題類型", "忘記密碼/無法重設密碼"
        return "APP帳號設定問題類型", "無法接收簡訊驗證碼"
    if any(k in t for k in ["修改", "更改", "更換"]) and any(k in t for k in ["帳號", "手機", "電話", "號碼"]):
        return "APP帳號設定問題類型", "帳號資訊修改/設定"

    # 登入問題
    if any(k in t for k in ["登入", "登不進去"]) and any(k in t for k in ["螢幕", "機台", "黑掉"]) and any(k in t for k in ["無法", "不能", "失敗", "不了"]):
        return "機台問題類型", "機台操作畫面無法登入"
    if any(k in t for k in ["無法登入", "不能登入", "登不進去", "登入失敗", "登入不了"]):
        return "APP帳號設定問題類型", "APP無法登入"

    # APP使用
    if "可投數量" in t or ("app" in t and "顯示" in t and "0" not in t):
        return "APP使用問題類型", "APP畫面顯示與機台狀態不符"
    if "顯示" in t and "不符" in t:
        return "APP使用問題類型", "APP畫面顯示與機台狀態不符"
    if any(k in t for k in ["app異常", "閃退", "轉圈", "更新"]):
        return "APP使用問題類型", "APP多重異常狀況"

    # 點數
    if "點數" in t and any(k in t for k in ["未累積", "未增加", "沒有入帳", "未入帳"]):
        return "回收點數問題類型", "點數未入帳號"
    if ("點數" in t or "沒入點" in t or "計點" in t) and any(k in t for k in ["未入", "沒入", "不見", "沒記", "沒收到"]):
        return "回收點數問題類型", "點數未入帳號"
    if "點數" in t and any(k in t for k in ["重複", "多給", "多入"]):
        return "回收點數問題類型", "點數重複入點"

    # 優惠券
    if any(k in t for k in ["優惠券", "兌換券", "折價", "序號", "抵用", "對換券", "票卷", "票夾", "條碼", "換這個"]):
        if any(k in t for k in ["提前按下", "操作錯誤", "系統還沒更新", "已更換", "限制", "期限", "規則"]):
            return "優惠券問題類型", "使用規則/限制條件說明"
        if any(k in t for k in ["過期", "還原", "點到", "沒按出條碼"]):
            return "優惠券問題類型", "無法進行兌換操作"
        if any(k in t for k in ["已使用", "失敗", "錯誤", "不能用", "刷不過", "沒有跑出條碼", "這怎麼一回事"]):
            return "優惠券問題類型", "兌換失敗/顯示錯誤"
        if any(k in t for k in ["查詢", "紀錄", "找不到", "在哪"]):
            return "優惠券問題類型", "查詢優惠券序號紀錄"
        return "優惠券問題類型", "無法進行兌換操作"

    # 機台問題
    if any(k in t for k in ["處理中", "卡住"]) and "暫停不動" in t:
        return "機台問題類型", "機台當機/無回應"
    if "寶特瓶卡住" in t or "卡在" in t or ("黑色門" in t and "卡住" in t) or "卡瓶" in t:
        return "機台問題類型", "投入物卡住_瓶罐/電池"
    if any(k in t for k in ["投很多次", "無法辨識", "一直顯示", "不顯示綠燈", "辨識失敗", "辨識異常"]):
        return "機台問題類型", "辨識失敗異常或錯誤"
    if any(k in t for k in ["顯示0都沒有更新", "通報維修"]):
        return "機台問題類型", "機台需維護/故障提醒"
    if any(k in t for k in ["關閉", "設備不動", "不能使用", "撤機", "故障快", "沒開", "未開啟", "關機"]):
        return "機台問題類型", "機台關閉/無法啟動"
    if any(k in t for k in ["髒污不收", "清潔", "髒污"]):
        return "機台問題類型", "機台髒污/需要清潔"
    if any(k in t for k in ["當機", "故障訊息", "沒反應", "lag", "機台異常"]):
        return "機台問題類型", "機台當機/無回應"
    if any(k in t for k in ["滿倉", "收滿", "滿台"]):
        return "機台問題類型", "瓶蓋桶已滿"
    if "運轉不會停止" in t:
        return "機台問題類型", "操作流程異常/無法正常操作"

    # 其他原手機台問題
    if any(k in t for k in ["履帶", "輸送帶", "傳送帶"]) and any(k in t for k in ["不動", "不轉", "異常"]):
        return "機台問題類型", "履帶未作動或異常抖動"
    if any(k in t for k in ["黑屏", "黑畫面", "螢幕異常", "畫面異常", "反光", "黑掉"]):
        return "機台問題類型", "螢幕異常顯示/畫面異常"
    if any(k in t for k in ["維護", "維修", "需維修", "故障提醒"]):
        return "機台問題類型", "機台需維護/故障提醒"
    if any(k in t for k in ["網路連線失敗", "連不上", "連線失敗"]) and "機台" in t:
        return "機台問題類型", "機台網路連線失敗"
    if any(k in t for k in ["網路不穩", "網路中斷"]):
        return "機台問題類型", "網路中斷或不穩定"
    if any(k in t for k in ["重量", "秤重", "偵測重量"]):
        return "機台問題類型", "重量偵測異常"
    if any(k in t for k in ["無法操作", "流程異常", "不能操作"]):
        return "機台問題類型", "操作流程異常/無法正常操作"
    if any(k in t for k in ["投入後沒點", "未獲點數", "未記錄"]):
        return "機台問題類型", "投入後未獲點數/點數未記錄"
    if any(k in t for k in ["中斷", "重啟", "重開機"]):
        return "機台問題類型", "機台運作中斷/重啟"
    if any(k in t for k in ["艙門", "門沒關", "回收艙門"]):
        return "機台問題類型", "回收艙門開啟"
    if "綠燈" in t and "不能" in t:
        return "機台問題類型", "投口綠燈拒收容器"

    return "顧客關係類型", "其他建議"


def parse_pdf_to_df(file_obj) -> pd.DataFrame:
    if pdfplumber is None:
        raise RuntimeError("未安裝 pdfplumber，無法解析 PDF。")
    rows: list[dict] = []
    with pdfplumber.open(file_obj) as pdf:
        for p_idx, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            for ln_idx, line in enumerate(text.splitlines(), start=1):
                cleaned = re.sub(r"\s+", " ", line).strip()
                if cleaned:
                    rows.append({"page": p_idx, "line": ln_idx, "content": cleaned})
    return pd.DataFrame(rows if rows else [{"content": ""}])


def load_input_file(uploaded_file, filename: str = "") -> pd.DataFrame:
    """Load file from a Streamlit UploadedFile or BytesIO. Pass filename when using BytesIO."""
    name = filename or getattr(uploaded_file, "name", "")
    suffix = Path(name).suffix.lower()
    if suffix in [".xlsx", ".xls"]:
        return pd.read_excel(uploaded_file)
    if suffix == ".csv":
        for enc in ["utf-8-sig", "utf-8", "cp950", "big5"]:
            try:
                uploaded_file.seek(0)
                return pd.read_csv(uploaded_file, encoding=enc)
            except (UnicodeDecodeError, AttributeError):
                continue
        uploaded_file.seek(0)
        return pd.read_csv(uploaded_file, encoding="utf-8", errors="replace")
    if suffix == ".pdf":
        return parse_pdf_to_df(uploaded_file)
    raise ValueError(f"僅支援 excel / csv / pdf，收到：{suffix or name}")


def make_unique_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = []
    seen = {}
    for c in df.columns:
        name = str(c)
        if name not in seen:
            seen[name] = 0
            cols.append(name)
        else:
            seen[name] += 1
            cols.append(f"{name}_{seen[name]}")
    out = df.copy()
    out.columns = cols
    return out


# ---- valid type set for fast lookup (all keys + known variant spellings from template) ----
_VALID_TYPES = set(TOPIC_DETAIL_MAP.keys())

# All valid details (flattened from TOPIC_DETAIL_MAP) for quick check
_VALID_DETAILS_FLAT: set[str] = {d for lst in TOPIC_DETAIL_MAP.values() for d in lst}


def _is_valid_pair(t: str, d: str) -> bool:
    """Return True if both type and detail are non-empty and the detail belongs to the type."""
    t, d = t.strip(), d.strip()
    if not t or not d:
        return False
    # Accept if type is valid AND detail is in that type's list
    if t in TOPIC_DETAIL_MAP and d in TOPIC_DETAIL_MAP[t]:
        return True
    # Also accept if type exists but detail is in the FULL detail pool (legacy data)
    if t in _VALID_TYPES and d in _VALID_DETAILS_FLAT:
        return True
    return False


def analyze_dataframe(df: pd.DataFrame, cfg: AnalysisConfig) -> pd.DataFrame:
    out = make_unique_columns(df.copy())

    # ------ Preserve existing valid 問題類型 + 問題細項 from source file ------
    existing_type   = out["問題類型"].copy()   if "問題類型" in out.columns else pd.Series([""] * len(out))
    existing_detail = out["問題細項"].copy()   if "問題細項" in out.columns else pd.Series([""] * len(out))

    # Drop internal columns before re-adding
    for c in ["問題類型", "問題細項", "選取", "部門", "日期", "_ai_filled"]:
        if c in out.columns:
            out = out.drop(columns=[c])

    # Run auto-classification for every row
    preds = out.apply(
        lambda r: analyze_complaint(str(r.get(cfg.subject_col, "")), str(r.get(cfg.content_col, ""))),
        axis=1,
        result_type="expand",
    )
    preds.columns = ["問題類型", "問題細項"]
    out = pd.concat([out, preds], axis=1)

    # ------ Merge: prefer original valid pair; fall back to AI prediction ------
    ai_filled_flags = []
    for idx in range(len(out)):
        orig_type   = str(existing_type.iloc[idx]).strip()
        orig_detail = str(existing_detail.iloc[idx]).strip()
        if _is_valid_pair(orig_type, orig_detail):
            # Original is valid → keep it, NOT AI-filled
            out.iloc[idx, out.columns.get_loc("問題類型")] = orig_type
            out.iloc[idx, out.columns.get_loc("問題細項")] = orig_detail
            ai_filled_flags.append(False)
        else:
            # Original missing/invalid → use AI prediction, mark as AI-filled
            ai_filled_flags.append(True)

    out["_ai_filled"] = ai_filled_flags

    # Final guard: ensure detail always belongs to its topic
    out["問題細項"] = out.apply(
        lambda r: r["問題細項"] if r["問題細項"] in TOPIC_DETAIL_MAP.get(r["問題類型"], [])
                  else TOPIC_DETAIL_MAP.get(r["問題類型"], ["其他建議"])[0],
        axis=1,
    )
    out["選取"] = False
    out["部門"] = out["問題類型"].map(DEPT_MAP).fillna("")
    if cfg.date_col and cfg.date_col in out.columns:
        out["日期"] = pd.to_datetime(out[cfg.date_col], errors="coerce")
    return out


def save_history(df: pd.DataFrame, source_name: str, existing_id: str = "") -> tuple[Path, str]:
    today = datetime.now().strftime("%Y%m%d")
    ts = existing_id if existing_id else datetime.now().strftime("%Y%m%d_%H%M%S")
    output_name = f"{today}_分析.xlsx"
    output_path = HISTORY_DIR / f"{ts}_{output_name}"
    # Remove stale file if overwriting
    if existing_id:
        for old in HISTORY_DIR.glob(f"{existing_id}_*.xlsx"):
            try: old.unlink()
            except: pass
    df.to_excel(output_path, index=False)
    meta = {
        "id": ts,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "source_name": source_name,
        "output_name": output_name,
        "output_path": str(output_path),
        "rows": int(len(df)),
    }
    history = []
    if META_FILE.exists():
        history = json.loads(META_FILE.read_text(encoding="utf-8"))
    # Remove existing entry with same id if overwriting
    history = [i for i in history if i["id"] != ts]
    history.insert(0, meta)
    META_FILE.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")
    return output_path, output_name


def load_history() -> list[dict]:
    if not META_FILE.exists():
        return []
    return json.loads(META_FILE.read_text(encoding="utf-8"))


def safe_filename(text: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", str(text))


def delete_history(item_id: str):
    if not META_FILE.exists(): return
    history = json.loads(META_FILE.read_text(encoding="utf-8"))
    for item in history:
        if item["id"] == item_id:
            try:
                Path(item["output_path"]).unlink(missing_ok=True)
            except Exception:
                pass
    history = [i for i in history if i["id"] != item_id]
    META_FILE.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")


def generate_ai_summary(df: pd.DataFrame) -> str:
    if df.empty:
        return "目前沒有可分析資料。"
    total = len(df)
    type_count = df["問題類型"].value_counts()
    detail_count = df["問題細項"].value_counts()
    top_type = type_count.index[0]
    top_type_count = int(type_count.iloc[0])
    top_detail = detail_count.index[0]
    top_detail_count = int(detail_count.iloc[0])
    return (
        f"1) 目前主力問題為「{top_type}」，共 {top_type_count} 件，占比 {top_type_count/total:.1%}。\n"
        f"2) 最常見細項是「{top_detail}」，共 {top_detail_count} 件，建議列為優先改善。\n"
        "3) 建議以 TOP3 問題建立跨部門改善任務，並每週追蹤件數變化與結案率。"
    )


def generate_ai_summary_llm(df: pd.DataFrame, model_name: str = "gpt-4o-mini") -> str:
    api_key = None
    if hasattr(st, "secrets"):
        try:
            api_key = st.secrets.get("OPENAI_API_KEY", None)
        except Exception:
            api_key = None
    if not api_key:
        api_key = st.session_state.get("OPENAI_API_KEY", "")
    if not api_key or OpenAI is None:
        return generate_ai_summary(df)
    sample = df[["問題類型", "問題細項", "部門"]].head(300).to_dict(orient="records")
    payload = {
        "total_rows": len(df),
        "top_types": df["問題類型"].value_counts().head(6).to_dict(),
        "top_details": df["問題細項"].value_counts().head(10).to_dict(),
        "sample_rows": sample,
    }
    prompt = (
        "你是客服品質分析顧問。請用繁體中文輸出3-5點重點，格式精簡，"
        "包含: 高頻問題、可能根因、跨部門優先改善建議。資料如下:\n"
        f"{json.dumps(payload, ensure_ascii=False)}\n"
        "請特別列出：1. 站點城市分布熱點 (如果從內容看得出來) 2. 問題類型與細項的熱點(最高頻的異常)。"
    )
    try:
        client = OpenAI(api_key=api_key)
        res = client.responses.create(model=model_name, input=prompt)
        text = getattr(res, "output_text", "").strip()
        return text if text else generate_ai_summary(df)
    except Exception:
        return generate_ai_summary(df)


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="analysis")
    return buffer.getvalue()


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def to_pdf_bytes(df: pd.DataFrame) -> bytes:
    """Generate PDF using reportlab for proper CJK text wrapping in cells."""
    try:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.units import mm
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        pdfmetrics.registerFont(UnicodeCIDFont('HeiseiMin-W3'))
        BASE_FONT = 'HeiseiMin-W3'
    except Exception:
        BASE_FONT = 'Helvetica'
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.units import mm

    buf = io.BytesIO()
    page_w, page_h = landscape(A4)
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                            leftMargin=10*mm, rightMargin=10*mm,
                            topMargin=10*mm, bottomMargin=10*mm)

    styles = getSampleStyleSheet()
    cell_style = styles['Normal'].clone('cell')
    cell_style.fontSize = 7
    cell_style.fontName = BASE_FONT
    cell_style.wordWrap = 'CJK'
    cell_style.leading = 10

    header_style = styles['Normal'].clone('hdr')
    header_style.fontSize = 7
    header_style.fontName = BASE_FONT
    header_style.textColor = colors.white
    header_style.wordWrap = 'CJK'

    table_df = df.head(50).copy()
    # Drop internal helper columns
    drop_cols = [c for c in ['選取'] if c in table_df.columns]
    table_df = table_df.drop(columns=drop_cols).fillna('')

    WRAP_COLS = {'用戶內容', '主旨', '問題主旨'}
    COL_W_MAP = {}  # col_name -> width in mm
    num_cols = len(table_df.columns)
    base_w = (page_w - 20*mm) / num_cols

    def make_cell(val, is_header=False):
        s = str(val)
        st = header_style if is_header else cell_style
        return Paragraph(s, st)

    data = [[make_cell(c, True) for c in table_df.columns]]
    for _, row in table_df.iterrows():
        data.append([make_cell(str(v)) for v in row.values])

    col_widths = [base_w] * num_cols

    t = Table(data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#3E75A0')),
        ('TEXTCOLOR',  (0,0), (-1,0), colors.white),
        ('FONTNAME',   (0,0), (-1,-1), BASE_FONT),
        ('FONTSIZE',   (0,0), (-1,-1), 7),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#EBF4FA')]),
        ('GRID',       (0,0), (-1,-1), 0.4, colors.HexColor('#CCCCCC')),
        ('VALIGN',     (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING',(0,0), (-1,-1), 3),
        ('RIGHTPADDING',(0,0),(-1,-1), 3),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING',(0,0),(-1,-1), 2),
    ]))

    elements = [t]
    doc.build(elements)
    return buf.getvalue()


def build_chart_pack(df: pd.DataFrame) -> dict[str, bytes]:
    """Build chart images (PNG) for download/PPT."""
    data = df.copy()
    stats = data["問題類型"].value_counts().rename_axis("問題類型").reset_index(name="件數")
    stats["百分比"] = (stats["件數"] / max(stats["件數"].sum(), 1) * 100).round(1)
    detail_stats = data["問題細項"].value_counts().reset_index().head(10)
    detail_stats.columns = ["問題細項", "件數"]

    # enhanced font detection for CJK
    import matplotlib.font_manager as fm
    import os
    import requests
    
    FONT_PATH = "NotoSansTC-Regular.otf"
    if not os.path.exists(FONT_PATH):
        try:
            # Download a CJK font subset if not exists
            resp = requests.get("https://raw.githubusercontent.com/googlefonts/noto-cjk/main/Sans/SubsetOTF/TC/NotoSansCJKtc-Regular.otf", timeout=10)
            if resp.status_code == 200:
                with open(FONT_PATH, "wb") as f:
                    f.write(resp.content)
        except Exception:
            pass
            
    if os.path.exists(FONT_PATH):
        try:
            fm.fontManager.addfont(FONT_PATH)
            plt.rcParams["font.family"] = fm.FontProperties(fname=FONT_PATH).get_name()
        except:
            pass
    else:
        # fallback to system CJK fonts
        sys_fonts = [f.name for f in fm.fontManager.ttflist if any(x in f.name for x in ["CJK", "TC", "TW", "JhengHei", "SimHei", "Fallback"])]
        if sys_fonts:
            plt.rcParams["font.family"] = sys_fonts[0]
            
    plt.rcParams["axes.unicode_minus"] = False

    # 1) type distribution
    fig1, ax1 = plt.subplots(figsize=(8, 4.5))
    colors = ["#FF5000", "#060E9F", "#FFCE00", "#8EB9C9", "#0076A9", "#FAE0B8"]
    ax1.bar(stats["問題類型"], stats["件數"], color=colors[: len(stats)])
    ax1.set_title("問題類型分布")
    ax1.set_ylabel("件數")
    ax1.tick_params(axis="x", rotation=20)
    for i, r in stats.iterrows():
        ax1.text(i, r["件數"], f'{r["百分比"]:.1f}%', ha="center", va="bottom", fontsize=9)
    fig1.tight_layout()
    b1 = io.BytesIO()
    fig1.savefig(b1, format="png", dpi=180)
    plt.close(fig1)

    # 2) machine ratio pie
    fig2, ax2 = plt.subplots(figsize=(6.2, 4.5))
    df_machine = data[data["問題類型"] == "機台問題類型"].copy()
    if df_machine.empty:
        ax2.text(0.5, 0.5, "無機台相關資料", ha="center", va="center")
    else:
        def get_machine_type(row):
            txt = str(row.get("用戶內容", "")) + " " + str(row.get("主旨", ""))
            if "方舟" in txt:
                return "方舟站"
            if "電池" in txt:
                return "電池機"
            return "收瓶機"
        df_machine["機台機型"] = df_machine.apply(get_machine_type, axis=1)
        pie_stats = df_machine["機台機型"].value_counts()
        ax2.pie(pie_stats.values, labels=pie_stats.index, autopct="%1.1f%%", colors=["#0076A9", "#8EB9C9", "#FAE0B8"])
        ax2.set_title("機台問題類型分布")
    fig2.tight_layout()
    b2 = io.BytesIO()
    fig2.savefig(b2, format="png", dpi=180)
    plt.close(fig2)

    # 3) top detail horizontal bar
    fig3, ax3 = plt.subplots(figsize=(8, 4.5))
    d = detail_stats.sort_values("件數", ascending=True)
    ax3.barh(d["問題細項"], d["件數"], color="#1f77b4")
    ax3.set_title("十大問題細項分布")
    ax3.set_xlabel("件數")
    fig3.tight_layout()
    b3 = io.BytesIO()
    fig3.savefig(b3, format="png", dpi=180)
    plt.close(fig3)

    # dashboard merged image
    fig4 = plt.figure(figsize=(14, 5))
    gs = fig4.add_gridspec(1, 3)
    a1 = fig4.add_subplot(gs[0, 0])
    a2 = fig4.add_subplot(gs[0, 1])
    a3 = fig4.add_subplot(gs[0, 2])
    a1.bar(stats["問題類型"], stats["件數"], color=colors[: len(stats)])
    a1.set_title("問題類型分布")
    a1.tick_params(axis="x", rotation=18)
    if df_machine.empty:
        a2.text(0.5, 0.5, "無機台資料", ha="center", va="center")
    else:
        pie_stats = df_machine["機台機型"].value_counts()
        a2.pie(pie_stats.values, labels=pie_stats.index, autopct="%1.1f%%")
        a2.set_title("機台問題占比")
    a3.barh(d["問題細項"], d["件數"], color="#1f77b4")
    a3.set_title("十大細項")
    fig4.tight_layout()
    b4 = io.BytesIO()
    fig4.savefig(b4, format="png", dpi=180)
    plt.close(fig4)

    return {
        "chart_問題類型分布.png": b1.getvalue(),
        "chart_機台問題占比.png": b2.getvalue(),
        "chart_十大問題細項.png": b3.getvalue(),
        "chart_dashboard.png": b4.getvalue(),
    }


def build_ppt_bytes(stats: pd.DataFrame, ai_text: str, source_name: str,
                    template_path: str = r"C:\Users\fen\Desktop\簡報範本.pptx",
                    chart_pack: Optional[dict[str, bytes]] = None) -> bytes:
    # Load template if exists, otherwise blank.
    # Keep template background/master unchanged.
    prs = Presentation(template_path) if Path(template_path).exists() else Presentation()
    # Try to find a master layout that looks like a title + content or blank
    slide_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]

    W = prs.slide_width
    H = prs.slide_height

    def add_slide(prs):
        return prs.slides.add_slide(slide_layout)

    # Slide 1: Summary
    slide1 = add_slide(prs)

    def add_text_box(slide, text, left, top, width, height, size=18, bold=False, color=(0x1a,0x1a,0x1a)):
        txb = slide.shapes.add_textbox(left, top, width, height)
        tf = txb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(size)
        p.font.bold = bold
        p.font.color.rgb = RGBColor(*color)
        return txb

    add_text_box(slide1, "【客服課】ECOCO 客訴分析簡報", Inches(0.5), Inches(0.28), W - Inches(1), Inches(0.65),
                 size=24, bold=True, color=(0x06,0x0E,0x9F))
    add_text_box(slide1, f"來源檔案：{source_name}　產出日期：{datetime.now().strftime('%Y-%m-%d')}",
                 Inches(0.5), Inches(0.9), W - Inches(1), Inches(0.45), size=12)
    
    # AI summary text box
    txb2 = slide1.shapes.add_textbox(Inches(0.6), Inches(1.45), W - Inches(1.2), H - Inches(1.85))
    tf2 = txb2.text_frame
    tf2.word_wrap = True
    first = True
    for ln in ["【客服課產出分析重點】"] + ai_text.splitlines():
        if not ln.strip(): continue
        if first:
            p = tf2.paragraphs[0]; first = False
        else:
            p = tf2.add_paragraph()
        p.text = ln.strip()
        p.font.size = Pt(11)
        p.font.color.rgb = RGBColor(0x31, 0x33, 0x3F)

    # Slide 2: Table
    slide2 = add_slide(prs)
    add_text_box(slide2, "【客服課】問題類型件數與占比", Inches(0.5), Inches(0.2), W - Inches(1), Inches(0.6),
                 size=20, bold=True, color=(0x06,0x0E,0x9F))

    rows_n = min(len(stats) + 1, 15)
    cols_n = 4
    table_left = Inches(0.4)
    table_top = Inches(1.1)
    table_width = W - Inches(0.8)
    table_height = Inches(0.45) * rows_n # Auto-calc height based on rows
    tbl_shape = slide2.shapes.add_table(rows_n, cols_n, table_left, table_top, table_width, table_height)
    tbl = tbl_shape.table
    tbl.columns[0].width = Inches(4.0)
    tbl.columns[1].width = Inches(1.3)
    tbl.columns[2].width = Inches(1.4)
    tbl.columns[3].width = Inches(6.0)
    
    headers = ["問題類型", "件數", "百分比", "歸屬部門"]
    for ci, h in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.text = h
        cell.fill.solid()
        # brand blue header
        cell.fill.fore_color.rgb = RGBColor(0x06, 0x0E, 0x9F)
        for para in cell.text_frame.paragraphs:
            para.alignment = 1 # Center
            for run in para.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.size = Pt(14)
                # Try to force font family for Chinese
                try: run.font.name = 'Microsoft JhengHei'
                except: pass

    for ri, (_, r) in enumerate(stats.head(rows_n - 1).iterrows(), start=1):
        # 確保百分比讀取回整數字串且具有 %
        try:
            pct_val = f'{int(float(r["百分比"]))}%'
        except:
            pct_val = f'{r["百分比"]}%'
        vals = [str(r["問題類型"]), str(r["件數"]), pct_val, str(r.get("歸屬部門", ""))]
        for ci, v in enumerate(vals):
            cell = tbl.cell(ri, ci)
            cell.text = v
            cell.fill.solid()
            # brand alternating row color
            if ri % 2 == 0:
                cell.fill.fore_color.rgb = RGBColor(0xE8, 0xF1, 0xF5)  # light blue
            else:
                cell.fill.fore_color.rgb = RGBColor(0xFA, 0xE0, 0xB8)  # beige
            for para in cell.text_frame.paragraphs:
                para.alignment = 1 # Center
                for run in para.runs:
                    run.font.size = Pt(13)
                    run.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
                    try: run.font.name = 'Microsoft JhengHei'
                    except: pass

    # Slide 3: chart dashboard preview (optional)
    if chart_pack and "chart_dashboard.png" in chart_pack:
        slide3 = add_slide(prs)
        add_text_box(slide3, "【客服課】圖表分析總覽", Inches(0.5), Inches(0.2), W - Inches(1), Inches(0.6),
                     size=20, bold=True, color=(0x06, 0x0E, 0x9F))
        img_stream = io.BytesIO(chart_pack["chart_dashboard.png"])
        slide3.shapes.add_picture(img_stream, Inches(0.45), Inches(0.95), width=W - Inches(0.9), height=H - Inches(1.3))

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def upload_to_google_sheet(df: pd.DataFrame, credentials_json: dict, spreadsheet_id: str, worksheet_name: str) -> None:
    if gspread is None or Credentials is None:
        raise RuntimeError("尚未安裝 gspread 或 google-auth。")
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(credentials_json, scopes=scopes)
    client = gspread.authorize(creds)
    sh = client.open_by_key(spreadsheet_id)
    try:
        ws = sh.worksheet(worksheet_name)
        ws.clear()
    except Exception:
        ws = sh.add_worksheet(title=worksheet_name, rows=1000, cols=30)
    values = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
    ws.update(values)


def section_1():
    st.subheader("功能一：檔案上傳與分析區")
    st.markdown("<div class='ecoco-card'>支援上傳 excel / csv / pdf，分析並產出【問題類型、問題細項】。</div>", unsafe_allow_html=True)

    # File info badge — no long text, just a compact pill with truncated name
    if st.session_state.get("_uploaded_bytes") and st.session_state.get("_uploaded_name"):
        fname_short = st.session_state['_uploaded_name']
        if len(fname_short) > 30:
            fname_short = fname_short[:14] + "..." + fname_short[-12:]
        col_badge, col_clear = st.columns([9, 1])
        col_badge.markdown(
            f"<span class='file-badge'>&#128196; {fname_short}</span>",
            unsafe_allow_html=True
        )
        if col_clear.button("x 清除", help="清除目前檔案，重新上傳"):
            for key in ["_uploaded_bytes", "_uploaded_name", "_uploaded_type", "analysis_df", "source_name"]:
                st.session_state.pop(key, None)
            st.rerun()

    uploaded = st.file_uploader("上傳新檔案", type=["xlsx", "xls", "csv", "pdf"], key="uploader")
    # Persist file bytes across menu switches
    if uploaded is not None:
        st.session_state["_uploaded_bytes"] = uploaded.read()
        st.session_state["_uploaded_name"] = uploaded.name
        st.session_state["_uploaded_type"] = uploaded.type

    # Restore from session if user switched tabs and came back
    if uploaded is None and st.session_state.get("_uploaded_bytes") is not None:
        saved_name = st.session_state.get("_uploaded_name", "file")
        buf = io.BytesIO(st.session_state["_uploaded_bytes"])
        df_raw_bytes = load_input_file(buf, filename=saved_name)
        st.caption(f"已載入 {saved_name}（從記憶復原），資料筆數：{len(df_raw_bytes)}")
        df_raw = make_unique_columns(df_raw_bytes)
        uploaded_name = saved_name
    elif uploaded is not None:
        fname = st.session_state.get("_uploaded_name", uploaded.name)
        df_raw = make_unique_columns(load_input_file(
            io.BytesIO(st.session_state["_uploaded_bytes"]), filename=fname
        ))
        uploaded_name = uploaded.name
        st.caption(f"已載入 {uploaded.name}，資料筆數：{len(df_raw)}")
    else:
        if "analysis_df" not in st.session_state:
            st.info("請上傳檔案開始分析。")
            return
        # Already analysed, show results without needing the raw file
        df_raw = None
        uploaded_name = st.session_state.get("source_name", "")

    if df_raw is not None:
        cols = list(df_raw.columns)
        if not cols:
            st.warning("檔案沒有可用欄位。")
            return

        st.markdown("##### 分析前篩選與欄位設定")
        subject_col = st.selectbox("用戶填寫的主題欄位", options=cols, index=0)
        content_col = st.selectbox("用戶內容欄位", options=cols, index=min(1, len(cols) - 1))
        date_opt = ["(無)"] + cols
        date_col = st.selectbox("日期欄位（選填）", options=date_opt, index=0)
        pre_keyword = st.text_input("分析前篩選關鍵字（主題/內容，選填）")
        cfg = AnalysisConfig(subject_col=subject_col, content_col=content_col,
                             date_col=None if date_col == "(無)" else date_col)

        if st.button("開始分析", type="primary"):
            work = df_raw.copy()
            if pre_keyword:
                work = work[
                    work[subject_col].astype(str).str.contains(pre_keyword, case=False, na=False)
                    | work[content_col].astype(str).str.contains(pre_keyword, case=False, na=False)
                ]
            st.session_state["analysis_df"] = analyze_dataframe(work, cfg)
            st.session_state["source_name"] = uploaded_name

    if "analysis_df" not in st.session_state:
        return
    df = st.session_state["analysis_df"]
    c1, c2, c3 = st.columns([2, 2, 1])
    keyword = c1.text_input("篩選：關鍵字（主題/內容）")
    filter_type = c2.multiselect("篩選：問題類型", options=TYPE_OPTIONS, default=[])
    
    valid_details = DETAIL_OPTIONS
    if filter_type:
        valid_details = []
        for t in filter_type:
            valid_details.extend(TOPIC_DETAIL_MAP.get(t, []))
            
    filter_detail = c3.multiselect("篩選：問題細項", options=valid_details, default=[])

    show = make_unique_columns(df.copy())
    # hide_index=True alone sometimes still shows original integer index;
    # reset to guarantee no row numbers in data_editor
    show = show.reset_index(drop=True)
    if keyword:
        show = show[
            show[subject_col].astype(str).str.contains(keyword, case=False, na=False)
            | show[content_col].astype(str).str.contains(keyword, case=False, na=False)
        ]
    if filter_type:
        show = show[show["問題類型"].isin(filter_type)]
    if filter_detail:
        show = show[show["問題細項"].isin(filter_detail)]

    st.markdown("#### 可編輯標記表（支援下拉 + 手動編輯）")

    # ---- AI填入標示 ---
    ai_col = "_ai_filled"
    MARKER_COL = "AI標記"  # kept for save compatibility only
    has_ai_col = ai_col in show.columns
    n_ai = 0
    if has_ai_col:
        n_ai = int(show[ai_col].fillna(False).astype(bool).sum())

    if n_ai > 0:
        st.markdown(
            f"""
            <div style='background:#fff5f5; border:1px solid #ffb3b3; border-radius:8px;
                        padding:8px 14px; margin-bottom:8px; font-size:0.85rem;'>
              <b style='color:#cc0000;'>● AI 自動標記</b>：共 <b style='color:#cc0000;'>{n_ai} 筆</b> 原始欄位空白或無效，
              已由 AI 根據客訴內容自動分析填入。
              請針對這幾筆核對，如需修改請直接在表格中下拉選擇，再點「💾 儲存修改」確認。
            </div>
            """,
            unsafe_allow_html=True
        )

    st.caption("💡 直接在表格中下拉選擇問題類型 / 問題細項，調整完成後點擊「💾 儲存修改」。")

    # 重新處理要顯示的欄位，確保原本隱藏的 MARKER_COL 正確加入
    display_cols = [c for c in show.columns if c not in (ai_col, MARKER_COL)]
    show_display = show[display_cols].reset_index(drop=True)

    # 新增一欄字號標記給 AI 填入的資料
    if has_ai_col:
        flags = show[ai_col].reset_index(drop=True)
        marker_vals = flags.map(lambda x: "⭐(AI填寫)" if x else "")
    else:
        marker_vals = [""] * len(show_display)
        
    insert_idx = 1
    show_display.insert(insert_idx, MARKER_COL, marker_vals)

    # --- Select All Trigger ---
    # Streamlit header clicks are not native. 
    # We include a "Select" label above the header that acts as a toggle.
    # We use a button to mimic a header-style toggle.
    cols_h = st.columns([13, 2])
    if cols_h[1].button("⬓ 選取 / 取消", key="toggle_all_btn", help="點擊此處可全選或取消全選"):
        all_sel = bool(df["選取"].all()) if "選取" in df.columns and not df.empty else False
        st.session_state["analysis_df"]["選取"] = not all_sel
        st.rerun()

    edited = st.data_editor(
        show_display,
        use_container_width=True,
        num_rows="dynamic",
        hide_index=True,
        column_config={
            "選取": st.column_config.CheckboxColumn(help="勾選要批次處理的列"),
            MARKER_COL: st.column_config.TextColumn("備註", disabled=True),
            "問題類型": st.column_config.SelectboxColumn(options=TYPE_OPTIONS, required=True),
            "問題細項": st.column_config.SelectboxColumn(options=DETAIL_OPTIONS, required=True),
            "部門": st.column_config.SelectboxColumn(options=DEPT_OPTIONS),
        },
        key="editor_table",
    )

    # 儲存按鈕在表格下方
    sv_col1, sv_col2, sv_col3 = st.columns([2, 2, 6])
    if sv_col1.button("💾 儲存修改", use_container_width=True):
        full_df = st.session_state["analysis_df"].copy()
        # Drop the AI marker column and 選取 before saving back
        save_edited = edited.drop(columns=["選取", MARKER_COL], errors="ignore")
        full_df.update(save_edited)
        # Clear _ai_filled flags for all saved rows (user has confirmed)
        if "_ai_filled" in full_df.columns:
            full_df["_ai_filled"] = False
        st.session_state["analysis_df"] = full_df
        # Also push to drafts list
        src_name = st.session_state.get("source_name", "未命名")
        if "_draft_list" not in st.session_state:
            st.session_state["_draft_list"] = []
        # Avoid duplicate same name drafts – update existing
        draft_ids = [d["name"] for d in st.session_state["_draft_list"]]
        if src_name not in draft_ids:
            st.session_state["_draft_list"].insert(0, {"name": src_name, "df": full_df.copy()})
        else:
            for d in st.session_state["_draft_list"]:
                if d["name"] == src_name:
                    d["df"] = full_df.copy()
        st.success(f"已儲存「{src_name}」")

    # 已儲存草稿列表
    if st.session_state.get("_draft_list"):
        st.markdown("---")
        st.markdown("##### 已儲存的草稿")
        for idx, draft in enumerate(st.session_state["_draft_list"]):
            d_col1, d_col2, d_col3, d_col4 = st.columns([5, 1, 1, 1])
            d_col1.markdown(
                f"<div style='padding-top:0.45rem; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-weight:600;'>"
                f"📄 {draft['name']}</div>",
                unsafe_allow_html=True
            )
            if d_col2.button("[載入]", key=f"draft_load_{idx}", use_container_width=True):
                st.session_state["analysis_df"] = draft["df"].copy()
                st.session_state["source_name"] = draft["name"]
                st.success(f"已載入「{draft['name']}」，可繼續編輯。")
            if d_col3.button("[修改]", key=f"draft_edit_{idx}", use_container_width=True):
                st.session_state["analysis_df"] = draft["df"].copy()
                st.session_state["source_name"] = draft["name"]
                st.rerun()
            if d_col4.button("[X]", key=f"draft_del_{idx}", use_container_width=True):
                st.session_state["_draft_list"].pop(idx)
                st.rerun()


    b1, b2, b3, b4 = st.columns(4)
    batch_type = b1.selectbox("批次設定問題類型", options=["(不變更)"] + TYPE_OPTIONS)
    detail_candidates = ["(不變更)"] + (TOPIC_DETAIL_MAP.get(batch_type, DETAIL_OPTIONS) if batch_type != "(不變更)" else DETAIL_OPTIONS)
    batch_detail = b2.selectbox("批次設定問題細項", options=detail_candidates)
    if b3.button("套用到勾選列"):
        mask = edited["選取"] == True
        if batch_type != "(不變更)":
            edited.loc[mask, "問題類型"] = batch_type
            edited.loc[mask, "部門"] = edited.loc[mask, "問題類型"].map(DEPT_MAP).fillna("")
        if batch_detail != "(不變更)":
            edited.loc[mask, "問題細項"] = batch_detail
        # Auto-fix rows whose detail mismatches topic
        edited["問題細項"] = edited.apply(
            lambda r: r["問題細項"] if r["問題細項"] in TOPIC_DETAIL_MAP.get(r["問題類型"], []) else TOPIC_DETAIL_MAP.get(r["問題類型"], ["其他建議"])[0],
            axis=1,
        )
        st.session_state["analysis_df"] = edited.copy()
        st.session_state["_batch_applied"] = True
    if st.session_state.pop("_batch_applied", False):
        st.success("已套用批次編輯。")
    if b4.button("刪除勾選列"):
        st.session_state["analysis_df"] = edited[edited["選取"] != True].copy()
        st.success("已刪除勾選列。")

    final_df = st.session_state["analysis_df"]
    
    st.markdown("#### 下載分析結果 (下載後自動歸檔至歷史紀錄)")
    dl_format = st.radio("選擇下載格式", ["Excel", "CSV", "PDF"], horizontal=True)
    
    def on_download():
        existing_id = st.session_state.pop("_editing_history_id", "")
        save_history(final_df, st.session_state.get("source_name", "unknown"), existing_id=existing_id)
        st.session_state["history_saved_msg"] = True

    if dl_format == "Excel":
        out_name = f"{datetime.now().strftime('%Y%m%d')}_分析.xlsx"
        data_bytes = to_excel_bytes(final_df)
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif dl_format == "CSV":
        out_name = f"{datetime.now().strftime('%Y%m%d')}_分析.csv"
        data_bytes = to_csv_bytes(final_df)
        mime = "text/csv"
    else:
        out_name = f"{datetime.now().strftime('%Y%m%d')}_分析單.pdf"
        try:
            data_bytes = to_pdf_bytes(final_df)
            mime = "application/pdf"
        except Exception as e:
            st.error(f"PDF 產生錯誤: {e}")
            data_bytes = b""
            mime = "application/pdf"

    st.download_button(
        label=f"📥 下載 {dl_format} 格式分析",
        data=data_bytes,
        file_name=out_name,
        mime=mime,
        on_click=on_download
    )
    
    if st.session_state.get("history_saved_msg"):
        st.success("檔案已下載，並自動保存至歷史紀錄。")
        st.session_state["history_saved_msg"] = False

    st.markdown("#### 分析文字產出")
    summary_text = generate_ai_summary(final_df)
    st.text_area("分析結果文字", summary_text, height=120)
    st.download_button(
        "下載分析文字（txt）",
        data=summary_text.encode("utf-8"),
        file_name=f"{datetime.now().strftime('%Y%m%d')}_分析文字.txt",
        mime="text/plain",
    )

    with st.expander("上傳到 Google Sheet"):
        st.write("請提供 Service Account JSON 與 Spreadsheet ID")
        cred_file = st.file_uploader("Google Service Account JSON", type=["json"], key="gcp_json")
        spreadsheet_id = st.text_input("Spreadsheet ID")
        ws_name = st.text_input("Worksheet 名稱", value=datetime.now().strftime("%Y%m%d_分析"))
        if st.button("上傳 Google Sheet"):
            if not cred_file or not spreadsheet_id:
                st.error("請先上傳 JSON 並填寫 Spreadsheet ID。")
            else:
                credentials_json = json.loads(cred_file.getvalue().decode("utf-8"))
                upload_to_google_sheet(final_df, credentials_json, spreadsheet_id, ws_name)
                st.success(f"已上傳到 Google Sheet 工作表：{ws_name}")


def render_charts_from_stats(stats: pd.DataFrame, df: pd.DataFrame, key_prefix: str = ""):
    """Render charts. stats may be manually edited in section_2."""
    c1, c2, c3 = st.columns(3)
    
    fig1 = px.bar(
        stats, x="問題類型", y="件數", color="歸屬部門", text="百分比", title="問題類型分布",
        color_discrete_sequence=["#FF5000", "#060E9F", "#FFCE00", "#8EB9C9", "#0076A9", "#FAE0B8"]
    )
    fig1.update_traces(texttemplate="%{text}%", textposition="outside")
    fig1.update_layout(height=400)
    c1.plotly_chart(fig1, use_container_width=True, key=f"{key_prefix}_fig1" if key_prefix else None)

    df_machine = df[df["問題類型"] == "機台問題類型"].copy()
    if not df_machine.empty:
        def get_machine_type(row):
            txt = str(row.get("用戶內容", "")) + " " + str(row.get("主旨", ""))
            if "方舟" in txt: return "方舟站"
            if "電池" in txt: return "電池機"
            return "收瓶機"
        df_machine["機台機型"] = df_machine.apply(get_machine_type, axis=1)
        m_stats = df_machine["機台機型"].value_counts().reset_index()
        m_stats.columns = ["機型", "件數"]
        fig2 = px.pie(m_stats, names="機型", values="件數", title="機台問題細分比較", hole=0.3)
        fig2.update_layout(height=400, margin=dict(t=40, b=0, l=0, r=0))
        c2.plotly_chart(fig2, use_container_width=True, key=f"{key_prefix}_fig2" if key_prefix else None)
    else:
        c2.info("無機台相關數據")

    detail_stats = df["問題細項"].value_counts().reset_index().head(10)
    detail_stats.columns = ["問題細項", "件數"]
    fig3 = px.bar(detail_stats, x="件數", y="問題細項", orientation='h', title="十大問題細項分佈")
    fig3.update_layout(height=400, yaxis={'categoryorder':'total ascending'}, margin=dict(t=40, b=0, l=0, r=0))
    c3.plotly_chart(fig3, use_container_width=True, key=f"{key_prefix}_fig3" if key_prefix else None)


def render_charts(df: pd.DataFrame, key_prefix: str = ""):
    date_cols = [c for c in df.columns if "日期" in c or "date" in c.lower()]
    if date_cols:
        dcol = date_cols[0]
        try:
            df[dcol] = pd.to_datetime(df[dcol], errors="coerce")
            valid_dates = df[dcol].dropna()
            if not valid_dates.empty:
                min_d = valid_dates.min().date()
                max_d = valid_dates.max().date()
                st.markdown("##### 分析日期區間")
                c_d1, c_d2 = st.columns(2)
                start_d = c_d1.date_input("起始日期", value=min_d, min_value=min_d, max_value=max_d, key=f"{key_prefix}_sd")
                end_d   = c_d2.date_input("結束日期", value=max_d, min_value=min_d, max_value=max_d, key=f"{key_prefix}_ed")
                df = df[(df[dcol].dt.date >= start_d) & (df[dcol].dt.date <= end_d)]
        except Exception:
            pass

    stats = df["問題類型"].value_counts().rename_axis("問題類型").reset_index(name="件數")
    stats["百分比"] = (stats["件數"] / max(stats["件數"].sum(), 1) * 100).round(0).astype(int)
    stats["歸屬部門"] = stats["問題類型"].map(DEPT_MAP).fillna("未分配")

    c1, c2, c3 = st.columns(3)
    
    fig1 = px.bar(
        stats, x="問題類型", y="件數", color="歸屬部門", text="百分比", title="問題類型分布",
        color_discrete_sequence=["#FF5000", "#060E9F", "#FFCE00", "#8EB9C9", "#0076A9", "#FAE0B8"]
    )
    fig1.update_traces(texttemplate="%{text}%", textposition="outside")
    fig1.update_layout(height=400)
    c1.plotly_chart(fig1, use_container_width=True, key=f"{key_prefix}_fig1" if key_prefix else None)

    df_machine = df[df["問題類型"] == "機台問題類型"].copy()
    if not df_machine.empty:
        def get_machine_type(row):
            txt = str(row.get("用戶內容", "")) + " " + str(row.get("主旨", ""))
            if "方舟" in txt: return "方舟站"
            if "電池" in txt: return "電池機"
            return "收瓶機"
        df_machine["機台機型"] = df_machine.apply(get_machine_type, axis=1)
        m_stats = df_machine["機台機型"].value_counts().reset_index()
        m_stats.columns = ["機型", "件數"]
        fig2 = px.pie(m_stats, names="機型", values="件數", title="機台問題細分比較", hole=0.3)
        fig2.update_layout(height=400, margin=dict(t=40, b=0, l=0, r=0))
        c2.plotly_chart(fig2, use_container_width=True, key=f"{key_prefix}_fig2" if key_prefix else None)
    else:
        c2.info("無機台相關數據")

    detail_stats = df["問題細項"].value_counts().reset_index().head(10)
    detail_stats.columns = ["問題細項", "件數"]
    fig3 = px.bar(detail_stats, x="件數", y="問題細項", orientation='h', title="十大問題細項分佈")
    fig3.update_layout(height=400, yaxis={'categoryorder':'total ascending'}, margin=dict(t=40, b=0, l=0, r=0))
    c3.plotly_chart(fig3, use_container_width=True, key=f"{key_prefix}_fig3" if key_prefix else None)


def section_2():
    st.subheader("功能二：圖表化與 AI 重點分析")
    if "analysis_df" not in st.session_state:
        st.info("請先在功能一完成分析。")
        return
    df_full = st.session_state["analysis_df"]
    if df_full.empty:
        st.warning("目前沒有資料。")
        return

    # --- Date range filter ---
    date_cols = [c for c in df_full.columns if "日期" in c or "date" in c.lower()]
    df = df_full.copy()
    if date_cols:
        dcol = date_cols[0]
        try:
            df[dcol] = pd.to_datetime(df[dcol], errors="coerce")
            valid_dates = df[dcol].dropna()
            if not valid_dates.empty:
                min_d = valid_dates.min().date()
                max_d = valid_dates.max().date()
                st.markdown("##### 分析日期區間")
                dr_col1, dr_col2 = st.columns(2)
                start_d = dr_col1.date_input("起始日期", value=min_d, min_value=min_d, max_value=max_d)
                end_d   = dr_col2.date_input("結束日期", value=max_d, min_value=min_d, max_value=max_d)
                df = df[(df[dcol].dt.date >= start_d) & (df[dcol].dt.date <= end_d)]
                st.caption(f"目前顯示 {len(df)} 筆 / 共 {len(df_full)} 筆")
        except Exception:
            pass

    stats = df["問題類型"].value_counts().rename_axis("問題類型").reset_index(name="件數")
    stats["百分比"] = (stats["件數"] / max(stats["件數"].sum(), 1) * 100).round(0).astype(int)
    stats["歸屬部門"] = stats["問題類型"].map(DEPT_MAP).fillna("")

    # Build totals row
    total_count = int(stats["件數"].sum())
    dept_totals = stats.groupby("歸屬部門")["件數"].sum()
    dept_summary = "  ".join([f"{d}:{int(n)}件" for d, n in dept_totals.items() if d])
    totals_row = pd.DataFrame([{
        "問題類型": "[ 合計 ]",
        "件數": total_count,
        "百分比": 100,
        "歸屬部門": dept_summary,
    }])
    stats_with_total = pd.concat([stats, totals_row], ignore_index=True)

    st.markdown("#### 類型件數與部門 (可直接編輯，圖表即時同步)")
    edited_stats = st.data_editor(
        stats_with_total,
        use_container_width=True,
        hide_index=True,
        column_config={
            "歸屬部門": st.column_config.SelectboxColumn(options=DEPT_OPTIONS + [dept_summary]),
            "百分比": st.column_config.NumberColumn(format="%d %%")
        },
        key="stats_editor",
        num_rows="fixed",
    )
    # Use main stats (drop totals row) for charts
    chart_stats = edited_stats[edited_stats["問題類型"] != "[ 合計 ]"]
    render_charts_from_stats(chart_stats, df, key_prefix="sec2")

    st.markdown("#### AI 問題重點分析")
    st.markdown("##### AI 設定（選填）")
    col_ai_1, col_ai_2 = st.columns([3, 2])
    key_input = col_ai_1.text_input("OpenAI API Key（若留空則使用內建規則摘要）", type="password")
    model_name = col_ai_2.text_input("模型", value="gpt-4o-mini")
    if key_input:
        st.session_state["OPENAI_API_KEY"] = key_input

    ai_text = generate_ai_summary_llm(df, model_name=model_name)
    st.text_area("分析摘要預覽", ai_text, height=140)
    chart_pack = build_chart_pack(df)
    st.download_button(
        "下載 AI 分析文字檔",
        data=ai_text.encode("utf-8"),
        file_name=f"{datetime.now().strftime('%Y%m%d')}_AI重點分析.txt",
        mime="text/plain",
    )
    # one-click chart image package
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for fn, b in chart_pack.items():
            zf.writestr(fn, b)
    st.download_button(
        "下載圖表圖檔（ZIP）",
        data=zip_buf.getvalue(),
        file_name=f"{datetime.now().strftime('%Y%m%d')}_圖表圖檔.zip",
        mime="application/zip",
    )
    ppt_bytes = build_ppt_bytes(
        stats_with_total, ai_text,
        st.session_state.get("source_name", "unknown"),
        chart_pack=chart_pack,
    )
    st.download_button(
        "一鍵下載分析簡報 PPT",
        data=ppt_bytes,
        file_name=f"{datetime.now().strftime('%Y%m%d')}_分析簡報.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


def section_3():
    st.subheader("功能三：歷史分析紀錄")
    history = load_history()
    if not history:
        st.info("尚無歷史紀錄。")
        return

    # De-duplicate: keep only latest entry per source_name
    seen_names: dict = {}
    deduped = []
    for item in history:
        sn = item.get("source_name", "")
        if sn not in seen_names:
            seen_names[sn] = item
            deduped.append(item)
        else:
            # Keep the newer one (history is already newest-first)
            pass
    history = deduped

    for item in history:
        out_path = Path(item["output_path"])
        if not out_path.exists():
            continue
        # Pure ASCII/CJK label — no emoji that Streamlit converts to text
        sname = item.get('source_name', '')
        if len(sname) > 28:
            sname = sname[:14] + "..." + sname[-10:]
        label = f"{item['created_at'][:16]}  {sname}  ({item['rows']} 筆)"
        with st.expander(label):
            df_hist = pd.read_excel(out_path)
            tab_data, tab_chart, tab_ai = st.tabs(["資料預覽", "圖表分析", "AI 重點摘要"])
            
            with tab_data:
                st.dataframe(df_hist.head(30), use_container_width=True, hide_index=True)
                col1, col2, col3 = st.columns([1, 1, 1])
                col1.download_button(
                    "下載該分析檔",
                    data=out_path.read_bytes(),
                    file_name=item["output_name"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{item['id']}",
                )
                if col2.button("[編輯]", key=f"edit_{item['id']}"):
                    st.session_state["analysis_df"] = df_hist.copy()
                    st.session_state["source_name"] = item["source_name"]
                    st.session_state["_editing_history_id"] = item["id"]
                    st.session_state["menu"] = "上傳檔案區（分析區）"
                    st.rerun()
                if col3.button("[刪除]", key=f"del_{item['id']}"):
                    delete_history(item["id"])
                    st.rerun()
            
            with tab_chart:
                if not df_hist.empty:
                    render_charts(df_hist, key_prefix=f"hist_{item['id']}")
                    cdl1, cdl2 = st.columns(2)
                    hist_stats = df_hist["問題類型"].value_counts().rename_axis("問題類型").reset_index(name="件數")
                    hist_stats["百分比"] = (hist_stats["件數"] / max(hist_stats["件數"].sum(), 1) * 100).round(0).astype(int)
                    hist_stats["歸屬部門"] = hist_stats["問題類型"].map(DEPT_MAP).fillna("")
                    hist_ai = generate_ai_summary(df_hist)
                    hist_chart_pack = build_chart_pack(df_hist)

                    hist_ppt = build_ppt_bytes(
                        hist_stats,
                        hist_ai,
                        item.get("source_name", "history"),
                        chart_pack=hist_chart_pack,
                    )
                    cdl1.download_button(
                        "一鍵下載PPT",
                        data=hist_ppt,
                        file_name=f"{datetime.now().strftime('%Y%m%d')}_{safe_filename(item.get('source_name','history'))}_圖表分析.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key=f"hist_ppt_{item['id']}",
                    )
                    hist_zip = io.BytesIO()
                    with zipfile.ZipFile(hist_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                        for fn, b in hist_chart_pack.items():
                            zf.writestr(fn, b)
                    cdl2.download_button(
                        "下載圖檔（ZIP）",
                        data=hist_zip.getvalue(),
                        file_name=f"{datetime.now().strftime('%Y%m%d')}_{safe_filename(item.get('source_name','history'))}_圖表.zip",
                        mime="application/zip",
                        key=f"hist_img_{item['id']}",
                    )
                else:
                    st.info("無資料可繪圖")
                    
            with tab_ai:
                st.info("點擊下方按鈕即時生成本檔案的 AI 重點摘要")
                if st.button("[產生 AI 摘要]", key=f"ai_btn_{item['id']}"):
                    with st.spinner("AI 分析中..."):
                        ai_result = generate_ai_summary_llm(df_hist)
                        st.markdown(ai_result)


def main():
    apply_brand_theme()
    st.markdown("<div class='ecoco-banner'>ECOCO 客訴智能分析平台</div>", unsafe_allow_html=True)
    with st.sidebar:
        st.markdown("<div class='side-title'>ECOCO AI</div>", unsafe_allow_html=True)
        st.markdown("<div class='side-sub'>客訴分析處理室</div>", unsafe_allow_html=True)
        if "menu" not in st.session_state:
            st.session_state["menu"] = "功能列表區"
        if st.button("🧩 功能列表區", use_container_width=True, type="primary" if st.session_state["menu"] == "功能列表區" else "secondary"):
            st.session_state["menu"] = "功能列表區"
        if st.button("📤 上傳檔案區（分析區）", use_container_width=True, type="primary" if st.session_state["menu"] == "上傳檔案區（分析區）" else "secondary"):
            st.session_state["menu"] = "上傳檔案區（分析區）"
        if st.button("📊 圖表與 AI 分析", use_container_width=True, type="primary" if st.session_state["menu"] == "圖表與 AI 分析" else "secondary"):
            st.session_state["menu"] = "圖表與 AI 分析"
        if st.button("🗂️ 歷史紀錄", use_container_width=True, type="primary" if st.session_state["menu"] == "歷史紀錄" else "secondary"):
            st.session_state["menu"] = "歷史紀錄"
        menu = st.session_state["menu"]

    if menu == "功能列表區":
        st.markdown(
            """
            <div class="ecoco-card">
              <b>功能 1</b>：上傳 excel/csv/pdf，分析並標記【問題類型、問題細項】；支援下拉選填、編輯、篩選、批次勾選編輯/刪除、下載 Excel、上傳 Google Sheet。
            </div>
            <div class="ecoco-card">
              <b>功能 2</b>：將分析結果圖表化，顯示各類型件數與百分比，並標示歸屬部門；可預覽與下載 AI 重點分析。
            </div>
            <div class="ecoco-card">
              <b>功能 3</b>：歷史分析紀錄管理（最新置頂），可預覽與下載。
            </div>
            """,
            unsafe_allow_html=True,
        )
    elif menu == "上傳檔案區（分析區）":
        section_1()
    elif menu == "圖表與 AI 分析":
        section_2()
    else:
        section_3()
        
    # Use a fixed-position div to stay at the absolute bottom of the viewport
    st.markdown(
        """
        <style>
            .fixed-footer {
                position: fixed;
                bottom: 15px;
                left: 0;
                width: 100%;
                text-align: center;
                color: #888888;
                font-size: 14px;
                z-index: 99;
                pointer-events: none; /* Don't block clicks to elements behind it */
            }
            /* Adjust for sidebar visibility if needed */
            @media (min-width: 768px) {
                .fixed-footer {
                    padding-left: 5rem; /* Offset slightly to be visually centered in the main area */
                }
            }
        </style>
        <div class="fixed-footer">
            202603© ECOCO宜可可循環經濟 客服課 ※ 請尊重智慧財產權 ※
        </div>
        """,
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
