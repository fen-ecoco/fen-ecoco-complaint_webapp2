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

# ── ECOCO 品牌色（Pantone 對應）──────────────────────────────
BRAND_ORANGE  = "#FF5000"   # Pantone Orange 021 C  → 營運部
BRAND_BLUE    = "#060E9F"   # Pantone Blue 072 C    → 資訊部 / 主圖色
BRAND_YELLOW  = "#FFCE00"   # Pantone 116 C         → 行銷部
BRAND_LBLUE   = "#8EB9C9"   # Pantone 550 C
BRAND_BEIGE   = "#FAE0B8"   # Pantone P17-2 C
BRAND_TEAL    = "#0076A9"   # Pantone 7690 C
BRAND_WHITE   = "#FFFFFF"   # Pantone White C

# 部門固定色（Plotly color_discrete_map 用）
DEPT_COLOR_MAP: dict[str, str] = {
    "營運部": BRAND_ORANGE,
    "行銷部": BRAND_YELLOW,
    "資訊部": BRAND_BLUE,
    "研發部": BRAND_TEAL,
    "廠務部": BRAND_LBLUE,
    "人資部": BRAND_BEIGE,
    "企劃部": "#A0C878",
    "財務部": "#C8A0E0",
    "開發部": "#E0C8A0",
    "總經理室": "#A0E0C8",
    "未分配":  "#CCCCCC",
    "":        "#CCCCCC",
}

# 圓餅圖 / 橫條圖單色排序
BRAND_PALETTE = [
    BRAND_BLUE, BRAND_ORANGE, BRAND_YELLOW,
    BRAND_LBLUE, BRAND_BEIGE, BRAND_TEAL,
]


# ── 啟動時確保 CJK 字型可用（下載備援）──────────────────────────
@st.cache_resource(show_spinner=False)
def _ensure_cjk_font() -> str:
    """回傳可用的 CJK 字型路徑；若系統沒有則下載到 /tmp。"""
    import os, glob
    CANDIDATES = [
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Medium.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/arphic/uming.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
        "/tmp/NotoSansCJK.ttc",
    ]
    CANDIDATES += glob.glob("/usr/share/fonts/**/NotoSansCJK*.ttc", recursive=True)
    found = next((p for p in CANDIDATES if os.path.exists(p)), None)
    if found:
        return found
    # 下載到 /tmp
    _dl = "/tmp/NotoSansCJK.ttc"
    URLS = [
        "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTC/NotoSansCJK-Regular.ttc",
        "https://github.com/notofonts/noto-cjk/raw/main/Sans/OTC/NotoSansCJK-Regular.ttc",
    ]
    for url in URLS:
        try:
            import urllib.request
            urllib.request.urlretrieve(url, _dl)
            if os.path.exists(_dl) and os.path.getsize(_dl) > 100_000:
                return _dl
        except Exception:
            continue
    return ""

HISTORY_DIR = Path("history_reports")
HISTORY_DIR.mkdir(exist_ok=True)
META_FILE = HISTORY_DIR / "history.json"

# 範本路徑：優先使用與程式同目錄的 簡報範本.pptx（已隨程式一起部署）
TEMPLATE_PATH = Path(__file__).parent / "簡報範本.pptx"


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


# ── Google Sheets 歷史紀錄持久化 ────────────────────────────────────────────
# Render 的磁碟每次重啟會清空；使用 Google Sheets 作為永久儲存後端。
# 需在 Streamlit Secrets 設定：
#   HISTORY_SHEET_ID = "<your_spreadsheet_id>"
#   [google_credentials]   ← service account JSON 欄位

def _get_gsheet_client():
    """從環境變數或 st.secrets 取得 gspread client。"""
    try:
        import gspread as _gs
        from google.oauth2.service_account import Credentials as _Creds
    except ImportError:
        return None
    try:
        import os, json as _json
        # ── 1. 優先讀 Render 環境變數 ──
        creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
        if creds_json:
            creds_dict = _json.loads(creds_json)
        else:
            # ── 2. 備用：本機 st.secrets（捕捉所有例外）──
            try:
                raw = st.secrets.get("google_credentials", {})
                creds_dict = dict(raw) if raw else {}
            except Exception:
                creds_dict = {}
        if not creds_dict:
            return None
        creds = _Creds.from_service_account_info(
            creds_dict,
            scopes=["https://spreadsheets.google.com/feeds",
                    "https://www.googleapis.com/auth/drive"],
        )
        return _gs.authorize(creds)
    except Exception:
        return None


def _history_sheet(log_error: bool = False):
    """回傳歷史紀錄工作表；失敗回傳 None。log_error=True 時把錯誤存入 session_state。"""
    import os
    client = _get_gsheet_client()
    if client is None:
        if log_error:
            st.session_state["_gsheet_error"] = "無法建立 Google API 連線（請確認 GOOGLE_CREDENTIALS_JSON 環境變數格式正確）"
        return None
    try:
        sid = os.environ.get("HISTORY_SHEET_ID", "").strip()
        if not sid:
            try:
                sid = str(st.secrets.get("HISTORY_SHEET_ID", "")).strip()
            except Exception:
                sid = ""
        if not sid:
            if log_error:
                st.session_state["_gsheet_error"] = "未設定 HISTORY_SHEET_ID 環境變數"
            return None
        ss = client.open_by_key(sid)
        try:
            ws = ss.worksheet("歷史紀錄")
            st.session_state.pop("_gsheet_error", None)
            return ws
        except Exception:
            ws = ss.add_worksheet("歷史紀錄", rows=500, cols=6)
            ws.append_row(["id", "created_at", "source_name", "rows", "excel_b64"])
            st.session_state.pop("_gsheet_error", None)
            return ws
    except Exception as e:
        err_str = str(e)
        if "PERMISSION_DENIED" in err_str or "403" in err_str:
            msg = (f"Google Sheets API 權限錯誤。請至 Google Cloud Console 確認已啟用：\n"
                   f"1. Google Sheets API\n2. Google Drive API\n"
                   f"錯誤：{err_str[:200]}")
        elif "NOT_FOUND" in err_str or "404" in err_str:
            msg = f"試算表不存在（ID 可能錯誤）：{err_str[:200]}"
        else:
            msg = f"Google Sheets 連線錯誤：{err_str[:300]}"
        if log_error:
            st.session_state["_gsheet_error"] = msg
        return None


def save_history(df: pd.DataFrame, source_name: str, existing_id: str = "") -> tuple[Path, str]:
    import base64
    today = datetime.now().strftime("%Y%m%d")
    ts = existing_id if existing_id else datetime.now().strftime("%Y%m%d_%H%M%S")
    output_name = f"{today}_分析.xlsx"
    excel_bytes = to_excel_bytes(df)
    excel_b64 = base64.b64encode(excel_bytes).decode()

    meta = {
        "id": ts, "created_at": datetime.now().isoformat(timespec="seconds"),
        "source_name": source_name, "output_name": output_name,
        "output_path": "", "rows": int(len(df)),
    }

    # 1. session_state 快取
    if "_history_cache" not in st.session_state:
        st.session_state["_history_cache"] = {}
    st.session_state["_history_cache"][ts] = {"meta": meta, "excel_bytes": excel_bytes}

    # 2. Google Sheets（永久）— 歷史索引 + 獨立工作表
    ws = _history_sheet()
    if ws:
        try:
            if existing_id:
                rows = ws.get_all_values()
                for i, row in enumerate(rows[1:], start=2):
                    if row and row[0] == existing_id:
                        ws.delete_rows(i); break
            ws.append_row([ts, meta["created_at"], source_name, str(len(df)), excel_b64])

            # ── 新增獨立分析工作表（每次創建唯一 sheet）──
            try:
                _ss = ws.spreadsheet
                _sheet_name = f"{ts[:8]}_{source_name[:12]}"
                # 清除非法字元
                import re as _re2
                _sheet_name = _re2.sub(r"[\[\]:*?/\\]", "_", _sheet_name)[:31]
                # 若同名已存在則刪除舊的
                try:
                    _old = _ss.worksheet(_sheet_name)
                    _ss.del_worksheet(_old)
                except Exception:
                    pass
                _new_ws = _ss.add_worksheet(title=_sheet_name, rows=max(len(df)+5, 50), cols=max(len(df.columns)+2, 20))
                # 寫入標題列
                _new_ws.append_row(df.columns.tolist())
                # 寫入資料（分批，每次200列）
                _rows_data = df.fillna("").astype(str).values.tolist()
                for _i in range(0, len(_rows_data), 200):
                    _new_ws.append_rows(_rows_data[_i:_i+200])
            except Exception:
                pass  # 獨立 sheet 失敗不影響主流程
        except Exception:
            pass

    # 3. 本機磁碟（輔助）
    output_path = HISTORY_DIR / f"{ts}_{output_name}"
    try:
        output_path.write_bytes(excel_bytes)
        history = []
        if META_FILE.exists():
            try: history = json.loads(META_FILE.read_text(encoding="utf-8"))
            except: pass
        history = [i for i in history if i["id"] != ts]
        history.insert(0, meta)
        META_FILE.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass
    return output_path, output_name


def load_history() -> list[dict]:
    import base64
    merged: dict[str, dict] = {}

    # 本機 JSON
    if META_FILE.exists():
        try:
            for item in json.loads(META_FILE.read_text(encoding="utf-8")):
                merged[item["id"]] = item
        except Exception:
            pass

    # Google Sheets（覆蓋本機，最可靠）
    ws = _history_sheet()
    if ws:
        try:
            for row in ws.get_all_values()[1:]:
                if not row or not row[0]:
                    continue
                rid = row[0]
                created_at = row[1] if len(row) > 1 else ""
                sname = row[2] if len(row) > 2 else ""
                rows_str = row[3] if len(row) > 3 else "0"
                excel_b64 = row[4] if len(row) > 4 else ""
                meta = {
                    "id": rid, "created_at": created_at,
                    "source_name": sname,
                    "rows": int(rows_str) if rows_str.isdigit() else 0,
                    "output_name": f"{rid}_分析.xlsx", "output_path": "",
                }
                merged[rid] = meta
                if "_history_cache" not in st.session_state:
                    st.session_state["_history_cache"] = {}
                if rid not in st.session_state["_history_cache"] and excel_b64:
                    try:
                        st.session_state["_history_cache"][rid] = {
                            "meta": meta,
                            "excel_bytes": base64.b64decode(excel_b64),
                        }
                    except Exception:
                        pass
        except Exception:
            pass

    # session_state 補充當次新增
    for rid, v in st.session_state.get("_history_cache", {}).items():
        if rid not in merged:
            merged[rid] = v["meta"]

    return sorted(merged.values(), key=lambda x: x.get("created_at", ""), reverse=True)


def safe_filename(text: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", str(text))


def delete_history(item_id: str):
    ws = _history_sheet()
    if ws:
        try:
            for i, row in enumerate(ws.get_all_values()[1:], start=2):
                if row and row[0] == item_id:
                    ws.delete_rows(i); break
        except Exception:
            pass
    if META_FILE.exists():
        try:
            history = json.loads(META_FILE.read_text(encoding="utf-8"))
            history = [i for i in history if i["id"] != item_id]
            META_FILE.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass
    cache = st.session_state.get("_history_cache", {})
    cache.pop(item_id, None)
    st.session_state["_history_cache"] = cache



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
    """Generate PDF using fpdf2 + Noto CJK for Traditional Chinese support."""
    from fpdf import FPDF
    from fpdf.enums import XPos, YPos
    import os, glob

    # ── 找字型：多重備援路徑 ──
    CJK_FONT_CANDIDATES = [
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Medium.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJKtc-Regular.otf",
        "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJKtc-Regular.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/arphic/uming.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    ]
    CJK_FONT_CANDIDATES += glob.glob("/usr/share/fonts/**/NotoSansCJK*.ttc", recursive=True)
    CJK_FONT_CANDIDATES += glob.glob("/usr/share/fonts/**/NotoSansCJK*.otf", recursive=True)
    font_path = next((p for p in CJK_FONT_CANDIDATES if os.path.exists(p)), None)
    # 使用 _ensure_cjk_font 確保有可用字型
    if not font_path:
        font_path = _ensure_cjk_font()

    table_df = df.copy()
    drop_cols = [c for c in ["選取"] if c in table_df.columns]
    table_df = table_df.drop(columns=drop_cols).fillna("")

    PAGE_W_MM = 277.0
    WIDE_COLS   = {"用戶內容", "主旨", "問題主旨"}
    MEDIUM_COLS = {"問題細項", "問題類型", "進件日期", "日期時間", "站點名稱", "問題細項"}
    num_cols = len(table_df.columns)
    wide_count   = sum(1 for c in table_df.columns if c in WIDE_COLS)
    medium_count = sum(1 for c in table_df.columns if c in MEDIUM_COLS)
    narrow_count = num_cols - wide_count - medium_count
    unit = PAGE_W_MM / max(wide_count*4 + medium_count*2 + narrow_count, 1)
    col_widths = {}
    for c in table_df.columns:
        if c in WIDE_COLS:       col_widths[c] = unit * 4
        elif c in MEDIUM_COLS:   col_widths[c] = unit * 2
        else:                    col_widths[c] = unit

    pdf = FPDF(orientation="L", format="A4")
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()

    if font_path:
        try:
            pdf.add_font("CJK", style="", fname=font_path)
            FONT = "CJK"
        except Exception:
            FONT = "Helvetica"
    else:
        FONT = "Helvetica"

    ROW_H = 7.0; HDR_H = 8.0; FS_HDR = 8; FS_CELL = 7

    def safe_text(s):
        if FONT == "Helvetica":
            return s.encode("ascii", "replace").decode()
        return s

    def fit_text_in_cell(pdf_obj, text, max_width, max_size=8, min_size=5):
        """縮小字型直到文字適合欄寬"""
        for fs in range(max_size, min_size-1, -1):
            pdf_obj.set_font(FONT, size=fs)
            if pdf_obj.get_string_width(text) <= max_width - 1:
                return fs
        return min_size

    # 表頭（自動縮小以適應欄寬）
    pdf.set_fill_color(0x06, 0x0E, 0x9F)
    pdf.set_text_color(255, 255, 255)
    for col in table_df.columns:
        cw = col_widths[col]
        header_text = safe_text(col)
        fs = fit_text_in_cell(pdf, header_text, cw, max_size=FS_HDR, min_size=4)
        pdf.set_font(FONT, size=fs)
        pdf.cell(cw, HDR_H, header_text, border=1, fill=True,
                 new_x=XPos.RIGHT, new_y=YPos.TOP, align="C")
    pdf.ln(HDR_H)

    # ── 資料列：框線統一 row_h 高度，文字疊加換行 ──
    pdf.set_font(FONT, size=FS_CELL)
    col_list = list(table_df.columns)

    for i, (_, row) in enumerate(table_df.iterrows()):
        fill_rgb = (0xEB, 0xF4, 0xFA) if i % 2 == 0 else (0xFF, 0xFF, 0xFF)

        # ── 準備文字（去除多餘空白與換行）──
        cell_texts = {
            col: safe_text(
                " ".join(str(row[col]).split())   # 多空白合併為單一空格
                .replace("  ", " ").strip()
            )
            for col in col_list
        }

        # ── 精確計算每欄行數 ──
        col_lines: dict[str, int] = {}
        for col in col_list:
            cw = col_widths[col] - 2
            text = cell_texts[col]
            if not text:
                col_lines[col] = 1
                continue
            n = 0
            for para in text.replace("\r", "").split("\n"):
                if not para:
                    n += 1
                    continue
                line_w = 0.0
                for ch in para:
                    ch_w = pdf.get_string_width(ch)
                    if line_w + ch_w > cw:
                        n += 1
                        line_w = ch_w
                    else:
                        line_w += ch_w
                n += 1
            col_lines[col] = max(1, n)

        max_lines = max(col_lines.values())
        row_h = max_lines * ROW_H

        # ── 換頁檢查 ──
        if pdf.get_y() + row_h > pdf.page_break_trigger:
            pdf.add_page()
            pdf.set_fill_color(0x06, 0x0E, 0x9F)
            pdf.set_text_color(255, 255, 255)
            for col in col_list:
                cw = col_widths[col]
                hdr_t = safe_text(col)
                fs = fit_text_in_cell(pdf, hdr_t, cw, max_size=FS_HDR, min_size=4)
                pdf.set_font(FONT, size=fs)
                pdf.cell(cw, HDR_H, hdr_t, border=1, fill=True,
                         new_x=XPos.RIGHT, new_y=YPos.TOP, align="C")
            pdf.ln(HDR_H)
            pdf.set_font(FONT, size=FS_CELL)

        x0 = pdf.get_x()
        y0 = pdf.get_y()
        pdf.set_text_color(0x22, 0x22, 0x22)

        # ── Step 1：先畫每欄的底色 + 完整框線（統一 row_h）──
        x_cursor = x0
        for col in col_list:
            cw = col_widths[col]
            # 底色填滿整格
            pdf.set_fill_color(*fill_rgb)
            pdf.rect(x_cursor, y0, cw, row_h, style="F")
            # 外框線（整格高度）
            pdf.set_draw_color(0x99, 0x99, 0x99)
            pdf.rect(x_cursor, y0, cw, row_h, style="D")
            x_cursor += cw

        # ── Step 2：疊加文字（multi_cell 只畫文字，不畫框線）──
        x_cursor = x0
        for col in col_list:
            cw = col_widths[col]
            val = cell_texts[col]
            pdf.set_xy(x_cursor + 0.5, y0 + 0.5)   # 0.5mm 內縮 padding
            pdf.set_fill_color(*fill_rgb)
            pdf.multi_cell(
                cw - 1, ROW_H, val,
                border=0,          # 不畫框線（已在 Step 1 畫好）
                align="L", fill=False,
                new_x=XPos.RIGHT, new_y=YPos.TOP,
                max_line_height=ROW_H,
            )
            x_cursor += cw

        pdf.set_draw_color(0, 0, 0)  # 恢復黑色
        pdf.set_xy(x0, y0 + row_h)

    # 頁尾
    pdf.set_y(-12)
    pdf.set_font(FONT, size=6)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(0, 6, safe_text(f"ECOCO 客訴分析報告  共 {len(table_df)} 筆  產出日期：{datetime.now().strftime('%Y/%m/%d')}"), align="C")

    return bytes(pdf.output())

def _setup_cjk_font() -> None:
    """設定 matplotlib 中文字型，使用 _ensure_cjk_font 取得字型路徑。"""
    import matplotlib.font_manager as fm
    import os

    # ── 已設定過就直接返回 ──
    current = plt.rcParams.get("font.family", "")
    if current and "sans-serif" not in str(current) and current != ["DejaVu Sans"]:
        return

    # ── 優先使用 _ensure_cjk_font 確保字型存在 ──
    fp = _ensure_cjk_font()
    if fp and os.path.exists(fp):
        try:
            fm.fontManager.addfont(fp)
            plt.rcParams["font.family"] = fm.FontProperties(fname=fp).get_name()
            plt.rcParams["axes.unicode_minus"] = False
            return
        except Exception:
            pass

    # ── 1. 已知路徑（Ubuntu / Render / Debian）──
    KNOWN_PATHS = [
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Medium.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJKtc-Regular.otf",
        "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJKtc-Regular.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/arphic/uming.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    ]
    for fp in KNOWN_PATHS:
        if os.path.exists(fp):
            try:
                fm.fontManager.addfont(fp)
                fname = fm.FontProperties(fname=fp).get_name()
                plt.rcParams["font.family"] = fname
                plt.rcParams["axes.unicode_minus"] = False
                return
            except Exception:
                continue

    # ── 2. 掃描所有已安裝字型找 CJK ──
    CJK_KEYWORDS = [
        "Noto Sans CJK", "Noto Serif CJK", "Noto Sans TC",
        "MingLiU", "PMingLiU", "Microsoft JhengHei",
        "WenQuanYi", "Droid Sans Fallback", "AR PL UMing",
    ]
    fm._load_fontmanager(try_read_cache=False)   # 強制重新掃描
    for kw in CJK_KEYWORDS:
        for f in fm.fontManager.ttflist:
            if kw.lower() in f.name.lower():
                plt.rcParams["font.family"] = f.name
                plt.rcParams["axes.unicode_minus"] = False
                return

    # ── 3. 最終 fallback：至少關掉負號亂碼 ──
    plt.rcParams["axes.unicode_minus"] = False


def build_chart_pack(df: pd.DataFrame,
                     color_bar: str | None = None,
                     color_pie: list[str] | None = None,
                     color_hbar: str | None = None) -> dict[str, bytes]:
    """Build chart PNG images for download/PPT.
    color_bar  : 問題類型直條圖 — None = 依部門品牌色; 或傳入單一 hex 強制套用
    color_pie  : 機台圓餅圖各扇形顏色 list，None = BRAND_PALETTE
    color_hbar : 十大細項橫條圖顏色，None = BRAND_BLUE
    """
    _setup_cjk_font()

    data = df.copy()
    # ── 機台類型正規化：方舟/方舟站 → 收瓶機 ──
    if "機台類型" in data.columns:
        data["機台類型"] = data["機台類型"].apply(
            lambda v: "收瓶機" if ("方舟" in str(v) or "收瓶" in str(v))
            else ("電池機" if "電池" in str(v) else str(v))
        )
    stats = data["問題類型"].value_counts().rename_axis("問題類型").reset_index(name="件數")
    stats["百分比"] = (stats["件數"] / max(stats["件數"].sum(), 1) * 100).round(1)
    detail_stats = data["問題細項"].value_counts().reset_index().head(10)
    detail_stats.columns = ["問題細項", "件數"]
    d = detail_stats.sort_values("件數", ascending=True)

    # ── resolve colors ──
    _pie_palette  = color_pie  if color_pie  else BRAND_PALETTE
    _hbar_color   = color_hbar if color_hbar else BRAND_BLUE

    def _bar_colors_for(series):
        if color_bar:
            return [color_bar] * len(series)
        return [DEPT_COLOR_MAP.get(DEPT_MAP.get(t, ""), BRAND_ORANGE) for t in series]

    # 1) 問題類型直條圖
    fig1, ax1 = plt.subplots(figsize=(8, 4.5))
    bc = _bar_colors_for(stats["問題類型"])
    ax1.bar(stats["問題類型"], stats["件數"], color=bc)
    ax1.set_title("問題類型分布")
    ax1.set_ylabel("件數")
    ax1.yaxis.set_major_locator(plt.MaxNLocator(integer=True))
    ax1.tick_params(axis="x", rotation=20)
    for i, r in stats.iterrows():
        ax1.text(i, r["件數"], f'{int(r["百分比"])}%', ha="center", va="bottom", fontsize=9)
    fig1.tight_layout()
    b1 = io.BytesIO(); fig1.savefig(b1, format="png", dpi=180); plt.close(fig1)

    # 2) 機台圓餅圖
    fig2, ax2 = plt.subplots(figsize=(6.2, 4.5))
    df_machine = data[data["問題類型"] == "機台問題類型"].copy()
    if df_machine.empty:
        ax2.text(0.5, 0.5, "無機台相關資料", ha="center", va="center", transform=ax2.transAxes)
        pie_counts = None
    else:
        def _get_mtype(row):
            txt = str(row.get("用戶內容", "")) + " " + str(row.get("主旨", ""))
            if "方舟" in txt: return "方舟站"
            if "電池" in txt: return "電池機"
            return "收瓶機"
        df_machine["機台機型"] = df_machine.apply(_get_mtype, axis=1)
        pie_counts = df_machine["機台機型"].value_counts()
        pc = _pie_palette[:len(pie_counts)]
        wedges, texts, autotexts = ax2.pie(
            pie_counts.values, labels=pie_counts.index, autopct="%1.1f%%",
            colors=pc, wedgeprops=dict(linewidth=1.5, edgecolor="white"),
        )
        for at in autotexts: at.set_fontsize(10)
    ax2.set_title("機台問題類型分布")
    fig2.tight_layout()
    b2 = io.BytesIO(); fig2.savefig(b2, format="png", dpi=180); plt.close(fig2)

    # 3) 十大細項橫條圖  ── 強制品牌主藍 #060E9F，整數刻度
    fig3, ax3 = plt.subplots(figsize=(8, 4.5))
    _hbar = _hbar_color if _hbar_color else "#060E9F"
    ax3.barh(d["問題細項"], d["件數"], color=_hbar)
    ax3.set_title("十大問題細項分布")
    ax3.set_xlabel("件數")
    # 強制整數刻度（件數必為整數）
    from matplotlib.ticker import MultipleLocator
    ax3.xaxis.set_major_locator(MultipleLocator(1))
    ax3.xaxis.set_minor_locator(MultipleLocator(1))
    ax3.set_xlim(left=0)
    fig3.tight_layout()
    b3 = io.BytesIO(); fig3.savefig(b3, format="png", dpi=180); plt.close(fig3)

    # 4) Dashboard 合圖
    fig4 = plt.figure(figsize=(14, 5))
    gs = fig4.add_gridspec(1, 3)
    a1 = fig4.add_subplot(gs[0, 0])
    a2 = fig4.add_subplot(gs[0, 1])
    a3 = fig4.add_subplot(gs[0, 2])
    a1.bar(stats["問題類型"], stats["件數"], color=bc)
    a1.set_title("問題類型分布")
    a1.yaxis.set_major_locator(MultipleLocator(1))
    a1.tick_params(axis="x", rotation=18)
    if pie_counts is None:
        a2.text(0.5, 0.5, "無機台資料", ha="center", va="center", transform=a2.transAxes)
    else:
        a2.pie(pie_counts.values, labels=pie_counts.index, autopct="%1.1f%%",
               colors=_pie_palette[:len(pie_counts)],
               wedgeprops=dict(linewidth=1.5, edgecolor="white"))
    a2.set_title("機台問題占比")
    a3.barh(d["問題細項"], d["件數"], color=_hbar)
    a3.xaxis.set_major_locator(MultipleLocator(1))
    a3.set_xlim(left=0)
    a3.set_title("十大細項")
    fig4.tight_layout()
    b4 = io.BytesIO(); fig4.savefig(b4, format="png", dpi=180); plt.close(fig4)

    return {
        "chart_問題類型分布.png": b1.getvalue(),
        "chart_機台問題占比.png": b2.getvalue(),
        "chart_十大問題細項.png": b3.getvalue(),
        "chart_dashboard.png":    b4.getvalue(),
    }


def build_ppt_bytes(stats: pd.DataFrame, ai_text: str, source_name: str,
                    template_path: str = "",
                    chart_pack: Optional[dict[str, bytes]] = None) -> bytes:
    """
    Build a PPT presentation.
    優先使用同目錄的範本；若找不到則從零構建符合 ECOCO 品牌風格的投影片。
    """
    from pptx.util import Emu, Inches, Pt
    from pptx.enum.text import PP_ALIGN

    BLUE   = RGBColor(0x06, 0x0E, 0x9F)
    ORANGE = RGBColor(0xFF, 0x50, 0x00)
    WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
    BEIGE  = RGBColor(0xFA, 0xE0, 0xB8)
    DARK   = RGBColor(0x22, 0x22, 0x22)
    LGRAY  = RGBColor(0xE8, 0xF1, 0xF5)
    FONT   = "MingLiU"   # 細明體

    # ── 嘗試載入範本 ──
    _tpath = Path(template_path) if template_path else TEMPLATE_PATH
    use_template = _tpath.exists()
    prs = Presentation(str(_tpath)) if use_template else Presentation()
    if not use_template:
        # 設定投影片大小為寬螢幕 16:9
        prs.slide_width  = Inches(13.33)
        prs.slide_height = Inches(7.5)

    SW = prs.slide_width
    SH = prs.slide_height

    # ── 小工具 ──
    def blank_layout():
        for lay in prs.slide_layouts:
            if lay.name.lower() in ("blank", "空白"):
                return lay
        return prs.slide_layouts[-1]

    def add_rect(slide, l, t, w, h, fill_rgb, line=False):
        shp = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill_rgb
        if not line:
            shp.line.fill.background()
        return shp

    def add_text(slide, text, l, t, w, h, font_size, bold=False,
                 color=None, align=PP_ALIGN.LEFT, wrap=True):
        txb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
        tf  = txb.text_frame
        tf.word_wrap = wrap
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.name  = FONT
        run.font.size  = Pt(font_size)
        run.font.bold  = bold
        run.font.color.rgb = color or DARK
        return txb

    def add_img(slide, img_bytes, l, t, w, h):
        if img_bytes:
            try:
                slide.shapes.add_picture(io.BytesIO(img_bytes),
                    Inches(l), Inches(t), Inches(w), Inches(h))
            except Exception:
                pass

    def add_header(slide, title_text, subtitle_text=""):
        """加上 ECOCO 品牌頁首（藍色長條 + 標題）"""
        add_rect(slide, 0, 0, SW/914400, 1.05, BLUE)
        add_text(slide, title_text, 0.3, 0.08, 9.0, 0.55,
                 20, bold=True, color=WHITE)
        if subtitle_text:
            add_text(slide, subtitle_text, 0.3, 0.62, 10.0, 0.38,
                     11, color=BEIGE)

    def delete_shape(sp):
        sp.element.getparent().remove(sp.element)

    # ════════════════════════════════════════════════════════
    #  使用範本：覆寫文字、表格、圖片
    # ════════════════════════════════════════════════════════
    if use_template:
        slides = list(prs.slides)

        # --- 封面 (slide 0) ---
        # 範本版面：左側藍色面板（版面配置提供），右側三個文字框：
        #   Shape;99  → 主標題「營運周報」 (l≈5.66, t≈2.18)
        #   Shape;98  → 日期/資料列 (l≈6.14, t≈3.48)
        #   Shape;96  → 公司名藍底白字 (l≈6.67, t≈5.04)
        s0 = slides[0]
        for sp in s0.shapes:
            if not sp.has_text_frame:
                continue
            l_in = sp.left / 914400
            t_in = sp.top  / 914400
            raw  = sp.text_frame.text.strip()

            # ── 主標題（在 x>5" 且 y<3"）
            if l_in > 5.0 and t_in < 3.0:
                tf = sp.text_frame
                tf.clear()
                p = tf.paragraphs[0]
                run = p.add_run()
                run.text = "客訴分析簡報"
                run.font.name  = FONT
                run.font.bold  = True
                run.font.size  = Pt(32)
                run.font.color.rgb = RGBColor(0x16, 0x2B, 0x7E)

            # ── 日期/資料欄（在 x>5" 且 y 在 3~5"）
            elif l_in > 5.0 and 3.0 <= t_in < 5.0:
                tf = sp.text_frame
                tf.clear()
                for label, val in [
                    ("報告日期", datetime.now().strftime("%Y/%m/%d")),
                    ("報告資料", source_name),
                ]:
                    p = tf.add_paragraph()
                    run = p.add_run()
                    run.text = f"{label}:{val}"
                    run.font.name  = FONT
                    run.font.bold  = True
                    run.font.size  = Pt(18)
                    run.font.color.rgb = RGBColor(0x1A, 0x2A, 0x7F)

            # ── 公司名（在 x>6" 且 y>=5" 或有填色藍底）
            elif l_in > 6.0 and t_in >= 4.8:
                pass   # 保留原樣「凡立橙股份有限公司」

        def _fill_slide(slide, title_txt, chart_key_list, add_table=True):
            SWi = prs.slide_width  / 914400
            SHi = prs.slide_height / 914400

            # 更新標題文字（比對關鍵字）
            for sp in slide.shapes:
                if sp.has_text_frame:
                    txt = sp.text_frame.text
                    if any(k in txt for k in ("客訴問題分析", "機台問題佔比", "機台與細項",
                                               "客訴問題", "問題分析", "20260")):
                        tf = sp.text_frame; tf.clear()
                        p = tf.paragraphs[0]
                        run = p.add_run()
                        run.text = title_txt
                        run.font.name = FONT
                        run.font.bold = True
                        run.font.size = Pt(16)
                        run.font.color.rgb = BLUE

            # 收集現有 Table / Picture 位置後刪除（清空舊內容）
            tbl_rect = None
            pic_rects = []
            for sp in list(slide.shapes):
                if sp.shape_type == 19:   # Table
                    tbl_rect = (sp.left, sp.top, sp.width, sp.height)
                    delete_shape(sp)
                elif sp.shape_type == 13:  # Picture
                    pic_rects.append((sp.left, sp.top, sp.width, sp.height))
                    delete_shape(sp)
            pic_rects.sort(key=lambda x: x[0])

            # ── 圖表插入：優先使用範本佔位位置，否則用固定座標 ──
            if chart_pack:
                if add_table:
                    # slide 2（問題分析）：表格左半 + 圖表右半
                    # 固定座標：圖表放右側
                    chart_fixed = [
                        (6.2, 1.15, SWi - 6.5, SHi - 1.4),   # 問題類型分布
                    ]
                else:
                    # slide 3（機台細項）：左右各放一張圖
                    chart_fixed = [
                        (0.3,              1.15, (SWi - 0.6) / 2,       SHi - 1.4),
                        (0.3 + (SWi-0.6)/2 + 0.15, 1.15, (SWi-0.6)/2, SHi - 1.4),
                    ]

                for idx, key in enumerate(chart_key_list):
                    if key not in chart_pack:
                        continue
                    if idx < len(pic_rects):
                        # 範本有佔位圖片 → 用原始位置
                        add_img(slide, chart_pack[key],
                                *[v / 914400 for v in pic_rects[idx]])
                    elif idx < len(chart_fixed):
                        # 範本沒有佔位 → 用固定座標
                        add_img(slide, chart_pack[key], *chart_fixed[idx])

            # ── 重建資料表格 ──
            if add_table:
                # 如果範本有舊表格位置就沿用，否則預設左側
                if tbl_rect:
                    tb_l, tb_t, tb_w, tb_h = tbl_rect
                else:
                    tb_l = Inches(0.25)
                    tb_t = Inches(1.15)
                    tb_w = Inches(5.8)
                    tb_h = Inches(SHi - 1.4)
                rows_n = min(len(stats) + 1, 12)
                tb = slide.shapes.add_table(rows_n, 4, tb_l, tb_t, tb_w, tb_h).table
                col_ws = [Inches(2.4), Inches(0.8), Inches(1.0), Inches(1.5)]
                for ci, cw in enumerate(col_ws):
                    tb.columns[ci].width = cw
                for ci, hdr in enumerate(["問題類型", "件數", "百分比", "歸屬部門"]):
                    cell = tb.cell(0, ci)
                    cell.text = hdr
                    cell.fill.solid(); cell.fill.fore_color.rgb = BLUE
                    for para in cell.text_frame.paragraphs:
                        para.alignment = PP_ALIGN.CENTER
                        for run in para.runs:
                            run.font.bold  = True
                            run.font.color.rgb = WHITE
                            run.font.size  = Pt(13)
                            run.font.name  = FONT
                for ri, (_, r) in enumerate(stats.head(rows_n - 1).iterrows(), 1):
                    try:   pct = f'{int(float(r["百分比"]))}%'
                    except: pct = f'{r["百分比"]}%'
                    dept = str(r.get("歸屬部門", ""))
                    vals = [str(r["問題類型"]), str(int(r["件數"])), pct, dept]
                    # 依部門套用品牌色為列底色
                    dept_hex = DEPT_COLOR_MAP.get(dept, "")
                    if dept_hex:
                        r_bg = RGBColor(
                            int(dept_hex[1:3], 16),
                            int(dept_hex[3:5], 16),
                            int(dept_hex[5:7], 16),
                        )
                        # 淡化：混入白色 80%
                        r_bg = RGBColor(
                            min(255, int(r_bg[0] * 0.25 + 255 * 0.75)),
                            min(255, int(r_bg[1] * 0.25 + 255 * 0.75)),
                            min(255, int(r_bg[2] * 0.25 + 255 * 0.75)),
                        )
                    else:
                        r_bg = LGRAY if ri % 2 == 0 else BEIGE
                    for ci, v in enumerate(vals):
                        cell = tb.cell(ri, ci)
                        cell.text = v
                        cell.fill.solid(); cell.fill.fore_color.rgb = r_bg
                        for para in cell.text_frame.paragraphs:
                            para.alignment = PP_ALIGN.CENTER
                            for run in para.runs:
                                run.font.size  = Pt(12)
                                run.font.color.rgb = DARK
                                run.font.name  = FONT

        if len(slides) >= 2:
            _fill_slide(slides[1],
                        f"{source_name} 客訴問題分析",
                        ["chart_問題類型分布.png"],
                        add_table=True)
        if len(slides) >= 3:
            _fill_slide(slides[2],
                        f"{source_name} 機台與細項分析",
                        ["chart_十大問題細項.png", "chart_機台問題占比.png"],
                        add_table=False)

    # ════════════════════════════════════════════════════════
    #  從零構建（範本不存在時）
    # ════════════════════════════════════════════════════════
    else:
        SWi = SW / 914400   # EMU → inches
        SHi = SH / 914400

        # ── Slide 1: 封面 ──
        s0 = prs.slides.add_slide(blank_layout())
        add_rect(s0, 0, 0, SWi, SHi, BLUE)      # 全藍背景
        add_text(s0, "ECOCO 客訴分析簡報",
                 1.0, SHi*0.25, SWi-2, 1.2, 36, bold=True,
                 color=WHITE, align=PP_ALIGN.CENTER)
        add_text(s0, f"報告日期：{datetime.now().strftime('%Y/%m/%d')}",
                 1.0, SHi*0.52, SWi-2, 0.5, 16,
                 color=BEIGE, align=PP_ALIGN.CENTER)
        add_text(s0, f"資料來源：{source_name}",
                 1.0, SHi*0.64, SWi-2, 0.5, 14,
                 color=BEIGE, align=PP_ALIGN.CENTER)
        add_text(s0, "凡立橙股份有限公司",
                 1.0, SHi*0.82, SWi-2, 0.4, 13,
                 color=WHITE, align=PP_ALIGN.CENTER)

        # ── Slide 2: 問題類型分析 ──
        s1 = prs.slides.add_slide(blank_layout())
        add_header(s1, f"客訴問題分析 — {source_name}",
                   f"報告日期：{datetime.now().strftime('%Y/%m/%d')}　資料來源：{source_name}")
        # 表格（左半）
        rows_n = min(len(stats) + 1, 10)
        tbl_left = Inches(0.3); tbl_top = Inches(1.15)
        tbl_w    = Inches(5.8); tbl_h   = Inches(SHi - 1.4)
        tb = s1.shapes.add_table(rows_n, 4, tbl_left, tbl_top, tbl_w, tbl_h).table
        tb.columns[0].width = Inches(2.2)
        tb.columns[1].width = Inches(0.9)
        tb.columns[2].width = Inches(1.0)
        tb.columns[3].width = Inches(1.5)
        for ci, hdr in enumerate(["問題類型", "件數", "百分比", "歸屬部門"]):
            c = tb.cell(0, ci); c.text = hdr
            c.fill.solid(); c.fill.fore_color.rgb = BLUE
            for para in c.text_frame.paragraphs:
                para.alignment = PP_ALIGN.CENTER
                for run in para.runs:
                    run.font.bold = True; run.font.color.rgb = WHITE
                    run.font.size = Pt(12); run.font.name = FONT
        for ri, (_, r) in enumerate(stats.head(rows_n - 1).iterrows(), 1):
            try:   pct = f'{int(float(r["百分比"]))}%'
            except: pct = f'{r["百分比"]}%'
            vals = [str(r["問題類型"]), str(r["件數"]), pct,
                    str(r.get("歸屬部門", ""))]
            bg = LGRAY if ri % 2 == 0 else BEIGE
            for ci, v in enumerate(vals):
                c = tb.cell(ri, ci); c.text = v
                c.fill.solid(); c.fill.fore_color.rgb = bg
                for para in c.text_frame.paragraphs:
                    para.alignment = PP_ALIGN.CENTER
                    for run in para.runs:
                        run.font.size = Pt(11); run.font.color.rgb = DARK
                        run.font.name = FONT
        # 圖表（右半）
        if chart_pack and "chart_問題類型分布.png" in chart_pack:
            add_img(s1, chart_pack["chart_問題類型分布.png"],
                    6.25, 1.15, SWi - 6.55, SHi - 1.4)

        # ── Slide 3: 機台與細項分析 ──
        s2 = prs.slides.add_slide(blank_layout())
        add_header(s2, f"機台與細項分析 — {source_name}",
                   f"報告日期：{datetime.now().strftime('%Y/%m/%d')}")
        half_w = (SWi - 0.6) / 2
        ch_t = 1.15; ch_h = SHi - 1.4
        if chart_pack and "chart_機台問題占比.png" in chart_pack:
            add_img(s2, chart_pack["chart_機台問題占比.png"],
                    0.3, ch_t, half_w, ch_h)
        if chart_pack and "chart_十大問題細項.png" in chart_pack:
            add_img(s2, chart_pack["chart_十大問題細項.png"],
                    0.3 + half_w + 0.15, ch_t, half_w, ch_h)

    # ── 最終：AI 重點分析投影片（所有路徑都加）──
    s_ai = prs.slides.add_slide(blank_layout())
    SWi2 = prs.slide_width  / 914400
    SHi2 = prs.slide_height / 914400
    # 藍色頁首
    add_rect(s_ai, 0, 0, SWi2, 1.05, BLUE)
    add_text(s_ai, "AI 重點問題分析",
             0.3, 0.08, 9.0, 0.55, 20, bold=True, color=WHITE)
    add_text(s_ai,
             f"資料來源：{source_name}　產出日期：{datetime.now().strftime('%Y/%m/%d')}",
             0.3, 0.65, 10.5, 0.35, 11, color=BEIGE)
    # 橘色左邊框裝飾
    add_rect(s_ai, 0.25, 1.15, 0.08, SHi2 - 1.35, ORANGE)
    # AI 文字框
    txb = s_ai.shapes.add_textbox(Inches(0.45), Inches(1.2),
                                   Inches(SWi2 - 0.65), Inches(SHi2 - 1.35))
    tf = txb.text_frame; tf.word_wrap = True
    first = True
    for line in ai_text.split('\n'):
        line = line.strip()
        if not line:
            continue
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_before = Pt(4)
        is_head = line[:2] in ('1)', '2)', '3)', '4)', '5)', '一、', '二、', '三、')
        run = p.add_run()
        run.text = line
        run.font.name  = FONT
        run.font.size  = Pt(14 if is_head else 13)
        run.font.bold  = is_head
        run.font.color.rgb = BLUE if is_head else DARK

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()



def upload_to_google_sheet(df: pd.DataFrame, credentials_json: dict, spreadsheet_id: str, worksheet_name: str) -> None:
    import gspread as _gs
    from google.oauth2.service_account import Credentials as _Creds
    # 必須同時包含 spreadsheets 和 drive scope
    scopes = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = _Creds.from_service_account_info(credentials_json, scopes=scopes)
    client = _gs.authorize(creds)
    try:
        sh = client.open_by_key(spreadsheet_id)
    except Exception as e:
        raise PermissionError(
            f"無法存取試算表（ID: {spreadsheet_id}）。\n"
            f"請確認已將試算表共用給：{credentials_json.get('client_email', '?')}\n"
            f"原始錯誤：{e}"
        )
    try:
        ws = sh.worksheet(worksheet_name)
        ws.clear()
    except Exception:
        ws = sh.add_worksheet(title=worksheet_name, rows=1000, cols=30)
    values = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
    ws.update(values)
    return ws.url if hasattr(ws, 'url') else ""


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

    # ── 欄位顯示控制（隱藏/釘選/刪除）────────────────────────────
    if "hidden_cols" not in st.session_state:
        st.session_state["hidden_cols"] = set()
    if "pinned_cols" not in st.session_state:
        st.session_state["pinned_cols"] = []

    with st.expander("⚙️ 欄位設定（隱藏 / 釘選 / 刪除）", expanded=False):
        all_data_cols = [c for c in show.columns if c not in ("選取", MARKER_COL)]
        col_chunks = [all_data_cols[i:i+5] for i in range(0, len(all_data_cols), 5)]
        for chunk in col_chunks:
            ck_cols = st.columns(len(chunk))
            for ci, col_name in enumerate(chunk):
                with ck_cols[ci]:
                    st.markdown(f"**{col_name}**")
                    h_key = f"hide_{col_name}"
                    p_key = f"pin_{col_name}"
                    d_key = f"del_{col_name}"
                    if st.checkbox("隱藏", key=h_key,
                                   value=col_name in st.session_state["hidden_cols"]):
                        st.session_state["hidden_cols"].add(col_name)
                    else:
                        st.session_state["hidden_cols"].discard(col_name)
                    if st.checkbox("釘選", key=p_key,
                                   value=col_name in st.session_state["pinned_cols"]):
                        if col_name not in st.session_state["pinned_cols"]:
                            st.session_state["pinned_cols"].append(col_name)
                    else:
                        if col_name in st.session_state["pinned_cols"]:
                            st.session_state["pinned_cols"].remove(col_name)
                    if st.button("🗑 刪除欄", key=d_key):
                        for target_df in ["analysis_df"]:
                            if target_df in st.session_state and col_name in st.session_state[target_df].columns:
                                st.session_state[target_df].drop(columns=[col_name], inplace=True)
                        st.session_state["hidden_cols"].discard(col_name)
                        if col_name in st.session_state["pinned_cols"]:
                            st.session_state["pinned_cols"].remove(col_name)
                        st.rerun()

    # ── 依設定重排欄位：釘選在前，其次一般，隱藏排除 ──────────────
    pinned = [c for c in st.session_state["pinned_cols"] if c in show_display.columns]
    others = [c for c in show_display.columns
              if c not in pinned and c not in st.session_state["hidden_cols"]]
    display_order = ["選取", MARKER_COL] + pinned + [c for c in others if c not in ("選取", MARKER_COL)]
    display_order = [c for c in display_order if c in show_display.columns]   # 確保欄位存在
    show_display_ordered = show_display[display_order]

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
    if "選取" in show_display.columns:
        insert_idx = show_display.columns.get_loc("選取") + 1
    show_display.insert(insert_idx, MARKER_COL, marker_vals)

    # --- Select All Trigger ---
    cols_h = st.columns([13, 2])
    if cols_h[1].button("⬓ 選取 / 取消", key="toggle_all_btn", help="全選或取消全選"):
        all_sel = bool(df["選取"].all()) if "選取" in df.columns and not df.empty else False
        st.session_state["analysis_df"]["選取"] = not all_sel
        st.rerun()

    edited = st.data_editor(
        show_display_ordered,
        use_container_width=True,
        num_rows="dynamic",
        hide_index=True,
        column_config={
            "選取": st.column_config.CheckboxColumn("選取", help="勾選要批次處理的列"),
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

    st.markdown("##### 批次處理與儲存")

    b1, b2, b3, b4 = st.columns([2, 2, 2, 2])
    batch_type   = b1.selectbox("批次問題類型", ["(不變更)"] + TYPE_OPTIONS, key="batch_type_sel")
    valid_batch_det = ["(不變更)"]
    if batch_type != "(不變更)":
        valid_batch_det += TOPIC_DETAIL_MAP.get(batch_type, [])
    batch_detail = b2.selectbox("批次問題細項", valid_batch_det, key="batch_cat_sel")

    if b3.button("將上方設定套用到所有勾選列", type="primary"):
        if "選取" not in edited.columns or not edited["選取"].any():
            st.warning("請先在表格內勾選要處理的資料列！")
        else:
            full_df = st.session_state["analysis_df"].copy()
            save_edited = edited.drop(columns=["選取", MARKER_COL], errors="ignore")
            # 只更新存在於 full_df 的欄位，避免因隱藏欄位造成欄位消失
            common_cols = [c for c in save_edited.columns if c in full_df.columns]
            full_df.update(save_edited[common_cols])

            mask_idx = edited.index[edited["選取"] == True]
            if batch_type != "(不變更)":
                full_df.loc[mask_idx, "問題類型"] = batch_type
                full_df.loc[mask_idx, "部門"] = DEPT_MAP.get(batch_type, "")
            if batch_detail != "(不變更)":
                full_df.loc[mask_idx, "問題細項"] = batch_detail
            full_df["問題細項"] = full_df.apply(
                lambda r: r["問題細項"]
                if r["問題細項"] in TOPIC_DETAIL_MAP.get(r["問題類型"], [])
                else TOPIC_DETAIL_MAP.get(r["問題類型"], ["其他建議"])[0],
                axis=1,
            )
            st.session_state["analysis_df"] = full_df
            # ── 保留隱藏/釘選狀態（不重置）──
            st.session_state["_batch_applied"] = True
            # 不呼叫 st.rerun() 而是直接更新，避免清空 hidden_cols/pinned_cols
            st.rerun()

    if st.session_state.pop("_batch_applied", False):
        st.success("✅ 已套用批次編輯，勾選狀態與欄位隱藏/釘選設定均已保留。")

    if b4.button("刪除勾選列"):
        if "選取" not in edited.columns or not edited["選取"].any():
            st.warning("請先在表格內勾選要刪除的資料列！")
        else:
            full_df = st.session_state["analysis_df"].copy()
            save_edited = edited.drop(columns=["選取", MARKER_COL], errors="ignore")
            full_df.update(save_edited)
            full_df = full_df.drop(index=edited.index[edited["選取"] == True]).reset_index(drop=True)
            st.session_state["analysis_df"] = full_df
            st.success("已刪除勾選列。")
            st.rerun()

    # ── 刪除整欄功能 ──────────────────────────────────────────────
    with st.expander("🗑️ 刪除整欄（永久移除不需要的欄位）", expanded=False):
        PROTECTED = {"選取", "問題類型", "問題細項", "部門", "_ai_filled", MARKER_COL}
        deletable = [c for c in st.session_state["analysis_df"].columns if c not in PROTECTED]
        if deletable:
            del_col_sel = st.multiselect("選擇要刪除的欄位", deletable, key="del_col_ms")
            if st.button("確認刪除所選欄位", key="del_col_btn", type="primary"):
                if del_col_sel:
                    st.session_state["analysis_df"] = st.session_state["analysis_df"].drop(
                        columns=del_col_sel, errors="ignore"
                    )
                    st.success(f"已刪除欄位：{', '.join(del_col_sel)}")
                    st.rerun()
        else:
            st.info("無可刪除的欄位（核心分析欄位受保護）")

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
                try:
                    credentials_json = json.loads(cred_file.getvalue().decode("utf-8"))
                    upload_to_google_sheet(final_df, credentials_json, spreadsheet_id, ws_name)
                    st.success(f"✅ 已上傳到 Google Sheet 工作表：{ws_name}")
                    st.info(f"📋 Service Account：{credentials_json.get('client_email', '')}")
                except PermissionError as e:
                    st.error(str(e))
                    st.warning("👉 請到 Google 試算表右上角「共用」，加入上方 Service Account email，並給予「編輯者」權限")
                except Exception as e:
                    st.error(f"上傳失敗：{e}")


def render_charts_from_stats(stats: pd.DataFrame, df: pd.DataFrame, key_prefix: str = ""):
    """Render interactive Plotly charts with per-chart color pickers."""

    # ── 顏色設定 expander ──────────────────────────────────────────
    kp = key_prefix or "main"
    with st.expander("🎨 調整圖表顏色（可個別修改）", expanded=False):
        ca, cb, cc = st.columns(3)
        # 問題類型直條圖：預設「依部門品牌色」，勾選後可指定單色
        use_single_bar = ca.checkbox("直條圖使用單一顏色", key=f"{kp}_cb_bar")
        c_bar_single   = ca.color_picker("直條圖顏色", value=BRAND_ORANGE, key=f"{kp}_cp_bar") if use_single_bar else None

        # 圓餅圖：最多3個扇形獨立調色
        pie_c1 = cb.color_picker("圓餅圖 第1色（主）", value=BRAND_BLUE,   key=f"{kp}_cp_pie1")
        pie_c2 = cb.color_picker("圓餅圖 第2色（次）", value=BRAND_ORANGE, key=f"{kp}_cp_pie2")
        pie_c3 = cb.color_picker("圓餅圖 第3色",       value=BRAND_LBLUE,  key=f"{kp}_cp_pie3")

        c_hbar = cc.color_picker("細項橫條圖顏色", value=BRAND_BLUE, key=f"{kp}_cp_hbar")

    custom_pie   = [pie_c1, pie_c2, pie_c3] + BRAND_PALETTE[3:]
    custom_hbar  = c_hbar

    # ── 圓餅圖資料（供 Plotly + matplotlib 共用）────────────────
    df_machine = df[df["問題類型"] == "機台問題類型"].copy()
    m_stats = None
    if not df_machine.empty:
        def _gmt(row):
            txt = str(row.get("用戶內容", "")) + " " + str(row.get("主旨", ""))
            if "方舟" in txt: return "方舟站"
            if "電池" in txt: return "電池機"
            return "收瓶機"
        df_machine["機台機型"] = df_machine.apply(_gmt, axis=1)
        m_stats = df_machine["機台機型"].value_counts().reset_index()
        m_stats.columns = ["機型", "件數"]

    detail_stats = df["問題細項"].value_counts().reset_index().head(10)
    detail_stats.columns = ["問題細項", "件數"]

    c1, c2, c3 = st.columns(3)

    # ── 圖1：問題類型直條圖 ────────────────────────────────────
    if use_single_bar:
        fig1 = px.bar(stats, x="問題類型", y="件數", text="百分比",
                      title="問題類型分布", color_discrete_sequence=[c_bar_single])
        fig1.update_traces(marker_color=c_bar_single)
    else:
        fig1 = px.bar(stats, x="問題類型", y="件數",
                      color="歸屬部門", text="百分比", title="問題類型分布",
                      color_discrete_map=DEPT_COLOR_MAP)
    fig1.update_traces(texttemplate="%{text}%", textposition="outside")
    fig1.update_layout(height=420, yaxis=dict(dtick=1, tickformat="d"),
                       margin=dict(t=45, b=0))
    c1.plotly_chart(fig1, use_container_width=True, key=f"{kp}_fig1")

    # ── 圖2：機台圓餅圖 ────────────────────────────────────────
    if m_stats is not None:
        cmap = {row["機型"]: custom_pie[i % len(custom_pie)]
                for i, row in m_stats.iterrows()}
        fig2 = px.pie(m_stats, names="機型", values="件數",
                      title="機台問題細分比較", hole=0.3,
                      color="機型", color_discrete_map=cmap)
        fig2.update_traces(texttemplate="%{percent:.1%}", textinfo="percent+label")
        fig2.update_layout(height=420, margin=dict(t=45, b=0, l=0, r=0))
        c2.plotly_chart(fig2, use_container_width=True, key=f"{kp}_fig2")
    else:
        c2.info("無機台相關數據")

    # ── 圖3：十大細項橫條圖 ────────────────────────────────────
    fig3 = px.bar(detail_stats, x="件數", y="問題細項",
                  orientation="h", title="十大問題細項分布",
                  color_discrete_sequence=[custom_hbar])
    fig3.update_traces(marker_color=custom_hbar)
    fig3.update_layout(height=420, yaxis={"categoryorder": "total ascending"},
                       xaxis=dict(dtick=1, tickformat="d"),
                       margin=dict(t=45, b=0, l=0, r=0))
    c3.plotly_chart(fig3, use_container_width=True, key=f"{kp}_fig3")

    # ── 把用戶自選顏色存進 session_state 供 PPT/ZIP 使用 ────────
    st.session_state[f"chart_colors_{kp}"] = {
        "bar":  c_bar_single if use_single_bar else None,
        "pie":  custom_pie,
        "hbar": custom_hbar,
    }


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
        color_map = {row["機型"]: BRAND_PALETTE[i % len(BRAND_PALETTE)]
                     for i, row in m_stats.iterrows()}
        fig2 = px.pie(
            m_stats, names="機型", values="件數",
            title="機台問題細分比較", hole=0.3,
            color="機型", color_discrete_map=color_map,
        )
        fig2.update_traces(texttemplate="%{percent:.1%}", textinfo="percent+label")
        fig2.update_layout(height=400, margin=dict(t=40, b=0, l=0, r=0))
        c2.plotly_chart(fig2, use_container_width=True, key=f"{key_prefix}_fig2" if key_prefix else None)
    else:
        c2.info("無機台相關數據")

    detail_stats = df["問題細項"].value_counts().reset_index().head(10)
    detail_stats.columns = ["問題細項", "件數"]
    fig3 = px.bar(
        detail_stats, x="件數", y="問題細項",
        orientation="h", title="十大問題細項分布",
        color_discrete_sequence=[BRAND_BLUE],
    )
    fig3.update_traces(marker_color=BRAND_BLUE)
    fig3.update_layout(
        height=400,
        yaxis={"categoryorder": "total ascending"},
        xaxis=dict(dtick=1, tickformat="d"),
        margin=dict(t=40, b=0, l=0, r=0),
    )
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
    start_d = end_d = None
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

    # 組合 source_name = 日期區間（用於 PPT 封面）
    if start_d and end_d:
        ppt_source = f"{start_d.strftime('%Y/%m/%d')}～{end_d.strftime('%Y/%m/%d')}"
    else:
        ppt_source = st.session_state.get("source_name", "unknown")

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

    # ── 預先產生所有下載檔案（避免 Streamlit on_click 時檔案還未產生）──
    chart_colors = st.session_state.get("chart_colors_sec2", {})

    # 用 session_state 快取，避免每次重繪都重新產生大檔
    cache_key = f"chart_pack_{ppt_source}"
    if cache_key not in st.session_state:
        with st.spinner("正在產生圖表與簡報..."):
            try:
                st.session_state[cache_key] = build_chart_pack(
                    df,
                    color_bar=chart_colors.get("bar"),
                    color_pie=chart_colors.get("pie"),
                    color_hbar=chart_colors.get("hbar"),
                )
            except Exception as e:
                st.error(f"圖表產生失敗：{e}")
                st.session_state[cache_key] = {}

    chart_pack = st.session_state[cache_key]

    ppt_cache_key = f"ppt_bytes_{ppt_source}"
    if ppt_cache_key not in st.session_state:
        with st.spinner("正在產生 PPT 簡報..."):
            try:
                st.session_state[ppt_cache_key] = build_ppt_bytes(
                    chart_stats, ai_text, ppt_source, chart_pack=chart_pack,
                )
            except Exception as e:
                st.error(f"PPT 產生失敗：{e}")
                st.session_state[ppt_cache_key] = b""

    ppt_bytes = st.session_state[ppt_cache_key]

    # ── 產生 ZIP ──
    zip_cache_key = f"zip_bytes_{ppt_source}"
    if zip_cache_key not in st.session_state:
        try:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for fn, b in chart_pack.items():
                    zi = zipfile.ZipInfo(fn)
                    zi.flag_bits |= 0x800
                    zi.compress_type = zipfile.ZIP_DEFLATED
                    zf.writestr(zi, b)
            st.session_state[zip_cache_key] = zip_buf.getvalue()
        except Exception as e:
            st.error(f"ZIP 產生失敗：{e}")
            st.session_state[zip_cache_key] = b""

    zip_bytes = st.session_state[zip_cache_key]

    # ── 下載按鈕（檔案已預先備好）──
    dl_col1, dl_col2, dl_col3 = st.columns(3)
    dl_col1.download_button(
        "⬇️ 下載 AI 分析文字檔",
        data=ai_text.encode("utf-8"),
        file_name=f"{datetime.now().strftime('%Y%m%d')}_AI重點分析.txt",
        mime="text/plain",
        use_container_width=True,
    )
    dl_col2.download_button(
        "⬇️ 下載圖表圖檔（ZIP）",
        data=zip_bytes,
        file_name=f"{datetime.now().strftime('%Y%m%d')}_圖表圖檔.zip",
        mime="application/zip",
        use_container_width=True,
        disabled=not zip_bytes,
    )
    dl_col3.download_button(
        "⬇️ 一鍵下載分析簡報 PPT",
        data=ppt_bytes,
        file_name=f"{datetime.now().strftime('%Y%m%d')}_分析簡報.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
        disabled=not ppt_bytes,
    )



def section_3():
    st.subheader("功能三：歷史分析紀錄")

    # ── Google Sheets 連線狀態 ──
    import os
    has_creds = bool(os.environ.get("GOOGLE_CREDENTIALS_JSON", ""))
    has_sid   = bool(os.environ.get("HISTORY_SHEET_ID", ""))
    ws_test   = _history_sheet()
    if ws_test is not None:
        st.success("☁️ Google Sheets 已連線，歷史紀錄永久保存")
    elif has_creds and has_sid:
        ws_test2 = _history_sheet(log_error=True)
        err_detail = st.session_state.get("_gsheet_error", "")
        st.warning(f"⚠️ 環境變數已設定但連線失敗\n{err_detail}")
        st.info("💡 請到 Google Cloud Console 確認已啟用 **Google Sheets API** 與 **Google Drive API**：\nhttps://console.cloud.google.com/apis/library")
    else:
        st.info("ℹ️ 未連線 Google Sheets，歷史紀錄僅限本次瀏覽")

    history = load_history()
    if not history:
        st.info("尚無歷史紀錄。")
        return

    # De-duplicate
    seen_names: dict = {}
    deduped = []
    for item in history:
        sn = item.get("source_name", "")
        if sn not in seen_names:
            seen_names[sn] = item
            deduped.append(item)
    history = deduped

    # ── 日期區間篩選器 ─────────────────────────────────────────
    st.markdown("---")
    st.markdown("##### 📅 日期區間篩選（有篩選時才顯示紀錄）")
    f_col1, f_col2, f_col3 = st.columns([2, 2, 1])

    # 取得所有紀錄的日期範圍
    all_dates = []
    for item in history:
        try:
            all_dates.append(datetime.fromisoformat(item["created_at"]).date())
        except Exception:
            pass

    if all_dates:
        min_date = min(all_dates)
        max_date = max(all_dates)
    else:
        min_date = max_date = datetime.now().date()

    start_filter = f_col1.date_input("開始日期", value=None, min_value=min_date, max_value=max_date,
                                      key="s3_start", format="YYYY/MM/DD")
    end_filter   = f_col2.date_input("結束日期", value=None, min_value=min_date, max_value=max_date,
                                      key="s3_end", format="YYYY/MM/DD")
    do_filter = f_col3.button("🔍 篩選", key="s3_filter", use_container_width=True)

    # 是否已啟動篩選
    filter_active = start_filter is not None or end_filter is not None

    if not filter_active:
        st.caption("請選擇日期區間後按「篩選」按鈕，即可顯示該區間的歷史紀錄。")
        return

    # 依日期區間過濾
    filtered = []
    for item in history:
        try:
            item_date = datetime.fromisoformat(item["created_at"]).date()
            if start_filter and item_date < start_filter:
                continue
            if end_filter and item_date > end_filter:
                continue
            filtered.append(item)
        except Exception:
            filtered.append(item)

    if not filtered:
        st.info(f"所選日期區間（{start_filter} ～ {end_filter}）無歷史紀錄。")
        return

    st.caption(f"共找到 {len(filtered)} 筆紀錄")
    history = filtered

    for item in history:
        out_path = Path(item.get("output_path", ""))
        cache = st.session_state.get("_history_cache", {})
        item_id = item["id"]

        # 取得 excel bytes：磁碟 → session_state 快取（已由 load_history 從 Sheets 填入）
        dl_bytes = None
        df_hist  = None
        if out_path.exists():
            try:
                dl_bytes = out_path.read_bytes()
                df_hist  = pd.read_excel(io.BytesIO(dl_bytes))
            except Exception:
                dl_bytes = None
        if dl_bytes is None and item_id in cache:
            try:
                dl_bytes = cache[item_id]["excel_bytes"]
                df_hist  = pd.read_excel(io.BytesIO(dl_bytes))
            except Exception:
                dl_bytes = None

        if dl_bytes is None:
            continue   # 真的找不到，跳過
        
        sname = item.get('source_name', '')
        if len(sname) > 28:
            sname = sname[:14] + "..." + sname[-10:]
        label = f"{item['created_at'][:16]}  {sname}  ({item['rows']} 筆)"
        with st.expander(label):
            tab_data, tab_chart, tab_ai = st.tabs(["資料預覽", "圖表分析", "AI 重點摘要"])
            
            with tab_data:
                st.dataframe(df_hist.head(30), use_container_width=True, hide_index=True)
                col1, col2, col3 = st.columns([1, 1, 1])
                col1.download_button(
                    "下載該分析檔",
                    data=dl_bytes,
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
                            zi = zipfile.ZipInfo(fn)
                            zi.flag_bits |= 0x800  # UTF-8 filename flag，避免中文亂碼
                            zi.compress_type = zipfile.ZIP_DEFLATED
                            zf.writestr(zi, b)
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



def section_4():
    """功能四：週/月/季/年度趨勢分析儀表板"""

    # ── ECOCO 品牌 CSS（對齊 HTML 範本風格）──────────────────────
    st.markdown("""<style>
    .s4-header{background:#060E9F;color:#fff;padding:22px 26px;border-radius:12px;
               border-bottom:6px solid #FF5000;margin-bottom:18px}
    .s4-header h2{margin:0;font-size:20px;font-weight:700;letter-spacing:.3px}
    .s4-header p{margin:4px 0 0;opacity:.85;font-size:13px}
    .s4-section{border-left:6px solid #FF5000;padding-left:14px;
                color:#060E9F;font-size:17px;font-weight:700;margin:22px 0 14px}
    .s4-card{background:#fff;border-radius:12px;padding:20px 24px;
             box-shadow:0 4px 10px rgba(0,0,0,.06);margin-bottom:16px}
    .s4-kpi-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:18px}
    .s4-kpi{background:#fff;border-radius:10px;padding:18px 16px;text-align:center;
            border-top:4px solid #FFCE00;box-shadow:0 2px 6px rgba(0,0,0,.05)}
    .s4-kpi-val{font-size:30px;font-weight:700;color:#FF5000}
    .s4-kpi-lbl{font-size:12px;color:#666;margin-top:4px}
    .s4-kpi-delta{font-size:11px;margin-top:3px}
    .delta-up{color:#c03000} .delta-dn{color:#0a6e44} .delta-flat{color:#888}
    .s4-rank-table{width:100%;border-collapse:collapse}
    .s4-rank-table th{background:#060E9F;color:#fff;padding:10px 12px;font-size:13px;text-align:center}
    .s4-rank-table td{padding:10px 12px;text-align:center;border-bottom:1px solid #eee;font-size:13px}
    .s4-rank-val{color:#FF5000;font-weight:700}
    .filter-chip-row{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px}
    .filter-chip{padding:4px 14px;border-radius:20px;font-size:12px;font-weight:600;cursor:pointer;
                 border:1.5px solid #060E9F;background:#fff;color:#060E9F}
    .filter-chip.active{background:#060E9F;color:#fff}
    </style>""", unsafe_allow_html=True)

    # ── 頁首 ──────────────────────────────────────────────────────
    st.markdown("""
    <div class="s4-header">
      <h2>📈 ECOCO 客訴趨勢分析儀表板</h2>
      <p>城市・站點・部門・問題類型・機台比例 | 自訂日期區間 + 維度篩選</p>
    </div>""", unsafe_allow_html=True)

    # ── 資料來源 ──────────────────────────────────────────────────
    src_tab1, src_tab2 = st.tabs(["📂 歷史紀錄資料", "🔗 填入 Google Sheets 網址"])
    all_dfs: list[pd.DataFrame] = []

    with src_tab1:
        ws = _history_sheet()
        if ws:
            import base64 as _b64
            try:
                for grow in ws.get_all_values()[1:]:
                    if not grow or not grow[0]: continue
                    excel_b64 = grow[4] if len(grow) > 4 else ""
                    if excel_b64:
                        try: all_dfs.append(pd.read_excel(io.BytesIO(_b64.b64decode(excel_b64))))
                        except Exception: pass
            except Exception: pass
        for v in st.session_state.get("_history_cache", {}).values():
            try: all_dfs.append(pd.read_excel(io.BytesIO(v["excel_bytes"])))
            except Exception: pass
        st.caption(f"已載入 {len(all_dfs)} 份歷史紀錄" if all_dfs else "尚無歷史資料")

    with src_tab2:
        gs_url = st.text_input("Google Sheets 網址", placeholder="https://docs.google.com/spreadsheets/d/xxxxx/edit", key="s4v3_gsurl")
        gs_sheet = st.text_input("工作表名稱（留空讀取第一張）", key="s4v3_gssheet", value="")
        if st.button("📥 讀取", key="s4v3_load_gs"):
            if not gs_url:
                st.error("請填入網址")
            else:
                try:
                    import re as _re
                    m = _re.search(r"/spreadsheets/d/([^/]+)", gs_url)
                    if not m:
                        st.error("無法解析試算表 ID")
                    else:
                        _client = _get_gsheet_client()
                        if not _client:
                            st.error("未連線 Google API")
                        else:
                            _ss = _client.open_by_key(m.group(1))
                            _ws = _ss.worksheet(gs_sheet) if gs_sheet else _ss.get_worksheet(0)
                            _rows = _ws.get_all_values()
                            if _rows:
                                _df = pd.DataFrame(_rows[1:], columns=_rows[0])
                                all_dfs.append(_df)
                                st.session_state["_s4v3_gs_df"] = _df
                                st.success(f"✅ 已讀取「{_ws.title}」，共 {len(_df)} 列")
                except Exception as e:
                    st.error(f"讀取失敗：{e}")
        if st.session_state.get("_s4v3_gs_df") is not None:
            all_dfs.append(st.session_state["_s4v3_gs_df"])

    if not all_dfs:
        st.info("尚無資料，請先在功能一完成分析儲存，或填入 Google Sheets 網址。")
        return

    df_all = pd.concat(all_dfs, ignore_index=True).drop_duplicates()

    # ── 欄位自動偵測 ──────────────────────────────────────────────
    date_col   = next((c for c in df_all.columns if "日期" in c or "date" in c.lower()), None)
    type_col   = next((c for c in df_all.columns if "問題類型" in c), None)
    detail_col = next((c for c in df_all.columns if "問題細項" in c), None)
    dept_col   = next((c for c in df_all.columns if "部門" in c or "歸屬" in c), None)
    city_col   = next((c for c in df_all.columns if "站點區域" in c or "城市" in c or "區域" in c), None)
    station_col= next((c for c in df_all.columns if c == "站點名稱"), None) or \
                 next((c for c in df_all.columns if "站點名稱" in c and "編號" not in c), None)
    machine_col= next((c for c in df_all.columns if "機台類型" in c or "機台" in c), None)

    if not date_col:
        st.warning("找不到日期欄位")
        return

    df_all[date_col] = pd.to_datetime(df_all[date_col], errors="coerce")
    df_all = df_all.dropna(subset=[date_col])
    if df_all.empty:
        st.warning("無有效日期資料")
        return

    # ── 時間區間選擇（維度 + 自訂日期）────────────────────────────
    st.markdown('<div class="s4-section">⚙️ 篩選條件</div>', unsafe_allow_html=True)
    filter_c1, filter_c2, filter_c3 = st.columns([2, 3, 2])

    dim_mode = filter_c1.radio("時間模式", ["維度選擇", "自訂日期區間"], horizontal=True, key="s4v3_dimmode")

    DIM_FREQ = {"週": "W", "月": "M", "季": "Q", "年度": "Y"}
    period_sel = period_prev = None
    df_cur = df_prev = pd.DataFrame()

    if dim_mode == "維度選擇":
        dim = filter_c2.selectbox("分析維度", ["週", "月", "季", "年度"], index=1, key="s4v3_dim")
        df_all["_period"] = df_all[date_col].dt.to_period(DIM_FREQ[dim]).astype(str)
        periods = sorted(df_all["_period"].unique(), reverse=True)
        if not periods:
            st.warning("資料不足")
            return
        period_sel = filter_c3.selectbox(f"本期", periods, key="s4v3_period")
        p_idx = periods.index(period_sel)
        period_prev = periods[p_idx + 1] if p_idx + 1 < len(periods) else None
        df_cur  = df_all[df_all["_period"] == period_sel].copy()
        df_prev = df_all[df_all["_period"] == period_prev].copy() if period_prev else pd.DataFrame()
        period_label = period_sel
    else:
        min_d = df_all[date_col].min().date()
        max_d = df_all[date_col].max().date()
        st.markdown("**本期：**")
        d_col1, d_col2 = filter_c2.columns(2)
        start_d = d_col1.date_input("開始", value=min_d, min_value=min_d, max_value=max_d, key="s4v3_sd")
        end_d   = d_col2.date_input("結束", value=max_d, min_value=min_d, max_value=max_d, key="s4v3_ed")
        # 對照期
        st.markdown("**對照期（可選）：**")
        cp_col1, cp_col2, cp_col3 = filter_c3.columns(3)
        use_cmp = cp_col1.checkbox("啟用對照期", key="s4v3_use_cmp")
        cmp_start = cp_col2.date_input("對照開始", value=min_d, min_value=min_d, max_value=max_d, key="s4v3_csd",
                                        disabled=not use_cmp)
        cmp_end   = cp_col3.date_input("對照結束", value=start_d, min_value=min_d, max_value=max_d, key="s4v3_ced",
                                        disabled=not use_cmp)
        df_cur  = df_all[(df_all[date_col].dt.date >= start_d) & (df_all[date_col].dt.date <= end_d)].copy()
        period_label = f"{start_d} ～ {end_d}"
        if use_cmp:
            df_prev = df_all[(df_all[date_col].dt.date >= cmp_start) & (df_all[date_col].dt.date <= cmp_end)].copy()
            period_prev = f"{cmp_start} ～ {cmp_end}"
        else:
            period_prev = None
            df_prev = pd.DataFrame()

    # ── 多維篩選 chips（城市/部門/問題類型/機台）──────────────────
    st.markdown("**篩選維度：**")
    chip_cols = st.columns(4)

    city_filter   = chip_cols[0].multiselect("🏙️ 城市", sorted(df_cur[city_col].dropna().unique().tolist()) if city_col and city_col in df_cur.columns else [], key="s4v3_city")
    dept_filter   = chip_cols[1].multiselect("🏢 部門", sorted(df_cur[dept_col].dropna().unique().tolist()) if dept_col and dept_col in df_cur.columns else [], key="s4v3_dept")
    type_filter   = chip_cols[2].multiselect("❓ 問題類型", sorted(df_cur[type_col].dropna().unique().tolist()) if type_col and type_col in df_cur.columns else [], key="s4v3_type")
    mach_filter   = chip_cols[3].multiselect("🔧 機台類型", sorted(df_cur[machine_col].dropna().unique().tolist()) if machine_col and machine_col in df_cur.columns else [], key="s4v3_mach")

    df_filt = df_cur.copy()
    if city_filter  and city_col:   df_filt = df_filt[df_filt[city_col].isin(city_filter)]
    if dept_filter  and dept_col:   df_filt = df_filt[df_filt[dept_col].isin(dept_filter)]
    if type_filter  and type_col:   df_filt = df_filt[df_filt[type_col].isin(type_filter)]
    if mach_filter  and machine_col: df_filt = df_filt[df_filt[machine_col].isin(mach_filter)]

    n_cur  = len(df_filt)
    n_prev = len(df_prev)

    def pct_change(cur, prev):
        if prev == 0: return None
        return (cur - prev) / prev * 100

    # ── 機台類型：將「方舟」歸類為「收瓶機」────────────────────────
    if machine_col and machine_col in df_filt.columns:
        def _normalize_machine(val):
            v = str(val).strip()
            if "方舟" in v or "收瓶" in v: return "收瓶機"
            if "電池" in v: return "電池機"
            return v
        df_filt = df_filt.copy()
        df_filt[machine_col] = df_filt[machine_col].apply(_normalize_machine)
    if machine_col and machine_col in df_all.columns:
        df_all = df_all.copy()
        df_all[machine_col] = df_all[machine_col].apply(
            lambda v: "收瓶機" if ("方舟" in str(v) or "收瓶" in str(v)) else ("電池機" if "電池" in str(v) else str(v))
        )

    # ── KPI 卡片（用 st.metric 避免 HTML escape 問題）────────────────
    st.markdown(f'<div class="s4-section">📊 本期即時統計（{period_label}）</div>', unsafe_allow_html=True)
    st.caption(f"📅 資料區間：{period_label}　篩選後共 **{n_cur}** 筆")

    kpi_items = [("🗂️ 總進件數", n_cur, n_prev)]
    if type_col and type_col in df_filt.columns:
        for t, tc in df_filt[type_col].value_counts().items():
            prev_tc = int(df_prev[type_col].eq(t).sum()) if not df_prev.empty and type_col in df_prev.columns else 0
            kpi_items.append((str(t)[:12], int(tc), prev_tc))
            if len(kpi_items) >= 4: break

    kpi_cols = st.columns(len(kpi_items[:4]))
    for col_i, (lbl, cur, prev) in enumerate(kpi_items[:4]):
        p = pct_change(cur, prev)
        delta_str = None
        if p is not None:
            sym = "+" if p >= 0 else ""
            delta_str = f"{sym}{p:.1f}% vs 上期"
        kpi_cols[col_i].metric(label=lbl, value=cur, delta=delta_str)

    # ── 排行統計（區域/站點/問題細項）───────────────────────────────
    st.markdown(f'<div class="s4-section">🏆 案件排行統計 Top 5 ── {period_label}</div>', unsafe_allow_html=True)

    rank_cols = st.columns(3)
    MEDAL = ["🥇","🥈","🥉","4️⃣","5️⃣"]

    def rank_table_html(series, header1, header2):
        rows = ""
        for idx, (k, v) in enumerate(series.head(5).items()):
            m = MEDAL[idx] if idx < len(MEDAL) else str(idx+1)
            rows += f'<tr><td>{m} {str(k)[:20]}</td><td class="s4-rank-val">{int(v)}</td></tr>'
        return f'''<table class="s4-rank-table">
          <thead><tr><th>{header1}</th><th>{header2}</th></tr></thead>
          <tbody>{rows}</tbody>
        </table>'''

    with rank_cols[0]:
        st.markdown('<div style="font-weight:700;color:#060E9F;margin-bottom:8px">📍 區域排行</div>', unsafe_allow_html=True)
        if city_col and city_col in df_filt.columns and not df_filt[city_col].dropna().empty:
            st.markdown('<div class="s4-card">' + rank_table_html(df_filt[city_col].value_counts(), "城市/區域", "件數") + '</div>', unsafe_allow_html=True)
        else:
            st.info("無城市資料")

    with rank_cols[1]:
        st.markdown('<div style="font-weight:700;color:#060E9F;margin-bottom:8px">🏬 站點排行</div>', unsafe_allow_html=True)
        if station_col and station_col in df_filt.columns and not df_filt[station_col].dropna().empty:
            st.markdown('<div class="s4-card">' + rank_table_html(df_filt[station_col].value_counts(), "站點名稱", "件數") + '</div>', unsafe_allow_html=True)
        else:
            st.info("無站點資料")

    with rank_cols[2]:
        st.markdown('<div style="font-weight:700;color:#060E9F;margin-bottom:8px">🔍 問題細項排行</div>', unsafe_allow_html=True)
        if detail_col and detail_col in df_filt.columns and not df_filt[detail_col].dropna().empty:
            st.markdown('<div class="s4-card">' + rank_table_html(df_filt[detail_col].value_counts(), "問題細項", "件數") + '</div>', unsafe_allow_html=True)
        else:
            st.info("無細項資料")

    # ── 圖表：問題類型 + 機台佔比（對齊 HTML 範本）─────────────────
    st.markdown(f'<div class="s4-section">📉 數據可視化分析 ── {period_label}</div>', unsafe_allow_html=True)
    chart_col1, chart_col2 = st.columns(2)

    with chart_col1:
        if type_col and type_col in df_filt.columns:
            _tc = df_filt[type_col].value_counts()
            _total_tc = _tc.sum()
            _CLRS_PIE = ["#060E9F","#FF5000","#FFCE00","#8EB9C9","#0076A9","#FAE0B8"]
            fig_pie = px.pie(
                values=_tc.values, names=_tc.index,
                title=f"{period_label} 客訴類別分佈",
                hole=0.32,
                color_discrete_sequence=_CLRS_PIE,
            )
            fig_pie.update_traces(
                texttemplate="%{percent:.0%}",
                textposition="inside",
                textfont_size=13,
                textfont_color="white",
                hovertemplate="<b>%{label}</b><br>%{value} 件　%{percent:.1%}<extra></extra>",
                showlegend=False,
            )
            # 右側緊湊圖例：■ 類別  XX%（N件）
            _n_items = len(_tc)
            _step = min(0.12, 0.85 / max(_n_items, 1))
            _anns = []
            for _i, (_k, _v) in enumerate(_tc.items()):
                _pct = int(_v) / max(_total_tc, 1) * 100
                _clr = _CLRS_PIE[_i % len(_CLRS_PIE)]
                _anns.append(dict(
                    x=1.01, y=0.96 - _i * _step,
                    xref="paper", yref="paper",
                    text=f'<span style="color:{_clr};font-size:14px">■</span>'
                         f' <b>{_k}</b>　{_pct:.0f}%（{int(_v)}件）',
                    showarrow=False, align="left", font=dict(size=11),
                    xanchor="left",
                ))
            fig_pie.update_layout(
                height=420, showlegend=False,
                annotations=_anns,
                margin=dict(t=50, b=10, l=10, r=160),
                title_font_size=14,
            )
            st.plotly_chart(fig_pie, use_container_width=True)

    with chart_col2:
        if machine_col and machine_col in df_filt.columns and not df_filt[machine_col].dropna().empty:
            _mc = df_filt[machine_col].value_counts()
            _total_mc = _mc.sum()
            _CLRS_MAC = ["#FF5000","#060E9F","#8EB9C9","#FFCE00"]
            fig_mac = px.pie(
                values=_mc.values, names=_mc.index,
                title=f"{period_label} 機台客訴佔比",
                color_discrete_sequence=_CLRS_MAC,
            )
            _n_mac = len(_mc)
            _step_mac = min(0.15, 0.85 / max(_n_mac, 1))
            _anns_mac = []
            for _i, (_k, _v) in enumerate(_mc.items()):
                _pct = int(_v) / max(_total_mc, 1) * 100
                _clr = _CLRS_MAC[_i % len(_CLRS_MAC)]
                _anns_mac.append(dict(
                    x=1.01, y=0.92 - _i * _step_mac,
                    xref="paper", yref="paper",
                    text=f'<span style="color:{_clr};font-size:14px">■</span>'
                         f' <b>{_k}</b>　{_pct:.0f}%（{int(_v)}件）',
                    showarrow=False, align="left", font=dict(size=12),
                    xanchor="left",
                ))
            fig_mac.update_traces(
                texttemplate="%{percent:.0%}",
                textposition="inside",
                textfont_size=14,
                textfont_color="white",
                hovertemplate="<b>%{label}</b><br>%{value} 件　%{percent:.1%}<extra></extra>",
                showlegend=False,
            )
            fig_mac.update_layout(
                height=420, showlegend=False,
                annotations=_anns_mac,
                margin=dict(t=50, b=10, l=10, r=160),
            )
            st.plotly_chart(fig_mac, use_container_width=True)
        elif detail_col and detail_col in df_filt.columns:
            _dc = df_filt[detail_col].value_counts().head(8)
            fig_det = px.bar(
                x=list(_dc.values)[::-1], y=list(_dc.index)[::-1],
                orientation="h", title=f"{period_label} TOP 8 問題細項",
                color_discrete_sequence=["#060E9F"],
            )
            fig_det.update_layout(height=430, xaxis=dict(dtick=1,tickformat="d"),
                                   margin=dict(t=45,b=0,l=0,r=0))
            st.plotly_chart(fig_det, use_container_width=True)

    # ── 趨勢折線圖 ────────────────────────────────────────────────
    st.markdown(f'<div class="s4-section">📈 客訴趨勢分析 ── {period_label}</div>', unsafe_allow_html=True)
    if dim_mode == "維度選擇" and len(df_all["_period"].unique()) >= 2:
        _trend = df_all.groupby("_period").size().reset_index(name="件數").sort_values("_period")
        fig_line = px.line(
            _trend, x="_period", y="件數",
            title=f"歷史件數趨勢",
            markers=True,
            color_discrete_sequence=["#FF5000"],
        )
        fig_line.update_traces(fill="tozeroy", fillcolor="rgba(255,80,0,0.1)")
        if period_sel and period_sel in _trend["_period"].values:
            _sel_i = _trend.index[_trend["_period"] == period_sel].tolist()
            if _sel_i:
                fig_line.add_vline(x=_sel_i[0], line_dash="dash", line_color="#060E9F",
                                   annotation_text="本期", annotation_font_color="#060E9F")
        fig_line.update_layout(
            height=320, xaxis_title="期間",
            yaxis=dict(dtick=1, tickformat="d"),
            paper_bgcolor="white", plot_bgcolor="rgba(250,224,184,0.15)",
            margin=dict(t=45,b=0),
        )
        st.plotly_chart(fig_line, use_container_width=True)
    else:
        # 自訂日期：每日件數，最高值標橘色
        _daily = df_filt.groupby(df_filt[date_col].dt.date).size().reset_index(name="件數")
        _daily.columns = ["日期", "件數"]
        if not _daily.empty:
            _max_cnt = _daily["件數"].max()
            _bar_colors = ["#FF5000" if v == _max_cnt else "#060E9F" for v in _daily["件數"]]
            fig_daily = px.bar(
                _daily, x="日期", y="件數",
                title="期間內每日件數（橘色＝最高日）",
                text="件數",
            )
            fig_daily.update_traces(
                marker_color=_bar_colors,
                texttemplate="%{text}",
                textposition="outside",
                textfont_size=10,
            )
            fig_daily.update_layout(
                height=320,
                yaxis=dict(dtick=1, tickformat="d"),
                margin=dict(t=50,b=0),
                showlegend=False,
            )
            # 對照期疊加折線
            if not df_prev.empty and date_col in df_prev.columns:
                _daily_prev = df_prev.groupby(df_prev[date_col].dt.date).size().reset_index(name="件數")
                _daily_prev.columns = ["日期", "件數"]
                if not _daily_prev.empty:
                    import plotly.graph_objects as go
                    fig_daily.add_trace(go.Scatter(
                        x=[str(d) for d in _daily_prev["日期"]],
                        y=list(_daily_prev["件數"]),
                        mode="lines+markers",
                        name=f"對照期（{period_prev}）",
                        line=dict(color="#FFCE00", width=2, dash="dot"),
                        marker=dict(size=6),
                    ))
            st.plotly_chart(fig_daily, use_container_width=True)

    # ── 城市展開排行（可折疊）────────────────────────────────────
    if city_col and city_col in df_filt.columns and not df_filt.empty:
        st.markdown(f'<div class="s4-section">🏙️ 區域排行榜 ── {period_label}</div>', unsafe_allow_html=True)
        city_rank = df_filt[city_col].value_counts()
        MEDAL_LIST = ["🥇","🥈","🥉","4️⃣","5️⃣","6️⃣","7️⃣","8️⃣","9️⃣","🔟"]
        for ri, (city, cnt) in enumerate(city_rank.items()):
            prev_cnt = int(df_prev[city_col].eq(city).sum()) if not df_prev.empty and city_col in df_prev.columns else 0
            p = pct_change(int(cnt), prev_cnt)
            delta_s = (f"　{'▲' if p>0 else '▼'}{abs(p):.1f}%") if p is not None else ""
            medal = MEDAL_LIST[ri] if ri < len(MEDAL_LIST) else f"#{ri+1}"
            with st.expander(f"{medal} **{city}**　{int(cnt)} 件{delta_s}", expanded=(ri==0)):
                df_city = df_filt[df_filt[city_col] == city]
                ec1, ec2 = st.columns(2)
                with ec1:
                    st.markdown("**📍 站點排行**")
                    if station_col and station_col in df_city.columns:
                        _sr = df_city[station_col].value_counts().head(8)
                        _sr_prev = df_prev[df_prev[city_col]==city][station_col].value_counts() if not df_prev.empty and city_col in df_prev.columns and station_col in df_prev.columns else pd.Series(dtype=int)
                        _html = ""
                        for si, (sn, sv) in enumerate(_sr.items()):
                            _sp = int(_sr_prev.get(sn, 0))
                            _sd = (f"　{'▲' if int(sv)-_sp>0 else '▼'}{abs(pct_change(int(sv),_sp)):.0f}%") if _sp and pct_change(int(sv),_sp) is not None else ""
                            _sm = MEDAL_LIST[si] if si < len(MEDAL_LIST) else f"#{si+1}"
                            _html += f'<div style="padding:5px 0;border-bottom:.5px solid #eee;font-size:13px;display:flex;justify-content:space-between"><span>{_sm} {str(sn)[:18]}</span><b style="color:#FF5000">{int(sv)}{_sd}</b></div>'
                        st.markdown(_html, unsafe_allow_html=True)
                with ec2:
                    st.markdown("**🔍 問題細項排行**")
                    if detail_col and detail_col in df_city.columns:
                        _dr = df_city[detail_col].value_counts().head(8)
                        _max = int(_dr.max()) if not _dr.empty else 1
                        _dhtml = ""
                        for di, (dn, dv) in enumerate(_dr.items()):
                            _dm = MEDAL_LIST[di] if di < len(MEDAL_LIST) else f"#{di+1}"
                            _bp = int(dv)/_max*100
                            _dhtml += f'''<div style="padding:5px 0;border-bottom:.5px solid #eee">
                              <div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:3px">
                                <span>{_dm} {str(dn)[:14]}</span><b style="color:#060E9F">{int(dv)}</b></div>
                              <div style="background:#f0f0f0;border-radius:3px;height:5px">
                                <div style="background:#060E9F;width:{_bp:.0f}%;height:100%;border-radius:3px"></div></div>
                            </div>'''
                        st.markdown(_dhtml, unsafe_allow_html=True)

    # ── 部門分析 ─────────────────────────────────────────────────
    if dept_col and dept_col in df_filt.columns:
        st.markdown(f'<div class="s4-section">🏢 各部門件數分析 ── {period_label}</div>', unsafe_allow_html=True)
        dept_rank = df_filt[dept_col].replace("","未分配").value_counts()
        DEPT_COLOR = {"營運部":"#FF5000","行銷部":"#FFCE00","資訊部":"#060E9F"}
        fig_dept = px.bar(
            dept_rank.reset_index(), x=dept_col, y="count",
            title="各部門件數",
            color=dept_col,
            color_discrete_map=DEPT_COLOR,
        )
        fig_dept.update_layout(height=300, yaxis=dict(dtick=1,tickformat="d"),
                                showlegend=False, margin=dict(t=45,b=0))
        st.plotly_chart(fig_dept, use_container_width=True)

    # ── 完整頁面 PDF 下載 ────────────────────────────────────────
    st.markdown("---")
    st.markdown(f'<div class="s4-section">⬇️ 下載完整分析報告</div>', unsafe_allow_html=True)

    if st.button("📄 產生完整分析 PDF", key="s4_full_pdf", use_container_width=False):
        with st.spinner("正在產生多頁 PDF 報告..."):
            try:
                from fpdf import FPDF
                from fpdf.enums import XPos, YPos
                import os, glob, matplotlib.pyplot as _mplt
                from matplotlib.ticker import MaxNLocator as _MNL

                _FONT_CANDS = [
                    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
                    "/usr/share/fonts/opentype/noto/NotoSansCJK-Medium.ttc",
                    "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
                    "/usr/share/fonts/truetype/arphic/uming.ttc",
                    "/tmp/NotoSansCJK.ttc",
                ]
                _FONT_CANDS += glob.glob("/usr/share/fonts/**/NotoSansCJK*.ttc", recursive=True)
                _font_path = _ensure_cjk_font()  # 使用已快取的字型路徑
                if not _font_path:
                    _font_path = next((p for p in _FONT_CANDS if os.path.exists(p)), None)

                _setup_cjk_font()

                class EcocoPDF(FPDF):
                    def __init__(self, font_path, font_name):
                        super().__init__(orientation="P", format="A4")
                        self.fp = font_path
                        self.fn = font_name
                        self.set_auto_page_break(auto=True, margin=15)
                        if font_path:
                            self.add_font(font_name, style="", fname=font_path)
                        self.set_margins(15, 15, 15)

                    def header(self):
                        # 藍色頁首條
                        self.set_fill_color(6, 14, 159)
                        self.rect(0, 0, 210, 14, style="F")
                        self.set_font(self.fn, size=9)
                        self.set_text_color(255, 255, 255)
                        self.set_xy(5, 3)
                        self.cell(0, 8, self._s(f"ECOCO 客訴趨勢分析報告　{period_label}"))
                        self.set_draw_color(255, 80, 0)
                        self.set_line_width(1.5)
                        self.line(0, 14, 210, 14)
                        self.set_line_width(0.2)
                        self.set_text_color(30, 30, 30)
                        self.ln(6)

                    def footer(self):
                        self.set_y(-12)
                        self.set_font(self.fn, size=8)
                        self.set_text_color(150, 150, 150)
                        self.cell(0, 8, self._s(f"第 {self.page_no()} 頁　產出日期：{datetime.now().strftime('%Y/%m/%d')}"), align="C")

                    def _s(self, s):
                        return s if self.fn != "Helvetica" else s.encode("ascii", "replace").decode()

                    def section_title(self, title):
                        self.ln(3)
                        self.set_fill_color(255, 80, 0)
                        self.rect(self.get_x(), self.get_y(), 6, 9, style="F")
                        self.set_font(self.fn, size=13)
                        self.set_text_color(6, 14, 159)
                        self.set_x(self.get_x() + 9)
                        self.cell(0, 9, self._s(title), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                        self.ln(2)

                    def full_table(self, headers, rows, col_widths):
                        """全幅表格，支援多行自動換行"""
                        # 表頭
                        self.set_fill_color(6, 14, 159)
                        self.set_text_color(255, 255, 255)
                        self.set_font(self.fn, size=9)
                        x0 = self.get_x()
                        for h, w in zip(headers, col_widths):
                            self.cell(w, 8, self._s(str(h)), border=1, fill=True, align="C",
                                      new_x=XPos.RIGHT, new_y=YPos.TOP)
                        self.ln(8)
                        # 資料列
                        self.set_text_color(30, 30, 30)
                        self.set_font(self.fn, size=9)
                        for i, row in enumerate(rows):
                            bg = (235, 244, 250) if i % 2 == 0 else (255, 255, 255)
                            self.set_fill_color(*bg)
                            self.set_x(x0)
                            for val, w in zip(row, col_widths):
                                self.cell(w, 7, self._s(str(val)[:40]), border=1, fill=True,
                                          new_x=XPos.RIGHT, new_y=YPos.TOP)
                            self.ln(7)
                        self.ln(3)

                    def embed_image(self, fig, w=180, h=110):
                        """嵌入 matplotlib 圖表"""
                        _b = io.BytesIO()
                        fig.savefig(_b, format="png", dpi=180, bbox_inches="tight",
                                    facecolor="white")
                        _mplt.close(fig)
                        _b.seek(0)
                        x = (210 - w) / 2  # 置中
                        self.image(_b, x=x, y=self.get_y(), w=w, h=h)
                        self.set_y(self.get_y() + h + 4)

                F = "CJK" if _font_path else "Helvetica"

                pdf = EcocoPDF(_font_path, F)

                # ════════════════════════════════════════════
                # Page 1：封面 + KPI 摘要
                # ════════════════════════════════════════════
                pdf.add_page()
                # 大標題框
                pdf.set_fill_color(6, 14, 159)
                pdf.rect(15, 20, 180, 38, style="F")
                pdf.set_font(F, size=20)
                pdf.set_text_color(255, 255, 255)
                pdf.set_xy(15, 26)
                pdf.cell(180, 12, pdf._s("ECOCO 客訴趨勢分析報告"), align="C",
                         new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                pdf.set_font(F, size=11)
                pdf.set_xy(15, 42)
                pdf.cell(180, 8, pdf._s(f"資料區間：{period_label}"), align="C",
                         new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                pdf.set_draw_color(255, 80, 0)
                pdf.set_line_width(2)
                pdf.line(15, 58, 195, 58)
                pdf.set_line_width(0.2)
                pdf.set_y(68)

                # KPI 卡片（橫排）
                pdf.section_title("本期即時統計摘要")
                kpi_data = [("🗂️ 總進件數", n_cur)]
                if type_col and type_col in df_filt.columns:
                    for t, tc in df_filt[type_col].value_counts().head(3).items():
                        kpi_data.append((str(t), int(tc)))
                card_w = 170 // len(kpi_data)
                card_x = 20
                for lbl, val in kpi_data:
                    pdf.set_fill_color(255, 206, 0)  # 黃色頂線
                    pdf.rect(card_x, pdf.get_y(), card_w - 4, 3, style="F")
                    pdf.set_fill_color(248, 249, 252)
                    pdf.rect(card_x, pdf.get_y() + 3, card_w - 4, 28, style="F")
                    pdf.set_font(F, size=22)
                    pdf.set_text_color(255, 80, 0)
                    pdf.set_xy(card_x, pdf.get_y() + 5)
                    pdf.cell(card_w - 4, 14, str(val), align="C",
                             new_x=XPos.RIGHT, new_y=YPos.TOP)
                    pdf.set_font(F, size=8)
                    pdf.set_text_color(80, 80, 80)
                    pdf.set_xy(card_x, pdf.get_y() + 20)
                    pdf.cell(card_w - 4, 8, pdf._s(str(lbl)[:10]), align="C",
                             new_x=XPos.RIGHT, new_y=YPos.TOP)
                    card_x += card_w
                pdf.set_y(pdf.get_y() + 36)

                # 城市 KPI
                if city_col and city_col in df_filt.columns:
                    pdf.ln(4)
                    pdf.set_font(F, size=9)
                    pdf.set_text_color(60, 60, 60)
                    city_top = df_filt[city_col].value_counts().head(3)
                    line = "　|　".join([f"{c}：{int(v)} 件" for c, v in city_top.items()])
                    pdf.set_x(20)
                    pdf.cell(0, 7, pdf._s(f"前三大城市：{line}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)

                # ════════════════════════════════════════════
                # Page 2：排行統計
                # ════════════════════════════════════════════
                pdf.add_page()
                pdf.section_title("案件排行統計")

                # 區域排行
                if city_col and city_col in df_filt.columns:
                    pdf.set_font(F, size=10); pdf.set_text_color(6,14,159)
                    pdf.cell(0, 7, pdf._s("📍 城市/區域排行"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    _city_v = df_filt[city_col].value_counts()
                    rows_c = [[i+1, c, int(v), f"{int(v)/n_cur*100:.0f}%"]
                               for i, (c, v) in enumerate(_city_v.head(10).items())]
                    pdf.full_table(["排名","城市/區域","件數","佔比"], rows_c, [15,110,25,30])

                # 站點排行
                if station_col and station_col in df_filt.columns:
                    pdf.set_font(F, size=10); pdf.set_text_color(6,14,159)
                    pdf.cell(0, 7, pdf._s("🏬 站點排行"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    _sta_v = df_filt[station_col].value_counts()
                    rows_s = [[i+1, str(s)[:30], int(v)]
                               for i, (s, v) in enumerate(_sta_v.head(10).items())]
                    pdf.full_table(["排名","站點名稱","件數"], rows_s, [15,140,25])

                # ════════════════════════════════════════════
                # Page 3：問題細項排行
                # ════════════════════════════════════════════
                pdf.add_page()
                pdf.section_title("問題細項排行")

                if detail_col and detail_col in df_filt.columns:
                    _det_v = df_filt[detail_col].value_counts()
                    rows_d = [[i+1, str(d)[:35], int(v), f"{int(v)/n_cur*100:.0f}%"]
                               for i, (d, v) in enumerate(_det_v.head(15).items())]
                    pdf.full_table(["排名","問題細項","件數","佔比"], rows_d, [15,120,20,25])

                # ════════════════════════════════════════════
                # Page 4：圖表（圓餅 + 機台）
                # ════════════════════════════════════════════
                pdf.add_page()
                pdf.section_title("數據可視化分析")

                if type_col and type_col in df_filt.columns:
                    _tc4 = df_filt[type_col].value_counts()
                    _total4 = _tc4.sum()
                    # 大圖（figsize 更寬），標籤改用圖例避免重疊
                    _f4, _a4 = _mplt.subplots(figsize=(9, 6))
                    _clrs4 = ["#060E9F","#FF5000","#FFCE00","#8EB9C9","#0076A9","#FAE0B8"]
                    _labels4 = [f"{k}（{int(v)}件）" for k, v in _tc4.items()]
                    wedges, texts, autotexts = _a4.pie(
                        list(_tc4.values),
                        labels=None,           # 不在扇形上顯示標籤，改用圖例
                        autopct=lambda p: f"{p:.0f}%" if p >= 5 else "",  # 小扇形不顯示%
                        colors=_clrs4[:len(_tc4)],
                        startangle=90,
                        pctdistance=0.75,
                        wedgeprops={"linewidth": 1.5, "edgecolor": "white"},
                    )
                    for _at in autotexts:
                        _at.set_fontsize(11)
                        _at.set_fontweight("bold")
                    _a4.legend(
                        wedges, _labels4,
                        loc="center left",
                        bbox_to_anchor=(1.0, 0.5),
                        fontsize=9,
                        frameon=False,
                    )
                    _a4.set_title(f"{period_label}　客訴類別分佈", fontsize=13, pad=12)
                    _f4.tight_layout()
                    pdf.embed_image(_f4, w=175, h=120)

                if machine_col and machine_col in df_filt.columns and not df_filt[machine_col].dropna().empty:
                    _mc4 = df_filt[machine_col].value_counts()
                    _labels_mc = [f"{k}（{int(v)}件）" for k, v in _mc4.items()]
                    _f5, _a5 = _mplt.subplots(figsize=(7, 5))
                    wedges5, texts5, autotexts5 = _a5.pie(
                        list(_mc4.values),
                        labels=None,
                        autopct=lambda p: f"{p:.0f}%" if p >= 5 else "",
                        colors=["#FF5000","#060E9F","#8EB9C9","#FFCE00"],
                        startangle=90,
                        pctdistance=0.72,
                        wedgeprops={"linewidth": 1.5, "edgecolor": "white"},
                    )
                    for _at5 in autotexts5:
                        _at5.set_fontsize(12)
                        _at5.set_fontweight("bold")
                    _a5.legend(
                        wedges5, _labels_mc,
                        loc="center left",
                        bbox_to_anchor=(1.0, 0.5),
                        fontsize=10,
                        frameon=False,
                    )
                    _a5.set_title(f"{period_label}　機台客訴佔比", fontsize=13, pad=12)
                    _f5.tight_layout()
                    pdf.embed_image(_f5, w=175, h=110)

                # ════════════════════════════════════════════
                # Page 5：趨勢 + 部門分析
                # ════════════════════════════════════════════
                pdf.add_page()
                pdf.section_title("客訴趨勢分析")

                _daily3 = df_filt.groupby(df_filt[date_col].dt.date).size().reset_index(name="件數")
                if len(_daily3) > 1:
                    _f6, _a6 = _mplt.subplots(figsize=(10, 4))
                    _a6.bar([str(d) for d in _daily3.iloc[:,0]], list(_daily3["件數"]),
                            color="#060E9F", edgecolor="white", linewidth=0.5)
                    _a6.set_title(f"{period_label}　每日件數趨勢", fontsize=13)
                    _a6.tick_params(axis="x", rotation=30, labelsize=8)
                    _a6.yaxis.set_major_locator(_MNL(integer=True))
                    _a6.set_ylabel("件數", fontsize=10)
                    _a6.grid(axis="y", alpha=0.3)
                    _f6.tight_layout()
                    pdf.embed_image(_f6, w=180, h=100)

                if dept_col and dept_col in df_filt.columns:
                    pdf.section_title("各部門件數分析")
                    _dp = df_filt[dept_col].replace("","未分配").value_counts()
                    rows_dp = [[i+1, str(d), int(v), f"{int(v)/n_cur*100:.0f}%"]
                                for i, (d, v) in enumerate(_dp.items())]
                    pdf.full_table(["排名","部門","件數","佔比"], rows_dp, [15,80,20,20])

                _pdf_bytes = bytes(pdf.output())
                st.session_state["_s4_pdf_bytes"] = _pdf_bytes
                st.session_state["_s4_pdf_label"] = period_label
                st.success(f"✅ PDF 已產生，共 {pdf.page_no()} 頁（{len(_pdf_bytes)//1024} KB）")
            except Exception as _e:
                import traceback
                st.error(f"PDF 產生失敗：{_e}")
                st.code(traceback.format_exc())

    if st.session_state.get("_s4_pdf_bytes"):
        _label = st.session_state.get("_s4_pdf_label", period_label)
        st.download_button(
            "⬇️ 下載完整分析 PDF（多頁）",
            data=st.session_state["_s4_pdf_bytes"],
            file_name=f"ECOCO_客訴分析_{_label.replace(' ','').replace('～','-').replace('/','-')}.pdf",
            mime="application/pdf",
            use_container_width=False,
            key="s4_dl_full_pdf",
        )

    # ── AI 口說報告 ────────────────────────────────────────────────
    st.markdown("---")
    st.markdown(f'<div class="s4-section">🎙️ AI 口說報告產生器</div>', unsafe_allow_html=True)
    rep_type = st.radio("報告類型", ["週會報告","月會報告","季報","年度報告"], horizontal=True, key="s4v3_rep")

    if st.button("🚀 產生 AI 口說報告", type="primary", key="s4v3_gen"):
        total_cur  = len(df_filt)
        total_prev = len(df_prev) if not df_prev.empty else None
        pct_chg    = pct_change(total_cur, total_prev) if total_prev else None

        type_summary = ""
        if type_col and type_col in df_filt.columns:
            _cvs = df_filt[type_col].value_counts()
            _pvs = df_prev[type_col].value_counts() if not df_prev.empty and type_col in df_prev.columns else pd.Series(dtype=int)
            for cat, cnt in _cvs.items():
                prev_cnt = int(_pvs.get(cat, 0))
                d = int(cnt) - prev_cnt
                pct_share = f"{int(cnt)/max(total_cur,1)*100:.0f}%"
                pline = f"（占 {pct_share}，較上期 {d:+d} 件，{pct_change(int(cnt),prev_cnt):+.1f}%）" if prev_cnt else f"（占 {pct_share}）"
                type_summary += f"- {cat}：{int(cnt)} 件{pline}\n"

        city_summary = ""
        if city_col and city_col in df_filt.columns:
            _cc = df_filt[city_col].value_counts()
            _pc = df_prev[city_col].value_counts() if not df_prev.empty and city_col in df_prev.columns else pd.Series(dtype=int)
            for city, cnt in _cc.head(8).items():
                prev_cnt = int(_pc.get(city, 0))
                d = int(cnt) - prev_cnt
                pct_share = f"{int(cnt)/max(total_cur,1)*100:.0f}%"
                pline = f"（占 {pct_share}，較上期 {d:+d} 件）" if prev_cnt else f"（占 {pct_share}）"
                city_summary += f"- {city}：{int(cnt)} 件{pline}\n"

        detail_summary = ""
        if detail_col and detail_col in df_filt.columns:
            _dv = df_filt[detail_col].value_counts()
            _dp = df_prev[detail_col].value_counts() if not df_prev.empty and detail_col in df_prev.columns else pd.Series(dtype=int)
            for det, cnt in _dv.head(8).items():
                prev_cnt = int(_dp.get(det, 0))
                d = int(cnt) - prev_cnt
                pct_share = f"{int(cnt)/max(total_cur,1)*100:.0f}%"
                pline = f"（占 {pct_share}，較上期 {d:+d} 件）" if prev_cnt else f"（占 {pct_share}）"
                detail_summary += f"- {det}：{int(cnt)} 件{pline}\n"

        station_summary = ""
        if station_col and station_col in df_filt.columns:
            for sta, cnt in df_filt[station_col].value_counts().head(5).items():
                station_summary += f"- {sta}：{int(cnt)} 件\n"

        dept_summary = ""
        if dept_col and dept_col in df_filt.columns:
            _dept = df_filt[dept_col].replace("","未分配").value_counts()
            _dprev = df_prev[dept_col].replace("","未分配").value_counts() if not df_prev.empty and dept_col in df_prev.columns else pd.Series(dtype=int)
            for dep, cnt in _dept.items():
                prev_cnt = int(_dprev.get(dep, 0))
                d = int(cnt) - prev_cnt
                pline = f"（較上期 {d:+d} 件）" if prev_cnt else ""
                dept_summary += f"- {dep}：{int(cnt)} 件{pline}\n"

        trend_word = "增加" if (pct_chg or 0) > 0 else ("減少" if (pct_chg or 0) < 0 else "持平")
        _upper_cmp = ""
        if pct_chg is not None:
            _upper_cmp = f"（對照期 {period_prev} 共 {total_prev} 件，本期{trend_word} {abs(pct_chg):.1f}%）"

        prompt = (
            f"你是 ECOCO 宜可可循環經濟客服部門的資深數據分析專員，擁有豐富的循環經濟、設備維運與客服管理知識。\n"
            f"請根據以下【完整數據摘要】，為本次{rep_type}撰寫一份詳盡的「口說報告」。\n\n"
            f"═══════════ 數據摘要 ═══════════\n"
            f"📅 分析區間：{period_label}{_upper_cmp}\n"
            f"📊 本期總進件：{total_cur} 件\n\n"
            f"【問題類型分布】：\n{type_summary or '（無資料）'}\n"
            f"【問題細項排行 TOP8】：\n{detail_summary or '（無資料）'}\n"
            f"【城市件數排行 TOP8】：\n{city_summary or '（無資料）'}\n"
            f"【熱門站點 TOP5】：\n{station_summary or '（無資料）'}\n"
            f"【各部門負責件數】：\n{dept_summary or '（無資料）'}\n"
            f"═══════════════════════════════\n\n"
            f"【報告段落結構】—— 嚴格按照以下六段撰寫，每段不少於 4 句，共約 800-1000 字：\n\n"
            f"━━ 一、開場白與執行摘要 ━━\n"
            f"• 用一句有力的話點出本期最重要的發現（數字+意義）\n"
            f"• 點明整體趨勢方向及幅度，說明是否達到預期目標\n"
            f"• 預告本次報告的三個核心重點\n\n"
            f"━━ 二、數據概覽與趨勢分析 ━━\n"
            f"• 詳細說明本期與上期總件數增減的具體數字與可能原因（季節性？活動推廣？設備老化？）\n"
            f"• 分析各問題類型的比例消長，哪類問題突增、哪類改善\n"
            f"• 以「每XX件客訴就有X件是...」方式說明佔比的實際意義\n\n"
            f"━━ 三、前三大痛點深度解析 ━━\n"
            f"• 針對問題細項排名前三，逐一分析：(1)件數與佔比 (2)主要發生城市/站點 (3)對用戶的實際影響 (4)與上期比較\n"
            f"• 深入推測問題根本原因（非表面描述），如硬體耗損、韌體問題、清運頻率不足等\n"
            f"• 說明若不及時處理可能引發的連鎖效應\n\n"
            f"━━ 四、城市、站點與機台分析 ━━\n"
            f"• 點出本期件數最高的城市及其熱門站點，分析地理集中原因\n"
            f"• 有顯著增加的區域要點出並推測原因，有改善的要給予肯定\n"
            f"• 說明機台類型（收瓶機/電池機）的問題分布是否均衡，有無特定機型問題集中\n\n"
            f"━━ 五、部門成效與改善追蹤 ━━\n"
            f"• 各部門（營運部、行銷部、資訊部）本期負責件數及較上期增減\n"
            f"• 哪些指標已有改善跡象（件數下降、比例縮小），哪些仍需持續關注\n"
            f"• 若有客訴反覆發生的情況，點名指出並說明影響\n\n"
            f"━━ 六、下階段行動建議 ━━\n"
            f"• 提出 4～5 項具體可執行的改善行動，需指定負責部門、執行時程、預期效果\n"
            f"• 建議本期需重點追蹤的 KPI 指標（例如：瓶蓋桶滿件數週降 20%）\n"
            f"• 以鼓勵性語氣結語，帶出下一期的期望目標\n\n"
            f"【撰寫規範】\n"
            f"✅ 繁體中文，語氣正式自然，如現場向長官口頭簡報\n"
            f"✅ 數字要說明「意義」：例如「機台問題佔49%，即每兩件客訴就有一件涉及硬體」\n"
            f"✅ 善用過渡詞：「首先」「值得特別注意的是」「從區域分布來看」「相較上期」「建議優先」\n"
            f"✅ 避免套語：「如上所述」「綜上所述」「總的來說」請用更具體的描述替代\n"
            f"✅ 報告總長度：800-1000 字，資訊完整、重點突出\n"
        )

        with st.spinner("AI 正在撰寫口說報告..."):
            try:
                import anthropic as _anth, os
                _api_key = (os.environ.get("ANTHROPIC_API_KEY","") or
                            str(st.secrets.get("ANTHROPIC_API_KEY","")))
                _client_ai = _anth.Anthropic(api_key=_api_key)
                _msg = _client_ai.messages.create(
                    model="claude-haiku-4-5-20251001",
                    max_tokens=2000,
                    messages=[{"role":"user","content":prompt}],
                )
                report_text = _msg.content[0].text
            except Exception as e:
                try:
                    report_text = generate_ai_summary_llm(df_filt, model_name="haiku")
                    report_text = f"【口說報告】\n\n{report_text}"
                except Exception:
                    report_text = f"⚠️ AI 暫時無法使用（{e}）\n\n數據摘要：\n\n{prompt}"

        st.text_area("📋 口說報告（可複製）", report_text, height=460, key="s4v3_report_out")

        dl_c1, dl_c2 = st.columns(2)
        dl_c1.download_button("⬇️ 下載口說報告（TXT）",
                              data=report_text.encode("utf-8"),
                              file_name=f"{period_label}_{rep_type}.txt",
                              mime="text/plain", key="s4v3_dl_txt",
                              use_container_width=True)
        try:
            _setup_cjk_font()
            import matplotlib.pyplot as _mplt
            _charts = {}
            if type_col and type_col in df_filt.columns:
                _tc2 = df_filt[type_col].value_counts()
                _f, _a = _mplt.subplots(figsize=(8,4))
                _colors = ["#060E9F","#FF5000","#FFCE00","#8EB9C9","#0076A9"]
                _a.pie(list(_tc2.values), labels=list(_tc2.index), autopct="%1.1f%%",
                       colors=_colors[:len(_tc2)], startangle=90)
                _a.set_title(f"{period_label} 客訴類別分佈")
                _b = io.BytesIO(); _f.savefig(_b, format="png", dpi=150, bbox_inches="tight")
                _mplt.close(_f); _charts["客訴類別分佈.png"] = _b.getvalue()
            if city_col and city_col in df_filt.columns:
                _cc2 = df_filt[city_col].value_counts().head(10)
                _f2, _a2 = _mplt.subplots(figsize=(8,4))
                _a2.bar(list(_cc2.index), list(_cc2.values), color="#FF5000")
                _a2.set_title(f"{period_label} 城市件數排行")
                _a2.yaxis.set_major_locator(_mplt.MaxNLocator(integer=True))
                _a2.tick_params(axis="x", rotation=15)
                _b2 = io.BytesIO(); _f2.savefig(_b2, format="png", dpi=150, bbox_inches="tight")
                _mplt.close(_f2); _charts["城市排行.png"] = _b2.getvalue()
            if detail_col and detail_col in df_filt.columns:
                _dc2 = df_filt[detail_col].value_counts().head(8)
                _f3, _a3 = _mplt.subplots(figsize=(8,4))
                _a3.barh(list(_dc2.index)[::-1], list(_dc2.values)[::-1], color="#060E9F")
                _a3.set_title(f"{period_label} TOP 8 問題細項")
                _b3 = io.BytesIO(); _f3.savefig(_b3, format="png", dpi=150, bbox_inches="tight")
                _mplt.close(_f3); _charts["問題細項排行.png"] = _b3.getvalue()
            _zbuf = io.BytesIO()
            with zipfile.ZipFile(_zbuf, "w", zipfile.ZIP_DEFLATED) as _zf:
                for _fn, _fb in _charts.items():
                    _zi = zipfile.ZipInfo(_fn); _zi.flag_bits |= 0x800
                    _zi.compress_type = zipfile.ZIP_DEFLATED; _zf.writestr(_zi, _fb)
                _zr = zipfile.ZipInfo(f"{period_label}_{rep_type}.txt")
                _zr.flag_bits |= 0x800; _zf.writestr(_zr, report_text.encode("utf-8"))
            dl_c2.download_button("⬇️ 下載圖表+報告（ZIP）",
                                  data=_zbuf.getvalue(),
                                  file_name=f"{period_label}_趨勢分析.zip",
                                  mime="application/zip", key="s4v3_dl_zip",
                                  use_container_width=True)
        except Exception as _ze:
            dl_c2.warning(f"ZIP 產生失敗：{_ze}")



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
        if st.button("📈 週/月/季/年度分析", use_container_width=True, type="primary" if st.session_state["menu"] == "趨勢分析" else "secondary"):
            st.session_state["menu"] = "趨勢分析"
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
            <div class="ecoco-card">
              <b>功能 4</b>：週/月/季/年度趨勢分析——從歷史紀錄聚合數據、趨勢對比、AI 口說報告產生器。
            </div>
            """,
            unsafe_allow_html=True,
        )
    elif menu == "上傳檔案區（分析區）":
        section_1()
    elif menu == "圖表與 AI 分析":
        section_2()
    elif menu == "趨勢分析":
        section_4()
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
