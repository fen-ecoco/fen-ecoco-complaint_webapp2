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


st.set_page_config(page_title="ECOCO 摰Ｚ迄??撟喳", page_icon="??", layout="wide")

TOPIC_DETAIL_MAP = {
    "APP雿輻??憿?": [
        "APP?恍憿舐內???啁???蝚?,
        "APP?振?蝛箇",
        "APP暺憿舐內?啣虜",
        "APP憭??啣虜?瘜?,
        "app?恍憿舐內???啁???蝚?,
        "app憭??啣虜?瘜?,
        "app暺憿舐內?啣虜",
        "app?振?蝛箇",
    ],
    "APP撣唾?閮剖???憿?": [
        "敹?撖Ⅳ/?⊥??身撖Ⅳ",
        "撣唾?鞈?靽格/閮剖?",
        "?⊥??交蝪∟?撽?蝣?,
        "APP?⊥??餃",
        "app?⊥??餃",
    ],
    "APP撣喳??餃??": [
        "APP?⊥??餃",
        "app?⊥??餃",
        "敹?撖Ⅳ/?⊥??身撖Ⅳ",
    ],
    "?芣??詨?憿???: [
        "??憭望?/憿舐內?航炊",
        "?⊥??脰?????",
        "雿輻閬?/?璇辣隤芣?",
        "?亥岷?芣??詨?????,
    ],
    "?暺??憿?": [
        "暺???仿?",
        "暺?芸撣唾?",
        "?敺?脤???暺?芾???,
    ],
    "璈??憿?": [
        "璈??銝剜/??",
        "暺???啣虜??嗅憛?,
        "???菜葫?啣虜",
        "??瘚??啣虜/?⊥?甇?虜??",
        "?Ｗ??啣虜憿舐內/?恍?啣虜",
        "撅亙葆?芯????啣虜??",
        "璈?嗆?/?∪???,
        "璈?蝬剛風/????",
        "璈蝬脰楝???憭望?",
        "璈擃情/?閬?瞏?,
        "蝬脰楝銝剜??蝛拙?",
        "璈??/?⊥???",
        "?蝬??摰孵",
        "??拙雿?嗥?/?餅?",
        "颲刻?憭望??啣虜?隤?,
        "璈???恍?⊥??餃",
        "?敺?脤???暺?芾???,
        "?Ｗ?镼踵撠暺?????,
        "?嗉?獢嗅歇皛?,
        "?????",
    ],
    "憿批恥??憿?": [
        "閮梢??啣?蝡?/閮剔?撱箄降",
        "?唾??芷撣唾?",
        "?湔?撣唾?",
        "?嗡?撱箄降",
        "??拐蝙?刻???,
        "?賊?瘣餃?閬???",
    ],
}

TYPE_OPTIONS = list(TOPIC_DETAIL_MAP.keys())
DETAIL_OPTIONS = [d for lst in TOPIC_DETAIL_MAP.values() for d in lst]

DEPT_OPTIONS = [
    "????, "???, "撱???, "鈭箄???, "銵??, 
    "鞈???, "隡???, "鞎∪???, "???, "蝮賜??恕"
]

DEPT_MAP = {
    "璈??憿?": "????,
    "璈?賊???": "????,
    "APP撣唾?閮剖???憿?": "鞈???,
    "APP雿輻??憿?": "鞈???,
    "APP撣喳??餃??": "鞈???,
    "?暺??憿?": "",
    "?芣??詨?憿???: "銵??,
    "憿批恥??憿?": "????,
}

# ?? ECOCO ???莎?Pantone 撠?嚗??????????????????????????????
BRAND_ORANGE  = "#FF5000"   # Pantone Orange 021 C  ??????
BRAND_BLUE    = "#060E9F"   # Pantone Blue 072 C    ??鞈???/ 銝餃???
BRAND_YELLOW  = "#FFCE00"   # Pantone 116 C         ??銵??
BRAND_LBLUE   = "#8EB9C9"   # Pantone 550 C
BRAND_BEIGE   = "#FAE0B8"   # Pantone P17-2 C
BRAND_TEAL    = "#0076A9"   # Pantone 7690 C
BRAND_WHITE   = "#FFFFFF"   # Pantone White C

# ?券??箏??莎?Plotly color_discrete_map ?剁?
DEPT_COLOR_MAP: dict[str, str] = {
    "????: BRAND_ORANGE,
    "銵??: BRAND_YELLOW,
    "鞈???: BRAND_BLUE,
    "???: BRAND_TEAL,
    "撱???: BRAND_LBLUE,
    "鈭箄???: BRAND_BEIGE,
    "隡???: "#A0C878",
    "鞎∪???: "#C8A0E0",
    "???: "#E0C8A0",
    "蝮賜??恕": "#A0E0C8",
    "?芸???:  "#CCCCCC",
    "":        "#CCCCCC",
}

# ????/ 璈急???脫?摨?
BRAND_PALETTE = [
    BRAND_BLUE, BRAND_ORANGE, BRAND_YELLOW,
    BRAND_LBLUE, BRAND_BEIGE, BRAND_TEAL,
]


# ?? ???Ⅱ靽?CJK 摮??舐嚗?頛??湛???????????????????????????
@st.cache_resource(show_spinner=False)
def _ensure_cjk_font() -> str:
    """??舐??CJK 摮?頝臬?嚗蝟餌絞瘝???頛 /tmp??""
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
    # 銝???/tmp
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
DEFAULT_HISTORY_SHEET_ID = "1Sqh_8bXtFw7jvmCPufTpStKxfIafDzwYJRlgc0HFBSs"
SHEET_CELL_CHAR_LIMIT = 49000

# 蝭頝臬?嚗?蝙?刻?蝔???? 蝪∪蝭.pptx嚗歇?函?撘?韏琿蝵莎?
TEMPLATE_PATH = Path(__file__).parent / "蝪∪蝭.pptx"


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
            font-size: 16px !important;
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
            font-size: 16px !important;
          }
          
          /* Use Noto Sans TC Medium (500) for everything ??no bold allowed */
          h1, h2, h3, h4, h5, h6, .ecoco-banner, strong, b, .side-title, section[data-testid="stSidebar"] .stButton > button {
            font-family: 'Noto Sans TC', 'Microsoft JhengHei', sans-serif !important;
            font-weight: 500 !important;
          }
          h1 { font-size: 36px !important; }
          h2 { font-size: 26px !important; }
          h3 { font-size: 21px !important; }
          .stDateInput input, .stTextInput input, .stSelectbox div, .stMultiSelect div {
            font-size: 16px !important;
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
            background: #ffce00;
            color:#333333; font-weight:500; margin-bottom: 12px;
            font-size: 36px !important;
          }
          .feature-title {
            color: #333333 !important;
            background: #fae0b8 !important;
            border-radius: 6px;
            padding: 8px 12px;
            font-size: 21px !important;
            font-weight: 500 !important;
            margin: 0 0 10px 0;
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
            background: #fae0b8;
            min-width: 19rem !important;
          }
          section[data-testid="stSidebar"] > div { width: 19rem !important; }
          
          /* Sidebar Text Overrides */
          .side-title {
            color: #333333 !important;
            font-weight: 500; font-size: 36px !important; margin-bottom: 8px;
          }
          .side-sub {
            color: #333333 !important;
            font-size: 26px !important; opacity: 0.9; margin-bottom: 14px;
          }
          
          /* Sidebar Buttons ??default = lightblue */
          section[data-testid="stSidebar"] .stButton > button {
            background-color: #fae0b8 !important;
            border-color: #e9cda4 !important;
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
          section[data-testid="stSidebar"] .stButton > button:focus-visible,
          section[data-testid="stSidebar"] .stButton > button:active,
          section[data-testid="stSidebar"] .stButton > button[kind="primary"],
          section[data-testid="stSidebar"] .stButton > button[data-testid="baseButton-primary"] {
            background-color: #FFFFFF !important;
            border-color: #FFFFFF !important;
            color: #333333 !important;
          }
          section[data-testid="stSidebar"] .stButton > button:hover *,
          section[data-testid="stSidebar"] .stButton > button:focus *,
          section[data-testid="stSidebar"] .stButton > button:focus-visible *,
          section[data-testid="stSidebar"] .stButton > button:active *,
          section[data-testid="stSidebar"] .stButton > button[kind="primary"] *,
          section[data-testid="stSidebar"] .stButton > button[data-testid="baseButton-primary"] * {
            color: #333333 !important;
          }
          
          /* Thicker scrollbar */
          ::-webkit-scrollbar { width: 10px; height: 10px; }
          ::-webkit-scrollbar { width: 0.5cm; height: 0.5cm; }
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
          .editor-toolbar-title {
            font-size: 12px !important;
            color: #333333 !important;
            margin: 6px 0 4px;
          }
          [data-testid="stDataFrame"], [data-testid="stDataEditor"] {
            border: 1.5px solid #555555 !important;
            border-radius: 6px !important;
            overflow-x: auto !important;
          }
          
          /* 蝘駁 arrow_down ?撱箏?蝷綽??踹??啣虜憿舐內蝝?摮?*/
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

    # 憿批恥??憿?
    if "閮餃?" in t and "?⊥?" in t:
        return "憿批恥??憿?", "?嗡?撱箄降"
    if any(k in t for k in ["銝???, "?漲", "??暻?, "銝???]):
        return "憿批恥??憿?", "?嗡?撱箄降"
    if any(k in t for k in ["?芷撣唾?", "閮駁"]):
        return "憿批恥??憿?", "?唾??芷撣唾?"
    if any(k in t for k in ["???Ⅳ", "?董??]) and any(k in t for k in ["霈", "?湔", "靽格"]):
        return "憿批恥??憿?", "?湔?撣唾?"
    if any(k in t for k in ["?湔?撣唾?", "?董??]):
        return "憿批恥??憿?", "?湔?撣唾?"
    if any(k in t for k in ["?啣?蝡?", "閮剔?", "撱箄降", "閮梢?"]):
        return "憿批恥??憿?", "閮梢??啣?蝡?/閮剔?撱箄降"
    if any(k in t for k in ["?閬?", "?釭", "?臬???]):
        return "憿批恥??憿?", "??拐蝙?刻???

    # APP撣唾?閮剖?
    if any(k in t for k in ["撽?蝣?, "隤?蝣?, "otp", "蝪∟?"]):
        if "敹?撖Ⅳ" in t:
            return "APP撣唾?閮剖???憿?", "敹?撖Ⅳ/?⊥??身撖Ⅳ"
        return "APP撣唾?閮剖???憿?", "?⊥??交蝪∟?撽?蝣?
    if any(k in t for k in ["靽格", "?湔", "?湔?"]) and any(k in t for k in ["撣唾?", "??", "?餉店", "?Ⅳ"]):
        return "APP撣唾?閮剖???憿?", "撣唾?鞈?靽格/閮剖?"

    # ?餃??
    if any(k in t for k in ["?餃", "?颱??脣"]) and any(k in t for k in ["?Ｗ?", "璈", "暺?"]) and any(k in t for k in ["?⊥?", "銝", "憭望?", "銝?"]):
        return "璈??憿?", "璈???恍?⊥??餃"
    if any(k in t for k in ["?⊥??餃", "銝?餃", "?颱??脣", "?餃憭望?", "?餃銝?"]):
        return "APP撣唾?閮剖???憿?", "APP?⊥??餃"

    # APP雿輻
    if "?舀??賊?" in t or ("app" in t and "憿舐內" in t and "0" not in t):
        return "APP雿輻??憿?", "APP?恍憿舐內???啁???蝚?
    if "憿舐內" in t and "銝泵" in t:
        return "APP雿輻??憿?", "APP?恍憿舐內???啁???蝚?
    if any(k in t for k in ["app?啣虜", "?", "頧?", "?湔"]):
        return "APP雿輻??憿?", "APP憭??啣虜?瘜?

    # 暺
    if "暺" in t and any(k in t for k in ["?芰敞蝛?, "?芸???, "瘝??亙董", "?芸撣?]):
        return "?暺??憿?", "暺?芸撣唾?"
    if ("暺" in t or "瘝暺? in t or "閮?" in t) and any(k in t for k in ["?芸", "瘝", "銝?", "瘝?", "瘝??]):
        return "?暺??憿?", "暺?芸撣唾?"
    if "暺" in t and any(k in t for k in ["??", "憭策", "憭"]):
        return "?暺??憿?", "暺???仿?"

    # ?芣???
    if any(k in t for k in ["?芣???, "????, "?", "摨?", "?萇", "撠???, "蟡典", "蟡典冗", "璇Ⅳ", "??]):
        if any(k in t for k in ["????", "???航炊", "蝟餌絞???湔", "撌脫??, "?", "??", "閬?"]):
            return "?芣??詨?憿???, "雿輻閬?/?璇辣隤芣?"
        if any(k in t for k in ["??", "??", "暺", "瘝??箸?蝣?]):
            return "?芣??詨?憿???, "?⊥??脰?????"
        if any(k in t for k in ["撌脖蝙??, "憭望?", "?航炊", "銝??, "?瑚???, "瘝?頝璇Ⅳ", "?獐銝??"]):
            return "?芣??詨?憿???, "??憭望?/憿舐內?航炊"
        if any(k in t for k in ["?亥岷", "蝝??, "?曆???, "?典"]):
            return "?芣??詨?憿???, "?亥岷?芣??詨?????
        return "?芣??詨?憿???, "?⊥??脰?????"

    # 璈??
    if any(k in t for k in ["??銝?, "?∩?"]) and "?怠?銝?" in t:
        return "璈??憿?", "璈?嗆?/?∪???
    if "撖嗥?嗅雿? in t or "?∪" in t or ("暺?" in t and "?∩?" in t) or "?∠" in t:
        return "璈??憿?", "??拙雿?嗥?/?餅?"
    if any(k in t for k in ["??憭活", "?⊥?颲刻?", "銝?湧＊蝷?, "銝＊蝷箇???, "颲刻?憭望?", "颲刻??啣虜"]):
        return "璈??憿?", "颲刻?憭望??啣虜?隤?
    if any(k in t for k in ["憿舐內0?賣????, "?蝬凋耨"]):
        return "璈??憿?", "璈?蝬剛風/????"
    if any(k in t for k in ["??", "閮剖?銝?", "銝雿輻", "?斗?", "??敹?, "瘝?", "?芷???, "??"]):
        return "璈??憿?", "璈??/?⊥???"
    if any(k in t for k in ["擃情銝", "皜?", "擃情"]):
        return "璈??憿?", "璈擃情/?閬?瞏?
    if any(k in t for k in ["?嗆?", "??閮", "瘝???, "lag", "璈?啣虜"]):
        return "璈??憿?", "璈?嗆?/?∪???
    if any(k in t for k in ["皛踹?, "?嗆遛", "皛踹"]):
        return "璈??憿?", "?嗉?獢嗅歇皛?
    if "??銝??迫" in t:
        return "璈??憿?", "??瘚??啣虜/?⊥?甇?虜??"

    # ?嗡???璈??
    if any(k in t for k in ["撅亙葆", "頛賊葆", "?喲葆"]) and any(k in t for k in ["銝?", "銝?", "?啣虜"]):
        return "璈??憿?", "撅亙葆?芯????啣虜??"
    if any(k in t for k in ["暺?", "暺??, "?Ｗ??啣虜", "?恍?啣虜", "??", "暺?"]):
        return "璈??憿?", "?Ｗ??啣虜憿舐內/?恍?啣虜"
    if any(k in t for k in ["蝬剛風", "蝬凋耨", "?蝬凋耨", "????"]):
        return "璈??憿?", "璈?蝬剛風/????"
    if any(k in t for k in ["蝬脰楝???憭望?", "???銝?, "???憭望?"]) and "璈" in t:
        return "璈??憿?", "璈蝬脰楝???憭望?"
    if any(k in t for k in ["蝬脰楝銝帘", "蝬脰楝銝剜"]):
        return "璈??憿?", "蝬脰楝銝剜??蝛拙?"
    if any(k in t for k in ["??", "蝘日?", "?菜葫??"]):
        return "璈??憿?", "???菜葫?啣虜"
    if any(k in t for k in ["?⊥???", "瘚??啣虜", "銝??"]):
        return "璈??憿?", "??瘚??啣虜/?⊥?甇?虜??"
    if any(k in t for k in ["?敺?暺?, "?芰暺", "?芾???]):
        return "璈??憿?", "?敺?脤???暺?芾???
    if any(k in t for k in ["銝剜", "??", "??璈?]):
        return "璈??憿?", "璈??銝剜/??"
    if any(k in t for k in ["??", "?瘝?", "???"]):
        return "璈??憿?", "?????"
    if "蝬?" in t and "銝" in t:
        return "璈??憿?", "?蝬??摰孵"

    return "憿批恥??憿?", "?嗡?撱箄降"


def parse_pdf_to_df(file_obj) -> pd.DataFrame:
    if pdfplumber is None:
        raise RuntimeError("?芸?鋆?pdfplumber嚗瘜圾??PDF??)
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
    raise ValueError(f"???excel / csv / pdf嚗?堆?{suffix or name}")


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


def _normalize_english_runs(text: str, mode: str) -> str:
    if not isinstance(text, str):
        return text
    repl = (lambda m: m.group(0).upper()) if mode == "upper" else (lambda m: m.group(0).lower())
    return re.sub(r"[A-Za-z]+", repl, text)


def normalize_problem_labels(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "??憿?" in out.columns:
        out["??憿?"] = out["??憿?"].map(lambda v: _normalize_english_runs(v, "upper"))
    if "??蝝圈?" in out.columns:
        out["??蝝圈?"] = out["??蝝圈?"].map(lambda v: _normalize_english_runs(v, "lower"))
    return out


def mask_sensitive_text(value):
    if not isinstance(value, str):
        return value
    text = re.sub(r"([A-Za-z0-9._%+-]{2})[A-Za-z0-9._%+-]*(@[A-Za-z0-9.-]+\.[A-Za-z]{2,})", r"\1***\2", value)
    text = re.sub(r"\b(09\d{2})\d{3}(\d{3})\b", r"\1***\2", text)
    text = re.sub(r"\b(0\d{1,2}-?\d{2,4})\d{3,4}(\d{2,4})\b", r"\1***\2", text)
    return text


def mask_sensitive_df(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    text_cols = [
        c for c in out.columns
        if any(k in str(c).lower() for k in ["?批捆", "銝餅", "憪?", "email", "mail", "?餉店", "??", "?舐窗"])
    ]
    for col in text_cols:
        out[col] = out[col].map(mask_sensitive_text)
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
    out = make_unique_columns(mask_sensitive_df(df.copy()))

    # ------ Preserve existing valid ??憿? + ??蝝圈? from source file ------
    existing_type   = out["??憿?"].copy()   if "??憿?" in out.columns else pd.Series([""] * len(out))
    existing_detail = out["??蝝圈?"].copy()   if "??蝝圈?" in out.columns else pd.Series([""] * len(out))

    # Drop internal columns before re-adding
    for c in ["??憿?", "??蝝圈?", "?詨?", "?券?", "?交?", "_ai_filled"]:
        if c in out.columns:
            out = out.drop(columns=[c])

    # Run auto-classification for every row
    preds = out.apply(
        lambda r: analyze_complaint(str(r.get(cfg.subject_col, "")), str(r.get(cfg.content_col, ""))),
        axis=1,
        result_type="expand",
    )
    preds.columns = ["??憿?", "??蝝圈?"]
    out = pd.concat([out, preds], axis=1)

    # ------ Merge: prefer original valid pair; fall back to AI prediction ------
    ai_filled_flags = []
    for idx in range(len(out)):
        orig_type   = str(existing_type.iloc[idx]).strip()
        orig_detail = str(existing_detail.iloc[idx]).strip()
        if _is_valid_pair(orig_type, orig_detail):
            # Original is valid ??keep it, NOT AI-filled
            out.iloc[idx, out.columns.get_loc("??憿?")] = orig_type
            out.iloc[idx, out.columns.get_loc("??蝝圈?")] = orig_detail
            ai_filled_flags.append(False)
        else:
            # Original missing/invalid ??use AI prediction, mark as AI-filled
            ai_filled_flags.append(True)

    out["_ai_filled"] = ai_filled_flags

    # Final guard: ensure detail always belongs to its topic
    out["??蝝圈?"] = out.apply(
        lambda r: r["??蝝圈?"] if r["??蝝圈?"] in TOPIC_DETAIL_MAP.get(r["??憿?"], [])
                  else TOPIC_DETAIL_MAP.get(r["??憿?"], ["?嗡?撱箄降"])[0],
        axis=1,
    )
    out["?詨?"] = False
    out["?券?"] = out["??憿?"].map(DEPT_MAP).fillna("")
    if cfg.date_col and cfg.date_col in out.columns:
        out["?交?"] = pd.to_datetime(out[cfg.date_col], errors="coerce")
    return normalize_problem_labels(out)


# ?? Google Sheets 甇瑕蝝??銋? ????????????????????????????????????????????
# Render ??蝣?甈⊿???皜征嚗蝙??Google Sheets 雿瘞訾??脣?敺垢??
# ???Streamlit Secrets 閮剖?嚗?
#   HISTORY_SHEET_ID = "<your_spreadsheet_id>"
#   [google_credentials]   ??service account JSON 甈?

def _get_gsheet_client():
    """敺憓??豢? st.secrets ?? gspread client??""
    try:
        import gspread as _gs
        from google.oauth2.service_account import Credentials as _Creds
    except ImportError:
        return None
    try:
        import os, json as _json
        # ?? 1. ?芸?霈 Render ?啣?霈 ??
        creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
        if creds_json:
            creds_dict = _json.loads(creds_json)
        else:
            # ?? 2. ?嚗璈?st.secrets嚗?????憭???
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
    """?甇瑕蝝?極雿”嚗仃????None?og_error=True ???航炊摮 session_state??""
    import os
    client = _get_gsheet_client()
    if client is None:
        if log_error:
            st.session_state["_gsheet_error"] = "?⊥?撱箇? Google API ???嚗?蝣箄? GOOGLE_CREDENTIALS_JSON ?啣?霈?澆?甇?Ⅱ嚗?
        return None
    try:
        sid = os.environ.get("HISTORY_SHEET_ID", "").strip()
        if not sid:
            try:
                sid = str(st.secrets.get("HISTORY_SHEET_ID", "")).strip()
            except Exception:
                sid = ""
        if not sid:
            sid = DEFAULT_HISTORY_SHEET_ID
        if not sid:
            if log_error:
                st.session_state["_gsheet_error"] = "?芾身摰?HISTORY_SHEET_ID ?啣?霈"
            return None
        ss = client.open_by_key(sid)
        try:
            ws = ss.worksheet("甇瑕蝝??)
            try:
                header = ws.row_values(1)
                if header[:5] != ["id", "created_at", "source_name", "rows", "data_ref"]:
                    ws.update(values=[["id", "created_at", "source_name", "rows", "data_ref"]], range_name="A1:E1")
            except Exception:
                pass
            st.session_state.pop("_gsheet_error", None)
            return ws
        except Exception:
            ws = ss.add_worksheet("甇瑕蝝??, rows=500, cols=6)
            ws.append_row(["id", "created_at", "source_name", "rows", "data_ref"])
            st.session_state.pop("_gsheet_error", None)
            return ws
    except Exception as e:
        err_str = str(e)
        if "PERMISSION_DENIED" in err_str or "403" in err_str:
            msg = (f"Google Sheets API 甈??航炊????Google Cloud Console 蝣箄?撌脣??剁?\n"
                   f"1. Google Sheets API\n2. Google Drive API\n"
                   f"?航炊嚗err_str[:200]}")
        elif "NOT_FOUND" in err_str or "404" in err_str:
            msg = f"閰衣?銵其?摮嚗D ?航?航炊嚗?{err_str[:200]}"
        else:
            msg = f"Google Sheets ????航炊嚗err_str[:300]}"
        if log_error:
            st.session_state["_gsheet_error"] = msg
        return None


def _sanitize_sheet_value(value, max_chars: int = SHEET_CELL_CHAR_LIMIT) -> str:
    if pd.isna(value):
        return ""
    text = str(value)
    if text.lower() in {"nan", "inf", "-inf", "infinity", "-infinity"}:
        return ""
    return text[:max_chars]


def _sanitize_df_for_sheet(df: pd.DataFrame, max_chars: int = SHEET_CELL_CHAR_LIMIT) -> pd.DataFrame:
    out = df.copy()
    out = out.replace([float("inf"), float("-inf")], pd.NA)
    out = out.astype(object).where(pd.notna(out), "")
    mapper = lambda v: _sanitize_sheet_value(v, max_chars=max_chars)
    if hasattr(out, "map"):
        return out.map(mapper)
    return out.apply(lambda col: col.map(mapper))


def _history_data_sheet_name(item_id: str) -> str:
    safe_id = re.sub(r"[^0-9A-Za-z_\\-]+", "_", str(item_id))[:80]
    return f"history_{safe_id}"


def _write_history_data_sheet(spreadsheet, worksheet_name: str, df: pd.DataFrame):
    clean_df = _sanitize_df_for_sheet(df)
    values = [clean_df.columns.tolist()] + clean_df.values.tolist()
    rows = max(len(values), 1)
    cols = max(len(clean_df.columns), 1)
    try:
        ws_data = spreadsheet.worksheet(worksheet_name)
        ws_data.clear()
        ws_data.resize(rows=max(rows, 100), cols=max(cols, 10))
    except Exception:
        ws_data = spreadsheet.add_worksheet(title=worksheet_name, rows=max(rows, 100), cols=max(cols, 10))
    if values:
        ws_data.update(values=values, range_name="A1")
    return ws_data


def _worksheet_to_dataframe(ws) -> pd.DataFrame:
    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()
    header = values[0]
    rows = values[1:]
    width = len(header)
    normalized = [(row + [""] * width)[:width] for row in rows]
    return pd.DataFrame(normalized, columns=header)


def save_history(df: pd.DataFrame, source_name: str, existing_id: str = "") -> tuple[Path, str, str]:
    today = datetime.now().strftime("%Y%m%d")
    ts = existing_id if existing_id else datetime.now().strftime("%Y%m%d_%H%M%S")
    output_name = f"{today}_??.xlsx"
    excel_bytes = to_excel_bytes(df)
    data_sheet_name = _history_data_sheet_name(ts)

    meta = {
        "id": ts, "created_at": datetime.now().isoformat(timespec="seconds"),
        "source_name": source_name, "output_name": output_name,
        "output_path": "", "rows": int(len(df)),
    }

    saved_to_gsheet = False
    # 1. Google Sheets嚗偶銋?
    ws = _history_sheet(log_error=True)
    if ws is not None:
        try:
            _write_history_data_sheet(ws.spreadsheet, data_sheet_name, df)
            if existing_id:
                rows = ws.get_all_values()
                for i, row in enumerate(rows[1:], start=2):
                    if row and row[0] == existing_id:
                        ws.delete_rows(i); break
            ws.append_row([ts, meta["created_at"], source_name, str(len(df)), f"sheet:{data_sheet_name}"])
            st.session_state.pop("_gsheet_error", None)
            saved_to_gsheet = True
        except Exception as e:
            st.session_state["_gsheet_error"] = f"甇瑕蝝?神??Google Sheets 憭望?嚗str(e)[:300]}"

    if saved_to_gsheet:
        if "_history_cache" not in st.session_state:
            st.session_state["_history_cache"] = {}
        st.session_state["_history_cache"][ts] = {"meta": meta, "excel_bytes": excel_bytes}

    # 2. ?祆?蝤?嚗靽?瑼?嚗?雿甇瑕蝝??皞?
    output_path = HISTORY_DIR / f"{ts}_{output_name}"
    try:
        output_path.write_bytes(excel_bytes)
    except Exception:
        pass
    return output_path, output_name, ts


def load_history() -> list[dict]:
    import base64
    merged: dict[str, dict] = {}

    # ?芸? Google Sheets 霈?風?脩????踹??脩垢?∠???蝬脤?畾???
    ws = _history_sheet()
    if ws is None:
        st.session_state["_history_cache"] = {}
        return []
    try:
        for row in ws.get_all_values()[1:]:
            if not row or not row[0]:
                continue
            rid = row[0]
            created_at = row[1] if len(row) > 1 else ""
            sname = row[2] if len(row) > 2 else ""
            rows_str = row[3] if len(row) > 3 else "0"
            data_ref = row[4] if len(row) > 4 else ""
            meta = {
                "id": rid, "created_at": created_at,
                "source_name": sname,
                "rows": int(rows_str) if rows_str.isdigit() else 0,
                "output_name": f"{rid}_??.xlsx", "output_path": "",
            }
            merged[rid] = meta
            if "_history_cache" not in st.session_state:
                st.session_state["_history_cache"] = {}
            if rid not in st.session_state["_history_cache"] and data_ref:
                try:
                    if data_ref.startswith("sheet:"):
                        data_ws = ws.spreadsheet.worksheet(data_ref.split(":", 1)[1])
                        hist_df = _worksheet_to_dataframe(data_ws)
                        excel_bytes = to_excel_bytes(hist_df)
                    else:
                        excel_bytes = base64.b64decode(data_ref)
                    st.session_state["_history_cache"][rid] = {
                        "meta": meta,
                        "excel_bytes": excel_bytes,
                    }
                except Exception:
                    pass
    except Exception:
        st.session_state["_history_cache"] = {}
        return []

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
        return "?桀?瘝??臬?????
    total = len(df)
    type_count = df["??憿?"].value_counts()
    detail_count = df["??蝝圈?"].value_counts()
    top_type = type_count.index[0]
    top_type_count = int(type_count.iloc[0])
    top_detail = detail_count.index[0]
    top_detail_count = int(detail_count.iloc[0])
    dept_text = ""
    if "?券?" in df.columns and not df["?券?"].dropna().empty:
        dept_top = df["?券?"].replace("", "?芸???).value_counts().head(3)
        dept_text = "嚗?.join([f"{k} {int(v)} 隞? for k, v in dept_top.items()])
    detail_lines = []
    for name, count in detail_count.head(5).items():
        detail_lines.append(f"{name} {int(count)} 隞?)
    detail_text = "嚗?.join(detail_lines)
    return (
        f"1) ?芸???嚗甈∪ {total} 隞塚?銝餃????箝top_type}?top_type_count} 隞塚??? {top_type_count/total:.1%}?n"
        f"2) 蝝圈?隤芣?嚗?擃蝝圈??箝top_detail}?top_detail_count} 隞塚?TOP5 ??{detail_text}?n"
        f"3) ?券?閫撖?{dept_text or '?桀??⊥?蝣粹?甈??臬霈'}?n"
        "4) ?郊?方?嚗??芸?瑼Ｘ擃蝝圈??臬?葉?潛摰?暺身??????瘚?嚗蒂瘥?餈??臬?雁靽柴暑??蝟餌絞?啣??n"
        "5) 撱箄降銵?嚗誑 TOP3 ??撱箇??孵?隞餃?嚗?摰?鞎祇???閮?????梯蕭頩斗?璅?
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
    sample = df[["??憿?", "??蝝圈?", "?券?"]].head(300).to_dict(orient="records")
    payload = {
        "total_rows": len(df),
        "top_types": df["??憿?"].value_counts().head(6).to_dict(),
        "top_details": df["??蝝圈?"].value_counts().head(10).to_dict(),
        "sample_rows": sample,
    }
    prompt = (
        "雿摰Ｘ??釭??憿批????函?擃葉?撓??-5暺?暺??澆?蝎曄陛嚗?
        "?: 擃????賣?楊?券??芸??孵?撱箄降????銝?\n"
        f"{json.dumps(payload, ensure_ascii=False)}\n"
        "隢?亙??綽?1. 蝡??????梢? (憒?敺摰寧?敺靘? 2. ??憿??敦???梢?(?擃?撣???
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


def to_pdf_bytes(df: pd.DataFrame, source_name: str = "", download_count: int = 1) -> bytes:
    """Generate PDF using fpdf2 + Noto CJK for Traditional Chinese support."""
    from fpdf import FPDF
    from fpdf.enums import XPos, YPos
    import os, glob

    # ?? ?曉???憭??頝臬? ??
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
    # 雿輻 _ensure_cjk_font 蝣箔???典???
    if not font_path:
        font_path = _ensure_cjk_font()

    table_df = df.copy()
    drop_cols = [c for c in ["?詨?"] if c in table_df.columns]
    table_df = table_df.drop(columns=drop_cols).fillna("")

    PAGE_W_MM = 277.0
    WIDE_COLS   = {"?冽?批捆", "銝餅", "??銝餅"}
    MEDIUM_COLS = {"??蝝圈?", "??憿?", "?脖辣?交?", "?交???", "蝡??迂", "??蝝圈?"}
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
        """蝮桀?摮??游???拙?甈祝"""
        for fs in range(max_size, min_size-1, -1):
            pdf_obj.set_font(FONT, size=fs)
            if pdf_obj.get_string_width(text) <= max_width - 1:
                return fs
        return min_size

    def draw_page_label():
        label = f"{source_name or '??瑼?'}  {datetime.now().strftime('%Y/%m/%d')}  蝚?{download_count} 甈?
        pdf.set_xy(8, 4)
        pdf.set_text_color(80, 80, 80)
        pdf.set_font(FONT, size=7)
        pdf.cell(0, 5, safe_text(label), align="L")
        pdf.set_xy(10, 10)

    draw_page_label()

    # 銵券嚗?葬撠誑?拇?甈祝嚗?
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

    # ?? 鞈???獢?蝯曹? row_h 擃漲嚗?摮???銵???
    pdf.set_font(FONT, size=FS_CELL)
    col_list = list(table_df.columns)

    for i, (_, row) in enumerate(table_df.iterrows()):
        fill_rgb = (0xEB, 0xF4, 0xFA) if i % 2 == 0 else (0xFF, 0xFF, 0xFF)

        # ?? 皞???嚗?文?擗征?質???嚗??
        cell_texts = {
            col: safe_text(
                " ".join(str(row[col]).split())   # 憭征?賢?雿萇?桐?蝛箸
                .replace("  ", " ").strip()
            )
            for col in col_list
        }

        # ?? 蝎曄Ⅱ閮?瘥?銵 ??
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

        # ?? ??瑼Ｘ ??
        if pdf.get_y() + row_h > pdf.page_break_trigger:
            pdf.add_page()
            draw_page_label()
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

        # ?? Step 1嚗??急?甈?摨 + 摰獢?嚗絞銝 row_h嚗??
        x_cursor = x0
        for col in col_list:
            cw = col_widths[col]
            # 摨憛急遛?湔
            pdf.set_fill_color(*fill_rgb)
            pdf.rect(x_cursor, y0, cw, row_h, style="F")
            # 憭?蝺??湔擃漲嚗?
            pdf.set_draw_color(0x99, 0x99, 0x99)
            pdf.rect(x_cursor, y0, cw, row_h, style="D")
            x_cursor += cw

        # ?? Step 2嚗???摮?multi_cell ?芰??嚗??急?蝺???
        x_cursor = x0
        for col in col_list:
            cw = col_widths[col]
            val = cell_texts[col]
            pdf.set_xy(x_cursor + 0.5, y0 + 0.5)   # 0.5mm ?抒葬 padding
            pdf.set_fill_color(*fill_rgb)
            pdf.multi_cell(
                cw - 1, ROW_H, val,
                border=0,          # 銝獢?嚗歇??Step 1 ?怠末嚗?
                align="L", fill=False,
                new_x=XPos.RIGHT, new_y=YPos.TOP,
                max_line_height=ROW_H,
            )
            x_cursor += cw

        pdf.set_draw_color(0, 0, 0)  # ?Ｗ儔暺
        pdf.set_xy(x0, y0 + row_h)

    # ?偏
    pdf.set_y(-12)
    pdf.set_font(FONT, size=6)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(0, 6, safe_text(f"ECOCO 摰Ｚ迄???勗?  ??{len(table_df)} 蝑? ?Ｗ?交?嚗datetime.now().strftime('%Y/%m/%d')}"), align="C")

    return bytes(pdf.output())

def _setup_cjk_font() -> None:
    """閮剖? matplotlib 銝剜?摮?嚗蝙??_ensure_cjk_font ??摮?頝臬???""
    import matplotlib.font_manager as fm
    import os

    # ?? 撌脰身摰?撠梁?亥?????
    current = plt.rcParams.get("font.family", "")
    if current and "sans-serif" not in str(current) and current != ["DejaVu Sans"]:
        return

    # ?? ?芸?雿輻 _ensure_cjk_font 蝣箔?摮?摮 ??
    fp = _ensure_cjk_font()
    if fp and os.path.exists(fp):
        try:
            fm.fontManager.addfont(fp)
            plt.rcParams["font.family"] = fm.FontProperties(fname=fp).get_name()
            plt.rcParams["axes.unicode_minus"] = False
            return
        except Exception:
            pass

    # ?? 1. 撌脩頝臬?嚗buntu / Render / Debian嚗??
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

    # ?? 2. ????歇摰?摮???CJK ??
    CJK_KEYWORDS = [
        "Noto Sans CJK", "Noto Serif CJK", "Noto Sans TC",
        "MingLiU", "PMingLiU", "Microsoft JhengHei",
        "WenQuanYi", "Droid Sans Fallback", "AR PL UMing",
    ]
    fm._load_fontmanager(try_read_cache=False)   # 撘瑕???
    for kw in CJK_KEYWORDS:
        for f in fm.fontManager.ttflist:
            if kw.lower() in f.name.lower():
                plt.rcParams["font.family"] = f.name
                plt.rcParams["axes.unicode_minus"] = False
                return

    # ?? 3. ?蝯?fallback嚗撠?????蝣???
    plt.rcParams["axes.unicode_minus"] = False


def build_chart_pack(df: pd.DataFrame,
                     color_bar: str | None = None,
                     color_pie: list[str] | None = None,
                     color_hbar: str | None = None) -> dict[str, bytes]:
    """Build chart PNG images for download/PPT.
    color_bar  : ??憿??湔?????None = 靘????? ??亙銝 hex 撘瑕憟
    color_pie  : 璈?????耦憿 list嚗one = BRAND_PALETTE
    color_hbar : ?之蝝圈?璈急????莎?None = BRAND_BLUE
    """
    _setup_cjk_font()

    data = df.copy()
    # ?? 璈憿?甇?????寡?/?寡?蝡????嗥璈???
    if "璈憿?" in data.columns:
        data["璈憿?"] = data["璈憿?"].apply(
            lambda v: "?嗥璈? if ("?寡?" in str(v) or "?嗥" in str(v))
            else ("?餅?璈? if "?餅?" in str(v) else str(v))
        )
    stats = data["??憿?"].value_counts().rename_axis("??憿?").reset_index(name="隞嗆")
    stats["?曉?瘥?] = (stats["隞嗆"] / max(stats["隞嗆"].sum(), 1) * 100).round(1)
    detail_stats = data["??蝝圈?"].value_counts().reset_index().head(10)
    detail_stats.columns = ["??蝝圈?", "隞嗆"]
    d = detail_stats.sort_values("隞嗆", ascending=True)

    # ?? resolve colors ??
    _pie_palette  = color_pie  if color_pie  else BRAND_PALETTE
    _hbar_color   = color_hbar if color_hbar else BRAND_BLUE

    def _bar_colors_for(series):
        if color_bar:
            return [color_bar] * len(series)
        return [DEPT_COLOR_MAP.get(DEPT_MAP.get(t, ""), BRAND_ORANGE) for t in series]

    # 1) ??憿??湔???
    fig1, ax1 = plt.subplots(figsize=(8, 4.5))
    bc = _bar_colors_for(stats["??憿?"])
    ax1.bar(stats["??憿?"], stats["隞嗆"], color=bc)
    ax1.set_title("??憿???")
    ax1.set_ylabel("隞嗆")
    ax1.yaxis.set_major_locator(plt.MaxNLocator(integer=True))
    ax1.tick_params(axis="x", rotation=20)
    for i, r in stats.iterrows():
        ax1.text(i, r["隞嗆"], f'{float(r["?曉?瘥?]):.1f}%', ha="center", va="bottom", fontsize=9)
    fig1.tight_layout()
    b1 = io.BytesIO(); fig1.savefig(b1, format="png", dpi=180); plt.close(fig1)

    # 2) 璈????
    fig2, ax2 = plt.subplots(figsize=(6.2, 4.5))
    df_machine = data[data["??憿?"] == "璈??憿?"].copy()
    if df_machine.empty:
        ax2.text(0.5, 0.5, "?⊥??啁????, ha="center", va="center", transform=ax2.transAxes)
        pie_counts = None
    else:
        def _get_mtype(row):
            txt = str(row.get("?冽?批捆", "")) + " " + str(row.get("銝餅", ""))
            if "?寡?" in txt: return "?寡?蝡?
            if "?餅?" in txt: return "?餅?璈?
            return "?嗥璈?
        df_machine["璈璈?"] = df_machine.apply(_get_mtype, axis=1)
        pie_counts = df_machine["璈璈?"].value_counts()
        pc = _pie_palette[:len(pie_counts)]
        wedges, texts, autotexts = ax2.pie(
            pie_counts.values, labels=pie_counts.index, autopct="%1.1f%%",
            colors=pc, wedgeprops=dict(linewidth=1.5, edgecolor="white"),
        )
        for at in autotexts: at.set_fontsize(10)
    ax2.set_title("璈??憿???")
    fig2.tight_layout()
    b2 = io.BytesIO(); fig2.savefig(b2, format="png", dpi=180); plt.close(fig2)

    # 3) ?之蝝圈?璈急??? ?? 撘瑕??銝餉? #060E9F嚗?詨摨?
    fig3, ax3 = plt.subplots(figsize=(8, 4.5))
    _hbar = _hbar_color if _hbar_color else "#060E9F"
    ax3.barh(d["??蝝圈?"], d["隞嗆"], color=_hbar)
    ax3.set_title("?之??蝝圈???")
    ax3.set_xlabel("?賊?(隞?")
    ax3.set_ylabel("?")
    # 撘瑕?湔?餃漲嚗辣?詨??箸?賂?
    from matplotlib.ticker import MultipleLocator
    ax3.xaxis.set_major_locator(MultipleLocator(1))
    ax3.xaxis.set_minor_locator(MultipleLocator(1))
    ax3.set_xlim(left=0)
    fig3.tight_layout()
    b3 = io.BytesIO(); fig3.savefig(b3, format="png", dpi=180); plt.close(fig3)

    # 4) Dashboard ??
    fig4 = plt.figure(figsize=(14, 5))
    gs = fig4.add_gridspec(1, 3)
    a1 = fig4.add_subplot(gs[0, 0])
    a2 = fig4.add_subplot(gs[0, 1])
    a3 = fig4.add_subplot(gs[0, 2])
    a1.bar(stats["??憿?"], stats["隞嗆"], color=bc)
    a1.set_title("??憿???")
    a1.yaxis.set_major_locator(MultipleLocator(1))
    a1.tick_params(axis="x", rotation=18)
    if pie_counts is None:
        a2.text(0.5, 0.5, "?⊥??啗???, ha="center", va="center", transform=a2.transAxes)
    else:
        a2.pie(pie_counts.values, labels=pie_counts.index, autopct="%1.1f%%",
               colors=_pie_palette[:len(pie_counts)],
               wedgeprops=dict(linewidth=1.5, edgecolor="white"))
    a2.set_title("璈????")
    a3.barh(d["??蝝圈?"], d["隞嗆"], color=_hbar)
    a3.xaxis.set_major_locator(MultipleLocator(1))
    a3.set_xlim(left=0)
    a3.set_title("?之蝝圈?")
    fig4.tight_layout()
    b4 = io.BytesIO(); fig4.savefig(b4, format="png", dpi=180); plt.close(fig4)

    return {
        "chart_??憿???.png": b1.getvalue(),
        "chart_璈????.png": b2.getvalue(),
        "chart_?之??蝝圈?.png": b3.getvalue(),
        "chart_dashboard.png":    b4.getvalue(),
    }


def build_ppt_bytes(stats: pd.DataFrame, ai_text: str, source_name: str,
                    template_path: str = "",
                    chart_pack: Optional[dict[str, bytes]] = None) -> bytes:
    """
    Build a PPT presentation.
    ?芸?雿輻???蝭嚗?曆??啣?敺瑽遣蝚血? ECOCO ??憸冽??敶梁???
    """
    from pptx.util import Emu, Inches, Pt
    from pptx.enum.text import PP_ALIGN

    BLUE   = RGBColor(0x06, 0x0E, 0x9F)
    ORANGE = RGBColor(0xFF, 0x50, 0x00)
    WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
    BEIGE  = RGBColor(0xFA, 0xE0, 0xB8)
    DARK   = RGBColor(0x22, 0x22, 0x22)
    LGRAY  = RGBColor(0xE8, 0xF1, 0xF5)
    FONT   = "MingLiU"   # 蝝唳?擃?

    # ?? ?岫頛蝭 ??
    _tpath = Path(template_path) if template_path else TEMPLATE_PATH
    use_template = _tpath.exists()
    prs = Presentation(str(_tpath)) if use_template else Presentation()
    if not use_template:
        # 閮剖??蔣?之撠撖祈撟?16:9
        prs.slide_width  = Inches(13.33)
        prs.slide_height = Inches(7.5)

    SW = prs.slide_width
    SH = prs.slide_height

    # ?? 撠極????
    def blank_layout():
        for lay in prs.slide_layouts:
            if lay.name.lower() in ("blank", "蝛箇"):
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
        """?? ECOCO ????嚗??脤璇?+ 璅?嚗?""
        add_rect(slide, 0, 0, SW/914400, 1.05, BLUE)
        add_text(slide, title_text, 0.3, 0.08, 9.0, 0.55,
                 20, bold=True, color=WHITE)
        if subtitle_text:
            add_text(slide, subtitle_text, 0.3, 0.62, 10.0, 0.38,
                     11, color=BEIGE)

    def delete_shape(sp):
        sp.element.getparent().remove(sp.element)

    # ????????????????????????????????????????????????????????
    #  雿輻蝭嚗?撖急?摮”?潦???
    # ????????????????????????????????????????????????????????
    if use_template:
        slides = list(prs.slides)

        # --- 撠 (slide 0) ---
        # 蝭?嚗椰?渲??脤?選???蔭??嚗??喳銝?摮?嚗?
        #   Shape;99  ??銝餅?憿???晞?(l??.66, t??.18)
        #   Shape;98  ???交?/鞈???(l??.14, t??.48)
        #   Shape;96  ???砍??摨摮?(l??.67, t??.04)
        s0 = slides[0]
        for sp in s0.shapes:
            if not sp.has_text_frame:
                continue
            l_in = sp.left / 914400
            t_in = sp.top  / 914400
            raw  = sp.text_frame.text.strip()

            # ?? 銝餅?憿???x>5" 銝?y<3"嚗?
            if l_in > 5.0 and t_in < 3.0:
                tf = sp.text_frame
                tf.clear()
                p = tf.paragraphs[0]
                run = p.add_run()
                run.text = "摰Ｚ迄??蝪∪"
                run.font.name  = FONT
                run.font.bold  = True
                run.font.size  = Pt(32)
                run.font.color.rgb = RGBColor(0x16, 0x2B, 0x7E)

            # ?? ?交?/鞈?甈???x>5" 銝?y ??3~5"嚗?
            elif l_in > 5.0 and 3.0 <= t_in < 5.0:
                tf = sp.text_frame
                tf.clear()
                for label, val in [
                    ("?勗??交?", datetime.now().strftime("%Y/%m/%d")),
                    ("?勗?鞈?", source_name),
                ]:
                    p = tf.add_paragraph()
                    run = p.add_run()
                    run.text = f"{label}:{val}"
                    run.font.name  = FONT
                    run.font.bold  = True
                    run.font.size  = Pt(18)
                    run.font.color.rgb = RGBColor(0x1A, 0x2A, 0x7F)

            # ?? ?砍????x>6" 銝?y>=5" ??憛怨??嚗?
            elif l_in > 6.0 and t_in >= 4.8:
                pass   # 靽??見?蝡??∩遢???砍??

        def _fill_slide(slide, title_txt, chart_key_list, add_table=True):
            SWi = prs.slide_width  / 914400
            SHi = prs.slide_height / 914400

            # ?湔璅???嚗?撠??萄?嚗?
            for sp in slide.shapes:
                if sp.has_text_frame:
                    txt = sp.text_frame.text
                    if any(k in txt for k in ("摰Ｚ迄????", "璈??雿?", "璈?敦??,
                                               "摰Ｚ迄??", "????", "20260")):
                        tf = sp.text_frame; tf.clear()
                        p = tf.paragraphs[0]
                        run = p.add_run()
                        run.text = title_txt
                        run.font.name = FONT
                        run.font.bold = True
                        run.font.size = Pt(16)
                        run.font.color.rgb = BLUE

            # ?園??暹? Table / Picture 雿蔭敺?歹?皜征?摰對?
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

            # ?? ?”?嚗?蝙?函??砌?雿?蝵殷??血??典摰漣璅???
            if chart_pack:
                if add_table:
                    # slide 2嚗?憿???嚗”?澆椰??+ ?”?喳?
                    # ?箏?摨扳?嚗?銵冽?喳
                    chart_fixed = [
                        (6.2, 1.15, SWi - 6.5, SHi - 1.4),   # ??憿???
                    ]
                else:
                    # slide 3嚗??啁敦??嚗椰?喳??曆?撘萄?
                    chart_fixed = [
                        (0.3,              1.15, (SWi - 0.6) / 2,       SHi - 1.4),
                        (0.3 + (SWi-0.6)/2 + 0.15, 1.15, (SWi-0.6)/2, SHi - 1.4),
                    ]

                for idx, key in enumerate(chart_key_list):
                    if key not in chart_pack:
                        continue
                    if idx < len(pic_rects):
                        # 蝭??雿??????典?憪?蝵?
                        add_img(slide, chart_pack[key],
                                *[v / 914400 for v in pic_rects[idx]])
                    elif idx < len(chart_fixed):
                        # 蝭瘝?雿? ???典摰漣璅?
                        add_img(slide, chart_pack[key], *chart_fixed[idx])

            # ?? ?遣鞈?銵冽 ??
            if add_table:
                # 憒?蝭??銵冽雿蔭撠望窒?剁??血??身撌血
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
                for ci, hdr in enumerate(["??憿?", "隞嗆", "?曉?瘥?, "甇詨惇?券?"]):
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
                    try:   pct = f'{float(r["?曉?瘥?]):.1f}%'
                    except: pct = f'{r["?曉?瘥?]}%'
                    dept = str(r.get("甇詨惇?券?", ""))
                    vals = [str(r["??憿?"]), str(int(r["隞嗆"])), pct, dept]
                    # 靘?憟???脩????
                    dept_hex = DEPT_COLOR_MAP.get(dept, "")
                    if dept_hex:
                        r_bg = RGBColor(
                            int(dept_hex[1:3], 16),
                            int(dept_hex[3:5], 16),
                            int(dept_hex[5:7], 16),
                        )
                        # 瘛∪?嚗毽?亦??80%
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
                        f"{source_name} 摰Ｚ迄????",
                        ["chart_??憿???.png"],
                        add_table=True)
        if len(slides) >= 3:
            _fill_slide(slides[2],
                        f"{source_name} 璈?敦????,
                        ["chart_?之??蝝圈?.png", "chart_璈????.png"],
                        add_table=False)

    # ????????????????????????????????????????????????????????
    #  敺瑽遣嚗??砌?摮??
    # ????????????????????????????????????????????????????????
    else:
        SWi = SW / 914400   # EMU ??inches
        SHi = SH / 914400

        # ?? Slide 1: 撠 ??
        s0 = prs.slides.add_slide(blank_layout())
        add_rect(s0, 0, 0, SWi, SHi, BLUE)      # ?刻??
        add_text(s0, "ECOCO 摰Ｚ迄??蝪∪",
                 1.0, SHi*0.25, SWi-2, 1.2, 36, bold=True,
                 color=WHITE, align=PP_ALIGN.CENTER)
        add_text(s0, f"?勗??交?嚗datetime.now().strftime('%Y/%m/%d')}",
                 1.0, SHi*0.52, SWi-2, 0.5, 16,
                 color=BEIGE, align=PP_ALIGN.CENTER)
        add_text(s0, f"鞈?靘?嚗source_name}",
                 1.0, SHi*0.64, SWi-2, 0.5, 14,
                 color=BEIGE, align=PP_ALIGN.CENTER)
        add_text(s0, "?∠?璈隞賣????,
                 1.0, SHi*0.82, SWi-2, 0.4, 13,
                 color=WHITE, align=PP_ALIGN.CENTER)

        # ?? Slide 2: ??憿??? ??
        s1 = prs.slides.add_slide(blank_layout())
        add_header(s1, f"摰Ｚ迄???? ??{source_name}",
                   f"?勗??交?嚗datetime.now().strftime('%Y/%m/%d')}?鞈?靘?嚗source_name}")
        # 銵冽嚗椰??
        rows_n = min(len(stats) + 1, 10)
        tbl_left = Inches(0.3); tbl_top = Inches(1.15)
        tbl_w    = Inches(5.8); tbl_h   = Inches(SHi - 1.4)
        tb = s1.shapes.add_table(rows_n, 4, tbl_left, tbl_top, tbl_w, tbl_h).table
        tb.columns[0].width = Inches(2.2)
        tb.columns[1].width = Inches(0.9)
        tb.columns[2].width = Inches(1.0)
        tb.columns[3].width = Inches(1.5)
        for ci, hdr in enumerate(["??憿?", "隞嗆", "?曉?瘥?, "甇詨惇?券?"]):
            c = tb.cell(0, ci); c.text = hdr
            c.fill.solid(); c.fill.fore_color.rgb = BLUE
            for para in c.text_frame.paragraphs:
                para.alignment = PP_ALIGN.CENTER
                for run in para.runs:
                    run.font.bold = True; run.font.color.rgb = WHITE
                    run.font.size = Pt(12); run.font.name = FONT
        for ri, (_, r) in enumerate(stats.head(rows_n - 1).iterrows(), 1):
            try:   pct = f'{float(r["?曉?瘥?]):.1f}%'
            except: pct = f'{r["?曉?瘥?]}%'
            vals = [str(r["??憿?"]), str(r["隞嗆"]), pct,
                    str(r.get("甇詨惇?券?", ""))]
            bg = LGRAY if ri % 2 == 0 else BEIGE
            for ci, v in enumerate(vals):
                c = tb.cell(ri, ci); c.text = v
                c.fill.solid(); c.fill.fore_color.rgb = bg
                for para in c.text_frame.paragraphs:
                    para.alignment = PP_ALIGN.CENTER
                    for run in para.runs:
                        run.font.size = Pt(11); run.font.color.rgb = DARK
                        run.font.name = FONT
        # ?”嚗??
        if chart_pack and "chart_??憿???.png" in chart_pack:
            add_img(s1, chart_pack["chart_??憿???.png"],
                    6.25, 1.15, SWi - 6.55, SHi - 1.4)

        # ?? Slide 3: 璈?敦??????
        s2 = prs.slides.add_slide(blank_layout())
        add_header(s2, f"璈?敦??????{source_name}",
                   f"?勗??交?嚗datetime.now().strftime('%Y/%m/%d')}")
        half_w = (SWi - 0.6) / 2
        ch_t = 1.15; ch_h = SHi - 1.4
        if chart_pack and "chart_璈????.png" in chart_pack:
            add_img(s2, chart_pack["chart_璈????.png"],
                    0.3, ch_t, half_w, ch_h)
        if chart_pack and "chart_?之??蝝圈?.png" in chart_pack:
            add_img(s2, chart_pack["chart_?之??蝝圈?.png"],
                    0.3 + half_w + 0.15, ch_t, half_w, ch_h)

    # ?? ?蝯?AI ?????蔣????楝敺????
    s_ai = prs.slides.add_slide(blank_layout())
    SWi2 = prs.slide_width  / 914400
    SHi2 = prs.slide_height / 914400
    # ???
    add_rect(s_ai, 0, 0, SWi2, 1.05, BLUE)
    add_text(s_ai, "AI ??????",
             0.3, 0.08, 9.0, 0.55, 20, bold=True, color=WHITE)
    add_text(s_ai,
             f"鞈?靘?嚗source_name}??Ｗ?交?嚗datetime.now().strftime('%Y/%m/%d')}",
             0.3, 0.65, 10.5, 0.35, 11, color=BEIGE)
    # 璈撌阡?獢?憌?
    add_rect(s_ai, 0.25, 1.15, 0.08, SHi2 - 1.35, ORANGE)
    # AI ??獢?
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
        is_head = line[:2] in ('1)', '2)', '3)', '4)', '5)', '銝??, '鈭?, '銝?)
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
    # 敹???? spreadsheets ??drive scope
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
            f"?⊥?摮?閰衣?銵剁?ID: {spreadsheet_id}嚗n"
            f"隢Ⅱ隤歇撠岫蝞”?梁蝯佗?{credentials_json.get('client_email', '?')}\n"
            f"???航炊嚗e}"
        )
    clean_df = _sanitize_df_for_sheet(df)
    values = [clean_df.columns.tolist()] + clean_df.values.tolist()
    try:
        ws = sh.worksheet(worksheet_name)
        ws.clear()
        ws.resize(rows=max(len(values), 100), cols=max(len(clean_df.columns), 10))
    except Exception:
        ws = sh.add_worksheet(
            title=worksheet_name,
            rows=max(len(values), 100),
            cols=max(len(clean_df.columns), 10),
        )
    if values:
        ws.update(values=values, range_name="A1")
    return ws.url if hasattr(ws, 'url') else ""


def section_1():
    st.markdown('<div class="feature-title">?銝嚗?獢??唾????</div>', unsafe_allow_html=True)
    st.markdown("<div class='ecoco-card'>?舀銝 excel / csv / pdf嚗??蒂?Ｗ??憿???憿敦??/div>", unsafe_allow_html=True)

    # File info badge ??no long text, just a compact pill with truncated name
    if st.session_state.get("_uploaded_bytes") and st.session_state.get("_uploaded_name"):
        fname_short = st.session_state['_uploaded_name']
        if len(fname_short) > 30:
            fname_short = fname_short[:14] + "..." + fname_short[-12:]
        st.markdown(
            f"<span class='file-badge'>&#128196; {fname_short}</span>",
            unsafe_allow_html=True
        )

    uploaded = st.file_uploader("銝?唳?獢?, type=["xlsx", "xls", "csv", "pdf"], key="uploader")
    # Persist file bytes across menu switches
    if uploaded is not None:
        if uploaded.name != st.session_state.get("_uploaded_name"):
            st.session_state.pop("_editing_history_id", None)
            st.session_state.pop("_saved_history_id", None)
        st.session_state["_uploaded_bytes"] = uploaded.read()
        st.session_state["_uploaded_name"] = uploaded.name
        st.session_state["_uploaded_type"] = uploaded.type

    # Restore from session if user switched tabs and came back
    if uploaded is None and st.session_state.get("_uploaded_bytes") is not None:
        saved_name = st.session_state.get("_uploaded_name", "file")
        buf = io.BytesIO(st.session_state["_uploaded_bytes"])
        df_raw_bytes = load_input_file(buf, filename=saved_name)
        st.caption(f"撌脰???{saved_name}嚗?閮敺拙?嚗?鞈?蝑嚗len(df_raw_bytes)}")
        df_raw = make_unique_columns(df_raw_bytes)
        uploaded_name = saved_name
    elif uploaded is not None:
        fname = st.session_state.get("_uploaded_name", uploaded.name)
        df_raw = make_unique_columns(load_input_file(
            io.BytesIO(st.session_state["_uploaded_bytes"]), filename=fname
        ))
        uploaded_name = uploaded.name
        st.caption(f"撌脰???{uploaded.name}嚗????賂?{len(df_raw)}")
    else:
        if "analysis_df" not in st.session_state:
            st.info("隢??單?獢?憪???)
            return
        # Already analysed, show results without needing the raw file
        df_raw = None
        uploaded_name = st.session_state.get("source_name", "")

    if df_raw is not None:
        cols = list(df_raw.columns)
        if not cols:
            st.warning("瑼?瘝??舐甈???)
            return

        st.markdown("##### ???祟?貉?甈?閮剖?")
        subject_col = st.selectbox("?冽憛怠神?蜓憿?雿?, options=cols, index=0)
        content_col = st.selectbox("?冽?批捆甈?", options=cols, index=min(1, len(cols) - 1))
        date_opt = ["(??"] + cols
        date_col = st.selectbox("?交?甈?嚗憛恬?", options=date_opt, index=0)
        cfg = AnalysisConfig(subject_col=subject_col, content_col=content_col,
                             date_col=None if date_col == "(??" else date_col)

        if st.button("????", type="primary"):
            st.session_state["analysis_subject_col"] = subject_col
            st.session_state["analysis_content_col"] = content_col
            with st.spinner("甇???鞈?嚗???頛之??蝔?.."):
                st.session_state["analysis_df"] = analyze_dataframe(df_raw.copy(), cfg)
            st.session_state["source_name"] = uploaded_name

    if "analysis_df" not in st.session_state:
        return
    df = st.session_state["analysis_df"]
    subject_col = st.session_state.get("analysis_subject_col", "銝餅")
    content_col = st.session_state.get("analysis_content_col", "?冽?批捆")
    c1, c2, c3 = st.columns([2, 2, 2])
    keyword = c1.text_input("蝭拚嚗??萄?嚗蜓憿??批捆嚗?)
    filter_type = c2.multiselect("蝭拚嚗?憿???, options=TYPE_OPTIONS, default=[])
    
    valid_details = DETAIL_OPTIONS
    if filter_type:
        valid_details = []
        for t in filter_type:
            valid_details.extend(TOPIC_DETAIL_MAP.get(t, []))
            
    filter_detail = c3.multiselect("蝭拚嚗?憿敦??, options=valid_details, default=[])

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
        show = show[show["??憿?"].isin(filter_type)]
    if filter_detail:
        show = show[show["??蝝圈?"].isin(filter_detail)]

    st.markdown('<div class="editor-toolbar-title">?舐楊頛舀?閮”嚗?港???+ ??蝺刻摩嚗?/div>', unsafe_allow_html=True)

    # ---- AI憛怠璅內 ---
    ai_col = "_ai_filled"
    MARKER_COL = "AI璅?"  # kept for save compatibility only
    has_ai_col = ai_col in show.columns
    n_ai = 0
    if has_ai_col:
        n_ai = int(show[ai_col].fillna(False).astype(bool).sum())

    if n_ai > 0:
        st.markdown(
            f"""
            <div style='background:#fff5f5; border:1px solid #ffb3b3; border-radius:8px;
                        padding:8px 14px; margin-bottom:8px; font-size:0.85rem;'>
              <b style='color:#cc0000;'>??AI ?芸?璅?</b>嚗 <b style='color:#cc0000;'>{n_ai} 蝑?/b> ??甈?蝛箇???
              撌脩 AI ?寞?摰Ｚ迄?批捆?芸???憛怠??
              隢?撠嗾蝑撠?憒?靽格隢?亙銵冽銝凋??????????脣?靽格?Ⅱ隤?
            </div>
            """,
            unsafe_allow_html=True
        )

    st.caption("? ?湔?刻”?潔葉銝??豢???憿? / ??蝝圈?嚗矽?游???暺?????脣?靽格??)

    tool_add, tool_add_btn, tool_del, tool_del_btn, tool_toggle = st.columns([2.5, 1, 2.5, 1, 1.2])
    new_col_name = tool_add.text_input("?啣??渡?甈?", value="", key="editor_new_col", placeholder="頛詨甈??迂")
    if tool_add_btn.button("?啣?甈?", key="editor_add_col", use_container_width=True):
        col_name = new_col_name.strip()
        if not col_name:
            st.warning("隢撓?交?雿?蝔晞?)
        elif col_name in st.session_state["analysis_df"].columns:
            st.warning("甈?撌脣??具?)
        else:
            st.session_state["analysis_df"][col_name] = ""
            st.session_state.pop("editor_table", None)
            st.rerun()
    protected_cols = {"?詨?", MARKER_COL, ai_col}
    deletable_cols = [c for c in st.session_state["analysis_df"].columns if c not in protected_cols]
    del_col_name = tool_del.selectbox("?詨?甈?", options=deletable_cols, key="editor_delete_col")
    if tool_del_btn.button("?芷?湔?", key="editor_del_col", use_container_width=True):
        if del_col_name:
            st.session_state["analysis_df"] = st.session_state["analysis_df"].drop(columns=[del_col_name], errors="ignore")
            st.session_state.pop("editor_table", None)
            st.rerun()
    if tool_toggle.button("?詨?/??", key="toggle_all_btn", help="?券??瘨??, use_container_width=True):
        all_sel = bool(df["?詨?"].all()) if "?詨?" in df.columns and not df.empty else False
        st.session_state["analysis_df"]["?詨?"] = not all_sel
        st.session_state.pop("editor_table", None)
        st.rerun()

    batch_c1, batch_c2, batch_c3, batch_c4 = st.columns([2, 2, 1.5, 1.2])
    batch_type = batch_c1.selectbox("?寞活??憿?", ["(銝???"] + TYPE_OPTIONS, key="batch_type_sel")
    valid_batch_det = ["(銝???"]
    if batch_type != "(銝???":
        valid_batch_det += TOPIC_DETAIL_MAP.get(batch_type, [])
    batch_detail = batch_c2.selectbox("?寞活??蝝圈?", valid_batch_det, key="batch_cat_sel")

    # ???閬＊蝷箇?甈?嚗Ⅱ靽??祇?? MARKER_COL 甇?Ⅱ?
    display_cols = [c for c in show.columns if c not in (ai_col, MARKER_COL)]
    if "?詨?" in display_cols:
        display_cols = ["?詨?"] + [c for c in display_cols if c != "?詨?"]
    show_display = show[display_cols].reset_index(drop=True)

    # ?啣?銝甈???閮策 AI 憛怠????
    if has_ai_col:
        flags = show[ai_col].reset_index(drop=True)
        marker_vals = flags.map(lambda x: "潃?AI憛怠神)" if x else "")
    else:
        marker_vals = [""] * len(show_display)
        
    insert_idx = 1 if "?詨?" in show_display.columns else 0
    show_display.insert(insert_idx, MARKER_COL, marker_vals)

    edited = st.data_editor(
        show_display,
        use_container_width=True,
        num_rows="dynamic",
        hide_index=True,
        column_config={
            "?詨?": st.column_config.CheckboxColumn("?詨?", help="?暸閬甈∟?????),
            MARKER_COL: st.column_config.TextColumn("?酉", disabled=True),
            "??憿?": st.column_config.SelectboxColumn(options=TYPE_OPTIONS, required=True),
            "??蝝圈?": st.column_config.SelectboxColumn(options=DETAIL_OPTIONS, required=True),
            "?券?": st.column_config.SelectboxColumn(options=DEPT_OPTIONS),
        },
        key="editor_table",
    )

    # ?脣????刻”?潔???
    sv_col1, sv_col2, sv_col3 = st.columns([2, 2, 6])
    if sv_col1.button("? ?脣?靽格", use_container_width=True):
        full_df = st.session_state["analysis_df"].copy()
        # Drop the AI marker column and ?詨? before saving back
        save_edited = edited.drop(columns=["?詨?", MARKER_COL], errors="ignore")
        full_df.update(save_edited)
        # Clear _ai_filled flags for all saved rows (user has confirmed)
        if "_ai_filled" in full_df.columns:
            full_df["_ai_filled"] = False
        st.session_state["analysis_df"] = full_df
        # Also push to drafts list
        src_name = st.session_state.get("source_name", "?芸??)
        if "_draft_list" not in st.session_state:
            st.session_state["_draft_list"] = []
        # Avoid duplicate same name drafts ??update existing
        draft_ids = [d["name"] for d in st.session_state["_draft_list"]]
        if src_name not in draft_ids:
            st.session_state["_draft_list"].insert(0, {"name": src_name, "df": full_df.copy()})
        else:
            for d in st.session_state["_draft_list"]:
                if d["name"] == src_name:
                    d["df"] = full_df.copy()
        existing_id = st.session_state.get("_editing_history_id") or st.session_state.get("_saved_history_id", "")
        _, _, history_id = save_history(full_df, src_name, existing_id=existing_id)
        st.session_state["_saved_history_id"] = history_id
        if st.session_state.get("_gsheet_error"):
            st.warning(st.session_state["_gsheet_error"])
        else:
            st.success(f"撌脣摮src_name}??銝血歇?甇瑕蝝??)

    # 撌脣摮?蝔踹?銵?
    if st.session_state.get("_draft_list"):
        st.markdown("---")
        st.markdown("##### 撌脣摮??阮")
        for idx, draft in enumerate(st.session_state["_draft_list"]):
            d_col1, d_col2, d_col3, d_col4 = st.columns([5, 1, 1, 1])
            d_col1.markdown(
                f"<div style='padding-top:0.45rem; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-weight:600;'>"
                f"?? {draft['name']}</div>",
                unsafe_allow_html=True
            )
            if d_col2.button("[頛]", key=f"draft_load_{idx}", use_container_width=True):
                st.session_state["analysis_df"] = draft["df"].copy()
                st.session_state["source_name"] = draft["name"]
                st.success(f"撌脰??乓draft['name']}???舐匱蝥楊頛胯?)
            if d_col3.button("[靽格]", key=f"draft_edit_{idx}", use_container_width=True):
                st.session_state["analysis_df"] = draft["df"].copy()
                st.session_state["source_name"] = draft["name"]
                st.rerun()
            if d_col4.button("[X]", key=f"draft_del_{idx}", use_container_width=True):
                st.session_state["_draft_list"].pop(idx)
                st.rerun()

    if batch_c3.button("憟?暸??, type="primary", use_container_width=True):
        if "?詨?" not in edited.columns or not edited["?詨?"].any():
            st.warning("隢??刻”?澆?暸閬???鞈???")
        else:
            mask = edited["?詨?"] == True
            if batch_type != "(銝???":
                edited.loc[mask, "??憿?"] = batch_type
                edited.loc[mask, "?券?"] = edited.loc[mask, "??憿?"].map(DEPT_MAP).fillna("")
            if batch_detail != "(銝???":
                edited.loc[mask, "??蝝圈?"] = batch_detail
            # Auto-fix rows whose detail mismatches topic
            edited["??蝝圈?"] = edited.apply(
                lambda r: r["??蝝圈?"] if r["??蝝圈?"] in TOPIC_DETAIL_MAP.get(r["??憿?"], []) else TOPIC_DETAIL_MAP.get(r["??憿?"], ["?嗡?撱箄降"])[0],
                axis=1,
            )
            st.session_state["analysis_df"] = edited.copy()
            st.session_state.pop("editor_table", None)
            st.session_state["_batch_applied"] = True
            st.rerun()
            
    if st.session_state.pop("_batch_applied", False):
        st.success("撌脣??冽甈∠楊頛胯?)
        
    if batch_c4.button("?芷?暸??, use_container_width=True):
        if "?詨?" not in edited.columns or not edited["?詨?"].any():
            st.warning("隢??刻”?澆?暸閬?斤?鞈???")
        else:
            st.session_state["analysis_df"] = edited[edited["?詨?"] != True].copy()
            st.session_state.pop("editor_table", None)
            st.success("撌脣?文?詨???)
            st.rerun()

    final_df = st.session_state["analysis_df"]
    
    st.markdown("#### 銝???蝯? (銝?敺?飛瑼甇瑕蝝??")
    dl_format = st.radio("?豢?銝??澆?", ["Excel", "CSV", "PDF"], horizontal=True)
    
    def on_download():
        existing_id = st.session_state.pop("_editing_history_id", "") or st.session_state.get("_saved_history_id", "")
        _, _, history_id = save_history(final_df, st.session_state.get("source_name", "unknown"), existing_id=existing_id)
        st.session_state["_saved_history_id"] = history_id
        st.session_state["history_saved_msg"] = True

    if dl_format == "Excel":
        out_name = f"{datetime.now().strftime('%Y%m%d')}_??.xlsx"
        data_bytes = to_excel_bytes(final_df)
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif dl_format == "CSV":
        out_name = f"{datetime.now().strftime('%Y%m%d')}_??.csv"
        data_bytes = to_csv_bytes(final_df)
        mime = "text/csv"
    else:
        out_name = f"{datetime.now().strftime('%Y%m%d')}_????pdf"
        try:
            dl_key = f"_pdf_download_count_{st.session_state.get('source_name', 'unknown')}"
            st.session_state[dl_key] = int(st.session_state.get(dl_key, 0)) + 1
            data_bytes = to_pdf_bytes(final_df, st.session_state.get("source_name", "unknown"), st.session_state[dl_key])
            mime = "application/pdf"
        except Exception as e:
            st.error(f"PDF ?Ｙ??航炊: {e}")
            data_bytes = b""
            mime = "application/pdf"

    st.download_button(
        label=f"? 銝? {dl_format} ?澆???",
        data=data_bytes,
        file_name=out_name,
        mime=mime,
        on_click=on_download
    )
    
    if st.session_state.get("history_saved_msg"):
        if st.session_state.get("_gsheet_error"):
            st.warning(st.session_state["_gsheet_error"])
        else:
            st.success("瑼?撌脖?頛?銝西??摮甇瑕蝝??)
        st.session_state["history_saved_msg"] = False

    st.markdown("#### ?????Ｗ")
    summary_text = generate_ai_summary(final_df)
    st.text_area("??蝯???", summary_text, height=120)
    st.download_button(
        "銝?????嚗xt嚗?,
        data=summary_text.encode("utf-8"),
        file_name=f"{datetime.now().strftime('%Y%m%d')}_????.txt",
        mime="text/plain",
    )

    with st.expander("銝??Google Sheet"):
        st.write("隢?靘?Service Account JSON ??Spreadsheet ID")
        cred_file = st.file_uploader("Google Service Account JSON", type=["json"], key="gcp_json")
        spreadsheet_id = st.text_input("Spreadsheet ID")
        ws_name = st.text_input("Worksheet ?迂", value=datetime.now().strftime("%Y%m%d_??"))
        if st.button("銝 Google Sheet"):
            if not cred_file or not spreadsheet_id:
                st.error("隢?銝 JSON 銝血‵撖?Spreadsheet ID??)
            else:
                try:
                    credentials_json = json.loads(cred_file.getvalue().decode("utf-8"))
                    upload_to_google_sheet(final_df, credentials_json, spreadsheet_id, ws_name)
                    st.success(f"??撌脖??喳 Google Sheet 撌乩?銵剁?{ws_name}")
                    st.info(f"?? Service Account嚗credentials_json.get('client_email', '')}")
                except PermissionError as e:
                    st.error(str(e))
                    st.warning("?? 隢 Google 閰衣?銵典銝???具??銝 Service Account email嚗蒂蝯虫??楊頛航???)
                except Exception as e:
                    st.error(f"銝憭望?嚗e}")


def render_charts_from_stats(stats: pd.DataFrame, df: pd.DataFrame, key_prefix: str = ""):
    """Render interactive Plotly charts with per-chart color pickers."""

    # ?? 憿閮剖? expander ??????????????????????????????????????????
    kp = key_prefix or "main"
    with st.expander("? 隤踵?”憿嚗?靽格嚗?, expanded=False):
        ca, cb, cc = st.columns(3)
        # ??憿??湔????身???券????脯??暸敺???株
        use_single_bar = ca.checkbox("?湔??蝙?典銝憿", key=f"{kp}_cb_bar")
        c_bar_single   = ca.color_picker("?湔?????, value=BRAND_ORANGE, key=f"{kp}_cp_bar") if use_single_bar else None

        # ?????憭???敶Ｙ蝡矽??
        pie_c1 = cb.color_picker("????蝚??莎?銝鳴?", value=BRAND_BLUE,   key=f"{kp}_cp_pie1")
        pie_c2 = cb.color_picker("????蝚??莎?甈∴?", value=BRAND_ORANGE, key=f"{kp}_cp_pie2")
        pie_c3 = cb.color_picker("????蝚???,       value=BRAND_LBLUE,  key=f"{kp}_cp_pie3")

        c_hbar = cc.color_picker("蝝圈?璈急?????, value=BRAND_BLUE, key=f"{kp}_cp_hbar")

    custom_pie   = [pie_c1, pie_c2, pie_c3] + BRAND_PALETTE[3:]
    custom_hbar  = c_hbar

    # ?? ??????靘?Plotly + matplotlib ?梁嚗????????????????
    df_machine = df[df["??憿?"] == "璈??憿?"].copy()
    m_stats = None
    if not df_machine.empty:
        def _gmt(row):
            txt = str(row.get("?冽?批捆", "")) + " " + str(row.get("銝餅", ""))
            if "?寡?" in txt: return "?寡?蝡?
            if "?餅?" in txt: return "?餅?璈?
            return "?嗥璈?
        df_machine["璈璈?"] = df_machine.apply(_gmt, axis=1)
        m_stats = df_machine["璈璈?"].value_counts().reset_index()
        m_stats.columns = ["璈?", "隞嗆"]

    detail_stats = df["??蝝圈?"].value_counts().reset_index().head(10)
    detail_stats.columns = ["??蝝圈?", "隞嗆"]

    c1, c2, c3 = st.columns(3)

    # ?? ??嚗?憿??璇? ????????????????????????????????????
    if use_single_bar:
        fig1 = px.bar(stats, x="??憿?", y="隞嗆", text="?曉?瘥?,
                      title="??憿???", color_discrete_sequence=[c_bar_single])
        fig1.update_traces(marker_color=c_bar_single)
    else:
        fig1 = px.bar(stats, x="??憿?", y="隞嗆",
                      color="甇詨惇?券?", text="?曉?瘥?, title="??憿???",
                      color_discrete_map=DEPT_COLOR_MAP)
    fig1.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    fig1.update_layout(height=420, yaxis=dict(tickformat="d", nticks=6),
                       xaxis_title="?", yaxis_title="?賊?(隞?",
                       margin=dict(t=45, b=0))
    c1.plotly_chart(fig1, use_container_width=True, key=f"{kp}_fig1")

    # ?? ??嚗??啣?擗? ????????????????????????????????????????
    if m_stats is not None:
        cmap = {row["璈?"]: custom_pie[i % len(custom_pie)]
                for i, row in m_stats.iterrows()}
        fig2 = px.pie(m_stats, names="璈?", values="隞嗆",
                      title="璈??蝝啣?瘥?", hole=0.3,
                      color="璈?", color_discrete_map=cmap)
        fig2.update_traces(texttemplate="%{percent:.1%}", textinfo="percent+label")
        fig2.update_layout(height=420, margin=dict(t=45, b=0, l=0, r=0))
        c2.plotly_chart(fig2, use_container_width=True, key=f"{kp}_fig2")
    else:
        c2.info("?⊥??啁???)

    # ?? ??嚗?憭抒敦?帖璇? ????????????????????????????????????
    fig3 = px.bar(detail_stats, x="隞嗆", y="??蝝圈?",
                  orientation="h", title="?之??蝝圈???",
                  color_discrete_sequence=[custom_hbar])
    fig3.update_traces(marker_color=custom_hbar)
    fig3.update_layout(height=420, yaxis={"categoryorder": "total ascending"},
                       xaxis=dict(tickformat="d", nticks=6),
                       xaxis_title="?賊?(隞?", yaxis_title="?",
                       margin=dict(t=45, b=0, l=0, r=0))
    c3.plotly_chart(fig3, use_container_width=True, key=f"{kp}_fig3")

    # ?? ??嗉?賊??脣???session_state 靘?PPT/ZIP 雿輻 ????????
    st.session_state[f"chart_colors_{kp}"] = {
        "bar":  c_bar_single if use_single_bar else None,
        "pie":  custom_pie,
        "hbar": custom_hbar,
    }


def render_charts(df: pd.DataFrame, key_prefix: str = ""):
    if df.empty:
        st.info("沒有資料紀錄")
        return
    date_cols = [c for c in df.columns if "?交?" in c or "date" in c.lower()]
    if date_cols:
        dcol = date_cols[0]
        try:
            df[dcol] = pd.to_datetime(df[dcol], errors="coerce")
            valid_dates = df[dcol].dropna()
            if not valid_dates.empty:
                min_d = valid_dates.min().date()
                max_d = valid_dates.max().date()
                st.markdown("##### ???交????)
                c_d1, c_d2 = st.columns(2)
                start_d = c_d1.date_input("韏瑕??交?", value=min_d, min_value=min_d, max_value=max_d, key=f"{key_prefix}_sd")
                end_d   = c_d2.date_input("蝯??交?", value=max_d, min_value=min_d, max_value=max_d, key=f"{key_prefix}_ed")
                df = df[(df[dcol].dt.date >= start_d) & (df[dcol].dt.date <= end_d)]
        except Exception:
            pass

    if "??憿?" not in df.columns or "??蝝圈?" not in df.columns:
        st.info("沒有資料紀錄")
        return

    stats = df["??憿?"].value_counts().rename_axis("??憿?").reset_index(name="隞嗆")
    if stats.empty:
        st.info("沒有資料紀錄")
        return
    stats["?曉?瘥?] = (stats["隞嗆"] / max(stats["隞嗆"].sum(), 1) * 100).round(1)
    stats["甇詨惇?券?"] = stats["??憿?"].map(DEPT_MAP).fillna("?芸???)

    c1, c2, c3 = st.columns(3)
    
    fig1 = px.bar(
        stats, x="??憿?", y="隞嗆", color="甇詨惇?券?", text="?曉?瘥?, title="??憿???",
        color_discrete_sequence=["#FF5000", "#060E9F", "#FFCE00", "#8EB9C9", "#0076A9", "#FAE0B8"]
    )
    fig1.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    fig1.update_layout(height=400, xaxis_title="?", yaxis_title="?賊?(隞?",
                       yaxis=dict(tickformat="d", nticks=6))
    c1.plotly_chart(fig1, use_container_width=True, key=f"{key_prefix}_fig1" if key_prefix else None)

    df_machine = df[df["??憿?"] == "璈??憿?"].copy()
    if not df_machine.empty:
        def get_machine_type(row):
            txt = str(row.get("?冽?批捆", "")) + " " + str(row.get("銝餅", ""))
            if "?寡?" in txt: return "?寡?蝡?
            if "?餅?" in txt: return "?餅?璈?
            return "?嗥璈?
        df_machine["璈璈?"] = df_machine.apply(get_machine_type, axis=1)
        m_stats = df_machine["璈璈?"].value_counts().reset_index()
        m_stats.columns = ["璈?", "隞嗆"]
        color_map = {row["璈?"]: BRAND_PALETTE[i % len(BRAND_PALETTE)]
                     for i, row in m_stats.iterrows()}
        fig2 = px.pie(
            m_stats, names="璈?", values="隞嗆",
            title="璈??蝝啣?瘥?", hole=0.3,
            color="璈?", color_discrete_map=color_map,
        )
        fig2.update_traces(texttemplate="%{percent:.1%}", textinfo="percent+label")
        fig2.update_layout(height=400, margin=dict(t=40, b=0, l=0, r=0))
        c2.plotly_chart(fig2, use_container_width=True, key=f"{key_prefix}_fig2" if key_prefix else None)
    else:
        c2.info("?⊥??啁???)

    detail_stats = df["??蝝圈?"].value_counts().reset_index().head(10)
    detail_stats.columns = ["??蝝圈?", "隞嗆"]
    fig3 = px.bar(
        detail_stats, x="隞嗆", y="??蝝圈?",
        orientation="h", title="?之??蝝圈???",
        color_discrete_sequence=[BRAND_BLUE],
    )
    fig3.update_traces(marker_color=BRAND_BLUE)
    fig3.update_layout(
        height=400,
        yaxis={"categoryorder": "total ascending"},
        xaxis=dict(tickformat="d", nticks=6),
        xaxis_title="?賊?(隞?",
        yaxis_title="?",
        margin=dict(t=40, b=0, l=0, r=0),
    )
    c3.plotly_chart(fig3, use_container_width=True, key=f"{key_prefix}_fig3" if key_prefix else None)


def section_2():
    st.markdown('<div class="feature-title">?鈭??”?? AI ????</div>', unsafe_allow_html=True)
    if "analysis_df" not in st.session_state:
        st.info("隢??典??賭?摰?????)
        return
    df_full = st.session_state["analysis_df"]
    if df_full.empty:
        st.warning("?桀?瘝?鞈???)
        return

    # --- Date range filter ---
    date_cols = [c for c in df_full.columns if "?交?" in c or "date" in c.lower()]
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
                st.markdown("##### ???交????)
                dr_col1, dr_col2 = st.columns(2)
                start_d = dr_col1.date_input("韏瑕??交?", value=min_d, min_value=min_d, max_value=max_d)
                end_d   = dr_col2.date_input("蝯??交?", value=max_d, min_value=min_d, max_value=max_d)
                df = df[(df[dcol].dt.date >= start_d) & (df[dcol].dt.date <= end_d)]
                st.caption(f"?桀?憿舐內 {len(df)} 蝑?/ ??{len(df_full)} 蝑?)
        except Exception:
            pass

    # 蝯? source_name = ?交?????冽 PPT 撠嚗?
    if start_d and end_d:
        ppt_source = f"{start_d.strftime('%Y/%m/%d')}嚚end_d.strftime('%Y/%m/%d')}"
    else:
        ppt_source = st.session_state.get("source_name", "unknown")

    stats = df["??憿?"].value_counts().rename_axis("??憿?").reset_index(name="隞嗆")
    stats["?曉?瘥?] = (stats["隞嗆"] / max(stats["隞嗆"].sum(), 1) * 100).round(1)
    stats["甇詨惇?券?"] = stats["??憿?"].map(DEPT_MAP).fillna("")

    # Build totals row
    total_count = int(stats["隞嗆"].sum())
    dept_totals = stats.groupby("甇詨惇?券?")["隞嗆"].sum()
    dept_summary = "  ".join([f"{d}:{int(n)}隞? for d, n in dept_totals.items() if d])
    totals_row = pd.DataFrame([{
        "??憿?": "[ ?? ]",
        "隞嗆": total_count,
        "?曉?瘥?: 100.0,
        "甇詨惇?券?": dept_summary,
    }])
    stats_with_total = pd.concat([stats, totals_row], ignore_index=True)

    st.markdown("#### 憿?隞嗆?? (?舐?亦楊頛荔??”?單??郊)")
    edited_stats = st.data_editor(
        stats_with_total,
        use_container_width=True,
        hide_index=True,
        column_config={
            "甇詨惇?券?": st.column_config.SelectboxColumn(options=DEPT_OPTIONS + [dept_summary]),
            "?曉?瘥?: st.column_config.NumberColumn(format="%.1f %%")
        },
        key="stats_editor",
        num_rows="fixed",
    )
    # Use main stats (drop totals row) for charts
    chart_stats = edited_stats[edited_stats["??憿?"] != "[ ?? ]"]
    render_charts_from_stats(chart_stats, df, key_prefix="sec2")

    st.markdown("#### AI ??????")
    st.markdown("##### AI 閮剖?嚗憛恬?")
    col_ai_1, col_ai_2 = st.columns([3, 2])
    key_input = col_ai_1.text_input("OpenAI API Key嚗?征?蝙?典撱箄???閬?", type="password")
    model_name = col_ai_2.text_input("璅∪?", value="gpt-4o-mini")
    if key_input:
        st.session_state["OPENAI_API_KEY"] = key_input

    ai_text = generate_ai_summary_llm(df, model_name=model_name)
    st.text_area("?????汗", ai_text, height=140)

    # ?? ???Ｙ????頛?獢??踹? Streamlit on_click ??獢??芰????
    chart_colors = st.session_state.get("chart_colors_sec2", {})

    # ??session_state 敹怠?嚗??甈⊿?蝜芷??Ｙ?憭扳?
    cache_key = f"chart_pack_{ppt_source}"
    if cache_key not in st.session_state:
        with st.spinner("甇??Ｙ??”?陛??.."):
            try:
                st.session_state[cache_key] = build_chart_pack(
                    df,
                    color_bar=chart_colors.get("bar"),
                    color_pie=chart_colors.get("pie"),
                    color_hbar=chart_colors.get("hbar"),
                )
            except Exception as e:
                st.error(f"?”?Ｙ?憭望?嚗e}")
                st.session_state[cache_key] = {}

    chart_pack = st.session_state[cache_key]

    ppt_cache_key = f"ppt_bytes_{ppt_source}"
    if ppt_cache_key not in st.session_state:
        with st.spinner("甇??Ｙ? PPT 蝪∪..."):
            try:
                st.session_state[ppt_cache_key] = build_ppt_bytes(
                    chart_stats, ai_text, ppt_source, chart_pack=chart_pack,
                )
            except Exception as e:
                st.error(f"PPT ?Ｙ?憭望?嚗e}")
                st.session_state[ppt_cache_key] = b""

    ppt_bytes = st.session_state[ppt_cache_key]

    # ?? ?Ｙ? ZIP ??
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
            st.error(f"ZIP ?Ｙ?憭望?嚗e}")
            st.session_state[zip_cache_key] = b""

    zip_bytes = st.session_state[zip_cache_key]

    # ?? 銝???嚗?獢歇???末嚗??
    dl_col1, dl_col2, dl_col3 = st.columns(3)
    dl_col1.download_button(
        "漎? 銝? AI ????瑼?,
        data=ai_text.encode("utf-8"),
        file_name=f"{datetime.now().strftime('%Y%m%d')}_AI????.txt",
        mime="text/plain",
        use_container_width=True,
    )
    dl_col2.download_button(
        "漎? 銝??”??嚗IP嚗?,
        data=zip_bytes,
        file_name=f"{datetime.now().strftime('%Y%m%d')}_?”??.zip",
        mime="application/zip",
        use_container_width=True,
        disabled=not zip_bytes,
    )
    dl_col3.download_button(
        "漎? 銝?萎?頛??陛??PPT",
        data=ppt_bytes,
        file_name=f"{datetime.now().strftime('%Y%m%d')}_??蝪∪.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
        disabled=not ppt_bytes,
    )



def section_3():
    st.markdown('<div class="feature-title">?銝?甇瑕??蝝??/div>', unsafe_allow_html=True)

    # ?? Google Sheets ????????
    import os
    has_creds = bool(os.environ.get("GOOGLE_CREDENTIALS_JSON", ""))
    has_sid   = bool(os.environ.get("HISTORY_SHEET_ID", ""))
    ws_test   = _history_sheet()
    if ws_test is not None:
        st.success("?? Google Sheets 撌脤??嚗風?脩??偶銋?摮?)
    elif has_creds and has_sid:
        ws_test2 = _history_sheet(log_error=True)
        err_detail = st.session_state.get("_gsheet_error", "")
        st.warning(f"?? ?啣?霈撌脰身摰????憭望?\n{err_detail}")
        st.info("? 隢 Google Cloud Console 蝣箄?撌脣???**Google Sheets API** ??**Google Drive API**嚗nhttps://console.cloud.google.com/apis/library")
    else:
        st.info("?對? ?芷?? Google Sheets嚗風?脩????甈∠汗")

    history = load_history()
    if not history:
        st.info("撠甇瑕蝝??)
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

    # ?? ?交???祟?詨 ?????????????????????????????????????????
    st.markdown("---")
    st.markdown("##### ?? ?交???祟?賂??祟?豢??＊蝷箇???")
    f_col1, f_col2, f_col3 = st.columns([2, 2, 1])

    # ????????交?蝭?
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

    start_filter = f_col1.date_input("???交?", value=None, min_value=min_date, max_value=max_date,
                                      key="s3_start", format="YYYY/MM/DD")
    end_filter   = f_col2.date_input("蝯??交?", value=None, min_value=min_date, max_value=max_date,
                                      key="s3_end", format="YYYY/MM/DD")
    do_filter = f_col3.button("?? 蝭拚", key="s3_filter", use_container_width=True)

    # ?臬撌脣??祟??
    filter_active = start_filter is not None or end_filter is not None

    if not filter_active:
        st.caption("隢??????祟?詻????喳憿舐內閰脣???甇瑕蝝??)
        return

    # 靘????瞈?
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
        st.info(f"??豢????{start_filter} 嚚?{end_filter}嚗甇瑕蝝??)
        return

    st.caption(f"?望??{len(filtered)} 蝑???)
    history = filtered

    for item in history:
        out_path = Path(item.get("output_path", ""))
        cache = st.session_state.get("_history_cache", {})
        item_id = item["id"]

        # ?? excel bytes嚗?蝣???session_state 敹怠?嚗歇??load_history 敺?Sheets 憛怠嚗?
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
            continue   # ???曆??堆?頝喲?
        
        sname = item.get('source_name', '')
        if len(sname) > 28:
            sname = sname[:14] + "..." + sname[-10:]
        label = f"{item['created_at'][:16]}  {sname}  ({item['rows']} 蝑?"
        with st.expander(label):
            tab_data, tab_chart, tab_ai = st.tabs(["鞈??汗", "?”??", "AI ????"])
            
            with tab_data:
                st.dataframe(df_hist.head(30), use_container_width=True, hide_index=True)
                col1, col2, col3 = st.columns([1, 1, 1])
                col1.download_button(
                    "銝?閰脣???",
                    data=dl_bytes,
                    file_name=item["output_name"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{item['id']}",
                )
                if col2.button("[蝺刻摩]", key=f"edit_{item['id']}"):
                    st.session_state["analysis_df"] = df_hist.copy()
                    st.session_state["source_name"] = item["source_name"]
                    st.session_state["_editing_history_id"] = item["id"]
                    st.session_state["menu"] = "銝瑼??嚗???嚗?
                    st.rerun()
                if col3.button("[?芷]", key=f"del_{item['id']}"):
                    delete_history(item["id"])
                    st.rerun()
            
            with tab_chart:
                if not df_hist.empty:
                    render_charts(df_hist, key_prefix=f"hist_{item['id']}")
                    cdl1, cdl2 = st.columns(2)
                    hist_stats = df_hist["??憿?"].value_counts().rename_axis("??憿?").reset_index(name="隞嗆")
                    hist_stats["?曉?瘥?] = (hist_stats["隞嗆"] / max(hist_stats["隞嗆"].sum(), 1) * 100).round(1)
                    hist_stats["甇詨惇?券?"] = hist_stats["??憿?"].map(DEPT_MAP).fillna("")
                    hist_ai = generate_ai_summary(df_hist)
                    hist_chart_pack = build_chart_pack(df_hist)

                    hist_ppt = build_ppt_bytes(
                        hist_stats,
                        hist_ai,
                        item.get("source_name", "history"),
                        chart_pack=hist_chart_pack,
                    )
                    cdl1.download_button(
                        "銝?萎?頛PT",
                        data=hist_ppt,
                        file_name=f"{datetime.now().strftime('%Y%m%d')}_{safe_filename(item.get('source_name','history'))}_?”??.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key=f"hist_ppt_{item['id']}",
                    )
                    hist_zip = io.BytesIO()
                    with zipfile.ZipFile(hist_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                        for fn, b in hist_chart_pack.items():
                            zi = zipfile.ZipInfo(fn)
                            zi.flag_bits |= 0x800  # UTF-8 filename flag嚗?葉??蝣?
                            zi.compress_type = zipfile.ZIP_DEFLATED
                            zf.writestr(zi, b)
                    cdl2.download_button(
                        "銝???嚗IP嚗?,
                        data=hist_zip.getvalue(),
                        file_name=f"{datetime.now().strftime('%Y%m%d')}_{safe_filename(item.get('source_name','history'))}_?”.zip",
                        mime="application/zip",
                        key=f"hist_img_{item['id']}",
                    )
                else:
                    st.info("?∟??蝜芸?")
                    
            with tab_ai:
                st.info("暺?銝???單????祆?獢? AI ????")
                if st.button("[?Ｙ? AI ??]", key=f"ai_btn_{item['id']}"):
                    with st.spinner("AI ??銝?.."):
                        ai_result = generate_ai_summary_llm(df_hist)
                        st.markdown(ai_result)



def section_4():
    """???????摮?撟游漲頞典???銵冽"""

    # ?? ECOCO ?? CSS嚗?朣?HTML 蝭憸冽嚗??????????????????????
    st.markdown("""<style>
    .s4-header{background:#060E9F;color:#fff;padding:22px 26px;border-radius:12px;
               border-bottom:6px solid #FF5000;margin-bottom:18px}
    .s4-header h2{margin:0;font-size:14px;font-weight:700;letter-spacing:.3px}
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

    # ?? ?? ??????????????????????????????????????????????????????
    st.markdown("""
    <div class="s4-header">
      <h2>?? ECOCO 摰Ｚ迄頞典???銵冽</h2>
      <p>???餌?暺?券??餃?憿??璈瘥? | ?芾??交????+ 蝬剖漲蝭拚</p>
    </div>""", unsafe_allow_html=True)

    # ?? 鞈?靘? ??????????????????????????????????????????????????
    src_tab1, src_tab2 = st.tabs(["?? 甇瑕蝝????, "?? 憛怠 Google Sheets 蝬脣?"])
    all_dfs: list[pd.DataFrame] = []

    with src_tab1:
        ws = _history_sheet()
        if ws:
            import base64 as _b64
            try:
                for grow in ws.get_all_values()[1:]:
                    if not grow or not grow[0]: continue
                    data_ref = grow[4] if len(grow) > 4 else ""
                    if data_ref:
                        try:
                            if data_ref.startswith("sheet:"):
                                data_ws = ws.spreadsheet.worksheet(data_ref.split(":", 1)[1])
                                all_dfs.append(_worksheet_to_dataframe(data_ws))
                            else:
                                all_dfs.append(pd.read_excel(io.BytesIO(_b64.b64decode(data_ref))))
                        except Exception: pass
            except Exception: pass
        st.caption(f"撌脰???{len(all_dfs)} 隞賣風?脩??? if all_dfs else "撠甇瑕鞈?")

    with src_tab2:
        gs_url = st.text_input("Google Sheets 蝬脣?", placeholder="https://docs.google.com/spreadsheets/d/xxxxx/edit", key="s4v3_gsurl")
        gs_sheet = st.text_input("撌乩?銵典?蝔梧??征霈?洵銝撘蛛?", key="s4v3_gssheet", value="")
        if st.button("? 霈??, key="s4v3_load_gs"):
            if not gs_url:
                st.error("隢‵?亦雯?")
            else:
                try:
                    import re as _re
                    m = _re.search(r"/spreadsheets/d/([^/]+)", gs_url)
                    if not m:
                        st.error("?⊥?閫??閰衣?銵?ID")
                    else:
                        _client = _get_gsheet_client()
                        if not _client:
                            st.error("?芷?? Google API")
                        else:
                            _ss = _client.open_by_key(m.group(1))
                            _ws = _ss.worksheet(gs_sheet) if gs_sheet else _ss.get_worksheet(0)
                            _rows = _ws.get_all_values()
                            if _rows:
                                _df = pd.DataFrame(_rows[1:], columns=_rows[0])
                                all_dfs.append(_df)
                                st.session_state["_s4v3_gs_df"] = _df
                                st.success(f"??撌脰??_ws.title}????{len(_df)} ??)
                except Exception as e:
                    st.error(f"霈?仃??{e}")
        if st.session_state.get("_s4v3_gs_df") is not None:
            all_dfs.append(st.session_state["_s4v3_gs_df"])

    if not all_dfs:
        st.info("撠鞈?嚗???銝摰????脣?嚗?憛怠 Google Sheets 蝬脣???)
        return

    df_all = pd.concat(all_dfs, ignore_index=True).drop_duplicates()

    # ?? 甈??芸??菜葫 ??????????????????????????????????????????????
    date_col   = next((c for c in df_all.columns if "?交?" in c or "date" in c.lower()), None)
    type_col   = next((c for c in df_all.columns if "??憿?" in c), None)
    detail_col = next((c for c in df_all.columns if "??蝝圈?" in c), None)
    dept_col   = next((c for c in df_all.columns if "?券?" in c or "甇詨惇" in c), None)
    city_col   = next((c for c in df_all.columns if "蝡???? in c or "??" in c or "??? in c), None)
    station_col= next((c for c in df_all.columns if c == "蝡??迂"), None) or \
                 next((c for c in df_all.columns if "蝡??迂" in c and "蝺刻?" not in c), None)
    machine_col= next((c for c in df_all.columns if "璈憿?" in c or "璈" in c), None)

    if not date_col:
        st.warning("?曆??唳??雿?)
        return

    df_all[date_col] = pd.to_datetime(df_all[date_col], errors="coerce")
    df_all = df_all.dropna(subset=[date_col])
    if df_all.empty:
        st.warning("?⊥??????)
        return

    # ?? ??????蝬剖漲 + ?芾??交?嚗????????????????????????????
    st.markdown('<div class="s4-section">?? 蝭拚璇辣</div>', unsafe_allow_html=True)
    filter_c1, filter_c2, filter_c3 = st.columns([2, 3, 2])

    dim_mode = filter_c1.radio("??璅∪?", ["蝬剖漲?豢?", "?芾??交????], horizontal=True, key="s4v3_dimmode")

    DIM_FREQ = {"??: "W", "??: "M", "摮?: "Q", "撟游漲": "Y"}
    period_sel = period_prev = None
    df_cur = df_prev = pd.DataFrame()

    if dim_mode == "蝬剖漲?豢?":
        dim = filter_c2.selectbox("??蝬剖漲", ["??, "??, "摮?, "撟游漲"], index=1, key="s4v3_dim")
        df_all["_period"] = df_all[date_col].dt.to_period(DIM_FREQ[dim]).astype(str)
        periods = sorted(df_all["_period"].unique(), reverse=True)
        if not periods:
            st.warning("鞈?銝雲")
            return
        period_sel = filter_c3.selectbox(f"?祆?", periods, key="s4v3_period")
        p_idx = periods.index(period_sel)
        period_prev = periods[p_idx + 1] if p_idx + 1 < len(periods) else None
        df_cur  = df_all[df_all["_period"] == period_sel].copy()
        df_prev = df_all[df_all["_period"] == period_prev].copy() if period_prev else pd.DataFrame()
        period_label = period_sel
    else:
        min_d = df_all[date_col].min().date()
        max_d = df_all[date_col].max().date()
        d_col1, d_col2 = filter_c2.columns(2)
        start_d = d_col1.date_input("??", value=min_d, min_value=min_d, max_value=max_d, key="s4v3_sd")
        end_d   = d_col2.date_input("蝯?", value=max_d, min_value=min_d, max_value=max_d, key="s4v3_ed")
        df_cur  = df_all[(df_all[date_col].dt.date >= start_d) & (df_all[date_col].dt.date <= end_d)].copy()
        period_label = f"{start_d} 嚚?{end_d}"
        enable_compare = filter_c3.checkbox("?撠??, value=False, key="s4v3_compare_on")
        if enable_compare:
            cmp_c1, cmp_c2 = filter_c3.columns(2)
            cmp_start = cmp_c1.date_input("撠??", value=min_d, min_value=min_d, max_value=max_d, key="s4v3_cmp_sd")
            cmp_end = cmp_c2.date_input("撠蝯?", value=min_d, min_value=min_d, max_value=max_d, key="s4v3_cmp_ed")
            df_prev = df_all[(df_all[date_col].dt.date >= cmp_start) & (df_all[date_col].dt.date <= cmp_end)].copy()
            period_prev = f"{cmp_start} 嚚?{cmp_end}"
        else:
            period_prev = None
            df_prev = pd.DataFrame()

    # ?? 憭雁蝭拚 chips嚗?撣??券?/??憿?/璈嚗??????????????????
    st.markdown("**蝭拚蝬剖漲嚗?*")
    chip_cols = st.columns(4)

    city_filter   = chip_cols[0].multiselect("??儭???", sorted(df_cur[city_col].dropna().unique().tolist()) if city_col and city_col in df_cur.columns else [], key="s4v3_city")
    dept_filter   = chip_cols[1].multiselect("? ?券?", sorted(df_cur[dept_col].dropna().unique().tolist()) if dept_col and dept_col in df_cur.columns else [], key="s4v3_dept")
    type_filter   = chip_cols[2].multiselect("????憿?", sorted(df_cur[type_col].dropna().unique().tolist()) if type_col and type_col in df_cur.columns else [], key="s4v3_type")
    mach_filter   = chip_cols[3].multiselect("? 璈憿?", sorted(df_cur[machine_col].dropna().unique().tolist()) if machine_col and machine_col in df_cur.columns else [], key="s4v3_mach")

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

    # ?? 璈憿?嚗???飛憿??嗆??????????????????????????
    if machine_col and machine_col in df_filt.columns:
        def _normalize_machine(val):
            v = str(val).strip()
            if "?寡?" in v or "?嗥" in v: return "?嗥璈?
            if "?餅?" in v: return "?餅?璈?
            return v
        df_filt = df_filt.copy()
        df_filt[machine_col] = df_filt[machine_col].apply(_normalize_machine)
    if machine_col and machine_col in df_all.columns:
        df_all = df_all.copy()
        df_all[machine_col] = df_all[machine_col].apply(
            lambda v: "?嗥璈? if ("?寡?" in str(v) or "?嗥" in str(v)) else ("?餅?璈? if "?餅?" in str(v) else str(v))
        )

    # ?? KPI ?∠?嚗 st.metric ?踹? HTML escape ??嚗????????????????
    st.markdown(f'<div class="s4-section">?? ?祆??單?蝯梯?嚗period_label}嚗?/div>', unsafe_allow_html=True)
    st.caption(f"?? 鞈????{period_label}?蝭拚敺 **{n_cur}** 蝑?)

    kpi_items = [("??儭?蝮賡脖辣??, n_cur, n_prev)]
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
            delta_str = f"{sym}{p:.1f}% vs 銝?"
        kpi_cols[col_i].metric(label=lbl, value=cur, delta=delta_str)

    # ?? ??蝯梯?嚗???蝡?/??蝝圈?嚗???????????????????????????????
    st.markdown(f'<div class="s4-section">?? 獢辣??蝯梯? Top 5 ?? {period_label}</div>', unsafe_allow_html=True)

    rank_cols = st.columns(3)
    MEDAL = ["??","??","??","4儭","5儭"]

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
        st.markdown('<div style="font-weight:700;color:#060E9F;margin-bottom:8px">?? ???銵?/div>', unsafe_allow_html=True)
        if city_col and city_col in df_filt.columns and not df_filt[city_col].dropna().empty:
            st.markdown('<div class="s4-card">' + rank_table_html(df_filt[city_col].value_counts(), "??/???, "隞嗆") + '</div>', unsafe_allow_html=True)
        else:
            st.info("?∪?撣???)

    with rank_cols[1]:
        st.markdown('<div style="font-weight:700;color:#060E9F;margin-bottom:8px">? 蝡???</div>', unsafe_allow_html=True)
        if station_col and station_col in df_filt.columns and not df_filt[station_col].dropna().empty:
            st.markdown('<div class="s4-card">' + rank_table_html(df_filt[station_col].value_counts(), "蝡??迂", "隞嗆") + '</div>', unsafe_allow_html=True)
        else:
            st.info("?∠?暺???)

    with rank_cols[2]:
        st.markdown('<div style="font-weight:700;color:#060E9F;margin-bottom:8px">?? ??蝝圈???</div>', unsafe_allow_html=True)
        if detail_col and detail_col in df_filt.columns and not df_filt[detail_col].dropna().empty:
            st.markdown('<div class="s4-card">' + rank_table_html(df_filt[detail_col].value_counts(), "??蝝圈?", "隞嗆") + '</div>', unsafe_allow_html=True)
        else:
            st.info("?∠敦????)

    # ?? ?”嚗?憿???+ 璈雿?嚗?朣?HTML 蝭嚗?????????????????
    st.markdown(f'<div class="s4-section">?? ?豢??航??????? {period_label}</div>', unsafe_allow_html=True)
    chart_col1, chart_col2 = st.columns(2)

    with chart_col1:
        if type_col and type_col in df_filt.columns:
            _tc = df_filt[type_col].value_counts()
            _total = _tc.sum()
            fig_pie = px.pie(
                values=_tc.values, names=_tc.index,
                title=f"{period_label} 摰Ｚ迄憿??",
                hole=0.35,
                color_discrete_sequence=["#060E9F","#FF5000","#FFCE00","#8EB9C9","#0076A9","#FAE0B8"],
            )
            # 撠5%??敶Ｗ憿舐內?典?靘??踹?璅惜??
            fig_pie.update_traces(
                texttemplate="%{label}<br>%{percent:.1%}",
                textposition="auto",
                textfont_size=12,
            )
            fig_pie.update_layout(
                height=420,
                showlegend=True,
                legend=dict(
                    orientation="v",
                    yanchor="middle", y=0.5,
                    xanchor="left", x=1.0,
                    font=dict(size=12),
                    itemsizing="constant",
                ),
                margin=dict(t=55, b=20, l=20, r=160),
                title_font_size=15,
                title_x=0.0,
            )
            st.plotly_chart(fig_pie, use_container_width=True)

    with chart_col2:
        if machine_col and machine_col in df_filt.columns and not df_filt[machine_col].dropna().empty:
            _mc = df_filt[machine_col].value_counts()
            fig_mac = px.pie(
                values=_mc.values, names=_mc.index,
                title=f"{period_label} 璈摰Ｚ迄雿?",
                color_discrete_sequence=["#FF5000","#060E9F","#8EB9C9","#FFCE00"],
            )
            fig_mac.update_traces(
                texttemplate="%{label}<br>%{percent:.1%}",
                textposition="auto",
                textfont_size=13,
            )
            fig_mac.update_layout(
                height=420,
                showlegend=True,
                legend=dict(
                    orientation="v",
                    yanchor="middle", y=0.5,
                    xanchor="left", x=1.0,
                    font=dict(size=12),
                ),
                margin=dict(t=55, b=20, l=20, r=160),
                title_font_size=15,
            )
            st.plotly_chart(fig_mac, use_container_width=True)
        elif detail_col and detail_col in df_filt.columns:
            _dc = df_filt[detail_col].value_counts().head(8)
            fig_det = px.bar(
                x=list(_dc.values)[::-1], y=list(_dc.index)[::-1],
                orientation="h", title=f"{period_label} TOP 8 ??蝝圈?",
                color_discrete_sequence=["#060E9F"],
            )
            fig_det.update_layout(height=420, xaxis=dict(tickformat="d", nticks=6),
                                   margin=dict(t=45,b=0,l=0,r=0))
            st.plotly_chart(fig_det, use_container_width=True)

    # ?? 頞典????????????????????????????????????????????????????
    st.markdown(f'<div class="s4-section">?? 摰Ｚ迄頞典?? ?? {period_label}</div>', unsafe_allow_html=True)
    if dim_mode == "蝬剖漲?豢?" and len(df_all["_period"].unique()) >= 2:
        _trend = df_all.groupby("_period").size().reset_index(name="隞嗆").sort_values("_period")
        fig_line = px.line(
            _trend, x="_period", y="隞嗆",
            title=f"甇瑕隞嗆頞典",
            markers=True,
            color_discrete_sequence=["#FF5000"],
        )
        fig_line.update_traces(fill="tozeroy", fillcolor="rgba(255,80,0,0.1)")
        if period_sel and period_sel in _trend["_period"].values:
            _sel_i = _trend.index[_trend["_period"] == period_sel].tolist()
            if _sel_i:
                fig_line.add_vline(x=_sel_i[0], line_dash="dash", line_color="#060E9F",
                                   annotation_text="?祆?", annotation_font_color="#060E9F")
        fig_line.update_layout(
            height=320, xaxis_title="??",
            yaxis=dict(tickformat="d", nticks=6),
            paper_bgcolor="white", plot_bgcolor="rgba(250,224,184,0.15)",
            margin=dict(t=45,b=0),
        )
        st.plotly_chart(fig_line, use_container_width=True)
    else:
        # ?芾??交?嚗??乩辣??
        _daily = df_filt.groupby(df_filt[date_col].dt.date).size().reset_index(name="隞嗆")
        _daily.columns = ["?交?", "隞嗆"]
        if not _daily.empty:
            fig_daily = px.bar(
                _daily, x="?交?", y="隞嗆",
                title="???扳??乩辣??,
                color_discrete_sequence=["#060E9F"],
            )
            fig_daily.update_layout(height=300, yaxis=dict(tickformat="d", nticks=6), margin=dict(t=45,b=0))
            st.plotly_chart(fig_daily, use_container_width=True)

    # ?? ??撅???嚗??嚗????????????????????????????????????
    if city_col and city_col in df_filt.columns and not df_filt.empty:
        st.markdown(f'<div class="s4-section">??儭????銵? ?? {period_label}</div>', unsafe_allow_html=True)
        city_rank = df_filt[city_col].value_counts()
        MEDAL_LIST = ["??","??","??","4儭","5儭","6儭","7儭","8儭","9儭","??"]
        for ri, (city, cnt) in enumerate(city_rank.items()):
            prev_cnt = int(df_prev[city_col].eq(city).sum()) if not df_prev.empty and city_col in df_prev.columns else 0
            p = pct_change(int(cnt), prev_cnt)
            delta_s = (f"?{'?? if p>0 else '??}{abs(p):.1f}%") if p is not None else ""
            medal = MEDAL_LIST[ri] if ri < len(MEDAL_LIST) else f"#{ri+1}"
            with st.expander(f"{medal} **{city}**?{int(cnt)} 隞閔delta_s}", expanded=(ri==0)):
                df_city = df_filt[df_filt[city_col] == city]
                ec1, ec2 = st.columns(2)
                with ec1:
                    st.markdown("**?? 蝡???**")
                    if station_col and station_col in df_city.columns:
                        _sr = df_city[station_col].value_counts().head(8)
                        _sr_prev = df_prev[df_prev[city_col]==city][station_col].value_counts() if not df_prev.empty and city_col in df_prev.columns and station_col in df_prev.columns else pd.Series(dtype=int)
                        _html = ""
                        for si, (sn, sv) in enumerate(_sr.items()):
                            _sp = int(_sr_prev.get(sn, 0))
                            _sd = (f"?{'?? if int(sv)-_sp>0 else '??}{abs(pct_change(int(sv),_sp)):.0f}%") if _sp and pct_change(int(sv),_sp) is not None else ""
                            _sm = MEDAL_LIST[si] if si < len(MEDAL_LIST) else f"#{si+1}"
                            _html += f'<div style="padding:5px 0;border-bottom:.5px solid #eee;font-size:13px;display:flex;justify-content:space-between"><span>{_sm} {str(sn)[:18]}</span><b style="color:#FF5000">{int(sv)}{_sd}</b></div>'
                        st.markdown(_html, unsafe_allow_html=True)
                with ec2:
                    st.markdown("**?? ??蝝圈???**")
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

    # ?? ?券??? ?????????????????????????????????????????????????
    if dept_col and dept_col in df_filt.columns:
        st.markdown(f'<div class="s4-section">? ??隞嗆?? ?? {period_label}</div>', unsafe_allow_html=True)
        dept_rank = df_filt[dept_col].replace("","?芸???).value_counts()
        DEPT_COLOR = {"????:"#FF5000","銵??:"#FFCE00","鞈???:"#060E9F"}
        fig_dept = px.bar(
            dept_rank.reset_index(), x=dept_col, y="count",
            title="??隞嗆",
            color=dept_col,
            color_discrete_map=DEPT_COLOR,
        )
        fig_dept.update_layout(height=300, yaxis=dict(tickformat="d", nticks=6),
                                showlegend=False, margin=dict(t=45,b=0))
        st.plotly_chart(fig_dept, use_container_width=True)

    # ?? 摰? PDF 銝? ????????????????????????????????????????
    st.markdown("---")
    st.markdown(f'<div class="s4-section">漎? 銝?摰???勗?</div>', unsafe_allow_html=True)

    if st.button("?? ?Ｙ?摰?? PDF", key="s4_full_pdf", use_container_width=False):
        with st.spinner("甇??Ｙ?憭? PDF ?勗?..."):
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
                _font_path = _ensure_cjk_font()  # 雿輻撌脣翰??摮?頝臬?
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
                        # ???璇?
                        self.set_fill_color(6, 14, 159)
                        self.rect(0, 0, 210, 14, style="F")
                        self.set_font(self.fn, size=9)
                        self.set_text_color(255, 255, 255)
                        self.set_xy(5, 3)
                        self.cell(0, 8, self._s(f"ECOCO 摰Ｚ迄頞典???勗??{period_label}"))
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
                        self.cell(0, 8, self._s(f"蝚?{self.page_no()} ??Ｗ?交?嚗datetime.now().strftime('%Y/%m/%d')}"), align="C")

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
                        """?典?銵冽嚗?游?銵??銵?""
                        # 銵券
                        self.set_fill_color(6, 14, 159)
                        self.set_text_color(255, 255, 255)
                        self.set_font(self.fn, size=9)
                        x0 = self.get_x()
                        for h, w in zip(headers, col_widths):
                            self.cell(w, 8, self._s(str(h)), border=1, fill=True, align="C",
                                      new_x=XPos.RIGHT, new_y=YPos.TOP)
                        self.ln(8)
                        # 鞈???
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
                        """撋 matplotlib ?”"""
                        _b = io.BytesIO()
                        fig.savefig(_b, format="png", dpi=180, bbox_inches="tight",
                                    facecolor="white")
                        _mplt.close(fig)
                        _b.seek(0)
                        x = (210 - w) / 2  # 蝵桐葉
                        self.image(_b, x=x, y=self.get_y(), w=w, h=h)
                        self.set_y(self.get_y() + h + 4)

                F = "CJK" if _font_path else "Helvetica"

                pdf = EcocoPDF(_font_path, F)

                # ????????????????????????????????????????????
                # Page 1嚗???+ KPI ??
                # ????????????????????????????????????????????
                pdf.add_page()
                # 憭扳?憿?
                pdf.set_fill_color(6, 14, 159)
                pdf.rect(15, 20, 180, 38, style="F")
                pdf.set_font(F, size=20)
                pdf.set_text_color(255, 255, 255)
                pdf.set_xy(15, 26)
                pdf.cell(180, 12, pdf._s("ECOCO 摰Ｚ迄頞典???勗?"), align="C",
                         new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                pdf.set_font(F, size=11)
                pdf.set_xy(15, 42)
                pdf.cell(180, 8, pdf._s(f"鞈????{period_label}"), align="C",
                         new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                pdf.set_draw_color(255, 80, 0)
                pdf.set_line_width(2)
                pdf.line(15, 58, 195, 58)
                pdf.set_line_width(0.2)
                pdf.set_y(68)

                # KPI ?∠?嚗帖??
                pdf.section_title("?祆??單?蝯梯???")
                kpi_data = [("??儭?蝮賡脖辣??, n_cur)]
                if type_col and type_col in df_filt.columns:
                    for t, tc in df_filt[type_col].value_counts().head(3).items():
                        kpi_data.append((str(t), int(tc)))
                card_w = 170 // len(kpi_data)
                card_x = 20
                for lbl, val in kpi_data:
                    pdf.set_fill_color(255, 206, 0)  # 暺??
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

                # ?? KPI
                if city_col and city_col in df_filt.columns:
                    pdf.ln(4)
                    pdf.set_font(F, size=9)
                    pdf.set_text_color(60, 60, 60)
                    city_top = df_filt[city_col].value_counts().head(3)
                    line = "?|?".join([f"{c}嚗int(v)} 隞? for c, v in city_top.items()])
                    pdf.set_x(20)
                    pdf.cell(0, 7, pdf._s(f"??憭批?撣?{line}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)

                # ????????????????????????????????????????????
                # Page 2嚗?銵絞閮?
                # ????????????????????????????????????????????
                pdf.add_page()
                pdf.section_title("獢辣??蝯梯?")

                # ???銵?
                if city_col and city_col in df_filt.columns:
                    pdf.set_font(F, size=10); pdf.set_text_color(6,14,159)
                    pdf.cell(0, 7, pdf._s("?? ??/???銵?), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    _city_v = df_filt[city_col].value_counts()
                    rows_c = [[i+1, c, int(v), f"{int(v)/n_cur*100:.0f}%"]
                               for i, (c, v) in enumerate(_city_v.head(10).items())]
                    pdf.full_table(["??","??/???,"隞嗆","雿?"], rows_c, [15,110,25,30])

                # 蝡???
                if station_col and station_col in df_filt.columns:
                    pdf.set_font(F, size=10); pdf.set_text_color(6,14,159)
                    pdf.cell(0, 7, pdf._s("? 蝡???"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    _sta_v = df_filt[station_col].value_counts()
                    rows_s = [[i+1, str(s)[:30], int(v)]
                               for i, (s, v) in enumerate(_sta_v.head(10).items())]
                    pdf.full_table(["??","蝡??迂","隞嗆"], rows_s, [15,140,25])

                # ????????????????????????????????????????????
                # Page 3嚗?憿敦??銵?
                # ????????????????????????????????????????????
                pdf.add_page()
                pdf.section_title("??蝝圈???")

                if detail_col and detail_col in df_filt.columns:
                    _det_v = df_filt[detail_col].value_counts()
                    rows_d = [[i+1, str(d)[:35], int(v), f"{int(v)/n_cur*100:.0f}%"]
                               for i, (d, v) in enumerate(_det_v.head(15).items())]
                    pdf.full_table(["??","??蝝圈?","隞嗆","雿?"], rows_d, [15,120,20,25])

                # ????????????????????????????????????????????
                # Page 4嚗?銵剁??? + 璈嚗?
                # ????????????????????????????????????????????
                pdf.add_page()
                pdf.section_title("?豢??航?????)

                if type_col and type_col in df_filt.columns:
                    _tc4 = df_filt[type_col].value_counts()
                    _total4 = _tc4.sum()
                    # 憭批?嚗igsize ?游祝嚗?璅惜?寧???踹???
                    _f4, _a4 = _mplt.subplots(figsize=(9, 6))
                    _clrs4 = ["#060E9F","#FF5000","#FFCE00","#8EB9C9","#0076A9","#FAE0B8"]
                    _labels4 = [f"{k}嚗int(v)}隞塚?" for k, v in _tc4.items()]
                    wedges, texts, autotexts = _a4.pie(
                        list(_tc4.values),
                        labels=None,           # 銝?耦銝＊蝷箸?蝐歹??寧??
                        autopct=lambda p: f"{p:.0f}%" if p >= 5 else "",  # 撠?敶Ｖ?憿舐內%
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
                    _a4.set_title(f"{period_label}?摰Ｚ迄憿??", fontsize=13, pad=12)
                    _f4.tight_layout()
                    pdf.embed_image(_f4, w=175, h=120)

                if machine_col and machine_col in df_filt.columns and not df_filt[machine_col].dropna().empty:
                    _mc4 = df_filt[machine_col].value_counts()
                    _labels_mc = [f"{k}嚗int(v)}隞塚?" for k, v in _mc4.items()]
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
                    _a5.set_title(f"{period_label}?璈摰Ｚ迄雿?", fontsize=13, pad=12)
                    _f5.tight_layout()
                    pdf.embed_image(_f5, w=175, h=110)

                # ????????????????????????????????????????????
                # Page 5嚗隅??+ ?券???
                # ????????????????????????????????????????????
                pdf.add_page()
                pdf.section_title("摰Ｚ迄頞典??")

                _daily3 = df_filt.groupby(df_filt[date_col].dt.date).size().reset_index(name="隞嗆")
                if len(_daily3) > 1:
                    _f6, _a6 = _mplt.subplots(figsize=(10, 4))
                    _a6.bar([str(d) for d in _daily3.iloc[:,0]], list(_daily3["隞嗆"]),
                            color="#060E9F", edgecolor="white", linewidth=0.5)
                    _a6.set_title(f"{period_label}?瘥隞嗆頞典", fontsize=13)
                    _a6.tick_params(axis="x", rotation=30, labelsize=8)
                    _a6.yaxis.set_major_locator(_MNL(integer=True))
                    _a6.set_ylabel("隞嗆", fontsize=10)
                    _a6.grid(axis="y", alpha=0.3)
                    _f6.tight_layout()
                    pdf.embed_image(_f6, w=180, h=100)

                if dept_col and dept_col in df_filt.columns:
                    pdf.section_title("??隞嗆??")
                    _dp = df_filt[dept_col].replace("","?芸???).value_counts()
                    rows_dp = [[i+1, str(d), int(v), f"{int(v)/n_cur*100:.0f}%"]
                                for i, (d, v) in enumerate(_dp.items())]
                    pdf.full_table(["??","?券?","隞嗆","雿?"], rows_dp, [15,80,20,20])

                _pdf_bytes = bytes(pdf.output())
                st.session_state["_s4_pdf_bytes"] = _pdf_bytes
                st.session_state["_s4_pdf_label"] = period_label
                st.success(f"??PDF 撌脩????{pdf.page_no()} ??{len(_pdf_bytes)//1024} KB嚗?)
            except Exception as _e:
                import traceback
                st.error(f"PDF ?Ｙ?憭望?嚗_e}")
                st.code(traceback.format_exc())

    if st.session_state.get("_s4_pdf_bytes"):
        _label = st.session_state.get("_s4_pdf_label", period_label)
        st.download_button(
            "漎? 銝?摰?? PDF嚗???",
            data=st.session_state["_s4_pdf_bytes"],
            file_name=f"ECOCO_摰Ｚ迄??_{_label.replace(' ','').replace('嚚?,'-').replace('/','-')}.pdf",
            mime="application/pdf",
            use_container_width=False,
            key="s4_dl_full_pdf",
        )

    # ?? AI ??牧?勗? ????????????????????????????????????????????????
    st.markdown("---")
    st.markdown(f'<div class="s4-section">??儭?AI ??牧?勗??Ｙ???/div>', unsafe_allow_html=True)
    rep_type = st.radio("?勗?憿?", ["?望??勗?","???勗?","摮?","撟游漲?勗?"], horizontal=True, key="s4v3_rep")

    if st.button("?? ?Ｙ? AI ??牧?勗?", type="primary", key="s4v3_gen"):
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
                pline = f"嚗?銝? {d:+d} 隞塚?{pct_change(int(cnt),prev_cnt):+.1f}%嚗? if prev_cnt else ""
                type_summary += f"- {cat}嚗int(cnt)} 隞閔pline}\n"

        city_summary = ""
        if city_col and city_col in df_filt.columns:
            _cc = df_filt[city_col].value_counts()
            _pc = df_prev[city_col].value_counts() if not df_prev.empty and city_col in df_prev.columns else pd.Series(dtype=int)
            for city, cnt in _cc.head(5).items():
                d = int(cnt) - int(_pc.get(city,0))
                city_summary += f"- {city}嚗int(cnt)} 隞塚?{d:+d}嚗n"

        top3 = ""
        if detail_col and detail_col in df_filt.columns:
            for _, r in df_filt[detail_col].value_counts().head(3).reset_index().iterrows():
                top3 += f"- {r[detail_col]}嚗r['count']} 隞跚n"

        _upper_cmp = (
            f"\n????瘥?{period_prev}嚗total_prev} 隞塚?蝮賭辣??{pct_chg:+.1f}%嚗?
            if pct_chg is not None else ""
        )
        prompt = (
            f"雿 ECOCO 摰?臬儐?啁?瞈恥???蝝????～n"
            f"隢?誑銝???Ｗ銝隞緹rep_type}?隤芸???拙??冽?霅唬葉撠摰陛?晞n\n"
            f"??瘞??撠平?????啜葆?遣霅唳改?憒?游隤?n"
            f"??瑽?\n"
            f"1. ??踝?暺?祆???嚗n"
            f"2. 蝮賡?頞典璁膩嚗摮?蝢抬??敹菜摮?\n"
            f"3. ??憭抒?暺楛摨西圾?????蔣?選?\n"
            f"4. ??/????漁暺n"
            f"5. ?孵???餈質馱\n"
            f"6. 銝?畾菔??遣霅豹n\n"
            f"????{period_label}嚗 {total_cur} 隞塚?嚗n"
            f"{type_summary or '嚗??憿?鞈?嚗?}\n\n"
            f"??撣?撣?TOP5??\n"
            f"{city_summary or '嚗??鞈?嚗?}\n\n"
            f"??銝之??蝝圈???\n"
            f"{top3 or '嚗蝝圈?鞈?嚗?}\n"
            f"{_upper_cmp}\n\n"
            f"隢誑蝜?銝剜??啣神嚗隤?嗡?銝仃撠平嚗?畾菔 2-4 ?乓?
        )

        with st.spinner("AI 甇??啣神??牧?勗?..."):
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
                    report_text = f"?隤芸?n\n{report_text}"
                except Exception:
                    report_text = f"?? AI ?急??⊥?雿輻嚗e}嚗n\n?豢???嚗n\n{prompt}"

        st.text_area("?? ??牧?勗?嚗銴ˊ嚗?, report_text, height=460, key="s4v3_report_out")

        dl_c1, dl_c2 = st.columns(2)
        dl_c1.download_button("漎? 銝???牧?勗?嚗XT嚗?,
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
                _a.set_title(f"{period_label} 摰Ｚ迄憿??")
                _b = io.BytesIO(); _f.savefig(_b, format="png", dpi=150, bbox_inches="tight")
                _mplt.close(_f); _charts["摰Ｚ迄憿??.png"] = _b.getvalue()
            if city_col and city_col in df_filt.columns:
                _cc2 = df_filt[city_col].value_counts().head(10)
                _f2, _a2 = _mplt.subplots(figsize=(8,4))
                _a2.bar(list(_cc2.index), list(_cc2.values), color="#FF5000")
                _a2.set_title(f"{period_label} ??隞嗆??")
                _a2.yaxis.set_major_locator(_mplt.MaxNLocator(integer=True))
                _a2.tick_params(axis="x", rotation=15)
                _b2 = io.BytesIO(); _f2.savefig(_b2, format="png", dpi=150, bbox_inches="tight")
                _mplt.close(_f2); _charts["????.png"] = _b2.getvalue()
            if detail_col and detail_col in df_filt.columns:
                _dc2 = df_filt[detail_col].value_counts().head(8)
                _f3, _a3 = _mplt.subplots(figsize=(8,4))
                _a3.barh(list(_dc2.index)[::-1], list(_dc2.values)[::-1], color="#060E9F")
                _a3.set_title(f"{period_label} TOP 8 ??蝝圈?")
                _b3 = io.BytesIO(); _f3.savefig(_b3, format="png", dpi=150, bbox_inches="tight")
                _mplt.close(_f3); _charts["??蝝圈???.png"] = _b3.getvalue()
            _zbuf = io.BytesIO()
            with zipfile.ZipFile(_zbuf, "w", zipfile.ZIP_DEFLATED) as _zf:
                for _fn, _fb in _charts.items():
                    _zi = zipfile.ZipInfo(_fn); _zi.flag_bits |= 0x800
                    _zi.compress_type = zipfile.ZIP_DEFLATED; _zf.writestr(_zi, _fb)
                _zr = zipfile.ZipInfo(f"{period_label}_{rep_type}.txt")
                _zr.flag_bits |= 0x800; _zf.writestr(_zr, report_text.encode("utf-8"))
            dl_c2.download_button("漎? 銝??”+?勗?嚗IP嚗?,
                                  data=_zbuf.getvalue(),
                                  file_name=f"{period_label}_頞典??.zip",
                                  mime="application/zip", key="s4v3_dl_zip",
                                  use_container_width=True)
        except Exception as _ze:
            dl_c2.warning(f"ZIP ?Ｙ?憭望?嚗_ze}")



def main():
    apply_brand_theme()
    st.markdown("<div class='ecoco-banner'>ECOCO 客訴智能分析平台</div>", unsafe_allow_html=True)
    st.session_state.setdefault("menu", "首頁")

    nav_labels = ["首頁", "功能一", "功能二", "功能三", "功能四"]
    nav_map = {
        "首頁": "首頁",
        "功能一": "上傳檔案區（分析區）",
        "功能二": "圖表與 AI 分析",
        "功能三": "歷史分析紀錄",
        "功能四": "趨勢分析",
    }
    nav_cols = st.columns(len(nav_labels))
    for idx, label in enumerate(nav_labels):
        with nav_cols[idx]:
            active = st.session_state["menu"] == nav_map[label]
            if st.button(label, key=f"nav_{idx}", use_container_width=True, type="primary" if active else "secondary"):
                st.session_state["menu"] = nav_map[label]

    menu = st.session_state["menu"]
    if menu == "首頁":
        st.markdown("<div class='ecoco-card'><b>功能 1</b>：檔案上傳與分析。</div>", unsafe_allow_html=True)
        st.markdown("<div class='ecoco-card'><b>功能 2</b>：圖表化與 AI 重點分析。</div>", unsafe_allow_html=True)
        st.markdown("<div class='ecoco-card'><b>功能 3</b>：歷史分析紀錄。</div>", unsafe_allow_html=True)
        st.markdown("<div class='ecoco-card'><b>功能 4</b>：ECOCO 客訴趨勢分析儀表板。</div>", unsafe_allow_html=True)
    elif menu == "上傳檔案區（分析區）":
        section_1()
    elif menu == "圖表與 AI 分析":
        section_2()
    elif menu == "趨勢分析":
        section_4()
    else:
        section_3()

    st.markdown(
        """
        <style>
            .fixed-footer { position: sticky; bottom: 0; width: 100%; text-align: center; color: #d9d9d9; font-size: 12px; margin: 36px 0 12px; padding: 12px 0; pointer-events: none; z-index: 0; }
            .scroll-top-btn { position: fixed; right: 24px; bottom: 24px; z-index: 1000; border: 1px solid #8EB9C9; background: #FFFFFF; color: #060E9F; border-radius: 999px; padding: 18px 22px; font-size: 22px; cursor: pointer; box-shadow: 0 2px 8px rgba(0,0,0,.12); }
        </style>
        <div class="fixed-footer">202603© ECOCO宜可可循環經濟 客服課 ※ 請尊重智慧財產權 ※</div>
        <button class="scroll-top-btn" onclick="window.parent.scrollTo({top:0, behavior:\"smooth\"});">置頂</button>
        """,
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
