import io
import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import plotly.express as px
import streamlit as st
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
    ],
    "APP帳號設定問題類型": [
        "忘記密碼/無法重設密碼",
        "帳號資訊修改/設定",
        "無法接收簡訊驗證碼",
        "APP無法登入",
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
    "APP帳號設定問題類型": "開發部",
    "APP使用問題類型": "開發部",
    "回收點數問題類型": "企劃部",
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
            color:white; font-weight:700; margin-bottom: 12px;
          }
          .ecoco-card{
            border:1px solid #e7e7e7; border-left:6px solid var(--ecoco-orange);
            border-radius:12px; padding:10px 14px; background:white; margin-bottom:10px;
            color: #555555 !important;
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
            font-weight: 800; font-size: 1.05rem; margin-bottom: 8px;
          }
          .side-sub {
            color: #ffffff !important;
            font-size: 0.78rem; opacity: 0.85; margin-bottom: 14px;
          }
          
          /* Sidebar Buttons Restyling */
          section[data-testid="stSidebar"] .stButton > button {
            background-color: var(--ecoco-lightblue) !important;
            border-color: var(--ecoco-lightblue) !important;
            color: #333333 !important;
            border-radius: 12px;
            min-height: 46px;
            font-weight: 700;
            text-align: left;
            transition: none !important;
          }
          
          section[data-testid="stSidebar"] .stButton > button * {
            color: #333333 !important;
          }
          
          /* Active Menu Button / Clicked */
          section[data-testid="stSidebar"] .stButton > button[kind="primary"],
          section[data-testid="stSidebar"] .stButton > button[data-testid="baseButton-primary"],
          section[data-testid="stSidebar"] .stButton > button:active,
          section[data-testid="stSidebar"] .stButton > button:focus {
            background-color: #FFFFFF !important;
            border-color: #FFFFFF !important;
            color: #333333 !important;
          }

          /* Hover state unchanged (Secondary) */
          section[data-testid="stSidebar"] .stButton > button:not([kind="primary"]):hover {
            background-color: var(--ecoco-lightblue) !important;
            border-color: var(--ecoco-lightblue) !important;
            color: #333333 !important;
          }
          
          /* Hover state unchanged (Primary) */
          section[data-testid="stSidebar"] .stButton > button[kind="primary"]:hover {
            background-color: #FFFFFF !important;
            border-color: #FFFFFF !important;
            color: #333333 !important;
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


def load_input_file(uploaded_file) -> pd.DataFrame:
    suffix = Path(uploaded_file.name).suffix.lower()
    if suffix in [".xlsx", ".xls"]:
        return pd.read_excel(uploaded_file)
    if suffix == ".csv":
        return pd.read_csv(uploaded_file)
    if suffix == ".pdf":
        return parse_pdf_to_df(uploaded_file)
    raise ValueError("僅支援 excel / csv / pdf")


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


def analyze_dataframe(df: pd.DataFrame, cfg: AnalysisConfig) -> pd.DataFrame:
    out = make_unique_columns(df.copy())
    # If source already has these columns, overwrite with fresh analysis values.
    for c in ["問題類型", "問題細項", "選取", "部門", "日期"]:
        if c in out.columns:
            out = out.drop(columns=[c])
    preds = out.apply(
        lambda r: analyze_complaint(str(r.get(cfg.subject_col, "")), str(r.get(cfg.content_col, ""))),
        axis=1,
        result_type="expand",
    )
    preds.columns = ["問題類型", "問題細項"]
    out = pd.concat([out, preds], axis=1)
    # Ensure detail always belongs to topic
    out["問題細項"] = out.apply(
        lambda r: r["問題細項"] if r["問題細項"] in TOPIC_DETAIL_MAP.get(r["問題類型"], []) else TOPIC_DETAIL_MAP.get(r["問題類型"], ["其他建議"])[0],
        axis=1,
    )
    out["選取"] = False
    out["部門"] = out["問題類型"].map(DEPT_MAP).fillna("未分配")
    if cfg.date_col and cfg.date_col in out.columns:
        out["日期"] = pd.to_datetime(out[cfg.date_col], errors="coerce")
    return out


def save_history(df: pd.DataFrame, source_name: str) -> tuple[Path, str]:
    today = datetime.now().strftime("%Y%m%d")
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_name = f"{today}_分析.xlsx"
    output_path = HISTORY_DIR / f"{ts}_{output_name}"
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
    history.insert(0, meta)
    META_FILE.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")
    return output_path, output_name


def load_history() -> list[dict]:
    if not META_FILE.exists():
        return []
    return json.loads(META_FILE.read_text(encoding="utf-8"))


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
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_pdf import PdfPages
    import io
    plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'PingFang HK', 'SimHei', 'Arial', 'sans-serif']
    fig, ax = plt.subplots(figsize=(10, min(10, max(2, int(len(df)*0.5)))))
    ax.axis('tight')
    ax.axis('off')
    table_df = df.head(40).copy()
    for c in table_df.columns:
        table_df[c] = table_df[c].astype(str).apply(lambda x: x[:15] + ".." if len(x)>15 else x)
    table = ax.table(cellText=table_df.values, colLabels=table_df.columns, loc='center', cellLoc='left')
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    buf = io.BytesIO()
    with PdfPages(buf) as pdf:
        pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)
    return buf.getvalue()


def build_ppt_bytes(stats: pd.DataFrame, ai_text: str, source_name: str) -> bytes:
    prs = Presentation()
    slide1 = prs.slides.add_slide(prs.slide_layouts[5])
    title = slide1.shapes.title
    title.text = "ECOCO 客訴分析簡報"
    tx = slide1.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(8.8), Inches(3.8)).text_frame
    tx.clear()
    p = tx.paragraphs[0]
    p.text = f"來源檔案：{source_name}"
    p.font.size = Pt(18)
    p = tx.add_paragraph()
    p.text = f"產出日期：{datetime.now().strftime('%Y-%m-%d')}"
    p.font.size = Pt(16)
    p = tx.add_paragraph()
    p.text = "重點摘要："
    p.font.bold = True
    p.font.size = Pt(16)
    for ln in ai_text.splitlines():
        if ln.strip():
            q = tx.add_paragraph()
            q.text = f"- {ln.strip()}"
            q.level = 1
            q.font.size = Pt(14)

    slide2 = prs.slides.add_slide(prs.slide_layouts[5])
    slide2.shapes.title.text = "問題類型件數與占比"
    rows = min(len(stats) + 1, 12)
    cols = 4
    table = slide2.shapes.add_table(rows, cols, Inches(0.6), Inches(1.3), Inches(11.8), Inches(5.0)).table
    table.cell(0, 0).text = "問題類型"
    table.cell(0, 1).text = "件數"
    table.cell(0, 2).text = "百分比"
    table.cell(0, 3).text = "歸屬部門"
    for i, (_, r) in enumerate(stats.head(rows - 1).iterrows(), start=1):
        table.cell(i, 0).text = str(r["問題類型"])
        table.cell(i, 1).text = str(r["件數"])
        table.cell(i, 2).text = f'{r["百分比"]}%'
        table.cell(i, 3).text = str(r["歸屬部門"])

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

    uploaded = st.file_uploader("上傳檔案", type=["xlsx", "xls", "csv", "pdf"], key="uploader")
    if not uploaded:
        return

    df_raw = make_unique_columns(load_input_file(uploaded))
    st.caption(f"已載入 `{uploaded.name}`，資料筆數：{len(df_raw)}")

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
    cfg = AnalysisConfig(subject_col=subject_col, content_col=content_col, date_col=None if date_col == "(無)" else date_col)

    if st.button("開始分析", type="primary"):
        work = df_raw.copy()
        if pre_keyword:
            work = work[
                work[subject_col].astype(str).str.contains(pre_keyword, case=False, na=False)
                | work[content_col].astype(str).str.contains(pre_keyword, case=False, na=False)
            ]
        st.session_state["analysis_df"] = analyze_dataframe(work, cfg)
        st.session_state["source_name"] = uploaded.name

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
    edited = st.data_editor(
        show,
        use_container_width=True,
        num_rows="dynamic",
        hide_index=True,
        column_config={
            "選取": st.column_config.CheckboxColumn(help="勾選要批次處理的列"),
            "問題類型": st.column_config.SelectboxColumn(options=TYPE_OPTIONS, required=True),
            "問題細項": st.column_config.SelectboxColumn(options=DETAIL_OPTIONS, required=True),
            "部門": st.column_config.SelectboxColumn(options=DEPT_OPTIONS),
        },
        key="editor_table",
    )
    st.session_state["analysis_df"] = edited.copy()

    b1, b2, b3, b4 = st.columns(4)
    batch_type = b1.selectbox("批次設定問題類型", options=["(不變更)"] + TYPE_OPTIONS)
    detail_candidates = ["(不變更)"] + (TOPIC_DETAIL_MAP.get(batch_type, DETAIL_OPTIONS) if batch_type != "(不變更)" else DETAIL_OPTIONS)
    batch_detail = b2.selectbox("批次設定問題細項", options=detail_candidates)
    if b3.button("套用到勾選列"):
        mask = edited["選取"] == True
        if batch_type != "(不變更)":
            edited.loc[mask, "問題類型"] = batch_type
            edited.loc[mask, "部門"] = edited.loc[mask, "問題類型"].map(DEPT_MAP).fillna("未分配")
        if batch_detail != "(不變更)":
            edited.loc[mask, "問題細項"] = batch_detail
        # Auto-fix rows whose detail mismatches topic
        edited["問題細項"] = edited.apply(
            lambda r: r["問題細項"] if r["問題細項"] in TOPIC_DETAIL_MAP.get(r["問題類型"], []) else TOPIC_DETAIL_MAP.get(r["問題類型"], ["其他建議"])[0],
            axis=1,
        )
        st.session_state["analysis_df"] = edited
        st.success("已套用批次編輯。")
    if b4.button("刪除勾選列"):
        st.session_state["analysis_df"] = edited[edited["選取"] != True].copy()
        st.success("已刪除勾選列。")

    final_df = st.session_state["analysis_df"]
    
    st.markdown("#### 下載分析結果 (下載後自動歸檔至歷史紀錄)")
    dl_format = st.radio("選擇下載格式", ["Excel", "CSV", "PDF"], horizontal=True)
    
    def on_download():
        save_history(final_df, st.session_state.get("source_name", "unknown"))
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


def render_charts(df: pd.DataFrame, key_prefix: str = ""):
    stats = df["問題類型"].value_counts().rename_axis("問題類型").reset_index(name="件數")
    stats["百分比"] = (stats["件數"] / stats["件數"].sum() * 100).round(2)
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
    df = st.session_state["analysis_df"]
    if df.empty:
        st.warning("目前沒有資料。")
        return

    stats = df["問題類型"].value_counts().rename_axis("問題類型").reset_index(name="件數")
    stats["百分比"] = (stats["件數"] / stats["件數"].sum() * 100).round(2)
    stats["歸屬部門"] = stats["問題類型"].map(DEPT_MAP).fillna("未分配")

    st.markdown("#### 類型件數與部門")
    st.dataframe(stats, use_container_width=True, hide_index=True)

    render_charts(df, key_prefix="sec2")

    st.markdown("#### AI 問題重點分析")
    st.markdown("##### AI 設定（選填）")
    col_ai_1, col_ai_2 = st.columns([3, 2])
    key_input = col_ai_1.text_input("OpenAI API Key（若留空則使用內建規則摘要）", type="password")
    model_name = col_ai_2.text_input("模型", value="gpt-4o-mini")
    if key_input:
        st.session_state["OPENAI_API_KEY"] = key_input

    ai_text = generate_ai_summary_llm(df, model_name=model_name)
    st.text_area("分析摘要預覽", ai_text, height=140)
    st.download_button(
        "下載 AI 分析文字檔",
        data=ai_text.encode("utf-8"),
        file_name=f"{datetime.now().strftime('%Y%m%d')}_AI重點分析.txt",
        mime="text/plain",
    )
    ppt_bytes = build_ppt_bytes(stats, ai_text, st.session_state.get("source_name", "unknown"))
    st.download_button(
        "下載分析簡報 PPT",
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
    for item in history:
        out_path = Path(item["output_path"])
        if not out_path.exists():
            continue
        with st.expander(f"{item['created_at']}｜{item['output_name']}｜來源：{item['source_name']}（{item['rows']}筆）"):
            df_hist = pd.read_excel(out_path)
            tab_data, tab_chart, tab_ai = st.tabs(["📄 資料預覽", "📊 圖表分析", "🤖 AI 重點摘要"])
            
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
                if col2.button("✏️ 編輯紀錄", key=f"edit_{item['id']}"):
                    st.session_state["analysis_df"] = df_hist
                    st.session_state["source_name"] = item["source_name"]
                    st.session_state["menu"] = "上傳檔案區（分析區）"
                    st.rerun()
                if col3.button("🗑️ 刪除紀錄", key=f"del_{item['id']}"):
                    delete_history(item["id"])
                    st.rerun()
            
            with tab_chart:
                if not df_hist.empty:
                    render_charts(df_hist, key_prefix=f"hist_{item['id']}")
                else:
                    st.info("無資料可繪圖")
                    
            with tab_ai:
                st.info("點擊下方按鈕即時生成本檔案的 AI 重點摘要")
                if st.button("🤖 產生 AI 摘要", key=f"ai_btn_{item['id']}"):
                    with st.spinner("AI 分析中..."):
                        ai_result = generate_ai_summary_llm(df_hist)
                        st.markdown(ai_result)


def main():
    apply_brand_theme()
    st.markdown("<div class='ecoco-banner'>ECOCO 客訴智能分析平台</div>", unsafe_allow_html=True)
    st.markdown(
        """
        <div class='small-muted'>
          功能列表：1) 檔案上傳分析 2) 圖表化 + AI重點 3) 歷史紀錄（最新置頂）
        </div>
        """,
        unsafe_allow_html=True,
    )
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
        
        st.markdown(
            """
            <div style='height: 40vh;'></div>
            <div style='text-align: center; font-size: 0.75rem; color: rgba(255,255,255,0.7);'>
            202603© ECOCO宜可可循環經濟 客服課<br>※ 請尊重智慧財產權 ※
            </div>
            """,
            unsafe_allow_html=True
        )

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


if __name__ == "__main__":
    main()
