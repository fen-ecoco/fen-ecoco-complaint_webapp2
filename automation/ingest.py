# -*- coding: utf-8 -*-
"""原始工單匯出 → 分析用欄位格式的自動轉換。

客服系統（工單平台）匯出的 CSV 有 70 幾個欄位，欄名是英文混中文標記
（`Created at`、`聯繫電話  [txt]`、`客服主題 [list]`…），而且
站點代號／區域／名稱三項擠在同一欄。原本要人工整理成 13 欄的分析格式，
這支模組把那段人工作業自動化。

對應關係是從公司既有的「清洗前 / 清洗後」範例逆推並逐格驗證的
（13 欄 × 38 列，其中 12 欄完全吻合；差異見 pick_detail 的註解）。
"""

from __future__ import annotations

import re

import pandas as pd

# 分析格式的欄位（順序就是輸出順序）
CLEAN_COLUMNS = [
    "進件日期", "進件編號", "帳號手機", "主題",
    "站點編號", "站點區域", "站點名稱",
    "日期時間", "機台類型",
    "問題主旨", "用戶內容", "問題類型", "問題細項",
]

# 原始匯出的欄名 → 分析格式欄位（一對一的部分）
DIRECT_MAP = {
    "Created at": "進件日期",
    "Id": "進件編號",
    "聯繫電話  [txt]": "帳號手機",
    "使用時間 [txt]": "日期時間",
    "機台類型 [list]": "機台類型",
    "客服主題 [list]": "問題類型",
}

# 問題類型 → 該類型的細項欄。工單系統每種主題各有一欄細項，
# 只有與「客服主題」相符的那一欄才是這張單真正的細項。
TOPIC_DETAIL_COLUMN = {
    "機台需要協助(滿袋無需回報)": "機台問題類型 [list]",
    "回收點數問題": "回收點數問題類型 [list]",
    "註冊帳號問題": "APP帳號設定問題類型 [list]",
    "優惠券兌換問題": "優惠券問題類型 [list]",
    "APP 使用問題": "APP使用問題類型 [list]",
    "顧客關係": "顧客關係類型 [list]",
}
DETAIL_COLUMNS = list(TOPIC_DETAIL_COLUMN.values())

# 判斷「這是不是原始工單匯出」用的指紋欄位
_SIGNATURE = ("Id", "Subject", "Created at", "客服主題 [list]")

# 站點欄格式：`es0249【新北】全聯福利中心中和新生店站`
_SITE_RE = re.compile(r"^\s*(\S+?)\s*【(.*?)】\s*(.*?)\s*$")

_BLANK = {"", "-", "nan", "none", "null"}


def _clean(value) -> str:
    """把工單系統的空值表示（`-`、NaN）正規化成空字串。"""
    v = str(value).strip()
    return "" if v.lower() in _BLANK else v


def _val(row: pd.Series, col: str) -> str:
    """取某一欄的值並正規化；沒有該欄就回空字串。"""
    return _clean(row[col]) if col in row.index else ""


def looks_like_raw_export(df: pd.DataFrame) -> bool:
    """這份資料是不是尚未整理的原始工單匯出。"""
    cols = set(map(str, df.columns))
    return sum(1 for c in _SIGNATURE if c in cols) >= 3


def split_site(value: str) -> tuple[str, str, str]:
    """`es0249【新北】全聯福利中心中和新生店站` → (編號, 區域, 名稱)。

    格式不符時整串當站點名稱，不硬拆（寧可留著給人看，也不要拆錯）。
    """
    v = str(value or "").strip()
    if not v or v.lower() in _BLANK:
        return "", "", ""
    m = _SITE_RE.match(v)
    if not m:
        return "", "", v
    return m.group(1), m.group(2), m.group(3)


def pick_detail(row: pd.Series, topic: str) -> str:
    """挑出這張單的問題細項。

    依「客服主題」取對應的細項欄。公司原本的人工整理是把所有非空的細項欄
    直接串起來，38 列裡有 5 列因此變成「瓶蓋桶已滿APP商家頁面空白」
    這種不存在的標記 —— 串接後的字串不在分類法裡，反而會被當成無效標記。
    這裡改成依主題挑對應那一欄，其餘情況才退回串接。
    """
    col = TOPIC_DETAIL_COLUMN.get(str(topic).strip())
    if col:
        picked = _val(row, col)
        if picked:
            return picked
    rest = [_val(row, c) for c in DETAIL_COLUMNS]
    rest = [v for v in rest if v]
    if len(rest) == 1:
        return rest[0]
    return "／".join(rest)


def clean_raw_export(df: pd.DataFrame, drop_empty_content: bool = True) -> pd.DataFrame:
    """把原始工單匯出轉成 13 欄的分析格式。

    drop_empty_content：沒有任何文字內容的列無法分析，預設剔除。
    """
    out = pd.DataFrame(index=df.index, columns=CLEAN_COLUMNS, dtype=object)

    for src, tgt in DIRECT_MAP.items():
        out[tgt] = df[src].map(_clean) if src in df.columns else ""

    sites = df["站點名稱 [list]"] if "站點名稱 [list]" in df.columns else pd.Series([""] * len(df), index=df.index)
    parsed = sites.map(split_site)
    out["站點編號"] = [p[0] for p in parsed]
    out["站點區域"] = [p[1] for p in parsed]
    out["站點名稱"] = [p[2] for p in parsed]

    # 用戶內容＝工單標題＋商家門市。門市名稱是客訴定位的關鍵字
    #（「全聯台中中清店」），併進內容才查得到。
    contents = []
    for idx, row in df.iterrows():
        text = _val(row, "Subject")
        shop = _val(row, "商家名稱與門市 [txt]")
        # 工單另有「聯絡訊息」長文欄位，有填就接在後面
        msg = _val(row, "聯絡訊息 [txt]")
        joined = text + shop
        if msg and msg not in joined:
            joined = f"{joined} {msg}".strip()
        contents.append(joined)
    out["用戶內容"] = contents

    # 主題與問題主旨：原始匯出沒有對應欄位，留空給人工或分析後回填
    out["主題"] = ""
    out["問題主旨"] = ""

    out["問題細項"] = [pick_detail(row, out.at[idx, "問題類型"]) for idx, row in df.iterrows()]

    out = out.fillna("")
    if drop_empty_content:
        out = out[out["用戶內容"].astype(str).str.strip() != ""]
    return out.reset_index(drop=True)


def drop_annotation_row(df: pd.DataFrame) -> pd.DataFrame:
    """剔除「清洗後」檔案第二列的欄位對照註記。

    公司現行的清洗後 CSV 第二列不是資料，而是記錄每一欄來自原始匯出的哪個
    欄位（`Created at,Id,聯繫電話  [txt],…`）。那一列混進分析會變成一筆假客訴。
    """
    if df.empty or "進件編號" not in df.columns:
        return df
    first = str(df["進件編號"].iloc[0]).strip()
    if first in ("Id", "進件編號") or (first and not any(ch.isdigit() for ch in first)):
        return df.iloc[1:].reset_index(drop=True)
    return df


def normalize_upload(df: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    """上傳檔案的統一入口：需要轉換就轉換，回傳 (資料, 說明文字)。

    說明文字為空字串代表沒有做任何轉換。
    """
    if looks_like_raw_export(df):
        before = len(df)
        out = clean_raw_export(df)
        note = (f"偵測到原始工單匯出（{len(df.columns)} 欄），"
                f"已自動轉換成分析格式：{before} 列 → {len(out)} 列"
                f"（剔除 {before - len(out)} 列無內容）")
        return out, note

    trimmed = drop_annotation_row(df)
    if len(trimmed) != len(df):
        return trimmed, "已略過第 2 列的欄位對照註記（非客訴資料）"
    return df, ""
