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

import json
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
    "客服主題 [list]": "問題類型",
}

# 機台類型要併兩欄：主欄說是哪一種機台（收瓶機／電池機／方舟站），
# 方舟站另有子欄說它收什麼（寶特瓶），只取主欄會把這個資訊丟掉。
MACHINE_COLUMN = "機台類型 [list]"
MACHINE_SUB_COLUMN = "機台類型(方舟站) [txt]"
ARK_STATION = "方舟站"

# 第一層分類的各種寫法 → 統一後的問題類型。
# 同一個分類在 CSV 匯出是「機台需要協助(滿袋無需回報)」、JSON 匯出是代碼
# 「機台需要協助_滿袋無需回報_」，公司統一只要「機台需要協助」。
TOPIC_CANONICAL = {
    "機台需要協助_滿袋無需回報_": "機台需要協助",
    "機台需要協助(滿袋無需回報)": "機台需要協助",
    "機台需要協助（滿袋無需回報）": "機台需要協助",
    "顧客關係_": "顧客關係",
    "回收點數問題_": "回收點數問題",
    "優惠券兌換問題_": "優惠券兌換問題",
    "app_使用問題_": "APP 使用問題",
    "註冊帳號問題_": "註冊帳號問題",
}

# 問題類型 → 該類型的細項欄。工單系統每種主題各有一欄細項，
# 只有與「客服主題」相符的那一欄才是這張單真正的細項。
TOPIC_DETAIL_COLUMN = {
    "機台需要協助": "機台問題類型 [list]",
    "回收點數問題": "回收點數問題類型 [list]",
    "註冊帳號問題": "APP帳號設定問題類型 [list]",
    "優惠券兌換問題": "優惠券問題類型 [list]",
    "APP 使用問題": "APP使用問題類型 [list]",
    "顧客關係": "顧客關係類型 [list]",
}


def normalize_topic(value) -> str:
    """第一層分類的任何寫法 → 統一的問題類型。

    沒收錄的寫法只去掉代碼尾巴的底線，不自作主張改內容。
    """
    v = str(value).strip()
    if v.lower() in _BLANK:
        return ""
    if v in TOPIC_CANONICAL:
        return TOPIC_CANONICAL[v]
    return v.rstrip("_")
DETAIL_COLUMNS = list(TOPIC_DETAIL_COLUMN.values())

# 判斷「這是不是原始工單匯出」用的指紋欄位
_SIGNATURE = ("Id", "Subject", "Created at", "客服主題 [list]")

# 公司挑單規則：只處理網頁表單進來的工單。
# Line / IG / Mail / Facebook 的來信不走客訴流程，不納入分析。
VIA_COLUMN = "Via"
WEB_FORM = "Web form"
CREATED_COLUMN = "Created at"

# 歷史紀錄裡代表「客訴進件時間」的欄位（接續分析的水位就看這個）
_DATE_COLUMNS = ("進件日期", "日期")

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


# ── Zendesk JSON 匯出 ────────────────────────────────────────
# 工單平台也能直接匯出 JSON（NDJSON，一行一張工單）。欄位是數字 id，
# 這份對照表是拿 137 筆同時出現在 CSV 與 JSON 的工單逐欄比對出來的。
ZENDESK_FIELDS = {
    41243924547865: "聯繫電話  [txt]",
    41492126053913: "商家名稱與門市 [txt]",
    41492314326681: "使用日期",            # CSV 匯出沒有這一欄，JSON 才有
    41492327655449: "使用時間 [txt]",
    41492532308761: "站點名稱 [list]",
    41492601887897: "機台類型 [list]",
    41492613173401: "機台類型(方舟站) [txt]",
    49277052627993: "客服主題 [list]",
    49273684609305: "機台問題類型 [list]",
    49274860688921: "回收點數問題類型 [list]",
    49275362537881: "APP帳號設定問題類型 [list]",
    49275451718553: "優惠券問題類型 [list]",
    49275558878617: "APP使用問題類型 [list]",
    49275607775641: "顧客關係類型 [list]",
}

# via.channel → CSV 匯出的 Via 值（兩種匯出要篩出同一批工單）
CHANNEL_TO_VIA = {
    "web": "Web form",
    "line": "Line",
    "email": "Mail",
    "instagram_dm": "Instagram DM",
    "sunshine_conversations_facebook_messenger":
        "Sunshine Conversations Facebook Messenger",
}

# 工單時間是 UTC，客服看的與歷史紀錄用的都是台灣時間
TW_OFFSET = pd.Timedelta(hours=8)


def parse_zendesk_json(raw: str | bytes) -> list[dict]:
    """讀 Zendesk 匯出：整份是 JSON 陣列，或 NDJSON（一行一張工單）。"""
    if isinstance(raw, bytes):
        raw = raw.decode("utf-8-sig", errors="replace")
    raw = raw.lstrip("﻿").strip()
    if not raw:
        return []
    try:
        doc = json.loads(raw)
    except json.JSONDecodeError:
        out = []
        for line in raw.splitlines():
            line = line.strip().rstrip(",")
            if line in ("", "[", "]"):
                continue
            try:
                out.append(json.loads(line))
            except json.JSONDecodeError:
                continue
        return out
    if isinstance(doc, dict):
        for key in ("tickets", "results", "data"):
            if isinstance(doc.get(key), list):
                return doc[key]
        return [doc]
    return doc if isinstance(doc, list) else []


def looks_like_zendesk_json(tickets) -> bool:
    return bool(tickets) and isinstance(tickets[0], dict) and         {"id", "subject", "created_at"} <= set(tickets[0])


def zendesk_to_frame(tickets) -> pd.DataFrame:
    """把工單 JSON 攤平成與 CSV 匯出同名的欄位，之後共用同一套清洗邏輯。"""
    rows = []
    for t in tickets:
        row = {
            "Id": t.get("id"),
            "Subject": _clean(t.get("subject")),
            # JSON 多了工單內文，CSV 匯出沒有這一欄 —— 它才是用戶真正寫的內容
            "Description": _clean(t.get("description")),
            "Status": _clean(t.get("status")),
            "Via": CHANNEL_TO_VIA.get(
                ((t.get("via") or {}).get("channel") or ""),
                _clean((t.get("via") or {}).get("channel")),
            ),
        }
        created = pd.to_datetime(t.get("created_at"), errors="coerce", utc=True)
        row[CREATED_COLUMN] = "" if pd.isna(created) else             (created.tz_localize(None) + TW_OFFSET).strftime("%Y-%m-%d %H:%M:%S")
        for cf in t.get("custom_fields") or []:
            name = ZENDESK_FIELDS.get(cf.get("id"))
            if name:
                row[name] = _clean(cf.get("value"))
        # 使用日期與使用時間在 JSON 是分開的兩欄，併成 CSV 的「使用時間」
        date, time = row.pop("使用日期", ""), row.get("使用時間 [txt]", "")
        row["使用時間 [txt]"] = f"{date} {time}".strip() if date else time
        rows.append(row)
    cols = ["Id", "Subject", "Description", "Status", "Via", CREATED_COLUMN] +         [c for c in dict.fromkeys(ZENDESK_FIELDS.values()) if c != "使用日期"]
    return pd.DataFrame(rows).reindex(columns=cols).fillna("")


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


def latest_processed_at(frames) -> pd.Timestamp | None:
    """歷史紀錄裡最新的一筆進件時間，就是「上次分析到哪裡」的水位。"""
    best = None
    for frame in frames or []:
        for col in _DATE_COLUMNS:
            if col in getattr(frame, "columns", []):
                t = pd.to_datetime(frame[col], errors="coerce").max()
                if pd.notna(t) and (best is None or t > best):
                    best = t
                break
    return best


def select_rows(df: pd.DataFrame, since=None,
                web_form_only: bool = True) -> tuple[pd.DataFrame, dict]:
    """套用公司的挑單規則，回傳 (挑選後的資料, 各階段筆數)。

    規則由公司提供：只要 Via = Web form，且 Created at 接續上次分析的
    最後一筆之後（嚴格大於，避免同一筆重複分析）。
    """
    stats = {"total": len(df), "web_form": len(df), "new": len(df)}
    out = df

    if web_form_only and VIA_COLUMN in out.columns:
        out = out[out[VIA_COLUMN].astype(str).str.strip() == WEB_FORM]
    stats["web_form"] = len(out)

    if since is not None and CREATED_COLUMN in out.columns:
        created = pd.to_datetime(out[CREATED_COLUMN], errors="coerce")
        # 解析不出時間的列保留下來，寧可多看一筆也不要漏掉客訴
        out = out[created.isna() | (created > pd.Timestamp(since))]
    stats["new"] = len(out)
    return out, stats


def clean_raw_export(df: pd.DataFrame, drop_empty_content: bool = True) -> pd.DataFrame:
    """把原始工單匯出轉成 13 欄的分析格式。

    drop_empty_content：沒有任何文字內容的列無法分析，預設剔除。
    """
    out = pd.DataFrame(index=df.index, columns=CLEAN_COLUMNS, dtype=object)

    for src, tgt in DIRECT_MAP.items():
        out[tgt] = df[src].map(_clean) if src in df.columns else ""
    out["問題類型"] = out["問題類型"].map(normalize_topic)

    # 歷史紀錄裡的進件編號一律是 Z + 工單號（5354 筆全部如此），
    # 少了這個前綴，同一張工單會在歷史裡變成兩筆不同的客訴。
    out["進件編號"] = out["進件編號"].map(
        lambda v: v if not v or v[:1].isalpha() else f"Z{v}")

    machines = []
    for _idx, row in df.iterrows():
        main = _val(row, MACHINE_COLUMN)
        sub = _val(row, MACHINE_SUB_COLUMN)
        # 子欄只有方舟站才有意義（說明收的是哪一種回收物）；其他機台填的是
        # 「第2台收瓶機」這類自由描述，併進來只會讓機台類型沒辦法分組。
        machines.append(f"{main}({sub})" if main == ARK_STATION and sub else main)
    out["機台類型"] = machines

    sites = df["站點名稱 [list]"] if "站點名稱 [list]" in df.columns else pd.Series([""] * len(df), index=df.index)
    parsed = sites.map(split_site)
    out["站點編號"] = [p[0] for p in parsed]
    out["站點區域"] = [p[1] for p in parsed]
    out["站點名稱"] = [p[2] for p in parsed]

    # JSON 匯出另有工單內文（Description），那才是用戶真正寫的內容，
    # 這時主旨與內容各歸各位；CSV 匯出沒有這一欄，只能把標題當內容用。
    split_subject = "Description" in df.columns
    subjects, contents = [], []
    for _idx, row in df.iterrows():
        title = _val(row, "Subject")
        shop = _val(row, "商家名稱與門市 [txt]")
        body = _val(row, "Description") or _val(row, "聯絡訊息 [txt]")
        if split_subject:
            subjects.append(title)
            # 門市名稱是客訴定位的關鍵字（「全聯台中中清店」），併進內容才查得到
            text = body or title
            contents.append(f"{text}{shop}" if shop and shop not in text else text)
        else:
            subjects.append("")
            joined = title + shop
            if body and body not in joined:
                joined = f"{joined} {body}".strip()
            contents.append(joined)
    out["問題主旨"] = subjects
    out["用戶內容"] = contents

    # 主題：來源沒有對應欄位，一律留空
    out["主題"] = ""

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


def read_zendesk_json(raw: str | bytes) -> pd.DataFrame:
    """JSON 匯出 → 與 CSV 匯出同名欄位的 DataFrame。內容不是工單就回空表。"""
    tickets = parse_zendesk_json(raw)
    if not looks_like_zendesk_json(tickets):
        raise ValueError("這份 JSON 不是客服系統的工單匯出")
    return zendesk_to_frame(tickets)


def normalize_upload(df: pd.DataFrame, since=None,
                     web_form_only: bool = True) -> tuple[pd.DataFrame, str]:
    """上傳檔案的統一入口：需要轉換就轉換，回傳 (資料, 說明文字)。

    since 為上次分析到的進件時間，只處理之後的新進件；None 代表不限。
    挑單規則只套用在原始工單匯出上 —— 已經整理過的檔案是人工挑好的，
    再篩一次只會把對方刻意放進來的列吃掉。

    說明文字為空字串代表沒有做任何轉換。
    """
    if looks_like_raw_export(df):
        picked, stats = select_rows(df, since=since, web_form_only=web_form_only)
        out = clean_raw_export(picked)
        steps = [f"原始工單匯出 {stats['total']} 列"]
        if web_form_only:
            steps.append(f"Web form {stats['web_form']} 列")
        if since is not None:
            steps.append(f"{pd.Timestamp(since):%Y/%m/%d %H:%M} 之後的新進件 {stats['new']} 列")
        if len(out) != stats["new"]:
            steps.append(f"剔除 {stats['new'] - len(out)} 列無內容")
        note = f"已自動轉換成分析格式（{len(df.columns)} 欄 → 13 欄）：" + " → ".join(steps)
        return out, note

    trimmed = drop_annotation_row(df)
    if len(trimmed) != len(df):
        return trimmed, "已略過第 2 列的欄位對照註記（非客訴資料）"
    return df, ""
