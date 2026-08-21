"""Google Sheets 存取（不依賴 streamlit）。

供無人值守排程使用，也供 Streamlit 介面讀取歷史標記來建立知識庫。
憑證與試算表 ID 全部來自 automation/config.py（環境變數 / secrets）。
"""

from __future__ import annotations

import re
from datetime import datetime
from typing import Optional

import pandas as pd

from . import config

HISTORY_INDEX_SHEET = "歷史紀錄"
HISTORY_HEADER = ["id", "created_at", "source_name", "rows", "data_ref"]
CELL_CHAR_LIMIT = 45000
SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]


def get_client():
    """建立 gspread client；缺套件或缺憑證時回傳 None。"""
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        return None
    creds_dict = config.get_google_credentials()
    if not creds_dict:
        return None
    try:
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception:
        return None


def open_spreadsheet(sheet_id: str):
    client = get_client()
    if client is None or not sheet_id:
        return None
    try:
        return client.open_by_key(sheet_id)
    except Exception:
        return None


def history_index_sheet(create: bool = True):
    """取得歷史紀錄索引工作表。"""
    ss = open_spreadsheet(config.get_history_sheet_id())
    if ss is None:
        return None
    try:
        return ss.worksheet(HISTORY_INDEX_SHEET)
    except Exception:
        if not create:
            return None
        try:
            ws = ss.add_worksheet(HISTORY_INDEX_SHEET, rows=500, cols=6)
            ws.append_row(HISTORY_HEADER)
            return ws
        except Exception:
            return None


def worksheet_to_dataframe(ws) -> pd.DataFrame:
    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()
    header = values[0]
    width = len(header)
    rows = [(row + [""] * width)[:width] for row in values[1:]]
    return pd.DataFrame(rows, columns=header)


def read_history_frames(max_items: int = 60) -> list[pd.DataFrame]:
    """讀取歷史紀錄的資料工作表，供知識庫學習。"""
    ws = history_index_sheet(create=False)
    if ws is None:
        return []
    try:
        rows = ws.get_all_values()[1:]
    except Exception:
        return []
    rows = sorted(rows, key=lambda r: (r[1] if len(r) > 1 else ""), reverse=True)[:max_items]
    frames: list[pd.DataFrame] = []
    for row in rows:
        data_ref = row[4] if len(row) > 4 else ""
        if not data_ref.startswith("sheet:"):
            continue
        try:
            data_ws = ws.spreadsheet.worksheet(data_ref.split(":", 1)[1])
            df = worksheet_to_dataframe(data_ws)
            if not df.empty:
                frames.append(df)
        except Exception:
            continue
    return frames


def _sanitize_value(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value)
    if text.lower() in {"nan", "inf", "-inf", "infinity", "-infinity"}:
        return ""
    return text[:CELL_CHAR_LIMIT]


def _sanitize_frame(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy().replace([float("inf"), float("-inf")], pd.NA)
    out = out.astype(object).where(pd.notna(out), "")
    if hasattr(out, "map"):
        return out.map(_sanitize_value)
    return out.apply(lambda col: col.map(_sanitize_value))


def data_sheet_name(item_id: str) -> str:
    return f"history_{re.sub(r'[^0-9A-Za-z_-]+', '_', str(item_id))[:80]}"


def append_history(df: pd.DataFrame, source_name: str, item_id: str = "") -> Optional[str]:
    """把分析結果寫入歷史紀錄；成功回傳紀錄 id，失敗回傳 None。"""
    if config.history_readonly():
        return None
    ws = history_index_sheet(create=True)
    if ws is None:
        return None
    item_id = item_id or datetime.now().strftime("%Y%m%d_%H%M%S")
    name = data_sheet_name(item_id)
    clean = _sanitize_frame(df)
    values = [clean.columns.tolist()] + clean.values.tolist()
    try:
        try:
            data_ws = ws.spreadsheet.worksheet(name)
            data_ws.clear()
        except Exception:
            data_ws = ws.spreadsheet.add_worksheet(
                name, rows=max(len(values) + 10, 20), cols=max(len(clean.columns) + 2, 5)
            )
        data_ws.update(values=values, range_name="A1")

        for i, row in enumerate(ws.get_all_values()[1:], start=2):
            if row and row[0] == item_id:
                ws.delete_rows(i)
                break
        ws.append_row([
            item_id,
            datetime.now().isoformat(timespec="seconds"),
            source_name,
            str(len(df)),
            f"sheet:{name}",
        ])
        return item_id
    except Exception:
        return None


def read_source_frame() -> Optional[pd.DataFrame]:
    """讀取 SOURCE_SHEET_ID 指定的原始客訴表（排程模式的資料來源）。"""
    ss = open_spreadsheet(config.get_source_sheet_id())
    if ss is None:
        return None
    try:
        name = config.get_source_worksheet()
        ws = ss.worksheet(name) if name else ss.get_worksheet(0)
        return worksheet_to_dataframe(ws)
    except Exception:
        return None
