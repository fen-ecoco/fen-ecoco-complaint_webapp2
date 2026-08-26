"""欄位自動偵測。

原本需要人工從下拉選單挑「主旨欄 / 內容欄 / 日期欄」，
這裡改用「欄名字典 + 欄位內容啟發式」自動判斷，
只有在信心不足時才需要人工確認。
"""

from __future__ import annotations

import warnings
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd

# 輸出欄位（分析結果），不可被當成輸入欄
DERIVED_COLUMNS = {
    "問題類型", "問題細項", "部門", "已確認", "日期",
    "_ai_filled", "_confidence", "_source_layer", "_needs_review", "_reason",
    "AI標記",
}

SUBJECT_KEYWORDS = [
    "客訴主旨", "案件主旨", "主旨", "標題", "案由", "問題摘要", "摘要", "案件名稱",
    "subject", "title", "summary", "topic",
]
CONTENT_KEYWORDS = [
    "客訴內容", "問題內容", "案件內容", "內容", "描述", "詳細", "說明", "陳述",
    "意見", "反映事項", "content", "detail", "description", "body", "message", "remark",
]
DATE_KEYWORDS = [
    "客訴日期", "受理日期", "建立日期", "申請日期", "發生日期", "日期", "時間",
    "建立時間", "受理時間", "date", "time", "created", "timestamp",
]


@dataclass
class DetectedColumns:
    subject: Optional[str] = None
    content: Optional[str] = None
    date: Optional[str] = None
    confidence: dict = field(default_factory=dict)
    reasons: dict = field(default_factory=dict)

    @property
    def ok(self) -> bool:
        """主旨與內容都找到，且信心足夠 → 可直接自動分析。"""
        if not self.subject or not self.content:
            return False
        return min(
            self.confidence.get("subject", 0.0),
            self.confidence.get("content", 0.0),
        ) >= 0.5

    def as_dict(self) -> dict:
        return {"subject": self.subject, "content": self.content, "date": self.date}


def _name_score(col: str, keywords: list[str]) -> tuple[float, str]:
    """欄名比對分數：完全相等最高，其次為包含關係（越前面的關鍵字越優先）。"""
    name = str(col).strip().lower()
    for i, kw in enumerate(keywords):
        k = kw.lower()
        if name == k:
            return 0.98 - i * 0.001, f"欄名與「{kw}」完全相符"
        if k in name:
            return 0.88 - i * 0.001, f"欄名包含「{kw}」"
    return 0.0, ""


def _text_stats(series: pd.Series) -> tuple[float, float]:
    """回傳 (非空比例, 平均字數)。"""
    s = series.dropna().astype(str).str.strip()
    s = s[s != ""]
    if series.empty or s.empty:
        return 0.0, 0.0
    return len(s) / max(len(series), 1), float(s.str.len().mean())


def _date_ratio(series: pd.Series) -> float:
    """可被解析為日期的比例。"""
    s = series.dropna()
    if s.empty:
        return 0.0
    if pd.api.types.is_datetime64_any_dtype(series):
        return 1.0
    sample = s.head(200)
    with warnings.catch_warnings():
        # 來源檔日期格式不一致是常態，這裡只需要「可否解析」的比例
        warnings.simplefilter("ignore")
        parsed = pd.to_datetime(sample, errors="coerce")
    return float(parsed.notna().sum()) / max(len(sample), 1)


def detect_columns(df: pd.DataFrame) -> DetectedColumns:
    """從 DataFrame 推斷主旨 / 內容 / 日期欄位。"""
    out = DetectedColumns()
    if df is None or df.empty or len(df.columns) == 0:
        return out

    candidates = [c for c in df.columns if str(c) not in DERIVED_COLUMNS]
    if not candidates:
        return out

    # ── 日期欄 ──
    date_scored: list[tuple[float, str, str]] = []
    for c in candidates:
        ratio = _date_ratio(df[c])
        name_s, name_why = _name_score(c, DATE_KEYWORDS)
        if name_s > 0 and ratio >= 0.5:
            date_scored.append((min(0.98, name_s + 0.05), c, f"{name_why}，且 {ratio:.0%} 可解析為日期"))
        elif name_s > 0 and ratio >= 0.2:
            date_scored.append((0.6, c, f"{name_why}，但僅 {ratio:.0%} 可解析為日期"))
        elif ratio >= 0.85 and not pd.api.types.is_numeric_dtype(df[c]):
            date_scored.append((0.7, c, f"{ratio:.0%} 的值可解析為日期"))
    if date_scored:
        date_scored.sort(reverse=True, key=lambda x: x[0])
        score, col, why = date_scored[0]
        out.date, out.confidence["date"], out.reasons["date"] = col, score, why

    # ── 文字欄統計 ──
    text_cols: list[tuple[str, float, float]] = []   # (col, 非空比例, 平均字數)
    for c in candidates:
        if c == out.date:
            continue
        if pd.api.types.is_numeric_dtype(df[c]) or pd.api.types.is_datetime64_any_dtype(df[c]):
            continue
        filled, avg_len = _text_stats(df[c])
        if filled >= 0.2 and avg_len >= 2:
            text_cols.append((c, filled, avg_len))

    # ── 主旨欄 / 內容欄：先看欄名 ──
    subj_named = sorted(
        ((_name_score(c, SUBJECT_KEYWORDS)[0], c, _name_score(c, SUBJECT_KEYWORDS)[1]) for c, _, _ in text_cols),
        reverse=True, key=lambda x: x[0],
    )
    cont_named = sorted(
        ((_name_score(c, CONTENT_KEYWORDS)[0], c, _name_score(c, CONTENT_KEYWORDS)[1]) for c, _, _ in text_cols),
        reverse=True, key=lambda x: x[0],
    )
    if subj_named and subj_named[0][0] > 0:
        out.subject = subj_named[0][1]
        out.confidence["subject"] = subj_named[0][0]
        out.reasons["subject"] = subj_named[0][2]
    if cont_named and cont_named[0][0] > 0 and cont_named[0][1] != out.subject:
        out.content = cont_named[0][1]
        out.confidence["content"] = cont_named[0][0]
        out.reasons["content"] = cont_named[0][2]

    # ── 內容啟發式：字數最長的文字欄視為內容欄 ──
    by_len = sorted(text_cols, key=lambda x: x[2], reverse=True)
    if out.content is None:
        for c, _filled, avg_len in by_len:
            if c != out.subject:
                out.content = c
                out.confidence["content"] = 0.55 if avg_len >= 15 else 0.45
                out.reasons["content"] = f"平均字數最長（{avg_len:.0f} 字），推定為內容欄"
                break
    if out.subject is None:
        for c, _filled, avg_len in by_len:
            if c != out.content:
                out.subject = c
                out.confidence["subject"] = 0.5
                out.reasons["subject"] = f"次長的文字欄（平均 {avg_len:.0f} 字），推定為主旨欄"
                break

    # 只有一個文字欄時，主旨與內容共用
    if out.content and not out.subject:
        out.subject = out.content
        out.confidence["subject"] = 0.4
        out.reasons["subject"] = "檔案僅有一個文字欄，主旨與內容共用"
    if out.subject and not out.content:
        out.content = out.subject
        out.confidence["content"] = 0.4
        out.reasons["content"] = "檔案僅有一個文字欄，主旨與內容共用"

    return out
