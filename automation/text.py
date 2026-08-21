"""文字正規化與個資遮蔽。

送往任何外部 LLM 之前，一律先過 mask_sensitive_df / mask_sensitive_text。
"""

import re

import pandas as pd


def _normalize_english_runs(text: str, mode: str) -> str:
    if not isinstance(text, str):
        return text
    repl = (lambda m: m.group(0).upper()) if mode == "upper" else (lambda m: m.group(0).lower())
    return re.sub(r"[A-Za-z]+", repl, text)


def normalize_problem_labels(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "問題類型" in out.columns:
        out["問題類型"] = out["問題類型"].map(lambda v: _normalize_english_runs(v, "upper"))
    if "問題細項" in out.columns:
        out["問題細項"] = out["問題細項"].map(lambda v: _normalize_english_runs(v, "lower"))
    return out


def mask_sensitive_text(value):
    if not isinstance(value, str):
        return value
    text = re.sub(r"([A-Za-z0-9._%+-]{2})[A-Za-z0-9._%+-]*(@[A-Za-z0-9.-]+\.[A-Za-z]{2,})", r"\1***\2", value)
    text = re.sub(r"(?<!\d)(09\d{2})\d{3}(\d{3})(?!\d)", r"\1***\2", text)
    text = re.sub(r"(?<!\d)(0\d{1,2}-?\d{2,4})\d{3,4}(\d{2,4})(?!\d)", r"\1***\2", text)
    return text


def mask_sensitive_df(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    text_cols = [
        c for c in out.columns
        if any(k in str(c).lower() for k in ["電話", "主旨", "內容", "email", "mail", "信箱", "姓名", "地址"])
    ]
    for col in text_cols:
        out[col] = out[col].map(mask_sensitive_text)
    return out
