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


def lower_english(value):
    """把字串裡的英文字母轉小寫，中文與符號不動。

    問題細項的比對用；分類法裡同時存在 "APP無法登入" 與 "app無法登入"，
    介面與篩選都走這個函式，才不會把同一個細項當成兩個。
    """
    return _normalize_english_runs(value, "lower")


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


PHONE_COL_HINTS = ("手機", "電話", "門號", "phone", "mobile", "tel")


def mask_phone_value(value):
    """遮蔽整格就是一組號碼的欄位（例如「帳號手機」）。

    mask_sensitive_text 是針對「內文裡夾帶號碼」設計的，
    對 09xxxxxxxx 這種整格號碼要連空白、+886、分隔符一起處理，
    所以獨立一個函式，中間位數一律換成 ****。
    """
    if value is None:
        return value
    text = str(value).strip()
    if not text or text.lower() in ("nan", "none"):
        return value
    digits = re.sub(r"\D", "", text)
    if digits.startswith("886"):
        digits = "0" + digits[3:]
    # Excel 常把「0912345678」當數字存，開頭的 0 會不見；補回來才不會遮錯位數
    if len(digits) == 9 and digits.startswith("9"):
        digits = "0" + digits
    if len(digits) < 8:                       # 不像號碼就交給一般文字遮蔽
        return mask_sensitive_text(str(value))
    return f"{digits[:4]}****{digits[-2:]}" if len(digits) <= 10 else f"{digits[:4]}****{digits[-3:]}"


def mask_sensitive_df(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    phone_cols = [c for c in out.columns
                  if any(k in str(c).lower() for k in PHONE_COL_HINTS)]
    for col in phone_cols:
        out[col] = out[col].map(mask_phone_value)
    text_cols = [
        c for c in out.columns
        if c not in phone_cols
        and any(k in str(c).lower() for k in ["主旨", "內容", "email", "mail", "信箱", "姓名", "地址"])
    ]
    for col in text_cols:
        out[col] = out[col].map(mask_sensitive_text)
    return out
