"""分析核心：把一份原始客訴表轉成已標記的分析結果。

不依賴 streamlit，Streamlit 介面與無介面排程（automation/cli.py）共用同一套邏輯。
"""

from __future__ import annotations

import random
from dataclasses import dataclass
from typing import Optional

import pandas as pd

from . import config
from .classifier import AGREE_CONFLICT, CascadeClassifier, LAYER_SOURCE, Prediction
from .rules import _is_valid_pair
from .taxonomy import DEPT_MAP, TOPIC_DETAIL_MAP, default_detail_for
from .text import mask_sensitive_df, normalize_problem_labels

# 分析結果的內部欄位（介面上不顯示，但會隨結果一起儲存，作為稽核軌跡）
META_COLUMNS = [
    "_ai_filled", "_confidence", "_source_layer", "_needs_review", "_reason",
    "_agreement", "_candidates", "_review_cause", "_topic_confidence",
]


@dataclass
class AnalysisConfig:
    subject_col: str
    content_col: str
    date_col: Optional[str]


def make_unique_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = []
    seen: dict[str, int] = {}
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


def analyze_dataframe(
    df: pd.DataFrame,
    cfg: AnalysisConfig,
    classifier: Optional[CascadeClassifier] = None,
    progress=None,
) -> pd.DataFrame:
    """對整份資料做分類，回傳含標記、信心分數與待複核旗標的 DataFrame。

    來源檔已有合法的 (問題類型, 問題細項) → 原樣保留，不重新分類。
    其餘由 classifier 決定；信心低於門檻者標記 _needs_review=True。
    """
    out = make_unique_columns(mask_sensitive_df(df.copy()))
    if classifier is None:
        classifier = CascadeClassifier()

    # ------ 保留來源檔既有的合法標記 ------
    existing_type = (
        out["問題類型"].copy() if "問題類型" in out.columns else pd.Series([""] * len(out))
    )
    existing_detail = (
        out["問題細項"].copy() if "問題細項" in out.columns else pd.Series([""] * len(out))
    )

    for c in ["問題類型", "問題細項", "已確認", "部門", "日期"] + META_COLUMNS:
        if c in out.columns:
            out = out.drop(columns=[c])

    subjects = out.get(cfg.subject_col, pd.Series([""] * len(out))).fillna("").astype(str)
    contents = out.get(cfg.content_col, pd.Series([""] * len(out))).fillna("").astype(str)

    # 來源檔已標記的列不必送分類（也不必花 LLM 成本）
    keep_mask: list[bool] = []
    for idx in range(len(out)):
        t = str(existing_type.iloc[idx]).strip() if idx < len(existing_type) else ""
        d = str(existing_detail.iloc[idx]).strip() if idx < len(existing_detail) else ""
        keep_mask.append(_is_valid_pair(t, d))

    todo_idx = [i for i, keep in enumerate(keep_mask) if not keep]
    pairs = [(subjects.iloc[i], contents.iloc[i]) for i in todo_idx]
    todo_preds = classifier.classify_many(pairs, progress=progress) if pairs else []

    preds: list[Prediction] = [None] * len(out)  # type: ignore[list-item]
    for i, p in zip(todo_idx, todo_preds):
        preds[i] = p
    for i, keep in enumerate(keep_mask):
        if keep:
            t = str(existing_type.iloc[i]).strip()
            d = str(existing_detail.iloc[i]).strip()
            if d not in TOPIC_DETAIL_MAP.get(t, []):
                d = default_detail_for(t)
            preds[i] = Prediction(
                topic=t, detail=d, dept=DEPT_MAP.get(t, "") or "",
                confidence=1.0, layer=LAYER_SOURCE, reason="來源檔已有合法標記，原樣保留",
            )

    threshold = classifier.threshold
    out["問題類型"] = [p.topic for p in preds]
    out["問題細項"] = [p.detail for p in preds]
    out["部門"] = [p.dept or DEPT_MAP.get(p.topic, "") or "" for p in preds]
    out["已確認"] = False
    out["_ai_filled"] = [p.layer != LAYER_SOURCE for p in preds]
    out["_confidence"] = [round(p.confidence, 3) for p in preds]
    out["_source_layer"] = [p.layer for p in preds]
    out["_reason"] = [p.reason for p in preds]
    out["_agreement"] = [p.agreement for p in preds]
    out["_candidates"] = [p.candidates for p in preds]
    out["_topic_confidence"] = [round(p.topic_confidence or p.confidence, 3) for p in preds]

    # ── 審核判定：低信心、各層分歧、或被抽樣稽核抽中，都要人工看 ──
    causes: list[str] = []
    for p in preds:
        if p.layer == LAYER_SOURCE:
            causes.append("")
            continue
        if p.confidence < threshold:
            # 類型有把握、只有細項沒把握 → 人工只需要挑細項，
            # 類型與部門可以照系統的判斷走，複核成本低很多。
            if (p.topic_confidence or 0) >= threshold:
                causes.append("僅需確認細項")
            else:
                causes.append("信心不足")
        elif p.agreement == AGREE_CONFLICT:
            causes.append("各層判斷分歧")
        else:
            causes.append("")
    causes = _add_audit_samples(causes, preds, threshold)
    out["_review_cause"] = causes
    out["_needs_review"] = [bool(c) for c in causes]

    if cfg.date_col and cfg.date_col in out.columns:
        out["日期"] = pd.to_datetime(out[cfg.date_col], errors="coerce")

    return normalize_problem_labels(out)


def _add_audit_samples(causes: list[str], preds: list[Prediction],
                       threshold: float, rate: Optional[float] = None) -> list[str]:
    """從「自動採用」的列中隨機抽一小部分進人工稽核。

    自動採用的品質若不抽驗，就只能靠上線前那一次評估；
    抽樣讓實際準確率能持續被量測（抽中的列會標記為稽核抽樣）。
    抽樣用固定種子，同一份資料重跑抽到同一批，方便追蹤。
    """
    rate = config.audit_sample_rate() if rate is None else rate
    if rate <= 0:
        return causes
    auto_idx = [
        i for i, (c, p) in enumerate(zip(causes, preds))
        if not c and p.layer != LAYER_SOURCE
    ]
    if not auto_idx:
        return causes
    n_sample = max(1, int(round(len(auto_idx) * rate))) if len(auto_idx) >= 10 else 0
    if not n_sample:
        return causes
    rng = random.Random(len(preds) * 7919 + len(auto_idx))
    for i in rng.sample(auto_idx, min(n_sample, len(auto_idx))):
        causes[i] = "稽核抽樣"
    return causes


def review_summary(df: pd.DataFrame) -> dict:
    """回傳自動化程度指標，供介面／排程報表顯示。"""
    total = len(df)
    if total == 0:
        return {"total": 0, "auto": 0, "review": 0, "auto_rate": 0.0, "layers": {}}
    review = int(df["_needs_review"].fillna(False).astype(bool).sum()) if "_needs_review" in df.columns else 0
    layers = (
        df["_source_layer"].value_counts().to_dict() if "_source_layer" in df.columns else {}
    )
    causes = (
        df["_review_cause"].replace("", pd.NA).dropna().value_counts().to_dict()
        if "_review_cause" in df.columns else {}
    )
    agreement = (
        df["_agreement"].replace("", pd.NA).dropna().value_counts().to_dict()
        if "_agreement" in df.columns else {}
    )
    return {
        "total": total,
        "auto": total - review,
        "review": review,
        "auto_rate": (total - review) / total,
        "layers": layers,
        "review_causes": causes,
        "agreement": agreement,
    }


def needs_review_frame(df: pd.DataFrame) -> pd.DataFrame:
    if "_needs_review" not in df.columns:
        return df.iloc[0:0]
    return df[df["_needs_review"].fillna(False).astype(bool)]


def visible_columns(df: pd.DataFrame) -> list[str]:
    """介面表格要顯示的欄位：隱藏所有底線開頭的內部欄位。"""
    return [c for c in df.columns if not str(c).startswith("_")]
