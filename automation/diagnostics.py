"""分類品質診斷：用歷史資料量化「哪裡分不開、合併能拿回多少」。

分類法要怎麼調整是業務決策，但決策需要數字。這個模組提供：
  * 時間切分評估（早期資料建知識庫、晚期資料當測試，模擬真實上線）
  * 各細項的可辨識度
  * 最常互相混淆的細項配對（雙向混淆＝定義重疊的強訊號）
  * 合併模擬：把幾個細項併成一個之後，準確率與自動化率會變成多少

不改任何設定，只做量測。
"""

from __future__ import annotations

import warnings
from contextlib import contextmanager
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from typing import Callable, Iterable, Optional

import pandas as pd

from .classifier import AGREE_CONFLICT, CascadeClassifier
from .knowledge import CONTENT_HINTS, SUBJECT_HINTS, _pick_col, build_from_history
from .taxonomy import DETAIL_TO_TOPIC, POLICY_DETAILS, legalize_pair

DATE_HINTS = ("進件日期", "受理日期", "客訴日期", "日期", "date")


@dataclass
class LabeledRow:
    date: pd.Timestamp
    subject: str
    content: str
    topic: str
    detail: str


@dataclass
class EvalResult:
    total: int = 0
    topic_acc: float = 0.0
    detail_acc: float = 0.0
    auto_rate: float = 0.0
    auto_topic_acc: float = 0.0
    auto_detail_acc: float = 0.0
    topic_only_rate: float = 0.0
    full_manual_rate: float = 0.0
    rules: int = 0
    per_detail: dict = field(default_factory=dict)      # 細項 → (筆數, 正確數)
    confusion: Counter = field(default_factory=Counter)  # (人工標, 系統判) → 次數

    def describe(self) -> str:
        return (
            f"測試 {self.total} 筆／規則 {self.rules} 條\n"
            f"  全體：類型 {self.topic_acc:.1%}　細項 {self.detail_acc:.1%}\n"
            f"  A 完全自動採用 {self.auto_rate:.1%}"
            f"（類型 {self.auto_topic_acc:.1%}、細項 {self.auto_detail_acc:.1%}）\n"
            f"  B 僅需確認細項 {self.topic_only_rate:.1%}\n"
            f"  C 需完整判斷   {self.full_manual_rate:.1%}"
        )


def load_labeled(frames: Iterable[pd.DataFrame]) -> list[LabeledRow]:
    """把歷史資料攤成帶日期的已標記列（只保留合乎現行分類法者）。"""
    rows: list[LabeledRow] = []
    for df in frames:
        if df is None or df.empty:
            continue
        if "問題類型" not in df.columns or "問題細項" not in df.columns:
            continue
        subj = _pick_col(df, SUBJECT_HINTS)
        cont = _pick_col(df, CONTENT_HINTS)
        date_col = _pick_col(df, DATE_HINTS)
        if subj is None and cont is None:
            continue
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            dates = (pd.to_datetime(df[date_col], errors="coerce")
                     if date_col else pd.Series([pd.NaT] * len(df)))
        for (_, row), when in zip(df.iterrows(), dates):
            legal = legalize_pair(str(row.get("問題類型", "")), str(row.get("問題細項", "")))
            if legal is None or pd.isna(when):
                continue
            rows.append(LabeledRow(
                date=when,
                subject=str(row.get(subj, "")) if subj else "",
                content=str(row.get(cont, "")) if cont else "",
                topic=legal[0], detail=legal[1],
            ))
    rows.sort(key=lambda r: r.date)
    return rows


def _to_frame(rows: Iterable[LabeledRow], remap: Optional[Callable[[str], str]] = None):
    remap = remap or (lambda d: d)
    return pd.DataFrame([{
        "問題主旨": r.subject, "用戶內容": r.content,
        "問題類型": r.topic, "問題細項": remap(r.detail),
        "_source_layer": "人工確認",
    } for r in rows])


def evaluate(rows: list[LabeledRow], cut: pd.Timestamp,
             remap: Optional[Callable[[str], str]] = None) -> EvalResult:
    """時間切分評估：cut 之前建知識庫，cut 之後當測試集。"""
    remap = remap or (lambda d: d)
    train = [r for r in rows if r.date < cut]
    test = [r for r in rows if r.date >= cut]
    result = EvalResult(total=len(test))
    if not train or not test:
        return result

    kb = build_from_history([_to_frame(train, remap)])
    clf = CascadeClassifier(knowledge=kb)
    result.rules = len(getattr(kb, "rules", []))
    th = clf.threshold

    auto = topic_only = topic_ok = detail_ok = auto_t = auto_d = 0
    for r in test:
        gold_detail = remap(r.detail)
        p = clf.classify_fast(r.subject, r.content)
        pred_detail = remap(p.detail)
        t_ok, d_ok = p.topic == r.topic, pred_detail == gold_detail
        topic_ok += t_ok
        detail_ok += d_ok
        stat = result.per_detail.setdefault(gold_detail, [0, 0])
        stat[0] += 1
        stat[1] += d_ok
        if not d_ok:
            result.confusion[(gold_detail, pred_detail)] += 1

        clean = p.agreement != AGREE_CONFLICT
        if p.confidence >= th and clean:
            auto += 1
            auto_t += t_ok
            auto_d += d_ok
        elif (p.topic_confidence or 0) >= th and clean:
            topic_only += 1

    n = len(test)
    result.topic_acc = topic_ok / n
    result.detail_acc = detail_ok / n
    result.auto_rate = auto / n
    result.auto_topic_acc = auto_t / auto if auto else 0.0
    result.auto_detail_acc = auto_d / auto if auto else 0.0
    result.topic_only_rate = topic_only / n
    result.full_manual_rate = (n - auto - topic_only) / n
    return result


def confusion_pairs(result: EvalResult, top: int = 15) -> list[dict]:
    """最常混淆的細項配對；雙向都混淆代表兩者定義重疊。"""
    out = []
    for (gold, pred), n in result.confusion.most_common(top):
        reverse = result.confusion.get((pred, gold), 0)
        out.append({
            "人工標記": gold, "系統判斷": pred, "次數": n,
            "反向次數": reverse, "雙向混淆": "是" if reverse else "",
            "合計": n + reverse,
        })
    return out


def separability(result: EvalResult, min_count: int = 10) -> list[dict]:
    """各細項的可辨識度（測試期出現次數達門檻者）。"""
    rows = [
        {"問題細項": d, "測試筆數": tot, "判對筆數": ok, "可辨識度": round(ok / tot, 3)}
        for d, (tot, ok) in result.per_detail.items() if tot >= min_count
    ]
    rows.sort(key=lambda r: r["可辨識度"])
    return rows


def suggest_merges(result: EvalResult, min_pair: int = 6,
                   max_group: int = 3) -> dict[str, list[str]]:
    """依雙向混淆建議合併群組。

    三個必要的限制（少了任何一個，建議就會荒謬）：
      1. 只合併同一個「問題類型」底下的細項 —— 跨類型合併沒有意義
      2. 不碰 POLICY_DETAILS —— 那是公司明訂要分開的（例如三種滿艙）
      3. 不用連通分量串連 —— A 與 B 混淆、B 與 C 混淆，不代表 A 與 C 該併；
         只把「彼此都互相混淆」的細項放進同一組，且限制組員數
    """
    mutual: dict[tuple[str, str], int] = {}
    for (a, b), n in result.confusion.items():
        if a == b or a in POLICY_DETAILS or b in POLICY_DETAILS:
            continue
        if DETAIL_TO_TOPIC.get(a) != DETAIL_TO_TOPIC.get(b):
            continue                      # 跨類型不合併
        rev = result.confusion.get((b, a), 0)
        if rev <= 0 or n + rev < min_pair:
            continue                      # 必須雙向混淆
        mutual[tuple(sorted((a, b)))] = n + rev

    groups: list[list[str]] = []
    used: set[str] = set()
    for (a, b), _score in sorted(mutual.items(), key=lambda kv: -kv[1]):
        if a in used or b in used:
            continue
        group = [a, b]
        # 只有跟組內每個成員都互相混淆的細項才能加入
        for cand in {x for pair in mutual for x in pair} - used - set(group):
            if len(group) >= max_group:
                break
            if all(tuple(sorted((cand, g))) in mutual for g in group):
                group.append(cand)
        used.update(group)
        groups.append(sorted(group))

    return {
        f"{DETAIL_TO_TOPIC.get(g[0], '')}：合併建議{i + 1}": g
        for i, g in enumerate(groups)
    }


@contextmanager
def _taxonomy_with_merges(groups: dict[str, list[str]]):
    """暫時把合併後的細項名稱加進分類法。

    合併後的新名稱若不在 TOPIC_DETAIL_MAP 裡，legalize_pair() 會判定不合法，
    知識庫就會把那些列整批丟掉，模擬結果會嚴重失真（看起來像合併讓準確率變差）。
    """
    from . import taxonomy as tx

    orig_map = {t: list(ds) for t, ds in tx.TOPIC_DETAIL_MAP.items()}
    orig_detail_topic = dict(tx.DETAIL_TO_TOPIC)
    try:
        for group, details in groups.items():
            topic = tx.DETAIL_TO_TOPIC.get(details[0])
            if topic is None:
                continue
            merged = [d for d in tx.TOPIC_DETAIL_MAP[topic] if d not in details]
            merged.append(group)
            tx.TOPIC_DETAIL_MAP[topic] = merged
            tx.DETAIL_TO_TOPIC[group] = topic
            for d in details:
                tx.DETAIL_TO_TOPIC.pop(d, None)
        yield
    finally:
        tx.TOPIC_DETAIL_MAP.clear()
        tx.TOPIC_DETAIL_MAP.update(orig_map)
        tx.DETAIL_TO_TOPIC.clear()
        tx.DETAIL_TO_TOPIC.update(orig_detail_topic)


def simulate_merge(rows: list[LabeledRow], cut: pd.Timestamp,
                   groups: dict[str, list[str]]) -> EvalResult:
    """模擬「把這些細項併成一個」之後的表現（知識庫也用合併後的分類重建）。"""
    mapping = {d: g for g, ds in groups.items() for d in ds}
    with _taxonomy_with_merges(groups):
        return evaluate(rows, cut, remap=lambda d: mapping.get(d, d))
