"""從過往產出建立分類知識庫。

歷史紀錄裡每一筆都是已標記的資料，其中被人工改過（人工確認）的列是黃金標註。
這個模組把它們煉成三種可直接用於分類的資產：

  L0 指紋快取     正規化文字 → 當時的標記。重複／樣板客訴零成本命中。
  L1 挖掘規則     以 log-odds 算出每個 (類型, 細項) 的區辨詞，自動產生關鍵字規則，
                  取代必須人工不斷加字的規則鏈。
  L1c 相似案例    字元 n-gram + IDF 的倒排索引檢索最相似的歷史標註，
                  用相似度加權投票直接分類（免 API、不需訓練、不需額外套件）；
                  同一份索引也用來為 L2 LLM 挑 few-shot 範例。
                  信心來自留一法實測的準確率（knn_calibration）。

另外也從歷史補齊 DEPT_MAP 沒填的部門（例如「回收點數問題類型」）：用歷史多數決。
"""

from __future__ import annotations

import math
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from typing import Iterable, Optional

import pandas as pd

from . import config
from .taxonomy import FALLBACK_DETAIL, FALLBACK_TYPE, legalize_pair

SUBJECT_HINTS = ("主旨", "標題", "案由", "subject", "title")
CONTENT_HINTS = ("內容", "描述", "說明", "content", "detail", "description")
CONFIRMED_LAYERS = {"人工確認", "來源檔既有"}

_TOKEN_RE = re.compile(r"[A-Za-z0-9]+")
STOPWORDS = {
    "的", "了", "是", "我", "有", "在", "也", "很", "就", "都", "不", "與", "及", "或",
    "請", "謝謝", "你好", "您好", "麻煩", "一下", "可以", "問題", "請問", "為什麼",
}


def normalize_text(text: str) -> str:
    """正規化：去空白、統一大小寫，作為指紋比對的基礎。"""
    t = str(text or "").lower()
    t = re.sub(r"\s+", "", t)
    t = re.sub(r"[，。！？、；：,.!?;:~\-_/\\()（）\[\]「」『』]", "", t)
    return t


def tokenize(text: str) -> list[str]:
    """中文用 2-gram、英數用整詞，不需要外部分詞套件。"""
    t = normalize_text(text)
    if not t:
        return []
    tokens = _TOKEN_RE.findall(t)
    cjk = _TOKEN_RE.sub(" ", t)
    for run in cjk.split():
        if len(run) == 1:
            tokens.append(run)
        else:
            tokens.extend(run[i:i + 2] for i in range(len(run) - 1))
    return [tk for tk in tokens if tk not in STOPWORDS and len(tk) >= 1]


@dataclass
class MinedRule:
    topic: str
    detail: str
    terms: list[str]
    support: int
    score: float
    precision: float = 0.0      # 在歷史資料上實測的精確率（規則命中且標記正確的比例）
    fired: int = 0              # 命中次數（含判斷錯誤的）


@dataclass
class KnowledgeBase:
    """歷史知識庫。所有查詢在找不到資料時都回傳 None，呼叫端自然退回規則。"""

    fingerprints: dict[str, tuple[str, str, str, int]] = field(default_factory=dict)
    rules: list[MinedRule] = field(default_factory=list)
    dept_by_topic: dict[str, str] = field(default_factory=dict)
    builtin_precision: dict[tuple[str, str], float] = field(default_factory=dict)
    examples: list[tuple[str, str, str]] = field(default_factory=list)   # (text, topic, detail)
    example_tokens: list[set[str]] = field(default_factory=list)
    doc_freq: dict[str, int] = field(default_factory=dict)
    postings: dict[str, list[int]] = field(default_factory=dict)   # token → 範例編號（倒排索引）
    knn_calibration: dict[str, float] = field(default_factory=dict)  # 票數區間 → 實測精確率
    knn_topic_calibration: dict[str, float] = field(default_factory=dict)  # 同上，但只看類型
    stats: dict = field(default_factory=dict)

    # ── L0 ────────────────────────────────────────────────
    def lookup_fingerprint(self, text: str) -> Optional[tuple[str, str, str, int]]:
        return self.fingerprints.get(normalize_text(text))

    # ── L1 ────────────────────────────────────────────────
    def match_rules(self, text: str) -> Optional[tuple[str, str, float, str]]:
        """回傳 (類型, 細項, 信心, 依據)；沒有規則命中回傳 None。"""
        if not self.rules:
            return None
        tokens = set(tokenize(text))
        if not tokens:
            return None
        best: Optional[tuple[float, MinedRule, list[str]]] = None
        for rule in self.rules:
            hits = [t for t in rule.terms if t in tokens]
            if not hits:
                continue
            # 命中詞數與規則強度共同決定分數
            score = rule.score * (1 + 0.35 * (len(hits) - 1))
            if best is None or score > best[0]:
                best = (score, rule, hits)
        if best is None:
            return None
        score, rule, hits = best
        # 信心＝這條規則在歷史資料上實測的精確率（Laplace 平滑），
        # 而不是用區辨強度硬換算，否則會出現「信心很高但常常判錯」。
        confidence = round(min(0.95, max(0.20, rule.precision)), 3)
        why = (f"歷史規則命中「{'、'.join(hits[:3])}」"
               f"（歷史 {rule.support} 筆，實測準確率 {rule.precision:.0%}）")
        return rule.topic, rule.detail, confidence, why

    # ── L2 ────────────────────────────────────────────────
    def _scored_neighbours(self, text: str, k: int) -> list[tuple[float, int]]:
        """用倒排索引 + IDF 加權找最相似的歷史範例，回傳 [(相似度, 範例編號)]。"""
        if not self.examples:
            return []
        q = set(tokenize(text))
        if not q:
            return []
        n_docs = max(len(self.examples), 1)
        acc: dict[int, float] = {}
        for tk in q:
            ids = self.postings.get(tk)
            if not ids:
                continue
            # 極常見的詞（超過三成範例都有）幾乎沒有辨識力，跳過以免拖慢又失準
            if len(ids) > n_docs * 0.3:
                continue
            idf = math.log(1 + n_docs / (1 + self.doc_freq.get(tk, 1)))
            for i in ids:
                acc[i] = acc.get(i, 0.0) + idf
        if not acc:
            return []
        scored = [
            (score / math.sqrt(len(self.example_tokens[i]) or 1), i)
            for i, score in acc.items()
        ]
        scored.sort(reverse=True)
        return scored[:k]

    def similar_examples(self, text: str, k: int = 3) -> list[tuple[str, str, str]]:
        """最相似的歷史標註（供 LLM few-shot 使用）。"""
        return [self.examples[i] for _, i in self._scored_neighbours(text, k)]

    def knn_predict(self, text: str, k: int = 15) -> Optional[tuple[str, str, float, str]]:
        """相似案例投票分類（不需要 API，完全在本機跑）。

        取最相似的 k 筆歷史標註，用相似度當權重投票；
        信心來自「這個票數區間在歷史上實測的精確率」，不是憑票數硬換算。
        回傳 (類型, 細項, 信心, 依據)，找不到相似案例時回傳 None。
        """
        neighbours = self._scored_neighbours(text, k)
        if not neighbours:
            return None
        votes: dict[tuple[str, str], float] = {}
        total = 0.0
        for sim, i in neighbours:
            _txt, topic, detail = self.examples[i]
            votes[(topic, detail)] = votes.get((topic, detail), 0.0) + sim
            total += sim
        if total <= 0:
            return None
        (topic, detail), score = max(votes.items(), key=lambda kv: kv[1])
        share = score / total
        confidence = self._knn_confidence(share, len(neighbours))
        why = (f"與歷史 {len(neighbours)} 筆相似案例投票一致度 {share:.0%}"
               f"（實測準確率 {confidence:.0%}）")
        return topic, detail, confidence, why

    def knn_topic_predict(self, text: str, k: int = 15) -> Optional[tuple[str, float, str]]:
        """只判「問題類型」的相似案例投票。

        類型只有 6 種、細項有 45 種，所以類型的把握度通常遠高於細項。
        把兩者分開算，就能做到「類型可以自動採用、只有細項需要人工確認」，
        而且部門是由類型決定的，類型定了部門就能自動填。
        """
        neighbours = self._scored_neighbours(text, k)
        if not neighbours:
            return None
        votes: dict[str, float] = {}
        total = 0.0
        for sim, i in neighbours:
            _txt, topic, _detail = self.examples[i]
            votes[topic] = votes.get(topic, 0.0) + sim
            total += sim
        if total <= 0:
            return None
        topic, score = max(votes.items(), key=lambda kv: kv[1])
        share = score / total
        base = self.knn_topic_calibration.get(self._share_bucket(share))
        if base is None:
            base = min(0.95, 0.35 + share * 0.55)
        confidence = round(min(0.97, max(0.15, base)), 3)
        return topic, confidence, f"類型投票一致度 {share:.0%}（實測準確率 {confidence:.0%}）"

    @staticmethod
    def _share_bucket(share: float) -> str:
        for edge in (0.9, 0.75, 0.6, 0.45, 0.3):
            if share >= edge:
                return f"{edge:.2f}"
        return "0.00"

    def _knn_confidence(self, share: float, n_neighbours: int) -> float:
        base = self.knn_calibration.get(self._share_bucket(share))
        if base is None:
            base = min(0.9, 0.25 + share * 0.5)     # 沒校準資料時保守估
        if n_neighbours < 3:
            base *= 0.8                              # 鄰居太少就打折
        return round(min(0.95, max(0.15, base)), 3)

    # ── 內建規則的實測精確率 ────────────────────────────
    def builtin_confidence(self, topic: str, detail: str) -> Optional[float]:
        return self.builtin_precision.get((topic, detail))

    # ── 部門 ──────────────────────────────────────────────
    def dept_for_topic(self, topic: str) -> Optional[str]:
        return self.dept_by_topic.get(topic)


# ── 建立流程 ─────────────────────────────────────────────

def _pick_col(df: pd.DataFrame, hints: Iterable[str]) -> Optional[str]:
    for c in df.columns:
        name = str(c).lower()
        if any(h.lower() in name for h in hints):
            return c
    return None


def _iter_labeled_rows(frames: Iterable[pd.DataFrame]):
    """把歷史 DataFrame 攤成 (文字, 類型, 細項, 部門, 是否人工確認)。"""
    for df in frames:
        if df is None or df.empty:
            continue
        if "問題類型" not in df.columns or "問題細項" not in df.columns:
            continue
        subj_col = _pick_col(df, SUBJECT_HINTS)
        cont_col = _pick_col(df, CONTENT_HINTS)
        if subj_col is None and cont_col is None:
            continue
        layer_col = "_source_layer" if "_source_layer" in df.columns else None
        ai_col = "_ai_filled" if "_ai_filled" in df.columns else None
        dept_col = "部門" if "部門" in df.columns else None

        for _, row in df.iterrows():
            # 只學合乎現行分類法的標記，不合的舊資料直接略過
            legal = legalize_pair(str(row.get("問題類型", "")), str(row.get("問題細項", "")))
            if legal is None:
                continue
            topic, detail = legal
            text = " ".join(
                str(row.get(c, "")) for c in (subj_col, cont_col) if c is not None
            ).strip()
            if not text:
                continue
            dept = str(row.get(dept_col, "")).strip() if dept_col else ""
            confirmed = True
            if layer_col is not None:
                confirmed = str(row.get(layer_col, "")).strip() in CONFIRMED_LAYERS
            elif ai_col is not None:
                flag = str(row.get(ai_col, "")).strip().lower()
                confirmed = flag in ("false", "0", "", "nan")
            yield text, topic, detail, dept, confirmed


def _mine_rules(
    labeled: list[tuple[str, str, str]],
    min_support: int,
    max_terms: int = 8,
) -> list[MinedRule]:
    """以 log-odds 找出每個 (類型, 細項) 的區辨詞。"""
    by_label: dict[tuple[str, str], Counter] = defaultdict(Counter)
    label_docs: Counter = Counter()
    global_counts: Counter = Counter()

    for text, topic, detail in labeled:
        key = (topic, detail)
        toks = set(tokenize(text))
        if not toks:
            continue
        label_docs[key] += 1
        for t in toks:
            by_label[key][t] += 1
            global_counts[t] += 1

    total_docs = sum(label_docs.values())
    rules: list[MinedRule] = []
    for key, counts in by_label.items():
        support = label_docs[key]
        if support < min_support:
            continue
        if key == (FALLBACK_TYPE, FALLBACK_DETAIL):
            # 保底分類是「其他都丟這裡」的桶子，裡面的詞沒有辨識意義，
            # 挖成規則只會把新客訴也一起吸進來。
            continue
        scored_terms: list[tuple[float, str]] = []
        for term, cnt in counts.items():
            if cnt < max(2, min_support // 2):
                continue
            in_label = cnt / support
            elsewhere = (global_counts[term] - cnt) / max(total_docs - support, 1)
            # log-odds：詞在這個標籤內出現的機率相對於其他標籤的倍率
            lift = math.log((in_label + 0.01) / (elsewhere + 0.01))
            if lift <= 0.5 or len(term) < 2:
                continue
            scored_terms.append((lift * math.log(1 + cnt), term))
        if not scored_terms:
            continue
        scored_terms.sort(reverse=True)
        terms = [t for _, t in scored_terms[:max_terms]]
        strength = sum(s for s, _ in scored_terms[:max_terms]) / len(terms)
        rules.append(MinedRule(topic=key[0], detail=key[1], terms=terms,
                               support=support, score=strength))
    # 支持度高、區辨力強的規則優先比對
    rules.sort(key=lambda r: (r.score, r.support), reverse=True)
    _calibrate_rules(rules, labeled)
    return rules


def _calibrate_rules(rules: list[MinedRule], labeled: list[tuple[str, str, str]]) -> None:
    """用歷史資料實測每條規則的精確率，作為信心分數的依據。

    做法與 match_rules 相同（取分數最高的命中規則），
    因此測到的就是這條規則實際被採用時的正確率。
    """
    fired: dict[int, int] = {}
    correct: dict[int, int] = {}
    for text, topic, detail in labeled:
        tokens = set(tokenize(text))
        if not tokens:
            continue
        best_i = None
        best_score = 0.0
        for i, rule in enumerate(rules):
            hits = [t for t in rule.terms if t in tokens]
            if not hits:
                continue
            sc = rule.score * (1 + 0.35 * (len(hits) - 1))
            if best_i is None or sc > best_score:
                best_i, best_score = i, sc
        if best_i is None:
            continue
        fired[best_i] = fired.get(best_i, 0) + 1
        if (rules[best_i].topic, rules[best_i].detail) == (topic, detail):
            correct[best_i] = correct.get(best_i, 0) + 1
    for i, rule in enumerate(rules):
        f = fired.get(i, 0)
        c = correct.get(i, 0)
        rule.fired = f
        rule.precision = (c + 1) / (f + 2)   # Laplace 平滑，樣本少時自動保守


def _calibrate_knn(kb: "KnowledgeBase", sample: int = 800) -> None:
    """用留一法實測 kNN 各「投票一致度區間」的真實準確率。

    做法：拿知識庫自己的範例當查詢，但把自己從鄰居中排除，
    這樣測到的就是「遇到沒見過的相似案例時」的準確率，不會自我作弊。
    """
    n = len(kb.examples)
    if n < 30:
        return
    step = max(1, n // sample)
    hit: Counter = Counter()
    tot: Counter = Counter()
    t_hit: Counter = Counter()
    t_tot: Counter = Counter()
    for i in range(0, n, step):
        text, gold_topic, gold_detail = kb.examples[i]
        neighbours = [(sim, j) for sim, j in kb._scored_neighbours(text, 16) if j != i][:15]
        if not neighbours:
            continue
        votes: dict[tuple[str, str], float] = {}
        total = 0.0
        for sim, j in neighbours:
            _t, tp, dt = kb.examples[j]
            votes[(tp, dt)] = votes.get((tp, dt), 0.0) + sim
            total += sim
        if total <= 0:
            continue
        (tp, dt), score = max(votes.items(), key=lambda kv: kv[1])
        bucket = kb._share_bucket(score / total)
        tot[bucket] += 1
        if (tp, dt) == (gold_topic, gold_detail):
            hit[bucket] += 1

        # 同一批鄰居，只看類型的投票
        tvotes: dict[str, float] = {}
        for sim, j in neighbours:
            _t, tp2, _d = kb.examples[j]
            tvotes[tp2] = tvotes.get(tp2, 0.0) + sim
        top_topic, tscore = max(tvotes.items(), key=lambda kv: kv[1])
        tbucket = kb._share_bucket(tscore / total)
        t_tot[tbucket] += 1
        if top_topic == gold_topic:
            t_hit[tbucket] += 1

    kb.knn_calibration = {
        b: round((hit[b] + 1) / (t + 2), 3)      # Laplace 平滑
        for b, t in tot.items() if t >= 5
    }
    kb.knn_topic_calibration = {
        b: round((t_hit[b] + 1) / (t + 2), 3)
        for b, t in t_tot.items() if t >= 5
    }


def _dedupe_rows(rows: Iterable[tuple]) -> list[tuple]:
    """同一筆客訴只算一次。

    save_history() 會同時寫本機 history_reports/ 與 Google Sheets，
    而 build_knowledge() 兩邊都讀，所以每筆存過檔的資料都會出現兩次。
    指紋是用文字當 key 不受影響，但規則挖掘的支持度與 kNN 票數會被灌水一倍，
    連帶讓 min_support 的門檻形同減半。
    以「正規化後文字 + 標記」為鍵去重；同一段文字若被標成不同答案，
    那是真正的標記分歧，要保留下來讓後面的多數決處理。
    """
    pos: dict[tuple, int] = {}
    out: list[tuple] = []
    for row in rows:
        text, topic, detail, dept, confirmed = row
        key = (normalize_text(text), topic, detail, dept)
        i = pos.get(key)
        if i is not None:
            # 重複列裡只要有一筆是人工確認的，就保留確認狀態
            if confirmed and not out[i][4]:
                out[i] = out[i][:4] + (True,)
            continue
        pos[key] = len(out)
        out.append(row)
    return out


def build_from_history(
    frames: Iterable[pd.DataFrame],
    min_support: Optional[int] = None,
    max_examples: int = 4000,
) -> Optional[KnowledgeBase]:
    """從歷史 DataFrame 清單建立知識庫；資料太少時回傳 None。"""
    rows = _dedupe_rows(_iter_labeled_rows(frames))
    if not rows:
        return None

    kb = KnowledgeBase()
    min_support = min_support if min_support is not None else config.knowledge_min_support()

    # ── L0 指紋（人工確認過的優先，且記錄一致筆數）──
    grouped: dict[str, Counter] = defaultdict(Counter)
    dept_votes: dict[str, Counter] = defaultdict(Counter)
    labeled_for_rules: list[tuple[str, str, str]] = []
    confirmed_rows: list[tuple[str, str, str]] = []

    for text, topic, detail, dept, confirmed in rows:
        norm = normalize_text(text)
        if norm:
            weight = 3 if confirmed else 1
            grouped[norm][(topic, detail, dept)] += weight
        if confirmed:
            confirmed_rows.append((text, topic, detail))
        labeled_for_rules.append((text, topic, detail))
        if dept:
            dept_votes[topic][dept] += 3 if confirmed else 1

    for norm, counter in grouped.items():
        (topic, detail, dept), votes = counter.most_common(1)[0]
        kb.fingerprints[norm] = (topic, detail, dept, int(votes))

    # ── L1 規則：優先用人工確認過的資料學，數量不足才用全部 ──
    basis = confirmed_rows if len(confirmed_rows) >= max(20, min_support * 5) else labeled_for_rules
    kb.rules = _mine_rules(basis, min_support=min_support)

    # ── 內建規則的實測精確率（讓 L1 內建規則不再用固定信心）──
    from .rules import analyze_complaint

    bi_fired: Counter = Counter()
    bi_correct: Counter = Counter()
    for text, topic, detail in labeled_for_rules:
        pt, pd_ = analyze_complaint(text, "")
        bi_fired[(pt, pd_)] += 1
        if (pt, pd_) == (topic, detail):
            bi_correct[(pt, pd_)] += 1
    kb.builtin_precision = {
        key: (bi_correct[key] + 1) / (n + 2) for key, n in bi_fired.items()
    }

    # ── 部門多數決 ──
    for topic, counter in dept_votes.items():
        kb.dept_by_topic[topic] = counter.most_common(1)[0][0]

    # ── L2 few-shot 檢索池 ──
    pool = confirmed_rows if confirmed_rows else labeled_for_rules
    seen: set[str] = set()
    for text, topic, detail in pool:
        norm = normalize_text(text)
        if norm in seen:
            continue
        seen.add(norm)
        kb.examples.append((text, topic, detail))
        if len(kb.examples) >= max_examples:
            break
    df_counter: Counter = Counter()
    postings: dict[str, list[int]] = defaultdict(list)
    for idx, (text, _t, _d) in enumerate(kb.examples):
        toks = set(tokenize(text))
        kb.example_tokens.append(toks)
        for tk in toks:
            df_counter[tk] += 1
            postings[tk].append(idx)
    kb.doc_freq = dict(df_counter)
    kb.postings = dict(postings)
    _calibrate_knn(kb)

    kb.stats = {
        "knn_examples": len(kb.examples),
        "knn_calibration": kb.knn_calibration,
        "history_rows": len(rows),
        "confirmed_rows": len(confirmed_rows),
        "fingerprints": len(kb.fingerprints),
        "rules": len(kb.rules),
        "examples": len(kb.examples),
        "dept_learned": len(kb.dept_by_topic),
        "avg_rule_precision": (
            round(sum(r.precision for r in kb.rules) / len(kb.rules), 3) if kb.rules else 0.0
        ),
        "rule_basis": "人工確認" if basis is confirmed_rows else "全部歷史",
    }
    return kb
