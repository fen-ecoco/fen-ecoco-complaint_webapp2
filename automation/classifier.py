"""瀑布式分類器。

    L0  指紋快取   ── 與歷史紀錄文字相同 → 直接沿用當時（多為人工確認過）的標記，零成本
    L1a 內建規則   ── 人工維護的關鍵字規則；公司明訂的政策細項（POLICY_DETAILS）以此為準
    L1b 歷史規則   ── 從歷史自動挖掘的關鍵字規則（log-odds 區辨詞）
    L1c 相似案例   ── 歷史標註的 kNN 投票（免 API、約 1ms/列）
    L2  LLM        ── 選配：帶最相似歷史標註做 few-shot，沒有 API key 時整層停用

L1a/L1b/L1c 的信心都是「在歷史資料上實測的準確率」，因此可以直接比大小，
取最可靠的那一個。每列都會回傳信心分數與判斷依據；
信心低於門檻者標記為待複核，介面只需要人工看這些列，其餘自動採用。
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional, Sequence

from . import config
from .rules import analyze_complaint, full_bin_detail
from .taxonomy import (
    DEPT_MAP,
    FALLBACK_DETAIL,
    FALLBACK_TYPE,
    POLICY_CONFIDENCE,
    POLICY_DETAILS,
    coerce_pair,
)

LAYER_CACHE = "L0-指紋"
LAYER_RULE = "L1-規則"
LAYER_MINED = "L1-歷史規則"
LAYER_KNN = "L1-相似案例"
LAYER_LLM = "L2-LLM"
LAYER_SOURCE = "來源檔既有"
LAYER_FALLBACK = "保底"


AGREE_MULTI = "多層一致"
AGREE_SINGLE = "單一依據"
AGREE_CONFLICT = "各層分歧"


@dataclass
class Prediction:
    topic: str = FALLBACK_TYPE
    detail: str = FALLBACK_DETAIL
    dept: str = ""
    confidence: float = 0.0
    layer: str = LAYER_FALLBACK
    reason: str = ""
    agreement: str = AGREE_SINGLE      # 交叉驗證結果
    candidates: str = ""              # 各層分別給了什麼答案（稽核軌跡）
    topic_confidence: float = 0.0      # 只看「問題類型」的信心（通常遠高於細項）

    @property
    def is_fallback(self) -> bool:
        return self.topic == FALLBACK_TYPE and self.detail == FALLBACK_DETAIL


def coerce(topic: str, detail: str) -> tuple[str, str]:
    """把任意輸出夾回合法的 (類型, 細項) 組合。"""
    return coerce_pair(topic, detail)


class CascadeClassifier:
    """把 L0/L1/L2 串起來的分類器。

    knowledge: automation.knowledge.KnowledgeBase 或 None（Phase 2 起提供）
    llm:       automation.llm.LLMClassifier 或 None（Phase 3 起提供）
    """

    def __init__(self, knowledge=None, llm=None, threshold: Optional[float] = None):
        self.knowledge = knowledge
        self.llm = llm
        self.threshold = (
            threshold if threshold is not None else config.review_confidence_threshold()
        )
        self.stats: dict[str, int] = {}

    # ── 單列（僅 L0 + L1，不含 LLM）────────────────────────
    def classify_fast(self, subject: str, content: str) -> Prediction:
        text = f"{subject or ''} {content or ''}".strip()

        # L0 指紋快取
        if self.knowledge is not None:
            hit = self.knowledge.lookup_fingerprint(text)
            if hit is not None:
                topic, detail, dept, n = hit
                topic, detail = coerce(topic, detail)
                if detail in POLICY_DETAILS:
                    # 公司政策明訂要依內容區分的細項：即使歷史有相同文字，
                    # 也要用政策重判，避免沿用政策訂立前的舊標記。
                    detail = full_bin_detail(text)
                    return Prediction(
                        topic=topic, detail=detail, dept=dept,
                        confidence=POLICY_CONFIDENCE, layer=LAYER_RULE,
                        topic_confidence=POLICY_CONFIDENCE,
                        reason="依公司政策由客訴內容判定滿艙種類",
                    )
                conf = 0.99 if n > 1 else 0.95
                return Prediction(
                    topic=topic, detail=detail, dept=dept,
                    confidence=conf, layer=LAYER_CACHE, topic_confidence=conf,
                    reason=f"與歷史 {n} 筆相同內容的標記一致",
                )

        # L1a 內建人工規則（最高優先）
        # 一律過 coerce：分類法若有細項改名或合併，規則吐出的舊名稱會自動歸位，
        # 不必同步修改整條規則鏈。
        topic, detail = coerce(*analyze_complaint(subject, content))
        measured = None
        if self.knowledge is not None:
            measured = self.knowledge.builtin_confidence(topic, detail)
        pred = Prediction(
            topic=topic, detail=detail,
            confidence=measured if measured is not None else 0.82,
            layer=LAYER_RULE,
            reason=("內建關鍵字規則命中"
                    + (f"（歷史實測準確率 {measured:.0%}）" if measured is not None else "")),
        )
        if not pred.is_fallback and detail in POLICY_DETAILS:
            # 公司明訂要依內容區分的細項：以內建規則為準，
            # 不採用舊資料挖出的規則（舊標記把三種滿艙混在一起）。
            pred.dept = self._dept_for(pred.topic)
            pred.confidence = max(pred.confidence, POLICY_CONFIDENCE)
            pred.topic_confidence = pred.confidence
            pred.reason = "依公司政策由客訴內容判定（不採用舊標記挖出的規則）"
            return pred

        # L1b/L1c：歷史挖掘規則與相似案例投票。
        # 兩者的信心都是「在歷史資料上實測的準確率」，可以直接比大小，
        # 取最可靠的那一個；都沒有結果才落到保底。
        candidates: list[Prediction] = []
        if not pred.is_fallback:
            pred.dept = self._dept_for(pred.topic)
            candidates.append(pred)
        if self.knowledge is not None:
            mined = self.knowledge.match_rules(text)
            if mined is not None:
                topic_m, detail_m, score_m, why_m = coerce(mined[0], mined[1]) + mined[2:]
                candidates.append(Prediction(
                    topic=topic_m, detail=detail_m, confidence=score_m,
                    layer=LAYER_MINED, reason=why_m,
                ))
            knn = self.knowledge.knn_predict(text)
            if knn is not None:
                topic_k, detail_k, score_k, why_k = coerce(knn[0], knn[1]) + knn[2:]
                candidates.append(Prediction(
                    topic=topic_k, detail=detail_k, confidence=score_k,
                    layer=LAYER_KNN, reason=why_k,
                ))
        if candidates:
            best = self._cross_check(candidates)
            best.dept = self._dept_for(best.topic)
            best.topic_confidence = self._topic_confidence(text, best, candidates)
            return best

        # 沒有任何依據 → 低信心保底，交給人工
        pred.confidence = 0.30
        pred.layer = LAYER_FALLBACK
        pred.reason = "無規則命中，落入保底分類"
        pred.dept = self._dept_for(pred.topic)
        return pred

    # ── 整批（含 L2 LLM）──────────────────────────────────
    def classify_many(
        self,
        pairs: Sequence[tuple[str, str]],
        progress=None,
    ) -> list[Prediction]:
        preds: list[Prediction] = [self.classify_fast(s, c) for s, c in pairs]

        if self.llm is not None:
            todo = [i for i, p in enumerate(preds) if p.confidence < self.threshold]
            cap = config.llm_max_rows()
            if len(todo) > cap:
                todo = todo[:cap]
            if todo:
                examples_for = None
                if self.knowledge is not None:
                    examples_for = self.knowledge.similar_examples
                llm_preds = self.llm.classify(
                    [(pairs[i][0], pairs[i][1]) for i in todo],
                    examples_for=examples_for,
                    progress=progress,
                )
                for i, lp in zip(todo, llm_preds):
                    if lp is None:
                        continue
                    if not lp.dept:
                        lp.dept = self._dept_for(lp.topic)
                    preds[i] = lp

        self.stats = {}
        for p in preds:
            self.stats[p.layer] = self.stats.get(p.layer, 0) + 1
        return preds

    # ── 類型層級信心 ────────────────────────────────────
    def _topic_confidence(self, text: str, best: Prediction,
                          candidates: list[Prediction]) -> float:
        """算「問題類型」單獨的信心。

        類型只有 6 種、細項 45 種，所以類型的把握度通常遠高於細項。
        分開算之後就能做到：類型與部門自動採用，只有細項需要人工確認 ——
        人工從「45 個細項裡挑」變成「這個類型的少數幾個細項裡挑」。
        """
        best_conf = best.confidence
        if self.knowledge is not None:
            topic_vote = self.knowledge.knn_topic_predict(text)
            if topic_vote is not None and topic_vote[0] == best.topic:
                best_conf = max(best_conf, topic_vote[1])
        # 各層對「類型」的看法一致，也是類型可靠的證據
        if len(candidates) >= 2 and all(p.topic == best.topic for p in candidates):
            best_conf = max(best_conf, min(0.95, best.confidence + config.agreement_boost()))
        return round(min(0.97, best_conf), 3)

    # ── 交叉驗證：多層互相對答案 ──────────────────────────
    def _cross_check(self, candidates: list[Prediction]) -> Prediction:
        """比對各層答案：一致就加信心，分歧就扣信心並記錄到稽核軌跡。

        每一層都是獨立的判斷依據（人工規則／歷史挖掘規則／相似案例投票），
        兩層以上給出同一個答案，代表這個判斷有交叉驗證支持，
        比單一來源的高分更可靠；反之若各層互相矛盾，就算最高分很高
        也應該讓人看一眼。
        """
        best = max(candidates, key=lambda p: p.confidence)
        best.candidates = "｜".join(
            f"{p.layer}:{p.detail}({p.confidence:.0%})" for p in candidates
        )
        if len(candidates) < 2:
            best.agreement = AGREE_SINGLE
            return best

        same_pair = [p for p in candidates
                     if (p.topic, p.detail) == (best.topic, best.detail)]
        if len(same_pair) >= 2:
            best.agreement = AGREE_MULTI
            layers = "、".join(sorted({p.layer for p in same_pair}))
            best.confidence = round(min(0.98, best.confidence + config.agreement_boost()), 3)
            best.reason = f"{best.reason}；{layers} 判斷一致（交叉驗證通過）"
            return best

        # 類型一致但細項不同 → 只算部分一致，不加分也不重罰
        if any(p.topic == best.topic for p in candidates if p is not best):
            best.agreement = AGREE_CONFLICT
            best.confidence = round(max(0.15, best.confidence - config.disagreement_penalty() / 2), 3)
            best.reason = f"{best.reason}；各層類型一致但細項不同，建議人工確認"
            return best

        best.agreement = AGREE_CONFLICT
        best.confidence = round(max(0.15, best.confidence - config.disagreement_penalty()), 3)
        best.reason = f"{best.reason}；各層判斷分歧（{best.candidates}），建議人工確認"
        return best

    # ── 部門推導 ────────────────────────────────────────
    def _dept_for(self, topic: str) -> str:
        learned = ""
        if self.knowledge is not None:
            learned = self.knowledge.dept_for_topic(topic) or ""
        if learned and config.prefer_learned_dept():
            return learned
        return (DEPT_MAP.get(topic, "") or "") or learned


def build_default(knowledge=None) -> CascadeClassifier:
    """依設定組出分類器：有 API key 且開關開啟才掛上 L2。"""
    llm = None
    if config.use_llm_classifier():
        try:
            from .llm import LLMClassifier  # noqa: PLC0415

            llm = LLMClassifier()
        except Exception:
            llm = None
    return CascadeClassifier(knowledge=knowledge, llm=llm)
