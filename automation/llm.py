"""L2：LLM 分類與文字生成。

原則：
  * 送出前一律先遮蔽個資（automation/text.py）
  * 輸出限定在 TOPIC_DETAIL_MAP 之內，模型亂答會被夾回合法值
  * 帶入最相似的歷史標註做 few-shot，讓模型學公司實際的標記慣例
  * 只有信心不足的列才會走到這裡，成本受 LLM_MAX_ROWS 上限保護
"""

from __future__ import annotations

import json
import re
from typing import Callable, Optional, Sequence

from . import config
from .taxonomy import TOPIC_DETAIL_MAP
from .text import mask_sensitive_text

MAX_FIELD_CHARS = 300
MAX_EXAMPLES = 12


# ── 通用文字生成（摘要／報告共用）────────────────────────────

def complete_text(prompt: str, model: Optional[str] = None, max_tokens: int = 2000) -> Optional[str]:
    """呼叫可用的 LLM 產生文字；沒有可用 API 或失敗時回傳 None。"""
    anth_key = config.get_anthropic_key()
    if anth_key:
        try:
            import anthropic

            client = anthropic.Anthropic(api_key=anth_key)
            msg = client.messages.create(
                model=model or config.llm_report_model(),
                max_tokens=max_tokens,
                messages=[{"role": "user", "content": prompt}],
            )
            text = "".join(getattr(b, "text", "") for b in msg.content).strip()
            if text:
                return text
        except Exception:
            pass

    oa_key = config.get_openai_key()
    if oa_key:
        try:
            from openai import OpenAI

            client = OpenAI(api_key=oa_key)
            res = client.responses.create(
                model=model if (model and "gpt" in str(model)) else "gpt-4o-mini",
                input=prompt,
            )
            text = (getattr(res, "output_text", "") or "").strip()
            if text:
                return text
        except Exception:
            pass
    return None


def _taxonomy_block() -> str:
    lines = []
    for topic, details in TOPIC_DETAIL_MAP.items():
        lines.append(f"- {topic}：" + "、".join(details))
    return "\n".join(lines)


def _clip(text: str) -> str:
    text = mask_sensitive_text(str(text or "")).replace("\n", " ").strip()
    return text[:MAX_FIELD_CHARS]


def _parse_json_array(raw: str) -> list[dict]:
    """從模型輸出中撈出第一個 JSON 陣列。"""
    if not raw:
        return []
    fence = re.search(r"```(?:json)?\s*(.+?)```", raw, re.S)
    if fence:
        raw = fence.group(1)
    start, end = raw.find("["), raw.rfind("]")
    if start < 0 or end <= start:
        return []
    try:
        data = json.loads(raw[start:end + 1])
    except Exception:
        return []
    return [d for d in data if isinstance(d, dict)]


class LLMClassifier:
    """把低信心的列交給 LLM，帶歷史範例做 few-shot。"""

    def __init__(self, model: Optional[str] = None, batch_size: Optional[int] = None):
        self.model = model or config.llm_classifier_model()
        self.batch_size = batch_size or config.llm_batch_size()
        self.calls = 0

    def classify(
        self,
        pairs: Sequence[tuple[str, str]],
        examples_for: Optional[Callable[[str, int], list[tuple[str, str, str]]]] = None,
        progress=None,
    ) -> list:
        from .classifier import LAYER_LLM, Prediction, coerce  # 避免匯入循環

        results: list = [None] * len(pairs)
        total_batches = max(1, (len(pairs) + self.batch_size - 1) // self.batch_size)

        for b in range(total_batches):
            chunk_idx = list(range(b * self.batch_size, min(len(pairs), (b + 1) * self.batch_size)))
            if not chunk_idx:
                break
            chunk = [pairs[i] for i in chunk_idx]

            examples: list[tuple[str, str, str]] = []
            if examples_for is not None:
                seen = set()
                for subj, cont in chunk:
                    for ex in examples_for(f"{subj} {cont}", 3):
                        key = ex[0][:60]
                        if key not in seen:
                            seen.add(key)
                            examples.append(ex)
                examples = examples[:MAX_EXAMPLES]

            prompt = self._build_prompt(chunk, examples)
            raw = complete_text(prompt, model=self.model, max_tokens=1500)
            self.calls += 1
            if progress is not None:
                try:
                    progress((b + 1) / total_batches)
                except Exception:
                    pass
            if not raw:
                continue

            for item in _parse_json_array(raw):
                try:
                    pos = int(item.get("index", item.get("序號", 0))) - 1
                except (TypeError, ValueError):
                    continue
                if pos < 0 or pos >= len(chunk_idx):
                    continue
                topic, detail = coerce(
                    str(item.get("問題類型", item.get("topic", ""))),
                    str(item.get("問題細項", item.get("detail", ""))),
                )
                try:
                    conf = float(item.get("信心", item.get("confidence", 0.7)))
                except (TypeError, ValueError):
                    conf = 0.7
                conf = max(0.3, min(0.95, conf))
                reason = str(item.get("理由", item.get("reason", ""))).strip()[:120]
                results[chunk_idx[pos]] = Prediction(
                    topic=topic, detail=detail, confidence=conf,
                    layer=LAYER_LLM,
                    reason=f"LLM 判斷：{reason}" if reason else "LLM 判斷",
                )
        return results

    def _build_prompt(self, chunk: Sequence[tuple[str, str]], examples: Sequence[tuple[str, str, str]]) -> str:
        rows = "\n".join(
            f"{i}. 主旨：{_clip(s)}｜內容：{_clip(c)}"
            for i, (s, c) in enumerate(chunk, start=1)
        )
        example_block = ""
        if examples:
            example_block = "【公司過往的實際標記範例（請沿用同樣的判斷慣例）】\n" + "\n".join(
                f"- 「{_clip(text)}」→ {topic} / {detail}" for text, topic, detail in examples
            ) + "\n\n"

        return (
            "你是 ECOCO 資源回收公司的客訴分類助理。請依照公司既有分類法，為每一筆客訴指定問題類型與問題細項。\n\n"
            "【可用分類（只能從這裡挑，不可自創）】\n"
            f"{_taxonomy_block()}\n\n"
            f"{example_block}"
            "【待分類客訴】\n"
            f"{rows}\n\n"
            "請只輸出 JSON 陣列，每筆一個物件，不要有其他文字：\n"
            '[{"index":1,"問題類型":"...","問題細項":"...","信心":0.0~1.0,"理由":"20字內"}]\n'
            "信心代表你對這個分類的確定程度；若客訴描述不清或無法歸類，請給低於 0.5 的信心。"
        )
