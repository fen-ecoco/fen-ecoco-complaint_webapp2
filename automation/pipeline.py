"""端到端流程：資料進來 → 分類 → 統計 → 產出。

Streamlit 介面與排程 CLI 都走這裡，確保兩邊行為一致。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd

from . import config, sheets
from .classifier import build_default
from .columns import DetectedColumns, detect_columns
from .core import AnalysisConfig, analyze_dataframe, review_summary
from .knowledge import build_from_history


@dataclass
class PipelineResult:
    df: pd.DataFrame
    columns: DetectedColumns
    summary: dict
    source_name: str
    history_id: Optional[str] = None
    knowledge_stats: dict = field(default_factory=dict)
    outputs: dict = field(default_factory=dict)   # 檔名 → 路徑
    report_text: str = ""

    def describe(self) -> str:
        s = self.summary
        lines = [
            f"來源：{self.source_name}",
            f"欄位：主旨={self.columns.subject}／內容={self.columns.content}／日期={self.columns.date or '（無）'}",
            f"筆數：{s.get('total', 0)}　自動採用：{s.get('auto', 0)}"
            f"（{s.get('auto_rate', 0):.0%}）　待複核：{s.get('review', 0)}",
        ]
        if s.get("layers"):
            lines.append("判斷來源：" + "、".join(f"{k} {v}" for k, v in s["layers"].items()))
        if s.get("review_causes"):
            lines.append("需人工原因：" + "、".join(f"{k} {v}" for k, v in s["review_causes"].items()))
        if s.get("agreement"):
            lines.append("交叉驗證：" + "、".join(f"{k} {v}" for k, v in s["agreement"].items()))
        if self.knowledge_stats:
            k = self.knowledge_stats
            lines.append(
                f"知識庫：歷史 {k.get('history_rows', 0)} 筆／規則 {k.get('rules', 0)} 條"
                f"／指紋 {k.get('fingerprints', 0)} 組"
            )
        if self.history_id:
            lines.append(f"已寫入歷史紀錄：{self.history_id}")
        for name, path in self.outputs.items():
            lines.append(f"產出：{name} → {path}")
        return "\n".join(lines)


def load_frame(path: str | Path) -> pd.DataFrame:
    """讀取 xlsx / xls / csv（排程模式的本機檔案來源）。"""
    p = Path(path)
    suffix = p.suffix.lower()
    if suffix in (".xlsx", ".xls"):
        return pd.read_excel(p)
    if suffix in (".csv", ".tsv", ".txt"):
        for enc in ("utf-8-sig", "utf-8", "cp950", "big5"):
            try:
                return pd.read_csv(p, encoding=enc, sep=None, engine="python")
            except (UnicodeDecodeError, ValueError):
                continue
        return pd.read_csv(p, encoding="utf-8", errors="replace")
    raise ValueError(f"不支援的檔案格式：{suffix or p.name}")


def read_local_history_frames(folder: Optional[str | Path] = None,
                              max_items: int = 60) -> list[pd.DataFrame]:
    """讀取本機歷史紀錄（history_reports/ 下的 xlsx / xls / csv）。

    Render 沒有持久化磁碟所以只能靠 Google Sheets，
    但公司內部主機有磁碟，這裡就是最穩定的學習來源。
    """
    folder = Path(folder or config.get_local_history_dir())
    if not folder.exists():
        return []
    files = [f for f in folder.iterdir()
             if f.suffix.lower() in (".xlsx", ".xls", ".csv")]
    files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    frames: list[pd.DataFrame] = []
    for f in files[:max_items]:
        try:
            df = load_frame(f)
            if not df.empty:
                frames.append(df)
        except Exception:
            continue
    return frames


def build_knowledge(frames: Optional[list[pd.DataFrame]] = None):
    """建立知識庫；讀不到歷史時回傳 None（分類自動退回規則）。

    未指定 frames 時，同時學習雲端歷史紀錄與本機 history_reports/。
    """
    try:
        if frames is None:
            frames = read_local_history_frames()
            try:
                frames = frames + sheets.read_history_frames()
            except Exception:
                pass
        return build_from_history(frames or [])
    except Exception:
        return None


def run(
    source: pd.DataFrame | str | Path,
    source_name: str = "",
    out_dir: Optional[str | Path] = None,
    save_history: Optional[bool] = None,
    make_report: bool = False,
    knowledge=None,
    progress=None,
) -> PipelineResult:
    """跑完整流程。source 可以是 DataFrame 或檔案路徑。"""
    if isinstance(source, (str, Path)):
        df_raw = load_frame(source)
        source_name = source_name or Path(source).name
    else:
        df_raw = source.copy()
        source_name = source_name or f"資料集_{datetime.now():%Y%m%d_%H%M%S}"

    det = detect_columns(df_raw)
    if not det.subject or not det.content:
        raise ValueError("無法判斷主旨／內容欄位，請確認來源資料的欄位名稱")

    if knowledge is None:
        knowledge = build_knowledge()
    classifier = build_default(knowledge=knowledge)

    cfg = AnalysisConfig(subject_col=det.subject, content_col=det.content, date_col=det.date)
    df = analyze_dataframe(df_raw, cfg, classifier=classifier, progress=progress)

    result = PipelineResult(
        df=df,
        columns=det,
        summary=review_summary(df),
        source_name=source_name,
        knowledge_stats=getattr(knowledge, "stats", {}) or {},
    )

    if out_dir:
        out = Path(out_dir)
        out.mkdir(parents=True, exist_ok=True)
        stem = f"{datetime.now():%Y%m%d_%H%M%S}_客訴分析"
        xlsx_path = out / f"{stem}.xlsx"
        df.to_excel(xlsx_path, index=False)
        result.outputs["Excel"] = str(xlsx_path)

        review_df = df[df["_needs_review"].fillna(False).astype(bool)] if "_needs_review" in df.columns else df.iloc[0:0]
        if not review_df.empty:
            # 最沒把握的排最前面，人工先看最可能出錯的列
            if "_confidence" in review_df.columns:
                review_df = review_df.sort_values("_confidence", kind="stable")
            review_path = out / f"{stem}_待複核.xlsx"
            review_df.to_excel(review_path, index=False)
            result.outputs["待複核清單"] = str(review_path)

            # 稽核抽樣單獨一份：這批系統有把握，是抽驗品質用的
            if "_review_cause" in review_df.columns:
                audit_df = review_df[review_df["_review_cause"] == "稽核抽樣"]
                if not audit_df.empty:
                    audit_path = out / f"{stem}_稽核抽樣.xlsx"
                    audit_df.to_excel(audit_path, index=False)
                    result.outputs["稽核抽樣"] = str(audit_path)

    if make_report:
        result.report_text = build_report_text(df, source_name)
        if out_dir:
            report_path = Path(out_dir) / f"{datetime.now():%Y%m%d_%H%M%S}_分析報告.txt"
            report_path.write_text(result.report_text, encoding="utf-8")
            result.outputs["報告"] = str(report_path)

    do_save = config.auto_save_history() if save_history is None else save_history
    if do_save:
        result.history_id = sheets.append_history(df, source_name)

    return result


def build_report_text(df: pd.DataFrame, source_name: str = "") -> str:
    """產生文字報告：有 LLM 就用 LLM，沒有就用統計摘要。"""
    total = len(df)
    type_counts = df["問題類型"].value_counts() if "問題類型" in df.columns else pd.Series(dtype=int)
    detail_counts = df["問題細項"].value_counts() if "問題細項" in df.columns else pd.Series(dtype=int)
    dept_counts = df["部門"].value_counts() if "部門" in df.columns else pd.Series(dtype=int)
    summary = review_summary(df)

    stat_lines = [
        f"【{source_name}】客訴分析報告　產出時間：{datetime.now():%Y-%m-%d %H:%M}",
        f"總件數：{total} 件　自動分類採用：{summary['auto']} 件（{summary['auto_rate']:.0%}）"
        f"　待人工複核：{summary['review']} 件",
    ]
    if summary.get("review_causes"):
        stat_lines.append(
            "需人工原因：" + "　".join(f"{k} {v} 件" for k, v in summary["review_causes"].items())
        )
    if summary.get("agreement"):
        stat_lines.append(
            "交叉驗證：" + "　".join(f"{k} {v} 件" for k, v in summary["agreement"].items())
        )
    stat_lines += ["", "問題類型分布："]
    for name, cnt in type_counts.head(8).items():
        stat_lines.append(f"  - {name}：{int(cnt)} 件（{int(cnt) / max(total, 1):.0%}）")
    stat_lines.append("")
    stat_lines.append("問題細項 TOP5：")
    for name, cnt in detail_counts.head(5).items():
        stat_lines.append(f"  - {name}：{int(cnt)} 件")
    if not dept_counts.empty:
        stat_lines.append("")
        stat_lines.append("責任部門分布：")
        for name, cnt in dept_counts.head(8).items():
            stat_lines.append(f"  - {name or '（未指定）'}：{int(cnt)} 件")
    stats_text = "\n".join(stat_lines)

    try:
        from .llm import complete_text

        prompt = (
            "你是 ECOCO 資源回收公司的客訴分析助理。請依下列統計資料，用繁體中文寫一份"
            "給主管閱讀的分析報告，包含：整體概況、主要問題與可能成因、各部門建議行動、"
            "下一步追蹤重點。語氣專業、條列清楚，控制在 4 段以內。\n\n"
            f"{stats_text}\n"
        )
        text = complete_text(prompt, max_tokens=1500)
        if text:
            return f"{stats_text}\n\n{'=' * 40}\nAI 分析與建議\n{'=' * 40}\n{text}"
    except Exception:
        pass
    return stats_text
