"""無人值守入口：不需要 Streamlit，可由 Render Cron 或 Windows 工作排程器呼叫。

用法：
    # 分析單一檔案，產出到 output/
    python -m automation.cli run --input 客訴清單.xlsx --out output --report

    # 讀 SOURCE_SHEET_ID 設定的 Google Sheet（只處理尚未分析過的新資料）
    python -m automation.cli run --from-sheet --out output --report

    # 監看資料夾裡所有還沒處理過的檔案
    python -m automation.cli watch --dir inbox --out output

    # 檢查設定與憑證是否就緒
    python -m automation.cli doctor

    # 政策改變後，依現行分類法重新標記歷史（產出供檢視，不覆蓋原檔）
    python -m automation.cli relabel --input history_reports/xxx.csv

    # 分類品質診斷：哪些細項分不開、合併能拿回多少準確率
    python -m automation.cli taxonomy-report --cut 2026-07-01

退出碼：0 正常；1 執行失敗；2 沒有資料可處理。
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

from . import config, pipeline, sheets

STATE_FILE = Path(".automation_state.json")
SUPPORTED_SUFFIXES = {".xlsx", ".xls", ".csv"}


def _log(msg: str) -> None:
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}", flush=True)


def _load_state() -> dict:
    if STATE_FILE.exists():
        try:
            return json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _save_state(state: dict) -> None:
    try:
        STATE_FILE.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        _log(f"警告：狀態檔寫入失敗（{exc}）")


def cmd_doctor(_args) -> int:
    creds = config.get_google_credentials()
    checks = [
        ("Google 憑證", bool(creds), (creds or {}).get("client_email", "未設定")),
        ("歷史紀錄試算表 HISTORY_SHEET_ID", bool(config.get_history_sheet_id()),
         config.get_history_sheet_id() or "未設定"),
        ("來源試算表 SOURCE_SHEET_ID", bool(config.get_source_sheet_id()),
         config.get_source_sheet_id() or "未設定（排程讀 Sheet 時才需要）"),
        ("Anthropic API Key", bool(config.get_anthropic_key()), "已設定" if config.get_anthropic_key() else "未設定"),
        ("OpenAI API Key", bool(config.get_openai_key()), "已設定" if config.get_openai_key() else "未設定"),
        ("L2 LLM 分類", config.use_llm_classifier(), config.llm_classifier_model()),
    ]
    for name, ok, detail in checks:
        _log(f"{'✓' if ok else '✗'} {name}：{detail}")

    if config.get_history_sheet_id():
        frames = sheets.read_history_frames(max_items=5)
        _log(f"歷史紀錄連線測試：讀到 {len(frames)} 份資料表")
        kb = pipeline.build_knowledge(frames) if frames else None
        if kb is not None:
            _log(f"知識庫試建：{kb.stats}")
    return 0


def _run_one(source, source_name: str, args, knowledge) -> int:
    try:
        result = pipeline.run(
            source,
            source_name=source_name,
            out_dir=args.out,
            # 未指定 --no-history 時交由設定（AUTO_SAVE_HISTORY）決定，
            # 不要用旗標的反向值硬蓋掉設定。
            save_history=False if args.no_history else None,
            make_report=args.report,
            knowledge=knowledge,
        )
    except Exception as exc:
        _log(f"失敗：{source_name}：{exc}")
        return 1
    _log("完成：\n" + result.describe())
    if args.fail_on_review and result.summary.get("review", 0) > 0:
        _log(f"注意：有 {result.summary['review']} 筆待人工複核")
    return 0


def cmd_history(args) -> int:
    """檢視與清空歷史紀錄。

    歷史紀錄存在三個地方，影響的東西不一樣：
      雲端索引分頁「歷史紀錄」   → 功能三的清單
      雲端資料分頁 history_*     → 功能三的預覽、功能四的儀表板
      本機 history_reports/      → 知識庫的主要學習來源

    預設只列出現況（不改任何東西）。清除動作一律要加 --yes，
    因為清掉本機歷史等於把知識庫歸零：實測自動採用率會從 95% 掉回 58%。
    """
    from pathlib import Path as _P

    ss = sheets.open_spreadsheet(config.get_history_sheet_id())
    tabs, data_tabs, rows, refs = [], [], [], []
    if ss is not None:
        try:
            tabs = [w.title for w in ss.worksheets()]
            data_tabs = [t for t in tabs if t.startswith("history_")]
            rows = ss.worksheet("歷史紀錄").get_all_values()[1:]
            refs = [r[4].split(":", 1)[1] for r in rows
                    if len(r) > 4 and r[4].startswith("sheet:")]
        except Exception as exc:
            _log(f"讀取雲端歷史失敗：{exc}")

    dangling = [r for r in rows
                if len(r) > 4 and r[4].startswith("sheet:")
                and r[4].split(":", 1)[1] not in data_tabs]
    orphans = [t for t in data_tabs if t not in refs]

    local_dir = _P(config.get_local_history_dir())
    local_files = ([f for f in local_dir.iterdir()
                    if f.suffix.lower() in (".xlsx", ".xls", ".csv")]
                   if local_dir.exists() else [])
    local_rows = 0
    for f in local_files:
        try:
            local_rows += len(pipeline.load_frame(f))
        except Exception:
            pass

    _log("── 雲端（Google Sheets）──")
    _log(f"   索引列 {len(rows)} 筆　可讀資料分頁 {len(refs) - len(dangling)} 份")
    _log(f"   斷鏈索引（指向已刪除的分頁）{len(dangling)} 筆")
    _log(f"   孤兒分頁（索引沒指到）{len(orphans)} 份")
    _log("── 本機（知識庫來源）──")
    _log(f"   {local_dir}／{len(local_files)} 個檔案　合計約 {local_rows} 列")

    todo = args.clean_dangling or args.purge_cloud or args.purge_local
    if not todo:
        _log("")
        _log("要清理請加參數（預設只檢視，不動任何東西）：")
        _log("   --clean-dangling   只清掉斷鏈索引與孤兒分頁（不會少任何資料）")
        _log("   --purge-cloud      清空雲端歷史（功能三／四會變空，知識庫仍在）")
        _log("   --purge-local      清空本機歷史（知識庫歸零，自動採用率會大幅下降）")
        _log("   加上 --yes 才會真的執行")
        return 0

    if not args.yes:
        _log("")
        _log("這是預演，沒有實際刪除。確認無誤請加 --yes 重跑。")

    # ── 只清斷鏈與孤兒：不會少任何真正的資料 ──
    if args.clean_dangling and ss is not None:
        _log(f"清理斷鏈索引 {len(dangling)} 筆、孤兒分頁 {len(orphans)} 份")
        if args.yes:
            try:
                ws = ss.worksheet("歷史紀錄")
                ids = {r[0] for r in dangling}
                # 由下往上刪，才不會因為列號位移而刪錯
                for i, r in reversed(list(enumerate(ws.get_all_values()[1:], start=2))):
                    if r and r[0] in ids:
                        ws.delete_rows(i)
                for t in orphans:
                    ss.del_worksheet(ss.worksheet(t))
                _log("完成")
            except Exception as exc:
                _log(f"清理失敗：{exc}")
                return 1

    # ── 清空雲端 ──
    if args.purge_cloud and ss is not None:
        _log(f"清空雲端：刪除 {len(data_tabs)} 份資料分頁、清空索引 {len(rows)} 列")
        if args.yes:
            try:
                for t in data_tabs:
                    ss.del_worksheet(ss.worksheet(t))
                ws = ss.worksheet("歷史紀錄")
                if len(rows):
                    ws.delete_rows(2, len(rows) + 1)
                _log("完成")
            except Exception as exc:
                _log(f"清空雲端失敗：{exc}")
                return 1

    # ── 清空本機 ──
    if args.purge_local:
        _log(f"清空本機：刪除 {len(local_files)} 個檔案（約 {local_rows} 列）")
        _log("注意：知識庫會歸零，自動採用率會從約 95% 掉到約 58%")
        if args.yes:
            backup = local_dir.parent / f"{local_dir.name}_已清空備份"
            backup.mkdir(exist_ok=True)
            for f in local_files:
                try:
                    f.rename(backup / f.name)
                except Exception as exc:
                    _log(f"搬移 {f.name} 失敗：{exc}")
            _log(f"完成。原檔已搬到 {backup}（不是直接刪除，反悔可以搬回來）")

    return 0


def cmd_run(args) -> int:
    knowledge = pipeline.build_knowledge()
    if knowledge is not None:
        _log(f"知識庫就緒：{knowledge.stats}")
    else:
        _log("未取得歷史知識庫，使用內建規則分類")

    if args.from_sheet:
        df = sheets.read_source_frame()
        if df is None or df.empty:
            _log("來源試算表沒有資料或讀取失敗")
            return 2
        state = _load_state()
        processed = int(state.get("source_sheet_rows", 0))
        if args.only_new and len(df) <= processed:
            _log(f"沒有新資料（來源 {len(df)} 筆，已處理 {processed} 筆）")
            return 2
        if args.only_new and processed:
            df = df.iloc[processed:]
            _log(f"只處理新增的 {len(df)} 筆")
        rc = _run_one(df, args.name or "Google Sheet 來源", args, knowledge)
        if rc == 0 and args.only_new:
            state["source_sheet_rows"] = processed + len(df)
            state["last_run"] = datetime.now().isoformat(timespec="seconds")
            _save_state(state)
        return rc

    if not args.input:
        _log("請指定 --input 檔案，或用 --from-sheet 讀取來源試算表")
        return 2
    path = Path(args.input)
    if not path.exists():
        _log(f"找不到檔案：{path}")
        return 2
    return _run_one(path, args.name or path.name, args, knowledge)


def cmd_watch(args) -> int:
    folder = Path(args.dir)
    if not folder.exists():
        _log(f"找不到資料夾：{folder}")
        return 2
    state = _load_state()
    done = set(state.get("processed_files", []))
    targets = [
        p for p in sorted(folder.iterdir())
        if p.suffix.lower() in SUPPORTED_SUFFIXES and str(p.resolve()) not in done
    ]
    if not targets:
        _log("沒有待處理的新檔案")
        return 2

    knowledge = pipeline.build_knowledge()
    failed = 0
    for p in targets:
        _log(f"處理：{p.name}")
        if _run_one(p, p.name, args, knowledge) == 0:
            done.add(str(p.resolve()))
        else:
            failed += 1
    state["processed_files"] = sorted(done)
    state["last_run"] = datetime.now().isoformat(timespec="seconds")
    _save_state(state)
    return 1 if failed else 0


def cmd_relabel(args) -> int:
    """依現行分類法與政策，重新標記歷史資料（產出供人工檢視，不覆蓋原檔）。

    用途：政策改變時（例如三種滿艙改為依內容分開），
    舊標記會與新政策衝突，知識庫若繼續學舊標記就會教出舊行為。
    """
    import pandas as pd

    from .rules import full_bin_detail
    from .taxonomy import POLICY_DETAILS, RETIRED_TOPICS

    merge_map: dict[str, str] = {}
    if getattr(args, "merge_map", ""):
        try:
            groups = json.loads(Path(args.merge_map).read_text(encoding="utf-8"))
            merge_map = {d: g for g, ds in groups.items() for d in ds}
            _log(f"套用合併對照表：{len(groups)} 組、共 {len(merge_map)} 個細項")
        except Exception as exc:
            _log(f"合併對照表讀取失敗：{exc}")
            return 1

    src = Path(args.input)
    if not src.exists():
        _log(f"找不到檔案：{src}")
        return 2
    try:
        df = pipeline.load_frame(src)
    except Exception as exc:
        _log(f"讀取失敗：{exc}")
        return 1

    if "問題類型" not in df.columns or "問題細項" not in df.columns:
        _log("來源檔缺少「問題類型」或「問題細項」欄，無法重新標記")
        return 2

    from .knowledge import CONTENT_HINTS, SUBJECT_HINTS, _pick_col

    subj = _pick_col(df, SUBJECT_HINTS)
    cont = _pick_col(df, CONTENT_HINTS)
    out = df.copy()
    changes = []
    for i, row in df.iterrows():
        old_t = str(row["問題類型"]).strip()
        old_d = str(row["問題細項"]).strip()
        new_t, new_d = RETIRED_TOPICS.get(old_t, old_t), old_d
        # 已被人工認定屬於政策細項家族者，依內容重新指派
        if old_d in POLICY_DETAILS or "滿" in old_d:
            text = " ".join(str(row.get(c, "")) for c in (subj, cont) if c)
            new_d = full_bin_detail(text)
            new_t = "機台問題類型"
        elif new_d in merge_map:
            new_d = merge_map[new_d]
        if (new_t, new_d) != (old_t, old_d):
            out.at[i, "問題類型"] = new_t
            out.at[i, "問題細項"] = new_d
            changes.append({
                "列號": i + 2, "問題主旨": str(row.get(subj, ""))[:40],
                "原類型": old_t, "原細項": old_d, "新類型": new_t, "新細項": new_d,
            })

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    new_path = out_dir / f"{src.stem}_重新標記{src.suffix or '.csv'}"
    diff_path = out_dir / f"{src.stem}_變更清單.csv"
    if new_path.suffix.lower() in (".xlsx", ".xls"):
        out.to_excel(new_path, index=False)
    else:
        out.to_csv(new_path, index=False, encoding="utf-8-sig")
    pd.DataFrame(changes).to_csv(diff_path, index=False, encoding="utf-8-sig")

    _log(f"共 {len(df)} 筆，建議改標 {len(changes)} 筆（{len(changes)/max(len(df),1):.1%}）")
    counts: dict[str, int] = {}
    for c in changes:
        key = f"{c['原細項']} → {c['新細項']}"
        counts[key] = counts.get(key, 0) + 1
    for key, n in sorted(counts.items(), key=lambda x: -x[1])[:10]:
        _log(f"  {key}：{n} 筆")
    _log(f"產出：{new_path}")
    _log(f"變更清單：{diff_path}")
    _log("請人工檢視變更清單，確認後再用重新標記的檔案取代原檔。")
    return 0


def cmd_taxonomy_report(args) -> int:
    """分類品質診斷：哪些細項分不開、合併能拿回多少準確率。

    只做量測與模擬，不會改動任何設定或資料。
    """
    import json as _json

    import pandas as pd

    from . import diagnostics

    if args.input:
        src = Path(args.input)
        if not src.exists():
            _log(f"找不到檔案：{src}")
            return 2
        frames = [pipeline.load_frame(src)]
        label = src.name
    else:
        frames = pipeline.read_local_history_frames()
        try:
            frames += sheets.read_history_frames()
        except Exception:
            pass
        label = "全部歷史紀錄"

    rows = diagnostics.load_labeled(frames)
    if len(rows) < 100:
        _log(f"可用的已標記資料只有 {len(rows)} 筆，太少，無法做有意義的診斷")
        return 2

    cut = pd.Timestamp(args.cut) if args.cut else rows[int(len(rows) * 0.8)].date
    _log(f"資料來源：{label}　可評估 {len(rows)} 筆"
         f"（{rows[0].date.date()} ~ {rows[-1].date.date()}）")
    _log(f"時間切分點：{cut.date()}（之前建知識庫，之後當測試）")

    base = diagnostics.evaluate(rows, cut)
    _log("現況：" + chr(10) + base.describe())

    sep = diagnostics.separability(base, min_count=args.min_count)
    _log(f"可辨識度最低的細項（測試期出現 ≥{args.min_count} 次）：")
    for r in sep[:10]:
        _log(f"    {r['可辨識度']:.0%}  {r['問題細項']}（{r['測試筆數']} 筆）")

    pairs = diagnostics.confusion_pairs(base)
    _log("最常混淆的細項配對：")
    for r in pairs[:8]:
        mark = "  ← 雙向混淆" if r["雙向混淆"] else ""
        _log(f"    {r['合計']:3} 次  {r['人工標記']} ↔ {r['系統判斷']}{mark}")

    if args.merge_map:
        try:
            groups = _json.loads(Path(args.merge_map).read_text(encoding="utf-8"))
            _log(f"改用指定的合併方案：{Path(args.merge_map).name}")
        except Exception as exc:
            _log(f"合併方案讀取失敗：{exc}")
            return 1
    else:
        groups = diagnostics.suggest_merges(base, min_pair=args.min_pair)
    if groups:
        _log("依雙向混淆自動建議的合併群組：")
        for g, ds in groups.items():
            _log(f"    {g}：{'、'.join(ds)}")
        merged = diagnostics.simulate_merge(rows, cut, groups)
        _log("合併後模擬：" + chr(10) + merged.describe())
        _log(f"效益：細項正確率 {base.detail_acc:.1%} → {merged.detail_acc:.1%}"
             f"　完全自動採用 {base.auto_rate:.1%} → {merged.auto_rate:.1%}"
             f"　需完整人工 {base.full_manual_rate:.1%} → {merged.full_manual_rate:.1%}")
    else:
        _log("沒有偵測到明顯的雙向混淆，暫時不需要合併細項")

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    pd.DataFrame(sep).to_csv(out_dir / f"{stamp}_細項可辨識度.csv",
                             index=False, encoding="utf-8-sig")
    pd.DataFrame(pairs).to_csv(out_dir / f"{stamp}_混淆配對.csv",
                               index=False, encoding="utf-8-sig")
    if groups:
        (out_dir / f"{stamp}_合併建議.json").write_text(
            _json.dumps(groups, ensure_ascii=False, indent=2), encoding="utf-8")
    _log(f"報表已輸出到 {out_dir}")
    _log("合併是分類法的業務決策：確認群組後，用 relabel --merge-map 套用到歷史資料。")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="automation.cli",
        description="ECOCO 客訴分析自動化排程入口",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    def add_common(sp):
        sp.add_argument("--out", default="output", help="產出資料夾（預設 output）")
        sp.add_argument("--report", action="store_true", help="同時產生文字分析報告")
        sp.add_argument("--no-history", action="store_true", help="不寫入雲端歷史紀錄")
        sp.add_argument("--name", default="", help="這批資料的名稱（顯示在歷史紀錄）")
        sp.add_argument("--fail-on-review", action="store_true",
                        help="有待複核資料時在日誌中特別提示")

    sp_run = sub.add_parser("run", help="分析單一來源")
    sp_run.add_argument("--input", help="來源檔案（xlsx / xls / csv）")
    sp_run.add_argument("--from-sheet", action="store_true",
                        help="改讀 SOURCE_SHEET_ID 指定的 Google Sheet")
    sp_run.add_argument("--only-new", action="store_true",
                        help="搭配 --from-sheet：只處理上次之後新增的列")
    add_common(sp_run)
    sp_run.set_defaults(func=cmd_run)

    sp_watch = sub.add_parser("watch", help="處理資料夾內所有未處理過的檔案")
    sp_watch.add_argument("--dir", default="inbox", help="監看的資料夾（預設 inbox）")
    add_common(sp_watch)
    sp_watch.set_defaults(func=cmd_watch)

    sp_rel = sub.add_parser("relabel", help="依現行分類法與政策重新標記歷史資料（不覆蓋原檔）")
    sp_rel.add_argument("--input", required=True, help="要重新標記的歷史檔（csv / xlsx）")
    sp_rel.add_argument("--out", default="output/relabel_review", help="產出資料夾")
    sp_rel.add_argument("--merge-map", default="",
                        help="細項合併對照表 JSON（taxonomy-report 產出的合併建議）")
    sp_rel.set_defaults(func=cmd_relabel)

    sp_tax = sub.add_parser("taxonomy-report",
                            help="分類品質診斷：哪些細項分不開、合併能拿回多少準確率")
    sp_tax.add_argument("--input", default="", help="指定歷史檔；留空則用全部歷史紀錄")
    sp_tax.add_argument("--cut", default="", help="時間切分點（YYYY-MM-DD），留空取後 20%% 當測試")
    sp_tax.add_argument("--out", default="output/taxonomy_report", help="報表輸出資料夾")
    sp_tax.add_argument("--min-count", type=int, default=10, help="細項至少出現幾次才列入可辨識度表")
    sp_tax.add_argument("--min-pair", type=int, default=6, help="雙向混淆達幾次才建議合併")
    sp_tax.add_argument("--merge-map", default="",
                        help="用指定的合併方案 JSON 做模擬，取代系統的自動建議")
    sp_tax.set_defaults(func=cmd_taxonomy_report)

    sp_doc = sub.add_parser("doctor", help="檢查設定與憑證")
    sp_doc.set_defaults(func=cmd_doctor)

    sp_hist = sub.add_parser("history", help="檢視或清空歷史紀錄（預設只檢視）")
    sp_hist.add_argument("--clean-dangling", action="store_true",
                         help="清掉斷鏈索引與孤兒分頁（不會少任何資料）")
    sp_hist.add_argument("--purge-cloud", action="store_true",
                         help="清空雲端歷史紀錄")
    sp_hist.add_argument("--purge-local", action="store_true",
                         help="清空本機歷史紀錄（知識庫會歸零）")
    sp_hist.add_argument("--yes", action="store_true",
                         help="真的執行；沒有這個參數只會預演")
    sp_hist.set_defaults(func=cmd_history)

    return parser


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
