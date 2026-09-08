#!/usr/bin/env python3
from __future__ import annotations

"""把一次月更新的執行結果整理成 Markdown 檢討報告。

輸入：
  --summary        run_all_banks 的彙整 JSON（各家 _runner 含起訖時間、returncode）
  --update-output  update_credit_card_workbook 的 stdout JSON（含 timings_seconds、verification、percent_audit）
  --retry          （可多個）情境 B 補跑的單一銀行 JSON，用來統計重試次數
  --tokens         各階段 token / 成本快照 JSON（可省略）：
                   {"stages": [{"stage": "抓取", "input_tokens": .., "output_tokens": .., "cache_read_tokens": ..,
                                "cache_write_tokens": .., "cost_usd": ..}, ...]}
                   每筆為該階段結束時的累計值，報告會自行相減成各階段增量。
輸出：--output 指定的 .md（目錄不存在會建立），同時把重點印到 stdout。
"""

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

sys.dont_write_bytecode = True

SLOW_BANK_SECONDS = 60
SLOW_STAGE_SECONDS = 120
LONG_RUN_MINUTES = 20


def load_json(path: str | None) -> dict[str, Any]:
    if not path:
        return {}
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def seconds_between(start: str | None, end: str | None) -> float | None:
    try:
        return round((datetime.fromisoformat(end) - datetime.fromisoformat(start)).total_seconds(), 1)
    except Exception:
        return None


def bank_rows(summary: dict[str, Any], retries: list[dict[str, Any]]) -> list[dict[str, Any]]:
    retry_keys = {r.get("bank_key") for retry in retries for r in retry.get("results", [])}
    rows = []
    for r in summary.get("results", []):
        runner = r.get("_runner") or {}
        rows.append({
            "bank": r.get("bank_name") or r.get("bank_key"),
            "status": r.get("status"),
            "data_month": r.get("data_month"),
            "seconds": seconds_between(runner.get("started_at"), runner.get("finished_at")),
            "returncode": runner.get("returncode"),
            "retried": r.get("bank_key") in retry_keys,
            "error": (r.get("failure_reason") or "")[:120],
        })
    return rows


def token_deltas(tokens: dict[str, Any]) -> list[dict[str, Any]]:
    stages = tokens.get("stages") or []
    out = []
    prev = {"input_tokens": 0, "output_tokens": 0, "cache_read_tokens": 0, "cache_write_tokens": 0, "cost_usd": 0.0}
    for snap in stages:
        delta = {k: round((snap.get(k) or 0) - (prev.get(k) or 0), 4) for k in prev}
        delta["stage"] = snap.get("stage")
        out.append(delta)
        prev = {k: snap.get(k) or 0 for k in prev}
    return out


def fmt(v: Any) -> str:
    if v is None:
        return "-"
    if isinstance(v, float):
        return f"{v:,.4f}" if abs(v) < 1 else f"{v:,.1f}"
    if isinstance(v, int):
        return f"{v:,}"
    return str(v)


def build_report(summary: dict[str, Any], update: dict[str, Any], retries: list[dict[str, Any]], tokens: dict[str, Any], run_label: str) -> tuple[str, list[str]]:
    lines: list[str] = []
    findings: list[str] = []
    rows = bank_rows(summary, retries)
    timings = update.get("timings_seconds") or {}
    verification = update.get("verification") or {}
    audit = update.get("percent_audit") or {}
    warnings = list(update.get("warnings") or [])

    # ---- 摘要 ----
    ok = sum(1 for r in rows if r["status"] == "success")
    partial = sum(1 for r in rows if r["status"] == "partial_success")
    failed = sum(1 for r in rows if r["status"] == "failed")
    lines += [f"# 月更新檢討報告：{run_label}", "",
              f"- 產生時間：{datetime.now().isoformat(timespec='seconds')}",
              f"- 目標月：{update.get('target_month') or '-'}（YYYYMM {update.get('yyyymm') or '-'}），block {'新建' if update.get('block_created') else '既有'}，起始列 {update.get('block_start_row') or '-'}",
              f"- 銀行抓取：成功 {ok}／部分 {partial}／失敗 {failed}，缺少 {summary.get('missing_banks') or []}",
              f"- 讀回驗證：{verification.get('status') or '-'}；百分比稽核異常 {audit.get('anomaly_count', '-')} 格（掃 {audit.get('blocks_audited', '-')} 個 block）",
              f"- 回寫總耗時：{fmt(timings.get('total'))} 秒", ""]

    # ---- 抓取耗時 ----
    lines += ["## 各銀行抓取", "", "| 銀行 | 狀態 | 資料月 | 秒數 | 重試 | 錯誤 |", "|---|---|---|---|---|---|"]
    for r in sorted(rows, key=lambda x: -(x["seconds"] or 0)):
        lines.append(f"| {r['bank']} | {r['status']} | {r['data_month'] or '-'} | {fmt(r['seconds'])} | {'是' if r['retried'] else ''} | {r['error']} |")
    total_scrape = sum(r["seconds"] or 0 for r in rows)
    lines += ["", f"抓取合計約 {fmt(total_scrape)} 秒（循序）。", ""]
    for r in rows:
        if r["seconds"] and r["seconds"] > SLOW_BANK_SECONDS:
            findings.append(f"{r['bank']} 抓取花 {fmt(r['seconds'])} 秒，超過 {SLOW_BANK_SECONDS} 秒：檢查官網回應、重試次數或 timeout 設定。")
        if r["status"] != "success":
            findings.append(f"{r['bank']} 狀態 {r['status']}：{r['error'] or '無錯誤訊息'}。")
        if r["retried"]:
            findings.append(f"{r['bank']} 需要補跑一次才成功，若連續數月如此，優先修這支爬蟲。")
    months = {r["data_month"] for r in rows if r["data_month"]}
    if len(months) > 1:
        findings.append(f"各行資料月不一致 {sorted(months)}：落後的銀行由 backfill 之後補，屬正常，但 PR 要註明。")

    # ---- 回寫階段 ----
    lines += ["## 回寫各階段耗時", "", "| 階段 | 秒數 |", "|---|---|"]
    for stage, sec in timings.items():
        lines.append(f"| {stage} | {fmt(sec)} |")
        if stage != "total" and isinstance(sec, (int, float)) and sec > SLOW_STAGE_SECONDS:
            findings.append(f"階段 {stage} 花 {fmt(sec)} 秒，超過 {SLOW_STAGE_SECONDS} 秒：看該階段的 lookback 是否過大、是否重複下載金管會 ZIP。")
    lines.append("")
    for key, label in (("bank_bureau_backfill", "銀行局 backfill"), ("market_total", "市場總計"), ("jcic_avg_cards", "JCIC")):
        d = update.get(key) or {}
        if d:
            res = d.get("results") or []
            actions = sorted({str(r.get("action")) for r in res})
            lines.append(f"- {label}：子程序 {fmt(d.get('elapsed_seconds'))} 秒，處理 {len(res)} 個月份，動作 {actions}" + (f"，最新可用 {d.get('latest_available_month')}" if d.get("latest_available_month") else ""))
            for w in d.get("summary_warnings") or []:
                warnings.append(f"{label}: {w}")
            for r in res:
                if r.get("action") == "skip_fetch_failed":
                    findings.append(f"{label} 回補 {r.get('ad_yyyymm')} 失敗：{str(r.get('error'))[:100]}。若該月金管會檔已下架屬正常，否則要查。")
    lines.append("")

    # ---- 驗證與稽核 ----
    lines += ["## 驗證與稽核", ""]
    if verification:
        lines.append(f"- verification：{verification.get('status')}，驗證 {len(verification.get('verified_banks') or [])} 家，mismatch {len(verification.get('mismatches') or [])} 格，percent_anomalies {len(verification.get('percent_anomalies') or [])} 格")
        for m in (verification.get("mismatches") or [])[:5]:
            findings.append(f"讀回不符：{m}")
    if audit:
        lines.append(f"- percent_audit：{audit.get('anomaly_count')} 格異常")
        for a in (audit.get("anomalies") or [])[:5]:
            findings.append(f"百分比異常：{a.get('yyyymm')} {a.get('item')} {a.get('field')} = {a.get('value')}（{a.get('reason')}）")
    if warnings:
        lines += ["", "警告："] + [f"- {w}" for w in warnings]
    lines.append("")

    # ---- token ----
    deltas = token_deltas(tokens)
    if deltas:
        lines += ["## Token 與成本（各階段增量）", "", "| 階段 | 輸入 | 輸出 | 快取讀 | 快取寫 | 成本 USD |", "|---|---|---|---|---|---|"]
        rows_shown = [d for d in deltas[1:]] or deltas  # 第一筆是起點快照，增量為 0，不列
        for d in rows_shown:
            lines.append(f"| {d['stage']} | {fmt(d['input_tokens'])} | {fmt(d['output_tokens'])} | {fmt(d['cache_read_tokens'])} | {fmt(d['cache_write_tokens'])} | {fmt(d['cost_usd'])} |")
        total_out = sum(d["output_tokens"] for d in deltas)
        total_cost = sum(d["cost_usd"] for d in deltas)
        lines += ["", f"合計輸出 {fmt(total_out)} tokens，成本約 {fmt(total_cost)} USD。平台用量在回合結束後才更新，所以通常只有整次增量。", ""]
        if len(rows_shown) > 1:
            top = max(rows_shown, key=lambda d: d["cost_usd"])
            if top["cost_usd"] > 0:
                findings.append(f"成本最高的階段是「{top['stage']}」（{fmt(top['cost_usd'])} USD）。若它是抓取或判讀 JSON，優先減少把大型 JSON 印到對話裡的次數，改用 python 摘要。")
        if total_cost > 5:
            findings.append(f"整次成本 {fmt(total_cost)} USD 偏高：檢查是否有把整份 summary / update JSON 印進對話，或重跑了多次抓取。")
    else:
        lines += ["## Token 與成本", "", "未提供 token 快照（排程執行時由主 session 在各階段記錄 get_session 用量後傳入）。", ""]

    # ---- 檢討 ----
    total_min = (timings.get("total") or 0) / 60 + total_scrape / 60
    if total_min > LONG_RUN_MINUTES:
        findings.append(f"整體約 {total_min:.0f} 分鐘，超過 {LONG_RUN_MINUTES} 分鐘：抓取可考慮平行執行三組，或縮短 timeout。")
    lines += ["## 檢討與建議", ""]
    lines += [f"- {f}" for f in findings] if findings else ["- 本次沒有需要改進的環節。"]
    lines.append("")
    return "\n".join(lines), findings


def main() -> None:
    parser = argparse.ArgumentParser(description="產生月更新檢討報告（Markdown）")
    parser.add_argument("--summary", required=True)
    parser.add_argument("--update-output", required=True)
    parser.add_argument("--retry", action="append", default=[])
    parser.add_argument("--tokens", default="")
    parser.add_argument("--label", default="", help="報告標題用，例如 115年08月；未指定用 update 輸出的 target_month")
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    summary = load_json(args.summary)
    update = load_json(args.update_output)
    retries = [load_json(p) for p in args.retry]
    tokens = load_json(args.tokens)
    label = args.label or update.get("target_month") or "unknown"
    report, findings = build_report(summary, update, retries, tokens, label)
    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(report, encoding="utf-8")
    print(json.dumps({"report": str(out), "finding_count": len(findings), "findings": findings}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
