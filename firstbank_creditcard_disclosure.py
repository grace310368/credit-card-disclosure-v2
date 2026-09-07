#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import sys
import time
import urllib.request
import urllib.error
import ssl
import socket
from html.parser import HTMLParser
from typing import List, Dict, Any, Optional


def configure_utf8_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            if hasattr(stream, "reconfigure"):
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


configure_utf8_stdio()


# =============== Config ===============
BANK_KEY = "firstbank"
BANK_NAME = "第一銀行"
ENTRY = "https://ccard.firstbank.com.tw/cmsweb/home/creditcardreport"

# 第一銀行信用卡重要業務資訊頁面，金額欄位通常以「仟元」揭露。
SOURCE_AMOUNT_UNIT = "仟元"
STANDARD_CARD_UNIT = "張"

LABEL_MAP = {
    "流通卡數": "circulating_cards",
    "流通卡數": "circulating_cards",
    "有效卡數": "valid_cards",
    "有效卡數": "valid_cards",
    "當月發卡數": "new_cards_this_month",
    "當月發卡數": "new_cards_this_month",
    "當月停卡數": "cancelled_cards_this_month",
    "當月停卡數": "cancelled_cards_this_month",
    "當月簽帳金額": "signed_amount_thousand",
    "循環信用餘額": "revolving_balance_thousand",
    "未到期分期付款餘額": "installment_balance_not_yet_due_thousand",
    "當月預借現金金額": "cash_advance_amount_thousand",
    "逾期三個月以上帳款占應收帳款餘額(含催收款)之比率": "overdue_3m_ratio_percent",
    "逾期三個月以上帳款占應收帳款餘額(含催收款)之比率(%)": "overdue_3m_ratio_percent",
    "逾期三個月以上帳款占應收帳款餘額（含催收款）之比率": "overdue_3m_ratio_percent",
    "逾期三個月以上帳款占應收帳款餘額（含催收款）之比率(%)": "overdue_3m_ratio_percent",
    "逾期三個月以上帳款佔應收帳款餘額(含催收款)之比率": "overdue_3m_ratio_percent",
    "逾期三個月以上帳款佔應收帳款餘額(含催收款)之比率(%)": "overdue_3m_ratio_percent",
    "逾期三個月以上帳款佔應收帳款餘額（含催收款）之比率": "overdue_3m_ratio_percent",
    "逾期三個月以上帳款佔應收帳款餘額（含催收款）之比率(%)": "overdue_3m_ratio_percent",
    "逾期六個月以上帳款占應收帳款餘額(含催收款)之比率": "overdue_6m_ratio_percent",
    "逾期六個月以上帳款占應收帳款餘額(含催收款)之比率(%)": "overdue_6m_ratio_percent",
    "逾期六個月以上帳款占應收帳款餘額（含催收款）之比率": "overdue_6m_ratio_percent",
    "逾期六個月以上帳款占應收帳款餘額（含催收款）之比率(%)": "overdue_6m_ratio_percent",
    "逾期六個月以上帳款佔應收帳款餘額(含催收款)之比率": "overdue_6m_ratio_percent",
    "逾期六個月以上帳款佔應收帳款餘額(含催收款)之比率(%)": "overdue_6m_ratio_percent",
    "逾期六個月以上帳款佔應收帳款餘額（含催收款）之比率": "overdue_6m_ratio_percent",
    "逾期六個月以上帳款佔應收帳款餘額（含催收款）之比率(%)": "overdue_6m_ratio_percent",
    "備抵呆帳提足率": "allowance_coverage_ratio_percent",
    "當月轉銷呆帳金額": "charge_off_amount_this_month_thousand",
    "當年度累計轉銷呆帳金額": "charge_off_amount_ytd_thousand",
    "當年度轉銷呆帳金額累計至資料月份": "charge_off_amount_ytd_thousand",
}
REQUIRED = list(LABEL_MAP.values())
YM6_RE = re.compile(r"\b(\d{6})\b")
# 同時保留 canonical key 與現行 metrics key，方便主控腳本或下游程式讀取。
METRIC_UNITS = {
    "circulating_cards": STANDARD_CARD_UNIT,
    "valid_cards": STANDARD_CARD_UNIT,
    "new_cards_this_month": STANDARD_CARD_UNIT,
    "cancelled_cards_this_month": STANDARD_CARD_UNIT,
    "signed_amount": SOURCE_AMOUNT_UNIT,
    "revolving_balance": SOURCE_AMOUNT_UNIT,
    "installment_balance_not_yet_due": SOURCE_AMOUNT_UNIT,
    "cash_advance_amount": SOURCE_AMOUNT_UNIT,
    "charge_off_amount_this_month": SOURCE_AMOUNT_UNIT,
    "charge_off_amount_ytd": SOURCE_AMOUNT_UNIT,
    "signed_amount_thousand": SOURCE_AMOUNT_UNIT,
    "revolving_balance_thousand": SOURCE_AMOUNT_UNIT,
    "installment_balance_not_yet_due_thousand": SOURCE_AMOUNT_UNIT,
    "cash_advance_amount_thousand": SOURCE_AMOUNT_UNIT,
    "charge_off_amount_this_month_thousand": SOURCE_AMOUNT_UNIT,
    "charge_off_amount_ytd_thousand": SOURCE_AMOUNT_UNIT,
    "overdue_3m_ratio_percent": "%",
    "overdue_6m_ratio_percent": "%",
    "allowance_coverage_ratio_percent": "%",
}


# =============== Content sanity checks ===============
SUSPECT_MARKERS = [
    "captcha", "cloudflare", "attention required", "access denied",
    "request blocked", "forbidden", "not authorized",
    "robot", "verify you are", "驗證", "人機", "機器人",
    "請啟用javascript", "啟用 javascript", "啟用 javascript", "啟用 JavaScript",
    "安全性原則", "防火牆", "blocked", "deny", "拒絕存取",
    "proxy", "certificate", "憑證",
]


def _normalize_text_for_check(s: str) -> str:
    return s.replace("\ufeff", "").lower()


class FetchError(RuntimeError):
    pass


def validate_html_or_raise(html: str, steps: List[str], used_insecure: bool) -> None:
    """
    檢查抓回來的 HTML 是否「看起來像目標頁面」。
    used_insecure=True 時更嚴格（因為 CERT_NONE 風險較高）。
    """
    if html is None:
        raise FetchError("content_invalid: html is None")

    if len(html) < 800:
        steps.append(f"VALIDATE fail: html_too_short len={len(html)}")
        raise FetchError(f"content_invalid: html_too_short len={len(html)}")

    text = _normalize_text_for_check(html)

    hit = [m for m in SUSPECT_MARKERS if m in text]
    if hit:
        steps.append(f"VALIDATE warn: suspect_markers={','.join(hit[:5])}")
        if used_insecure:
            raise FetchError(f"content_invalid: suspect_markers={','.join(hit[:5])}")

    must_any = [
        "firstbank", "ccard.firstbank.com.tw",
        "信用卡", "credit card", "creditcard",
        "重要業務資訊", "揭露", "report"
    ]
    if not any(k.lower() in text for k in must_any):
        steps.append("VALIDATE fail: missing_basic_identity_keywords")
        raise FetchError("content_invalid: missing_basic_identity_keywords")

    labels_found = sum(1 for lab in LABEL_MAP.keys() if lab.replace(" ", "").lower() in text)
    steps.append(f"VALIDATE labels_found={labels_found} insecure={used_insecure}")

    if used_insecure:
        if labels_found < 2:
            raise FetchError(f"content_invalid: too_few_labels_found={labels_found} (insecure)")
    else:
        if labels_found < 1:
            raise FetchError(f"content_invalid: too_few_labels_found={labels_found}")

    if not YM6_RE.search(text):
        steps.append("VALIDATE warn: ym6_not_found")

    steps.append("VALIDATE ok")


# =============== Result Schema ===============
def build_result(
    bank_name: str,
    status: str,
    entry_url: str,
    month: str,
    report_url: str,
    metrics: Dict[str, str],
    message: str,
    errors: List[Dict[str, str]],
    steps: List[str],
) -> Dict[str, Any]:
    return {
        "bank": bank_name,
        "bank_key": BANK_KEY,
        "status": status,  # success / partial_success / failed
        "entry": entry_url,
        "report_url": report_url or "",
        "month": month or "",
        "metric_units": dict(METRIC_UNITS),
        "metrics": metrics or {},
        "message": message or "",
        "errors": errors or [],
        "steps": steps or [],
        "ts": int(time.time()),
    }


# =============== TLS / Fetch (three-stage) ===============
def _make_ssl_context(
    steps: List[str],
    ca_bundle: str = "",
    mode: str = "verify",     # verify | relax | insecure
) -> ssl.SSLContext:
    """
    mode:
      - verify  : normal certificate + hostname verification
      - relax   : verify but clears VERIFY_X509_STRICT if available
      - insecure: CERT_NONE (no verification) (high risk)
    """
    if mode == "insecure":
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        steps.append("TLS mode=insecure verify=CERT_NONE")
        return ctx

    if ca_bundle:
        ctx = ssl.create_default_context(cafile=ca_bundle)
        steps.append(f"TLS mode={mode} cafile={ca_bundle}")
    else:
        ctx = ssl.create_default_context()
        steps.append(f"TLS mode={mode} cafile=system_default")

    try:
        ctx.minimum_version = ssl.TLSVersion.TLSv1_2
        steps.append("TLS min_version=1.2")
    except Exception:
        pass

    if mode == "relax":
        # Clear strict checks; still verifies chain + hostname
        try:
            strict_flag = getattr(ssl, "VERIFY_X509_STRICT", None)
            if strict_flag is not None:
                ctx.verify_flags &= ~strict_flag
                steps.append("TLS relax_verify=ON (VERIFY_X509_STRICT cleared)")
            else:
                steps.append("TLS relax_verify=ON (VERIFY_X509_STRICT not available)")
        except Exception as e:
            steps.append(f"TLS relax_verify=FAILED ({e})")

    return ctx


def _attempt_fetch(
    url: str,
    steps: List[str],
    ctx: ssl.SSLContext,
    timeout: int,
    retries: int,
) -> str:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/122.0.0.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
        "Connection": "close",
    }

    last_err: Optional[Exception] = None

    for i in range(retries + 1):
        try:
            if i > 0:
                backoff = min(6.0, 0.8 * (2 ** (i - 1)))
                steps.append(f"FETCH retry={i} backoff={backoff:.1f}s")
                time.sleep(backoff)

            req = urllib.request.Request(url, headers=headers, method="GET")
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
                raw = resp.read()

                ctype = resp.headers.get("Content-Type", "")
                m = re.search(r"charset=([^\s;]+)", ctype, re.I)
                enc = (m.group(1).strip() if m else "").lower()

                for cand in [enc, "utf-8", "utf-8-sig", "big5", "cp950"]:
                    if not cand:
                        continue
                    try:
                        html = raw.decode(cand, errors="strict")
                        steps.append(f"FETCH ok decode={cand} bytes={len(raw)}")
                        return html
                    except Exception:
                        continue

                steps.append(f"FETCH ok decode=fallback bytes={len(raw)}")
                return raw.decode("utf-8", errors="replace")

        except urllib.error.HTTPError as e:
            last_err = e
            steps.append(f"FETCH HTTPError code={e.code} reason={getattr(e, 'reason', '')}")
        except ssl.SSLError as e:
            last_err = e
            steps.append(f"FETCH SSLError {e}")
        except (socket.timeout, TimeoutError) as e:
            last_err = e
            steps.append(f"FETCH Timeout {e}")
        except urllib.error.URLError as e:
            last_err = e
            steps.append(f"FETCH URLError reason={getattr(e, 'reason', '')}")
        except Exception as e:
            last_err = e
            steps.append(f"FETCH Exception {type(e).__name__}: {e}")

    raise FetchError(f"fetch_failed url={url} err={last_err}")


def safe_fetch_three_stage(
    url: str,
    steps: List[str],
    timeout: int = 20,
    retries: int = 3,
    ca_bundle: str = "",
    allow_insecure: bool = False,
) -> str:
    """
    Three-stage strategy:
      1) verify
      2) relax-verify
      3) insecure (only if allow_insecure=True)

    Note:
      - Stage 3 is intentionally gated to avoid silent downgrade.
    """
    # Stage 1: verify
    try:
        ctx1 = _make_ssl_context(steps, ca_bundle=ca_bundle, mode="verify")
        return _attempt_fetch(url, steps, ctx1, timeout, retries)
    except FetchError as e1:
        msg1 = str(e1)
        verify_failed = ("CERTIFICATE_VERIFY_FAILED" in msg1) or ("certificate verify failed" in msg1)
        steps.append(f"STAGE1 fail verify_failed={verify_failed}")
        # Stage 2: relax
        try:
            ctx2 = _make_ssl_context(steps, ca_bundle=ca_bundle, mode="relax")
            return _attempt_fetch(url, steps, ctx2, timeout, retries)
        except FetchError as e2:
            msg2 = str(e2)
            verify_failed2 = ("CERTIFICATE_VERIFY_FAILED" in msg2) or ("certificate verify failed" in msg2)
            steps.append(f"STAGE2 fail verify_failed={verify_failed2}")

            # Stage 3: insecure (gated)
            if allow_insecure:
                steps.append("STAGE3 insecure=ON (attempt CERT_NONE)")
                ctx3 = _make_ssl_context(steps, ca_bundle="", mode="insecure")
                # Stage3 不需要太多 retry，避免長時間掛住
                return _attempt_fetch(url, steps, ctx3, timeout, max(1, min(2, retries)))

            # If not allowed, raise the stage2 error (more recent)
            raise


# =============== Normalize Helpers ===============
def normalize_value_text(raw: str) -> Dict[str, Any]:
    if raw is None:
        return {"ok": False, "normalized": ""}

    s = str(raw).strip()
    if not s:
        return {"ok": False, "normalized": ""}

    s = s.replace("，", ",").replace(" ", "")
    s = re.sub(r"(千元|仟元|元|萬元|億元|億|萬|％|%)", "", s)

    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1]

    s = re.sub(r"[^0-9,\.\-]", "", s)
    s = s.replace("--", "-")
    s = s.replace(",", "")

    if not re.fullmatch(r"-?\d+(\.\d+)?", s):
        return {"ok": False, "normalized": ""}

    if neg and not s.startswith("-"):
        s = "-" + s

    return {"ok": True, "normalized": s}


def normalize_month_flexible(x: str) -> str:
    if not x:
        return ""

    s = str(x).strip()

    m = re.fullmatch(r"(\d{6})", s)
    if m:
        roc_y = int(m.group(1)[:3])
        mm = int(m.group(1)[3:])
        if 1 <= mm <= 12:
            return f"{roc_y}年{mm:02d}月"

    m = re.fullmatch(r"(\d{4})(\d{2})", s)
    if m:
        y, mm = int(m.group(1)), int(m.group(2))
        if 1 <= mm <= 12:
            roc_y = y - 1911 if y >= 1911 else y
            return f"{roc_y}年{mm:02d}月"

    m = re.search(r"(\d{2,3})\s*年\s*(\d{1,2})\s*月", s)
    if m:
        roc_y, mm = int(m.group(1)), int(m.group(2))
        if 1 <= mm <= 12:
            return f"{roc_y}年{mm:02d}月"

    m = re.fullmatch(r"(\d{2,4})[\/\-](\d{1,2})", s)
    if m:
        y, mm = int(m.group(1)), int(m.group(2))
        if 1 <= mm <= 12:
            roc_y = y - 1911 if y >= 1911 else y
            return f"{roc_y}年{mm:02d}月"

    m = YM6_RE.search(s)
    if m:
        return normalize_month_flexible(m.group(1))

    return ""


# =============== HTML Parsers ===============
class _TableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.in_tr = False
        self.in_cell = False
        self._cell_buf: List[str] = []
        self._row: List[str] = []
        self.rows: List[List[str]] = []

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == "tr":
            self.in_tr = True
            self._row = []
        elif tag in ("td", "th") and self.in_tr:
            self.in_cell = True
            self._cell_buf = []

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in ("td", "th") and self.in_cell:
            txt = "".join(self._cell_buf)
            txt = re.sub(r"\s+", " ", txt).strip()
            self._row.append(txt)
            self.in_cell = False
        elif tag == "tr" and self.in_tr:
            if self._row:
                self.rows.append(self._row)
            self.in_tr = False

    def handle_data(self, data):
        if self.in_cell:
            self._cell_buf.append(data)


def parse_html_table_rows(html: str) -> List[List[str]]:
    p = _TableParser()
    p.feed(html)
    return p.rows


class _Stripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts: List[str] = []

    def handle_data(self, data):
        if data:
            self.parts.append(data)

    def get_text(self) -> str:
        return re.sub(r"\s+", " ", "".join(self.parts)).strip()


def strip_tags(html: str) -> str:
    sp = _Stripper()
    sp.feed(html)
    return sp.get_text()


# =============== CLI ===============
def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="FirstBank credit card disclosure extractor")
    p.add_argument("--month", default="", help="Optional target month override, e.g. 115年04月")
    p.add_argument("--debug", action="store_true", help="Verbose steps")

    p.add_argument("--retries", type=int, default=3, help="Fetch retries (default: 3)")
    p.add_argument("--timeout", type=int, default=20, help="Fetch timeout seconds (default: 20)")
    p.add_argument("--ca-bundle", default="", help="Path to CA bundle PEM file for TLS verification")

    # IMPORTANT: gated insecure
    p.add_argument(
        "--allow-insecure",
        action="store_true",
        help="Allow stage3 insecure (CERT_NONE). High risk; not recommended.",
    )
    return p.parse_args()


# =============== Main Extractor ===============
def main() -> int:
    args = parse_args()
    steps: list[str] = []
    debug = args.debug

    def log(msg: str, important: bool = False) -> None:
        if important or debug:
            steps.append(msg)

    def put(metrics: dict[str, str], key: str, raw: str, label: str) -> None:
        nv = normalize_value_text(raw)
        if nv.get("ok"):
            metrics[key] = nv["normalized"]
            log(f"PARSE {label}:{key}={nv['normalized']}")
        else:
            log(f"WARN parse_failed {label}:{key} raw={raw}", important=True)

    def extract_table(html: str) -> tuple[str, dict[str, str]]:
        rows = parse_html_table_rows(html)
        ym_raw, metrics = "", {}
        for cells in rows:
            if len(cells) < 3:
                continue
            ym = cells[0].strip()
            label = cells[1].strip().replace(" ", "")
            value = cells[2].strip()

            if not ym_raw and ym.isdigit() and len(ym) == 6:
                ym_raw = ym

            for ktxt, mkey in LABEL_MAP.items():
                if ktxt in label:
                    put(metrics, mkey, value, ktxt)
                    break

        month = normalize_month_flexible(ym_raw) if ym_raw else ""
        if month:
            log(f"MONTH {ym_raw}->{month}")
        return month, metrics

    def extract_text(html: str) -> tuple[str, dict[str, str]]:
        text = strip_tags(html)
        m = YM6_RE.search(text)
        month = normalize_month_flexible(m.group(1)) if m else ""
        if month and m:
            log(f"MONTH_FALLBACK {m.group(1)}->{month}")

        patterns = {
            lab: re.compile(re.escape(lab) + r"(?:\s*（[^）]*）)?\s*([-]?\d[\d,]*\.?\d*)")
            for lab in LABEL_MAP
        }
        metrics: dict[str, str] = {}
        for lab, key in LABEL_MAP.items():
            mm = patterns[lab].search(text)
            if mm:
                put(metrics, key, mm.group(1), lab)
        return month, metrics

    try:
        html = safe_fetch_three_stage(
            ENTRY,
            steps,
            timeout=args.timeout,
            retries=args.retries,
            ca_bundle=args.ca_bundle,
            allow_insecure=args.allow_insecure,
        )
        log("FETCH ok", important=True)

        used_insecure = any("TLS mode=insecure" in s for s in steps)
        validate_html_or_raise(html, steps, used_insecure)

        month, metrics = extract_table(html)
        if not metrics:
            log("FALLBACK text_parse", important=True)
            month, metrics = extract_text(html)

        if not month and args.month:
            month = args.month
            log(f"MONTH_OVERRIDE {args.month}", important=True)

        missing = [k for k in REQUIRED if k not in metrics]
        if missing:
            result = build_result(
                BANK_NAME,
                "partial_success" if metrics else "failed",
                ENTRY,
                month,
                "",
                metrics,
                "部分欄位缺漏，需人工確認。" if metrics else "無法可靠取得13項揭露欄位。",
                [{"stage": "extract", "message": f"Missing metrics: {', '.join(missing)}"}],
                steps,
            )
        else:
            result = build_result(BANK_NAME, "success", ENTRY, month, "", metrics, "", [], steps)

        print(json.dumps(result, ensure_ascii=False, separators=(",", ":")))
        return 0 if result.get("status") != "failed" else 1

    except FetchError as exc:
        log(f"FETCH failed {exc}", important=True)
        result = build_result(
            BANK_NAME,
            "failed",
            ENTRY,
            args.month,
            "",
            {},
            "一銀抓取流程失敗。",
            [{"stage": "fetch", "message": str(exc)}],
            steps,
        )
        print(json.dumps(result, ensure_ascii=False, separators=(",", ":")))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
