#!/usr/bin/env python3
from __future__ import annotations

import argparse
import html as html_lib
import json
import re
import ssl
import sys
import unicodedata
from datetime import datetime, timedelta, timezone
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


def configure_utf8_stdio() -> None:
    """Ensure UTF-8 stdout/stderr on Windows / various terminals."""
    for stream in (sys.stdout, sys.stderr):
        try:
            if hasattr(stream, "reconfigure"):
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


configure_utf8_stdio()


def normalize_month_text(text: str) -> str:
    """月份文字正規化為民國格式（如 115年05月）；無法解析回傳空字串。"""
    if not text:
        return ""

    s = str(text).strip()
    s = s.replace("民國", "").strip()

    m = re.search(r"(?P<y>\d{3,4})\s*[\/\-.年]\s*(?P<m>\d{1,2})", s)
    if not m:
        m2 = re.search(r"(?P<y>\d{3,4})\s*(?P<m>\d{2})\b", s)
        if not m2:
            return ""
        y = int(m2.group("y"))
        mm = int(m2.group("m"))
    else:
        y = int(m.group("y"))
        mm = int(m.group("m"))

    if not (1 <= mm <= 12):
        return ""

    if y >= 1911:
        y = y - 1911

    if y <= 0 or y > 300:
        return ""

    return f"{y}年{mm:02d}月"


def normalize_value(value_text: str) -> dict:
    """清洗數值文字（去千分位、單位、貨幣符號），不做單位換算；回傳 {"ok", "normalized"}。"""
    if value_text is None:
        return {"ok": False, "normalized": ""}

    s = str(value_text).strip()
    if not s:
        return {"ok": False, "normalized": ""}

    if s in {"—", "-", "－", "–", "N/A", "NA", "n/a"}:
        return {"ok": False, "normalized": ""}

    s = s.replace("卡", "")
    s = s.replace("仟元", "")
    s = s.replace("千元", "")
    s = s.replace("元", "")
    s = s.replace("％", "%")
    s = s.replace("$", "")
    s = s.replace("\u00a0", " ")
    s = s.strip()

    m = re.search(r"-?\d[\d,]*\.?\d*", s)
    if not m:
        return {"ok": False, "normalized": ""}

    num = m.group(0).replace(",", "")
    if num.endswith("."):
        num = num[:-1]

    try:
        if "." in num:
            f = float(num)
            if f.is_integer():
                num = str(int(f))
            else:
                num = str(f).rstrip("0").rstrip(".")
        else:
            num = str(int(num))
    except Exception:
        return {"ok": False, "normalized": ""}

    return {"ok": True, "normalized": num}


BANK_KEY = "sinopac"
BANK_NAME = "永豐銀行"
ENTRY = "https://bank.sinopac.com/sinopacBT/about/introduction/announcement/legal-disclaimers/20180814174552445000000000000582.html"

# 永豐銀行這支腳本由官網 HTML 抽取信用卡重要業務資訊。
SOURCE_AMOUNT_UNIT = "仟元"
STANDARD_CARD_UNIT = "張"

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

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0 Safari/537.36",
    "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
}

TARGET_LABELS = {
    "流通卡數": "circulating_cards",
    "有效卡數": "valid_cards",
    "當月發卡數": "new_cards_this_month",
    "當月停卡數": "cancelled_cards_this_month",
    "當月簽帳金額": "signed_amount_thousand",
    "當月簽帳金額(仟元)": "signed_amount_thousand",
    "循環信用餘額": "revolving_balance_thousand",
    "循環信用餘額(仟元)": "revolving_balance_thousand",
    "未到期分期付款餘額": "installment_balance_not_yet_due_thousand",
    "未到期分期付款餘額(仟元)": "installment_balance_not_yet_due_thousand",
    "當月預借現金金額": "cash_advance_amount_thousand",
    "當月預借現金金額(仟元)": "cash_advance_amount_thousand",
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
    "備抵呆帳提足率(%)": "allowance_coverage_ratio_percent",
    "當月轉銷呆帳金額": "charge_off_amount_this_month_thousand",
    "當月轉銷呆帳金額(仟元)": "charge_off_amount_this_month_thousand",
    "當年度累計轉銷呆帳金額": "charge_off_amount_ytd_thousand",
    "當年度累計轉銷呆帳金額(仟元)": "charge_off_amount_ytd_thousand",
    "當年度轉銷呆帳金額累計至資料月份": "charge_off_amount_ytd_thousand",
    "當年度轉銷呆帳金額累計至資料月份(仟元)": "charge_off_amount_ytd_thousand",
}


NORMALIZED_TARGET_LABELS = {
    unicodedata.normalize("NFKC", key).replace("(", "（").replace(")", "）").replace(" ", ""): value
    for key, value in TARGET_LABELS.items()
}


def now_iso() -> str:
    tz = timezone(timedelta(hours=8))
    return datetime.now(tz).isoformat(timespec="seconds")


def request_text(url: str) -> str:
    req = Request(url, headers=HEADERS, method="GET")
    try:
        with urlopen(req, timeout=30) as resp:
            raw = resp.read()
            charset = resp.headers.get_content_charset() or "utf-8"
            return raw.decode(charset, errors="replace")
    except URLError as exc:
        reason = getattr(exc, "reason", exc)
        if isinstance(reason, ssl.SSLCertVerificationError):
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
            with urlopen(req, timeout=30, context=ctx) as resp:
                raw = resp.read()
                charset = resp.headers.get_content_charset() or "utf-8"
                return raw.decode(charset, errors="replace")
        raise


def strip_tags(raw_html: str) -> str:
    """Remove scripts/styles/tags and unescape HTML entities."""
    text = re.sub(r"<script[\s\S]*?</script>", "", raw_html, flags=re.I)
    text = re.sub(r"<style[\s\S]*?</style>", "", text, flags=re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = html_lib.unescape(text)
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_label_text(text: str) -> str:
    """Normalize label text for robust matching against compatibility characters like 數 / 金."""
    s = strip_tags(text)
    s = unicodedata.normalize("NFKC", s)
    s = re.sub(r"^\s*\d{1,2}[\.．、]?\s*", "", s)
    s = s.replace("(", "（").replace(")", "）")
    s = s.replace(" ", "")
    return s.strip()


def fetch_page(steps: list[str]) -> str:
    steps.append(f"STEP 1: 從固定入口開始：{ENTRY}")
    html = request_text(ENTRY)
    steps.append("STEP 2: 已取得永豐銀行信用卡重要業務資訊頁 HTML。")
    return html


def extract_data_month(html: str, steps: list[str], fallback_month: str = "") -> tuple[str, str]:
    """
    Returns (data_month, base_date_raw)
      - data_month normalized to ROC format like '115年05月'
      - base_date_raw keeps original base text if found
    """
    base_date_raw = ""
    data_month = ""

    patterns = [
        r"<td[^>]*>\s*<strong>\s*項目\s*</strong>\s*</td>\s*<td[^>]*>\s*<strong>\s*([^<]+?)\s*</strong>\s*</td>",
        r">\s*(\d{3,4}年\d{1,2}月)\s*<",
        r">\s*(\d{3,4}[\./-]\d{1,2})\s*<",
    ]

    for pattern in patterns:
        m = re.search(pattern, html, flags=re.I)
        if m:
            raw = strip_tags(m.group(1)).strip()
            normalized = normalize_month_text(raw)
            if normalized:
                base_date_raw = raw
                data_month = normalized
                steps.append(f"STEP 3: 找到資料月份：{raw} → {data_month}")
                return data_month, base_date_raw

    if not data_month and fallback_month:
        data_month = normalize_month_text(fallback_month) or fallback_month
        steps.append(f"STEP 3: 使用外部指定月份：{fallback_month} → {data_month}")

    return data_month, base_date_raw


def extract_metrics(html: str, steps: list[str]) -> dict:
    metrics: dict[str, str] = {}

    rows = re.findall(r"<tr[^>]*>([\s\S]*?)</tr>", html, flags=re.I)
    for row in rows:
        cells = re.findall(r"<t[dh][^>]*>([\s\S]*?)</t[dh]>", row, flags=re.I)
        if len(cells) < 2:
            continue

        label = normalize_label_text(cells[0])
        value_text = strip_tags(cells[1])

        for key_text, metric_key in NORMALIZED_TARGET_LABELS.items():
            if key_text in label:
                normalized = normalize_value(value_text)
                if normalized["ok"]:
                    metrics[metric_key] = normalized["normalized"]
                    steps.append(f"STEP 4: 擷取 {key_text} = {normalized['normalized']}")
                else:
                    steps.append(f"STEP 4: 找到 {key_text}，但數值清洗失敗：{value_text}")
                break

    return metrics


def build_result(
    status: str,
    source_url: str,
    data_month: str,
    base_date: str,
    metrics: dict,
    notes: str,
    errors: list[dict],
    steps: list[str],
) -> dict:
    return {
        "bank": BANK_NAME,
        "bank_key": BANK_KEY,
        "status": status,
        "data_month": data_month,
        "base_date": base_date,
        "metric_units": dict(METRIC_UNITS),
        "metrics": metrics,
        "source": {
            "source_type": "html",
            "source_url": source_url,
            "official_site": True,
        },
        "notes": notes,
        "errors": errors,
        "captured_at": now_iso(),
        "steps": steps,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="SinoPac disclosure extractor (standalone)")
    parser.add_argument("--month", default="", help="Optional target month, e.g. 115年05月 or 115/5")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    steps: list[str] = []

    try:
        html = fetch_page(steps)
        data_month, base_date_raw = extract_data_month(html, steps, args.month)
        metrics = extract_metrics(html, steps)

        required = [
            "circulating_cards",
            "valid_cards",
            "new_cards_this_month",
            "cancelled_cards_this_month",
            "signed_amount_thousand",
            "revolving_balance_thousand",
            "installment_balance_not_yet_due_thousand",
            "cash_advance_amount_thousand",
            "overdue_3m_ratio_percent",
            "overdue_6m_ratio_percent",
            "allowance_coverage_ratio_percent",
            "charge_off_amount_this_month_thousand",
            "charge_off_amount_ytd_thousand",
        ]
        missing = [k for k in required if not metrics.get(k)]

        if missing:
            result = build_result(
                "partial_success" if metrics else "failed",
                ENTRY,
                data_month,
                base_date_raw,
                metrics,
                "部分欄位缺漏，需人工確認。" if metrics else "無法可靠取得13項揭露欄位。",
                [{"stage": "extract", "message": f"Missing metrics: {', '.join(missing)}"}],
                steps,
            )
        else:
            result = build_result(
                "success",
                ENTRY,
                data_month,
                base_date_raw,
                metrics,
                "",
                [],
                steps,
            )

        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0

    except (HTTPError, URLError, ValueError, json.JSONDecodeError) as exc:
        result = build_result(
            "failed",
            ENTRY,
            normalize_month_text(args.month) or args.month,
            "",
            {},
            "永豐銀行真實抓取流程失敗。",
            [{"stage": "fetch", "message": str(exc)}],
            steps,
        )
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
