#!/usr/bin/env python3
from __future__ import annotations

import argparse
import html as html_lib
import json
import re
import ssl
import sys
from datetime import datetime, timedelta, timezone
from io import BytesIO
from urllib.error import HTTPError, URLError
from urllib.parse import urljoin
from urllib.request import Request, urlopen

try:
    from pypdf import PdfReader
except Exception:  # pragma: no cover - handled at runtime
    PdfReader = None  # type: ignore[assignment]


# -----------------------
# Inline replacements for common/*
# -----------------------


def configure_utf8_stdio() -> None:
    """Ensure UTF-8 stdout/stderr on Windows / various terminals."""
    for stream in (sys.stdout, sys.stderr):
        try:
            if hasattr(stream, "reconfigure"):
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


configure_utf8_stdio()


BANK_KEY = "fubon"
BANK_NAME = "台北富邦銀行"
ENTRY = "https://www.fubon.com/banking/public_info/index.htm"

# 台北富邦這支腳本由官網 PDF 抽取信用卡重要業務及財務資訊。
# 目的：抓取資料並正式紀錄來源單位，不在子腳本階段轉成百萬元。
# 後續跨銀行比較時，建議由 run_all_banks.py 依 metric_units 統一轉為百萬元。
SOURCE_AMOUNT_UNIT = "仟元"
STANDARD_CARD_UNIT = "張"
AMOUNT_UNIT_NORMALIZED = False

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

# 部分環境對站台 SSL 驗證較敏感，放寬以提高抓取成功率。
SSL_CONTEXT = ssl.create_default_context()
try:
    SSL_CONTEXT.check_hostname = False
    SSL_CONTEXT.verify_mode = ssl.CERT_NONE
except Exception:
    pass

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
    "逾期三個月以上帳款佔應收帳款餘額(含催收款)之比率": "overdue_3m_ratio_percent",
    "逾期三個月以上帳款佔應收帳款餘額(含催收款)之比率(%)": "overdue_3m_ratio_percent",
    "逾期六個月以上帳款占應收帳款餘額(含催收款)之比率": "overdue_6m_ratio_percent",
    "逾期六個月以上帳款占應收帳款餘額(含催收款)之比率(%)": "overdue_6m_ratio_percent",
    "逾期六個月以上帳款佔應收帳款餘額(含催收款)之比率": "overdue_6m_ratio_percent",
    "逾期六個月以上帳款佔應收帳款餘額(含催收款)之比率(%)": "overdue_6m_ratio_percent",
    "備抵呆帳提足率": "allowance_coverage_ratio_percent",
    "備抵呆帳提足率(%)": "allowance_coverage_ratio_percent",
    "當月轉銷呆帳金額": "charge_off_amount_this_month_thousand",
    "當月轉銷呆帳金額(仟元)": "charge_off_amount_this_month_thousand",
    "當年度累計轉銷呆帳金額": "charge_off_amount_ytd_thousand",
    "當年度累計轉銷呆帳金額(仟元)": "charge_off_amount_ytd_thousand",
    "當年度轉銷呆帳金額累計至資料月份": "charge_off_amount_ytd_thousand",
    "當年度轉銷呆帳金額累計至資料月份(仟元)": "charge_off_amount_ytd_thousand",
}

# 富邦 PDF 抽字時中文常會亂碼，因此依第一頁摘要區的固定順序擷取：
# 1. 流通卡數
# 2. 有效卡數
# 3. 當月發卡數
# 4. 當月停卡數
# 5. 循環信用餘額
# 6. 未到期分期付款餘額
# 7. 當月簽帳金額
# 8. 當月預借現金金額
# 9. 逾期三個月以上帳款占應收帳款餘額(含催收款)之比率
# 10. 逾期六個月以上帳款占應收帳款餘額(含催收款)之比率
# 11. 備抵呆帳提足率
# 12. 當月轉銷呆帳金額
# 13. 當年度累計轉銷呆帳金額
TARGET_VALUE_INDEX_MAP = {
    1: "circulating_cards",
    2: "valid_cards",
    3: "new_cards_this_month",
    4: "cancelled_cards_this_month",
    5: "revolving_balance_thousand",
    6: "installment_balance_not_yet_due_thousand",
    7: "signed_amount_thousand",
    8: "cash_advance_amount_thousand",
    9: "overdue_3m_ratio_percent",
    10: "overdue_6m_ratio_percent",
    11: "allowance_coverage_ratio_percent",
    12: "charge_off_amount_this_month_thousand",
    13: "charge_off_amount_ytd_thousand",
}

PDF_LINK_TEXT_CANDIDATES = [
    "信用卡發卡機構重要業務及財務資訊",
    "信用卡重要業務及財務資訊",
]


REQUIRED_METRICS = [
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


def now_iso() -> str:
    tz = timezone(timedelta(hours=8))
    return datetime.now(tz).isoformat(timespec="seconds")


def normalize_month_text(text: str) -> str:
    """
    Normalize month text into ROC year format: '115年04月'.
    Supported inputs:
      - '115/4', '115/04', '115-4', '115年4月', '民國115年4月', '115.4'
      - '2026/04', '2026-4', '2026年4月', '2026.4'
      - '115/05/31' -> '115年05月'
    If cannot parse, return ''.
    """
    if not text:
        return ""

    s = str(text).strip().replace("民國", "").strip()

    m = re.search(r"(?P<y>\d{3,4})\s*[\/\-.年]\s*(?P<m>\d{1,2})(?:\s*[\/\-.月]\s*\d{1,2})?", s)
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
        y -= 1911

    if y <= 0 or y > 300:
        return ""

    return f"{y}年{mm:02d}月"


def normalize_value(value_text: str) -> dict[str, object]:
    """
    只清洗數字，不進行單位換算。
    單位由 build_result() 輸出的 source_amount_unit / metric_units 正式紀錄。
    """
    if value_text is None:
        return {"ok": False, "normalized": ""}

    s = str(value_text).strip()
    if not s:
        return {"ok": False, "normalized": ""}

    if s in {"—", "-", "－", "–", "N/A", "NA", "n/a"}:
        return {"ok": False, "normalized": ""}

    # 僅移除顯示用單位文字，不做單位換算。
    s = s.replace("$", "")
    s = s.replace("卡", "")
    s = s.replace("仟元", "")
    s = s.replace("千元", "")
    s = s.replace("元", "")
    s = s.replace("％", "%")
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


def request_text(url: str) -> str:
    req = Request(url, headers=dict(HEADERS), method="GET")
    with urlopen(req, timeout=30, context=SSL_CONTEXT) as resp:
        raw = resp.read()
        charset = resp.headers.get_content_charset() or "utf-8"
        return raw.decode(charset, errors="replace")


def request_bytes(url: str) -> bytes:
    req = Request(url, headers=dict(HEADERS), method="GET")
    with urlopen(req, timeout=30, context=SSL_CONTEXT) as resp:
        return resp.read()


def strip_tags(raw_html: str) -> str:
    text = re.sub(r"<script[\s\S]*?</script>", "", raw_html, flags=re.I)
    text = re.sub(r"<style[\s\S]*?</style>", "", text, flags=re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = html_lib.unescape(text)
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def fetch_entry_page(steps: list[str]) -> str:
    steps.append(f"STEP 1: 從固定入口開始：{ENTRY}")
    html = request_text(ENTRY)
    steps.append("STEP 2: 已取得入口頁 HTML。")
    return html


def find_creditcard_pdf_url(entry_html: str, steps: list[str]) -> str:
    """從富邦入口頁的內嵌 jsonData 中尋找信用卡 PDF 連結。"""
    m = re.search(r"const\s+jsonData\s*=\s*(\{[\s\S]*?\})\s*;", entry_html)
    if m:
        block = m.group(1)
        try:
            data = json.loads(block)
            items = data.get("item", [])
            for candidate_text in PDF_LINK_TEXT_CANDIDATES:
                for item in items:
                    title = str(item.get("title", "")).strip()
                    url = str(item.get("url", "")).strip()
                    if candidate_text in title and url.lower().endswith(".pdf"):
                        pdf_url = urljoin(ENTRY, url)
                        steps.append(f"STEP 3: 入口頁 JSON 找到『{title}』PDF 連結：{pdf_url}")
                        return pdf_url
        except Exception:
            steps.append("STEP 3: 入口頁 JSON 解析失敗，改用 HTML 連結掃描。")

    links = re.findall(r"<a\b[^>]*href=[\"']([^\"']+)[\"'][^>]*>([\s\S]*?)</a>", entry_html, flags=re.I)
    for href, inner_html in links:
        label = strip_tags(inner_html)
        if any(candidate in label for candidate in PDF_LINK_TEXT_CANDIDATES):
            pdf_url = urljoin(ENTRY, html_lib.unescape(href).strip())
            steps.append(f"STEP 3: 入口頁 HTML 找到『{label}』PDF 連結：{pdf_url}")
            return pdf_url

    raise ValueError("入口頁找不到『信用卡發卡機構重要業務及財務資訊 / 信用卡重要業務及財務資訊』連結。")


def fetch_pdf(pdf_url: str, steps: list[str]) -> bytes:
    data = request_bytes(pdf_url)
    if not data.startswith(b"%PDF"):
        raise ValueError("找到連結但內容不是 PDF。")
    steps.append(f"STEP 4: 已下載 PDF，大小 {len(data)} bytes。")
    return data


def extract_pdf_pages(pdf_bytes: bytes, steps: list[str]) -> list[str]:
    if PdfReader is None:
        raise RuntimeError("缺少 pypdf 套件，請先安裝：pip install pypdf")

    reader = PdfReader(BytesIO(pdf_bytes))
    if not reader.pages:
        raise ValueError("PDF 沒有頁面。")

    pages: list[str] = []
    for i, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        pages.append(text)
        steps.append(f"STEP 5: 已抽取 PDF 第 {i} 頁文字，長度 {len(text)}。")

    if not any(page.strip() for page in pages):
        raise ValueError("PDF 無法抽出文字，可能需要 OCR。")

    return pages


def extract_data_month(text: str, steps: list[str], fallback_month: str = "") -> tuple[str, str]:
    base_date_raw = ""
    data_month = ""

    patterns = [
        r"\b(\d{3,4}[\/\-.]\d{1,2}[\/\-.]\d{1,2})\b",
        r"\b(\d{3,4}[\/\-.]\d{1,2})\b",
        r"\b(\d{3,4}年\d{1,2}月)\b",
        r"\b(\d{3,4}\.\d{1,2})\b",
    ]
    for pattern in patterns:
        m = re.search(pattern, text)
        if m:
            raw = m.group(1).strip()
            normalized = normalize_month_text(raw)
            if normalized:
                base_date_raw = raw
                data_month = normalized
                steps.append(f"STEP 6: 從 PDF 文字找到資料基期：{raw} → {data_month}")
                return data_month, base_date_raw

    if fallback_month:
        data_month = normalize_month_text(fallback_month) or fallback_month
        steps.append(f"STEP 6: 使用外部指定月份：{fallback_month} → {data_month}")

    return data_month, base_date_raw


def extract_candidate_numbers_from_first_page(first_page_text: str) -> list[str]:
    """
    從第一頁摘要區抓取候選數值。
    PDF 常在頁首出現無逗號的長文號／編號，因此先找到第一個帶千分位逗號的數值，
    再將其後的數值視為摘要區候選值。
    """
    values = re.findall(r"\b\d{1,3}(?:,\d{3})+(?:\.\d+)?\b|\b\d+(?:\.\d+)?\b", first_page_text)
    if not values:
        return []

    start_index = 0
    for i, value in enumerate(values):
        if "," in value:
            start_index = i
            break

    trimmed = values[start_index:]
    return [value for value in trimmed if value]


def first_page_summary_lines(first_page_text: str) -> list[str]:
    summary_text = first_page_text.split("揭露項目及認定標準", 1)[0]
    lines: list[str] = []
    for raw_line in summary_text.splitlines():
        line = re.sub(r"\s+", " ", raw_line).strip()
        if line:
            lines.append(line)
    return lines


def extract_metrics_by_label(first_page_text: str, steps: list[str]) -> dict[str, str]:
    metrics: dict[str, str] = {}
    lines = first_page_summary_lines(first_page_text)
    label_pairs = sorted(TARGET_LABELS.items(), key=lambda item: len(item[0]), reverse=True)

    for line in lines:
        compact_line = re.sub(r"\s+", "", line)
        for label_text, metric_key in label_pairs:
            compact_label = re.sub(r"\s+", "", label_text)
            if not compact_line.startswith(compact_label):
                continue
            raw_value_text = compact_line[len(compact_label):]
            normalized = normalize_value(raw_value_text)
            if normalized["ok"]:
                metrics[metric_key] = str(normalized["normalized"])
                steps.append(f"STEP 7: 依標籤擷取 {metric_key} = {normalized['normalized']}")
            else:
                steps.append(f"STEP 7: 找到標籤 {label_text}，但數值清洗失敗：{line}")
            break

    return metrics


def extract_metrics(first_page_text: str, steps: list[str]) -> dict[str, str]:
    metrics = extract_metrics_by_label(first_page_text, steps)
    if len(metrics) >= len(REQUIRED_METRICS):
        steps.append("STEP 7: 已透過標籤比對完整抓取富邦 13 項揭露欄位。")
        return metrics

    ordered_values = extract_candidate_numbers_from_first_page(first_page_text)
    if ordered_values:
        preview = ", ".join(ordered_values[:13])
        steps.append(f"STEP 8: 從第一頁摘要區擷取到前 13 個候選數值：{preview}")
    else:
        steps.append("STEP 8: 未找到可用的候選數值。")
        return metrics

    if len(ordered_values) < 13:
        steps.append(f"STEP 8: 候選數值不足 13 個，僅有 {len(ordered_values)} 個。")
        return metrics

    leading_values = ordered_values[:13]
    for position, metric_key in TARGET_VALUE_INDEX_MAP.items():
        if metrics.get(metric_key):
            continue
        raw_value = leading_values[position - 1]
        normalized = normalize_value(raw_value)
        if normalized["ok"]:
            metrics[metric_key] = str(normalized["normalized"])
            steps.append(f"STEP 8: 依固定順序第 {position} 個值補擷取 {metric_key} = {normalized['normalized']}")
        else:
            steps.append(f"STEP 8: 第 {position} 個候選值清洗失敗：{raw_value}")

    return metrics


def build_result(
    status: str,
    source_url: str,
    data_month: str,
    base_date: str,
    metrics: dict[str, str],
    notes: str,
    errors: list[dict[str, str]],
    steps: list[str],
) -> dict[str, object]:
    return {
        "bank": BANK_NAME,
        "bank_key": BANK_KEY,
        "status": status,
        "data_month": data_month,
        "base_date": base_date,

        # ===== 單位紀錄 =====
        # source_amount_unit：本銀行來源金額單位。
        # metric_units：各 metrics 欄位對應來源單位。
        # amount_unit_normalized：False 表示本子腳本只紀錄來源單位，尚未轉成百萬元。
        "source_amount_unit": SOURCE_AMOUNT_UNIT,
        "metric_units": dict(METRIC_UNITS),
        "amount_unit_normalized": AMOUNT_UNIT_NORMALIZED,
        "unit_note": "本腳本保留台北富邦來源 PDF 單位；金額欄位為仟元，卡數欄位為張。若需統一為百萬元，請由主控腳本集中轉換。",

        "metrics": metrics,
        "source": {
            "source_type": "pdf",
            "source_url": source_url,
            "official_site": True,
        },
        "notes": notes,
        "errors": errors,
        "captured_at": now_iso(),
        "steps": steps,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Fubon credit card disclosure extractor (standalone)")
    parser.add_argument("--month", default="", help="Optional target month, e.g. 115年05月 or 115/5")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    steps: list[str] = []

    try:
        entry_html = fetch_entry_page(steps)
        pdf_url = find_creditcard_pdf_url(entry_html, steps)
        pdf_bytes = fetch_pdf(pdf_url, steps)
        page_texts = extract_pdf_pages(pdf_bytes, steps)

        merged_text = "\n".join(page_texts)
        first_page_text = page_texts[0] if page_texts else ""

        data_month, base_date_raw = extract_data_month(merged_text, steps, args.month)
        metrics = extract_metrics(first_page_text, steps)

        missing = [k for k in REQUIRED_METRICS if not metrics.get(k)]

        if missing:
            result = build_result(
                "partial_success" if metrics else "failed",
                pdf_url,
                data_month,
                base_date_raw,
                metrics,
                "部分欄位缺漏，需人工確認。富邦 PDF 文字抽取存在亂碼風險，本版以第一頁固定數值順序擷取主要指標。" if metrics else "無法可靠取得13項揭露欄位。",
                [{"stage": "extract", "message": f"Missing metrics: {', '.join(missing)}"}],
                steps,
            )
        else:
            result = build_result(
                "success",
                pdf_url,
                data_month,
                base_date_raw,
                metrics,
                "富邦 PDF 文字抽取可能出現亂碼，因此本版依第一頁固定揭露順序擷取主要數值。",
                [],
                steps,
            )

        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0

    except (HTTPError, URLError, ValueError, json.JSONDecodeError, RuntimeError) as exc:
        result = build_result(
            "failed",
            ENTRY,
            normalize_month_text(args.month) or args.month,
            "",
            {},
            "富邦真實抓取流程失敗。",
            [{"stage": "fetch", "message": str(exc)}],
            steps,
        )
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
