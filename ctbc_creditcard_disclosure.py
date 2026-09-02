#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import unicodedata
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

# 避免在唯讀執行沙箱中產生 __pycache__/*.pyc（會觸發 Refusing to overwrite）
sys.dont_write_bytecode = True

try:
    from pypdf import PdfReader
except Exception:  # pragma: no cover - optional dependency
    PdfReader = None  # type: ignore[assignment]


BASE_DIR = Path(__file__).resolve().parent
WORKSPACE_DIR = BASE_DIR.parent
BANK_KEY = "ctbc"
BANK_NAME = "中國信託商業銀行"

# 中信這支腳本是「由使用者上傳/本機檔案」解析資料。
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
    "debit_card_signed_amount": SOURCE_AMOUNT_UNIT,
    "signed_amount_thousand": SOURCE_AMOUNT_UNIT,
    "revolving_balance_thousand": SOURCE_AMOUNT_UNIT,
    "installment_balance_not_yet_due_thousand": SOURCE_AMOUNT_UNIT,
    "cash_advance_amount_thousand": SOURCE_AMOUNT_UNIT,
    "charge_off_amount_this_month_thousand": SOURCE_AMOUNT_UNIT,
    "charge_off_amount_ytd_thousand": SOURCE_AMOUNT_UNIT,
    "debit_card_signed_amount_thousand": SOURCE_AMOUNT_UNIT,
    # 逾期比率在解析階段就統一轉成小數比率（normalize_metric_value），
    # 這裡宣告 decimal_ratio 讓 run_all_banks 知道不可再除以 100。
    "overdue_3m_ratio_percent": "decimal_ratio",
    "overdue_6m_ratio_percent": "decimal_ratio",
    "allowance_coverage_ratio_percent": "%",
}

SEARCH_ROOTS = [
    WORKSPACE_DIR / "input",
    BASE_DIR / "input",
    WORKSPACE_DIR,
]

IGNORED_DIR_NAMES = {
    "__pycache__", ".git", ".venv", "venv", "node_modules", "creditcardinfo",
    "bank_pages", "fresh_scrape", "out", "ppt_images", "tmp", "test",
}

SUPPORTED_EXTENSIONS = {
    ".txt", ".md", ".csv", ".tsv", ".json", ".xlsx", ".xlsm", ".docx", ".pptx", ".pdf",
}

SOURCE_TYPE_BY_EXT = {
    ".txt": "manual_upload_text",
    ".md": "manual_upload_text",
    ".csv": "manual_upload_table",
    ".tsv": "manual_upload_table",
    ".json": "manual_upload_json",
    ".xlsx": "manual_upload_excel",
    ".xlsm": "manual_upload_excel",
    ".docx": "manual_upload_docx",
    ".pptx": "manual_upload_pptx",
    ".pdf": "manual_upload_pdf",
}

FILENAME_KEYWORDS = [
    "中信", "中國信託", "ctbc", "中國信託商業銀行", "ctbc bank", "chinatrust", "中信銀", "中信卡",
]
CONTENT_KEYWORDS = [
    "中信", "中國信託", "CTBC", "信用卡", "中國信託商業銀行", "中信銀", "重要業務及財務資訊",
    "流通卡數", "有效卡數", "當月發卡數", "當月停卡數", "當月簽帳金額", "循環信用餘額", "未到期分期付款餘額", "當月預借現金金額", "逾期三個月以上帳款占應收帳款餘額(含催收款)之比率", "逾期六個月以上帳款占應收帳款餘額(含催收款)之比率", "備抵呆帳提足率", "當月轉銷呆帳金額", "當年度累計轉銷呆帳金額",
]
PREFERRED_DIR_NAMES = ["input", "uploads", "upload", "manual", "manual_upload", "ctbc", "中信"]
PREFERRED_FILE_TOKENS = ["credit", "card", "信用卡", "業務", "財務", "揭露", "資料", "月報", "report", "stat"]
FAST_SCAN_BYTES = 65536
MAX_CANDIDATES_TO_PARSE = 12

LABEL_ALIASES = {
    "circulating_cards": ["流通卡數", "circulating_cards", " y q d  "],
    "valid_cards": ["有效卡數", "active_cards", "valid_cards", "   ĥd  "],
    "new_cards_this_month": ["當月發卡數", "new_cards", "new_cards_this_month", "    o d  "],
    "cancelled_cards_this_month": ["當月停卡數", "cancelled_cards", "cancelled_cards_this_month", "   백 d  "],
    "signed_amount_thousand": ["當月簽帳金額", "簽帳金額", "monthly_spending", "signed_amount_thousand", "    ñ b   B"],
    "revolving_balance_thousand": ["循環信用餘額", "revolving_balance", "revolving_balance_thousand", " `   H ξl B"],
    "installment_balance_not_yet_due_thousand": ["未到期分期付款餘額", "installment_balance_not_yet_due", "installment_balance_not_yet_due_thousand"],
    "cash_advance_amount_thousand": ["當月預借現金金額", "cash_advance_amount", "cash_advance_amount_thousand"],
    "overdue_3m_ratio_percent": ["逾期三個月以上帳款占應收帳款餘額(含催收款)之比率", "逾期三個月以上帳款佔應收帳款餘額(含催收款)之比率", "overdue_3m_ratio_percent"],
    "overdue_6m_ratio_percent": ["逾期六個月以上帳款占應收帳款餘額(含催收款)之比率", "逾期六個月以上帳款佔應收帳款餘額(含催收款)之比率", "overdue_6m_ratio_percent"],
    "allowance_coverage_ratio_percent": ["備抵呆帳提足率", "allowance_coverage_ratio_percent"],
    "charge_off_amount_this_month_thousand": ["當月轉銷呆帳金額", "charge_off_amount_this_month", "charge_off_amount_this_month_thousand"],
    "charge_off_amount_ytd_thousand": ["當年度累計轉銷呆帳金額", "當年度轉銷呆帳金額累計至資料月份", "charge_off_amount_ytd", "charge_off_amount_ytd_thousand"],
    "debit_card_signed_amount_thousand": ["當月轉帳卡簽帳金額", "轉帳卡簽帳金額", "debit_card_signed_amount_thousand", "      b dñ b   B"],
}

TEMPLATE_ROW_ORDER = [
    "circulating_cards",
    "valid_cards",
    "new_cards_this_month",
    "cancelled_cards_this_month",
    "revolving_balance_thousand",
    "installment_balance_not_yet_due_thousand",
    "signed_amount_thousand",
    "cash_advance_amount_thousand",
    "overdue_3m_ratio_percent",
    "overdue_6m_ratio_percent",
    "allowance_coverage_ratio_percent",
    "charge_off_amount_this_month_thousand",
    "charge_off_amount_ytd_thousand",
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
OPTIONAL_METRICS = ["debit_card_signed_amount_thousand"]


def configure_utf8_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            if hasattr(stream, "reconfigure"):
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


configure_utf8_stdio()


def now_iso() -> str:
    tz = timezone(timedelta(hours=8))
    return datetime.now(tz).isoformat(timespec="seconds")


def normalize_month_text(text: str) -> str:
    if not text:
        return ""
    s = str(text).strip().replace("民國", "").strip()
    m = re.search(r"(?P<y>\d{3,4})\s*[\/\-.年]\s*(?P<m>\d{1,2})(?:\s*[\/\-.月]\s*\d{1,2})?", s)
    if not m:
        m2 = re.search(r"(?P<y>\d{3,4})\s*(?P<m>\d{2})\b", s)
        if not m2:
            return ""
        y = int(m2.group("y")); mm = int(m2.group("m"))
    else:
        y = int(m.group("y")); mm = int(m.group("m"))
    if not (1 <= mm <= 12):
        return ""
    if y >= 1911:
        y -= 1911
    if y <= 0 or y > 300:
        return ""
    return f"{y}年{mm:02d}月"


def normalize_value(value_text: Any) -> dict[str, Any]:
    """只清洗數字，不進行單位換算；單位由 metric_units 正式紀錄。"""
    if value_text is None:
        return {"ok": False, "normalized": ""}
    s = str(value_text).strip()
    if not s or s in {"—", "-", "－", "–", "N/A", "NA", "n/a"}:
        return {"ok": False, "normalized": ""}
    s = (s.replace("卡", "").replace("仟元", "").replace("千元", "").replace("元", "")
           .replace("$", "").replace("％", "%").replace("\u00a0", " ").strip())
    m = re.search(r"-?\d[\d,]*\.?\d*", s)
    if not m:
        return {"ok": False, "normalized": ""}
    num = m.group(0).replace(",", "")
    if num.endswith("."):
        num = num[:-1]
    try:
        if "." in num:
            f = float(num)
            num = str(int(f)) if f.is_integer() else str(f).rstrip("0").rstrip(".")
        else:
            num = str(int(num))
    except Exception:
        return {"ok": False, "normalized": ""}
    return {"ok": True, "normalized": num}


def read_text_with_fallbacks(path: Path) -> str:
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "utf-8", "cp950", "big5", "utf-16"):
        try:
            return raw.decode(enc)
        except Exception:
            continue
    return raw.decode("utf-8", errors="replace")


def strip_tags(raw_html: str) -> str:
    text = re.sub(r"<script[\s\S]*?</script>", " ", raw_html, flags=re.I)
    text = re.sub(r"<style[\s\S]*?</style>", " ", text, flags=re.I)
    text = re.sub(r"</?(w:tab|a:tab)[^>]*>", "\t", text, flags=re.I)
    text = re.sub(r"</?(w:br|a:br|w:cr)[^>]*>", "\n", text, flags=re.I)
    text = re.sub(r"</?(w:p|a:p|a:t|w:t|row|c|si)[^>]*>", " ", text, flags=re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = text.replace("&nbsp;", " ")
    text = re.sub(r"\r\n?", "\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def extract_text_from_xlsx(path: Path) -> str:
    wb = load_workbook(path, data_only=True, read_only=True)
    chunks: list[str] = []
    for ws in wb.worksheets:
        chunks.append(f"[sheet]{ws.title}")
        for row in ws.iter_rows(values_only=True):
            values = [str(cell).strip() for cell in row if cell not in (None, "")]
            if values:
                chunks.append("\t".join(values))
    return "\n".join(chunks)


def extract_xlsx_rows(path: Path) -> list[list[Any]]:
    wb = load_workbook(path, data_only=True, read_only=True)
    rows: list[list[Any]] = []
    for ws in wb.worksheets:
        rows.append([f"[sheet]{ws.title}"])
        for row in ws.iter_rows(values_only=True):
            values = list(row)
            if any(cell not in (None, "") for cell in values):
                rows.append(values)
    return rows


def normalize_label_text(text: Any) -> str:
    s = unicodedata.normalize("NFKC", str(text or ""))
    s = s.replace(" ", " ")
    s = s.replace("佔", "占")
    s = s.replace("（", "(").replace("）", ")")
    s = s.replace("金", "金").replace("年", "年").replace("度", "度").replace("累", "累")
    s = s.replace("％", "%")
    s = re.sub(r"\s+", "", s)
    s = s.replace("(%)", "").replace("%", "")
    s = s.replace("之比率", "比率")
    return s.strip()


PERCENT_METRIC_KEYS = {
    "overdue_3m_ratio_percent",
    "overdue_6m_ratio_percent",
    "allowance_coverage_ratio_percent",
}

CTBC_FIXED_XLSX_ROW_ORDER = [
    ("circulating_cards", "流通卡數"),
    ("valid_cards", "有效卡數"),
    ("new_cards_this_month", "當月發卡數"),
    ("cancelled_cards_this_month", "當月停卡數"),
    ("revolving_balance_thousand", "循環信用餘額"),
    ("signed_amount_thousand", "當月簽帳金額"),
    ("cash_advance_amount_thousand", "當月預借現金金額"),
    ("overdue_3m_ratio_percent", "逾期三個月以上帳款占應收帳款餘額(含催收款)之比率"),
    ("overdue_6m_ratio_percent", "逾期六個月以上帳款占應收帳款餘額(含催收款)之比率"),
    ("allowance_coverage_ratio_percent", "備抵呆帳提足率"),
    ("charge_off_amount_this_month_thousand", "當月轉銷呆帳金額"),
    ("charge_off_amount_ytd_thousand", "當年度累計轉銷呆帳金額"),
    ("installment_balance_not_yet_due_thousand", "未到期分期付款餘額"),
    ("debit_card_signed_amount_thousand", "當月轉帳卡簽帳金額"),
]
CTBC_FIXED_XLSX_ROW_ORDER_NORMALIZED = [
    (metric_key, normalize_label_text(label)) for metric_key, label in CTBC_FIXED_XLSX_ROW_ORDER
]


STRUCTURED_XLSX_LABEL_MAP = {
    normalize_label_text("流通卡數"): "circulating_cards",
    normalize_label_text("有效卡數"): "valid_cards",
    normalize_label_text("當月發卡數"): "new_cards_this_month",
    normalize_label_text("當月停卡數"): "cancelled_cards_this_month",
    normalize_label_text("循環信用餘額"): "revolving_balance_thousand",
    normalize_label_text("當月簽帳金額"): "signed_amount_thousand",
    normalize_label_text("當月預借現金金額"): "cash_advance_amount_thousand",
    normalize_label_text("逾期三個月以上帳款占應收帳款餘額(含催收款)之比率"): "overdue_3m_ratio_percent",
    normalize_label_text("逾期六個月以上帳款占應收帳款餘額(含催收款)之比率"): "overdue_6m_ratio_percent",
    normalize_label_text("備抵呆帳提足率"): "allowance_coverage_ratio_percent",
    normalize_label_text("當月轉銷呆帳金額"): "charge_off_amount_this_month_thousand",
    normalize_label_text("當年度累計轉銷呆帳金額"): "charge_off_amount_ytd_thousand",
    normalize_label_text("未到期分期付款餘額"): "installment_balance_not_yet_due_thousand",
    normalize_label_text("當月轉帳卡簽帳金額"): "debit_card_signed_amount_thousand",
}


def row_text(row: list[Any]) -> str:
    return " ".join(str(cell).strip() for cell in row if cell not in (None, "")).strip()


def row_number_candidates(row: list[Any]) -> list[str]:
    values: list[str] = []
    for cell in row:
        parsed = normalize_value(cell)
        if parsed["ok"]:
            values.append(str(parsed["normalized"]))
    return values


# vendored from percent_utils.py — 修改時請與 percent_utils.py 同步
def to_number(value: Any) -> float | int | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return int(value) if isinstance(value, float) and value.is_integer() else value

    text = str(value).strip().replace(',', '')
    if text in {'', '-', '－', '—'}:
        return None
    try:
        num = float(text)
        return int(num) if num.is_integer() else num
    except ValueError:
        return None


def normalize_percent_value(value: Any, *, number_format: str = '', assume_percent_input: bool = False, force_percent_input: bool = False) -> float | None:
    if value is None or value == '':
        return None

    raw_text = str(value).strip()
    raw_has_percent = '%' in raw_text or '％' in raw_text
    normalized_text = raw_text.replace('％', '%')
    if raw_has_percent:
        normalized_text = re.sub(r'\s*%\s*', '', normalized_text)

    number = to_number(normalized_text if raw_has_percent else value)
    if number is None:
        return None

    decimal = float(number)
    if raw_has_percent:
        return decimal / 100.0
    if '%' in str(number_format or ''):
        return decimal
    if force_percent_input:
        return decimal / 100.0
    if assume_percent_input:
        return decimal / 100.0 if abs(decimal) > 1 else decimal
    if abs(decimal) > 1:
        return decimal / 100.0
    return decimal


def format_decimal_ratio(value: float) -> str:
    text = repr(float(value))
    if "e" not in text and "E" not in text and "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


def normalize_metric_value(metric_key: str, value: Any, number_format: str = "") -> str | None:
    parsed = normalize_value(value)
    if not parsed["ok"]:
        return None
    normalized = str(parsed["normalized"])
    if metric_key == "allowance_coverage_ratio_percent":
        percent_value = normalize_percent_value(
            normalized,
            number_format=number_format,
            force_percent_input=False,
        )
        if percent_value is None:
            return normalized
        return format_decimal_ratio(percent_value * 100)
    if metric_key in PERCENT_METRIC_KEYS:
        # 逾期比率一律輸出小數比率：儲存格是 % 格式時 raw 值已是小數（保留），
        # 文字含 % 或裸數字（含 (%) 欄的顯示值，如 0.12 代表 0.12%）一律 /100。
        raw_text = str(value or "").strip()
        percent_value = normalize_percent_value(
            raw_text or normalized,
            number_format=number_format,
            force_percent_input=True,
        )
        if percent_value is not None:
            return format_decimal_ratio(percent_value)
    return normalized


def cell_raw_value(item: Any) -> Any:
    return getattr(item, "value", item)


def parse_ctbc_fixed_layout_rows(first_sheet_rows: list[list[Any]]) -> dict[str, str]:
    """rows 內容可為儲存格物件或純值；傳儲存格才能取得 number_format，
    讓百分比欄位正確判斷「% 格式 raw 已是小數」與「裸數字要 /100」。"""
    if len(first_sheet_rows) < len(CTBC_FIXED_XLSX_ROW_ORDER) + 1:
        return {}
    header = first_sheet_rows[0] if first_sheet_rows else []
    if len(header) < 4:
        return {}
    header_col0 = str(cell_raw_value(header[0]) or "")
    header_col1 = str(cell_raw_value(header[1]) or "") if len(header) > 1 else ""
    if "應揭露項目" not in header_col1 and "應揭露項目" not in header_col0:
        return {}
    fixed_metrics: dict[str, str] = {}
    for idx, (metric_key, expected_label) in enumerate(CTBC_FIXED_XLSX_ROW_ORDER_NORMALIZED, start=1):
        row_index = idx
        if row_index >= len(first_sheet_rows):
            return {}
        row = first_sheet_rows[row_index]
        label = normalize_label_text(cell_raw_value(row[1]) if len(row) > 1 else "")
        # 來源表順序穩定，優先按列序抓；若標籤偏移就記錄 hint，仍繼續以固定順序取值。
        value = row[3] if len(row) > 3 else None
        if metric_key in PERCENT_METRIC_KEYS:
            normalized = normalize_metric_value(metric_key, cell_raw_value(value), getattr(value, 'number_format', ''))
        else:
            parsed = normalize_value(cell_raw_value(value))
            normalized = str(parsed['normalized']) if parsed['ok'] else None
        if normalized is not None:
            fixed_metrics[metric_key] = normalized
    return fixed_metrics


def parse_metrics_from_xlsx_rows(path: Path, requested_month: str = "") -> tuple[dict[str, str], str, str, list[str]]:
    wb = load_workbook(path, data_only=False, read_only=True)
    metrics: dict[str, str] = {}
    hints: list[str] = []
    data_month = ""
    base_date = ""
    first_sheet_rows: list[list[Any]] = []
    first_sheet_cell_rows: list[list[Any]] = []

    for sheet_index, ws in enumerate(wb.worksheets):
        if not data_month:
            data_month, base_date = extract_data_month(ws.title, requested_month, path)
        for row in ws.iter_rows():
            values = [cell.value for cell in row]
            if not any(value not in (None, "") for value in values):
                continue
            if sheet_index == 0:
                first_sheet_rows.append(values)
                # 保留儲存格物件，固定版型解析才能取得 number_format 判斷百分比欄位
                first_sheet_cell_rows.append(list(row))
            text = row_text(values)
            if not data_month and text:
                data_month, base_date = extract_data_month(text, requested_month, path)
            if not text:
                continue

            structured_label = values[1] if len(values) > 1 else None
            structured_metric_key = None
            if structured_label not in (None, ""):
                structured_metric_key = STRUCTURED_XLSX_LABEL_MAP.get(normalize_label_text(structured_label))
            if structured_metric_key and not metrics.get(structured_metric_key):
                value_cell = row[3] if len(row) > 3 else row[-1]
                normalized = normalize_metric_value(structured_metric_key, value_cell.value, value_cell.number_format)
                if normalized is not None:
                    metrics[structured_metric_key] = normalized
                    hints.append(f"structured:{structured_metric_key}")
                    continue

            lower_text = text.lower()
            for metric_key, aliases in LABEL_ALIASES.items():
                if metrics.get(metric_key):
                    continue
                if any(alias.lower() in lower_text for alias in aliases):
                    nums = row_number_candidates(values)
                    if nums:
                        metrics[metric_key] = nums[-1]
                        hints.append(f"alias:{metric_key}")

    fixed_layout_metrics = parse_ctbc_fixed_layout_rows(first_sheet_cell_rows)
    if fixed_layout_metrics:
        metrics.update(fixed_layout_metrics)
        hints.append("ctbc_fixed_row_order")

    if len(metrics) >= len(REQUIRED_METRICS):
        return metrics, data_month, base_date, hints

    template_title = row_text(first_sheet_rows[0]) if first_sheet_rows else ""
    template_month = row_text(first_sheet_rows[0][:4]) if first_sheet_rows else ""
    title_has_month = bool(normalize_month_text(template_month) or normalize_month_text(template_title))
    indexed_rows = [row for row in first_sheet_rows[1:] if row and row[0] not in (None, "")]
    if len(indexed_rows) >= len(TEMPLATE_ROW_ORDER) and title_has_month:
        fallback_metrics = dict(metrics)
        for idx, metric_key in enumerate(TEMPLATE_ROW_ORDER, start=1):
            if idx > len(indexed_rows):
                break
            if fallback_metrics.get(metric_key):
                continue
            nums = row_number_candidates(indexed_rows[idx - 1])
            if nums:
                fallback_metrics[metric_key] = nums[-1]
        if len(fallback_metrics) > len(metrics):
            metrics = fallback_metrics
            hints.append("template_row_order")
    if len(metrics) >= len(REQUIRED_METRICS):
        return metrics, data_month, base_date, hints

    sequential_rows: list[list[Any]] = []
    for row in first_sheet_rows:
        first = row[0] if row else None
        if isinstance(first, (int, float)) or re.fullmatch(r"\d+", str(first).strip() if first not in (None, "") else ""):
            sequential_rows.append(row)
    if len(sequential_rows) >= len(TEMPLATE_ROW_ORDER):
        fallback_metrics = dict(metrics)
        for idx, metric_key in enumerate(TEMPLATE_ROW_ORDER, start=1):
            if fallback_metrics.get(metric_key):
                continue
            matches = [row for row in sequential_rows if str(row[0]).strip() == str(idx)]
            if matches:
                nums = row_number_candidates(matches[0])
                if nums:
                    fallback_metrics[metric_key] = nums[-1]
        if len(fallback_metrics) > len(metrics):
            metrics = fallback_metrics
            hints.append("sequential_row_fallback")
    return metrics, data_month, base_date, hints


def extract_text_from_zip_xml(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".pptx":
        with zipfile.ZipFile(path) as zf:
            slide_members = sorted(name for name in zf.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", name))
            return "\n".join(strip_tags(zf.read(name).decode("utf-8", errors="replace")) for name in slide_members)
    members = ["word/document.xml"] if suffix == ".docx" else []
    with zipfile.ZipFile(path) as zf:
        return "\n".join(strip_tags(zf.read(name).decode("utf-8", errors="replace")) for name in members if name in zf.namelist())


def extract_text_from_pdf(path: Path) -> str:
    if PdfReader is None:
        return ""
    try:
        reader = PdfReader(str(path))
        return "\n".join((page.extract_text() or "") for page in reader.pages)
    except Exception:
        return ""


def extract_text_from_file(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in {".txt", ".md", ".csv", ".tsv", ".json"}:
        return read_text_with_fallbacks(path)
    if suffix in {".xlsx", ".xlsm"}:
        return extract_text_from_xlsx(path)
    if suffix in {".docx", ".pptx"}:
        return extract_text_from_zip_xml(path)
    if suffix == ".pdf":
        return extract_text_from_pdf(path)
    return ""


def normalize_search_text(text: str) -> str:
    s = unicodedata.normalize("NFKC", text or "")
    s = s.replace("\u00a0", " ")
    s = re.sub(r"\r\n?", "\n", s)
    s = re.sub(r"[ \t]+", " ", s)
    return s


def safe_mtime(path: Path) -> float:
    try:
        return path.stat().st_mtime
    except Exception:
        return 0.0


def recent_score(path: Path) -> int:
    try:
        age_seconds = max(0.0, datetime.now().timestamp() - safe_mtime(path))
    except Exception:
        return 0
    if age_seconds <= 3600:
        return 80
    if age_seconds <= 86400:
        return 50
    if age_seconds <= 7 * 86400:
        return 25
    return 0


def extension_priority(path: Path) -> int:
    return {".xlsx": 35, ".xlsm": 35, ".txt": 30, ".csv": 28, ".tsv": 28, ".docx": 26, ".md": 22, ".json": 22, ".pptx": 18, ".pdf": 18}.get(path.suffix.lower(), 0)


def quick_preview_text(path: Path) -> str:
    if path.suffix.lower() not in {".txt", ".md", ".csv", ".tsv", ".json"}:
        return ""
    try:
        raw = path.read_bytes()[:FAST_SCAN_BYTES]
    except Exception:
        return ""
    for enc in ("utf-8-sig", "utf-8", "cp950", "big5", "utf-16"):
        try:
            return raw.decode(enc, errors="replace")
        except Exception:
            continue
    return raw.decode("utf-8", errors="replace")


def packed_month_text(month: str) -> str:
    m = re.fullmatch(r"(\d{3})年(\d{2})月", month or "")
    return f"{m.group(1)}{m.group(2)}" if m else ""


def path_hint_score(path: Path, requested_month: str) -> tuple[int, list[str]]:
    score = extension_priority(path) + recent_score(path); hints: list[str] = []
    normalized_name = normalize_search_text(path.name).lower()
    normalized_path = normalize_search_text(str(path)).lower()
    lower_parts = [part.lower() for part in path.parts]
    alias_hits = [keyword for keyword in FILENAME_KEYWORDS if keyword.lower() in normalized_name]
    if alias_hits:
        score += 180 + min(60, (len(alias_hits) - 1) * 20); hints.append("檔名含中信別名")
    token_hits = [token for token in PREFERRED_FILE_TOKENS if token.lower() in normalized_name]
    if token_hits:
        score += min(90, 25 + len(token_hits) * 15); hints.append("檔名含信用卡/揭露語意")
    preferred_dir_hits = [d for d in PREFERRED_DIR_NAMES if any(d.lower() in part for part in lower_parts)]
    if preferred_dir_hits:
        score += 35 + min(30, (len(preferred_dir_hits) - 1) * 10); hints.append("位於偏好資料夾")
    preview = normalize_search_text(quick_preview_text(path))
    if preview:
        preview_lower = preview.lower()
        matched = [keyword for keyword in CONTENT_KEYWORDS if keyword.lower() in preview_lower]
        if matched:
            score += min(160, 30 + len(matched) * 14); hints.append("前段內容含中信/欄位關鍵字")
        if requested_month and requested_month in preview:
            score += 50; hints.append("前段內容符合指定月份")
    if requested_month:
        packed = packed_month_text(requested_month)
        if requested_month.lower() in normalized_path:
            score += 40; hints.append("檔名含指定月份")
        elif packed and packed in re.sub(r"\D+", "", normalized_name):
            score += 35; hints.append("檔名含壓縮月份")
    return score, hints


def extract_data_month(text: str, fallback_month: str = "", path: Path | None = None) -> tuple[str, str]:
    normalized = normalize_search_text(text)
    patterns = [
        r"(?:資料年月|資料月份|資料基期|查詢年月|年月|month|data_month|base_date)\s*[:：]?\s*(\d{3,4}[\/\-.年]\d{1,2}月?)",
        r"(\d{3,4}年\d{1,2}月)",
        r"(\d{3,4}[\/\-.]\d{1,2})",
    ]
    for pattern in patterns:
        m = re.search(pattern, normalized, flags=re.I)
        if m:
            raw = m.group(1); month = normalize_month_text(raw)
            if month:
                return month, raw
    if path:
        m = re.search(r"(\d{3,4}[年\/\-.]\d{1,2}月?)", path.name)
        if m:
            raw = m.group(1); month = normalize_month_text(raw)
            if month:
                return month, raw
    fallback = normalize_month_text(fallback_month)
    return fallback, fallback_month


def extract_metrics(text: str) -> dict[str, str]:
    normalized = normalize_search_text(text); metrics: dict[str, str] = {}
    for metric_key, aliases in LABEL_ALIASES.items():
        for alias in aliases:
            pattern = re.compile(re.escape(alias) + r"(?:\s*[（(][^)）\n]{0,20}[)）])?\s*[:：]?\s*(-?\d[\d,]*\.?\d*)", flags=re.I)
            m = pattern.search(normalized)
            if not m:
                continue
            # 文字/PDF 顯示的是百分比數字（0.12 代表 0.12%），
            # 逾期比率交給 normalize_metric_value 統一 /100 成小數比率。
            normalized_value = normalize_metric_value(metric_key, m.group(1))
            if normalized_value is not None:
                metrics[metric_key] = normalized_value; break
    return metrics


def has_ctbc_signal(path: Path, text: str) -> bool:
    haystacks = [normalize_search_text(path.name).lower(), normalize_search_text(text).lower()]
    return any(keyword.lower() in hay for hay in haystacks for keyword in FILENAME_KEYWORDS + CONTENT_KEYWORDS)


def iter_candidate_files(root: Path, max_depth: int = 2) -> list[Path]:
    if not root.exists() or not root.is_dir():
        return []
    results: list[Path] = []
    for current, dirs, files in os.walk(root):
        current_path = Path(current)
        rel_depth = len(current_path.relative_to(root).parts)
        dirs[:] = [d for d in dirs if d not in IGNORED_DIR_NAMES and not d.startswith(".")]
        if rel_depth >= max_depth:
            dirs[:] = []
        for filename in files:
            path = current_path / filename
            if path.suffix.lower() in SUPPORTED_EXTENSIONS and path.stat().st_size > 0:
                results.append(path)
    return results


def score_candidate(path: Path, text: str, metrics: dict[str, str], requested_month: str, pre_score: int = 0) -> int:
    score = pre_score + len(metrics) * 100
    normalized_name = normalize_search_text(path.name).lower()
    normalized_text = normalize_search_text(text)
    if any(keyword.lower() in normalized_name for keyword in FILENAME_KEYWORDS):
        score += 60
    if "信用卡" in normalized_name or "card" in normalized_name:
        score += 25
    if has_ctbc_signal(path, text):
        score += 25
    if requested_month and requested_month in normalized_text:
        score += 40
    return score


def evaluate_candidate(path: Path, requested_month: str, steps: list[str], pre_score: int = 0, pre_hints: list[str] | None = None) -> dict[str, Any]:
    suffix = path.suffix.lower(); xlsx_metrics: dict[str, str] = {}; xlsx_data_month = ""; xlsx_base_date = ""; xlsx_parse_hints: list[str] = []
    try:
        text = extract_text_from_file(path)
        if suffix in {".xlsx", ".xlsm"}:
            xlsx_metrics, xlsx_data_month, xlsx_base_date, xlsx_parse_hints = parse_metrics_from_xlsx_rows(path, requested_month)
    except Exception as exc:
        steps.append(f"STEP: 略過檔案 {path}，讀取失敗：{exc}")
        return {"path": path, "metrics": {}, "score": -1, "data_month": "", "base_date": "", "text": "", "hints": pre_hints or []}
    if not text.strip() and not xlsx_metrics:
        steps.append(f"STEP: 略過檔案 {path}，未讀到可解析文字。")
        return {"path": path, "metrics": {}, "score": -1, "data_month": "", "base_date": "", "text": "", "hints": pre_hints or []}
    data_month, base_date = extract_data_month(text, requested_month, path)
    metrics = extract_metrics(text)
    if len(xlsx_metrics) > len(metrics):
        metrics = xlsx_metrics
    if xlsx_data_month and not data_month:
        data_month = xlsx_data_month
    if xlsx_base_date and not base_date:
        base_date = xlsx_base_date
    score = score_candidate(path, text, metrics, requested_month, pre_score=pre_score)
    hints = list(pre_hints or [])
    if data_month:
        hints.append(f"月份={data_month}")
    if has_ctbc_signal(path, text):
        hints.append("含中信/信用卡訊號")
    hints.extend(xlsx_parse_hints)
    if len(metrics) >= len(REQUIRED_METRICS):
        hints.append(f"抓到{len(metrics)}個指標")
    steps.append(f"STEP: 掃描候選檔案 {path}，找到 {len(metrics)} 個指標" + (f"，月份={data_month}" if data_month else "") + (f"，快速分數={pre_score}" if pre_score else "") + (f"，線索={'/'.join(hints[:6])}" if hints else ""))
    return {"path": path, "metrics": metrics, "score": score, "data_month": data_month, "base_date": base_date, "text": text, "hints": hints}


def resolve_existing_path(path_text: str) -> Path:
    raw = Path(path_text); candidates: list[Path] = []
    if raw.is_absolute():
        candidates.append(raw)
    else:
        candidates.extend([Path.cwd() / raw, WORKSPACE_DIR / raw, BASE_DIR / raw])
    seen: set[str] = set()
    for candidate in candidates:
        key = str(candidate.resolve(strict=False))
        if key in seen:
            continue
        seen.add(key)
        if candidate.exists():
            return candidate
    raise FileNotFoundError(f"找不到指定輸入檔案：{path_text}")


def collect_ranked_candidates(requested_month: str, steps: list[str]) -> list[dict[str, Any]]:
    ranked: list[dict[str, Any]] = []; seen: set[str] = set()
    for root in SEARCH_ROOTS:
        found = iter_candidate_files(root)
        if found:
            steps.append(f"STEP 2: 於 {root} 找到 {len(found)} 個可檢查檔案。")
        for path in found:
            resolved_key = str(path.resolve(strict=False))
            if resolved_key in seen:
                continue
            seen.add(resolved_key)
            pre_score, hints = path_hint_score(path, requested_month)
            ranked.append({"path": path, "pre_score": pre_score, "hints": hints, "mtime": safe_mtime(path)})
    ranked.sort(key=lambda item: (item["pre_score"], item["mtime"]), reverse=True)
    return ranked


def build_result(status: str, source_path: Path | None, data_month: str, base_date: str, metrics: dict[str, str], notes: str, errors: list[dict[str, str]], steps: list[str]) -> dict[str, Any]:
    suffix = (source_path.suffix.lower() if source_path else "")
    return {
        "bank": BANK_NAME,
        "bank_key": BANK_KEY,
        "status": status,
        "data_month": data_month,
        "base_date": base_date,

        # ===== 單位紀錄 =====
        "source_amount_unit": SOURCE_AMOUNT_UNIT,
        "metric_units": dict(METRIC_UNITS),
        "amount_unit_normalized": AMOUNT_UNIT_NORMALIZED,
        "unit_note": "本腳本保留中國信託來源/上傳檔案單位；金額欄位為仟元，卡數欄位為張。若需統一為百萬元，請由主控腳本集中轉換。",

        "metrics": metrics,
        "source": {
            "source_type": SOURCE_TYPE_BY_EXT.get(suffix, "manual_upload_local_file"),
            "source_url": str(source_path.resolve()) if source_path else "",
            "official_site": False,
            "manual_upload": True,
            "manual_copy_from_official_site": True,
        },
        "notes": notes,
        "errors": errors,
        "captured_at": now_iso(),
        "steps": steps,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="CTBC / 中信信用卡揭露資料（由本機上傳檔案讀取）")
    parser.add_argument("--month", default="", help="可選，指定月份，例如 115年05月 或 115/5")
    parser.add_argument("--input", default="", help="可選，直接指定使用者上傳檔案路徑")
    return parser.parse_args()


def main() -> int:
    args = parse_args(); steps: list[str] = []
    requested_month = normalize_month_text(args.month) if args.month else ""
    try:
        if args.input:
            source_path = resolve_existing_path(args.input)
            steps.append(f"STEP 1: 使用指定檔案作為中信來源：{source_path}")
            ranked_candidates = [{"path": source_path, "pre_score": 9999, "hints": ["使用者明確指定檔案"], "mtime": safe_mtime(source_path)}]
        else:
            steps.append("STEP 1: 中信官網不直接爬取，改從使用者上傳到 Spaid 工作目錄的檔案中找資料。")
            ranked_candidates = collect_ranked_candidates(requested_month, steps)
            if ranked_candidates:
                preview = ranked_candidates[:5]
                steps.append("STEP 2-1: 快速排序前五名候選：" + "；".join(f"{item['path'].name}(預分={item['pre_score']}, 線索={','.join(item['hints'][:3]) or '無'})" for item in preview))
        if not ranked_candidates:
            result = build_result("failed", None, requested_month or args.month, "", {}, "找不到中信信用卡上傳檔案。請先將從中信官網複製的資料存成 txt / xlsx / docx / pptx / json / csv 後放到 Spaid 根目錄，再重跑。", [{"stage": "input", "message": "No CTBC manual-upload file found in workspace."}], steps)
            print(json.dumps(result, ensure_ascii=False, indent=2)); return 1

        best: dict[str, Any] | None = None
        for item in ranked_candidates[:MAX_CANDIDATES_TO_PARSE]:
            candidate = item["path"]
            current = evaluate_candidate(candidate, requested_month, steps, pre_score=item["pre_score"], pre_hints=item.get("hints") or [])
            if best is None or current["score"] > best["score"] or (current["score"] == best["score"] and safe_mtime(candidate) > safe_mtime(best["path"])):
                best = current
        assert best is not None
        source_path = best["path"]; metrics = best["metrics"]; data_month = best["data_month"]; base_date = best["base_date"]
        steps.append(f"STEP 3: 採用最佳候選檔案：{source_path}" + (f"（線索：{'/'.join(best.get('hints') or [])}）" if best.get("hints") else ""))

        if not metrics:
            result = build_result("failed", source_path, data_month or requested_month or args.month, base_date, {}, "已找到候選檔案，但無法完整辨識中信信用卡揭露欄位。請確認檔案內含資料年月與13項信用卡重要業務及財務資訊。", [{"stage": "extract", "message": f"No CTBC metrics parsed from file: {source_path.name}"}], steps)
            print(json.dumps(result, ensure_ascii=False, indent=2)); return 1

        missing = [key for key in REQUIRED_METRICS if not metrics.get(key)]
        notes_list = ["中信官網不直接爬取，本次改由使用者上傳至 Spaid 工作目錄的檔案讀取。"]
        if requested_month and data_month and requested_month != data_month:
            notes_list.append(f"指定月份為 {requested_month}，但上傳檔案辨識到的月份為 {data_month}，已回傳檔案中的資料。")
        elif requested_month and not data_month:
            notes_list.append(f"未能從檔案辨識資料月份，先沿用指定月份 {requested_month}。")
            data_month = requested_month
        if missing:
            notes_list.append("部分欄位缺漏，需人工確認。")
            result = build_result("partial_success", source_path, data_month, base_date, metrics, "；".join(notes_list), [{"stage": "extract", "message": f"Missing metrics: {', '.join(missing)}"}], steps)
        else:
            result = build_result("success", source_path, data_month, base_date, metrics, "；".join(notes_list), [], steps)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result.get("status") != "failed" else 1
    except Exception as exc:
        result = build_result("failed", None, requested_month or args.month, "", {}, "中信手動上傳檔案解析失敗。", [{"stage": "runner", "message": str(exc)}], steps)
        print(json.dumps(result, ensure_ascii=False, indent=2)); return 1


if __name__ == "__main__":
    raise SystemExit(main())
