from __future__ import annotations

import argparse
import io
import json
import re
import ssl
import sys
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from copy import copy
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator

# 避免在唯讀執行沙箱中產生 __pycache__/*.pyc（會觸發 Refusing to overwrite）
sys.dont_write_bytecode = True

from bank_aliases import BANK_BUREAU_ALIASES, BANK_ITEM_ALIASES, ITEM_RANKS
from percent_utils import PERCENT_DECIMAL_FIELDS, apply_percent_number_format, normalize_percent_value
from workbook_block_helpers import (
    apply_default_font,
    canonical_block_item as shared_canonical_block_item,
    find_month_block as shared_find_month_block,
    find_or_create_month_block as shared_find_or_create_month_block,
    refresh_block_metadata as shared_refresh_block_metadata,
)

DEFAULT_WORKBOOK = Path("銀行局信用卡公開資料.xlsx")
# 回補起始月（AD YYYYMM）：自動模式與 --month 指定模式都只處理此月（含）以後的資料。
BACKFILL_START_YYYYMM = 202601
DEFAULT_ZIP_SEARCH_DIRS = [Path("."), Path("input")]
ZIP_SUFFIX = "信用卡重要資訊揭露.zip"

BANK_ORDER = ["ctbc", "fubon", "cathay", "esun", "taishin", "dbs", "ubot", "sinopac", "firstbank", "feib"]
BANK_NAMES = {
    "ctbc": "中信",
    "fubon": "富邦",
    "cathay": "國泰",
    "esun": "玉山",
    "taishin": "台新",
    "dbs": "星展",
    "ubot": "聯邦",
    "sinopac": "永豐",
    "firstbank": "第一",
    "feib": "遠東",
}
CARD_ROW_STARTS = {
    "circulating_cards": 5,
    "valid_cards": 22,
    "valid_rate": 37,
    "net_issue": 50,
    "new_cards_this_month": 63,
    "cancelled_cards_this_month": 78,
    "signed_amount_million": 93,
    "revolving_balance_million": 108,
}
SECTION_HEADER_ROWS = [4, 21, 36, 49, 62, 77, 92, 107, 122]
REQUIRED_METRICS = [
    "circulating_cards",
    "valid_cards",
    "new_cards_this_month",
    "cancelled_cards_this_month",
    "signed_amount_million",
    "revolving_balance_million",
]
RAW_MONTH_HEADER_MAP = {
    "資料年月": "data_month",
    "銀行名稱": "bank_name",
    "流通卡數": "circulating_cards",
    "有效卡數": "valid_cards",
    "當月發卡數": "new_cards_this_month",
    "當月停卡數": "cancelled_cards_this_month",
    "當月簽帳金額(百萬元)": "signed_amount_million",
    "循環信用餘額(百萬元)": "revolving_balance_million",
    "執行結果": "status",
    "失敗原因": "failure_reason",
}
HEADER_ALIASES = {
    "bank_name": ["發卡機構名稱", "金融機構名稱", "發卡銀行名稱", "銀行名稱", "機構名稱"],
    "circulating_cards": ["流通卡數"],
    "valid_cards": ["有效卡數"],
    "new_cards_this_month": ["當月發卡數"],
    "cancelled_cards_this_month": ["當月停卡數"],
    "signed_amount_raw": ["當月簽帳金額"],
    "revolving_balance_raw": ["循環信用餘額"],
    # 別名以實際檔案表頭為準（占/佔、有無 (%) 後綴都涵蓋；全形括號由 normalize 處理）。
    "installment_balance_raw": ["未到期分期付款餘額"],
    "cash_advance_raw": ["當月預借現金金額", "預借現金金額"],
    "overdue_3m_ratio_raw": [
        "逾期三個月以上帳款占應收帳款餘額(含催收款)之比率(%)",
        "逾期三個月以上帳款佔應收帳款餘額(含催收款)之比率(%)",
        "逾期三個月以上帳款占應收帳款餘額(含催收款)之比率",
        "逾期三個月以上帳款佔應收帳款餘額(含催收款)之比率",
    ],
    "overdue_6m_ratio_raw": [
        "逾期六個月以上帳款占應收帳款餘額(含催收款)之比率(%)",
        "逾期六個月以上帳款佔應收帳款餘額(含催收款)之比率(%)",
        "逾期六個月以上帳款占應收帳款餘額(含催收款)之比率",
        "逾期六個月以上帳款佔應收帳款餘額(含催收款)之比率",
    ],
    "allowance_ratio_raw": ["備抵呆帳提足率(%)", "備抵呆帳提足率"],
    "charge_off_this_month_raw": ["當月轉銷呆帳金額"],
    "charge_off_ytd_raw": ["當年度轉銷呆帳金額累計至資料月份", "當年度累計轉銷呆帳金額"],
}



@dataclass
class DownloadMeta:
    mode: str | None
    ssl_fallback_used: bool
    warning: str | None = None


@dataclass
class WorkbookResolution:
    path: Path
    fallback_used: bool
    warnings: list[str]


@dataclass
class FetchOptions:
    insecure: bool
    prefer_local_zip: bool
    zip_search_dirs: list[Path]
    explicit_zip_file: Path | None = None


def normalize_text(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "").replace("\u3000", "")).strip()


def normalize_bank_name(value: Any) -> str:
    text = normalize_text(value)
    text = text.replace("（", "(").replace("）", ")")
    return text


def parse_roc_month(text: str) -> tuple[int, int]:
    match = re.search(r"(\d{3})\D*(\d{1,2})", str(text))
    if not match:
        raise ValueError(f"無法解析民國月份: {text}")
    y, m = int(match.group(1)), int(match.group(2))
    if not 1 <= m <= 12:
        raise ValueError(f"月份不合法: {text}")
    return y, m


def roc_month_text(y: int, m: int) -> str:
    return f"{y}年{m:02d}月"


def roc_to_ad_yyyymm(y: int, m: int) -> int:
    return (y + 1911) * 100 + m


def ad_yyyymm_to_roc(yyyymm: int) -> tuple[int, int]:
    ad_y, m = divmod(int(yyyymm), 100)
    return ad_y - 1911, m


def build_zip_url(y: int, m: int) -> str:
    return f"https://www.fsc.gov.tw/userfiles/file/{y:03d}{m:02d}_{urllib.parse.quote(ZIP_SUFFIX)}"


def is_ssl_verification_error(exc: Exception) -> bool:
    if isinstance(exc, ssl.SSLCertVerificationError):
        return True
    if isinstance(exc, ssl.SSLError):
        return "CERTIFICATE_VERIFY_FAILED" in str(exc).upper()
    if isinstance(exc, urllib.error.URLError):
        reason = getattr(exc, "reason", None)
        if isinstance(reason, Exception):
            return is_ssl_verification_error(reason)
    return "CERTIFICATE_VERIFY_FAILED" in str(exc).upper()


def make_ssl_context(insecure: bool) -> ssl.SSLContext:
    ctx = ssl.create_default_context()
    if insecure:
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
    return ctx


def download_bytes(url: str, timeout: int = 60, insecure: bool = False) -> tuple[bytes, DownloadMeta]:
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})

    if insecure:
        with urllib.request.urlopen(req, context=make_ssl_context(True), timeout=timeout) as resp:
            return resp.read(), DownloadMeta("https_insecure_forced", False)

    try:
        with urllib.request.urlopen(req, context=make_ssl_context(False), timeout=timeout) as resp:
            return resp.read(), DownloadMeta("https_verified", False)
    except Exception as exc:
        if not is_ssl_verification_error(exc):
            raise
        with urllib.request.urlopen(req, context=make_ssl_context(True), timeout=timeout) as resp:
            return resp.read(), DownloadMeta(
                "https_insecure_fallback",
                True,
                f"SSL 憑證驗證失敗，已自動改用不驗證模式: {type(exc).__name__}: {exc}",
            )


def to_number(value: Any) -> float | int | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return int(value) if isinstance(value, float) and value.is_integer() else value
    text = str(value).strip().replace(",", "")
    if text in {"", "-", "－", "—"}:
        return None
    try:
        num = float(text)
        return int(num) if num.is_integer() else num
    except ValueError:
        return None


def validate_xlsx(path: Path) -> str | None:
    if not path.exists():
        return f"檔案不存在: {path}"
    if path.suffix.lower() != ".xlsx":
        return f"副檔名不是 .xlsx: {path}"
    if not zipfile.is_zipfile(path):
        return f"不是有效的 xlsx/zip 檔: {path}"
    try:
        with zipfile.ZipFile(path) as zf:
            if "[Content_Types].xml" not in zf.namelist():
                return f"缺少 [Content_Types].xml: {path}"
        wb = load_workbook(path, read_only=True, data_only=True)
        wb.close()
        return None
    except Exception as exc:
        return f"{type(exc).__name__}: {exc}"


def workbook_sort_key(path: Path) -> tuple[int, int, str]:
    name = path.name.lower()
    if name == "銀行局信用卡公開資料.xlsx":
        rank = 0
    elif name.startswith("銀行局信用卡公開資料_") and "backup" not in name:
        rank = 1
    elif name.startswith("銀行局信用卡公開資料_backup_"):
        rank = 2
    elif name.startswith("workbook"):
        rank = 3
    else:
        rank = 9
    try:
        mtime = -int(path.stat().st_mtime)
    except Exception:
        mtime = 0
    return rank, mtime, name


def list_workbook_candidates() -> list[Path]:
    patterns = ["銀行局信用卡公開資料*.xlsx", "workbook*.xlsx", "*.xlsx"]
    seen: set[Path] = set()
    out: list[Path] = []
    for pattern in patterns:
        for path in Path(".").glob(pattern):
            if path.name.startswith("~$"):
                continue
            resolved = path.resolve()
            if resolved in seen:
                continue
            seen.add(resolved)
            out.append(path)
    return sorted(out, key=workbook_sort_key)


def resolve_workbook(requested: Path, strict: bool) -> WorkbookResolution:
    error = validate_xlsx(requested)
    if error is None:
        return WorkbookResolution(requested, False, [])
    warnings = [f"指定工作簿不可用: {requested} ({error})"]
    if strict:
        raise FileNotFoundError(warnings[-1])
    requested_resolved = requested.resolve()
    for candidate in list_workbook_candidates():
        if candidate.resolve() == requested_resolved:
            continue
        error = validate_xlsx(candidate)
        if error is None:
            warnings.append(f"已自動改用可用工作簿: {candidate}")
            return WorkbookResolution(candidate, True, warnings)
        warnings.append(f"略過不可用工作簿: {candidate} ({error})")
    raise FileNotFoundError("找不到可用工作簿。" + ("；" + "；".join(warnings) if warnings else ""))


def describe_non_zip_content(data: bytes) -> str | None:
    head = data[:200].lstrip().lower()
    if head.startswith(b"<!doctype html") or head.startswith(b"<html"):
        return "內容看起來是 HTML，不是 ZIP，可能是下載到錯誤頁或另存了網頁內容"
    if head.startswith(b"{") or head.startswith(b"["):
        return "內容看起來是 JSON，不是 ZIP"
    return None


def load_zip_bytes(source: Path | bytes, source_label: str) -> bytes:
    data = source if isinstance(source, bytes) else source.read_bytes()
    if zipfile.is_zipfile(io.BytesIO(data)):
        return data
    hint = describe_non_zip_content(data)
    if hint:
        raise zipfile.BadZipFile(f"{source_label}: {hint}")
    raise zipfile.BadZipFile(f"{source_label}: 不是有效的 ZIP 檔")


def scan_bureau_month(ws) -> tuple[int, int]:
    candidates = [ws["J3"].value]
    for row in range(1, min(ws.max_row, 10) + 1):
        for col in range(1, min(ws.max_column, 15) + 1):
            value = ws.cell(row, col).value
            if value is not None and re.search(r"\d{3}\D*\d{1,2}", str(value)):
                candidates.append(value)
    for candidate in candidates:
        try:
            return parse_roc_month(str(candidate))
        except ValueError:
            continue
    raise ValueError("無法在銀行局檔案中確認資料月份")


def normalize_header_text(value: Any) -> str:
    return normalize_text(value).replace("（", "(").replace("）", ")")


def match_header_key(header_value: Any) -> str | None:
    normalized = normalize_header_text(header_value)
    for key, aliases in HEADER_ALIASES.items():
        if any(normalized == normalize_header_text(alias) for alias in aliases):
            return key
    return None


def bureau_bank_key(name: Any) -> str | None:
    normalized = normalize_bank_name(name)
    for bank_key, aliases in BANK_BUREAU_ALIASES.items():
        if any(normalized == normalize_bank_name(alias) for alias in aliases):
            return bank_key
    return None


def parse_bank_rows_from_xlsx_bytes(xlsx_bytes: bytes, source_label: str) -> dict[str, Any]:
    wb = load_workbook(io.BytesIO(xlsx_bytes), data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    roc_y, roc_m = scan_bureau_month(ws)
    month = roc_month_text(roc_y, roc_m)
    ad_yyyymm = roc_to_ad_yyyymm(roc_y, roc_m)

    header_map: dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        key = match_header_key(ws.cell(4, col).value)
        if key and key not in header_map:
            header_map[key] = col

    missing_headers = [k for k in HEADER_ALIASES if k not in header_map]
    if missing_headers:
        raise ValueError(f"銀行局檔案缺少欄位: {missing_headers}; source={source_label}")

    def cell_thousand_to_million(row: int, key: str) -> float | int | None:
        col = header_map.get(key)
        if col is None:
            return None
        num = to_number(ws.cell(row, col).value)
        return None if num is None else num / 1000

    def cell_percent_decimal(row: int, key: str, *, force_percent_input: bool) -> float | None:
        col = header_map.get(key)
        if col is None:
            return None
        cell = ws.cell(row, col)
        return normalize_percent_value(
            cell.value,
            number_format=str(cell.number_format or ""),
            force_percent_input=force_percent_input,
            assume_percent_input=not force_percent_input,
        )

    banks: dict[str, dict[str, Any]] = {}
    matched_rows: list[dict[str, Any]] = []
    for row in range(5, ws.max_row + 1):
        bank_name = ws.cell(row, header_map["bank_name"]).value
        bank_key = bureau_bank_key(bank_name)
        if bank_key is None:
            continue
        raw = {
            "circulating_cards": to_number(ws.cell(row, header_map["circulating_cards"]).value),
            "valid_cards": to_number(ws.cell(row, header_map["valid_cards"]).value),
            "new_cards_this_month": to_number(ws.cell(row, header_map["new_cards_this_month"]).value),
            "cancelled_cards_this_month": to_number(ws.cell(row, header_map["cancelled_cards_this_month"]).value),
            "signed_amount_million": cell_thousand_to_million(row, "signed_amount_raw"),
            "revolving_balance_million": cell_thousand_to_million(row, "revolving_balance_raw"),
            "installment_balance_not_yet_due_million": cell_thousand_to_million(row, "installment_balance_raw"),
            "cash_advance_amount_million": cell_thousand_to_million(row, "cash_advance_raw"),
            # 逾期比率是裸的百分比數字（0.12 代表 0.12%），一律 /100 成小數比率；
            # 備抵呆帳提足率恆大於 1（如 429.67），沿用 abs > 1 才 /100 的啟發式。
            "overdue_3m_ratio_percent": cell_percent_decimal(row, "overdue_3m_ratio_raw", force_percent_input=True),
            "overdue_6m_ratio_percent": cell_percent_decimal(row, "overdue_6m_ratio_raw", force_percent_input=True),
            "allowance_coverage_ratio_percent": cell_percent_decimal(row, "allowance_ratio_raw", force_percent_input=False),
            "charge_off_amount_this_month_million": cell_thousand_to_million(row, "charge_off_this_month_raw"),
            "charge_off_amount_ytd_million": cell_thousand_to_million(row, "charge_off_ytd_raw"),
        }
        banks[bank_key] = {
            "bank_key": bank_key,
            "bank_name": BANK_NAMES[bank_key],
            **raw,
        }
        matched_rows.append({"row": row, "bank_key": bank_key, "bank_name": BANK_NAMES[bank_key]})

    missing_banks = [k for k in BANK_ORDER if k not in banks]
    return {
        "month": month,
        "roc_year": roc_y,
        "roc_month": roc_m,
        "ad_yyyymm": ad_yyyymm,
        "banks": banks,
        "matched_rows": matched_rows,
        "missing_banks": missing_banks,
    }


def parse_bank_rows_from_zip_bytes(zip_bytes: bytes, source_label: str) -> dict[str, Any]:
    zip_bytes = load_zip_bytes(zip_bytes, source_label)
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        xlsx_names = [name for name in zf.namelist() if name.lower().endswith(".xlsx")]
        if not xlsx_names:
            raise ValueError(f"ZIP 內找不到 xlsx: {source_label}")
        return parse_bank_rows_from_xlsx_bytes(zf.read(xlsx_names[0]), source_label)


def parse_bank_rows_from_zip_file(path: Path) -> dict[str, Any]:
    payload = parse_bank_rows_from_zip_bytes(path.read_bytes(), str(path))
    payload.update({
        "source_kind": "local_zip_explicit",
        "source_url": None,
        "local_zip_path": str(path),
        "download_mode": None,
        "ssl_fallback_used": False,
        "warnings": [],
    })
    return payload


def local_zip_sort_key(path: Path, roc_yyyymm: str) -> tuple[int, int, str]:
    name = path.name.lower()
    if name == f"market_{roc_yyyymm}.zip":
        rank = 0
    elif roc_yyyymm in name and "信用卡重要資訊揭露" in path.name:
        rank = 1
    elif roc_yyyymm in name:
        rank = 2
    else:
        rank = 9
    try:
        mtime = -int(path.stat().st_mtime)
    except Exception:
        mtime = 0
    return rank, mtime, name


def find_local_zip_candidates(y: int, m: int, search_dirs: list[Path]) -> list[Path]:
    roc_yyyymm = f"{y:03d}{m:02d}"
    patterns = [f"market_{roc_yyyymm}.zip", f"*{roc_yyyymm}*.zip", "*.zip"]
    seen: set[Path] = set()
    out: list[Path] = []
    for base_dir in search_dirs:
        if not base_dir.exists() or not base_dir.is_dir():
            continue
        for pattern in patterns:
            for path in base_dir.glob(pattern):
                if not path.is_file():
                    continue
                if pattern == "*.zip" and roc_yyyymm not in path.name and "信用卡重要資訊揭露" not in path.name:
                    continue
                resolved = path.resolve()
                if resolved in seen:
                    continue
                seen.add(resolved)
                out.append(path)
    return sorted(out, key=lambda path: local_zip_sort_key(path, roc_yyyymm))


def fetch_bank_bureau_month(y: int, m: int, options: FetchOptions) -> dict[str, Any]:
    expected_yyyymm = roc_to_ad_yyyymm(y, m)
    warnings: list[str] = []

    if options.explicit_zip_file is not None:
        payload = parse_bank_rows_from_zip_file(options.explicit_zip_file)
        if payload["ad_yyyymm"] != expected_yyyymm:
            raise ValueError(f"指定 ZIP 月份不符: 預期 {expected_yyyymm}，實際 {payload['ad_yyyymm']}")
        return payload

    if options.prefer_local_zip:
        for path in find_local_zip_candidates(y, m, options.zip_search_dirs):
            try:
                payload = parse_bank_rows_from_zip_bytes(path.read_bytes(), str(path))
            except Exception as exc:
                warnings.append(f"本機 ZIP 不可用，略過 {path}: {type(exc).__name__}: {exc}")
                continue
            if payload["ad_yyyymm"] != expected_yyyymm:
                warnings.append(f"本機 ZIP 月份不符，略過 {path}: 預期 {expected_yyyymm}，實際 {payload['ad_yyyymm']}")
                continue
            payload.update({
                "source_kind": "local_zip_auto",
                "source_url": None,
                "local_zip_path": str(path),
                "download_mode": None,
                "ssl_fallback_used": False,
                "warnings": warnings,
            })
            return payload

    url = build_zip_url(y, m)
    zip_bytes, meta = download_bytes(url, insecure=options.insecure)
    payload = parse_bank_rows_from_zip_bytes(zip_bytes, url)
    if payload["ad_yyyymm"] != expected_yyyymm:
        raise ValueError(f"銀行局月份不一致: 預期 {roc_month_text(y, m)}，實際 {payload['month']}，url={url}")
    if meta.warning:
        warnings.append(meta.warning)
    payload.update({
        "source_kind": "download",
        "source_url": url,
        "local_zip_path": None,
        "download_mode": meta.mode,
        "ssl_fallback_used": meta.ssl_fallback_used,
        "warnings": warnings,
    })
    return payload


MAIN_SHEET = "歷史資料(年+月)"
MARKET_TOTAL_ITEM = "市場總計(銀行局)"
BANK_BUREAU_ITEM = "銀行局"
JCIC_ITEM = "財團法人金融聯合徵信中心"
BLOCK_ITEMS = [BANK_NAMES[key] for key in BANK_ORDER] + [MARKET_TOTAL_ITEM, BANK_BUREAU_ITEM, JCIC_ITEM]
BLOCK_SIZE = len(BLOCK_ITEMS)

FIELD_ALIASES = {
    "yyyymm": ["YYYYMM"],
    "ad_year": ["年度"],
    "month_number": ["月份"],
    "rank": ["Rank"],
    "item": ["Item"],
    "circulating_cards": ["流通卡數"],
    "valid_cards": ["有效卡數"],
    "new_cards_this_month": ["當月發卡數"],
    "cancelled_cards_this_month": ["當月停卡數"],
    "signed_amount_million": ["當月簽帳金額"],
    "revolving_balance_million": ["循環信用餘額"],
    "installment_balance_not_yet_due_million": ["未到期分期付款餘額"],
    "cash_advance_amount_million": ["當月預借現金金額"],
    "overdue_3m_ratio_percent": [
        "逾期三個月以上比率",
        "逾期三個月比率",
        "逾期三個月以上帳款占應收帳款餘額（含催收款）之比率(%)",
        "逾期三個月以上帳款占應收帳款餘額(含催收款)之比率(%)",
        "逾期三個月以上帳款佔應收帳款餘額（含催收款）之比率(%)",
        "逾期三個月以上帳款佔應收帳款餘額(含催收款)之比率(%)",
    ],
    "overdue_6m_ratio_percent": [
        "逾期六個月以上比率",
        "逾期六個月比率",
        "逾期六個月以上帳款占應收帳款餘額（含催收款）之比率(%)",
        "逾期六個月以上帳款占應收帳款餘額(含催收款)之比率(%)",
        "逾期六個月以上帳款佔應收帳款餘額（含催收款）之比率(%)",
        "逾期六個月以上帳款佔應收帳款餘額(含催收款)之比率(%)",
    ],
    "allowance_coverage_ratio_percent": ["備抵呆帳提足率", "備抵呆帳提足率(%)"],
    "charge_off_amount_this_month_million": ["當月轉銷呆帳金額", "當月轉銷呆帳金額"],
    "charge_off_amount_ytd_million": ["當年度累計轉銷呆帳金額", "當年度累計轉銷呆帳金額"],
    "market_total": ["市場總計"],
    "top5_circulating_cards": ["TOP 5 流通卡"],
    "top10_circulating_cards": ["TOP 10 流通卡"],
    "top5_valid_cards": ["TOP 5 有效卡"],
    "top10_valid_cards": ["TOP 10 有效卡"],
    "top5_new_cards": ["TOP 5 當月發卡數"],
    "top10_new_cards": ["TOP 10 當月發卡數"],
    "top5_cancelled_cards": ["TOP 5 當月停卡數"],
    "top10_cancelled_cards": ["TOP 10 當月停卡數"],
    "top5_revolving_balance": ["TOP 5 循環信用餘額"],
    "top10_revolving_balance": ["TOP 10 循環信用餘額"],
    "top5_signed_amount": ["TOP 5 當月簽帳金額"],
    "top10_signed_amount": ["TOP 10 當月簽帳金額"],
}

BANK_WRITABLE_FIELDS = [
    "circulating_cards",
    "valid_cards",
    "new_cards_this_month",
    "cancelled_cards_this_month",
    "signed_amount_million",
    "revolving_balance_million",
    "installment_balance_not_yet_due_million",
    "cash_advance_amount_million",
    "overdue_3m_ratio_percent",
    "overdue_6m_ratio_percent",
    "allowance_coverage_ratio_percent",
    "charge_off_amount_this_month_million",
    "charge_off_amount_ytd_million",
]

BANK_BUREAU_FORMULA_FIELDS = [
    "top5_circulating_cards",
    "top10_circulating_cards",
    "top5_valid_cards",
    "top10_valid_cards",
    "top5_new_cards",
    "top10_new_cards",
    "top5_cancelled_cards",
    "top10_cancelled_cards",
    "top5_revolving_balance",
    "top10_revolving_balance",
    "top5_signed_amount",
    "top10_signed_amount",
]


def normalize_month_yyyymm(value: Any) -> int | None:
    if value is None:
        return None
    try:
        yyyymm = int(value)
    except Exception:
        text = str(value).strip()
        if re.fullmatch(r"20\d{4}", text):
            yyyymm = int(text)
        else:
            return None
    return yyyymm if 200001 <= yyyymm <= 299912 and 1 <= yyyymm % 100 <= 12 else None


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def get_main_sheet(wb):
    if MAIN_SHEET in wb.sheetnames:
        return wb[MAIN_SHEET]
    return wb[wb.sheetnames[0]]


def canonical_block_item(value: Any) -> str:
    return shared_canonical_block_item(
        value, special_items=[MARKET_TOTAL_ITEM, BANK_BUREAU_ITEM, JCIC_ITEM], bank_names=BANK_NAMES, bank_item_aliases=BANK_ITEM_ALIASES
    )


def normalize_item_labels(ws, index_map: dict[str, int]) -> dict[str, Any]:
    """把歷史資料與新寫入的 Item 全部統一成簡稱。"""
    item_col = index_map.get("item")
    if item_col is None:
        return {"changed_count": 0, "changed_rows": []}
    changed_rows: list[dict[str, Any]] = []
    for row_no in range(2, ws.max_row + 1):
        raw_value = ws.cell(row_no, item_col).value
        normalized = canonical_block_item(raw_value)
        raw_text = str(raw_value or "").strip()
        if normalized != raw_text:
            ws.cell(row_no, item_col).value = normalized
            changed_rows.append({"row": row_no, "from": raw_text, "to": normalized})
    return {"changed_count": len(changed_rows), "changed_rows": changed_rows}


def header_index_map(ws, field_aliases: dict[str, list[str]]) -> dict[str, int]:
    normalized_headers = {normalize_text(ws.cell(1, col).value): col for col in range(1, ws.max_column + 1)}
    out: dict[str, int] = {}
    for field, aliases in field_aliases.items():
        for alias in aliases:
            col = normalized_headers.get(normalize_text(alias))
            if col is not None:
                out[field] = col
                break
    return out


def require_headers(idx: dict[str, int], required: list[str], sheet_name: str) -> None:
    missing = [x for x in required if x not in idx]
    if missing:
        raise ValueError(f"{sheet_name} 缺少欄位: {missing}")


def find_month_block(ws, index_map: dict[str, int], yyyymm: int) -> dict[str, Any] | None:
    return shared_find_month_block(
        ws, index_map, yyyymm, normalize_yyyymm=normalize_month_yyyymm, block_items=BLOCK_ITEMS, canonicalize_item=canonical_block_item, block_label='月 block '
    )


def find_or_create_month_block(ws, index_map: dict[str, int], ad_year: int, month_number: int) -> dict[str, Any]:
    return shared_find_or_create_month_block(
        ws, index_map, ad_year=ad_year, month_number=month_number, normalize_yyyymm=normalize_month_yyyymm, block_items=BLOCK_ITEMS, bank_ranks=ITEM_RANKS, canonicalize_item=canonical_block_item, block_label='月 block '
    )


def refresh_block_metadata(ws, index_map: dict[str, int], block_info: dict[str, Any], ad_year: int, month_number: int) -> None:
    shared_refresh_block_metadata(
        ws, index_map, block_info, ad_year=ad_year, month_number=month_number, block_items=BLOCK_ITEMS, bank_ranks=ITEM_RANKS
    )


def write_bank_rows_to_block(ws, index_map: dict[str, int], block_info: dict[str, Any], bank_payload: dict[str, dict[str, Any]], overwrite: bool) -> dict[str, Any]:
    updated_rows = 0
    written: dict[str, dict[str, Any]] = {}
    skipped_existing: dict[str, dict[str, Any]] = {}

    for bank_key in BANK_ORDER:
        item = bank_payload.get(bank_key)
        if item is None:
            continue
        row_no = block_info["item_to_row"][BANK_NAMES[bank_key]]
        row_written: dict[str, Any] = {}
        row_skipped: dict[str, Any] = {}
        for field in BANK_WRITABLE_FIELDS:
            column = index_map.get(field)
            if column is None:
                continue
            cell = ws.cell(row_no, column)
            value = item.get(field)
            if overwrite or is_blank(cell.value):
                cell.value = value
                if value is not None:
                    apply_default_font(cell)
                if field in PERCENT_DECIMAL_FIELDS and value is not None:
                    apply_percent_number_format(cell)
                row_written[field] = value
            else:
                row_skipped[field] = cell.value
        if row_written:
            written[bank_key] = row_written
        if row_skipped:
            skipped_existing[bank_key] = row_skipped
        updated_rows += 1

    return {
        "bank_rows_updated": updated_rows,
        "written": written,
        "skipped_existing": skipped_existing,
    }


def has_any_blank_bank_metric(ws, index_map: dict[str, int], block_info: dict[str, Any]) -> bool:
    for bank_key in BANK_ORDER:
        row_no = block_info["item_to_row"][BANK_NAMES[bank_key]]
        for field in BANK_WRITABLE_FIELDS:
            column = index_map.get(field)
            if column is None:
                continue
            if is_blank(ws.cell(row_no, column).value):
                return True
    return False


def find_formula_source_row(ws, index_map: dict[str, int], target_yyyymm: int) -> int | None:
    candidates: list[tuple[int, int]] = []
    for row_no in range(2, ws.max_row + 1):
        if str(ws.cell(row_no, index_map["item"]).value or "").strip() != BANK_BUREAU_ITEM:
            continue
        yyyymm = normalize_month_yyyymm(ws.cell(row_no, index_map["yyyymm"]).value)
        if yyyymm is None or yyyymm == target_yyyymm:
            continue
        if any(
            ws.cell(row_no, index_map[field]).data_type == "f" and isinstance(ws.cell(row_no, index_map[field]).value, str)
            for field in BANK_BUREAU_FORMULA_FIELDS
            if field in index_map
        ):
            candidates.append((abs(yyyymm - target_yyyymm), row_no))
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0])
    return candidates[0][1]


def repair_bank_bureau_formulas(ws, index_map: dict[str, int], block_info: dict[str, Any], overwrite: bool) -> dict[str, Any]:
    target_row = block_info["item_to_row"][BANK_BUREAU_ITEM]
    source_row = find_formula_source_row(ws, index_map, block_info["yyyymm"])
    if source_row is None:
        return {
            "formula_source_row": None,
            "formulas_written": [],
            "formula_skipped_existing": [],
            "warning": "找不到可用的銀行局公式樣板列",
        }

    formulas_written: list[str] = []
    formula_skipped_existing: list[str] = []
    for field in BANK_BUREAU_FORMULA_FIELDS:
        column = index_map.get(field)
        if column is None:
            continue
        source = ws.cell(source_row, column)
        target = ws.cell(target_row, column)
        if not (source.data_type == "f" and isinstance(source.value, str) and source.value.strip()):
            continue
        if overwrite or is_blank(target.value):
            try:
                target.value = Translator(source.value, origin=source.coordinate).translate_formula(target.coordinate)
            except Exception:
                target.value = source.value
            if source.has_style:
                target._style = copy(source._style)
            target.font = copy(source.font)
            target.fill = copy(source.fill)
            target.border = copy(source.border)
            target.alignment = copy(source.alignment)
            target.protection = copy(source.protection)
            target.number_format = source.number_format
            formulas_written.append(field)
        else:
            formula_skipped_existing.append(field)

    return {
        "formula_source_row": source_row,
        "formulas_written": formulas_written,
        "formula_skipped_existing": formula_skipped_existing,
    }


def has_missing_bank_bureau_formula(ws, index_map: dict[str, int], block_info: dict[str, Any]) -> bool:
    row_no = block_info["item_to_row"][BANK_BUREAU_ITEM]
    for field in BANK_BUREAU_FORMULA_FIELDS:
        column = index_map.get(field)
        if column is None:
            continue
        value = ws.cell(row_no, column).value
        if is_blank(value):
            return True
    return False


def existing_months(ws, index_map: dict[str, int]) -> list[int]:
    months = sorted({
        yyyymm
        for row_no in range(2, ws.max_row + 1)
        for yyyymm in [normalize_month_yyyymm(ws.cell(row_no, index_map["yyyymm"]).value)]
        if yyyymm is not None
    }, reverse=True)
    return months


def iter_months_descending(newest_yyyymm: int, oldest_yyyymm: int):
    """由 newest 往前（含兩端）逐月 yield AD YYYYMM，正確跨年。"""
    y, m = divmod(int(newest_yyyymm), 100)
    oy, om = divmod(int(oldest_yyyymm), 100)
    while (y, m) >= (oy, om):
        yield y * 100 + m
        m -= 1
        if m == 0:
            m = 12
            y -= 1


def month_candidates_to_process(
    ws,
    index_map: dict[str, int],
    lookback_months: int,
    newest_yyyymm: int | None = None,
) -> list[int]:
    """
    自動模式要處理的月份（由新到舊）：
    1. 日曆缺口：newest 往前到 BACKFILL_START_YYYYMM 之間工作簿完全沒有的年月。
    2. 既有 block 中 10 家銀行指標有空白、或銀行局列 Top5/Top10 公式缺漏的年月。
    newest 優先用呼叫端傳入的資料月，未指定時用工作簿最新月；lookback_months 限制掃描月數。
    """
    existing = existing_months(ws, index_map)
    existing_set = set(existing)

    anchor_candidates = [v for v in (newest_yyyymm, existing[0] if existing else None) if v is not None]
    if not anchor_candidates:
        return []
    newest = max(anchor_candidates)
    if newest < BACKFILL_START_YYYYMM:
        return []
    oldest = BACKFILL_START_YYYYMM

    candidates: list[int] = []
    checked = 0
    for yyyymm in iter_months_descending(newest, oldest):
        if checked >= lookback_months:
            break
        checked += 1
        if yyyymm not in existing_set:
            candidates.append(yyyymm)
            continue
        block_info = find_month_block(ws, index_map, yyyymm)
        if block_info is None:
            continue
        if has_any_blank_bank_metric(ws, index_map, block_info) or has_missing_bank_bureau_formula(ws, index_map, block_info):
            candidates.append(yyyymm)
    return candidates


def set_full_calc_on_load(wb) -> None:
    if hasattr(wb, "calculation"):
        try:
            wb.calculation.fullCalcOnLoad = True
            wb.calculation.forceFullCalc = True
        except Exception:
            pass


def make_json_safe(value: Any) -> Any:
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, list):
        return [make_json_safe(item) for item in value]
    if isinstance(value, dict):
        return {str(k): make_json_safe(v) for k, v in value.items()}
    return value


def configure_utf8_stdio() -> None:
    """盡量把 stdout / stderr 設為 UTF-8，避免中文輸出亂碼。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            if hasattr(stream, "reconfigure"):
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="用銀行局 ZIP 回補『歷史資料(年+月)』中的各銀行舊月份空白（13 欄：卡數 4 欄＋金額 6 欄＋比率 3 欄），並修復銀行局 Top5/Top10 公式。")
    parser.add_argument("--workbook", default=str(DEFAULT_WORKBOOK), help="目標 Excel 檔，預設為 銀行局信用卡公開資料.xlsx")
    parser.add_argument("--month", default="", help="指定月份，例如 115年04月；若未指定，會做日曆缺口偵測 + 掃描既有 block 的空白月份")
    parser.add_argument("--newest-month", default="", help="日曆缺口偵測的最新錨點月份（民國年，如 115年05月）；未指定時用工作簿最新月。通常由主流程帶入 10 家銀行的資料月，用來往前找缺漏年月")
    parser.add_argument("--lookback-months", type=int, default=24, help="未指定 --month 時，從最新錨點往前掃描幾個『日曆月』（含缺漏月）。預設 24")
    parser.add_argument("--overwrite", action="store_true", help="若指定，會覆蓋既有資料與公式；預設只補空白")
    parser.add_argument("--insecure", action="store_true", help="直接略過 SSL 憑證驗證；若不指定，會先正常驗證，失敗時自動 fallback")
    parser.add_argument("--zip-file", help="指定本機銀行局 ZIP 檔；若提供，優先使用此檔案")
    parser.add_argument("--zip-search-dir", action="append", default=[], help="搜尋本機 ZIP 檔的目錄，可重複指定；預設會搜尋 . 與 input")
    parser.add_argument("--no-local-zip", action="store_true", help="不搜尋本機 ZIP，直接嘗試從網站下載")
    parser.add_argument("--strict-workbook", action="store_true", help="若指定工作簿不可用時，直接報錯，不自動改用其他 xlsx")
    return parser.parse_args()


def main() -> None:
    configure_utf8_stdio()
    args = parse_args()
    requested_workbook = Path(args.workbook)
    zip_search_dirs = [Path(path) for path in args.zip_search_dir] or DEFAULT_ZIP_SEARCH_DIRS

    explicit_zip = Path(args.zip_file) if args.zip_file else None
    if explicit_zip and not explicit_zip.exists():
        raise FileNotFoundError(f"指定的 ZIP 檔不存在: {explicit_zip}")

    workbook_resolution = resolve_workbook(requested_workbook, strict=args.strict_workbook)
    wb = load_workbook(workbook_resolution.path)
    ws = get_main_sheet(wb)
    index_map = header_index_map(ws, FIELD_ALIASES)
    require_headers(index_map, ["yyyymm", "ad_year", "month_number", "rank", "item", *BANK_WRITABLE_FIELDS], MAIN_SHEET)

    fetch_options = FetchOptions(
        insecure=args.insecure,
        prefer_local_zip=not args.no_local_zip,
        zip_search_dirs=zip_search_dirs,
        explicit_zip_file=explicit_zip,
    )

    if args.month:
        roc_y, roc_m = parse_roc_month(args.month)
        specified_yyyymm = roc_to_ad_yyyymm(roc_y, roc_m)
        if specified_yyyymm < BACKFILL_START_YYYYMM:
            raise ValueError(
                f"指定月份 {args.month}（{specified_yyyymm}）早於回補起始月 {BACKFILL_START_YYYYMM}，"
                f"依設定不處理 2026年01月 以前的資料"
            )
        months_to_process = [specified_yyyymm]
    else:
        newest_yyyymm = None
        if args.newest_month:
            newest_y, newest_m = parse_roc_month(args.newest_month)
            newest_yyyymm = roc_to_ad_yyyymm(newest_y, newest_m)
        months_to_process = month_candidates_to_process(ws, index_map, args.lookback_months, newest_yyyymm)

    cache: dict[int, dict[str, Any]] = {}
    results: list[dict[str, Any]] = []
    item_label_fix_summary: dict[str, Any] = {"changed_count": 0, "changed_rows": []}
    summary_warnings = list(workbook_resolution.warnings)

    for requested_yyyymm in months_to_process:
        roc_y, roc_m = ad_yyyymm_to_roc(requested_yyyymm)
        target_month = roc_month_text(roc_y, roc_m)
        ad_year = roc_y + 1911
        month_number = roc_m

        payload = cache.get(requested_yyyymm)
        if payload is None:
            payload = fetch_bank_bureau_month(roc_y, roc_m, fetch_options)
            cache[requested_yyyymm] = payload
        summary_warnings.extend(payload.get("warnings", []))

        if payload.get("missing_banks"):
            results.append({
                "month": target_month,
                "ad_yyyymm": requested_yyyymm,
                "action": "skip_missing_banks",
                "missing_banks": payload["missing_banks"],
                "source_url": payload.get("source_url"),
                "local_zip_path": payload.get("local_zip_path"),
            })
            continue

        block_info = find_or_create_month_block(ws, index_map, ad_year, month_number)
        refresh_block_metadata(ws, index_map, block_info, ad_year, month_number)
        bank_result = write_bank_rows_to_block(ws, index_map, block_info, payload["banks"], overwrite=args.overwrite)
        formula_result = repair_bank_bureau_formulas(ws, index_map, block_info, overwrite=args.overwrite)

        results.append({
            "month": target_month,
            "ad_yyyymm": requested_yyyymm,
            "block_created": block_info["created"],
            "block_start_row": block_info["start_row"],
            "action": (
                "specified_month_update" if args.month
                else ("backfill_missing_month" if block_info["created"] else "backfill_blank_bank_metrics")
            ),
            "source_url": payload.get("source_url"),
            "source_kind": payload.get("source_kind"),
            "local_zip_path": payload.get("local_zip_path"),
            "download_mode": payload.get("download_mode"),
            "ssl_fallback_used": payload.get("ssl_fallback_used", False),
            "matched_rows": payload.get("matched_rows", []),
            "bank_result": bank_result,
            "bank_bureau_formula_result": formula_result,
        })

    item_label_fix_summary = normalize_item_labels(ws, index_map)

    set_full_calc_on_load(wb)
    wb.save(workbook_resolution.path)

    payload = {
        "status": "success",
        "requested_workbook": str(requested_workbook),
        "selected_workbook": str(workbook_resolution.path),
        "workbook_fallback_used": workbook_resolution.fallback_used,
        "sheet": ws.title,
        "overwrite": args.overwrite,
        "month": args.month or None,
        "lookback_months": args.lookback_months,
        "zip_file": str(explicit_zip) if explicit_zip else None,
        "zip_search_dirs": [str(path) for path in zip_search_dirs],
        "summary_warnings": summary_warnings,
        "processed_month_count": len(results),
        "results": results,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "item_label_fix_summary": item_label_fix_summary,
    }
    print(json.dumps(make_json_safe(payload), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(json.dumps({"status": "error", "error_type": type(exc).__name__, "error": str(exc)}, ensure_ascii=False, indent=2), file=sys.stderr)
        raise
