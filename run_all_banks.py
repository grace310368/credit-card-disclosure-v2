#!/usr/bin/env python3
from __future__ import annotations

import argparse
import importlib
import json
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

# 避免在唯讀執行沙箱中產生 __pycache__/*.pyc（會觸發 Refusing to overwrite）
sys.dont_write_bytecode = True

BASE_DIR = Path(__file__).resolve().parent

BANK_SCRIPTS = {
    "esun": "esun_creditcard_disclosure.py",
    "cathay": "cathay_creditcard_disclosure.py",
    "firstbank": "firstbank_creditcard_disclosure.py",
    "feib": "feib_creditcard_disclosure.py",
    "taishin": "taishin_creditcard_disclosure.py",
    "sinopac": "sinopac_creditcard_disclosure.py",
    "dbs": "dbs_creditcard_disclosure.py",
    "fubon": "fubon_creditcard_disclosure.py",
    "ubot": "ubot_creditcard_disclosure.py",
    "ctbc": "ctbc_creditcard_disclosure.py",
}

BANK_ORDER = [
    "ctbc",
    "fubon",
    "cathay",
    "esun",
    "taishin",
    "dbs",
    "ubot",
    "sinopac",
    "firstbank",
    "feib",
]

BANK_NAMES = {
    "esun": "玉山",
    "cathay": "國泰",
    "firstbank": "第一",
    "feib": "遠東",
    "taishin": "台新",
    "sinopac": "永豐",
    "dbs": "星展",
    "fubon": "富邦",
    "ubot": "聯邦",
    "ctbc": "中信",
}

BANK_NAME_TO_KEY = {value: key for key, value in BANK_NAMES.items()}

# 卡數類欄位：單位為「張」
CARD_METRIC_KEYS = [
    "circulating_cards",
    "valid_cards",
    "new_cards_this_month",
    "cancelled_cards_this_month",
]

# 金額類欄位：子腳本若回傳仟元欄位，主控腳本統一轉成百萬元欄位
AMOUNT_METRIC_CONVERSIONS = {
    "signed_amount_thousand": "signed_amount_million",
    "revolving_balance_thousand": "revolving_balance_million",
    "installment_balance_not_yet_due_thousand": "installment_balance_not_yet_due_million",
    "cash_advance_amount_thousand": "cash_advance_amount_million",
    "charge_off_amount_this_month_thousand": "charge_off_amount_this_month_million",
    "charge_off_amount_ytd_thousand": "charge_off_amount_ytd_million",
}

# 標準化後建議使用的單位
METRIC_UNITS = {
    "circulating_cards": "張",
    "valid_cards": "張",
    "new_cards_this_month": "張",
    "cancelled_cards_this_month": "張",
    "signed_amount_million": "百萬元",
    "revolving_balance_million": "百萬元",
    "installment_balance_not_yet_due_million": "百萬元",
    "cash_advance_amount_million": "百萬元",
    "charge_off_amount_this_month_million": "百萬元",
    "charge_off_amount_ytd_million": "百萬元",
    "overdue_3m_ratio_percent": "decimal_ratio",
    "overdue_6m_ratio_percent": "decimal_ratio",
    "allowance_coverage_ratio_percent": "decimal_ratio",
}

# 舊欄位來源單位，保留供追溯使用
SOURCE_METRIC_UNITS = {
    "signed_amount_thousand": "仟元",
    "revolving_balance_thousand": "仟元",
    "installment_balance_not_yet_due_thousand": "仟元",
    "cash_advance_amount_thousand": "仟元",
    "charge_off_amount_this_month_thousand": "仟元",
    "charge_off_amount_ytd_thousand": "仟元",
}


def configure_utf8_stdio() -> None:
    """盡量把 stdout / stderr 設為 UTF-8，避免中文輸出亂碼。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            if hasattr(stream, "reconfigure"):
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def zh_print(text: str) -> None:
    """在 Windows Big5 / cp950 console 下，盡量安全輸出中文。"""
    encoding = (getattr(sys.stdout, "encoding", "") or "").lower()
    if "utf" in encoding:
        print(text)
        return
    try:
        safe = text.encode("cp950", errors="replace").decode("cp950", errors="replace")
    except Exception:
        safe = text
    print(safe)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="批次抓取 10 家銀行信用卡財務業務資訊；只回傳 JSON 數據，不產出 Excel。金額統一轉成百萬元，百分比統一轉成小數比率。"
    )
    parser.add_argument(
        "--banks",
        default="all",
        help="all 或逗號分隔銀行代碼，例如 esun,cathay,dbs",
    )
    parser.add_argument(
        "--month",
        default="",
        help="可選，傳給各銀行腳本的月份，例如 115年05月 或 115/5；未指定則由各銀行腳本自行抓最新月份",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=120,
        help="每家銀行子程序逾時秒數，預設 120。注意：總執行時間約為 銀行數 × timeout，需低於執行工具上限（900 秒）；firstbank 內部重試最長約 170 秒，單獨重跑 firstbank 時建議 --timeout 200",
    )
    parser.add_argument(
        "--allow-insecure",
        action="store_true",
        help="轉傳 --allow-insecure 給 firstbank 子腳本（啟用其第三段 CERT_NONE fallback；其他銀行腳本會自動 fallback，不需此參數）",
    )
    parser.add_argument(
        "--collect-only",
        nargs="+",
        default=[],
        metavar="JSON",
        help="只彙整既有 JSON 結果，不重跑任何銀行腳本；可提供多個檔案（run_all_banks 彙整 payload 或單一銀行 stdout JSON），後面的檔案覆蓋前面同銀行的結果",
    )
    parser.add_argument(
        "--preflight",
        action="store_true",
        help="不執行抓取，只檢查執行環境與必要檔案：openpyxl/pypdf 可載入、各銀行腳本存在、指定 --month 時中信本機來源檔存在",
    )
    parser.add_argument(
        "--json-output",
        default="",
        help="可選：若指定路徑，會把本次完整 JSON 結果另存成檔案；未指定則只印到 stdout",
    )
    return parser.parse_args()


def resolve_banks(banks_arg: str) -> list[str]:
    if not banks_arg or banks_arg.strip().lower() == "all":
        return list(BANK_ORDER)

    requested = [item.strip().lower() for item in banks_arg.split(",") if item.strip()]
    invalid = [item for item in requested if item not in BANK_SCRIPTS]
    if invalid:
        raise SystemExit(f"不支援的銀行代碼: {', '.join(invalid)}")
    return requested


def resolve_existing_path(path_text: str) -> Path:
    raw = Path(path_text)
    candidates: list[Path] = []

    if raw.is_absolute():
        candidates.append(raw)
    else:
        candidates.append(Path.cwd() / raw)
        candidates.append(BASE_DIR / raw)
        if raw.parts and raw.parts[0] == BASE_DIR.name:
            candidates.append(BASE_DIR.parent / raw)

    seen: set[str] = set()
    for candidate in candidates:
        key = str(candidate.resolve(strict=False))
        if key in seen:
            continue
        seen.add(key)
        if candidate.exists():
            return candidate

    raise SystemExit(f"找不到檔案：{path_text}")


def is_normalized_result(result: dict[str, Any]) -> bool:
    """已標準化的紀錄 metrics_units 為 decimal_ratio；不可再跑 flatten_result_fields，否則比率會重複 /100。"""
    metrics_units = result.get("metrics_units")
    return isinstance(metrics_units, dict) and metrics_units.get("overdue_3m_ratio_percent") == "decimal_ratio"


def load_collect_results(path_text: str) -> list[dict[str, Any]]:
    """讀取 --collect-only 檔案：results list、run_all_banks 彙整 payload 或單一銀行 stdout JSON 皆可。"""
    path = resolve_existing_path(path_text)
    data = json.loads(path.read_text(encoding="utf-8-sig"))

    if isinstance(data, list):
        records = data
    elif isinstance(data, dict) and isinstance(data.get("results"), list):
        records = data["results"]
    elif isinstance(data, dict):
        records = [data]
    else:
        raise SystemExit(f"--collect-only 檔案格式不支援（最外層應為 list 或 dict）：{path}")

    if not all(isinstance(record, dict) for record in records):
        raise SystemExit(f"--collect-only 檔案含非物件的紀錄，無法解析：{path}")

    return [
        record if is_normalized_result(record) else flatten_result_fields(record)
        for record in records
    ]


def result_bank_key(result: dict[str, Any]) -> str:
    runner = result.get("_runner") or {}
    key = str(runner.get("bank_key") or "").strip().lower()
    if key in BANK_SCRIPTS:
        return key

    bank_name = str(result.get("bank") or "").strip()
    key = BANK_NAME_TO_KEY.get(bank_name, "")
    return key


def sort_results(results: list[dict[str, Any]]) -> list[dict[str, Any]]:
    order_index = {key: idx for idx, key in enumerate(BANK_ORDER)}

    def sort_key(result: dict[str, Any]) -> tuple[int, str]:
        key = result_bank_key(result)
        return (order_index.get(key, len(BANK_ORDER)), key or str(result.get("bank") or ""))

    return sorted(results, key=sort_key)


def merge_results(previous_results: list[dict[str, Any]], rerun_results: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rerun_map = {result_bank_key(item): item for item in rerun_results if result_bank_key(item)}
    merged_map: dict[str, dict[str, Any]] = {}
    extras: list[dict[str, Any]] = []

    for old in previous_results:
        key = result_bank_key(old)
        if key and key in rerun_map:
            merged_map[key] = rerun_map[key]
        elif key:
            merged_map[key] = old
        else:
            extras.append(old)

    for key, item in rerun_map.items():
        if key not in merged_map:
            merged_map[key] = item

    ordered = [merged_map[key] for key in BANK_ORDER if key in merged_map]
    remaining = [item for key, item in merged_map.items() if key not in BANK_ORDER]
    return ordered + sort_results(remaining + extras)


def decode_subprocess_output(raw: bytes | str | None) -> str:
    if raw is None:
        return ""
    if isinstance(raw, str):
        return raw
    for enc in ("utf-8", "utf-8-sig", "cp950", "big5"):
        try:
            return raw.decode(enc)
        except Exception:
            continue
    return raw.decode("utf-8", errors="replace")


def extract_json_from_output(text: str) -> dict[str, Any]:
    """
    從子程序輸出中取出 JSON。
    若子腳本只輸出 JSON，直接解析；若混有 log，會抓第一個 { 到最後一個 }。
    """
    payload = (text or "").strip()
    if not payload:
        raise ValueError("子程序沒有輸出 JSON")

    try:
        return json.loads(payload)
    except json.JSONDecodeError:
        start = payload.find("{")
        end = payload.rfind("}")
        if start >= 0 and end > start:
            return json.loads(payload[start : end + 1])
        raise


def parse_number(value: Any) -> int | float | None:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return value

    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        if "." in text:
            num = float(text)
            return int(num) if num.is_integer() else num
        return int(text)
    except ValueError:
        return None


# vendored from percent_utils.py — 修改時請與 percent_utils.py 同步
# 來源一律是百分比顯示數字（0.12 代表 0.12%）、且典型值 < 1 的欄位：
# 解析時必須用 force_percent_input=True，不可依賴 abs > 1 的啟發式。
# 與 percent_utils.PERCENT_SANITY_BOUNDS 同步維護：小數比率的可信範圍 (nonzero_min, max)
PERCENT_SANITY_BOUNDS = {
    "overdue_3m_ratio_percent": (0.00005, 0.05),
    "overdue_6m_ratio_percent": (None, 0.02),
    "allowance_coverage_ratio_percent": (0.2, 50.0),
}


def percent_sanity_error(field: str, value: Any) -> str | None:
    bounds = PERCENT_SANITY_BOUNDS.get(field)
    if bounds is None or value is None or value == "":
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return f"{field} 不是數字：{value!r}"
    nonzero_min, maximum = bounds
    if number < 0:
        return f"{field} 為負值：{number}"
    if number > maximum:
        return f"{field}={number}（顯示 {number * 100:.2f}%）超過上限，疑似百分比數字未除以 100"
    if nonzero_min is not None and 0 < number < nonzero_min:
        return f"{field}={number} 低於下限，疑似小數比率被重複除以 100"
    return None


FORCE_PERCENT_INPUT_FIELDS = {
    "overdue_3m_ratio_percent",
    "overdue_6m_ratio_percent",
}


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


def thousand_to_million(value: Any) -> int | float | None:
    """
    將仟元轉成百萬元。
    1 百萬元 = 1,000 仟元
    """
    num = parse_number(value)
    if num is None:
        return None

    converted = num / 1000
    return int(converted) if float(converted).is_integer() else converted


def normalize_metrics(result: dict[str, Any]) -> dict[str, Any]:
    """
    標準化 metrics：卡數轉數字（張）、仟元金額轉百萬元（保留原始 *_thousand）、
    百分比轉小數比率，並寫入 metrics_units / metrics_source_units。
    子腳本已回傳 *_million 時直接採用，不再除以 1000。
    """
    metrics = result.get("metrics")
    if not isinstance(metrics, dict):
        result["metrics"] = {}
        result["metrics_units"] = dict(METRIC_UNITS)
        result["metrics_source_units"] = dict(SOURCE_METRIC_UNITS)
        return result

    # 卡數類欄位：轉成數字，單位維持「張」
    for key in CARD_METRIC_KEYS:
        if key in metrics:
            metrics[key] = parse_number(metrics.get(key))

    source_metric_units = {}
    for unit_key in ("metric_units", "metrics_units"):
        units = result.get(unit_key)
        if isinstance(units, dict):
            source_metric_units.update(units)

    percent_metric_keys = [
        "overdue_3m_ratio_percent",
        "overdue_6m_ratio_percent",
        "allowance_coverage_ratio_percent",
    ]
    for key in percent_metric_keys:
        if key in metrics:
            source_unit = str(source_metric_units.get(key) or "").strip()
            if source_unit == "decimal_ratio":
                # 子腳本已在解析階段轉成小數比率（例如 ctbc），不可再除以 100
                metrics[key] = to_number(metrics.get(key))
            elif key in FORCE_PERCENT_INPUT_FIELDS:
                # 逾期比率來源恆為百分比顯示數字（0.12 代表 0.12%），一律 /100
                metrics[key] = normalize_percent_value(metrics.get(key), force_percent_input=True)
            else:
                metrics[key] = normalize_percent_value(metrics.get(key), assume_percent_input=source_unit == "%")
            # 正規化後仍不在可信範圍 → 來源格式可能變了；提早標記，讓步驟 4 的 partial_success_count 看得到
            reason = percent_sanity_error(key, metrics.get(key))
            if reason:
                result.setdefault("errors", []).append({"stage": "normalize_percent", "message": reason})
                if result.get("status") == "success":
                    result["status"] = "partial_success"

    # 金額類欄位：由仟元轉成百萬元，新增 *_million 欄位
    for source_key, target_key in AMOUNT_METRIC_CONVERSIONS.items():
        if source_key in metrics:
            metrics[source_key] = parse_number(metrics.get(source_key))
            metrics[target_key] = thousand_to_million(metrics.get(source_key))

    # 若子腳本已經回傳 million 欄位，也一併標準化成數字
    # 不再除以 1000，避免重複轉換。
    for target_key in AMOUNT_METRIC_CONVERSIONS.values():
        if target_key in metrics:
            metrics[target_key] = parse_number(metrics.get(target_key))

    # 中信轉帳卡簽帳金額：若來源為仟元，也同步補出百萬元欄位。
    if "debit_card_signed_amount_thousand" in metrics:
        metrics["debit_card_signed_amount_thousand"] = parse_number(metrics.get("debit_card_signed_amount_thousand"))
        metrics["debit_card_signed_amount_million"] = thousand_to_million(metrics.get("debit_card_signed_amount_thousand"))
    if "debit_card_signed_amount_million" in metrics:
        metrics["debit_card_signed_amount_million"] = parse_number(metrics.get("debit_card_signed_amount_million"))

    result["metrics"] = metrics
    result["metrics_units"] = dict(METRIC_UNITS)
    result["metrics_source_units"] = dict(SOURCE_METRIC_UNITS)

    return result


def flatten_result_fields(result: dict[str, Any]) -> dict[str, Any]:
    """補出 update_credit_card_workbook.py 直接讀取的扁平欄位；metrics/source/errors/_runner 原樣保留。"""
    result = normalize_metrics(result)

    metrics = result.get("metrics") if isinstance(result.get("metrics"), dict) else {}
    source = result.get("source") if isinstance(result.get("source"), dict) else {}
    errors = result.get("errors") if isinstance(result.get("errors"), list) else []

    bank_key = str(result.get("bank_key") or "").strip().lower()
    bank_name = str(result.get("bank_name") or result.get("bank") or BANK_NAMES.get(bank_key, "")).strip()

    result["bank_key"] = bank_key or result.get("bank_key")
    result["bank_name"] = bank_name
    if not result.get("data_month") and result.get("month"):
        result["data_month"] = result["month"]

    # 讓 update_credit_card_workbook.py 的 first_present(...) 可直接取值。
    flattened_metric_keys = [
        "circulating_cards",
        "valid_cards",
        "new_cards_this_month",
        "cancelled_cards_this_month",
        "signed_amount_thousand",
        "signed_amount_million",
        "revolving_balance_thousand",
        "revolving_balance_million",
        "installment_balance_not_yet_due_thousand",
        "installment_balance_not_yet_due_million",
        "cash_advance_amount_thousand",
        "cash_advance_amount_million",
        "overdue_3m_ratio_percent",
        "overdue_6m_ratio_percent",
        "allowance_coverage_ratio_percent",
        "charge_off_amount_this_month_thousand",
        "charge_off_amount_this_month_million",
        "charge_off_amount_ytd_thousand",
        "charge_off_amount_ytd_million",
        "debit_card_signed_amount_thousand",
        "debit_card_signed_amount_million",
    ]
    for key in flattened_metric_keys:
        if key in metrics:
            result[key] = metrics.get(key)

    result["source_type"] = source.get("source_type", result.get("source_type"))
    result["source_url"] = source.get("source_url", result.get("source_url"))
    result["error_count"] = len(errors)

    if errors and not result.get("failure_reason"):
        messages = [str(item.get("message") or "").strip() for item in errors if isinstance(item, dict)]
        messages = [msg for msg in messages if msg]
        if messages:
            result["failure_reason"] = "; ".join(messages)

    return result


def run_bank(bank_key: str, month: str, timeout: int, allow_insecure: bool = False) -> dict[str, Any]:
    script_name = BANK_SCRIPTS[bank_key]
    script_path = BASE_DIR / script_name
    cmd = [sys.executable, "-B", str(script_path)]
    if month:
        cmd.extend(["--month", month])
    if bank_key == "firstbank" and allow_insecure:
        cmd.append("--allow-insecure")

    started_at = datetime.now().isoformat(timespec="seconds")

    if not script_path.exists():
        return flatten_result_fields({
            "bank": BANK_NAMES.get(bank_key, bank_key),
            "bank_key": bank_key,
            "status": "failed",
            "metrics": {},
            "metrics_units": dict(METRIC_UNITS),
            "metrics_source_units": dict(SOURCE_METRIC_UNITS),
            "errors": [{"stage": "runner", "message": f"找不到子腳本：{script_name}"}],
            "notes": "主控腳本找不到對應的銀行子腳本。",
            "_runner": {
                "bank_key": bank_key,
                "script": script_name,
                "returncode": -3,
                "started_at": started_at,
                "finished_at": datetime.now().isoformat(timespec="seconds"),
                "stderr": f"missing script: {script_path}",
            },
        })

    try:
        proc = subprocess.run(
            cmd,
            cwd=BASE_DIR,
            capture_output=True,
            text=False,
            timeout=timeout,
        )
        stdout_text = decode_subprocess_output(proc.stdout)
        stderr_text = decode_subprocess_output(proc.stderr)
        raw = stdout_text.strip() or stderr_text.strip()
        payload = extract_json_from_output(raw)
        payload = flatten_result_fields(payload)
        payload["_runner"] = {
            "bank_key": bank_key,
            "script": script_name,
            "returncode": proc.returncode,
            "started_at": started_at,
            "finished_at": datetime.now().isoformat(timespec="seconds"),
            "stderr": stderr_text.strip(),
        }
        return payload
    except subprocess.TimeoutExpired:
        return flatten_result_fields({
            "bank": BANK_NAMES.get(bank_key, bank_key),
            "bank_key": bank_key,
            "status": "failed",
            "metrics": {},
            "metrics_units": dict(METRIC_UNITS),
            "metrics_source_units": dict(SOURCE_METRIC_UNITS),
            "errors": [{"stage": "runner", "message": f"timeout after {timeout} seconds"}],
            "notes": "主控腳本執行子程序逾時。",
            "_runner": {
                "bank_key": bank_key,
                "script": script_name,
                "returncode": -1,
                "started_at": started_at,
                "finished_at": datetime.now().isoformat(timespec="seconds"),
                "stderr": "timeout",
            },
        })
    except Exception as exc:
        return flatten_result_fields({
            "bank": BANK_NAMES.get(bank_key, bank_key),
            "bank_key": bank_key,
            "status": "failed",
            "metrics": {},
            "metrics_units": dict(METRIC_UNITS),
            "metrics_source_units": dict(SOURCE_METRIC_UNITS),
            "errors": [{"stage": "runner", "message": str(exc)}],
            "notes": "主控腳本無法解析子程序結果。",
            "_runner": {
                "bank_key": bank_key,
                "script": script_name,
                "returncode": -2,
                "started_at": started_at,
                "finished_at": datetime.now().isoformat(timespec="seconds"),
                "stderr": str(exc),
            },
        })


def write_json_if_requested(path_text: str, payload: dict[str, Any]) -> str:
    """若使用者指定 --json-output，將完整 payload 存成 JSON；否則不產生任何檔案。"""
    if not path_text:
        return ""

    path = Path(path_text)
    if not path.is_absolute():
        path = BASE_DIR / path
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return str(path)


def roc_month_to_ad_yyyymm(month_text: str) -> str:
    """把民國月份（例如 115年05月、115/5）轉成西元 yyyymm（例如 202605）。"""
    match = re.search(r"(\d{2,3})\s*[年/\-.]\s*(\d{1,2})", str(month_text or ""))
    if not match:
        return ""
    return f"{int(match.group(1)) + 1911}{int(match.group(2)):02d}"


def run_preflight(month: str) -> int:
    """不執行抓取，只檢查執行環境與必要檔案，並以 JSON 回報結果（僅回報、不視為失敗）。"""
    checks: list[dict[str, Any]] = []

    for module_name in ("openpyxl", "pypdf"):
        try:
            importlib.import_module(module_name)
            checks.append({"name": f"import:{module_name}", "ok": True, "detail": "可正常載入"})
        except Exception as exc:
            checks.append({"name": f"import:{module_name}", "ok": False, "detail": f"無法載入：{exc}"})

    for bank_key in BANK_ORDER:
        script_path = BASE_DIR / BANK_SCRIPTS[bank_key]
        checks.append({
            "name": f"script:{bank_key}",
            "ok": script_path.exists(),
            "detail": str(script_path),
        })

    if month:
        yyyymm = roc_month_to_ad_yyyymm(month)
        if not yyyymm:
            checks.append({
                "name": "ctbc_source_file",
                "ok": False,
                "detail": f"無法解析月份：{month}",
            })
        else:
            search_dirs = [BASE_DIR.parent / "input", BASE_DIR / "input", BASE_DIR.parent, BASE_DIR]
            matches: list[str] = []
            for directory in search_dirs:
                if directory.is_dir():
                    matches.extend(str(path) for path in sorted(directory.glob(f"*中信*{yyyymm}*.xls*")))
            checks.append({
                "name": "ctbc_source_file",
                "ok": bool(matches),
                "detail": (
                    "; ".join(matches)
                    if matches
                    else f"找不到 *中信*{yyyymm}*.xls*；搜尋目錄：{'、'.join(str(d) for d in search_dirs)}"
                ) + "（僅淺層搜尋，近似 ctbc 子腳本本身的搜尋方式）",
            })

    payload = {
        "status": "success",
        "preflight": True,
        "preflight_ok": all(check["ok"] for check in checks),
        "checks": checks,
    }
    zh_print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


def main() -> int:
    configure_utf8_stdio()
    args = parse_args()

    if args.preflight:
        return run_preflight(args.month)

    collect_only_mode = bool(args.collect_only)
    missing_banks: list[str] = []

    if collect_only_mode:
        merged_results: list[dict[str, Any]] = []
        for path_text in args.collect_only:
            merged_results = merge_results(merged_results, load_collect_results(path_text))
        results = sort_results(merged_results)
        rerun_results = []
        banks = [key for key in (result_bank_key(item) for item in results) if key]
        missing_banks = [key for key in BANK_ORDER if key not in banks]
    else:
        banks = resolve_banks(args.banks)
        rerun_results = [run_bank(bank, args.month, args.timeout, args.allow_insecure) for bank in banks]
        results = sort_results(rerun_results)

    success_count = sum(1 for item in results if item.get("status") == "success")
    partial_count = sum(1 for item in results if item.get("status") == "partial_success")
    failed_count = sum(1 for item in results if item.get("status") == "failed")

    output_payload: dict[str, Any] = {
        "status": "success",
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "banks": banks,
        "rerun_count": len(rerun_results),
        "output_result_count": len(results),
        "success_count": success_count,
        "partial_success_count": partial_count,
        "failed_count": failed_count,
        "amount_unit_standard": "百萬元",
        "json_output": "",
    }
    if collect_only_mode:
        output_payload["collect_only"] = True
        output_payload["source_files"] = list(args.collect_only)
        output_payload["missing_banks"] = missing_banks
    output_payload["results"] = results

    saved_path = write_json_if_requested(args.json_output, output_payload)
    if saved_path:
        output_payload["json_output"] = saved_path
        # 重新寫一次，讓檔案內也包含 json_output 路徑。
        write_json_if_requested(args.json_output, output_payload)

    zh_print(json.dumps(output_payload, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
