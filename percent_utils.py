from __future__ import annotations

"""Shared percentage-field rules for credit card disclosure scripts.

Canonical rule:
- Python / JSON 內部值一律使用小數比率(decimal ratio)。
  例如 5% -> 0.05
- Excel 只靠 number_format 顯示成百分比。
- 若來源提供的是顯示用百分比數值（例如 5 或 "5%"），
  應在解析階段正規化成小數值，而不是在寫入工作簿時臨時 /100。
- 逾期三個月/六個月比率的來源值恆小於 1（例如 0.12 代表 0.12%），
  「絕對值 > 1 才 /100」的啟發式會誤判，這兩欄一律用
  force_percent_input=True 強制 /100。
"""

import re
from typing import Any

PERCENT_DECIMAL_FIELDS = {
    "overdue_3m_ratio_percent",
    "overdue_6m_ratio_percent",
    "allowance_coverage_ratio_percent",
}

# 來源一律是百分比顯示數字（0.12 代表 0.12%）、且典型值 < 1 的欄位：
# 解析時必須用 force_percent_input=True，不可依賴 abs > 1 的啟發式。
FORCE_PERCENT_INPUT_FIELDS = {
    "overdue_3m_ratio_percent",
    "overdue_6m_ratio_percent",
}

PERCENT_NUMBER_FORMAT = '0.00%'


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


def apply_percent_number_format(cell) -> None:
    cell.number_format = PERCENT_NUMBER_FORMAT


# ---- 寫入前合理性閘門 --------------------------------------------------------
# 小數比率的可信範圍 (nonzero_min, max)。超出代表單位錯誤（漏除或多除 100），寫入端應 fail-fast，
# 不要把值寫進工作簿。逾期三個月比率 100 倍會變成 ≥ 5%、六個月比率 ≥ 2%，備抵呆帳提足率 100 倍會變成 ≥ 5000%。
PERCENT_SANITY_BOUNDS: dict[str, tuple[float | None, float]] = {
    'overdue_3m_ratio_percent': (0.00005, 0.05),
    'overdue_6m_ratio_percent': (None, 0.02),
    'allowance_coverage_ratio_percent': (0.2, 50.0),
}

PERCENT_FIELD_LABELS = {
    'overdue_3m_ratio_percent': '逾期三個月以上比率',
    'overdue_6m_ratio_percent': '逾期六個月以上比率',
    'allowance_coverage_ratio_percent': '備抵呆帳提足率',
}


def percent_sanity_error(field: str, value: Any) -> str | None:
    """value 為小數比率；不在可信範圔時回傳原因文字，否則 None。非百分比欄位或空值一律視為正常。"""
    bounds = PERCENT_SANITY_BOUNDS.get(field)
    if bounds is None or value is None or value == '':
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return f'{PERCENT_FIELD_LABELS.get(field, field)} 不是數字：{value!r}'
    nonzero_min, maximum = bounds
    label = PERCENT_FIELD_LABELS.get(field, field)
    if number < 0:
        return f'{label} 為負值：{number}'
    if number > maximum:
        return f'{label}={number}（顯示 {number * 100:.2f}%）超過上限 {maximum * 100:.0f}%，疑似百分比數字未除以 100'
    if nonzero_min is not None and 0 < number < nonzero_min:
        return f'{label}={number}（顯示 {number * 100:.4f}%）低於下限 {nonzero_min * 100:.3f}%，疑似小數比率被重複除以 100'
    return None


def assert_percent_sane(field: str, value: Any, *, context: str = '') -> None:
    """寫入工作簿前呼叫；不合理就拋 ValueError，讓整支腳本在存檔前停下。"""
    reason = percent_sanity_error(field, value)
    if reason:
        prefix = f'拒絕寫入 {context}：' if context else '拒絕寫入：'
        raise ValueError(f'{prefix}{reason}。請先修正來源解析（force_percent_input 規則），不要直接改工作簿。')
