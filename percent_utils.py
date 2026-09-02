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
