from __future__ import annotations

from copy import copy
from typing import Any, Callable

DEFAULT_FONT_NAME = '微軟正黑體'


def apply_default_font(cell, font_name: str = DEFAULT_FONT_NAME) -> None:
    new_font = copy(cell.font)
    new_font.name = font_name
    cell.font = new_font


def apply_default_font_to_row(ws, row_no: int, *, start_column: int = 1, end_column: int | None = None, font_name: str = DEFAULT_FONT_NAME) -> None:
    if end_column is None:
        end_column = ws.max_column
    for column in range(start_column, end_column + 1):
        cell = ws.cell(row_no, column)
        apply_default_font(cell, font_name=font_name)


def canonical_block_item(
    value: Any,
    *,
    special_items: list[str] | set[str] | tuple[str, ...],
    bank_names: dict[str, str],
    bank_item_aliases: dict[str, list[str]],
) -> str:
    text = str(value or "").strip()
    if text in set(special_items):
        return text
    for bank_key, aliases in bank_item_aliases.items():
        if text in aliases:
            return bank_names[bank_key]
    return text


def copy_row_style(ws, src_row: int, dst_row: int) -> None:
    for column in range(1, ws.max_column + 1):
        src = ws.cell(src_row, column)
        dst = ws.cell(dst_row, column)
        if src.has_style:
            dst._style = copy(src._style)
        dst.font = copy(src.font)
        apply_default_font(dst)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.protection = copy(src.protection)
        dst.number_format = src.number_format
    if src_row in ws.row_dimensions:
        ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height
        ws.row_dimensions[dst_row].hidden = ws.row_dimensions[src_row].hidden


def set_row_value_if_present(ws, row_no: int, index_map: dict[str, int], field: str, value: Any) -> None:
    column = index_map.get(field)
    if column is not None:
        ws.cell(row_no, column).value = value


def set_block_row_metadata(
    ws,
    row_no: int,
    index_map: dict[str, int],
    *,
    yyyymm: int,
    ad_year: int,
    month_number: int | str,
    rank: int | None,
    item: str,
    year_field: str = 'ad_year',
    month_field: str = 'month_number',
) -> None:
    set_row_value_if_present(ws, row_no, index_map, 'yyyymm', yyyymm)
    set_row_value_if_present(ws, row_no, index_map, year_field, ad_year)
    set_row_value_if_present(ws, row_no, index_map, month_field, month_number)
    set_row_value_if_present(ws, row_no, index_map, 'rank', rank)
    set_row_value_if_present(ws, row_no, index_map, 'item', item)


def validate_block_items(
    ws,
    index_map: dict[str, int],
    rows: list[int],
    yyyymm: int,
    *,
    block_items: list[str],
    canonicalize_item: Callable[[Any], str],
    block_label: str = 'block',
) -> dict[str, int]:
    if len(rows) != len(block_items):
        raise ValueError(f'YYYYMM={yyyymm} 的{block_label}列數不是 {len(block_items)}：實際 {len(rows)} 列')
    expected_rows = list(range(rows[0], rows[0] + len(block_items)))
    if rows != expected_rows:
        raise ValueError(f'YYYYMM={yyyymm} 的{block_label}不是連續 {len(block_items)} 列：{rows}')

    item_to_row: dict[str, int] = {}
    for offset, expected_item in enumerate(block_items):
        row_no = rows[offset]
        actual_item = canonicalize_item(ws.cell(row_no, index_map['item']).value)
        if actual_item != expected_item:
            raise ValueError(
                f"YYYYMM={yyyymm} 的{block_label}Item 不符合固定模板：row {row_no} 預期 {expected_item!r}，實際 {ws.cell(row_no, index_map['item']).value!r}"
            )
        item_to_row[expected_item] = row_no
    return item_to_row


def find_month_block(
    ws,
    index_map: dict[str, int],
    yyyymm: int,
    *,
    normalize_yyyymm: Callable[[Any], int | None],
    block_items: list[str],
    canonicalize_item: Callable[[Any], str],
    block_label: str = 'block',
) -> dict[str, Any] | None:
    rows = [
        row_no
        for row_no in range(2, ws.max_row + 1)
        if normalize_yyyymm(ws.cell(row_no, index_map['yyyymm']).value) == yyyymm
    ]
    if not rows:
        return None
    item_to_row = validate_block_items(
        ws,
        index_map,
        rows,
        yyyymm,
        block_items=block_items,
        canonicalize_item=canonicalize_item,
        block_label=block_label,
    )
    return {
        'yyyymm': yyyymm,
        'start_row': rows[0],
        'rows': rows,
        'item_to_row': item_to_row,
        'created': False,
    }


def append_month_block(
    ws,
    index_map: dict[str, int],
    *,
    ad_year: int,
    month_number: int,
    block_items: list[str],
    bank_ranks: dict[str, int],
    year_field: str = 'ad_year',
    month_field: str = 'month_number',
) -> dict[str, Any]:
    old_max_row = ws.max_row
    start_row = old_max_row + 1
    yyyymm = ad_year * 100 + month_number
    style_template_start = old_max_row - len(block_items) + 1 if old_max_row >= len(block_items) + 1 else None
    item_to_row: dict[str, int] = {}
    rows: list[int] = []

    for offset, item in enumerate(block_items):
        row_no = start_row + offset
        if style_template_start and style_template_start >= 2:
            copy_row_style(ws, style_template_start + offset, row_no)
        rank = bank_ranks.get(item)
        set_block_row_metadata(
            ws,
            row_no,
            index_map,
            yyyymm=yyyymm,
            ad_year=ad_year,
            month_number=month_number,
            rank=rank,
            item=item,
            year_field=year_field,
            month_field=month_field,
        )
        item_to_row[item] = row_no
        rows.append(row_no)

    return {
        'yyyymm': yyyymm,
        'start_row': start_row,
        'rows': rows,
        'item_to_row': item_to_row,
        'created': True,
    }


def find_or_create_month_block(
    ws,
    index_map: dict[str, int],
    *,
    ad_year: int,
    month_number: int,
    normalize_yyyymm: Callable[[Any], int | None],
    block_items: list[str],
    bank_ranks: dict[str, int],
    canonicalize_item: Callable[[Any], str],
    block_label: str = 'block',
    year_field: str = 'ad_year',
    month_field: str = 'month_number',
) -> dict[str, Any]:
    yyyymm = ad_year * 100 + month_number
    found = find_month_block(
        ws,
        index_map,
        yyyymm,
        normalize_yyyymm=normalize_yyyymm,
        block_items=block_items,
        canonicalize_item=canonicalize_item,
        block_label=block_label,
    )
    if found:
        return found
    return append_month_block(
        ws,
        index_map,
        ad_year=ad_year,
        month_number=month_number,
        block_items=block_items,
        bank_ranks=bank_ranks,
        year_field=year_field,
        month_field=month_field,
    )


def refresh_block_metadata(
    ws,
    index_map: dict[str, int],
    block_info: dict[str, Any],
    *,
    ad_year: int,
    month_number: int,
    block_items: list[str],
    bank_ranks: dict[str, int],
    year_field: str = 'ad_year',
    month_field: str = 'month_number',
) -> None:
    yyyymm = ad_year * 100 + month_number
    for item in block_items:
        row_no = block_info['item_to_row'][item]
        rank = bank_ranks.get(item)
        set_block_row_metadata(
            ws,
            row_no,
            index_map,
            yyyymm=yyyymm,
            ad_year=ad_year,
            month_number=month_number,
            rank=rank,
            item=item,
            year_field=year_field,
            month_field=month_field,
        )
