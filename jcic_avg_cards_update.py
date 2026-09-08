from __future__ import annotations

import argparse
import csv
import io
import json
import re
import ssl
import sys
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from bank_aliases import MARKET_TOTAL_ITEM
from workbook_block_helpers import apply_default_font, canonical_block_item

# 避免在唯讀執行沙箱中產生 __pycache__/*.pyc（會觸發 Refusing to overwrite）
sys.dont_write_bytecode = True

DEFAULT_WORKBOOK = Path('銀行局信用卡公開資料.xlsx')
DEFAULT_ENTRY_URL = 'https://www.jcic.org.tw/main_ch/download_page.aspx?uid=213&pid=190'
DEFAULT_SOURCE_URL = 'https://www.jcic.org.tw/main_ch/fileRename/fileRename.aspx?uid=213&fid=910&kid=4'
TARGET_LABEL = '平均每人持卡張數'
SOURCE_HEADER_KEY = '信用卡平均每戶持卡張數'
DEFAULT_SHEET_NAME = '歷史資料(年+月)'
TARGET_ITEM = MARKET_TOTAL_ITEM  # 平均每人持卡張數放在市場總計列
TARGET_FID = '910'

FIELD_ALIASES = {
    'yyyymm': ['YYYYMM'],
    'ad_year': ['年度'],
    'month_number': ['月份'],
    'rank': ['Rank', '排序編號'],
    'item': ['Item'],
    'avg_cards_per_person': ['平均每人持卡張數'],
}

SSL_FALLBACK_USED = False


# 與 run_all_banks.py 的 configure_utf8_stdio 保持同步
def configure_utf8_stdio() -> None:
    """盡量把 stdout / stderr 設為 UTF-8，避免中文輸出亂碼。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            if hasattr(stream, "reconfigure"):
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description='抓取 JCIC 平均每人持卡張數 CSV，依 base month 往前回補 銀行局信用卡公開資料.xlsx 的空白欄位。'
    )
    p.add_argument('--workbook', default=str(DEFAULT_WORKBOOK), help='目標 Excel 檔，預設為 銀行局信用卡公開資料.xlsx')
    p.add_argument('--sheet', default=DEFAULT_SHEET_NAME, help='目標工作表名稱，預設為 歷史資料(年+月)')
    p.add_argument('--entry-url', default=DEFAULT_ENTRY_URL, help='JCIC 下載頁入口，用於解析實際下載連結')
    p.add_argument('--source-url', default=DEFAULT_SOURCE_URL, help='已知可用的 JCIC CSV 下載連結；若入口頁解析失敗，會回退使用此連結')
    p.add_argument('--base-month', required=True, help='以指定月份為基期，往前檢查 --backfill-months 個月份的空白/缺列後回補，例如 115年05月')
    p.add_argument('--backfill-months', type=int, default=24, help='搭配 --base-month 使用，從基期月份起最多往前檢查幾個月份（含基期），預設 24')
    p.add_argument('--label', default=TARGET_LABEL, help='目標欄位名稱，預設為 平均每人持卡張數')
    p.add_argument('--overwrite', action='store_true', help='若指定，會覆蓋既有資料；預設只補空白')
    p.add_argument('--insecure', action='store_true', help='若公司環境 SSL 憑證驗證失敗，可加此參數略過憑證驗證')
    return p.parse_args()


def get_ssl_context(insecure: bool = False) -> ssl.SSLContext:
    ctx = ssl.create_default_context()
    if insecure:
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
    return ctx


# 與 run_market_total.py 的 is_ssl_verification_error 保持同步
def is_ssl_verification_error(exc: Exception) -> bool:
    if isinstance(exc, ssl.SSLCertVerificationError):
        return True
    if isinstance(exc, ssl.SSLError):
        return 'CERTIFICATE_VERIFY_FAILED' in str(exc).upper()
    if isinstance(exc, urllib.error.URLError):
        reason = getattr(exc, 'reason', None)
        if isinstance(reason, Exception):
            return is_ssl_verification_error(reason)
    return 'CERTIFICATE_VERIFY_FAILED' in str(exc).upper()


def download_bytes(url: str, insecure: bool = False, timeout: int = 60) -> bytes:
    global SSL_FALLBACK_USED
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    if insecure:
        with urllib.request.urlopen(req, context=get_ssl_context(True), timeout=timeout) as resp:
            return resp.read()
    try:
        with urllib.request.urlopen(req, context=get_ssl_context(False), timeout=timeout) as resp:
            return resp.read()
    except Exception as exc:
        if not is_ssl_verification_error(exc):
            raise
        SSL_FALLBACK_USED = True
        with urllib.request.urlopen(req, context=get_ssl_context(True), timeout=timeout) as resp:
            return resp.read()


def decode_text(raw: bytes) -> str:
    for enc in ('utf-8-sig', 'utf-8', 'cp950', 'big5'):
        try:
            return raw.decode(enc)
        except Exception:
            continue
    return raw.decode('utf-8', errors='replace')


def download_text(url: str, insecure: bool = False, timeout: int = 60) -> str:
    return decode_text(download_bytes(url, insecure=insecure, timeout=timeout))


def normalize_text(text: Any) -> str:
    return re.sub(r'\s+', '', str(text or '').replace('\u3000', ''))


def parse_roc_month(text: str) -> tuple[int, int]:
    m = re.search(r'(\d{3})\D*(\d{1,2})', str(text))
    if not m:
        raise ValueError(f'無法解析民國月份: {text}')
    y, mm = int(m.group(1)), int(m.group(2))
    if not 1 <= mm <= 12:
        raise ValueError(f'月份不合法: {text}')
    return y, mm


def roc_month_text(y: int, m: int) -> str:
    return f'{y}年{m:02d}月'


def roc_to_ad_yyyymm(y: int, m: int) -> int:
    return (y + 1911) * 100 + m


def shift_roc_month(y: int, m: int, delta: int) -> tuple[int, int]:
    total = (y + 1911) * 12 + (m - 1) + delta
    ad_y, month_index = divmod(total, 12)
    return ad_y - 1911, month_index + 1


def month_window_from_base(base_month: str, count: int) -> list[dict[str, Any]]:
    y, m = parse_roc_month(base_month)
    total = max(0, int(count))
    out = []
    for offset in range(total):
        yy, mm = shift_roc_month(y, m, -offset)
        out.append({
            'month': roc_month_text(yy, mm),
            'ad_yyyymm': roc_to_ad_yyyymm(yy, mm),
        })
    return out


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == '')


def build_index_map(ws) -> dict[str, int]:
    normalized_headers = {normalize_text(ws.cell(1, col).value): col for col in range(1, ws.max_column + 1)}
    index_map: dict[str, int] = {}
    for field, aliases in FIELD_ALIASES.items():
        for alias in aliases:
            col = normalized_headers.get(normalize_text(alias))
            if col is not None:
                index_map[field] = col
                break
    required = ['yyyymm', 'ad_year', 'month_number', 'rank', 'item', 'avg_cards_per_person']
    missing = [field for field in required if field not in index_map]
    if missing:
        raise ValueError(f'工作表 {ws.title} 缺少必要欄位: {missing}')
    return index_map


def build_month_item_rows(ws, index_map: dict[str, int]) -> dict[tuple[int, str], int]:
    mapping: dict[tuple[int, str], int] = {}
    yyyymm_col = index_map['yyyymm']
    item_col = index_map['item']
    for row_no in range(2, ws.max_row + 1):
        yyyymm = ws.cell(row_no, yyyymm_col).value
        item = canonical_block_item(ws.cell(row_no, item_col).value)
        try:
            yyyymm_int = int(yyyymm)
        except Exception:
            continue
        if not item:
            continue
        mapping[(yyyymm_int, item)] = row_no
    return mapping


def require_target_row(row_map: dict[tuple[int, str], int], ad_yyyymm: int) -> int:
    row_no = row_map.get((ad_yyyymm, TARGET_ITEM))
    if row_no is None:
        raise ValueError(f'YYYYMM={ad_yyyymm} 在工作簿中沒有月 block，拒絕單獨建立列；請先由 update_credit_card_workbook.py 或 run_bank_bureau_bank_backfill.py 建立該月 11 列 block')
    return row_no


def resolve_source_url(entry_url: str, fallback_url: str, insecure: bool = False) -> tuple[str, str]:
    html = download_text(entry_url, insecure=insecure)

    fid_match = re.search(
        rf'href=["\'](?P<href>[^"\']*fileRename/fileRename\.aspx\?[^"\']*fid={TARGET_FID}[^"\']*)["\']',
        html,
        flags=re.I,
    )
    if fid_match:
        full = urllib.parse.urljoin(entry_url, fid_match.group('href').replace('&amp;', '&'))
        return full, 'entry_page_fid_910_match'

    direct_candidates = re.findall(r'fileRename/fileRename\.aspx\?[^"\'<>\s]+', html, flags=re.I)
    for candidate in direct_candidates:
        full = urllib.parse.urljoin(entry_url, candidate.replace('&amp;', '&'))
        if 'uid=213' in full and f'fid={TARGET_FID}' in full:
            return full, 'entry_page_direct_match'

    for m in re.finditer(r'<a\b[^>]*href=["\']([^"\']+)["\'][^>]*>([\s\S]*?)</a>', html, flags=re.I):
        href, inner = m.group(1), m.group(2)
        label = re.sub(r'<[^>]+>', ' ', inner)
        label = re.sub(r'\s+', ' ', label).strip()
        if '6-16' in label or SOURCE_HEADER_KEY in label or '平均每戶持卡張數' in label or '平均每人持卡張數' in label:
            full = urllib.parse.urljoin(entry_url, href.replace('&amp;', '&'))
            return full, 'entry_page_label_match'

    return fallback_url, 'fallback_source_url'


def extract_csv_text(payload: bytes) -> str:
    if payload.startswith(b'PK'):
        with zipfile.ZipFile(io.BytesIO(payload)) as zf:
            csv_names = [name for name in zf.namelist() if name.lower().endswith('.csv')]
            if not csv_names:
                raise ValueError('下載內容是 ZIP，但裡面找不到 CSV')
            return decode_text(zf.read(csv_names[0]))
    return decode_text(payload)


def fetch_series(source_url: str, insecure: bool = False) -> tuple[str, dict[str, Any]]:
    payload = download_bytes(source_url, insecure=insecure)
    csv_text = extract_csv_text(payload)
    stripped = csv_text.lstrip('\ufeff\r\n\t ')
    if stripped.startswith('<'):
        preview = stripped[:200].replace('\n', ' ')
        raise ValueError(f'下載內容疑似 HTML，非 CSV: {preview}')
    return csv_text, parse_csv_series(csv_text)


def parse_csv_series(csv_text: str) -> dict[str, Any]:
    normalized_text = csv_text.replace('\r\n', '\n').replace('\r', '\n')
    rows = list(csv.reader(io.StringIO(normalized_text), skipinitialspace=True))
    if not rows:
        raise ValueError('JCIC CSV 內容為空')

    header = rows[0]
    if len(header) < 3:
        raise ValueError(f'JCIC CSV 欄位數不足: {header!r}')

    metric_name = str(header[2]).strip().lstrip('\ufeff')
    if SOURCE_HEADER_KEY not in metric_name and '平均每戶持卡張數' not in metric_name:
        raise ValueError(f'JCIC CSV 指標欄位異常: {metric_name}')

    data: dict[str, dict[str, Any]] = {}
    for row in rows[1:]:
        if len(row) < 3:
            continue
        y_text = str(row[0]).strip()
        m_text = str(row[1]).strip()
        v_text = str(row[2]).strip()
        if not y_text or not m_text or not v_text:
            continue
        try:
            ad_y = int(y_text)
            mm = int(m_text)
            value = float(v_text)
        except Exception:
            continue
        if not 1 <= mm <= 12:
            continue
        roc_y = ad_y - 1911
        month = roc_month_text(roc_y, mm)
        data[month] = {
            'month': month,
            'roc_year': roc_y,
            'month_number': mm,
            'ad_yyyymm': ad_y * 100 + mm,
            'value': value,
        }

    if not data:
        raise ValueError('JCIC CSV 找不到可用資料')

    return {
        'source_header': metric_name,
        'latest_ad_yyyymm': max(item['ad_yyyymm'] for item in data.values()),
        'data': data,
    }


def pending_months_from_base(ws, index_map: dict[str, int], row_map: dict[tuple[int, str], int], series_data: dict[str, dict[str, Any]], base_month: str, backfill_months: int, overwrite: bool) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    value_col = index_map['avg_cards_per_person']
    for req in month_window_from_base(base_month, backfill_months):
        item = series_data.get(req['month'])
        if item is None:
            continue
        row_no = row_map.get((req['ad_yyyymm'], TARGET_ITEM))
        if row_no is None:
            out.append({
                'month': item['month'],
                'ad_yyyymm': item['ad_yyyymm'],
                'row': None,
                'reason': 'missing_target_row_within_base_window',
            })
            continue
        current = ws.cell(row_no, value_col).value
        if overwrite or is_blank(current):
            out.append({
                'month': item['month'],
                'ad_yyyymm': item['ad_yyyymm'],
                'row': row_no,
                'reason': 'overwrite' if overwrite else 'blank_value_within_base_window',
            })
    return out


def write_single_month(ws, index_map: dict[str, int], row_map: dict[tuple[int, str], int], item: dict[str, Any], overwrite: bool) -> dict[str, Any]:
    ad_yyyymm = int(item['ad_yyyymm'])
    row_no = require_target_row(row_map, ad_yyyymm)
    value_col = index_map['avg_cards_per_person']
    cell = ws.cell(row_no, value_col)
    before = cell.value
    if overwrite or is_blank(before):
        cell.value = item['value']
        if item['value'] is not None:
            apply_default_font(cell)
        action = 'written'
    else:
        action = 'skipped_existing'
    return {
        'month': item['month'],
        'ad_yyyymm': ad_yyyymm,
        'row': row_no,
        'before_value': before,
        'written_value': ws.cell(row_no, value_col).value,
        'action': action,
    }


def main() -> None:
    configure_utf8_stdio()
    args = parse_args()
    workbook = Path(args.workbook)
    if not workbook.exists():
        raise FileNotFoundError(f'找不到工作簿: {workbook}')
    if workbook.stat().st_size == 0:
        raise ValueError(f'工作簿檔案大小為 0 bytes，無法更新: {workbook}')

    resolved_source_url, source_resolution = resolve_source_url(args.entry_url, args.source_url, insecure=args.insecure)
    _, series = fetch_series(resolved_source_url, insecure=args.insecure)

    wb = load_workbook(workbook)
    if args.sheet not in wb.sheetnames:
        raise KeyError(f'找不到工作表: {args.sheet}; 現有工作表: {wb.sheetnames}')
    ws = wb[args.sheet]
    index_map = build_index_map(ws)
    row_map = build_month_item_rows(ws, index_map)

    output: dict[str, Any] = {
        'status': 'success',
        'workbook': str(workbook),
        'sheet': ws.title,
        'entry_url': args.entry_url,
        'resolved_source_url': resolved_source_url,
        'source_resolution': source_resolution,
        'ssl_fallback_used': SSL_FALLBACK_USED,
        'source_header': series['source_header'],
        'target_item': TARGET_ITEM,
        'target_label': args.label,
        'overwrite': args.overwrite,
    }

    normalized_base_month = roc_month_text(*parse_roc_month(args.base_month))
    pending = pending_months_from_base(
        ws,
        index_map,
        row_map,
        series['data'],
        normalized_base_month,
        args.backfill_months,
        overwrite=args.overwrite,
    )
    skipped_missing_block = [{**req, 'action': 'skip_missing_month_block'} for req in pending if req['row'] is None]
    updated = []
    for req in pending:
        if req['row'] is None:
            continue
        item = series['data'][req['month']]
        updated.append(write_single_month(ws, index_map, row_map, item, overwrite=args.overwrite))
    output.update({
        'mode': 'base_month_window_backfill_longform',
        'base_month': normalized_base_month,
        'backfill_months': args.backfill_months,
        'detected_pending_months': pending,
        'skipped_missing_block': skipped_missing_block,
        'updated_count': len(updated),
        'updated': updated,
        'latest_available_month': max(series['data'].values(), key=lambda x: x['ad_yyyymm'])['month'],
    })
    if skipped_missing_block:
        warnings = output.setdefault('warnings', [])
        for req in skipped_missing_block:
            warnings.append(f"YYYYMM={req['ad_yyyymm']} 無 11 列 block，未寫入，留待 block 建立後回補")

    if hasattr(wb, 'calculation'):
        try:
            wb.calculation.fullCalcOnLoad = True
            wb.calculation.forceFullCalc = True
        except Exception:
            pass

    wb.save(workbook)
    print(json.dumps(output, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        print(json.dumps({
            'status': 'error',
            'error_type': type(exc).__name__,
            'error': str(exc),
        }, ensure_ascii=False, indent=2), file=sys.stderr)
        raise
