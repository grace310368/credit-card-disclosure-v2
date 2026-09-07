from __future__ import annotations

"""Shared bank / block constants for credit card disclosure scripts.

工作簿 `歷史資料(年+月)` 的 block 版型：每個年月（或年度 `YYYY--`）固定 11 列，
依 `排序編號` 1–10 為十家銀行、11 為 `市場總計`；`Item` 欄為「NN 名稱」（如 `01 中信`）、
`Bank` 欄為簡稱。市場總計列同時承載 `平均每人持卡張數`、`市場總計` 與 TOP5/TOP10 公式。
"""

import re

BANK_ITEM_ALIASES = {
    "ctbc": ["中國信託商業銀行", "中國信託銀行", "中國信託", "中信"],
    "fubon": ["台北富邦銀行", "台北富邦商業銀行", "富邦銀行", "富邦"],
    "cathay": ["國泰世華銀行", "國泰世華商業銀行", "國泰世華", "國泰"],
    "esun": ["玉山銀行", "玉山商業銀行", "玉山"],
    "taishin": ["台新銀行", "台新國際商業銀行", "台新"],
    "dbs": ["星展銀行", "星展(台灣)商業銀行", "星展(台灣)", "星展"],
    "ubot": ["聯邦銀行", "聯邦商業銀行", "聯邦"],
    "sinopac": ["永豐銀行", "永豐商業銀行", "永豐"],
    "firstbank": ["第一銀行", "第一商業銀行", "一銀", "第一"],
    "feib": ["遠東商銀", "遠東國際商業銀行", "遠銀", "遠東"],
}

# 銀行局原始檔的銀行名稱比工作簿 Item 更保守，避免用過短別名誤判。
BANK_BUREAU_ALIASES = {
    "ctbc": ["中國信託商業銀行", "中國信託銀行", "中國信託"],
    "fubon": ["台北富邦商業銀行", "台北富邦銀行", "富邦銀行"],
    "cathay": ["國泰世華商業銀行", "國泰世華銀行", "國泰世華"],
    "esun": ["玉山商業銀行", "玉山銀行", "玉山"],
    "taishin": ["台新國際商業銀行", "台新銀行", "台新"],
    "dbs": ["星展(台灣)商業銀行", "星展銀行", "星展(台灣)", "DBS"],
    "ubot": ["聯邦商業銀行", "聯邦銀行", "聯邦"],
    "sinopac": ["永豐商業銀行", "永豐銀行", "永豐"],
    "firstbank": ["第一商業銀行", "第一銀行", "一銀"],
    "feib": ["遠東國際商業銀行", "遠東商銀", "遠銀"],
}

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

MARKET_TOTAL_ITEM = "市場總計"
SPECIAL_ITEMS = [MARKET_TOTAL_ITEM]
# 舊版工作簿曾用「市場總計(銀行局)」與「銀行局」兩列承載市場總計與 TOP 公式，一律視為市場總計列。
SPECIAL_ITEM_ALIASES = {MARKET_TOTAL_ITEM: ["市場總計", "市場總計(銀行局)", "銀行局"]}

BLOCK_ITEMS = [BANK_NAMES[key] for key in BANK_ORDER] + SPECIAL_ITEMS
BLOCK_SIZE = len(BLOCK_ITEMS)
ITEM_RANKS = {item: rank for rank, item in enumerate(BLOCK_ITEMS, start=1)}

_ITEM_PREFIX_RE = re.compile(r"^\d{1,2}\s*")


def strip_item_prefix(text: str) -> str:
    """去掉 Item 欄的排序前綴：'01 中信' -> '中信'。"""
    return _ITEM_PREFIX_RE.sub("", str(text or "").strip()).strip()


def item_label(item: str) -> str:
    """Item 欄的標準寫法：'中信' -> '01 中信'、'市場總計' -> '11 市場總計'。"""
    rank = ITEM_RANKS.get(item)
    return f"{rank:02d} {item}" if rank else item


def year_block_key(ad_year: int) -> str:
    """年度 block 的 YYYYMM 欄寫法：2025 -> '2025--'。"""
    return f"{int(ad_year)}--"
