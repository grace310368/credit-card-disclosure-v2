from __future__ import annotations

"""Shared bank alias constants for credit card disclosure scripts."""

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
MARKET_TOTAL_ITEM = "市場總計(銀行局)"
BANK_BUREAU_ITEM = "銀行局"
JCIC_ITEM = "財團法人金融聯合徵信中心"
SPECIAL_ITEMS = [MARKET_TOTAL_ITEM, BANK_BUREAU_ITEM, JCIC_ITEM]
ITEM_RANKS = {BANK_NAMES[key]: rank for rank, key in enumerate(BANK_ORDER, start=1)}
ITEM_RANKS.update({MARKET_TOTAL_ITEM: 11, BANK_BUREAU_ITEM: 12, JCIC_ITEM: 13})
