# creditcardinfo 維運說明

抓取台灣 10 家銀行信用卡官方揭露資料，並維護工作簿 `銀行局信用卡公開資料.xlsx`。
沙箱化執行環境（input_files ≤ 10 檔、timeout ≤ 900 秒）的精確配方見 `SKILL.md` 情境 A–D；本文說明架構、分工與本機用法。

---

## 支援銀行

| bank_key | 銀行名稱 |
|---|---|
| `ctbc` | 中國信託商業銀行 |
| `fubon` | 台北富邦銀行 |
| `cathay` | 國泰世華銀行 |
| `esun` | 玉山銀行 |
| `taishin` | 台新銀行 |
| `dbs` | 星展銀行 |
| `ubot` | 聯邦銀行 |
| `sinopac` | 永豐銀行 |
| `firstbank` | 第一銀行 |
| `feib` | 遠東商銀 |

---

## 工作簿結構

- 預設工作簿：`銀行局信用卡公開資料*.xlsx`（檔名帶資料區間與日期）；主要工作表：`歷史資料(年+月)`，整張表是 Excel Table `creditcard`
- 前 7 欄：`YYYYMM`、`年度`、`月份`、`季度`、`排序編號`、`Bank`、`Item`；其中年度/月份/季度是結構化公式（由 `YYYYMM` 推得）
- 14 個指標欄：流通卡數、有效卡數、當月發卡數、當月停卡數、循環信用餘額、未到期分期付款餘額、當月簽帳金額、
  當月預借現金金額、逾期三個月以上比率、逾期六個月以上比率、備抵呆帳提足率、當月轉銷呆帳金額、
  當年度累計轉銷呆帳金額、當月轉帳卡簽帳金額
- 衍生欄：`平均每人持卡張數`、`市場總計`、TOP 5 / TOP 10 各種占比欄
- 每個年月（或年度）固定 **11 列一組**：`排序編號` 1–10 為十家銀行、11 為 `市場總計`；`Item` 寫「NN 名稱」（`01 中信`、`11 市場總計`），`Bank` 為簡稱
- `市場總計` 列承載官方市場總計 13 欄、`平均每人持卡張數`，以及 `市場總計` 與 TOP5/TOP10 占比公式（SUMIFS，`[#This Row]` 結構化參照）
- 定位邏輯：先用 `YYYYMM` 找 block，再用 `Item`（去前綴）找列；新增 block 時複製上一個 block 的樣式與公式，並延伸 Table 範圍
- 年度 block 的 `YYYYMM` 欄寫 `YYYY--`（如 `2025--`）
- block 依新增順序排在表尾，不保證依年月排序；公式以 `YYYYMM` 比對，順序不影響結果

---

## 腳本分工

### 資料抓取層（不寫 Excel）

- `Script/<bank>_creditcard_disclosure.py` × 10：從官網（中信為本機上傳檔）抓資料，輸出標準 metrics JSON。
  金額欄位以來源單位（仟元）回傳，並在 `metric_units` 宣告；不依賴共用模組，各自內建所需函式。
- `Script/run_all_banks.py`：循序執行各銀行腳本、正規化並彙整為 summary JSON。
  - 卡數轉數字；`*_thousand` 金額轉 `*_million`（保留原始仟元欄）；百分比轉小數比率；補出扁平欄位供回寫層讀取
  - `--collect-only`：從既有 JSON（彙整 payload 或單一銀行 stdout）組 summary，不重抓；後面的檔案覆蓋同銀行舊結果，已正規化的紀錄不會重複轉換
  - `--preflight`：檢查 openpyxl/pypdf、各銀行腳本、中信來源檔
  - `--allow-insecure`：轉傳給 firstbank 的第三段 TLS fallback

### Excel 回寫層

- `Script/update_credit_card_workbook.py`：主控。寫入 10 家銀行月 block → 委派銀行局 backfill → 市場總計 → JCIC →
  該年 1–12 月齊全時自動整理年度 block → 重新載入工作簿讀回驗證（輸出 `verification` 段）。
  另有 `--repair-partial-blocks`（刪除表尾孤兒列並回報 `deleted`；表中的只列 `needs_manual_review`）與 `--annual-only`。
  一般月更新前會預掃不完整月 block 並 fail-fast。
- `Script/run_market_total.py`：金管會 ZIP → 更新 `市場總計` 列的 13 個指標欄（`市場總計` 欄本身是公式，不動）；支援 lookback 回補與本機 ZIP。
- `Script/jcic_avg_cards_update.py`：JCIC CSV → 更新 `市場總計` 列的 `平均每人持卡張數`；以 `--base-month` 往前回補空白。
- `Script/run_bank_bureau_bank_backfill.py`：金管會 ZIP → 回補 10 家銀行舊月份空白（13 欄：卡數 4＋金額 6＋比率 3），
  並補齊 `市場總計` 列的 `市場總計`/TOP5/TOP10 公式（從鄰近月份原樣複製）。金額千元→百萬；逾期比率 /100 存小數並套 `0.00%`；13 欄全為必要欄位，缺任一欄該月報錯。
  回補下限 202601；預設只補空白（`--overwrite` 才覆蓋既有值）。該月 ZIP 未發布（金管會回 404 或導回首頁）時記 `skip_missing_banks`，不建空 block。

### 共用模組（只在回寫層 import）

- `bank_aliases.py`：銀行別名、`BANK_ORDER`、`BLOCK_ITEMS`/`ITEM_RANKS`（11 列）、Item 標籤與年度鍵工具
- `percent_utils.py`：百分比正規化（`force_percent_input`）與 `0.00%` 格式
- `workbook_block_helpers.py`：11 列 block 的尋找 / 建立 / 驗證、樣式與公式複製、Table 範圍延伸、Item 標籤正規化

**月 block 只能由 update / backfill 建立**；market 與 jcic 對沒有 block 的月份一律跳過並輸出 `skip_missing_month_block` warning。

---

## 本機用法

### 抓各銀行資料

```bash
python creditcardinfo/Script/run_all_banks.py --banks all --month 115年04月 --json-output out/summary_11504.json
```

總時間 ≈ 銀行數 × `--timeout`。沙箱檔案數不夠帶 `--banks all` 時分組跑再合併：

```bash
python creditcardinfo/Script/run_all_banks.py --collect-only out/g1.json out/g2.json --json-output out/summary_11504.json
```

### 回寫工作簿

```bash
python creditcardinfo/Script/update_credit_card_workbook.py \
  --summary out/summary_11504.json \
  --month 115年04月 \
  --workbook 銀行局信用卡公開資料.xlsx
```

不帶 `--month` 時取 summary 各行 `data_month` 最新月為 target。判讀：`verification.status == "ok"`、`mismatches` 為空；
`warnings` 中「落後月份銀行已跳過」屬正常，留給 backfill。

### 補舊月份

```bash
python creditcardinfo/Script/run_bank_bureau_bank_backfill.py --workbook 銀行局信用卡公開資料.xlsx --month 115年01月
```

---

## summary JSON 欄位

每家銀行結果保留巢狀 `metrics` / `metrics_units` / `metrics_source_units` / `source` / `errors` / `steps` / `_runner` 供追溯，
並補出以下扁平欄位供回寫層讀取：

`circulating_cards`、`valid_cards`、`new_cards_this_month`、`cancelled_cards_this_month`、
`signed_amount_million`、`revolving_balance_million`、`installment_balance_not_yet_due_million`、`cash_advance_amount_million`、
`overdue_3m_ratio_percent`、`overdue_6m_ratio_percent`、`allowance_coverage_ratio_percent`、
`charge_off_amount_this_month_million`、`charge_off_amount_ytd_million`、`debit_card_signed_amount_million`、
`status`、`data_month`、`base_date`、`notes`、`failure_reason`、`source_type`、`source_url`、`error_count`
（各 `*_million` 皆另保留對應的 `*_thousand` 原始值）

---

## 百分比單位規則

- JSON 與工作簿內部一律存**小數比率**（5% → `0.05`），Excel 靠 `0.00%` 格式顯示
- 逾期三/六個月比率的來源（官網、銀行局檔）是百分比顯示數字且恆小於 1（`0.12` 代表 0.12%），
  解析時用 `force_percent_input=True` 一律 /100，不可用 abs>1 啟發式
- 儲存格本身是 % 格式時 raw 已是小數，直接保留；中信腳本在解析階段就轉小數並在 `metric_units` 宣告 `decimal_ratio`，
  `run_all_banks` 看到 `decimal_ratio` 不再除，整條管線只除一次
- 備抵呆帳提足率（值恆 >1，如 `429.67` → `4.2967`）沿用 abs>1 才 /100 的啟發式
- 先前被寫成 100 倍的既有儲存格不會被預設回補覆蓋，需要時對受影響月份跑 `--overwrite`

---

## 中信（CTBC）來源檔

`中信信用卡<yyyymm>.xlsx` 的 14 個項目順序固定，解析腳本依固定順序抓取：

1. 流通卡數
2. 有效卡數
3. 當月發卡數
4. 當月停卡數
5. 循環信用餘額
6. 當月簽帳金額
7. 當月預借現金金額
8. 逾期三個月以上帳款占應收帳款餘額(含催收款)之比率
9. 逾期六個月以上帳款占應收帳款餘額(含催收款)之比率
10. 備抵呆帳提足率
11. 當月轉銷呆帳金額
12. 當年度累計轉銷呆帳金額
13. 未到期分期付款餘額
14. 當月轉帳卡簽帳金額

---

## 改文件或程式時要避免的說法

- 預設工作簿是 `信用卡資料彙總.xlsx`，或主要工作表是 `信用卡 / 原始資料(月) / 原始資料(年)`（舊版三表模型已淘汰）
- 月份在欄、指標在列
- 每月 13 列、另有 `銀行局` 與 `財團法人金融聯合徵信中心` 兩列（v4 起為 11 列，TOP 公式與 JCIC 值都在 `市場總計` 列）
- market / jcic 腳本可以替不存在的月份自建列（一律跳過；月 block 只由 update / backfill 建立）
- 沙箱中可以一次帶 `--banks all` 的完整檔案組合（實需 12+ 檔，改用分組＋`--collect-only`）
