# creditcardinfo 維運說明

本資料夾用來抓取台灣多家銀行信用卡官方揭露資料，並維護新版工作簿：`銀行局信用卡公開資料.xlsx`。

> 這份 README 以 **新版報表結構** 為準，整理目前實際分工與維運方式：
>
> - 哪些部分是資料抓取層
> - 哪些部分是 Excel 回寫層
> - 哪些文件段落是舊版對照／過時說法提醒
>
> 若程式碼或舊資料仍保留舊版三張 sheet 痕跡，請把它視為 **相容舊資料的歷史包袱**，不要再把舊結構當成新規格。

---

## 先看這 8 件事

1. 預設工作簿是：`銀行局信用卡公開資料.xlsx`
2. 主要工作表名稱是：`歷史資料(年+月)`
3. 新版報表是 **單一工作表模型**，不是舊版三張 sheet
4. 每個年月或年度資料區塊是 **固定 13 列一組**
5. 10 家銀行抓取腳本本體目前大致 **不用因新報表格式而修改**
6. `run_all_banks.py` 仍是**資料抓取 / JSON 彙整層**，不直接寫 Excel
7. 目前 Excel 回寫層已分工為：
   - `update_credit_card_workbook.py`：寫月 block、委派下列三支、年度整理、`--repair-partial-blocks` 修復、結尾讀回驗證（輸出 `verification` 段）
   - `run_market_total.py`：更新 `市場總計(銀行局)` 與 `銀行局.市場總計`
   - `jcic_avg_cards_update.py`：更新 `財團法人金融聯合徵信中心 / 平均每人持卡張數`
   - `run_bank_bureau_bank_backfill.py`：回補 10 家銀行舊月份空白，並修復 `銀行局` Top5/Top10 公式
8. **月 block（13 列）只能由 update / backfill 建立**：market 與 jcic 對「沒有 block 的月份」
   一律跳過並輸出 `skip_missing_month_block` warning，不會再自建孤兒列
9. 新版目前已支援：
   - 月資料回寫（含結尾讀回驗證）
   - 市場總計回寫
   - JCIC 回寫
   - 舊月份 backfill
   - 年資料自動整理
   - `--collect-only`（從既有 JSON 組 summary，不重抓）
   - `--preflight`（無網路環境檢查）
   - `--repair-partial-blocks`（刪除表尾孤兒列並回報被移除的值）

---

## 目錄重點

### 主要腳本

- `Script/update_credit_card_workbook.py`：工作簿更新主控；寫入月 block、委派三支子腳本、年度整理、
  `--repair-partial-blocks` 修復模式、預掃 fail-fast（發現不完整月 block 先擋下）、結尾讀回驗證
- `Script/run_all_banks.py`：批次執行各銀行腳本，輸出 JSON 彙整結果；支援 `--collect-only`（免重抓合併）、
  `--preflight`（環境檢查）、`--allow-insecure`（轉傳給 firstbank）
- `Script/run_market_total.py`：更新 `市場總計(銀行局)` 列，並同步更新 `銀行局` 列的 `市場總計`；月 block 不存在時跳過
- `Script/jcic_avg_cards_update.py`：更新 `財團法人金融聯合徵信中心` 列的 `平均每人持卡張數`；月 block 不存在時跳過
- `Script/run_bank_bureau_bank_backfill.py`：用銀行局逐家銀行資料回補舊月份 10 家銀行缺漏，並修復 `銀行局` Top5/Top10 公式

### 主要檔案

- `銀行局信用卡公開資料.xlsx`
- `SKILL.md`
- `README.md`

---

## 支援銀行

固定 10 家：

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

## 新版工作簿結構

### 工作簿與工作表

- 預設工作簿：`銀行局信用卡公開資料.xlsx`
- 主要工作表：`歷史資料(年+月)`

### 欄位結構

新版報表共有約 33 欄，前幾欄包含：

- `YYYYMM`
- `年度`
- `月份`
- `Rank`
- `Item`

其餘資料欄與衍生欄包含：

- 流通卡數
- 有效卡數
- 當月發卡數
- 當月停卡數
- 循環信用餘額
- 未到期分期付款餘額
- 當月簽帳金額
- 當月預借現金金額
- 逾期三個月以上比率
- 逾期六個月以上比率
- 備抵呆帳提足率
- 當月轉銷呆帳金額
- 當年度累計轉銷呆帳金額
- 當月轉帳卡簽帳金額
- 平均每人持卡張數
- 市場總計
- TOP 5 / TOP 10 各種占比欄

### Block 結構

每個年月或年度 total 都是 **固定 13 列一組**：

1. 10 家銀行
2. `市場總計(銀行局)`
3. `銀行局`
4. `財團法人金融聯合徵信中心`

### 這代表什麼

新版邏輯應改成：

- 先用 `YYYYMM` 定位區塊
- 再用 `Item` 定位該 block 內的列
- 針對對應欄位回寫資料、公式或補值

**不要再用舊版「月份在欄、指標在列」的設計思維。**

---

## 哪些部分不用改

### 1. 10 家銀行抓取腳本本體

以下腳本目前看起來 **不用因為新報表格式而修改**：

- `creditcardinfo/Script/ctbc_creditcard_disclosure.py`
- `creditcardinfo/Script/fubon_creditcard_disclosure.py`
- `creditcardinfo/Script/cathay_creditcard_disclosure.py`
- `creditcardinfo/Script/esun_creditcard_disclosure.py`
- `creditcardinfo/Script/taishin_creditcard_disclosure.py`
- `creditcardinfo/Script/dbs_creditcard_disclosure.py`
- `creditcardinfo/Script/ubot_creditcard_disclosure.py`
- `creditcardinfo/Script/sinopac_creditcard_disclosure.py`
- `creditcardinfo/Script/firstbank_creditcard_disclosure.py`
- `creditcardinfo/Script/feib_creditcard_disclosure.py`

原因：

- 這些腳本的責任是從銀行官網抓資料
- 再轉成標準 metrics JSON
- 它們不直接依賴舊 workbook 的 sheet 結構

### 2. `run_all_banks.py` 核心抓取邏輯

- `creditcardinfo/Script/run_all_banks.py`

這支目前看起來 **核心抓取流程大致不用改**，但文件敘述要更新。

原因：

- 它不直接寫 Excel
- 只負責彙整各銀行 JSON
- 已能輸出新版工作簿可用的多數扁平欄位

目前已確認可扁平化的欄位包含：

- `circulating_cards`
- `valid_cards`
- `new_cards_this_month`
- `cancelled_cards_this_month`
- `signed_amount_million`
- `revolving_balance_million`
- `installment_balance_not_yet_due_million`
- `cash_advance_amount_million`
- `overdue_3m_ratio_percent`
- `overdue_6m_ratio_percent`
- `allowance_coverage_ratio_percent`
- `charge_off_amount_this_month_million`
- `charge_off_amount_ytd_million`
- `debit_card_signed_amount_million`

結論：

- 新報表需要的資料來源層大致已夠
- 主要問題在 **怎麼回寫新版 Excel**

---

## 目前腳本分工

### 1. `update_credit_card_workbook.py`

- `creditcardinfo/Script/update_credit_card_workbook.py`

目前職責：

- 以 `歷史資料(年+月)` 為主要工作表
- 以 `YYYYMM` / `Item` / 13-row block 為定位基礎
- 將 `run_all_banks.py` 的 summary JSON 寫入月 block
- 可委派市場總計、JCIC、銀行局 backfill 腳本
- 當某年度 1–12 月資料完整時，自動整理成年資料
- 亦支援 `--annual-only` 單獨執行年度整理

### 2. `run_market_total.py`

- `creditcardinfo/Script/run_market_total.py`

目前職責：

- 依 `YYYYMM` 找到對應月份 block
- 更新 `市場總計(銀行局)` 那一列
- 更新 `銀行局` 那一列的 `市場總計` 欄位
- 支援 lookback / backfill / workbook fallback

### 3. `jcic_avg_cards_update.py`

- `creditcardinfo/Script/jcic_avg_cards_update.py`

目前職責：

- 以 `歷史資料(年+月)` 為目標工作表
- 依年月 block 定位 `財團法人金融聯合徵信中心` 列
- 回寫 `平均每人持卡張數`
- 支援 base month window backfill

### 4. `run_bank_bureau_bank_backfill.py`

- `creditcardinfo/Script/run_bank_bureau_bank_backfill.py`

目前職責：

- 用銀行局 ZIP / XLSX 補 10 家銀行舊月份空白（v3 起共 **13 欄**：
  流通卡數、有效卡數、當月發卡數、當月停卡數、當月簽帳金額、循環信用餘額、
  未到期分期付款餘額、當月預借現金金額、逾期三/六個月比率、備抵呆帳提足率、
  當月轉銷呆帳金額、當年度累計轉銷呆帳金額）
- 金額欄由千元轉百萬元；逾期比率一律 /100 存小數並套 `0.00%` 格式；
  13 個資料欄全為必要欄位：來源缺任一欄時該月整月報錯（fail-fast）
- 回補下限固定為 202601（2026年01月），自動模式與 `--month` 都不處理更早月份
- 直接回補 `歷史資料(年+月)` 中指定月份 block
- 修復 `銀行局` 列的 Top5 / Top10 公式
- 相依模組：`bank_aliases.py`、`percent_utils.py`、`workbook_block_helpers.py`
- 不再處理 `原始資料(月)` / `信用卡` / 年度同步
- 不再負責 `市場總計(銀行局)` 或 `銀行局.市場總計` 更新

---

## README / SKILL.md 已同步更新的重點

目前文件已統一反映以下新版現況：

- 預設工作簿為 `銀行局信用卡公開資料.xlsx`
- 主要工作表為 `歷史資料(年+月)`
- 不再把舊版三張 sheet 當成預設模型
- 明確標示四支主要回寫腳本的目前分工
- 說明新版月 block、backfill、市場總計、JCIC、年度整理的責任分界

---

## 實際流程

> 沙箱化執行環境（input_files ≤ 10 檔、timeout ≤ 900 秒）的精確配方見 `SKILL.md` 情境 A–D；
> 以下為在一般環境（本機終端）的用法。

### 流程 A：抓各銀行資料

```bash
python creditcardinfo/Script/run_all_banks.py --banks all --month 115年04月 --json-output out/summary_11504.json
```

用途：

- 跑 10 家銀行子腳本（循序；總時間 ≈ 銀行數 × `--timeout`）
- 輸出 JSON 到 stdout；`--json-output` 另存 JSON 檔
- 沙箱環境檔案數不夠帶 `--banks all` 時，分兩組跑再用 `--collect-only` 合併：

```bash
python creditcardinfo/Script/run_all_banks.py --collect-only out/g1.json out/g2.json --json-output out/summary_11504.json
```

- `--collect-only` 接受彙整 payload 或單一銀行 stdout JSON；後面的檔案覆蓋前面同銀行的結果，
  已正規化的紀錄不會被重複轉換

### 流程 B：回寫新版工作簿

目前新版主流程已可使用：

```bash
python creditcardinfo/Script/update_credit_card_workbook.py \
  --summary out/banks_11504.json \
  --month 115年04月 \
  --workbook 銀行局信用卡公開資料.xlsx
```

用途：

- 依 `YYYYMM` 找到指定月份 block（發現不完整月 block 會先 fail-fast，指引跑 `--repair-partial-blocks`）
- 更新 10 家銀行對應列
- 依固定順序委派：銀行局逐家銀行 backfill → 市場總計 → JCIC 補值
- 維持 `銀行局` 與 `財團法人金融聯合徵信中心` 列的正確性
- 若該年度 1–12 月資料完整，會自動整理成年資料
- 結尾重新載入工作簿做讀回驗證，輸出 JSON 的 `verification` 段（`status`/`mismatches`/`special_rows`）

---

## `run_all_banks.py` 與新版工作簿的接軌方式

目前 `run_all_banks.py` 已調整為：

- **保留** `results[].metrics` 巢狀欄位，供除錯與追溯
- **同時補出扁平欄位**，可作為新版工作簿回寫輸入來源

### 每家銀行結果常見欄位

- `bank`
- `bank_key`
- `bank_name`
- `status`
- `data_month`
- `base_date`
- `metrics`
- `metrics_units`
- `metrics_source_units`
- `notes`
- `errors`
- `source`
- `captured_at`
- `steps`
- `_runner`

### 已知重要扁平欄位

- `circulating_cards`
- `valid_cards`
- `new_cards_this_month`
- `cancelled_cards_this_month`
- `signed_amount_thousand`
- `signed_amount_million`
- `revolving_balance_thousand`
- `revolving_balance_million`
- `installment_balance_not_yet_due_million`
- `cash_advance_amount_million`
- `overdue_3m_ratio_percent`
- `overdue_6m_ratio_percent`
- `allowance_coverage_ratio_percent`
- `charge_off_amount_this_month_million`
- `charge_off_amount_ytd_million`
- `debit_card_signed_amount_thousand`
- `debit_card_signed_amount_million`
- `status`
- `data_month`
- `base_date`
- `notes`
- `failure_reason`
- `source_type`
- `source_url`
- `error_count`

### 金額標準化

`run_all_banks.py` 會把：

- `signed_amount_thousand` → `signed_amount_million`
- `revolving_balance_thousand` → `revolving_balance_million`

也會保留原始仟元欄位，供追溯與除錯。

---

## 建議維運流程

### 日常月更新

1. 先跑 `run_all_banks.py`（沙箱環境分兩組＋`--collect-only`；見 SKILL.md 情境 A）
2. 檢查 JSON 結果與失敗銀行（`failed_count`、`missing_banks`）
3. 確認 summary JSON 內容完整
4. 再執行 `update_credit_card_workbook.py`（一條龍，不加 `--skip-*`）
5. 看輸出 JSON 的 `verification` 段：`status == "ok"`、`mismatches` 為空
6. `warnings` 中「落後月份銀行已跳過」屬正常，留給 backfill

### 補舊月份 / 回補缺漏

適合情境：

- 銀行官網已查不到舊資料
- 舊月份有空白
- 需要以銀行局揭露資料補足歷史區塊

目前做法：

- 使用 `run_bank_bureau_bank_backfill.py` 直接回補 `歷史資料(年+月)` 的指定月份 block
- 補 10 家銀行資料空白（13 欄；舊月份只缺新欄位時，自動模式也會掃出來補齊）
- 預設只補空白、不覆蓋既有值（既有錯值需 `--overwrite` 才會被蓋掉）
- 修復 `銀行局` 列的 Top5 / Top10 公式
- 不再沿用舊版 `原始資料(月)` / `信用卡` / 年度同步 模型

---

## 修改文件或程式時要避免的過時說法

以下說法在新版工作簿情境下都不準，請避免：

- 預設工作簿是 `信用卡資料彙總.xlsx`
- 主要工作表是 `信用卡 / 原始資料(月) / 原始資料(年)`
- 月份一定在欄、指標一定在列
- 市場總計腳本只要找月份欄與固定 row 即可
- JCIC 腳本只要找 `平均每人持卡張數` 那一列即可
- 回補腳本仍以 `原始資料(月)` 為主
- 年度同步仍沿用舊版三表模型
- 主控腳本會自動執行 `run_all_banks.py`
- `--summary` 支援 `summary.xlsx`
- `--target-row` / `--jcic-target-row` 只保留作為相容舊參數；新版 block 定位不依賴它們
- market / jcic 腳本可以替不存在的月份自建列（v2 起一律跳過；月 block 只由 update / backfill 建立）
- 沙箱中可以一次帶 `--banks all` 的完整檔案組合（實需 12+ 檔，超過 input_files 上限；改用分組＋`--collect-only`）

---

## v2 強韌化摘要（2026-07-08）

針對「弱模型＋沙箱化 python 工具（input_files ≤ 10 檔、timeout ≤ 900 秒、拒絕覆寫未列檔案）」
執行環境所做的修改：

1. **孤兒列防護**：`jcic_avg_cards_update.py` 的 `ensure_target_row` 與 `run_market_total.py` 的
   `ensure_special_item_row` 不再替沒有 13 列 block 的月份建列（曾在生產工作簿留下單列 202604 block，
   會讓後續 backfill / 年度同步 crash）
2. **修復模式**：`update_credit_card_workbook.py --repair-partial-blocks` 刪除表尾 partial block
   （表中的只回報 `needs_manual_review`，因 openpyxl `delete_rows` 不平移下方公式參照）；
   預設月更新路徑加入預掃 fail-fast
3. **讀回驗證**：update 結尾重載工作簿逐欄比對，輸出 `verification` 段，取代人工 openpyxl 驗證
4. **bytecode 抑制**：會 import 共用模組的 6 支 entrypoint（run_all / ctbc / update / market / jcic / backfill）
   設 `sys.dont_write_bytecode`、子行程一律帶 `-B`，不再產生 `__pycache__/*.pyc`
   （沙箱「Refusing to overwrite」的根源；其餘 9 支 scraper 無 sibling import，直接執行本就不會寫 .pyc）
5. **UTF-8 stdout**：四支 writer 腳本補上 `configure_utf8_stdio()`（先前輸出 cp950 亂碼）
6. **collect-only / preflight / --allow-insecure**：`run_all_banks.py` 支援免重抓合併、
   無網路環境檢查、firstbank 第三段 TLS 轉傳
7. **scraper 自帶 percent 函式**：`run_all_banks.py` 與 `ctbc_...py` 內嵌
   `to_number`/`normalize_percent_value`（與 `percent_utils.py` 同步維護），抓取層不再依賴共用模組
8. **SSL fallback 補齊**：esun / dbs / ubot 加上 CERT_NONE 自動 fallback；jcic 比照 market 加自動
   fallback 與 stderr JSON 錯誤輸出

## v3 修正摘要（2026-07-09）

1. **逾期比率單位修正（全管線）**：逾期三/六個月比率的來源值是百分比顯示數字且恆小於 1
   （如 `0.12` 代表 0.12%），舊的「abs > 1 才 /100」啟發式會誤把 `0.12` 原樣寫入，
   Excel `0.00%` 格式顯示成 12%（放大 100 倍）。`percent_utils.normalize_percent_value`
   新增 `force_percent_input`：這兩欄裸數字**一律 /100**；文字帶 `%`（/100）與
   儲存格本身是 % 格式（raw 已是小數，保留）兩條路徑不變。
   套用點：`run_all_banks.normalize_metrics`、`update_credit_card_workbook.load_summary_results`、
   `run_market_total`（市場總計列）、`run_bank_bureau_bank_backfill`（銀行局回補）。
   備抵呆帳提足率（值恆 >1，如 `429.67`→`4.2967`）沿用既有啟發式，不受影響。
2. **中信解析階段統一轉小數**：`ctbc_...py` 的 `normalize_metric_value` 三條路徑
   （% 文字、% 格式儲存格、裸數字）現在都輸出小數比率，固定版型解析改收儲存格物件
   以取得 `number_format`；這兩欄 `metric_units` 宣告 `decimal_ratio`，
   `run_all_banks` 看到 `decimal_ratio` 直接採用、不再除，保證整條管線只除一次。
3. **銀行局 backfill 擴為 13 欄**：新增未到期分期付款餘額、當月預借現金金額、
   逾期三/六個月比率、備抵呆帳提足率、當月轉銷呆帳金額、當年度累計轉銷呆帳金額。
   金額欄千元→百萬；比率欄寫入時套 `0.00%` 格式；來源表頭別名涵蓋 占/佔、全形括號、
   有無 `(%)` 後綴與「當年度轉銷呆帳金額累計至資料月份」（115年04月實際表頭）。
   新 7 欄原屬選配，現已升級為**必要**：與原 6 欄相同，來源缺任一欄即整月 fail-fast。
   另訂回補下限 `BACKFILL_START_YYYYMM = 202601`，2026年01月以前的月份一律不處理。
   空白偵測涵蓋 13 欄，舊月份只缺新欄位也會被自動模式列入回補（預設仍只補空白）。
   backfill 新增相依 `percent_utils.py`（沙箱 input_files 4→5 個）。
4. **歷史錯值不自動修正**：修正只影響之後寫入的資料；先前被寫成 100 倍的既有儲存格
   非空白、預設回補不會覆蓋，需要時對受影響月份跑 `--overwrite` 用銀行局資料整批蓋回。

## 維護時建議一起檢查的檔案

若要調整這套流程，建議一起檢查：

- `README.md`
- `SKILL.md`
- `Script/update_credit_card_workbook.py`
- `Script/run_all_banks.py`
- `Script/run_market_total.py`
- `Script/jcic_avg_cards_update.py`
- `Script/run_bank_bureau_bank_backfill.py`

---

## 結論

目前這個專案的關鍵不是「重新抓資料」，而是：

- 保留既有 10 家銀行抓取能力
- 延續 `run_all_banks.py` 的 JSON / 扁平欄位輸出
- 把 Excel 回寫層全面改成 **新版單一工作表 `歷史資料(年+月)` 的 block 模型**

簡單說：

- **抓取層大致夠用**
- **回寫層一定要重構**
- **README / SKILL.md 必須同步改成新版說法**


## 中信（CTBC）特殊來源說明

`中信信用卡資料.xlsx` 的 14 個項目目前順序穩定，CTBC 解析腳本應優先依固定順序抓取：

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

如果原始檔版型未變，這段不應再靠列順序猜測，也不應用最後一個數字碰運氣。
