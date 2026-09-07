---
name: credit-card-disclosure
description: 當使用者要抓取、整理、回補、同步或維護台灣銀行信用卡官方揭露資料時使用。高優先觸發詞包含：銀行局信用卡公開資料.xlsx、信用卡資料表20260706.xlsx、信用卡資料更新、信用卡追蹤表、run_all_banks.py、update_credit_card_workbook.py、銀行局信用卡重要資訊揭露、JCIC 平均每人持卡張數、歷史資料(年+月)、市場總計、流通卡數、有效卡數、當月發卡數、當月停卡數、當月簽帳金額、循環信用餘額、未到期分期付款餘額、當月預借現金金額、轉帳卡簽帳金額、補舊月份、補空白、改信用卡腳本、改 README、改 SKILL.md，以及中信/富邦/國泰/玉山/台新/星展/聯邦/永豐/一銀/遠銀等銀行信用卡揭露需求。若主題是信用卡推薦、權益比較、申辦資格或行銷文案，通常不要使用本 skill。
---

# Credit Card Disclosure Skill

> 版本 v3.1（2026-09-07）。部署同步檢查：`run_all_banks.py --help` 應有 `--collect-only`，
> `run_bank_bureau_bank_backfill.py --help` 描述應提到「13 欄」；缺任一項代表部署過期，先提醒使用者。

處理 **台灣銀行信用卡官方揭露資料** 的抓取、彙整、回寫 Excel、舊月份回補與腳本維護。

- 預設工作簿：`銀行局信用卡公開資料.xlsx`（工作區根目錄）；主要工作表：`歷史資料(年+月)`（單一工作表，約 33 欄）
- 每個年月（或年度 Total）固定 **13 列一組**：10 家銀行 → `市場總計(銀行局)` → `銀行局` → `財團法人金融聯合徵信中心`
- `Rank`：銀行列 1–10、`市場總計(銀行局)`=11、`銀行局`=12、`財團法人金融聯合徵信中心`=13
- 定位方式：先找 `YYYYMM` block，再用 `Item` 找列

---

## 情境 A：標準月更新（照抄，共 5 次執行）

> 執行工具限制：input_files ≤ 10 個檔案（只能是檔案）、timeout ≤ 900 秒。以下配方已按限制設計好，**不要自行合併或重新設計**。
> `<yyyymm>` = 西元年月（如 202605）、`<月>` = 民國格式（如 `115年05月`）。
> 若使用者沒指定月份，先讀工作簿最新 `YYYYMM`，目標月通常是它的下一個月（官方揭露約落後 2 個月）。

**步驟 0（僅首次或環境有疑慮時）**：確認中信來源檔 `creditcardinfo/input/中信信用卡<yyyymm>.xlsx` 已存在，
**缺這個檔 ctbc 一定 failed，先請使用者上傳**。可用 preflight 檢查：
`python creditcardinfo/Script/run_all_banks.py --preflight --month <月>`
（input_files 同步驟 1；只看 `import:openpyxl`、`import:pypdf`、`ctbc_source_file` 三項，未 staged 的腳本檢查失敗屬預期）

**步驟 1：抓取 A 組（6 家）**
```
python creditcardinfo/Script/run_all_banks.py --banks ctbc,fubon,cathay,esun,taishin,feib \
  --month <月> --timeout 95 --json-output out/banks_groupA_<yyyymm>.json
```
input_files（8 個）：`creditcardinfo/Script/run_all_banks.py`、`creditcardinfo/Script/ctbc_creditcard_disclosure.py`、
`creditcardinfo/Script/fubon_creditcard_disclosure.py`、`creditcardinfo/Script/cathay_creditcard_disclosure.py`、
`creditcardinfo/Script/esun_creditcard_disclosure.py`、`creditcardinfo/Script/taishin_creditcard_disclosure.py`、
`creditcardinfo/Script/feib_creditcard_disclosure.py`、`creditcardinfo/input/中信信用卡<yyyymm>.xlsx`；timeout_secs=600
（6×95=570 ≤ 600）

**步驟 2：抓取 B 組（3 家）**
```
python creditcardinfo/Script/run_all_banks.py --banks dbs,ubot,sinopac \
  --month <月> --timeout 120 --json-output out/banks_groupB_<yyyymm>.json
```
input_files（4 個）：`creditcardinfo/Script/run_all_banks.py` + 該 3 支 `*_creditcard_disclosure.py`；timeout_secs=450

**步驟 3：firstbank 單獨跑**（內建三段 TLS 重試，最壞約 234 秒，必須給足逾時、不能跟別家併組）
```
python creditcardinfo/Script/run_all_banks.py --banks firstbank \
  --month <月> --timeout 250 --allow-insecure --json-output out/banks_firstbank_<yyyymm>.json
```
input_files（2 個）：`creditcardinfo/Script/run_all_banks.py`、`creditcardinfo/Script/firstbank_creditcard_disclosure.py`；timeout_secs=300

**步驟 4：合併 summary（不重抓）**
```
python creditcardinfo/Script/run_all_banks.py \
  --collect-only out/banks_groupA_<yyyymm>.json out/banks_groupB_<yyyymm>.json out/banks_firstbank_<yyyymm>.json \
  --json-output out/summary_<yyyymm>.json
```
input_files（4 個）：`creditcardinfo/Script/run_all_banks.py` + 三個抓取結果 JSON；timeout_secs=60
→ `missing_banks` 應為空、`failed_count` 應為 0；有 failed 先照「情境 B」補跑該銀行再重做本步驟。

**步驟 5：一條龍回寫工作簿（不要加任何 `--skip-*`）**
```
python creditcardinfo/Script/update_credit_card_workbook.py \
  --workbook 銀行局信用卡公開資料.xlsx --summary out/summary_<yyyymm>.json
```
input_files（9 個）：`creditcardinfo/Script/update_credit_card_workbook.py`、
`creditcardinfo/Script/run_market_total.py`、`creditcardinfo/Script/jcic_avg_cards_update.py`、
`creditcardinfo/Script/run_bank_bureau_bank_backfill.py`、`creditcardinfo/Script/bank_aliases.py`、
`creditcardinfo/Script/percent_utils.py`、`creditcardinfo/Script/workbook_block_helpers.py`、
`out/summary_<yyyymm>.json`、`銀行局信用卡公開資料.xlsx`；timeout_secs=900
（若 summary 裡 ctbc 缺「當月轉帳卡簽帳金額」，再多帶第 10 個檔 `中信信用卡資料.xlsx`）

不帶 `--month` 時腳本自動取各行 `data_month` 最新月為 target；內部順序固定：
**寫入最新月 block → 銀行局 backfill（補日曆缺口與舊月空白、修 Top5/Top10 公式）→ 市場總計 → JCIC → 年度整理（該年 1–12 月齊全才觸發）**。

**判讀結果（看步驟 5 的 stdout JSON，不要自己寫 openpyxl 驗證）**
- `verification.status == "ok"` → 成功；`"mismatch"` → 列出 `mismatches` 回報使用者
- `warnings` 有「資料月份與目標月不符，已跳過」屬正常（落後銀行交給 backfill），照實回報
- 若一開始就報「**工作簿存在不完整的月 block…請先執行 --repair-partial-blocks**」→ 先跑情境 D 再重跑本步驟

---

## 情境 B：單獨補跑某一家銀行

```
python creditcardinfo/Script/run_all_banks.py --banks <bank> --month <月> --timeout 250 \
  --json-output out/<bank>_retry_<yyyymm>.json
```
（bank = ctbc/fubon/cathay/esun/taishin/dbs/ubot/sinopac/firstbank/feib；firstbank 加 `--allow-insecure`）
input_files：`run_all_banks.py` + 該銀行腳本（ctbc 另加 `creditcardinfo/input/中信信用卡<yyyymm>.xlsx`）；timeout_secs=300
然後用 collect-only 合併（後面的檔案覆蓋同銀行舊結果）：
```
python creditcardinfo/Script/run_all_banks.py \
  --collect-only out/summary_<yyyymm>.json out/<bank>_retry_<yyyymm>.json \
  --json-output out/summary_<yyyymm>.json
```

## 情境 C：回補舊月份 / 補空白

交給一條龍（步驟 5 指令加 `--month <目標月>`）或單獨跑 backfill：
```
python creditcardinfo/Script/run_bank_bureau_bank_backfill.py \
  --workbook 銀行局信用卡公開資料.xlsx --month <月>
```
input_files（5 個，另可選本機 ZIP）：backfill 腳本、`bank_aliases.py`、`percent_utils.py`、
`workbook_block_helpers.py`、工作簿；timeout_secs=600
- `--month`＝只補該單一月（會建立該月 block）；`--newest-month`＝以此為錨點往前掃日曆缺口
- 該月官方 ZIP 未發布會 `skip_missing_banks`，不會建空 block，屬正常
- 回補 **13 欄**（卡數 4＋金額 6＋比率 3），全為必要欄位，來源缺任一欄該月整月報錯；預設只補空白、不覆蓋既有值
- 回補下限固定為 202601

## 情境 D：修復不完整的月 block（孤兒列）

```
python creditcardinfo/Script/update_credit_card_workbook.py \
  --workbook 銀行局信用卡公開資料.xlsx --repair-partial-blocks
```
input_files（5 個）：`update_credit_card_workbook.py`、`bank_aliases.py`、`percent_utils.py`、
`workbook_block_helpers.py`、工作簿；timeout_secs=300
- 只會刪除**位於工作簿尾端**的 partial block（被移除的值列在輸出 JSON 的 `deleted`）
- `needs_manual_review` 非空 → 停下來回報使用者，不要自己動表中列
- 修復後把被刪的月份用情境 C 或步驟 5 重建

---

## 腳本快查表

| 腳本 | 角色 / 資料來源 | 關鍵 CLI | SSL | 沙箱相依（input_files） |
|---|---|---|---|---|
| `run_all_banks.py` | 批次抓取＋彙整（自帶 percent 函式） | `--banks` `--month` `--timeout` `--json-output` `--collect-only` `--preflight` `--allow-insecure` | 不上網 | 自身＋要跑的銀行腳本 |
| `ctbc_...py` | **本機檔** `input/中信信用卡<yyyymm>.xlsx`，不上網 | `--month` `--input` | — | 自身＋來源 xlsx（需 openpyxl） |
| `fubon_...py` | 官網 PDF（需 pypdf） | `--month` | 恆用 insecure | 自身 |
| `dbs_...py` | 官網 PDF（需 pypdf） | `--month` | 自動 fallback | 自身 |
| `cathay` `esun` `taishin` `ubot` `sinopac` `feib` `_...py` | 官網 HTML / API | `--month` | 自動 fallback | 自身 |
| `firstbank_...py` | 官網 HTML（內建重試最長 ~170s） | `--month` `--allow-insecure` `--retries` `--timeout` | 第三段需 `--allow-insecure` | 自身 |
| `update_credit_card_workbook.py` | 主控：寫月 block＋委派＋年度整理＋讀回驗證 | `--workbook` `--summary` `--month` `--repair-partial-blocks` `--annual-only`（`--skip-*` 平常不要用） | 不上網 | 自身＋3 委派腳本＋3 共用模組＋summary＋工作簿 |
| `run_market_total.py` | 金管會 ZIP → `市場總計(銀行局)`＋`銀行局.市場總計` | `--workbook` `--base-month` `--lookback-months` `--zip-file` | 自動 fallback | 自身＋`percent_utils.py`＋`bank_aliases.py`＋`workbook_block_helpers.py`＋工作簿 |
| `jcic_avg_cards_update.py` | JCIC CSV → `平均每人持卡張數` | `--workbook` `--base-month`（必填） `--backfill-months` | 自動 fallback | 自身＋`bank_aliases.py`＋`workbook_block_helpers.py`＋工作簿 |
| `run_bank_bureau_bank_backfill.py` | 金管會 ZIP → 補舊月 10 行空白（13 欄）＋修公式 | `--workbook` `--month` / `--newest-month` `--lookback-months` | 自動 fallback | 自身＋3 共用模組＋工作簿 |

共用模組（不可執行）：`percent_utils.py`、`bank_aliases.py`、`workbook_block_helpers.py`。
**月 block 只由 update / backfill 建立**；market 與 jcic 對不存在的月份一律跳過並回報 warning。

---

## 執行前必檢查

- [ ] 目標月份已確認（`115年05月` 格式；未指定就讓 summary 推定最新月）
- [ ] **中信本機來源檔已就位**（`creditcardinfo/input/中信信用卡<yyyymm>.xlsx`）
- [ ] 目標月官方資料已發布（未發布月抓回來是 HTML 錯誤頁；backfill 會自動 skip）

## 常見卡點與 fallback

1. **input_files ≤ 10 檔**：`--banks all` 需要 12+ 檔，必然超限，一律照情境 A 分組跑再 `--collect-only` 合併。
2. **`__pycache__` / `.pyc`**：腳本已設 `sys.dont_write_bytecode`、子行程帶 `-B`；**永遠不要把 `.pyc` 加進 input_files**。
   若仍遇到「Refusing to overwrite ... .pyc」，先清空 `creditcardinfo/Script/__pycache__/`。
3. **cwd 不是 repo root**：一律用 repo-relative 路徑（`creditcardinfo/Script/xxx.py`）。
4. **不要自己直連下載金管會 ZIP / JCIC CSV**：URL quote、憑證、未發布月回 HTML 等問題腳本都處理好了。
5. **結果判讀看 stdout JSON 的 `verification` 段**：不要手寫 openpyxl 逐格驗證；真要抽查只讀目標 `YYYYMM` 的 13 列。
6. **多月份 data_month 不是錯誤**：update 取最新月為 target，落後銀行列進 `warnings` 留給 backfill。
7. **timeout 數學**：run_all 循序執行，總時間 ≈ 銀行數 × `--timeout`，不可超過工具的 900 秒上限。

---

## 資料模型關鍵事實（判讀 / 改碼時用）

- `市場總計(銀行局)`＝回寫結果列；`銀行局`＝公式列（Top5/Top10，要用腳本修復）；`財團法人金融聯合徵信中心`＝JCIC 補值列
- 年度 block 的 `YYYYMM` 欄放 4 位數年份、`月份` 欄為 `Total`；年度整理只在該年 1–12 月齊全時自動觸發
- **百分比欄單位鐵則**：JSON 與工作簿內部一律存**小數比率**，Excel 靠 `0.00%` 格式顯示。
  逾期三/六個月比率的來源都是百分比顯示數字（`0.12` 代表 0.12%）→ 解析時**一律 /100**（`force_percent_input=True`）；
  唯一例外是儲存格本身是 % 格式（如中信 xlsx raw `0.0012`）→ 直接保留。中信腳本這兩欄在解析階段已轉小數，
  `metric_units` 宣告 `decimal_ratio`。工作簿出現 `0.12`（顯示 12%）、`17`、`85659` 這類值代表 parser 壞了要先修；
  備抵呆帳提足率（值恆 >1，如 `429.67`→`4.2967`）沿用 abs>1 才 /100 的啟發式
- **中信來源檔 14 個項目順序固定**（流通卡數→有效卡數→當月發卡數→當月停卡數→循環信用餘額→當月簽帳金額→
  當月預借現金金額→逾期三個月比率→逾期六個月比率→備抵呆帳提足率→當月轉銷呆帳金額→當年度累計轉銷呆帳金額→
  未到期分期付款餘額→當月轉帳卡簽帳金額），照固定順序抓

## 不要用這個 skill 的情況

信用卡推薦、權益比較、回饋排行、申辦資格、行銷文案，以及與本資料夾揭露流程無關的通用問題。

## 維護提醒

架構與分工看 `README.md`；改完 SKILL.md 記得讓部署端同步並更新頂部版本戳記。
