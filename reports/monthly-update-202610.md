# 月更新檢討報告：115年10月（GitHub Actions）

- 產生時間：2026-11-18T07:58:09
- 目標月：115年10月（YYYYMM 202610），block 新建，起始列 1014
- 銀行抓取：成功 10／部分 0／失敗 0，缺少 []
- 讀回驗證：ok；百分比稽核異常 0 格（掃 92 個 block）
- 回寫總耗時：33.0 秒

## 各銀行抓取

| 銀行 | 狀態 | 資料月 | 秒數 | 重試 | 錯誤 |
|---|---|---|---|---|---|
| 台北富邦銀行 | success | 115年10月 | 5.0 |  |  |
| 中國信託商業銀行 | success | 115年10月 | 4.0 |  |  |
| 國泰世華銀行 | success | 115年10月 | 4.0 |  |  |
| 台新銀行 | success | 115年10月 | 4.0 |  |  |
| 玉山銀行 | success | 115年10月 | 3.0 |  |  |
| 永豐銀行 | success | 115年10月 | 2.0 |  |  |
| 遠東商銀 | success | 115年10月 | 2.0 |  |  |
| 星展銀行 | success | 115年10月 | 1.0 |  |  |
| 聯邦銀行 | success | 115年10月 | 1.0 |  |  |
| 第一銀行 | success | 115年10月 | 1.0 |  |  |

抓取合計約 27.0 秒（循序）。

## 回寫各階段耗時

| 階段 | 秒數 |
|---|---|
| write_month_block | 3.5 |
| bank_bureau_backfill | 3.6 |
| market_total | 9.0 |
| jcic_avg_cards | 7.5 |
| annual_sync | 3.4 |
| verification | 2.7 |
| percent_audit | 3.2 |
| total | 33.0 |

- 銀行局 backfill：子程序 3.6 秒，處理 0 個月份，動作 []
- 市場總計：子程序 9.0 秒，處理 2 個月份，動作 ['latest_month_update', 'skip_source_unavailable']
  - 來源已下架或尚未發布、無法回補的月份：[202504]（舊月份屬已知永久缺口，最新月等下月再補）
- JCIC：子程序 7.5 秒，處理 0 個月份，動作 []，最新可用 115年07月

## 驗證與稽核

- verification：ok，驗證 10 家，mismatch 0 格，percent_anomalies 0 格
- percent_audit：0 格異常

## Token 與成本

本次未使用 Claude token（GitHub Actions 執行），或未提供開始／結束兩筆用量快照。

## 檢討與建議

- 本次沒有需要改進的環節。


---
PR：https://github.com/grace310368/credit-card-disclosure-v2/pull/7
