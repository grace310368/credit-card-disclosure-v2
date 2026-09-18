# 月更新檢討報告：115年08月（GitHub Actions）

- 產生時間：2026-09-18T05:38:07
- 目標月：115年08月（YYYYMM 202608），block 既有，起始列 992
- 銀行抓取：成功 10／部分 0／失敗 0，缺少 []
- 讀回驗證：ok；百分比稽核異常 0 格（掃 90 個 block）
- 回寫總耗時：53.1 秒

## 各銀行抓取

| 銀行 | 狀態 | 資料月 | 秒數 | 重試 | 錯誤 |
|---|---|---|---|---|---|
| 台北富邦銀行 | success | 115年08月 | 5.0 |  |  |
| 台新銀行 | success | 115年08月 | 5.0 |  |  |
| 玉山銀行 | success | 115年08月 | 4.0 |  |  |
| 中國信託商業銀行 | success | 115年08月 | 3.0 |  |  |
| 國泰世華銀行 | success | 115年08月 | 3.0 |  |  |
| 遠東商銀 | success | 115年08月 | 3.0 |  |  |
| 永豐銀行 | success | 115年08月 | 2.0 |  |  |
| 第一銀行 | success | 115年08月 | 2.0 |  |  |
| 星展銀行 | success | 115年08月 | 1.0 |  |  |
| 聯邦銀行 | success | 115年08月 | 1.0 |  |  |

抓取合計約 29.0 秒（循序）。

## 回寫各階段耗時

| 階段 | 秒數 |
|---|---|
| write_month_block | 3.2 |
| bank_bureau_backfill | 3.4 |
| market_total | 32.0 |
| jcic_avg_cards | 5.8 |
| annual_sync | 3.3 |
| verification | 2.6 |
| percent_audit | 2.8 |
| total | 53.1 |

- 銀行局 backfill：子程序 3.4 秒，處理 0 個月份，動作 []
- 市場總計：子程序 32.0 秒，處理 2 個月份，動作 ['latest_month_update', 'skip_fetch_failed']
- JCIC：子程序 5.8 秒，處理 0 個月份，動作 []，最新可用 115年04月

## 驗證與稽核

- verification：ok，驗證 10 家，mismatch 0 格，percent_anomalies 0 格
- percent_audit：0 格異常

## Token 與成本

本次未使用 Claude token（GitHub Actions 執行），或未提供開始／結束兩筆用量快照。

## 檢討與建議

- 市場總計 回補 202504 失敗：BadZipFile: https://www.fsc.gov.tw/userfiles/file/11404_%E4%BF%A1%E7%94%A8%E5%8D%A1%E9%87%8D%E8%A6%8。若該月金管會檔已下架屬正常，否則要查。
