# 派出所綜合管制系統 v2.2

> 全台 1300+ 派出所通用的 Excel 管制系統 ｜ 無巨集 ｜ 一檔搞定刑案/毒調/交通/績效/歷史

[![License: CC BY 4.0](https://img.shields.io/badge/License-CC%20BY%204.0-lightgrey.svg)](https://creativecommons.org/licenses/by/4.0/)
[![Version: v2.2](https://img.shields.io/badge/Version-v2.2-blue.svg)](#)
[![Tests](https://img.shields.io/badge/Tests-1426%2F1426-brightgreen.svg)](#)
[![Excel](https://img.shields.io/badge/Excel-2016%2B-success.svg)](#)

---

## 📦 直接下載使用

點下面任一個檔案 → 右上「Download raw file」即可下載：

| 檔案 | 用途 | 大小 |
|---|---|---|
| **[PoliceStation_v2.2_部署版.xlsx](build/output/PoliceStation_v2.2_部署版.xlsx)** | ⭐ **派出所首次安裝（空白資料）** | 132.8 KB |
| [PoliceStation_v2.1.xlsx](build/output/PoliceStation_v2.1.xlsx) | 含 sample 示範資料版 | 134.8 KB |
| [📖 README_部署手冊_v2.2.md](docs/README_部署手冊_v2.2.md) | 部署 SOP（必讀）| — |

---

## 🎯 設計理念（三柱）

1. **功能覆蓋完整度** — 五大領域（刑案/毒調/交通/績效/歷史）一檔搞定
2. **可擴充性 L1** — 設定表清單值層全可調，全台 1300+ 派出所都能直接用
3. **專業數據呈現** — 麥肯錫顧問風（A）+ 政府公文風（B）雙風格混搭

服務雙對象：長官 5 秒看完（A 風格）／ 上級機關 30 分鐘審閱（B 風格）

---

## 📊 17 分頁完整功能

| # | 分頁 | 角色 | 風格 |
|---|---|---|---|
| 1 | 首頁_操作說明 | 第一次拿到檔案的員警 | A 麥肯錫 |
| 2 | 管制總覽 | 所長 30 秒掃完全所 | A — 9 KPI 九宮格 + sparkline |
| 2.5 | 全般刑案管制情形分析 | 上級機關承辦 | A+B 章戳 + 雙環形圖 |
| 3 | 長官 5 分鐘報告 | 督導官 | A — Pyramid Principle |
| 4 | 對上級機關上呈 | 對分局/警察局/警政署 | B — CHOOSE 切三格式 |
| 5 | 案件資料庫 | 員警 — 唯一輸入點 SSOT | B — 25 欄 1000 列 |
| 6 | 刑案管制 | 偵查佐 | B — 兩區段並呈 |
| 7 | 發展中案件管制 | 偵查佐 | B — 70/30 主表+候選 |
| 8 | 毒品調驗人口管制 | 偵查佐 | B — v2.2 加最後到驗日精確 30 日 |
| 9 | 交通績效管制 | 所長 | B — 6 違規達成率看板 |
| 9.5 | 交通取締明細 | 交通組員警 | B |
| 10 | 績效統計 | 報表組 | B — 案類 + Top 10 員警 |
| 11 | 員警個人績效卡 | 督導官評鑑 | A — A4 一頁交付 |
| 12 | 跨期間趨勢分析 | 所長/督導 | A — 5 KPI × 12 月走勢 |
| 13 | 設定表 | 部署人員/承辦 | B — 8 大分區 |
| 14 | 歷史資料 | 承辦人月底抄寫 | B — v2.2 加即時 KPI 抄寫區 |
| 15 | 系統檢核 | 督導稽核 | B — 45 條規則 7 大類 |

---

## 🚀 15 分鐘安裝

1. 下載 [`PoliceStation_v2.2_部署版.xlsx`](build/output/PoliceStation_v2.2_部署版.xlsx)
2. 開檔 → 跳到「設定表」分頁
3. 填 **A 區 三層組織**：警察局/分局/派出所/英文名
4. 填 **B 區 員警名冊**（最多 60 名）
5. 改密碼（預設 `KD2026`）：審閱 → 取消保護工作表 → 設新密碼
6. 跳到「**系統檢核**」確認紅燈 = 0

詳細 SOP 看 [docs/README_部署手冊_v2.2.md](docs/README_部署手冊_v2.2.md)。

---

## 📅 月例行作業（5-10 分/月）

### 月底承辦人
- 跳到「歷史資料」row 6 看「🚨 本月尚未抄寫」紅燈
- **選 A7:N7 整列 → Ctrl+C → 到下方主表第一筆空白列 → Ctrl+Shift+V → 值**
- 紅燈自動轉綠

### 員警日常
- 跳到「案件資料庫」第一筆空白列 → key 入 22 欄
- 案件狀況 4 選 1：尚未偵破 / 已破獲未移送 / 已移送 / 簽結
- **雙軌欄位**「發生管轄」+「查獲管轄」要分清

---

## 🏗 技術架構

### 純公式 / 無巨集
- 開檔不會彈宏警示
- 11 個 sheet 受密碼保護（預設 `KD2026`）
- 案件資料庫輸入欄解鎖，22 欄結構化引用 `[@欄名]`

### 4 大資料表（SSOT）
- `Tbl案件` 25 欄 / 1000 列容量
- `Tbl交通` 6 欄 / 196 列
- `Tbl毒調` **14 欄**（v2.2 加最後到驗日 + 距今天數）/ 70 列
- `Tbl歷史` 14 欄 / 140 列（11 年資料）

### 32 個命名範圍
組織抬頭、員警名冊、案類分類、違規項目、閾值（毒調率目標、新進/老案天數）、今日（TODAY 集中策略）

### 條件格式 7 大規範
- 紅黃綠燈動態染色（頁 2 KPI 9 卡 × 6 條 / 頁 6/7/9 各類 / 頁 15 系統檢核）
- 容量警示（案件 ≥800 黃 / ≥950 紅）
- 抄寫狀態（已抄寫綠 / 未抄寫紅）

### 真原生 Sparkline
- 頁 2 七條 native line sparkline（Excel 2010+）
- XML 後處理注入 `extLst`（openpyxl 不支援直接寫）

### 系統檢核 45 條 7 大類
- A 部署完整性 5 / B 資料合理性 8 / C 業務邏輯 10
- D 時效警示 8 / E 跨頁一致性 6 / F 性能監控 5 / G 環境檢測 3
- 每條附 HYPERLINK 跳轉到問題頁

---

## 🛠 開發者指南

### 環境
- Python 3.12
- openpyxl 3.1.5
- 工具：`build/scripts/build.py --phase {1|1.5|2|3.*|4.*|5|6|8|9|11|all}`

### 重 build 全檔
```powershell
$env:PYTHONIOENCODING = 'utf-8'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
python build/scripts/build.py --phase all
```

### 跑全 14 個測試（1426 條）
```powershell
Get-ChildItem build/scripts/tests/test_phase*.py |
  ForEach-Object { python $_.FullName }
```

### Phase 列表
```
1     17 工作表骨架
1.5   設定表 8 大分區
2     三大資料表（案件/交通/歷史）
3     頁 1 首頁 + 頁 7 發展中
3.2   頁 2 管制總覽（九宮格 + sparkline）
3.25  頁 2.5 全般刑案分析（雙環形圖）
3.6   頁 6 刑案管制（兩區段並呈）
3.8   頁 8 毒調（Tbl毒調 14 欄）
3.9   頁 9 交通看板
3.10  頁 10 績效統計
3.11  頁 11 員警個人卡
3.12  頁 12 跨期間趨勢
4.3   頁 3 長官 5 分鐘報告
4.4   頁 4 對上級機關上呈
8     頁 14 即時 KPI 抄寫區
6     頁 15 系統檢核 45 條
5     條件格式 + 工作表保護 + 列印區
9     真原生 Excel Sparkline（XML 注入）
11    部署前清資料（產 v2.2_部署版.xlsx）
```

### 14 大踩雷點
詳見 [docs/00_FROM_HERE_v2_2.md](docs/00_FROM_HERE_v2_2.md)：
1. `unmerge_all_in_sheet` 後必重 merge_cells
2. openpyxl 無 sparkline API（要 XML 注入）
3. 結構引用 `[@欄名]` 欄名要完全一致
4. 日期 sample 必須是 datetime 物件
5. 17 頁全 freeze_panes = None
6. PowerShell UTF-8 環境變數
7. openpyxl 公式不可起始空白
8. 條件格式套法 FormulaRule
9. **CF Font 只能用 bold/color/italic**（不接受 name/size）
10. **CF 公式不支援結構化引用 `Tbl...[]`**（要用普通範圍）
11. **Merged cell `protection` 只看 anchor**（左上格）
12. openpyxl `add_table` 後不可重設 calc cell
13. CF 規則非冪等（重 build 要先 `_cf_rules = {}`）
14. Phase 5 保護必須在 Phase 6/8 填內容之後跑

---

## 🧪 測試結果（1426/1426 全綠燈）

| 測試檔 | 通過 |
|---|---|
| test_phase3_page2.py | 100/100 |
| test_phase3_page2_5.py | 77/77 |
| test_phase3_page6.py | 103/103 |
| test_phase3_page8.py | 102/102 |
| test_phase3_page9.py | 46/46 |
| test_phase3_page10.py | 173/173 |
| test_phase3_page11.py | 128/128 |
| test_phase3_page12.py | 119/119 |
| test_phase4_page3.py | 51/51 |
| test_phase4_page4.py | 40/40 |
| test_phase5_polish.py | 58/58 |
| test_phase6_inspection.py | 298/298 |
| test_phase8_history.py | 33/33 |
| test_phase10_system_stress.py | 98/98 |

每個測試含 6 大類：公式語法 / 邊界資料 / 性能基準 / 跨平台 / 真實情境 / 錯誤恢復

---

## 📞 回報 Bug / 建議

- [GitHub Issues](https://github.com/final90006-sketch/police-station-control-v2/issues)
- 工程方：林錦瑞

---

## 📜 授權

**CC BY 4.0** — 可自由使用、修改、商用，需標示作者。詳見 [LICENSE](LICENSE)。

> 派出所綜合管制系統 v2.2 — 林錦瑞 (2026)
> 取自 https://github.com/final90006-sketch/police-station-control-v2

---

© 林錦瑞 ｜ KKEVIN-LIN-2026-V2.2 ｜ 2026-05-26
