# 派出所綜合管制系統 v2.2 — 交接包（給下個 Claude）

> **這份取代舊的 00_FROM_HERE_v2_1.md。** 前面 session 已把 v2.1 升級到 v2.2，加了 Phase 5（保護+CF+列印）+ Phase 6（系統檢核 45 條）+ Phase 7（Tbl毒調 7 條接通 + 頁 2 KPI 4/5 接通）+ Phase 8（頁 14 即時 KPI 抄寫區）。

製作日期：2026-05-26 ｜ 上一 Claude session 結束時 context ~80%

---

## 🎯 一句話交接

派出所 v2 已 **97% 完成（工程建檔層）**。17 頁全交付 + 4 大資料表 + 45 條系統檢核 + 工作表保護 + 列印區 + 即時 KPI 抄寫區 + 1319/1319 公式測試全綠。**剩 3% 是交接文件 + 部署準備**；剩下要做的是實機跨平台測試（Mac / LibreOffice）、列印實測、三所試用 1 月。

---

## ✅ v2.2 已完成項目（本 session 新交付）

### Phase 5 — 條件格式 + 工作表保護 + 列印區

| 項目 | 結果 |
|---|---|
| **頁 6 條件格式** | 4 條（兩區段案件狀況淡紅/淡黃自動染色）|
| **頁 7 條件格式** | 5 條（K 欄狀態三色 + 整列依狀態淡染）|
| **頁 9 條件格式** | 6 條（D 欄達成率三色 + F 欄狀態三色）|
| **頁 5 案件資料庫容量警示** | A2 動態公式 + 2 條 CF（≥800 黃 / ≥950 紅）|
| **工作表保護 11 sheets** | 密碼 KD2026；input 欄解鎖、calc 欄鎖 |
| **列印區 4 頁** | 頁 2.5 / 3 / 4 / 11 → A4 直式 1 頁 |

### Phase 6 — 頁 15 系統檢核 45 條 7 大類

| 類別 | 條數 |
|---|---|
| A 部署完整性 | 5 |
| B 資料合理性 | 8 |
| C 業務邏輯 | 10 |
| D 時效警示 | 8 |
| E 跨頁一致性 | 6 |
| F 性能監控 | 5（資訊性）|
| G 環境檢測 | 3（資訊性）|

- 頂部健康度 4 大字：紅燈總數 / 黃燈總數 / 綠燈總數 / 其他
- 7 大類縮影條（row 11 深底白字類別名 + row 12 紅 X 黃 Y 綠 Z）
- 每條規則含 HYPERLINK 跳轉到問題頁
- 6 條條件格式（紅/黃/綠/灰 + 整列淡染）

### Phase 7 — Tbl毒調 7 條 pending 規則 + 頁 2 KPI 接通

| 規則 | 公式 |
|---|---|
| B7 身分證長度 = 10 | `LEN(Tbl毒調[身分證])≠10` 計數 |
| C4 毒調率分子 ≤ 列管總數 | `SUM(是否到驗) - COUNTA(姓名) ≤ 0` |
| C10 強採案必到驗 | 含「強採」但是否到驗=0 計數 |
| D2 毒調容量 | `COUNTA(姓名)` ≥50 黃 / ≥65 紅（上限 70）|
| D5 未到驗件數 | `COUNTIF("*未到驗*")` 0綠/1-5黃/>5紅 |
| D8 已解除列管 | `COUNTIF("*解除*")` 資訊性 |
| E6 毒調率 ≥ 目標 | 即時 vs 設定表「毒調率目標」 |
| **頁 2 KPI 4 毒調率** | 從「—」變實際 %（依毒調率目標判燈號）|
| **頁 2 KPI 5 未到驗** | 從「—」變實際件數 |

### Phase 8 — 頁 14 即時 KPI 抄寫區

- Row 6 動態標題 + 抄寫狀態（CF 紅綠燈：已抄寫綠 / 未抄寫紅）
- Row 7 14 欄即時 KPI 公式（淡黃底，提示「這列要複製」）
  - 統計年月 / 全般發生 / 未破 / 破獲 / 破獲率 / 竊盜 / 詐欺 / 已破未送 / 發展中 / 毒調率 / 未到驗 / 交通達標 / 員警冠軍 / 備註
  - 用 `DATE` + `EOMONTH` 取當月期間（Excel 2007+ 全相容）
- Row 8 操作流程（① 複製 A7:N7 ② 貼下方 ③ Enter）

## 🧪 高強度測試現況

**全 13 個測試檔 1319/1319 全綠燈**

| 測試檔 | 通過 |
|---|---|
| test_phase3_page2.py | 100/100 |
| test_phase3_page2_5.py | 77/77 |
| test_phase3_page6.py | 103/103 |
| test_phase3_page8.py | 94/94 |
| test_phase3_page9.py | 46/46 |
| test_phase3_page10.py | 173/173 |
| test_phase3_page11.py | 128/128 |
| test_phase3_page12.py | 119/119 |
| test_phase4_page3.py | 51/51 |
| test_phase4_page4.py | 40/40 |
| **test_phase5_polish.py** | 58/58 |
| **test_phase6_inspection.py** | 298/298 |
| **test_phase8_history.py** | 32/32 |

---

## 📁 工程目錄

```
派出所v2_build/
├── config/
│   ├── phase1_skeleton.json
│   ├── phase1_5_settings.json
│   └── phase2_data_tables.json
├── scripts/
│   ├── build.py                    CLI: --phase {...|5|6|8|all}
│   ├── styles.py
│   ├── phases/
│   │   ├── _base.py                + unmerge_all_in_sheet() helper
│   │   ├── phase1_skeleton.py
│   │   ├── phase1_5_settings.py
│   │   ├── phase2_data_tables.py
│   │   ├── phase3_pages_1_7.py
│   │   ├── phase3_page2_overview.py    ★ v2.2 KPI 4/5 接通
│   │   ├── phase3_page2_5_crime_analysis.py
│   │   ├── phase3_page6_case_control.py
│   │   ├── phase3_page8_drug.py
│   │   ├── phase3_page9_traffic.py
│   │   ├── phase3_page10_performance.py
│   │   ├── phase3_page11_officer.py
│   │   ├── phase3_page12_trend.py
│   │   ├── phase4_page3_brief.py
│   │   ├── phase4_page4_upreport.py
│   │   ├── phase5_polish.py            ★ v2.2 新
│   │   ├── phase6_inspection.py        ★ v2.2 新
│   │   └── phase8_history_writeback.py ★ v2.2 新
│   └── tests/
│       ├── framework.py
│       ├── test_phase3_*.py
│       ├── test_phase4_*.py
│       ├── test_phase5_polish.py       ★ v2.2 新
│       ├── test_phase6_inspection.py   ★ v2.2 新
│       └── test_phase8_history.py      ★ v2.2 新
├── output/
│   ├── PoliceStation_v2.0.xlsx    最新 build 輸出（v2.2 內容）
│   └── PoliceStation_v2.1.xlsx    spec 命名版（cp 自 v2.0）
└── log/
```

## 🚀 開發/驗證指令

```powershell
$env:PYTHONIOENCODING = 'utf-8'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$py = 'C:\Users\User\AppData\Local\Programs\Python\Python312\python.exe'
$build = 'C:\Users\User\Desktop\02【工作】案件管制與自動化系統\派出所v2_v43_開發檔案\派出所v2_build'

# 一鍵 build 全部（17 個 phase）
& $py "$build\scripts\build.py" --phase all

# v2.2 新 phase
& $py "$build\scripts\build.py" --phase 5    # 保護 + CF + 列印
& $py "$build\scripts\build.py" --phase 6    # 系統檢核 45 條
& $py "$build\scripts\build.py" --phase 8    # 即時 KPI 抄寫區

# 跑全 13 個測試（1319 條）
foreach ($t in (ls "$build\scripts\tests\test_phase*.py")) {
    & $py $t.FullName 2>&1 | Select-String "總計" | Select-Object -Last 1
}

# 複製 v2.0 → v2.1（spec 命名）
Copy-Item "$build\output\PoliceStation_v2.0.xlsx" "$build\output\PoliceStation_v2.1.xlsx" -Force
```

**⚠ Phase 順序很重要**：`--phase all` 中 8 → 6 → 5 順序確保「先填內容、後上保護」。dict 插入順序在 build.py 中已設好。

---

## ⚠ 已知踩雷點（v2.2 補充 — 全 12 點）

### 1. ★ `unmerge_all_in_sheet(ws)` 是必要前置動作
每個頁面 builder 開頭呼叫 `unmerge_all_in_sheet(ws)`，**記得 ws.merge_cells(...) 重新建 banner merge**（曾在頁 7 漏掉，banner 文字塞 col A）。

### 2. ★ openpyxl 沒有 sparkline API
頁 2 + 頁 12 用 **Unicode 區塊字元**（`▁▂▃▄▅▆▇█`）+ CHOOSE+CONCAT 公式，Consolas 等寬字。

### 3. ★ 結構引用 [@欄名] 對欄名要完全一致
Tbl案件 header 不可加任何前綴（★ 標示等）。

### 4. ★ 日期 sample 必須是 datetime 物件
phase2_data_tables.py 有 `_coerce_value()` 處理。

### 5. ★ 17 頁全 freeze_panes = None
使用者明確要求。設定 None 後易產生 orphan `<selection pane="bottomLeft">` 警示——一律從零重 build。

### 6. ★ 用戶 PowerShell 環境
必加：`$env:PYTHONIOENCODING = 'utf-8'` + `[Console]::OutputEncoding = [System.Text.Encoding]::UTF8`。

### 7. ★ openpyxl 公式 cell 開頭不可有空白
寫 `'   ="..."` 會被 Excel 當文字。公式必須以 `=` 起頭。

### 8. ★ 條件格式套法（FormulaRule）
```python
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill, Font

ws.conditional_formatting.add("E8",
    FormulaRule(formula=['ISNUMBER(SEARCH("紅",E8))'],
                fill=PatternFill("solid", fgColor="FFFEE2E2"),
                font=Font(name="微軟正黑體", bold=True, color="FFB91C1C")))
```

### 9. ★★ CF Font 只能用 bold/color/italic/underline/strike
**不接受 `name=`、`size=`**（其實 name 可以但 size 不可）。設了會被 Excel 開檔時「修復」掉整個 CF 區塊，跳「我們發現部分內容有問題」警示。
（v2.2 修：phase5 cf_page5_capacity 原本用 `size=11, italic=True` 觸發此 bug）

### 10. ★★★ CF 公式 **不支援 Excel Table 結構化引用**
寫 `COUNTA(Tbl案件[編號])>=950` 在 CF 公式中會被 Excel 開檔時整段移除 → 「已移除的功能 sheet6.xml CF」警示。
**改用普通範圍引用**：`COUNTA($A$15:$A$1014)>=950`（cell 內公式可以用結構化引用，CF 公式不行）
（v2.2 大坑，phase5 cf_page5_capacity / phase8 row 6 CF 都中過）

### 11. ★★ Merged cell 的 `cell.protection` 只看 anchor（左上格）
phase 5 解鎖頁 3 行動方案區用 `_unlock_range(ws, "B24:G27")` 不生效，因為 row 24-27 是 `A:H` merged，anchor 在 A 不在 B。
改為 `_unlock_range(ws, "A24:H27")` 才對。

### 12. ★ openpyxl `add_table` 後不可重設 cell.value 為公式字串
若 Excel Table 已建好，再對 calc 欄寫入公式會被 Table style 覆蓋。應該在 phase 2 建表時就一次寫入公式（已落實）。

### 13. ★ CF 規則 非冪等（重 build 會疊加）
openpyxl `conditional_formatting.add` 是累加式。Phase 5 每個 cf_pageN 函數要先 `ws.conditional_formatting._cf_rules = {}` 才冪等。
（v2.2 fix：phase5/phase6/phase8 各 builder 開頭都已加 _clear_all_cf）

### 14. ★ Phase 5（保護）必須在 Phase 6/8（填內容）之後跑
否則 Phase 5 把 sheet 鎖了，後面 Phase 6/8 也能寫（openpyxl 不受 sheet protection 限制），但邏輯顛倒。build.py PHASES dict 順序已設好（8 → 6 → 5）。

---

## 📌 業務規則（v2.2 沿用）

### 案件狀況 4 值
- **尚未偵破** = 派出所主管，要破
- **已破獲未移送** = 偵查隊卡，催行政
- **已移送** = 地檢端，派出所不管細列
- **簽結** = 結案，只算發生不算破獲

### 是否破獲 計算欄
`=IFERROR(IF(OR([@案件狀況]="已破獲未移送",[@案件狀況]="已移送"),"是","否"),"")`

### 計入發生數
`=IFERROR(IF([@績效類別]="","是","否"),"是")`
→ 線上查獲（毒品/酒駕/其他）不計發生

### 案類分類 精準比對
舊（v2）：`COUNTIFS(Tbl案件[案類],"*竊盜*",...)`
新（v2.1+）：`COUNTIFS(Tbl案件[案類分類],"竊盜",...)`

### 毒調業務（v2.2 接通）
通緝中／已聲強採／在監 三狀態都算到驗 → Tbl毒調[是否到驗] 計算欄含這 4 種

### 違規項目 6 項（v2.1+）
車輛不停讓行人 / 行人違規 / 車輛行駛人行道 / 移動式測速取締 / 大型車動態違規 / 慢車易肇事違規

### 預設組織
- 警察局：臺北市政府警察局
- 分局：萬華分局
- 派出所：康定路派出所
- 英文：KANGDING POLICE STATION

---

## ⏭ 接下來該做什麼（剩 3% 工程 + 15% 部署）

### 工程建檔層（剩 3%）

1. **頁 8 加「最後到驗日」精確欄**（plan 提的 B 方案）
   - 目前 D5 規則用 `COUNTIF("*未到驗*")` 文字比對
   - 升級加 1 欄淡黃輸入「最後到驗日」+ 計算欄「距今天數」+ CF 紅黃綠
   - 影響：頁 15 D5 從 0/1-5/>5 件數警示 → 改為「>30 日未到驗」精確警示

2. **xlsxwriter post-process 升級 Sparkline**
   - 目前用 Unicode block 模擬（▁▂▃▄▅▆▇█）
   - 升級成 Excel 2010+ 原生 sparkline（openpyxl 不支援）
   - 用 xlsxwriter 二次處理 PoliceStation.xlsx → 寫真 sparkline → 保留其他內容

### 部署準備層（剩 15%）

3. **跨平台測試**：Win + Mac + LibreOffice 開檔測試
4. **列印實測**：A4 直/橫 印頁 2.5 / 3 / 4 / 11 看實際效果
5. **三所試用 1 月**：小所（離島）/ 中所（市區）/ 大所（都會）
6. **部署文件**：
   - 安裝 SOP（密碼 KD2026 修改流程）
   - 月例行作業手冊（抄寫 + 設定表維護）
   - 督導稽核 SOP（頁 15 系統檢核 → 修復）

---

## 🔍 接手第一步建議

1. **讀完這份 + 03_plan_當前進度.md**
2. 跑全測試確認 1319/1319 仍綠
3. 開 v2.1.xlsx 人眼掃 17 頁，重點看：
   - 頁 14 row 6-7 即時 KPI 抄寫區（紅黃綠燈正常切換）
   - 頁 15 健康度大字 + 45 條規則 + HYPERLINK 跳轉
   - 頁 2 KPI 4 毒調率 / KPI 5 未到驗（有實際數字非「—」）
4. 詢問使用者：要做工程剩 3%（最後到驗日 / 真 sparkline）還是部署文件？

## 對使用者開場建議

> 您好，前一 session 已完成 v2.2 升級：Phase 5 保護+CF+列印、Phase 6 系統檢核 45 條、Phase 7 毒調 7 條接通 + 頁 2 KPI 4/5 接通、Phase 8 頁 14 即時 KPI 抄寫區。1319/1319 公式測試全綠。檔案 131.9 KB。
>
> 工程建檔層 97% 完成，剩 3%：
> - 頁 8 加最後到驗日精確 30 日逾期欄
> - xlsxwriter 升級真 sparkline
>
> 部署準備層 0%（您階段），需要：跨平台測試 + 列印實機測試 + 三所試用。

---

© 林錦瑞 ｜ KKEVIN-LIN-2026-V2.2 ｜ 交接包 v2.2 產生於 2026-05-26
