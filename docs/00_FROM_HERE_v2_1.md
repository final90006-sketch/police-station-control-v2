# 派出所綜合管制系統 v2.1 — 交接包（給下個 Claude）

> **這份取代舊的 00_FROM_HERE.md。** 前面 session 已把 v2.0 升級到 v2.1，加了 v2.1 8 大改 + Phase 4 兩頁 + 大量視覺修整。

製作日期：2026-05-26 ｜ 上一 Claude session 結束時 context ~80%

---

## 🎯 一句話交接

派出所 v2 已完成 **Phase 1-4**（17 頁中 14 頁有實質內容 + 案件資料庫等 4 大資料表 + 雙環形圖 + sparkline）。**v2.1 修改規格 M1-M8 全執行完畢**。剩 **Phase 5（條件格式/工作表保護/列印區）+ Phase 6（頁 15 系統檢核 45 條）**。

---

## ✅ v2.1 已完成項目（前一 session 全部交付）

| 範疇 | 項目 | 結果 |
|---|---|---|
| **M1-M2 公式 Bug** | T 案類分類 / U 是否納入統計 修復 | ✓ |
| **M3 案類字典** | 18 → **61 案類**（含 26 主分類）| ✓ |
| **M4 命名範圍** | 案類代碼/名稱/分類/納入全般/績效類別 5 個就緒 | ✓ |
| **M5 績效類別欄** | 案件資料庫 W 欄 + 設定表 H 區（毒品線上/酒駕線上/其他線上） | ✓ |
| **M6 計入發生數** | X 欄自動公式 = IF([@績效類別]="","是","否") | ✓ |
| **M7 自填案類** | Y 欄手填 | ✓ |
| **M8 5 統計頁公式升級** | 頁 2/2.5/10/11/6 都用 [案類分類] 精準比對 + [計入發生數] 條件 | ✓ |
| **Phase 4 頁 3** | 長官 5 分鐘報告（Pyramid Principle）| ✓ |
| **Phase 4 頁 4** | 對上級機關上呈（CHOOSE 切 3 對象）| ✓ |
| **頁 6 重設計** | **移除視圖下拉，改兩區段並呈**（尚未偵破 + 已破獲未移送） | ✓ |
| **頁 7 強化** | +承辦人(G)/線索來源(H)/重大性(I) 3 欄，閾值 14/60 → **7/30 日** | ✓ |
| **違規項目** | 12 → **6 項**（康定所實務：車輛不停讓行人/行人違規/車輛行駛人行道/移動式測速取締/大型車動態違規/慢車易肇事違規）| ✓ |
| **預設值** | **臺北市政府警察局 / 萬華分局 / 康定路派出所** | ✓ |
| **日期顯示** | datetime 從 `hh:mm` 改 `h"時"`（不顯示分鐘） | ✓ |
| **日期 DV** | 發生時間/破獲時間/建立日期/取締日期 都加 Date 資料驗證 (2020-2035) | ✓ |
| **欄寬全面加大** | 案件資料庫以外 8 頁加寬 2-6 字 | ✓ |
| **凍結窗格** | 全 17 頁 freeze_panes = None（使用者要求）| ✓ |
| **下拉清單** | 新增 線索來源清單（6 選項）+ 重大性清單（高/中/低）| ✓ |
| **結案 → 簽結** | 全域改名（案件狀況/偵辦進度/發展中進度 全用「簽結」）| ✓ |
| **未破 → 尚未偵破** | 全域改名 | ✓ |
| **已破未移送 → 已破獲未移送** | 全域改名 | ✓ |
| **頁 2 紅黃綠燈動態上色** | 加 54 條條件格式（9 KPI × 2 cells × 3 規則）燈號 + 大字 cell 依「紅/黃/綠」自動染底色 | ✓ |
| **頁 1 footer 公式 bug 修** | 原本 `'   ="..."` 開頭空白導致 Excel 視為文字，移除前綴 → 正常顯示日期 | ✓ |
| **頁 1 5 步驟區美化** | 3 列等高 48pt（深藍編號 / 淡藍標題 / 灰底說明），全有框線 | ✓ |
| **頁 7 banner 修 overflow** | unmerge 後忘了重新 ws.merge_cells("A1:O1")，導致 banner 塞在 col A — 已加 merge | ✓ |

## 🧪 高強度測試現況

**全 10 頁 公式測試 931/931 全綠燈**
- 頁 2 管制總覽: 100/100
- 頁 2.5 全般刑案分析: 77/77
- 頁 6 刑案管制（v2.1 兩區段版）: 103/103
- 頁 8 毒品調驗: 94/94
- 頁 9 交通績效: 46/46
- 頁 10 績效統計: 173/173
- 頁 11 員警個人卡: 128/128
- 頁 12 跨期間趨勢: 119/119
- 頁 3 5 分鐘報告: 51/51
- 頁 4 對上呈: 40/40

---

## 📁 工程目錄

```
派出所v2_build/
├── config/
│   ├── phase1_skeleton.json        17 sheets + freeze=null
│   ├── phase1_5_settings.json      設定表 8 大分區（61 案類 / 6 違規 / 績效類別等 10 dropdowns）
│   └── phase2_data_tables.json     Tbl案件 25 欄 / Tbl交通 6 欄 / Tbl歷史 14 欄
├── scripts/
│   ├── build.py                    CLI: --phase {1|1.5|2|3|3.2|3.25|3.6|3.8|3.9|3.10|3.11|3.12|4.3|4.4|all}
│   ├── styles.py                   COLOR/FONT/FILL/BORDER/ALIGN/NUM_FMT
│   ├── phases/
│   │   ├── _base.py                + unmerge_all_in_sheet() helper ★ 重要
│   │   ├── phase1_skeleton.py
│   │   ├── phase1_5_settings.py
│   │   ├── phase2_data_tables.py   含 _coerce_value() 日期字串→datetime
│   │   ├── phase3_pages_1_7.py     頁 1 首頁 + 頁 7 發展中（v2.1 9 欄版）
│   │   ├── phase3_page2_overview.py
│   │   ├── phase3_page2_5_crime_analysis.py
│   │   ├── phase3_page6_case_control.py  ★ v2.1 兩區段並呈版
│   │   ├── phase3_page8_drug.py
│   │   ├── phase3_page9_traffic.py
│   │   ├── phase3_page10_performance.py
│   │   ├── phase3_page11_officer.py
│   │   ├── phase3_page12_trend.py
│   │   ├── phase4_page3_brief.py   ★ Pyramid Principle 5 分鐘簡報
│   │   └── phase4_page4_upreport.py ★ CHOOSE 切 3 對象上呈
│   └── tests/
│       ├── framework.py            6 大類測試框架
│       └── test_phase3_page*.py / test_phase4_page*.py
├── output/
│   ├── PoliceStation_v2.0.xlsx    最新 build 輸出（v2.1 內容）
│   └── PoliceStation_v2.1.xlsx    spec 要求的命名版本（cp 自 v2.0）
└── log/                            build_*.log 每次 build 自動記錄
```

## 🚀 開發/驗證指令

```powershell
$py = 'C:\Users\User\AppData\Local\Programs\Python\Python312\python.exe'
$build = 'C:\Users\User\Desktop\02【工作】案件管制與自動化系統\派出所v2_v43_開發檔案\派出所v2_build'

# 重 build 全部
& $py "$build\scripts\build.py" --phase all

# 重 build 單頁
& $py "$build\scripts\build.py" --phase 3.6      # 頁 6
& $py "$build\scripts\build.py" --phase 4.3      # 頁 3

# 跑單頁測試
& $py "$build\scripts\tests\test_phase3_page6.py"

# 跑全部測試
foreach ($t in (ls "$build\scripts\tests\test_phase3_*.py","$build\scripts\tests\test_phase4_*.py")) {
    & $py $t.FullName 2>&1 | Select-String "通過" | Select-Object -Last 1
}

# 複製 v2.0 → v2.1 (因 spec 命名)
Copy-Item "$build\output\PoliceStation_v2.0.xlsx" "$build\output\PoliceStation_v2.1.xlsx" -Force
```

---

## ⚠ 已知踩雷點（前面 session 累積，必記）

### 1. ★ `unmerge_all_in_sheet(ws)` 是必要前置動作
Phase 1 預設給每張 sheet 一個 `A1:G1`/`A2:G2` 的 banner merge。任何 Phase 3/4 頁面若 banner 範圍 ≠ A:G 就會與 Phase 1 重疊，Excel 開檔報「移除合併儲存格」+ 欄寬亂掉。

**解法**：每個頁面 builder 開頭呼叫 `unmerge_all_in_sheet(ws)`，然後 **記得 ws.merge_cells(...) 重新建 banner merge**（曾在頁 7 漏掉，導致 banner 文字塞在 col A 1 格內）。

### 2. ★ openpyxl 沒有 sparkline API
頁 2 + 頁 12 用 **Unicode 區塊字元**模擬（`▁▂▃▄▅▆▇█`），CHOOSE+CONCAT 公式，Consolas 等寬字。

### 3. ★ 結構引用 [@欄名] 對欄名要完全一致
Tbl案件 header 不可加任何前綴（★ 標示等），否則 `[@案件狀況]` 找不到欄位 → Excel 把公式修成 `[[#標題],[XX]]OR(#REF!=...)`。

### 4. ★ 日期 sample 必須是 datetime 物件
不能直接寫 `"2026-05-03 14:20"` 字串。phase2_data_tables.py 有 `_coerce_value()` 處理。

### 5. ★ 17 頁全 freeze_panes = None
使用者明確要求。設定 None 後 sheet view 容易產生 orphan `<selection pane="bottomLeft">` 警示——一律從零重 build 而非 patch。

### 6. 凍結窗格 plan 內提的 A8/A14 等 都不適用
全部 None。

### 7. 用戶 PowerShell 環境
必加：`$env:PYTHONIOENCODING = 'utf-8'` 跟 `[Console]::OutputEncoding = [System.Text.Encoding]::UTF8`。

### 8. ★ openpyxl 公式 cell 開頭不可有空白
寫 `'   ="..."` 會被 Excel 當作文字。公式必須以 `=` 起頭，空白要放進字串內 `'="   ..."`。
（頁 1 footer 因為這個 bug 顯示原始公式文字）

### 9. ★ 條件格式套法 (頁 2 已上)
```python
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill, Font

ws.conditional_formatting.add("E8",
    FormulaRule(formula=['ISNUMBER(SEARCH("紅",E8))'],
                fill=PatternFill("solid", fgColor="FFFEE2E2"),
                font=Font(name="微軟正黑體", bold=True, color="FFB91C1C")))
```
頁 6/7/9 等的燈號 cell 還沒加 CF（Phase 5 該做）。

---

## 📌 業務規則（沿用 v2 plan + v2.1 升級）

### 案件狀況 4 值（已改新名）
- **尚未偵破** = 派出所主管，要破
- **已破獲未移送** = 偵查隊卡，催行政
- **已移送** = 地檢端，派出所不管細列
- **簽結** = 結案，只算發生不算破獲

### 是否破獲 計算欄
`=IFERROR(IF(OR([@案件狀況]="已破獲未移送",[@案件狀況]="已移送"),"是","否"),"")`
→ **簽結 ≠ 破獲**（v2.1 業務規則）

### 計入發生數（v2.1 新增）
`=IFERROR(IF([@績效類別]="","是","否"),"是")`
→ 一般受理 = 是；線上查獲（毒品/酒駕/其他）= 否
→ 統計頁分母都用 `[計入發生數]="是"` 過濾

### 案類分類精準比對（v2.1 從 wildcard 升級）
舊：`COUNTIFS(Tbl案件[案類],"*竊盜*",...)`
新：`COUNTIFS(Tbl案件[案類分類],"竊盜",...)` ← 用 D 區計算欄反查

### 毒調業務（沿用 v2）
通緝中／已聲強採／在監 三狀態都算到驗 → 公式 J 是否到驗

### 違規項目 6 項（v2.1 改）
車輛不停讓行人 / 行人違規 / 車輛行駛人行道 / 移動式測速取締（康定所目標 0）/ 大型車動態違規 / 慢車易肇事違規

### 預設組織
- 警察局：臺北市政府警察局
- 分局：萬華分局
- 派出所：康定路派出所
- 英文：KANGDING POLICE STATION

---

## ⏭ 接下來該做什麼

### Phase 5 — 條件格式 + 工作表保護 + 列印區（建議優先）

1. **條件格式**：
   - ~~頁 2 KPI 卡~~ ✅ 已完成（前一 session 上 54 條 CF）
   - 頁 6 兩區段案件狀況欄 / 頁 7 狀態欄 / 頁 9 達成率 — 還沒上 CF
   - 案件資料庫 容量警示（≥ 800 列變紅）
2. **工作表保護** (密碼 KD2026)：
   - 案件資料庫 鎖住計算欄 R/S/T/U/V/X，輸入欄解鎖
   - 頁 13 設定表 鎖住標題列 + 公式區
   - 其他統計頁全部鎖（只能看不能改）
3. **列印區設定**：
   - 頁 3 5 分鐘報告 → A4 直式 1 頁
   - 頁 4 對上呈 → A4 直式 1-3 頁（依下拉對象）
   - 頁 11 員警卡 → A4 直式 1 頁
   - 頁 2.5 全般刑案分析 → A4 直式 1 頁

### Phase 6 — 頁 15 系統檢核 45 條
（已設計完成，待實作）見舊 03_plan_當前進度.md line 591+

### 未來可能還要做（用戶可能再要求）
- 頁 14 歷史資料的「本月即時 KPI 抄寫區」實作公式
- 頁 8 毒調 加「最後到驗日」欄做精確 30 日逾期計算
- xlsxwriter post-process 把 Unicode block sparkline 升級成真正 Excel sparkline

---

## 🔍 接手第一步建議

1. **讀完這份 + 03_plan_當前進度.md**（plan 那份是完整設計細節）
2. 開檔人眼看 `PoliceStation_v2.1.xlsx`，特別檢視：
   - 頁 6 兩區段並呈是否正確
   - 頁 7 9 欄 + 下拉是否好用
   - 頁 9 違規項目是否變成 6 項
   - 各頁 banner 是否完整顯示（不要再有 col A 塞死）
3. 詢問使用者：要 Phase 5 還是 Phase 6 優先？視覺還要調哪裡？

## 對使用者開場建議

> 您好，前面 session 已完成 v2.1 升級（M1-M8 全做完 + Phase 4 兩頁上線）。我接手後讀完所有交接 + plan，跑完 931/931 測試確認系統穩定。
>
> 接下來可以選：
> - **Phase 5**：上條件格式紅黃綠燈、工作表保護、列印區
> - **Phase 6**：頁 15 系統檢核 45 條
> - 或您有任何頁面想調整？

---

© 林錦瑞 ｜ KKEVIN-LIN-2026-V2.1 ｜ 交接包 v2.1 產生於 2026-05-26
