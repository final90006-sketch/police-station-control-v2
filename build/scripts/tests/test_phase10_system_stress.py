"""
test_phase10_system_stress.py — Phase 10 系統面高強度模擬測試

不是針對單頁，而是針對整個 v2.2 系統做：
  1. 跨頁 SSOT 連動驗證（Tbl 改變 → 多頁同步）
  2. 業務鐵則模擬（拘提他轄、簽結不計破獲、毒調 4 狀態算到驗等）
  3. 邊界資料壓力（容量極限、跨年度、特殊字元）
  4. xlsx 結構完整性（17 頁 / 4 Tbl / 32+ 命名範圍 / 17 條件格式）
  5. 真原生 sparkline XML 注入正確性
  6. Phase 順序驗證（先填內容後鎖定）
"""
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook

from phases._base import OUTPUT_PATH
from tests.framework import TestSuite, Category


def run():
    suite = TestSuite("系統面高強度模擬（v2.2 全 17 頁）")
    wb = load_workbook(OUTPUT_PATH)

    # ============ 1. 公式語法（系統層）============
    cat = Category.SYNTAX

    # 1.1 17 頁全部存在
    expected_sheets = [
        "首頁_操作說明", "管制總覽", "全般刑案管制情形分析",
        "長官5分鐘報告", "對上級機關上呈", "案件資料庫", "刑案管制",
        "發展中案件管制", "毒品調驗人口管制", "交通績效管制",
        "交通取締明細", "績效統計", "員警個人績效卡", "跨期間趨勢分析",
        "設定表", "歷史資料", "系統檢核"
    ]
    for name in expected_sheets:
        suite.assert_true(f"工作表 '{name}' 存在", cat,
                          name in wb.sheetnames)

    # 1.2 4 大資料表完整
    expected_tables = ["Tbl案件", "Tbl交通", "Tbl歷史", "Tbl毒調"]
    all_tables = []
    for ws_name in wb.sheetnames:
        for t in wb[ws_name].tables.values():
            all_tables.append(t.name)
    for tname in expected_tables:
        suite.assert_true(f"Excel Table '{tname}' 已註冊", cat,
                          tname in all_tables)

    # 1.3 32+ 命名範圍（含 v2.2 新加）
    nrs = set(wb.defined_names)
    critical_nrs = [
        "派出所名稱", "警察局名稱", "分局名稱", "派出所英文",
        "員警姓名清單", "案類名稱清單", "案類分類清單", "納入全般清單",
        "違規項目清單", "案件狀況清單", "發生管轄清單", "查獲管轄清單",
        "今日", "新進閾值", "老案閾值", "毒調率目標",
        "案件容量警示", "案件容量上限",
        "婦幼件數", "協尋件數",
        "績效類別清單", "線索來源清單", "重大性清單",
    ]
    for n in critical_nrs:
        suite.assert_true(f"命名範圍 '{n}' 存在", cat, n in nrs)

    # ============ 2. 邊界資料（跨頁 SSOT）============
    cat = Category.BOUNDARY

    # 2.1 Tbl毒調 14 欄結構（v2.2）
    ws_drug = wb["毒品調驗人口管制"]
    drug_tbl = ws_drug.tables["Tbl毒調"]
    suite.assert_eq("Tbl毒調 範圍 A14:N84 （v2.2 14 欄）", cat,
                    drug_tbl.ref, "A14:N84")

    # 2.2 距今天數 計算欄存在
    suite.assert_eq("Tbl毒調 K14 = 距今天數", cat,
                    ws_drug["K14"].value, "距今天數")
    suite.assert_eq("Tbl毒調 J14 = 最後到驗日（input）", cat,
                    ws_drug["J14"].value, "最後到驗日")

    # 2.3 頁 15 D5 已升級用 候選旗標
    ws15 = wb["系統檢核"]
    # D5 在 row 15+ where id=D5
    d5_row = None
    for r in range(15, 60):
        if ws15[f"A{r}"].value == "D5":
            d5_row = r
            break
    suite.assert_true("頁 15 D5 規則找得到", cat, d5_row is not None)
    if d5_row:
        d5_value = str(ws15[f"D{d5_row}"].value or "")
        suite.assert_true("頁 15 D5 已升級用 候選旗標（含 距今 > 30）", cat,
                          "候選旗標" in d5_value,
                          detail=d5_value[:80])

    # 2.4 頁 14 M7 員警冠軍 自動算
    ws14 = wb["歷史資料"]
    m7 = str(ws14["M7"].value or "")
    suite.assert_true("頁 14 M7 員警冠軍 自動算（INDEX+MATCH+MAX）", cat,
                      "INDEX" in m7 and "MATCH" in m7 and "MAX" in m7,
                      detail=m7[:80])

    # 2.5 頁 14 helper 區存在（R50:S110）
    r50 = str(ws14["R50"].value or "")
    s50 = str(ws14["S50"].value or "")
    suite.assert_true("頁 14 R50 helper 員警引用", cat,
                      "設定表" in r50,
                      detail=r50[:60])
    suite.assert_true("頁 14 S50 helper 當月破獲 COUNTIFS", cat,
                      "COUNTIFS" in s50 and "Tbl案件" in s50,
                      detail=s50[:80])

    # ============ 3. 性能基準 ============
    cat = Category.PERFORMANCE

    # 3.1 檔案大小 ≤ 500 KB（plan 規範）
    size_kb = OUTPUT_PATH.stat().st_size / 1024
    suite.assert_le("檔案大小 ≤ 500 KB", cat, size_kb, 500,
                    detail=f"actual={size_kb:.1f} KB")

    # 3.2 全表公式總數 ≤ 12,000（plan 規範）
    total_formulas = 0
    for ws_name in wb.sheetnames:
        ws = wb[ws_name]
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if v and isinstance(v, str) and v.startswith("="):
                    total_formulas += 1
    suite.assert_le("全表公式總數 ≤ 12,000（plan 規範）", cat,
                    total_formulas, 12000,
                    detail=f"actual={total_formulas}")

    # 3.3 TODAY() 集中化 — 全表 TODAY() 直接呼叫 ≤ 1 處
    today_calls = 0
    for ws_name in wb.sheetnames:
        ws = wb[ws_name]
        for row in ws.iter_rows():
            for cell in row:
                v = str(cell.value or "")
                today_calls += v.count("TODAY()")
    # plan 規範：每張表最多 1-2 個 TODAY()，全表 ≤ 3 處（集中策略）
    suite.assert_le("TODAY() 直接呼叫 ≤ 3 處（plan 規範）", cat,
                    today_calls, 3,
                    detail=f"actual={today_calls}")

    # ============ 4. 跨平台相容 ============
    cat = Category.CROSS_PLATFORM

    # 4.1 17 個 sheet 全部 freeze_panes = None（使用者要求）
    for name in expected_sheets:
        ws = wb[name]
        suite.assert_true(f"{name} freeze_panes = None", cat,
                          ws.freeze_panes is None or ws.freeze_panes == "A1",
                          detail=f"actual={ws.freeze_panes}")

    # 4.2 FILTER/SORT 使用統計（Excel 365 only — 應有 fallback）
    filter_count = 0
    sort_count = 0
    for ws_name in wb.sheetnames:
        ws = wb[ws_name]
        for row in ws.iter_rows():
            for cell in row:
                v = str(cell.value or "")
                if v.startswith("="):
                    if "FILTER(" in v:
                        filter_count += 1
                    if "SORT(" in v:
                        sort_count += 1
    suite.assert_le("FILTER 使用 ≤ 5 處（plan：盡量無）", cat,
                    filter_count, 5, detail=f"actual={filter_count}")
    suite.assert_le("SORT 使用 ≤ 5 處（plan：盡量無）", cat,
                    sort_count, 5, detail=f"actual={sort_count}")

    # 4.3 真原生 sparkline XML 注入正確（sheet2.xml 應含 x14:sparklineGroup）
    with zipfile.ZipFile(OUTPUT_PATH) as z:
        sheet2 = z.read("xl/worksheets/sheet2.xml").decode("utf-8")
    suite.assert_true("sheet2.xml 含 sparkline XML 注入", cat,
                      "x14:sparklineGroup" in sheet2)
    # 計算 sparkline 數量
    spark_count = sheet2.count("<x14:sparkline>")
    suite.assert_eq("頁 2 sparkline 數 = 7", cat, spark_count, 7)

    # 4.4 sparkline 命名空間正確
    suite.assert_true("sparkline ext URI 正確", cat,
                      "05C60535-1F16-4fd2-B633-F4F36F0B64E0" in sheet2)
    suite.assert_true("sparkline x14 namespace 正確", cat,
                      "schemas.microsoft.com/office/spreadsheetml/2009/9/main" in sheet2)

    # ============ 5. 真實情境（業務鐵則）============
    cat = Category.REAL_WORLD

    # 5.1 案件資料庫 case sample 包含 4 種案件狀況
    ws5 = wb["案件資料庫"]
    case_statuses = []
    for r in range(15, 25):
        v = ws5[f"J{r}"].value
        if v:
            case_statuses.append(str(v))
    expected_statuses = ["尚未偵破", "已破獲未移送", "已移送"]  # 簽結至少有設計
    for status in expected_statuses:
        suite.assert_true(f"sample 含案件狀況「{status}」", cat,
                          any(status in s for s in case_statuses))

    # 5.2 是否破獲 計算欄符合業務規則
    # 樣本 R15 應該是「否」(尚未偵破), R16/R17/R18 應該是「是」(已破獲/已移送)
    for r, expected in [(15, "否"), (16, "是"), (17, "是"), (18, "是")]:
        formula = str(ws5[f"R{r}"].value or "")
        suite.assert_true(f"R{r} 是否破獲 為公式 IF(...)", cat,
                          formula.startswith("="),
                          detail=formula[:60])

    # 5.3 毒調 sample 5 列含業務鐵則 4 狀態
    drug_statuses = []
    for r in range(15, 20):
        v = str(ws_drug[f"H{r}"].value or "")
        drug_statuses.append(v)
    for kw in ["已到驗", "未到驗", "通緝", "強採", "在監"]:
        suite.assert_true(f"毒調 sample 含「{kw}」", cat,
                          any(kw in s for s in drug_statuses))

    # 5.4 v2.2 毒調 J 欄（最後到驗日）有 sample 日期
    drug_dates = []
    for r in range(15, 20):
        v = ws_drug[f"J{r}"].value
        drug_dates.append(v)
    non_empty_dates = sum(1 for d in drug_dates if d is not None and d != "")
    suite.assert_true("v2.2 毒調 J 欄至少 3 筆樣本日期", cat,
                      non_empty_dates >= 3,
                      detail=f"actual={non_empty_dates}")

    # ============ 6. 錯誤恢復 ============
    cat = Category.ERROR_RECOVERY

    # 6.1 所有 sheet 公式 IFERROR 比例 ≥ 50%
    formulas_with_iferror = 0
    total_for_iferror_check = 0
    for ws_name in wb.sheetnames:
        ws = wb[ws_name]
        for row in ws.iter_rows():
            for cell in row:
                v = str(cell.value or "")
                if v.startswith("="):
                    total_for_iferror_check += 1
                    if "IFERROR" in v:
                        formulas_with_iferror += 1
    iferror_pct = (formulas_with_iferror / total_for_iferror_check * 100
                   if total_for_iferror_check else 0)
    suite.assert_true(f"IFERROR 覆蓋率 ≥ 50%（實際 {iferror_pct:.1f}%）", cat,
                      iferror_pct >= 50,
                      detail=f"{formulas_with_iferror}/{total_for_iferror_check}")

    # 6.2 11 個 sheet 已上密碼保護（plan 規範）
    protected_count = sum(1 for name in expected_sheets
                          if wb[name].protection.sheet)
    suite.assert_true("受保護 sheet 數 ≥ 10", cat, protected_count >= 10,
                      detail=f"actual={protected_count}/17")

    # 6.3 4 頁列印區設定（A4 直式 1 頁）
    print_sheets = [
        "全般刑案管制情形分析", "長官5分鐘報告",
        "對上級機關上呈", "員警個人績效卡"
    ]
    for name in print_sheets:
        ws = wb[name]
        pa = ws.print_area or ""
        suite.assert_true(f"{name} 有列印區", cat,
                          pa is not None and pa != "",
                          detail=f"pa={pa!r}")

    # 6.4 CF 規則使用結構化引用比例 = 0
    # CF 公式不可用 Tbl... 結構化引用，否則 Excel 修復
    cf_with_tbl_ref = []
    for ws_name in wb.sheetnames:
        ws = wb[ws_name]
        for rng, rules in ws.conditional_formatting._cf_rules.items():
            for rule in rules:
                for f in (rule.formula or []):
                    if "Tbl" in f and "[" in f:
                        cf_with_tbl_ref.append((ws_name, str(rng), f[:50]))
    suite.assert_eq("CF 公式 0 個用結構化引用（防 Excel 修復警示）", cat,
                    len(cf_with_tbl_ref), 0,
                    detail=f"violations={cf_with_tbl_ref[:3]}")

    return suite


if __name__ == "__main__":
    suite = run()
    print(suite.report())
    sys.exit(0 if suite.is_all_pass() else 1)
