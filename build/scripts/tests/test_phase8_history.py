"""
test_phase8_history.py — Phase 8 頁 14 即時 KPI 抄寫區 6 大類測試
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook

from phases._base import OUTPUT_PATH
from tests.framework import TestSuite, Category


def run():
    suite = TestSuite("頁 14 即時 KPI 抄寫區")
    wb = load_workbook(OUTPUT_PATH)
    ws = wb["歷史資料"]

    # ============ 1. 公式語法 ============
    cat = Category.SYNTAX

    # 1.1 Row 6 section title + 抄寫狀態
    a6 = str(ws["A6"].value or "")
    suite.assert_true("A6 為動態公式", cat, a6.startswith("="),
                      detail=a6[:60])
    suite.assert_true("A6 含「已抄寫/未抄寫」判斷", cat,
                      "已抄寫" in a6 and "未抄寫" in a6,
                      detail=a6[:80])

    # 1.2 Row 7 14 欄即時 KPI 公式
    for col in "ABCDEFGHIJKLMN":
        v = str(ws[f"{col}7"].value or "")
        suite.assert_true(f"{col}7 為公式或字面值", cat,
                          v.startswith("=") or v.startswith('='),
                          detail=f"{col}7={v[:40]}")

    # 1.3 關鍵 KPI 引用正確 Tbl
    b7 = str(ws["B7"].value or "")
    suite.assert_true("B7 全般發生 引用 Tbl案件", cat,
                      "Tbl案件" in b7 and "COUNTIFS" in b7,
                      detail=b7[:60])
    j7 = str(ws["J7"].value or "")
    suite.assert_true("J7 毒調率 引用 Tbl毒調", cat,
                      "Tbl毒調" in j7,
                      detail=j7[:60])
    l7 = str(ws["L7"].value or "")
    suite.assert_true("L7 交通達標 引用 交通績效管制", cat,
                      "交通績效管制" in l7,
                      detail=l7[:60])

    # ============ 2. 邊界資料 ============
    cat = Category.BOUNDARY

    # 2.1 統計年月 A7 用民國年 + 月格式
    a7 = str(ws["A7"].value or "")
    suite.assert_true("A7 統計年月用 YY/MM 格式", cat,
                      "YEAR(今日)" in a7 and "MONTH(今日)" in a7
                      and "1911" in a7,
                      detail=a7[:80])

    # 2.2 發展中 為 placeholder（承辦自填）
    i7 = str(ws["I7"].value or "")
    suite.assert_true("I7 發展中 為占位符（承辦自填）", cat,
                      i7 == '="—"', detail=i7)
    # 2.2b v2.2 M7 員警冠軍 自動算（INDEX+MATCH+MAX）
    m7 = str(ws["M7"].value or "")
    suite.assert_true("M7 員警冠軍 自動算（INDEX+MATCH+MAX）", cat,
                      "INDEX" in m7 and "MATCH" in m7 and "MAX" in m7,
                      detail=m7[:80])
    suite.assert_true("M7 IFERROR fallback 「—」", cat,
                      "IFERROR" in m7 and '"—"' in m7)

    # 2.3 破獲率 E7 = D7/B7（依賴前面欄位）
    e7 = str(ws["E7"].value or "")
    suite.assert_true("E7 破獲率 = D7/B7", cat,
                      "D7/B7" in e7 and "IFERROR" in e7,
                      detail=e7)

    # ============ 3. 性能基準 ============
    cat = Category.PERFORMANCE

    # 3.1 公式新增量 ≤ 20（row 6 + row 7 = 1+14 = 15 個 cell）
    new_formulas = 0
    for col in "ABCDEFGHIJKLMN":
        if str(ws[f"{col}7"].value or "").startswith("="):
            new_formulas += 1
    suite.assert_le("Row 7 新公式 ≤ 14 個", cat, new_formulas, 14)

    # 3.2 EOMONTH / DATE / YEAR / MONTH 都是 Excel 2007+ 內建
    has_eomonth = any("EOMONTH" in str(ws[f"{c}7"].value or "") for c in "BCDFGH")
    suite.assert_true("使用 EOMONTH 取月底（Excel 2007+ 全相容）", cat,
                      has_eomonth)

    # ============ 4. 跨平台相容 ============
    cat = Category.CROSS_PLATFORM

    # 4.1 無 FILTER/SORT 依賴
    no_filter = True
    for col in "ABCDEFGHIJKLMN":
        v = str(ws[f"{col}7"].value or "")
        if "FILTER(" in v or "SORT(" in v:
            no_filter = False
    suite.assert_true("無 FILTER/SORT 依賴", cat, no_filter)

    # 4.2 引用命名範圍「今日」（TODAY 集中策略）
    has_jin_ri = "今日" in a7
    suite.assert_true("A7 統計年月 引用「今日」（TODAY 集中）", cat,
                      has_jin_ri,
                      detail="使用命名範圍而非揮發 TODAY()")

    # ============ 5. 真實情境 ============
    cat = Category.REAL_WORLD

    # 5.1 Row 8 操作流程說明
    a8 = str(ws["A8"].value or "")
    suite.assert_true("A8 含操作流程", cat,
                      "複製" in a8 and "貼" in a8,
                      detail=a8[:60])

    # 5.2 Row 7 淡黃底（提示「這列是要複製的」）
    b7_fill = ws["B7"].fill
    suite.assert_true("Row 7 用淡黃底（提示複製來源）", cat,
                      b7_fill.fgColor.rgb == "FFFEF3C7"
                      or "FEF3C7" in str(b7_fill.fgColor.rgb or ""),
                      detail=f"fgColor={b7_fill.fgColor.rgb}")

    # ============ 6. 錯誤恢復 ============
    cat = Category.ERROR_RECOVERY

    # 6.1 所有公式包 IFERROR
    iferror_count = 0
    for col in "BCDEFGHJKL":   # 數字欄
        v = str(ws[f"{col}7"].value or "")
        if "IFERROR" in v:
            iferror_count += 1
    suite.assert_true(f"Row 7 數字欄公式 IFERROR 容錯 ≥ 8 條", cat,
                      iferror_count >= 8,
                      detail=f"actual={iferror_count}")

    # 6.2 條件格式 ≥ 2 條（紅綠燈）
    cf_count = sum(len(rules) for rules in ws.conditional_formatting._cf_rules.values())
    suite.assert_true("CF ≥ 2 條（已抄寫綠 / 未抄寫紅）", cat,
                      cf_count >= 2,
                      detail=f"actual={cf_count}")

    # 6.3 CF 公式無結構化引用（避免 Excel 修復警示）
    cf_no_table_ref = True
    for rng, rules in ws.conditional_formatting._cf_rules.items():
        for rule in rules:
            for f in (rule.formula or []):
                if "Tbl" in f and "[" in f:
                    cf_no_table_ref = False
                    break
    suite.assert_true("CF 公式無結構化引用", cat, cf_no_table_ref,
                      detail="避免 sheet xml 修復警示")

    return suite


if __name__ == "__main__":
    suite = run()
    print(suite.report())
    sys.exit(0 if suite.is_all_pass() else 1)
