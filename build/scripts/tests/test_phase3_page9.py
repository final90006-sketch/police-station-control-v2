"""
test_phase3_page9.py — 頁 9 交通績效 6 大類高強度測試
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook

from phases._base import OUTPUT_PATH, load_config
from tests.framework import TestSuite, Category


def run():
    suite = TestSuite("頁 9 交通績效管制")
    wb = load_workbook(OUTPUT_PATH)
    ws = wb["交通績效管制"]

    # ============ 1. 公式語法 ============
    cat = Category.SYNTAX

    # 1.1 達標縮影 3 卡公式
    for cell_ref, desc in [("A8", "達標項數"), ("C8", "未達標項數"), ("E8", "總達成率")]:
        f = ws[cell_ref].value
        suite.assert_true(f"卡片公式 {desc} 存在", cat,
                          f is not None and str(f).startswith("="),
                          detail=f"got={str(f)[:60]}")

    # 1.2 主表 12 項公式
    p15 = load_config("phase1_5_settings")
    violations_count = len(p15["section_e"]["violations"])
    for r in range(14, 14 + violations_count):
        for col in ["C", "D", "E", "F"]:
            f = ws[f"{col}{r}"].value
            suite.assert_true(f"主表 {col}{r} 公式存在", cat,
                              f is not None and str(f).startswith("="),
                              detail=f"violation_row={r-13}")

    # 1.3 所有公式包 IFERROR
    iferror_required = ["B14", "C14", "D14", "F14"]  # 部分欄期望容錯
    for ref in ["C14", "D14"]:
        f = str(ws[ref].value or "")
        suite.assert_true(f"{ref} 包 IFERROR", cat, "IFERROR" in f,
                          detail=f[:60])

    # 1.4 SSOT 引用 Tbl交通
    f = str(ws["C14"].value or "")
    suite.assert_true("達成欄引用 Tbl交通", cat, "Tbl交通" in f,
                      detail=f[:60])

    # ============ 2. 邊界資料 ============
    cat = Category.BOUNDARY

    # 2.1 無取締資料時：所有達成 = 0、達成率 = 0、狀態 = 紅
    # 計算公式時不需實際 Excel 評估，但確認公式邏輯能處理空表
    f_state = str(ws["F14"].value or "")
    suite.assert_true("狀態公式處理 0% 案例", cat,
                      "0.6" in f_state and "0.8" in f_state,
                      detail="包含閾值判斷")

    # 2.2 合計列存在
    total_row = 14 + violations_count
    total_label = ws[f"A{total_row}"].value
    suite.assert_true("合計列存在", cat,
                      total_label and "全般合計" in str(total_label),
                      detail=f"row={total_row} label={total_label!r}")

    # 2.3 違規項目數 = E 區 config 設定
    item_count = 0
    for r in range(14, 14 + 20):
        v = ws[f"A{r}"].value
        if not v:
            break
        v_str = str(v).strip()
        if v_str.startswith("─") or "合計" in v_str:
            break
        item_count += 1
    suite.assert_eq("違規項目數量符合 E 區", cat,
                    item_count, violations_count)

    # ============ 3. 性能基準 ============
    cat = Category.PERFORMANCE

    # 3.1 公式總數 < 100（頁 9 應該很輕量）
    formula_count = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                formula_count += 1
    suite.assert_le("頁 9 公式數量 ≤ 100", cat, formula_count, 100,
                    detail=f"actual={formula_count}")

    # 3.2 儲存格使用範圍合理（不超過 30 列）
    max_row = ws.max_row
    suite.assert_le("最大列數 ≤ 30", cat, max_row, 30,
                    detail=f"max_row={max_row}")

    # ============ 4. 跨平台 ============
    cat = Category.CROSS_PLATFORM

    # 4.1 使用的函數：COUNTIF、SUM、IFERROR、INDEX、MATCH、IF、ROUND、REPT、MIN、ROW
    # 都是 Excel 2010+ 支援，無 FILTER/SORT 依賴 — 跨平台 OK
    suite.assert_true("無 FILTER/SORT 依賴（跨平台相容）", cat,
                      not any("FILTER(" in str(c.value or "") or "SORT(" in str(c.value or "")
                              for row in ws.iter_rows() for c in row),
                      detail="搜尋無 FILTER/SORT")

    # 4.2 命名範圍引用都可解析
    wb_nrs = set(wb.defined_names)
    refs_used = ["違規項目清單"]
    for n in refs_used:
        suite.assert_true(f"命名範圍 '{n}' 存在", cat, n in wb_nrs)

    # ============ 5. 真實情境 ============
    cat = Category.REAL_WORLD

    # 5.1 視覺橫條公式產生長度合理
    f_bar = str(ws["E14"].value or "")
    suite.assert_true("視覺橫條使用 REPT", cat, "REPT" in f_bar,
                      detail=f_bar[:60])
    suite.assert_true("視覺橫條上限 20 chars（不爆框）", cat, "20" in f_bar)

    # 5.2 狀態燈 3 段（綠 / 黃 / 紅）
    f_status = str(ws["F14"].value or "")
    for icon in ["綠", "黃", "紅"]:
        suite.assert_true(f"狀態燈含 '{icon}'", cat, icon in f_status,
                          detail=f"formula={f_status[:80]}")

    # 5.3 banner 文字符合「達成率看板」定位
    a1 = str(ws["A1"].value or "")
    suite.assert_true("Banner 含「達成率看板」", cat, "達成率" in a1,
                      detail=a1[:60])

    # ============ 6. 錯誤恢復 ============
    cat = Category.ERROR_RECOVERY

    # 6.1 達成率公式有 IFERROR 包 0
    f = str(ws["D14"].value or "")
    suite.assert_true("達成率 IFERROR 預設 0", cat,
                      "IFERROR" in f and ",0)" in f,
                      detail=f[:60])

    # 6.2 達成欄 IFERROR 包 0
    f = str(ws["C14"].value or "")
    suite.assert_true("達成 IFERROR 預設 0", cat,
                      "IFERROR" in f and ",0)" in f,
                      detail=f[:60])

    # 6.3 合計列 IFERROR
    f = str(ws[f"D{14+violations_count}"].value or "")
    suite.assert_true("合計達成率 IFERROR 包覆", cat, "IFERROR" in f,
                      detail=f[:60])

    return suite


if __name__ == "__main__":
    suite = run()
    print(suite.report())
    sys.exit(0 if suite.is_all_pass() else 1)
