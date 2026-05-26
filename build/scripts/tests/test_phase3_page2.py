"""
test_phase3_page2.py — 頁 2 管制總覽 6 大類高強度測試

依測試框架 6 大類：
  1. 公式語法（IFERROR / 命名範圍 / 結構引用 / 燈號）
  2. 邊界資料（空表 / 全綠 / 全紅 / placeholder）
  3. 性能基準（公式總數 / 揮發性函數）
  4. 跨平台相容（無 FILTER/SORT 依賴 / Excel 2010+ 函數）
  5. 真實情境（30 秒掃完 / 紅燈 insight）
  6. 錯誤恢復（IFERROR 包覆 / 分母 0 / 空 Tbl歷史）
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook

from phases._base import OUTPUT_PATH
from tests.framework import TestSuite, Category


SHEET = "管制總覽"


def run():
    suite = TestSuite("頁 2 管制總覽")
    wb = load_workbook(OUTPUT_PATH)
    ws = wb[SHEET]

    # =========================================================
    # 1. 公式語法
    # =========================================================
    cat = Category.SYNTAX

    # 1.1 Banner sub-formula 引用三層組織命名範圍
    a2 = str(ws["A2"].value or "")
    for nr in ["警察局名稱", "分局名稱", "派出所名稱", "今日"]:
        suite.assert_true(f"Banner 副標引用 '{nr}'", cat, nr in a2,
                          detail=a2[:80])

    # 1.2 Insight headline 公式存在且引用 S50:S58 燈號區
    a4 = str(ws["A4"].value or "")
    suite.assert_true("Insight 公式存在", cat, a4.startswith("="),
                      detail=a4[:80])
    suite.assert_true("Insight 引用 S50:S58 燈號區", cat, "S50:S58" in a4,
                      detail=a4[:80])
    for word in ["紅", "黃", "綠"]:
        suite.assert_true(f"Insight 含 '{word}'", cat, word in a4)

    # 1.3 9 KPI 現值 + 燈號 公式存在
    kpi_cells = [
        # (label_cell, value_cell, light_cell, name)
        ("A8",  "A9",  "E8",  "全般破獲率"),
        ("G8",  "G9",  "K8",  "竊盜破獲率"),
        ("M8",  "M9",  "Q8",  "詐欺破獲率"),
        ("A13", "A14", "E13", "毒調率"),
        ("G13", "G14", "K13", "未到驗人口"),
        ("M13", "M14", "Q13", "未破案件"),
        ("A18", "A19", "E18", "交通達成率"),
        ("G18", "G19", "K18", "酒駕取締"),
        ("M18", "M19", "Q18", "闖紅燈取締"),
    ]
    for (lab, val, lig, name) in kpi_cells:
        label_v = str(ws[lab].value or "")
        val_v = str(ws[val].value or "")
        lig_v = str(ws[lig].value or "")
        suite.assert_true(f"KPI '{name}' label 存在", cat, name in label_v,
                          detail=label_v[:40])
        suite.assert_true(f"KPI '{name}' 現值公式存在", cat,
                          val_v.startswith("="),
                          detail=val_v[:60])
        suite.assert_true(f"KPI '{name}' 燈號公式存在", cat,
                          lig_v.startswith("="),
                          detail=lig_v[:60])

    # 1.4 所有非 placeholder 現值公式必須包 IFERROR（除了引用其他 sheet 的）
    non_placeholder_value_cells = ["A9", "G9", "M9", "M14", "A19", "G19", "M19"]
    for ref in non_placeholder_value_cells:
        f = str(ws[ref].value or "")
        suite.assert_true(f"{ref} 包 IFERROR", cat, "IFERROR" in f,
                          detail=f[:60])

    # 1.5 KPI 4/5（毒調率/未到驗）已被頁 8 build 接通真實 Tbl毒調 公式
    # （update_page2_drug_kpis Part C 設計，跑 3.8 後生效）
    for ref in ["A14", "G14"]:
        f = str(ws[ref].value or "")
        suite.assert_true(f"{ref} 接通 Tbl毒調（頁 8 update_page2_drug_kpis）", cat,
                          "Tbl毒調" in f, detail=f[:60])

    # 1.6 KPI 現值公式引用 Tbl案件 / Tbl交通 / 跨頁
    suite.assert_true("KPI 1 引用 Tbl案件", cat,
                      "Tbl案件" in str(ws["A9"].value or ""),
                      detail="全般破獲率")
    suite.assert_true("KPI 6 引用 Tbl案件", cat,
                      "Tbl案件" in str(ws["M14"].value or ""),
                      detail="未破案件")
    suite.assert_true("KPI 7 引用交通績效管制", cat,
                      "交通績效管制" in str(ws["A19"].value or ""),
                      detail="交通達成率")
    suite.assert_true("KPI 8 引用 Tbl交通", cat,
                      "Tbl交通" in str(ws["G19"].value or ""),
                      detail="酒駕")

    # 1.7 helper sparkline CONCAT 公式存在
    for i in range(9):
        cell = f"AR{50+i}"
        f = str(ws[cell].value or "")
        # 4/5/8/9 是 placeholder 或無歷史
        spark_col_none_idx = [7, 8]  # KPI 8、9 spark_col 是 None（i=7, 8）
        if i in spark_col_none_idx:
            suite.assert_true(f"{cell} (KPI {i+1}) 為待擴充字串", cat,
                              "待擴充" in f, detail=f[:40])
        else:
            suite.assert_true(f"{cell} (KPI {i+1}) CONCAT 公式存在", cat,
                              "CONCAT" in f, detail=f[:60])

    # =========================================================
    # 2. 邊界資料
    # =========================================================
    cat = Category.BOUNDARY

    # 2.1 警示摘要三卡公式存在
    for ref, name in [("A25", "紅燈總數"), ("G25", "黃燈總數"), ("M25", "綠燈總數")]:
        f = str(ws[ref].value or "")
        suite.assert_true(f"{name} 公式存在", cat, "COUNTIF" in f and "S50:S58" in f,
                          detail=f[:60])

    # 2.2 Sparkline block 字元公式處理 MIN=MAX 情況
    af50 = str(ws["AF50"].value or "")
    suite.assert_true("block 公式處理 MIN=MAX（全相同）", cat,
                      'MAX($T50:$AE50)=MIN($T50:$AE50)' in af50,
                      detail="避免除零")

    # 2.3 sparkline 資料 12 個月（T..AE）
    for i in range(12):
        cell = f"{chr(ord('T')+i) if i < 7 else 'A'+chr(ord('A')+i-6)}50"
        # Actually 用 get_column_letter
    from openpyxl.utils import get_column_letter
    spark_data_cols_present = 0
    for col_offset in range(12):
        col = get_column_letter(20 + col_offset)
        v = ws[f"{col}50"].value
        if v is not None and str(v) != "":
            spark_data_cols_present += 1
    suite.assert_eq("KPI 1 sparkline 12 個資料 cell 都有公式", cat,
                    spark_data_cols_present, 12)

    # 2.4 placeholder KPI（4、5）的 spark_cell 仍能取歷史資料（毒調率、未到驗有 Tbl歷史 欄）
    suite.assert_true("KPI 4 毒調率 sparkline 有資料公式", cat,
                      "INDEX(Tbl歷史" in str(ws["T53"].value or ""),
                      detail=str(ws["T53"].value or "")[:60])

    # =========================================================
    # 3. 性能基準
    # =========================================================
    cat = Category.PERFORMANCE

    # 3.1 公式總數合理（< 350，含 9 KPI × ~30 公式 + 警示三卡 + headline + helper）
    fcount = 0
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if v and isinstance(v, str) and v.startswith("="):
                fcount += 1
    suite.assert_le("頁 2 公式總數 ≤ 350", cat, fcount, 350,
                    detail=f"actual={fcount}")

    # 3.2 揮發性函數（TODAY / NOW / OFFSET / INDIRECT）使用克制
    volatile = ["TODAY()", "NOW()", "OFFSET(", "INDIRECT("]
    counts = {v: 0 for v in volatile}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                for v in volatile:
                    counts[v] += cell.value.count(v)
    for v, n in counts.items():
        suite.assert_le(f"揮發性函數 '{v}' ≤ 5", cat, n, 5,
                        detail=f"actual={n}")
    # 3.3 TODAY() 不應直接出現（要走「今日」命名範圍）
    suite.assert_eq("頁 2 不直接呼叫 TODAY()", cat, counts["TODAY()"], 0,
                    detail="應全部引用「今日」命名範圍")

    # 3.4 max_row 合理（含隱藏 helper 區域）
    suite.assert_le("頁 2 max_row ≤ 60（含 helper）", cat, ws.max_row, 60,
                    detail=f"max_row={ws.max_row}")

    # =========================================================
    # 4. 跨平台相容
    # =========================================================
    cat = Category.CROSS_PLATFORM

    # 4.1 無 FILTER/SORT 依賴（Excel 2019- 也能跑）
    has_filter_sort = False
    for row in ws.iter_rows():
        for cell in row:
            v = str(cell.value or "")
            if "FILTER(" in v or "SORT(" in v:
                has_filter_sort = True
                break
    suite.assert_true("頁 2 無 FILTER/SORT 依賴", cat, not has_filter_sort,
                      detail="Excel 2019- 兼容")

    # 4.2 CONCAT 可用（Excel 2016+ 標準）
    # CONCAT 在 Excel 2019 完整、365 完整、2016 部分 — 標註但不阻擋
    # 若有疑慮可改 CONCATENATE，但 CONCATENATE 不支援範圍引用
    has_concat = any("CONCAT(" in str(c.value or "")
                     for row in ws.iter_rows() for c in row)
    suite.assert_true("使用 CONCAT 函數（Excel 2016+）", cat, has_concat,
                      detail="LibreOffice 也支援")

    # 4.3 命名範圍存在
    wb_nrs = set(wb.defined_names)
    for nr in ["今日", "警察局名稱", "分局名稱", "派出所名稱"]:
        suite.assert_true(f"命名範圍 '{nr}' 存在", cat, nr in wb_nrs)

    # 4.4 Tbl 結構引用（不是 fixed range）
    sample_formula = str(ws["A9"].value or "")
    suite.assert_true("使用結構引用 Tbl案件[...]", cat,
                      "Tbl案件[" in sample_formula,
                      detail=sample_formula[:60])

    # =========================================================
    # 5. 真實情境
    # =========================================================
    cat = Category.REAL_WORLD

    # 5.1 Banner 含「管制總覽」+「九宮格」
    a1 = str(ws["A1"].value or "")
    suite.assert_true("Banner 含「管制總覽」", cat, "管制總覽" in a1)
    suite.assert_true("Banner 含「九宮格」", cat, "九宮格" in a1)

    # 5.2 9 KPI 全到位（業務組合：刑案 3 + 毒調 2 + 案件 1 + 交通 3）
    expected_kpis = ["全般破獲率", "竊盜破獲率", "詐欺破獲率",
                     "毒調率", "未到驗人口", "未破案件",
                     "交通達成率", "酒駕取締", "闖紅燈取締"]
    label_cells = ["A8","G8","M8","A13","G13","M13","A18","G18","M18"]
    for (cell, expected) in zip(label_cells, expected_kpis):
        v = str(ws[cell].value or "")
        suite.assert_true(f"KPI 名稱 '{expected}' 在 {cell}", cat,
                          expected in v, detail=v[:30])

    # 5.3 警示摘要三色齊全
    for cell, color in [("A24", "紅"), ("G24", "黃"), ("M24", "綠")]:
        v = str(ws[cell].value or "")
        suite.assert_true(f"警示摘要 {cell} 含 '{color}'", cat, color in v,
                          detail=v[:30])

    # 5.4 凍結窗格 A8（標題以上凍結）
    suite.assert_eq("凍結窗格 = None（使用者要求全頁無凍結）", cat,
                    ws.freeze_panes, None)

    # 5.5 頁尾含 KKEVIN-LIN
    a28 = str(ws["A28"].value or "")
    suite.assert_true("頁尾含 KKEVIN-LIN", cat, "KKEVIN-LIN" in a28,
                      detail=a28[:60])

    # =========================================================
    # 6. 錯誤恢復
    # =========================================================
    cat = Category.ERROR_RECOVERY

    # 6.1 所有 sparkline data 公式包 IFERROR（避免歷史不足時 #REF）
    # 取 KPI 1 第 1 月公式
    f = str(ws["T50"].value or "")
    suite.assert_true("Sparkline 資料公式包 IFERROR", cat, "IFERROR" in f,
                      detail=f[:60])

    # 6.2 sparkline CONCAT 包 IFERROR
    f = str(ws["AR50"].value or "")
    suite.assert_true("Sparkline CONCAT 包 IFERROR", cat, "IFERROR" in f,
                      detail=f[:60])

    # 6.3 block 字元公式處理空值（ISNUMBER 檢查）
    f = str(ws["AF50"].value or "")
    suite.assert_true("block 公式檢查 ISNUMBER 處理空值", cat,
                      "ISNUMBER" in f, detail=f[:60])

    # 6.4 KPI 1-3 破獲率分母為 0 時不爆（IFERROR 包覆）
    f = str(ws["A9"].value or "")
    suite.assert_true("破獲率 IFERROR(...,0)", cat,
                      "IFERROR" in f and ",0)" in f.replace(" ", ""),
                      detail=f[:60])

    # 6.5 KPI 4/5 燈號已接通 Tbl毒調 + 毒調率目標（非 placeholder）
    e13 = str(ws["E13"].value or "")
    suite.assert_true("E13 毒調率燈號 引用毒調率目標", cat,
                      "毒調率目標" in e13, detail=e13[:60])
    k13 = str(ws["K13"].value or "")
    suite.assert_true("K13 未到驗燈號 引用候選旗標", cat,
                      "候選旗標" in k13 or "Tbl毒調" in k13,
                      detail=k13[:60])

    return suite


if __name__ == "__main__":
    suite = run()
    print(suite.report())
    sys.exit(0 if suite.is_all_pass() else 1)
