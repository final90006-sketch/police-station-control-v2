"""
test_phase3_page11.py — 頁 11 員警個人卡 6 大類高強度測試
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook

from phases._base import OUTPUT_PATH
from tests.framework import TestSuite, Category


SHEET = "員警個人績效卡"
OFFICER_CELL = "B3"


def run():
    suite = TestSuite("頁 11 員警個人卡")
    wb = load_workbook(OUTPUT_PATH)
    ws = wb[SHEET]

    # =========================================================
    # 1. 公式語法
    # =========================================================
    cat = Category.SYNTAX

    # 1.1 Banner / sub-banner
    a1 = str(ws["A1"].value or "")
    suite.assert_true("Banner 含「員警個人績效卡」", cat, "員警個人績效卡" in a1)
    a2 = str(ws["A2"].value or "")
    for nr in ["警察局名稱", "分局名稱", "派出所名稱", "今日"]:
        suite.assert_true(f"sub-banner 引用 '{nr}'", cat, nr in a2)

    # 1.2 員警下拉位置 B3 + DataValidation
    suite.assert_true("B3 預設值存在", cat, ws[OFFICER_CELL].value is not None)
    dvs = list(ws.data_validations.dataValidation)
    has_officer_dv = False
    for dv in dvs:
        if "員警姓名清單" in str(dv.formula1 or ""):
            for r in dv.sqref.ranges:
                if "B3" in str(r):
                    has_officer_dv = True
    suite.assert_true("B3 有員警姓名清單 DataValidation", cat, has_officer_dv)

    # 1.3 基本資料條：職稱 / 在職（從設定表抓）
    c3 = str(ws["C3"].value or "")
    suite.assert_true("C3 職稱用 INDEX/MATCH 抓設定表", cat,
                      "INDEX" in c3 and "MATCH" in c3 and "員警姓名清單" in c3,
                      detail=c3[:80])
    e3 = str(ws["E3"].value or "")
    suite.assert_true("E3 在職用 INDEX/MATCH", cat,
                      "INDEX" in e3 and "MATCH" in e3, detail=e3[:80])

    # 1.4 4 KPI 卡 (rows 5-7) values
    kpi_value_cells = [("A6", "總破獲"), ("C6", "全所排名"),
                       ("E6", "全所佔比"), ("G6", "本月新增")]
    for cell, name in kpi_value_cells:
        f = str(ws[cell].value or "")
        suite.assert_true(f"KPI '{name}' ({cell}) 公式存在", cat,
                          f.startswith("="), detail=f[:60])

    # 1.5 4 KPI 卡公式特性
    # 總破獲 用 COUNTIFS + 承辦人 B3
    f_a6 = str(ws["A6"].value or "")
    suite.assert_true("總破獲 用 COUNTIFS Tbl案件 + 承辦人 B3", cat,
                      "COUNTIFS" in f_a6 and "承辦人" in f_a6 and OFFICER_CELL in f_a6,
                      detail=f_a6[:80])

    # 全所排名 用 SUMPRODUCT 比較
    f_c6 = str(ws["C6"].value or "")
    suite.assert_true("全所排名 用 SUMPRODUCT 比較", cat,
                      "SUMPRODUCT" in f_c6, detail=f_c6[:80])
    # 引用頁 10 helper area
    suite.assert_true("全所排名 引用頁 10 helper（績效統計）", cat,
                      "績效統計" in f_c6, detail=f_c6[:80])

    # 全所佔比 = 個人 / 全所
    f_e6 = str(ws["E6"].value or "")
    suite.assert_true("全所佔比 = 個人 / 全所（IFERROR）", cat,
                      "IFERROR" in f_e6 and "/" in f_e6, detail=f_e6[:80])

    # 本月新增 用 DATE 區間
    f_g6 = str(ws["G6"].value or "")
    suite.assert_true("本月新增 用 DATE 區間 + 破獲時間", cat,
                      "DATE(YEAR(今日)" in f_g6 and "破獲時間" in f_g6,
                      detail=f_g6[:80])

    # 1.6 比較表 5 項
    compare_rows = [(10, "總破獲"), (11, "竊盜"), (12, "詐欺"),
                    (13, "毒品"), (14, "暴力犯罪")]
    for r, name in compare_rows:
        suite.assert_eq(f"比較項 {name} 在 A{r}", cat,
                        ws[f"A{r}"].value, name)
        # B 個人 / C 全所平均 / D 差距 / E 評級 公式存在
        for col in ["B", "C", "D", "E"]:
            f = str(ws[f"{col}{r}"].value or "")
            suite.assert_true(f"比較表 {name} {col}{r} 公式", cat,
                              f.startswith("="), detail=f[:60])

    # 1.7 比較表 C 全所平均 = 全所破獲 / 在職員警數
    f_c10 = str(ws["C10"].value or "")
    suite.assert_true("全所平均 用 COUNTIF 在職員警", cat,
                      "COUNTIF" in f_c10 and "設定表" in f_c10,
                      detail=f_c10[:80])

    # 1.8 評級欄含 4 級
    f_e10 = str(ws["E10"].value or "")
    for level in ["優異", "達標", "待加強", "落後"]:
        suite.assert_true(f"評級含「{level}」", cat, level in f_e10,
                          detail=f_e10[:80])

    # 1.9 辦案清單 header (row 17) 8 欄
    expected_list = ["序", "案類", "發生日", "發生地", "破獲日",
                     "案件狀況", "涉案金額", "備註"]
    for i, h in enumerate(expected_list, start=1):
        col = chr(ord("A") + i - 1)
        suite.assert_eq(f"辦案清單 header {col}17", cat, ws[f"{col}17"].value, h)

    # 1.10 辦案清單 30 列用 AGGREGATE + INDEX
    for k in [1, 15, 30]:
        row = 17 + k
        f_b = str(ws[f"B{row}"].value or "")
        suite.assert_true(f"辦案清單列 {k} (B{row}) 用 AGGREGATE+INDEX", cat,
                          "AGGREGATE" in f_b and "INDEX" in f_b,
                          detail=f_b[:80])
        suite.assert_true(f"辦案清單列 {k} (B{row}) 含承辦人={OFFICER_CELL}", cat,
                          OFFICER_CELL in f_b, detail=f_b[:80])
        suite.assert_true(f"辦案清單列 {k} (B{row}) 含是否破獲", cat,
                          "是否破獲" in f_b)

    # 1.11 簽章區（row 49）
    a49 = str(ws["A49"].value or "")
    e49 = str(ws["E49"].value or "")
    suite.assert_true("簽章區所長", cat, "所長" in a49 and "簽章" in a49)
    suite.assert_true("簽章區督導官", cat, "督導官" in e49 and "簽章" in e49)

    # =========================================================
    # 2. 邊界資料
    # =========================================================
    cat = Category.BOUNDARY

    # 2.1 未選員警時 KPI 公式 IFERROR → 0
    f = str(ws["A6"].value or "")
    suite.assert_true("總破獲 IFERROR 預設 0", cat,
                      "IFERROR" in f and ",0)" in f, detail=f[:60])

    # 2.2 全所平均分母 0（無在職員警）IFERROR 容錯
    f = str(ws["C10"].value or "")
    suite.assert_true("全所平均 IFERROR 容錯", cat, "IFERROR" in f)

    # 2.3 辦案清單超過 30 列：第 30 列以後不顯示（無 row 48 公式）
    suite.assert_true("辦案清單第 31 列無公式（限 30 列）", cat,
                      ws["B48"].value is None or not str(ws["B48"].value or "").startswith("="),
                      detail=f"B48={ws['B48'].value!r}")

    # 2.4 員警下拉 prompt + error
    has_prompt = False
    has_error = False
    for dv in list(ws.data_validations.dataValidation):
        if "員警姓名清單" in str(dv.formula1 or ""):
            has_prompt = bool(dv.prompt)
            has_error = bool(dv.error)
    suite.assert_true("員警下拉有 prompt 提示", cat, has_prompt)
    suite.assert_true("員警下拉有 error 警告", cat, has_error)

    # =========================================================
    # 3. 性能基準
    # =========================================================
    cat = Category.PERFORMANCE

    fcount = 0
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if v and isinstance(v, str) and v.startswith("="):
                fcount += 1
    suite.assert_le("頁 11 公式總數 ≤ 350", cat, fcount, 350,
                    detail=f"actual={fcount}")

    # 揮發性
    volatile_counts = {"TODAY()": 0, "NOW()": 0, "OFFSET(": 0, "INDIRECT(": 0}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                for k in volatile_counts:
                    volatile_counts[k] += cell.value.count(k)
    suite.assert_eq("不直接呼叫 TODAY()", cat, volatile_counts["TODAY()"], 0)
    suite.assert_le("OFFSET ≤ 3", cat, volatile_counts["OFFSET("], 3)
    suite.assert_le("INDIRECT ≤ 3", cat, volatile_counts["INDIRECT("], 3)

    suite.assert_le("max_row ≤ 60", cat, ws.max_row, 60,
                    detail=f"max_row={ws.max_row}")

    # =========================================================
    # 4. 跨平台相容
    # =========================================================
    cat = Category.CROSS_PLATFORM

    # 4.1 無 FILTER/SORT — 用 AGGREGATE Fallback
    has_filter = False
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                if "FILTER(" in cell.value or "SORT(" in cell.value:
                    has_filter = True
                    break
    suite.assert_true("頁 11 無 FILTER/SORT 依賴（用 AGGREGATE）", cat, not has_filter)

    # 4.2 AGGREGATE 使用（Excel 2010+ 標準）
    has_aggregate = False
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and "AGGREGATE(" in cell.value:
                has_aggregate = True
                break
    suite.assert_true("辦案清單用 AGGREGATE Fallback", cat, has_aggregate)

    # 4.3 命名範圍存在
    wb_nrs = set(wb.defined_names)
    for nr in ["今日", "員警姓名清單"]:
        suite.assert_true(f"命名範圍 '{nr}' 存在", cat, nr in wb_nrs)

    # 4.4 結構引用 Tbl案件[...]
    f = str(ws["A6"].value or "")
    suite.assert_true("使用 Tbl案件 結構引用", cat, "Tbl案件[" in f)

    # =========================================================
    # 5. 真實情境
    # =========================================================
    cat = Category.REAL_WORLD

    # 5.1 凍結 A8（rows 1-7 含員警下拉 + KPI 卡 留在頂部）
    suite.assert_eq("凍結窗格 = None（使用者要求全頁無凍結）", cat,
                    ws.freeze_panes, None)

    # 5.2 4 KPI 卡 labels 含對應 emoji + 名稱
    expected = [("A5", "總破獲"), ("C5", "全所排名"),
                ("E5", "全所佔比"), ("G5", "本月新增")]
    for cell, name in expected:
        v = str(ws[cell].value or "")
        suite.assert_true(f"{cell} 含 KPI 名稱「{name}」", cat, name in v,
                          detail=v[:40])

    # 5.3 比較表評級含 ★ 標星
    f = str(ws["E10"].value or "")
    suite.assert_true("評級含 ★ 標星", cat, "★" in f)

    # 5.4 頁尾含 KKEVIN-LIN + 「A4 直式」
    f = str(ws["A50"].value or "")
    suite.assert_true("頁尾含 KKEVIN-LIN", cat, "KKEVIN-LIN" in f)
    suite.assert_true("頁尾說明 A4 直式列印", cat, "A4 直式" in f)

    # 5.5 員警下拉預設值合理（示範員警之一）
    default = ws[OFFICER_CELL].value
    suite.assert_true("員警下拉預設值合理（示範員警）", cat,
                      default in ["王小明", "李大華", "陳美玲"],
                      detail=f"default={default}")

    # =========================================================
    # 6. 錯誤恢復
    # =========================================================
    cat = Category.ERROR_RECOVERY

    # 6.1 KPI 4 卡全包 IFERROR
    for cell in ["A6", "C6", "E6", "G6"]:
        f = str(ws[cell].value or "")
        suite.assert_true(f"{cell} 包 IFERROR", cat, "IFERROR" in f,
                          detail=f[:60])

    # 6.2 比較表 B/C/D/E 全包 IFERROR
    for r in range(10, 15):
        for col in ["B", "C", "D", "E"]:
            f = str(ws[f"{col}{r}"].value or "")
            suite.assert_true(f"{col}{r} IFERROR", cat, "IFERROR" in f,
                              detail=f[:60])

    # 6.3 辦案清單 30 列全包 IFERROR
    for k in [1, 10, 20, 30]:
        row = 17 + k
        for col in ["B", "C", "F"]:
            f = str(ws[f"{col}{row}"].value or "")
            suite.assert_true(f"辦案 {col}{row} IFERROR", cat,
                              "IFERROR" in f, detail=f[:60])

    # 6.4 職稱/在職資料缺失時 IFERROR 顯示「—」
    f = str(ws["C3"].value or "")
    suite.assert_true("職稱 IFERROR 顯示「—」", cat,
                      'IFERROR' in f and '"—"' in f, detail=f[:80])

    return suite


if __name__ == "__main__":
    suite = run()
    print(suite.report())
    sys.exit(0 if suite.is_all_pass() else 1)
