"""
test_phase3_page2_5.py — 頁 2.5 全般刑案分析 6 大類高強度測試
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook

from phases._base import OUTPUT_PATH
from tests.framework import TestSuite, Category


SHEET = "全般刑案管制情形分析"


def run():
    suite = TestSuite("頁 2.5 全般刑案分析")
    wb = load_workbook(OUTPUT_PATH)
    ws = wb[SHEET]

    # =========================================================
    # 1. 公式語法
    # =========================================================
    cat = Category.SYNTAX

    # 1.1 章戳頁首
    a1 = str(ws["A1"].value or "")
    suite.assert_true("Banner 含「全般刑案管制情形分析」", cat,
                      "全般刑案管制情形分析" in a1)

    # 1.2 sub-banner 含英文機關名 + REPORT DATE
    a2 = str(ws["A2"].value or "")
    suite.assert_true("sub-banner 引用 派出所英文", cat,
                      "派出所英文" in a2)
    suite.assert_true("sub-banner 含 REPORT DATE", cat,
                      "REPORT DATE" in a2)

    # 1.3 統計區間
    a3 = str(ws["A3"].value or "")
    suite.assert_true("統計區間公式含 YEAR(今日)-1911", cat,
                      "YEAR(今日)-1911" in a3, detail=a3[:80])
    suite.assert_true("統計區間含同期對照", cat, "同期對照" in a3)

    # 1.4 三大 KPI 卡（rows 6-7）
    kpi_cells = [("A7", "本所統計·破獲率"),
                 ("D7", "發生件數"),
                 ("F7", "破獲件數")]
    for val_cell, name in kpi_cells:
        f = str(ws[val_cell].value or "")
        suite.assert_true(f"KPI '{name}' ({val_cell}) 公式存在", cat,
                          f.startswith("="), detail=f[:60])
        suite.assert_true(f"KPI '{name}' 包 IFERROR", cat,
                          "IFERROR" in f, detail=f[:60])

    # 1.5 本所統計 用 查獲管轄 = 本轄
    a7 = str(ws["A7"].value or "")
    suite.assert_true("本所破獲率 用 查獲管轄=本轄 + 發生管轄=本轄", cat,
                      "查獲管轄" in a7 and "發生管轄" in a7 and "本轄" in a7,
                      detail=a7[:120])

    # 1.6 比較長條（row 8-9）— 本所 vs 分局
    c8 = str(ws["C8"].value or "")
    suite.assert_true("本所長條用 REPT", cat,
                      "REPT" in c8 and "查獲管轄" in c8, detail=c8[:80])
    c9 = str(ws["C9"].value or "")
    suite.assert_true("分局長條用 REPT + 發生管轄+是否破獲", cat,
                      "REPT" in c9 and "發生管轄" in c9 and "是否破獲" in c9,
                      detail=c9[:80])

    # 1.7 差異原因卡 (row 10) — 自動算 拘提他轄 + 已破未送
    a10 = str(ws["A10"].value or "")
    suite.assert_true("差異卡含拘提他轄項", cat,
                      "拘提" in a10 and "他轄" in a10, detail=a10[:80])
    suite.assert_true("差異卡含已破未移送項", cat,
                      "已破未移送" in a10)
    suite.assert_true("差異卡用 COUNTIFS", cat,
                      "COUNTIFS" in a10)

    # 1.8 重點案類（row 13）— 竊盜 + 詐欺
    a13 = str(ws["A13"].value or "")
    suite.assert_true("竊盜卡片含發生+破獲+破獲率", cat,
                      "發生" in a13 and "破獲" in a13 and "破獲率" in a13)
    suite.assert_true("竊盜卡片 [案類分類]=竊盜 (v2.1 精準比對)", cat,
                      '"竊盜"' in a13 and '案類分類' in a13, detail=a13[:80])
    e13 = str(ws["E13"].value or "")
    suite.assert_true("詐欺卡片 [案類分類]=詐欺 (v2.1 精準比對)", cat,
                      '"詐欺"' in e13 and '案類分類' in e13, detail=e13[:80])

    # 1.9 環形圖 helper data (Q100/R100 竊盜, S100/T100 詐欺)
    suite.assert_eq("竊盜 helper Q100 = 已破", cat, ws["Q100"].value, "已破")
    suite.assert_eq("竊盜 helper Q101 = 尚未偵破", cat,
                    ws["Q101"].value, "尚未偵破")
    r100 = str(ws["R100"].value or "")
    suite.assert_true("竊盜 helper R100 用 COUNTIFS [案類分類]=竊盜", cat,
                      '"竊盜"' in r100 and '案類分類' in r100 and "是否破獲" in r100, detail=r100[:80])

    suite.assert_eq("詐欺 helper S100 = 已破", cat, ws["S100"].value, "已破")
    t100 = str(ws["T100"].value or "")
    suite.assert_true("詐欺 helper T100 用 COUNTIFS [案類分類]=詐欺", cat,
                      '"詐欺"' in t100 and '案類分類' in t100 and "是否破獲" in t100)

    # 1.10 婦幼 + 協尋（row 30）
    a30 = str(ws["A30"].value or "")
    suite.assert_true("婦幼狀態引用 婦幼件數 命名範圍", cat,
                      "婦幼件數" in a30, detail=a30[:80])
    suite.assert_true("婦幼 = 0 顯示「目前尚無管制案件」", cat,
                      "目前尚無管制案件" in a30)

    e30 = str(ws["E30"].value or "")
    suite.assert_true("協尋狀態引用 協尋件數 命名範圍", cat,
                      "協尋件數" in e30)

    # 1.11 未破刑案三欄表（rows 25-29）
    # Header
    a25 = str(ws["A25"].value or "")
    suite.assert_true("未破刑案三欄 header 含「分類原因」", cat, "分類原因" in a25)
    d25 = str(ws["D25"].value or "")
    suite.assert_true("D25 header 含「件數」", cat, "件數" in d25)
    e25 = str(ws["E25"].value or "")
    suite.assert_true("E25 header 含「補充說明」", cat, "補充說明" in e25)

    # 總計列 D29
    d29 = str(ws["D29"].value or "")
    suite.assert_true("總計 D29 = SUM(D26:D28)", cat, "SUM(D26:D28)" in d29,
                      detail=d29[:60])

    # 自動算未破比對 (E29)
    e29 = str(ws["E29"].value or "")
    suite.assert_true("E29 顯示自動算未破", cat,
                      "自動算未破" in e29 and "COUNTIFS" in e29,
                      detail=e29[:80])

    # =========================================================
    # 2. 邊界資料
    # =========================================================
    cat = Category.BOUNDARY

    # 2.1 分母 0 處理 — 所有破獲率公式 IFERROR
    for cell in ["A7", "C8", "C9", "A13", "E13"]:
        f = str(ws[cell].value or "")
        if "/" in f:    # 含除法的才檢查 IFERROR
            suite.assert_true(f"{cell} 包 IFERROR（分母 0 容錯）", cat,
                              "IFERROR" in f, detail=f[:80])

    # 2.2 環形圖數據 IFERROR
    for cell in ["R100", "R101", "T100", "T101"]:
        f = str(ws[cell].value or "")
        suite.assert_true(f"環形圖 {cell} IFERROR / MAX(0,...)", cat,
                          "IFERROR" in f or "MAX(0" in f, detail=f[:60])

    # 2.3 未破數可能為負時 MAX(0, ...) 容錯
    f = str(ws["R101"].value or "")
    suite.assert_true("未破數 MAX(0,...) 防負", cat, "MAX(0" in f, detail=f[:60])

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
    suite.assert_le("頁 2.5 公式總數 ≤ 50", cat, fcount, 50,
                    detail=f"actual={fcount}")

    # 揮發性
    has_today = False
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and "TODAY()" in cell.value:
                has_today = True
                break
    suite.assert_true("不直接呼叫 TODAY()", cat, not has_today)

    # Charts 數
    suite.assert_eq("Charts 數 = 2（雙環形圖）", cat, len(ws._charts), 2)

    # =========================================================
    # 4. 跨平台相容
    # =========================================================
    cat = Category.CROSS_PLATFORM

    has_filter = False
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                if "FILTER(" in cell.value or "SORT(" in cell.value:
                    has_filter = True
                    break
    suite.assert_true("頁 2.5 無 FILTER/SORT 依賴", cat, not has_filter)

    # 命名範圍
    wb_nrs = set(wb.defined_names)
    for nr in ["今日", "派出所名稱", "派出所英文", "婦幼件數", "協尋件數"]:
        suite.assert_true(f"命名範圍 '{nr}' 存在", cat, nr in wb_nrs)

    # 結構引用
    sample = str(ws["A7"].value or "")
    suite.assert_true("使用 Tbl案件 結構引用", cat, "Tbl案件[" in sample)

    # =========================================================
    # 5. 真實情境
    # =========================================================
    cat = Category.REAL_WORLD

    # 5.1 凍結 A8
    suite.assert_eq("凍結窗格 = None（使用者要求全頁無凍結）", cat,
                    ws.freeze_panes, None)

    # 5.2 三大 Section
    sections = [(5, "壹  全般刑案總覽"),
                (12, "貳  重點案類"),
                (24, "參  未破案件")]
    for row, name in sections:
        v = str(ws[f"A{row}"].value or "")
        suite.assert_true(f"Section row {row} 含「{name}」", cat,
                          name in v, detail=v[:40])

    # 5.3 頁尾含「報告人」「破獲率基準 100%」「本所/分局統計分流」
    f = str(ws["A32"].value or "")
    suite.assert_true("頁尾含「報告人」", cat, "報告人" in f)
    suite.assert_true("頁尾含「破獲率基準」", cat, "破獲率基準" in f)
    suite.assert_true("頁尾明示分流邏輯", cat,
                      "查獲管轄" in f and "發生管轄" in f, detail=f[:80])

    # 5.4 未破三欄表預留 3 列 + 總計
    suite.assert_true("未破列 26 預留輸入", cat,
                      ws["A26"].value is not None, detail=ws["A26"].value)
    suite.assert_true("未破列 28 預留輸入", cat,
                      ws["A28"].value is not None)

    # 5.5 婦幼 + 協尋 採方案 A 邏輯
    a30 = str(ws["A30"].value or "")
    suite.assert_true("婦幼狀態：件數=0 顯示無管制", cat,
                      "IF(婦幼件數=0" in a30 and "尚無管制" in a30)

    # =========================================================
    # 6. 錯誤恢復
    # =========================================================
    cat = Category.ERROR_RECOVERY

    # 6.1 所有 KPI / 比較長條 / 差異卡 / 重點案類 / 環形圖數 IFERROR
    cells_iferror = ["A7", "D7", "F7", "C8", "C9", "A10",
                     "A13", "E13", "R100", "R101", "T100", "T101"]
    for cell in cells_iferror:
        f = str(ws[cell].value or "")
        suite.assert_true(f"{cell} 容錯（IFERROR 或 MAX/SUM）", cat,
                          "IFERROR" in f or "MAX(" in f or "SUM(" in f,
                          detail=f[:60])

    # 6.2 差異卡 IFERROR
    f = str(ws["A10"].value or "")
    suite.assert_true("差異卡 IFERROR", cat, "IFERROR" in f)

    # 6.3 自動 vs 手動總計不一致時提示
    e29 = str(ws["E29"].value or "")
    suite.assert_true("E29 不一致時建議補分類", cat,
                      "若不一致" in e29 or "補分類" in e29,
                      detail=e29[:80])

    return suite


if __name__ == "__main__":
    suite = run()
    print(suite.report())
    sys.exit(0 if suite.is_all_pass() else 1)
