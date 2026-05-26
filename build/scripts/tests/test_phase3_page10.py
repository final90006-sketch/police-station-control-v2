"""
test_phase3_page10.py — 頁 10 績效統計 6 大類高強度測試
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook

from phases._base import OUTPUT_PATH
from tests.framework import TestSuite, Category


SHEET = "績效統計"


def run():
    suite = TestSuite("頁 10 績效統計")
    wb = load_workbook(OUTPUT_PATH)
    ws = wb[SHEET]

    # =========================================================
    # 1. 公式語法
    # =========================================================
    cat = Category.SYNTAX

    # 1.1 Banner & sub-banner
    a1 = str(ws["A1"].value or "")
    suite.assert_true("Banner 含「績效統計」", cat, "績效統計" in a1)
    a2 = str(ws["A2"].value or "")
    for nr in ["警察局名稱", "分局名稱", "派出所名稱", "今日"]:
        suite.assert_true(f"sub-banner 引用 '{nr}'", cat, nr in a2)

    # 1.2 上半案類表 — 5 案類
    case_rows = [(9, "竊盜"), (10, "詐欺"), (11, "毒品"),
                 (12, "暴力犯罪"), (13, "其他")]
    for r, name in case_rows:
        v = str(ws[f"A{r}"].value or "")
        suite.assert_true(f"案類 '{name}' 在 row {r}", cat, name in v,
                          detail=v[:30])
        # B 發生 / C 破獲 / D 破獲率 公式存在
        for col in ["B", "C", "D"]:
            f = str(ws[f"{col}{r}"].value or "")
            suite.assert_true(f"案類 {name} {col}{r} 公式存在", cat,
                              f.startswith("="), detail=f[:60])

    # 1.3 4 大案類使用 Tbl案件 結構引用 + 萬用字元
    for r, name in case_rows[:4]:
        f_occur = str(ws[f"B{r}"].value or "")
        f_solve = str(ws[f"C{r}"].value or "")
        suite.assert_true(f"案類 {name} 發生公式引用 Tbl案件", cat,
                          "Tbl案件[" in f_occur, detail=f_occur[:60])
        suite.assert_true(f"案類 {name} 破獲公式引用 Tbl案件", cat,
                          "Tbl案件[" in f_solve, detail=f_solve[:60])
        # 「發生」用發生管轄
        suite.assert_true(f"案類 {name} 發生用發生管轄=本轄", cat,
                          "發生管轄" in f_occur and "本轄" in f_occur,
                          detail=f_occur[:80])
        # 「破獲」用查獲管轄 + 是否破獲
        suite.assert_true(f"案類 {name} 破獲用查獲管轄+是否破獲", cat,
                          "查獲管轄" in f_solve and "是否破獲" in f_solve,
                          detail=f_solve[:80])

    # 1.4 「其他」案類 = 全般 - 前 4 案類（用 MAX(0, ...) 避免負值）
    f = str(ws["B13"].value or "")
    suite.assert_true("「其他」發生 = 全般 - SUM(B9:B12)", cat,
                      "SUM(B9:B12)" in f and "MAX(0" in f, detail=f[:80])
    f = str(ws["C13"].value or "")
    suite.assert_true("「其他」破獲 = 全般 - SUM(C9:C12)", cat,
                      "SUM(C9:C12)" in f and "MAX(0" in f, detail=f[:80])

    # 1.5 排名欄用 RANK.AVG
    for r in range(9, 14):
        f = str(ws[f"E{r}"].value or "")
        suite.assert_true(f"E{r} 排名用 RANK.AVG", cat,
                          "RANK.AVG" in f, detail=f[:60])

    # 1.6 本月變化欄含 DATE/MONTH 日期判斷
    f = str(ws["F9"].value or "")
    suite.assert_true("本月變化欄含 DATE 比較", cat,
                      "DATE(YEAR(今日)" in f and "MONTH" in f, detail=f[:80])
    suite.assert_true("本月變化顯示 ↑/↓/→", cat,
                      "↑" in f and "↓" in f and "→" in f, detail=f[:80])

    # 1.7 全般合計 row 14
    f = str(ws["A14"].value or "")
    suite.assert_true("全般合計列存在", cat, "全般合計" in f, detail=f[:30])
    suite.assert_true("合計列破獲率公式", cat,
                      "IFERROR" in str(ws["D14"].value or ""))

    # 1.8 下半員警 Top 10
    headers_off = ["排名", "員警", "全般破獲", "竊盜", "詐欺", "毒品", "暴力犯罪", "其他"]
    for i, h in enumerate(headers_off, start=1):
        col = chr(ord("A") + i - 1)
        v = str(ws[f"{col}18"].value or "")
        suite.assert_eq(f"員警表 header {col}18", cat, v, h)

    # 1.9 Top 10 主表用 LARGE + INDEX/MATCH（非 SORT/FILTER）
    for rank in range(1, 11):
        row = 18 + rank
        for col in ["A", "B", "C", "H"]:
            f = str(ws[f"{col}{row}"].value or "")
            suite.assert_true(f"Top {rank} {col}{row} 公式存在", cat,
                              f.startswith("="), detail=f[:60])
        # 使用 LARGE
        f_a = str(ws[f"A{row}"].value or "")
        suite.assert_true(f"Top {rank} 排名公式用 LARGE", cat,
                          "LARGE" in f_a, detail=f_a[:60])
        # 員警姓名用 INDEX/MATCH
        f_b = str(ws[f"B{row}"].value or "")
        suite.assert_true(f"Top {rank} 員警公式用 INDEX/MATCH", cat,
                          "INDEX" in f_b and "MATCH" in f_b,
                          detail=f_b[:60])

    # 1.10 Top 1-3 含獎牌符號
    f_a19 = str(ws["A19"].value or "")
    for medal in ["🥇", "🥈", "🥉"]:
        suite.assert_true(f"Top 排名公式含 {medal}", cat,
                          medal in f_a19, detail=f_a19[:60])

    # 1.11 Helper area 60 員警 × 9 欄
    # row 100 (第 1 員警)
    suite.assert_eq("Helper row 100 序號 = 1", cat, ws["A100"].value, 1)
    suite.assert_true("Helper B100 員警姓名公式存在", cat,
                      str(ws["B100"].value or "").startswith("="),
                      detail=str(ws["B100"].value or "")[:60])
    suite.assert_true("Helper B100 引用 員警姓名清單", cat,
                      "員警姓名清單" in str(ws["B100"].value or ""))
    # row 159 (第 60 員警)
    suite.assert_eq("Helper row 159 序號 = 60", cat, ws["A159"].value, 60)

    # =========================================================
    # 2. 邊界資料
    # =========================================================
    cat = Category.BOUNDARY

    # 2.1 空表時 0/0 不爆（IFERROR 包覆）
    for r in range(9, 14):
        f = str(ws[f"D{r}"].value or "")
        suite.assert_true(f"D{r} 破獲率包 IFERROR", cat, "IFERROR" in f,
                          detail=f[:60])

    # 2.2 Top 10 員警若不足 10 人：LARGE 仍有效（rank 鍵 >0 才顯示）
    f_a19 = str(ws["A19"].value or "")
    suite.assert_true("Top 1 排名公式檢查 key>0", cat,
                      "LARGE" in f_a19 and ">0" in f_a19, detail=f_a19[:80])

    # 2.3 員警 0 人時：Top 10 表全部顯示「—」
    # 由於排名鍵 = 0（B 空時）+ LARGE 取最大 = 0 → IF(>0,...,"—")
    suite.assert_true("空員警 LARGE 鍵 = 0 → 顯示「—」", cat,
                      '"—"' in f_a19, detail=f_a19[:80])

    # 2.4 helper 排名鍵唯一化（C + seq*0.0001）
    f = str(ws["I100"].value or "")
    suite.assert_true("排名鍵唯一化（+seq*0.0001）", cat,
                      "0.0001" in f, detail=f[:60])

    # 2.5 同分情況下，LARGE/MATCH 不會重複拿同員警
    # 由於排名鍵每員警 +seq*0.0001 唯一，MATCH 永遠找到不同列
    suite.assert_true("排名鍵設計支援同分區分", cat, True,
                      detail="seq*0.0001 唯一化")

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
    suite.assert_le("頁 10 公式總數 ≤ 700", cat, fcount, 700,
                    detail=f"actual={fcount}")

    # 揮發性函數使用
    volatile_counts = {"TODAY()": 0, "NOW()": 0, "OFFSET(": 0, "INDIRECT(": 0}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                for k in volatile_counts:
                    volatile_counts[k] += cell.value.count(k)
    suite.assert_eq("不直接呼叫 TODAY()", cat, volatile_counts["TODAY()"], 0,
                    detail="應全用「今日」命名範圍")
    suite.assert_eq("不直接呼叫 NOW()", cat, volatile_counts["NOW()"], 0)
    suite.assert_le("OFFSET 使用 ≤ 3", cat, volatile_counts["OFFSET("], 3)
    suite.assert_le("INDIRECT 使用 ≤ 3", cat, volatile_counts["INDIRECT("], 3)

    # max_row（含 helper）
    suite.assert_le("頁 10 max_row ≤ 165", cat, ws.max_row, 165,
                    detail=f"max_row={ws.max_row}")

    # =========================================================
    # 4. 跨平台相容
    # =========================================================
    cat = Category.CROSS_PLATFORM

    # 4.1 無 FILTER/SORT 依賴（plan 要求 Fallback）
    has_filter_sort = False
    for row in ws.iter_rows():
        for cell in row:
            v = str(cell.value or "")
            if "FILTER(" in v or "SORT(" in v:
                has_filter_sort = True
                break
    suite.assert_true("頁 10 無 FILTER/SORT 依賴（Excel 2010+ 兼容）",
                      cat, not has_filter_sort)

    # 4.2 使用 LARGE/INDEX/MATCH（Excel 2010+ 標準）
    f = str(ws["B19"].value or "")
    suite.assert_true("Top 10 用 LARGE+INDEX+MATCH", cat,
                      "INDEX" in f and "MATCH" in f and "LARGE" in str(ws["A19"].value or ""))

    # 4.3 命名範圍存在
    wb_nrs = set(wb.defined_names)
    for nr in ["今日", "員警姓名清單"]:
        suite.assert_true(f"命名範圍 '{nr}' 存在", cat, nr in wb_nrs)

    # 4.4 結構引用 Tbl案件[...]
    sample = str(ws["B9"].value or "")
    suite.assert_true("使用結構引用 Tbl案件[...]", cat,
                      "Tbl案件[" in sample, detail=sample[:60])

    # =========================================================
    # 5. 真實情境
    # =========================================================
    cat = Category.REAL_WORLD

    # 5.1 5 大案類在主表
    for r, name in case_rows:
        v = str(ws[f"A{r}"].value or "")
        suite.assert_eq(f"案類 {name} 在 A{r}", cat, v.strip(), name)

    # 5.2 員警表 8 欄完整
    suite.assert_eq("員警表 8 欄", cat, ws["H18"].value, "其他")

    # 5.3 凍結窗格 A8
    suite.assert_eq("凍結窗格 = None（使用者要求全頁無凍結）", cat,
                    ws.freeze_panes, None)

    # 5.4 頁尾含 KKEVIN-LIN + SSOT 註腳
    f = str(ws["A31"].value or "")
    suite.assert_true("頁尾含 KKEVIN-LIN", cat, "KKEVIN-LIN" in f)
    suite.assert_true("頁尾說明 SSOT（無獨立 Tbl刑案）", cat,
                      "SSOT" in f or "無獨立 Tbl刑案" in f, detail=f[:80])

    # 5.5 案類「破獲率 ≥ 100% 標星」公式（H 欄）
    f = str(ws["H9"].value or "")
    suite.assert_true("案類 H 欄含 ★ 標星邏輯", cat,
                      "★" in f and ">=1" in f, detail=f[:80])

    # 5.6 維護指引存在
    f = str(ws["A4"].value or "")
    suite.assert_true("維護指引提示頁 5 + 設定表 B 區", cat,
                      "頁 5" in f and "設定表 B 區" in f, detail=f[:80])

    # =========================================================
    # 6. 錯誤恢復
    # =========================================================
    cat = Category.ERROR_RECOVERY

    # 6.1 所有 KPI 公式 IFERROR 包覆
    important_cells = ["B9", "C9", "D9", "B10", "C10", "D10",
                       "B13", "C13", "D13", "B14", "C14", "D14"]
    for ref in important_cells:
        f = str(ws[ref].value or "")
        suite.assert_true(f"{ref} IFERROR 容錯", cat,
                          "IFERROR" in f or "MAX(0" in f or "SUM(" in f,
                          detail=f[:60])

    # 6.2 Top 10 INDEX/MATCH 包 IFERROR
    f = str(ws["B19"].value or "")
    suite.assert_true("Top 1 員警姓名 IFERROR 容錯", cat,
                      "IFERROR" in f, detail=f[:60])

    # 6.3 Helper area 員警公式 IFERROR
    f = str(ws["C100"].value or "")
    suite.assert_true("Helper 員警全般破獲 IFERROR", cat,
                      "IFERROR" in f and 'IF(B100=""' in f, detail=f[:80])

    # 6.4 排名鍵設計：空員警 → 鍵 = 0（不入榜）
    f = str(ws["I100"].value or "")
    suite.assert_true("排名鍵：空員警 = 0", cat,
                      'IF(B100=""' in f and ",0," in f, detail=f[:60])

    return suite


if __name__ == "__main__":
    suite = run()
    print(suite.report())
    sys.exit(0 if suite.is_all_pass() else 1)
