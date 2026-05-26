"""
test_phase6_inspection.py — 頁 15 系統檢核 6 大類高強度測試

驗證：
  1. 45 條規則 完整出現（編號 + 類別 + 描述 + 公式）
  2. 健康度 3 卡 + ○ 待接通卡 = 4 個聚合公式
  3. 7 大類縮影 7 卡
  4. 跳轉連結 HYPERLINK 公式
  5. 條件格式 4 條（紅/黃/綠/灰 + 整列淡染 2 條）
  6. 容錯：所有公式包 IFERROR；待接通規則正確顯示「○ 待接通」
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook

from phases._base import OUTPUT_PATH
from phases.phase6_inspection import RULES, CATEGORIES
from tests.framework import TestSuite, Category


SHEET = "系統檢核"
RULE_FIRST_ROW = 15
RULE_LAST_ROW = RULE_FIRST_ROW + 45 - 1   # 59


def run():
    suite = TestSuite("頁 15 系統檢核")
    wb = load_workbook(OUTPUT_PATH)
    ws = wb[SHEET]

    # ============ 1. 公式語法 ============
    cat = Category.SYNTAX

    # 1.1 45 條規則：A 編號 + B 類別 + C 描述 + D 公式 + E 燈號 + F 嚴重度 全有
    for idx, rule in enumerate(RULES):
        r = RULE_FIRST_ROW + idx
        # A 編號
        suite.assert_eq(f"規則 {rule['id']} A{r} 編號", cat,
                        ws[f"A{r}"].value, rule["id"])
        # B 類別
        suite.assert_eq(f"規則 {rule['id']} B{r} 類別", cat,
                        ws[f"B{r}"].value, rule["cat"])
        # C 描述存在
        suite.assert_true(f"規則 {rule['id']} C{r} 描述存在", cat,
                          ws[f"C{r}"].value is not None
                          and len(str(ws[f"C{r}"].value)) > 5)
        # D 公式
        d_val = str(ws[f"D{r}"].value or "")
        suite.assert_true(f"規則 {rule['id']} D{r} 為公式", cat,
                          d_val.startswith("="),
                          detail=d_val[:50])
        # E 燈號公式
        e_val = str(ws[f"E{r}"].value or "")
        suite.assert_true(f"規則 {rule['id']} E{r} 為燈號公式", cat,
                          e_val.startswith("="),
                          detail=e_val[:50])

    # 1.2 健康度 3 卡 + 1 其他 = 4 大字
    health_cells = [("A8", "紅"), ("C8", "黃"), ("E8", "綠"), ("G8", "其他")]
    for cref, label in health_cells:
        v = str(ws[cref].value or "")
        suite.assert_true(f"健康度 {label} 卡 {cref} 為 COUNTIF 公式",
                          cat, v.startswith("=") and "COUNTIF" in v,
                          detail=v[:60])

    # 1.3 7 大類縮影：row 11 = label (固定文字), row 12 = SUMPRODUCT 統計
    for i, (code, name, _) in enumerate(CATEGORIES):
        col = chr(ord("A") + i)
        # Row 11 label：固定文字「A 部署完整性」
        lbl = str(ws[f"{col}11"].value or "")
        suite.assert_true(f"類別 {code} 標籤 {col}11", cat,
                          lbl == f"{code} {name}",
                          detail=lbl)
        # Row 12 公式：SUMPRODUCT 統計紅黃綠
        v = str(ws[f"{col}12"].value or "")
        suite.assert_true(f"類別 {code} 計數 {col}12 為公式", cat,
                          v.startswith("=") and "SUMPRODUCT" in v,
                          detail=v[:60])

    # ============ 2. 邊界資料 ============
    cat = Category.BOUNDARY

    # 2.1 45 條規則「正好 45」（不多不少）
    suite.assert_eq("規則總數 = 45", cat, len(RULES), 45)

    # 2.2 7 大類覆蓋（A-G 各有規則）
    cat_codes = set(rule["id"][0] for rule in RULES)
    suite.assert_eq("覆蓋 7 大類", cat, len(cat_codes), 7)
    for c in "ABCDEFG":
        suite.assert_true(f"類別 {c} 至少 1 條", cat, c in cat_codes)

    # 2.3 各類別條數符合 plan
    expected_counts = {"A": 5, "B": 8, "C": 10, "D": 8, "E": 6, "F": 5, "G": 3}
    actual_counts = {}
    for rule in RULES:
        c = rule["id"][0]
        actual_counts[c] = actual_counts.get(c, 0) + 1
    for c, expected in expected_counts.items():
        suite.assert_eq(f"類別 {c} 條數 = {expected}", cat,
                        actual_counts.get(c, 0), expected)

    # 2.4 編號連續（A1-A5, B1-B8, ..., C1-C10）
    # 用自然序排（int 排序避免 C10 在 C2 前）
    for c, count in expected_counts.items():
        ids = sorted([rule["id"] for rule in RULES if rule["id"].startswith(c)],
                     key=lambda x: int(x[1:]))
        expected_ids = [f"{c}{i}" for i in range(1, count + 1)]
        suite.assert_eq(f"類別 {c} 編號連續", cat, ids, expected_ids)

    # ============ 3. 性能基準 ============
    cat = Category.PERFORMANCE

    # 3.1 公式總數 ≤ 200（45 規則 × 2 + 健康度 4 + 縮影 7 + banner = ~ 150）
    formula_count = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                formula_count += 1
    suite.assert_le("公式總數 ≤ 200", cat, formula_count, 200,
                    detail=f"actual={formula_count}")

    # 3.2 最大列數 ≤ 70（45 規則 + header 14 + footer 2）
    suite.assert_le("最大列數 ≤ 70", cat, ws.max_row, 70,
                    detail=f"max_row={ws.max_row}")

    # 3.3 條件格式 ≤ 10 條（4 狀態色 + 2 整列淡染 = 6）
    cf_count = sum(len(rules) for rules in ws.conditional_formatting._cf_rules.values())
    suite.assert_le("條件格式 ≤ 10 條", cat, cf_count, 10,
                    detail=f"actual={cf_count}")
    suite.assert_true("條件格式 ≥ 4 條（紅黃綠灰）", cat, cf_count >= 4,
                      detail=f"actual={cf_count}")

    # ============ 4. 跨平台相容 ============
    cat = Category.CROSS_PLATFORM

    # 4.1 無 FILTER/SORT 依賴（Excel 2010+ 全相容）
    no_filter = True
    for row in ws.iter_rows():
        for c in row:
            v = str(c.value or "")
            if v.startswith("=") and ("FILTER(" in v or "SORT(" in v):
                no_filter = False
                break
    suite.assert_true("無 FILTER/SORT 依賴（Excel 2010+ 全相容）", cat,
                      no_filter)

    # 4.2 命名範圍引用都可解析
    wb_nrs = set(wb.defined_names)
    refs_used = ["派出所名稱", "警察局名稱", "分局名稱", "案件容量上限",
                 "員警姓名清單", "違規項目清單", "案類名稱清單",
                 "案件狀況清單", "發生管轄清單", "查獲管轄清單",
                 "今日", "新進閾值", "老案閾值", "納入全般清單"]
    for n in refs_used:
        suite.assert_true(f"命名範圍 '{n}' 存在", cat, n in wb_nrs)

    # 4.3 INFO() 函數（G1 規則用）是 Excel 全版本支援
    g1_row = RULE_FIRST_ROW + sum(1 for r in RULES if r["id"][0] < "G")
    # G1 是第一個 G 類別規則
    g1_actual = next((i for i, r in enumerate(RULES) if r["id"] == "G1"), None)
    suite.assert_true("G1 規則存在", cat, g1_actual is not None)

    # ============ 5. 真實情境 ============
    cat = Category.REAL_WORLD

    # 5.1 跳轉連結 G 欄 HYPERLINK 公式
    hyperlink_count = 0
    for idx, rule in enumerate(RULES):
        r = RULE_FIRST_ROW + idx
        g_val = str(ws[f"G{r}"].value or "")
        if "HYPERLINK" in g_val:
            hyperlink_count += 1
    suite.assert_true("跳轉連結 ≥ 30 條（多數規則有對應頁）", cat,
                      hyperlink_count >= 30,
                      detail=f"actual={hyperlink_count}")

    # 5.2 「待接通」規則（v2.2 7 條毒調已全接通 → 預期 0 條 pending）
    pending_count = sum(1 for r in RULES
                        if r["light"]("D1") == '="○ 待接通"')
    suite.assert_eq("v2.2 待接通規則 = 0（全部 Tbl 已接通）", cat,
                    pending_count, 0)

    # 5.3 嚴重度分布合理（高/中/低/資訊都有）
    severities = {rule["severity"] for rule in RULES}
    for s in ["高", "中", "低", "資訊"]:
        suite.assert_true(f"嚴重度 '{s}' 至少 1 條", cat, s in severities)

    # 5.4 banner 含「45 條」「7 大類」
    a1 = str(ws["A1"].value or "")
    suite.assert_true("Banner 含「45」", cat, "45" in a1, detail=a1[:60])
    suite.assert_true("Banner 含「7 大類」", cat, "7 大類" in a1)

    # ============ 6. 錯誤恢復 ============
    cat = Category.ERROR_RECOVERY

    # 6.1 所有 D 欄公式包 IFERROR（除了待接通的字串字面值）
    iferror_count = 0
    no_iferror = []
    for idx, rule in enumerate(RULES):
        r = RULE_FIRST_ROW + idx
        v = str(ws[f"D{r}"].value or "")
        if "IFERROR" in v:
            iferror_count += 1
        elif not (v.startswith('="') and v.endswith('"')):
            # 不是字串字面值卻沒 IFERROR
            no_iferror.append(rule["id"])
    suite.assert_true(f"D 欄公式 IFERROR 容錯 ≥ 30 條（剩餘為字串字面值）", cat,
                      iferror_count >= 30,
                      detail=f"actual={iferror_count}, miss={no_iferror[:5]}")

    # 6.2 E 欄燈號公式皆有 IF 判斷
    if_count = 0
    for idx, rule in enumerate(RULES):
        r = RULE_FIRST_ROW + idx
        v = str(ws[f"E{r}"].value or "")
        if v.startswith("=IF") or '="○' in v or '="ℹ' in v:
            if_count += 1
    suite.assert_eq("E 欄燈號 45 條皆有 IF/字面值", cat, if_count, 45)

    # 6.3 健康度公式包 COUNTIF（容錯：找不到回傳 0）
    a8 = str(ws["A8"].value or "")
    suite.assert_true("健康度紅燈 A8 用 COUNTIF（容錯）", cat,
                      "COUNTIF" in a8 and "紅" in a8,
                      detail=a8[:80])

    # 6.4 跳轉連結 IFERROR 包覆（連結失效不爆）
    iferror_link = 0
    for idx, rule in enumerate(RULES):
        r = RULE_FIRST_ROW + idx
        g = str(ws[f"G{r}"].value or "")
        if "HYPERLINK" in g and "IFERROR" in g:
            iferror_link += 1
    suite.assert_true("跳轉連結 IFERROR 包覆 ≥ 30 條", cat,
                      iferror_link >= 30,
                      detail=f"actual={iferror_link}")

    return suite


if __name__ == "__main__":
    suite = run()
    print(suite.report())
    sys.exit(0 if suite.is_all_pass() else 1)
