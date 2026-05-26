"""
test_phase3_page8.py — 頁 8 毒品調驗 6 大類高強度測試
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook

from phases._base import OUTPUT_PATH
from tests.framework import TestSuite, Category


DRUG_SHEET = "毒品調驗人口管制"
OVERVIEW_SHEET = "管制總覽"


def run():
    suite = TestSuite("頁 8 毒品調驗")
    wb = load_workbook(OUTPUT_PATH)
    ws = wb[DRUG_SHEET]
    ws2 = wb[OVERVIEW_SHEET]

    # =========================================================
    # 1. 公式語法
    # =========================================================
    cat = Category.SYNTAX

    # 1.1 Tbl毒調 已註冊（v2.2 擴 14 欄：A14:N84）
    tables = list(ws.tables.values())
    drug_tbls = [t for t in tables if t.name == "Tbl毒調"]
    suite.assert_eq("Tbl毒調 已註冊", cat, len(drug_tbls), 1)
    if drug_tbls:
        suite.assert_eq("Tbl毒調 ref = A14:N84（v2.2 擴 14 欄）", cat,
                        drug_tbls[0].ref, "A14:N84")

    # 1.2 v43 9 欄 header 維持不動
    expected_v43_headers = [
        "編號", "姓名", "身分證", "性別", "住所", "聯繫", "通緝", "管制情形", "備註"
    ]
    for i, h in enumerate(expected_v43_headers, start=1):
        col = chr(ord("A") + i - 1)
        suite.assert_eq(f"v43 header {col}14 = {h}", cat,
                        ws[f"{col}14"].value, h)

    # 1.3 v2.2 5 個 J-N header（J 最後到驗日 input + K-N 計算欄）
    for col, name in [("J", "最後到驗日"), ("K", "距今天數"),
                      ("L", "是否到驗"), ("M", "狀態燈"), ("N", "候選旗標")]:
        suite.assert_eq(f"欄 {col}14 = {name}", cat,
                        ws[f"{col}14"].value, name)

    # 1.4 K 距今天數公式（=今日-[@最後到驗日]）
    k_fml = str(ws["K15"].value or "")
    suite.assert_true("K15 距今天數 = 今日-[@最後到驗日]", cat,
                      "今日" in k_fml and "[@最後到驗日]" in k_fml,
                      detail=k_fml[:80])

    # 1.5 L 是否到驗（v2.2 升級：含 4 種狀態 OR 距今 ≤ 30）
    l_fml = str(ws["L15"].value or "")
    suite.assert_true("L15 是否到驗包 SUMPRODUCT+SEARCH 多關鍵字", cat,
                      "SUMPRODUCT" in l_fml and "SEARCH" in l_fml,
                      detail=l_fml[:80])
    for keyword in ["已驗", "通緝", "強採", "在監"]:
        suite.assert_true(f"是否到驗公式含 '{keyword}'", cat,
                          keyword in l_fml, detail=l_fml[:80])
    suite.assert_true("v2.2 升級：是否到驗含 距今 ≤ 30 判斷", cat,
                      "[@距今天數]" in l_fml and "<=30" in l_fml,
                      detail=l_fml[:120])

    # 1.6 M 狀態燈公式
    m_fml = str(ws["M15"].value or "")
    suite.assert_true("狀態燈包 IFERROR", cat, "IFERROR" in m_fml)
    for keyword in ["未填", "結束", "● 紅", "● 綠", "⚠ 黃"]:
        suite.assert_true(f"狀態燈含 '{keyword}'", cat,
                          keyword in m_fml, detail=m_fml[:80])
    suite.assert_true("v2.2 升級：狀態燈含 距今 > 30 判紅", cat,
                      "[@距今天數]" in m_fml and ">30" in m_fml,
                      detail=m_fml[:120])

    # 1.7 N 候選旗標：未到驗 OR 距今 > 30（v2.2 升級）
    n_fml = str(ws["N15"].value or "")
    suite.assert_true("候選旗標 = 未到驗 OR 距今>30", cat,
                      "未到驗" in n_fml and "[@距今天數]" in n_fml,
                      detail=n_fml[:80])

    # 1.8 計算欄使用結構引用 [@...]
    for col, ref_keyword in [("K", "[@最後到驗日]"), ("L", "[@管制情形]"),
                             ("M", "[@是否到驗]"), ("N", "[@管制情形]")]:
        f = str(ws[f"{col}15"].value or "")
        suite.assert_true(f"{col}15 用結構引用 {ref_keyword}", cat,
                          ref_keyword in f, detail=f[:80])

    # 1.8 status cards 公式（row 9）
    expected_cards = {
        "A9": "COUNTA(Tbl毒調[姓名])",
        "C9": "SUM(Tbl毒調[是否到驗])",
        "G9": "SUM(Tbl毒調[候選旗標])",
    }
    for cell, fragment in expected_cards.items():
        f = str(ws[cell].value or "")
        suite.assert_true(f"{cell} status card 含 '{fragment}'", cat,
                          fragment in f, detail=f[:80])

    # 1.9 毒調率 KPI bar (row 11) 引用毒調率目標
    a11 = str(ws["A11"].value or "")
    suite.assert_true("毒調率 KPI bar 引用毒調率目標", cat,
                      "毒調率目標" in a11, detail=a11[:80])
    suite.assert_true("毒調率 KPI bar 公式包 IFERROR", cat,
                      "IFERROR" in a11)
    suite.assert_true("毒調率 KPI bar 顯示達標狀態", cat,
                      "達標" in a11 and "未達標" in a11)

    # 1.10 進度條（row 12）REPT 滿格 + 空格分明
    a12 = str(ws["A12"].value or "")
    suite.assert_true("進度條用 REPT('█',n)", cat,
                      'REPT("█"' in a12 and "20" in a12, detail=a12[:80])
    suite.assert_true("進度條空段 REPT('░',n)", cat,
                      'REPT("░"' in a12, detail=a12[:80])

    # 1.11 候選清單主表（v2.2: P-S row 15）用 INDEX+MATCH
    for col, src_field in [("P", "Tbl毒調[編號]"), ("Q", "Tbl毒調[姓名]"),
                           ("R", "Tbl毒調[管制情形]")]:
        f = str(ws[f"{col}15"].value or "")
        suite.assert_true(f"候選 {col}15 用 INDEX+MATCH 取 {src_field}", cat,
                          "INDEX" in f and "MATCH" in f and src_field in f,
                          detail=f[:80])
        suite.assert_true(f"候選 {col}15 包 IFERROR", cat, "IFERROR" in f)

    # 1.12 helper area v2.2 移到 T/U/V cols (rows 100-169)
    suite.assert_eq("helper T100 序號 = 1", cat, ws["T100"].value, 1)
    u100 = str(ws["U100"].value or "")
    suite.assert_true("helper U100 引用 Tbl毒調[候選旗標]", cat,
                      "INDEX(Tbl毒調[候選旗標]" in u100, detail=u100[:60])
    v100 = str(ws["V100"].value or "")
    suite.assert_true("helper V100 累加候選順序", cat,
                      "SUM" in v100, detail=v100[:60])

    # 1.13 頁 2 KPI 4 毒調率 已接通 Tbl毒調
    p2_a14 = str(ws2["A14"].value or "")
    suite.assert_true("頁 2 KPI 4 接通 Tbl毒調", cat,
                      "Tbl毒調" in p2_a14, detail=p2_a14[:80])
    suite.assert_true("頁 2 KPI 4 包 IFERROR", cat, "IFERROR" in p2_a14)

    # 1.14 頁 2 KPI 5 未到驗 已接通
    p2_g14 = str(ws2["G14"].value or "")
    suite.assert_true("頁 2 KPI 5 接通 Tbl毒調", cat,
                      "Tbl毒調" in p2_g14, detail=p2_g14[:80])

    # 1.15 頁 2 KPI 4 燈號 引用毒調率目標
    p2_e13 = str(ws2["E13"].value or "")
    suite.assert_true("頁 2 KPI 4 燈號引用毒調率目標", cat,
                      "毒調率目標" in p2_e13, detail=p2_e13[:80])

    # =========================================================
    # 2. 邊界資料
    # =========================================================
    cat = Category.BOUNDARY

    # 2.1 5 sample rows 含業務規則 4 狀態（已驗/通緝/強採/在監/未到驗）
    sample_status = []
    for r in range(15, 20):
        v = str(ws[f"H{r}"].value or "")
        sample_status.append(v)
    suite.assert_true("sample 含「已到驗」", cat,
                      any("已到驗" in s for s in sample_status))
    suite.assert_true("sample 含「未到驗」", cat,
                      any("未到驗" in s for s in sample_status))
    suite.assert_true("sample 含「通緝中」", cat,
                      any("通緝中" in s for s in sample_status))
    suite.assert_true("sample 含「已聲強採」", cat,
                      any("已聲強採" in s for s in sample_status))
    suite.assert_true("sample 含「在監」", cat,
                      any("在監" in s for s in sample_status))

    # 2.2 空表時 status cards 不爆（IFERROR / SUM 處理空表 = 0）
    # SUM/COUNTA 對空表都返回 0，不會錯誤
    # 主要看毒調率分母 0 處理
    a11 = str(ws["A11"].value or "")
    suite.assert_true("毒調率分母 0 處理（IFERROR + COUNTA）", cat,
                      "IFERROR" in a11 and "COUNTA" in a11,
                      detail=a11[:80])

    # 2.3 容量 70 列：Tbl毒調 ref 應包含 A14:L84
    # 已在 1.1 驗證 ref

    # 2.4 v2.2 隱藏 cols K/L/M/N (Tbl 計算欄) + T/U/V (helper)
    # J 最後到驗日 input 可見、K-N 隱藏；T-V helper 隱藏
    hidden = [letter for letter, d in ws.column_dimensions.items() if d.hidden]
    for col in ["K", "L", "M", "N", "T", "U", "V"]:
        suite.assert_true(f"col {col} 已隱藏", cat, col in hidden,
                          detail=f"hidden={hidden}")

    # 2.5 候選清單最多 16 列，註腳明示（v2.2: P{31}）
    note_row = 15 + 16  # row 31
    note = str(ws[f"P{note_row}"].value or "")
    suite.assert_true("候選清單註腳「最多顯示 16 個」", cat,
                      "16" in note and "候選" in note, detail=note[:60])

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
    suite.assert_le("頁 8 公式總數 ≤ 500", cat, fcount, 500,
                    detail=f"actual={fcount}")

    # 揮發性函數
    volatile_counts = {"TODAY()": 0, "NOW()": 0, "OFFSET(": 0, "INDIRECT(": 0}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                for k in volatile_counts:
                    volatile_counts[k] += cell.value.count(k)
    suite.assert_eq("不直接呼叫 TODAY()", cat, volatile_counts["TODAY()"], 0)
    suite.assert_eq("不直接呼叫 NOW()", cat, volatile_counts["NOW()"], 0)
    suite.assert_le("OFFSET 使用 ≤ 3", cat, volatile_counts["OFFSET("], 3)
    suite.assert_le("INDIRECT 使用 ≤ 3", cat, volatile_counts["INDIRECT("], 3)

    # max_row 包含 Tbl 70 列 + helper 70 列
    suite.assert_le("頁 8 max_row ≤ 175", cat, ws.max_row, 175,
                    detail=f"max_row={ws.max_row}")

    # =========================================================
    # 4. 跨平台相容
    # =========================================================
    cat = Category.CROSS_PLATFORM

    has_filter_sort = False
    for row in ws.iter_rows():
        for cell in row:
            v = str(cell.value or "")
            if "FILTER(" in v or "SORT(" in v:
                has_filter_sort = True
                break
    suite.assert_true("頁 8 無 FILTER/SORT 依賴", cat, not has_filter_sort)

    # 命名範圍
    wb_nrs = set(wb.defined_names)
    for nr in ["毒調率目標", "性別清單", "通緝清單", "毒品管制情形清單"]:
        suite.assert_true(f"命名範圍 '{nr}' 存在", cat, nr in wb_nrs)

    # 資料驗證 (下拉) 已設
    dvs = list(ws.data_validations.dataValidation)
    suite.assert_true("資料驗證至少 3 項（性別/通緝/管制情形）", cat,
                      len(dvs) >= 3, detail=f"count={len(dvs)}")

    # =========================================================
    # 5. 真實情境
    # =========================================================
    cat = Category.REAL_WORLD

    # Banner / sub-banner / 維護指引
    a1 = str(ws["A1"].value or "")
    suite.assert_true("Banner 含「毒品調驗」", cat, "毒品調驗" in a1)
    suite.assert_true("Banner 註記 v43 9 欄基礎", cat,
                      "v43" in a1 and "9 欄" in a1)

    a2 = str(ws["A2"].value or "")
    suite.assert_true("sub-banner 註記業務規則", cat,
                      "通緝/強採/在監" in a2 and "算到驗" in a2)

    a4 = str(ws["A4"].value or "")
    suite.assert_true("維護指引提示主表 9 欄不變", cat,
                      "9 欄 v43 不動" in a4)

    # 凍結 A14
    suite.assert_eq("凍結窗格 = None（使用者要求全頁無凍結）", cat,
                    ws.freeze_panes, None)

    # 業務規則：sample 003 通緝中 / 005 在監 → 「算到驗」
    # v2.2: 公式 L15: 含 SEARCH "通緝" "在監"
    l_fml = str(ws["L15"].value or "")
    for status in ["通緝", "強採", "在監"]:
        suite.assert_true(f"算到驗業務規則含「{status}」", cat,
                          status in l_fml)

    # 候選清單放右側 (v2.2: cols P-S)，與主表並讀
    cand_cols = ["P", "Q", "R", "S"]
    for col in cand_cols:
        suite.assert_true(f"候選欄 {col}14 header 存在", cat,
                          ws[f"{col}14"].value is not None,
                          detail=f"val={ws[f'{col}14'].value}")

    # =========================================================
    # 6. 錯誤恢復
    # =========================================================
    cat = Category.ERROR_RECOVERY

    # 6.1 v2.2 計算欄 4 欄全包 IFERROR (K/L/M/N)
    for col in ["K", "L", "M", "N"]:
        f = str(ws[f"{col}15"].value or "")
        suite.assert_true(f"計算欄 {col}15 包 IFERROR", cat,
                          "IFERROR" in f, detail=f[:60])

    # 6.2 status cards 沒用 IFERROR 也安全（SUM/COUNTA 對空 = 0）
    # 但毒調率 (除法) 必須 IFERROR
    a11 = str(ws["A11"].value or "")
    suite.assert_true("毒調率公式 IFERROR 包覆", cat,
                      "IFERROR" in a11)

    # 6.3 進度條 ROUND 限制 0-20
    a12 = str(ws["A12"].value or "")
    suite.assert_true("進度條 MIN(20,...) 限制最大值", cat,
                      "MIN(20" in a12, detail=a12[:80])

    # 6.4 候選清單 INDEX/MATCH 包 IFERROR（v2.2: P15）
    p15 = str(ws["P15"].value or "")
    suite.assert_true("候選 P15 INDEX/MATCH 包 IFERROR", cat,
                      "IFERROR" in p15)

    # 6.5 helper V 公式不會錯（U=0 時返回 0）
    v100 = str(ws["V100"].value or "")
    suite.assert_true("helper V 順序公式：非候選 = 0", cat,
                      "IF(U100=1" in v100 and ",0)" in v100,
                      detail=v100[:60])

    return suite


if __name__ == "__main__":
    suite = run()
    print(suite.report())
    sys.exit(0 if suite.is_all_pass() else 1)
