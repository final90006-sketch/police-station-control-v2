"""
test_phase5_polish.py — Phase 5 高強度測試（6 大類）

驗證：
  1. CF 規則確實掛在指定範圍上（頁 5/6/7/9）
  2. 案件資料庫 A2 容量警示公式 + 2 條 CF
  3. 工作表保護（密碼 KD2026 + 鎖定設定 + 解鎖區域）
  4. 列印區 + A4 直式 + fitToPage（頁 2.5 / 3 / 4 / 11）
  5. 既有 v2.1 公式不被破壞（取樣關鍵 cell 確認 still =...）
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook

from phases._base import OUTPUT_PATH, load_config
from tests.framework import TestSuite, Category


PASSWORD = "KD2026"


def _count_cf_rules(ws, range_str=None):
    """數工作表上 CF 規則總數。range_str=None 時數全表。"""
    total = 0
    for rng, rules in ws.conditional_formatting._cf_rules.items():
        if range_str is None:
            total += len(rules)
        elif str(rng) == range_str:
            total += len(rules)
    return total


def _has_cf_on_range(ws, range_str):
    """檢查指定範圍是否掛有 CF。"""
    for rng, rules in ws.conditional_formatting._cf_rules.items():
        if str(rng) == range_str or range_str in str(rng):
            return len(rules) > 0
    return False


def run():
    suite = TestSuite("Phase 5 — CF + 保護 + 列印區")
    wb = load_workbook(OUTPUT_PATH)

    # ============ 1. 公式語法 ============
    cat = Category.SYNTAX

    # 1.1 頁 5 A2 容量警示公式
    ws5 = wb["案件資料庫"]
    a2 = str(ws5["A2"].value or "")
    suite.assert_true("頁 5 A2 容量警示為公式", cat,
                      a2.startswith("="),
                      detail=a2[:60])
    suite.assert_true("頁 5 A2 公式引用 Tbl案件[編號]", cat,
                      "Tbl案件[編號]" in a2)
    suite.assert_true("頁 5 A2 公式含 800 / 950 雙閾值", cat,
                      "800" in a2 and "950" in a2)

    # 1.2 頁 6/7/9 既有公式未被破壞（取樣）
    ws6 = wb["刑案管制"]
    f_a7 = str(ws6["A7"].value or "")
    suite.assert_true("頁 6 KPI 縮影 A7 公式 still active", cat,
                      f_a7.startswith("=") and "尚未偵破" in f_a7)

    ws9 = wb["交通績效管制"]
    f_d14 = str(ws9["D14"].value or "")
    suite.assert_true("頁 9 達成率 D14 公式 still active", cat,
                      f_d14.startswith("=") and "IFERROR" in f_d14)

    # ============ 2. 邊界資料 ============
    cat = Category.BOUNDARY

    # 2.1 CF 規則數量符合預期（頁 6/7/9 各自）
    ws6_cf_count = _count_cf_rules(ws6)
    suite.assert_true("頁 6 CF 規則 ≥ 4 條（兩區段案件狀況）", cat,
                      ws6_cf_count >= 4,
                      detail=f"actual={ws6_cf_count}")

    ws7 = wb["發展中案件管制"]
    ws7_cf_count = _count_cf_rules(ws7)
    suite.assert_true("頁 7 CF 規則 ≥ 5 條（狀態燈 + 整列淡染）", cat,
                      ws7_cf_count >= 5,
                      detail=f"actual={ws7_cf_count}")

    ws9_cf_count = _count_cf_rules(ws9)
    suite.assert_true("頁 9 CF 規則 ≥ 6 條（達成率 + 狀態欄）", cat,
                      ws9_cf_count >= 6,
                      detail=f"actual={ws9_cf_count}")

    ws5_cf_count = _count_cf_rules(ws5)
    suite.assert_true("頁 5 CF 規則 ≥ 2 條（容量警示）", cat,
                      ws5_cf_count >= 2,
                      detail=f"actual={ws5_cf_count}")

    # 2.2 CF 範圍正確掛載
    suite.assert_true("頁 6 CF 掛 A11:I25（尚未偵破）", cat,
                      _has_cf_on_range(ws6, "A11:I25"))
    suite.assert_true("頁 6 CF 掛 A29:I43（已破獲未移送）", cat,
                      _has_cf_on_range(ws6, "A29:I43"))
    suite.assert_true("頁 7 CF 掛 K14:K28（狀態欄）", cat,
                      _has_cf_on_range(ws7, "K14:K28"))
    suite.assert_true("頁 9 CF 掛 D14（達成率）", cat,
                      any("D14" in str(rng) for rng in ws9.conditional_formatting._cf_rules))
    suite.assert_true("頁 5 CF 掛 A2（容量警示）", cat,
                      _has_cf_on_range(ws5, "A2"))

    # ============ 3. 性能基準 ============
    cat = Category.PERFORMANCE

    # 3.1 每頁 CF 規則上限（plan 規範：≤ 8 為目標）
    # 頁 2 有 54 條（前一 session 已上的 9 KPI × 2 cells × 3 規則，特殊豁免）
    # 頁 7/9 是文字 + 整列染色雙層 → 上限放寬到 8（status 3 + row tint 2 / status 3 + rate 3）
    thresholds = {
        "管制總覽": 60,         # 頁 2 豁免
        "刑案管制": 8,
        "發展中案件管制": 8,    # 5 條（3 status + 2 row tint）
        "交通績效管制": 8,      # 6 條（3 rate + 3 status）
        "案件資料庫": 4,
    }
    for sheet_name, threshold in thresholds.items():
        ws = wb[sheet_name]
        n = _count_cf_rules(ws)
        suite.assert_le(f"{sheet_name} CF 規則 ≤ {threshold}", cat,
                        n, threshold, detail=f"actual={n}")

    # 3.2 檔案大小 < 500 KB
    size_kb = OUTPUT_PATH.stat().st_size / 1024
    suite.assert_le("檔案大小 ≤ 500 KB", cat, size_kb, 500,
                    detail=f"actual={size_kb:.1f} KB")

    # ============ 4. 跨平台 — 工作表保護 ============
    cat = Category.CROSS_PLATFORM

    # 4.1 案件資料庫保護啟用 + 密碼正確
    suite.assert_true("頁 5 案件資料庫 protection.sheet=True", cat,
                      ws5.protection.sheet)
    suite.assert_true("頁 5 案件資料庫 密碼設定（hash 存在）", cat,
                      ws5.protection.password is not None
                      and ws5.protection.password != "")

    # 4.2 統計頁全鎖
    locked_sheets = ["管制總覽", "全般刑案管制情形分析", "刑案管制",
                     "毒品調驗人口管制", "交通績效管制",
                     "績效統計", "跨期間趨勢分析", "系統檢核"]
    for name in locked_sheets:
        ws = wb[name]
        suite.assert_true(f"{name} 保護啟用", cat,
                          ws.protection.sheet,
                          detail=f"sheet={ws.protection.sheet}")

    # 4.3 解鎖 cell 確實 locked=False
    # 頁 5 A15（編號欄第一筆 input） — 應解鎖
    a15 = ws5["A15"]
    suite.assert_true("頁 5 A15（編號 input） protection.locked=False", cat,
                      a15.protection.locked is False,
                      detail=f"locked={a15.protection.locked}")
    # 頁 5 R15（是否破獲 calc） — 應鎖（預設）
    r15 = ws5["R15"]
    suite.assert_true("頁 5 R15（是否破獲 calc） 預設鎖", cat,
                      r15.protection.locked is True or r15.protection.locked is None,
                      detail=f"locked={r15.protection.locked}")

    # 頁 3 行動方案區 A24 解鎖（A:H merged，anchor 在 A）
    ws3 = wb["長官5分鐘報告"]
    a24 = ws3["A24"]
    suite.assert_true("頁 3 A24（行動方案手動 merged anchor） 解鎖", cat,
                      a24.protection.locked is False,
                      detail=f"locked={a24.protection.locked}")

    # 頁 11 員警下拉 B3 解鎖
    ws11 = wb["員警個人績效卡"]
    b3 = ws11["B3"]
    suite.assert_true("頁 11 B3（員警下拉） 解鎖", cat,
                      b3.protection.locked is False,
                      detail=f"locked={b3.protection.locked}")

    # ============ 5. 真實情境 — 列印區 ============
    cat = Category.REAL_WORLD

    # 5.1 4 頁列印區 + 頁面設定
    # openpyxl 寫入後格式為 '<sheet>'!$A$1:$H$32（含 $）
    print_sheets = [
        ("全般刑案管制情形分析", "$A$1:$H$32"),
        ("長官5分鐘報告",       "$A$1:$H$36"),
        ("對上級機關上呈",       "$A$1:$H$28"),
        ("員警個人績效卡",       "$A$1:$H$50"),
    ]
    for name, expected_range in print_sheets:
        ws = wb[name]
        pa = ws.print_area or ""
        suite.assert_true(f"{name} print_area 含 {expected_range}",
                          cat, expected_range in str(pa),
                          detail=f"actual={pa!r}")
        # A4 直式
        suite.assert_eq(f"{name} 方向=portrait", cat,
                        ws.page_setup.orientation, "portrait")
        # fitToPage
        suite.assert_true(f"{name} fitToPage 啟用", cat,
                          bool(ws.sheet_properties.pageSetUpPr.fitToPage),
                          detail=f"actual={ws.sheet_properties.pageSetUpPr.fitToPage}")
        # PAPERSIZE_A4 = 9
        suite.assert_eq(f"{name} 紙張=A4", cat,
                        ws.page_setup.paperSize, 9)

    # 5.2 容量警示文字含使用者友善訊息
    a2_val = str(ws5["A2"].value or "")
    for keyword in ["容量", "歸檔", "受保護", "KD2026"]:
        suite.assert_true(f"A2 容量訊息含「{keyword}」", cat,
                          keyword in a2_val,
                          detail=a2_val[:80])

    # ============ 6. 錯誤恢復 ============
    cat = Category.ERROR_RECOVERY

    # 6.1 CF 公式都包 ISNUMBER 或 AND 容錯（避免空 cell 觸發誤判）
    # 取樣頁 7 K14:K28
    cf_formulas = []
    for rng, rules in ws7.conditional_formatting._cf_rules.items():
        if "K14" in str(rng):
            for rule in rules:
                if rule.formula:
                    cf_formulas.extend(rule.formula)
    suite.assert_true("頁 7 K 欄 CF 公式用 ISNUMBER(SEARCH)（容空白）", cat,
                      any("ISNUMBER(SEARCH" in f for f in cf_formulas),
                      detail=f"got {len(cf_formulas)} formulas")

    # 6.2 頁 9 達成率 CF 用 AND(ISNUMBER, ...)（空 cell 不誤觸發）
    cf_formulas9 = []
    for rng, rules in ws9.conditional_formatting._cf_rules.items():
        if "D14" in str(rng):
            for rule in rules:
                if rule.formula:
                    cf_formulas9.extend(rule.formula)
    suite.assert_true("頁 9 D 欄 CF 公式用 AND(ISNUMBER)（容空白）", cat,
                      any("AND(ISNUMBER" in f for f in cf_formulas9),
                      detail=f"got {len(cf_formulas9)} formulas")

    # 6.3 容量警示三段 IF（容量 ≥950 / ≥800 / else）
    suite.assert_true("A2 公式含三段警示（即將滿 / 警戒 / 充裕）", cat,
                      "即將滿" in a2_val and "警戒" in a2_val and "充裕" in a2_val,
                      detail=a2_val[:120])

    # 6.4 既有 v2.1 公式樣本 — 頁 2 KPI / 頁 6 helper / 頁 11 員警 KPI 仍 active
    ws2 = wb["管制總覽"]
    suite.assert_true("頁 2 A8 KPI label 完整", cat,
                      ws2["A8"].value is not None
                      and "全般破獲率" in str(ws2["A8"].value))

    return suite


if __name__ == "__main__":
    suite = run()
    print(suite.report())
    sys.exit(0 if suite.is_all_pass() else 1)
