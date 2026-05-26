"""
phase5_polish.py — Phase 5：條件格式 + 工作表保護 + 列印區

收尾打磨層。不重 build 任何頁，只把 v2.1 已產出的 PoliceStation_v2.0.xlsx
加上：
  1. 紅黃綠燈條件格式（頁 6 / 7 / 9 — 頁 2 已上 54 條，由 phase3_page2_overview 負責）
  2. 案件資料庫容量警示 CF（≥800 黃 / ≥950 紅）
  3. 工作表保護（密碼 KD2026 — 案件資料庫鎖計算欄、統計頁全鎖、設定表分區鎖）
  4. 列印區 + 頁面設定（頁 2.5 / 3 / 4 / 11 → A4 直式 1 頁）

設計原則：
- 不動既有 cell 內容（公式 / 樣式都已就緒）
- 只加 CF 規則 / cell.protection / ws.print_area / ws.page_setup
- 每 sheet 上 CF 規則控制在 ≤ 8 條（per plan 性能規範）
- 保護模式：sheet=True + password；input 欄 cell.protection.locked=False 解鎖
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill, Font, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from phases._base import backup_existing, OUTPUT_PATH, get_logger, load_config


# ============================================================
# 共用樣式（與頁 2 已上的 CF 顏色一致 — styles.py COLOR 對應）
# ============================================================
FILL_RED    = PatternFill("solid", fgColor="FFFEE2E2")   # danger_bg
FILL_YELLOW = PatternFill("solid", fgColor="FFFEF3C7")   # warn_bg
FILL_GREEN  = PatternFill("solid", fgColor="FFDCFCE7")   # pass_bg

FONT_RED    = Font(name="微軟正黑體", bold=True, color="FFB91C1C")
FONT_YELLOW = Font(name="微軟正黑體", bold=True, color="FFB45309")
FONT_GREEN  = Font(name="微軟正黑體", bold=True, color="FF15803D")

PASSWORD = "KD2026"


def _clear_all_cf(ws):
    """清掉 sheet 上所有條件格式（讓 Phase 5 冪等）。

    Phase 5 寫 CF 到頁 5/6/7/9，那些頁原本 NO CF（phase 2/3 不上 CF），
    所以全清不會破壞別人。頁 2 的 54 條 CF 由 phase3_page2_overview 負責，
    Phase 5 不該動。
    """
    ws.conditional_formatting._cf_rules = {}


# ============================================================
# (1) 條件格式：頁 6 兩區段案件狀況
# ============================================================
def cf_page6(wb, log):
    """頁 6 刑案管制 — 兩區段案件狀況欄自動上色

    結構（v2.1 兩區段並呈版）：
      Section 1 尚未偵破：rows 11-25，狀態固定「尚未偵破」→ 紅底
      Section 2 已破獲未移送：rows 29-43，狀態固定「已破獲未移送」→ 黃底

    CF 規則用文字偵測，即使 INDEX 抓到空白也不誤觸發。
    """
    ws = wb["刑案管制"]
    _clear_all_cf(ws)   # 冪等：重 build 不重複疊加
    rules_added = 0

    # Section 1 尚未偵破清單 — 整列染淡紅，狀態欄 F 染深紅
    sec1_range = "A11:I25"
    # 整列淡紅底：當 F 欄有「尚未偵破」字樣
    ws.conditional_formatting.add(sec1_range,
        FormulaRule(formula=['ISNUMBER(SEARCH("尚未偵破",$F11))'],
                    fill=FILL_RED, stopIfTrue=False))
    rules_added += 1
    # F 欄狀態文字加紅
    ws.conditional_formatting.add("F11:F25",
        FormulaRule(formula=['ISNUMBER(SEARCH("尚未偵破",F11))'],
                    fill=FILL_RED, font=FONT_RED, stopIfTrue=False))
    rules_added += 1

    # Section 2 已破獲未移送清單 — 整列染淡黃，狀態欄 F 染深橘
    sec2_range = "A29:I43"
    ws.conditional_formatting.add(sec2_range,
        FormulaRule(formula=['ISNUMBER(SEARCH("已破獲未移送",$F29))'],
                    fill=FILL_YELLOW, stopIfTrue=False))
    rules_added += 1
    ws.conditional_formatting.add("F29:F43",
        FormulaRule(formula=['ISNUMBER(SEARCH("已破獲未移送",F29))'],
                    fill=FILL_YELLOW, font=FONT_YELLOW, stopIfTrue=False))
    rules_added += 1

    # KPI 縮影 3 卡（row 7）— 已寫死視覺色，不需 CF

    log.info(f"  頁 6 條件格式：{rules_added} 條（兩區段案件狀況自動染色）")
    return rules_added


# ============================================================
# (2) 條件格式：頁 7 發展中狀態燈
# ============================================================
def cf_page7(wb, log):
    """頁 7 發展中案件管制 — 狀態欄 K 自動染色

    狀態欄 K 公式回傳「● 綠」「⚠ 黃」「✕ 紅」
    範圍：K14:K28（5 示範 + 10 預留 = 15 列）
    """
    ws = wb["發展中案件管制"]
    _clear_all_cf(ws)
    rules_added = 0

    status_range = "K14:K28"

    # 紅燈
    ws.conditional_formatting.add(status_range,
        FormulaRule(formula=['ISNUMBER(SEARCH("紅",K14))'],
                    fill=FILL_RED, font=FONT_RED, stopIfTrue=False))
    rules_added += 1
    # 黃燈
    ws.conditional_formatting.add(status_range,
        FormulaRule(formula=['ISNUMBER(SEARCH("黃",K14))'],
                    fill=FILL_YELLOW, font=FONT_YELLOW, stopIfTrue=False))
    rules_added += 1
    # 綠燈
    ws.conditional_formatting.add(status_range,
        FormulaRule(formula=['ISNUMBER(SEARCH("綠",K14))'],
                    fill=FILL_GREEN, font=FONT_GREEN, stopIfTrue=False))
    rules_added += 1

    # 整列依狀態淡染（A:K，col K 是 status）
    row_range = "A14:K28"
    ws.conditional_formatting.add(row_range,
        FormulaRule(formula=['ISNUMBER(SEARCH("紅",$K14))'],
                    fill=FILL_RED, stopIfTrue=False))
    rules_added += 1
    ws.conditional_formatting.add(row_range,
        FormulaRule(formula=['ISNUMBER(SEARCH("黃",$K14))'],
                    fill=FILL_YELLOW, stopIfTrue=False))
    rules_added += 1

    log.info(f"  頁 7 條件格式：{rules_added} 條（狀態欄燈號 + 整列淡染）")
    return rules_added


# ============================================================
# (3) 條件格式：頁 9 達成率燈號
# ============================================================
def cf_page9(wb, log):
    """頁 9 交通績效看板 — 達成率欄 D + 狀態欄 F 自動染色

    範圍：D14:D{14+violations}、F14:F{14+violations}（含合計列）
    達成率：>=80% 綠 / >=60% 黃 / <60% 紅
    狀態欄公式回傳「● 綠」「⚠ 黃」「● 紅」
    """
    ws = wb["交通績效管制"]
    _clear_all_cf(ws)
    p15 = load_config("phase1_5_settings")
    violations_count = len(p15["section_e"]["violations"])
    total_row = 14 + violations_count   # 含合計列

    rate_range = f"D14:D{total_row}"
    status_range = f"F14:F{total_row}"

    rules_added = 0

    # 達成率 D 欄：用數值判斷
    ws.conditional_formatting.add(rate_range,
        FormulaRule(formula=[f'AND(ISNUMBER(D14),D14>=0.8)'],
                    fill=FILL_GREEN, font=FONT_GREEN, stopIfTrue=False))
    rules_added += 1
    ws.conditional_formatting.add(rate_range,
        FormulaRule(formula=[f'AND(ISNUMBER(D14),D14>=0.6,D14<0.8)'],
                    fill=FILL_YELLOW, font=FONT_YELLOW, stopIfTrue=False))
    rules_added += 1
    ws.conditional_formatting.add(rate_range,
        FormulaRule(formula=[f'AND(ISNUMBER(D14),D14<0.6)'],
                    fill=FILL_RED, font=FONT_RED, stopIfTrue=False))
    rules_added += 1

    # 狀態欄 F：用文字偵測
    ws.conditional_formatting.add(status_range,
        FormulaRule(formula=['ISNUMBER(SEARCH("紅",F14))'],
                    fill=FILL_RED, font=FONT_RED, stopIfTrue=False))
    rules_added += 1
    ws.conditional_formatting.add(status_range,
        FormulaRule(formula=['ISNUMBER(SEARCH("黃",F14))'],
                    fill=FILL_YELLOW, font=FONT_YELLOW, stopIfTrue=False))
    rules_added += 1
    ws.conditional_formatting.add(status_range,
        FormulaRule(formula=['ISNUMBER(SEARCH("綠",F14))'],
                    fill=FILL_GREEN, font=FONT_GREEN, stopIfTrue=False))
    rules_added += 1

    log.info(f"  頁 9 條件格式：{rules_added} 條（達成率 + 狀態欄）")
    return rules_added


# ============================================================
# (4) 案件資料庫容量警示
# ============================================================
def cf_page5_capacity(wb, log):
    """案件資料庫 A2 sub-banner — 容量動態警示

    覆寫 A2 為公式版（顯示即時 COUNTA），加 3 條 CF：
      ≥ 950 列 → 紅底（即將滿）
      ≥ 800 列 → 黃底（建議歸檔）
      其他    → 維持深藍 banner 底

    註：A2 cell 既有 fill 是 accent_dark（深藍底白字），CF 觸發才會覆蓋。
    """
    ws = wb["案件資料庫"]
    _clear_all_cf(ws)

    # 動態公式：顯示即時容量
    cap_formula = (
        '="📊 容量：" & IFERROR(COUNTA(Tbl案件[編號]),0) & " / 1000 列  ｜  '
        '" & IF(COUNTA(Tbl案件[編號])>=950,"🚨 容量即將滿，請立即歸檔",'
        'IF(COUNTA(Tbl案件[編號])>=800,"⚠ 已達 80% 容量警戒，建議規劃歸檔",'
        '"✓ 容量充裕")) & '
        '"  ｜  🔒 受保護（密碼：KD2026）"'
    )
    ws["A2"].value = cap_formula
    # 字色字型已由 phase2 設定，不重複

    # ★ 重要：Excel 條件格式公式 **不支援結構化引用**（如 Tbl案件[編號]）
    #   會被 Excel 開檔時整段移除 → 「已移除的功能 sheet6.xml CF」警示
    #   改用同 sheet 普通範圍引用：編號欄在 A15:A1014（phase 2 config）
    cap_count = 'COUNTA($A$15:$A$1014)'

    # CF：紅燈（≥950）
    ws.conditional_formatting.add("A2",
        FormulaRule(formula=[f'{cap_count}>=950'],
                    fill=PatternFill("solid", fgColor="FFB91C1C"),  # 深紅
                    font=Font(bold=True, color="FFFFFFFF"),
                    stopIfTrue=True))
    # CF：黃燈（≥800）
    ws.conditional_formatting.add("A2",
        FormulaRule(formula=[f'{cap_count}>=800'],
                    fill=PatternFill("solid", fgColor="FFB45309"),  # 深橘
                    font=Font(bold=True, color="FFFFFFFF"),
                    stopIfTrue=True))

    log.info("  頁 5 案件資料庫：A2 容量警示動態公式 + 2 條 CF（≥950 紅 / ≥800 黃）")
    return 2


# ============================================================
# (5) 工作表保護（密碼 KD2026）
# ============================================================
def _unlock_range(ws, cell_range: str):
    """解鎖儲存格範圍（cell.protection.locked=False）。"""
    for row in ws[cell_range]:
        for cell in row:
            cell.protection = Protection(locked=False, hidden=False)


def protect_page5_case_database(wb, log):
    """案件資料庫：鎖計算欄（R/S/T/U/V/X），解鎖輸入欄

    columns（phase2 config 對應）：
      A-Q (1-17)  input  → 解鎖
      R   (18)    calc   → 鎖
      S   (19)    calc   → 鎖
      T   (20)    calc   → 鎖
      U   (21)    calc   → 鎖
      V   (22)    calc   → 鎖
      W   (23)    input  → 解鎖（績效類別）
      X   (24)    calc   → 鎖
      Y   (25)    input  → 解鎖（自填案類）
    """
    ws = wb["案件資料庫"]
    cfg = load_config("phase2_data_tables")
    p5 = cfg["page5_case_database"]
    data_start = p5["data_start_row"]
    data_end = data_start + len(p5.get("sample_rows", [])) + p5.get("blank_rows", 200) - 1

    # 解鎖 input 欄
    input_cols = [i for i, c in enumerate(p5["columns"], start=1) if c["kind"] == "input"]
    unlocked = 0
    for col_idx in input_cols:
        col_letter = get_column_letter(col_idx)
        rng = f"{col_letter}{data_start}:{col_letter}{data_end}"
        _unlock_range(ws, rng)
        unlocked += data_end - data_start + 1

    # 開啟工作表保護
    ws.protection.password = PASSWORD
    ws.protection.sheet = True
    ws.protection.enable()
    # 允許使用者選擇 / 排序 / 自動篩選
    ws.protection.sort = False
    ws.protection.autoFilter = False
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False

    locked_cols = [get_column_letter(i) for i, c in enumerate(p5["columns"], start=1)
                   if c["kind"] == "calc"]
    log.info(f"  頁 5 案件資料庫：密碼 KD2026 ｜ 鎖計算欄 {','.join(locked_cols)} ｜ 解鎖 {unlocked} cells")


def protect_settings_sheet(wb, log):
    """設定表：標題列鎖住，使用者輸入區（員警/案類/閾值等）解鎖

    保守策略：整表上鎖，只解鎖使用者明確該改的儲存格範圍。
    若使用者要改更多區域，可用密碼 KD2026 暫時解保護。
    """
    ws = wb["設定表"]
    # 整表鎖住但解鎖以下範圍（A-G 欄各區的「值」儲存格）
    # 採保守做法：解鎖 B/C/D/E/F 欄 row 5 以下（A 欄為標籤、第 1-4 列為 header）
    unlock_ranges = [
        "B5:G500",   # 各分區的 value/說明等
    ]
    unlocked = 0
    for rng in unlock_ranges:
        _unlock_range(ws, rng)
        # 粗估解鎖 cell 數
        unlocked += 496 * 6

    ws.protection.password = PASSWORD
    ws.protection.sheet = True
    ws.protection.enable()
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False

    log.info(f"  設定表：密碼 KD2026 ｜ 解鎖 B5:G500（使用者輸入區）")


def protect_stats_sheets(wb, log):
    """統計頁：全鎖（只能看不能改）

    包含頁 2 / 2.5 / 3 / 4 / 6 / 8 / 9 / 10 / 11 / 12
    例外：頁 3 「下月行動方案」(手動輸入區) / 頁 4 上呈對象下拉 / 頁 11 員警下拉 — 解鎖
    """
    # ★ 統計頁全鎖
    locked_sheets = [
        "管制總覽", "全般刑案管制情形分析", "刑案管制",
        "毒品調驗人口管制", "交通績效管制",
        "績效統計", "跨期間趨勢分析", "歷史資料", "系統檢核"
    ]
    for name in locked_sheets:
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        ws.protection.password = PASSWORD
        ws.protection.sheet = True
        ws.protection.enable()
        ws.protection.selectLockedCells = False
        ws.protection.selectUnlockedCells = False
    log.info(f"  全鎖：{len(locked_sheets)} 頁（{','.join(locked_sheets)}）")

    # ★ 頁 3 長官 5 分鐘報告 — 解鎖「下月行動方案」區
    if "長官5分鐘報告" in wb.sheetnames:
        ws = wb["長官5分鐘報告"]
        # 依 phase4_page3 版型，行動方案區 row 24-27 是 A:H merged
        # merged 中的 cell.protection 只看 anchor (A 欄)，所以解鎖 A24:A27 即可
        # 但保險起見全 row 一起解鎖（merged 非 anchor cell 設了也不傷）
        _unlock_range(ws, "A24:H27")
        ws.protection.password = PASSWORD
        ws.protection.sheet = True
        ws.protection.enable()
        ws.protection.selectLockedCells = False
        ws.protection.selectUnlockedCells = False
        log.info("  頁 3：解鎖 A24:H27（下月行動方案手動輸入區，merged）")

    # ★ 頁 4 對上級機關上呈 — 解鎖上呈對象下拉
    if "對上級機關上呈" in wb.sheetnames:
        ws = wb["對上級機關上呈"]
        _unlock_range(ws, "B3:B3")
        ws.protection.password = PASSWORD
        ws.protection.sheet = True
        ws.protection.enable()
        ws.protection.selectLockedCells = False
        ws.protection.selectUnlockedCells = False
        log.info("  頁 4：解鎖 B3（上呈對象下拉）")

    # ★ 頁 11 員警個人績效卡 — 解鎖員警下拉
    if "員警個人績效卡" in wb.sheetnames:
        ws = wb["員警個人績效卡"]
        _unlock_range(ws, "B3:B3")
        ws.protection.password = PASSWORD
        ws.protection.sheet = True
        ws.protection.enable()
        ws.protection.selectLockedCells = False
        ws.protection.selectUnlockedCells = False
        log.info("  頁 11：解鎖 B3（員警下拉）")

    # ★ 頁 7 發展中、頁 9.5 交通取締明細 — 主表輸入，採類似頁 5 的策略
    # 頁 7：A14:I28 解鎖（9 input cols × 15 rows）
    if "發展中案件管制" in wb.sheetnames:
        ws = wb["發展中案件管制"]
        _unlock_range(ws, "A14:I28")   # 9 input cols
        ws.protection.password = PASSWORD
        ws.protection.sheet = True
        ws.protection.enable()
        ws.protection.selectLockedCells = False
        ws.protection.selectUnlockedCells = False
        log.info("  頁 7：解鎖 A14:I28（主表輸入區）")

    # 頁 9.5 交通取締明細：A15 以下全是 input
    if "交通取締明細" in wb.sheetnames:
        ws = wb["交通取締明細"]
        _unlock_range(ws, "A15:F200")
        ws.protection.password = PASSWORD
        ws.protection.sheet = True
        ws.protection.enable()
        ws.protection.selectLockedCells = False
        ws.protection.selectUnlockedCells = False
        log.info("  頁 9.5：解鎖 A15:F200（取締明細輸入區）")

    # 頁 14 歷史資料：A11 以下全是 input（每月抄寫）
    if "歷史資料" in wb.sheetnames:
        ws = wb["歷史資料"]
        # 歷史資料表的保護由前面 locked_sheets 已上；但要解鎖抄寫區
        _unlock_range(ws, "A11:N200")
        log.info("  頁 14：解鎖 A11:N200（每月抄寫區）")


# ============================================================
# (6) 列印區 + 頁面設定
# ============================================================
def set_print_layout(ws: Worksheet, print_range: str, *,
                     orientation="portrait", fit_height=1):
    """設置 A4 列印區 + fit-to-page。"""
    ws.print_area = print_range
    ws.page_setup.orientation = orientation
    ws.page_setup.paperSize = ws.PAPERSIZE_A4   # 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = fit_height
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    # 設邊界
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5
    ws.page_margins.header = 0.3
    ws.page_margins.footer = 0.3


def apply_print_areas(wb, log):
    """4 頁設 A4 直式列印區。

    print_range 依各頁版型尾列：
      頁 2.5 全般刑案管制情形分析 → A1:H32
      頁 3   長官 5 分鐘報告       → A1:H36
      頁 4   對上級機關上呈         → A1:H28（單對象 1 頁，多對象 fitToHeight=3）
      頁 11  員警個人績效卡         → A1:H50
    """
    sheets_to_setup = [
        ("全般刑案管制情形分析", "A1:H32", "portrait", 1),
        ("長官5分鐘報告",       "A1:H36", "portrait", 1),
        ("對上級機關上呈",       "A1:H28", "portrait", 3),
        ("員警個人績效卡",       "A1:H50", "portrait", 1),
    ]
    count = 0
    for name, rng, orient, fith in sheets_to_setup:
        if name not in wb.sheetnames:
            log.warning(f"  ⚠ sheet 不存在：{name}")
            continue
        ws = wb[name]
        set_print_layout(ws, rng, orientation=orient, fit_height=fith)
        log.info(f"  {name}：print_area={rng} ｜ A4 {orient} ｜ fit={fith}")
        count += 1

    return count


# ============================================================
# 主入口
# ============================================================
def build():
    log = get_logger("phase5_polish")
    log.info("===== Phase 5（條件格式 + 保護 + 列印區）開始 =====")

    wb = load_workbook(OUTPUT_PATH)

    # === 1. 條件格式 ===
    log.info("--- (1) 條件格式 ---")
    n6 = cf_page6(wb, log)
    n7 = cf_page7(wb, log)
    n9 = cf_page9(wb, log)
    n5 = cf_page5_capacity(wb, log)
    cf_total = n6 + n7 + n9 + n5
    log.info(f"  Phase 5 新增 CF 共 {cf_total} 條（不含頁 2 的 54 條既有）")

    # === 2. 工作表保護 ===
    log.info("--- (2) 工作表保護（密碼 KD2026）---")
    protect_page5_case_database(wb, log)
    protect_settings_sheet(wb, log)
    protect_stats_sheets(wb, log)

    # === 3. 列印區 + 頁面設定 ===
    log.info("--- (3) 列印區（A4 直式）---")
    print_count = apply_print_areas(wb, log)

    # === 儲存 ===
    bak = backup_existing()
    if bak:
        log.info(f"  舊版備份：{bak}")
    wb.save(OUTPUT_PATH)
    size_kb = OUTPUT_PATH.stat().st_size / 1024
    log.info(f"  儲存：{OUTPUT_PATH.name} ({size_kb:.1f} KB)")
    log.info(f"  CF 新增：{cf_total} 條 ｜ 列印區設定：{print_count} 頁")
    log.info("===== Phase 5 完成 =====\n")
    return OUTPUT_PATH


if __name__ == "__main__":
    build()
