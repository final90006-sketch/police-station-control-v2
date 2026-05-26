"""
phase3_page2_5_crime_analysis.py — Phase 3：頁 2.5 全般刑案管制情形分析（A+B 混搭升級）

版型（A4 直式一頁，可獨立列印交付）：
  Row 1-2   深藍章戳頁首（英文機關名 + 主標 + 統計區間）
  Row 3     金色裝飾線 / 同期對照註腳
  Row 4     spacer
  Row 5     Section「壹 全般刑案總覽」
  Row 6-7   大字 KPI 3 卡（本所破獲率 / 發生件數 / 破獲件數）
  Row 8-9   比較長條（本所 vs 分局 + 紅線 100% 基準）
  Row 10    差異原因卡（自動算 + 手動填補充）
  Row 11    spacer
  Row 12    Section「貳 重點案類（竊盜 / 詐欺）」
  Row 13-22 雙環形圖 (DoughnutChart) + 數字
  Row 23    spacer
  Row 24    Section「參 未破案件與其他管制」
  Row 25-29 未破刑案三欄表（手動輸入）
  Row 30    婦幼／協尋狀態
  Row 31    spacer
  Row 32    頁尾（報告人 + 破獲率基準）

業務鐵則：
  - 本所統計 = 查獲管轄 = "本轄"（含拘提他轄）
  - 分局統計 = 發生管轄 = "本轄"（純看轄內發生）
  - 差異 = 拘提他轄破 + 一般已破未送
  - 婦幼/協尋：設定表 F 區手動 4 格輸入

環形圖：openpyxl DoughnutChart
  竊盜：已破 / 未破 → 深藍系
  詐欺：已破 / 未破 → 翡翠綠系
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.chart import DoughnutChart, Reference
from openpyxl.chart.layout import Layout, ManualLayout

from phases._base import backup_existing, OUTPUT_PATH, get_logger, unmerge_all_in_sheet
import styles as S


SHEET = "全般刑案管制情形分析"
ROC_YEAR = "(YEAR(今日)-1911)"
CASE_TBL = "Tbl案件"


def build_page2_5(wb, log):
    log.info("--- 頁 2.5 全般刑案管制情形分析 ---")
    ws = wb[SHEET]

    cleared = unmerge_all_in_sheet(ws)
    log.info(f"  清空既有 merges：{cleared} 個")

    # === 欄寬（v2.1 全面加大）===
    col_widths = {
        "A": 14, "B": 20, "C": 20, "D": 20, "E": 20, "F": 20, "G": 20, "H": 16,
    }
    for c, w in col_widths.items():
        ws.column_dimensions[c].width = w

    # === Row 1-2 深藍章戳頁首 ===
    ws.merge_cells("A1:H1")
    S.set_cell(ws, "A1",
               "   ⚖  全般刑案管制情形分析",
               font_key=S.font(20, bold=True, color="white"),
               fill_key=S.fill("accent_dark"),
               align_key="center", border_key="bottom_thick")
    ws.row_dimensions[1].height = 56

    ws.merge_cells("A2:H2")
    sub_formula = (
        '=派出所英文&" · "&警察局名稱&" · "&分局名稱&" · "&派出所名稱'
        '&"  ｜  REPORT DATE: "&TEXT(今日,"yyyy/mm/dd")'
    )
    S.set_cell(ws, "A2", sub_formula,
               font_key=S.font(11, italic=True, color="white"),
               fill_key=S.fill("accent_dark"), align_key="center")
    ws.row_dimensions[2].height = 26

    # === Row 3 統計區間 + 同期對照 ===
    ws.merge_cells("A3:H3")
    range_formula = (
        '="統計區間：中華民國 "&(YEAR(今日)-1911)&" 年 01/01 ~ "&TEXT(今日,"mm/dd")'
        '&"   ｜  同期對照：中華民國 "&(YEAR(今日)-1912)&" 年同期間"'
    )
    S.set_cell(ws, "A3", range_formula,
               font_key=S.font(11, bold=True, color="accent"),
               fill_key=S.fill("warn_bg"),     # 金色底
               align_key="center",
               border_key="bottom_medium")
    ws.row_dimensions[3].height = 26

    # === Row 4 spacer ===
    ws.row_dimensions[4].height = 10

    # === Row 5 Section「壹 全般刑案總覽」===
    ws.merge_cells("A5:H5")
    S.set_cell(ws, "A5", "   壹  全般刑案總覽",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[5].height = S.ROW_HEIGHT["header"]

    # === Row 6-7 大字 KPI 3 卡 ===
    # 本所統計：查獲管轄 = 本轄 (含拘提他轄)
    own_solved = (
        f'COUNTIFS({CASE_TBL}[查獲管轄],"本轄",'
        f'{CASE_TBL}[是否破獲],"是",'
        f'{CASE_TBL}[歸屬年度],{ROC_YEAR})'
    )
    local_occur = (
        f'COUNTIFS({CASE_TBL}[發生管轄],"本轄",'
        f'{CASE_TBL}[歸屬年度],{ROC_YEAR},'
        f'{CASE_TBL}[計入發生數],"是")'
    )
    # 分局統計：發生管轄 = 本轄 + 是否破獲 = 是
    div_solved = (
        f'COUNTIFS({CASE_TBL}[發生管轄],"本轄",'
        f'{CASE_TBL}[是否破獲],"是",'
        f'{CASE_TBL}[歸屬年度],{ROC_YEAR})'
    )

    kpi_cards = [
        (("A6:C6", "A7:C7"),
            "🎯 本所統計·破獲率", "accent", "accent_light",
            f'=IFERROR({own_solved}/{local_occur},0)', "percent_one",
            "查獲管轄=本轄 ÷ 發生管轄=本轄"),
        (("D6:E6", "D7:E7"),
            "📊 發生件數", "warn", "warn_bg",
            f'=IFERROR({local_occur},0)', "integer",
            "發生管轄=本轄"),
        (("F6:H6", "F7:H7"),
            "✓ 破獲件數（本所）", "pass", "pass_bg",
            f'=IFERROR({own_solved},0)', "integer",
            "查獲管轄=本轄"),
    ]
    for (lab_rng, val_rng), label, fg, bg, fml, fmt, sub in kpi_cards:
        ws.merge_cells(lab_rng)
        first = lab_rng.split(":")[0]
        S.set_cell(ws, first, f"  {label}\n  {sub}",
                   font_key=S.font(11, bold=True, color=fg),
                   fill_key=S.fill(bg),
                   align_key="left", border_key="top_bottom")
        ws.merge_cells(val_rng)
        S.set_cell(ws, val_rng.split(":")[0], fml,
                   font_key=S.font(30, bold=True, color=fg),
                   fill_key=S.fill(bg),
                   align_key="center", number_format=fmt,
                   border_key="bottom_thick")
    ws.row_dimensions[6].height = 36
    ws.row_dimensions[7].height = 56

    # === Row 8-9 比較長條（本所 vs 分局）===
    ws.merge_cells("A8:B8")
    S.set_cell(ws, "A8", "  本所破獲率：",
               font_key="body_bold", fill_key="calc",
               align_key="right", border_key="all_thin")
    ws.merge_cells("C8:F8")
    S.set_cell(ws, "C8",
               f'=REPT("█",MIN(20,ROUND(IFERROR({own_solved}/{local_occur},0)*20,0)))'
               f'&" "&TEXT(IFERROR({own_solved}/{local_occur},0),"0.0%")',
               font_key=S.font(13, bold=True, color="accent",
                               family="Consolas"),
               fill_key="calc",
               align_key="left", border_key="all_thin")
    ws.merge_cells("G8:H8")
    S.set_cell(ws, "G8",
               f'=IF(IFERROR({own_solved}/{local_occur},0)>=1,"✓ 達標","● 待達標")',
               font_key=S.font(11, bold=True, color="accent"),
               fill_key="calc",
               align_key="center", border_key="all_thin")

    ws.merge_cells("A9:B9")
    S.set_cell(ws, "A9", "  分局破獲率：",
               font_key="body_bold", fill_key="calc",
               align_key="right", border_key="all_thin")
    ws.merge_cells("C9:F9")
    S.set_cell(ws, "C9",
               f'=REPT("█",MIN(20,ROUND(IFERROR({div_solved}/{local_occur},0)*20,0)))'
               f'&" "&TEXT(IFERROR({div_solved}/{local_occur},0),"0.0%")',
               font_key=S.font(13, bold=True, color="warn",
                               family="Consolas"),
               fill_key="calc",
               align_key="left", border_key="all_thin")
    ws.merge_cells("G9:H9")
    S.set_cell(ws, "G9", "  100% 基準",
               font_key=S.font(10, color="muted", italic=True),
               fill_key="calc",
               align_key="center", border_key="all_thin")
    ws.row_dimensions[8].height = 26
    ws.row_dimensions[9].height = 26

    # === Row 10 差異原因卡 ===
    # 差異 = own_solved - div_solved (拘提他轄 + 已破未送)
    # 自動算：拘提他轄破 + 已破未送
    own_div_diff = f'IFERROR({own_solved}-{div_solved},0)'
    bring_other = (
        f'COUNTIFS({CASE_TBL}[發生管轄],"他轄",'
        f'{CASE_TBL}[查獲管轄],"本轄",'
        f'{CASE_TBL}[歸屬年度],{ROC_YEAR})'
    )
    not_sent = (
        f'COUNTIFS({CASE_TBL}[發生管轄],"本轄",'
        f'{CASE_TBL}[案件狀況],"已破獲未移送",'
        f'{CASE_TBL}[歸屬年度],{ROC_YEAR})'
    )
    ws.merge_cells("A10:H10")
    diff_fml = (
        f'="差異分析：本所 vs 分局 差 "&{own_div_diff}&" 件 = '
        f'拘提掃碼車手破他轄 "&IFERROR({bring_other},0)&" 件 + '
        f'本轄已破未移送 "&IFERROR({not_sent},0)&" 件"'
    )
    S.set_cell(ws, "A10", diff_fml,
               font_key=S.font(11, color="accent", italic=True),
               fill_key=S.fill("accent_light"),
               align_key="center", border_key="all_thin")
    ws.row_dimensions[10].height = 28

    # === Row 11 spacer ===
    ws.row_dimensions[11].height = 10

    # === Row 12 Section「貳 重點案類」===
    ws.merge_cells("A12:H12")
    S.set_cell(ws, "A12", "   貳  重點案類（竊盜 / 詐欺）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[12].height = S.ROW_HEIGHT["header"]

    # === Row 13-22 雙環形圖 ===
    # 左半 A-D 竊盜環形圖；右半 E-H 詐欺環形圖
    # KPI 數字在 row 13-14；DoughnutChart anchor 在 row 15

    # 竊盜 KPI (M8 升級：[案類分類] 精準比對 + 發生加 [計入發生數])
    theft_occur = f'COUNTIFS({CASE_TBL}[案類分類],"竊盜",{CASE_TBL}[發生管轄],"本轄",{CASE_TBL}[歸屬年度],{ROC_YEAR},{CASE_TBL}[計入發生數],"是")'
    theft_solved = f'COUNTIFS({CASE_TBL}[案類分類],"竊盜",{CASE_TBL}[查獲管轄],"本轄",{CASE_TBL}[是否破獲],"是",{CASE_TBL}[歸屬年度],{ROC_YEAR})'
    # 詐欺 KPI
    fraud_occur = f'COUNTIFS({CASE_TBL}[案類分類],"詐欺",{CASE_TBL}[發生管轄],"本轄",{CASE_TBL}[歸屬年度],{ROC_YEAR},{CASE_TBL}[計入發生數],"是")'
    fraud_solved = f'COUNTIFS({CASE_TBL}[案類分類],"詐欺",{CASE_TBL}[查獲管轄],"本轄",{CASE_TBL}[是否破獲],"是",{CASE_TBL}[歸屬年度],{ROC_YEAR})'

    # Row 13 標籤
    ws.merge_cells("A13:D13")
    S.set_cell(ws, "A13",
               f'="🔵 竊盜：發生 "&IFERROR({theft_occur},0)&" / 破獲 "'
               f'&IFERROR({theft_solved},0)&" / 破獲率 "'
               f'&TEXT(IFERROR({theft_solved}/{theft_occur},0),"0.0%")',
               font_key=S.font(13, bold=True, color="accent"),
               fill_key=S.fill("accent_light"),
               align_key="center", border_key="all_thin")
    ws.merge_cells("E13:H13")
    S.set_cell(ws, "E13",
               f'="🟢 詐欺：發生 "&IFERROR({fraud_occur},0)&" / 破獲 "'
               f'&IFERROR({fraud_solved},0)&" / 破獲率 "'
               f'&TEXT(IFERROR({fraud_solved}/{fraud_occur},0),"0.0%")',
               font_key=S.font(13, bold=True, color="pass"),
               fill_key=S.fill("pass_bg"),
               align_key="center", border_key="all_thin")
    ws.row_dimensions[13].height = 28

    # Row 14-22 環形圖 chart anchor 區域
    # 預留 9 列高度 for 圖表
    for r in range(14, 23):
        ws.row_dimensions[r].height = 22

    # === Helper data for DoughnutChart（隱藏在 row 100+）===
    # 竊盜：cols Q (label) / R (value)
    S.set_cell(ws, "Q99", "[helper] 竊盜環形圖",
               font_key=S.font(9, color="muted"))
    S.set_cell(ws, "Q100", "已破",
               font_key=S.font(9, color="muted"))
    S.set_cell(ws, "R100", f'=IFERROR({theft_solved},0)',
               font_key=S.font(9, color="muted"))
    S.set_cell(ws, "Q101", "尚未偵破",
               font_key=S.font(9, color="muted"))
    S.set_cell(ws, "R101",
               f'=MAX(0,IFERROR({theft_occur}-{theft_solved},0))',
               font_key=S.font(9, color="muted"))
    # 詐欺
    S.set_cell(ws, "S99", "[helper] 詐欺環形圖",
               font_key=S.font(9, color="muted"))
    S.set_cell(ws, "S100", "已破",
               font_key=S.font(9, color="muted"))
    S.set_cell(ws, "T100", f'=IFERROR({fraud_solved},0)',
               font_key=S.font(9, color="muted"))
    S.set_cell(ws, "S101", "尚未偵破",
               font_key=S.font(9, color="muted"))
    S.set_cell(ws, "T101",
               f'=MAX(0,IFERROR({fraud_occur}-{fraud_solved},0))',
               font_key=S.font(9, color="muted"))

    # 隱藏 helper rows 99-105 + cols Q-T
    for r in range(99, 106):
        ws.row_dimensions[r].hidden = True
    for col in ["Q", "R", "S", "T"]:
        ws.column_dimensions[col].hidden = True

    # === DoughnutChart 1 — 竊盜 (anchor A14) ===
    chart_theft = DoughnutChart()
    chart_theft.title = "竊盜破獲率"
    chart_theft.style = 26
    data = Reference(ws, min_col=18, min_row=100, max_col=18, max_row=101)  # R100:R101
    labels = Reference(ws, min_col=17, min_row=100, max_col=17, max_row=101)  # Q100:Q101
    chart_theft.add_data(data, titles_from_data=False)
    chart_theft.set_categories(labels)
    chart_theft.height = 5     # cm
    chart_theft.width = 8
    ws.add_chart(chart_theft, "A14")

    # === DoughnutChart 2 — 詐欺 (anchor E14) ===
    chart_fraud = DoughnutChart()
    chart_fraud.title = "詐欺破獲率"
    chart_fraud.style = 28
    data = Reference(ws, min_col=20, min_row=100, max_col=20, max_row=101)  # T100:T101
    labels = Reference(ws, min_col=19, min_row=100, max_col=19, max_row=101)  # S100:S101
    chart_fraud.add_data(data, titles_from_data=False)
    chart_fraud.set_categories(labels)
    chart_fraud.height = 5
    chart_fraud.width = 8
    ws.add_chart(chart_fraud, "E14")

    # === Row 23 spacer ===
    ws.row_dimensions[23].height = 10

    # === Row 24 Section「參 未破案件 + 其他管制」===
    ws.merge_cells("A24:H24")
    S.set_cell(ws, "A24", "   參  未破案件與其他管制",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[24].height = S.ROW_HEIGHT["header"]

    # === Row 25-29 未破刑案三欄表（手動輸入，公式總計）===
    # Header row 25
    ws.merge_cells("A25:C25")
    S.set_cell(ws, "A25", "  分類原因（手動填）",
               font_key="header", fill_key="header",
               align_key="left", border_key="all_thin")
    ws.merge_cells("D25:D25")
    S.set_cell(ws, "D25", "件數（手動）",
               font_key="header", fill_key="header",
               align_key="center", border_key="all_thin")
    ws.merge_cells("E25:H25")
    S.set_cell(ws, "E25", "  補充說明（手動填）",
               font_key="header", fill_key="header",
               align_key="left", border_key="all_thin")
    ws.row_dimensions[25].height = S.ROW_HEIGHT["header"]

    # 預留 3 列分類（淡黃輸入）
    sample_categories = [
        ("他轄發生暫掛", 0, ""),
        ("網路詐欺跨境", 0, ""),
        ("無監視器卷宗", 0, ""),
    ]
    for r_idx, (cat, count, note) in enumerate(sample_categories, start=26):
        ws.merge_cells(f"A{r_idx}:C{r_idx}")
        S.set_cell(ws, f"A{r_idx}", cat,
                   font_key="body", fill_key="input",
                   align_key="left", border_key="all_thin")
        S.set_cell(ws, f"D{r_idx}", count,
                   font_key="body", fill_key="input",
                   align_key="center", border_key="all_thin",
                   number_format="integer")
        ws.merge_cells(f"E{r_idx}:H{r_idx}")
        S.set_cell(ws, f"E{r_idx}", note,
                   font_key="body", fill_key="input",
                   align_key="left", border_key="all_thin")
        ws.row_dimensions[r_idx].height = S.ROW_HEIGHT["default"]

    # 總計列 row 29
    unsolved_local = (
        f'COUNTIFS({CASE_TBL}[發生管轄],"本轄",'
        f'{CASE_TBL}[案件狀況],"尚未偵破",'
        f'{CASE_TBL}[歸屬年度],{ROC_YEAR},'
        f'{CASE_TBL}[計入發生數],"是")'
    )
    ws.merge_cells("A29:C29")
    S.set_cell(ws, "A29", "  ─ 總計（手動 vs 自動）─",
               font_key=S.font(12, bold=True, color="white"),
               fill_key="banner", align_key="left",
               border_key="all_thin")
    S.set_cell(ws, "D29",
               f'=SUM(D26:D28)',
               font_key=S.font(12, bold=True, color="white"),
               fill_key="banner",
               align_key="center", border_key="all_thin",
               number_format="integer")
    ws.merge_cells("E29:H29")
    S.set_cell(ws, "E29",
               f'="  自動算未破："&IFERROR({unsolved_local},0)&" 件  ｜  '
               f'若不一致請補分類"',
               font_key=S.font(11, color="white", italic=True),
               fill_key="banner",
               align_key="left", border_key="all_thin")
    ws.row_dimensions[29].height = 26

    # === Row 30 婦幼／協尋狀態 ===
    ws.merge_cells("A30:D30")
    woman_fml = (
        '="🛡 婦幼案件："&IF(婦幼件數=0,"目前尚無管制案件",婦幼件數&" 件 列管中")'
    )
    S.set_cell(ws, "A30", woman_fml,
               font_key=S.font(11, bold=True, color="warn"),
               fill_key=S.fill("warn_bg"),
               align_key="left", border_key="all_thin")
    ws.merge_cells("E30:H30")
    missing_fml = (
        '="🔍 協尋人口："&IF(協尋件數=0,"目前尚無管制案件",協尋件數&" 件 列管中")'
    )
    S.set_cell(ws, "E30", missing_fml,
               font_key=S.font(11, bold=True, color="warn"),
               fill_key=S.fill("warn_bg"),
               align_key="left", border_key="all_thin")
    ws.row_dimensions[30].height = 28

    # === Row 31 spacer ===
    ws.row_dimensions[31].height = 10

    # === Row 32 頁尾 ===
    ws.merge_cells("A32:H32")
    footer = (
        '="報告人："&派出所名稱&"所長  ｜  破獲率基準：100%  ｜  '
        '本所統計：查獲管轄=本轄；分局統計：發生管轄=本轄  ｜  '
        '製表："&TEXT(今日,"e/mm/dd")'
    )
    S.set_cell(ws, "A32", footer,
               font_key=S.font(10, color="muted", italic=True),
               fill_key=S.fill("accent_dark"),
               align_key="center", border_key="top_bottom")
    # 重設文字顏色為白色（因為 accent_dark 底）
    ws["A32"].font = S.font(10, italic=True, color="white")
    ws.row_dimensions[32].height = 28

    # 公式計數
    fcount = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                fcount += 1
    log.info(f"  KPI 3 卡 + 比較長條 + 差異卡 + 雙環形圖")
    log.info(f"  公式總數：{fcount}")
    log.info(f"  Charts：2 個 DoughnutChart（竊盜 + 詐欺）")
    log.info(f"  凍結窗格：A8")
    return fcount


def build():
    log = get_logger("phase3_p2_5")
    log.info("===== Phase 3 — 頁 2.5 全般刑案分析 開始 =====")

    wb = load_workbook(OUTPUT_PATH)
    fcount = build_page2_5(wb, log)

    bak = backup_existing()
    if bak:
        log.info(f"  舊版備份：{bak}")
    wb.save(OUTPUT_PATH)
    size_kb = OUTPUT_PATH.stat().st_size / 1024
    log.info(f"  儲存：{OUTPUT_PATH.name} ({size_kb:.1f} KB)")
    log.info(f"  公式：{fcount}")
    log.info("===== Phase 3 — 頁 2.5 完成 =====\n")
    return OUTPUT_PATH


if __name__ == "__main__":
    build()
