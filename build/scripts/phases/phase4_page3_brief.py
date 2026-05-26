"""
phase4_page3_brief.py — Phase 4：頁 3 長官 5 分鐘報告（A 麥肯錫風 — Pyramid Principle）

版型（A4 直式，可獨立列印交付督導官）：
  Row 1-2  Banner（派出所名 ｜ 民國年月勤務狀況報告）
  Row 3    報告人：所長 ◯◯◯
  Row 4    spacer
  Row 5    Section title「★ 結論（一句話）」
  Row 6-7  結論文字（公式自動，依紅黃綠燈動態組句）
  Row 8    spacer
  Row 9    Section title「★ 三大重點」
  Row 10   重點 1：刑案績效（header）
  Row 11   bullet：本所破獲率
  Row 12   bullet：未破/已破未送/已移送 件數
  Row 13   bullet：本月變化
  Row 14   重點 2：毒品調驗（header）
  Row 15   bullet：毒調率
  Row 16   bullet：列管/已驗/未到驗
  Row 17   bullet：應強採候選
  Row 18   重點 3：交通取締（header）
  Row 19   bullet：達標項數
  Row 20   bullet：總達成率
  Row 21   bullet：酒駕/闖紅燈
  Row 22   spacer
  Row 23   Section title「★ 下月行動方案（手動輸入區）」
  Row 24-27 4 行手動輸入區（淡黃）— 所長決定
  Row 28   spacer
  Row 29   Section title「★ 本月 Top 3 員警」
  Row 30   header
  Row 31-33 Top 3 員警（INDEX from 頁 10 helper）
  Row 34   spacer
  Row 35   簽章區
  Row 36   頁尾

設計理念（plan）：
  - 純文字 + 數字嵌套，無圖表（長官不想看圖只想看話）
  - 結論 + 三大重點自動生成 → 所長省時
  - 下月行動方案手動 → 保留判斷權威（行動是判斷題，公式不能代生）
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook

from phases._base import backup_existing, OUTPUT_PATH, get_logger, unmerge_all_in_sheet
import styles as S


SHEET = "長官5分鐘報告"
ROC_YEAR = "(YEAR(今日)-1911)"
CASE_TBL = "Tbl案件"


def build_page3_brief(wb, log):
    log.info("--- 頁 3 長官 5 分鐘報告 ---")
    ws = wb[SHEET]

    cleared = unmerge_all_in_sheet(ws)
    log.info(f"  清空既有 merges：{cleared} 個")

    # === 欄寬（A4 直式，8 cols）===
    col_widths = {
        "A": 8, "B": 22, "C": 16, "D": 16, "E": 16, "F": 16, "G": 16, "H": 14,
    }
    for c, w in col_widths.items():
        ws.column_dimensions[c].width = w

    # === Row 1-2 Banner ===
    ws.merge_cells("A1:H1")
    banner_fml = (
        '="   📜  " & 派出所名稱 & "  ｜  中華民國 " & (YEAR(今日)-1911) '
        '& " 年 " & TEXT(MONTH(今日),"00") & " 月勤務狀況報告"'
    )
    S.set_cell(ws, "A1", banner_fml,
               font_key="banner", fill_key="banner",
               align_key="left", border_key="bottom_thick")
    ws.row_dimensions[1].height = S.ROW_HEIGHT["banner"]

    ws.merge_cells("A2:H2")
    sub_fml = (
        '="   " & 警察局名稱 & " · " & 分局名稱 & " · " & 派出所名稱 '
        '& "  ｜  製表 " & TEXT(今日,"e/mm/dd") & "  ｜  Pyramid Principle 5 分鐘簡報"'
    )
    S.set_cell(ws, "A2", sub_fml,
               font_key=S.font(11, color="white", italic=True),
               fill_key=S.fill("accent_dark"), align_key="left")
    ws.row_dimensions[2].height = S.ROW_HEIGHT["sub_banner"]

    # === Row 3 報告人 ===
    ws.merge_cells("A3:H3")
    S.set_cell(ws, "A3",
               '="   📌  報告人：" & 派出所名稱 & "所長  ｜  呈：分局督導官  ｜  本報告由系統自動聚合 + 所長行動方案"',
               font_key=S.font(12, bold=True, color="accent"),
               fill_key=S.fill("accent_light"),
               align_key="left", border_key="all_thin")
    ws.row_dimensions[3].height = 30

    # === Row 4 spacer ===
    ws.row_dimensions[4].height = 8

    # === Row 5 Section「★ 結論」===
    ws.merge_cells("A5:H5")
    S.set_cell(ws, "A5", "   ★  結論（一句話總結，公式自動）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[5].height = S.ROW_HEIGHT["header"]

    # Row 6-7 結論文字（merge 2 rows for height）
    # 依頁 2 警示摘要 紅/黃/綠燈數動態組句
    ws.merge_cells("A6:H7")
    conclusion_fml = (
        '=IF(管制總覽!A25>0,'
        '"🚨 本月狀態嚴峻：9 大 KPI 有 " & 管制總覽!A25 & " 項紅燈，需立即處置（建議優先處理未破刑案與毒調候選）",'
        'IF(管制總覽!G25>0,'
        '"⚠ 本月狀態尚可：" & 管制總覽!G25 & " 項黃燈需關注，建議排入週會檢討；其餘 KPI 達標",'
        '"✓ 本月狀態良好：9 大 KPI 全綠燈通過，建議持續維持"))'
    )
    S.set_cell(ws, "A6", conclusion_fml,
               font_key=S.font(15, bold=True, color="accent"),
               fill_key=S.fill("accent_light"),
               align_key="center", border_key="all_thin")
    ws.row_dimensions[6].height = 32
    ws.row_dimensions[7].height = 32

    # === Row 8 spacer ===
    ws.row_dimensions[8].height = 10

    # === Row 9 Section「★ 三大重點」===
    ws.merge_cells("A9:H9")
    S.set_cell(ws, "A9", "   ★  三大重點（按優先順序：刑案 → 毒調 → 交通）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[9].height = S.ROW_HEIGHT["header"]

    # 三大重點佈局
    # Row 10/14/18 = 大標
    # Row 11-13, 15-17, 19-21 = 三個 bullets each
    # 引用既有頁面公式以保持 SSOT
    own_solved = (
        f'COUNTIFS({CASE_TBL}[查獲管轄],"本轄",'
        f'{CASE_TBL}[是否破獲],"是",'
        f'{CASE_TBL}[歸屬年度],{ROC_YEAR})'
    )
    local_occur_in = (
        f'COUNTIFS({CASE_TBL}[發生管轄],"本轄",'
        f'{CASE_TBL}[歸屬年度],{ROC_YEAR},'
        f'{CASE_TBL}[計入發生數],"是")'
    )
    unsolved = (
        f'COUNTIFS({CASE_TBL}[案件狀況],"尚未偵破",'
        f'{CASE_TBL}[發生管轄],"本轄",'
        f'{CASE_TBL}[歸屬年度],{ROC_YEAR})'
    )
    not_sent = (
        f'COUNTIFS({CASE_TBL}[案件狀況],"已破獲未移送",'
        f'{CASE_TBL}[歸屬年度],{ROC_YEAR})'
    )
    sent = (
        f'COUNTIFS({CASE_TBL}[案件狀況],"已移送",'
        f'{CASE_TBL}[歸屬年度],{ROC_YEAR})'
    )

    points = [
        # (row, emoji, title, bullets list of formulas)
        (10, "🔍", "刑案績效", "accent", [
            f'="    ① 本所統計破獲率：" & TEXT(IFERROR({own_solved}/{local_occur_in},0),"0.0%") & "（年度累計）"',
            f'="    ② 案件狀況分布：尚未偵破 " & IFERROR({unsolved},0) & " 件　已破獲未移送 " & IFERROR({not_sent},0) & " 件　已移送 " & IFERROR({sent},0) & " 件"',
            f'="    ③ 重點案類：竊盜破獲率 " & TEXT(管制總覽!G9,"0.0%") & "　詐欺破獲率 " & TEXT(管制總覽!M9,"0.0%")',
        ]),
        (14, "💊", "毒品調驗", "warn", [
            '="    ① 毒調率：" & TEXT(IFERROR(SUM(Tbl毒調[是否到驗])/COUNTA(Tbl毒調[姓名]),0),"0.0%") & "（目標 " & TEXT(毒調率目標,"0%") & "）"',
            '="    ② 列管 " & COUNTA(Tbl毒調[姓名]) & " 人　完成到驗（含通緝/強採/在監）" & SUM(Tbl毒調[是否到驗]) & " 人"',
            '="    ③ 應強採候選：" & SUM(Tbl毒調[候選旗標]) & " 人未到驗待處理"',
        ]),
        (18, "🚦", "交通取締", "pass", [
            '="    ① 總達成率：" & TEXT(IFERROR(交通績效管制!E8,0),"0.0%")',
            '="    ② 達標項數：" & COUNTIF(交通績效管制!D14:D19,">="&0.8) & " ／ 6 項違規"',
            '="    ③ 主要違規：酒駕 " & IFERROR(COUNTIF(Tbl交通[違規項目],"*酒駕*"),0) & " 件　闖紅燈 " & IFERROR(COUNTIF(Tbl交通[違規項目],"*闖紅燈*"),0) & " 件"',
        ]),
    ]

    for row_start, emoji, title, color, bullets in points:
        # 大標
        ws.merge_cells(f"A{row_start}:H{row_start}")
        S.set_cell(ws, f"A{row_start}",
                   f"   {emoji}  {title}",
                   font_key=S.font(13, bold=True, color=color),
                   fill_key=S.fill(color + "_bg" if color != "accent" else "accent_light"),
                   align_key="left", border_key="all_thin")
        ws.row_dimensions[row_start].height = 28
        # 3 bullets
        for i, b in enumerate(bullets):
            r = row_start + 1 + i
            ws.merge_cells(f"A{r}:H{r}")
            S.set_cell(ws, f"A{r}", b,
                       font_key=S.font(12, color="text"),
                       fill_key="calc",
                       align_key="left", border_key="all_thin")
            ws.row_dimensions[r].height = 24

    # === Row 22 spacer ===
    ws.row_dimensions[22].height = 10

    # === Row 23 Section「★ 下月行動方案」===
    ws.merge_cells("A23:H23")
    S.set_cell(ws, "A23",
               "   ★  下月行動方案（手動輸入區 — 所長判斷裁示，公式不代生）",
               font_key="section_title", fill_key=S.fill("danger"),
               align_key="left")
    ws.row_dimensions[23].height = S.ROW_HEIGHT["header"]

    # Row 24-27 4 行手動輸入
    action_placeholders = [
        "  行動 1：（請所長手填，例：本月加強 XX 路段巡邏，目標降低詐欺 X 件）",
        "  行動 2：（請所長手填）",
        "  行動 3：（請所長手填）",
        "  備註：（必要時補充說明）",
    ]
    for i, txt in enumerate(action_placeholders):
        r = 24 + i
        ws.merge_cells(f"A{r}:H{r}")
        S.set_cell(ws, f"A{r}", txt,
                   font_key=S.font(11, color="muted", italic=True),
                   fill_key="input",
                   align_key="left", border_key="all_thin")
        ws.row_dimensions[r].height = 28

    # === Row 28 spacer ===
    ws.row_dimensions[28].height = 10

    # === Row 29 Section「★ 本月 Top 3 員警」===
    ws.merge_cells("A29:H29")
    S.set_cell(ws, "A29", "   🏆  本月 Top 3 員警（公式自動，來源頁 10 績效統計）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[29].height = S.ROW_HEIGHT["header"]

    # Row 30 header
    headers = ["排名", "員警姓名", "全般破獲", "竊盜", "詐欺", "毒品", "暴力犯罪", "其他"]
    for i, h in enumerate(headers, start=1):
        col = chr(ord("A") + i - 1)
        S.set_cell(ws, f"{col}30", h,
                   font_key="header", fill_key="header",
                   align_key="center", border_key="all_thin")
    ws.row_dimensions[30].height = S.ROW_HEIGHT["header"]

    # Row 31-33 Top 3 — 直接從頁 10 績效統計 row 19-21 引用
    for rank in range(1, 4):
        r = 30 + rank
        page10_row = 18 + rank   # 頁 10 Top 1 在 row 19
        # 8 欄全部從頁 10 引用（排名/員警/全般破獲/竊盜/詐欺/毒品/暴力犯罪/其他）
        sources = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
        for i, src_col in enumerate(sources):
            col = chr(ord("A") + i)
            fml = f'=IFERROR(績效統計!{src_col}{page10_row},"")'
            S.set_cell(ws, f"{col}{r}", fml,
                       font_key="body_bold" if i < 3 else "body",
                       fill_key="calc",
                       align_key="left" if i == 1 else "center",
                       border_key="all_thin",
                       number_format="integer" if i >= 2 else "general")
        ws.row_dimensions[r].height = S.ROW_HEIGHT["default"]

    # === Row 34 spacer ===
    ws.row_dimensions[34].height = 12

    # === Row 35 簽章區 ===
    ws.merge_cells("A35:D35")
    S.set_cell(ws, "A35", "  所長：__________________  簽章",
               font_key=S.font(11, color="text"),
               fill_key=None, align_key="left",
               border_key="bottom_medium")
    ws.merge_cells("E35:H35")
    S.set_cell(ws, "E35", "  督導官：__________________  簽章",
               font_key=S.font(11, color="text"),
               fill_key=None, align_key="left",
               border_key="bottom_medium")
    ws.row_dimensions[35].height = 36

    # === Row 36 頁尾 ===
    ws.merge_cells("A36:H36")
    footer = (
        '="© KKEVIN-LIN-2026-V2.0  ｜  Pyramid Principle 5 分鐘簡報  ｜  A4 直式可獨立交付督導官  ｜  '
        '製表："&TEXT(今日,"e/mm/dd")'
    )
    S.set_cell(ws, "A36", footer,
               font_key=S.font(9, color="muted", italic=True),
               fill_key="calc", align_key="center")
    ws.row_dimensions[36].height = 22

    # 公式計數
    fcount = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                fcount += 1
    log.info(f"  Banner / 結論 / 三大重點 / 行動方案 / Top 3")
    log.info(f"  公式總數：{fcount}")
    return fcount


def build():
    log = get_logger("phase4_p3")
    log.info("===== Phase 4 — 頁 3 長官 5 分鐘報告 開始 =====")

    wb = load_workbook(OUTPUT_PATH)
    fcount = build_page3_brief(wb, log)

    bak = backup_existing()
    if bak:
        log.info(f"  舊版備份：{bak}")
    wb.save(OUTPUT_PATH)
    size_kb = OUTPUT_PATH.stat().st_size / 1024
    log.info(f"  儲存：{OUTPUT_PATH.name} ({size_kb:.1f} KB)")
    log.info(f"  公式：{fcount}")
    log.info("===== Phase 4 — 頁 3 完成 =====\n")
    return OUTPUT_PATH


if __name__ == "__main__":
    build()
