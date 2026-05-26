"""
phase4_page4_upreport.py — Phase 4：頁 4 對上級機關上呈（B 公文風）

頂部下拉切換對象（對分局 / 對警察局 / 對警政署），公文抬頭 + 附件清單動態切換。
KPI 主表沿用 SSOT，三種對象共用同一組數字。

版型：
  Row 1-2  Banner
  Row 3    上呈對象下拉 + 動態抬頭
  Row 4    spacer
  Row 5    Section「★ 公文抬頭」
  Row 6-8  公文標題 / 主旨 / 受文者（CHOOSE 切三組）
  Row 9    spacer
  Row 10   Section「★ 主要 KPI（年度累計）」
  Row 11   KPI 表 header
  Row 12-19 8 KPI 列
  Row 20   spacer
  Row 21   Section「★ 必附附件」
  Row 22-25 附件清單（CHOOSE 切三組）
  Row 26   spacer
  Row 27   簽章區
  Row 28   頁尾
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

from phases._base import backup_existing, OUTPUT_PATH, get_logger, unmerge_all_in_sheet
import styles as S


SHEET = "對上級機關上呈"
TARGET_CELL = "B3"
ROC_YEAR = "(YEAR(今日)-1911)"
CASE_TBL = "Tbl案件"

TARGETS = ["對分局", "對警察局", "對警政署"]


def build_page4_upreport(wb, log):
    log.info("--- 頁 4 對上級機關上呈 ---")
    ws = wb[SHEET]

    cleared = unmerge_all_in_sheet(ws)
    log.info(f"  清空既有 merges：{cleared} 個")

    # === 欄寬 ===
    col_widths = {
        "A": 10, "B": 20, "C": 20, "D": 16, "E": 16, "F": 16, "G": 16, "H": 14,
    }
    for c, w in col_widths.items():
        ws.column_dimensions[c].width = w

    # === Row 1-2 Banner ===
    ws.merge_cells("A1:H1")
    S.set_cell(ws, "A1", "   📋  對上級機關上呈（公文格式 — 三層動態切換）",
               font_key="banner", fill_key="banner",
               align_key="left", border_key="bottom_thick")
    ws.row_dimensions[1].height = S.ROW_HEIGHT["banner"]

    ws.merge_cells("A2:H2")
    sub_fml = (
        '="   " & 警察局名稱 & " · " & 分局名稱 & " · " & 派出所名稱 '
        '& "  ｜  中華民國 " & (YEAR(今日)-1911) & " 年 " '
        '& TEXT(MONTH(今日),"00") & " 月  ｜  公文體例：三線表 / 官方藍 / 章戳"'
    )
    S.set_cell(ws, "A2", sub_fml,
               font_key=S.font(11, color="white", italic=True),
               fill_key=S.fill("accent_dark"), align_key="left")
    ws.row_dimensions[2].height = S.ROW_HEIGHT["sub_banner"]

    # === Row 3 下拉 + 動態抬頭 ===
    S.set_cell(ws, "A3", "  上呈對象：",
               font_key=S.font(13, bold=True, color="accent"),
               fill_key="accent_light",
               align_key="right", border_key="all_thin")
    S.set_cell(ws, TARGET_CELL, "對分局",   # 預設
               font_key=S.font(14, bold=True, color="white"),
               fill_key="banner",
               align_key="center", border_key="all_thin")
    # DataValidation
    target_list = "\"" + ",".join(TARGETS) + "\""
    dv = DataValidation(type="list", formula1=target_list, allow_blank=False)
    dv.add(TARGET_CELL)
    dv.prompt = "請選擇對象：分局／警察局／警政署"
    dv.promptTitle = "上呈對象切換"
    ws.add_data_validation(dv)

    # C3:H3 動態抬頭
    ws.merge_cells("C3:H3")
    head_fml = (
        '=CHOOSE(MATCH(' + TARGET_CELL + ',{"對分局","對警察局","對警政署"},0),'
        '"📂 詳細五領域 KPI ｜ 列印 A4 ×2-3 頁",'
        '"📂 月報摘要 + 全市比較 ｜ 列印 A4 ×1-2 頁",'
        '"📂 全國標準格式 ｜ 列印 A4 ×1 頁")'
    )
    S.set_cell(ws, "C3", head_fml,
               font_key=S.font(12, bold=True, color="accent"),
               fill_key="accent_light",
               align_key="left", border_key="all_thin")
    ws.row_dimensions[3].height = 32

    # === Row 4 spacer ===
    ws.row_dimensions[4].height = 8

    # === Row 5 Section「★ 公文抬頭」===
    ws.merge_cells("A5:H5")
    S.set_cell(ws, "A5", "   ★  公文抬頭（依上呈對象動態切換）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[5].height = S.ROW_HEIGHT["header"]

    # Row 6 公文標題
    ws.merge_cells("A6:H6")
    title_fml = (
        '=CHOOSE(MATCH(' + TARGET_CELL + ',{"對分局","對警察局","對警政署"},0),'
        '派出所名稱 & " 中華民國 " & (YEAR(今日)-1911) & " 年 " & TEXT(MONTH(今日),"00") & " 月勤務報告（呈分局）",'
        '派出所名稱 & " 中華民國 " & (YEAR(今日)-1911) & " 年 " & TEXT(MONTH(今日),"00") & " 月績效月報（呈警察局）",'
        '派出所名稱 & " 中華民國 " & (YEAR(今日)-1911) & " 年 " & TEXT(MONTH(今日),"00") & " 月標準統計（呈警政署）")'
    )
    S.set_cell(ws, "A6", title_fml,
               font_key=S.font(16, bold=True, color="accent"),
               fill_key=S.fill("accent_light"),
               align_key="center", border_key="all_thin")
    ws.row_dimensions[6].height = 36

    # Row 7 受文者
    ws.merge_cells("A7:B7")
    S.set_cell(ws, "A7", "  受文者：",
               font_key="body_bold", fill_key="calc",
               align_key="right", border_key="all_thin")
    ws.merge_cells("C7:H7")
    recv_fml = (
        '=CHOOSE(MATCH(' + TARGET_CELL + ',{"對分局","對警察局","對警政署"},0),'
        '分局名稱 & " 督導官",'
        '警察局名稱 & " 刑事警察大隊",'
        '"內政部警政署刑事警察局")'
    )
    S.set_cell(ws, "C7", recv_fml,
               font_key="body_bold", fill_key="calc",
               align_key="left", border_key="all_thin")
    ws.row_dimensions[7].height = 24

    # Row 8 主旨
    ws.merge_cells("A8:B8")
    S.set_cell(ws, "A8", "  主旨：",
               font_key="body_bold", fill_key="calc",
               align_key="right", border_key="all_thin")
    ws.merge_cells("C8:H8")
    subject_fml = (
        '=CHOOSE(MATCH(' + TARGET_CELL + ',{"對分局","對警察局","對警政署"},0),'
        '"呈報本所本月勤務狀況、案件管制、毒調人口、交通取締五大領域 KPI",'
        '"呈報本所本月績效摘要，並與全市平均比較",'
        '"呈報本所本月全國統一格式刑案統計")'
    )
    S.set_cell(ws, "C8", subject_fml,
               font_key="body", fill_key="calc",
               align_key="left", border_key="all_thin")
    ws.row_dimensions[8].height = 24

    # === Row 9 spacer ===
    ws.row_dimensions[9].height = 10

    # === Row 10 Section「★ 主要 KPI」===
    ws.merge_cells("A10:H10")
    S.set_cell(ws, "A10",
               "   ★  主要 KPI（年度累計，SSOT 自動聚合）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[10].height = S.ROW_HEIGHT["header"]

    # Row 11 KPI header
    headers = ["編號", "KPI 項目", "數值", "—", "目標", "—", "達標狀態", "—"]
    for i, h in enumerate(headers, start=1):
        col = chr(ord("A") + i - 1)
        if h != "—":
            S.set_cell(ws, f"{col}11", h,
                       font_key="header", fill_key="header",
                       align_key="center", border_key="all_thin")
    # merge 數值/目標/狀態 each 2 cols
    ws.merge_cells("C11:D11")
    S.set_cell(ws, "C11", "數值",
               font_key="header", fill_key="header",
               align_key="center", border_key="all_thin")
    ws.merge_cells("E11:F11")
    S.set_cell(ws, "E11", "目標",
               font_key="header", fill_key="header",
               align_key="center", border_key="all_thin")
    ws.merge_cells("G11:H11")
    S.set_cell(ws, "G11", "達標狀態",
               font_key="header", fill_key="header",
               align_key="center", border_key="all_thin")
    ws.row_dimensions[11].height = S.ROW_HEIGHT["header"]

    # Row 12-19: 8 KPI 列
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
    overall_rate = f'IFERROR({own_solved}/{local_occur_in},0)'

    kpi_list = [
        # (no, name, value_formula, fmt, target, target_fmt, threshold)
        ("01", "全般破獲率", overall_rate, "percent_one", "0.85", "percent_one", ">=0.85"),
        ("02", "竊盜破獲率", "=管制總覽!G9", "percent_one", "0.80", "percent_one", ">=0.80"),
        ("03", "詐欺破獲率", "=管制總覽!M9", "percent_one", "0.60", "percent_one", ">=0.60"),
        ("04", "毒調率", "=管制總覽!A14", "percent_one", "毒調率目標", "percent_one", ">=毒調率目標"),
        ("05", "未到驗人口", "=管制總覽!G14", "integer", "5", "integer", "<=5"),
        ("06", "未破案件", "=管制總覽!M14", "integer", "5", "integer", "<=5"),
        ("07", "交通達成率", "=交通績效管制!E8", "percent_one", "0.80", "percent_one", ">=0.80"),
        ("08", "本所員警在職數", '=COUNTIF(設定表!$D$13:$D$72,"是")', "integer", "—", "general", "info"),
    ]

    for idx, (no, name, vf, fmt, target, target_fmt, threshold) in enumerate(kpi_list):
        r = 12 + idx
        # A 編號
        S.set_cell(ws, f"A{r}", no,
                   font_key="body_bold", fill_key="calc",
                   align_key="center", border_key="all_thin")
        # B KPI 名
        S.set_cell(ws, f"B{r}", name,
                   font_key="body_bold", fill_key="calc",
                   align_key="left", border_key="all_thin")
        # C-D 數值 (merge)
        ws.merge_cells(f"C{r}:D{r}")
        # vf might be a formula starting with =
        val = vf if vf.startswith("=") else f"={vf}"
        S.set_cell(ws, f"C{r}", val,
                   font_key=S.font(13, bold=True, color="accent"),
                   fill_key="calc",
                   align_key="center", border_key="all_thin",
                   number_format=fmt)
        # E-F 目標
        ws.merge_cells(f"E{r}:F{r}")
        if target_fmt == "general":
            S.set_cell(ws, f"E{r}", target,
                       font_key="body", fill_key="calc",
                       align_key="center", border_key="all_thin")
        else:
            tval = target
            if target.startswith("毒調率"):
                tval = f"=毒調率目標"
            else:
                tval = f"={target}"
            S.set_cell(ws, f"E{r}", tval,
                       font_key="body", fill_key="calc",
                       align_key="center", border_key="all_thin",
                       number_format=target_fmt)
        # G-H 達標狀態
        ws.merge_cells(f"G{r}:H{r}")
        if threshold == "info":
            S.set_cell(ws, f"G{r}", '="—"',
                       font_key=S.font(11, color="muted"),
                       fill_key="calc",
                       align_key="center", border_key="all_thin")
        else:
            status_fml = f'=IF(IFERROR(C{r}{threshold},FALSE),"● 達標","● 未達標")'
            # special: threshold starts with operator
            S.set_cell(ws, f"G{r}", status_fml,
                       font_key="body_bold", fill_key="calc",
                       align_key="center", border_key="all_thin")

        ws.row_dimensions[r].height = S.ROW_HEIGHT["default"]

    # === Row 20 spacer ===
    ws.row_dimensions[20].height = 10

    # === Row 21 Section「★ 必附附件」===
    ws.merge_cells("A21:H21")
    S.set_cell(ws, "A21",
               "   ★  必附附件（依上呈對象動態切換）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[21].height = S.ROW_HEIGHT["header"]

    # Row 22-25 附件 (CHOOSE 4 行)
    appendix_lines = [
        # (row_idx_in_choose, 對分局, 對警察局, 對警政署)
        (1, "附件 1：未破案件明細表（頁 6 列印）", "附件 1：案類分類統計（頁 10 列印）", "附件 1：標準破獲率表（頁 2 列印）"),
        (2, "附件 2：員警勤務統計（頁 10 列印）", "附件 2：達標未達標項目清單（頁 9 列印）", "附件 2：列管人口處理（頁 8 列印）"),
        (3, "附件 3：毒品調驗人口明細（頁 8 列印）", "附件 3：與全市平均比較表", "附件 3：違規取締達成統計（頁 9 列印）"),
        (4, "附件 4：跨期間趨勢分析（頁 12 列印）", "（無）", "（無）"),
    ]
    for r_idx, (n, opt1, opt2, opt3) in enumerate(appendix_lines):
        r = 21 + r_idx + 1
        ws.merge_cells(f"A{r}:H{r}")
        fml = (
            f'=CHOOSE(MATCH({TARGET_CELL},{{"對分局","對警察局","對警政署"}},0),'
            f'"  {opt1}",'
            f'"  {opt2}",'
            f'"  {opt3}")'
        )
        S.set_cell(ws, f"A{r}", fml,
                   font_key=S.font(11, color="text"),
                   fill_key="calc",
                   align_key="left", border_key="all_thin")
        ws.row_dimensions[r].height = 24

    # === Row 26 spacer ===
    ws.row_dimensions[26].height = 12

    # === Row 27 簽章區 ===
    ws.merge_cells("A27:D27")
    S.set_cell(ws, "A27", "  發文人：所長 __________________  簽章",
               font_key=S.font(11, color="text"),
               fill_key=None, align_key="left",
               border_key="bottom_medium")
    ws.merge_cells("E27:H27")
    S.set_cell(ws, "E27", "  收文機關承辦：__________________  簽章",
               font_key=S.font(11, color="text"),
               fill_key=None, align_key="left",
               border_key="bottom_medium")
    ws.row_dimensions[27].height = 36

    # === Row 28 頁尾 ===
    ws.merge_cells("A28:H28")
    footer = (
        '="© KKEVIN-LIN-2026-V2.0  ｜  CHOOSE/INDIRECT 動態切換  ｜  '
        '上呈對象：" & ' + TARGET_CELL + ' & "  ｜  製表："&TEXT(今日,"e/mm/dd")'
    )
    S.set_cell(ws, "A28", footer,
               font_key=S.font(9, color="muted", italic=True),
               fill_key="calc", align_key="center")
    ws.row_dimensions[28].height = 22

    # 公式計數
    fcount = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                fcount += 1
    log.info(f"  上呈對象下拉：{TARGET_CELL} (3 選項)")
    log.info(f"  8 KPI 主表 + 4 行動態附件")
    log.info(f"  公式總數：{fcount}")
    return fcount


def build():
    log = get_logger("phase4_p4")
    log.info("===== Phase 4 — 頁 4 對上級機關上呈 開始 =====")

    wb = load_workbook(OUTPUT_PATH)
    fcount = build_page4_upreport(wb, log)

    bak = backup_existing()
    if bak:
        log.info(f"  舊版備份：{bak}")
    wb.save(OUTPUT_PATH)
    size_kb = OUTPUT_PATH.stat().st_size / 1024
    log.info(f"  儲存：{OUTPUT_PATH.name} ({size_kb:.1f} KB)")
    log.info(f"  公式：{fcount}")
    log.info("===== Phase 4 — 頁 4 完成 =====\n")
    return OUTPUT_PATH


if __name__ == "__main__":
    build()
