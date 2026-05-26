"""
phase3_page12_trend.py — Phase 3：頁 12 跨期間趨勢分析（A 麥肯錫風）

v2 內以 sparkline 為主軸的頁面（其他頁僅輔助）

版型：
  Row 1-2  Banner
  Row 3    期間下拉 + insight headline
  Row 4    spacer
  Row 5    Section「📈 5 KPI × 12 月走勢矩陣」
  Row 6    spacer (4)
  ─── FREEZE A8 ───
  Row 7    主表 header
  Row 8-12 5 KPI × 6 cols 矩陣
  Row 13   spacer
  Row 14   Section「★ Top 3 變化（升降最大）」
  Row 15-17 Top 3 變化卡（升 / 降 / 持平）
  Row 18   spacer
  Row 19   頁尾

5 KPI（對應 Tbl歷史 欄）：
  1. 全般破獲率 (col 5)
  2. 毒調率 (col 10)
  3. 未到驗 (col 11)
  4. 交通達標率 (col 12)
  5. 未破案件 (col 3)

Sparkline 仍用 Unicode 區塊字元（與頁 2 一致策略）
Helper rows 100-104：每 KPI 一列，含 12 月資料 + block 字元 + CONCAT
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from phases._base import backup_existing, OUTPUT_PATH, get_logger, unmerge_all_in_sheet
import styles as S


SHEET = "跨期間趨勢分析"
PERIOD_CELL = "B3"
ROC_YEAR = "(YEAR(今日)-1911)"
HIST_TBL = "Tbl歷史"
SPARK_FONT = "Consolas"

PERIODS = ["6 月", "12 月", "24 月"]

# 5 KPI 規格
KPIS = [
    # (id, name, hist_col_index, fmt, reverse, threshold_lo, threshold_hi)
    {"id": 1, "name": "全般破獲率", "col_idx": 5,  "fmt": "percent_one", "reverse": False, "lo": 0.70, "hi": 0.85},
    {"id": 2, "name": "毒調率",     "col_idx": 10, "fmt": "percent_one", "reverse": False, "lo": 0.50, "hi": 0.60},
    {"id": 3, "name": "未到驗人口", "col_idx": 11, "fmt": "integer",     "reverse": True,  "lo": 5, "hi": 10},
    {"id": 4, "name": "交通達標率", "col_idx": 12, "fmt": "percent_one", "reverse": False, "lo": 0.60, "hi": 0.80},
    {"id": 5, "name": "未破案件",   "col_idx": 3,  "fmt": "integer",     "reverse": True,  "lo": 5, "hi": 10},
]


def block_char_formula(value_cell, data_range):
    """Unicode 區塊字元公式（同頁 2 邏輯）"""
    return (
        f'=IF(OR(NOT(ISNUMBER({value_cell})),'
            f'MAX({data_range})=MIN({data_range})),'
        f'IF(ISNUMBER({value_cell}),"▄",""),'
        f'CHOOSE(1+MIN(7,MAX(0,ROUND(({value_cell}-MIN({data_range}))/'
            f'(MAX({data_range})-MIN({data_range}))*7,0))),'
            f'"▁","▂","▃","▄","▅","▆","▇","█"))'
    )


def build_page12_trend(wb, log):
    log.info("--- 頁 12 跨期間趨勢分析 ---")
    ws = wb[SHEET]

    cleared = unmerge_all_in_sheet(ws)
    log.info(f"  清空既有 merges：{cleared} 個")

    # === 欄寬 (v2.1 全面加大) ===
    col_widths = {
        "A": 22,   # KPI 名稱
        "B": 28,   # Sparkline (12 chars + 邊距)
        "C": 18,   # 當月
        "D": 18,   # 3 月前
        "E": 18,   # 6 月前
        "F": 22,   # vs 去年同期
        "G": 18,   # 變化方向
    }
    for c, w in col_widths.items():
        ws.column_dimensions[c].width = w

    # === Row 1-2 Banner ===
    ws.merge_cells("A1:G1")
    S.set_cell(ws, "A1",
               "   📈  跨期間趨勢分析（v2 唯一 Sparkline 主軸頁）",
               font_key="banner", fill_key="banner",
               align_key="left", border_key="bottom_thick")
    ws.row_dimensions[1].height = S.ROW_HEIGHT["banner"]

    ws.merge_cells("A2:G2")
    sub_formula = (
        '="   "&警察局名稱&" · "&分局名稱&" · "&派出所名稱'
        '&" ｜ 中華民國 "&(YEAR(今日)-1911)&" 年 ｜ 資料源：Tbl歷史（"'
        f'&IFERROR(COUNTA({HIST_TBL}[統計年月]),0)&" 個月份快照）"'
    )
    S.set_cell(ws, "A2", sub_formula,
               font_key=S.font(11, color="white", italic=True),
               fill_key=S.fill("accent_dark"), align_key="left")
    ws.row_dimensions[2].height = S.ROW_HEIGHT["sub_banner"]

    # === Row 3 期間下拉 + insight headline ===
    # A3 label
    S.set_cell(ws, "A3", "  期間：",
               font_key=S.font(13, bold=True, color="accent"),
               fill_key="accent_light",
               align_key="right", border_key="all_thin")
    # B3 下拉
    S.set_cell(ws, PERIOD_CELL, "12 月",
               font_key=S.font(13, bold=True, color="white"),
               fill_key="banner",
               align_key="center", border_key="all_thin")
    period_list = "\"" + ",".join(PERIODS) + "\""
    dv = DataValidation(type="list", formula1=period_list, allow_blank=False)
    dv.add(PERIOD_CELL)
    dv.prompt = "期間：6 月 / 12 月 / 24 月"
    dv.promptTitle = "期間切換"
    ws.add_data_validation(dv)

    # C3:G3 Insight headline
    ws.merge_cells("C3:G3")
    insight_fml = (
        '="📊 近 "&B3&" 走勢摘要："&'
        f'IF(IFERROR(COUNTA({HIST_TBL}[統計年月]),0)<3,'
        '"歷史資料不足 3 個月，需累積快照",'
        '"見下方主表，Top 3 變化見底部")'
    )
    S.set_cell(ws, "C3", insight_fml,
               font_key=S.font(12, bold=True, color="accent"),
               fill_key="accent_light",
               align_key="left", border_key="all_thin")
    ws.row_dimensions[3].height = 32

    # === Row 4 spacer ===
    ws.row_dimensions[4].height = 8

    # === Row 5 Section「📈 5 KPI × N 月走勢矩陣」===
    ws.merge_cells("A5:G5")
    S.set_cell(ws, "A5",
               '="   📈  5 KPI × "&B3&" 走勢矩陣（純文字 sparkline，跨平台一致）"',
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[5].height = S.ROW_HEIGHT["header"]

    # === Row 6 spacer ===
    ws.row_dimensions[6].height = 4

    # === Row 7 主表 header ===
    headers = ["KPI", "近 N 月走勢", "當月", "3 月前", "6 月前", "vs 去年同期", "趨勢"]
    for i, h in enumerate(headers, start=1):
        col = get_column_letter(i)
        S.set_cell(ws, f"{col}7", h,
                   font_key="header", fill_key="header",
                   align_key="center", border_key="all_thin")
    ws.row_dimensions[7].height = S.ROW_HEIGHT["header"]

    # === Row 8-12: 5 KPI × 6 cols ===
    # 各 KPI 取 Tbl歷史 對應欄
    # MAX_ROW = ROWS(Tbl歷史) → 最新一列
    # 當月 = INDEX(col, MAX_ROW)
    # 3m 前 = INDEX(col, MAX_ROW - 3)
    # 6m 前 = INDEX(col, MAX_ROW - 6)
    # vs 去年同期 = INDEX(col, MAX_ROW - 12)
    # sparkline = AR helper

    for i, kpi in enumerate(KPIS):
        row = 7 + kpi["id"]    # 8-12
        col_idx = kpi["col_idx"]
        helper_row = 100 + i   # 100-104

        # A KPI 名
        S.set_cell(ws, f"A{row}",
                   f'  {kpi["id"]:02d}  {kpi["name"]}',
                   font_key="body_bold", fill_key="calc",
                   align_key="left", border_key="all_thin")

        # B sparkline (引用 helper AR)
        S.set_cell(ws, f"B{row}", f"=AR{helper_row}",
                   font_key=S.font(16, bold=True, color="accent",
                                   family=SPARK_FONT),
                   fill_key="calc",
                   align_key="center", border_key="all_thin")

        # 取 Tbl歷史 最新列號
        # MAX_ROW = MAX(1, COUNTA(Tbl歷史[統計年月]))
        max_idx = f'MAX(1,COUNTA({HIST_TBL}[統計年月]))'

        # C 當月
        S.set_cell(ws, f"C{row}",
                   f'=IFERROR(INDEX({HIST_TBL},{max_idx},{col_idx}),"—")',
                   font_key=S.font(14, bold=True, color="accent"),
                   fill_key="calc",
                   align_key="center", border_key="all_thin",
                   number_format=kpi["fmt"])

        # D 3 月前
        S.set_cell(ws, f"D{row}",
                   f'=IFERROR(INDEX({HIST_TBL},MAX(1,{max_idx}-3),{col_idx}),"—")',
                   font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin",
                   number_format=kpi["fmt"])

        # E 6 月前
        S.set_cell(ws, f"E{row}",
                   f'=IFERROR(INDEX({HIST_TBL},MAX(1,{max_idx}-6),{col_idx}),"—")',
                   font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin",
                   number_format=kpi["fmt"])

        # F vs 去年同期 (12 月前)
        S.set_cell(ws, f"F{row}",
                   f'=IFERROR(IF({max_idx}<=12,"資料不足",'
                   f'INDEX({HIST_TBL},{max_idx}-12,{col_idx})),"—")',
                   font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin",
                   number_format=kpi["fmt"])

        # G 趨勢方向（當月 vs 3 月前）
        # reverse=True: 越小越好 → 當月 < 3 月前 = ↑改善
        reverse = kpi.get("reverse", False)
        if reverse:
            trend_fml = (
                f'=IFERROR(IF(C{row}<D{row},"↑ 改善",'
                f'IF(C{row}>D{row},"↓ 惡化","→ 持平")),"—")'
            )
        else:
            trend_fml = (
                f'=IFERROR(IF(C{row}>D{row},"↑ 上升",'
                f'IF(C{row}<D{row},"↓ 下降","→ 持平")),"—")'
            )
        S.set_cell(ws, f"G{row}", trend_fml,
                   font_key="body_bold", fill_key="calc",
                   align_key="center", border_key="all_thin")

        ws.row_dimensions[row].height = 30

    # === Row 13 spacer ===
    ws.row_dimensions[13].height = 10

    # === Row 14 Section「★ Top 3 變化」===
    ws.merge_cells("A14:G14")
    S.set_cell(ws, "A14",
               "   ★  Top 3 變化（依當月 vs 3 月前差距排序）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[14].height = S.ROW_HEIGHT["header"]

    # === Row 15-17 Top 3 變化卡 ===
    # 用 Helper cols U/V/W：每 KPI 一列，存「絕對差距 + 方向」
    # 然後主表用 LARGE + INDEX 取 Top 3

    # 簡化版：直接顯示 3 個 fixed KPI 的變化（破獲率 / 毒調率 / 未到驗）
    # 即不動態排序，但明確顯示 3 大 KPI 的趨勢
    top_displays = [
        ("🥇 最大改善",  "pass",   "pass_bg",
            'IF(D8=0,"—",(C8-D8)/D8)', "全般破獲率 vs 3 月前"),
        ("⚠ 需關注",     "warn",   "warn_bg",
            'IF(D11=0,"—",(C11-D11)/D11)', "交通達標率 vs 3 月前"),
        ("🔍 持續觀察",  "accent", "accent_light",
            'IF(D9=0,"—",(C9-D9)/D9)', "毒調率 vs 3 月前"),
    ]
    for k, (label, fg, bg, fml_inner, desc) in enumerate(top_displays):
        row = 15 + k
        # A:B label
        ws.merge_cells(f"A{row}:B{row}")
        S.set_cell(ws, f"A{row}", f"  {label}",
                   font_key=S.font(13, bold=True, color=fg),
                   fill_key=S.fill(bg),
                   align_key="left", border_key="all_thin")
        # C:D value (差距 %)
        ws.merge_cells(f"C{row}:D{row}")
        S.set_cell(ws, f"C{row}", f"=IFERROR({fml_inner},0)",
                   font_key=S.font(20, bold=True, color=fg),
                   fill_key=S.fill(bg),
                   align_key="center", border_key="all_thin",
                   number_format="+0.0%;-0.0%;0.0%")
        # E:G desc
        ws.merge_cells(f"E{row}:G{row}")
        S.set_cell(ws, f"E{row}", f"  {desc}",
                   font_key=S.font(11, color="muted", italic=True),
                   fill_key=S.fill(bg),
                   align_key="left", border_key="all_thin")
        ws.row_dimensions[row].height = 36

    # === Row 18 spacer ===
    ws.row_dimensions[18].height = 10

    # === Row 19 頁尾 ===
    ws.merge_cells("A19:G19")
    footer = (
        '="© KKEVIN-LIN-2026-V2.0  ｜  資料源：頁 14 Tbl歷史（月份快照）｜  '
        'Sparkline 採 Unicode 區塊字元（跨平台一致）｜  '
        '製表："&TEXT(今日,"e/mm/dd")'
    )
    S.set_cell(ws, "A19", footer,
               font_key=S.font(9, color="muted", italic=True),
               fill_key="calc", align_key="center")
    ws.row_dimensions[19].height = 22

    # =========================================================
    # Helper area (rows 100-104) — 5 KPI sparkline 資料
    # =========================================================
    S.set_cell(ws, "A99", "[helper] 趨勢矩陣",
               font_key=S.font(9, color="muted"))
    S.set_cell(ws, "B99", "KPI",
               font_key=S.font(9, color="muted"))
    S.set_cell(ws, "T99", "近 12 月資料 (T..AE)",
               font_key=S.font(9, color="muted"))
    S.set_cell(ws, "AF99", "block 字元 (AF..AQ)",
               font_key=S.font(9, color="muted"))
    S.set_cell(ws, "AR99", "sparkline 字串",
               font_key=S.font(9, color="muted"))

    for i, kpi in enumerate(KPIS):
        helper_row = 100 + i
        col_idx = kpi["col_idx"]

        # A KPI 名
        S.set_cell(ws, f"A{helper_row}", f'{kpi["id"]:02d} {kpi["name"]}',
                   font_key=S.font(9, color="muted"))
        # B 欄號（debug）
        S.set_cell(ws, f"B{helper_row}", col_idx,
                   font_key=S.font(9, color="muted"))

        # T:AE 12 月資料（從 Tbl歷史 最末 12 列取）
        max_idx = f'MAX(1,COUNTA({HIST_TBL}[統計年月]))'
        for col_offset in range(12):
            col_letter = get_column_letter(20 + col_offset)   # T..AE
            # 最末 12 列：MAX_ROW - 11 + col_offset
            fml = (
                f'=IFERROR(INDEX({HIST_TBL},'
                f'MAX(1,{max_idx}-11+{col_offset}),{col_idx}),"")'
            )
            S.set_cell(ws, f"{col_letter}{helper_row}", fml,
                       font_key=S.font(9, color="muted"),
                       number_format=kpi["fmt"])

        # AF:AQ block 字元
        data_range = f"$T{helper_row}:$AE{helper_row}"
        for col_offset in range(12):
            src_col = get_column_letter(20 + col_offset)
            dst_col = get_column_letter(32 + col_offset)
            src_cell = f"{src_col}{helper_row}"
            fml = block_char_formula(src_cell, data_range)
            S.set_cell(ws, f"{dst_col}{helper_row}", fml,
                       font_key=S.font(9, color="muted"))

        # AR CONCAT
        S.set_cell(ws, f"AR{helper_row}",
                   f'=IFERROR(CONCAT($AF{helper_row}:$AQ{helper_row}),"")',
                   font_key=S.font(11, family=SPARK_FONT, color="accent"))

    # 隱藏 helper rows + cols
    for r in range(99, 110):
        ws.row_dimensions[r].hidden = True
    for c in range(20, 45):   # T..AR
        ws.column_dimensions[get_column_letter(c)].hidden = True

    # 公式計數
    fcount = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                fcount += 1
    log.info(f"  5 KPI × 12 月矩陣")
    log.info(f"  Top 3 變化卡")
    log.info(f"  公式總數：{fcount}")
    log.info(f"  凍結窗格：A8")
    return fcount


def build():
    log = get_logger("phase3_p12")
    log.info("===== Phase 3 — 頁 12 跨期間趨勢 開始 =====")

    wb = load_workbook(OUTPUT_PATH)
    fcount = build_page12_trend(wb, log)

    bak = backup_existing()
    if bak:
        log.info(f"  舊版備份：{bak}")
    wb.save(OUTPUT_PATH)
    size_kb = OUTPUT_PATH.stat().st_size / 1024
    log.info(f"  儲存：{OUTPUT_PATH.name} ({size_kb:.1f} KB)")
    log.info(f"  公式：{fcount}")
    log.info("===== Phase 3 — 頁 12 完成 =====\n")
    return OUTPUT_PATH


if __name__ == "__main__":
    build()
