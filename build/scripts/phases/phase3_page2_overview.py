"""
phase3_page2_overview.py — Phase 3：頁 2 管制總覽（A 麥肯錫風 — 九宮格小圖）

版型（plan B = 完整規格）：
  Row 1-2  Banner（派出所抬頭 + 民國年月）
  Row 3    spacer
  Row 4    Insight headline（一句話結論，公式組句）
  Row 5    spacer
  Row 6    Section title「📊 9 KPI 九宮格」
  Row 7    spacer
  ─── FREEZE A8 ───
  Row 8-11   KPI #1 全般破獲率 / #2 竊盜破獲率 / #3 詐欺破獲率
  Row 12     spacer
  Row 13-16  KPI #4 毒調率 / #5 未到驗 / #6 未破案件
  Row 17     spacer
  Row 18-21  KPI #7 交通達成率 / #8 酒駕 / #9 闖紅燈
  Row 22     spacer
  Row 23     Section title「⚠ 警示摘要」
  Row 24-26  紅燈總數 / 黃燈總數 / 綠燈總數
  Row 27     spacer
  Row 28     頁尾

每 KPI 4 列 × 5 欄：
  row+0    label（KPI 名稱）+ status light（最右欄）
  row+1    大字現值（merged，~56pt 高）
  row+2    threshold / vs 比較行
  row+3    sparkline 儲存格（Unicode 區塊字元，~24pt）

★ Sparkline 採用方案 A：Unicode 區塊字元（▁▂▃▄▅▆▇█）
  - openpyxl 3.1.5 無原生 sparkline 支援
  - 用 helper rows（隱藏）儲存最近 12 月資料 + 對應 block 字元
  - sparkline cell = CONCAT(12 個 block 字元)
  - 等寬字 Consolas 保證對齊

注意（資料源缺口）：
  - KPI 4 毒調率 / 5 未到驗 — 目前無 Tbl毒調，現值顯示「—」，sparkline 用 Tbl歷史 可繪。
  - KPI 8 酒駕 / 9 闖紅燈 — Tbl歷史無對應欄，現值用 Tbl交通 可算，sparkline 顯示「—」。
  - 待頁 8 Tbl毒調 建好 + Tbl歷史擴充酒駕/闖紅燈欄後回補（標註 TODO）。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill, Font

from phases._base import backup_existing, OUTPUT_PATH, get_logger, unmerge_all_in_sheet
import styles as S


# ============================================================
# 9 KPI 規格表
# ============================================================
# 民國當年 = YEAR(今日) - 1911
ROC_YEAR = "(YEAR(今日)-1911)"

KPIS = [
    # threshold: 預設 reverse=False（值越大越好）
    {
        "id": 1,
        "name": "全般破獲率",
        "value_formula":
            f'=IFERROR('
            f'COUNTIFS(Tbl案件[查獲管轄],"本轄",Tbl案件[是否破獲],"是",Tbl案件[歸屬年度],{ROC_YEAR})'
            f'/'
            f'COUNTIFS(Tbl案件[發生管轄],"本轄",Tbl案件[歸屬年度],{ROC_YEAR},Tbl案件[計入發生數],"是")'
            f',0)',
        "fmt": "percent_one",
        "threshold_lo": 0.70, "threshold_hi": 0.85,
        "threshold_note": "目標 ≥ 85% ｜ 黃燈 70-85% ｜ 紅燈 < 70%",
        "spark_col": "破獲率",
        "value_placeholder": False,
    },
    {
        "id": 2,
        "name": "竊盜破獲率",
        "value_formula":
            f'=IFERROR('
            f'COUNTIFS(Tbl案件[案類分類],"竊盜",Tbl案件[查獲管轄],"本轄",Tbl案件[是否破獲],"是",Tbl案件[歸屬年度],{ROC_YEAR})'
            f'/'
            f'COUNTIFS(Tbl案件[案類分類],"竊盜",Tbl案件[發生管轄],"本轄",Tbl案件[歸屬年度],{ROC_YEAR},Tbl案件[計入發生數],"是")'
            f',0)',
        "fmt": "percent_one",
        "threshold_lo": 0.60, "threshold_hi": 0.80,
        "threshold_note": "目標 ≥ 80% ｜ 黃燈 60-80% ｜ 紅燈 < 60%",
        "spark_col": "竊盜",
        "value_placeholder": False,
    },
    {
        "id": 3,
        "name": "詐欺破獲率",
        "value_formula":
            f'=IFERROR('
            f'COUNTIFS(Tbl案件[案類分類],"詐欺",Tbl案件[查獲管轄],"本轄",Tbl案件[是否破獲],"是",Tbl案件[歸屬年度],{ROC_YEAR})'
            f'/'
            f'COUNTIFS(Tbl案件[案類分類],"詐欺",Tbl案件[發生管轄],"本轄",Tbl案件[歸屬年度],{ROC_YEAR},Tbl案件[計入發生數],"是")'
            f',0)',
        "fmt": "percent_one",
        "threshold_lo": 0.40, "threshold_hi": 0.60,
        "threshold_note": "目標 ≥ 60% ｜ 黃燈 40-60% ｜ 紅燈 < 40%",
        "spark_col": "詐欺",
        "value_placeholder": False,
    },
    {
        "id": 4,
        "name": "毒調率",
        # v2.2 接通 Tbl毒調：(已驗+通緝+強採+在監) / 列管總數
        # Tbl毒調[是否到驗] 計算欄已含 4 種「算到驗」狀態，直接 SUM 即可
        "value_formula":
            '=IFERROR(SUM(Tbl毒調[是否到驗])/COUNTA(Tbl毒調[姓名]),0)',
        "fmt": "percent_one",
        "threshold_lo": 0.50, "threshold_hi": 0.60,
        "threshold_note": "目標 ≥ 60% ｜ 黃燈 50-60% ｜ 紅燈 < 50%",
        "spark_col": "毒調率",
        "value_placeholder": False,
    },
    {
        "id": 5,
        "name": "未到驗人口",
        # v2.2 接通 Tbl毒調：管制情形含「未到驗」字樣的件數
        "value_formula":
            '=IFERROR(COUNTIF(Tbl毒調[管制情形],"*未到驗*"),0)',
        "fmt": "integer",
        "threshold_lo": 5, "threshold_hi": 10,
        "threshold_note": "目標 < 5 人 ｜ 黃燈 5-10 ｜ 紅燈 > 10（值越小越好）",
        "spark_col": "未到驗",
        "value_placeholder": False,
        "reverse": True,
    },
    {
        "id": 6,
        "name": "未破案件",
        "value_formula":
            f'=IFERROR('
            f'COUNTIFS(Tbl案件[案件狀況],"尚未偵破",Tbl案件[發生管轄],"本轄",Tbl案件[歸屬年度],{ROC_YEAR},Tbl案件[計入發生數],"是")'
            f',0)',
        "fmt": "integer",
        "threshold_lo": 5, "threshold_hi": 10,
        "threshold_note": "目標 ≤ 5 件 ｜ 黃燈 5-10 件 ｜ 紅燈 > 10 件（值越小越好）",
        "spark_col": "全般未破",
        "value_placeholder": False,
        "reverse": True,
    },
    {
        "id": 7,
        "name": "交通達成率",
        "value_formula": "=IFERROR(交通績效管制!E8,0)",
        "fmt": "percent_one",
        "threshold_lo": 0.60, "threshold_hi": 0.80,
        "threshold_note": "目標 ≥ 80% ｜ 黃燈 60-80% ｜ 紅燈 < 60%",
        "spark_col": "交通達標",
        "value_placeholder": False,
    },
    {
        "id": 8,
        "name": "酒駕取締",
        "value_formula": '=IFERROR(COUNTIF(Tbl交通[違規項目],"*酒駕*"),0)',
        "fmt": "integer",
        "threshold_lo": 1, "threshold_hi": 3,
        "threshold_note": "目標 ≥ 3 件 ｜ 黃燈 1-3 件 ｜ 紅燈 0 件",
        "spark_col": None,   # TODO: Tbl歷史 待擴充酒駕欄
        "value_placeholder": False,
    },
    {
        "id": 9,
        "name": "闖紅燈取締",
        "value_formula": '=IFERROR(COUNTIF(Tbl交通[違規項目],"*闖紅燈*"),0)',
        "fmt": "integer",
        "threshold_lo": 1, "threshold_hi": 5,
        "threshold_note": "目標 ≥ 5 件 ｜ 黃燈 1-5 件 ｜ 紅燈 0 件",
        "spark_col": None,   # TODO: Tbl歷史 待擴充闖紅燈欄
        "value_placeholder": False,
    },
]


# Tbl歷史 結構：14 欄
# 1 統計年月 / 2 全般發生 / 3 全般未破 / 4 全般破獲 / 5 破獲率 / 6 竊盜 / 7 詐欺
# 8 已破未送 / 9 發展中 / 10 毒調率 / 11 未到驗 / 12 交通達標 / 13 員警冠軍 / 14 備註
HIST_COL_INDEX = {
    "破獲率": 5, "竊盜": 6, "詐欺": 7, "毒調率": 10,
    "未到驗": 11, "全般未破": 3, "交通達標": 12,
}

# Sparkline 等寬字（保證對齊）
SPARK_FONT = "Consolas"


# ============================================================
# 燈號公式產生器
# ============================================================
def light_formula(value_cell, lo, hi, reverse=False, placeholder=False):
    """產生燈號公式（● 綠 / ⚠ 黃 / ● 紅 / ○ 待接通）。

    reverse=False：值越大越好。>=hi 綠，>=lo 黃，else 紅。
    reverse=True ：值越小越好。<=lo 綠，<=hi 黃，else 紅。
    placeholder=True：直接顯示「○ 待接通」。
    """
    if placeholder:
        return '="○ 待接通"'
    if reverse:
        return f'=IF({value_cell}<={lo},"● 綠",IF({value_cell}<={hi},"⚠ 黃","● 紅"))'
    return f'=IF({value_cell}>={hi},"● 綠",IF({value_cell}>={lo},"⚠ 黃","● 紅"))'


# ============================================================
# Sparkline helper：產生 12 個 block 字元的公式
# ============================================================
def block_char_formula(value_cell, data_range):
    """value_cell：當期值；data_range：12 cells 的範圍（用 MIN/MAX normalize）

    回傳公式：依 value 在 data_range 的 min..max 之間落點 → 8 級 block 字元
    ▁=2581 ▂=2582 ▃=2583 ▄=2584 ▅=2585 ▆=2586 ▇=2587 █=2588

    處理空值：IFERROR(...,"") + IF(value_cell="" → "")
    處理同值（MIN=MAX）：用 MAX(1,MAX-MIN) 避免除零
    """
    return (
        f'=IF(OR(NOT(ISNUMBER({value_cell})),'
            f'MAX({data_range})=MIN({data_range})),'
        # 全相同或空：顯示中間級
        f'IF(ISNUMBER({value_cell}),"▄",""),'
        # 正常情況：normalize 到 0..7 → CHOOSE
        f'CHOOSE(1+MIN(7,MAX(0,ROUND(({value_cell}-MIN({data_range}))/'
            f'(MAX({data_range})-MIN({data_range}))*7,0))),'
            f'"▁","▂","▃","▄","▅","▆","▇","█"))'
    )


# ============================================================
# 主建檔函式
# ============================================================
def build_page2_overview(wb, log):
    log.info("--- 頁 2 管制總覽 ---")
    ws = wb["管制總覽"]

    # 清掉 Phase 1 既有 banner merge（A1:G1, A2:G2），避免與下面 A1:Q1 重疊
    cleared = unmerge_all_in_sheet(ws)
    log.info(f"  清空既有 merges：{cleared} 個（避免與新 banner 重疊）")

    # 欄寬：3 大區（A-E, G-K, M-Q）+ 2 spacer (F, L)
    # v2.1 加大：每 KPI 5 cols × 10 = 50 chars 寬
    col_widths = {
        "A": 10, "B": 10, "C": 10, "D": 10, "E": 10,
        "F": 3,
        "G": 10, "H": 10, "I": 10, "J": 10, "K": 10,
        "L": 3,
        "M": 10, "N": 10, "O": 10, "P": 10, "Q": 10,
    }
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    # === Row 1-2 Banner ===
    ws.merge_cells("A1:Q1")
    S.set_cell(ws, "A1", "   📊  管制總覽（V4 九宮格 Small Multiples）",
               font_key="banner", fill_key="banner",
               align_key="left", border_key="bottom_thick")
    ws.row_dimensions[1].height = S.ROW_HEIGHT["banner"]

    ws.merge_cells("A2:Q2")
    sub_formula = (
        '="   "&警察局名稱&" · "&分局名稱&" · "&派出所名稱'
        '&" ｜ 中華民國 "&(YEAR(今日)-1911)&" 年 "&TEXT(MONTH(今日),"00")&" 月 ｜ 製表："&TEXT(今日,"e/mm/dd")'
    )
    S.set_cell(ws, "A2", sub_formula,
               font_key=S.font(11, color="white", italic=True),
               fill_key=S.fill("accent_dark"), align_key="left")
    ws.row_dimensions[2].height = S.ROW_HEIGHT["sub_banner"]

    # === Row 3 spacer ===
    ws.row_dimensions[3].height = 8

    # === Row 4 Insight Headline ===
    ws.merge_cells("A4:Q4")
    headline_formula = (
        '=IF(COUNTIF(S50:S58,"*紅*")>0,'
            '"🚨 全所有 "&COUNTIF(S50:S58,"*紅*")&" 項紅燈需立即處理，請查看下方警示摘要",'
        'IF(COUNTIF(S50:S58,"*黃*")>0,'
            '"⚠ 全所狀態尚可，"&COUNTIF(S50:S58,"*黃*")&" 項黃燈需關注，建議排入週會檢討",'
            '"✓ 全所狀態良好：9 大 KPI 全綠燈通過"))'
    )
    S.set_cell(ws, "A4", headline_formula,
               font_key=S.font(14, bold=True, color="accent"),
               fill_key=S.fill("accent_light"),
               align_key="center", border_key="all_thin")
    ws.row_dimensions[4].height = 36

    # === Row 5 spacer ===
    ws.row_dimensions[5].height = 8

    # === Row 6 section title ===
    ws.merge_cells("A6:Q6")
    S.set_cell(ws, "A6",
               "   📊  9 KPI 九宮格 ｜ 每格：當期現值 + 狀態燈 + 近 12 月走勢",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[6].height = S.ROW_HEIGHT["header"]

    # === Row 7 spacer ===
    ws.row_dimensions[7].height = 6

    # === Rows 8-21 — KPI 3x3 grid ===
    kpi_row_starts = [8, 13, 18]
    kpi_col_starts = ["A", "G", "M"]
    kpi_col_ends = ["E", "K", "Q"]

    kpi_anchors = []

    for i, kpi in enumerate(KPIS):
        row_grp = i // 3
        col_grp = i % 3
        r0 = kpi_row_starts[row_grp]
        col_start = kpi_col_starts[col_grp]
        col_end = kpi_col_ends[col_grp]
        light_col = col_end                       # 最右欄放燈號
        label_end_col = chr(ord(col_end) - 1)     # 倒數第二欄

        label_range = f"{col_start}{r0}:{label_end_col}{r0}"
        light_cell = f"{light_col}{r0}"
        value_cell = f"{col_start}{r0+1}"
        threshold_cell = f"{col_start}{r0+2}"
        spark_cell = f"{col_start}{r0+3}"

        # Row+0: label
        ws.merge_cells(label_range)
        S.set_cell(ws, f"{col_start}{r0}",
                   f"  {i+1:02d}  {kpi['name']}",
                   font_key=S.font(11, bold=True, color="muted"),
                   fill_key="calc",
                   align_key="left", border_key="top_bottom")

        # 燈號 cell
        light_fml = light_formula(value_cell, kpi["threshold_lo"], kpi["threshold_hi"],
                                  reverse=kpi.get("reverse", False),
                                  placeholder=kpi.get("value_placeholder", False))
        S.set_cell(ws, light_cell, light_fml,
                   font_key=S.font(13, bold=True, color="accent"),
                   fill_key="calc",
                   align_key="center", border_key="top_bottom")

        # Row+1: 大字現值
        value_range = f"{col_start}{r0+1}:{col_end}{r0+1}"
        ws.merge_cells(value_range)
        S.set_cell(ws, value_cell, kpi["value_formula"],
                   font_key="kpi_value",
                   fill_key=None,
                   align_key="center",
                   number_format=kpi["fmt"])
        ws.row_dimensions[r0+1].height = 56

        # Row+2: threshold note
        th_range = f"{col_start}{r0+2}:{col_end}{r0+2}"
        ws.merge_cells(th_range)
        S.set_cell(ws, threshold_cell, f"  {kpi['threshold_note']}",
                   font_key=S.font(9, color="muted", italic=True),
                   fill_key=None,
                   align_key="left", border_key="bottom_thick")
        ws.row_dimensions[r0+2].height = 22

        # Row+3: sparkline cell（公式之後填入）
        spark_range = f"{col_start}{r0+3}:{col_end}{r0+3}"
        ws.merge_cells(spark_range)
        # 先設樣式，公式之後在 helper 段一起設
        S.set_cell(ws, spark_cell, "",
                   font_key=S.font(16, bold=True, color="accent", family=SPARK_FONT),
                   align_key="center")
        ws.row_dimensions[r0+3].height = 26

        kpi_anchors.append({
            "kpi": kpi,
            "value_cell": value_cell,
            "light_cell": light_cell,
            "spark_cell": spark_cell,
            "row": r0,
        })

        ws.row_dimensions[r0].height = 26

    # spacer rows between 3 grid rows
    for r in [12, 17]:
        ws.row_dimensions[r].height = 10

    # === Row 22 spacer ===
    ws.row_dimensions[22].height = 14

    # === Row 23 警示摘要 section title ===
    ws.merge_cells("A23:Q23")
    S.set_cell(ws, "A23",
               "   ⚠  警示摘要（公式自動聚合 9 KPI 燈號）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[23].height = S.ROW_HEIGHT["header"]

    # === Row 24-26 三卡 ===
    summary_cards = [
        ("● 紅燈總數", "danger", "danger_bg",
            '=COUNTIF(S50:S58,"*紅*")', "integer",
            "需立即處理"),
        ("⚠ 黃燈總數", "warn", "warn_bg",
            '=COUNTIF(S50:S58,"*黃*")', "integer",
            "建議週會檢討"),
        ("● 綠燈總數", "pass", "pass_bg",
            '=COUNTIF(S50:S58,"*綠*")', "integer",
            "持續維持"),
    ]
    for j, (label, fg, bg, fml, fmt, sub) in enumerate(summary_cards):
        c_start = kpi_col_starts[j]
        c_end = kpi_col_ends[j]
        ws.merge_cells(f"{c_start}24:{c_end}24")
        S.set_cell(ws, f"{c_start}24", f"  {label}",
                   font_key=S.font(11, bold=True, color=fg),
                   fill_key=S.fill(bg),
                   align_key="left", border_key="top_bottom")
        ws.merge_cells(f"{c_start}25:{c_end}25")
        S.set_cell(ws, f"{c_start}25", fml,
                   font_key=S.font(28, bold=True, color=fg),
                   fill_key=S.fill(bg),
                   align_key="center", number_format=fmt)
        ws.merge_cells(f"{c_start}26:{c_end}26")
        S.set_cell(ws, f"{c_start}26", f"  {sub}",
                   font_key=S.font(9, color="muted"),
                   fill_key=S.fill(bg),
                   align_key="left", border_key="bottom_thick")
    ws.row_dimensions[24].height = 24
    ws.row_dimensions[25].height = 44
    ws.row_dimensions[26].height = 20

    # === Row 27 spacer ===
    ws.row_dimensions[27].height = 12

    # === Row 28 頁尾 ===
    ws.merge_cells("A28:Q28")
    footer = (
        '="© KKEVIN-LIN-2026-V2.0  ｜  資料源：Tbl案件 + Tbl交通 + Tbl歷史  ｜  '
        '製表："&TEXT(今日,"e/mm/dd")&"  ｜  下次重算：開檔自動"'
    )
    S.set_cell(ws, "A28", footer,
               font_key=S.font(9, color="muted", italic=True),
               fill_key="calc", align_key="center")
    ws.row_dimensions[28].height = 22

    # ============================================================
    # Helper area（隱藏）：燈號 + sparkline 資料 + block 字元
    # ============================================================
    # 配置：
    #   R 欄 = KPI 序號與名稱（debug）
    #   S 欄 = 燈號文字（COUNTIF 用）
    #   T:AE = 12 個月原始資料（從 Tbl歷史 OFFSET 取）
    #   AF:AQ = 12 個對應的 block 字元
    #   AR = CONCAT 出的 12 字元 sparkline 字串

    S.set_cell(ws, "R49", "[helper] KPI",
               font_key=S.font(9, color="muted"))
    S.set_cell(ws, "S49", "燈",
               font_key=S.font(9, color="muted"))
    S.set_cell(ws, "T49", "近 12 月資料 (T..AE) →",
               font_key=S.font(9, color="muted"))
    S.set_cell(ws, "AF49", "block 字元 (AF..AQ) →",
               font_key=S.font(9, color="muted"))
    S.set_cell(ws, "AR49", "sparkline 字串",
               font_key=S.font(9, color="muted"))

    NUM_FMT_FOR_SPARK = {
        "percent_one": "0.0%",
        "integer": "#,##0",
        "general": "General",
    }

    for i, anchor in enumerate(kpi_anchors):
        helper_row = 50 + i
        kpi = anchor["kpi"]

        # R 欄：KPI 名稱
        S.set_cell(ws, f"R{helper_row}", f"{i+1:02d} {kpi['name']}",
                   font_key=S.font(9, color="muted"))
        # S 欄：燈號（引用 light_cell）
        S.set_cell(ws, f"S{helper_row}",
                   f"={anchor['light_cell']}",
                   font_key=S.font(9, color="muted"))

        # T:AE 欄：12 個月原始資料
        if kpi["spark_col"] is None:
            # 無歷史欄，全部空白
            for col_offset in range(12):
                col_letter = get_column_letter(20 + col_offset)
                S.set_cell(ws, f"{col_letter}{helper_row}", "",
                           font_key=S.font(9, color="muted"))
        else:
            col_idx = HIST_COL_INDEX[kpi["spark_col"]]
            for col_offset in range(12):
                col_letter = get_column_letter(20 + col_offset)
                fml = (
                    f'=IFERROR(INDEX(Tbl歷史,'
                    f'MAX(1,ROWS(Tbl歷史)-11+{col_offset}),'
                    f'{col_idx}),"")'
                )
                S.set_cell(ws, f"{col_letter}{helper_row}", fml,
                           font_key=S.font(9, color="muted"),
                           number_format=NUM_FMT_FOR_SPARK.get(kpi["fmt"], "General"))

        # AF:AQ 欄：對應 block 字元（依 T:AE 範圍 normalize）
        data_range = f"$T{helper_row}:$AE{helper_row}"
        for col_offset in range(12):
            src_col = get_column_letter(20 + col_offset)   # T..AE
            dst_col = get_column_letter(32 + col_offset)   # AF..AQ
            src_cell = f"{src_col}{helper_row}"
            if kpi["spark_col"] is None:
                S.set_cell(ws, f"{dst_col}{helper_row}", "",
                           font_key=S.font(9, color="muted"))
            else:
                fml = block_char_formula(src_cell, data_range)
                S.set_cell(ws, f"{dst_col}{helper_row}", fml,
                           font_key=S.font(9, color="muted"))

        # AR 欄：CONCAT 12 個 block 字元
        if kpi["spark_col"] is None:
            concat_fml = '="─ 待擴充 ─"'
        else:
            concat_fml = (
                f'=IFERROR(CONCAT($AF{helper_row}:$AQ{helper_row}),"")'
            )
        S.set_cell(ws, f"AR{helper_row}", concat_fml,
                   font_key=S.font(11, family=SPARK_FONT, color="accent"))

        # 把 sparkline 字串引用回主 spark_cell
        spark_cell = anchor["spark_cell"]
        spark_ref_fml = f"=AR{helper_row}"
        # 重設 spark_cell 公式
        ws[spark_cell] = spark_ref_fml
        # 樣式已在前面設好（Consolas 16pt accent）

    # 隱藏 helper rows 49-65 與 cols R-AR
    for r in range(49, 66):
        ws.row_dimensions[r].hidden = True
    # R=18 .. AR=44
    for c in range(18, 45):
        ws.column_dimensions[get_column_letter(c)].hidden = True

    # === 條件格式：燈號 cell + 大字 KPI cell 依「紅/黃/綠」自動上色 ===
    fill_red    = PatternFill("solid", fgColor="FFFEE2E2")  # danger_bg
    fill_yellow = PatternFill("solid", fgColor="FFFEF3C7")  # warn_bg
    fill_green  = PatternFill("solid", fgColor="FFDCFCE7")  # pass_bg
    font_red    = Font(name="微軟正黑體", bold=True, color="FFB91C1C")
    font_yellow = Font(name="微軟正黑體", bold=True, color="FFB45309")
    font_green  = Font(name="微軟正黑體", bold=True, color="FF15803D")

    for anchor in kpi_anchors:
        light_cell = anchor["light_cell"]
        value_cell = anchor["value_cell"]
        # 兩個 cell 都套同樣 3 條 CF 規則
        for cell in [light_cell, value_cell]:
            # 紅燈
            ws.conditional_formatting.add(cell,
                FormulaRule(formula=[f'ISNUMBER(SEARCH("紅",{light_cell}))'],
                            fill=fill_red, font=font_red, stopIfTrue=False))
            # 黃燈
            ws.conditional_formatting.add(cell,
                FormulaRule(formula=[f'ISNUMBER(SEARCH("黃",{light_cell}))'],
                            fill=fill_yellow, font=font_yellow, stopIfTrue=False))
            # 綠燈
            ws.conditional_formatting.add(cell,
                FormulaRule(formula=[f'ISNUMBER(SEARCH("綠",{light_cell}))'],
                            fill=fill_green, font=font_green, stopIfTrue=False))
    log.info(f"  條件格式：9 KPI × 2 cells × 3 規則 = 54 條（紅/黃/綠 動態上色）")

    log.info(f"  KPI 卡：{len(kpi_anchors)}")
    log.info(f"  Sparkline：Unicode 區塊字元方案（Consolas 16pt）")
    log.info(f"  Helper：rows 49-65 隱藏；cols R-AR 隱藏")
    log.info(f"  凍結窗格：A8（Phase 1 既有）")

    return len(kpi_anchors)


def build():
    log = get_logger("phase3_p2")
    log.info("===== Phase 3 — 頁 2 管制總覽 開始 =====")

    wb = load_workbook(OUTPUT_PATH)
    kpi_count = build_page2_overview(wb, log)

    bak = backup_existing()
    if bak:
        log.info(f"  舊版備份：{bak}")
    wb.save(OUTPUT_PATH)
    size_kb = OUTPUT_PATH.stat().st_size / 1024
    log.info(f"  儲存：{OUTPUT_PATH.name} ({size_kb:.1f} KB)")
    log.info(f"  KPI 卡：{kpi_count}")
    log.info("===== Phase 3 — 頁 2 完成 =====\n")
    return OUTPUT_PATH


if __name__ == "__main__":
    build()
