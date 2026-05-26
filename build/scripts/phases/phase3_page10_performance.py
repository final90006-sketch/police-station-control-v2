"""
phase3_page10_performance.py — Phase 3：頁 10 績效統計（B 公文風 — 二合一）

版型：
  Row 1-2  Banner
  Row 3    spacer
  Row 4    維護指引（資料源說明）
  Row 5    spacer
  Row 6    Section title「上半：按案類績效統計（5 大案類）」
  Row 7    spacer
  ─── FREEZE A8 ───
  Row 8    案類表 header
  Row 9-13 5 案類資料（竊盜 / 詐欺 / 毒品 / 暴力 / 其他）
  Row 14   小計列
  Row 15   spacer
  Row 16   Section title「下半：員警破獲 Top 10」
  Row 17   spacer
  Row 18   員警表 header
  Row 19-28 10 員警資料
  Row 29   小計列
  Row 30   spacer
  Row 31   頁尾

業務規則：
  - 拿掉 v43 獨立 Tbl刑案，績效全部從 Tbl案件 聚合（SSOT 徹底落實）
  - 5 大案類：竊盜 / 詐欺 / 毒品 / 暴力 / 其他
  - 案類「其他」 = 全般總和 - 前 4 大案類
  - 員警 Top 10 用 LARGE + INDEX/MATCH（非 SORT/FILTER，跨平台 Excel 2010+）

Helper area（隱藏，rows 100-159）：
  - col A: 員警序號 1..60
  - col B: 員警姓名 = IFERROR(INDEX(員警姓名清單, A), "")
  - col C: 全般破獲（COUNTIFS Tbl案件）
  - col D: 竊盜破獲
  - col E: 詐欺破獲
  - col F: 毒品破獲
  - col G: 暴力破獲
  - col H: 其他破獲 = C - D - E - F - G
  - col I: 排名鍵 = C + A*0.0001（唯一化避免同分）
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from phases._base import backup_existing, OUTPUT_PATH, get_logger, unmerge_all_in_sheet
import styles as S


ROC_YEAR = "(YEAR(今日)-1911)"

# 5 大案類聚合（v2.1 升級：用 [案類分類] 精準比對 + 暴力 → 暴力犯罪）
CASE_CATEGORIES = [
    {"name": "竊盜",     "category": "竊盜"},
    {"name": "詐欺",     "category": "詐欺"},
    {"name": "毒品",     "category": "毒品"},
    {"name": "暴力犯罪", "category": "暴力犯罪"},
    # 「其他」由公式算（全般 - 前 4 大），不直接聚合
]

MAX_OFFICERS = 60       # 員警上限
TOP_N = 10               # Top 10 員警

# Helper area 位置
HELPER_START_ROW = 100   # 員警 helper 從 row 100 開始（rows 100-159 = 60 員警）
HELPER_NAME_LIST = "員警姓名清單"


def build_page10_performance(wb, log):
    log.info("--- 頁 10 績效統計 ---")
    ws = wb["績效統計"]

    cleared = unmerge_all_in_sheet(ws)
    log.info(f"  清空既有 merges：{cleared} 個")

    # 欄寬（v2.1 全面加大；非輸入表都要寬一點）
    col_widths = {
        "A": 16,   # 排名 / 序號（容納「🥇 1」、「─ Top 10 合計 ─」等）
        "B": 24,   # 案類 / 員警（容納「暴力犯罪」等 4-5 字 +emoji）
        "C": 16,   # 發生 / 全般破獲
        "D": 16,   # 破獲 / 竊盜
        "E": 18,   # 破獲率 / 詐欺
        "F": 14,   # 排名 / 毒品
        "G": 22,   # 本月變化 / 暴力犯罪（"↑ 3 件" 等）
        "H": 16,   # —      / 其他
    }
    for c, w in col_widths.items():
        ws.column_dimensions[c].width = w

    # === Row 1-2 Banner ===
    ws.merge_cells("A1:H1")
    S.set_cell(ws, "A1", "   📈  績效統計（按案類 + 員警 Top 10）",
               font_key="banner", fill_key="banner",
               align_key="left", border_key="bottom_thick")
    ws.row_dimensions[1].height = S.ROW_HEIGHT["banner"]

    ws.merge_cells("A2:H2")
    sub_formula = (
        '="   "&警察局名稱&" · "&分局名稱&" · "&派出所名稱'
        '&" ｜ 中華民國 "&(YEAR(今日)-1911)&" 年累計（截至 "&TEXT(今日,"mm/dd")&"）｜ 資料源：Tbl案件（SSOT）"'
    )
    S.set_cell(ws, "A2", sub_formula,
               font_key=S.font(11, color="white", italic=True),
               fill_key=S.fill("accent_dark"), align_key="left")
    ws.row_dimensions[2].height = S.ROW_HEIGHT["sub_banner"]

    # === Row 3 spacer ===
    ws.row_dimensions[3].height = 8

    # === Row 4 維護指引 ===
    ws.merge_cells("A4:H4")
    S.set_cell(ws, "A4",
               "   💡  本頁公式全自動：頁 5 案件資料庫每新增 1 件 → 案類聚合 + 員警 Top 10 即時更新。"
               "員警名冊請去設定表 B 區維護（最多 60 名）。",
               font_key=S.font(11, color="warn", italic=True),
               fill_key=S.fill("warn_bg"), align_key="left",
               border_key="all_thin")
    ws.row_dimensions[4].height = 30

    # === Row 5 spacer ===
    ws.row_dimensions[5].height = 8

    # =========================================================
    # 上半：案類績效統計（Row 6-14）
    # =========================================================
    ws.merge_cells("A6:H6")
    S.set_cell(ws, "A6", "   📊  上半：按案類績效統計（5 大案類）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[6].height = S.ROW_HEIGHT["header"]

    ws.row_dimensions[7].height = 6

    # Row 8 header
    headers_case = ["案類", "發生", "破獲", "破獲率", "排名", "本月變化（vs 上月）", "視覺", "—"]
    for i, h in enumerate(headers_case, start=1):
        col = get_column_letter(i)
        S.set_cell(ws, f"{col}8", h,
                   font_key="header", fill_key="header",
                   align_key="center", border_key="all_thin")
    ws.row_dimensions[8].height = S.ROW_HEIGHT["header"]

    # Row 9-12: 4 大案類（竊盜 / 詐欺 / 毒品 / 暴力犯罪）
    case_rate_cells = []        # 用於排名計算（破獲率欄）
    case_value_rows = []
    for idx, cat in enumerate(CASE_CATEGORIES, start=9):
        cat_name = cat["category"]    # v2.1 用 [案類分類] 精準比對

        # A 案類名稱
        S.set_cell(ws, f"A{idx}", cat["name"],
                   font_key="body_bold", fill_key="calc",
                   align_key="left", border_key="all_thin")

        # B 發生（年度累計）— M8 升級：加 [計入發生數]="是"
        S.set_cell(ws, f"B{idx}",
                   f'=IFERROR(COUNTIFS(Tbl案件[案類分類],"{cat_name}",'
                   f'Tbl案件[發生管轄],"本轄",'
                   f'Tbl案件[歸屬年度],{ROC_YEAR},'
                   f'Tbl案件[計入發生數],"是"),0)',
                   font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin",
                   number_format="integer")

        # C 破獲（年度累計，含本所拘提他轄即用查獲管轄）— M8 原則 3：破獲不加 [計入發生數]
        S.set_cell(ws, f"C{idx}",
                   f'=IFERROR(COUNTIFS(Tbl案件[案類分類],"{cat_name}",'
                   f'Tbl案件[查獲管轄],"本轄",'
                   f'Tbl案件[是否破獲],"是",'
                   f'Tbl案件[歸屬年度],{ROC_YEAR}),0)',
                   font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin",
                   number_format="integer")

        # D 破獲率
        S.set_cell(ws, f"D{idx}",
                   f'=IFERROR(C{idx}/B{idx},0)',
                   font_key="body_bold", fill_key="calc",
                   align_key="center", border_key="all_thin",
                   number_format="percent_one")
        case_rate_cells.append(f"D{idx}")

        # E 排名（5 案類間用 RANK.AVG 排序破獲率，降序）
        # 注意：必須等所有 5 列建完才能用 RANK，所以排名公式後面再寫
        # 先預留 cell
        S.set_cell(ws, f"E{idx}", "",
                   font_key="body_bold", fill_key="calc",
                   align_key="center", border_key="all_thin",
                   number_format="integer")

        # F 本月變化（vs 上月）— 升級：用 [案類分類] 精準 + 加 [計入發生數]
        f_curr = (
            f'COUNTIFS(Tbl案件[案類分類],"{cat_name}",'
            f'Tbl案件[發生管轄],"本轄",'
            f'Tbl案件[發生時間],">="&DATE(YEAR(今日),MONTH(今日),1),'
            f'Tbl案件[發生時間],"<"&DATE(YEAR(今日),MONTH(今日)+1,1),'
            f'Tbl案件[計入發生數],"是")'
        )
        f_prev = (
            f'COUNTIFS(Tbl案件[案類分類],"{cat_name}",'
            f'Tbl案件[發生管轄],"本轄",'
            f'Tbl案件[發生時間],">="&DATE(YEAR(今日),MONTH(今日)-1,1),'
            f'Tbl案件[發生時間],"<"&DATE(YEAR(今日),MONTH(今日),1),'
            f'Tbl案件[計入發生數],"是")'
        )
        delta_fml = (
            f'=IFERROR(IF({f_curr}>{f_prev},"↑ "&({f_curr}-{f_prev})&" 件",'
            f'IF({f_curr}<{f_prev},"↓ "&({f_prev}-{f_curr})&" 件","→ 持平")),"—")'
        )
        S.set_cell(ws, f"F{idx}", delta_fml,
                   font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin")

        # G 視覺（橫條 REPT）
        S.set_cell(ws, f"G{idx}",
                   f'=REPT("█",MIN(15,ROUND(D{idx}*15,0)))',
                   font_key=S.font(11, bold=True, color="accent"),
                   fill_key="calc",
                   align_key="left", border_key="all_thin")

        # H 預留（破獲率 ★ 標星：≥100% 標星）
        S.set_cell(ws, f"H{idx}",
                   f'=IF(D{idx}>=1,"★ 滿分",IF(D{idx}>=0.85,"● 達標",""))',
                   font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin")

        ws.row_dimensions[idx].height = S.ROW_HEIGHT["default"]
        case_value_rows.append(idx)

    # Row 13: 「其他」案類 — 全般總和 - 前 4 案類
    other_row = 13
    # A
    S.set_cell(ws, f"A{other_row}", "其他",
               font_key="body_bold", fill_key="calc",
               align_key="left", border_key="all_thin")
    # B 發生 = 全般發生 - SUM(前 4 發生) — 加 [計入發生數]
    all_occur = (
        f'IFERROR(COUNTIFS(Tbl案件[發生管轄],"本轄",'
        f'Tbl案件[歸屬年度],{ROC_YEAR},'
        f'Tbl案件[計入發生數],"是"),0)'
    )
    S.set_cell(ws, f"B{other_row}",
               f'=MAX(0,{all_occur}-SUM(B9:B12))',
               font_key="body", fill_key="calc",
               align_key="center", border_key="all_thin",
               number_format="integer")
    # C 破獲 = 全般破獲 - SUM(前 4 破獲)
    all_solved = (
        f'IFERROR(COUNTIFS(Tbl案件[查獲管轄],"本轄",'
        f'Tbl案件[是否破獲],"是",'
        f'Tbl案件[歸屬年度],{ROC_YEAR}),0)'
    )
    S.set_cell(ws, f"C{other_row}",
               f'=MAX(0,{all_solved}-SUM(C9:C12))',
               font_key="body", fill_key="calc",
               align_key="center", border_key="all_thin",
               number_format="integer")
    # D 破獲率
    S.set_cell(ws, f"D{other_row}",
               f'=IFERROR(C{other_row}/B{other_row},0)',
               font_key="body_bold", fill_key="calc",
               align_key="center", border_key="all_thin",
               number_format="percent_one")
    case_rate_cells.append(f"D{other_row}")
    # E 排名（之後填）
    S.set_cell(ws, f"E{other_row}", "",
               font_key="body_bold", fill_key="calc",
               align_key="center", border_key="all_thin",
               number_format="integer")
    # F 本月變化（不算「其他」，留 "—"）
    S.set_cell(ws, f"F{other_row}", "—",
               font_key=S.font(11, color="muted"),
               fill_key="calc",
               align_key="center", border_key="all_thin")
    # G 視覺
    S.set_cell(ws, f"G{other_row}",
               f'=REPT("█",MIN(15,ROUND(D{other_row}*15,0)))',
               font_key=S.font(11, bold=True, color="accent"),
               fill_key="calc",
               align_key="left", border_key="all_thin")
    # H 標星
    S.set_cell(ws, f"H{other_row}",
               f'=IF(D{other_row}>=1,"★ 滿分",IF(D{other_row}>=0.85,"● 達標",""))',
               font_key="body", fill_key="calc",
               align_key="center", border_key="all_thin")
    ws.row_dimensions[other_row].height = S.ROW_HEIGHT["default"]

    # 現在補上排名欄 E9:E13（5 案類間排名，破獲率降序）
    rate_range = f"$D$9:$D$13"
    for row in range(9, 14):
        rank_fml = f'=IFERROR(RANK.AVG(D{row},{rate_range},0),"—")'
        ws[f"E{row}"] = rank_fml

    # Row 14: 小計列（全 5 案類合計）
    total_row = 14
    S.set_cell(ws, f"A{total_row}", "  ─ 全般合計 ─",
               font_key=S.font(12, bold=True, color="white"),
               fill_key="banner", align_key="left",
               border_key="all_thin")
    S.set_cell(ws, f"B{total_row}",
               f'=SUM(B9:B13)',
               font_key=S.font(12, bold=True, color="white"),
               fill_key="banner", align_key="center",
               border_key="all_thin", number_format="integer")
    S.set_cell(ws, f"C{total_row}",
               f'=SUM(C9:C13)',
               font_key=S.font(12, bold=True, color="white"),
               fill_key="banner", align_key="center",
               border_key="all_thin", number_format="integer")
    S.set_cell(ws, f"D{total_row}",
               f'=IFERROR(C{total_row}/B{total_row},0)',
               font_key=S.font(12, bold=True, color="white"),
               fill_key="banner", align_key="center",
               border_key="all_thin", number_format="percent_one")
    S.set_cell(ws, f"E{total_row}", "—",
               font_key=S.font(12, bold=True, color="white"),
               fill_key="banner", align_key="center",
               border_key="all_thin")
    S.set_cell(ws, f"F{total_row}", "—",
               font_key=S.font(12, bold=True, color="white"),
               fill_key="banner", align_key="center",
               border_key="all_thin")
    S.set_cell(ws, f"G{total_row}",
               f'=REPT("█",MIN(15,ROUND(D{total_row}*15,0)))',
               font_key=S.font(11, bold=True, color="white"),
               fill_key="banner", align_key="left",
               border_key="all_thin")
    S.set_cell(ws, f"H{total_row}",
               f'=IF(D{total_row}>=0.85,"● 達標","")',
               font_key=S.font(12, bold=True, color="white"),
               fill_key="banner", align_key="center",
               border_key="all_thin")
    ws.row_dimensions[total_row].height = S.ROW_HEIGHT["header"]

    # === Row 15 spacer ===
    ws.row_dimensions[15].height = 14

    # =========================================================
    # 下半：員警 Top 10
    # =========================================================
    ws.merge_cells("A16:H16")
    S.set_cell(ws, "A16", "   🏆  下半：員警破獲 Top 10（依全般破獲降序排序）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[16].height = S.ROW_HEIGHT["header"]

    ws.row_dimensions[17].height = 6

    # Row 18 header
    headers_off = ["排名", "員警", "全般破獲", "竊盜", "詐欺", "毒品", "暴力犯罪", "其他"]
    for i, h in enumerate(headers_off, start=1):
        col = get_column_letter(i)
        S.set_cell(ws, f"{col}18", h,
                   font_key="header", fill_key="header",
                   align_key="center", border_key="all_thin")
    ws.row_dimensions[18].height = S.ROW_HEIGHT["header"]

    # Row 19-28: Top 10
    helper_name_range = f"$B${HELPER_START_ROW}:$B${HELPER_START_ROW + MAX_OFFICERS - 1}"
    helper_key_range = f"$I${HELPER_START_ROW}:$I${HELPER_START_ROW + MAX_OFFICERS - 1}"

    for rank in range(1, TOP_N + 1):
        row = 18 + rank
        # A 排名
        rank_label = (
            f'=IF(LARGE({helper_key_range},{rank})>0,'
            f'IF({rank}=1,"🥇 ",IF({rank}=2,"🥈 ",IF({rank}=3,"🥉 ","")))&{rank},"—")'
        )
        S.set_cell(ws, f"A{row}", rank_label,
                   font_key="body_bold", fill_key="calc",
                   align_key="center", border_key="all_thin")

        # MATCH 索引（不直接寫，每欄重算）
        # B 員警姓名 = INDEX(姓名, MATCH(LARGE(key, rank), key, 0))
        match_expr = (
            f'MATCH(LARGE({helper_key_range},{rank}),{helper_key_range},0)'
        )
        S.set_cell(ws, f"B{row}",
                   f'=IFERROR(INDEX({helper_name_range},{match_expr}),"")',
                   font_key="body_bold", fill_key="calc",
                   align_key="left", border_key="all_thin")

        # C..H 各欄破獲（INDEX 各 helper 欄）
        helper_value_cols = ["C", "D", "E", "F", "G", "H"]
        # helper col 對應：C=全般 D=竊盜 E=詐欺 F=毒品 G=暴力 H=其他
        for col_letter in helper_value_cols:
            src_range = (
                f"${col_letter}${HELPER_START_ROW}:${col_letter}"
                f"${HELPER_START_ROW + MAX_OFFICERS - 1}"
            )
            fml = f'=IFERROR(INDEX({src_range},{match_expr}),0)'
            S.set_cell(ws, f"{col_letter}{row}", fml,
                       font_key="body", fill_key="calc",
                       align_key="center", border_key="all_thin",
                       number_format="integer")

        ws.row_dimensions[row].height = S.ROW_HEIGHT["default"]

    # Row 29: 小計列（Top 10 員警合計）
    sum_row = 29
    S.set_cell(ws, f"A{sum_row}", "  ─ Top 10 合計 ─",
               font_key=S.font(12, bold=True, color="white"),
               fill_key="banner", align_key="left",
               border_key="all_thin")
    S.set_cell(ws, f"B{sum_row}", "",
               font_key=S.font(12, bold=True, color="white"),
               fill_key="banner", align_key="center",
               border_key="all_thin")
    for col_letter in ["C", "D", "E", "F", "G", "H"]:
        S.set_cell(ws, f"{col_letter}{sum_row}",
                   f'=SUM({col_letter}19:{col_letter}28)',
                   font_key=S.font(12, bold=True, color="white"),
                   fill_key="banner", align_key="center",
                   border_key="all_thin", number_format="integer")
    ws.row_dimensions[sum_row].height = S.ROW_HEIGHT["header"]

    # === Row 30 spacer ===
    ws.row_dimensions[30].height = 12

    # === Row 31 頁尾 ===
    ws.merge_cells("A31:H31")
    footer = (
        '="© KKEVIN-LIN-2026-V2.0  ｜  資料源：Tbl案件（SSOT，無獨立 Tbl刑案）  ｜  '
        '個別員警深度分析請見頁 11 員警個人卡  ｜  製表："&TEXT(今日,"e/mm/dd")'
    )
    S.set_cell(ws, "A31", footer,
               font_key=S.font(9, color="muted", italic=True),
               fill_key="calc", align_key="center")
    ws.row_dimensions[31].height = 22

    # =========================================================
    # Helper area（隱藏，rows 100-159 = 60 員警）
    # =========================================================
    # Row 99 標頭（debug）
    S.set_cell(ws, "A99", "[helper] 員警聚合",
               font_key=S.font(9, color="muted"))
    helper_headers = ["序號", "員警姓名", "全般破獲", "竊盜", "詐欺", "毒品", "暴力犯罪", "其他", "排名鍵"]
    for i, h in enumerate(helper_headers, start=1):
        col = get_column_letter(i)
        S.set_cell(ws, f"{col}99", h,
                   font_key=S.font(9, color="muted"))

    for k in range(MAX_OFFICERS):
        r = HELPER_START_ROW + k
        seq = k + 1
        # A 序號
        S.set_cell(ws, f"A{r}", seq,
                   font_key=S.font(9, color="muted"),
                   number_format="integer")
        # B 姓名 = INDEX(員警姓名清單, seq)
        S.set_cell(ws, f"B{r}",
                   f'=IFERROR(INDEX({HELPER_NAME_LIST},{seq}),"")',
                   font_key=S.font(9, color="muted"))
        # C 全般破獲
        S.set_cell(ws, f"C{r}",
                   f'=IF(B{r}="",0,IFERROR('
                   f'COUNTIFS(Tbl案件[承辦人],B{r},'
                   f'Tbl案件[是否破獲],"是",'
                   f'Tbl案件[歸屬年度],{ROC_YEAR}),0))',
                   font_key=S.font(9, color="muted"),
                   number_format="integer")
        # D 竊盜 — v2.1：[案類分類] 精準
        S.set_cell(ws, f"D{r}",
                   f'=IF(B{r}="",0,IFERROR('
                   f'COUNTIFS(Tbl案件[承辦人],B{r},'
                   f'Tbl案件[案類分類],"竊盜",'
                   f'Tbl案件[是否破獲],"是",'
                   f'Tbl案件[歸屬年度],{ROC_YEAR}),0))',
                   font_key=S.font(9, color="muted"),
                   number_format="integer")
        # E 詐欺
        S.set_cell(ws, f"E{r}",
                   f'=IF(B{r}="",0,IFERROR('
                   f'COUNTIFS(Tbl案件[承辦人],B{r},'
                   f'Tbl案件[案類分類],"詐欺",'
                   f'Tbl案件[是否破獲],"是",'
                   f'Tbl案件[歸屬年度],{ROC_YEAR}),0))',
                   font_key=S.font(9, color="muted"),
                   number_format="integer")
        # F 毒品
        S.set_cell(ws, f"F{r}",
                   f'=IF(B{r}="",0,IFERROR('
                   f'COUNTIFS(Tbl案件[承辦人],B{r},'
                   f'Tbl案件[案類分類],"毒品",'
                   f'Tbl案件[是否破獲],"是",'
                   f'Tbl案件[歸屬年度],{ROC_YEAR}),0))',
                   font_key=S.font(9, color="muted"),
                   number_format="integer")
        # G 暴力犯罪 — v2.1：分類名改為「暴力犯罪」
        S.set_cell(ws, f"G{r}",
                   f'=IF(B{r}="",0,IFERROR('
                   f'COUNTIFS(Tbl案件[承辦人],B{r},'
                   f'Tbl案件[案類分類],"暴力犯罪",'
                   f'Tbl案件[是否破獲],"是",'
                   f'Tbl案件[歸屬年度],{ROC_YEAR}),0))',
                   font_key=S.font(9, color="muted"),
                   number_format="integer")
        # H 其他 = C - D - E - F - G
        S.set_cell(ws, f"H{r}",
                   f'=MAX(0,C{r}-D{r}-E{r}-F{r}-G{r})',
                   font_key=S.font(9, color="muted"),
                   number_format="integer")
        # I 排名鍵 = C + seq*0.0001（避免同分）
        # 若 B 空，鍵 = 0（不入榜）
        S.set_cell(ws, f"I{r}",
                   f'=IF(B{r}="",0,C{r}+{seq}*0.0001)',
                   font_key=S.font(9, color="muted"),
                   number_format="0.0000")

    # 隱藏 helper rows 99-160
    for r in range(99, 161):
        ws.row_dimensions[r].hidden = True

    # 公式計數
    fcount = 0
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if v and isinstance(v, str) and v.startswith("="):
                fcount += 1

    log.info(f"  案類表：5 案類 + 全般合計")
    log.info(f"  員警 Top 10 + 小計")
    log.info(f"  Helper：員警 {MAX_OFFICERS} 列 × 9 欄聚合（rows 99-160 隱藏）")
    log.info(f"  公式總數：{fcount}")
    log.info(f"  凍結窗格：A8")

    return fcount


def build():
    log = get_logger("phase3_p10")
    log.info("===== Phase 3 — 頁 10 績效統計 開始 =====")

    wb = load_workbook(OUTPUT_PATH)
    fcount = build_page10_performance(wb, log)

    bak = backup_existing()
    if bak:
        log.info(f"  舊版備份：{bak}")
    wb.save(OUTPUT_PATH)
    size_kb = OUTPUT_PATH.stat().st_size / 1024
    log.info(f"  儲存：{OUTPUT_PATH.name} ({size_kb:.1f} KB)")
    log.info(f"  公式：{fcount}")
    log.info("===== Phase 3 — 頁 10 完成 =====\n")
    return OUTPUT_PATH


if __name__ == "__main__":
    build()
