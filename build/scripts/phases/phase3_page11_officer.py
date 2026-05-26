"""
phase3_page11_officer.py — Phase 3：頁 11 員警個人績效卡（A 麥肯錫風）

版型（A 直式一頁，可獨立列印 A4 給督導官）：
  Row 1-2  Banner
  Row 3    [員警下拉] + 基本資料條（職稱 / 在職 / 排名提示）
  Row 4    spacer
  Row 5    4 KPI 卡 labels（總破獲 / 全所排名 / 全所佔比 / 本月新增）
  Row 6    4 KPI 卡 values（大字）
  Row 7    4 KPI 卡 sub text
  ─── FREEZE A8 ───
  Row 8    Section「📈 個人 vs 全所平均比較表」
  Row 9    比較表 header
  Row 10-14 5 比較項目（總破獲 / 竊盜 / 詐欺 / 毒品 / 暴力）
  Row 15   spacer
  Row 16   Section「📋 本年度辦案清單」
  Row 17   辦案清單 header
  Row 18-47 30 列辦案（AGGREGATE + INDEX，2019- Fallback）
  Row 48   spacer
  Row 49   簽章區（所長 / 督導官）
  Row 50   頁尾

切換機制：員警下拉 B3，DataValidation 列表來自員警姓名清單
        所有公式引用 $B$3 切換

Fallback 設計：
  辦案清單不用 FILTER（Excel 365）
  用 AGGREGATE(15, 6, ROW/條件, k) 取第 k 個符合列號
  - AGGREGATE option 6 = 忽略錯誤值（除以 0 會產生 #DIV/0，正好被忽略）
  - 條件為布林，TRUE/FALSE → 1/0，分母 0 時除錯被忽略
  - 兼容 Excel 2010+ 全版本（含 Mac、LibreOffice）

Tbl案件 header row 為 14（rough hardcode；變動時需同步更新此檔）
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from phases._base import backup_existing, OUTPUT_PATH, get_logger, unmerge_all_in_sheet
import styles as S


SHEET = "員警個人績效卡"
OFFICER_CELL = "B3"      # 員警下拉位置
ROC_YEAR = "(YEAR(今日)-1911)"

# Tbl案件 header row（Tbl案件 註冊在 case_database 表 row 14）
CASE_HEADER_ROW = 14
CASE_TBL = "Tbl案件"

CASE_LIST_ROWS = 30      # 顯示辦案清單行數


def build_page11_officer(wb, log):
    log.info("--- 頁 11 員警個人績效卡 ---")
    ws = wb[SHEET]

    cleared = unmerge_all_in_sheet(ws)
    log.info(f"  清空既有 merges：{cleared} 個")

    # === 欄寬（v2.1 全面加大）===
    col_widths = {
        "A": 14,   # 序號 / KPI 卡 label
        "B": 24,   # 員警下拉 / 案類（容納「暴力犯罪」）
        "C": 18,   # 發生日 / 基本資料
        "D": 30,   # 發生地
        "E": 18,   # 破獲日
        "F": 18,   # 案件狀況
        "G": 16,   # 涉案金額
        "H": 30,   # 備註
    }
    for c, w in col_widths.items():
        ws.column_dimensions[c].width = w

    # === Row 1-2 Banner ===
    ws.merge_cells("A1:H1")
    S.set_cell(ws, "A1", "   👮  員警個人績效卡（A 麥肯錫風 — A4 直式可獨立列印）",
               font_key="banner", fill_key="banner",
               align_key="left", border_key="bottom_thick")
    ws.row_dimensions[1].height = S.ROW_HEIGHT["banner"]

    ws.merge_cells("A2:H2")
    sub_formula = (
        '="   "&警察局名稱&" · "&分局名稱&" · "&派出所名稱'
        '&" ｜ 中華民國 "&(YEAR(今日)-1911)&" 年度績效卡 ｜ 製表："&TEXT(今日,"e/mm/dd")'
    )
    S.set_cell(ws, "A2", sub_formula,
               font_key=S.font(11, color="white", italic=True),
               fill_key=S.fill("accent_dark"), align_key="left")
    ws.row_dimensions[2].height = S.ROW_HEIGHT["sub_banner"]

    # === Row 3 員警下拉 + 基本資料條 ===
    # A3 label
    S.set_cell(ws, "A3", "  選擇員警：",
               font_key=S.font(13, bold=True, color="accent"),
               fill_key="accent_light",
               align_key="right", border_key="all_thin")
    # B3 員警下拉（DataValidation）
    S.set_cell(ws, OFFICER_CELL, "王小明",  # 預設選第一位示範員警
               font_key=S.font(15, bold=True, color="white"),
               fill_key="banner",
               align_key="center", border_key="all_thin")
    # 加下拉
    dv = DataValidation(type="list", formula1="=員警姓名清單", allow_blank=False)
    dv.add(OFFICER_CELL)
    dv.error = "請選擇有效員警"
    dv.errorTitle = "員警未在名冊"
    dv.prompt = "請從下拉選擇員警（來源：設定表 B 區員警名冊）"
    dv.promptTitle = "員警下拉"
    ws.add_data_validation(dv)

    # C3:H3 基本資料條
    # C3 職稱（用 INDEX/MATCH 從設定表 B 區員警名冊抓）
    # 設定表員警名冊在 row 13+，col B=姓名, col C=職稱
    # MATCH 姓名 → 對應的職稱
    ws.merge_cells("C3:D3")
    title_fml = (
        '="職稱：" & IFERROR(INDEX(設定表!$C$13:$C$72,'
        f'MATCH({OFFICER_CELL},員警姓名清單,0)),"—")'
    )
    S.set_cell(ws, "C3", title_fml,
               font_key=S.font(12, bold=True, color="accent"),
               fill_key="accent_light",
               align_key="left", border_key="all_thin")

    # E3:F3 在職狀態
    ws.merge_cells("E3:F3")
    active_fml = (
        '="在職：" & IFERROR(INDEX(設定表!$D$13:$D$72,'
        f'MATCH({OFFICER_CELL},員警姓名清單,0)),"—")'
    )
    S.set_cell(ws, "E3", active_fml,
               font_key=S.font(12, bold=True, color="accent"),
               fill_key="accent_light",
               align_key="left", border_key="all_thin")

    # G3:H3 提示
    ws.merge_cells("G3:H3")
    S.set_cell(ws, "G3",
               '="📌 完整名冊請至設定表 B 區"',
               font_key=S.font(10, color="muted", italic=True),
               fill_key="accent_light",
               align_key="right", border_key="all_thin")
    ws.row_dimensions[3].height = 32

    # === Row 4 spacer ===
    ws.row_dimensions[4].height = 6

    # === Row 5-7 大字 KPI 4 卡 ===
    # 4 卡：A:B / C:D / E:F / G:H
    # Card 1: 總破獲 / Card 2: 全所排名 / Card 3: 全所佔比 / Card 4: 本月新增
    common = (
        f'COUNTIFS({CASE_TBL}[承辦人],{OFFICER_CELL},'
        f'{CASE_TBL}[是否破獲],"是",'
        f'{CASE_TBL}[歸屬年度],{ROC_YEAR})'
    )
    all_solved = (
        f'COUNTIFS({CASE_TBL}[查獲管轄],"本轄",'
        f'{CASE_TBL}[是否破獲],"是",'
        f'{CASE_TBL}[歸屬年度],{ROC_YEAR})'
    )
    this_month = (
        f'COUNTIFS({CASE_TBL}[承辦人],{OFFICER_CELL},'
        f'{CASE_TBL}[是否破獲],"是",'
        f'{CASE_TBL}[破獲時間],">="&DATE(YEAR(今日),MONTH(今日),1),'
        f'{CASE_TBL}[破獲時間],"<"&DATE(YEAR(今日),MONTH(今日)+1,1))'
    )

    kpi_cards = [
        # (cols_range, label, fg, bg, value_formula, fmt, sub)
        (("A5:B5", "A6:B6", "A7:B7"),
            "🏆 總破獲", "accent", "accent_light",
            f'=IFERROR({common},0)', "integer",
            "本年度累計"),
        (("C5:D5", "C6:D6", "C7:D7"),
            "📊 全所排名", "warn", "warn_bg",
            # 排名：在頁 10 helper 算過。直接用該頁 helper area
            # 簡化：用 SUMPRODUCT 算「破獲數比我多的員警數」+1
            # 或從頁 10 helper 拉
            f'=IFERROR(SUMPRODUCT(--(績效統計!$C$100:$C$159>IFERROR({common},0))*'
            f'(績效統計!$B$100:$B$159<>""))+1,"—")', "integer",
            "全員警比較"),
        (("E5:F5", "E6:F6", "E7:F7"),
            "📈 全所佔比", "pass", "pass_bg",
            f'=IFERROR({common}/{all_solved},0)', "percent_one",
            "個人 / 全所破獲"),
        (("G5:H5", "G6:H6", "G7:H7"),
            "🆕 本月新增", "accent_dark", "accent_light",
            f'=IFERROR({this_month},0)', "integer",
            "本月破獲"),
    ]
    for (lab_rng, val_rng, sub_rng), label, fg, bg, fml, fmt, sub in kpi_cards:
        ws.merge_cells(lab_rng)
        S.set_cell(ws, lab_rng.split(":")[0], f"  {label}",
                   font_key=S.font(11, bold=True, color=fg),
                   fill_key=S.fill(bg),
                   align_key="left", border_key="top_bottom")
        ws.merge_cells(val_rng)
        S.set_cell(ws, val_rng.split(":")[0], fml,
                   font_key=S.font(28, bold=True, color=fg),
                   fill_key=S.fill(bg),
                   align_key="center", number_format=fmt)
        ws.merge_cells(sub_rng)
        S.set_cell(ws, sub_rng.split(":")[0], f"  {sub}",
                   font_key=S.font(9, color="muted"),
                   fill_key=S.fill(bg),
                   align_key="left", border_key="bottom_thick")
    ws.row_dimensions[5].height = 24
    ws.row_dimensions[6].height = 48
    ws.row_dimensions[7].height = 20

    # === Row 8 Section「個人 vs 全所平均比較」===
    ws.merge_cells("A8:H8")
    S.set_cell(ws, "A8", "   📈  個人 vs 全所平均比較表（按案類）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[8].height = S.ROW_HEIGHT["header"]

    # === Row 9 比較表 header ===
    compare_headers = ["案類", "個人破獲", "全所平均", "差距", "評級", "—", "—", "—"]
    for i, h in enumerate(compare_headers, start=1):
        col = get_column_letter(i)
        S.set_cell(ws, f"{col}9", h,
                   font_key="header", fill_key="header",
                   align_key="center", border_key="all_thin")
    ws.row_dimensions[9].height = S.ROW_HEIGHT["header"]

    # === Row 10-14 5 比較項目 ===
    # 全所員警數 (在職)：COUNTIF(設定表!$D$13:$D$72, "是")
    officers_active = 'COUNTIF(設定表!$D$13:$D$72,"是")'

    # v2.1 升級：使用 [案類分類] 精準比對；「暴力」→「暴力犯罪」
    compare_rows = [
        # (label, category_name, all_count_formula)
        ("總破獲",
            None,
            f'COUNTIFS({CASE_TBL}[查獲管轄],"本轄",{CASE_TBL}[是否破獲],"是",{CASE_TBL}[歸屬年度],{ROC_YEAR})'),
        ("竊盜",
            "竊盜",
            f'COUNTIFS({CASE_TBL}[案類分類],"竊盜",{CASE_TBL}[查獲管轄],"本轄",{CASE_TBL}[是否破獲],"是",{CASE_TBL}[歸屬年度],{ROC_YEAR})'),
        ("詐欺",
            "詐欺",
            f'COUNTIFS({CASE_TBL}[案類分類],"詐欺",{CASE_TBL}[查獲管轄],"本轄",{CASE_TBL}[是否破獲],"是",{CASE_TBL}[歸屬年度],{ROC_YEAR})'),
        ("毒品",
            "毒品",
            f'COUNTIFS({CASE_TBL}[案類分類],"毒品",{CASE_TBL}[查獲管轄],"本轄",{CASE_TBL}[是否破獲],"是",{CASE_TBL}[歸屬年度],{ROC_YEAR})'),
        ("暴力犯罪",
            "暴力犯罪",
            f'COUNTIFS({CASE_TBL}[案類分類],"暴力犯罪",{CASE_TBL}[查獲管轄],"本轄",{CASE_TBL}[是否破獲],"是",{CASE_TBL}[歸屬年度],{ROC_YEAR})'),
    ]
    for i, (label, cat_name, all_fml) in enumerate(compare_rows, start=10):
        # A 案類名
        S.set_cell(ws, f"A{i}", label,
                   font_key="body_bold", fill_key="calc",
                   align_key="left", border_key="all_thin")

        # B 個人破獲
        if cat_name is None:
            personal = (
                f'COUNTIFS({CASE_TBL}[承辦人],{OFFICER_CELL},'
                f'{CASE_TBL}[是否破獲],"是",'
                f'{CASE_TBL}[歸屬年度],{ROC_YEAR})'
            )
        else:
            personal = (
                f'COUNTIFS({CASE_TBL}[承辦人],{OFFICER_CELL},'
                f'{CASE_TBL}[案類分類],"{cat_name}",'
                f'{CASE_TBL}[是否破獲],"是",'
                f'{CASE_TBL}[歸屬年度],{ROC_YEAR})'
            )
        S.set_cell(ws, f"B{i}",
                   f'=IFERROR({personal},0)',
                   font_key=S.font(12, bold=True, color="accent"),
                   fill_key="calc",
                   align_key="center", border_key="all_thin",
                   number_format="integer")

        # C 全所平均 = 全所破獲 / 在職員警數
        S.set_cell(ws, f"C{i}",
                   f'=IFERROR({all_fml}/{officers_active},0)',
                   font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin",
                   number_format="0.0")

        # D 差距 = 個人 - 全所平均
        S.set_cell(ws, f"D{i}",
                   f'=IFERROR(B{i}-C{i},0)',
                   font_key="body_bold", fill_key="calc",
                   align_key="center", border_key="all_thin",
                   number_format="+0.0;-0.0;0.0")

        # E 評級
        S.set_cell(ws, f"E{i}",
                   f'=IFERROR(IF(B{i}>=C{i}*1.5,"★ 優異",'
                   f'IF(B{i}>=C{i},"● 達標",'
                   f'IF(B{i}>=C{i}*0.5,"⚠ 待加強","● 落後"))),"—")',
                   font_key="body_bold", fill_key="calc",
                   align_key="center", border_key="all_thin")

        # F-H 留白
        for col in ["F", "G", "H"]:
            S.set_cell(ws, f"{col}{i}", "",
                       fill_key="calc", border_key="all_thin")

        ws.row_dimensions[i].height = S.ROW_HEIGHT["default"]

    # === Row 15 spacer ===
    ws.row_dimensions[15].height = 12

    # === Row 16 Section「本年度辦案清單」===
    ws.merge_cells("A16:H16")
    S.set_cell(ws, "A16",
               "   📋  本年度辦案清單（只列已破獲案件，依破獲日期序）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[16].height = S.ROW_HEIGHT["header"]

    # === Row 17 辦案清單 header ===
    list_headers = ["序", "案類", "發生日", "發生地", "破獲日", "案件狀況", "涉案金額", "備註"]
    for i, h in enumerate(list_headers, start=1):
        col = get_column_letter(i)
        S.set_cell(ws, f"{col}17", h,
                   font_key="header", fill_key="header",
                   align_key="center", border_key="all_thin")
    ws.row_dimensions[17].height = S.ROW_HEIGHT["header"]

    # === Row 18-47 30 列辦案 (AGGREGATE+INDEX Fallback) ===
    # 條件：承辦人 = OFFICER_CELL AND 是否破獲 = "是" AND 歸屬年度 = 當年
    # AGGREGATE(15, 6, ROW/condition, k) — k 從 1 開始
    # 取出 k 個符合的 sheet row (絕對 row 號)，減 header_row 取 table index
    # Tbl案件 第 i 列 = 表內 (i - 14) 列
    #
    # ROW(Tbl案件[編號]) = 此欄各 cell 的 sheet row (15-1014)
    # 條件 = (承辦人=B3)*(是否破獲="是")*(歸屬年度=當年)
    # 公式：AGGREGATE(15,6,ROW(Tbl案件[編號])/條件,k) → 第 k 個 sheet row
    # 再 - CASE_HEADER_ROW = table-relative row
    # INDEX(Tbl案件[案類], 該 index) 取值

    case_columns = [
        # (col_letter, src_field, fmt)
        ("B", "案類", "general"),
        ("C", "發生時間", "date_roc"),
        ("D", "發生地點", "general"),
        ("E", "破獲時間", "date_roc"),
        ("F", "案件狀況", "general"),
        ("G", "涉案金額", "money"),
        ("H", "備註", "general"),
    ]

    cond_expr = (
        f'(({CASE_TBL}[承辦人]={OFFICER_CELL})'
        f'*({CASE_TBL}[是否破獲]="是")'
        f'*({CASE_TBL}[歸屬年度]={ROC_YEAR}))'
    )

    for k in range(1, CASE_LIST_ROWS + 1):
        row = 17 + k       # 18-47

        # A 序號（顯示用，僅當該列有資料才顯示）
        # 若 AGGREGATE 返回有效 row 才顯示 k
        S.set_cell(ws, f"A{row}",
                   f'=IFERROR(IF(AGGREGATE(15,6,'
                   f'ROW({CASE_TBL}[編號])/{cond_expr},{k})>0,{k},""),"")',
                   font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin",
                   number_format="integer")

        # 各欄
        for col, field, fmt in case_columns:
            fml = (
                f'=IFERROR(INDEX({CASE_TBL}[{field}],'
                f'AGGREGATE(15,6,ROW({CASE_TBL}[編號])/{cond_expr},{k})'
                f'-{CASE_HEADER_ROW}),"")'
            )
            S.set_cell(ws, f"{col}{row}", fml,
                       font_key="body", fill_key="calc",
                       align_key="left" if col in ("B", "D", "H") else "center",
                       border_key="all_thin",
                       number_format=fmt)
        ws.row_dimensions[row].height = S.ROW_HEIGHT["default"]

    # === Row 48 spacer ===
    ws.row_dimensions[48].height = 12

    # === Row 49 簽章區 ===
    ws.merge_cells("A49:D49")
    S.set_cell(ws, "A49", "  所長：__________________  簽章",
               font_key=S.font(11, color="text"),
               fill_key=None, align_key="left",
               border_key="bottom_medium")
    ws.merge_cells("E49:H49")
    S.set_cell(ws, "E49", "  督導官：__________________  簽章",
               font_key=S.font(11, color="text"),
               fill_key=None, align_key="left",
               border_key="bottom_medium")
    ws.row_dimensions[49].height = 32

    # === Row 50 頁尾 ===
    ws.merge_cells("A50:H50")
    footer = (
        '="© KKEVIN-LIN-2026-V2.0  ｜  個人卡資料源：Tbl案件（SSOT）｜  '
        '製表："&TEXT(今日,"e/mm/dd")&"  ｜  A4 直式可獨立列印交付督導官"'
    )
    S.set_cell(ws, "A50", footer,
               font_key=S.font(9, color="muted", italic=True),
               fill_key="calc", align_key="center")
    ws.row_dimensions[50].height = 22

    # 公式計數
    fcount = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                fcount += 1
    log.info(f"  員警下拉：{OFFICER_CELL}（DataValidation = 員警姓名清單）")
    log.info(f"  4 KPI 卡：總破獲 / 全所排名 / 全所佔比 / 本月新增")
    log.info(f"  比較表：5 案類（總破獲 / 竊盜 / 詐欺 / 毒品 / 暴力）")
    log.info(f"  辦案清單：{CASE_LIST_ROWS} 列（AGGREGATE 2010+ Fallback）")
    log.info(f"  公式總數：{fcount}")
    log.info(f"  凍結窗格：A8（rows 1-7 留在頂部）")
    return fcount


def build():
    log = get_logger("phase3_p11")
    log.info("===== Phase 3 — 頁 11 員警個人卡 開始 =====")

    wb = load_workbook(OUTPUT_PATH)
    fcount = build_page11_officer(wb, log)

    bak = backup_existing()
    if bak:
        log.info(f"  舊版備份：{bak}")
    wb.save(OUTPUT_PATH)
    size_kb = OUTPUT_PATH.stat().st_size / 1024
    log.info(f"  儲存：{OUTPUT_PATH.name} ({size_kb:.1f} KB)")
    log.info(f"  公式：{fcount}")
    log.info("===== Phase 3 — 頁 11 完成 =====\n")
    return OUTPUT_PATH


if __name__ == "__main__":
    build()
