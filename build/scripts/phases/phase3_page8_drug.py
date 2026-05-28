"""
phase3_page8_drug.py — Phase 3：頁 8 毒品調驗人口管制（B 公文風 — 輕量主表 + 雙翼升級）

包含兩部分：
  Part A：建立 Tbl毒調（v43 9 欄維持不動 + v2 升級 3 個計算欄）
  Part B：頁 8 layout（頂部 4 段燈 + 毒調率 KPI 條 + 主表 + 70/30 候選清單）
  Part C：回頭把頁 2 KPI 4 毒調率 / KPI 5 未到驗 接上 Tbl毒調

業務規則（plan 校正 2026-05-24）：
  - 通緝中、已聲強採、在監 三狀態都算到驗（驗不到的也算）
  - 毒調率 = (已驗 + 通緝 + 強採 + 在監) / 列管總數
  - 應強採候選 = 只列「未到驗」（通緝/強採/在監 不入候選）

Tbl毒調 12 欄：
  1 編號 / 2 姓名 / 3 身分證 / 4 性別 / 5 住所 / 6 聯繫 / 7 通緝 / 8 管制情形 / 9 備註
  10 是否到驗（計算欄，0/1）
  11 狀態燈（計算欄，● 綠 / ⚠ 紅 / ○ 結束）
  12 候選旗標（計算欄，1 if 未到驗 else 0）

頁面 layout：
  Cols A-I  主表 9 欄（v43 不動）
  Cols J-L  Tbl毒調 計算欄（隱藏）
  Col  M    spacer
  Cols N-Q  應強採候選清單 4 欄（編號 / 姓名 / 狀態 / 操作建議）

  Row 1-2   Banner
  Row 4     維護指引
  Row 6     Section title「📊 狀態縮影 + 毒調率」(A:I) ｜「⚠ 應強採候選」(N:Q)
  Row 8-9   4 status cards（列管總數 / 已到驗 / 未到驗 / 強採候選）
  Row 11    毒調率 大字 KPI
  Row 12    毒調率 進度條 (REPT)
  Row 13    Section title「📋 列管人口主表」(A:I)
  Row 14    主表 + 候選 header  ← FREEZE A14
  Row 15-84 Tbl毒調 資料 + 候選清單前 16 列
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo, TableColumn, TableFormula
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from phases._base import backup_existing, OUTPUT_PATH, get_logger, unmerge_all_in_sheet
import styles as S


ROC_YEAR = "(YEAR(今日)-1911)"

DRUG_SHEET = "毒品調驗人口管制"
DRUG_TABLE = "Tbl毒調"
DRUG_HEADER_ROW = 14
DRUG_DATA_START = 15
DRUG_CAPACITY = 70                 # 70 列 ≈ 一般派出所列管上限
DRUG_DATA_END = DRUG_DATA_START + DRUG_CAPACITY - 1  # 84

# 14 欄定義（v2.2 加最後到驗日 + 距今天數 精確 30 日逾期）
DRUG_COLUMNS = [
    # (header, width, dv_named_range, kind, format)
    ("編號",       8, None,                  "input", "general"),
    ("姓名",       12, None,                 "input", "general"),
    ("身分證",     14, None,                 "input", "general"),
    ("性別",       6, "性別清單",            "input", "general"),
    ("住所",       22, None,                 "input", "general"),
    ("聯繫",       14, None,                 "input", "general"),
    ("通緝",       8, "通緝清單",            "input", "general"),
    ("管制情形",   16, "毒品管制情形清單",   "input", "general"),
    ("備註",       20, None,                 "input", "general"),
    # v2.2 新增：最後到驗日（input, datetime）+ 距今天數（calc）
    ("最後到驗日", 14, None, "input", "date_roc"),
    ("距今天數",   10, None, "calc", "integer"),
    # 計算欄（隱藏）— 業務規則：通緝/強採/在監/已驗 都算到驗
    ("是否到驗",   10, None, "calc", "integer"),
    ("狀態燈",     10, None, "calc", "general"),
    ("候選旗標",   10, None, "calc", "integer"),
]

from datetime import datetime as _dt

# 5 筆示範資料（v2.2 加最後到驗日）
DRUG_SAMPLE = [
    ["001", "王小美", "A123456789", "女", "中正路 12 號", "0912-345-678",
     "否", "已到驗（5/3）", "", _dt(2026, 5, 3)],
    ["002", "李志明", "B234567890", "男", "和平路 56 號", "0923-456-789",
     "否", "未到驗（5/3 應到）", "需聲請強採書", _dt(2026, 5, 3)],
    ["003", "陳大山", "C345678901", "男", "復興路 78 號", "通緝",
     "是", "通緝中", "失聯，算到驗", None],  # 通緝中沒到驗日
    ["004", "張小宇", "D456789012", "男", "民生路 99 號", "0934-567-890",
     "否", "已聲強採（5/15）", "強採書已聲請", _dt(2026, 5, 15)],
    ["005", "林月華", "E567890123", "女", "中山路 33 號", "在監",
     "否", "在監", "高雄監獄", None],  # 在監沒到驗日
]

# 候選清單顯示行數
CANDIDATE_ROWS = 16


def _drug_calc_formulas(r):
    """v2.2 第二版修：Tbl毒調 計算欄純儲存格參照（防 #REF! 地雷）。

    欄位對應：H=管制情形 J=最後到驗日 K=距今天數 L=是否到驗
    回傳 (K 距今天數, L 是否到驗, M 狀態燈, N 候選旗標)
    """
    K = f'=IFERROR(IF(J{r}="","",今日-J{r}),"")'
    L = (f'=IFERROR(IF(H{r}="",0,IF(OR('
         f'SUMPRODUCT(--ISNUMBER(SEARCH({{"已驗","已到驗","通緝","強採","在監"}},H{r})))>0,'
         f'AND(ISNUMBER(K{r}),K{r}<=30)),1,0)),0)')
    M = (f'=IFERROR(IF(H{r}="","○ 未填",'
         f'IF(ISNUMBER(SEARCH("解除",H{r})),"○ 結束",'
         f'IF(AND(ISNUMBER(K{r}),K{r}>30),"● 紅",'
         f'IF(L{r}=1,"● 綠","⚠ 黃")))),"")')
    N = (f'=IFERROR(IF(OR(ISNUMBER(SEARCH("未到驗",H{r})),'
         f'AND(ISNUMBER(K{r}),K{r}>30)),1,0),0)')
    return K, L, M, N


def build_tbl_drug(wb, log):
    """Part A: 建立 Tbl毒調"""
    ws = wb[DRUG_SHEET]
    log.info(f"  [A] 建 Tbl毒調 ({DRUG_CAPACITY} 列容量)")

    cleared = unmerge_all_in_sheet(ws)
    log.info(f"  清空既有 merges：{cleared} 個")

    # === 設定欄寬（v2.2：14 欄 + 候選清單往後挪 2 欄）===
    visible_widths = {
        "A": 10, "B": 14, "C": 16, "D": 8, "E": 24, "F": 16, "G": 10, "H": 18,
        "I": 22,
        "J": 14,                                 # J 最後到驗日（input, 顯示）
        "K": 10, "L": 10, "M": 10, "N": 10,     # K-N 計算欄稍後隱藏
        "O": 3,                                  # spacer
        "P": 10, "Q": 16, "R": 16, "S": 24,    # 候選清單
    }
    for c, w in visible_widths.items():
        ws.column_dimensions[c].width = w

    # === Header row 14（Tbl毒調 12 欄）===
    for i, (h, w, dv, kind, fmt) in enumerate(DRUG_COLUMNS, start=1):
        col = get_column_letter(i)
        S.set_cell(ws, f"{col}{DRUG_HEADER_ROW}", h,
                   font_key="header", fill_key="header",
                   align_key="center", border_key="all_thin")
    ws.row_dimensions[DRUG_HEADER_ROW].height = S.ROW_HEIGHT["header"]

    # === Sample 5 列資料（rows 15-19）+ 計算欄公式 ===
    # v2.2：14 欄（9 input + 1 日期 input + 4 calc）
    #   J 最後到驗日（input, 樣本帶 datetime）
    #   K 距今天數（calc）/ L 是否到驗 / M 狀態燈 / N 候選旗標
    for r_idx, sample in enumerate(DRUG_SAMPLE, start=DRUG_DATA_START):
        for c_idx, val in enumerate(sample, start=1):
            col = get_column_letter(c_idx)
            fmt = DRUG_COLUMNS[c_idx - 1][4]   # date_roc for J
            S.set_cell(ws, f"{col}{r_idx}", val,
                       font_key="body", fill_key="input",
                       align_key="left" if c_idx <= 9 else "center",
                       border_key="all_thin",
                       number_format=fmt if fmt != "general" else None)
        # K/L/M/N 計算欄（v2.2 第二版：純參照防 #REF!）
        fK, fL, fM, fN = _drug_calc_formulas(r_idx)
        S.set_cell(ws, f"K{r_idx}", fK, font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin", number_format="integer")
        S.set_cell(ws, f"L{r_idx}", fL, font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin", number_format="integer")
        S.set_cell(ws, f"M{r_idx}", fM, font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin")
        S.set_cell(ws, f"N{r_idx}", fN, font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin", number_format="integer")
        ws.row_dimensions[r_idx].height = S.ROW_HEIGHT["default"]

    # === 空白列 (預留容量，v2.2：input cols 1-10 含最後到驗日；calc K-N)===
    for r_idx in range(DRUG_DATA_START + len(DRUG_SAMPLE), DRUG_DATA_END + 1):
        # Input cols 1-10（含最後到驗日 J）：空白底色
        for c_idx in range(1, 11):
            col = get_column_letter(c_idx)
            fmt = DRUG_COLUMNS[c_idx - 1][4]
            S.set_cell(ws, f"{col}{r_idx}", None,
                       font_key="body", fill_key="input",
                       align_key="left" if c_idx <= 9 else "center",
                       border_key="all_thin",
                       number_format=fmt if fmt != "general" else None)
        # K/L/M/N 計算欄（v2.2 第二版：純參照防 #REF!）
        fK, fL, fM, fN = _drug_calc_formulas(r_idx)
        S.set_cell(ws, f"K{r_idx}", fK, font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin", number_format="integer")
        S.set_cell(ws, f"L{r_idx}", fL, font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin", number_format="integer")
        S.set_cell(ws, f"M{r_idx}", fM, font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin")
        S.set_cell(ws, f"N{r_idx}", fN, font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin", number_format="integer")
        ws.row_dimensions[r_idx].height = S.ROW_HEIGHT["default"]

    # === 註冊 Excel Table ===（v2.2 擴 12→14 欄）
    # ★ v2.2 第二版修：計算欄改純參照（見 _drug_calc_formulas），
    #   table 不加 calculatedColumnFormula（會被 Excel 判定不合法移除 table）
    table_ref = f"A{DRUG_HEADER_ROW}:N{DRUG_DATA_END}"
    tbl = Table(displayName=DRUG_TABLE, name=DRUG_TABLE, ref=table_ref)
    tbl.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium5",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tbl)

    # === 隱藏計算欄 K/L/M/N（最後到驗日 J 不隱藏 — 使用者要填）===
    for col in ["K", "L", "M", "N"]:
        ws.column_dimensions[col].hidden = True

    # === 資料驗證（下拉清單）===
    dv_specs = [
        ("D", "性別清單"),
        ("G", "通緝清單"),
        ("H", "毒品管制情形清單"),
    ]
    for col, named in dv_specs:
        dv = DataValidation(type="list", formula1=f"={named}", allow_blank=True)
        dv.add(f"{col}{DRUG_DATA_START}:{col}{DRUG_DATA_END}")
        ws.add_data_validation(dv)

    # 最後到驗日（J）— 日期格式驗證
    dv_date = DataValidation(
        type="date", operator="between",
        formula1="DATE(2020,1,1)", formula2="DATE(2035,12,31)",
        allow_blank=True, showErrorMessage=True,
        errorTitle="日期格式錯誤",
        error="最後到驗日必須是有效日期\n範例：2026/05/03"
    )
    dv_date.add(f"J{DRUG_DATA_START}:J{DRUG_DATA_END}")
    ws.add_data_validation(dv_date)

    log.info(f"    Tbl毒調 v2.2 註冊：{table_ref}（14 欄：9 input + 1 日期 + 4 calc）")
    log.info(f"    計算欄 K/L/M/N 已隱藏；下拉 D/G/H + 日期 J 已設驗證")


def build_page8_layout(wb, log):
    """Part B: 頁 8 layout（status cards + 毒調率 + 候選清單）"""
    ws = wb[DRUG_SHEET]
    log.info("  [B] 建頁 8 layout")

    # === Row 1-2 Banner ===（v2.2：擴 A:S 因 Tbl毒調 多 2 欄）
    ws.merge_cells("A1:S1")
    S.set_cell(ws, "A1",
               "   💊  毒品調驗人口管制（v43 9 欄 + v2.2 加最後到驗日精確 30 日逾期）",
               font_key="banner", fill_key="banner",
               align_key="left", border_key="bottom_thick")
    ws.row_dimensions[1].height = S.ROW_HEIGHT["banner"]

    ws.merge_cells("A2:S2")
    sub_formula = (
        '="   "&警察局名稱&" · "&分局名稱&" · "&派出所名稱'
        '&" ｜ 中華民國 "&(YEAR(今日)-1911)&" 年累計（截至 "&TEXT(今日,"mm/dd")'
        '&"）｜ 業務規則：通緝/強採/在監 算到驗 ｜ 距今 > 30 日紅燈"'
    )
    S.set_cell(ws, "A2", sub_formula,
               font_key=S.font(11, color="white", italic=True),
               fill_key=S.fill("accent_dark"), align_key="left")
    ws.row_dimensions[2].height = S.ROW_HEIGHT["sub_banner"]

    # === Row 3 spacer ===
    ws.row_dimensions[3].height = 8

    # === Row 4 維護指引 ===
    ws.merge_cells("A4:S4")
    S.set_cell(ws, "A4",
               "   💡  主表 9 欄 v43 不動 + 最後到驗日（J）。狀態燈/候選由公式自動驅動 "
               "（含距今 > 30 日精確判斷）。毒調率目標可在設定表 C 區調整。",
               font_key=S.font(11, color="warn", italic=True),
               fill_key=S.fill("warn_bg"), align_key="left",
               border_key="all_thin")
    ws.row_dimensions[4].height = 30

    # === Row 5 spacer ===
    ws.row_dimensions[5].height = 8

    # === Row 6 Section titles ===
    ws.merge_cells("A6:I6")
    S.set_cell(ws, "A6", "   📊  頂部狀態縮影 + 毒調率（公式自動）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    # 右：應強採候選（v2.2: P:S）
    ws.merge_cells("P6:S6")
    S.set_cell(ws, "P6", "   ⚠  應強採候選清單",
               font_key="section_title", fill_key=S.fill("danger"),
               align_key="left")
    ws.row_dimensions[6].height = S.ROW_HEIGHT["header"]

    # === Row 7 spacer ===
    ws.row_dimensions[7].height = 6

    # === Row 8-9 — 4 status cards (A:B / C:D / E:F / G:I) ===
    # 公式範圍：Tbl毒調 計算欄
    range_total = f'COUNTA(Tbl毒調[姓名])'                # 列管總數 = 有姓名的列
    range_check = f'SUM(Tbl毒調[是否到驗])'                # 已到驗（含通緝/強採/在監）
    range_cand = f'SUM(Tbl毒調[候選旗標])'                 # 候選 = 未到驗

    status_cards = [
        # (cols_range, label, fg, bg, formula, sub)
        (("A8:B8", "A9:B9"),
            "👥 列管總數", "accent", "accent_light",
            f'={range_total}', "本所列管"),
        (("C8:D8", "C9:D9"),
            "● 已到驗（含 4 種）", "pass", "pass_bg",
            f'={range_check}', "驗+通緝+強採+在監"),
        (("E8:F8", "E9:F9"),
            "● 未到驗", "danger", "danger_bg",
            f'={range_total}-{range_check}-COUNTIF(Tbl毒調[狀態燈],"*結束*")', "需追"),
        (("G8:I8", "G9:I9"),
            "⚠ 應強採候選", "warn", "warn_bg",
            f'={range_cand}', "未到驗待聲請強採"),
    ]
    for (lab_rng, val_rng), label, fg, bg, fml, sub in status_cards:
        ws.merge_cells(lab_rng)
        S.set_cell(ws, lab_rng.split(":")[0], f"  {label}",
                   font_key=S.font(11, bold=True, color=fg),
                   fill_key=S.fill(bg),
                   align_key="left", border_key="top_bottom")
        ws.merge_cells(val_rng)
        S.set_cell(ws, val_rng.split(":")[0], fml,
                   font_key=S.font(26, bold=True, color=fg),
                   fill_key=S.fill(bg),
                   align_key="center", border_key="bottom_thick",
                   number_format="integer")
    ws.row_dimensions[8].height = 24
    ws.row_dimensions[9].height = 44

    # === Row 10 spacer ===
    ws.row_dimensions[10].height = 8

    # === Row 11-12 毒調率 KPI 大字 + 進度條 ===
    # 毒調率 = 已到驗（含 4 種）/ 列管總數
    ws.merge_cells("A11:I11")
    drug_rate_fml = (
        f'="毒調率：" & TEXT(IFERROR({range_check}/{range_total},0),"0.0%")'
        f'&"   ｜  目標 ≥"&TEXT(毒調率目標,"0%")'
        f'&"   ｜  狀態："'
        f'&IF(IFERROR({range_check}/{range_total},0)>=毒調率目標,"● 達標","● 未達標")'
    )
    S.set_cell(ws, "A11", drug_rate_fml,
               font_key=S.font(20, bold=True, color="accent"),
               fill_key=S.fill("accent_light"),
               align_key="center", border_key="all_thin")
    ws.row_dimensions[11].height = 40

    # Row 12 進度條（REPT），含 20 字元滿格
    ws.merge_cells("A12:I12")
    progress_fml = (
        f'=REPT("█",MIN(20,ROUND(IFERROR({range_check}/{range_total},0)*20,0)))'
        f'&REPT("░",20-MIN(20,ROUND(IFERROR({range_check}/{range_total},0)*20,0)))'
    )
    S.set_cell(ws, "A12", progress_fml,
               font_key=S.font(14, bold=True, color="accent",
                               family="Consolas"),
               fill_key="calc",
               align_key="center", border_key="bottom_thick")
    ws.row_dimensions[12].height = 22

    # === Row 13 Section title「📋 列管人口主表」===
    ws.merge_cells("A13:I13")
    S.set_cell(ws, "A13",
               "   📋  列管人口主表（Tbl毒調 9 欄維持 v43 不動 + 3 計算欄隱藏）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[13].height = S.ROW_HEIGHT["header"]

    # === Row 14 主表 header 已由 build_tbl_drug 設置 ===
    # v2.2 候選清單 header (P:S row 14；helper T/U/V 隱藏)
    cand_headers = ["編號", "姓名", "管制狀態", "操作建議"]
    for i, h in enumerate(cand_headers, start=16):  # P=16
        col = get_column_letter(i)
        S.set_cell(ws, f"{col}{DRUG_HEADER_ROW}", h,
                   font_key="header", fill_key=S.fill("danger"),
                   align_key="center", border_key="all_thin")

    # === Row 15-30 候選清單 (16 個位置) ===
    cand_data_start = DRUG_DATA_START
    candidate_helper_start = 100

    # Helper area 100-169 (70 candidates) v2.2 移到 T/U/V (避開 P-S)
    # col T = ROW (1..70)
    # col U = 對應 Tbl毒調[候選旗標]
    # col V = 候選順序 (cumulative)
    for k in range(DRUG_CAPACITY):
        helper_row = candidate_helper_start + k
        seq = k + 1
        S.set_cell(ws, f"T{helper_row}", seq,
                   font_key=S.font(9, color="muted"),
                   number_format="integer")
        S.set_cell(ws, f"U{helper_row}",
                   f'=IFERROR(INDEX(Tbl毒調[候選旗標],{seq}),0)',
                   font_key=S.font(9, color="muted"),
                   number_format="integer")
        S.set_cell(ws, f"V{helper_row}",
                   f'=IF(U{helper_row}=1,SUM($U$100:U{helper_row}),0)',
                   font_key=S.font(9, color="muted"),
                   number_format="integer")

    for r in range(100, 170):
        ws.row_dimensions[r].hidden = True
    for col in ["T", "U", "V"]:
        ws.column_dimensions[col].hidden = True

    S.set_cell(ws, "T99", "[helper] 候選 seq",
               font_key=S.font(9, color="muted"))
    S.set_cell(ws, "U99", "候選旗標",
               font_key=S.font(9, color="muted"))
    S.set_cell(ws, "V99", "候選順序",
               font_key=S.font(9, color="muted"))

    # 候選清單主表 (rows 15-30, 16 個位置)
    helper_seq_range = f"$V$100:$V${candidate_helper_start + DRUG_CAPACITY - 1}"
    for k in range(CANDIDATE_ROWS):
        cand_row = cand_data_start + k
        rank = k + 1
        match_expr = f"MATCH({rank},{helper_seq_range},0)"
        # P 編號
        S.set_cell(ws, f"P{cand_row}",
                   f'=IFERROR(INDEX(Tbl毒調[編號],{match_expr}),"")',
                   font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin")
        # Q 姓名
        S.set_cell(ws, f"Q{cand_row}",
                   f'=IFERROR(INDEX(Tbl毒調[姓名],{match_expr}),"")',
                   font_key="body_bold", fill_key="calc",
                   align_key="left", border_key="all_thin")
        # R 管制狀態
        S.set_cell(ws, f"R{cand_row}",
                   f'=IFERROR(INDEX(Tbl毒調[管制情形],{match_expr}),"")',
                   font_key=S.font(11, color="danger"),
                   fill_key="calc",
                   align_key="left", border_key="all_thin")
        # S 操作建議
        S.set_cell(ws, f"S{cand_row}",
                   f'=IF(P{cand_row}="","","聲請強採書（30 日內）")',
                   font_key=S.font(11, color="warn", italic=True),
                   fill_key="calc",
                   align_key="left", border_key="all_thin")
        ws.row_dimensions[cand_row].height = S.ROW_HEIGHT["default"]

    # 候選清單末端註腳
    note_row = cand_data_start + CANDIDATE_ROWS
    ws.merge_cells(f"P{note_row}:S{note_row}")
    S.set_cell(ws, f"P{note_row}",
               f'="本表最多顯示 {CANDIDATE_ROWS} 個候選；超出請至主表 H 欄篩選『未到驗』或最後到驗日逾 30 日"',
               font_key=S.font(9, color="muted", italic=True),
               fill_key="calc", align_key="center")

    log.info(f"    layout：banner / 4 status cards / 毒調率 KPI / 主表 / 候選 {CANDIDATE_ROWS} 列")


def update_page2_drug_kpis(wb, log):
    """Part C: 把頁 2 KPI 4 毒調率 / 5 未到驗 接上 Tbl毒調"""
    log.info("  [C] 更新頁 2 KPI 4/5 公式（接 Tbl毒調）")
    ws2 = wb["管制總覽"]

    range_total = 'COUNTA(Tbl毒調[姓名])'
    range_check = 'SUM(Tbl毒調[是否到驗])'

    # KPI 4 毒調率：A14（值）+ E13（燈）
    drug_rate_fml = f'=IFERROR({range_check}/{range_total},0)'
    ws2["A14"] = drug_rate_fml
    ws2["A14"].number_format = "0.0%"
    # 燈號公式：閾值 lo=0.5, hi=0.6（用毒調率目標可動態）
    ws2["E13"] = (
        '=IF(IFERROR(SUM(Tbl毒調[是否到驗])/COUNTA(Tbl毒調[姓名]),0)>=毒調率目標,"● 綠",'
        'IF(IFERROR(SUM(Tbl毒調[是否到驗])/COUNTA(Tbl毒調[姓名]),0)>=(毒調率目標-0.1),"⚠ 黃","● 紅"))'
    )

    # KPI 5 未到驗：G14（值）+ K13（燈）
    untested_fml = (
        f'=IFERROR({range_total}-{range_check}'
        f'-COUNTIF(Tbl毒調[狀態燈],"*結束*"),0)'
    )
    ws2["G14"] = untested_fml
    ws2["G14"].number_format = "#,##0"
    # 燈號：reverse=True，值越小越好（lo=5 綠, hi=10 黃）
    ws2["K13"] = (
        '=IF(IFERROR(SUM(Tbl毒調[候選旗標]),0)<=5,"● 綠",'
        'IF(IFERROR(SUM(Tbl毒調[候選旗標]),0)<=10,"⚠ 黃","● 紅"))'
    )

    log.info("    KPI 4 毒調率 ✓ KPI 5 未到驗 ✓")


def build_page8(wb, log):
    """主入口：Part A + B + C"""
    log.info("--- 頁 8 毒品調驗人口管制 ---")
    build_tbl_drug(wb, log)
    build_page8_layout(wb, log)
    update_page2_drug_kpis(wb, log)

    # 公式計數
    ws = wb[DRUG_SHEET]
    fcount = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                fcount += 1
    log.info(f"  頁 8 公式總數：{fcount}")
    log.info(f"  凍結窗格：A14")
    return fcount


def build():
    log = get_logger("phase3_p8")
    log.info("===== Phase 3 — 頁 8 毒品調驗 開始 =====")

    wb = load_workbook(OUTPUT_PATH)
    fcount = build_page8(wb, log)

    bak = backup_existing()
    if bak:
        log.info(f"  舊版備份：{bak}")
    wb.save(OUTPUT_PATH)
    size_kb = OUTPUT_PATH.stat().st_size / 1024
    log.info(f"  儲存：{OUTPUT_PATH.name} ({size_kb:.1f} KB)")
    log.info(f"  公式：{fcount}")
    log.info("===== Phase 3 — 頁 8 完成 =====\n")
    return OUTPUT_PATH


if __name__ == "__main__":
    build()
