"""
phase3_pages_1_7.py — Phase 3：頁 1 首頁 + 頁 7 發展中案件管制
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

from phases._base import backup_existing, OUTPUT_PATH, get_logger
import styles as S


# ============================================================
# 頁 1：首頁_操作說明
# ============================================================
def build_page1_home(wb, log):
    log.info("--- 頁 1 首頁_操作說明 ---")
    ws = wb["首頁_操作說明"]

    # 覆寫 Phase 1 banner
    S.set_cell(ws, "A1", "   🏛  派出所綜合管制系統 v2",
               font_key="banner", fill_key="banner",
               align_key="left", border_key="bottom_thick")
    S.set_cell(ws, "A2",
               "   KKEVIN-LIN-2026-V2.0 ｜ 全台通用 ｜ 無巨集 ｜ 一檔搞定刑案/毒調/交通/績效/歷史",
               font_key=S.font(11, color="white", italic=True),
               fill_key=S.fill("accent_dark"), align_key="left")

    # Row 4-6 Insight headline
    ws.merge_cells("A4:G4")
    S.set_cell(ws, "A4",
               "   ★  本月一句話",
               font_key=S.font(11, bold=True, color="warn"),
               fill_key="warn_bg", align_key="left")
    ws.row_dimensions[4].height = 24

    ws.merge_cells("A5:G5")
    S.set_cell(ws, "A5",
               '   ="系統健檢 ✓ 通過 ｜ 案件 " & '
               'IFERROR(COUNTA(Tbl案件[編號]), 0) & '
               '"/" & 案件容量上限 & " ｜ 健康狀態：穩定運作"',
               font_key=S.font(13, bold=True, color="text"),
               fill_key="warn_bg", align_key="left",
               border_key="bottom_thick")
    ws.row_dimensions[5].height = 34

    ws.row_dimensions[6].height = 12

    # Row 7-13 角色分流區（兩欄）
    ws.merge_cells("A7:C7")
    S.set_cell(ws, "A7", "   👤  我是…",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.merge_cells("D7:G7")
    S.set_cell(ws, "D7", "   📍  我要去…（點擊一鍵跳轉）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[7].height = S.ROW_HEIGHT["header"]

    roles_targets = [
        ("第一次用",          "→ 設定表（先填派出所名稱與員警名冊）"),
        ("員警／承辦人",      "→ 案件資料庫（登錄案件）"),
        ("所長",              "→ 管制總覽（看全所狀態）"),
        ("督導官",            "→ 長官5分鐘報告（A4 一頁列印交付）"),
        ("上級機關承辦",      "→ 對上級機關上呈（公文體例三層格式）"),
    ]
    r = 8
    for role, target in roles_targets:
        ws.merge_cells(f"A{r}:C{r}")
        S.set_cell(ws, f"A{r}", f"   ▸  {role}",
                   font_key="body_bold", fill_key="input",
                   align_key="left", border_key="all_thin")
        ws.merge_cells(f"D{r}:G{r}")
        S.set_cell(ws, f"D{r}", f"   {target}",
                   font_key="body", fill_key="calc",
                   align_key="left", border_key="all_thin")
        ws.row_dimensions[r].height = S.ROW_HEIGHT["default"]
        r += 1

    r += 1
    ws.row_dimensions[r - 1].height = 12

    # 5 分鐘上手三步驟
    ws.merge_cells(f"A{r}:G{r}")
    S.set_cell(ws, f"A{r}", "   📦  5 分鐘上手三步驟",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[r].height = S.ROW_HEIGHT["header"]
    r += 1

    steps = [
        ("①", "填設定表", "派出所名稱、警察局、分局、員警名冊 — 一次設定終身受用"),
        ("②", "案件資料庫試打", "到「案件資料庫」第一列填試打資料，下拉選單會逐欄帶您"),
        ("③", "回管制總覽", "看九宮格 KPI 是否點亮，狀態燈正常即代表系統運作健康"),
    ]
    for num, title, desc in steps:
        # 編號 (A) — 深藍底 / 白大字
        S.set_cell(ws, f"A{r}", num,
                   font_key=S.font(24, bold=True, color="white"),
                   fill_key="banner", align_key="center",
                   border_key="all_thin")
        # 標題 (B) — accent_light 底 / 粗體
        S.set_cell(ws, f"B{r}", title,
                   font_key=S.font(14, bold=True, color="accent"),
                   fill_key="accent_light", align_key="center",
                   border_key="all_thin")
        # 說明 (C-G merged) — 灰底
        ws.merge_cells(f"C{r}:G{r}")
        S.set_cell(ws, f"C{r}", f"  {desc}",
                   font_key=S.font(12, color="text"),
                   fill_key="calc",
                   align_key="left", border_key="all_thin")
        ws.row_dimensions[r].height = 48
        r += 1

    r += 1
    ws.row_dimensions[r - 1].height = 12

    # 頁尾 — 修 bug：原本前 3 個空白導致 Excel 視為文字
    ws.merge_cells(f"A{r}:G{r}")
    S.set_cell(ws, f"A{r}",
               '="    © 林錦瑞    ｜    KKEVIN-LIN-2026-V2.1    ｜    今日:" & TEXT(今日, "[$-zh-TW]yyyy/mm/dd")',
               font_key=S.font(11, color="muted", italic=True),
               fill_key="accent_light", align_key="center",
               border_key="all_thin")
    ws.row_dimensions[r].height = 32

    log.info(f"  頁 1 完成（共 {r} 列）")


# ============================================================
# 頁 7：發展中案件管制（狀態縮影 + 70/30 主表+候選）
# ============================================================
def build_page7_developing(wb, log):
    log.info("--- 頁 7 發展中案件管制 ---")
    ws = wb["發展中案件管制"]

    # 先清掉既有 merges（避免與新 layout 重疊）
    from phases._base import unmerge_all_in_sheet
    cleared = unmerge_all_in_sheet(ws)
    log.info(f"  清空既有 merges：{cleared} 個")

    # 覆寫 banner（必須重新 merge — unmerge_all 清掉 Phase 1 的 A1:G1）
    ws.merge_cells("A1:O1")
    S.set_cell(ws, "A1", "   📌  發展中案件管制（v2.1 加 承辦人 / 線索來源 / 重大性 三欄）",
               font_key="banner", fill_key="banner",
               align_key="left", border_key="bottom_thick")
    ws.row_dimensions[1].height = S.ROW_HEIGHT["banner"]

    ws.merge_cells("A2:O2")
    S.set_cell(ws, "A2",
               "   主表 9 欄（v43 6 欄 + v2.1 3 欄）+ 頂部三段狀態縮影 + 右側待轉案件候選（70/30 兩翼防呆）",
               font_key=S.font(11, color="white", italic=True),
               fill_key=S.fill("accent_dark"), align_key="left")
    ws.row_dimensions[2].height = S.ROW_HEIGHT["sub_banner"]

    # Row 4-7：三段狀態縮影（卡片式）
    ws.row_dimensions[3].height = 8
    ws.merge_cells("A4:E4")
    S.set_cell(ws, "A4", "   ⚡  狀態縮影（v2.1 閾值改 7/30 日 — 更敏感）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[4].height = S.ROW_HEIGHT["header"]

    ws.row_dimensions[5].height = 8

    # Row 6-8：3 個狀態卡（v2.1 改 7/30 日閾值）
    cards = [
        ('="● 新進（≤"&新進閾值&" 日）"',       6,   "pass_bg",    "pass"),
        ('="⚠ 偵辦中（"&(新進閾值+1)&"-"&老案閾值&" 日）"',  7,   "warn_bg",    "warn"),
        ('="✕ 老案（>"&老案閾值&" 日）"',       8,   "danger_bg",  "danger"),
    ]
    for label, row_idx, bg, fg in cards:
        ws.merge_cells(f"A{row_idx}:B{row_idx}")
        S.set_cell(ws, f"A{row_idx}", label,
                   font_key=S.font(12, bold=True, color=fg),
                   fill_key=S.fill(bg), align_key="left",
                   border_key="all_thin")
        # 值（範例 — 後續 Phase 接公式時改成 COUNTIF）
        S.set_cell(ws, f"C{row_idx}",
                   f"={3 if row_idx==6 else (4 if row_idx==7 else 1)}",
                   font_key=S.font(20, bold=True, color=fg),
                   fill_key=S.fill(bg), align_key="center",
                   border_key="all_thin")
        ws.row_dimensions[row_idx].height = 32

    # Row 10：閾值說明條
    ws.row_dimensions[9].height = 8
    ws.merge_cells("A10:O10")
    S.set_cell(ws, "A10",
               '="🔧 閾值：新進 ≤ " & 新進閾值 & " 日 ｜ 老案 > " & 老案閾值 & " 日 ｜ 可在設定表 C 區調整"',
               font_key=S.font(11, color="muted", italic=True),
               fill_key="code_bg", align_key="left",
               border_key="all_thin")
    ws.row_dimensions[10].height = 26

    # ===== Row 12+：70/30 主表（A-K）+ 候選（M-O）=====
    ws.row_dimensions[11].height = 12

    # 主表 banner（含 9 input + 2 calc = 11 cols A-K）
    ws.merge_cells("A12:K12")
    S.set_cell(ws, "A12",
               "   📋  發展中案件主表（v43 6 欄 + v2.1 3 欄輸入 + 2 計算欄）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[12].height = S.ROW_HEIGHT["header"]

    # 候選 banner (M-O)
    ws.merge_cells("M12:O12")
    S.set_cell(ws, "M12", "   📌  待轉案件候選",
               font_key="section_title", fill_key="banner",
               align_key="left")

    # 主表 header row 13 (9 input cols A-I)
    headers_main = [
        ("編號",     10),
        ("案由",     24),
        ("進度",     12),
        ("執行時間", 14),
        ("查緝結果", 22),
        ("備註",     22),
        ("承辦人",   12),  # v2.1 新增
        ("線索來源", 14),  # v2.1 新增
        ("重大性",   10),  # v2.1 新增
    ]
    for i, (h, w) in enumerate(headers_main, start=1):
        col = get_column_letter(i)
        S.set_cell(ws, f"{col}13", h,
                   font_key="header", fill_key="header",
                   align_key="center", border_key="all_thin")
        ws.column_dimensions[col].width = w
    # 計算欄 J/K (距今 / 狀態)
    for i, h in enumerate(["距今", "狀態"], start=10):
        col = get_column_letter(i)
        S.set_cell(ws, f"{col}13", h,
                   font_key="header", fill_key="header",
                   align_key="center", border_key="all_thin")
        ws.column_dimensions[col].width = 10

    # L spacer
    ws.column_dimensions["L"].width = 2

    # 候選 header row 13 (M-O)
    headers_cand = [("#", 6), ("案由 / 狀態", 32), ("操作", 16)]
    for i, (h, w) in enumerate(headers_cand, start=13):
        col = get_column_letter(i)
        S.set_cell(ws, f"{col}13", h,
                   font_key="header", fill_key="header",
                   align_key="center", border_key="all_thin")
        ws.column_dimensions[col].width = w
    ws.row_dimensions[13].height = S.ROW_HEIGHT["header"]

    # 示範資料 5 筆（row 14-18，9 input + 2 calc）
    samples = [
        # (編號, 案由, 進度, 執行時間, 查緝結果, 備註, 承辦人, 線索來源, 重大性, 距今, 狀態)
        ("001", "○○公園販毒線索",  "偵辦中", "2026/05/22", "監控中⋯",     "刑大協助",  "王小明", "巡邏發現", "高",  4,  "● 綠"),
        ("002", "XX 路詐騙集團",    "偵辦中", "2026/05/12", "蒐證中⋯",     "民眾照片",  "李大華", "民眾報案", "中", 14, "⚠ 黃"),
        ("003", "夜店毒品流通",     "已成熟", "2026/05/02", "嫌犯身分確認", "擬聲搜",    "陳美玲", "同仁通報", "高", 24, "⚠ 黃"),
        ("004", "○○街頭聚賭",      "偵辦中", "2026/04/20", "查無實證⋯",   "建議簽結",  "王小明", "巡邏發現", "低", 36, "✕ 紅"),
        ("005", "外勞集資詐騙",     "已成熟", "2026/05/06", "通譯協助",     "擬移送地檢","李大華", "上級交辦", "高", 20, "⚠ 黃"),
    ]
    for r_idx, sample in enumerate(samples, start=14):
        for c_idx, val in enumerate(sample, start=1):
            col = get_column_letter(c_idx)
            fill_key = "calc" if c_idx > 9 else "input"
            font_key = "body"
            if c_idx == 11:  # 狀態欄 K
                if "綠" in str(val):
                    font_key = S.font(12, bold=True, color="pass")
                elif "黃" in str(val):
                    font_key = S.font(12, bold=True, color="warn")
                elif "紅" in str(val):
                    font_key = S.font(12, bold=True, color="danger")
            S.set_cell(ws, f"{col}{r_idx}", val,
                       font_key=font_key, fill_key=fill_key,
                       align_key="center" if c_idx not in [2, 5, 6] else "left",
                       border_key="all_thin")
        ws.row_dimensions[r_idx].height = S.ROW_HEIGHT["default"]

    # 預留 10 列空白給未來 (rows 19-28) — 含 DV 範圍
    for r in range(19, 29):
        for c_idx in range(1, 12):   # A-K
            col = get_column_letter(c_idx)
            fill = "calc" if c_idx > 9 else "input"
            S.set_cell(ws, f"{col}{r}", None,
                       fill_key=fill, border_key="all_thin")
        ws.row_dimensions[r].height = S.ROW_HEIGHT["default"]

    # === DataValidation ===
    dv_specs = [
        ("C", "發展中進度清單"),
        ("G", "員警姓名清單"),
        ("H", "線索來源清單"),
        ("I", "重大性清單"),
    ]
    for col, named in dv_specs:
        dv = DataValidation(type="list", formula1=f"={named}", allow_blank=True)
        dv.add(f"{col}14:{col}28")
        ws.add_data_validation(dv)

    # 候選清單（row 14-16，3 筆）— M/N/O
    candidates = [
        ("003", "夜店毒品流通\n已成熟 · 24 日", "轉案件",  "warn_bg"),
        ("005", "外勞集資詐騙\n已成熟 · 20 日", "轉案件",  "warn_bg"),
        ("004", "○○街頭聚賭\n老案逾 30 · 36 日", "評估簽結", "danger_bg"),
    ]
    for r_idx, (no, desc, action, bg) in enumerate(candidates, start=14):
        S.set_cell(ws, f"M{r_idx}", no,
                   font_key="body_bold", fill_key=S.fill(bg),
                   align_key="center", border_key="all_thin")
        S.set_cell(ws, f"N{r_idx}", desc,
                   font_key="body", fill_key=S.fill(bg),
                   align_key="left", border_key="all_thin")
        S.set_cell(ws, f"O{r_idx}", action,
                   font_key=S.font(11, bold=True, color="white"),
                   fill_key=S.fill("danger" if "簽結" in action else "accent"),
                   align_key="center", border_key="all_thin")
        ws.row_dimensions[r_idx].height = 40

    # 候選底部說明
    ws.merge_cells("M17:O17")
    S.set_cell(ws, "M17",
               "   ✏ 偵查佐評估後：「轉案件」者複製欄位 → 貼到頁 5",
               font_key=S.font(10, color="muted", italic=True),
               fill_key="code_bg", align_key="left",
               border_key="all_thin")
    ws.row_dimensions[17].height = 30

    log.info(f"  頁 7 完成（主表 9 欄 + 計算 2 欄 + 候選 3 欄；5 示範 + 10 預留；4 DV）")


# ============================================================
# Main
# ============================================================
def build():
    log = get_logger("phase3")
    log.info("===== Phase 3（頁 1 + 頁 7）開始 =====")

    wb = load_workbook(OUTPUT_PATH)
    build_page1_home(wb, log)
    build_page7_developing(wb, log)

    bak = backup_existing()
    if bak:
        log.info(f"  舊版備份：{bak}")
    wb.save(OUTPUT_PATH)
    size_kb = OUTPUT_PATH.stat().st_size / 1024
    log.info(f"  儲存：{OUTPUT_PATH.name} ({size_kb:.1f} KB)")
    log.info("===== Phase 3 完成 =====\n")
    return OUTPUT_PATH


if __name__ == "__main__":
    build()
