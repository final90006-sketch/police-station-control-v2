"""
phase3_page9_traffic.py — Phase 3：頁 9 交通績效管制（達成率看板）

設計：
- 頂部 banner（覆寫 Phase 1 通用 banner）
- 達標縮影（達標 / 未達標 / 總達成率 三卡）
- 違規項目達成率主表：12 項違規（v43 既有，從 E 區同步）
  - 違規項目 ｜ 目標 ｜ 達成 ｜ 達成率 ｜ 橫條視覺 ｜ 狀態燈
- 公式來源：
  - 目標 ← 設定表 E 區（手動同步：頁 9 與 E 區用同一份違規清單）
  - 達成 ← COUNTIFS(Tbl交通[違規項目], 該違規項目)
  - 達成率 = 達成 / 目標（IFERROR 包覆）
  - 狀態 = IF >=80% 綠 / >=60% 黃 / else 紅
  - 橫條視覺 = REPT("█", 達成率*20)
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from phases._base import load_config, backup_existing, OUTPUT_PATH, get_logger
import styles as S


def build_page9_traffic(wb, log):
    log.info("--- 頁 9 交通績效管制（達成率看板）---")
    ws = wb["交通績效管制"]

    # === 覆寫 banner ===
    S.set_cell(ws, "A1", "   🚦  交通績效管制（達成率看板）",
               font_key="banner", fill_key="banner",
               align_key="left", border_key="bottom_thick")
    S.set_cell(ws, "A2",
               "   12 項違規取締目標達成率 ｜ 資料來源：Tbl交通（頁 9.5 取締明細）｜ 目標：設定表 E 區",
               font_key=S.font(11, color="white", italic=True),
               fill_key=S.fill("accent_dark"), align_key="left")

    # === Row 4 維護指引 ===
    ws.row_dimensions[3].height = 8
    ws.merge_cells("A4:F4")
    S.set_cell(ws, "A4",
               "   💡  本頁公式自動：到頁 9.5 取締明細填新案件 → 本頁達成率/狀態燈即時更新。"
               "目標值請去設定表 E 區調整",
               font_key=S.font(11, color="warn", italic=True),
               fill_key=S.fill("warn_bg"), align_key="left",
               border_key="all_thin")
    ws.row_dimensions[4].height = 30

    # === Row 6-8 達標縮影（3 卡片）===
    ws.row_dimensions[5].height = 10
    ws.merge_cells("A6:F6")
    S.set_cell(ws, "A6", "   📊  達標縮影（公式自動）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[6].height = S.ROW_HEIGHT["header"]

    # 載入 phase 1.5 config 取得違規清單
    p15 = load_config("phase1_5_settings")
    violations = p15["section_e"]["violations"]  # [(項目, 目標, 警示閾值)...]

    # 計算公式：總達成率、達標數、未達標數
    target_cells_range = "B14:B" + str(13 + len(violations))
    achieved_range = "C14:C" + str(13 + len(violations))
    rate_range = "D14:D" + str(13 + len(violations))

    # 3 卡片 row 7-8（fg color / bg color 分開指定避免 "accent_bg" 不存在）
    cards = [
        ("● 達標項數",      "pass",     "pass_bg",      f'=COUNTIF({rate_range},">="&0.75)', "integer"),
        ("⚠ 未達標項數",   "warn",     "warn_bg",      f'=COUNTIF({rate_range},"<"&0.75)',  "integer"),
        ("📊 總達成率",    "accent",   "accent_light", f'=IFERROR(SUM({achieved_range})/SUM({target_cells_range}),0)', "percent_one"),
    ]
    for i, (label, fg, bg, formula, fmt) in enumerate(cards):
        col_start = 1 + i * 2
        col_end = col_start + 1
        cl = get_column_letter(col_start)
        cr = get_column_letter(col_end)

        # 上半 label
        ws.merge_cells(f"{cl}7:{cr}7")
        S.set_cell(ws, f"{cl}7", f"  {label}",
                   font_key=S.font(11, color=fg),
                   fill_key=S.fill(bg),
                   align_key="left", border_key="all_thin")
        # 下半 value
        ws.merge_cells(f"{cl}8:{cr}8")
        S.set_cell(ws, f"{cl}8", formula,
                   font_key=S.font(24, bold=True, color=fg),
                   fill_key=S.fill(bg),
                   align_key="center", border_key="all_thin",
                   number_format=fmt)
    ws.row_dimensions[7].height = 26
    ws.row_dimensions[8].height = 48

    # === Row 11+ 違規項目達成率主表（凍結窗格在 row 10 設定）===
    # Phase 1 設凍結 A10，header 放 row 10
    ws.row_dimensions[9].height = 14
    # 但 row 10 已被 Phase 1 freeze_panes 預期，我們把 header 放這
    # 注意：Phase 1 banner 用了 row 1-2，所以 row 10 是 header 沒問題

    # 主表 header at row 13 (留些 spacer)
    ws.row_dimensions[10].height = 14
    ws.merge_cells("A11:F11")
    S.set_cell(ws, "A11", "   📋  違規項目達成率（依設定表 E 區清單同步）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[11].height = S.ROW_HEIGHT["header"]

    ws.row_dimensions[12].height = 10

    # Header row 13
    headers = ["違規項目", "目標", "達成", "達成率", "視覺（█）", "狀態"]
    widths  = [26, 12, 12, 14, 28, 12]
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        col = get_column_letter(i)
        S.set_cell(ws, f"{col}13", h,
                   font_key="header", fill_key="header",
                   align_key="center", border_key="all_thin")
        ws.column_dimensions[col].width = w
    ws.row_dimensions[13].height = S.ROW_HEIGHT["header"]

    # Data rows 14+
    for r_idx, (item_name, target, _warn) in enumerate(violations, start=14):
        # A 違規項目
        S.set_cell(ws, f"A{r_idx}", item_name,
                   font_key="body_bold", fill_key="calc",
                   align_key="left", border_key="all_thin")
        # B 目標（用 VLOOKUP/INDEX 從設定表 E 區）
        # 簡化：直接寫入值（部署後使用者改 E 區，本頁要手動同步）
        # 進階：用 INDEX+MATCH 動態抓取
        S.set_cell(ws, f"B{r_idx}",
                   f'=IFERROR(INDEX(設定表!B:B,MATCH(A{r_idx},違規項目清單,0)+ROW(違規項目清單)-1),{target})',
                   font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin",
                   number_format="integer")
        # 簡化版（保守，怕跨年改變）— 改用直接寫入
        ws[f"B{r_idx}"] = target

        # C 達成（COUNTIFS Tbl交通）
        S.set_cell(ws, f"C{r_idx}",
                   f'=IFERROR(COUNTIFS(Tbl交通[違規項目],A{r_idx}),0)',
                   font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin",
                   number_format="integer")
        # D 達成率
        S.set_cell(ws, f"D{r_idx}",
                   f'=IFERROR(C{r_idx}/B{r_idx},0)',
                   font_key="body_bold", fill_key="calc",
                   align_key="center", border_key="all_thin",
                   number_format="percent_one")
        # E 視覺橫條（REPT）
        S.set_cell(ws, f"E{r_idx}",
                   f'=REPT("█",MIN(20,ROUND(D{r_idx}*20,0)))',
                   font_key=S.font(11, bold=True, color="accent"),
                   fill_key="calc",
                   align_key="left", border_key="all_thin")
        # F 狀態燈
        S.set_cell(ws, f"F{r_idx}",
                   f'=IF(D{r_idx}>=0.8,"● 綠",IF(D{r_idx}>=0.6,"⚠ 黃","● 紅"))',
                   font_key="body_bold", fill_key="calc",
                   align_key="center", border_key="all_thin")
        ws.row_dimensions[r_idx].height = S.ROW_HEIGHT["default"]

    # 合計列
    total_row = 14 + len(violations)
    S.set_cell(ws, f"A{total_row}", "  ─ 全般合計 ─",
               font_key=S.font(12, bold=True, color="white"),
               fill_key="banner", align_key="left",
               border_key="all_thin")
    S.set_cell(ws, f"B{total_row}",
               f'=SUM(B14:B{total_row-1})',
               font_key=S.font(12, bold=True, color="white"),
               fill_key="banner", align_key="center",
               border_key="all_thin", number_format="integer")
    S.set_cell(ws, f"C{total_row}",
               f'=SUM(C14:C{total_row-1})',
               font_key=S.font(12, bold=True, color="white"),
               fill_key="banner", align_key="center",
               border_key="all_thin", number_format="integer")
    S.set_cell(ws, f"D{total_row}",
               f'=IFERROR(C{total_row}/B{total_row},0)',
               font_key=S.font(12, bold=True, color="white"),
               fill_key="banner", align_key="center",
               border_key="all_thin", number_format="percent_one")
    S.set_cell(ws, f"E{total_row}",
               f'=REPT("█",MIN(20,ROUND(D{total_row}*20,0)))',
               font_key=S.font(11, bold=True, color="white"),
               fill_key="banner", align_key="left",
               border_key="all_thin")
    S.set_cell(ws, f"F{total_row}",
               f'=IF(D{total_row}>=0.8,"● 綠",IF(D{total_row}>=0.6,"⚠ 黃","● 紅"))',
               font_key=S.font(12, bold=True, color="white"),
               fill_key="banner", align_key="center",
               border_key="all_thin")
    ws.row_dimensions[total_row].height = S.ROW_HEIGHT["header"]

    log.info(f"  頁 9 完成（{len(violations)} 項違規 + 合計列；公式 SSOT 連動 Tbl交通）")
    return total_row


def build():
    log = get_logger("phase3_p9")
    log.info("===== Phase 3 — 頁 9 交通績效 開始 =====")

    wb = load_workbook(OUTPUT_PATH)
    build_page9_traffic(wb, log)

    bak = backup_existing()
    if bak:
        log.info(f"  舊版備份：{bak}")
    wb.save(OUTPUT_PATH)
    size_kb = OUTPUT_PATH.stat().st_size / 1024
    log.info(f"  儲存：{OUTPUT_PATH.name} ({size_kb:.1f} KB)")
    log.info("===== Phase 3 — 頁 9 完成 =====\n")
    return OUTPUT_PATH


if __name__ == "__main__":
    build()
