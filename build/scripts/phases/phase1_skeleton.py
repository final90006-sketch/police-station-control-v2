"""
phase1_skeleton.py — Phase 1：17 工作表骨架（重排版面）

職責：
- 建立 17 個工作表（中文名稱、順序、凍結窗格）
- 統一欄寬（A:G 7 欄）
- 每張表頂部：banner + 副標
- 設定表額外：TODAY 集中區（獨立 row，視覺清楚）
- 命名範圍「今日」
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

from phases._base import load_config, backup_existing, OUTPUT_PATH, get_logger
import styles as S


def _set_column_widths(ws):
    """套用設定表統一欄寬到所有 sheet（A:G 7 欄）"""
    for col, width in S.SETTINGS_COL_WIDTHS.items():
        ws.column_dimensions[col].width = width


def _add_top_banner(ws, sheet_order: int, sheet_name: str, purpose: str):
    """每張 sheet 共用的頂部 banner（row 1-2）"""
    # Row 1: 主 banner
    ws.merge_cells("A1:G1")
    S.set_cell(ws, "A1",
               f"   📄  頁 {sheet_order:>2} — {sheet_name}",
               font_key="banner",
               fill_key="banner",
               align_key="left",
               border_key="bottom_thick")
    ws.row_dimensions[1].height = S.ROW_HEIGHT["banner"]

    # Row 2: 副標
    ws.merge_cells("A2:G2")
    sub_text = f"   📌  {purpose}    ｜  狀態：骨架（Phase 1）— 內容於後續 Phase 填入"
    S.set_cell(ws, "A2", sub_text,
               font_key=S.font(11, color="white", italic=True),
               fill_key=S.fill("accent_dark"),
               align_key="left")
    ws.row_dimensions[2].height = S.ROW_HEIGHT["sub_banner"]

    # Row 3: spacer
    ws.row_dimensions[3].height = S.ROW_HEIGHT["spacer"]


def _add_today_section(ws):
    """設定表 row 4-5 的 TODAY 集中區（資訊卡風格、視覺清楚）"""
    # Row 4: TODAY 卡片
    # A4: 簡短標籤、右對齊、深藍 accent
    S.set_cell(ws, "A4", "📅 本系統日期",
               font_key=S.font(13, bold=True, color="accent"),
               fill_key="accent_light",
               align_key="right",
               border_key="all_thin")

    # B4: TODAY 公式、大字、置中、淡黃輸入色
    S.set_cell(ws, "B4", "=TODAY()",
               font_key=S.font(18, bold=True, color="accent"),
               fill_key="input",
               align_key="center",
               border_key="all_thin",
               number_format='[$-zh-TW]e/mm/dd')  # 民國年帶 padding 115/05/25

    # C4:G4 merged: 簡短註解（不過長）
    ws.merge_cells("C4:G4")
    S.set_cell(ws, "C4",
               "   ← 此格命名「今日」，全 v2 公式引用此格（避免 TODAY() 散落，性能優化）",
               font_key=S.font(10, color="muted", italic=True),
               fill_key="code_bg",
               align_key="left",
               border_key="all_thin")

    ws.row_dimensions[4].height = 44   # 更高的 TODAY 卡片列高

    # Row 5: spacer
    ws.row_dimensions[5].height = S.ROW_HEIGHT["spacer"]


def build():
    log = get_logger("phase1")
    log.info("===== Phase 1（骨架，重排版面）開始 =====")

    cfg = load_config("phase1_skeleton")
    wb = Workbook()
    wb.remove(wb.active)

    sheets_sorted = sorted(cfg["sheets"], key=lambda s: s["order"])

    # === 建立 17 工作表 ===
    for s in sheets_sorted:
        ws = wb.create_sheet(title=s["name"])

        # 凍結窗格
        if s.get("freeze"):
            ws.freeze_panes = s["freeze"]

        # 欄寬統一
        _set_column_widths(ws)

        # 頂部 banner
        _add_top_banner(ws, s["order"], s["name"], s.get("purpose", ""))

        log.info(f"  [{s['order']:>2}] {s['name']:<24} freeze={s['freeze']}")

    # === 設定表 — TODAY 集中區（row 4） ===
    settings_ws = wb["設定表"]
    _add_today_section(settings_ws)
    log.info(f"  TODAY 集中：設定表!$C$4（民國年格式）")

    # === 命名範圍「今日」← 設定表!$B$4 ===
    wb.defined_names["今日"] = DefinedName(name="今日", attr_text="設定表!$B$4")
    log.info(f"  命名範圍：今日 = 設定表!$B$4")

    # === 設定表 row 6 = 後續 Section A 將從這開始 ===
    # （此處不做事，留給 Phase 1.5）

    # === 開檔預設停在「設定表」（多重設定確保生效）===
    settings_idx = wb.sheetnames.index("設定表")
    wb.active = settings_idx
    for i, sht in enumerate(wb.worksheets):
        sht.sheet_view.tabSelected = (i == settings_idx)
        sht.sheet_view.topLeftCell = "A1"
    log.info(f"  開檔預設 sheet：設定表（index {settings_idx}）")

    # 設定表頁籤顏色
    settings_ws.sheet_properties.tabColor = S.COLOR["accent"][2:]

    # === 儲存 ===
    bak = backup_existing()
    if bak:
        log.info(f"  舊版備份：{bak}")
    wb.save(OUTPUT_PATH)
    size_kb = OUTPUT_PATH.stat().st_size / 1024
    log.info(f"  儲存：{OUTPUT_PATH.name} ({size_kb:.1f} KB)")
    log.info("===== Phase 1 完成 =====\n")
    return OUTPUT_PATH


if __name__ == "__main__":
    build()
