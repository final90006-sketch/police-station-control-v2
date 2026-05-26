"""
phase1_5_settings.py — Phase 1.5：設定表 8 大分區（精緻版）

視覺亮點：
- 頂部大 banner（深藍底白字、版本資訊）
- 每區獨立 banner + icon + title
- 區塊間 spacer 留白
- 全表格邊框、字體、行高、欄寬精調
- TODAY 集中儲存格高亮
- 22 個命名範圍動態計算位置（不再寫死）
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter

from phases._base import load_config, backup_existing, OUTPUT_PATH, get_logger
import styles as S


# ============================================================
# 共用 helper
# ============================================================
def update_top_banner(ws, title: str, subtitle: str, version: str):
    """覆寫 Phase 1 已建立的頂部 banner（row 1-2），改成設定表專屬版本"""
    # Row 1: 主 banner（覆寫）
    S.set_cell(ws, "A1",
               f"   🏢  {title}",
               font_key="banner",
               fill_key="banner",
               align_key="left",
               border_key="bottom_thick")
    ws.row_dimensions[1].height = S.ROW_HEIGHT["banner"]

    # Row 2: 副標（覆寫）
    S.set_cell(ws, "A2",
               f"   {subtitle}    ｜   版本：{version}",
               font_key=S.font(11, color="white", italic=True),
               fill_key=S.fill("accent_dark"),
               align_key="left")
    ws.row_dimensions[2].height = S.ROW_HEIGHT["sub_banner"]


def add_section_banner(ws, row: int, icon: str, code: str, name: str,
                       desc: str = "", span_cols: int = 5):
    """區塊 banner — 含 icon、區代碼、名稱、簡短描述"""
    end_col_letter = get_column_letter(span_cols)
    ws.merge_cells(f"A{row}:{end_col_letter}{row}")
    text = f"  {icon}  {code} 區 ─ {name}"
    if desc:
        text += f"    ｜ {desc}"
    S.set_cell(ws, f"A{row}", text,
               font_key=S.font(13, bold=True, color="white"),
               fill_key="banner",
               align_key="left",
               border_key="bottom_thick")
    ws.row_dimensions[row].height = S.ROW_HEIGHT["section_title"]
    return row + 1


def add_table_header(ws, row: int, headers: list, col_widths: list = None):
    """寫入表頭列（深藍底白字 + 邊框）
    注：col_widths 參數已棄用 — 全頁用 Phase 1 設的 SETTINGS_COL_WIDTHS 統一寬度
    避免每區段不一致的寬度覆蓋。"""
    for i, h in enumerate(headers, start=1):
        col = get_column_letter(i)
        S.set_cell(ws, f"{col}{row}", h,
                   font_key="header",
                   fill_key="header",
                   align_key="center",
                   border_key="all_thin")
    ws.row_dimensions[row].height = S.ROW_HEIGHT["header"]
    # 不再覆蓋全頁欄寬（保持 Phase 1 設定一致）
    return row + 1


def add_data_row(ws, row: int, values: list, *, input_cols: set = None,
                 num_formats: dict = None):
    """寫入一列資料（輸入欄淡黃、邊框）"""
    input_cols = input_cols or set(range(1, len(values) + 1))  # 預設全部淡黃
    num_formats = num_formats or {}
    for i, val in enumerate(values, start=1):
        col = get_column_letter(i)
        fill_key = "input" if i in input_cols else "calc"
        S.set_cell(ws, f"{col}{row}", val,
                   font_key="body",
                   fill_key=fill_key,
                   align_key="center",
                   border_key="all_thin",
                   number_format=num_formats.get(i))
    ws.row_dimensions[row].height = S.ROW_HEIGHT["default"]
    return row + 1


def add_spacer(ws, row: int, height: int = 12):
    """區塊間 spacer 列"""
    ws.row_dimensions[row].height = height
    return row + 1


def add_section_instruction(ws, row: int, text: str, span_cols: int = 7):
    """在區塊 banner 下方加一條「怎麼新增/維護」小提示（淡黃底、左對齊、italic）
    span_cols: 跨多少欄合併（依該區實際表格寬度，避免黃色突出）"""
    end_col = get_column_letter(span_cols)
    ws.merge_cells(f"A{row}:{end_col}{row}")
    S.set_cell(ws, f"A{row}", f"   💡 {text}",
               font_key=S.font(11, color="warn", italic=True),
               fill_key=S.fill("warn_bg"),
               align_key="left",
               border_key="all_thin")
    ws.row_dimensions[row].height = 30
    return row + 1


def add_maintenance_guide(ws, row: int):
    """設定表頂部「新增與維護指引」資訊卡 — 一表看完所有怎麼新增/停用"""
    # Title banner
    ws.merge_cells(f"A{row}:G{row}")
    S.set_cell(ws, f"A{row}",
               "   📌  新增與維護指引（部署後可隨時更新；不需刪整列、改「是否啟用」即可保留歷史）",
               font_key=S.font(13, bold=True, color="white"),
               fill_key="banner",
               align_key="left",
               border_key="bottom_thick")
    ws.row_dimensions[row].height = S.ROW_HEIGHT["header"]
    row += 1

    # Table header
    headers = ["區", "項目", "新增方式", "上限", "停用/刪除", "連動影響"]
    for i, h in enumerate(headers, start=1):
        col = get_column_letter(i)
        S.set_cell(ws, f"{col}{row}", h,
                   font_key="header",
                   fill_key="header",
                   align_key="center",
                   border_key="all_thin")
    # G 欄空白填頭
    S.set_cell(ws, f"G{row}", "",
               fill_key="header", border_key="all_thin")
    ws.row_dimensions[row].height = S.ROW_HEIGHT["header"]
    row += 1

    # 維護指引 5 列
    guide_rows = [
        ("B", "員警名冊",     "下方空白列 key 入",       "60 名",    "改「是否在職=否」",   "下拉清單自動擴展"),
        ("D", "案類分類",     "下方空白列 key 入",       "67 列",    "改「納入全般=否」",   "頁 5/10 案類即時更新"),
        ("E", "違規項目",     "下方空白列 key 入",       "16 項",    "改目標=0 或留白",       "頁 9 達成率自動更新"),
        ("F", "婦幼／協尋",  "直接改件數＋更新日期",  "—",         "件數=0 自動顯示「尚無」", "頁 2.5「其他管制」連動"),
        ("C", "全所閾值",     "改數值即可",                "—",         "改回預設值",                "各頁警示燈閾值跟著動"),
    ]
    for code, item, action, limit, disable, impact in guide_rows:
        # 區 (中央對齊、藍色 bold)
        S.set_cell(ws, f"A{row}", code,
                   font_key=S.font(13, bold=True, color="accent"),
                   fill_key="accent_light",
                   align_key="center", border_key="all_thin")
        S.set_cell(ws, f"B{row}", item,
                   font_key="body_bold", fill_key="calc",
                   align_key="left", border_key="all_thin")
        # 跨 C:D 合併「新增方式」
        ws.merge_cells(f"C{row}:D{row}")
        S.set_cell(ws, f"C{row}", action,
                   font_key="body", fill_key="input",
                   align_key="left", border_key="all_thin")
        S.set_cell(ws, f"E{row}", limit,
                   font_key="body", fill_key="calc",
                   align_key="center", border_key="all_thin")
        S.set_cell(ws, f"F{row}", disable,
                   font_key="body", fill_key="calc",
                   align_key="left", border_key="all_thin")
        S.set_cell(ws, f"G{row}", impact,
                   font_key="footnote", fill_key="calc",
                   align_key="left", border_key="all_thin")
        ws.row_dimensions[row].height = S.ROW_HEIGHT["default"]
        row += 1

    return row


# ============================================================
# 各區建構函式（每個都 return 下一個可用列）
# ============================================================
def build_section_a(ws, wb, row: int, sec: dict, log):
    """A 區 三層組織"""
    row = add_section_banner(ws, row, "🏢", "A", sec["name"], span_cols=3)
    row = add_table_header(ws, row, ["項目", "值", "說明"],
                            col_widths=[20, 32, 36])
    for field in sec["fields"]:
        # 「項目」(灰計算) | 「值」(淡黃輸入) | 「說明」(輔助色)
        S.set_cell(ws, f"A{row}", field["label"],
                   font_key="body_bold", fill_key="calc",
                   align_key="left", border_key="all_thin")
        S.set_cell(ws, f"B{row}", field.get("default", ""),
                   font_key="body", fill_key="input",
                   align_key="left", border_key="all_thin")
        S.set_cell(ws, f"C{row}", field.get("note", ""),
                   font_key="footnote", fill_key="calc",
                   align_key="left", border_key="all_thin")
        ws.row_dimensions[row].height = S.ROW_HEIGHT["default"]

        # 命名範圍
        if field.get("named_range"):
            cell_ref = f"設定表!$B${row}"
            wb.defined_names[field["named_range"]] = DefinedName(
                name=field["named_range"], attr_text=cell_ref)
            log.info(f"  命名範圍：{field['named_range']} = {cell_ref}")
        row += 1

    log.info(f"  A 區（三層組織 + 勤區）完成，{len(sec['fields'])} 項")
    return row


def build_section_b(ws, wb, row: int, sec: dict, log):
    """B 區 員警名冊（7 欄）+ 員警姓名清單命名範圍"""
    row = add_section_banner(ws, row, "👮", "B",
                              f"{sec['name']}（最多 60 名）", span_cols=7)
    row = add_section_instruction(ws, row,
        "新增員警：① 點下方任一空白列 → ② 直接 key 入 7 欄資料（編號/姓名/職稱⋯）→ ③ 員警下拉清單自動擴展。"
        "離職：改「是否在職=否」（保留歷史紀錄）",
        span_cols=7)
    row = add_table_header(ws, row, sec["headers"],
                            col_widths=[8, 14, 12, 12, 14, 8, 24])

    data_start = row  # 員警資料起始列（姓名在 B 欄）
    for officer in sec["sample_officers"]:
        row = add_data_row(ws, row, officer)
    for _ in range(sec["reserved_blank_rows"]):
        for col_idx in range(1, len(sec["headers"]) + 1):
            col = get_column_letter(col_idx)
            S.set_cell(ws, f"{col}{row}", "",
                       fill_key="input",
                       border_key="all_thin")
        ws.row_dimensions[row].height = S.ROW_HEIGHT["default"]
        row += 1
    data_end = row - 1

    # 命名範圍：員警姓名清單（B 欄）= 設定表!$B$<start>:$B$<end>
    name_range = f"設定表!$B${data_start}:$B${data_end}"
    wb.defined_names["員警姓名清單"] = DefinedName(
        name="員警姓名清單", attr_text=name_range)
    log.info(f"  命名範圍：員警姓名清單 = {name_range}")

    log.info(f"  B 區（員警名冊）完成，{len(sec['sample_officers'])} 示範 + "
             f"{sec['reserved_blank_rows']} 預留")
    return row


def build_section_c(ws, wb, row: int, sec: dict, log):
    """C 區 全所閾值（含 TODAY 集中註釋）"""
    row = add_section_banner(ws, row, "⚙", "C",
                              sec["name"],
                              desc="性能關鍵：TODAY() 集中於此區",
                              span_cols=3)
    row = add_table_header(ws, row, ["項目", "值", "說明"],
                            col_widths=[24, 16, 40])

    for thr in sec["thresholds"]:
        S.set_cell(ws, f"A{row}", thr["label"],
                   font_key="body_bold", fill_key="calc",
                   align_key="left", border_key="all_thin")

        val = thr["default"]
        cell_b = S.set_cell(ws, f"B{row}", val,
                            font_key="body_bold", fill_key="input",
                            align_key="center", border_key="all_thin")
        if thr.get("number_format"):
            cell_b.number_format = thr["number_format"]

        S.set_cell(ws, f"C{row}", thr.get("note", ""),
                   font_key="footnote", fill_key="calc",
                   align_key="left", border_key="all_thin")
        ws.row_dimensions[row].height = S.ROW_HEIGHT["default"]

        if thr.get("named_range"):
            cell_ref = f"設定表!$B${row}"
            wb.defined_names[thr["named_range"]] = DefinedName(
                name=thr["named_range"], attr_text=cell_ref)
            log.info(f"  命名範圍：{thr['named_range']} = {cell_ref}")
        row += 1

    log.info(f"  C 區（全所閾值）完成，{len(sec['thresholds'])} 項")
    return row


def build_section_d(ws, wb, row: int, sec: dict, log):
    """D 區 案類分類（5 欄）"""
    row = add_section_banner(ws, row, "📂", "D",
                              f"{sec['name']}（最多 67 列）",
                              desc="納入全般統計 + 納入破獲率",
                              span_cols=5)
    row = add_section_instruction(ws, row,
        "新增案類：① 點下方任一空白列 → ② 填「代碼/名稱/分類」並設「納入全般」「納入破獲率」是/否。"
        "停用：改「納入全般=否」（保留歷史紀錄）",
        span_cols=5)
    row = add_table_header(ws, row, sec["headers"],
                            col_widths=[10, 18, 12, 14, 14])
    data_start = row  # 案類資料起始列（案類名稱在 B 欄）
    for ct in sec["case_types"]:
        row = add_data_row(ws, row, ct)
    # v2.1 改：named range 僅含實際 case_types（不含 blank 預留），方便公式精準對齊
    data_end = row - 1
    # 保留 5 列空白預留（不影響 named range）
    for _ in range(5):
        for col_idx in range(1, len(sec["headers"]) + 1):
            cell = ws.cell(row=row, column=col_idx)
            cell.fill = S.FILL["input"]
            cell.border = S.BORDER["all_thin"]
        ws.row_dimensions[row].height = S.ROW_HEIGHT["default"]
        row += 1

    # 命名範圍（v2.1 共 4 個）
    # 案類代碼清單 / 案類名稱清單 / 案類分類清單（C 欄，v2.1 新增）/ 納入全般清單（D 欄，v2.1 新增）
    wb.defined_names["案類代碼清單"] = DefinedName(
        name="案類代碼清單", attr_text=f"設定表!$A${data_start}:$A${data_end}")
    wb.defined_names["案類名稱清單"] = DefinedName(
        name="案類名稱清單", attr_text=f"設定表!$B${data_start}:$B${data_end}")
    wb.defined_names["案類分類清單"] = DefinedName(
        name="案類分類清單", attr_text=f"設定表!$C${data_start}:$C${data_end}")
    wb.defined_names["納入全般清單"] = DefinedName(
        name="納入全般清單", attr_text=f"設定表!$D${data_start}:$D${data_end}")
    log.info(f"  命名範圍：案類代碼/名稱/分類/納入全般 = 設定表!$X${data_start}:$X${data_end}")

    log.info(f"  D 區（案類分類）完成，{len(sec['case_types'])} 類 + 5 空白預留")
    return row


def build_section_e(ws, wb, row: int, sec: dict, log):
    """E 區 違規項目（3 欄）"""
    row = add_section_banner(ws, row, "🚦", "E",
                              f"{sec['name']}（最多 16 項）",
                              desc="目標值與警示閾值",
                              span_cols=3)
    row = add_section_instruction(ws, row,
        "新增違規項目：① 點下方任一空白列 → ② 填「項目/目標/警示閾值」。"
        "頁 9 達成率自動算 = 取締數/目標。停用：目標=0 即不列入考核",
        span_cols=3)
    row = add_table_header(ws, row, sec["headers"],
                            col_widths=[24, 10, 18])
    data_start = row
    for v in sec["violations"]:
        row = add_data_row(ws, row, v,
                            num_formats={2: S.NUM_FMT["integer"]})
    for _ in range(4):
        for col_idx in range(1, len(sec["headers"]) + 1):
            cell = ws.cell(row=row, column=col_idx)
            cell.fill = S.FILL["input"]
            cell.border = S.BORDER["all_thin"]
        ws.row_dimensions[row].height = S.ROW_HEIGHT["default"]
        row += 1
    data_end = row - 1

    # 命名範圍：違規項目清單（A 欄）
    wb.defined_names["違規項目清單"] = DefinedName(
        name="違規項目清單", attr_text=f"設定表!$A${data_start}:$A${data_end}")
    log.info(f"  命名範圍：違規項目清單 = 設定表!$A${data_start}:$A${data_end}")

    log.info(f"  E 區（違規項目）完成，{len(sec['violations'])} 項 + 4 空白預留")
    return row


def build_section_f(ws, wb, row: int, sec: dict, log):
    """F 區 婦幼/協尋（4 格半自動）"""
    row = add_section_banner(ws, row, "👶", "F",
                              sec["name"],
                              desc="連動頁 2.5「其他管制項目」",
                              span_cols=3)
    row = add_table_header(ws, row, ["項目", "件數", "最後更新日"],
                            col_widths=[24, 12, 16])

    for item in sec["items"]:
        S.set_cell(ws, f"A{row}", item["label"],
                   font_key="body_bold", fill_key="calc",
                   align_key="left", border_key="all_thin")
        S.set_cell(ws, f"B{row}", item.get("default_count", 0),
                   font_key="body_bold", fill_key="input",
                   align_key="center", border_key="all_thin",
                   number_format="integer")
        S.set_cell(ws, f"C{row}", item.get("default_date", ""),
                   font_key="body", fill_key="input",
                   align_key="center", border_key="all_thin")
        ws.row_dimensions[row].height = S.ROW_HEIGHT["default"]

        if item.get("count_named_range"):
            cell_ref = f"設定表!$B${row}"
            wb.defined_names[item["count_named_range"]] = DefinedName(
                name=item["count_named_range"], attr_text=cell_ref)
            log.info(f"  命名範圍：{item['count_named_range']} = {cell_ref}")
        row += 1

    log.info(f"  F 區（婦幼/協尋）完成，{len(sec['items'])} 項")
    return row


def build_section_h(ws, wb, row: int, sec: dict, log):
    """H 區 安全密碼 + 9 個下拉清單命名範圍"""
    row = add_section_banner(ws, row, "🔒", "H",
                              sec["name"],
                              desc="部署密碼 + 9 個下拉清單",
                              span_cols=3)

    # 密碼列
    row = add_table_header(ws, row, ["項目", "值", "說明"],
                            col_widths=[24, 16, 50])
    S.set_cell(ws, f"A{row}", "案件資料庫密碼",
               font_key="body_bold", fill_key="calc",
               align_key="left", border_key="all_thin")
    S.set_cell(ws, f"B{row}", sec["default_password"],
               font_key="body_bold", fill_key="input",
               align_key="center", border_key="all_thin")
    S.set_cell(ws, f"C{row}",
               "⚠ 預設 KD2026，強烈建議派出所自訂個別密碼",
               font_key="footnote_danger", fill_key="danger_bg",
               align_key="left", border_key="all_thin")
    ws.row_dimensions[row].height = S.ROW_HEIGHT["default"]
    row += 1

    row = add_spacer(ws, row, height=20)

    # 9 個下拉清單
    # 排兩欄：左清單 + 右清單 並排，較不浪費垂直空間
    # 但為求簡潔仍採一欄式
    for ld in sec["dropdown_lists"]:
        # 清單標題（小型 banner）
        ws.merge_cells(f"A{row}:C{row}")
        S.set_cell(ws, f"A{row}",
                   f"  📋 {ld['label']}（命名範圍：{ld['named_range']}）",
                   font_key="section_title",
                   fill_key="accent_light",
                   align_key="left",
                   border_key="bottom_medium")
        ws.row_dimensions[row].height = 26
        row += 1

        start_row = row
        for val in ld["values"]:
            S.set_cell(ws, f"A{row}", val,
                       font_key="body", fill_key="input",
                       align_key="center", border_key="all_thin")
            ws.row_dimensions[row].height = S.ROW_HEIGHT["default"]
            row += 1
        end_row = row - 1

        # 命名範圍
        cell_ref = f"設定表!$A${start_row}:$A${end_row}"
        wb.defined_names[ld["named_range"]] = DefinedName(
            name=ld["named_range"], attr_text=cell_ref)
        log.info(f"  命名範圍：{ld['named_range']} = {cell_ref} "
                 f"（{end_row - start_row + 1} 值）")

        row = add_spacer(ws, row, height=10)  # 清單之間小 spacer

    log.info(f"  H 區（安全密碼 + 下拉清單）完成，{len(sec['dropdown_lists'])} 個清單")
    return row


# ============================================================
# Main build
# ============================================================
def build():
    log = get_logger("phase1_5")
    log.info("===== Phase 1.5（設定表精緻版）開始 =====")

    if not OUTPUT_PATH.exists():
        log.error("Phase 1 output 不存在，請先跑 phase 1")
        sys.exit(1)

    cfg = load_config("phase1_5_settings")
    wb = load_workbook(OUTPUT_PATH)
    ws = wb["設定表"]

    # === 覆寫 Phase 1 的頂部 banner，改成設定表專屬 ===
    pb = cfg.get("page_banner", {})
    update_top_banner(ws,
                       title=pb.get("title", "派出所綜合管制系統 v2 — 設定表"),
                       subtitle=pb.get("subtitle",
                                        "全所 SSOT 底層 ｜ 22 個命名範圍 ｜ 部署時一次設定"),
                       version=pb.get("version", "KKEVIN-LIN-2026-V2.0"))

    # Phase 1 已佈置：
    #   row 1-2: banner + 副標
    #   row 3: spacer
    #   row 4: TODAY 集中區
    #   row 5: spacer
    # 接著加「新增與維護指引」資訊卡，然後各區從下方開始
    row = 6
    row = add_maintenance_guide(ws, row)
    row = add_spacer(ws, row, height=20)
    log.info("  📌 新增與維護指引 完成")

    # === 依序建構 8 大區 ===
    row = build_section_a(ws, wb, row, cfg["section_a"], log)
    row = add_spacer(ws, row, height=20)

    row = build_section_b(ws, wb, row, cfg["section_b"], log)
    row = add_spacer(ws, row, height=20)

    row = build_section_c(ws, wb, row, cfg["section_c"], log)
    row = add_spacer(ws, row, height=20)

    row = build_section_d(ws, wb, row, cfg["section_d"], log)
    row = add_spacer(ws, row, height=20)

    row = build_section_e(ws, wb, row, cfg["section_e"], log)
    row = add_spacer(ws, row, height=20)

    row = build_section_f(ws, wb, row, cfg["section_f"], log)
    row = add_spacer(ws, row, height=20)

    row = build_section_h(ws, wb, row, cfg["section_h"], log)

    # === 設定表頁籤顏色 ===
    ws.sheet_properties.tabColor = S.COLOR["accent"][2:]  # 去掉 alpha prefix

    # === 開檔時預設停在「設定表」（多重設定確保生效）===
    settings_idx = wb.sheetnames.index("設定表")
    wb.active = settings_idx
    # 關鍵：tabSelected 屬性才是 Excel 真正讀的旗標
    for i, sheet in enumerate(wb.worksheets):
        sheet.sheet_view.tabSelected = (i == settings_idx)
        sheet.sheet_view.topLeftCell = "A1"
    log.info(f"  開檔預設 sheet：設定表（index {settings_idx}）")

    # === 儲存 ===
    bak = backup_existing()
    if bak:
        log.info(f"  舊版備份：{bak}")
    wb.save(OUTPUT_PATH)
    size_kb = OUTPUT_PATH.stat().st_size / 1024
    log.info(f"  儲存：{OUTPUT_PATH.name} ({size_kb:.1f} KB)")
    log.info(f"  總命名範圍數：{len(wb.defined_names)}")
    log.info("===== Phase 1.5 完成 =====\n")
    return OUTPUT_PATH


if __name__ == "__main__":
    build()
