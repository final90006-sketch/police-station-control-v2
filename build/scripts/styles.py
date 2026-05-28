"""
styles.py — 集中樣式管理（顏色 / 字型 / 填色 / 邊框 / 對齊 / 數字格式）

所有顏色 / 字型 / 填色都集中在這，避免散落各頁。
顏色用 ARGB 格式（前 2 碼 alpha 通常為 FF）。
"""
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


# ============================================================
# 顏色（ARGB 格式）
# ============================================================
COLOR = {
    # 官方體例
    "accent":         "FF1F4E79",   # 官方藍（與 v43 一致）
    "accent_light":   "FFDBEAFE",   # 淡藍
    "accent_dark":    "FF14253D",   # 深藍（章戳用）

    # 狀態
    "pass":           "FF15803D",   # 通過綠
    "pass_bg":        "FFDCFCE7",   # 淡綠底
    "warn":           "FFB45309",   # 警告橘
    "warn_bg":        "FFFEF3C7",   # 淡黃橘
    "danger":         "FFB91C1C",   # 警示紅
    "danger_bg":      "FFFEE2E2",   # 淡紅底

    # 中性色
    "text":           "FF1E293B",   # 主文色
    "muted":          "FF64748B",   # 輔助色
    "border":         "FFE2E8F0",   # 淡邊框
    "code_bg":        "FFF1F5F9",   # 代碼底（淡灰）

    # Excel 體例
    "excel_header":   "FF1F4E79",   # 表頭藍
    "excel_input":    "FFFFF8E1",   # 淡黃輸入欄
    "excel_calc":     "FFF1EFE8",   # 淡灰計算欄

    # 白與黑
    "white":          "FFFFFFFF",
    "black":          "FF000000",
}


# ============================================================
# 字型（含繁中字體支援）— 加大、清晰、UI 級別
# ============================================================
FONT_FAMILY_DEFAULT = "微軟正黑體"  # Windows 標準繁中字
FONT_FAMILY_BANNER = "微軟正黑體"   # banner 用同字體保證一致

def font(size=12, bold=False, color="text", italic=False, family=None):
    """字型工廠：以 styles.font(size=14, bold=True, color='accent') 呼叫"""
    return Font(
        name=family or FONT_FAMILY_DEFAULT,
        size=size,
        bold=bold,
        italic=italic,
        color=COLOR.get(color, color),
    )


# 預設字型集合（全面加大 1-2pt）
FONT = {
    "body":           font(12),                           # 一般儲存格（11→12）
    "body_bold":      font(12, bold=True),
    "footnote":       font(10, color="muted"),
    "footnote_warn":  font(10, color="warn"),
    "footnote_danger":font(10, color="danger"),

    "header":         font(13, bold=True, color="white"), # 表頭（12→13）
    "section_title":  font(15, bold=True, color="white"), # 區塊標題（14→15，深藍底白字）
    "section_title_accent": font(15, bold=True, color="accent"),  # 區塊標題（無底色）
    "page_title":     font(22, bold=True, color="accent"),# 頁面標題（20→22）

    "kpi_value":      font(32, bold=True, color="accent"),# 大字 KPI（28→32）
    "kpi_value_warn": font(32, bold=True, color="warn"),
    "kpi_value_danger":font(32, bold=True, color="danger"),
    "kpi_value_pass": font(32, bold=True, color="pass"),
    "kpi_label":      font(12, color="muted"),

    "banner":         font(18, bold=True, color="white", family=FONT_FAMILY_BANNER),  # 章戳 banner（16→18）
    "sub_banner":     font(11, color="white"),            # 副標
}


# ============================================================
# 填色
# ============================================================
def fill(color_key):
    return PatternFill("solid", fgColor=COLOR.get(color_key, color_key))


FILL = {
    "header":         fill("excel_header"),     # 表頭深藍
    "input":          fill("excel_input"),      # 淡黃輸入
    "calc":           fill("excel_calc"),       # 淡灰計算
    "accent_light":   fill("accent_light"),     # 淡藍底
    "pass_bg":        fill("pass_bg"),
    "warn_bg":        fill("warn_bg"),
    "danger_bg":      fill("danger_bg"),
    "banner":         fill("accent"),           # 章戳深藍底
    "code_bg":        fill("code_bg"),          # 淡灰程式碼底
}


# ============================================================
# 邊框
# ============================================================
SIDE_THIN     = Side(border_style="thin",     color=COLOR["border"])
SIDE_MEDIUM   = Side(border_style="medium",   color=COLOR["accent"])
SIDE_THICK    = Side(border_style="thick",    color=COLOR["accent"])
SIDE_DASHED   = Side(border_style="dashed",   color=COLOR["muted"])
SIDE_DOTTED   = Side(border_style="dotted",   color=COLOR["muted"])

BORDER = {
    "all_thin":       Border(left=SIDE_THIN, right=SIDE_THIN, top=SIDE_THIN, bottom=SIDE_THIN),
    "bottom_thick":   Border(bottom=SIDE_THICK),
    "bottom_medium":  Border(bottom=SIDE_MEDIUM),
    "top_bottom":     Border(top=SIDE_THIN, bottom=SIDE_THIN),
    "header":         Border(left=SIDE_THIN, right=SIDE_THIN, top=SIDE_THIN, bottom=SIDE_MEDIUM),
    "section_divider":Border(bottom=SIDE_MEDIUM),
}


# ============================================================
# 對齊
# ============================================================
ALIGN = {
    "center":         Alignment(horizontal="center", vertical="center", wrap_text=True),
    "left":           Alignment(horizontal="left",   vertical="center", wrap_text=True),
    "right":          Alignment(horizontal="right",  vertical="center", wrap_text=True),
    "left_top":       Alignment(horizontal="left",   vertical="top",    wrap_text=True),
    "center_top":     Alignment(horizontal="center", vertical="top",    wrap_text=True),
    # 表頭專用：強制單行不折行（欄寬已足夠容納），標題整齊不長長短短
    "center_nowrap":  Alignment(horizontal="center", vertical="center", wrap_text=False),
}


# ============================================================
# 數字格式
# ============================================================
NUM_FMT = {
    "integer":        "#,##0",
    "money":          "#,##0",                # 元（無小數）
    "money_decimal":  "#,##0.00",
    "percent":        "0%",
    "percent_one":    "0.0%",
    "percent_two":    "0.00%",
    "date_roc":       'yyyy/mm/dd',           # 西元；ROC 顯示需另設
    "date_iso":       "yyyy-mm-dd",
    # datetime 改成「到幾點」（不顯示分鐘，v2.1 使用者要求）
    "datetime":       'yyyy/mm/dd h"時"',
    "datetime_full":  "yyyy/mm/dd hh:mm",     # 保留全精度給特殊場合
    "general":        "General",
}


# ============================================================
# 列高 / 欄寬（全面加大，更舒適）
# ============================================================
ROW_HEIGHT = {
    "default":  30,    # 加 4pt 給更多呼吸空間（26 → 30）
    "header":   34,    # 32 → 34
    "section_title": 38,
    "kpi":      60,
    "banner":   52,
    "sub_banner": 28,
    "spacer":   12,
}

COL_WIDTH = {
    "narrow":   10,
    "default":  16,
    "wide":     22,
    "text":     32,
    "long_text":50,
}

# 設定表通用欄寬（精緻寬鬆版，避免文字擠在窄欄）
SETTINGS_COL_WIDTHS = {
    "A": 26,    # 項目 / 編號（足夠放 8 中文字）
    "B": 48,    # 值 / 姓名 / 案類名稱（足夠放「高雄市政府警察局」+ 空間）
    "C": 52,    # 說明 / 職稱 / 案類分類（長中文說明）
    "D": 16,    # 在職 / 納入全般
    "E": 20,    # 納入統計 / 納入破獲率
    "F": 14,    # 排序 / 目標
    "G": 32,    # 備註 / 達成率警示
}


# ============================================================
# Helper：批次套用樣式
# ============================================================
def apply_to_range(ws, cell_range: str, *, font_key=None, fill_key=None,
                   border_key=None, align_key=None, number_format=None):
    """批次套用樣式到範圍"""
    for row in ws[cell_range]:
        for cell in row:
            if font_key:
                cell.font = FONT[font_key] if font_key in FONT else font_key
            if fill_key:
                cell.fill = FILL[fill_key] if fill_key in FILL else fill_key
            if border_key:
                cell.border = BORDER[border_key] if border_key in BORDER else border_key
            if align_key:
                cell.alignment = ALIGN[align_key] if align_key in ALIGN else align_key
            if number_format:
                cell.number_format = NUM_FMT.get(number_format, number_format)


def set_cell(ws, coord, value, *, font_key=None, fill_key=None,
             border_key=None, align_key=None, number_format=None):
    """設定單一儲存格 + 樣式"""
    cell = ws[coord]
    cell.value = value
    if font_key:
        cell.font = FONT[font_key] if font_key in FONT else font_key
    if fill_key:
        cell.fill = FILL[fill_key] if fill_key in FILL else fill_key
    if border_key:
        cell.border = BORDER[border_key] if border_key in BORDER else border_key
    if align_key:
        cell.alignment = ALIGN[align_key] if align_key in ALIGN else align_key
    if number_format:
        cell.number_format = NUM_FMT.get(number_format, number_format)
    return cell
