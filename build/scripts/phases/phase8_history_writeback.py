"""
phase8_history_writeback.py — Phase 8：頁 14 歷史資料即時 KPI 抄寫區

業務情境：
  - 月底承辦人要把當月 14 欄 KPI 寫入 Tbl歷史 一列
  - v43 純手抄表 → 易遺漏、易計算錯
  - v2.1 plan 提出：頁面頂部加「即時 KPI 抄寫區」公式自動算當月值
    使用者只需「複製整列 → 貼到下方主表」即完成抄寫

實作：
  Row 6  Section title「★ 本月即時 KPI 抄寫區」+ 動態抄寫狀態
         （已抄寫=綠／未抄寫=紅，由 D6 規則同源邏輯驅動）
  Row 7  14 欄即時 KPI 公式（同 Tbl歷史 column 結構）
         使用者複製此列 → 貼到 row 11+ 即完成抄寫
  Row 8-9 spacer

當月期間：DATE(YEAR(今日),MONTH(今日),1) ~ EOMONTH(今日,0)
  - EOMONTH 是 Excel 2007+ 內建，全平台相容

公式設計：
  - 大部分 KPI 由 Tbl案件 / Tbl毒調 / Tbl交通 即時聚合
  - 發展中 / 員警冠軍 用 placeholder（"—"）讓承辦人自填
    （理由：發展中沒 Excel Table 無法聚合；員警冠軍要 RANK + INDEX 太重）
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

from phases._base import backup_existing, OUTPUT_PATH, get_logger
import styles as S


SHEET = "歷史資料"
ROC_YEAR = "(YEAR(今日)-1911)"
CASE_TBL = "Tbl案件"
DRUG_TBL = "Tbl毒調"

# 當月期間範圍（共用 sub-formula）
MONTH_START = "DATE(YEAR(今日),MONTH(今日),1)"
MONTH_END = "EOMONTH(今日,0)"


def build_page14_writeback(wb, log):
    log.info("--- 頁 14 即時 KPI 抄寫區 ---")
    ws = wb[SHEET]

    # === Row 6 Section title + 抄寫狀態 ===
    # 動態：檢查 Tbl歷史[統計年月] 是否含本月，依此顯示綠/紅文字
    title_fml = (
        f'="   ★  本月即時 KPI 抄寫區（YY/MM = "&{ROC_YEAR}&"/"&'
        f'TEXT(MONTH(今日),"00")&"）  ｜  "'
        f'&IF(COUNTIF(Tbl歷史[統計年月],{ROC_YEAR}&"/"&TEXT(MONTH(今日),"00"))>=1,'
        f'"✓ 本月已抄寫","🚨 本月尚未抄寫 — 請複製下方整列貼到主表")'
    )
    ws.merge_cells("A6:N6")
    S.set_cell(ws, "A6", title_fml,
               font_key=S.font(12, bold=True, color="white"),
               fill_key=S.fill("accent_dark"),
               align_key="left", border_key="all_thin")
    ws.row_dimensions[6].height = 30

    # === Row 7 即時 KPI 14 欄公式（Tbl歷史 column 結構同步）===
    # 14 欄定義：(column letter, value formula, number format)
    kpis = [
        # 1. 統計年月（字串）
        ("A", f'={ROC_YEAR}&"/"&TEXT(MONTH(今日),"00")', "general"),
        # 2. 全般發生（本轄當月，計入發生數=是）
        ("B", (f'=IFERROR(COUNTIFS({CASE_TBL}[發生時間],">="&{MONTH_START},'
               f'{CASE_TBL}[發生時間],"<="&{MONTH_END},'
               f'{CASE_TBL}[發生管轄],"本轄",'
               f'{CASE_TBL}[計入發生數],"是"),0)'), "integer"),
        # 3. 全般未破（本轄當月，案件狀況=尚未偵破）
        ("C", (f'=IFERROR(COUNTIFS({CASE_TBL}[發生時間],">="&{MONTH_START},'
               f'{CASE_TBL}[發生時間],"<="&{MONTH_END},'
               f'{CASE_TBL}[發生管轄],"本轄",'
               f'{CASE_TBL}[案件狀況],"尚未偵破"),0)'), "integer"),
        # 4. 全般破獲（查獲本轄當月，是否破獲=是）
        ("D", (f'=IFERROR(COUNTIFS({CASE_TBL}[破獲時間],">="&{MONTH_START},'
               f'{CASE_TBL}[破獲時間],"<="&{MONTH_END},'
               f'{CASE_TBL}[查獲管轄],"本轄",'
               f'{CASE_TBL}[是否破獲],"是"),0)'), "integer"),
        # 5. 破獲率 = 破獲 / 發生
        ("E", '=IFERROR(D7/B7,0)', "percent_one"),
        # 6. 竊盜 (本轄當月 案類分類=竊盜)
        ("F", (f'=IFERROR(COUNTIFS({CASE_TBL}[發生時間],">="&{MONTH_START},'
               f'{CASE_TBL}[發生時間],"<="&{MONTH_END},'
               f'{CASE_TBL}[發生管轄],"本轄",'
               f'{CASE_TBL}[案類分類],"竊盜"),0)'), "integer"),
        # 7. 詐欺
        ("G", (f'=IFERROR(COUNTIFS({CASE_TBL}[發生時間],">="&{MONTH_START},'
               f'{CASE_TBL}[發生時間],"<="&{MONTH_END},'
               f'{CASE_TBL}[發生管轄],"本轄",'
               f'{CASE_TBL}[案類分類],"詐欺"),0)'), "integer"),
        # 8. 已破獲未移送（查獲本轄當月）
        ("H", (f'=IFERROR(COUNTIFS({CASE_TBL}[破獲時間],">="&{MONTH_START},'
               f'{CASE_TBL}[破獲時間],"<="&{MONTH_END},'
               f'{CASE_TBL}[查獲管轄],"本轄",'
               f'{CASE_TBL}[案件狀況],"已破獲未移送"),0)'), "integer"),
        # 9. 發展中（頁 7 沒 Excel Table，承辦自填）
        ("I", '="—"', "general"),
        # 10. 毒調率（即時 Tbl毒調 公式）
        ("J", (f'=IFERROR(SUM({DRUG_TBL}[是否到驗])/'
               f'COUNTA({DRUG_TBL}[姓名]),0)'), "percent_one"),
        # 11. 未到驗
        ("K", f'=IFERROR(COUNTIF({DRUG_TBL}[管制情形],"*未到驗*"),0)', "integer"),
        # 12. 交通達標（引用頁 9 E8 — 總達成率大字）
        ("L", '=IFERROR(交通績效管制!$E$8,0)', "percent_one"),
        # 13. 員警冠軍（v2.2 自動算）— INDEX+MATCH+MAX 取本月破獲最高員警
        #   依賴 helper area row 50-110 計算每員警的當月破獲數
        ("M",
            '=IFERROR(INDEX($R$50:$R$110,MATCH(MAX($S$50:$S$110),$S$50:$S$110,0)),"—")',
            "general"),
        # 14. 備註（手填）
        ("N", '=""', "general"),
    ]

    for col, fml, fmt in kpis:
        S.set_cell(ws, f"{col}7", fml,
                   font_key="body_bold",
                   fill_key="warn_bg",   # 淡黃，提示「這列是要複製的」
                   align_key="center" if col != "N" else "left",
                   border_key="all_thin",
                   number_format=fmt)
    ws.row_dimensions[7].height = 32

    # === Row 8 操作說明 ===
    ws.merge_cells("A8:N8")
    S.set_cell(ws, "A8",
               "   📋 操作流程：①  選 A7:N7 整列 → 複製（Ctrl+C）  "
               "②  到下方主表第一筆空白列 → 貼上值（Ctrl+Shift+V → 值）  "
               "③  按 Enter 完成抄寫（上方紅燈將自動轉綠）",
               font_key=S.font(10, color="muted", italic=True),
               fill_key="code_bg", align_key="left",
               border_key="all_thin")
    ws.row_dimensions[8].height = 24
    ws.row_dimensions[9].height = 8

    # === Row 6 條件格式：已抄寫=綠底 / 未抄寫=紅底 ===
    # ★ 用同 sheet 普通範圍引用避免結構化引用觸發 Excel 修復警示
    # Tbl歷史 統計年月在 A11:A150（capacity 140 + header）
    cap_count = ('COUNTIF($A$11:$A$150,'
                 f'({ROC_YEAR}&"/"&TEXT(MONTH(今日),"00")))')

    # 清掉既有 CF（冪等）
    ws.conditional_formatting._cf_rules = {}

    # 紅燈：本月未抄寫
    ws.conditional_formatting.add("A6",
        FormulaRule(formula=[f'{cap_count}=0'],
                    fill=PatternFill("solid", fgColor="FFB91C1C"),
                    font=Font(bold=True, color="FFFFFFFF"),
                    stopIfTrue=True))
    # 綠燈：本月已抄寫
    ws.conditional_formatting.add("A6",
        FormulaRule(formula=[f'{cap_count}>=1'],
                    fill=PatternFill("solid", fgColor="FF15803D"),
                    font=Font(bold=True, color="FFFFFFFF"),
                    stopIfTrue=True))

    # ==========================================================
    # Helper 區（rows 49-110，隱藏）：每員警的當月破獲數
    # 用於 M7 員警冠軍 INDEX+MATCH+MAX 自動算
    # ==========================================================
    S.set_cell(ws, "R49", "[helper] 員警",
               font_key=S.font(9, color="muted"))
    S.set_cell(ws, "S49", "當月破獲數",
               font_key=S.font(9, color="muted"))

    # 員警名冊：設定表!$B$24:$B$83 (60 員警)
    for k in range(1, 61):   # rows 50-109
        helper_row = 49 + k
        officer_src = f"設定表!$B${23+k}"
        # R 員警姓名
        S.set_cell(ws, f"R{helper_row}",
                   f'=IFERROR(IF({officer_src}="","",{officer_src}),"")',
                   font_key=S.font(9, color="muted"))
        # S 該員警當月破獲數
        s_fml = (
            f'=IFERROR(IF(R{helper_row}="",0,'
            f'COUNTIFS(Tbl案件[承辦人],R{helper_row},'
            f'Tbl案件[是否破獲],"是",'
            f'Tbl案件[破獲時間],">="&{MONTH_START},'
            f'Tbl案件[破獲時間],"<="&{MONTH_END})),0)'
        )
        S.set_cell(ws, f"S{helper_row}", s_fml,
                   font_key=S.font(9, color="muted"),
                   number_format="integer")

    # 隱藏 helper rows + cols R/S
    for r in range(49, 111):
        ws.row_dimensions[r].hidden = True
    ws.column_dimensions["R"].hidden = True
    ws.column_dimensions["S"].hidden = True

    log.info("  Row 6: section title + 抄寫狀態（CF 紅綠燈）")
    log.info("  Row 7: 14 欄即時 KPI 公式（含 M7 員警冠軍自動算）")
    log.info("  Row 8: 操作流程說明")
    log.info("  Helper rows 50-109: 60 員警 × 當月破獲數（隱藏）")
    log.info("  CF: 2 條（已抄寫綠 / 未抄寫紅）")


def build():
    log = get_logger("phase8_history")
    log.info("===== Phase 8 — 頁 14 即時 KPI 抄寫區 開始 =====")

    wb = load_workbook(OUTPUT_PATH)
    build_page14_writeback(wb, log)

    bak = backup_existing()
    if bak:
        log.info(f"  舊版備份：{bak}")
    wb.save(OUTPUT_PATH)
    size_kb = OUTPUT_PATH.stat().st_size / 1024
    log.info(f"  儲存：{OUTPUT_PATH.name} ({size_kb:.1f} KB)")
    log.info("===== Phase 8 完成 =====\n")
    return OUTPUT_PATH


if __name__ == "__main__":
    build()
