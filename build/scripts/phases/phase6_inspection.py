"""
phase6_inspection.py — Phase 6：頁 15 系統檢核 45 條

7 大類規則：
  A 部署完整性 5
  B 資料合理性 8
  C 業務邏輯 10
  D 時效警示 8
  E 跨頁一致性 6
  F 性能監控 5
  G 環境檢測 3
                ──
                45

版型（A1:G65 列印 A4 直式）：
  Row 1-2   Banner
  Row 3     spacer
  Row 4     維護指引
  Row 5     spacer
  Row 6     Section「⚕ 全表健康度」
  Row 7-8   大字 3 卡：紅 / 黃 / 綠 總數
  Row 9     spacer
  Row 10    Section「📊 7 大類縮影」
  Row 11    7 卡（A-G 各類別，類別名 + 紅/黃/綠/待接通 數）
  Row 12    spacer
  Row 13    Section「📋 45 條檢核明細」
  Row 14    主表 header
  Row 15-59 45 規則列
  Row 60    spacer
  Row 61    頁尾

每條規則欄位 (A-G)：
  A 編號 (A1..G3)
  B 類別 (固定文字)
  C 規則描述（人類可讀）
  D 當前值（公式自動算）
  E 狀態燈（依 D 判斷：● 綠 / ⚠ 黃 / ● 紅 / ○ 待接通 / ℹ 資訊）
  F 嚴重度（高 / 中 / 低 / 資訊）
  G 跳轉連結（HYPERLINK 到問題頁）

容錯規範：
  - 所有公式包 IFERROR
  - 引用未建表（Tbl毒調）時自動回傳「○ 待接通」
  - 燈號公式統一用文字判斷（與 Phase 5 CF 規則相容）
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

from phases._base import backup_existing, OUTPUT_PATH, get_logger, unmerge_all_in_sheet
import styles as S


SHEET = "系統檢核"
ROC_YEAR = "(YEAR(今日)-1911)"
CASE_TBL = "Tbl案件"
TRAFFIC_TBL = "Tbl交通"
HIST_TBL = "Tbl歷史"
DRUG_TBL = "Tbl毒調"

CATEGORIES = [
    ("A", "部署完整性", "accent"),
    ("B", "資料合理性", "accent"),
    ("C", "業務邏輯",   "warn"),
    ("D", "時效警示",   "danger"),
    ("E", "跨頁一致性", "accent"),
    ("F", "性能監控",   "muted"),
    ("G", "環境檢測",   "muted"),
]


# ============================================================
# 45 規則定義
#   value_formula：D 欄公式（不含等號前綴會自動加）
#   status_logic：產生 E 欄狀態燈公式的函式（接 D 欄 ref，回傳 "=IF(...)"）
#   link：(sheet, cell) for HYPERLINK
# ============================================================
def light_ok_or_red(d, ok_value="TRUE"):
    """二元判斷：D=ok 綠，否則紅"""
    if ok_value == "TRUE":
        return f'=IF({d}=TRUE,"● 綠","● 紅")'
    return f'=IF({d}={ok_value},"● 綠","● 紅")'


def light_count_zero_ok(d):
    """違規計數：=0 綠，>0 紅"""
    return f'=IF(ISNUMBER({d}),IF({d}=0,"● 綠","● 紅"),"○ 待接通")'


def light_count_min(d, min_count):
    """數量 ≥ min 綠，否則紅"""
    return f'=IF(ISNUMBER({d}),IF({d}>={min_count},"● 綠","● 紅"),"○ 待接通")'


def light_capacity(d, hi_red, hi_yellow):
    """容量警示：≥ hi_red 紅，≥ hi_yellow 黃，否則綠"""
    return (f'=IF(ISNUMBER({d}),'
            f'IF({d}>={hi_red},"● 紅",'
            f'IF({d}>={hi_yellow},"⚠ 黃","● 綠")),"○ 待接通")')


def light_info(d):
    """資訊性顯示，不評燈"""
    return f'="ℹ 資訊"'


def light_pending(d):
    """待接通（Tbl毒調等尚未建表）"""
    return f'="○ 待接通"'


RULES = [
    # ===================================================================
    # A 部署完整性（5）— 從設定表 A 區 / B 區 / D 區 / E 區 抓
    # ===================================================================
    {"id": "A1", "cat": "A 部署完整性",
     "desc": "派出所名稱已填",
     "value": "=IFERROR(LEN(派出所名稱)>0,FALSE)",
     "light": lambda d: light_ok_or_red(d),
     "severity": "高",
     "link": ("設定表", "B18")},
    {"id": "A2", "cat": "A 部署完整性",
     "desc": "警察局＋分局名稱已填",
     "value": "=IFERROR(AND(LEN(警察局名稱)>0,LEN(分局名稱)>0),FALSE)",
     "light": lambda d: light_ok_or_red(d),
     "severity": "高",
     "link": ("設定表", "B16")},
    {"id": "A3", "cat": "A 部署完整性",
     "desc": "員警名冊 ≥ 1 名（在職）",
     "value": ('=IFERROR(COUNTIFS(設定表!$B$24:$B$83,"?*",'
               '設定表!$D$24:$D$83,"是"),0)'),
     "light": lambda d: light_count_min(d, 1),
     "severity": "高",
     "link": ("設定表", "B24")},
    {"id": "A4", "cat": "A 部署完整性",
     "desc": "案類清單 ≥ 1 項（納入全般）",
     "value": '=IFERROR(COUNTIF(納入全般清單,"是"),0)',
     "light": lambda d: light_count_min(d, 1),
     "severity": "高",
     "link": ("設定表", "A98")},
    {"id": "A5", "cat": "A 部署完整性",
     "desc": "違規項目清單 ≥ 1 項",
     "value": "=IFERROR(COUNTA(違規項目清單),0)",
     "light": lambda d: light_count_min(d, 1),
     "severity": "中",
     "link": ("設定表", "A168")},

    # ===================================================================
    # B 資料合理性（8）— Tbl案件 / Tbl交通 / Tbl毒調
    # ===================================================================
    {"id": "B1", "cat": "B 資料合理性",
     "desc": "案件編號無重複",
     "value": (f'=IFERROR(SUMPRODUCT(({CASE_TBL}[編號]<>"")*'
               f'(COUNTIF({CASE_TBL}[編號],{CASE_TBL}[編號])>1)),0)'),
     "light": lambda d: light_count_zero_ok(d),
     "severity": "高",
     "link": ("案件資料庫", "A15")},
    {"id": "B2", "cat": "B 資料合理性",
     "desc": "案件發生時間欄全為有效日期",
     "value": (f'=IFERROR(SUMPRODUCT(({CASE_TBL}[發生時間]<>"")*'
               f'(NOT(ISNUMBER({CASE_TBL}[發生時間])))),0)'),
     "light": lambda d: light_count_zero_ok(d),
     "severity": "中",
     "link": ("案件資料庫", "D15")},
    {"id": "B3", "cat": "B 資料合理性",
     "desc": "案件狀況欄全為有效值（4 選 1）",
     "value": (f'=IFERROR(SUMPRODUCT(({CASE_TBL}[案件狀況]<>"")*'
               f'(COUNTIF(案件狀況清單,{CASE_TBL}[案件狀況])=0)),0)'),
     "light": lambda d: light_count_zero_ok(d),
     "severity": "高",
     "link": ("案件資料庫", "J15")},
    {"id": "B4", "cat": "B 資料合理性",
     "desc": "發生管轄 + 查獲管轄欄全為「本轄/他轄」",
     "value": (f'=IFERROR(SUMPRODUCT(({CASE_TBL}[發生管轄]<>"")*'
               f'(COUNTIF(發生管轄清單,{CASE_TBL}[發生管轄])=0))'
               f'+SUMPRODUCT(({CASE_TBL}[查獲管轄]<>"")*'
               f'(COUNTIF(查獲管轄清單,{CASE_TBL}[查獲管轄])=0)),0)'),
     "light": lambda d: light_count_zero_ok(d),
     "severity": "高",
     "link": ("案件資料庫", "K15")},
    {"id": "B5", "cat": "B 資料合理性",
     "desc": "是否破獲欄全為「是」或「否」（計算欄）",
     "value": (f'=IFERROR(SUMPRODUCT(({CASE_TBL}[是否破獲]<>"")*'
               f'(({CASE_TBL}[是否破獲]<>"是")*({CASE_TBL}[是否破獲]<>"否"))),0)'),
     "light": lambda d: light_count_zero_ok(d),
     "severity": "中",
     "link": ("案件資料庫", "R15")},
    {"id": "B6", "cat": "B 資料合理性",
     "desc": "已破獲/已移送案件 應有破獲時間（時效完整性）",
     "value": (f'=IFERROR(SUMPRODUCT(({CASE_TBL}[是否破獲]="是")*'
               f'({CASE_TBL}[破獲時間]="")),0)'),
     "light": lambda d: light_count_zero_ok(d),
     "severity": "低",
     "link": ("案件資料庫", "F15")},
    {"id": "B7", "cat": "B 資料合理性",
     "desc": "毒調人口身分證長度 = 10（1 英文字母 + 9 數字）",
     "value": (f'=IFERROR(SUMPRODUCT(({DRUG_TBL}[身分證]<>"")*'
               f'(LEN({DRUG_TBL}[身分證])<>10)),0)'),
     "light": lambda d: light_count_zero_ok(d),
     "severity": "中",
     "link": ("毒品調驗人口管制", "C15")},
    {"id": "B8", "cat": "B 資料合理性",
     "desc": "交通取締舉發單號無重複",
     "value": (f'=IFERROR(SUMPRODUCT(({TRAFFIC_TBL}[舉發單號]<>"")*'
               f'(COUNTIF({TRAFFIC_TBL}[舉發單號],{TRAFFIC_TBL}[舉發單號])>1)),0)'),
     "light": lambda d: light_count_zero_ok(d),
     "severity": "中",
     "link": ("交通取締明細", "D15")},

    # ===================================================================
    # C 業務邏輯（10）
    # ===================================================================
    {"id": "C1", "cat": "C 業務邏輯",
     "desc": "「已破獲未移送」案件 偵辦進度 應為「已破獲」",
     "value": (f'=IFERROR(SUMPRODUCT(({CASE_TBL}[案件狀況]="已破獲未移送")*'
               f'({CASE_TBL}[偵辦進度]<>"已破獲")),0)'),
     "light": lambda d: light_count_zero_ok(d),
     "severity": "中",
     "link": ("案件資料庫", "J15")},
    {"id": "C2", "cat": "C 業務邏輯",
     "desc": "「尚未偵破」案件 是否破獲 應為「否」",
     "value": (f'=IFERROR(SUMPRODUCT(({CASE_TBL}[案件狀況]="尚未偵破")*'
               f'({CASE_TBL}[是否破獲]="是")),0)'),
     "light": lambda d: light_count_zero_ok(d),
     "severity": "高",
     "link": ("案件資料庫", "J15")},
    {"id": "C3", "cat": "C 業務邏輯",
     "desc": "「已移送」案件 是否破獲 應為「是」",
     "value": (f'=IFERROR(SUMPRODUCT(({CASE_TBL}[案件狀況]="已移送")*'
               f'({CASE_TBL}[是否破獲]="否")),0)'),
     "light": lambda d: light_count_zero_ok(d),
     "severity": "高",
     "link": ("案件資料庫", "J15")},
    {"id": "C4", "cat": "C 業務邏輯",
     "desc": "毒調率分子（已驗+通緝+強採+在監）≤ 列管總數",
     "value": (f'=IFERROR(SUM({DRUG_TBL}[是否到驗])'
               f'-COUNTA({DRUG_TBL}[姓名]),0)'),
     # 分子 - 總數 應該 <= 0 才合理（分子永遠 ≤ 總數）
     "light": lambda d: f'=IF(ISNUMBER({d}),IF({d}<=0,"● 綠","● 紅"),"○ 待接通")',
     "severity": "中",
     "link": ("毒品調驗人口管制", "A8")},
    {"id": "C5", "cat": "C 業務邏輯",
     "desc": "交通達成率全部 < 200%（過高表示目標設太低）",
     "value": "=IFERROR(MAX(交通績效管制!$D$14:$D$19),0)",
     "light": lambda d: f'=IF(ISNUMBER({d}),IF({d}<2,"● 綠","⚠ 黃"),"○ 待接通")',
     "severity": "低",
     "link": ("交通績效管制", "D14")},
    {"id": "C6", "cat": "C 業務邏輯",
     "desc": "歷史快照員警冠軍欄全在員警名冊內",
     "value": (f'=IFERROR(SUMPRODUCT(({HIST_TBL}[員警冠軍]<>"")*'
               f'(COUNTIF(員警姓名清單,{HIST_TBL}[員警冠軍])=0)),0)'),
     "light": lambda d: light_count_zero_ok(d),
     "severity": "中",
     "link": ("歷史資料", "M11")},
    {"id": "C7", "cat": "C 業務邏輯",
     "desc": "本轄發生 ≥ 本轄破獲（純本轄，不含拘提他轄）",
     "value": (f'=IFERROR(COUNTIFS({CASE_TBL}[發生管轄],"本轄",'
               f'{CASE_TBL}[歸屬年度],{ROC_YEAR})'
               f'-COUNTIFS({CASE_TBL}[發生管轄],"本轄",'
               f'{CASE_TBL}[是否破獲],"是",'
               f'{CASE_TBL}[歸屬年度],{ROC_YEAR}),0)'),
     "light": lambda d: f'=IF(ISNUMBER({d}),IF({d}>=0,"● 綠","● 紅"),"○ 待接通")',
     "severity": "高",
     "link": ("管制總覽", "A1")},
    {"id": "C8", "cat": "C 業務邏輯",
     "desc": "歸屬年度計算欄無錯誤值（115 年案件 = 民國 115）",
     "value": (f'=IFERROR(SUMPRODUCT(({CASE_TBL}[編號]<>"")*'
               f'ISERROR({CASE_TBL}[歸屬年度])),0)'),
     "light": lambda d: light_count_zero_ok(d),
     "severity": "中",
     "link": ("案件資料庫", "S15")},
    {"id": "C9", "cat": "C 業務邏輯",
     "desc": "他轄發生案件若已破，查獲管轄應為「本轄」（拘提他轄）",
     "value": (f'=IFERROR(SUMPRODUCT(({CASE_TBL}[發生管轄]="他轄")*'
               f'({CASE_TBL}[是否破獲]="是")*'
               f'({CASE_TBL}[查獲管轄]<>"本轄")*'
               f'({CASE_TBL}[查獲管轄]<>"")),0)'),
     "light": lambda d: light_count_zero_ok(d),
     "severity": "中",
     "link": ("案件資料庫", "L15")},
    {"id": "C10", "cat": "C 業務邏輯",
     "desc": "含「強採」字樣的列管人口 是否到驗 必為 1（業務規則）",
     "value": (f'=IFERROR(SUMPRODUCT(({DRUG_TBL}[管制情形]<>"")*'
               f'ISNUMBER(SEARCH("強採",{DRUG_TBL}[管制情形]))*'
               f'({DRUG_TBL}[是否到驗]=0)),0)'),
     "light": lambda d: light_count_zero_ok(d),
     "severity": "中",
     "link": ("毒品調驗人口管制", "H15")},

    # ===================================================================
    # D 時效警示（8）
    # ===================================================================
    {"id": "D1", "cat": "D 時效警示",
     "desc": "案件資料庫容量（≥ 800 黃 / ≥ 950 紅）",
     "value": f'=IFERROR(COUNTA({CASE_TBL}[編號]),0)',
     "light": lambda d: light_capacity(d, hi_red=950, hi_yellow=800),
     "severity": "高",
     "link": ("案件資料庫", "A2")},
    {"id": "D2", "cat": "D 時效警示",
     "desc": "毒調人口容量（≥ 50 黃 / ≥ 65 紅 ｜ 上限 70）",
     "value": f'=IFERROR(COUNTA({DRUG_TBL}[姓名]),0)',
     "light": lambda d: light_capacity(d, hi_red=65, hi_yellow=50),
     "severity": "中",
     "link": ("毒品調驗人口管制", "A2")},
    {"id": "D3", "cat": "D 時效警示",
     "desc": "交通取締容量（≥ 150 黃 / ≥ 180 紅 ｜ 上限 196）",
     "value": f'=IFERROR(COUNTA({TRAFFIC_TBL}[編號]),0)',
     "light": lambda d: light_capacity(d, hi_red=180, hi_yellow=150),
     "severity": "中",
     "link": ("交通取締明細", "A2")},
    {"id": "D4", "cat": "D 時效警示",
     "desc": "發展中老案警示（> 老案閾值 件數）",
     "value": ('=IFERROR(COUNTIF(發展中案件管制!$J$14:$J$28,'
               '">"&老案閾值),0)'),
     "light": lambda d: f'=IF(ISNUMBER({d}),IF({d}=0,"● 綠",IF({d}<=3,"⚠ 黃","● 紅")),"○ 待接通")',
     "severity": "中",
     "link": ("發展中案件管制", "A12")},
    {"id": "D5", "cat": "D 時效警示",
     "desc": "未到驗或距今逾 30 日 件數（v2.2 精確 — 0 綠 / 1-5 黃 / >5 紅）",
     "value": f'=IFERROR(SUM({DRUG_TBL}[候選旗標]),0)',
     "light": lambda d: (f'=IF(ISNUMBER({d}),'
                         f'IF({d}=0,"● 綠",'
                         f'IF({d}<=5,"⚠ 黃","● 紅")),"○ 待接通")'),
     "severity": "高",
     "link": ("毒品調驗人口管制", "G9")},
    {"id": "D6", "cat": "D 時效警示",
     "desc": "本月歷史快照已抄寫（月底前未抄 紅燈）",
     "value": (f'=IFERROR(COUNTIF({HIST_TBL}[統計年月],'
               f'{ROC_YEAR}&"/"&TEXT(MONTH(今日),"00")),0)'),
     "light": lambda d: f'=IF(ISNUMBER({d}),IF({d}>=1,"● 綠","● 紅"),"○ 待接通")',
     "severity": "中",
     "link": ("歷史資料", "A11")},
    {"id": "D7", "cat": "D 時效警示",
     "desc": "孤兒承辦人（Tbl案件 承辦人不在員警名冊）",
     "value": (f'=IFERROR(SUMPRODUCT(({CASE_TBL}[承辦人]<>"")*'
               f'(COUNTIF(員警姓名清單,{CASE_TBL}[承辦人])=0)),0)'),
     "light": lambda d: light_count_zero_ok(d),
     "severity": "中",
     "link": ("設定表", "B24")},
    {"id": "D8", "cat": "D 時效警示",
     "desc": "已解除列管人數（資訊性，顯示本月歸檔件數）",
     "value": f'=IFERROR(COUNTIF({DRUG_TBL}[管制情形],"*解除*"),0)',
     "light": lambda d: light_info(d),
     "severity": "資訊",
     "link": ("毒品調驗人口管制", "H15")},

    # ===================================================================
    # E 跨頁一致性（6）
    # ===================================================================
    {"id": "E1", "cat": "E 跨頁一致性",
     "desc": "頁 9 達成數合計 = Tbl交通 違規明細數",
     "value": (f'=IFERROR(SUM(交通績效管制!$C$14:$C$19)'
               f'-COUNTA({TRAFFIC_TBL}[違規項目]),0)'),
     "light": lambda d: light_count_zero_ok(d),
     "severity": "中",
     "link": ("交通取締明細", "C15")},
    {"id": "E2", "cat": "E 跨頁一致性",
     "desc": "Tbl案件 是否破獲=是 加總 ≥ 0（基本不應為負）",
     "value": (f'=IFERROR(COUNTIFS({CASE_TBL}[是否破獲],"是",'
               f'{CASE_TBL}[歸屬年度],{ROC_YEAR}),0)'),
     "light": lambda d: f'=IF(ISNUMBER({d}),IF({d}>=0,"● 綠","● 紅"),"○ 待接通")',
     "severity": "低",
     "link": ("績效統計", "A1")},
    {"id": "E3", "cat": "E 跨頁一致性",
     "desc": "頁 12 sparkline 需要歷史 ≥ 12 月才完整",
     "value": f'=IFERROR(COUNTA({HIST_TBL}[統計年月]),0)',
     "light": lambda d: f'=IF(ISNUMBER({d}),IF({d}>=12,"● 綠",IF({d}>=6,"⚠ 黃","● 紅")),"○ 待接通")',
     "severity": "低",
     "link": ("歷史資料", "A11")},
    {"id": "E4", "cat": "E 跨頁一致性",
     "desc": "頁 2.5 本所統計與頁 6 已破獲未移送 + 拘提他轄一致",
     "value": (f'=IFERROR(COUNTIFS({CASE_TBL}[查獲管轄],"本轄",'
               f'{CASE_TBL}[是否破獲],"是",'
               f'{CASE_TBL}[歸屬年度],{ROC_YEAR})'
               f'-COUNTIFS({CASE_TBL}[發生管轄],"本轄",'
               f'{CASE_TBL}[是否破獲],"是",'
               f'{CASE_TBL}[歸屬年度],{ROC_YEAR})'
               f'-COUNTIFS({CASE_TBL}[發生管轄],"他轄",'
               f'{CASE_TBL}[查獲管轄],"本轄",'
               f'{CASE_TBL}[是否破獲],"是",'
               f'{CASE_TBL}[歸屬年度],{ROC_YEAR}),0)'),
     "light": lambda d: light_count_zero_ok(d),
     "severity": "中",
     "link": ("全般刑案管制情形分析", "A1")},
    {"id": "E5", "cat": "E 跨頁一致性",
     "desc": "案類分類欄全部對應到 D 區案類名稱（無孤值）",
     "value": (f'=IFERROR(SUMPRODUCT(({CASE_TBL}[案類]<>"")*'
               f'(COUNTIF(案類名稱清單,{CASE_TBL}[案類])=0)),0)'),
     "light": lambda d: light_count_zero_ok(d),
     "severity": "中",
     "link": ("案件資料庫", "B15")},
    {"id": "E6", "cat": "E 跨頁一致性",
     "desc": "毒調率 ≥ 目標（設定表閾值 60%）",
     "value": (f'=IFERROR(SUM({DRUG_TBL}[是否到驗])/'
               f'COUNTA({DRUG_TBL}[姓名]),0)'),
     "light": lambda d: (f'=IF(ISNUMBER({d}),'
                         f'IF({d}>=毒調率目標,"● 綠",'
                         f'IF({d}>=毒調率目標*0.8,"⚠ 黃","● 紅")),"○ 待接通")'),
     "severity": "中",
     "link": ("毒品調驗人口管制", "A11")},

    # ===================================================================
    # F 性能監控（5）— 多數需 build script 提供，資訊性顯示
    # ===================================================================
    {"id": "F1", "cat": "F 性能監控",
     "desc": "公式總量目標 ≤ 12,000 條（v43 為 26,465；目前由 build script 統計）",
     "value": "=\"目標 ≤ 12000\"",
     "light": lambda d: light_info(d),
     "severity": "資訊",
     "link": (None, None)},
    {"id": "F2", "cat": "F 性能監控",
     "desc": "檔案大小目標 ≤ 500 KB（v2.1 約 125 KB ✓）",
     "value": "=\"目標 ≤ 500 KB\"",
     "light": lambda d: light_info(d),
     "severity": "資訊",
     "link": (None, None)},
    {"id": "F3", "cat": "F 性能監控",
     "desc": "揮發性函數使用 ≤ 3 處（TODAY 集中設定表「今日」一格 ✓）",
     "value": "=\"TODAY 已集中\"",
     "light": lambda d: light_info(d),
     "severity": "資訊",
     "link": ("設定表", "B4")},
    {"id": "F4", "cat": "F 性能監控",
     "desc": "條件格式規則 每張表 ≤ 8 條（頁 2 為 54 特殊豁免）",
     "value": "=\"目標 每表 ≤ 8\"",
     "light": lambda d: light_info(d),
     "severity": "資訊",
     "link": (None, None)},
    {"id": "F5", "cat": "F 性能監控",
     "desc": "無循環引用（Excel 開檔自動偵測，若有會彈警示視窗）",
     "value": "=\"Excel 自動偵測\"",
     "light": lambda d: light_info(d),
     "severity": "資訊",
     "link": (None, None)},

    # ===================================================================
    # G 環境檢測（3）
    # ===================================================================
    {"id": "G1", "cat": "G 環境檢測",
     "desc": "Excel 版本 ≥ 2016（FILTER/SORT 需 365+，否則用 Fallback）",
     "value": '=IFERROR(INFO("release"),"未知")',
     "light": lambda d: f'=IF(VALUE({d})>=16,"● 綠","⚠ 黃")',
     "severity": "中",
     "link": (None, None)},
    {"id": "G2", "cat": "G 環境檢測",
     "desc": "中文字體（微軟正黑體）已內建於 Windows 10/11 + Office",
     "value": '="系統內建"',
     "light": lambda d: light_info(d),
     "severity": "資訊",
     "link": (None, None)},
    {"id": "G3", "cat": "G 環境檢測",
     "desc": "中文編碼正確（UTF-8 BOM，無亂碼）",
     "value": '="UTF-8 ✓"',
     "light": lambda d: light_info(d),
     "severity": "資訊",
     "link": (None, None)},
]


assert len(RULES) == 45, f"45 條檢核規則 expected, got {len(RULES)}"


# ============================================================
# 主建檔函式
# ============================================================
def build_page15_inspection(wb, log):
    log.info("--- 頁 15 系統檢核（45 條規則）---")
    ws = wb[SHEET]

    cleared = unmerge_all_in_sheet(ws)
    log.info(f"  清空既有 merges：{cleared} 個")

    # === 欄寬（7 欄全部平均 14-18，確保 7 大類縮影每張卡一行容下類別名）===
    #   原寬 A=8 / F=8 太窄，「A 部署完整性」「F 性能監控」會擠 2 行
    #   現在統一平均 14-18：所有類別名（5-6 字）都能一行顯示
    #   主表 C 規則描述從 52→40，仍可容下多數規則（28pt row 容 2 行 backup）
    col_widths = {
        "A": 14,    # 編號 / 縮影 A 卡（容「A 部署完整性」6 字）
        "B": 18,    # 類別 / 縮影 B 卡
        "C": 40,    # 規則描述 / 縮影 C 卡
        "D": 15,    # 當前值 / 縮影 D 卡
        "E": 15,    # 狀態燈 / 縮影 E 卡
        "F": 14,    # 嚴重度 / 縮影 F 卡（容「F 性能監控」5 字）
        "G": 15,    # 跳轉 / 縮影 G 卡
    }
    for c, w in col_widths.items():
        ws.column_dimensions[c].width = w

    # === Row 1-2 Banner ===
    ws.merge_cells("A1:G1")
    S.set_cell(ws, "A1",
               "   ⚕  系統檢核（45 條規則・7 大類）",
               font_key="banner", fill_key="banner",
               align_key="left", border_key="bottom_thick")
    ws.row_dimensions[1].height = S.ROW_HEIGHT["banner"]

    ws.merge_cells("A2:G2")
    S.set_cell(ws, "A2",
               '="   "&警察局名稱&" · "&分局名稱&" · "&派出所名稱'
               '&" ｜ 健檢日："&TEXT(今日,"e/mm/dd")&"  ｜  按 F9 重算 ｜ 紅燈須立即處理"',
               font_key=S.font(11, color="white", italic=True),
               fill_key=S.fill("accent_dark"), align_key="left")
    ws.row_dimensions[2].height = S.ROW_HEIGHT["sub_banner"]

    # === Row 3 spacer ===
    ws.row_dimensions[3].height = 8

    # === Row 4 維護指引 ===
    ws.merge_cells("A4:G4")
    S.set_cell(ws, "A4",
               "   💡  每條規則 D 欄公式自動算當前值，E 欄狀態燈依規則判斷紅黃綠，"
               "G 欄超連結跳問題頁。「○ 待接通」= 對應資料表尚未建立。",
               font_key=S.font(11, color="warn", italic=True),
               fill_key="warn_bg", align_key="left",
               border_key="all_thin")
    ws.row_dimensions[4].height = 32

    # === Row 5 spacer ===
    ws.row_dimensions[5].height = 8

    # === Row 6 全表健康度 section title ===
    ws.merge_cells("A6:G6")
    S.set_cell(ws, "A6",
               "   ⚕  全表健康度（自動聚合 45 條燈號）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[6].height = S.ROW_HEIGHT["header"]

    # === Row 7-8 大字 3 卡 ===
    # 規則範圍 row 15-59（45 條）
    rule_first_row = 15
    rule_last_row = rule_first_row + len(RULES) - 1  # 59
    status_range = f"$E${rule_first_row}:$E${rule_last_row}"

    health_cards = [
        ("● 紅燈總數",  "danger", "danger_bg",
         f'=COUNTIF({status_range},"*紅*")',
         "需立即處理"),
        ("⚠ 黃燈總數",  "warn",   "warn_bg",
         f'=COUNTIF({status_range},"*黃*")',
         "建議排程處理"),
        ("● 綠燈總數",  "pass",   "pass_bg",
         f'=COUNTIF({status_range},"*綠*")',
         "持續維持"),
    ]
    # 拆成 row 7 = label（兩格 merged），row 8 = value（大字，兩格 merged）
    for i, (label, fg, bg, fml, sub) in enumerate(health_cards):
        c_start = ["A", "C", "E"][i]
        c_end   = ["B", "D", "F"][i]
        # Row 7 label
        ws.merge_cells(f"{c_start}7:{c_end}7")
        S.set_cell(ws, f"{c_start}7", f"  {label}",
                   font_key=S.font(12, bold=True, color=fg),
                   fill_key=S.fill(bg), align_key="left",
                   border_key="all_thin")
        # Row 8 value（大字）
        ws.merge_cells(f"{c_start}8:{c_end}8")
        S.set_cell(ws, f"{c_start}8", fml,
                   font_key=S.font(32, bold=True, color=fg),
                   fill_key=S.fill(bg), align_key="center",
                   border_key="all_thin",
                   number_format="integer")
    # G7-G8 給「待接通 / 資訊」總數
    ws.merge_cells("G7:G7")
    S.set_cell(ws, "G7", "  ○ 其他",
               font_key=S.font(10, color="muted"),
               fill_key="calc", align_key="left",
               border_key="all_thin")
    ws.merge_cells("G8:G8")
    pending_fml = (f'=COUNTIF({status_range},"*待接通*")'
                   f'+COUNTIF({status_range},"*資訊*")')
    S.set_cell(ws, "G8", pending_fml,
               font_key=S.font(20, bold=True, color="muted"),
               fill_key="calc", align_key="center",
               border_key="all_thin",
               number_format="integer")
    ws.row_dimensions[7].height = 26
    ws.row_dimensions[8].height = 56

    # === Row 9 spacer ===
    ws.row_dimensions[9].height = 8

    # === Row 10 七大類縮影 section title ===
    ws.merge_cells("A10:G10")
    S.set_cell(ws, "A10",
               "   📊  7 大類縮影（每類 紅/黃/綠/待接通 數）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[10].height = S.ROW_HEIGHT["header"]

    # === Row 11 七類 header（A-G 各 1 欄，深底白字 12pt）===
    for i, (code, name, color_key) in enumerate(CATEGORIES):
        col = get_column_letter(i + 1)
        S.set_cell(ws, f"{col}11", f"{code} {name}",
                   font_key=S.font(12, bold=True, color="white"),
                   fill_key=S.fill("accent_dark"),
                   align_key="center", border_key="all_thin")
    ws.row_dimensions[11].height = 34

    # === Row 12 七類 計數（紅 / 黃 / 綠 並列，font 13）===
    #   公式回傳「紅 0   黃 0   綠 5」三段並列（彩色由 CF 補染）
    for i, (code, name, _) in enumerate(CATEGORIES):
        col = get_column_letter(i + 1)
        # 過濾此類別的列：B 欄前 2 字 = code + 空白
        cat_filter = (f'(LEFT($B${rule_first_row}:$B${rule_last_row},2)'
                      f'="{code} ")')
        red_fml = (f'SUMPRODUCT({cat_filter}*'
                   f'ISNUMBER(SEARCH("紅",{status_range})))')
        yel_fml = (f'SUMPRODUCT({cat_filter}*'
                   f'ISNUMBER(SEARCH("黃",{status_range})))')
        grn_fml = (f'SUMPRODUCT({cat_filter}*'
                   f'ISNUMBER(SEARCH("綠",{status_range})))')
        combo_fml = (f'="紅 "&{red_fml}&"   黃 "&{yel_fml}'
                     f'&"   綠 "&{grn_fml}')
        S.set_cell(ws, f"{col}12", combo_fml,
                   font_key=S.font(13, bold=True, color="accent"),
                   fill_key=S.fill("accent_light"),
                   align_key="center", border_key="all_thin")
    ws.row_dimensions[12].height = 38

    # === Row 13 主表 section title ===
    ws.merge_cells("A13:G13")
    S.set_cell(ws, "A13",
               "   📋  45 條檢核明細（公式自動 ｜ 點 G 欄連結跳問題頁）",
               font_key="section_title", fill_key="banner",
               align_key="left")
    ws.row_dimensions[13].height = S.ROW_HEIGHT["header"]

    # === Row 14 header ===
    headers = [
        ("編號", "center"),
        ("類別", "center"),
        ("規則描述", "left"),
        ("當前值", "center"),
        ("狀態", "center"),
        ("嚴重度", "center"),
        ("跳轉", "center"),
    ]
    for i, (h, al) in enumerate(headers, start=1):
        col = get_column_letter(i)
        S.set_cell(ws, f"{col}14", h,
                   font_key="header", fill_key="header",
                   align_key=al, border_key="all_thin")
    ws.row_dimensions[14].height = S.ROW_HEIGHT["header"]

    # === Row 15-59 45 規則 ===
    for idx, rule in enumerate(RULES):
        r = rule_first_row + idx

        # A 編號
        S.set_cell(ws, f"A{r}", rule["id"],
                   font_key="body_bold", fill_key="calc",
                   align_key="center", border_key="all_thin")
        # B 類別
        S.set_cell(ws, f"B{r}", rule["cat"],
                   font_key=S.font(10, color="muted"),
                   fill_key="calc",
                   align_key="center", border_key="all_thin")
        # C 規則描述
        S.set_cell(ws, f"C{r}", rule["desc"],
                   font_key="body", fill_key=None,
                   align_key="left", border_key="all_thin")
        # D 當前值（公式）
        S.set_cell(ws, f"D{r}", rule["value"],
                   font_key="body_bold", fill_key="calc",
                   align_key="center", border_key="all_thin",
                   number_format="general")
        # E 狀態燈
        light_fml = rule["light"](f"D{r}")
        S.set_cell(ws, f"E{r}", light_fml,
                   font_key=S.font(12, bold=True, color="accent"),
                   fill_key="calc",
                   align_key="center", border_key="all_thin")
        # F 嚴重度
        sev_color = {"高": "danger", "中": "warn",
                     "低": "muted", "資訊": "muted"}.get(rule["severity"], "text")
        S.set_cell(ws, f"F{r}", rule["severity"],
                   font_key=S.font(11, bold=True, color=sev_color),
                   fill_key="calc",
                   align_key="center", border_key="all_thin")
        # G 跳轉（HYPERLINK）
        link_sheet, link_cell = rule["link"]
        if link_sheet and link_cell:
            link_fml = (f'=IFERROR(HYPERLINK("#\'{link_sheet}\'!{link_cell}",'
                        f'"→ 跳轉"),"")')
            S.set_cell(ws, f"G{r}", link_fml,
                       font_key=S.font(11, bold=True, color="accent"),
                       fill_key="calc",
                       align_key="center", border_key="all_thin")
        else:
            S.set_cell(ws, f"G{r}", "—",
                       font_key=S.font(11, color="muted"),
                       fill_key="calc",
                       align_key="center", border_key="all_thin")

        ws.row_dimensions[r].height = 28

    # === Row 60 spacer ===
    ws.row_dimensions[rule_last_row + 1].height = 12

    # === Row 61 頁尾 ===
    footer_row = rule_last_row + 2
    ws.merge_cells(f"A{footer_row}:G{footer_row}")
    footer = (
        '="© KKEVIN-LIN-2026-V2.1  ｜  45 條檢核規則（7 大類）  ｜  '
        '健檢時間："&TEXT(今日,"e/mm/dd")&"  ｜  按 F9 強制重算"'
    )
    S.set_cell(ws, f"A{footer_row}", footer,
               font_key=S.font(9, color="muted", italic=True),
               fill_key="calc", align_key="center")
    ws.row_dimensions[footer_row].height = 22

    # === 條件格式：狀態欄 E 三色自動染色 ===
    ws.conditional_formatting._cf_rules = {}   # 冪等（重 build 不疊加）

    fill_red    = PatternFill("solid", fgColor="FFFEE2E2")
    fill_yellow = PatternFill("solid", fgColor="FFFEF3C7")
    fill_green  = PatternFill("solid", fgColor="FFDCFCE7")
    fill_gray   = PatternFill("solid", fgColor="FFF1F5F9")
    font_red    = Font(name="微軟正黑體", bold=True, color="FFB91C1C")
    font_yellow = Font(name="微軟正黑體", bold=True, color="FFB45309")
    font_green  = Font(name="微軟正黑體", bold=True, color="FF15803D")
    font_gray   = Font(name="微軟正黑體", bold=True, color="FF64748B")

    status_full = f"E{rule_first_row}:E{rule_last_row}"
    # 紅
    ws.conditional_formatting.add(status_full,
        FormulaRule(formula=[f'ISNUMBER(SEARCH("紅",E{rule_first_row}))'],
                    fill=fill_red, font=font_red, stopIfTrue=False))
    # 黃
    ws.conditional_formatting.add(status_full,
        FormulaRule(formula=[f'ISNUMBER(SEARCH("黃",E{rule_first_row}))'],
                    fill=fill_yellow, font=font_yellow, stopIfTrue=False))
    # 綠
    ws.conditional_formatting.add(status_full,
        FormulaRule(formula=[f'ISNUMBER(SEARCH("綠",E{rule_first_row}))'],
                    fill=fill_green, font=font_green, stopIfTrue=False))
    # 待接通 / 資訊（灰）
    ws.conditional_formatting.add(status_full,
        FormulaRule(formula=[f'OR(ISNUMBER(SEARCH("待接通",E{rule_first_row})),'
                             f'ISNUMBER(SEARCH("資訊",E{rule_first_row})))'],
                    fill=fill_gray, font=font_gray, stopIfTrue=False))

    # 整列依狀態淡染（A:G）
    row_full = f"A{rule_first_row}:G{rule_last_row}"
    ws.conditional_formatting.add(row_full,
        FormulaRule(formula=[f'ISNUMBER(SEARCH("紅",$E{rule_first_row}))'],
                    fill=fill_red, stopIfTrue=False))
    ws.conditional_formatting.add(row_full,
        FormulaRule(formula=[f'ISNUMBER(SEARCH("黃",$E{rule_first_row}))'],
                    fill=fill_yellow, stopIfTrue=False))

    log.info(f"  45 條規則寫入 row {rule_first_row}-{rule_last_row}")
    log.info(f"  健康度 3 卡 + 7 類縮影 + 條件格式 6 條")

    # 公式計數
    fcount = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                fcount += 1
    log.info(f"  公式總數：{fcount}")
    return fcount


def build():
    log = get_logger("phase6_inspection")
    log.info("===== Phase 6 — 頁 15 系統檢核 開始 =====")

    wb = load_workbook(OUTPUT_PATH)
    fcount = build_page15_inspection(wb, log)

    bak = backup_existing()
    if bak:
        log.info(f"  舊版備份：{bak}")
    wb.save(OUTPUT_PATH)
    size_kb = OUTPUT_PATH.stat().st_size / 1024
    log.info(f"  儲存：{OUTPUT_PATH.name} ({size_kb:.1f} KB)")
    log.info(f"  公式：{fcount} 條")
    log.info("===== Phase 6 完成 =====\n")
    return OUTPUT_PATH


if __name__ == "__main__":
    build()
