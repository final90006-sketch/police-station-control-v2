"""
build.py — 派出所綜合管制系統 v2 主建檔 CLI

用法:
  python build.py --phase 1
  python build.py --phase 1.5
  python build.py --phase all

實際工作交給 phases/ 子模組。本檔只負責 dispatch。
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from phases import (phase1_skeleton, phase1_5_settings, phase2_data_tables,
                     phase3_pages_1_7, phase3_page9_traffic,
                     phase3_page2_overview, phase3_page10_performance,
                     phase3_page8_drug, phase3_page11_officer,
                     phase3_page6_case_control, phase3_page2_5_crime_analysis,
                     phase3_page12_trend, phase4_page3_brief,
                     phase4_page4_upreport, phase5_polish,
                     phase6_inspection, phase8_history_writeback,
                     phase9_native_sparklines)


PHASES = {
    "1":    ("Phase 1：17 工作表骨架", phase1_skeleton.build),
    "1.5":  ("Phase 1.5：設定表 8 大分區（精緻版）", phase1_5_settings.build),
    "2":    ("Phase 2：三大資料表(案件/取締/歷史)", phase2_data_tables.build),
    "3":    ("Phase 3：頁 1 首頁 + 頁 7 發展中", phase3_pages_1_7.build),
    "3.2":  ("Phase 3 — 頁 2 管制總覽（九宮格 + sparkline）", phase3_page2_overview.build),
    "3.25": ("Phase 3 — 頁 2.5 全般刑案分析（章戳 + 雙環形圖）", phase3_page2_5_crime_analysis.build),
    "3.6":  ("Phase 3 — 頁 6 刑案管制（3 視圖下拉 + 流程圖）", phase3_page6_case_control.build),
    "3.8":  ("Phase 3 — 頁 8 毒品調驗（Tbl毒調 + 雙翼防呆）", phase3_page8_drug.build),
    "3.9":  ("Phase 3 — 頁 9 交通績效看板", phase3_page9_traffic.build),
    "3.10": ("Phase 3 — 頁 10 績效統計（案類 + 員警 Top 10）", phase3_page10_performance.build),
    "3.11": ("Phase 3 — 頁 11 員警個人績效卡（下拉切員警）", phase3_page11_officer.build),
    "3.12": ("Phase 3 — 頁 12 跨期間趨勢（5 KPI sparkline 矩陣）", phase3_page12_trend.build),
    "4.3":  ("Phase 4 — 頁 3 長官 5 分鐘報告（Pyramid Principle）", phase4_page3_brief.build),
    "4.4":  ("Phase 4 — 頁 4 對上級機關上呈（CHOOSE 切三格式）", phase4_page4_upreport.build),
    # ★ Phase 8（頁 14 即時 KPI 抄寫區）放在 Phase 6 之前
    #   讓 Phase 6 D6 規則能跨入 row 6 文字（其實 D6 用統計年月 COUNTIF 與 row 6 無關，但保持結構整潔）
    "8":    ("Phase 8 — 頁 14 即時 KPI 抄寫區", phase8_history_writeback.build),
    # ★ Phase 6（頁 15 系統檢核）放在 Phase 5（保護）之前，
    #   讓 build all 先填內容再上鎖
    "6":    ("Phase 6 — 頁 15 系統檢核 45 條", phase6_inspection.build),
    "5":    ("Phase 5 — 條件格式 + 工作表保護 + 列印區", phase5_polish.build),
    # ★ Phase 9（真原生 Sparkline）必須是最後一個 — XML 後處理已 openpyxl 寫完的 xlsx
    "9":    ("Phase 9 — 真原生 Excel Sparkline（XML 注入）", phase9_native_sparklines.build),
}


def main():
    ap = argparse.ArgumentParser(
        description="派出所綜合管制系統 v2 主建檔",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="\n".join([f"  {k:<6} {v[0]}" for k, v in PHASES.items()] +
                          ["  all     依序執行所有已實作 phase"]),
    )
    ap.add_argument("--phase", required=True,
                     choices=list(PHASES.keys()) + ["all"])
    args = ap.parse_args()

    if args.phase == "all":
        for code, (desc, fn) in PHASES.items():
            print(f"\n>>> {code} :: {desc}")
            fn()
    else:
        desc, fn = PHASES[args.phase]
        print(f">>> {args.phase} :: {desc}")
        fn()


if __name__ == "__main__":
    main()
