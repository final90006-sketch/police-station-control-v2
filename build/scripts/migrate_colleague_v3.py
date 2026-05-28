"""
migrate_colleague_v3.py — 同仁檔遷移到 v2.2 最新結構

把同仁最終版（舊欄序）的資料搬進全新模板（含即時統計 + 計算欄前移 +
新案件狀況）。用「全新模板 + 依欄名搬資料」確保結構乾淨。

保留：
  - 同仁 92 筆案件（依欄名對應，自動落到新位置）
  - 自訂員警名冊（設定表 B 區）
  - 組織設定
新模板提供：
  - 頂端即時統計（全般/竊盜/詐欺）
  - 是否破獲/計入發生數 前移到案件狀況後
  - 案件狀況新增「已移請偵查隊續辦」（排除發生）
"""
import sys
import os
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

SRC = r"C:\Users\User\Desktop\PoliceStation_v2.2_最終版.xlsx"   # 桌面備份（D 槽已拔）
TEMPLATE = (r"C:\Users\User\Desktop\02【工作】案件管制與自動化系統"
            r"\派出所v2_v43_開發檔案\派出所v2_build\output\PoliceStation_v2.2_部署版.xlsx")
DST = r"C:\Users\User\Desktop\PoliceStation_v2.2_最終版_新結構.xlsx"

# 要搬移的 input 欄（依欄名；calc 欄由模板公式自動算）
INPUT_COLS = ["編號", "案類", "E化案號", "發生時間", "發生地點", "破獲時間",
              "破獲地點", "案情摘要", "偵辦進度", "案件狀況", "發生管轄",
              "查獲管轄", "承辦人", "績效類別", "自填案類"]


def name_to_col(ws, header_row=14, max_col=30):
    d = {}
    for c in range(1, max_col + 1):
        h = ws[f"{get_column_letter(c)}{header_row}"].value
        if h:
            d[h] = get_column_letter(c)
    return d


def main():
    print("===== 同仁檔遷移新結構 開始 =====")
    for p in (SRC, TEMPLATE):
        if not os.path.exists(p):
            print(f"  ✗ 找不到：{p}")
            return

    # 來源
    wb_src = load_workbook(SRC)
    ws_src = wb_src["案件資料庫"]
    src_map = name_to_col(ws_src)
    # 找資料列範圍
    src_tbl = ws_src.tables["Tbl案件"]
    import re
    m = re.match(r"[A-Z]+\d+:[A-Z]+(\d+)", src_tbl.ref)
    src_end = int(m.group(1))

    # 讀同仁 92 筆 input 資料
    rows_data = []
    col_no = src_map["編號"]
    for r in range(15, src_end + 1):
        if ws_src[f"{col_no}{r}"].value in (None, ""):
            continue
        rec = {}
        for name in INPUT_COLS:
            if name in src_map:
                rec[name] = ws_src[f"{src_map[name]}{r}"].value
        rows_data.append(rec)
    print(f"  讀到同仁 {len(rows_data)} 筆案件")

    # 讀員警名冊（設定表 B24:H83）+ 組織
    ws_set_src = wb_src["設定表"]
    officers = []
    for r in range(24, 84):
        row = [ws_set_src[f"{get_column_letter(c)}{r}"].value for c in range(1, 8)]
        if any(v not in (None, "") for v in row):
            officers.append((r, row))
    org = {c: ws_set_src[c].value for c in ("B16", "B17", "B18", "B19")}
    print(f"  讀到員警名冊 {len(officers)} 列；組織 {org['B18']}")

    # === 載入全新模板 ===
    wb = load_workbook(TEMPLATE)
    ws = wb["案件資料庫"]
    dst_map = name_to_col(ws)
    print(f"  模板欄序: 是否破獲={dst_map.get('是否破獲')} 計入發生數={dst_map.get('計入發生數')}")

    # 寫入案件 input 資料（從 row 15）
    for i, rec in enumerate(rows_data):
        r = 15 + i
        for name, val in rec.items():
            if name in dst_map:
                ws[f"{dst_map[name]}{r}"] = val
    print(f"  寫入 {len(rows_data)} 筆到新模板（依欄名落新位置）")

    # 寫員警名冊 + 組織
    ws_set = wb["設定表"]
    for r, row in officers:
        for c, v in enumerate(row, start=1):
            ws_set[f"{get_column_letter(c)}{r}"] = v
    for cell, v in org.items():
        ws_set[cell] = v
    print(f"  還原員警名冊 {len(officers)} 列 + 組織設定")

    wb.calculation.fullCalcOnLoad = True
    wb.save(DST)

    # 驗證
    wb2 = load_workbook(DST)
    ws2 = wb2["案件資料庫"]
    dm = name_to_col(ws2)
    print()
    print("=== 驗證 ===")
    print(f"  即時統計 B8(全般發生公式): {str(ws2['B8'].value)[:40]}")
    print(f"  是否破獲欄位置: {dm.get('是否破獲')}（應在案件狀況 {dm.get('案件狀況')} 之後）")
    no_col = dm['編號']
    v15 = ws2[f"{no_col}15"].value
    v106 = ws2[f"{no_col}106"].value
    print(f"  資料: row15 編號={v15} / row106 編號={v106}")
    print(f"  員警 B24={wb2['設定表']['B24'].value}")
    print(f"  ✓ 完成：{os.path.basename(DST)} ({os.path.getsize(DST)/1024:.1f} KB)")
    print("===== 完成 =====")


if __name__ == "__main__":
    main()
