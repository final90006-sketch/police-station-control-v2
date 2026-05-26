"""
phase11_deploy_clear.py — Phase 11：部署前清資料

把 5 份 sample data 清空（保留結構 / 公式 / 樣式 / 驗證 / CF）：
  - Tbl案件     rows 15-18  (4 筆 — 詐欺/竊盜/毒品/車手)
  - Tbl交通     rows 15-19  (5 筆 取締)
  - Tbl歷史     rows 11-15  (5 月 KPI 快照)
  - Tbl毒調     rows 15-19  (5 筆 列管人口)
  - 頁 7 發展中  rows 14-18  (5 筆 線報)
  - 設定表 B 區員警 rows 24-26 (3 筆 示範員警)

保留：
  - 17 工作表結構
  - 全部公式 / 條件格式 / 工作表保護 / 列印區
  - Excel Table 註冊 / DataValidation 下拉
  - 設定表 A 組織預設值 / C 閾值 / D 案類 / E 違規 / H 下拉清單

策略：只清「input cells」value，calc cells 公式保留。
產生新檔 PoliceStation_v2.2_部署版.xlsx，原 v2.0 / v2.1 保留供開發用。
"""
import sys
import shutil
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from phases._base import OUTPUT_PATH, OUTPUT_DIR, get_logger
from phases import phase9_native_sparklines


# ============================================================
# 清資料規格
# ============================================================
CLEAR_SPECS = [
    # (sheet_name, row_start, row_end, col_start, col_end, description)
    ("案件資料庫",     15, 18, 1, 17,  "Tbl案件 4 sample（input cols A-Q，保留 R-Y calc）"),
    ("案件資料庫",     15, 18, 23, 23, "Tbl案件 W 績效類別 input"),
    ("案件資料庫",     15, 18, 25, 25, "Tbl案件 Y 自填案類 input"),
    ("交通取締明細",   15, 19, 1, 6,   "Tbl交通 5 sample（input 全欄 A-F）"),
    ("歷史資料",       11, 15, 1, 14,  "Tbl歷史 5 月快照（input 全欄 A-N）"),
    ("毒品調驗人口管制", 15, 19, 1, 10, "Tbl毒調 5 sample（input cols A-J 含最後到驗日）"),
    ("發展中案件管制", 14, 18, 1, 9,   "頁 7 主表 5 sample（input cols A-I）"),
    # 設定表 B 區員警名冊（單筆 7 欄）— sample 在 row 24-26
    ("設定表",         24, 26, 1, 7,   "員警名冊 3 sample"),
]


def clear_sample_data(wb, log):
    total_cleared = 0
    for sheet_name, r_start, r_end, c_start, c_end, desc in CLEAR_SPECS:
        if sheet_name not in wb.sheetnames:
            log.warning(f"  ⚠ sheet 不存在：{sheet_name}")
            continue
        ws = wb[sheet_name]
        count = 0
        for r in range(r_start, r_end + 1):
            for c in range(c_start, c_end + 1):
                col_letter = get_column_letter(c)
                cell = ws[f"{col_letter}{r}"]
                # 只清 input：value 是字串 / 數字 / datetime 但不是公式
                if cell.value is None:
                    continue
                v = cell.value
                if isinstance(v, str) and v.startswith("="):
                    continue   # calc 公式保留
                # 清值（保留樣式 / 填色 / 邊框）
                cell.value = None
                count += 1
        total_cleared += count
        log.info(f"  {sheet_name} row {r_start}-{r_end} col {get_column_letter(c_start)}-{get_column_letter(c_end)}: 清 {count} cells ({desc})")
    log.info(f"  總清空 cells：{total_cleared}")
    return total_cleared


def build():
    log = get_logger("phase11_deploy")
    log.info("===== Phase 11 — 部署前清資料 開始 =====")

    # 開 v2.0 → 清資料 → 另存 v2.2_部署版
    deploy_path = OUTPUT_DIR / "PoliceStation_v2.2_部署版.xlsx"
    log.info(f"  source: {OUTPUT_PATH.name}")
    log.info(f"  target: {deploy_path.name}")

    wb = load_workbook(OUTPUT_PATH)
    cleared = clear_sample_data(wb, log)

    # 加部署版識別到 設定表 A1（避免被誤用為 dev 檔）
    if "設定表" in wb.sheetnames:
        ws_set = wb["設定表"]
        a1_val = ws_set["A1"].value
        if a1_val and "部署版" not in str(a1_val):
            # 不改 A1（會破壞 layout），改加 build_info row 2 末段
            pass

    wb.save(deploy_path)
    log.info(f"  暫存：{deploy_path.name}（openpyxl 載入時 sparkline 已被剝除，需重注入）")

    # ★ 重要：openpyxl 載入會剝除 native sparkline → 對部署版重注入
    # 用 phase 9 邏輯：解 zip → 注入 extLst → 重 zip
    import zipfile, os, shutil
    tmp_dir = OUTPUT_DIR / "_deploy_sparkline_tmp"
    tmp_xlsx = OUTPUT_DIR / "_deploy_sparkline_inject.xlsx"
    if tmp_dir.exists():
        shutil.rmtree(tmp_dir)

    tmp_dir.mkdir(parents=True)
    with zipfile.ZipFile(deploy_path) as z:
        z.extractall(tmp_dir)

    sheet2_path = tmp_dir / "xl" / "worksheets" / "sheet2.xml"
    ext_xml = phase9_native_sparklines.build_sparkline_ext(
        "管制總覽", phase9_native_sparklines.PAGE2_SPARKLINES)
    with open(sheet2_path, "r", encoding="utf-8") as f:
        content = f.read()
    new_content = phase9_native_sparklines.inject_into_sheet(content, ext_xml)
    with open(sheet2_path, "w", encoding="utf-8") as f:
        f.write(new_content)

    with zipfile.ZipFile(tmp_xlsx, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(tmp_dir):
            for f in files:
                fp = Path(root) / f
                arcname = fp.relative_to(tmp_dir)
                z.write(fp, arcname)
    shutil.move(str(tmp_xlsx), str(deploy_path))
    shutil.rmtree(tmp_dir)
    log.info(f"  ✓ 重注入 7 條 native sparkline 到 {deploy_path.name}")

    size_kb = deploy_path.stat().st_size / 1024
    log.info(f"  ✓ 部署版產出：{deploy_path.name} ({size_kb:.1f} KB)")
    log.info(f"  共清 {cleared} 個 sample cells；公式/樣式/CF/保護/列印區/Tbl/驗證/sparkline 全保留")
    log.info("===== Phase 11 完成 =====\n")
    return deploy_path


if __name__ == "__main__":
    build()
