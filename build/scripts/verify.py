"""
verify.py — 驗證建出的 xlsx 結構
用法:  python verify.py
"""
import json
import sys
from datetime import datetime, date
from pathlib import Path

from openpyxl import load_workbook


SCRIPT_DIR = Path(__file__).resolve().parent
BUILD_DIR = SCRIPT_DIR.parent
CONFIG_DIR = BUILD_DIR / "config"
OUTPUT_DIR = BUILD_DIR / "output"

OUTPUT_PATH = OUTPUT_DIR / "PoliceStation_v2.0.xlsx"


def main():
    if not OUTPUT_PATH.exists():
        print(f"[ERROR] File not found: {OUTPUT_PATH}")
        sys.exit(1)

    size_kb = OUTPUT_PATH.stat().st_size / 1024
    print(f"File: {OUTPUT_PATH.name} ({size_kb:.1f} KB)")
    print()

    cfg = json.loads((CONFIG_DIR / "phase1_skeleton.json").read_text(encoding="utf-8"))

    # Load with data_only=False to see formulas; True to evaluate
    wb_formulas = load_workbook(OUTPUT_PATH, data_only=False)
    wb_values = load_workbook(OUTPUT_PATH, data_only=True)

    pass_count = 0
    fail_count = 0

    def check(name: str, cond: bool, detail: str = ""):
        nonlocal pass_count, fail_count
        if cond:
            print(f"  [PASS] {name}" + (f" :: {detail}" if detail else ""))
            pass_count += 1
        else:
            print(f"  [FAIL] {name}" + (f" :: {detail}" if detail else ""))
            fail_count += 1

    # === Phase 1 checks ===
    print("--- Phase 1: 骨架驗證 ---")
    expected_sheets = sorted(cfg["sheets"], key=lambda s: s["order"])

    # 1. 工作表數量
    check("工作表數量 = 17", len(wb_formulas.sheetnames) == 17,
          f"actual={len(wb_formulas.sheetnames)}")

    # 2. 工作表順序與名稱
    for idx, exp in enumerate(expected_sheets):
        if idx < len(wb_formulas.sheetnames):
            actual = wb_formulas.sheetnames[idx]
            check(f"Sheet[{idx+1}] = {exp['name']}", actual == exp["name"],
                  f"actual={actual}")

    # 3. TODAY 集中
    today_cfg = cfg["today_centralization"]
    settings_ws = wb_formulas[today_cfg["sheet"]]
    formula = settings_ws[today_cfg["formula_cell"]].value
    check(f"TODAY 公式存在", formula == today_cfg["formula"],
          f"got={formula!r}")

    settings_ws_v = wb_values[today_cfg["sheet"]]
    today_val = settings_ws_v[today_cfg["formula_cell"]].value
    # openpyxl 不會自己 evaluate TODAY；要等 Excel 開過再存
    if today_val is None:
        print(f"  [INFO] TODAY 值尚未計算（需 Excel 打開過再存）")
    else:
        is_today = isinstance(today_val, (date, datetime)) and (
            today_val == datetime.now().date() or
            (isinstance(today_val, datetime) and today_val.date() == datetime.now().date())
        )
        check("TODAY 值 = 今天", is_today, f"got={today_val}")

    # 4. 命名範圍
    for nr in cfg["named_ranges_phase1"]:
        exists = nr["name"] in wb_formulas.defined_names
        check(f"命名範圍 '{nr['name']}' 存在", exists)
        if exists:
            actual_ref = wb_formulas.defined_names[nr["name"]].attr_text
            expected_ref = nr["refers_to"].lstrip("=")
            check(f"命名範圍 '{nr['name']}' 指向", actual_ref == expected_ref,
                  f"got={actual_ref!r}")

    # 5. 凍結窗格抽樣（首頁、案件資料庫、設定表）
    for exp in [expected_sheets[0], expected_sheets[5], expected_sheets[14]]:
        ws = wb_formulas[exp["name"]]
        freeze = ws.freeze_panes
        check(f"凍結窗格 {exp['name']} = {exp['freeze']}",
              freeze == exp["freeze"], f"got={freeze}")

    # 6. A1 marker 存在
    for exp in expected_sheets:
        ws = wb_formulas[exp["name"]]
        val = ws["A1"].value
        check(f"A1 marker on {exp['name']}", val is not None and len(str(val)) > 0)

    # 7. Phase 1.5 額外檢查（若 config 存在）
    p15_cfg_path = CONFIG_DIR / "phase1_5_settings.json"
    if p15_cfg_path.exists():
        print()
        print("--- Phase 1.5: 設定表 8 大分區驗證 ---")
        p15 = json.loads(p15_cfg_path.read_text(encoding="utf-8"))
        ws = wb_formulas["設定表"]

        # 動態搜尋區塊位置（掃 A 欄）
        def find_section_row(text_fragment: str) -> int:
            for r in range(1, ws.max_row + 1):
                val = ws.cell(row=r, column=1).value
                if val and text_fragment in str(val):
                    return r
            return -1

        # 頂部 banner
        a1 = ws["A1"].value
        check("頁面頂部 banner 存在", a1 and "v2" in str(a1), f"got={a1!r}")

        # 各區搜尋
        for code in ["A 區", "B 區", "C 區", "D 區", "E 區", "F 區", "H 區"]:
            r = find_section_row(code)
            check(f"{code} banner 找到", r > 0, f"row={r}")

        # 員警示範資料（B 區 banner 後第 3 列）
        b_row = find_section_row("B 區")
        if b_row > 0:
            b_data_start = b_row + 2
            sample_name = ws.cell(row=b_data_start, column=2).value
            check("B 區示範員警有資料",
                  sample_name is not None and len(str(sample_name)) > 0,
                  f"row={b_data_start} col=B got={sample_name!r}")

        # H 區命名範圍
        for ld in p15["section_h"]["dropdown_lists"]:
            exists = ld["named_range"] in wb_formulas.defined_names
            check(f"命名範圍 '{ld['named_range']}' (H 區下拉)", exists)

        # A 區 5 命名範圍
        for fld in p15["section_a"]["fields"]:
            if fld.get("named_range"):
                exists = fld["named_range"] in wb_formulas.defined_names
                check(f"命名範圍 '{fld['named_range']}' (A 區)", exists)

        # C 區 7 命名範圍
        for thr in p15["section_c"]["thresholds"]:
            if thr.get("named_range"):
                exists = thr["named_range"] in wb_formulas.defined_names
                check(f"命名範圍 '{thr['named_range']}' (C 區)", exists)

        # F 區 2 命名範圍
        for item in p15["section_f"]["items"]:
            if item.get("count_named_range"):
                exists = item["count_named_range"] in wb_formulas.defined_names
                check(f"命名範圍 '{item['count_named_range']}' (F 區)", exists)

        # 總命名範圍數
        total_names = len(wb_formulas.defined_names)
        check(f"總命名範圍數 >= 22", total_names >= 22, f"got={total_names}")

    print()
    total = pass_count + fail_count
    print(f"===== Verify ended | Pass: {pass_count}/{total} | Fail: {fail_count} =====")
    if fail_count > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
