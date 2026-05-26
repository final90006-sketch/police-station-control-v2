"""
_base.py — 各 phase 共用工具：路徑、logger、config loader、backup
"""
import json
import logging
import sys
from datetime import datetime
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent.parent  # scripts/
BUILD_DIR = SCRIPT_DIR.parent
CONFIG_DIR = BUILD_DIR / "config"
OUTPUT_DIR = BUILD_DIR / "output"
LOG_DIR = BUILD_DIR / "log"

OUTPUT_NAME = "PoliceStation_v2.0.xlsx"
OUTPUT_PATH = OUTPUT_DIR / OUTPUT_NAME

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)


def get_logger(phase_name: str):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = LOG_DIR / f"build_{phase_name}_{ts}.log"

    logger = logging.getLogger(f"build.{phase_name}")
    logger.setLevel(logging.INFO)
    # 避免重複 handler
    logger.handlers.clear()

    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(fh)

    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
    logger.addHandler(ch)

    logger.info(f"Log file: {log_path.name}")
    return logger


def load_config(name: str) -> dict:
    path = CONFIG_DIR / f"{name}.json"
    if not path.exists():
        raise FileNotFoundError(f"Config not found: {path}")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def unmerge_all_in_sheet(ws):
    """清掉該 sheet 上所有 merged ranges。

    使用情境：phase 3 頁面 builder 想用比 Phase 1 banner 更寬的範圍 merge
    （Phase 1 預設 banner = A1:G1，A2:G2），不先 unmerge 會與新 merge 重疊
    導致 Excel 開檔報「已移除的記錄: ... 部分的 合併儲存格」並破壞欄寬。

    呼叫此函數會清掉 ALL merges（含 banner、各區塊），caller 必須自己重建。
    """
    # list() 複製避免 iteration 時修改集合
    existing = [str(r) for r in ws.merged_cells.ranges]
    for r in existing:
        ws.unmerge_cells(r)
    return len(existing)


def backup_existing():
    if OUTPUT_PATH.exists():
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        bak = OUTPUT_PATH.with_name(f"PoliceStation_v2.0_bak_{ts}.xlsx")
        # 若同秒內已建過備份，加流水號
        if bak.exists():
            n = 2
            while True:
                bak = OUTPUT_PATH.with_name(f"PoliceStation_v2.0_bak_{ts}_{n}.xlsx")
                if not bak.exists():
                    break
                n += 1
        OUTPUT_PATH.rename(bak)
        return bak.name
    return None
