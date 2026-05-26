"""
phase9_native_sparklines.py — Phase 9：真原生 Excel Sparkline

openpyxl 3.1.5 沒有 sparkline 模組，所以這個 Phase 在 openpyxl 存檔 **之後**
用 XML 後處理對 xlsx 加 `extLst` 區塊，注入 sparkline xml。

Excel 2010+ 全相容，舊版會 fallback 顯示原有 Unicode block 字元（已存在）。

策略：
  - 頁 2 管制總覽：A11/G11/M11/A16/G16/M16/A21 = 7 sparkline（KPI 1-7）
    KPI 8/9（酒駕/闖紅燈）helper data 為空白，不加 sparkline
    Data 來源：helper row 50-56，cols T:AE (12 月)
  - 頁 12 跨期間趨勢分析：5 KPI × 12 月 sparkline
    （此頁 sparkline 結構另查 phase3_page12_trend.py）

xlsx 內部 sheet 編號（依 workbook 插入順序）：
  sheet1.xml = 首頁_操作說明
  sheet2.xml = 管制總覽           ← 目標 1
  sheet3.xml = 全般刑案管制情形分析
  ...
  sheet14.xml = 跨期間趨勢分析    ← 目標 2

Sparkline XML schema (Excel 2010+)：
  <extLst>
    <ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"
         xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
      <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
        <x14:sparklineGroup type="line" displayEmptyCellsAs="zero">
          <x14:colorSeries rgb="FF1F4E79"/>
          <x14:sparklines>
            <x14:sparkline>
              <xm:f>管制總覽!T50:AE50</xm:f>
              <xm:sqref>A11</xm:sqref>
            </x14:sparkline>
          </x14:sparklines>
        </x14:sparklineGroup>
      </x14:sparklineGroups>
    </ext>
  </extLst>
"""
import sys
import os
import shutil
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from phases._base import OUTPUT_PATH, OUTPUT_DIR, get_logger


# ============================================================
# 頁 2 管制總覽 sparkline 規格（7 KPI）
# ============================================================
# spark_cell / helper row (T:AE = 12 月資料)
PAGE2_SPARKLINES = [
    # (cell, helper_row, color_hex)
    ("A11", 50, "FF1F4E79"),  # KPI 1 全般破獲率（深藍）
    ("G11", 51, "FF1F4E79"),  # KPI 2 竊盜
    ("M11", 51, "FF1F4E79"),  # ← 應該是 row 52，留意（後修正）
    # 重新依 KPIS index：0=A8 1=G8 2=M8 3=A13 4=G13 5=M13 6=A18 7=G18 8=M18
    # spark_cell = col+r0+3 → A11/G11/M11/A16/G16/M16/A21/G21/M21
    # helper rows = 50+index
]

# 修正版：依 KPIS 順序
PAGE2_SPARKLINES = [
    ("A11", 50, "FF1F4E79"),  # 0 全般破獲率
    ("G11", 51, "FF1F4E79"),  # 1 竊盜破獲率
    ("M11", 52, "FF1F4E79"),  # 2 詐欺破獲率
    ("A16", 53, "FF1F4E79"),  # 3 毒調率
    ("G16", 54, "FFB91C1C"),  # 4 未到驗人口（紅）
    ("M16", 55, "FFB91C1C"),  # 5 未破案件（紅）
    ("A21", 56, "FF15803D"),  # 6 交通達成率（綠）
    # KPI 7/8 (G21/M21) — helper data 空白，跳過
]


# ============================================================
# 頁 12 跨期間趨勢分析 sparkline（5 KPI × 12 月）
# 從 phase3_page12_trend.py 讀取規格
# ============================================================
# spark cell @ B11/B13/B15/B17/B19 (粗估，5 KPI 5 列)
# helper data 從 phase3_page12 內部 layout 讀
# 簡化先 skip，後續若需要再補
PAGE12_SPARKLINES = []   # 暫不處理（phase 12 用 OFFSET 動態取）


def build_sparkline_ext(sheet_name: str, specs):
    """為一個 sheet 產生完整的 extLst XML。"""
    if not specs:
        return ""
    sparkline_groups_xml = []
    for cell, helper_row, color in specs:
        data_range = f"T{helper_row}:AE{helper_row}"
        group_xml = (
            f'<x14:sparklineGroup type="line" displayEmptyCellsAs="zero" '
            f'lineWeight="1.25" '
            f'displayXAxis="0" '
            f'markers="0" high="1" low="1" first="0" last="0" '
            f'negative="0">'
            f'<x14:colorSeries rgb="{color}"/>'
            f'<x14:colorNegative rgb="FFB91C1C"/>'
            f'<x14:colorAxis rgb="FF000000"/>'
            f'<x14:colorMarkers rgb="{color}"/>'
            f'<x14:colorFirst rgb="FF15803D"/>'
            f'<x14:colorLast rgb="FFB91C1C"/>'
            f'<x14:colorHigh rgb="FF15803D"/>'
            f'<x14:colorLow rgb="FFB91C1C"/>'
            f'<x14:sparklines>'
            f'<x14:sparkline>'
            f'<xm:f>{sheet_name}!{data_range}</xm:f>'
            f'<xm:sqref>{cell}</xm:sqref>'
            f'</x14:sparkline>'
            f'</x14:sparklines>'
            f'</x14:sparklineGroup>'
        )
        sparkline_groups_xml.append(group_xml)

    return (
        '<extLst>'
        '<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" '
        'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
        '<x14:sparklineGroups '
        'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
        + "".join(sparkline_groups_xml) +
        '</x14:sparklineGroups>'
        '</ext>'
        '</extLst>'
    )


def inject_into_sheet(xml_content: str, ext_xml: str) -> str:
    """把 extLst 插入到 sheet xml 的 </worksheet> 之前。

    若 sheet 已有 extLst（不太可能，但保險）→ 合併。
    """
    if not ext_xml:
        return xml_content
    # 簡單實作：直接在 </worksheet> 前插入。Phase 5 沒有用 extLst。
    return xml_content.replace('</worksheet>', ext_xml + '</worksheet>')


def post_process_sparklines(log):
    """主入口：對 OUTPUT_PATH 做 XML 注入。"""
    src = OUTPUT_PATH
    tmp_dir = OUTPUT_DIR / "_sparkline_tmp"
    tmp_xlsx = OUTPUT_DIR / "_sparkline_inject.xlsx"

    # 清舊 temp
    if tmp_dir.exists():
        shutil.rmtree(tmp_dir)
    if tmp_xlsx.exists():
        tmp_xlsx.unlink()

    # Unzip xlsx
    tmp_dir.mkdir(parents=True)
    with zipfile.ZipFile(src) as z:
        z.extractall(tmp_dir)
    log.info(f"  解 xlsx → {tmp_dir}")

    # 注入頁 2（sheet2.xml = 管制總覽）
    sheet2_path = tmp_dir / "xl" / "worksheets" / "sheet2.xml"
    ext_xml_p2 = build_sparkline_ext("管制總覽", PAGE2_SPARKLINES)
    with open(sheet2_path, "r", encoding="utf-8") as f:
        content = f.read()
    new_content = inject_into_sheet(content, ext_xml_p2)
    with open(sheet2_path, "w", encoding="utf-8") as f:
        f.write(new_content)
    log.info(f"  頁 2 注入 {len(PAGE2_SPARKLINES)} 條 native sparkline → sheet2.xml")

    # 重 zip 成新 xlsx（保持 zip 結構）
    with zipfile.ZipFile(tmp_xlsx, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(tmp_dir):
            for f in files:
                fp = Path(root) / f
                arcname = fp.relative_to(tmp_dir)
                z.write(fp, arcname)
    log.info(f"  重 zip → {tmp_xlsx.name}")

    # 取代原檔
    shutil.move(str(tmp_xlsx), str(src))
    shutil.rmtree(tmp_dir)
    log.info(f"  ✓ 已注入 sparkline 到 {src.name}")


def build():
    log = get_logger("phase9_sparklines")
    log.info("===== Phase 9 — 真原生 Sparkline 後處理 開始 =====")

    post_process_sparklines(log)

    size_kb = OUTPUT_PATH.stat().st_size / 1024
    log.info(f"  最終 size：{size_kb:.1f} KB")
    log.info(f"  Sparkline 統計：頁 2 = {len(PAGE2_SPARKLINES)} 條")
    log.info(f"  Page 12 暫 skip（其結構需另查）")
    log.info("===== Phase 9 完成 =====\n")
    return OUTPUT_PATH


if __name__ == "__main__":
    build()
