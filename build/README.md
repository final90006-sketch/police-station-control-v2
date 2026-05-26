# 派出所綜合管制系統 v2 — 工程建檔
> 用 PowerShell + Excel COM 自動產生原生 xlsx，無第三方依賴

## 📁 目錄結構

```
派出所v2_build/
├── README.md              本檔
├── config/                各階段的設定 JSON（UTF-8，含中文）
│   ├── phase1_skeleton.json     17 工作表骨架定義
│   ├── phase2_settings.json     頁 13 設定表內容（待）
│   └── phase3_data.json         頁 5/9.5/14 表頭與計算欄（待）
├── scripts/               PowerShell 腳本（純 ASCII）
│   ├── build_phase1.ps1   建立骨架 xlsx
│   ├── verify.ps1         結構驗證
│   └── test.ps1           高強度測試（公式邊界等）
├── output/                建檔產出
│   └── 派出所綜合管制系統_v2.0.xlsx
├── log/                   每次建檔的時間戳記日誌
└── test/                  測試資料、案例
```

## 🚦 階段路線

| Phase | 內容 | 預估時間 | 狀態 |
|---|---|---|---|
| **0 工程基礎** | 目錄/config/build/verify 框架 | 30 分 | ✓ |
| **1 骨架** | 17 工作表名稱 + 凍結 + TODAY 集中 + 6 命名範圍 | 60 分 | ⏳ |
| **1.5 設定表完整** | 頁 13 A-H 8 大分區 + 所有下拉清單 | 60 分 | ⏸ |
| **2 SSOT 主表** | 頁 5 案件資料庫（22 欄+1000 列+Excel Table）/ 頁 9.5 / 頁 14 | 90 分 | ⏸ |
| **3 衍生看板** | 頁 2/2.5/6/7/8/9/10/11/12 | 120 分 | ⏸ |
| **4 報告頁** | 頁 1 首頁、頁 3 5 分鐘、頁 4 上呈、頁 11 員警卡 | 90 分 | ⏸ |
| **5 細節保護** | 條件格式 + 資料驗證 + 工作表受保護 + 列印區 | 60 分 | ⏸ |
| **6 系統檢核** | 頁 15 45 條規則 | 90 分 | ⏸ |

每階段：**build → verify → test → 通過才進下一階段**

## 🚀 執行方式

```powershell
# 建檔
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\build_phase1.ps1

# 驗證
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\verify.ps1

# 測試
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\test.ps1
```

## ⚠ 中文處理規範

- PowerShell 腳本（.ps1）**純 ASCII**（PS5.1 ANSI 讀 .ps1 會毀中文）
- 中文資料 → UTF-8 JSON sidecar（config/*.json）
- xlsx 含中文 → COM 物件原生支援（無需特殊處理）
- 輸出 xlsx 檔名與路徑可含中文（PowerShell `-LiteralPath` 處理）

## 🛡 容錯設計

每個腳本：
- 用 `$ErrorActionPreference = 'Stop'`
- `try/finally` 確保 Excel COM 釋放
- 失敗時詳細 log 寫到 `log/build_YYYYMMDD_HHMMSS.log`
- 不直接覆寫舊版 xlsx（保留 `.bak`）

---
© 林錦瑞 KKEVIN-LIN-2026-V2.0
