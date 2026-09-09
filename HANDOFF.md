# 當前交接狀態 (Current Handoff)

- **本輪目標**：重整左側「案件資料、專案管理、清冊版型與匯出」資訊架構，不變更既有輸出與專案資料格式。
- **已完成**：
  1. 「案件資料、清冊預設與版型」改為「案件資料與清冊預設」，並移出 `layoutSelect`。
  2. 新增精簡「專案」列，保留既有 `btnSaveProject`、`importProjectInput` 與行為。
  3. 「準備產出正式現場照片清冊」改為「版型與匯出」；`layoutSelect` 改標示「清冊版型」。
  4. 新增三種版型的即時說明；Word/PDF 共用原有版型值與匯出邏輯。
  5. Word 保持主要操作；Excel 匯入／匯出歸入「資料交換」，ZIP 歸入「其他輸出」。
  6. 擴充 E2E：驗證左側分群、既有 ID、版型提示、專案開啟／儲存、Excel 匯入與 ZIP 匯出。
  7. 擴充匯出基準測試：三種版型皆實際產生有效 Word 與 PDF。
- **刻意未修改（保留範圍）**：Word/PDF/Excel/ZIP exporter 核心、三種正式版型尺寸、專案檔格式、Undo/Redo、拖曳排序、照片載入與 EXIF 均未修改。
- **驗證結果與測試證據**：
  - `npm test`：4/4 單元測試通過。
  - `npm run test:e2e`：左側 UX、完整度提醒、專案、Excel、ZIP 流程通過。
  - `npm run test:baseline`：既有 Golden Baseline 與三種 Word/PDF 版型通過。
  - `scripts/qa.ps1`、`git diff --check`：通過。
  - `scripts/build-portable.ps1`：通過；本輪產物為 `照片清冊產生器_2.2.1_x64-setup.exe` 與 `照片清冊產生器_2.2.1_x64_portable.zip`（2026-09-09 12:48）。
- **已知事項與注意事項**：
  - `Browserslist` 顯示 caniuse-lite 更新提示，未影響建置或測試；本輪不擴大處理相依更新。
  - 已完成自動化網頁與建置驗證；Tauri 桌面視窗的人工肉眼手感驗收仍可在實機進行。
- **下一步建議**：若要發布下一正式版本，另行決定版本號、Release 與 GitHub Pages 部署；本輪未修改版本號。
- **目前狀態判定**：可交付
