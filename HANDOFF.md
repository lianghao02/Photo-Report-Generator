# HANDOFF

## 目前狀態
可交付

## 本輪目標
升級大圖燈箱 (Image Lightbox) 至鑑識級規格（對齊 PaperSwitch 視圖標準）：
1. 雙層舞台解耦（杜絕旋轉後滑鼠平移軸向反轉）。
2. 游標錨點平滑縮放（0.35x ～ 4.0x，步進 0.15）。
3. 滑鼠抓手平移拖曳（放大狀態下按住左鍵自由平移）。
4. 順逆時針旋轉（`R` / `Shift+R`）並即時同步畫布卡片與表格縮圖。
5. 延遲歷史記錄提交（Lazy Commit，防止連續旋轉污染 Undo/Redo 佇列）。
6. 工具列控制項（縮放百分比徽章、適合視窗、順逆旋轉、關閉）與快捷鍵（`0`、`Space`、`Enter`、`+`/`-`）。
7. 平滑降級（優先顯示縮圖以保證即時反饋，背景載入原圖後無縫替換）。

## 已完成
1. **雙層容器隔離座標系**：
   - 燈箱架構重構為 `#lightboxViewport`（視口與抓手狀態） > `#lightboxStage`（負責 Pan 平移與 Zoom 縮放） > `#lightboxImg`（負責 Rotation 旋轉）。
   - 徹底隔離旋轉與平移矩陣，解決了旋轉 90° 或 270° 後滑鼠拖曳方向反轉的常見幾何問題。
2. **游標錨點平滑縮放與平移抓手**：
   - 實作 `handleLightboxWheel` 與 `zoomLightboxBy`，根據滑鼠目前游標在視口中的相對位置計算縮放平移偏移量，縮放範圍限制於 0.35x 至 4.0x，重設（<= 1.0x）時自動置中歸零。
   - 實作 `handleLightboxPointerDown`、`handleLightboxPointerMove`、`handleLightboxPointerUp`，支援放大狀態下按住左鍵平移，並具有最大平移範圍邊界限制。
   - Viewport 支援 `.is-zoomed`（grab 抓手）與 `.is-panning`（grabbing 抓取中）視覺狀態。
3. **即時旋轉與主畫布雙向聯動**：
   - 燈箱內點擊「順旋」、「逆旋」或按下 `R` / `Shift + R` 快捷鍵，即時旋轉照片資料 `photo.rotation`。
   - `syncPhotoRotationToCanvas` 即時更新主畫布卡片（`.photo-thumb-card .thumbnail-stage img`）與試算表表格列（DataGrid）之旋轉樣式。
4. **旋轉歷史延遲提交 (Lazy Commit)**：
   - 燈箱內連續旋轉時僅更新視覺與資料，標記 `this.lightboxHasRotated = true`。
   - 僅在「翻頁 (`navigateLightbox`)」或「關閉燈箱 (`closeModal`)」時，才統一呼叫 `commitLightboxHistoryIfNeeded` 寫入一筆 `燈箱旋轉照片` 歷史記錄，避免 Undo 歷史遭連續旋轉佔滿。
5. **快捷鍵與控制列全面增強**：
   - 頂部工具列加入：即時縮放百分比徽章（如 `100%`、`180%`）、`[適合視窗]`（快速鍵 `0`）、`[↺ 逆旋]`（`Shift+R`）、`[↻ 順旋]`（`R`）、`[關閉 (ESC)]`。
   - 主畫布支援選取照片後按 `Space` 或 `Enter` 直接開啟大圖燈箱（遇到輸入框、文字區塊與按鈕時自動跳過）。
   - 保留原有的「雙擊縮圖卡片 = 展開單張編輯並自動聚焦說明欄位」。
6. **平滑降級載入**：
   - 開啟燈箱或翻頁時，優先直接指派 `photo.thumbnailSrc` 即時呈現預覽，避免任何空白等待；同時以非同步方式載入完整原圖，載入完成後無縫替換。
7. **快捷鍵說明彈窗同步更新**：
   - 在 `#shortcutsModal` 擴充「縮放與大圖燈箱」分類說明（涵蓋 Space、滾輪/+-、拖曳平移、R/Shift+R、0 適合視窗）。
8. **自動化測試擴充與全數通過**：
   - `tests/e2e/photo-report.spec.js` 擴充 `[7/5]` 驗證點擊放大按鈕、縮放徽章百分比、鍵盤放大、抓手平移、快速鍵 0 還原、順逆旋轉、延遲歷史提交、主畫布 Space 鍵開啟燈箱，以及雙擊展開編輯器。
   - 執行 `npm test`（4/4 通過）、`npm run test:e2e`（全部通過）、`npm run test:baseline`（全部通過）、`scripts/qa.ps1`（通過）。
   - 產出高畫質展示截圖（100% 預設、放大 180% 平移、旋轉 90 度、主畫布旋轉同步、快捷鍵指南）。

## 刻意未修改
- 未改動 Word、PDF、Excel、ZIP 匯出結構（Baseline 測試 100% 通過）。
- 未改動主畫布卡片拖曳排序與框選邏輯。

## 尚未完成
無

## 驗證結果
### 已執行
- `npm test`：全部通過 (4/4)
- `npm run test:e2e`：全部通過 (包含燈箱放大、平移、旋轉、還原、Space 開啟、雙擊展開編輯)
- `npm run test:baseline`：全部通過 (Word, Excel, PDF 結構符合 Golden Baseline)
- `pwsh -NoProfile -File ./scripts/qa.ps1`：通過
- Playwright 實測截圖驗證：
  - `09_lightbox_100_default.png`（鑑識級大圖燈箱 100% 預設視圖）
  - `10_lightbox_zoomed_panned.png`（放大 180% 與抓手平移狀態）
  - `11_lightbox_rotated_90.png`（燈箱內旋轉 90 度）
  - `12_canvas_synced_rotation.png`（關閉燈箱後主畫布縮圖同步旋轉）
  - `13_shortcuts_guide_lightbox.png`（快捷鍵操作指南面板）

### 尚未驗證
無阻斷性未驗證項目。

### 已知風險
無阻斷性風險。

## Git 狀態
- Commit：未提交（待提交）
- Push：否
- Working Tree：Modified
- Branch: main

## 下一步
- 提交 Git Commit。
