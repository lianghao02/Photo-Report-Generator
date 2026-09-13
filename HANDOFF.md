# HANDOFF

## 目前狀態
可交付

## 本輪目標
修復縮圖卡片「點擊放大圖示開啟燈箱」與「雙擊縮圖卡片自動展開單張編輯並聚焦說明欄位」兩項核心操作互動。

## 已完成
1. **點擊放大圖示開啟燈箱大圖**：
   - 診斷出 PaperSwitch 指標拖曳在 `beginCardPointerDrag` 時過早執行 `card.setPointerCapture`，導致點擊事件直接派發給父容器 `card`，而忽略子層 `.thumbnail-stage` 點擊的問題。
   - 將 `setPointerCapture` 移至 `updateCardPointerDrag`（僅在滑鼠移動距離超過 6px 門檻時才觸發 capture），徹底還原原生點擊與雙擊事件傳遞機制。
   - 將縮圖卡片上的「放大圖示」獨立為右下角優雅的 `.btn-zoom-preview` 按鈕，設定明確的 `position: absolute; bottom: 6px; right: 6px; pointer-events: auto !important;`，排除預編譯 Tailwind 樣式失效隱患。
   - 點擊按鈕時透過 `event.stopPropagation()` 與 `window.app.openLightboxByIndex(idx)` 順暢開啟 `#imageLightboxModal` 燈箱，按 ESC 鍵或關閉按鈕即可退出。
2. **雙擊縮圖自動展開單張編輯**：
   - 在縮圖卡片模板（以及試算表 DataGrid 的 `<tr>`）綁定 `ondblclick="window.app.handleCardDblClick(${idx}, event)"`。
   - 實作 `handleCardDblClick(index, event)` 方法：
     - 精確排除按鈕、核取方塊、連結等次要控制項。
     - 自動選取該張照片並設定焦點索引（`currentIndex = index`）。
     - 若右側單張編輯面板為收合狀態（`editorCollapsed === true`），自動呼叫 `this.toggleEditor()` 展開面板。
     - 自動將游標聚焦（`focus()`）於現場跡證說明欄位（`#editDesc`），並將文字游標移至末端，支援立即鍵入說明。
3. **快速鍵與操作指南彈窗同步更新**：
   - 在「選取操作」分類新增「單張編輯：雙擊縮圖展開」之指南項目。
4. **全套自動化測試與 E2E 擴充**：
   - 在 `tests/e2e/photo-report.spec.js` 擴充 `[7/5]` 測試案例：實測點擊縮圖放大圖示、驗證燈箱開啟與 ESC 關閉、驗證收合狀態下雙擊卡片自動展開編輯器並聚焦 `#editDesc`。
   - 執行 `npm run test:all`（Phase 0A, 0B, 0C 全部通過）。
   - 執行 `scripts/qa.ps1` 通過。

## 刻意未修改
- 未改動 PaperSwitch 拖曳排版演算法與拖曳門檻（6px），確保卡片排序手感依然平順穩定。
- 未改動 Word、PDF、Excel、ZIP 匯出結構。

## 尚未完成
無

## 驗證結果
### 已執行
- `npm test`：全部通過 (4/4)
- `npm run test:e2e`：全部通過 (7/5 階段完整涵蓋放大燈箱與雙擊展開)
- `npm run test:baseline`：全部通過 (Word, Excel, PDF 結構符合 Golden Baseline)
- `npm run test:all`：全部通過 (Phase 0A, 0B, 0C)
- `powershell -ExecutionPolicy Bypass -File scripts\qa.ps1`：通過
- Playwright 實測截圖存檔：
  - `06_thumbnail_zoom_hover.png`（右下角放大圖示浮現效果）
  - `07_double_click_expanded_editor.png`（雙擊後自動展開右側單張編輯並聚焦 #editDesc）
  - `08_lightbox_zoomed.png`（點擊放大後燈箱大圖開啟）

### 尚未驗證
- 無阻斷性未驗證項目。

### 已知風險
- 無阻斷性風險。所有功能皆受單元、E2E 與 Baseline 回歸測試保護。

## Git 狀態
- Commit：40c9d79 (fix: 修復縮圖點擊放大燈箱與雙擊展開單張編輯功能)
- Push：否
- Working Tree：Clean
- Branch: main

## 下一步
- 成果交付完成。

