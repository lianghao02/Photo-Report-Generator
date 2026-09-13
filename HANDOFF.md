# HANDOFF

## 目前狀態
可交付

## 本輪目標
升級現場照片清冊生成工具之 UI/UX 至商用軟體水準，包含左側邊欄人體工學、版型微縮示意圖、中央智慧工具列模組化分群、商業級 Hero 空狀態、獨立快捷鍵與操作指南彈窗，並保證 100% 相容所有測試與匯出引擎。

## 已完成
1. **左側邊欄人體工學與首屏完整度（Above the Fold）**：
   - 案件資料欄位全面雙欄化（「採證日期/地點」、「製作單位/製作人」並排），大幅節省垂直高度。
   - 修正預編譯 Tailwind CSS 缺乏 `w-8` 導致 `#layoutMiniPreview` 撐爆版面問題，以專屬 CSS 明確鎖定 36px × 48px。
   - 確保在 800px 視窗高度下，專案檔、案件資訊、版型、Word 匯出、PDF 匯出、Excel 匯入/匯出、ZIP 打包「全數免滾動、一屏可見」，並附帶自適應捲動保護。
2. **中央智慧工具列結構化雙層設計**：
   - 將單一隨機換行的 `flex-wrap` 工具列重構為嚴謹的「雙層指揮台」結構：
     - 第一層：左側為批次編輯動作（全選、流水號、檔名時間、復原重做、旋轉、位移），右側為危險操作（移除選取、清空）。
     - 第二層：左側為檢視模式分段切換（工作台、試算表、分頁預覽），右側為縮放滑桿、排序、單張編輯。
   - 徹底根除危險操作按鈕被擠出第三行孤立於右側的版面破綻。
3. **商業級 Hero 空狀態與動態降噪**：
   - 0 張照片時自動隱藏頂部次要匯入列，消除畫面上下重複「加入照片/資料夾」的認知雜訊。
   - 0 張照片時工具列依賴照片之按鈕自動半透明禁用，焦點完全留給中央高挑舒適的 Hero 拖曳區。
   - 載入照片後，頂部匯入列、審計篩選列、縮圖網格與各操作按鈕無縫切換啟動。
4. **快捷鍵與操作指南彈窗**：
   - 移除非結構化的狀態列長文案，改為「[ ⌨️ 快速鍵指南 ]」按鈕與獨立 `#shortcutsModal`。
   - 以標準鍵盤鍵位（`<kbd>`）呈現四維度操作指南，支援背景點擊、關閉按鈕與 ESC 鍵快速關閉。
5. **完整回歸驗證**：
   - `npm test`（4/4 單元測試通過）
   - `npm run test:e2e`（Playwright UI 測試 6/6 階段通過）
   - `npm run test:baseline`（Word/Excel/PDF 匯出結構比對通過）
   - `scripts/qa.ps1`（QA 腳本通過）
   - Playwright 800px 高度空狀態與載入狀態實測截圖全數通過。

## 刻意未修改
- 未改動任何 DOM 元素 ID、Class 依賴與事件繫結架構。
- 未更動 Word、PDF、Excel、ZIP 匯出演算法或範本結構。

## 尚未完成
無

## 驗證結果
### 已執行
- `npm test`：全部通過 (4/4)
- `npm run test:e2e`：全部通過 (6/6 階段)
- `npm run test:baseline`：全部通過 (Word, Excel, PDF)
- `powershell -ExecutionPolicy Bypass -File scripts\qa.ps1`：通過
- Playwright 800px Viewport 實測截圖驗證（04_empty_state_ergonomic.png, 05_populated_state_ergonomic.png）
- `git diff --check`：通過

### 尚未驗證
- 無阻斷性未驗證項目。

### 已知風險
- 無阻斷性風險。所有匯出與既有功能皆受 E2E 與 Baseline 比對防護。

## Git 狀態
- Commit：cd83d7b (design: 優化左側首屏佈局與雙層工具列，消弭換行孤立與空狀態雜訊)
- Push：否
- Working Tree：Clean
- Branch: main

## 下一步
- 成果交付完成。

