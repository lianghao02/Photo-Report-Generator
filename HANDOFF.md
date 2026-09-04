# HANDOFF

## 目前狀態
可交付（Phase 1 純邏輯抽離完成，待進行 Phase 2）

## 本輪目標
落實 Phase 1：低風險純邏輯拆分（Pure Logic Extraction）。
抽離無副作用、純資料計算、不依賴 DOM 與 App State 的獨立邏輯：
1. `js/validation.js`：`isValidMinguoDate`、`isValidTimeFormat`。
2. `js/audit.js`：`buildDuplicateNameSet`、`auditPhotosCompleteness`。
3. `index.html` 引入新模組，宿主同名方法改為委派調用以兼顧 100% 向後相容。
4. 更新 `scripts/prepare-web.ps1`，確保建構時自動將 `js/` 模組目錄同步至 `web/js/`（Tauri 桌面端相容）。
5. 執行 `npm run test:all` 全套自動化回歸測試確認零回歸。

## 已完成
1. **建立獨立純邏輯模組 (`js/`)**：
   - [`js/validation.js`](file:///C:/Development/GitHub/04_Photo-Report-Generator/js/validation.js)：
     - 封裝為 UMD 模式（相容 Node.js 與純瀏覽器環境）。
     - 提供 `isValidMinguoDate(raw)`、`isValidTimeFormat(raw)`。
     - 零 DOM 依賴、無副作用。
   - [`js/audit.js`](file:///C:/Development/GitHub/04_Photo-Report-Generator/js/audit.js)：
     - 封裝為 UMD 模式，依賴 `validation.js`。
     - 提供 `buildDuplicateNameSet(photos)`、`auditPhotosCompleteness(photos, defaultLocation)`。
     - 保證 `photos` 唯讀、不修改傳入資料、不呼叫任何 UI 函式。
2. **`index.html` 整合與委派**：
   - 引入 `<script src="js/validation.js"></script>` 與 `<script src="js/audit.js"></script>`。
   - `PhotoReportApp` 內的 `isValidMinguoDate`、`isValidTimeFormat`、`_buildDupSet`、`auditPhotosCompleteness` 改為安全委派至新模組（包含降級容錯邏輯），維持 100% 呼叫相容性。
3. **構建與部署腳本強化**：
   - 更新 [`scripts/prepare-web.ps1`](file:///C:/Development/GitHub/04_Photo-Report-Generator/scripts/prepare-web.ps1)，加入同步 `js/` 目錄至 `web/js/` 的邏輯，保障 Tauri 桌面端與 `web/` 發布檔完整可用。
4. **單元測試套件升級**：
   - 更新 `tests/unit/validation.test.js` 與 `audit.test.js`，同時直接測試獨立模組導出函式與 App 委派方法，確保雙軌 100% 通過。

## 刻意未修改
- **不碰觸 UI 與匯出層**：嚴格遵守 Phase 1 邊界，不碰觸選取、歷史、DOM 渲染或 Exporter。
- **保留既有 API 簽名**：`PhotoReportApp` 既有同名方法均保留為轉發層，外部調用端零破壞。

## 尚未完成
- Phase 2：選取與歷史責任拆分（`js/selection.js`、`js/history.js`）。
- Phase 3：UI Controller 漸進拆分。
- Phase 4：匯出模組分離。

## 驗證結果
### 已執行
1. **Phase 0 全套自動化回歸測試 100% 通過**：`npm run test:all`
   - Phase 0A 單元測試：3/3 套件通過（涵蓋獨立模組直接斷言與宿主委派斷言）。
   - Phase 0B E2E 測試：5/5 流程通過（模擬檔案上傳、篩選列 Badge 統計、卡片可見度切換、匯出確認 Modal）。
   - Phase 0C 基準比對：3/3 格式完全吻合（Word DOCX XML 表格數/關鍵字、Excel 欄位資料行、PDF 頁數/尺寸）。
2. **共用 QA 檢驗通過**：`powershell -ExecutionPolicy Bypass -File scripts\qa.ps1`（exit code 0）。
3. **前端資產打包構建通過**：`powershell -ExecutionPolicy Bypass -File scripts\prepare-web.ps1`，成功驗證 `web/js/validation.js` 與 `web/js/audit.js` 正確產出。

### 尚未驗證
- 全模組化完成後之最終人工視覺版型驗收（依策略放至 Phase 4 結束後）。

### 已知風險
- 無阻斷性風險。

## Git 狀態
- Commit：待提交
- Push：待推送
- Working Tree：Clean（提交後）
- Branch：main

## 下一步
啟動 **Phase 2：選取與歷史責任拆分 (Selection & History Models)**：
1. 抽離 `js/selection.js`：可見照片過濾、鍵盤導航索引、多選計算。
2. 抽離 `js/history.js`：`HistoryManager`、`historySignature`、快照堆疊管理。
3. 每次調整均以 `npm run test:all` 確保全套回歸零錯誤。
