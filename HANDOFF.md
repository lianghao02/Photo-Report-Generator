# HANDOFF

## 目前狀態
可交付（Phase 2 選取與歷史責任拆分完成，待進行 Phase 3）

## 本輪目標
落實 Phase 2：選取與歷史責任拆分 (Selection & History Models)。
將複雜的多選過濾、鍵盤導航與 Undo/Redo 快照簽名／歷史管理器獨立為專門模組：
1. `js/selection.js`：`getVisiblePhotoIndices`、`getSelectedIndicesOrCurrent`、`computeKeyboardNavIndex`。
2. `js/history.js`：`historySignature`、`projectSignature`、`HistoryManager`。
3. `index.html` 引入新模組，宿主內部同名方法改為委派調用以兼顧 100% 向後相容。
4. 更新單元測試 `tests/unit/selection.test.js` 與 `history.test.js`，將新測試整合入 `tests/run-all-unit.js`。
5. 執行 `prepare-web.ps1` 確保 `web/js/` 完整同步（Tauri 桌面端相容）。
6. 執行 `npm run test:all` 全套自動化回歸測試確認零回歸。

## 已完成
1. **建立獨立選取模組 (`js/selection.js`)**：
   - 封裝為 UMD 模式，依賴 `validation.js` 與 `audit.js`。
   - 提供 `getVisiblePhotoIndices`（支援 all / missingLocation / missingDesc / invalidDateTime / duplicatePhotos 篩選）。
   - 提供 `getSelectedIndicesOrCurrent`（取得勾選或焦點照片索引）。
   - 提供 `computeKeyboardNavIndex`（純邏輯計算鍵盤前後與跨列目標索引）。
   - 僅處理純資料索引計算，不介入任何指標與拖曳事件。
2. **建立獨立歷史管理器模組 (`js/history.js`)**：
   - 封裝為 UMD 模式。
   - 提供 `historySignature(state)`：精確提取純資料欄位（`uid`, `rotation`, `seq`, `date`, `time`, `location`, `desc`, `selected`, `stageX`, `stageY`、`caseData`、索引），排除 UI 暫存屬性。
   - 提供 `projectSignature(caseData, photos)`：計算未儲存變更特徵簽名。
   - 提供 `HistoryManager` 類別：管理快照堆疊、Undo/Redo、邊界判定與上限控制。
3. **`index.html` 整合與同名委派**：
   - 引入 `<script src="js/selection.js"></script>` 與 `<script src="js/history.js"></script>`。
   - `PhotoReportApp` 內部 `getVisiblePhotoIndices`、`getSelectedIndicesOrCurrent`、`selectPhotoByKeyboard`、`historySignature`、`projectSignature` 安全委派至新模組。
4. **單元測試套件整合與擴充**：
   - 建立 [`tests/unit/selection.test.js`](file:///C:/Development/GitHub/04_Photo-Report-Generator/tests/unit/selection.test.js)：完整覆蓋各篩選模式過濾、選取判定與鍵盤導航索引計算。
   - 強化 [`tests/unit/history.test.js`](file:///C:/Development/GitHub/04_Photo-Report-Generator/tests/unit/history.test.js)：覆蓋 `HistoryManager` 堆疊推移、Undo、Redo 與相同簽名防抖測試。
   - 更新 [`tests/run-all-unit.js`](file:///C:/Development/GitHub/04_Photo-Report-Generator/tests/run-all-unit.js) 加入 `selection.test.js`（測試總數升至 4 套件）。

## 刻意未修改
- **不碰觸 UI 渲染與 Exporter 模組**：維持 Phase 2 邊界，UI Controller 與 Exporter 留待 Phase 3 與 Phase 4 處理。
- **保留 API 簽名與呼叫約定**：既有外部調用完全透明相容。

## 尚未完成
- Phase 3：UI Controller 漸進拆分（`js/ui/audit-ui.js`、`modal-ui.js`、`photo-grid-ui.js`）。
- Phase 4：匯出模組分離（Word / PDF / Excel / ZIP）。

## 驗證結果
### 已執行
1. **Phase 0 全套自動化回歸測試 100% 通過**：`npm run test:all`
   - Phase 0A 單元測試：4/4 套件通過（`validation`, `audit`, `history`, `selection`）。
   - Phase 0B E2E 測試：5/5 流程通過。
   - Phase 0C 基準比對：3/3 格式完全吻合（Word 表格數/關鍵字、Excel 欄位資料行、PDF 頁數/尺寸）。
2. **共用 QA 檢驗通過**：`powershell -ExecutionPolicy Bypass -File scripts\qa.ps1`（exit code 0）。
3. **前端資產打包構建通過**：`powershell -ExecutionPolicy Bypass -File scripts\prepare-web.ps1`，成功驗證 `web/js/selection.js` 與 `web/js/history.js` 正確產出。

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
啟動 **Phase 3：UI Controller 漸進拆分 (UI Presentation Components)**：
1. 建立 `js/ui/` 目錄。
2. 抽離 `audit-ui.js`（完整度篩選列按鈕樣式與即時徽章渲染）。
3. 抽離 `modal-ui.js`（匯出確認與各 Modal 開閉、鍵盤 ESC 互動）。
4. 每次調整均以 `npm run test:all` 確保全套回歸零錯誤。
