# Phase 0 自動化回歸基準與測試基礎建設 (Automated Regression Baseline)

> **定位與原則**：
> 本文件記錄 `04_Photo-Report-Generator` 在啟動漸進模組化（Phase 1～4）前的**自動化回歸測試基礎建設**與**回歸基準矩陣**。
> 核心策略：**自動化測試管「資料與結構對不對」，最終人工驗收管「視覺版型與操作手感」**。
> 各 Phase（抽離 `validation.js`、`audit.js`、`selection.js`、`history.js`、UI Controller、Exporter）完成後，均須執行自動化回歸測試，確認行為零偏差。

---

## 一、回歸測試四層防線架構

```text
tests/
├─ fixtures/               # 固定測試圖檔與預設案件資料 (sample01.jpg, sample02.jpg...)
├─ unit/                   # 1. 純邏輯自動化測試 (Node.js 100% 自動化)
│  ├─ validation.test.js   # 民國日期、時間格式合法性檢驗
│  ├─ audit.test.js        # 完整度稽核、同名照片計算、問題過濾
│  └─ history.test.js      # 快照簽名、Undo/Redo 狀態不污染檢驗
├─ e2e/                    # 2. Web UI 自動化測試 (Playwright 模擬真實操作)
│  └─ photo-report.spec.js # 載入、篩選列、Badge 同步、Modal 切換、快捷鍵
└─ baseline/               # 3. 匯出結構 Golden Baseline (DOCX, PDF, Excel 結構比對)
   ├─ docx-structure.json  # 表格行列數、欄寬、單元格結構
   ├─ excel-data.json      # 工作表欄位、標題、資料筆數比對
   └─ pdf-metadata.json    # 頁數、版面方向與文字分佈
```

---

## 二、第一層防線：純邏輯 100% 自動化 (Unit Tests)

- **執行工具**：Node.js (原生斷言庫，零額外相依套件負擔)
- **測試標的**：
  1. **`validation.test.js`**：
     - 民國日期合法性：7 碼數字、閏年判斷（民國 113 年 2 月有 29 日，民國 114 年無 29 日）、大小月、非數字過濾。
     - 時間合法性：4 碼 (HH:MM)、6 碼 (HH:MM:SS)、時分秒範圍 (0-23, 0-59, 0-59)。
  2. **`audit.test.js`**：
     - 同名照片集合計算：`_buildDupSet()` 在多組重複與不重複檔名下的精確 Set 判定。
     - 完整度統計：缺失地點、缺失說明、時間異常、同名照片之精確統計與首要過濾器指標。
     - 唯讀性驗證：確認呼叫 `auditPhotosCompleteness()` 前後，照片陣列物件未被新增、刪除或修改。
  3. **`history.test.js`**：
     - 快照簽名比對：確認 `historySignature()` 僅擷取純資料欄位（`uid`, `seq`, `date`, `time`, `location`, `desc`, `stageX`, `stageY`）。
     - 視圖隔離驗證：確認切換 `activeFilter` 或縮放比例時，歷史特徵簽名維持不變。

---

## 三、第二層防線：Web UI 自動化 (Playwright E2E)

- **執行工具**：Playwright (Chromium / Edge WebView2 核心)
- **測試標的**：
  1. **照片載入與 DOM 生成**：自動上傳 fixture 照片，驗證縮圖卡片正確出現在畫布。
  2. **完整度篩選列與 Badge 同步**：
     - 檢查「全部 / 未填地點 / 未填說明 / 日期時間待確認 / 同名照片」徽章數字。
     - 點擊「同名照片」篩選按鈕，驗證畫布卡片顯示數量等於同名照片數。
     - 點擊「全部」恢復顯示全部照片。
  3. **匯出前非阻斷提醒 Modal**：
     - 在照片有缺失資料時觸發匯出按鈕（Word / PDF / Excel）。
     - 斷言 `#exportAuditModal` 移除 `hidden` 類別。
     - 斷言標題文字包含對應匯出清冊名稱（如 `匯出前確認（Word 清冊）`）。
     - 測試點擊「查看問題照片」：Modal 自動關閉，且篩選列自動切換至首個問題分類。
     - 測試點擊「仍要匯出」：暫存之匯出回呼動作正確執行。
  4. **鍵盤導航與單次事件**：
     - 驗證點擊篩選按鈕無重複呼叫（單一事件來源）。
     - 驗證篩選模式下鍵盤導航僅在可見照片間切換焦點。

---

## 四、第三層防線：匯出內容與結構 Golden Baseline

- **核心原則**：不只驗證「檔案有產出」，而是將產出檔案解構為資料結構，與 Baseline 進行精準 Diff 比對。
- **比對機制**：
  1. **Word (`.docx`)**：
     - 將產出的 `.docx`（ZIP 格式）解壓縮，解析內部 `word/document.xml`。
     - 比對 XML 中的 `<w:tbl>` 表格數量、欄列數、欄寬（如標準 8302 dxa）、儲存格文字與圖片項目。
  2. **Excel (`.xlsx`)**：
     - 使用現有 `xlsx` 工具讀取產出之活頁簿。
     - 驗證工作表名稱、標題欄位（案由、日期、地點、序號、說明）與資料筆數。
  3. **PDF (`.pdf`)**：
     - 驗證輸出二進位結構、頁數、頁面寬高比例（直式 A4 與橫式 A4）。

---

## 五、第四層防線：Tauri Smoke Test 與最終人工視覺驗收

### 1. 各 Phase 的 Tauri Smoke Test
在模組化推進期間，各 Phase 僅執行：
- 檢查 `scripts/prepare-web.ps1` 是否成功產出 `web/`。
- 檢查 `src-tauri/tauri.conf.json` 設定無語法錯誤。
- 確保無任何破壞 WebView2 載入之 ESM 匯入路徑問題。

### 2. 全部 Phase 完成後之「最終人工視覺驗收」
人工驗收集中於全案模組化完成後進行**單次全面驗收**，專注於「自動化測試不擅長的體驗細節」：
- **Word / PDF 視覺版型**：以 Microsoft Word / 閱讀器實際開啟，肉眼檢查公務表格線條、5 點中繼段落行高、標楷體渲染與圖片採證質感。
- **操作手感**：實際操作拖曳排序時指示線跟隨手感、畫布平滑滾輪縮放、Marquee 框選流暢度。
- **Tauri 桌面端實裝**：啟動 Windows 獨立 Exe，確認原生視窗縮放與記憶體管理正常。

---

## 六、Phase 0 執行路徑 (Phase 0A ~ 0C)

```text
Phase 0A｜自動測試基礎建設
  ├─ 建立 tests/ 目錄架構與 fixtures 測試圖檔
  ├─ 完成 validation.test.js、audit.test.js、history.test.js
  └─ npm test 指令整合，確認 100% 通過

Phase 0B｜Web E2E 自動化腳本
  └─ 完成 Playwright 測試腳本 (photo-report.spec.js) 驗證 UI 流程

Phase 0C｜匯出結構 Golden Baseline
  └─ 建立固定案例之 DOCX/PDF/Excel 結構基準檔與比對工具

=> 完成 Phase 0 全部自動化基準後，方可解鎖 Phase 1 (抽離 validation.js / audit.js)
```
