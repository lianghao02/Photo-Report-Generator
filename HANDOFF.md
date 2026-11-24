# HANDOFF

## 核心元資料 (Metadata)
- **Repository**：lianghao02/Photo-Report-Generator
- **Branch**：main
- **Commit SHA**：未提交（前次遠端基線 `d4766d0`，v2.5.0）
- **Skill Version**：v1.0.0
- **Task Type**：FIX
- **Local Path Hint**：04_Photo-Report-Generator

---

## 目前狀態
可交付（本機待使用者確認是否提交並推送 `origin/main`）；全套單元測試、Playwright E2E、Golden Baseline 與共用 QA 皆已通過，依先前指示未建立桌面打包檔。

## 本輪目標
依全專案深度 AUDIT 清單，在不影響任何既有功能與三大 Word／PDF 版型尺寸的前提下，一次修訂 9 項 P1／P2／P3 問題：Word 手動換行、DOM XSS／URL 協定防線、Excel 匯出匯入修復、自動暫存排隊防覆蓋、試算表套用篩選與按視圖條件渲染、PDF 逐頁寫入降記憶體、預設參數集中管理、打包腳本可攜性與四檔版本檢查、拆出 3 個模組並全面改為事件委派以移除 Tauri CSP `script-src-attr 'unsafe-inline'`，以及瘦身本 `HANDOFF.md`。

## 基準與已確認事實 (Baseline & Confirmed Facts)
- 版本號四檔一致為 `v2.5.0`（`version.txt`、`package.json`、`src-tauri/tauri.conf.json`、`src-tauri/Cargo.toml`）。
- 全專案靜態 HTML（`index.html`）與動態 JS 渲染已達成 **0 處行內 `on*` 事件屬性**，Tauri CSP 已收緊移除 `script-src-attr 'unsafe-inline'`。
- 三大公務版型（`up_down_2`、`left_right_2`、`landscape_3`）之 Word／PDF 表格尺寸、段落數與 Golden Baseline 完全維持不變。

## 已完成 (Completed)
- **P1 #1 Word 手動換行 `<w:br/>`**：於 `js/exporters/docx-exporter.js` 加入 `buildMultilineTextRuns`，三種版型說明欄遇 `\n` 皆輸出正式 `<w:br/>`，且單行段落結構維持原樣通過 Golden Baseline。
- **P1 #2 XSS 與 URL 協定防線**：`js/project-format.js` 驗證 `photo.uid` 禁止不合法 HTML 字元；卡片 `id="thumb_..."` 與 `updateAnnotationElementDom` 補上跳脫防護；`showUpdateAvailable` 驗證 `release.html_url` 僅允許 `https:`／`http:`。
- **P2 #3 Excel 往返修復**：`js/exporters/excel-exporter.js` 將編號／日期／時間鎖定文字型別（`t: 's'`、`z: '@'`）保留前導零；`js/app.js` `importExcel` 改用 `raw: false`、補上前導零數值容錯比對、支援清空欄位，並補上 `recordHistory()` 支援 `Ctrl+Z` 復原與自動暫存。
- **P2 #4 自動暫存排隊防競態**：`saveAutoDraftNow` 改採 `autoDraftInFlight`／`autoDraftQueued`／`autoDraftSequence` 排隊與過期丟棄機制，並在 `clearAutoDraftDb` 時立即中止進行中的舊寫入。
- **P2 #5 試算表篩選與按需渲染＋就地選取**：`render()` 依目前視圖（`workspace`／`grid`）條件渲染不白重繪隱藏畫面；DataGrid 改用 `visibleEntries` 套用篩選，並以 `data-grid-row-index` 精準同步單列；單擊縮圖改為 `syncWorkspaceSelectionInPlace()` 就地更新選取樣式，兼顧大量照片順暢度與雙擊穩定性。
- **P2 #6 PDF 逐頁寫入記憶體控管**：`js/exporters/pdf-exporter.js` 改為每頁僅壓縮與轉換該頁 2～3 張照片、畫完即釋放暫存並以 `try/finally` 確保進度條清除。
- **P3 #7 拆檔與全站事件委派＋收緊 CSP**：新拆 `js/config.js`、`js/storage/auto-draft-store.js`、`js/storage/export-directory.js`、`js/ui/workspace-render-ui.js`；移除 `index.html` 全數 45 處與動態樣板全數 13 處行內 `on*` 屬性；`src-tauri/tauri.conf.json` 移除 `script-src-attr 'unsafe-inline'`；`scripts/prepare-assets.ps1` 補入新 UI 模組至 Tailwind `--content`。
- **P3 #8 預設參數集中與尊重清空**：透過 `js/config.js` 集中預設值，使用者刻意清空單位／製作人欄位時，Word／PDF／Excel 皆尊重空白不靜默填回。
- **P3 #9 打包腳本強化與文件瘦身**：`scripts/build-portable.ps1` 支援自動偵測 MinGW／MSVC 工具鏈並一次核對四檔版本號；更新 `AGENTS.md` 模組地圖並將 `HANDOFF.md` 瘦身僅保留最新交接重點（歷史紀錄由 Git 保存）。

## 異動檔案 (Changed Files)
- 新增：`js/config.js`、`js/storage/auto-draft-store.js`、`js/storage/export-directory.js`、`js/ui/workspace-render-ui.js`
- 修改：`index.html`、`js/app.js`、`js/project-format.js`、`js/exporters/docx-exporter.js`、`js/exporters/pdf-exporter.js`、`js/exporters/excel-exporter.js`、`src-tauri/tauri.conf.json`、`scripts/prepare-assets.ps1`、`scripts/build-portable.ps1`、`AGENTS.md`、`HANDOFF.md`、`tests/unit/project-core.test.js`、`tests/e2e/photo-report.spec.js`、`tests/e2e/desktop-policy.spec.js`

## 刻意未修改 (Do Not Do / Deliberately Omitted)
- 依使用者先前指示未建立 Tauri `.exe`／Portable 打包檔；未更動 `vendor/` 第三方套件版本或三大 Word／PDF 版型尺寸。

## 尚未完成 (Remaining Work)
- **P1 (阻斷/必須)**：無。
- **P2 (重要/當次)**：無。
- **P3 (改善建議/暫緩)**：未來若有需要可再將燈箱標記與並排裁切邏輯獨立拆分；Tauri 桌面 Release 待正式發版授權時再打包。

## 驗證結果 (Validation)
### 已執行測試與結果
- `powershell -ExecutionPolicy Bypass -File scripts\prepare-web.ps1`：通過，已重建 Tailwind 並自動同步 `index.html` `?v=` 快取標籤與 `web/`。
- `npm run test:all`：全部通過（單元測試 `5/5`、Playwright 4 份 E2E `photo-report.spec.js`〔含第 11 節綜合回歸〕／`desktop-policy.spec.js`〔含嚴格 CSP 零行內屬性驗證〕／`layout.spec.js`／`project-integrity.spec.js`、Word／Excel／PDF Golden Baseline 及三種版型）。
- `powershell -ExecutionPolicy Bypass -File scripts\qa.ps1`：通過。
### 尚未驗證項目
- 本輪未重建 Tauri 桌面板 `.exe`（依指示略過），亦未提交／推送至 `origin/main`。
### 已知風險 (Known Risks)
- 受機關 Chrome GPO（`DefaultCookiesSetting = 4`）控管環境下，File System Access 寫入固定資料夾仍需改用 Edge 或維持預設下載區。

## Git 狀態
- Commit：本輪待提交（遠端基線 `d4766d0`）
- Push：否（待使用者確認）
- Working Tree：Modified
- Branch：main

## 下一步建議動作 (Next Recommended Action)
待使用者確認後執行 `git commit` 並 `git push origin main` 更新線上版。

## 發布狀態 (Release Status)
線上網頁版待本輪提交推送後生效；桌面版暫不建立。
