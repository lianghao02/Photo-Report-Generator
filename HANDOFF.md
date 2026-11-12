# HANDOFF

## 核心元資料 (Metadata)
- **Repository**：lianghao02/Photo-Report-Generator
- **Branch**：main
- **Commit SHA**：隨本輪提交（父提交 392e0a12d634fe50b7b41eb089d1a46379963378；本檔包含於提交，實際 SHA 見 git log -1）
- **Skill Version**：v1.0.0
- **Task Type**：IMPROVE
- **Local Path Hint**：04_Photo-Report-Generator

---

## 目前狀態
可交付（本機修訂版）；五階段計畫、完整測試、Portable跨路徑原生流程及自我檢查已完成。無阻斷與重要問題，依使用者指示停止擴大修改。

## 本輪目標
依IMPLEMENTATION_PLAN.md保存現有修復與UI基準、補強測試失敗判定、抽出照片／進度格式／彙整邏輯，完成存檔狀態、開檔檢查及彙整明細。使用D:\PicPick照片副本實測，原始資料保持不變。使用者於2026-11-12追加授權「開始修正提交」，本輪固定已驗證成果至本機Git；不推送、不升版、不發布。

## 基準與已確認事實 (Baseline & Confirmed Facts)
- 開始本計畫時main已有前輪桌面修復、UI與計畫等未提交修改；不得覆寫、reset或只以HEAD當作回復點。
- 已保存99個正式來源／新增檔案、baseline.patch、baseline-status.txt與前輪Portable產物，均位於被忽略的build/plan-implementation/baseline及baseline-artifacts。
- 進度格式維持version 1。相容樣本為修改前程式實際產生的v1檔案，另外合成省略版本與錯誤案例；未取得更早正式發行版的真實進度檔，不宣稱全部歷史發行版皆已實測。
- 原有彙整判定以50組修改前規則輸出及完整既有E2E比對；既有非空案件資料等合併行為保持。
- 使用者提供的4張照片只複製至被忽略的測試區；來源路徑／雜湊紀錄不加入版本庫。

## 已完成 (Completed)
- 繼承前輪CSP、加入照片按鈕、拖放設定與六尺寸UI修訂；保存穩定基準。
- 照片資料統一至js/photo-model.js，區分可保存欄位與File／Object URL執行期資源；歷程副本隔離可變資料。
- js/project-format.js統一進度與草稿格式、支援版本、欄位結構、ZIP CRC、圖片清單及解碼檢查。開啟／彙整先驗證；異常、取消保留目前工作，釋放臨時URL。保留舊超標文字。
- js/project-merge.js集中純資料彙整規則與固定前後差異；App保留工作流程、圖片資源與畫面協調。
- 新增最近成功彙整明細，由js/ui/merge-report-ui.js安全呈現文字／旋轉／標記／特寫與案件欄位前後值；支援Esc、焦點與刪除快捷鍵防護。失敗不覆蓋前次成功紀錄，開啟另一進度／關閉程式不保留，Undo不改寫歷史快照。
- 存檔狀態明確區分「已產生下載檔」、「已寫入檔案」、「已載入工作進度」及最新自動草稿。實際寫入須等write與close完成，失敗嘗試abort。
- 修正非同步儲存／草稿期間繼續編輯的狀態誤判：固定開始時資料快照，後續修改仍標未存，舊草稿不顯示為新內容已暫存；補足競態測試。
- 主要E2E遇pageerror真正失敗；以刻意拋錯／缺模組反例驗證。新增直接測模組的單元測試與進度檔完整性E2E。
- 已更新README、AGENTS、Tailwind掃描來源及完成計畫。
- 提交前核對分支main與origin、25個檔案的完整範圍、既有測試與原生驗證證據；更新本輪提交授權與交接狀態，不擴大產品修改。

## 異動檔案 (Changed Files)
- 累計既有與本輪修改：AGENTS.md、HANDOFF.md、README.md、IMPLEMENTATION_PLAN.md、index.html、css/app.css、js/app.js、package.json、scripts/prepare-assets.ps1、src-tauri/tauri.conf.json、vendor/tailwind.css。
- 新模組：js/photo-model.js、js/project-format.js、js/project-merge.js、js/ui/merge-report-ui.js。
- 測試：tests/e2e/photo-report.spec.js、desktop-policy.spec.js、layout.spec.js、project-integrity.spec.js；tests/helpers/page-errors.js、tests/run-all-unit.js、tests/unit/project-core.test.js；tests/fixtures/project-v1.photo-report、merge-compatibility.json、project-fixtures.md。
- build、web、target及真實照片副本均屬被忽略的本機產物，不應加入Git。

## 刻意未修改 (Do Not Do / Deliberately Omitted)
使用者排除輸出前預覽，大量照片效能暫緩；不拆影像／燈箱、不新增範本、不新增桌面原生另存API、不更換框架。不修改三種匯出版型尺寸、不升級套件；版本維持2.4.1。

## 尚未完成 (Remaining Work)
- **P1 (阻斷/必須)**：無。
- **P2 (重要/當次)**：無。
- **P3 (改善建議/暫緩)**：Windows 10、其他電腦、125%／150%系統DPI與安裝版實際安裝驗證，正式發布時另安排。

## 驗證結果 (Validation)
### 已執行測試與結果
- 最終npm run test:all通過：單元5/5（含50組彙整相容規則）、完整既有E2E、桌面CSP、六尺寸排版、進度完整性／失敗保護／非同步競態、Word／Excel／PDF Golden Baseline及三種版型。
- 測試使用既有Node Playwright；同時驗證file://直接開啟，未新增測試相依套件。刻意觸發pageerror與必要模組缺少，防線確實拒絕。
- 以PHOTO_REPORT_TEST_DIR指向D槽照片副本完成4張匯入、保存、重開與文字內容比對，並實際輸出DOCX／PDF。
- 最新原始碼已重建release EXE、Setup及Portable ZIP，未正式發布。原始碼與web的8個關鍵資源逐位元組一致。
- 最終Portable解壓至本輪C槽工作區outputs/photo-report-plan-complete/中文 Portable，跨磁碟及含中文、空白路徑正常啟動。
- 原生視窗完成：開啟4張進度、彙整另1份（去重4張、較新更新1張）、查看說明前後明細及Esc關閉、Ctrl+S正確顯示下載模式、匯出Word與PDF。
- 原生視窗另以檔案選擇器加入1張照片副本並修改說明、儲存、關閉重啟、開啟下載檔；確認共5張照片與「桌面原生編輯與重開驗證」說明完整還原。
- 原生下載產物實際解壓與內容檢查：4／5張進度檔CRC及圖片均完整；DOCX含案件名稱、彙整後說明與4張圖片；PDF標頭正確且2頁。
- D:\PicPick的4張來源SHA-256與開始時紀錄全部一致，原始檔未變更。
- 解壓EXE與建置EXE逐位元組一致，SHA-256：d4106cf2079c29ca025a22b13759ecf35fab6e2cf18123da81d37290e841fcf5。
- node --check、git diff --check及scripts/qa.ps1通過。最後文件更新後再核對Git與QA；無敏感資料警訊。
- 本輪證據位於被忽略的build/plan-implementation/results：網頁產物／截圖、native實際產物與verification.json、SELF_REVIEW.md。含使用者照片的產物不加入版本庫。
### 尚未驗證項目
Windows 10、其他電腦、125%／150%系統DPI、檔案總管原生拖入及安裝版實際安裝未另實測。原生測試使用本機預設視窗，六尺寸是瀏覽器可視區測試；不混稱為多DPI實機驗證。原生直接寫入分支使用瀏覽器測試控制代碼驗證，Portable實際採下載流程。
### 已知風險 (Known Risks)
下載流程只能確認交付瀏覽器，無法可靠確認使用者最終磁碟存檔；介面已明白提示檢查下載資料夾。彙整明細只保留當次工作階段最近成功紀錄，不是永久修改歷程。沒有本輪已知阻斷或重要問題。

## Git 狀態
- Commit：本輪提交（包含本檔；實際SHA見git log -1，父提交392e0a12d634fe50b7b41eb089d1a46379963378）
- Push：否
- Working Tree：本輪25個檔案提交後確認Clean；包含前輪桌面修復／UI及本輪計畫實作，最終狀態以git status為準
- Branch：main

## 下一步建議動作 (Next Recommended Action)
本輪已達完成條件，依授權完成本機Git提交後停止。使用者可測試最新Portable；後續推送或正式發布另依明確授權及Release Gate辦理。

## 發布狀態 (Release Status)
本機修訂版可交付，已依使用者授權納入本輪Git提交。維持v2.4.1，Setup與Portable已重新建置；未推送、未更新線上版或遠端Release。

---

以下為先前交接紀錄，現況以本輪狀態與 Git 為準。

# HANDOFF

## 核心元資料 (Metadata)
- **Repository**：lianghao02/Photo-Report-Generator
- **Branch**：main
- **Commit SHA**：未提交（基準 392e0a12d634fe50b7b41eb089d1a46379963378）
- **Skill Version**：v1.0.0
- **Task Type**：IMPROVE
- **Local Path Hint**：04_Photo-Report-Generator

---

## 目前狀態
可交付（本機 UI 修訂版）；最終完整測試、六尺寸版面測試、桌面重新建置及原生工作台驗證全部通過。

## 本輪目標
承接前輪桌面修復成果，修訂網頁與桌面共用 UI：匯出列不遮擋、案件資料可收合、照片與表單獨立捲動、固定照片切換、改善工具列換行。不提交、不推送、不發布遠端版本。

## 基準與已確認事實 (Baseline & Confirmed Facts)
- 任務開始時 main 工作目錄乾淨，HEAD 為 392e0a1。
- 原桌面 EXE、Setup 與 Portable 為 2026-10-22 產物，未包含 11 月功能。
- 前輪實機觀察空白工作台「加入照片」無檔案選擇視窗；CSP 根因尚未直接取得舊 EXE 主控台證據。
- Tauri 預設原生拖放會攔截 Windows HTML5 拖放；前端使用 HTML5 drop，未接收原生事件。
- 原 E2E 寫死個人截圖路徑並因 EPERM 失敗。

## 已完成 (Completed)
- 空白工作台加入照片／資料夾改為具 ID 按鈕與 addEventListener。
- Tauri 設定 dragDropEnabled=false；script-src 限制為 self，獨立 script-src-attr 保留現有行內事件，避免 nonce／hash 使 unsafe-inline 失效。
- 個人截圖改為選用 PHOTO_REPORT_SCREENSHOT_DIR；未設定時不寫入除錯圖片。
- 新增 desktop-policy.spec.js：本機 HTTP CSP 加入 nonce，驗證照片／資料夾選擇、照片載入、動態縮圖燈箱與拖放設定；納入 test:e2e 與 test:all。

- UI 排版修訂：桌面尺寸以可視高度分配工作區；匯出列獨立、不覆蓋照片或說明欄；照片、表格與編輯欄各自捲動，編輯欄底部照片切換固定可見。
- 案件欄新增原生 details 收合；1200px 以下預設收合，768～1199px 展開使用浮動面板，不壓縮照片導覽。
- 主編輯欄維持至少280px、最多320px，短欄位並排；寬視窗最大工作區擴至1800px。小於768px維持文件捲動。
- 操作說明改為明確的「操作提示」按鈕，工具列以群組換行並保護按鈕文字不被拆行。
- 新增 tests/e2e/layout.spec.js 並納入完整測試，驗證1920×1080、1366×768、1100×760、900×600、768×600、390×844；涵蓋匯出列／照片不重疊、照片導覽可見、編輯欄寬、案件欄展開、編輯欄收合、表格與燈箱。
- 重建 Tailwind，逐一展開選擇器比對：原有選擇器零刪除，新增 .underline；未升級套件。

## 異動檔案 (Changed Files)
css/app.css、index.html、js/app.js、src-tauri/tauri.conf.json、package.json、vendor/tailwind.css、tests/e2e/photo-report.spec.js、tests/e2e/desktop-policy.spec.js、tests/e2e/layout.spec.js、HANDOFF.md（含前輪修復尚未提交範圍）。

## 刻意未修改 (Do Not Do / Deliberately Omitted)
三大匯出版型、業務資料模型、vendor 套件版本、應用版本號與遠端版本皆未調整；打包產物為本機修復驗證版 v2.4.1，不代表已正式發布。

## 尚未完成 (Remaining Work)
- **P1 (阻斷/必須)**：無。
- **P2 (重要/當次)**：無。
- **P3 (改善建議/暫緩)**：Windows 10、其他電腦及不同系統 DPI 環境驗收另行安排。

## 驗證結果 (Validation)
### 已執行測試與結果
- npm run test:all：單元 4/4、完整 E2E、桌面 CSP 回歸、Word／Excel／PDF Golden Baseline 及三種版型匯出全部通過，未使用測試繞過。
- node --check：js/app.js、tests/e2e/desktop-policy.spec.js 通過。
- scripts/qa.ps1 通過；vendor 重建後無實質差異。
- web/index.html、web/js/app.js 與來源 SHA-256 一致。
- 本輪已以最終 UI 原始碼重新建置 release EXE、NSIS Setup 與 Portable ZIP；版本維持2.4.1，未發布遠端版本。
- 前輪修復驗證：Portable ZIP 解壓至 C 槽「中文路徑 Portable 163439」，從 D 槽開發目錄以外位置兩次正常啟動；EXE 與建置產物 SHA-256 同為 2AF1E7E218573EEF8E24CED9E03BB3F04D374B30CE68D2088904D21A42D2B95F。
- 原生視窗實測：加入照片開啟檔案選擇視窗、載入 sample01.jpg／sample02.jpg、編輯說明及 16／68 全形計數、向下填滿、Ctrl+S 存檔、匯出 Word／PDF，全部通過。
- 結束程式後重新啟動，再由原生檔案選擇視窗開啟 .photo-report，顯示「已開啟工作進度，共 2 張照片」；案件名稱、照片與說明完整還原，資料檢查通過。
- 實際產物檢查：.photo-report ZIP 包含 manifest.json 與 2 張圖片，兩張 desc 均一致；DOCX XML 包含驗證案件名稱與說明且含 2 個圖片媒體檔；PDF 標頭有效且為 1 頁。
- 驗證產物已另複製至本輪工作區 outputs/photo-report-desktop-20261112；只使用合成測試照片，未處理使用者案件資料。
- 最終 npm run test:all 再次全數通過（單元4/4、完整E2E、桌面CSP、六尺寸排版、Word／Excel／PDF Golden Baseline）。
- 本輪原生桌面：最終Portable解壓至C槽「桌面 排版修訂版」，正常啟動並還原前輪兩張合成測試照片。說明欄、上一張／下一張與匯出列在首屏可見，點擊下一張成功切換第2／2張。
- 最終桌面截圖：本輪工作區 outputs/photo-report-ui-revision/desktop-final.png；網頁六尺寸截圖亦位於該資料夾。
- 原始碼與web打包資源（index.html、css/app.css）SHA-256一致；Tailwind原有選擇器零刪除。
### 尚未驗證項目
Windows 10、其他電腦、系統125%／150% DPI、由檔案總管拖入檔案及安裝版實際安裝未另驗證。六尺寸含900×600為瀏覽器可視區測試；原生視窗本輪驗證其實際預設尺寸，未聲稱完成原生900×600驗證。
### 已知風險 (Known Risks)
打包曾因前輪開啟的舊 EXE 鎖住檔案而失敗；已透過桌面 API 關閉該空白檢查視窗後重試。

## Git 狀態
- Commit：未提交（基準 392e0a1）
- Push：否
- Working Tree：Modified，均為本輪修改
- Branch：main

## 下一步建議動作 (Next Recommended Action)
使用者已要求先建立白話實作計畫，見 IMPLEMENTATION_PLAN.md；本輪僅新增計畫，尚未開始分檔或新功能實作。後續接獲開始指示後，依計畫第1階段保存完整工作目錄基準；輸出預覽排除、大量照片處理暫緩。正式發布另依Release Gate與授權處理。

## 發布狀態 (Release Status)
本機修復版可交付；維持 v2.4.1，未升版、未提交、未推送、未發布遠端版本。

---

以下為先前交接紀錄，現況以本輪狀態與 Git 為準。

# HANDOFF

## 核心元資料 (Metadata)
- **Repository**：lianghao02/Photo-Report-Generator
- **Branch**：main
- **Commit SHA**：隨本輪一併提交推送（承接基準 `c02ad13c20ab8232c798dee26f2f3434cdc6c2fe`，實際 SHA 見 `git log -1`）
- **Skill Version**：v1.0.0
- **Task Type**：FIX
- **Local Path Hint**：04_Photo-Report-Generator

---

## 目前狀態
可發布（已完成 Word 案由／日期／時間置中對齊修正、地點 17 全形／說明 68 全形與 2 行防呆控制，以及縮圖「並排／特寫／翻轉／標記」整併為右下角單一膠囊並通過全套自動化測試與實機截圖驗證，準備執行本輪 Commit 與 Push 更新線上版）。

## 本輪目標
1. 修正 Word 匯出之「案由、日期、時間」欄位對齊方式為 **水平置中對齊**（`AlignmentType.CENTER`），取代原本受 `JUSTIFIED` 影響造成單行靠左顯示的問題。
2. 解決 A4 直式上下兩張（`up_down_2`）中「說明」欄位超過 2 行變成 3 行時將同頁第二張說明推至下一頁的跑版問題：
   - 說明欄位限制最多 **68 個全形佔位**（136 半形單位）且最多 **2 視覺行**（單行 34 全形），防止超出兩行。
   - 地點欄位（案件預設「採證地點」與單張「地點」）限制最多 **17 個全形佔位**（34 半形單位），單行不可換行。
   - 落實 **兩個半形字元（例如 `()`、ASCII 英數／空白／半形標點）計為 1 個全形佔位**，並妥善處理 Unicode Surrogate Pair 與 Windows 注音輸入法（IME）組字過程。
   - 保留已開啟之舊 `.photo-report` 專案檔既有超標文字不被靜默截斷，改以警示、逐步刪減、縮圖／試算表紅框、完整性篩選（「字數過長」）與匯出前提醒處理。
3. 消除小縮圖下左側 `.thumb-sidecrop-badge` 與右側 `.thumb-annotation-badge` 互疊遮擋照片底部的問題：依使用者選定之方案 **A2（合併成單一右下角膠囊）** 整合特寫／並排／翻轉／標記數，並依縮圖寬度自動切換清楚與精簡標記。

## 基準與已確認事實 (Baseline & Confirmed Facts)
- 前一輪 Commit `c02ad13c20ab8232c798dee26f2f3434cdc6c2fe` 已同步至 `origin/main`（GitHub Pages 線上版承接該次提交）。
- 實測 `vendor/docx.umd.js`（`docx@8.5.0`）：`AlignmentType.CENTER` 輸出 `<w:jc w:val="center"/>`；含 `\n` 的 `TextRun` 輸出 `<w:t xml:space="preserve">...</w:t>`，截圖之所以出現 3 行是因為含半形字元合計逾 68 全形（單行 34 全形，雙行合計 68 全形）折行超過表格餘裕高度（~130 dxa）。
- 所有新增 HTML 類別經確認皆已存在於 `vendor/tailwind.css`（無須異動 `vendor/`）。
- 縮圖於 `90px～110px` 緊湊模式下內部舞台寬度僅約 `74～94px`，既有左右兩個絕對定位膠囊（合計逾 `82px`）必然碰撞並遮蓋照片底緣；統一整合至右下單一膠囊並保留完整滑鼠懸停 `title` 可完全避免重疊。

## 已完成 (Completed)
- **Word 匯出對齊修正（`js/exporters/docx-exporter.js`）**：
  - `up_down_2` 第 1 列：`案由`、`日期`、`時間` 之數值儲存格由 `AlignmentType.JUSTIFIED` 改為 `AlignmentType.CENTER`（與標籤、編號、地點、製作人一致）。
  - `landscape_3` 第 1 列與第 2 列：`蒐證時間`、時間值、`蒐證地點`、地點值、`製作人`、製作人值與 `照片編號` 統一為 `AlignmentType.CENTER` ＋ `VerticalAlign.CENTER`。
- **全形／半形寬度計算與截斷核心（`js/validation.js`）**：
  - 匯出 `MAX_LOCATION_FULLWIDTH = 17`、`MAX_DESC_FULLWIDTH = 68`、`DESC_LINE_FULLWIDTH = 34`、`MAX_DESC_LINES = 2`，以及 `isHalfWidthChar`、`getTextHalfWidthUnits`、`getTextFullWidthLength`、`formatFullWidthCount`、`estimateDescVisualLines`、`truncateToFullWidth`、`inspectTextLimits`。
  - 支援 UTF-16 Surrogate Pair 安全切字，並於多餘換行超出 `maxLines` 時僅略過多餘換行符號而不刪除後方文字。
- **即時介面提示、IME 組字與舊檔相容保護（`index.html`、`css/app.css`、`js/app.js`）**：
  - `#defaultLocation`、`#editLocation`、`#editDesc` 加入即時全形計數器（`X / 17 全形`、`X / 68 全形`）與剩餘／超限警示提示（接近上限轉琥珀色，超限或觸及截斷轉紅色閃爍 2.6 秒提醒）。
  - 實作 `bindLimitedTextControl` 與 `bindGridTextLimitEvents`，於 `compositionstart`／`compositionend` 保護注音等 IME 組字不被提早截斷，且工作台右側欄與試算表視圖（DataGrid）雙向即時同步。
  - 若載入含超標文字的舊 `.photo-report` 專案檔，保留原文字不被破壞，顯示紅框警示，編輯時允許逐步刪減至上限內但不允許繼續增字；變更案件預設採證地點時同步重刷各縮圖狀態燈與稽核列。
- **縮圖單一合併膠囊（A2，防小圖重疊，改動 `css/app.css`、`js/app.js`）**：
  - 移除左下各自定位的 `.thumb-sidecrop-badge` 與 `.thumb-crop-badge`，於 `getPhotoAnnotationBadgeHtml()` 將「特寫／◫並排／⇄翻轉／標記N」整合進右下角單一 `.thumb-annotation-badge`（加掛 `data-badges` 與完整 `title`）。
  - 單一並排／單一特寫分別維持墨綠（`--sidecrop`）與天藍（`--crop`）視覺色系；一般寬度（`> 125px`）且僅 1～2 項時顯示清楚全名（如 `◫並排・標記2`），於小縮圖（`<= 125px`）或 3 項以上並存時自動切換緊湊符號（如 `◫・標2`、`◫・⇄・標1`、`特・◫・標1`），搭配 `max-width: calc(100% - 8px)` 與 `text-overflow: ellipsis` 防溢位。
- **完整度稽核、篩選與匯出前提醒（`js/audit.js`、`js/selection.js`、`js/ui/audit-ui.js`、`js/ui/modal-ui.js`、`js/app.js`）**：
  - 新增 `overTextLimit` 稽核統計（同時檢查單張地點、繼承案件地點與說明）、篩選按鈕 `#filterBtn_overTextLimit`（`字數過長`），以及匯出確認對話框項目 `地點或說明過長（恐影響排版）：X 張`。
- **自動化測試與實機視覺驗證（`tests/`）**：
  - `tests/unit/app-bridge.js` 修正擷取方法邊界以正確支援預設物件參數簽名，並新增全形計算與超標稽核測試（`validation.test.js`、`audit.test.js`、`selection.test.js`）。
  - `tests/e2e/photo-report.spec.js` 更新第 `2E` 節驗證單一合併膠囊於 150px／95px 縮圖之顯示與幾何邊界防溢位，並新增第 `[8/5]` 節實際解壓 `.docx` `word/document.xml` 驗證 `<w:jc w:val="center"/>`、半形 `()`＝1 全形、使用者 3 行範例截斷、多行防呆、注音 IME 模擬、DataGrid 雙向同步及舊超標專案檔保留／逐步刪減／篩選／匯出提醒。

## 異動檔案 (Changed Files)
css/app.css、index.html、js/app.js、js/audit.js、js/exporters/docx-exporter.js、js/selection.js、js/ui/audit-ui.js、js/ui/modal-ui.js、js/validation.js、tests/e2e/photo-report.spec.js、tests/unit/app-bridge.js、tests/unit/audit.test.js、tests/unit/selection.test.js、tests/unit/validation.test.js、HANDOFF.md。

## 刻意未修改 (Do Not Do / Deliberately Omitted)
`up_down_2`、`left_right_2`、`landscape_3` 之表格欄寬（DXA）、邊距與中繼 5 點行高結構；版本號與 Tauri 桌面版安裝檔打包（待功能定版後視需求進行）。

## 尚未完成 (Remaining Work)
- **P1 (阻斷/必須)**：無。
- **P2 (重要/當次)**：無。
- **P3 (改善建議/暫緩)**：日後發布新版桌面安裝檔時可一併升版與打包 Tauri MSI／NSIS。

## 驗證結果 (Validation)
### 已執行測試與結果
- `node --check`（全部修改與測試 `.js` 檔）：通過。
- `git diff --check` 與敏感資料掃描、`powershell -ExecutionPolicy Bypass -File scripts\qa.ps1`：通過。
- `npm test`（`node tests/run-all-unit.js`）：4/4 全部通過（涵蓋模組與 `MockApp` 宿主委派）。
- `npm run test:baseline`（`node tests/run-baseline-test.js`）：Word／Excel／PDF Golden Baseline 與三種版型匯出全數通過。
- `npm run test:e2e`（`node tests/e2e/photo-report.spec.js`）：全部通過（含單一合併膠囊雙尺寸實測與第 `[8/5]` 節 DOCX XML 置中／字數限制全套驗證）。
- Playwright 實機視覺截圖檢視（1366×768 表單上限／閃爍提示、100px 與 150px 縮圖合併膠囊）：無水平溢位、無徽章重疊、主控台零錯誤。
### 尚未驗證項目
未重新升版或打包 Tauri 桌面版安裝檔。
### 已知風險 (Known Risks)
無本輪已知阻斷性問題。

## Git 狀態
- Commit：隨本輪提交至 `main`（實際最新 SHA 請見 `git log -1`，承接基準 `c02ad13`）
- Push：將隨本輪驗證通過後推送至 `origin/main`
- Working Tree：Clean（提交後）
- Branch：main

## 下一步建議動作 (Next Recommended Action)
線上版使用者可於 GitHub Pages 部署完成後實地匯出一份 `up_down_2` 清冊確認置中對齊、68／17 全形限制與小縮圖單一合併膠囊手感。

## 發布狀態 (Release Status)
可發布（全套單元測試、Golden Baseline、Playwright E2E、QA 檢核與實機截圖皆已通過）。

---

以下為先前交接紀錄，現況以本輪狀態與 Git 為準。

## 2026-11-10 v2.4.1 工作進度儲存強化、依採證時間排序與多檔智慧彙整（已提交 c02ad13）
- **核心元資料**：`lianghao02/Photo-Report-Generator`、`main`、承接 `27521102`、`IMPROVE`。
- **目前狀態**：已完成並推送至 `origin/main`（Commit `c02ad13`）。
- **本輪目標**：
1. 全面改善「儲存／開啟工作進度（`.photo-report`）」之日常操作體驗與底層穩定性：
   - 採用方案 A：一般「儲存工作進度（`Ctrl+S`）」與「另存新進度檔…（`Ctrl+Shift+S`）」預設自動執行 220 ppi 等級（長邊 `<= 2000px`）圖片最佳化壓縮並以 `STORE` 快速封裝，不另立重複的「儲存輕量進度檔」按鈕，同時 100% 保留黃框、姓名標籤、旋轉翻轉與裁切並排之非破壞性可編輯狀態。
   - 存檔／開檔／草稿恢復／彙整合併支援進度條「取消」按鈕；修正錯誤提示、補齊 `finally { this.clearProgress(); }`、防止中途失敗之 Object URL 記憶體洩漏，並於開檔與彙整時自動關閉燈箱與重置篩選器為 `all`。
   - 專案檔名自動過濾非法字元（`sanitizeFilename`）；存檔與合併前自動定稿輸入中欄位（`finalizeActiveInputsBeforeSave`）。
   - 支援同檔原地覆寫（File System Access API `currentProjectFileHandle`）、`Ctrl+S` 全域快速存檔、頂部導覽列「未儲存／已儲存」即時狀態徽章（`#btnQuickSaveStatus`）、本機 IndexedDB 自動暫存草稿與恢復橫幅（`#autoDraftBanner`），以及直接拖放 `.photo-report` 進畫面開檔。
2. 支援多 `.photo-report` 工作進度檔智慧彙整合併：
   - 提供「檔案 ▾ → 併入其他工作進度…」（`#mergeProjectInput`，支援複選）與一次拖放 2 個以上 `.photo-report` 進畫面直接自動彙整之雙入口。
   - 支援同源 `uid` 配對與跨電腦／經過 2000px 重壓縮之雙重身分識別（正規化檔名 ＋ 長寬比 ＋ 16×16 dHash 與 2×2 象限色彩視覺指紋 `visualFingerprint`），同時避免同檔名不同畫面照片誤合併。
   - 依使用者定案規則進行無衝突預設自動合併：「一空白、一有字 ➔ 自動補齊」、「兩邊都有且不同 ➔ 以比較晚存檔／編輯的那一份覆蓋」，並具備單張未存修改防護（沒存檔的新修改不會被較舊進度檔蓋回）、未手動寫入真實時間預設 `12:00`（`timeAutoDefaulted`）不誤蓋真實時間、不同案件預設日期／地點保護、子特寫緊跟母圖去重或續編、新照片依序接續與 `Ctrl+Z` 完整復原等安全防護。
3. 在「更多 ▾ → 排序與編號」選單新增 **「依採證時間排序」**（`#btnSortByCaptureTime`），依每張照片之採證日期（空白時繼承案件日期）與採證時間由早到晚排序，並維持 `-1`、`-2` 子特寫緊跟母圖連動。
4. 依使用者指示精簡介面：移除不再需要的「分頁預覽」、「從 Excel 匯入說明」、「匯出 Excel 清單」與「下載處理後照片 (ZIP)」選單項目。

## 基準與已確認事實 (Baseline & Confirmed Facts)
- 前一輪 Commit `275211027f7c4b07ad95e0e710063f83f5434d95` 已推送至 `origin/main`。
- `tests/e2e/photo-report.spec.js` 第 `1A` 節要求 `btnSaveProject`、`mergeProjectInput` 等控制項位於 `details.app-menu` 內，因此頂部導覽列快捷存檔狀態徽章使用獨立 ID `#btnQuickSaveStatus`。
- 未標記、無裁切之一般匯出流程與 `.photo-report` 新舊格式維持 100% 向下相容。

## 已完成 (Completed)
- **底層穩定度與預設 220 ppi 最佳化壓縮**：
  - `js/app.js`（`saveProject`、`createLitePhotoEntry`、`openProjectFromFile`、`handleProjectError`、`sanitizeFilename`、`finalizeActiveInputsBeforeSave`）：
    - `saveProject()` 預設啟用 `createLitePhotoEntry(photo, 2000, 0.88)`，自動將超過 `2000px` 之超大原圖依 EXIF 轉正並縮放為長邊 `<= 2000px` JPEG（同時將存入的 `exifRotation` 歸零避免二次旋轉，並完整保留 `rotation`、`flipH`、`annotations`、`cropRect`、`sideCrop` 與 `visualFingerprint`），再以 `compression: 'STORE'` 直接封裝並於 `manifest.json` 寫入 ISO `savedAt`。
    - 存檔、開檔、彙整合併與自動草稿恢復迴圈皆接入 `this.throwIfProcessingCancelled()`，支援進度條「取消」按鈕並於取消或異常時即時釋放中途建立之 Object URL。
    - 補齊 `saveProject`、`openProjectFromFile` 與 `mergeProjectFiles` 的 `finally { this.clearProgress(); }`，修正 `handleProjectError(error, action)` 區分「無法儲存／無法開啟／無法合併」，並自動關閉燈箱、重置 `activeFilter = 'all'`、等待 `await this.syncAllCompositeThumbnails()` 完成。
    - 新增 `sanitizeFilename()` 清洗案由中的 `[\\/:*?"<>|]` 等非法檔名字元；新增 `finalizeActiveInputsBeforeSave()` 於存檔、合併與排序前自動觸發輸入中欄位的 `blur()` 與歷史紀錄提交。
- **日常操作體驗強化、新增「依採證時間排序」與選單精簡**：
  - `index.html`、`css/app.css`、`vendor/tailwind.css`、`js/app.js`：
    - 頂部導覽列新增 `#btnQuickSaveStatus` 即時顯示琥珀色「還有修改沒存（Ctrl+S）・已自動暫存」或翠綠色「工作進度已儲存／已儲存：檔名」（對比度 > 8:1），點擊即可存檔；全域攔截 `Ctrl+S`（存檔）與 `Ctrl+Shift+S`（另存新檔）。
    - 「檔案 ▾」選單精簡為「開啟工作進度…」、「併入其他工作進度…」、「儲存工作進度（Ctrl+S）」與「另存新進度檔…」（移除重複的輕量存檔鈕、Excel 匯入／匯出與下載 ZIP）。
    - 「更多 ▾」選單新增 `#btnSortByCaptureTime`（「依採證時間排序」），實作 `getPhotoCaptureSortKeys()` 與 `sortPhotosByCaptureTime()`，依採證日期（空白時自動採用案件預設日期）與時間由早到晚排列，同時確保局部特寫子圖緊跟母圖且保留原編號格式；並從選單與畫面移除「分頁預覽」。
    - 擴充 `bindExternalFileDropTarget()` 支援拖放單一 `.photo-report` 開啟與 2 個以上 `.photo-report` 自動彙整；實作 IndexedDB `PhotoReportAutoDraftDB` 背景自動暫存與 `#autoDraftBanner` 恢復橫幅。
- **多 `.photo-report` 工作進度檔智慧彙整合併**：
  - `index.html`、`js/app.js`（`mergeProjects`、`mergeProjectFiles`、`mergeMatchedPhotoRecord`、`computeVisualFingerprintFromCanvas`、`computeVisualFingerprintFromBlob`、`compareVisualFingerprints`、`normalizePhotoFilenameKey`、`allocateMergedMainSeq`、`markProjectSaved`、`getPhotoContentSignature`）：
    - 實作雙重照片身分比對：優先以 `uid` 配對，未命中時再比對正規化檔名、長寬比（誤差 `<= 0.05`）與 256-bit dHash ＋ 2×2 象限 RGB 色彩 `visualFingerprint`，跨電腦或經過 2000px 重壓縮之相同照片可精準合併去重，同檔名不同畫面照片各自保留不誤合。
    - 依存檔時間由早到晚處理多份來源，並追蹤目前工作台各照片是否相較最近存檔異動過（`savedPhotosByUid`）：畫面上未存檔的局部修改優先保護，未動過的照片則依較晚時間公平更新。
    - 保護不同 `caseDate`／`defaultLocation` 案件之專屬照片顯示不互染，新增主照順延編號不撞號、子特寫緊跟母圖（依 `cropRect` 近似配對去重或新增 `-2`）；合併後清除 `currentProjectFileHandle` 防 `Ctrl+S` 誤蓋來源小檔並推入完整 `recordHistory('彙整工作進度檔')` 供 `Ctrl+Z` 一鍵退回。
- **測試與驗證擴充**：
  - `tests/e2e/photo-report.spec.js`：更新 `1A` 資訊架構檢查（含 `btnSortByCaptureTime`、`btnSaveProjectAs`、`mergeProjectInput` 與已移除項目確認），並於第 5 節（`5A`～`5G`）完整覆蓋狀態徽章、`Ctrl+S`、檔名非法字元清洗、時間遮罩自動定稿、原地覆寫 File Handle、預設 2000px 最佳化壓縮與標記保留、拖放單檔開啟、依採證時間排序、損毀檔案防漏 URL、IndexedDB 自動草稿，以及**多 `.photo-report` 進度檔跨機重壓縮去重／防誤合／一空一有補齊／較晚覆蓋／未存保護／子特寫合併／壞檔防護／`Ctrl+Z` 復原**全套 E2E 驗證。

## 異動檔案 (Changed Files)
index.html、css/app.css、js/app.js、vendor/tailwind.css、tests/e2e/photo-report.spec.js、HANDOFF.md。

## 刻意未修改 (Do Not Do / Deliberately Omitted)
Word／PDF／Excel 匯出結構、版本號與 Tauri 桌面版安裝檔打包（待使用者確認後處理）。

## 尚未完成 (Remaining Work)
- **P1 (阻斷/必須)**：無。
- **P2 (重要/當次)**：無。
- **P3 (改善建議/暫緩)**：視後續需求決定是否升版或執行 `js/app.js` 模組拆分。

## 驗證結果 (Validation)
### 已執行測試與結果
- `node --check js/app.js`：通過。
- `powershell -ExecutionPolicy Bypass -File scripts\qa.ps1`：通過（含 `git diff --check`）。
- `npm test`（`node tests/run-all-unit.js`）：4/4 通過。
- `npm run test:baseline`（`node tests/run-baseline-test.js`）：Word／Excel／PDF 結構與 3 種版型全部通過。
- `npm run test:e2e`（`node tests/e2e/photo-report.spec.js`）：全部通過（含第 `5A`～`5G` 節全套實測）。
### 尚未驗證項目
未重新升版或建置 Tauri 桌面版安裝檔。
### 已知風險 (Known Risks)
無本輪已知阻斷性問題。

## Git 狀態
- Commit：承接基準 `275211027f7c4b07ad95e0e710063f83f5434d95`（本輪修改尚未 Commit，待使用者確認）。
- Push：前一輪 `2752110` 已 Push；本輪修改待使用者指示後再 Commit 與 Push。
- Working Tree：有本輪未提交修改（`index.html`、`css/app.css`、`js/app.js`、`vendor/tailwind.css`、`tests/e2e/photo-report.spec.js`、`HANDOFF.md`）。
- Branch：main。

## 下一步建議動作 (Next Recommended Action)
請使用者實測本輪工作進度強化、依採證時間排序與多進度檔智慧彙整功能，確認無誤後執行 Commit 與 Push。

## 發布狀態 (Release Status)
可交付（Web 本機功能與全套測試通過；待使用者確認後提交推送）。
---

以下為先前交接紀錄，現況以本輪狀態與 Git 為準。

## 2026-10-22 v2.4.1 發布：拆出內嵌樣式與主程式（Claude）
- **發布結果**：使用者授權後，分支快轉合併回 `main`，提交 d9d6ec8 並推送；annotated tag `v2.4.1`、GitHub Release（含 `Photo-Report-Generator-2.4.1-Setup.exe` 2.37 MB、`-Portable.zip` 2.58 MB）已建立，下載回來雜湊與本機一致。GitHub Pages 建置完成（built d9d6ec8），正式站 `index.html`、`css/app.css`、`js/app.js`、`vendor/tailwind.css`、`version.txt` 與提交逐位元組一致（5/5）。
- **樣式驗證**：以「屬性名稱排序後」的計算後樣式指紋比對，拆分前 v2.4.0、拆分後 v2.4.1 本機、正式站三者皆為 `1b93b46f`（1366 寬）。注意：未排序時 CSS 變數的列舉順序會因來源不同而變，指紋不可跨來源比；新連接埠第一次載入時 Font Awesome 樣式可能尚未套用，要等載入完成再量。
- **桌面版**：以 Windows PowerShell 5.1 執行 `build-portable.ps1` 建置成功；免安裝版解壓到含中文與空白的路徑可啟動、8 秒後仍正常執行。**未確認桌面視窗內的畫面內容**（未擷取螢幕）。
- **另修**：`build-portable.ps1` 原本無 BOM，Windows PowerShell 5.1 解析失敗（提交 05e0bf0 補 BOM）。
- `package-lock.json` 的版本仍是 2.3.2（既有狀況，未更動）。
- 下方為拆分過程的原記錄。

### 拆分過程（原記錄）
- 依「有條件、分批、先補測試再拆」原則，兩步各一個提交、不夾帶功能修改：
  1. `c37fe94` 內嵌 `<style>` 630 行 → `css/app.css`（逐字，只去縮排）。
  2. `a28fb87` 內嵌 `<script>` 3,174 行 → `js/app.js`（原封不動，保留縮排）。
- `index.html` 4,581 → 775 行。
- 抓到並修正兩個只有重建／打包時才會出現的問題：
  - Tailwind 只掃描 `index.html`，搬走後會少產生規則（CSS 少 3 條；JS 少約 1.7KB）。`prepare-assets.ps1` 掃描範圍加入兩個新檔後，產物與拆分前**逐位元組相同**。
  - 桌面版 `prepare-web.ps1` 原本不複製 `css/`，已加入；實跑確認 `web/` 與來源一致。
- 驗證：單元 4/4、Golden Baseline（Word／Excel／PDF）、Playwright E2E（file:// 開啟）、`qa.ps1` 全部通過；真實瀏覽器以「每個元素的計算後樣式」雜湊比對，1366 與 390 寬度皆與拆分前完全相同；主程式正常啟動、主控台零錯誤。
- 未做：未合併回 `main`、未推送、未升版；未實際執行 Tauri 打包（`cargo` 建置），只驗證了它的前置腳本 `prepare-web.ps1`。合併後建議打包一次桌面版並實際開啟確認。
- 下一步可選：`js/app.js`（3,174 行）之後要大改某區塊時，再依 AGENTS.md 第 4 節逐步拆出模組。

## 核心元資料 (Metadata)
- **Repository**：lianghao02/Photo-Report-Generator
- **Branch**：main
- **Commit SHA**：d9f8b8bc（本輪提交前基準；最新提交以 Git 記錄為準）
- **Skill Version**：v1.0.0
- **Task Type**：HANDOFF
- **Local Path Hint**：04_Photo-Report-Generator

---

## 2026-10-21 發布 v2.4.0（Claude）
- 依使用者指示升版 2.3.2 → 2.4.0：`version.txt`、`package.json`、`src-tauri/tauri.conf.json`、`src-tauri/Cargo.toml`／`Cargo.lock`（僅 `photo-report-generator` 套件）、`index.html` 內 4 處版本字串、README、CHANGELOG（兩個未發布區塊併入 v2.4.0）。
- 以 `scripts/build-portable.ps1` 建置 Setup.exe 與 Portable.zip；Release 附件為這兩個檔案。
- 推送 main 會同時觸發 GitHub Pages 更新線上版。

## 2026-10-21 縮圖卡片等高與雙擊放大（Claude）
- **起因**：使用者反映卡片上下參差、小圖看不到「放大」按鈕、文字不夠清楚。根因：`.thumbnail-stage` 的 `aspect-ratio` 取自照片比例，扁長圖框高只剩約 10 px，按鈕被 `overflow-hidden` 切掉。
- **變更**：`getPhotoDisplayLayout` 固定回傳 `safeAspect: 1`（旋轉只需 `rotate`，因為正方形旋轉後不變）；移除 `.thumbnail-zoom-trigger`／`.btn-zoom-preview`；縮圖框 `ondblclick` 開燈箱（`stopPropagation` 以免觸發編輯）；說明列雙擊才編輯；文字 12 px、提示色 `#4B5563`；預設縮圖 150 px。
- **測試**：e2e 新增「所有縮圖框同大小」「單擊不開燈箱、雙擊開」，編輯展開改為雙擊說明列。`npm run test:all` 全過；已重建 `vendor/tailwind.css`。以寬扁（1200×70、1800×90）、細長（420×1300）等混合照片截圖確認等高。
- **未做**：未升版、未重建 Tauri 安裝檔、未推送。

## 2026-10-21 方案 B 畫面重整與文字簡化（Claude）
- **起因**：使用者反映畫面太亂、太雜。分析後與使用者確認改用方案 B（單一工作台重整）。使用者表示以下功能都很少用：Excel 匯入／匯出、照片 ZIP、解析檔名時間、完整度篩選、分頁預覽、開啟／儲存專案檔。
- **變更內容**：
  - 頂部改為「檔案」與「說明」兩個 `details.app-menu` 選單。
  - 左欄只留案件資料；原本的「套用」改成整行按鈕「把日期與地點寫進已選照片」。
  - 照片工具列合併為一排，加上「更多」選單。
  - 新增 `#exportBar`：版型＋匯出 Word／PDF，sticky 固定在底部。
  - 漏填檢查列改為預設收合：新增 `#btnToggleAudit`、`applyAuditPanelState()`。
  - `setupMenus()`：開一個選單時關閉其他選單，點選單外或按 Esc 會收起。
- **相容性**：所有既有元素 ID 與 JS 行為保留，只刪除未使用的 `stageImportBar`；`recordHistory` 的內部標籤未更動。
- **偏好設定預設值**：編輯欄預設顯示（`editorCollapsed` 預設 false）；左欄預設寬 300px，最小 280px。
- **文字**：標籤、提示、alert／confirm、快速鍵說明、匯出提醒、更新視窗全面改為白話。
- **測試同步**：`tests/e2e/photo-report.spec.js` 改為驗證新的資訊架構，並把「已開啟專案」改為「已開啟工作進度」。
- **驗證**：
  - `npm run test:all` 全部通過：單元測試 4/4、Playwright Web UI、Golden Baseline。Word／Excel／PDF 匯出結構與修改前一致。
  - `scripts/qa.ps1` 通過。
  - 以 Playwright 截圖檢查 1366×768、1920×1040、1100×760：
    - 選單在點選單外後會正確收起。
    - 可見文字對比掃描零項低於 4.5:1。
- **未做**：沒有升版、沒有重建 Tauri 安裝檔、沒有推送。

## 2026-10-20 2D Q 版介面接手收尾（Claude）
- 前一輪（其他 Agent，已結束）未提交的 Q 版改版為 `index.html` 樣式（+298／−171），本輪已驗收並一併提交。
- 修正：
  - `vendor/tailwind.css` 未重建，新增的任意值樣式實際不存在；已以 `scripts/prepare-assets.ps1` 重建。
  - 主要藍色按鈕白字對比 3.7:1，改為 `#2563EB`（5.2:1）。
  - 紅色危險文字改為 `red-700`。
  - 停用按鈕不透明度從 0.45 調整為 0.62。
- 驗證：
  - `scripts/qa.ps1` 通過。
  - `npm run test:all` 通過：單元測試 4/4、Playwright Web UI、匯出 Golden Baseline。
  - 在本機以 `http.server` 開啟頁面，載入 `tests/fixtures` 照片後，檢查工作台與試算表兩種視圖：可見文字對比掃描零項低於 4.5:1。
- 未做：沒有升版、沒有重建 Tauri 安裝檔、沒有推送。

## 目前狀態
本輪已發布 v2.3.2 空白工作台操作密度修訂與桌面產物；以下待驗證事項維持。舊交接原文保留於下方，屬歷史，不代表本輪 Git 或測試狀態。

## 本輪目標
依已授權目錄配置與舊產物清理要求，保留現有功能。

## 基準與已確認事實 (Baseline & Confirmed Facts)
中央 docs/project-layout/baseline.json、operations.json 保存本輪基準，HEAD/分支保持，前輪功能成果繼承。

## 已完成 (Completed)
2026-10-08 UI/UX 微調：空白狀態隱藏僅適用於已載入照片的批次工具列；保留檢視模式與縮圖大小設定。版型微縮圖新增可見「預覽」標示，並在沒有照片時停用單張編輯。未變更照片載入後的工具列行為、匯出邏輯或資料模型。

2026-10-08 v2.3.2 發布完成：版本號已同步至 Web、Node 與 Tauri 設定；`Photo-Report-Generator-2.3.2-Setup.exe` 與 `Photo-Report-Generator-2.3.2-Portable.zip` 已建置並完成 SHA-256 與 ZIP 內容核對。正式 Release：https://github.com/lianghao02/Photo-Report-Generator/releases/tag/v2.3.2

2026-10-06 GitHub 同步交接：使用者已授權提交與推送前輪成果；本輪只提交已核對範圍。最新 Commit SHA、遠端同步與 CI 結果統一見控制中心 `docs/github-sync/RESULTS.md`，不將提交本身的 SHA 寫入同一份提交。

2026-10-05 README 文件更新：補齊專案概念、開發原因、典型流程、已知 Bug／限制及回報方式，並依實際入口校正必要操作說明。本次沒有修改產品程式、環境或個人資料，未 Commit／Push；前輪成果與既有待辦繼承。文件檢核與逐案索引由控制中心 docs/readme-refresh/RESULTS.md 彙整，不代表本次重新驗收全部功能。

刪除 4 張現行測試未引用的 test_photos 舊合成照片，更新 README。保留 v2.3.1 功能、tests/fixtures、Golden Baseline、vendor 與 Tauri/Web 配置。

## 異動檔案 (Changed Files)
上述明確項目與本交接；詳細清冊見中央 docs/project-layout/RESULTS.md。

## 刻意未修改 (Do Not Do / Deliberately Omitted)
業務演算法、現行環境、模型、有效測試素材及使用者原始資料未動；不覆寫未知修改，舊交接內容完整保留。

## 尚未完成 (Remaining Work)
- **P1 (阻斷/必須)**：無本輪整理阻斷。
- **P2 (重要/當次)**：無本輪未完成事項；前輪介面改善建議仍另案。
- **P3 (改善建議/暫緩)**：未因整理擴大重構；正式發布另依 release-gate 驗證。

## 驗證結果 (Validation)
### 已執行測試與結果
四組單元測試、Playwright 上傳/去重/燈箱/多選及 Word/Excel/PDF Golden Baseline 通過。

2026-10-08：`scripts/qa.ps1`、`npm run test:all`（單元、Playwright、Golden Baseline）均通過；另以本機預覽頁確認空白狀態工具列、版型預覽標示與單張編輯停用狀態。PowerShell 7 的 `scripts/build-portable.ps1` 建置成功，且已核對 Portable ZIP 內含執行檔、`WebView2Loader.dll` 與說明檔。
### 尚未驗證項目
未重新驗收全部原生功能或其他電腦/Windows 10 發布環境。
### 已知風險 (Known Risks)
保留上述既有待辦與驗證邊界，不把清理宣稱為其修復。

## Git 狀態
- Commit：上述 SHA 為提交前基準；最新 SHA 見 `git log -1` 與中央同步報告。
- Push：實際推送及遠端核對結果見中央 `docs/github-sync/RESULTS.md`。
- Working Tree：最終狀態見中央同步報告；不含被忽略的環境、成品與使用者資料。
- Working Tree（2026-10-08）：v2.3.2 程式、版本與文件已提交、推送；本交接狀態更新待提交。
- Branch：main。

## 下一步建議動作 (Next Recommended Action)
本輪停止擴大修改；日後提交前核對工作範圍並另取得授權。

## 發布狀態 (Release Status)
v2.3.2 已發布，含 Windows 安裝版與免安裝綠色版。

---

## 承接的前輪交接（原文保留，屬歷史）


> 2026-10-05 環境修復交接：本輪僅修正 AGENTS.md 的共用 Skill 正式來源為 configs/skills，程式碼與既有環境不變。Working Tree 為 Modified，未 Commit／Push；下列發布與功能紀錄為承接的前輪成果。

# HANDOFF

## 核心元資料 (Metadata)
- **Repository**：lianghao02/Photo-Report-Generator
- **Branch**：main
- **最新 Commit**：以 Repository 即時狀態為準
- **Skill Version**：v1.0.0
- **Task Type**：MAINTENANCE-CONVERGENCE

> **目前實際 Git 狀態（2026-10-02）**：已完成 Phase A/B 收斂與正式發布 v2.3.1；所有既有單元測試、E2E 測試、Golden Baseline 及桌面版建置全數通過。

## 目前狀態
可交付／Stable Maintenance (v2.3.1)

## 本輪目標
1. **Phase A：發布／版本／產品表面一致性收斂**：
   - 清除 Footer 殘留之舊版 v2.0 與內部「全域開發憲法 v8.1」治理字樣。
   - 清理更新 Modal 中的寫死 placeholder（`v2.1.3` $\rightarrow$ `--`）。
   - 修正建置腳本 `scripts/build-portable.ps1`，確保 Release Artifact 產物採用穩定的 ASCII 命名格式（`Photo-Report-Generator-<version>-Setup.exe` 與 `Photo-Report-Generator-<version>-Portable.zip`），不受語系編碼干擾。
   - 在 `CHANGELOG.md` 新增 `Unreleased` 段落，補登 9/13 於 main 已完成但尚未發布之改進項目。
2. **Phase B：文件治理與歷史資產歸檔**：
   - 從 `CHANGELOG.md` 移除未立項之空想性 Roadmap（GPS 反查、OCR、DOCX 逆向解析等）。
   - 建立 `docs/history/`，將已結案之歷史規劃文件歸檔（`IMPLEMENTATION_PLAN.md`、`MODULARIZATION_PLAN.md`、`REGRESSION_BASELINE.md`）。
   - 收斂根目錄治理結構，杜絕歷史待辦干擾 Agent 上下文。

## 已完成
1. **產品表面與版本一致性收斂 (Phase A)**：
   - `index.html` 頁尾版權宣告收斂為乾淨的中立產品資訊：`© [年份] LiangHao｜現場照片清冊生成工具・純本機離線處理`，徹底移除過期版本號與內部規範字眼。
   - `index.html` 更新對話框中 `#updateLatestVer` 的預設文字修正為中立佔位符 `--`，版本完全交由 API 非同步獲取寫入。
   - `scripts/build-portable.ps1` 產物命名重構為基於版本來源動態產生之標準 ASCII 格式：
     - 安裝檔：`Photo-Report-Generator-2.3.0-Setup.exe`（本機實機建置驗證通過，2.59 MB）
     - 可攜版：`Photo-Report-Generator-2.3.0-Portable.zip`（本機實機建置驗證通過，2.83 MB）
   - `CHANGELOG.md` 頂部新增 `## Unreleased`，忠實記錄 UI/UX 工具列重整、雙層舞台大圖燈箱及多選拖曳衝突修正。
2. **文件治理與歷史資產歸檔 (Phase B)**：
   - 移除 `CHANGELOG.md` 內之 `## 🔮 下一版本預計規劃 (Roadmap v2.4.0)`，消除過度承諾與技術偏離。
   - 透過 `git mv` 完整保留歷程，歸檔歷史文件至 `docs/history/`：
     - `IMPLEMENTATION_PLAN.md` $\rightarrow$ `docs/history/implementation-2026-08.md`
     - `MODULARIZATION_PLAN.md` $\rightarrow$ `docs/history/modularization-roadmap.md`
     - `REGRESSION_BASELINE.md` $\rightarrow$ `docs/history/regression-baseline.md`
   - 根目錄維持極簡治理清單（`README.md`、`AGENTS.md`、`HANDOFF.md`、`CHANGELOG.md`、`version.txt`）。

## 刻意未修改
- **未動核心業務邏輯**：未修改 Word、PDF、Excel、ZIP 匯出器與資料模型。
- **未動照片互動演算法**：未修改多選、拖曳排序、畫布縮放或快捷鍵核心邏輯。
- **未改動既有 v2.3.0 GitHub Release**：本輪僅修正未來建置檔名規則，不變更遠端既有 Release 與 Tag。
- **刻意延後項目**：Contextual Toolbar、10px/11px 字體優化、響應式 Smoke Test、`photo-grid-ui.js` 抽離均留待獨立專案階段評估。

## 尚未完成
無阻斷性待辦事項。

## 驗證結果
### 已執行
- **靜態檢查**：
  - `pwsh -NoProfile -File ./scripts/qa.ps1`：PASS（共用 QA 通過）
  - `git diff --check`：PASS（零格式或空白異常）
- **單元測試**：
  - `npm test`：PASS (4/4 全部斷言通過，validation, audit, history, selection)
- **匯出結構比對 (Golden Baseline)**：
  - `npm run test:baseline`：PASS（Word XML 結構、Excel 欄位、PDF 版面尺寸與三大版型 100% 符合）
- **E2E 自動化測試 (Playwright)**：
  - `npm run test:e2e`：PASS（完整度稽核、雙擊卡片、大圖燈箱、Ctrl/Shift 多選全部驗證通過）
- **建置驗證 (Tauri Desktop)**：
  - `pwsh -NoProfile -File ./scripts/build-portable.ps1`：PASS
  - 產出 `Photo-Report-Generator-2.3.0-Setup.exe` (2.59 MB)
  - 產出 `Photo-Report-Generator-2.3.0-Portable.zip` (2.83 MB)
- **回歸測試**：
  - 照片載入、多選、拖曳排序、大圖燈箱、Undo/Redo、三種版型匯出全數維持正常。

### 尚未驗證
無。

### 已知風險
無新增風險。

## Git 狀態
- Branch：main
- Push：已完成
- Working Tree：Clean
- 最新 Commit：以 Repository 即時狀態為準

## 下一步
僅在有已重現 Bug、實際業務需求或準備下一次正式 Release 時再開新任務。
