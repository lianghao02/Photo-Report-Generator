# HANDOFF

## 目前狀態
可交付

## 本輪目標
修正 ec8ef93 後續已確認之 4 個局部缺陷：
1. 日期／時間完整度驗證邏輯（依專案既有民國日期與時間範圍精確檢驗，非單純長度檢查）
2. 「疑似重複」與實際同檔名演算法一致，文案改為「同名照片」
3. 移除 audit filter 按鈕之 HTML inline onclick 雙重事件綁定
4. 修正 exportAuditTitle / exportAuditModalTitle ID 查詢一致性與標題更新

## 已完成
1. 在 `index.html` 新增 `isValidMinguoDate(raw)`：精準驗證民國年月日 7 碼、大小月、民國閏年 2 月天數合法性。
2. 在 `index.html` 新增 `isValidTimeFormat(raw)`：精準驗證 HH:MM 或 HH:MM:SS（含純數字 4/6 碼）之時分秒數值合法範圍（0-23、0-59、0-59）。
3. 升級 `auditPhotosCompleteness()` 與 `getVisiblePhotoIndices()` 內 `invalidDateTime` 判斷邏輯，改用 `isValidMinguoDate` 與 `isValidTimeFormat`，不合法即視為待確認。
4. 移除 `auditFilterBar` 內全部 5 個篩選按鈕之 HTML inline `onclick`，統一由 `initEvents()` 中的 `addEventListener` 處理，徹底消除單次點擊雙重觸發。
5. 將 `duplicatePhotos` 相關之按鈕文字、篩選提示文字與匯出提示列表文案由「疑似重複」統一修正為「同名照片」，精準符合依檔名比對之實際邏輯。
6. 在 `initElements()` 快取 `this.exportAuditModalTitle = document.getElementById('exportAuditModalTitle')`，並在 `confirmExportWithAudit()` 中正確更新匯出提示彈窗標題。
7. 執行 `scripts/prepare-web.ps1` 同步更新離線資源與 `web/index.html`。
8. 執行 `scripts/qa.ps1` 差異與機敏資料檢查通過。

## 刻意未修改
- Word / PDF / Excel 匯出核心邏輯（`exportDocx`、`exportPdf`、`exportExcel`）完全未動
- Word / PDF / Excel 正式版型與尺寸完全未動
- Undo / Redo 資料歷史簽名邏輯保持安全
- Tauri 雙模式架構與設定完全未動
- 演算法邊界：未引入圖片 Hash，維持依檔名比對

## 尚未完成
- 無（4 項已確認缺陷均已修復）

## 驗證結果
### 已執行
1. JavaScript 語法檢查（Node.js `new Function()` 檢驗 6 個 script 區塊）：全部通過（Block 0-5: OK）。
2. 日期與時間驗證單元測試（Node.js）：
   - `113/08/26`、`1130826` -> 合法
   - `113/02/29` (民國113年為西元2024閏年) -> 合法
   - `114/02/29` (民國114年為西元2025平年) -> 正確判定非法
   - `113/13/01` (月份非法) -> 正確判定非法
   - `113/04/31` (小月31日) -> 正確判定非法
   - `14:30`、`14:30:15`、`1430`、`143015` -> 合法
   - `24:00`、`12:60`、`12:30:60` -> 正確判定非法
3. `scripts/prepare-web.ps1`：執行成功（Done in 475ms，Local frontend assets updated）。
4. `scripts/qa.ps1`：共用 QA 通過（exit code 0）。
5. `git diff` 審查：確認僅修改 `index.html` 相關之 4 項缺陷，exporter 模組無任何異動。

### 尚未驗證
- 實際於 WebView2 桌面封裝程式執行點擊（已於純網頁環境完成語法與 QA 檢驗）

### 已知風險
- 無阻斷性風險

## Git 狀態
- Commit：未提交（待交付提交）
- Push：否
- Working Tree：Modified (index.html, HANDOFF.md)
- Branch：main

## 下一步
執行 commit 與 push。