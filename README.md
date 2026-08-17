# 📷 現場照片清冊自動套印工具 (v1.1.2)

[![Version](https://img.shields.io/badge/version-v1.1.2-blue.svg)](https://github.com/lianghao02/Photo-Report-Generator)
[![VBA](https://img.shields.io/badge/Engine-VBA%20Late%20Binding-green.svg)](https://microsoft.com)

## 下載、依賴與執行

- **必要軟體**：Windows 桌面版 Microsoft Excel 與 Word；不需要 Python 或 Node.js。
- **推薦下載**：從 GitHub 下載 ZIP，或使用 `downloads/Photo Report.rar` 發行封裝；解壓後不可拆散 Excel 巨集檔與 Word 範本。
- **執行入口**：開啟 `vba_src/清冊編輯.xlsm`，確認來源可信後依 Excel 提示啟用巨集，再選擇照片與版型。
- **範本**：`vba_src/左右兩張.docx` 與 `vba_src/上下兩張.docx` 必須和巨集檔維持相對位置。
- **打包／移機**：將上述 `.xlsm` 與兩份 `.docx` 一起壓縮即可；本專案沒有程式建置步驟。
- **網站展示**：`index.html` 僅為介紹頁，會從 CDN 載入 Tailwind CSS 與 Google Fonts，不是照片清冊執行入口。

## 🏆 v1.1 里程碑：Word/Excel 現場照片自動排版套印

## 📖 重大更新摘要 (Summary)

本版本提供警務與現場勘查照片的 Word／Excel 自動排版範本，採用 VBA Late Binding 呼叫本機 Office。

使用者可依既有 Word 範本與 Excel 巨集匯入照片、依比例縮放並放入表格。實際速度與相容性取決於照片數量、Office 版本、巨集安全性設定及電腦效能。

## 使用方式與限制

1. 從 `vba_src/` 取得 Excel 巨集檔與對應 Word 範本。
2. 先備份原始照片與範本，再於受信任的本機環境啟用巨集。
3. 確認 Word／Excel 輸出內容、照片順序及備註後再定稿。

本工具需要桌面版 Microsoft Office。巨集檔不可在來源不明或未確認內容時直接啟用。

## ✨ 重點更新特色

- 📄 **VBA Late Binding 跨版本相容引擎 (CreateObject Architecture)**：
  - 使用 `CreateObject("Word.Application")` 動態綁定，擺脫 Reference 遺失隱患。
  - 降低 Office 版本差異造成的參照遺失，但仍應在實際 Office 版本與巨集政策下測試。

- 🖼️ **相片比例自動鎖定與表格對齊 (Auto-Resizing)**：
  - 智慧計算表格儲存格邊界，自動將圖片依比例縮放並居中對齊。
  - 降低圖片變形或表格溢位，輸出後仍需人工確認版面與照片順序。
