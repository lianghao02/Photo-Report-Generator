# 📷 現場照片清冊自動套印工具 (v1.1.1)

[![Version](https://img.shields.io/badge/version-v1.1.1-blue.svg)](https://github.com/lianghao02/Photo-Report-Generator)
[![VBA](https://img.shields.io/badge/Engine-VBA%20Late%20Binding-green.svg)](https://microsoft.com)

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
  - 確保在 Office 2016 至 Office 365 各版本間皆能 100% 穩定執行。

- 🖼️ **相片比例自動鎖定與表格對齊 (Auto-Resizing)**：
  - 智慧計算表格儲存格邊界，自動將圖片依比例縮放並居中對齊。
  - 杜絕圖片變形或表格爆頁痛點，產出軍規級標準勘查報告。
