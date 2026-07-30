# 📷 現場照片清冊自動套印工具 (v1.1)

[![Version](https://img.shields.io/badge/version-v1.1-blue.svg)](https://github.com/lianghao02/Photo-Report-Generator)
[![VBA](https://img.shields.io/badge/Engine-VBA%20Late%20Binding-green.svg)](https://microsoft.com)

## 🏆 v1.1 里程碑：Word/Excel 現場照片自動排版套印

## 📖 重大更新摘要 (Summary)

本版本為警務與現場勘查照片自動化排版工具之穩定發行版本，採用 VBA Late Binding 技術與 Word/Excel 自動化物件溝通。

傳統勘查員在整理數百張現場照片時，必須手動貼入 Word 並逐一調整圖片尺寸、對齊表格與輸入備註，處理一份報告往往耗費半天以上時間。本工具透過 VBA 腳本可在 **5 秒內** 自動讀取資料夾中所有相片，依標準規格壓縮、對齊並自動填入表格，極致提升公務處理效率。

## ✨ 重點更新特色

- 📄 **VBA Late Binding 跨版本相容引擎 (CreateObject Architecture)**：
  - 使用 `CreateObject("Word.Application")` 動態綁定，擺脫 Reference 遺失隱患。
  - 確保在 Office 2016 至 Office 365 各版本間皆能 100% 穩定執行。

- 🖼️ **相片比例自動鎖定與表格對齊 (Auto-Resizing)**：
  - 智慧計算表格儲存格邊界，自動將圖片依比例縮放並居中對齊。
  - 杜絕圖片變形或表格爆頁痛點，產出軍規級標準勘查報告。