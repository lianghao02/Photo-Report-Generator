# 📷 現場照片清冊自動套印工具 (Photo Report Generator) v2.1.0

[![Version](https://img.shields.io/badge/version-v2.1.0-blue.svg)](https://github.com/lianghao02/Photo-Report-Generator)
[![Platform](https://img.shields.io/badge/Platform-Pure%20Web%20SPA%20%2F%20Tauri-emerald.svg)](https://lianghao02.github.io/Photo-Report-Generator/)
[![Security](https://img.shields.io/badge/Security-Zero%20Macro%20Warning-brightgreen.svg)](https://microsoft.com)

> **版本**：v2.1.0 (純前端現代化 SPA 工作台．Tauri v2 輕量桌面版．三大經典版型支援)
> **線上即用網址**：[https://lianghao02.github.io/Photo-Report-Generator/](https://lianghao02.github.io/Photo-Report-Generator/)
> **維護者**：LiangHao (梁巡官)

## 技術架構現況（2026-08-24）

主力版已是 **HTML5／CSS／JavaScript 純前端 SPA**，並已支援 **Tauri v2 輕量桌面版**，徹底取代舊 VBA 巨集與 Microsoft Office 執行依賴。專案支援 GitHub Pages 瀏覽器即開即用，以及 3.08 MB 單檔桌面執行檔。

---

## 📖 專案簡介

本專案專為警務同仁、刑事偵查與現場勘查人員打造之**現場照片清冊線上生成工具**。

v2.0/v2.1 版本迎來重大架構革新：**徹底淘汰舊有 VBA 巨集與本機 Office 依賴**，升級為純前端單頁 Web 應用程式 (SPA) 與 Tauri 桌面版。使用者只需使用瀏覽器打開網頁或雙擊桌面版，即可批次拖曳照片、智慧填寫時間與說明，並直接在記憶體中生成符合公務標準的 **原生 Word (`.docx`) 清冊**，100% 零巨集資安警示！

---

## ✨ v2.0 核心特色與創新

### 1. 📄 內建三大經典 Word 版型（100% 精準對齊原公務範本）
- 🔼 **上下兩張 (A4 直式．經典版型)**：每頁 2 個獨立表格，大圖呈現（限寬 14cm / 限高 9cm），配備案由、日期、時間、編號、地點、製作人與詳細說明。
- ⏸️ **左右兩張 (A4 直式雙欄．對比版型)**：每頁 1 個整合表格，頂部全寬案由與蒐證資訊，左右並排 2 張照片（限寬 8cm / 限高 18cm）與各自獨立說明。
- ⏹️ **橫式三張 (A4 橫式三欄．廣角版型)**：每頁 1 個整合表格，橫向並排 3 張照片與詳細說明，適合全景連續畫面。

### 2. 🕒 專為「監視器／通訊截圖／蒐證」打造的四大時間輸入模式
- ⚡ **從檔名智慧解析時間**：一鍵正則提取檔名中的日期時間（如 `YYYYMMDD_HHMMSS`、`Screenshot_...`、`LINE_...`），自動轉為民國年或時分秒。
- ⏱️ **基準時間 + 步進自動遞增**：設定首張監視器起始時間（如 `14:32:00`）與遞增間隔（如 `+5秒` 或 `+1分鐘`），全數照片自動推算。
- ⬇️ **一鍵向下填滿時間**：同一個蒐證時段只需輸入第一筆，一鍵套用至下方所有照片。
- 📋 **剪貼簿多行貼上**：複製外部 Excel 或記事本的多行時間文字，一鍵依序對齊填入各照片。

### 3. 🔢 智慧流水號與 Excel 級向下填滿
- 支援設定前綴（如 `照片-`、`證物 `）、起始編號與補零位數（`01`、`001`），一鍵全自動編號。
- 採證地點、製作人、跡證說明均配備 **⬇️ 向下填滿按鈕**，輸入一次即可批次套用。

### 4. 🖼️ Canvas 影像前處理與列印品質最佳化
- **EXIF 自動轉正**：自動解析手機直拍之 EXIF Orientation 標記，自動修正旋轉角度，亦提供手動 90° 旋轉按鈕。
- **等比列印壓縮**：前端自動等比縮圖至適合 A4 列印之最適解析度（最大 1600px，JPEG 85% 品質），將清冊檔案控制在輕盈的 3MB~8MB，徹底杜絕 Word 檔案過大膨脹。

### 5. 📊 雙向相容與全方位匯出
- 📄 **匯出 Word 清冊 (`.docx`)**：純前端生成原生標準 OpenXML 文件，各版本 Word / WPS / LibreOffice 通用且無巨集警告。
- 📊 **匯出 Excel 清單 (`.xlsx`)**：輸出結構化資料表格。
- 📥 **匯入 Excel 說明檔 (`.xlsx`)**：支援將外部編輯好的 Excel 檔案匯入，自動依編號或檔名對齊填回說明。
- 🗜️ **匯出最佳化照片 ZIP (`.zip`)**：打包經旋轉校正與壓縮後的清晰照片。

---

## 🚀 使用方式

### 線上即用（推薦）
直接點擊瀏覽器開啟：[https://lianghao02.github.io/Photo-Report-Generator/](https://lianghao02.github.io/Photo-Report-Generator/)

### 離線單機使用
1. 下載專案根目錄下的 `index.html`。
2. 直接雙擊 `index.html` 於瀏覽器開啟即可開始使用，完全不需要安裝 Python、Node.js 或 Office 巨集。

---

## 🗂️ 專案目錄結構

```text
04_Photo-Report-Generator/
├── .nojekyll
├── CHANGELOG.md             # 版本更新紀錄
├── README.md                # 專案說明文件
├── index.html               # 核心應用程式 (v2.0 純前端 SPA 工作台)
├── legacy_vba/              # 舊版 VBA 歷史封存目錄 (清冊編輯VBA-0820.xlsm、三大 docx 範本)
└── scripts/
    └── qa.ps1               # 專案品質與敏感資料檢核腳本
```

---

## 📜 歷史版本說明

舊版基於 Excel VBA 之巨集工具（`清冊編輯VBA-0820.xlsm`）與原始範本已完整封存於 `legacy_vba/` 資料夾內，僅供歷史查閱，新開發與實務作業全面推薦使用 `index.html` 網頁版。
