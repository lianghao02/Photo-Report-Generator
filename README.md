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

### 大量照片模式
載入時會以低解析縮圖預覽，原圖只在 Word 或 ZIP 匯出時處理，適合數百張照片的清冊。超過 500 張或合計 2GB 時，系統會先顯示確認提示；超大批次仍建議依案情分冊匯出。

### 工作台選取與排序
- 在照片畫布按住 `Ctrl` 後滾動滑鼠滾輪，可將縮圖尺寸連續調整為 90–400px；一般滾輪仍用於捲動。
- `Ctrl` 點選可多選、`Shift` 點選可連續選取，也可從畫布空白處拖曳框選；選取後直接拖曳任一卡片即可整組調整順序。
- `Ctrl + ← / →` 可讓已選照片前後移動一格，`Home / End` 可直接移至首尾。
- 在右側修改時間、地點或說明後，按「將已填欄位套用至已選照片」才會批次寫入；勾選「空白欄位也覆寫」可清除既有資料。

### 線上即用（推薦）
直接點擊瀏覽器開啟：[https://lianghao02.github.io/Photo-Report-Generator/](https://lianghao02.github.io/Photo-Report-Generator/)

### 離線單機使用
1. 下載發行版的 Portable ZIP，或完整下載專案中的 `index.html` 與 `vendor/` 目錄。
2. 保持兩者相對路徑不變後直接開啟 `index.html`，即可在無網路環境使用；完全不需要安裝 Python、Node.js 或 Office 巨集。

---

## 🗂️ 專案目錄結構

```text
04_Photo-Report-Generator/
├── .nojekyll                # 避免 GitHub Pages 略過下底線或 vendor 資源
├── CHANGELOG.md             # 版本更新歷程紀錄
├── README.md                # 專案說明文件
├── index.html               # 核心應用程式 (v2.1 純前端 SPA 工作台)
├── vendor/                  # 100% 離線第三方函式庫 (Tailwind CSS, FontAwesome, docx.js, JSZip 等)
├── test_photos/             # 現場照片測試資料集 (提供即時功能與排版驗證)
├── src-tauri/               # Tauri v2 輕量桌面應用程式設定與 Rust 核心
└── scripts/
    ├── build-portable.ps1   # 一鍵桌面版打包腳本 (產出 NSIS 安裝包 + Portable ZIP)
    ├── prepare-web.ps1      # 桌面版前端資源同步腳本 (自動整合 index.html 與 vendor/)
    ├── prepare-assets.ps1   # 離線靜態資源產製與驗證腳本
    ├── tailwind-input.css   # 離線樣式來源定義
    └── qa.ps1               # 專案品質與敏感資料檢核腳本
```

---

## 📜 歷史版本與現代化架構演進

本專案經歷了完整的公務自動化技術演進：

1. **早期舊版（Excel VBA 巨集）**：
   - 早期依賴 `清冊編輯.xlsm` 搭配本機 Office COM 物件進行巨集套印。
   - 缺點：易觸發公務電腦巨集資安警告、跨電腦版面易跑位、受限於 Office 版本安裝。已於 v2.0 全面除役退場。
2. **評估方案（Python 自動化腳本）**：
   - 曾評估以 Python (`python-docx` / `PySide6` / `PyInstaller`) 實作。
   - 缺點：打包後體積龐大（通常達 40MB~80MB）、在無 Python 環境或內網封閉電腦常遇相依性地獄。
3. **現行正式版（純前端 Web SPA + Tauri v2 桌面應用程式）**：
   - **首選 Web**：直接採用純前端技術（HTML5 + Canvas + `docx.js` 記憶體即時組裝 OpenXML），瀏覽器開箱即用、零環境安裝負擔。
   - **首選桌面**：以 **Tauri v2** 直接複用 Windows 內建 Edge WebView2 核心，打造出僅 **1.15 MB** 的安裝包與 **3.08 MB** 的免安裝單檔桌面程式，完美兼顧原生桌面手感與極致分發效率。
