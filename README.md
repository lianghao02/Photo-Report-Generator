# 📷 現場照片清冊自動套印工具 (Photo Report Generator) v2.1.2

[![Version](https://img.shields.io/badge/version-v2.1.2-blue.svg)](https://github.com/lianghao02/Photo-Report-Generator)
[![Platform](https://img.shields.io/badge/Platform-Pure%20Web%20SPA%20%2F%20Tauri%20Desktop-emerald.svg)](https://lianghao02.github.io/Photo-Report-Generator/)
[![Security](https://img.shields.io/badge/Security-Zero%20Macro%20Warning-brightgreen.svg)](https://microsoft.com)
[![License](https://img.shields.io/badge/License-MIT-gray.svg)](LICENSE)

> **版本**：v2.1.2 (純前端現代化 SPA 工作台．Tauri v2 輕量桌面版．三大經典版型支援)<br>
> **線上即用網址**：[https://lianghao02.github.io/Photo-Report-Generator/](https://lianghao02.github.io/Photo-Report-Generator/)<br>
> **最新釋出版下載**：[GitHub Releases 最新版本](https://github.com/lianghao02/Photo-Report-Generator/releases/latest)<br>
> **維護者**：LiangHao (梁巡官)

---

## 📥 該下載哪一個？（版本選擇指引）

依您的使用環境與公務需求選擇最適版本：

| 版本類型 | 適用情境 | 檔案大小 | 安裝與使用方式 |
| :--- | :--- | :---: | :--- |
| 🌐 **線上網頁版** | 有網際網路連線、臨時或非公務電腦 | **0 MB** | 直接點擊 [https://lianghao02.github.io/Photo-Report-Generator/](https://lianghao02.github.io/Photo-Report-Generator/) 即開即用。 |
| 📦 **免安裝綠色版 (ZIP)**<br>*(推薦公務同仁)* | 封閉內網、公務電腦無管理員權限、放置隨身碟隨插即用 | **~2.8 MB** | 1. 下載 `照片清冊產生器_x.x.x_x64_portable.zip`<br>2. **解壓縮至任意資料夾**<br>3. 雙擊 `照片清冊產生器.exe` 即可直接啟動。 |
| 💿 **Windows 安裝版 (EXE)** | 個人專用公務電腦、需開始功能表捷徑與固定檔案關聯 | **~2.6 MB** | 1. 下載 `照片清冊產生器_x.x.x_x64-setup.exe`<br>2. 雙擊執行安裝程式，依提示按「下一步」完成安裝。 |

> 🔒 **100% 離線安全保證**：所有相依資源（樣式、圖示、Word/Excel/ZIP 生成核心）均已內建打包。無論是網頁版或桌面版，**圖片處理與檔案匯出均在您的本機記憶體內完成，絕不上傳任何伺服器或外部網路**，完全符合公務機敏資安標準。

---

## 📖 專案簡介

本專案專為警務同仁、刑事偵查與現場勘查人員打造之**現場照片清冊線上生成工具**。

v2.0/v2.1/v2.2 版本迎來重大架構革新：**徹底淘汰舊有 VBA 巨集與本機 Office 依賴**，升級為純前端單頁 Web 應用程式 (SPA) 與 Tauri 桌面版。使用者只需使用瀏覽器打開網頁或雙擊桌面版，即可批次拖曳照片、智慧填寫時間與說明，並直接在記憶體中生成符合公務標準的 **原生 Word (`.docx`) 清冊**，100% 零巨集資安警示！

---

## ✨ 核心特色與亮點

### 1. 🕒 時間與民國年日期智慧遮罩 (Masked Input)
- **時間遮罩輸入**：輸入 4 碼純數字自動格式化為 `HH:MM`（如 `1432` $\rightarrow$ `14:32`）、輸入 6 碼自動格式化為 `HH:MM:SS`（如 `143210` $\rightarrow$ `14:32:10`），亦保留手動輸入冒號相容性。
- **個別照片採證日期**：支援個別照片自訂中華民國曆日期（輸入 7 碼 `1130826` $\rightarrow$ `113/08/26`），若留空則自動沿用左側全域案件日期，兼顧統一案件與多日蒐證需求。
- **說明向下填滿保護**：執行說明「⬇️ 向下填滿」時自動跳過已有說明的照片，杜絕誤按覆寫既有內容。

### 2. 🎮 工作台 PaperSwitch 手感與整組多向移動
- **流暢縮圖縮放**：按住 `Ctrl + 滑鼠滾輪` 可在 90–400px 之間平滑調整縮圖大小，一般滾輪正常捲動頁面。
- **彈性多選與框選**：支援滑鼠空白處拖曳框選、`Ctrl` 點選多選、`Shift` 連續範圍選取。
- **全向同步位移**：支援 `Ctrl + ← / →`（左右前後移動一格）與 `Ctrl + ↑ / ↓`（上下整列移動），選取多張時整組同步位移，排版直覺。

### 3. 📄 內建三大經典 Word 版型（100% 精準對齊原公務範本）
- 🔼 **上下兩張 (A4 直式．經典版型)**：每頁 2 個獨立表格，大圖呈現（限寬 14cm / 限高 9cm），配備案由、日期、時間、編號、地點、製作人與詳細說明。
- ⏸️ **左右兩張 (A4 直式雙欄．對比版型)**：每頁 1 個整合表格，頂部全寬案由與蒐證資訊，左右並排 2 張照片（限寬 8cm / 限高 18cm）與各自獨立說明。
- ⏹️ **橫式三張 (A4 橫式三欄．廣角版型)**：每頁 1 個整合表格，橫向並排 3 張照片與詳細說明，適合全景連續畫面。

### 4. 🖼️ Canvas 影像前處理與列印品質最佳化
- **EXIF 自動轉正**：自動解析手機直拍之 EXIF Orientation 標記，自動修正旋轉角度，亦提供手動 90° 旋轉按鈕。
- **等比列印壓縮**：前端自動等比縮圖至適合 A4 列印之最適解析度（最大 1600px，JPEG 88% 品質），將清冊檔案控制在輕盈的 3MB~8MB，徹底杜絕 Word 檔案過大膨脹。

### 5. 📊 雙向相容與全方位匯出
- 📄 **匯出 Word 清冊 (`.docx`)**：純前端生成原生標準 OpenXML 文件，各版本 Word / WPS / LibreOffice 通用且無巨集警告。
- 📊 **匯出 Excel 清單 (`.xlsx`)**：輸出結構化資料表格（包含自訂日期、時間、地點、說明）。
- 📥 **匯入 Excel 說明檔 (`.xlsx`)**：支援將外部編輯好的 Excel 檔案匯入，自動依編號或檔名對齊填回說明。
- 🗜️ **匯出最佳化照片 ZIP (`.zip`)**：打包經旋轉校正與壓縮後的清晰照片。

---

## 🚀 快速上手步驟

### 步驟 1：載入照片
- 點擊或直接將照片檔案／整個資料夾拖曳至上方虛線區塊。
- 支援低記憶體縮圖模式，載入上百張照片依然流暢。

### 步驟 2：填寫基本資料與排版
- 左側填寫案件名稱、全域日期、單位與製作人。
- 點選縮圖或框選多張照片，在右側編輯器填寫時間、地點與說明，或使用「⬇️ 向下填滿」。
- 使用 `Ctrl + 方向鍵` 調整照片順序。

### 步驟 3：一鍵匯出
- 選擇版型（上下兩張／左右兩張／橫式三張），點擊 **「📄 匯出 Word 現場照片清冊」** 即可於數秒內下載標準公務 Word 文件！


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
