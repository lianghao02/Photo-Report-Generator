# 專案特定細則：現況照片清冊生成工具 (Photo Report Generator)

> [!IMPORTANT]
> 本專案嚴格遵循「全域開發憲法 v1.1」。所有開發計畫 (Plan) 與任務 (Task) 必須使用繁體中文。

## 1. 核心業務邏輯 (Business Logic)
- **VBA 工具交付**：本專案核心為 Excel VBA 工具 (`.xlsm`)。網頁僅作為下載與說明頁面。
- **下載資源管理**：工具壓縮檔 (`.rar`) 應存放於 `downloads/` 目錄，並確保網頁連結正確。

## 2. 架構轉化規範 (Constitution Alignment)
- **Web 單檔化**：將 `index.html` 作為唯一入口。
  -移除本地 `output.css`，改用 Tailwind CDN。
  - CSS 寫在 `<style>` 或透過 Tailwind Config 定義。
- **外部資源清單 (CDN)**：
  - Tailwind CSS: `https://cdn.tailwindcss.com`
  - FontAwesome (若有): 使用 CDN 或 Inline SVG。
  - Google Fonts: `Noto Sans TC` (允許使用外部字體)。

## 3. UI配置 (UI Configuration)
- **靜態頁面**：本頁面為靜態 Landing Page，雖無複雜互動，但仍需將可配置項目（如：下載連結、版本號、主題色）提取至 `CONFIG` 物件或透過 Tailwind Config 管理，以便日後維護。

## 4. 檔案結構清理
- **移除**：`node_modules/`, `src/` (若為 web source), `package.json`, `tailwind.config.js`。
- **保留**：`vba_src/` (VBA原始碼), `downloads/` (發布檔案), `README.md`, `index.html`。
