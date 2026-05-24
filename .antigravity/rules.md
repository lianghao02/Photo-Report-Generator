# 專案特定細則：現況照片清冊生成工具 (Photo Report Generator Rules)

> [!IMPORTANT]
> 本專案嚴格遵循「全域大腦 v2.1.0 (Tech Lead)」。所有開發計畫 (Plan)、任務 (Task)、程式註解與使用者介面必須 100% 使用台灣繁體中文。保持直接、專業的 Tech Lead 溝通口吻，去除 AI 感。

## 1. 核心業務邏輯 (Business Logic)
- **VBA 工具交付**：本專案的真正核心為 Excel VBA 工具 (`.xlsm`)。網頁端的功能僅定位為兼具質感與實用性的下載與說明入口 (Landing Page)。
- **下載資源管理**：工具的壓縮檔 (`.rar` 或 `.zip`) 應妥善存放於 `downloads/` 目錄，並在 HTML 中確保下載連結的絕對正確性。

## 2. 架構轉化規範 (Architecture & Deployment)
- **Web 單檔化與精簡交付**：貫徹單檔交付精神，將 `index.html` 作為唯一的入口。
  - 果斷移除本地 `output.css` 或繁重的編譯配置，全面改用 Tailwind CDN。
  - 客製化 CSS 必須寫在 `<style>` 或透過 Tailwind Config Script 俐落定義。
- **外部資源清單 (CDN)**：
  - Tailwind CSS: `https://cdn.tailwindcss.com`
  - FontAwesome (若有): 優先使用 CDN 或輕量級 Inline SVG。
  - Google Fonts: `Noto Sans TC` (允許使用外部字體以提升質感)。

## 3. UI配置與極致美學 (Aesthetics & Configuration)
- **極致美學**：雖然本專案只是 Landing Page，但 UI 介面必須告別傳統公務系統的陽春感。請善用 Tailwind 的現代化佈局、漸層、陰影與過渡動畫 (Transitions)，打造兼具科技感與信賴感的視覺體驗。
- **配置提取**：將可配置項目（如：下載連結、當前版本號、主題色系）統一提取至 `CONFIG` 物件或透過 Tailwind Config 管理，確保未來維護的擴充性。

## 4. 檔案結構清理 (Defense & Housekeeping)
- **嚴格清理無用依賴**：
  - **移除**：`node_modules/`, `src/` (若是為了 Web 編譯而存在), `package.json`, `tailwind.config.js` 等冗餘檔案，降低維護成本。
  - **保留**：`vba_src/` (VBA 原始碼的備份), `downloads/` (發布檔案), `README.md`, 以及高內聚的 `index.html`。
