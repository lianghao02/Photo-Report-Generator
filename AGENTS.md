# 04_Photo-Report-Generator Agent 開發規範

本專案遵循目前有效之全域開發憲法；本檔僅定義專案專屬規則與例外。

---

## 1. 技術棧與前端架構邊界
- **主力架構**：純前端 Web SPA (HTML5 / Vanilla JS / Canvas / Tailwind CSS) + Tauri v2 桌面應用程式（雙模式並行）。
- **架構純粹性標準**：
  - 本工具核心為零環境依賴、免伺服器之純前端應用。**嚴禁因專案成長或功能擴充而盲目引入 React、Vue、Angular 等重量級框架**。
  - 當單檔 `index.html` 規模過大時，應採「漸進拆分原生 ES 模組（如資料模型、UI 互動、畫布縮放、各類匯出模組）」進行重構，維持純 Vanilla JS 與零打包工具負擔。
- **100% 離線與機敏資安保證**：
  - 所有第三方函式庫（Tailwind CSS, FontAwesome, docx.js, jsPDF, JSZip, xlsx）皆必須置於本機 `vendor/` 目錄。
  - 嚴禁透過 CDN 或外部網路請求載入任何代碼、樣式或字型；所有照片處理與文件生成必須於本機記憶體（Object URL）內完成，**絕不上傳任何伺服器**。

---

## 2. 業務領域與公務清冊核心邊界
- **版型精準度為最高優先級（不可破壞核心）**：
  - 內建三大經典公務版型（A4 直式上下兩張、A4 直式雙欄左右兩張、A4 橫式三欄三張）為法律採證之官方標準格式。
  - 任何重構、模組化或功能調整，**嚴禁更動、偏移或破壞 Word (`.docx`) 與 PDF 之表格尺寸、中繼段落固定行高（5 點）、頁邊距、等比縮放與公務欄位結構**。
- **匯出模組隔離原則**：
  - Word (`docx.js`)、PDF (`jsPDF + Canvas`)、Excel (`xlsx`) 與 ZIP 匯出模組必須維持職責隔離，彼此邏輯互不污染。
  - 匯出大型報表時必須採用批次讓步處理（Yield to event loop）並呈現進度百分比，嚴禁阻塞瀏覽器主執行緒。
- **大量照片管理與資料完整度**：
  - 必須維持 Object URL 記憶體控制機制，避免 Base64 膨脹；清空或刪除時確實釋放 `revokeObjectURL`。
  - 支援資料完整度檢視與篩選（未填地點、未填說明、時間缺失、多選批次套用），且批次編輯時嚴格保護個別照片編號不被誤覆寫。

---

## 3. UI/UX 與操作手感標準
- **工作台互動體驗**：
  - 維持平滑連續縮放畫布、空白處拖曳框選、多選整組拖曳位移與鍵盤精準排序。
  - 維持多步 Undo / Redo 操作防護，歷程僅記錄資料參照，不複製原始圖檔。

---

## 4. 核心驗證方式
- 修改 JavaScript 邏輯、樣式或配置後，必須執行專案 QA 檢測與離線資源檢查：
  ```powershell
  powershell -ExecutionPolicy Bypass -File scripts\qa.ps1
  ```

---

## 5. 共用 Skill 引用與動態解析 (Discovery Rule)
- 本專案遵循 LiangHao 全生態系標準規範 Skill：`lianghao-development`（Canonical Source 位於 `Dev-Control-Center/skills/lianghao-development`，v1.0.0）。
- Agent 開始工作時：
  1. 優先使用目前環境可自動發現的 `lianghao-development` Skill
  2. 若平台未自動載入，使用既有動態解析順序：環境變數 `LIANGHAO_SKILL_HOME` ➜ `%USERPROFILE%\.lianghao\config.json` ➜ 鄰近工作區探索 `..\00_Dev-Control-Center\skills\lianghao-development` 或呼叫其 `skill-resolver.ps1`
  3. 先讀 Shared Skill
  4. 再讀本專案 `AGENTS.md`
  5. 再讀目前 `HANDOFF.md` / `IMPLEMENTATION_PLAN.md`
  6. 專案、Git、測試與建置實際狀態優先於文件歷史
- 跨 Agent HANDOFF 一律使用 `lianghao-development` 標準範本，以 `Repository Full Name`、`Branch`、`Commit SHA`、`Task Type` 為主要識別，嚴禁複製 Skill 到本專案。
