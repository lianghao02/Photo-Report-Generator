const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { chromium } = require('playwright');

async function main() {
    const browser = await chromium.launch({ headless: true });
    try {
        for (const [width, height] of [[1920,1080],[1366,768],[1100,760],[900,600],[768,600],[390,844]]) {
            const context = await browser.newContext({ viewport: { width, height } });
            const page = await context.newPage();
            await page.goto(pathToFileURL(path.resolve(__dirname, '../../index.html')).href);
            await page.waitForFunction(() => Boolean(window.app));
            const buffer = fs.readFileSync(path.resolve(__dirname, '../fixtures/sample01.jpg'));
            await page.setInputFiles('#fileInput', Array.from({ length: 12 }, (_, index) => ({ name: `版面驗證${index}.jpg`, mimeType: 'image/jpeg', buffer })));
            await page.waitForFunction(() => window.app.photos.length === 12);
            const layout = await page.evaluate(() => {
                const rect = id => { const r = document.getElementById(id).getBoundingClientRect(); return { top:r.top, bottom:r.bottom, width:r.width, height:r.height }; };
                return { scrollWidth:document.documentElement.scrollWidth, scrollHeight:document.documentElement.scrollHeight, export:rect('exportBar'), grid:rect('thumbGrid'), editor:rect('editorPanel'), next:rect('btnNextPhoto'), caseOpen:document.getElementById('caseDetails').open };
            });
            assert(layout.scrollWidth <= width + 1, `水平溢位：${width}`);
            if (width >= 768) {
                assert(layout.scrollHeight <= height + 1, `整頁高度溢位：${width} ${JSON.stringify(layout)}`);
                assert(layout.grid.bottom <= layout.export.top + 1, `匯出列遮住照片：${width}`);
                assert(layout.next.bottom <= layout.export.top + 1 && layout.next.top >= 0, `切換照片按鈕不可見：${width}`);
                assert(layout.editor.width >= 280, `編輯欄過窄：${width}`);
                assert.equal(layout.caseOpen, width >= 1200);
            }
            if (process.env.PHOTO_REPORT_LAYOUT_DIR) {
                fs.mkdirSync(process.env.PHOTO_REPORT_LAYOUT_DIR, { recursive: true });
                await page.screenshot({ path: path.join(process.env.PHOTO_REPORT_LAYOUT_DIR, `layout-${width}.png`) });
            }
            if (width === 900) {
                await page.locator('#caseDetails > summary').click();
                assert(await page.locator('#caseTitle').isVisible(), '窄視窗應可展開案件資料');
                const expandedNext = await page.locator('#btnNextPhoto').boundingBox();
                assert(expandedNext.y + expandedNext.height < height, '案件面板展開不應擠走照片導覽');
                await page.locator('#caseDetails > summary').click();
                await page.click('#btnToggleEditor');
                assert(await page.locator('#editorPanel').isHidden(), '可收合編輯欄');
                await page.click('#btnToggleEditor');
                await page.click('#btnWorkspaceHelp');
                assert(await page.locator('#shortcutsModal').isVisible(), '操作提示應開啟說明');
                await page.keyboard.press('Escape');
            }
            if (width === 1366) {
                await page.click('#btnViewGrid');
                assert(await page.locator('#gridContainer').isVisible());
                const table = await page.locator('#gridContainer').boundingBox();
                const bar = await page.locator('#exportBar').boundingBox();
                assert(table.y + table.height <= bar.y + 1, '匯出列不應遮住資料表');
                await page.click('#btnViewWorkspace');
                await page.locator('.thumbnail-stage').first().dblclick();
                assert(await page.locator('#imageLightboxModal').isVisible(), '燈箱仍正常');
                await page.keyboard.press('Escape');
            }
            console.log(`版面 ${width}×${height} 通過；照片區高 ${Math.round(layout.grid.height)}px，編輯欄寬 ${Math.round(layout.editor.width)}px`);
            await context.close();
        }
    } finally { await browser.close(); }
}
main().catch(error => { console.error(error); process.exitCode = 1; });
