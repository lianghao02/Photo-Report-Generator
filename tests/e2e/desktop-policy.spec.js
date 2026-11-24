const assert = require('node:assert/strict');
const fs = require('node:fs');
const http = require('node:http');
const path = require('node:path');
const { chromium } = require('playwright');

// 重現桌面建置加入 nonce 後的 CSP，避免 file:// 測試漏掉行內事件失效。
async function main() {
    const root = path.resolve(__dirname, '../..');
    const config = JSON.parse(fs.readFileSync(path.join(root, 'src-tauri/tauri.conf.json'), 'utf8'));
    assert.equal(config.app.windows[0].dragDropEnabled, false, 'Windows 必須將 HTML5 拖放交給前端');
    assert.equal(config.app.security.csp.includes('script-src-attr'), false, '事件委派完成後 CSP 不應再保留 script-src-attr unsafe-inline');
    const policy = config.app.security.csp.replace(/script-src ([^;]+)/, "$& 'nonce-desktop-regression'");
    const types = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.woff2': 'font/woff2' };
    const server = http.createServer((req, res) => {
        const relative = decodeURIComponent(new URL(req.url, 'http://localhost').pathname).replace(/^\/+/, '') || 'index.html';
        const file = path.resolve(root, relative);
        if (!file.startsWith(root + path.sep) || !fs.existsSync(file) || !fs.statSync(file).isFile()) {
            res.writeHead(404).end();
            return;
        }
        res.writeHead(200, { 'Content-Type': types[path.extname(file)] || 'application/octet-stream', 'Content-Security-Policy': policy });
        fs.createReadStream(file).pipe(res);
    });
    await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
    let browser;
    try {
        browser = await chromium.launch({ headless: true });
        const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
        const errors = [];
        page.on('pageerror', error => errors.push(error.message));
        await page.goto(`http://127.0.0.1:${server.address().port}`);
        await page.waitForFunction(() => Boolean(window.app));
        for (const [button, input] of [['#btnEmptyAddFolder', 'folderInput'], ['#btnEmptyAddPhotos', 'fileInput']]) {
            const chooserPromise = page.waitForEvent('filechooser');
            await page.click(button);
            const chooser = await chooserPromise;
            assert.equal(await chooser.element().getAttribute('id'), input);
            if (input === 'fileInput') await chooser.setFiles(path.join(root, 'tests/fixtures/sample01.jpg'));

        }
        await page.waitForFunction(() => window.app.photos.length === 1);
        // 縮圖仍使用動態行內事件，須在桌面 CSP 下真正點擊驗證。
        await page.locator('.thumbnail-stage').first().dblclick();
        await page.waitForFunction(() => !document.getElementById('imageLightboxModal').classList.contains('hidden'));
        await page.keyboard.press('Escape');

        // 驗證安全來源 (http://127.0.0.1) 下導引選取案件資料夾、模擬權限重問與實際寫檔不退回下載區
        const secureZoneResult = await page.evaluate(async () => {
            const mockFiles = new Map();
            const downloadsFallback = [];
            const origSaveAs = window.saveAs;
            window.saveAs = (_blob, name) => downloadsFallback.push(String(name || ''));

            let permState = 'prompt';
            let reqActivationAtCall = null;
            const makeDir = (name, prefix = '') => ({
                name,
                kind: 'directory',
                async queryPermission() { return permState; },
                async requestPermission() {
                    reqActivationAtCall = Boolean(navigator.userActivation?.isActive);
                    if (!reqActivationAtCall) {
                        const err = new Error('User activation is required to request permissions.');
                        err.name = 'NotAllowedError';
                        throw err;
                    }
                    permState = 'granted';
                    return 'granted';
                },
                async getDirectoryHandle(sub) {
                    return makeDir(sub, prefix ? `${prefix}/${sub}` : sub);
                },
                async getFileHandle(fileName, opts = {}) {
                    const key = prefix ? `${prefix}/${fileName}` : fileName;
                    if (!opts.create && !mockFiles.has(key)) {
                        const err = new Error('Not found');
                        err.name = 'NotFoundError';
                        throw err;
                    }
                    return {
                        name: fileName,
                        kind: 'file',
                        async createWritable() {
                            const chunks = [];
                            return {
                                async write(d) { chunks.push(d); },
                                async close() { mockFiles.set(key, new Blob(chunks)); },
                                async abort() {}
                            };
                        }
                    };
                }
            });

            window.showDirectoryPicker = async () => makeDir('案件匯出區');
            document.getElementById('caseDate').value = '115/11/03';
            document.getElementById('caseSubject').value = '001';
            document.getElementById('caseTitle').value = '公職人員選舉罷免法';
            document.getElementById('caseSubject').dispatchEvent(new Event('input', { bubbles: true }));

            document.getElementById('btnPickExportDir').click();
            const guideVisible = !document.getElementById('exportDirGuideModal').classList.contains('hidden');
            const guidePreview = document.getElementById('exportDirGuidePreview')?.textContent?.trim();
            const guideTitle = document.getElementById('exportDirGuideTitle')?.textContent?.trim();

            document.getElementById('btnConfirmPickExportDir').click();
            for (let i = 0; i < 40 && !window.app.exportDirectoryHandle; i++) {
                await new Promise(r => setTimeout(r, 25));
            }
            const labelAfterPick = document.getElementById('exportDirLabel')?.textContent?.trim();

            // 模擬重新授權狀態：permState 重設回 'prompt'，必須在點擊當下完成 requestPermission
            permState = 'prompt';
            window.__testCaptured = { mockFiles, downloadsFallback, origSaveAs, reqActivationAtCall: () => reqActivationAtCall };
            return { guideVisible, guidePreview, guideTitle, labelAfterPick, permAfterPick: permState };
        });

        assert.equal(secureZoneResult.guideVisible, true, '應先開啟白話導引視窗');
        assert.equal(secureZoneResult.guidePreview, '1151103_001_公職人員選舉罷免法');
        assert.equal(secureZoneResult.guideTitle, '選一個專門放案件的資料夾');
        assert.equal(secureZoneResult.labelAfterPick, '存到：案件匯出區 \\ 1151103_001_公職人員選舉罷免法');

        // 真正透過滑鼠點擊 #btnExportDocx -> #btnAuditProceedExport 完成匯出，確認不因延遲退回下載區
        await page.click('#btnExportDocx');
        await page.waitForFunction(() => !document.getElementById('exportAuditModal').classList.contains('hidden'));
        await page.click('#btnAuditProceedExport');
        await page.waitForFunction(() => !document.getElementById('exportDirSavedNotice').classList.contains('hidden') && !window.app.hasUnsavedProjectChanges());

        const afterRealClickExport = await page.evaluate(() => {
            const cap = window.__testCaptured;
            window.saveAs = cap.origSaveAs;
            return {
                files: Array.from(cap.mockFiles.keys()),
                downloadsCount: cap.downloadsFallback.length,
                reqActivated: cap.reqActivationAtCall(),
                notice: document.getElementById('exportDirSavedNotice')?.textContent?.trim()
            };
        });
        assert.equal(afterRealClickExport.reqActivated, true, '權限請求必須在點擊有效期間完成');
        assert.equal(afterRealClickExport.downloadsCount, 0, '寫入指定資料夾時不應退回下載區');
        assert(afterRealClickExport.files.includes('1151103_001_公職人員選舉罷免法/1151103_001_公職人員選舉罷免法_現場照片清冊.docx'));
        assert.equal(afterRealClickExport.files.includes('1151103_001_公職人員選舉罷免法/001_公職人員選舉罷免法.prp'), true);
        assert.equal(afterRealClickExport.notice, '已存進「1151103_001_公職人員選舉罷免法」');

        // 驗證若瀏覽器於選擇資料夾超過 1.2 秒後無聲拋出 AbortError（如權限被關閉），畫面應顯示明確提醒而非靜默失敗
        const abortDiagnostic = await page.evaluate(async () => {
            await window.app.clearExportDirectory();
            window.showDirectoryPicker = async () => {
                await new Promise(r => setTimeout(r, 1250));
                const err = new Error('The user aborted a request.');
                err.name = 'AbortError';
                throw err;
            };
            const picked = await window.app.pickExportDirectory({ skipGuide: true });
            return {
                picked,
                noticeVisible: !document.getElementById('exportDirSavedNotice')?.classList.contains('hidden'),
                noticeText: document.getElementById('exportDirSavedNotice')?.textContent?.trim() || ''
            };
        });
        assert.equal(abortDiagnostic.picked, false);
        assert.equal(abortDiagnostic.noticeVisible, true);
        assert(abortDiagnostic.noticeText.includes('未完成資料夾授權'), 'AbortError 逾時應顯示明確權限診斷提醒');

        // 驗證嚴格 CSP 下零行內 on* 屬性，且靜態／動態事件委派皆正常運作
        await page.click('#btnViewGrid');
        const delegationCheck = await page.evaluate(() => {
            const allInlineHandlers = [];
            document.querySelectorAll('*').forEach(el => {
                for (const attr of Array.from(el.attributes)) {
                    if (/^on/i.test(attr.name)) allInlineHandlers.push(`${el.tagName}.${attr.name}`);
                }
            });
            return {
                inlineCount: allInlineHandlers.length,
                gridRows: document.querySelectorAll('#gridTbody tr').length
            };
        });
        assert.equal(delegationCheck.inlineCount, 0, '嚴格 CSP 下 DOM 不得殘留任何行內 on* 屬性');
        assert.equal(delegationCheck.gridRows, 1, '切換試算表視圖後應正常渲染資料列');
        await page.click('#btnViewWorkspace');

        assert.deepEqual(errors, [], '桌面 CSP 不應產生未捕捉例外');
        console.log('桌面 CSP：照片／資料夾選擇、載入照片、動態縮圖燈箱、指定案件資料夾匯出、零行內事件委派及拖放設定通過');
    } finally {
        if (browser) await browser.close();
        await new Promise(resolve => server.close(resolve));
    }
}
main().catch(error => { console.error(error); process.exitCode = 1; });
