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
        assert.deepEqual(errors, [], '桌面 CSP 不應產生未捕捉例外');
        console.log('桌面 CSP：照片／資料夾選擇、載入照片、動態縮圖燈箱及拖放設定通過');
    } finally {
        if (browser) await browser.close();
        await new Promise(resolve => server.close(resolve));
    }
}
main().catch(error => { console.error(error); process.exitCode = 1; });
