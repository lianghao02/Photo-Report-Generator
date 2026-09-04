const path = require('path');
const fs = require('fs');

const { chromium } = require('playwright');

async function runE2eTests() {
    console.log('========================================');
    console.log('開始執行 Phase 0B Web UI Playwright 測試');
    console.log('========================================\n');

    const htmlPath = path.resolve(__dirname, '../../index.html');
    const fileUrl = 'file:///' + htmlPath.replace(/\\/g, '/');

    const browser = await chromium.launch({ headless: true });
    const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
    const page = await context.newPage();

    // 監聽 console 警告或錯誤
    const pageErrors = [];
    page.on('pageerror', err => pageErrors.push(err.message));

    console.log(`[1/5] 開啟頁面: ${fileUrl}`);
    await page.goto(fileUrl, { waitUntil: 'load' });

    // 檢查 window.app 是否存在
    const isAppInitialized = await page.evaluate(() => typeof window.app !== 'undefined');
    if (!isAppInitialized) {
        throw new Error('window.app 未能成功初始化！');
    }
    console.log('  ✅ window.app 初始化完成');

    // [2/5] 上傳測試照片至 fileInput
    console.log('[2/5] 模擬檔案上傳 (fixtures)...');
    const fixturesDir = path.resolve(__dirname, '../fixtures');
    const filesToUpload = [
        path.join(fixturesDir, 'sample01.jpg'),
        path.join(fixturesDir, 'sample02.jpg'),
        path.join(fixturesDir, 'dup', 'sample01.jpg') // 同名 sample01.jpg
    ];

    const fileInput = await page.$('#fileInput');
    if (!fileInput) throw new Error('找不到 #fileInput 元素');
    await fileInput.setInputFiles(filesToUpload);

    // 等待縮圖渲染到 DOM
    await page.waitForFunction(() => {
        return window.app && window.app.photos && window.app.photos.length === 3;
    }, { timeout: 8000 });

    const photoCardsCount = await page.$$eval('.photo-thumb-card', cards => cards.length);
    if (photoCardsCount !== 3) {
        throw new Error(`預期有 3 張照片卡片，實際渲染了 ${photoCardsCount} 張`);
    }
    console.log('  ✅ 成功載入 3 張測試照片卡片並完成縮圖渲染');

    // [3/5] 驗證完整度篩選列與 Badge 數字
    console.log('[3/5] 驗證完整度篩選列 Badge 統計...');
    // 清空預設地點並重新呼叫 render()，檢驗預設地點對 missingLocation 的影響
    await page.evaluate(() => {
        document.getElementById('defaultLocation').value = '';
        window.app.render();
    });

    const badgesNoLocation = await page.evaluate(() => {
        return {
            all: document.getElementById('auditBadge_all')?.textContent?.trim(),
            missingLocation: document.getElementById('auditBadge_missingLocation')?.textContent?.trim(),
            missingDesc: document.getElementById('auditBadge_missingDesc')?.textContent?.trim(),
            duplicatePhotos: document.getElementById('auditBadge_duplicatePhotos')?.textContent?.trim()
        };
    });

    console.log('  當清空預設地點時之 Badge 數值:', badgesNoLocation);
    if (badgesNoLocation.all !== '3') throw new Error(`全部照片 Badge 預期 3，實際為 ${badgesNoLocation.all}`);
    if (badgesNoLocation.missingLocation !== '3') throw new Error(`未填地點 Badge 預期 3，實際為 ${badgesNoLocation.missingLocation}`);
    if (badgesNoLocation.missingDesc !== '3') throw new Error(`未填說明 Badge 預期 3，實際為 ${badgesNoLocation.missingDesc}`);
    if (badgesNoLocation.duplicatePhotos !== '2') throw new Error(`同名照片 Badge 預期 2，實際為 ${badgesNoLocation.duplicatePhotos}`);
    console.log('  ✅ 篩選列 Badge 統計數據計算完全正確');

    // 測試點擊「同名照片」篩選按鈕
    console.log('[4/5] 測試篩選視圖切換...');
    await page.click('#filterBtn_duplicatePhotos');
    
    // 檢查卡片可見數量
    const visibleCardsCount = await page.$$eval('.photo-thumb-card', cards => cards.length);
    if (visibleCardsCount !== 2) {
        throw new Error(`同名照片篩選下預期渲染 2 張卡片，實際渲染 ${visibleCardsCount} 張`);
    }
    console.log(`  ✅ 同名照片篩選視圖正常：正確顯示 ${visibleCardsCount} 張卡片`);

    // 切回「全部」
    await page.click('#filterBtn_all');
    const resetVisibleCards = await page.$$eval('.photo-thumb-card', cards => cards.length);
    if (resetVisibleCards !== 3) {
        throw new Error(`切回全部後預期渲染 3 張卡片，實際渲染 ${resetVisibleCards} 張`);
    }
    console.log('  ✅ 切回「全部」視圖正常：顯示全部 3 張照片');

    // [5/5] 測試匯出前非阻斷提醒彈窗 (exportAuditModal)
    console.log('[5/5] 測試匯出前非阻斷提醒 Modal...');
    // 呼叫 confirmExportWithAudit 觸發提示
    await page.evaluate(() => {
        window.testExportExecuted = false;
        window.app.confirmExportWithAudit(() => {
            window.testExportExecuted = true;
        }, 'Word 清冊');
    });

    // 驗證 Modal 顯示
    const isModalVisible = await page.$eval('#exportAuditModal', el => !el.classList.contains('hidden'));
    if (!isModalVisible) {
        throw new Error('匯出前確認 Modal 應顯示，但目前含有 hidden 類別');
    }

    const modalTitle = await page.$eval('#exportAuditModalTitle', el => el.textContent);
    if (!modalTitle.includes('Word 清冊')) {
        throw new Error(`Modal 標題應包含「Word 清冊」，實際為: ${modalTitle}`);
    }
    console.log(`  ✅ Modal 彈出成功，標題驗證正確: "${modalTitle}"`);

    // 測試點擊「查看問題照片」：應關閉 Modal 並自動切換篩選列
    await page.click('#btnAuditViewIssues');
    const isModalClosed = await page.$eval('#exportAuditModal', el => el.classList.contains('hidden'));
    if (!isModalClosed) {
        throw new Error('點擊「查看問題照片」後，Modal 應自動關閉');
    }

    const currentFilter = await page.evaluate(() => window.app.activeFilter);
    if (currentFilter === 'all') {
        throw new Error('點擊「查看問題照片」後，篩選模式應自動切換至首個問題分類，而非 all');
    }
    console.log(`  ✅ 「查看問題照片」行為正確：Modal 關閉且自動切換篩選視圖至 "${currentFilter}"`);

    // 測試「仍要匯出」按鈕
    await page.evaluate(() => {
        window.testExportExecuted = false;
        window.app.confirmExportWithAudit(() => {
            window.testExportExecuted = true;
        }, 'PDF 報表');
    });
    await page.click('#btnAuditProceedExport');

    const exportExecuted = await page.evaluate(() => window.testExportExecuted);
    if (!exportExecuted) {
        throw new Error('點擊「仍要匯出」後，暫存之匯出回呼動作應被執行！');
    }
    console.log('  ✅ 「仍要匯出」行為正確：Modal 關閉且匯出回呼正常觸發');

    if (pageErrors.length > 0) {
        console.warn('⚠️ 頁面出現 Uncaught Error:', pageErrors);
    }

    await browser.close();
    console.log('\n========================================');
    console.log('🎉 Phase 0B Playwright Web UI 測試全部通過！');
    console.log('========================================');
}

runE2eTests().catch(err => {
    console.error('\n❌ Phase 0B 測試失敗:', err);
    process.exit(1);
});
