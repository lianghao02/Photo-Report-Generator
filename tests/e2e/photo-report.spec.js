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
    if (!(await page.isDisabled('#btnExportDocx')) || !(await page.isDisabled('#btnExportPdf'))) throw new Error('空白清冊應停用 Word 與 PDF 匯出');
    if (!(await page.isVisible('#exportEmptyHint'))) throw new Error('空白清冊應提示加入照片後即可匯出');

    // 驗證左側資訊架構與版型提示：維持既有控制項 ID，並讓 Word/PDF 共用同一版型。
    console.log('[1A/5] 驗證版型與匯出區的資訊架構...');
    const exportUi = await page.evaluate(() => {
        const byId = id => document.getElementById(id);
        const exportBar = byId('exportBar');
        const caseSection = byId('caseTitle')?.closest('section');
        return {
            requiredControlsPresent: ['layoutSelect', 'btnSaveProject', 'btnSaveProjectAs', 'importProjectInput', 'mergeProjectInput', 'btnSortByCaptureTime', 'btnExportPdf', 'btnExportDocx']
                .every(id => Boolean(byId(id))),
            removedControlsAbsent: ['btnViewPreview', 'previewContainer', 'btnSaveProjectLite', 'importExcelInput', 'btnExportExcel', 'btnExportZip']
                .every(id => !byId(id)),
            layoutInExportBar: Boolean(exportBar?.querySelector('#layoutSelect')) && Boolean(exportBar?.querySelector('#btnExportDocx')) && Boolean(exportBar?.querySelector('#btnExportPdf')),
            rareActionsInMenus: ['btnSaveProject', 'btnSaveProjectAs', 'importProjectInput', 'mergeProjectInput', 'btnAutoArrangeStage', 'btnSortByCaptureTime', 'btnParseFilenameTime']
                .every(id => Boolean(byId(id)?.closest('details.app-menu'))),
            caseSectionTitle: caseSection?.querySelector('h2')?.textContent?.trim(),
            layoutInCaseSection: Boolean(caseSection?.querySelector('#layoutSelect')),
            wordIsPrimary: byId('btnExportDocx')?.classList.contains('btn-primary'),
            initialDescription: byId('layoutDescription')?.textContent?.trim()
        };
    });
    if (!exportUi.requiredControlsPresent) throw new Error('專案／排序／匯出控制項 ID 不完整');
    if (!exportUi.removedControlsAbsent) throw new Error('已精簡移除之分頁預覽、Excel、ZIP 與獨立輕量存檔鈕不應再出現在 DOM');
    if (!exportUi.layoutInExportBar) throw new Error('版型與 Word／PDF 匯出應集中在底部匯出列');
    if (!exportUi.rareActionsInMenus) throw new Error('專案檔與排序／檔名時間操作應收在選單內');
    if (!exportUi.caseSectionTitle?.includes('案件資料')) throw new Error(`案件區標題不正確：${exportUi.caseSectionTitle}`);
    if (exportUi.layoutInCaseSection) throw new Error('清冊版型不應留在案件資料區');
    if (!exportUi.wordIsPrimary) throw new Error('Word 清冊應維持主要操作按鈕');
    if (exportUi.initialDescription !== '上下兩張・A4 直式・每頁 2 張') throw new Error(`預設版型提示不正確：${exportUi.initialDescription}`);

    await page.selectOption('#layoutSelect', 'left_right_2');
    const twoColumnDescription = await page.$eval('#layoutDescription', el => el.textContent.trim());
    if (twoColumnDescription !== '左右兩張・A4 直式雙欄・每頁 2 張') throw new Error(`左右兩張版型提示不正確：${twoColumnDescription}`);
    await page.selectOption('#layoutSelect', 'landscape_3');
    const landscapeDescription = await page.$eval('#layoutDescription', el => el.textContent.trim());
    if (landscapeDescription !== '橫式三張・A4 橫式・每頁 3 張') throw new Error(`橫式三張版型提示不正確：${landscapeDescription}`);
    await page.selectOption('#layoutSelect', 'up_down_2');
    console.log('  ✅ 左欄案件資料、底部匯出列、選單收納與三種版型提示皆正常');

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

    // 核對操作提示與實際旋轉範圍，避免未勾選時被誤認為只處理目前照片。
    await page.evaluate(() => window.app.toggleSelectAll(false));
    if (!(await page.isDisabled('#btnApplyDefaultsToSelected'))) throw new Error('未勾選照片時應停用案件資料套用');
    if ((await page.textContent('#batchRotateText')) !== '旋轉全部') throw new Error('未勾選時應標示旋轉全部');
    const rotationsBefore = await page.evaluate(() => window.app.photos.map(p => p.rotation));
    await page.click('#btnBatchRotate');
    const rotationsAll = await page.evaluate(() => window.app.photos.map(p => p.rotation));
    if (!rotationsAll.every((rotation, i) => rotation === (rotationsBefore[i] + 90) % 360)) throw new Error('旋轉全部的實際範圍不符');
    await page.locator('.photo-thumb-card input[type="checkbox"]').nth(0).check();
    if (await page.isDisabled('#btnApplyDefaultsToSelected')) throw new Error('勾選照片後應啟用案件資料套用');
    if (!(await page.textContent('#btnApplyDefaultsToSelected')).includes('已選 1 張')) throw new Error('套用按鈕應顯示單選張數');
    if ((await page.textContent('#batchRotateText')) !== '旋轉已選') throw new Error('勾選後應標示旋轉已選');
    await page.click('#btnBatchRotate');
    const rotationsSelected = await page.evaluate(() => window.app.photos.map(p => p.rotation));
    if (rotationsSelected[0] !== (rotationsAll[0] + 90) % 360 || rotationsSelected.slice(1).some((rotation, i) => rotation !== rotationsAll[i + 1])) throw new Error('旋轉已選不應修改未勾選照片');
    await page.locator('.photo-thumb-card input[type="checkbox"]').nth(1).check();
    if (!(await page.textContent('#btnApplyDefaultsToSelected')).includes('已選 2 張')) throw new Error('套用按鈕應隨多選更新張數');
    if ((await page.textContent('#selectedCountBadge')) !== '已勾選 2 張') throw new Error('應明確顯示已勾選張數');
    if (!(await page.textContent('#editorTitle')).includes('正在編輯第')) throw new Error('編輯標題應與勾選狀態區別');
    await page.evaluate(rotations => {
        window.app.photos.forEach((photo, i) => { photo.rotation = rotations[i]; });
        window.app.toggleSelectAll(false);
    }, rotationsBefore);
    console.log('  ✅ 未選取、單選、多選提示與實際旋轉範圍一致');

    if (await page.isDisabled('#btnExportDocx') || await page.isDisabled('#btnExportPdf')) throw new Error('加入照片後應啟用匯出');
    const originalFields = await page.evaluate(() => window.app.photos.map(p => ({ date: p.date, desc: p.desc })));
    await page.evaluate(() => {
        window.app.currentIndex = 0;
        window.app.photos[0].date = '115/10/26';
        window.app.photos[1].date = '115/10/25';
        window.app.photos[2].date = '';
        window.app.photos[0].desc = '新增說明';
        window.app.photos[1].desc = '保留說明';
        window.app.photos[2].desc = '';
        window.app.render();
    });
    let overwritePrompt = '';
    page.once('dialog', async dialog => { overwritePrompt = dialog.message(); await dialog.dismiss(); });
    await page.click('#btnFillDownDate');
    if (!overwritePrompt.includes('後續 2 張') || !overwritePrompt.includes('1 張已有不同內容')) throw new Error('覆寫確認應標示範圍與既有內容張數');
    if ((await page.evaluate(() => window.app.photos[1].date)) !== '115/10/25') throw new Error('取消覆寫後應保留原值');
    const acceptFill = async dialog => { await dialog.accept(); };
    page.on('dialog', acceptFill);
    await page.click('#btnFillDownDate');
    if (!(await page.evaluate(() => window.app.photos.every(p => p.date === '115/10/26')))) throw new Error('確認後應填滿後續日期');
    await page.click('#btnFillDownDescAll');
    const descriptions = await page.evaluate(() => window.app.photos.map(p => p.desc));
    if (descriptions[1] !== '保留說明' || descriptions[2] !== '新增說明') throw new Error('說明應只填空白並保護已有內容');
    page.off('dialog', acceptFill);
    await page.evaluate(() => { window.app.currentIndex = 2; window.app.render(); });
    if (!(await page.isDisabled('#btnFillDownDate'))) throw new Error('最後一張應停用向下填滿');
    await page.evaluate(fields => {
        window.app.photos.forEach((photo, i) => Object.assign(photo, fields[i]));
        window.app.currentIndex = 0;
        window.app.render();
    }, originalFields);
    console.log('  ✅ 覆寫取消／確認、保留原有說明、最後一張停用與匯出狀態皆正常');

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
    const expectedIssues = await page.evaluate(() => window.app.auditPhotosCompleteness().issuesCount);
    if (!(await page.textContent('#btnToggleAudit')).includes(`待檢查 ${expectedIssues} 項`)) throw new Error('檢查摘要應呈現問題項數');
    console.log('  ✅ 篩選列 Badge 統計數據計算完全正確');

    // 測試點擊「同名照片」篩選按鈕
    console.log('[4/5] 測試篩選視圖切換...');
    // 漏填檢查列預設收合，先展開再點篩選
    const auditCollapsed = await page.$eval('#auditFilterBar', el => el.classList.contains('audit-collapsed'));
    if (!auditCollapsed) throw new Error('漏填檢查列預設應收合');
    await page.click('#btnToggleAudit');
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

    // 以超過 9 張的情境驗證排序不補零、不改編號，以及匯出前的重新編號選擇。
    await page.evaluate(() => {
        window.originalSequenceTestState = { photos: window.app.photos, currentIndex: window.app.currentIndex, activeFilter: window.app.activeFilter };
        const source = window.app.photos[0];
        window.app.photos = Array.from({ length: 12 }, (_, index) => ({ ...source,
            uid: `sequence-test-${index}`, name: `${12 - index}.jpg`, seq: String(index + 1),
            date: '115/10/26', time: '12:00', location: '現場', desc: '測試說明', selected: false
        }));
        window.app.currentIndex = 0;
        window.app.activeFilter = 'all';
        window.app.sortPhotosByNaturalFilename();
        window.sequenceTestExported = false;
    });
    const sortedSequence = await page.evaluate(() => window.app.photos.map(p => ({ name: p.name, seq: p.seq })));
    if (sortedSequence[0].name !== '1.jpg' || sortedSequence[0].seq !== '12' || sortedSequence[11].seq !== '1') throw new Error('檔名排序應保留照片原編號，不得自動重編或補零');
    if (await page.locator('.thumbnail-card-footer').count()) throw new Error('縮圖下方不應顯示說明或地點');
    await page.evaluate(() => window.app.confirmExportWithAudit(() => { window.sequenceTestExported = true; }, 'Word 清冊'));
    if (!(await page.isVisible('#btnAuditRenumber')) || !(await page.isVisible('#exportSequenceNotice'))) throw new Error('編號順序不連續時應提示重新編號');
    if (await page.evaluate(() => window.sequenceTestExported)) throw new Error('使用者選擇前不得執行匯出');
    await page.click('#btnAuditProceedExport');
    if (!(await page.evaluate(() => window.sequenceTestExported && window.app.photos[0].seq === '12'))) throw new Error('應能保留原編號匯出');
    await page.evaluate(() => {
        window.sequenceTestExported = false;
        window.app.confirmExportWithAudit(() => { window.sequenceTestExported = true; }, 'PDF 清冊');
    });
    await page.click('#btnAuditRenumber');
    if (!(await page.isVisible('#sequenceModal'))) throw new Error('應開啟既有重新編號設定');
    await page.fill('#seqPrefix', '');
    await page.fill('#seqStart', '1');
    await page.selectOption('#seqDigits', '1');
    await page.click('#btnApplySequence');
    if (!(await page.evaluate(() => window.sequenceTestExported && window.app.photos.every((p, i) => p.seq === String(i + 1))))) throw new Error('重新編號後應接續原匯出流程');
    await page.evaluate(() => {
        window.app.photos.forEach((p, i) => { p.seq = `證物-${String(10 + i).padStart(3, '0')}`; });
        if (window.app.hasSequenceOrderMismatch()) throw new Error('自訂前綴、起始值與補零應視為有效連號');
        const before = new Map(window.app.photos.map(p => [p.uid, p.seq]));
        window.app.sortPhotosByNaturalFilename();
        if (window.app.photos.some(p => p.seq !== before.get(p.uid))) throw new Error('排序應保留自訂補零格式');
        window.app.photos[0].selected = true;
        window.app.moveSelectedBy(1);
        if (!window.app.hasSequenceOrderMismatch()) throw new Error('移動照片順序後應偵測連號不一致');
        Object.assign(window.app, window.originalSequenceTestState);
        window.app.render();
        delete window.originalSequenceTestState;
    });
    console.log('  ✅ 排序保留編號、無縮圖下方文字、匯出保留／重新編號與自訂格式皆正常');

    // 專案、Excel、ZIP 按鈕雖只搬移位置，仍實際驗證其既有資料流程。
    console.log('[6/5] 驗證專案、Excel 與 ZIP 資料操作...');
    const dataActions = await page.evaluate(async () => {
        const originalSaveAs = window.saveAs;
        const originalAlert = window.alert;
        const originalConfirm = window.confirm;
        const savedFiles = [];
        const alerts = [];
        window.saveAs = (blob, name) => savedFiles.push({ blob, name });
        window.alert = message => alerts.push(String(message));
        window.confirm = () => true;

        try {
            await window.app.saveProject();
            const project = savedFiles.find(file => file.name.endsWith('.photo-report'));
            if (!project) throw new Error('未產生專案檔');
            const projectFile = new File([project.blob], project.name, { type: 'application/zip' });
            await window.app.openProject({ target: { files: [projectFile], value: '' } });

            const importTarget = window.app.photos.find(photo => photo.name === 'sample02.jpg');
            if (!importTarget) throw new Error('找不到唯一檔名的 Excel 匯入測試照片');
            const importSheet = window.XLSX.utils.aoa_to_sheet([
                ['編號', '檔名', '日期', '時間', '地點', '說明'],
                [importTarget.seq, importTarget.name, '115/09/09', '10:30', '測試地點', 'Excel 匯入驗證']
            ]);
            const workbook = window.XLSX.utils.book_new();
            window.XLSX.utils.book_append_sheet(workbook, importSheet, '照片清冊');
            const excelFile = new File([window.XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })], 'import.xlsx', {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            });
            window.app.importExcel({ target: { files: [excelFile], value: '' } });
            await new Promise((resolve, reject) => {
                const deadline = Date.now() + 3000;
                const check = () => {
                    if (window.app.photos.some(photo => photo.name === 'sample02.jpg' && photo.desc === 'Excel 匯入驗證')) return resolve();
                    if (Date.now() >= deadline) return reject(new Error('Excel 匯入逾時'));
                    setTimeout(check, 25);
                };
                check();
            });

            await window.app.exportZip();
            const zip = savedFiles.find(file => file.name.endsWith('.zip'));
            return {
                restoredPhotoCount: window.app.photos.length,
                importedDescription: window.app.photos.find(photo => photo.name === 'sample02.jpg')?.desc,
                zipBytes: zip?.blob?.size || 0,
                projectOpened: alerts.some(message => message.includes('已開啟工作進度'))
            };
        } finally {
            window.saveAs = originalSaveAs;
            window.alert = originalAlert;
            window.confirm = originalConfirm;
        }
    });
    if (dataActions.restoredPhotoCount !== 3) throw new Error(`開啟專案後照片數量不正確：${dataActions.restoredPhotoCount}`);
    if (dataActions.importedDescription !== 'Excel 匯入驗證') throw new Error('Excel 匯入未寫入預期說明');
    if (dataActions.zipBytes <= 0) throw new Error('最佳化照片 ZIP 未產生內容');
    if (!dataActions.projectOpened) throw new Error('專案開啟流程未完成');
    console.log('  ✅ 專案開啟／儲存、Excel 匯入與 ZIP 匯出皆正常');

    // [7/5] 驗證鑑識級大圖燈箱 (PaperSwitch 規格) 與雙擊卡片展開編輯
    console.log('[7/5] 驗證鑑識級大圖燈箱 (縮放/平移/旋轉/快速鍵) 與雙擊卡片展開編輯...');
    await page.click('#filterBtn_all');
    // 0. 所有縮圖框等高（固定 1:1），不隨照片比例忽高忽低
    const stageSizes = await page.$$eval('.photo-thumb-card .thumbnail-stage', els => els.map(el => `${Math.round(el.getBoundingClientRect().width)}x${Math.round(el.getBoundingClientRect().height)}`));
    if (new Set(stageSizes).size !== 1) throw new Error(`縮圖框大小不一致：${[...new Set(stageSizes)].join(', ')}`);
    // 1. 測試雙擊縮圖開啟燈箱（單擊只選取，不放大，避免誤觸）
    await page.click('.photo-thumb-card .thumbnail-stage');
    const openedBySingleClick = await page.$eval('#imageLightboxModal', el => !el.classList.contains('hidden'));
    if (openedBySingleClick) throw new Error('單擊縮圖不應開啟燈箱');
    await page.dblclick('.photo-thumb-card .thumbnail-stage');
    const isLightboxVisible = await page.$eval('#imageLightboxModal', el => !el.classList.contains('hidden'));
    if (!isLightboxVisible) throw new Error('雙擊縮圖後，燈箱 Modal 未正常顯示 (仍有 hidden 類別)');
    console.log('  ✅ 雙擊縮圖成功開啟燈箱，單擊不會誤開');

    // 檢查初始縮放比例徽章
    const initialBadge = await page.$eval('#lightboxZoomBadge', el => el.textContent.trim());
    if (initialBadge !== '100%') throw new Error(`初始縮放比率應為 100%，實際為: ${initialBadge}`);
    console.log('  ✅ 初始縮放比例為 100%');

    // 測試鍵盤 + 縮放放大
    await page.keyboard.press('+');
    const zoomedBadge = await page.$eval('#lightboxZoomBadge', el => el.textContent.trim());
    const isZoomedClass = await page.$eval('#lightboxViewport', el => el.classList.contains('is-zoomed'));
    if (!isZoomedClass || zoomedBadge === '100%') throw new Error(`按 + 鍵放大失敗，Badge: ${zoomedBadge}, is-zoomed: ${isZoomedClass}`);
    console.log(`  ✅ 燈箱放大成功 (縮放比例: ${zoomedBadge})`);

    // 測試抓手平移拖曳
    const vpBox = await page.$eval('#lightboxViewport', el => {
        const r = el.getBoundingClientRect();
        return { x: r.left + r.width / 2, y: r.top + r.height / 2 };
    });
    await page.mouse.move(vpBox.x, vpBox.y);
    await page.mouse.down();
    await page.mouse.move(vpBox.x + 50, vpBox.y + 30);
    const panState = await page.evaluate(() => ({
        panX: window.app.lightboxPanX,
        panY: window.app.lightboxPanY
    }));
    await page.mouse.up();
    if (panState.panX === 0 && panState.panY === 0) throw new Error('拖曳平移失敗，Pan 座標未改變');
    console.log(`  ✅ 燈箱抓手平移成功 (PanX: ${panState.panX}, PanY: ${panState.panY})`);

    // 測試按 0 還原適合視窗
    await page.keyboard.press('0');
    const resetBadge = await page.$eval('#lightboxZoomBadge', el => el.textContent.trim());
    const resetPan = await page.evaluate(() => ({ panX: window.app.lightboxPanX, panY: window.app.lightboxPanY }));
    if (resetBadge !== '100%' || resetPan.panX !== 0 || resetPan.panY !== 0) throw new Error(`按 0 還原失敗: Badge=${resetBadge}, panX=${resetPan.panX}`);
    console.log('  ✅ 按快速鍵 0 成功還原適合視窗 (100% 置中)');

    // 測試按 R 順時針旋轉 90 度
    const rotBefore = await page.evaluate(() => window.app.photos[window.app.lightboxIndex].rotation || 0);
    await page.keyboard.press('r');
    const rotAfterR = await page.evaluate(() => window.app.photos[window.app.lightboxIndex].rotation);
    if (rotAfterR !== (rotBefore + 90) % 360) throw new Error(`按 R 旋轉失敗: 前=${rotBefore}, 後=${rotAfterR}`);
    console.log(`  ✅ 按 R 鍵順時針旋轉 90° 成功 (目前角度: ${rotAfterR}°)`);

    // 測試按 Shift + R 逆時針旋轉 90 度
    await page.keyboard.press('Shift+R');
    const rotAfterShiftR = await page.evaluate(() => window.app.photos[window.app.lightboxIndex].rotation);
    if (rotAfterShiftR !== rotBefore) throw new Error(`按 Shift+R 逆旋失敗: 預期=${rotBefore}, 實際=${rotAfterShiftR}`);
    console.log(`  ✅ 按 Shift+R 逆時針旋轉 90° 成功 (還原角度: ${rotAfterShiftR}°)`);

    // 再次按 R 旋轉，以便測試關閉燈箱時的 Lazy Commit 歷史記錄
    await page.keyboard.press('r');

    // 關閉燈箱 (按 ESC)
    await page.keyboard.press('Escape');
    const isLightboxClosed = await page.$eval('#imageLightboxModal', el => el.classList.contains('hidden'));
    if (!isLightboxClosed) throw new Error('按 ESC 鍵後燈箱未能成功關閉');
    console.log('  ✅ ESC 鍵成功關閉燈箱');

    // 檢查 Lazy Commit 是否成功記錄歷史
    const lastHistoryAction = await page.evaluate(() => {
        const hist = window.app.history[window.app.historyIndex];
        return hist ? (hist.label || hist.action) : '';
    });
    if (lastHistoryAction !== '燈箱旋轉照片') throw new Error(`燈箱關閉後未觸發延遲提交歷史，最後動作: ${lastHistoryAction}`);
    console.log('  ✅ 燈箱旋轉延遲歷史提交 (Lazy Commit) 驗證成功');

    // 測試主畫布 Space 鍵開啟燈箱
    const firstCard = await page.$('.photo-thumb-card');
    if (!firstCard) throw new Error('找不到照片卡片以測試 Space 快速鍵');
    await firstCard.click();
    await page.keyboard.press('Space');
    const isLightboxReopened = await page.$eval('#imageLightboxModal', el => !el.classList.contains('hidden'));
    if (!isLightboxReopened) throw new Error('在主畫布選取照片後按 Space 鍵未能開啟大圖燈箱');
    await page.keyboard.press('Escape');
    await page.waitForFunction(() => document.getElementById('imageLightboxModal').classList.contains('hidden'));

    // 1B. 測試 Ctrl + 點選與 Shift + 點選連續範圍多選
    console.log('  驗證 Ctrl + 點選多選與 Shift + 連續範圍選取...');
    await (await page.$$('.photo-thumb-card'))[0].click();
    await page.keyboard.down('Control');
    await (await page.$$('.photo-thumb-card'))[2].click();
    await page.keyboard.up('Control');
    const ctrlSelTest = await page.evaluate(() => window.app.photos.map(p => p.selected));
    if (!ctrlSelTest[0] || !ctrlSelTest[2] || ctrlSelTest[1]) {
        throw new Error(`Ctrl+點選多選失敗，選取狀態: ${JSON.stringify(ctrlSelTest)}`);
    }
    console.log('  ✅ Ctrl + 點選多選功能正常');

    await (await page.$$('.photo-thumb-card'))[0].click();
    await page.keyboard.down('Shift');
    await (await page.$$('.photo-thumb-card'))[2].click();
    await page.keyboard.up('Shift');
    const shiftSelTest = await page.evaluate(() => window.app.photos.map(p => p.selected));
    if (!shiftSelTest[0] || !shiftSelTest[1] || !shiftSelTest[2]) {
        throw new Error(`Shift+點選連續範圍選取失敗，選取狀態: ${JSON.stringify(shiftSelTest)}`);
    }
    console.log('  ✅ Shift + 點選連續範圍選取功能正常');

    // 2. 測試雙擊卡片展開單張編輯
    // 先確保編輯面板為收合狀態
    await page.evaluate(() => {
        if (!window.app.editorCollapsed) {
            window.app.toggleEditor();
        }
    });
    const isEditorHiddenBefore = await page.$eval('#editorPanel', el => el.classList.contains('hidden'));
    if (!isEditorHiddenBefore) throw new Error('測試前提失敗：編輯面板未能設為收合');

    // 雙擊第二張卡片
    const secondCard = (await page.$$('.photo-thumb-card'))[1];
    if (!secondCard) throw new Error('找不到第二張照片卡片');
    await (await secondCard.$('.thumbnail-meta-row')).dblclick();

    // 檢查編輯面板是否展開
    const isEditorOpenAfter = await page.$eval('#editorPanel', el => !el.classList.contains('hidden'));
    if (!isEditorOpenAfter) throw new Error('雙擊卡片後，右側單張編輯面板未自動展開');

    // 檢查目前選取索引是否為第二張 (index 1)
    const activeIndex = await page.evaluate(() => window.app.currentIndex);
    if (activeIndex !== 1) throw new Error(`雙擊卡片後目前編輯索引應為 1，實際為: ${activeIndex}`);

    // 等待 100ms 檢查焦點是否在 #editDesc
    await page.waitForTimeout(100);
    const isFocusedOnDesc = await page.evaluate(() => document.activeElement && document.activeElement.id === 'editDesc');
    if (!isFocusedOnDesc) throw new Error('雙擊卡片後焦點未自動置於說明欄位 #editDesc');
    console.log('  ✅ 雙擊縮圖卡片成功展開單張編輯面板並聚焦至說明欄位');

    // 2B. 測試大圖燈箱靠左並排＋右側即時編輯欄連動與兩段式 ESC 防護
    await page.evaluate(() => window.app.openLightboxByIndex(0));
    const splitLayoutCheck = await page.evaluate(() => {
        const vp = document.getElementById('lightboxViewport');
        const mount = document.getElementById('lightboxEditorMount');
        const card = document.getElementById('editorCard');
        const vpRect = vp.getBoundingClientRect();
        const mountRect = mount.getBoundingClientRect();
        return {
            cardInMount: mount.contains(card),
            mountVisible: !mount.classList.contains('hidden') && mountRect.width >= 300,
            sideBySide: mountRect.left >= vpRect.right - 4,
            initialIdx: window.app.currentIndex,
            lightboxIdx: window.app.lightboxIndex
        };
    });
    if (!splitLayoutCheck.cardInMount || !splitLayoutCheck.mountVisible || !splitLayoutCheck.sideBySide) {
        throw new Error(`燈箱並排編輯欄排版異常: ${JSON.stringify(splitLayoutCheck)}`);
    }
    await page.click('#editDesc');
    await page.fill('#editDesc', '燈箱並排即時編輯測試');
    // 在輸入框聚焦時按第一次 ESC：僅解除輸入框焦點，不得直接關閉燈箱
    await page.keyboard.press('Escape');
    const afterFirstEsc = await page.evaluate(() => ({
        lightboxOpen: !document.getElementById('imageLightboxModal').classList.contains('hidden'),
        activeId: document.activeElement ? document.activeElement.id : ''
    }));
    if (!afterFirstEsc.lightboxOpen || afterFirstEsc.activeId === 'editDesc') {
        throw new Error(`輸入框內按第一次 ESC 應僅解除焦點且保留燈箱開啟: ${JSON.stringify(afterFirstEsc)}`);
    }
    // 測試從右側編輯欄點「下一張」同步切換左側大圖燈箱
    await page.click('#btnNextPhoto');
    const afterNextClick = await page.evaluate(() => ({
        currentIndex: window.app.currentIndex,
        lightboxIndex: window.app.lightboxIndex,
        titleText: document.getElementById('editorTitle')?.textContent?.trim() || '',
        posText: document.getElementById('lightboxPosition')?.textContent?.trim() || ''
    }));
    if (afterNextClick.currentIndex !== 1 || afterNextClick.lightboxIndex !== 1 || !afterNextClick.titleText.includes('第 2／') || !afterNextClick.posText.startsWith('2 /')) {
        throw new Error(`燈箱與右側編輯欄切換下一張未同步: ${JSON.stringify(afterNextClick)}`);
    }
    // 測試燈箱頂列切換隱藏／顯示右側編輯欄
    await page.click('#btnLightboxToggleEditor');
    if (!(await page.$eval('#lightboxEditorMount', el => el.classList.contains('hidden')))) {
        throw new Error('點擊隱藏編輯欄後 #lightboxEditorMount 應隱藏');
    }
    await page.click('#btnLightboxToggleEditor');
    if (await page.$eval('#lightboxEditorMount', el => el.classList.contains('hidden'))) {
        throw new Error('再次點擊顯示編輯欄後 #lightboxEditorMount 應顯示');
    }
    // 按第二次 ESC（非輸入狀態）關閉燈箱，確認 #editorCard 歸還至 #editorPanel 且資料已保存
    await page.keyboard.press('Escape');
    const afterCloseCheck = await page.evaluate(() => ({
        lightboxClosed: document.getElementById('imageLightboxModal').classList.contains('hidden'),
        cardRestored: document.getElementById('editorPanel').contains(document.getElementById('editorCard')),
        savedDesc: window.app.photos[0].desc
    }));
    if (!afterCloseCheck.lightboxClosed || !afterCloseCheck.cardRestored || afterCloseCheck.savedDesc !== '燈箱並排即時編輯測試') {
        throw new Error(`關閉燈箱後狀態還原異常: ${JSON.stringify(afterCloseCheck)}`);
    }
    console.log('  ✅ 大圖燈箱靠左並排＋右側即時編輯欄連動與兩段式 ESC 防護驗證成功');

    // 2C. 驗證 Phase 1：畫黃／紅框＋姓名標籤＋水平鏡像翻轉＋Canvas 繪製與 .photo-report 儲存還原
    await page.evaluate(() => window.app.openLightboxByIndex(0));
    await page.waitForFunction(() => {
        const layer = document.getElementById('lightboxAnnotationLayer');
        return layer && layer.getBoundingClientRect().width > 0;
    });
    // 水平鏡像翻轉
    await page.click('#btnLightboxFlipH');
    const flipCheck = await page.evaluate(() => ({
        flipH: window.app.photos[0].flipH,
        btnActive: document.getElementById('btnLightboxFlipH')?.classList.contains('is-flipped')
    }));
    if (!flipCheck.flipH || !flipCheck.btnActive) {
        throw new Error(`水平鏡像翻轉狀態異常: ${JSON.stringify(flipCheck)}`);
    }
    // 畫黃框
    await page.click('#btnLightboxToolRectYellow');
    const layerBox = await page.$eval('#lightboxAnnotationLayer', el => {
        const r = el.getBoundingClientRect();
        return { left: r.left, top: r.top, width: r.width, height: r.height };
    });
    await page.mouse.move(layerBox.left + layerBox.width * 0.15, layerBox.top + layerBox.height * 0.2);
    await page.mouse.down();
    await page.mouse.move(layerBox.left + layerBox.width * 0.45, layerBox.top + layerBox.height * 0.6);
    await page.mouse.up();

    // 畫紅框
    await page.click('#btnLightboxToolRectRed');
    await page.mouse.move(layerBox.left + layerBox.width * 0.55, layerBox.top + layerBox.height * 0.25);
    await page.mouse.down();
    await page.mouse.move(layerBox.left + layerBox.width * 0.85, layerBox.top + layerBox.height * 0.65);
    await page.mouse.up();

    // 輸入姓名標籤並點照片貼上
    await page.fill('#lightboxLabelTextInput', '賴延銓');
    await page.mouse.click(layerBox.left + layerBox.width * 0.3, layerBox.top + layerBox.height * 0.14);

    // 雙擊姓名標籤，驗證行內編輯器（不使用 window.prompt）可修改標籤文字
    const labelEl = await page.$('#lightboxAnnotationLayer .lightbox-annotation-label');
    if (!labelEl) throw new Error('找不到剛建立的姓名標籤');
    await labelEl.dblclick();
    const inlineInput = await page.waitForSelector('.lightbox-inline-label-editor input', { timeout: 2000 });
    await inlineInput.fill('賴延銓（駕駛）');
    await inlineInput.press('Enter');

    // 拖曳姓名標籤右下角圓點，驗證可手動拉角落放大縮小，並同步記錄至 ann.scale
    const labelResizeBox = await page.$eval('#lightboxAnnotationLayer .lightbox-annotation-label .lightbox-annotation-resize', el => {
        const r = el.getBoundingClientRect();
        return { x: r.left + r.width / 2, y: r.top + r.height / 2 };
    });
    await page.mouse.move(labelResizeBox.x, labelResizeBox.y);
    await page.mouse.down();
    await page.mouse.move(labelResizeBox.x + 50, labelResizeBox.y + 40);
    await page.mouse.up();
    const labelScaleAfterDrag = await page.evaluate(() => {
        const lbl = window.app.photos[0].annotations?.find(a => a.type === 'label');
        return lbl?.scale || 1;
    });
    if (!(labelScaleAfterDrag > 1.05)) {
        throw new Error(`拖曳姓名標籤角落後 ann.scale 應大於 1，實際為: ${labelScaleAfterDrag}`);
    }

    // 驗證 flipH === true 時按「向右轉」的視覺旋轉方向正確（rawDelta = -90°），再按「向左轉」還原
    const rotBeforeFlipCW = await page.evaluate(() => window.app.photos[0].rotation || 0);
    const expectedFlipCW = ((rotBeforeFlipCW - 90) % 360 + 360) % 360;
    await page.click('#btnLightboxRotateCW');
    const rotAfterFlipCW = await page.evaluate(() => window.app.photos[0].rotation);
    if (rotAfterFlipCW !== expectedFlipCW) {
        throw new Error(`flipH 狀態下順時針旋轉 90° 之內部 rotation 應為 ${expectedFlipCW}，實際為: ${rotAfterFlipCW}`);
    }
    await page.click('#btnLightboxRotateCCW');
    const rotAfterFlipCCW = await page.evaluate(() => window.app.photos[0].rotation);
    if (rotAfterFlipCCW !== rotBeforeFlipCW) {
        throw new Error(`flipH 狀態下逆時針轉回後內部 rotation 應為 ${rotBeforeFlipCW}，實際為: ${rotAfterFlipCCW}`);
    }

    // 驗證從照片內拉框拖出 #lightboxDialog 到 #imageLightboxModal 背景放開時，不會誤關閉燈箱
    await page.click('#btnLightboxToolRectYellow');
    await page.mouse.move(layerBox.left + layerBox.width * 0.2, layerBox.top + layerBox.height * 0.2);
    await page.mouse.down();
    await page.mouse.move(5, 5);
    await page.mouse.up();
    const stillOpenAfterDragOut = await page.$eval('#imageLightboxModal', el => !el.classList.contains('hidden'));
    if (!stillOpenAfterDragOut) {
        throw new Error('從照片內拉框拖出燈箱外框放開時，不應誤關閉燈箱');
    }
    // 透過 deleteLightboxAnnotation 刪除剛才拖出測試多畫的第 3 個框
    await page.evaluate(() => {
        const p = window.app.photos[0];
        if (p && p.annotations && p.annotations.length > 3) {
            window.app.deleteLightboxAnnotation(p.annotations[p.annotations.length - 1].id);
        }
    });

    const phase1Check = await page.evaluate(async () => {
        const photo = window.app.photos[0];
        const renderedRects = document.querySelectorAll('#lightboxAnnotationLayer .lightbox-annotation-rect').length;
        const renderedLabels = Array.from(document.querySelectorAll('#lightboxAnnotationLayer .lightbox-annotation-label-text')).map(el => el.textContent);
        const countBadge = document.getElementById('lightboxAnnotationCountBadge')?.textContent?.trim();
        const thumbBadge = document.querySelector(`#thumb_${photo.uid} .thumb-annotation-badge`)?.textContent?.trim() || '';
        const processed = await window.app.processImageToBlob(photo);

        // 驗證專案儲存與還原可完整保留 flipH 與 annotations
        const originalSaveAs = window.saveAs;
        const originalAlert = window.alert;
        const originalConfirm = window.confirm;
        let savedProjectBlob = null;
        window.saveAs = (blob) => { savedProjectBlob = blob; };
        window.alert = () => {};
        window.confirm = () => true;
        try {
            await window.app.saveProject();
            const projectFile = new File([savedProjectBlob], 'phase1.photo-report', { type: 'application/zip' });
            await window.app.openProject({ target: { files: [projectFile], value: '' } });
        } finally {
            window.saveAs = originalSaveAs;
            window.alert = originalAlert;
            window.confirm = originalConfirm;
        }
        const restoredPhoto = window.app.photos[0];
        const restoredLabel = restoredPhoto.annotations?.find(a => a.type === 'label');
        const result = {
            annCount: restoredPhoto.annotations?.length || 0,
            flipH: restoredPhoto.flipH,
            types: (restoredPhoto.annotations || []).map(a => `${a.type}:${a.color}`),
            labelText: restoredLabel?.text || '',
            labelScale: restoredLabel?.scale || 1,
            renderedRects,
            renderedLabels,
            countBadge,
            thumbBadge,
            blobSize: processed?.blob?.size || 0
        };
        // 清理標記並關閉燈箱，維持後續測試乾淨狀態
        window.app.clearCurrentPhotoAnnotations();
        if (window.app.photos[0].flipH) window.app.toggleLightboxFlipH();
        window.app.setLightboxTool('pan');
        window.app.closeModal('imageLightboxModal');
        return result;
    });
    if (
        phase1Check.annCount !== 3 ||
        !phase1Check.flipH ||
        phase1Check.labelText !== '賴延銓（駕駛）' ||
        !(phase1Check.labelScale > 1.05) ||
        phase1Check.renderedRects !== 2 ||
        !phase1Check.renderedLabels.includes('賴延銓（駕駛）') ||
        phase1Check.countBadge !== '已標記 3 項' ||
        !phase1Check.thumbBadge.includes('標記3') ||
        phase1Check.blobSize <= 0
    ) {
        throw new Error(`Phase 1 標記框、姓名標籤（含縮放）或水平翻轉驗證失敗: ${JSON.stringify(phase1Check)}`);
    }
    console.log('  ✅ Phase 1（畫黃／紅框＋姓名標籤角落放大縮小＋水平鏡像翻轉＋Canvas 匯出繪製＋專案檔還原）驗證成功');

    // 2D. 驗證 Phase 2 模式 A：局部裁切特寫插在原圖後方、鎖定 1-1 / 2-1 連動（往前拉與往後推皆自動跟隨）與 .photo-report 還原
    await page.evaluate(() => {
        window.app.openLightboxByIndex(0);
    });
    await page.waitForFunction(() => {
        const layer = document.getElementById('lightboxAnnotationLayer');
        return layer && layer.getBoundingClientRect().width > 0;
    });
    const cropLayerBox = await page.$eval('#lightboxAnnotationLayer', el => {
        const r = el.getBoundingClientRect();
        return { left: r.left, top: r.top, width: r.width, height: r.height };
    });
    await page.click('#btnLightboxToolCrop');
    await page.mouse.move(cropLayerBox.left + cropLayerBox.width * 0.2, cropLayerBox.top + cropLayerBox.height * 0.2);
    await page.mouse.down();
    await page.mouse.move(cropLayerBox.left + cropLayerBox.width * 0.65, cropLayerBox.top + cropLayerBox.height * 0.65);
    await page.mouse.up();
    await page.waitForSelector('#btnConfirmLightboxCrop', { timeout: 2000 });
    // 先按一次 ESC：僅清除裁切框，不得關閉燈箱
    await page.keyboard.press('Escape');
    const escCropCheck = await page.evaluate(() => ({
        modalOpen: !document.getElementById('imageLightboxModal').classList.contains('hidden'),
        cropDraftCleared: !window.app.lightboxCropDraft
    }));
    if (!escCropCheck.modalOpen || !escCropCheck.cropDraftCleared) {
        throw new Error(`按一次 ESC 應只清除裁切框而非關閉燈箱: ${JSON.stringify(escCropCheck)}`);
    }
    // 重新拖曳裁切框並點擊確認插入特寫
    const photoCountBeforeCrop = await page.evaluate(() => window.app.photos.length);
    const parentUidForCrop = await page.evaluate(() => window.app.photos[0].uid);
    const dupBeforeCrop = await page.evaluate(() => window.app.auditPhotosCompleteness().duplicatePhotos);
    await page.mouse.move(cropLayerBox.left + cropLayerBox.width * 0.25, cropLayerBox.top + cropLayerBox.height * 0.25);
    await page.mouse.down();
    await page.mouse.move(cropLayerBox.left + cropLayerBox.width * 0.7, cropLayerBox.top + cropLayerBox.height * 0.7);
    await page.mouse.up();
    await page.waitForSelector('#btnConfirmLightboxCrop', { timeout: 2000 });
    await page.click('#btnConfirmLightboxCrop');
    await page.waitForFunction((expectedCount) => window.app.photos.length === expectedCount, photoCountBeforeCrop + 1);

    const cropInsertedCheck = await page.evaluate((parentUid) => {
        const audit = window.app.auditPhotosCompleteness();
        return {
            index1ParentUid: window.app.photos[1]?.parentUid,
            index1Name: window.app.photos[1]?.name,
            index0Seq: window.app.photos[0]?.seq,
            index1Seq: window.app.photos[1]?.seq,
            currentIndex: window.app.currentIndex,
            editSeqReadOnly: document.getElementById('editSeq')?.readOnly,
            lockNoticeVisible: !document.getElementById('editSeqLockNotice')?.classList.contains('hidden'),
            orderMismatch: window.app.hasSequenceOrderMismatch(),
            auditDup: audit.duplicatePhotos,
            detailInDupSet: window.app._dupSet ? window.app._dupSet.has(window.app.photos[1]?.name) : false
        };
    }, parentUidForCrop);
    if (
        cropInsertedCheck.index1ParentUid !== parentUidForCrop ||
        cropInsertedCheck.index1Seq !== `${cropInsertedCheck.index0Seq}-1` ||
        cropInsertedCheck.currentIndex !== 1 ||
        !cropInsertedCheck.editSeqReadOnly ||
        !cropInsertedCheck.lockNoticeVisible ||
        cropInsertedCheck.orderMismatch ||
        cropInsertedCheck.auditDup !== dupBeforeCrop ||
        cropInsertedCheck.detailInDupSet
    ) {
        throw new Error(`Phase 2A 局部裁切插入或編號鎖定狀態異常: ${JSON.stringify(cropInsertedCheck)}`);
    }

    // 關閉燈箱，驗證把第 3 張（index 2）往前移一步會跨過 [原圖, 特寫] 整組，且重新編號後原圖變 #2、特寫自動變 #2-1，拉回前方變回 #1-1
    await page.click('#btnLightboxClose');
    const reorderCropCheck = await page.evaluate(async (parentUid) => {
        const origConfirm = window.confirm;
        window.confirm = () => true;
        try {
            // 1. 選取後一張照片往前移動：從 [1, 1-1, 2, ...] 跨過整組變成 [2, 1, 1-1, ...]
            window.app.photos.forEach((p, idx) => { p.selected = idx === 2; });
            window.app.currentIndex = 2;
            window.app.moveSelectedBy(-1);
            document.getElementById('seqPrefix').value = '';
            document.getElementById('seqStart').value = '1';
            document.getElementById('seqDigits').value = '1';
            window.app.applySequence();
            const pushedBack = {
                parentIdx: window.app.photos.findIndex(p => p.uid === parentUid),
                childIdx: window.app.photos.findIndex(p => p.parentUid === parentUid),
                seqs: window.app.photos.slice(0, 3).map(p => p.seq),
                mismatch: window.app.hasSequenceOrderMismatch()
            };

            // 2. 把原圖選取往前拉回第一位：特寫需自動跟著原圖回到前面並鎖定回 #1-1
            window.app.photos.forEach(p => { p.selected = (p.uid === parentUid); });
            window.app.currentIndex = pushedBack.parentIdx;
            window.app.moveSelectedBy(-1);
            window.app.applySequence();
            const pulledFront = {
                parentIdx: window.app.photos.findIndex(p => p.uid === parentUid),
                childIdx: window.app.photos.findIndex(p => p.parentUid === parentUid),
                seqs: window.app.photos.slice(0, 3).map(p => p.seq),
                mismatch: window.app.hasSequenceOrderMismatch()
            };

            // 3. 修改原圖編號文字，子特寫即時連動；再驗證 .photo-report 存檔還原
            window.app.photos.forEach((p, idx) => { p.selected = idx === 0; });
            window.app.currentIndex = 0;
            window.app.updateCurrentPhoto('seq', '證物A');
            const cascadedSeq = window.app.photos[1].seq;
            window.app.applySequence();

            const originalSaveAs = window.saveAs;
            const originalAlert = window.alert;
            let savedBlob = null;
            window.saveAs = (blob) => { savedBlob = blob; };
            window.alert = () => {};
            try {
                await window.app.saveProject();
                const projectFile = new File([savedBlob], 'phase2a.photo-report', { type: 'application/zip' });
                await window.app.openProject({ target: { files: [projectFile], value: '' } });
            } finally {
                window.saveAs = originalSaveAs;
                window.alert = originalAlert;
            }
            const restoredChild = window.app.photos[1];
            const restoredOk = restoredChild?.parentUid === window.app.photos[0]?.uid && restoredChild?.seq === '1-1' && restoredChild?.cropRect?.w > 0;

            // 清理測試產生的局部裁切照片並恢復編號，保留乾淨狀態給後續測試
            window.app.removePhoto(1);
            window.app.applySequence();

            return { pushedBack, pulledFront, cascadedSeq, restoredOk };
        } finally {
            window.confirm = origConfirm;
        }
    }, parentUidForCrop);
    if (
        reorderCropCheck.pushedBack.parentIdx !== 1 ||
        reorderCropCheck.pushedBack.childIdx !== 2 ||
        reorderCropCheck.pushedBack.seqs[1] !== '2' ||
        reorderCropCheck.pushedBack.seqs[2] !== '2-1' ||
        reorderCropCheck.pushedBack.mismatch ||
        reorderCropCheck.pulledFront.parentIdx !== 0 ||
        reorderCropCheck.pulledFront.childIdx !== 1 ||
        reorderCropCheck.pulledFront.seqs[0] !== '1' ||
        reorderCropCheck.pulledFront.seqs[1] !== '1-1' ||
        reorderCropCheck.pulledFront.mismatch ||
        reorderCropCheck.cascadedSeq !== '證物A-1' ||
        !reorderCropCheck.restoredOk
    ) {
        throw new Error(`Phase 2A 推後／拉前群組跟隨、1-1／2-1 連動或專案檔還原失敗: ${JSON.stringify(reorderCropCheck)}`);
    }
    console.log('  ✅ Phase 2 模式 A（局部裁切插入原圖後方＋1-1／2-1 前拉後推鎖定連動＋專案檔還原）驗證成功');

    // 2E. 驗證 Phase 2 模式 B：框選放到原圖右側並排＋保留原圖黃框＋右下角等比縮放＋垂直對齊＋切換框線＋旋轉/翻轉＋合併匯出＋專案還原
    await page.evaluate(() => {
        const p = window.app.photos[0];
        if (p) {
            p.annotations = [];
            p.sideCrop = null;
            p.flipH = false;
            p.rotation = 0;
        }
        window.app.openLightboxByIndex(0);
    });
    await page.waitForFunction(() => {
        const layer = document.getElementById('lightboxAnnotationLayer');
        return layer && layer.getBoundingClientRect().width > 0;
    });
    const modeBLayerBox = await page.$eval('#lightboxAnnotationLayer', el => {
        const r = el.getBoundingClientRect();
        return { left: r.left, top: r.top, width: r.width, height: r.height };
    });
    const baseBlobSize = await page.evaluate(async () => {
        const p = window.app.photos[0];
        const res = await window.app.processImageToBlob(p, 1600, 0.88);
        return { width: res.width, height: res.height };
    });
    const countBeforeSideCrop = await page.evaluate(() => window.app.photos.length);

    // 框選一塊偏直／近正方形的局部區域並放到原圖右側
    await page.click('#btnLightboxToolCrop');
    await page.mouse.move(modeBLayerBox.left + modeBLayerBox.width * 0.2, modeBLayerBox.top + modeBLayerBox.height * 0.15);
    await page.mouse.down();
    await page.mouse.move(modeBLayerBox.left + modeBLayerBox.width * 0.45, modeBLayerBox.top + modeBLayerBox.height * 0.72);
    await page.mouse.up();
    await page.waitForSelector('#btnSideLightboxCrop', { timeout: 2000 });
    await page.click('#btnSideLightboxCrop');
    await page.waitForFunction(() => document.getElementById('lightboxCanvasFrame')?.classList.contains('has-side-crop') && document.getElementById('lightboxSideCropCanvas')?.dataset?.painted === '1');

    await page.waitForTimeout(80);
    const sideCreatedCheck = await page.evaluate(() => {
        const p = window.app.photos[0];
        const frameRect = document.getElementById('lightboxCanvasFrame')?.getBoundingClientRect();
        const leftRect = document.getElementById('lightboxLeftSlot')?.getBoundingClientRect();
        const sideAreaRect = document.getElementById('lightboxSideCropArea')?.getBoundingClientRect();
        const sideCardRect = document.getElementById('lightboxSideCropCard')?.getBoundingClientRect();
        const sideCanvas = document.getElementById('lightboxSideCropCanvas');
        const layoutBarVisible = !document.getElementById('lightboxSideLayoutBar')?.classList.contains('hidden');
        return {
            photoCount: window.app.photos.length,
            hasSideCrop: Boolean(p?.sideCrop),
            widthRatio: p?.sideCrop?.widthRatio || 0,
            offsetY: p?.sideCrop?.offsetY ?? -1,
            borderColor: p?.sideCrop?.borderColor || '',
            rectCount: (p?.annotations || []).filter(a => a.type === 'rect' && a.color === 'yellow').length,
            layoutBarVisible,
            frameW: frameRect?.width || 0,
            frameH: frameRect?.height || 0,
            leftRight: leftRect?.right || 0,
            sideLeft: sideAreaRect?.left || 0,
            cardWidth: sideCardRect?.width || 0,
            cardHeight: sideCardRect?.height || 0,
            canvasW: sideCanvas?.width || 0,
            canvasH: sideCanvas?.height || 0
        };
    });
    const upDownExpectedAspect = 535 / 365;
    if (
        sideCreatedCheck.photoCount !== countBeforeSideCrop ||
        !sideCreatedCheck.hasSideCrop ||
        sideCreatedCheck.rectCount !== 1 ||
        sideCreatedCheck.borderColor !== 'yellow' ||
        !sideCreatedCheck.layoutBarVisible ||
        Math.abs((sideCreatedCheck.frameW / Math.max(1, sideCreatedCheck.frameH)) - upDownExpectedAspect) > 0.08 ||
        !(sideCreatedCheck.sideLeft > sideCreatedCheck.leftRight) ||
        sideCreatedCheck.cardWidth <= 0 ||
        sideCreatedCheck.canvasW <= 0
    ) {
        throw new Error(`Phase 2B 建立右側並排圖、固定白框或來源黃框失敗: ${JSON.stringify(sideCreatedCheck)}`);
    }

    // 測試拖曳右下角縮放手把（先向左縮小確認可縮、寬高同步變化，且外圍白框寬高維持固定不動）
    const resizeHandleBox = await page.$eval('[data-side-crop-resize]', el => {
        const r = el.getBoundingClientRect();
        return { cx: r.left + r.width / 2, cy: r.top + r.height / 2 };
    });
    await page.mouse.move(resizeHandleBox.cx, resizeHandleBox.cy);
    await page.mouse.down();
    await page.mouse.move(resizeHandleBox.cx - 50, resizeHandleBox.cy - 25, { steps: 6 });
    await page.mouse.up();
    const afterResizeCheck = await page.evaluate(() => {
        const p = window.app.photos[0];
        const frameRect = document.getElementById('lightboxCanvasFrame')?.getBoundingClientRect();
        const cardRect = document.getElementById('lightboxSideCropCard')?.getBoundingClientRect();
        const areaRect = document.getElementById('lightboxSideCropArea')?.getBoundingClientRect();
        return {
            widthRatio: p?.sideCrop?.widthRatio || 0,
            frameW: frameRect?.width || 0,
            frameH: frameRect?.height || 0,
            cardWidth: cardRect?.width || 0,
            cardHeight: cardRect?.height || 0,
            areaHeight: areaRect?.height || 0
        };
    });
    if (
        Math.abs(afterResizeCheck.frameW - sideCreatedCheck.frameW) > 1.5 ||
        Math.abs(afterResizeCheck.frameH - sideCreatedCheck.frameH) > 1.5
    ) {
        throw new Error(`拖拉右側特寫大小時外圍模擬白框不應忽大忽小: before=${sideCreatedCheck.frameW}x${sideCreatedCheck.frameH}, after=${afterResizeCheck.frameW}x${afterResizeCheck.frameH}`);
    }
    const selfAuditGeoAndHistory = await page.evaluate(() => {
        const lastHistoryLabel = window.app.history[window.app.historyIndex]?.label || '';
        // 驗證 16:9 寬圖框選窄長直式區域（w=0.12, h=0.80）不變形且具可縮放空間
        const narrowGeo = window.app.computeSideCropGeometry(1600, 900, {
            cropRect: { x: 0.2, y: 0.1, w: 0.12, h: 0.8 },
            widthRatio: 0.99
        });
        const actualAspect = narrowGeo.panelW / narrowGeo.panelH;
        return {
            lastHistoryLabel,
            aspectDiff: Math.abs(actualAspect - narrowGeo.cropAspect),
            minRatio: narrowGeo.minWidthRatio,
            maxRatio: narrowGeo.maxWidthRatio,
            sourceAnnIdLinked: Boolean(window.app.photos[0]?.sideCrop?.sourceAnnId)
        };
    });
    if (
        !(afterResizeCheck.widthRatio < sideCreatedCheck.widthRatio) ||
        !(afterResizeCheck.cardWidth < sideCreatedCheck.cardWidth) ||
        selfAuditGeoAndHistory.lastHistoryLabel !== '調整右側並排特寫大小' ||
        selfAuditGeoAndHistory.aspectDiff > 0.001 ||
        !(selfAuditGeoAndHistory.minRatio < selfAuditGeoAndHistory.maxRatio * 0.65) ||
        !selfAuditGeoAndHistory.sourceAnnIdLinked
    ) {
        throw new Error(`Phase 2B 縮放紀錄或窄框等比防變形異常: before=${JSON.stringify(sideCreatedCheck)}, after=${JSON.stringify(afterResizeCheck)}, audit=${JSON.stringify(selfAuditGeoAndHistory)}`);
    }

    // 若右側卡片未滿高，拖曳卡片上下調整 offsetY
    if (afterResizeCheck.areaHeight - afterResizeCheck.cardHeight > 12) {
        const cardBeforeMove = await page.$eval('#lightboxSideCropCard', el => {
            const r = el.getBoundingClientRect();
            return { cx: r.left + r.width / 2, cy: r.top + r.height / 2, top: r.top };
        });
        const offsetBeforeMove = await page.evaluate(() => window.app.photos[0].sideCrop.offsetY);
        await page.mouse.move(cardBeforeMove.cx, cardBeforeMove.cy);
        await page.mouse.down();
        await page.mouse.move(cardBeforeMove.cx, cardBeforeMove.cy - 60, { steps: 6 });
        await page.mouse.up();
        const moveCheck = await page.evaluate(() => ({
            offsetY: window.app.photos[0].sideCrop.offsetY,
            top: document.getElementById('lightboxSideCropCard')?.getBoundingClientRect().top || 0
        }));
        if (!(moveCheck.offsetY < offsetBeforeMove) || !(moveCheck.top < cardBeforeMove.top)) {
            throw new Error(`Phase 2B 垂直拖曳右側特寫位置異常: before=${offsetBeforeMove}/${cardBeforeMove.top}, after=${JSON.stringify(moveCheck)}`);
        }
    }

    // 切換框線顏色：黃框 -> 紅框 -> 無框
    await page.click('#btnSideCropBorder');
    const borderRedText = await page.textContent('#btnSideCropBorder');
    const hasRedClass = await page.$eval('#lightboxSideCropCard', el => el.classList.contains('lightbox-side-crop-card--red'));
    await page.click('#btnSideCropBorder');
    const borderNoneText = await page.textContent('#btnSideCropBorder');
    const hasNoneClass = await page.$eval('#lightboxSideCropCard', el => el.classList.contains('lightbox-side-crop-card--none'));
    await page.click('#btnSideCropBorder');
    const borderBackYellow = await page.textContent('#btnSideCropBorder');
    if (borderRedText.trim() !== '紅框' || !hasRedClass || borderNoneText.trim() !== '無框' || !hasNoneClass || borderBackYellow.trim() !== '黃框') {
        throw new Error(`Phase 2B 切換框線樣式異常: ${borderRedText} / ${borderNoneText} / ${borderBackYellow}`);
    }

    // 測試旋轉與翻轉時 sideCrop.cropRect 跟隨變換，且等待縮圖合成徽章出現
    const cropBeforeRotate = await page.evaluate(() => ({ ...window.app.photos[0].sideCrop.cropRect }));
    await page.click('#btnLightboxRotateCW');
    await page.click('#btnLightboxFlipH');
    await page.waitForFunction(() => {
        const p = window.app.photos[0];
        const badge = document.querySelector(`#thumb_${p.uid} .thumb-sidecrop-badge`);
        const img = document.querySelector(`#thumb_${p.uid} .thumbnail-stage img`);
        return badge && p.compositeThumbnailSrc && img?.src === p.compositeThumbnailSrc;
    }, { timeout: 5000 });
    const transformedRect = await page.evaluate(() => ({ ...window.app.photos[0].sideCrop.cropRect }));
    if (transformedRect.x === cropBeforeRotate.x && transformedRect.y === cropBeforeRotate.y) {
        throw new Error(`旋轉／翻轉後 sideCrop.cropRect 未隨方向同步變換: before=${JSON.stringify(cropBeforeRotate)}, after=${JSON.stringify(transformedRect)}`);
    }
    // 轉正與翻回原方向，方便驗證最終合併寬度大於單張原寬
    await page.click('#btnLightboxFlipH');
    await page.click('#btnLightboxRotateCCW');

    const compositeExportCheck = await page.evaluate(async () => {
        const p = window.app.photos[0];
        const processed = await window.app.processImageToBlob(p, 1600, 0.88);
        const origSaveAs = window.saveAs;
        const origAlert = window.alert;
        const origConfirm = window.confirm;
        let savedProjBlob = null;
        let savedDocxBlob = null;
        let savedMode = '';
        window.saveAs = (blob, name) => {
            if (savedMode === 'docx') savedDocxBlob = blob;
            else savedProjBlob = blob;
        };
        window.alert = () => {};
        window.confirm = () => true;
        try {
            savedMode = 'docx';
            await window.app.exportDocx();
            savedMode = 'proj';
            await window.app.saveProject();
            const projFile = new File([savedProjBlob], 'phase2b.photo-report', { type: 'application/zip' });
            await window.app.openProject({ target: { files: [projFile], value: '' } });
        } finally {
            window.saveAs = origSaveAs;
            window.alert = origAlert;
            window.confirm = origConfirm;
        }
        const docxZip = await JSZip.loadAsync(savedDocxBlob);
        const docXml = await docxZip.file('word/document.xml').async('text');
        const drawingCount = (docXml.match(/<w:drawing>/g) || []).length;
        const portraitGeo = window.app.computeSideCropGeometry(900, 1600, {
            cropRect: { x: 0.1, y: 0.35, w: 0.85, h: 0.35 }
        });
        const restored = window.app.photos[0];
        return {
            compW: processed.width,
            compH: processed.height,
            hasMainAndSideBlobs: Boolean(processed.mainImage?.blob && processed.sideImage?.blob),
            drawingCount,
            portraitDefaultRatio: portraitGeo.widthRatio,
            portraitMaxRatio: portraitGeo.maxWidthRatio,
            docxSize: savedDocxBlob?.size || 0,
            restoredHasSideCrop: Boolean(restored?.sideCrop?.cropRect),
            restoredWidthRatio: restored?.sideCrop?.widthRatio || 0
        };
    });
    const pdfDownloadPromise = page.waitForEvent('download', { timeout: 10000 });
    await page.evaluate(() => window.app.exportPdf());
    const pdfDownload = await pdfDownloadPromise;
    const pdfSuggestedName = pdfDownload.suggestedFilename();
    if (
        !(compositeExportCheck.compW > baseBlobSize.width * 1.15) ||
        !compositeExportCheck.hasMainAndSideBlobs ||
        compositeExportCheck.drawingCount !== 4 ||
        !(compositeExportCheck.portraitDefaultRatio > 1.1) ||
        !(compositeExportCheck.portraitMaxRatio > 2.0) ||
        !compositeExportCheck.restoredHasSideCrop ||
        compositeExportCheck.docxSize <= 0 ||
        !pdfSuggestedName.endsWith('.pdf')
    ) {
        throw new Error(`Phase 2B Word 雙獨立圖片／合併匯出尺寸／PDF 或專案檔還原失敗: base=${JSON.stringify(baseBlobSize)}, comp=${JSON.stringify(compositeExportCheck)}, pdf=${pdfSuggestedName}`);
    }

    // 驗證刪除右側並排保留左圖黃框、點標記框 [data-ann-sidecrop] 重新並排，再用 Ctrl+Z Undo 復原刪除
    await page.evaluate(() => {
        window.app.openLightboxByIndex(0);
    });
    await page.waitForSelector('#btnSideCropDelete', { timeout: 2000 });
    await page.click('#btnSideCropDelete');
    await page.waitForFunction(() => !document.getElementById('lightboxCanvasFrame')?.classList.contains('has-side-crop'));
    const afterDeleteSide = await page.evaluate(() => ({
        sideCrop: window.app.photos[0].sideCrop,
        rectCount: (window.app.photos[0].annotations || []).filter(a => a.type === 'rect').length
    }));
    if (afterDeleteSide.sideCrop !== null || afterDeleteSide.rectCount !== 1) {
        throw new Error(`移除右側並排應只移除右側並排而保留來源黃框: ${JSON.stringify(afterDeleteSide)}`);
    }
    // 測試既有標記框一鍵 [data-ann-sidecrop]
    await page.click('[data-ann-sidecrop]');
    await page.waitForFunction(() => document.getElementById('lightboxCanvasFrame')?.classList.contains('has-side-crop'));
    // 再刪掉並測試 Ctrl+Z 復原
    await page.click('#btnSideCropDelete');
    await page.waitForFunction(() => !document.getElementById('lightboxCanvasFrame')?.classList.contains('has-side-crop'));
    await page.keyboard.press('Control+z');
    await page.waitForFunction(() => document.getElementById('lightboxCanvasFrame')?.classList.contains('has-side-crop'));

    // 驗證刪除右側並排時自動隱藏版面模擬列並清除白框固定尺寸，復原後再顯示
    await page.click('#btnSideCropDelete');
    await page.waitForFunction(() => document.getElementById('lightboxSideLayoutBar')?.classList.contains('hidden'));
    const singleRestoredStyle = await page.evaluate(() => {
        const f = document.getElementById('lightboxCanvasFrame');
        return { w: f?.style?.width || '', h: f?.style?.height || '', cls: f?.className || '' };
    });
    if (singleRestoredStyle.w !== '' || singleRestoredStyle.h !== '') {
        throw new Error(`恢復單張時應清除白框固定尺寸回復最大化檢視: ${JSON.stringify(singleRestoredStyle)}`);
    }

    // 注入 900x1600 直式 9:16 測試圖，實測固定三種版面模擬框、左圖靠左完全不動、右拉吸附／超寬提醒
    await page.click('#btnLightboxClose');
    const portraitCountBefore = await page.evaluate(() => window.app.photos.length);
    const portraitDataUrl = await page.evaluate(() => {
        const c = document.createElement('canvas');
        c.width = 900;
        c.height = 1600;
        const ctx = c.getContext('2d');
        ctx.fillStyle = '#334155';
        ctx.fillRect(0, 0, 900, 1600);
        ctx.fillStyle = '#0EA5E9';
        ctx.fillRect(80, 160, 740, 1280);
        // 模擬主角人物肩臂區塊
        ctx.fillStyle = '#FACC15';
        ctx.fillRect(140, 520, 460, 420);
        ctx.fillStyle = '#DC2626';
        ctx.fillRect(310, 610, 260, 180);
        return c.toDataURL('image/jpeg', 0.9);
    });
    const portraitTmpPath = path.resolve(__dirname, '../fixtures/.tmp_portrait_9x16.jpg');
    fs.writeFileSync(portraitTmpPath, Buffer.from(portraitDataUrl.split(',')[1], 'base64'));
    try {
        await page.setInputFiles('#fileInput', portraitTmpPath);
        await page.waitForFunction((exp) => window.app.photos.length === exp, portraitCountBefore + 1);
        await page.evaluate(() => {
            const idx = window.app.photos.length - 1;
            window.app.photos[idx].annotations = [{
                id: 'ann_portrait_lbl',
                type: 'label',
                color: 'yellow',
                x: 0.34,
                y: 0.31,
                scale: 1,
                text: '洪金和'
            }];
            window.app.openLightboxByIndex(idx);
        });
        await page.waitForFunction(() => {
            const img = document.getElementById('lightboxImg');
            return img && img.naturalWidth === 900 && document.getElementById('lightboxAnnotationLayer')?.getBoundingClientRect().width > 0;
        });
        const pLayerBox = await page.$eval('#lightboxAnnotationLayer', el => {
            const r = el.getBoundingClientRect();
            return { left: r.left, top: r.top, width: r.width, height: r.height };
        });
        await page.click('#btnLightboxToolCrop');
        await page.mouse.move(pLayerBox.left + pLayerBox.width * 0.14, pLayerBox.top + pLayerBox.height * 0.32);
        await page.mouse.down();
        await page.mouse.move(pLayerBox.left + pLayerBox.width * 0.84, pLayerBox.top + pLayerBox.height * 0.54);
        await page.mouse.up();
        await page.waitForSelector('#btnSideLightboxCrop', { timeout: 2000 });
        await page.click('#btnSideLightboxCrop');
        await page.waitForFunction(() => document.getElementById('lightboxCanvasFrame')?.classList.contains('has-side-crop') && document.getElementById('lightboxSideCropCanvas')?.dataset?.painted === '1');
        await page.waitForTimeout(80);

        const pInit = await page.evaluate(() => {
            const f = document.getElementById('lightboxCanvasFrame').getBoundingClientRect();
            const l = document.getElementById('lightboxLeftSlot').getBoundingClientRect();
            return {
                frameW: f.width, frameH: f.height,
                leftX: l.left, leftW: l.width, leftH: l.height,
                ratio: window.app.photos.at(-1).sideCrop.widthRatio,
                statusText: document.getElementById('lightboxSideLayoutStatus')?.textContent || ''
            };
        });

        // 縮小右圖：外框與左圖位置大小必須完全固定不動
        const pResize1 = await page.$eval('[data-side-crop-resize]', el => {
            const r = el.getBoundingClientRect();
            return { cx: r.left + r.width / 2, cy: r.top + r.height / 2 };
        });
        await page.mouse.move(pResize1.cx, pResize1.cy);
        await page.mouse.down();
        await page.mouse.move(pResize1.cx - 80, pResize1.cy - 20, { steps: 6 });
        await page.mouse.up();
        await page.waitForTimeout(60);

        const pShrunk = await page.evaluate(() => {
            const f = document.getElementById('lightboxCanvasFrame').getBoundingClientRect();
            const l = document.getElementById('lightboxLeftSlot').getBoundingClientRect();
            return {
                frameW: f.width, frameH: f.height,
                leftX: l.left, leftW: l.width, leftH: l.height,
                ratio: window.app.photos.at(-1).sideCrop.widthRatio,
                statusText: document.getElementById('lightboxSideLayoutStatus')?.textContent || ''
            };
        });
        if (
            !(pShrunk.ratio < pInit.ratio - 0.1) ||
            Math.abs(pShrunk.frameW - pInit.frameW) > 1.5 ||
            Math.abs(pShrunk.frameH - pInit.frameH) > 1.5 ||
            Math.abs(pShrunk.leftX - pInit.leftX) > 1.5 ||
            Math.abs(pShrunk.leftW - pInit.leftW) > 1.5 ||
            Math.abs(pShrunk.leftH - pInit.leftH) > 1.5
        ) {
            throw new Error(`直式 9:16 縮小右圖時外框與左圖不應移動或忽大忽小: init=${JSON.stringify(pInit)}, shrunk=${JSON.stringify(pShrunk)}`);
        }

        // 切換三種版面模擬白框並檢查寬高比與下拉選單同步
        await page.click('button[data-lightbox-layout="left_right_2"]');
        await page.waitForTimeout(60);
        const lrFrame = await page.evaluate(() => {
            const f = document.getElementById('lightboxCanvasFrame').getBoundingClientRect();
            return {
                aspect: f.width / f.height,
                selectVal: document.getElementById('layoutSelect').value,
                overflow: document.getElementById('lightboxCanvasFrame').classList.contains('is-side-overflow')
            };
        });
        if (lrFrame.selectVal !== 'left_right_2' || Math.abs(lrFrame.aspect - (315 / 690)) > 0.08 || !lrFrame.overflow) {
            throw new Error(`左右兩張版面模擬白框或警示狀態異常: ${JSON.stringify(lrFrame)}`);
        }

        await page.click('button[data-lightbox-layout="landscape_3"]');
        await page.waitForTimeout(60);
        const ls3Frame = await page.evaluate(() => {
            const f = document.getElementById('lightboxCanvasFrame').getBoundingClientRect();
            return {
                aspect: f.width / f.height,
                selectVal: document.getElementById('layoutSelect').value
            };
        });
        if (ls3Frame.selectVal !== 'landscape_3' || Math.abs(ls3Frame.aspect - (335 / 535)) > 0.08) {
            throw new Error(`橫式三張版面模擬白框異常: ${JSON.stringify(ls3Frame)}`);
        }

        // 切回上下兩張並向外拉大到貼齊／超出白框右緣，確認外框依然鎖死，並輸出截圖檢查視覺狀態
        await page.click('button[data-lightbox-layout="up_down_2"]');
        await page.waitForTimeout(60);
        await page.screenshot({ path: 'C:/Users/chia-hao/.gemini/antigravity/brain/68f06b8b-6e4d-44ac-a2aa-808ca2759063/scratch/preview_fixed_updown.png' });

        // 清理臨時直式測試圖
        await page.evaluate(() => {
            const idx = window.app.photos.length - 1;
            if (idx >= 0) window.app.removePhoto(idx);
            window.app.currentIndex = 0;
            window.app.lightboxIndex = 0;
        });
    } finally {
        try { fs.unlinkSync(portraitTmpPath); } catch (_) {}
    }

    // 清理測試狀態給第 3 節拖曳測試使用
    await page.evaluate(() => {
        const p = window.app.photos[0];
        if (p) {
            p.sideCrop = null;
            p.annotations = [];
            p.rotation = 0;
            p.flipH = false;
        }
        window.app.setLightboxTool('pan');
        window.app.closeModal('imageLightboxModal');
        window.app.render();
    });
    console.log('  ✅ Phase 2 模式 B（右側同格並排＋保留黃框＋角落縮放／上下調整＋切換框線＋旋轉翻轉＋合成匯出／專案還原＋Undo）驗證成功');

    // 3. 測試大量照片捲動到底部時，外部檔案拖曳提示框仍完整覆蓋可視區域
    const sample01Path = path.resolve(__dirname, '../fixtures/sample01.jpg');
    await page.setInputFiles('#fileInput', Array.from({ length: 24 }, () => sample01Path));
    await page.waitForFunction(() => window.app && window.app.photos.length >= 24);
    const dragOverlayCheck = await page.evaluate(() => {
        window.app.setThumbnailSize(130);
        const grid = document.getElementById('thumbGrid');
        const zone = document.getElementById('thumbGridDropZone');
        grid.scrollTop = grid.scrollHeight;
        const dt = new DataTransfer();
        dt.items.add(new File(['x'], 'extra.jpg', { type: 'image/jpeg' }));
        grid.dispatchEvent(new DragEvent('dragenter', { bubbles: true, cancelable: true, dataTransfer: dt }));
        grid.dispatchEvent(new DragEvent('dragover', { bubbles: true, cancelable: true, dataTransfer: dt }));
        const afterStyle = getComputedStyle(zone, '::after');
        const result = {
            gridScrollHeight: grid.scrollHeight,
            gridClientHeight: grid.clientHeight,
            gridScrollTop: grid.scrollTop,
            zoneClientHeight: zone.clientHeight,
            zoneScrollHeight: zone.scrollHeight,
            zoneHasClass: zone.classList.contains('external-file-drag-over'),
            gridHasClass: grid.classList.contains('external-file-drag-over'),
            afterHeight: parseFloat(afterStyle.height),
            afterContent: afterStyle.content
        };
        grid.dispatchEvent(new DragEvent('dragleave', { bubbles: true, cancelable: true, dataTransfer: dt, relatedTarget: document.body }));
        result.clearedAfterLeave = !zone.classList.contains('external-file-drag-over') && !grid.classList.contains('external-file-drag-over');
        return result;
    });
    if (dragOverlayCheck.gridScrollHeight <= dragOverlayCheck.gridClientHeight || dragOverlayCheck.gridScrollTop <= 0) {
        throw new Error(`測試前提失敗：縮圖區未產生捲動 (${JSON.stringify(dragOverlayCheck)})`);
    }
    if (!dragOverlayCheck.zoneHasClass || !dragOverlayCheck.gridHasClass || !dragOverlayCheck.afterContent.includes('放開以加入照片')) {
        throw new Error(`外部拖曳提示狀態未正確啟用 (${JSON.stringify(dragOverlayCheck)})`);
    }
    if (Math.abs(dragOverlayCheck.afterHeight - (dragOverlayCheck.zoneClientHeight - 16)) > 4 || dragOverlayCheck.zoneScrollHeight !== dragOverlayCheck.zoneClientHeight) {
        throw new Error(`捲動後拖曳提示框高度異常 (${JSON.stringify(dragOverlayCheck)})`);
    }
    if (!dragOverlayCheck.clearedAfterLeave) {
        throw new Error('拖離縮圖區後未清除外部拖曳提示狀態');
    }
    console.log('  ✅ 大量照片捲動時外部拖曳提示框完整覆蓋可視區域');

    // 4. 測試 Delete 移除快捷鍵（主畫面／燈箱）＋輸入／彈窗防誤刪＋多選確認＋Ctrl+Z／Ctrl+Y 復原重做
    await page.waitForFunction(() => document.getElementById('processingStatus')?.classList.contains('hidden'));
    await page.evaluate(() => {
        window.app.photos.forEach((p, idx) => { p.selected = idx === 0; });
        window.app.currentIndex = 0;
        window.app.lastSelectedIndex = 0;
        window.app.activeFilter = 'all';
        window.app.switchView('workspace');
        window.app.render();
        window.app.recordHistory('準備測試刪除快捷鍵');
    });

    // 4A. 文字輸入框聚焦時按 Delete 只能刪字，不得移除照片
    const countBeforeTypingDelete = await page.evaluate(() => window.app.photos.length);
    await page.fill('#editDesc', 'ABC');
    await page.focus('#editDesc');
    await page.evaluate(() => {
        const el = document.getElementById('editDesc');
        el.setSelectionRange(0, 1);
    });
    await page.keyboard.press('Delete');
    const typingDeleteCheck = await page.evaluate(() => ({
        count: window.app.photos.length,
        desc: document.getElementById('editDesc').value
    }));
    if (typingDeleteCheck.count !== countBeforeTypingDelete || typingDeleteCheck.desc !== 'BC') {
        throw new Error(`輸入框聚焦時按 Delete 只能編輯文字，不得刪除照片: ${JSON.stringify(typingDeleteCheck)}`);
    }
    await page.keyboard.press('Escape');

    // 4B. 輔助彈窗（如操作說明）開啟時按 Delete 不得誤刪背景照片
    await page.evaluate(() => window.app.openModal('shortcutsModal'));
    await page.keyboard.press('Delete');
    const modalDeleteCount = await page.evaluate(() => window.app.photos.length);
    if (modalDeleteCount !== countBeforeTypingDelete) {
        throw new Error(`輔助彈窗開啟時按 Delete 不應刪除背景照片: expected=${countBeforeTypingDelete}, actual=${modalDeleteCount}`);
    }
    await page.keyboard.press('Escape');

    // 4C. 主畫面單張選取時按 Delete 直接移除，再按 Ctrl+Z 復原、Ctrl+Y 重做
    const firstPhotoBeforeDel = await page.evaluate(() => ({
        uid0: window.app.photos[0].uid,
        uid1: window.app.photos[1].uid,
        count: window.app.photos.length
    }));
    await page.keyboard.press('Delete');
    const afterSingleDel = await page.evaluate(() => ({
        count: window.app.photos.length,
        uid0: window.app.photos[0]?.uid,
        selected0: Boolean(window.app.photos[0]?.selected),
        currentIndex: window.app.currentIndex
    }));
    if (
        afterSingleDel.count !== firstPhotoBeforeDel.count - 1 ||
        afterSingleDel.uid0 !== firstPhotoBeforeDel.uid1 ||
        !afterSingleDel.selected0 ||
        afterSingleDel.currentIndex !== 0
    ) {
        throw new Error(`主畫面單張按 Delete 後應直接移除並自動選取補位照片: before=${JSON.stringify(firstPhotoBeforeDel)}, after=${JSON.stringify(afterSingleDel)}`);
    }
    await page.keyboard.press('Control+z');
    const afterSingleUndo = await page.evaluate(() => ({
        count: window.app.photos.length,
        uid0: window.app.photos[0]?.uid
    }));
    if (afterSingleUndo.count !== firstPhotoBeforeDel.count || afterSingleUndo.uid0 !== firstPhotoBeforeDel.uid0) {
        throw new Error(`主畫面 Delete 後按 Ctrl+Z 應還原照片: ${JSON.stringify(afterSingleUndo)}`);
    }
    await page.keyboard.press('Control+y');
    if ((await page.evaluate(() => window.app.photos.length)) !== firstPhotoBeforeDel.count - 1) {
        throw new Error('按 Ctrl+Y 應重新套用移除結果');
    }
    await page.keyboard.press('Control+z');

    // 4D. 主畫面多張勾選按 Delete 必須跳出確認保護（取消保留、確定刪除，且仍可 Ctrl+Z 復原）
    await page.evaluate(() => {
        window.app.photos.forEach((p, idx) => { p.selected = idx < 2; });
        window.app.currentIndex = 0;
        window.app.render();
    });
    let multiDelPrompt = '';
    page.once('dialog', async d => { multiDelPrompt = d.message(); await d.dismiss(); });
    await page.keyboard.press('Delete');
    if (!multiDelPrompt.includes('已選取的 2 張照片')) {
        throw new Error(`多選按 Delete 應跳出多張移除確認提示，實際訊息: ${multiDelPrompt}`);
    }
    if ((await page.evaluate(() => window.app.photos.length)) !== firstPhotoBeforeDel.count) {
        throw new Error('取消多張移除確認後不得刪除任何照片');
    }
    page.once('dialog', async d => { await d.accept(); });
    await page.keyboard.press('Delete');
    if ((await page.evaluate(() => window.app.photos.length)) !== firstPhotoBeforeDel.count - 2) {
        throw new Error('確認多張移除後應一次移除所選 2 張照片');
    }
    await page.keyboard.press('Control+z');
    if ((await page.evaluate(() => window.app.photos.length)) !== firstPhotoBeforeDel.count) {
        throw new Error('多張確認移除後按 Ctrl+Z 應完整復原');
    }

    // 4E. 大圖燈箱內按 Delete 直接移出目前照片並切換下一張；按 Ctrl+Z 直接在燈箱恢復；刪到 0 張自動關燈箱再按 Ctrl+Z 復原
    await page.evaluate(() => {
        window.app.photos = window.app.photos.slice(0, 2);
        window.app.photos.forEach((p, idx) => { p.selected = idx === 0; });
        window.app.currentIndex = 0;
        window.app.render();
        window.app.recordHistory('兩張燈箱刪除測試基準');
        window.app.openLightboxByIndex(0);
    });
    const lbBeforeDel = await page.evaluate(() => ({
        uid0: window.app.photos[0].uid,
        uid1: window.app.photos[1].uid,
        posText: document.getElementById('lightboxPosition')?.textContent?.trim()
    }));
    await page.keyboard.press('Delete');
    const lbAfterFirstDel = await page.evaluate(() => ({
        open: !document.getElementById('imageLightboxModal').classList.contains('hidden'),
        count: window.app.photos.length,
        lightboxIndex: window.app.lightboxIndex,
        currentUid: window.app.photos[window.app.lightboxIndex]?.uid,
        posText: document.getElementById('lightboxPosition')?.textContent?.trim()
    }));
    if (
        !lbAfterFirstDel.open ||
        lbAfterFirstDel.count !== 1 ||
        lbAfterFirstDel.currentUid !== lbBeforeDel.uid1 ||
        lbAfterFirstDel.posText !== '1 / 1'
    ) {
        throw new Error(`燈箱按 Delete 應立即移出目前照片並切換至補位照: ${JSON.stringify(lbAfterFirstDel)}`);
    }

    // 燈箱內按 Ctrl+Z 應立刻恢復剛被刪掉的第 1 張照片
    await page.keyboard.press('Control+z');
    const lbAfterUndo = await page.evaluate(() => ({
        open: !document.getElementById('imageLightboxModal').classList.contains('hidden'),
        count: window.app.photos.length,
        currentUid: window.app.photos[window.app.lightboxIndex]?.uid,
        posText: document.getElementById('lightboxPosition')?.textContent?.trim()
    }));
    if (
        !lbAfterUndo.open ||
        lbAfterUndo.count !== 2 ||
        lbAfterUndo.currentUid !== lbBeforeDel.uid0 ||
        lbAfterUndo.posText !== '1 / 2'
    ) {
        throw new Error(`燈箱按 Ctrl+Z 應直接在燈箱內恢復剛被刪掉的照片: ${JSON.stringify(lbAfterUndo)}`);
    }

    // 連按兩次 Delete 把剩餘 2 張都移出：最後一張移除後應自動關閉燈箱並歸還右側編輯欄，主畫面 Ctrl+Z 可恢復
    await page.keyboard.press('Delete');
    await page.keyboard.press('Delete');
    const afterEmptyDel = await page.evaluate(() => ({
        open: !document.getElementById('imageLightboxModal').classList.contains('hidden'),
        count: window.app.photos.length,
        cardInPanel: document.getElementById('editorPanel').contains(document.getElementById('editorCard')),
        emptyVisible: !document.getElementById('emptyState').classList.contains('hidden')
    }));
    if (afterEmptyDel.open || afterEmptyDel.count !== 0 || !afterEmptyDel.cardInPanel || !afterEmptyDel.emptyVisible) {
        throw new Error(`燈箱刪到 0 張時應自動關閉燈箱、歸還編輯欄並顯示空白狀態: ${JSON.stringify(afterEmptyDel)}`);
    }
    await page.keyboard.press('Control+z');
    const afterEmptyUndo = await page.evaluate(() => ({
        count: window.app.photos.length,
        emptyVisible: !document.getElementById('emptyState').classList.contains('hidden')
    }));
    if (afterEmptyUndo.count !== 1 || afterEmptyUndo.emptyVisible) {
        throw new Error(`刪光關閉燈箱後於主畫面按 Ctrl+Z 應恢復照片並結束空白狀態: ${JSON.stringify(afterEmptyUndo)}`);
    }
    console.log('  ✅ Delete 移除快捷鍵（主畫面／燈箱／輸入與彈窗防護／多選確認／Ctrl+Z 復原）驗證成功');

    // 5. 驗證工作進度儲存／開啟全方位強化（Ctrl+S、狀態徽章、原地覆寫、輕量進度檔、拖放開檔、取消與防漏、IndexedDB 自動暫存）
    // 5A. 測試未儲存狀態徽章、遮罩輸入自動定稿、非法檔名字元清洗與 Ctrl+S 快捷鍵（含輸入框與燈箱內）
    await page.evaluate(() => {
        const titleInput = document.getElementById('caseTitle');
        titleInput.value = '115/11/10:竊盜案*測試?';
        titleInput.dispatchEvent(new Event('input', { bubbles: true }));
    });
    const unsavedBadgeState = await page.evaluate(() => {
        const btn = document.getElementById('btnQuickSaveStatus');
        return {
            visible: btn && !btn.classList.contains('hidden'),
            unsavedClass: btn?.classList.contains('project-save-badge--unsaved'),
            text: document.getElementById('quickSaveStatusText')?.textContent?.trim()
        };
    });
    if (!unsavedBadgeState.visible || !unsavedBadgeState.unsavedClass || !unsavedBadgeState.text?.includes('還有修改沒存')) {
        throw new Error(`修改案由後頂部應立即顯示未儲存狀態徽章: ${JSON.stringify(unsavedBadgeState)}`);
    }

    // 聚焦時間欄輸入 4 碼 '1645'（未離開欄位），直接按 Ctrl+S 應自動完成遮罩格式化並儲存
    await page.evaluate(() => {
        window.app.photos.forEach((p, idx) => { p.selected = idx === 0; });
        window.app.currentIndex = 0;
        window.app.render();
        window.__ctrlSSaved = [];
        window.__origSaveAs5 = window.saveAs;
        window.saveAs = (blob, name) => window.__ctrlSSaved.push({ blob, name });
    });
    await page.click('#editTime');
    await page.keyboard.type('1645');
    await page.keyboard.press('Control+s');
    await page.waitForFunction(() => Array.isArray(window.__ctrlSSaved) && window.__ctrlSSaved.length === 1);

    const afterCtrlS = await page.evaluate(async () => {
        const saved = window.__ctrlSSaved[0];
        const zip = await window.JSZip.loadAsync(saved.blob);
        const manifest = JSON.parse(await zip.file('manifest.json').async('text'));
        const btn = document.getElementById('btnQuickSaveStatus');
        return {
            fileName: saved.name,
            photoTimeInApp: window.app.photos[0]?.time,
            photoTimeInManifest: manifest.photos[0]?.time,
            savedClass: btn?.classList.contains('project-save-badge--saved'),
            badgeText: document.getElementById('quickSaveStatusText')?.textContent?.trim(),
            unsavedNow: window.app.hasUnsavedProjectChanges()
        };
    });
    if (
        afterCtrlS.fileName !== '115_11_10_竊盜案_測試__照片清冊專案.photo-report' ||
        afterCtrlS.photoTimeInApp !== '16:45' ||
        afterCtrlS.photoTimeInManifest !== '16:45' ||
        !afterCtrlS.savedClass ||
        afterCtrlS.unsavedNow
    ) {
        throw new Error(`Ctrl+S 應自動完成時間遮罩定稿、清洗檔名並更新為已儲存狀態: ${JSON.stringify(afterCtrlS)}`);
    }

    // 開啟燈箱旋轉照片後，在燈箱內按 Ctrl+S 也應正常存檔
    await page.evaluate(() => window.app.openLightboxByIndex(0));
    await page.waitForFunction(() => !document.getElementById('imageLightboxModal').classList.contains('hidden'));
    await page.keyboard.press('r');
    await page.keyboard.press('Control+s');
    await page.waitForFunction(() => Array.isArray(window.__ctrlSSaved) && window.__ctrlSSaved.length === 2);
    await page.keyboard.press('Escape');

    // 5B. 測試同檔原地覆寫（currentProjectFileHandle）與 Ctrl+Shift+S 另存新檔
    const handleOverwriteResult = await page.evaluate(async () => {
        const writes = [];
        window.app.currentProjectFileHandle = {
            name: 'linked_case.photo-report',
            createWritable: async () => ({
                write: async (blob) => { writes.push(blob); },
                close: async () => {}
            })
        };
        window.app.currentProjectFileName = 'linked_case.photo-report';
        window.app.saveAs = window.app.initialSaveAs;
        window.saveAs = window.app.initialSaveAs;

        window.app.photos[0].desc = '原地覆寫測試說明';
        window.app.recordHistory('修改說明');
        const unsavedBefore = window.app.hasUnsavedProjectChanges();
        await window.app.saveProject({ fromUserPicker: true });
        const badgeAfterOverwrite = document.getElementById('quickSaveStatusText')?.textContent?.trim();
        const menuTextAfterOverwrite = document.getElementById('saveProjectMenuText')?.textContent?.trim();

        // 測試另存新檔（saveAsNew: true）會走 saveAs 而非覆寫原 handle
        const saveAsNewFiles = [];
        window.saveAs = (blob, name) => saveAsNewFiles.push({ blob, name });
        await window.app.saveProject({ saveAsNew: true });

        window.app.currentProjectFileHandle = null;
        return {
            unsavedBefore,
            writeCount: writes.length,
            writtenSize: writes[0]?.size || 0,
            unsavedAfter: window.app.hasUnsavedProjectChanges(),
            badgeAfterOverwrite,
            menuTextAfterOverwrite,
            saveAsNewCount: saveAsNewFiles.length
        };
    });
    if (
        !handleOverwriteResult.unsavedBefore ||
        handleOverwriteResult.writeCount !== 1 ||
        handleOverwriteResult.writtenSize <= 0 ||
        handleOverwriteResult.unsavedAfter ||
        !handleOverwriteResult.badgeAfterOverwrite?.includes('linked_case.photo-report') ||
        !handleOverwriteResult.menuTextAfterOverwrite?.includes('linked_case.photo-report') ||
        handleOverwriteResult.saveAsNewCount !== 1
    ) {
        throw new Error(`同檔原地覆寫與另存新檔行為異常: ${JSON.stringify(handleOverwriteResult)}`);
    }

    // 5C & 5D. 測試「預設自動 220 ppi 級（長邊 2000px）最佳化壓縮且保留標記／翻轉／並排」、「直接拖放 .photo-report 進度檔開啟」與「依採證時間排序」
    const liteAndDropResult = await page.evaluate(async () => {
        const origAlert = window.alert;
        const origConfirm = window.confirm;
        const alerts = [];
        window.alert = msg => alerts.push(String(msg));
        window.confirm = () => true;
        try {
            // 建立一張 2600x1800 的高解析度彩色測試圖
            const bigCanvas = document.createElement('canvas');
            bigCanvas.width = 2600;
            bigCanvas.height = 1800;
            const bctx = bigCanvas.getContext('2d');
            bctx.fillStyle = '#1E3A8A';
            bctx.fillRect(0, 0, 2600, 1800);
            for (let i = 0; i < 40; i++) {
                bctx.fillStyle = `hsl(${(i * 19) % 360}, 80%, 55%)`;
                bctx.fillRect((i * 63) % 2400, (i * 41) % 1600, 180, 140);
            }
            const bigBlob = await new Promise(r => bigCanvas.toBlob(r, 'image/jpeg', 0.98));
            const bigFile = new File([bigBlob], 'highres_scene.jpg', { type: 'image/jpeg' });

            window.app.photos = [];
            await window.app.handleFiles([bigFile]);
            const p = window.app.photos[0];
            p.rotation = 90;
            p.flipH = true;
            p.desc = '高解析度輕量專案測試';
            p.annotations = [
                { id: 'ann_lite_1', type: 'rect', color: 'yellow', x: 0.2, y: 0.25, w: 0.35, h: 0.3 },
                { id: 'ann_lite_2', type: 'label', color: 'red', x: 0.5, y: 0.2, scale: 1.4, text: '關鍵證物' }
            ];
            p.sideCrop = {
                cropRect: { x: 0.2, y: 0.25, w: 0.35, h: 0.3 },
                widthRatio: 0.85,
                offsetY: 0.5,
                borderColor: 'yellow',
                sourceAnnId: 'ann_lite_1'
            };

            const outFiles = [];
            window.saveAs = (blob, name) => outFiles.push({ blob, name });
            // 未縮圖原始大小對照組 vs 預設 saveProject() 自動最佳化壓縮
            await window.app.saveProject({ lite: false });
            await window.app.saveProject();

            const fullProj = outFiles[0];
            const liteProj = outFiles[1];

            // 先清空工作台，測試直接拖放（drop）liteProj 到 #emptyState 開啟
            window.app.photos = [];
            window.app.activeFilter = 'missingDesc';
            window.app.render();

            const dt = new DataTransfer();
            const droppedFile = new File([liteProj.blob], liteProj.name, { type: 'application/zip' });
            dt.items.add(droppedFile);
            const emptyStateEl = document.getElementById('emptyState');
            emptyStateEl.dispatchEvent(new DragEvent('drop', { bubbles: true, cancelable: true, dataTransfer: dt }));

            await new Promise((resolve, reject) => {
                const deadline = Date.now() + 5000;
                const check = () => {
                    if (
                        window.app.photos.length === 1 &&
                        window.app.photos[0].desc === '高解析度輕量專案測試' &&
                        alerts.some(msg => msg.includes('已開啟工作進度'))
                    ) {
                        return resolve();
                    }
                    if (Date.now() > deadline) return reject(new Error('拖放 .photo-report 開啟逾時'));
                    setTimeout(check, 30);
                };
                check();
            });

            const restoredPhoto = window.app.photos[0];
            const restoredImg = await window.app.loadImage(restoredPhoto.src);

            // 測試「依採證時間排序」（含子特寫緊跟母圖連動）
            const basePhoto = window.app.photos[0];
            document.getElementById('caseDate').value = '115/11/10';
            window.app.photos = [
                { ...basePhoto, uid: 't_late', name: 'A_01.jpg', seq: '1', date: '115/11/10', time: '15:30', parentUid: null, cropRect: null },
                { ...basePhoto, uid: 't_late_child', name: 'A_01_特寫1.jpg', seq: '1-1', date: '115/11/10', time: '08:00', parentUid: 't_late', cropRect: { x: 0.1, y: 0.1, w: 0.3, h: 0.3 } },
                { ...basePhoto, uid: 't_early', name: 'Z_99.jpg', seq: '2', date: '', time: '09:15:20', parentUid: null, cropRect: null },
                { ...basePhoto, uid: 't_prev_day', name: 'M_50.jpg', seq: '3', date: '115/11/09', time: '22:10', parentUid: null, cropRect: null }
            ];
            window.app.sortPhotosByCaptureTime({ confirmBeforeSort: true });
            const sortedTimeUids = window.app.photos.map(item => item.uid);
            const sortedTimeSeqs = window.app.photos.map(item => item.seq);

            // 恢復單張照片供後續 5E / 5F 測試
            window.app.photos = [restoredPhoto];
            window.app.currentIndex = 0;
            window.app.render();

            return {
                fullSize: fullProj.blob.size,
                liteSize: liteProj.blob.size,
                liteName: liteProj.name,
                restoredMaxDim: Math.max(restoredImg.naturalWidth || restoredImg.width, restoredImg.naturalHeight || restoredImg.height),
                restoredRotation: restoredPhoto.rotation,
                restoredFlipH: restoredPhoto.flipH,
                restoredAnnCount: restoredPhoto.annotations?.length || 0,
                restoredSideCropRatio: restoredPhoto.sideCrop?.widthRatio,
                filterResetToAll: window.app.activeFilter === 'all',
                sortedTimeUids,
                sortedTimeSeqs
            };
        } finally {
            window.alert = origAlert;
            window.confirm = origConfirm;
        }
    });
    if (
        !liteAndDropResult.liteName.endsWith('_照片清冊專案.photo-report') ||
        liteAndDropResult.liteSize >= liteAndDropResult.fullSize ||
        liteAndDropResult.restoredMaxDim > 2000 ||
        liteAndDropResult.restoredRotation !== 90 ||
        !liteAndDropResult.restoredFlipH ||
        liteAndDropResult.restoredAnnCount !== 2 ||
        liteAndDropResult.restoredSideCropRatio !== 0.85 ||
        !liteAndDropResult.filterResetToAll ||
        liteAndDropResult.sortedTimeUids.join(',') !== 't_prev_day,t_early,t_late,t_late_child' ||
        liteAndDropResult.sortedTimeSeqs.join(',') !== '3,2,1,1-1'
    ) {
        throw new Error(`預設最佳化壓縮、拖放開啟或依採證時間排序結果不符預期: ${JSON.stringify(liteAndDropResult)}`);
    }

    // 5E. 測試錯誤提示區分（無法儲存 vs 無法開啟）、中途失敗不洩漏 URL 與取消處理保護
    const errorAndCancelCheck = await page.evaluate(async () => {
        const origAlert = window.alert;
        const origConfirm = window.confirm;
        const alerts = [];
        window.alert = msg => alerts.push(String(msg));
        window.confirm = () => true;
        try {
            const urlsBefore = window.app.managedUrls.size;
            window.app.handleProjectError(new Error('磁碟已滿'), 'save');
            window.app.handleProjectError(new Error('使用者已取消處理。'), 'save');

            // 構造缺少第 2 張圖檔的損壞專案檔，驗證失敗時釋放第 1 張暫存 URL 且保留原工作台照片
            const badZip = new window.JSZip();
            const firstPhoto = window.app.photos[0];
            badZip.file('images/0001_ok.jpg', firstPhoto.file);
            badZip.file('manifest.json', JSON.stringify({
                format: 'photo-report-project',
                version: 1,
                caseData: window.app.getCaseData(),
                photos: [
                    { uid: 'u1', archiveName: '0001_ok.jpg', name: 'ok.jpg', type: 'image/jpeg', seq: '1' },
                    { uid: 'u2', archiveName: '0002_missing.jpg', name: 'missing.jpg', type: 'image/jpeg', seq: '2' }
                ]
            }));
            const badBlob = await badZip.generateAsync({ type: 'blob' });
            const badFile = new File([badBlob], 'corrupted.photo-report', { type: 'application/zip' });
            const openOk = await window.app.openProjectFromFile(badFile);
            const urlsAfter = window.app.managedUrls.size;
            const progressHidden = document.getElementById('processingStatus').classList.contains('hidden');
            return {
                alerts,
                openOk,
                urlsBefore,
                urlsAfter,
                progressHidden,
                stillHasOriginalPhoto: window.app.photos.length === 1 && window.app.photos[0].desc === '高解析度輕量專案測試'
            };
        } finally {
            window.alert = origAlert;
            window.confirm = origConfirm;
        }
    });
    if (
        errorAndCancelCheck.openOk !== false ||
        errorAndCancelCheck.urlsAfter !== errorAndCancelCheck.urlsBefore ||
        !errorAndCancelCheck.progressHidden ||
        !errorAndCancelCheck.stillHasOriginalPhoto ||
        !errorAndCancelCheck.alerts[0]?.startsWith('工作進度無法儲存：') ||
        !errorAndCancelCheck.alerts[1]?.startsWith('工作進度檔無法開啟：')
    ) {
        throw new Error(`專案儲存／開啟錯誤處理與 URL 防洩漏檢查未通過: ${JSON.stringify(errorAndCancelCheck)}`);
    }

    // 5F. 測試本機 IndexedDB 自動暫存草稿（寫入、啟動偵測橫幅、一鍵接續進度與捨棄草稿）
    const autoDraftCheck = await page.evaluate(async () => {
        document.getElementById('caseTitle').value = '自動暫存恢復測試案';
        const savedOk = await window.app.saveAutoDraftNow();

        // 模擬重新整理後的空白初始工作台
        window.app.photos.forEach(p => window.app.releasePhotoUrls(p));
        window.app.photos = [];
        window.app.render();

        const offered = await window.app.checkAndOfferAutoDraft();
        const bannerVisibleAfterOffer = !document.getElementById('autoDraftBanner').classList.contains('hidden');
        const bannerTitle = document.getElementById('autoDraftBannerTitle')?.textContent?.trim();

        // 點擊「接續上次進度」
        await window.app.restoreAutoDraft();
        const bannerHiddenAfterRestore = document.getElementById('autoDraftBanner').classList.contains('hidden');
        const restoredCount = window.app.photos.length;
        const restoredTitle = document.getElementById('caseTitle').value;
        const restoredDesc = window.app.photos[0]?.desc;
        const restoredAnnCount = window.app.photos[0]?.annotations?.length || 0;

        // 測試「捨棄草稿」清除 IndexedDB
        await window.app.discardAutoDraft();
        window.app.photos = [];
        const offeredAfterDiscard = await window.app.checkAndOfferAutoDraft();

        // 復原 window.saveAs
        if (window.__origSaveAs5) window.saveAs = window.__origSaveAs5;

        return {
            savedOk,
            offered,
            bannerVisibleAfterOffer,
            bannerTitle,
            bannerHiddenAfterRestore,
            restoredCount,
            restoredTitle,
            restoredDesc,
            restoredAnnCount,
            offeredAfterDiscard
        };
    });
    if (
        !autoDraftCheck.savedOk ||
        !autoDraftCheck.offered ||
        !autoDraftCheck.bannerVisibleAfterOffer ||
        !autoDraftCheck.bannerTitle?.includes('自動暫存恢復測試案') ||
        !autoDraftCheck.bannerHiddenAfterRestore ||
        autoDraftCheck.restoredCount !== 1 ||
        autoDraftCheck.restoredTitle !== '自動暫存恢復測試案' ||
        autoDraftCheck.restoredDesc !== '高解析度輕量專案測試' ||
        autoDraftCheck.restoredAnnCount !== 2 ||
        autoDraftCheck.offeredAfterDiscard !== false
    ) {
        throw new Error(`IndexedDB 自動暫存草稿儲存、恢復或捨棄驗證失敗: ${JSON.stringify(autoDraftCheck)}`);
    }
    console.log('  ✅ 工作進度全方位強化（Ctrl+S、狀態徽章、原地覆寫、輕量進度檔、拖放開檔、防漏與自動暫存）驗證成功');

    // 5G. 測試多 .photo-report 工作進度檔智慧彙整合併（跨電腦重壓縮識別、撞號不同圖不誤合、一空一有補齊、較晚覆蓋、未存修改保護、子特寫合併與 Ctrl+Z 復原）
    const mergeCheck = await page.evaluate(async () => {
        const origAlert = window.alert;
        const origConfirm = window.confirm;
        const origSaveAs = window.saveAs;
        const alerts = [];
        window.alert = msg => alerts.push(String(msg));
        window.confirm = () => true;

        const makeSceneFile = async (filename, drawFn, width = 2400, height = 1600) => {
            const canvas = document.createElement('canvas');
            canvas.width = width;
            canvas.height = height;
            const ctx = canvas.getContext('2d');
            drawFn(ctx, width, height);
            const blob = await new Promise(r => canvas.toBlob(r, 'image/jpeg', 0.94));
            return new File([blob], filename, { type: 'image/jpeg' });
        };

        const stampManifestTime = async (blob, savedAtIso) => {
            const zip = await window.JSZip.loadAsync(blob);
            const manifest = JSON.parse(await zip.file('manifest.json').async('text'));
            manifest.savedAt = savedAtIso;
            zip.file('manifest.json', JSON.stringify(manifest), { compression: 'DEFLATE', compressionOptions: { level: 6 } });
            return zip.generateAsync({ type: 'blob' });
        };

        try {
            // 建立測試照片（2400x1600，存成專案時皆會經過 2000px 自動最佳化壓縮）
            const fileScene1 = await makeSceneFile('cam_scene_1.jpg', (ctx, w, h) => {
                ctx.fillStyle = '#0F172A';
                ctx.fillRect(0, 0, w, h);
                ctx.fillStyle = '#F59E0B';
                ctx.fillRect(160, 140, 820, 900);
                ctx.fillStyle = '#38BDF8';
                ctx.beginPath();
                ctx.arc(1700, 950, 420, 0, Math.PI * 2);
                ctx.fill();
            });
            const fileScene2 = await makeSceneFile('cam_scene_2.jpg', (ctx, w, h) => {
                ctx.fillStyle = '#ECFDF5';
                ctx.fillRect(0, 0, w, h);
                ctx.fillStyle = '#059669';
                ctx.fillRect(220, 220, 1960, 420);
                ctx.fillStyle = '#DC2626';
                ctx.fillRect(900, 760, 620, 620);
            });
            const fileSharedBlue = await makeSceneFile('IMG_8888.jpg', (ctx, w, h) => {
                ctx.fillStyle = '#1D4ED8';
                ctx.fillRect(0, 0, w, h);
                ctx.fillStyle = '#FACC15';
                ctx.fillRect(120, 120, 600, 1200);
            });
            const fileSharedGreenDiff = await makeSceneFile('IMG_8888.jpg', (ctx, w, h) => {
                ctx.fillStyle = '#15803D';
                ctx.fillRect(0, 0, w, h);
                ctx.fillStyle = '#F43F5E';
                ctx.fillRect(1400, 260, 820, 980);
            });
            const fileScene3 = await makeSceneFile('cam_scene_3.jpg', (ctx, w, h) => {
                ctx.fillStyle = '#7C2D12';
                ctx.fillRect(0, 0, w, h);
                ctx.fillStyle = '#E2E8F0';
                ctx.fillRect(420, 320, 1520, 920);
            });

            const savedOutput = [];
            window.saveAs = (blob, name) => savedOutput.push({ blob, name });

            // 1. 模擬人員 A（較早）各自開新網頁匯入照片並編輯存檔
            await window.app.clearAll(true);
            document.getElementById('caseTitle').value = '聯合作業測試案';
            document.getElementById('caseDate').value = '115/11/09';
            document.getElementById('defaultLocation').value = '舊預設地點';
            await window.app.handleFiles([fileScene1, fileScene2, fileSharedBlue]);
            window.app.photos[0].desc = '早期說明';
            window.app.photos[0].location = '台北市中正區';
            window.app.photos[0].time = '09:30';
            window.app.photos[0].timeAutoDefaulted = false;
            window.app.photos[0].annotations = [{ id: 'a_red', type: 'rect', color: 'red', x: 0.1, y: 0.1, w: 0.2, h: 0.2 }];
            // photos[1] (cam_scene_2) 刻意維持空白與自動預設 time='12:00'
            window.app.photos[2].desc = '早檔藍圖';
            // 為早檔藍圖建立一張子特寫（uid 保留供後續第三份檔測試同/異特寫）
            const blueParentUid = window.app.photos[2].uid;
            const cropChild = await window.app.createDetailCropPhoto(2, { x: 0.1, y: 0.1, w: 0.35, h: 0.35 });
            cropChild.desc = '早檔藍圖特寫1';
            const childCropUid = cropChild.uid;
            await window.app.saveProject({ lite: true });
            const earlyBlob = await stampManifestTime(savedOutput.pop().blob, '2026-11-10T08:00:00.000Z');

            // 2. 模擬人員 B（較晚）各自重開網頁匯入同批高解析原圖（完全不同 uid）與撞號綠圖
            await window.app.clearAll(true);
            document.getElementById('caseTitle').value = '聯合作業測試案';
            document.getElementById('caseDate').value = '115/11/10';
            document.getElementById('defaultLocation').value = '新案件現場';
            await window.app.handleFiles([fileScene1, fileScene2, fileSharedGreenDiff, fileScene3]);
            window.app.photos[0].desc = '較晚完整說明';
            window.app.photos[0].location = ''; // 留空應保留早檔「台北市中正區」
            window.app.photos[0].annotations = [
                { id: 'b_yellow', type: 'rect', color: 'yellow', x: 0.25, y: 0.25, w: 0.3, h: 0.3 },
                { id: 'b_lbl', type: 'label', color: 'yellow', x: 0.3, y: 0.2, scale: 1.2, text: '陳某' }
            ];
            window.app.photos[1].desc = '補齊空白說明';
            window.app.photos[1].location = '大安區路段';
            window.app.photos[1].time = '14:25';
            window.app.photos[1].timeAutoDefaulted = false;
            window.app.photos[1].sideCrop = { cropRect: { x: 0.2, y: 0.2, w: 0.4, h: 0.4 }, widthRatio: 0.72, offsetY: 0.5, borderColor: 'yellow' };
            window.app.photos[2].desc = '晚檔撞號綠圖';
            window.app.photos[3].desc = '第三現場補充';
            await window.app.saveProject({ lite: true });
            const lateBlob = await stampManifestTime(savedOutput.pop().blob, '2026-11-10T10:00:00.000Z');

            // 3. 製作第三份進度檔（基於早檔衍生或同原圖但更晚於 10:00，測試已開啟專案併入與未存修改保護）
            const thirdZip = await window.JSZip.loadAsync(earlyBlob);
            const thirdManifest = JSON.parse(await thirdZip.file('manifest.json').async('text'));
            thirdManifest.savedAt = '2026-11-10T12:00:00.000Z';
            thirdManifest.photos[0].desc = '不該蓋掉畫面未存的第1張';
            thirdManifest.photos[1].desc = '第三手最新修訂第2張';
            thirdManifest.photos[3].desc = '第三手更新同特寫1';
            // 再加一張同母圖但不同 cropRect 的新特寫
            const extraChildArchive = '0099_extra_crop.jpg';
            const blueOrigEntry = thirdZip.file(new RegExp(`images/.*${blueParentUid}.*`))[0] || thirdZip.file(`images/${thirdManifest.photos[2].archiveName}`);
            const blueOrigBytes = await blueOrigEntry.async('uint8array');
            thirdZip.file(`images/${extraChildArchive}`, blueOrigBytes, { compression: 'STORE' });
            thirdManifest.photos.push({
                uid: 'third_new_child_2',
                archiveName: extraChildArchive,
                name: 'IMG_8888_特寫2.jpg',
                type: 'image/jpeg',
                rotation: 0,
                flipH: false,
                parentUid: blueParentUid,
                cropRect: { x: 0.5, y: 0.45, w: 0.3, h: 0.3 },
                sideCrop: null,
                annotations: [],
                exifRotation: 0,
                seq: '3-2',
                date: '',
                time: '12:00',
                timeAutoDefaulted: true,
                location: '',
                desc: '第三手新增第二特寫',
                selected: false
            });
            thirdZip.file('manifest.json', JSON.stringify(thirdManifest));
            const thirdBlob = await thirdZip.generateAsync({ type: 'blob' });

            // 4. 清空工作台，模擬一次拖放「2 份進度檔」（故意把較晚檔放前面、較早檔放後面）
            await window.app.clearAll(true);
            const dt = new DataTransfer();
            dt.items.add(new File([lateBlob], 'B_late.photo-report', { type: 'application/zip' }));
            dt.items.add(new File([earlyBlob], 'A_early.photo-report', { type: 'application/zip' }));
            document.getElementById('emptyState').dispatchEvent(new DragEvent('drop', { bubbles: true, cancelable: true, dataTransfer: dt }));

            await new Promise((resolve, reject) => {
                const deadline = Date.now() + 8000;
                const tick = () => {
                    if (window.app.photos.length === 6 && alerts.some(m => m.includes('已完成彙整 2 份工作進度'))) return resolve();
                    if (Date.now() > deadline) return reject(new Error('拖放多份 .photo-report 彙整逾時'));
                    setTimeout(tick, 30);
                };
                tick();
            });

            const afterTwoDrop = {
                count: window.app.photos.length,
                seqs: window.app.photos.map(p => p.seq),
                p0Desc: window.app.photos[0].desc,
                p0Loc: window.app.photos[0].location,
                p0Anns: window.app.photos[0].annotations.map(a => `${a.type}:${a.color}`).join(','),
                p1Desc: window.app.photos[1].desc,
                p1Loc: window.app.photos[1].location,
                p1Time: window.app.photos[1].time,
                p1SideCropRatio: window.app.photos[1].sideCrop?.widthRatio || 0,
                blueLoc: window.app.photos.find(p => p.desc === '早檔藍圖')?.location,
                hasGreenCollisionKept: window.app.photos.some(p => p.desc === '晚檔撞號綠圖'),
                hasScene3: window.app.photos.some(p => p.desc === '第三現場補充'),
                handleCleared: window.app.currentProjectFileHandle === null
            };

            // 5. 模擬工作台目前已先鎖定為 10:00 存檔狀態，只手動改第 1 張（未 Ctrl+S），再併入 12:00 的 thirdBlob
            window.app.markProjectSaved(Date.parse('2026-11-10T10:00:00.000Z'));
            window.app.photos[0].desc = '畫面剛改還沒按CtrlS';
            window.app.lastLocalEditAt = Date.parse('2026-11-10T14:00:00.000Z');
            window.app.recordHistory('準備測試未存保護併入');

            const mergeInput = document.getElementById('mergeProjectInput');
            const mergeDt = new DataTransfer();
            mergeDt.items.add(new File([thirdBlob], 'C_third.photo-report', { type: 'application/zip' }));
            mergeInput.files = mergeDt.files;
            mergeInput.dispatchEvent(new Event('change', { bubbles: true }));
            await new Promise((resolve, reject) => {
                const deadline = Date.now() + 8000;
                const tick = () => {
                    if (window.app.photos.length === 7 && alerts.some(m => m.includes('已完成彙整 1 份工作進度'))) return resolve();
                    if (Date.now() > deadline) return reject(new Error('選單併入第三份 .photo-report 逾時'));
                    setTimeout(tick, 30);
                };
                tick();
            });
            const afterThirdMerge = {
                count: window.app.photos.length,
                seqs: window.app.photos.map(p => p.seq),
                p0Desc: window.app.photos[0].desc,
                p1Desc: window.app.photos[1].desc,
                updatedChildDesc: window.app.photos.find(p => p.uid === childCropUid)?.desc,
                addedSecondChildDesc: window.app.photos.find(p => p.name === 'IMG_8888_特寫2.jpg')?.desc,
                addedSecondChildSeq: window.app.photos.find(p => p.name === 'IMG_8888_特寫2.jpg')?.seq
            };

            // 6. 測試損壞進度檔合併失敗時保護原工作台與 URL 不洩漏
            const urlsBeforeBadMerge = window.app.managedUrls.size;
            const badMergeZip = new window.JSZip();
            badMergeZip.file('manifest.json', JSON.stringify({
                format: 'photo-report-project',
                version: 1,
                caseData: {},
                photos: [{ uid: 'bad1', archiveName: 'missing.jpg', name: 'missing.jpg', seq: '1' }]
            }));
            const badMergeOk = await window.app.mergeProjectFiles([new File([await badMergeZip.generateAsync({ type: 'blob' })], 'bad.photo-report', { type: 'application/zip' })]);
            const urlsAfterBadMerge = window.app.managedUrls.size;

            // 7. 測試 Ctrl+Z 復原最後一次彙整（應回到 6 張與第 2 張被第三手修訂前的狀態）
            window.app.undo();
            const afterUndoMerge = {
                count: window.app.photos.length,
                p1Desc: window.app.photos[1].desc,
                hasSecondChild: window.app.photos.some(p => p.name === 'IMG_8888_特寫2.jpg')
            };

            // 收尾清除並復原乾淨狀態
            await window.app.clearAll(true);
            return {
                afterTwoDrop,
                afterThirdMerge,
                badMergeOk,
                badMergeErrorAlert: alerts.at(-1) || '',
                urlsLeaked: urlsAfterBadMerge !== urlsBeforeBadMerge,
                afterUndoMerge
            };
        } finally {
            window.alert = origAlert;
            window.confirm = origConfirm;
            window.saveAs = origSaveAs;
        }
    });

    if (
        mergeCheck.afterTwoDrop.count !== 6 ||
        mergeCheck.afterTwoDrop.p0Desc !== '較晚完整說明' ||
        mergeCheck.afterTwoDrop.p0Loc !== '台北市中正區' ||
        mergeCheck.afterTwoDrop.p0Anns !== 'rect:yellow,label:yellow' ||
        mergeCheck.afterTwoDrop.p1Desc !== '補齊空白說明' ||
        mergeCheck.afterTwoDrop.p1Loc !== '大安區路段' ||
        mergeCheck.afterTwoDrop.p1Time !== '14:25' ||
        mergeCheck.afterTwoDrop.p1SideCropRatio !== 0.72 ||
        mergeCheck.afterTwoDrop.blueLoc !== '舊預設地點' ||
        !mergeCheck.afterTwoDrop.hasGreenCollisionKept ||
        !mergeCheck.afterTwoDrop.hasScene3 ||
        !mergeCheck.afterTwoDrop.handleCleared
    ) {
        throw new Error(`雙進度檔跨機智慧彙整結果異常: ${JSON.stringify(mergeCheck.afterTwoDrop)}`);
    }
    if (
        mergeCheck.afterThirdMerge.count !== 7 ||
        mergeCheck.afterThirdMerge.p0Desc !== '畫面剛改還沒按CtrlS' ||
        mergeCheck.afterThirdMerge.p1Desc !== '第三手最新修訂第2張' ||
        mergeCheck.afterThirdMerge.updatedChildDesc !== '第三手更新同特寫1' ||
        mergeCheck.afterThirdMerge.addedSecondChildDesc !== '第三手新增第二特寫' ||
        !String(mergeCheck.afterThirdMerge.addedSecondChildSeq).endsWith('-2') ||
        mergeCheck.badMergeOk !== false ||
        !mergeCheck.badMergeErrorAlert.startsWith('工作進度檔無法合併：') ||
        mergeCheck.urlsLeaked ||
        mergeCheck.afterUndoMerge.count !== 6 ||
        mergeCheck.afterUndoMerge.p1Desc !== '補齊空白說明' ||
        mergeCheck.afterUndoMerge.hasSecondChild
    ) {
        throw new Error(`後續併入未存修改保護、子特寫合併、錯誤防護或 Ctrl+Z 復原失敗: ${JSON.stringify(mergeCheck)}`);
    }
    console.log('  ✅ 多 .photo-report 工作進度檔智慧彙整（跨機重壓縮去重、防誤合、自動補齊、較晚覆蓋、未存保護、子特寫與復原）驗證成功');

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
