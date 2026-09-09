const path = require('path');
const fs = require('fs');
const assert = require('assert');

const { chromium } = require('playwright');

async function runBaselineTests() {
    console.log('========================================');
    console.log('開始執行 Phase 0C 匯出結構 Golden Baseline 比對');
    console.log('========================================\n');

    const baselineDir = path.resolve(__dirname, 'baseline');
    const docxBaseline = JSON.parse(fs.readFileSync(path.join(baselineDir, 'docx-structure.json'), 'utf8'));
    const excelBaseline = JSON.parse(fs.readFileSync(path.join(baselineDir, 'excel-data.json'), 'utf8'));
    const pdfBaseline = JSON.parse(fs.readFileSync(path.join(baselineDir, 'pdf-metadata.json'), 'utf8'));

    const htmlPath = path.resolve(__dirname, '../index.html');
    const fileUrl = 'file:///' + htmlPath.replace(/\\/g, '/');

    const browser = await chromium.launch({ headless: true });
    const page = await browser.newPage();
    await page.goto(fileUrl, { waitUntil: 'load' });

    // 載入 fixtures 照片
    const fixturesDir = path.resolve(__dirname, 'fixtures');
    await page.setInputFiles('#fileInput', [
        path.join(fixturesDir, 'sample01.jpg'),
        path.join(fixturesDir, 'sample02.jpg')
    ]);
    await page.waitForFunction(() => window.app && window.app.photos.length === 2);

    // 注入與 baseline 完全相同的標準案件資料
    await page.evaluate(() => {
        document.getElementById('caseTitle').value = '現場勘驗採證案件';
        document.getElementById('caseDate').value = '113/08/26';
        document.getElementById('defaultLocation').value = '現場第一勘驗點';
        document.getElementById('deptName').value = '臺南市政府警察局新化分局';
        document.getElementById('officerName').value = '巡官梁家豪';
        window.app.photos[0].desc = '第一張跡證說明';
        window.app.photos[1].desc = '第二張跡證說明';
    });

    // 1. 測試 Word (.docx) 結構比對
    console.log('[1/3] 比對 Word (.docx) XML 結構與關鍵字...');
    const currentDocx = await page.evaluate(async () => {
        let capturedBlob = null;
        window.saveAs = (blob) => { capturedBlob = blob; };
        await window.app.exportDocx();
        const zip = await window.JSZip.loadAsync(capturedBlob);
        const xmlText = await zip.file('word/document.xml').async('string');
        const headerFile = zip.file(/word\/header\d+\.xml/)[0];
        const headerText = headerFile ? await headerFile.async('string') : '';
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlText, 'application/xml');
        const tables = xmlDoc.getElementsByTagName('w:tbl');
        const paragraphs = xmlDoc.getElementsByTagName('w:p');
        return {
            tableCount: tables.length,
            paragraphCount: paragraphs.length,
            keywords: [
                '現場勘驗採證案件',
                '113/08/26',
                '現場第一勘驗點',
                '巡官梁家豪',
                '第一張跡證說明',
                '第二張跡證說明'
            ].map(kw => ({ keyword: kw, present: xmlText.includes(kw) })),
            deptNameInHeader: headerText.includes('臺南市政府警察局新化分局')
        };
    });

    assert.strictEqual(currentDocx.tableCount, docxBaseline.tableCount, `DOCX 表格數量不一致：預期 ${docxBaseline.tableCount}，實際 ${currentDocx.tableCount}`);
    assert.strictEqual(currentDocx.paragraphCount, docxBaseline.paragraphCount, `DOCX 段落數量不一致：預期 ${docxBaseline.paragraphCount}，實際 ${currentDocx.paragraphCount}`);
    assert.strictEqual(currentDocx.deptNameInHeader, true, 'DOCX 頁首 Header 應包含機關名稱');
    docxBaseline.keywords.forEach(({ keyword, present }) => {
        if (keyword !== '臺南市政府警察局新化分局') {
            const actualKw = currentDocx.keywords.find(k => k.keyword === keyword);
            assert.strictEqual(actualKw && actualKw.present, present, `DOCX 內文關鍵字 [${keyword}] 存在性不符合基準！`);
        }
    });
    console.log('  ✅ Word (.docx) 結構與 XML 完全符合 Golden Baseline！');

    // 2. 測試 Excel (.xlsx) 欄位與資料比對
    console.log('[2/3] 比對 Excel (.xlsx) 工作表與欄位資料...');
    const currentExcel = await page.evaluate(() => {
        let capturedWb = null;
        const origWriteFile = window.XLSX.writeFile;
        window.XLSX.writeFile = (wb) => { capturedWb = wb; };
        window.app.exportExcel();
        window.XLSX.writeFile = origWriteFile;
        const sheetName = capturedWb.SheetNames[0];
        const sheet = capturedWb.Sheets[sheetName];
        const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1 });
        return {
            sheetNames: capturedWb.SheetNames,
            rowCount: rows.length,
            headers: rows[0],
            rows: rows.slice(1)
        };
    });

    assert.deepStrictEqual(currentExcel.sheetNames, excelBaseline.sheetNames, 'Excel 工作表名稱不一致');
    assert.strictEqual(currentExcel.rowCount, excelBaseline.rowCount, 'Excel 資料總列數不一致');
    assert.deepStrictEqual(currentExcel.headers, excelBaseline.headers, 'Excel 欄位標題不一致');
    assert.deepStrictEqual(currentExcel.rows, excelBaseline.rows, 'Excel 資料行內容不一致');
    console.log('  ✅ Excel (.xlsx) 工作表欄位與資料完全符合 Golden Baseline！');

    // 3. 測試 PDF (.pdf) 頁數與版面比對
    console.log('[3/3] 比對 PDF (.pdf) 頁數與版面尺寸...');
    const currentPdf = await page.evaluate(async () => {
        const origJsPdf = window.jspdf.jsPDF;
        let pdfOutput = null;
        window.jspdf.jsPDF = function(...args) {
            const doc = new origJsPdf(...args);
            doc.save = function() {
                pdfOutput = {
                    pageCount: doc.internal.getNumberOfPages(),
                    pageSize: {
                        width: Math.round(doc.internal.pageSize.getWidth()),
                        height: Math.round(doc.internal.pageSize.getHeight())
                    }
                };
            };
            return doc;
        };
        await window.app.exportPdf();
        window.jspdf.jsPDF = origJsPdf;
        return pdfOutput;
    });

    assert.strictEqual(currentPdf.pageCount, pdfBaseline.pageCount, 'PDF 頁數不一致');
    assert.strictEqual(currentPdf.pageSize.width, pdfBaseline.pageSize.width, 'PDF 頁面寬度不一致');
    assert.strictEqual(currentPdf.pageSize.height, pdfBaseline.pageSize.height, 'PDF 頁面高度不一致');
    console.log('  ✅ PDF (.pdf) 頁數與版面尺寸完全符合 Golden Baseline！');

    // 4. 三種清冊版型均需能實際產生 Word 與 PDF；UI 搬移不應影響匯出器讀取版型。
    console.log('[4/3] 驗證三種清冊版型的 Word／PDF 匯出...');
    const layoutExports = await page.evaluate(async () => {
        const variants = [
            { value: 'up_down_2', expectedLandscape: false },
            { value: 'left_right_2', expectedLandscape: false },
            { value: 'landscape_3', expectedLandscape: true }
        ];
        const originalSaveAs = window.saveAs;
        const originalJsPdf = window.jspdf.jsPDF;
        const results = [];

        try {
            for (const variant of variants) {
                const layoutSelect = document.getElementById('layoutSelect');
                layoutSelect.value = variant.value;
                layoutSelect.dispatchEvent(new Event('change', { bubbles: true }));

                let docxBlob = null;
                window.saveAs = blob => { docxBlob = blob; };
                await window.app.exportDocx();
                const docxZip = await window.JSZip.loadAsync(docxBlob);
                const docxXml = await docxZip.file('word/document.xml').async('string');

                let pdfOutput = null;
                window.jspdf.jsPDF = function(...args) {
                    const doc = new originalJsPdf(...args);
                    doc.save = function() {
                        pdfOutput = {
                            pageCount: doc.internal.getNumberOfPages(),
                            width: Math.round(doc.internal.pageSize.getWidth()),
                            height: Math.round(doc.internal.pageSize.getHeight())
                        };
                    };
                    return doc;
                };
                await window.app.exportPdf();

                results.push({
                    value: variant.value,
                    expectedLandscape: variant.expectedLandscape,
                    docxBytes: docxBlob?.size || 0,
                    docxIsValidDocument: docxXml.includes('<w:document') && docxXml.includes('<w:tbl'),
                    pdfOutput
                });
            }
        } finally {
            window.saveAs = originalSaveAs;
            window.jspdf.jsPDF = originalJsPdf;
        }

        return results;
    });

    layoutExports.forEach(result => {
        assert.ok(result.docxBytes > 0, `${result.value} Word 未產生檔案內容`);
        assert.strictEqual(result.docxIsValidDocument, true, `${result.value} Word 文件結構不正確`);
        assert.ok(result.pdfOutput?.pageCount >= 1, `${result.value} PDF 未產生頁面`);
        assert.strictEqual(result.pdfOutput.width > result.pdfOutput.height, result.expectedLandscape, `${result.value} PDF 方向不正確`);
    });
    console.log('  ✅ 上下兩張、左右兩張、橫式三張皆可正常產生 Word／PDF！');

    await browser.close();
    console.log('\n========================================');
    console.log('🎉 Phase 0C 匯出結構 Golden Baseline 比對全部通過！');
    console.log('========================================');
}

runBaselineTests().catch(err => {
    console.error('\n❌ Phase 0C 基準比對失敗:', err);
    process.exit(1);
});
