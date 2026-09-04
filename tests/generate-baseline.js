const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

(async () => {
    console.log('正在為 Phase 0C 產生 Golden Baseline 基準檔...');
    const browser = await chromium.launch();
    const page = await browser.newPage();
    const htmlPath = path.resolve('index.html');
    await page.goto('file:///' + htmlPath.replace(/\\/g, '/'));

    const fixturesDir = path.resolve('tests/fixtures');
    await page.setInputFiles('#fileInput', [
        path.join(fixturesDir, 'sample01.jpg'),
        path.join(fixturesDir, 'sample02.jpg')
    ]);
    await page.waitForFunction(() => window.app && window.app.photos.length === 2);

    // 注入固定標準案件資料
    await page.evaluate(() => {
        document.getElementById('caseTitle').value = '現場勘驗採證案件';
        document.getElementById('caseDate').value = '113/08/26';
        document.getElementById('defaultLocation').value = '現場第一勘驗點';
        document.getElementById('deptName').value = '臺南市政府警察局新化分局';
        document.getElementById('officerName').value = '巡官梁家豪';
        window.app.photos[0].desc = '第一張跡證說明';
        window.app.photos[1].desc = '第二張跡證說明';
    });

    // 1. 產生 DOCX 結構基準
    const docxBaseline = await page.evaluate(async () => {
        let capturedBlob = null;
        window.saveAs = (blob) => { capturedBlob = blob; };
        await window.app.exportDocx();
        const zip = await window.JSZip.loadAsync(capturedBlob);
        const xmlText = await zip.file('word/document.xml').async('string');
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
                '臺南市政府警察局新化分局',
                '巡官梁家豪',
                '第一張跡證說明',
                '第二張跡證說明'
            ].map(kw => ({ keyword: kw, present: xmlText.includes(kw) }))
        };
    });

    // 2. 產生 Excel 結構基準
    const excelBaseline = await page.evaluate(() => {
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

    // 3. 產生 PDF 結構基準
    const pdfBaseline = await page.evaluate(async () => {
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

    const baselineDir = path.resolve('tests/baseline');
    if (!fs.existsSync(baselineDir)) fs.mkdirSync(baselineDir, { recursive: true });

    fs.writeFileSync(path.join(baselineDir, 'docx-structure.json'), JSON.stringify(docxBaseline, null, 2), 'utf8');
    fs.writeFileSync(path.join(baselineDir, 'excel-data.json'), JSON.stringify(excelBaseline, null, 2), 'utf8');
    fs.writeFileSync(path.join(baselineDir, 'pdf-metadata.json'), JSON.stringify(pdfBaseline, null, 2), 'utf8');

    console.log('✅ 成功產出三份 Golden Baseline 基準檔：');
    console.log('  - tests/baseline/docx-structure.json');
    console.log('  - tests/baseline/excel-data.json');
    console.log('  - tests/baseline/pdf-metadata.json');

    await browser.close();
})();
