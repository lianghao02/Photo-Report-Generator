/**
 * excel-exporter.js - Excel 清冊匯出模組
 * 封裝 SheetJS 工作表轉換、欄位對齊與 .xlsx 檔案產生
 */
(function(root, factory) {
    if (typeof module === 'object' && typeof module.exports === 'object') {
        module.exports = factory();
    } else {
        root.PhotoReportExcelExporter = factory();
    }
})(typeof self !== 'undefined' ? self : this, function() {
    'use strict';

    /**
     * 匯出 Excel 清冊 (.xlsx)
     * @param {Object} options
     * @param {Array} options.photos - 照片物件陣列
     * @param {string} options.caseTitle - 案由名稱
     * @param {string} options.caseDate - 採證日期
     * @param {string} options.defaultLocation - 預設採證地點
     * @param {string} options.globalOfficer - 製作人姓名
     * @param {Object} [options.xlsxLib] - SheetJS XLSX 物件（預設自動讀取全域 XLSX）
     */
    async function exportExcel({
        photos = [],
        caseTitle = '詐欺案',
        caseSubject = '',
        caseDate = '',
        defaultLocation = '',
        globalOfficer,
        xlsxLib = (typeof XLSX !== 'undefined' ? XLSX : (typeof window !== 'undefined' ? window.XLSX : null)),
        saveAs = null
    } = {}) {
        if (!photos.length) {
            if (typeof alert === 'function') alert('請先載入照片！');
            return;
        }
        if (!xlsxLib) {
            if (typeof alert === 'function') alert('SheetJS 尚未載入！');
            return;
        }

        const cfg = (typeof PhotoReportConfig !== 'undefined' && PhotoReportConfig) ? PhotoReportConfig : {};
        const title = (caseTitle || '').trim() || cfg.DEFAULT_CASE_TITLE || '詐欺案';
        const date = (caseDate || '').trim() || '';
        const loc = (defaultLocation || '').trim() || '';
        const officer = globalOfficer !== undefined ? String(globalOfficer ?? '').trim() : (cfg.DEFAULT_OFFICER_NAME || '');

        const headers = ['編號', '檔名', '案由', '日期', '時間', '地點', '製作人', '說明'];
        const data = [headers];

        photos.forEach(p => {
            data.push([
                p.seq,
                p.name,
                title,
                p.date || date,
                p.time,
                p.location || loc,
                officer,
                p.desc
            ]);
        });

        const ws = xlsxLib.utils.aoa_to_sheet(data);
        // 強制編號、日期與時間欄位以純文字格式儲存，避免 Excel 開啟後吃掉前導零或誤轉為小數時間
        const textCols = [0, 3, 4];
        for (let r = 1; r < data.length; r++) {
            for (const c of textCols) {
                const ref = xlsxLib.utils.encode_cell({ r, c });
                const cell = ws[ref];
                if (cell) {
                    cell.t = 's';
                    cell.v = String(cell.v ?? '');
                    cell.z = '@';
                }
            }
        }
        const wb = xlsxLib.utils.book_new();
        xlsxLib.utils.book_append_sheet(wb, ws, '工作表1');

        const compactDate = String(date || '').replace(/[\\/:*?"<>|\s-]+/g, '');
        const cleanSubject = String(caseSubject || '').trim()
            .replace(/[\\/:*?"<>|\u0000-\u001F\s]+/g, '_')
            .replace(/_+/g, '_')
            .replace(/^_+|_+$/g, '');
        const shortSubject = Array.from(cleanSubject).slice(0, 15).join('').replace(/^_+|_+$/g, '');
        const cleanFileTitle = String(title || '').replace(/[\\/:*?"<>|\u0000-\u001F]+/g, '_').replace(/_+/g, '_').replace(/^_+|_+$/g, '') || '現場照片清冊';
        const fileBase = [compactDate, shortSubject, cleanFileTitle].filter(Boolean).join('_') || '現場照片清冊';
        const safeName = `${fileBase}_資料清單.xlsx`.replace(/[\\/:*?"<>|]/g, '_');
        if (typeof saveAs === 'function') {
            const wbOut = xlsxLib.write(wb, { bookType: 'xlsx', type: 'array' });
            await saveAs(
                new Blob([wbOut], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
                safeName
            );
        } else {
            xlsxLib.writeFile(wb, safeName);
        }
    }

    return {
        exportExcel: exportExcel
    };
});
