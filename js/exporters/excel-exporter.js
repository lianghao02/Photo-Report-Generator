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
    function exportExcel({
        photos = [],
        caseTitle = '詐欺案',
        caseDate = '',
        defaultLocation = '',
        globalOfficer = '巡官梁家豪',
        xlsxLib = (typeof XLSX !== 'undefined' ? XLSX : (typeof window !== 'undefined' ? window.XLSX : null))
    } = {}) {
        if (!photos.length) {
            if (typeof alert === 'function') alert('請先載入照片！');
            return;
        }
        if (!xlsxLib) {
            if (typeof alert === 'function') alert('SheetJS 尚未載入！');
            return;
        }

        const title = (caseTitle || '').trim() || '詐欺案';
        const date = (caseDate || '').trim() || '';
        const loc = (defaultLocation || '').trim() || '';
        const officer = (globalOfficer || '').trim() || '巡官梁家豪';

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
        const wb = xlsxLib.utils.book_new();
        xlsxLib.utils.book_append_sheet(wb, ws, '工作表1');

        const safeName = `${date || '報告'}_${title}_資料清單.xlsx`.replace(/[\\/:*?"<>|]/g, '_');
        xlsxLib.writeFile(wb, safeName);
    }

    return {
        exportExcel: exportExcel
    };
});
