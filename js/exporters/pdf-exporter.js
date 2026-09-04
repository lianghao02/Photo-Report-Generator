/**
 * pdf-exporter.js - PDF 清冊匯出模組
 * 封裝直橫式 A4、Canvas 中文字型繪製、多版型配置與 .pdf 檔案產生
 */
(function(root, factory) {
    if (typeof module === 'object' && typeof module.exports === 'object') {
        module.exports = factory();
    } else {
        root.PhotoReportPdfExporter = factory();
    }
})(typeof self !== 'undefined' ? self : this, function() {
    'use strict';

    /**
     * 將 Blob 轉成 Data URL
     * @param {Blob} blob
     * @returns {Promise<string>}
     */
    function blobToDataUrl(blob) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.onerror = () => reject(new Error('PDF 圖片轉換失敗。'));
            reader.readAsDataURL(blob);
        });
    }

    /**
     * 匯出 PDF 清冊
     * @param {Object} options
     * @param {Array} options.photos - 照片物件陣列
     * @param {string} options.caseTitle - 案由名稱
     * @param {string} options.caseDate - 採證日期
     * @param {string} options.defaultLocation - 預設地點
     * @param {string} options.globalOfficer - 製作人姓名
     * @param {string} options.deptName - 機關單位名稱
     * @param {string} options.layout - 排版版型 ('up_down_2' | 'left_right_2' | 'landscape_3')
     * @param {Function} options.processImageToBlob - 圖片壓縮處理函式 (photoItem, maxDimension, quality) => Promise<{blob, width, height}>
     * @param {Function} [options.setProgress] - 進度更新回呼 (msg, current, total)
     * @param {Function} [options.clearProgress] - 進度清除回呼 ()
     * @param {Function} [options.throwIfCancelled] - 取消檢查回呼 ()
     * @param {Function} [options.yieldToBrowser] - 瀏覽器讓渡回呼 ()
     * @param {Object} [options.jspdfLib] - jsPDF 函式庫實體（預設由全域或 window.jspdf 取得）
     */
    async function exportPdf({
        photos = [],
        caseTitle = '現場照片',
        caseDate = '',
        defaultLocation = '',
        globalOfficer = '',
        deptName = '臺南市政府警察局新化分局',
        layout = 'up_down_2',
        processImageToBlob,
        setProgress = () => {},
        clearProgress = () => {},
        throwIfCancelled = () => {},
        yieldToBrowser = () => Promise.resolve(),
        jspdfLib = (typeof window !== 'undefined' && window.jspdf ? window.jspdf : (typeof jspdf !== 'undefined' ? jspdf : null))
    } = {}) {
        if (!photos.length) {
            if (typeof alert === 'function') alert('請先載入照片再匯出！');
            return;
        }
        if (!jspdfLib?.jsPDF) {
            if (typeof alert === 'function') alert('PDF 函式庫尚未載入，請重新整理後再試一次。');
            return;
        }

        const { jsPDF } = jspdfLib;
        const isLandscape = layout === 'landscape_3';
        const perPage = layout === 'up_down_2' ? 2 : layout === 'left_right_2' ? 2 : 3;
        const totalPages = Math.ceil(photos.length / perPage);
        const pdf = new jsPDF({ orientation: isLandscape ? 'landscape' : 'portrait', unit: 'mm', format: 'a4', compress: true });
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const title = (caseTitle || '').trim() || '現場照片';
        const date = (caseDate || '').trim() || '';
        const loc = (defaultLocation || '').trim() || '';
        const officer = (globalOfficer || '').trim() || '';
        const dept = (deptName || '').trim() || '臺南市政府警察局新化分局';

        const drawText = (text, x, y, width, size = 9, align = 'left') => {
            // jsPDF 內建字型沒有中文；以 Canvas 使用 Windows 系統中文字型繪製後嵌入 PDF。
            const scale = 3;
            const pixelsPerMm = 3.78 * scale;
            const fontPixels = Math.max(12, size * 1.35 * scale);
            const canvas = document.createElement('canvas');
            const context = canvas.getContext('2d');
            context.font = `${fontPixels}px "Microsoft JhengHei", sans-serif`;
            const maxPixels = width * pixelsPerMm;
            const lines = [];
            let line = '';
            for (const character of String(text || '')) {
                const candidate = line + character;
                if (line && context.measureText(candidate).width > maxPixels) {
                    lines.push(line);
                    line = character;
                } else {
                    line = candidate;
                }
            }
            lines.push(line || ' ');
            const linePixels = Math.ceil(fontPixels * 1.45);
            canvas.width = Math.ceil(maxPixels);
            canvas.height = linePixels * lines.length + Math.ceil(fontPixels * 0.5);
            context.font = `${fontPixels}px "Microsoft JhengHei", sans-serif`;
            context.fillStyle = '#20242B';
            context.textBaseline = 'top';
            lines.forEach((lineText, lineIndex) => {
                const measured = context.measureText(lineText).width;
                const textX = align === 'center' ? (canvas.width - measured) / 2 : align === 'right' ? canvas.width - measured : 0;
                context.fillText(lineText, textX, lineIndex * linePixels);
            });
            const height = canvas.height / pixelsPerMm;
            pdf.addImage(canvas.toDataURL('image/png'), 'PNG', x, y - (fontPixels * 0.9 / pixelsPerMm), width, height, undefined, 'FAST');
            return height;
        };

        const drawPhoto = (dataUrl, x, y, maxWidth, maxHeight) => {
            const properties = pdf.getImageProperties(dataUrl);
            const ratio = Math.min(maxWidth / properties.width, maxHeight / properties.height);
            const width = properties.width * ratio;
            const height = properties.height * ratio;
            pdf.addImage(dataUrl, 'JPEG', x + (maxWidth - width) / 2, y + (maxHeight - height) / 2, width, height, undefined, 'FAST');
        };

        const unit = dept;
        const margin = 10;

        setProgress('正在壓縮照片並建立 PDF 清冊…', 0, photos.length);
        for (let index = 0; index < photos.length; index++) {
            const photo = photos[index];
            if (index > 0 && index % perPage === 0) pdf.addPage();
            const pageIndex = Math.floor(index / perPage) + 1;
            const slot = index % perPage;
            const processed = await processImageToBlob(photo, 1600, 0.88);
            const dataUrl = await blobToDataUrl(processed.blob);
            throwIfCancelled();

            pdf.setDrawColor(65, 88, 117);
            if (slot === 0) drawText(`${unit ? `${unit} ` : ''}蒐證照片`, margin, 8, pageWidth - margin * 2, 14, 'center');

            if (layout === 'up_down_2') {
                const top = 14 + slot * ((pageHeight - 22) / 2);
                const cardHeight = (pageHeight - 22) / 2 - 3;
                pdf.rect(margin, top, pageWidth - margin * 2, cardHeight);
                drawText(`案由：${title}`, margin + 4, top + 7, pageWidth - 28, 9);
                drawText(`日期：${photo.date || date}　時間：${photo.time || ''}　地點：${photo.location || loc}`, margin + 4, top + 13, pageWidth - 28, 8);
                drawPhoto(dataUrl, margin + 5, top + 17, pageWidth - margin * 2 - 10, cardHeight - 42);
                drawText(`說明：${photo.desc || '（無說明）'}`, margin + 4, top + cardHeight - 19, pageWidth - 28, 8);
                drawText(`編號：${photo.seq}　製作人：${officer}`, margin + 4, top + cardHeight - 7, pageWidth - 28, 8);
            } else if (layout === 'left_right_2') {
                const cardWidth = (pageWidth - margin * 2 - 4) / 2;
                const left = margin + slot * (cardWidth + 4);
                const top = 14;
                const cardHeight = pageHeight - 27;
                pdf.rect(left, top, cardWidth, cardHeight);
                drawText(`#${photo.seq}　${title}`, left + 3, top + 7, cardWidth - 6, 9);
                drawText(`${photo.date || date} ${photo.time || ''}`, left + 3, top + 13, cardWidth - 6, 8);
                drawText(photo.location || loc, left + 3, top + 19, cardWidth - 6, 8);
                drawPhoto(dataUrl, left + 4, top + 23, cardWidth - 8, cardHeight - 65);
                drawText(`說明：${photo.desc || '（無說明）'}`, left + 3, top + cardHeight - 34, cardWidth - 6, 8);
                drawText(`製作人：${officer}`, left + 3, top + cardHeight - 8, cardWidth - 6, 8);
            } else {
                const cardWidth = (pageWidth - margin * 2 - 4) / 3;
                const left = margin + slot * (cardWidth + 2);
                const top = 14;
                const cardHeight = pageHeight - 27;
                pdf.rect(left, top, cardWidth, cardHeight);
                drawText(`#${photo.seq}　${title}`, left + 3, top + 7, cardWidth - 6, 8);
                drawText(`${photo.date || date} ${photo.time || ''}`, left + 3, top + 12, cardWidth - 6, 7);
                drawText(photo.location || loc, left + 3, top + 17, cardWidth - 6, 7);
                drawPhoto(dataUrl, left + 3, top + 21, cardWidth - 6, cardHeight - 56);
                drawText(`說明：${photo.desc || '（無說明）'}`, left + 3, top + cardHeight - 25, cardWidth - 6, 7);
            }
            if (slot === perPage - 1 || index === photos.length - 1) {
                drawText(`第 ${pageIndex} 頁 / 共 ${totalPages} 頁`, margin, pageHeight - 5, pageWidth - margin * 2, 8, 'center');
            }
            setProgress('正在壓縮照片並建立 PDF 清冊…', index + 1, photos.length);
            if ((index + 1) % 3 === 0) await yieldToBrowser();
        }

        const safeName = `${date || '報告'}_${title}_現場照片清冊.pdf`.replace(/[\\/:*?"<>|]/g, '_');
        pdf.save(safeName);
        clearProgress();
    }

    return {
        blobToDataUrl: blobToDataUrl,
        exportPdf: exportPdf
    };
});
