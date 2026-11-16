/**
 * pdf-exporter.js - PDF 清冊匯出模組
 * 1:1 對齊 Word 三大公務版型（上下兩張 8302 dxa、左右兩張 9864 dxa、橫式三張 15648 dxa），
 * 使用標楷體優先字型、完整黑線表格儲存格、置中／靠左排版與跨頁 Header/Footer。
 */
(function(root, factory) {
    if (typeof module === 'object' && typeof module.exports === 'object') {
        module.exports = factory();
    } else {
        root.PhotoReportPdfExporter = factory();
    }
})(typeof self !== 'undefined' ? self : this, function() {
    'use strict';

    const DXA_TO_MM = 25.4 / 1440;
    const PX_TO_MM = 25.4 / 96;
    const FONT_FAMILY = '"DFKai-SB", "BiauKai", "標楷體", "Microsoft JhengHei", sans-serif';

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
     * 依儲存格寬度與換行符計算折行文字
     */
    function measureWrappedLines(context, text, maxPixels) {
        const normalized = String(text ?? '').replace(/\r\n?/g, '\n');
        const paragraphs = normalized.split('\n');
        const lines = [];
        for (const para of paragraphs) {
            if (!para.length) {
                lines.push('');
                continue;
            }
            let line = '';
            for (const ch of para) {
                const candidate = line + ch;
                if (line && context.measureText(candidate).width > maxPixels) {
                    lines.push(line);
                    line = ch;
                } else {
                    line = candidate;
                }
            }
            if (line) lines.push(line);
        }
        return lines.length ? lines : [''];
    }

    /**
     * 匯出 PDF 清冊（與 Word 三大版型同格線結構）
     */
    async function exportPdf({
        photos = [],
        caseTitle = '現場照片清冊',
        caseDate = '',
        defaultLocation = '',
        globalOfficer = '巡官梁家豪',
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
        const perPage = layout === 'landscape_3' ? 3 : 2;
        const totalPages = Math.ceil(photos.length / perPage);
        const pdf = new jsPDF({
            orientation: isLandscape ? 'landscape' : 'portrait',
            unit: 'mm',
            format: 'a4',
            compress: true
        });
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const title = (caseTitle || '').trim() || '現場照片清冊';
        const date = (caseDate || '').trim() || '';
        const loc = (defaultLocation || '').trim() || '';
        const officer = (globalOfficer || '').trim() || '巡官梁家豪';
        const dept = (deptName || '').trim() || '臺南市政府警察局新化分局';

        const pxPerMm = 4.2 * 3.78;
        const measureCanvas = document.createElement('canvas');
        const measureCtx = measureCanvas.getContext('2d');

        const buildFontStr = (fontPt, bold = false) => {
            const fontPx = Math.max(10, fontPt * 0.352778 * pxPerMm);
            return { fontPx, font: `${bold ? '700 ' : ''}${fontPx}px ${FONT_FAMILY}` };
        };

        /**
         * 計算單格或雙段文字在限定寬度下的最佳字級與自然高度（mm）
         */
        const layoutCellRuns = (runs, widthMm, padXMm = 1.4, padYMm = 1.0, maxHeightMm = Infinity) => {
            const innerW = Math.max(4, widthMm - padXMm * 2);
            const maxPixels = innerW * pxPerMm;
            const minScale = 0.65;
            let scaleFactor = 1.0;
            let renderedBlocks = [];
            let totalTextHeightMm = 0;

            while (true) {
                renderedBlocks = [];
                let totalPx = 0;
                for (const run of runs) {
                    const pt = (run.fontPt || 10) * scaleFactor;
                    const { fontPx, font } = buildFontStr(pt, Boolean(run.bold));
                    measureCtx.font = font;
                    const linePx = fontPx * 1.32;
                    let wrapped;
                    if (run.prefixBold) {
                        const prefixFont = buildFontStr(pt, true);
                        measureCtx.font = prefixFont.font;
                        const prefixW = measureCtx.measureText(run.prefixBold).width;
                        measureCtx.font = font;
                        wrapped = measureWrappedLines(measureCtx, run.text || '', Math.max(20, maxPixels - prefixW));
                        renderedBlocks.push({
                            fontPx,
                            font,
                            linePx,
                            prefixBold: run.prefixBold,
                            prefixFont: prefixFont.font,
                            prefixW,
                            lines: wrapped
                        });
                    } else {
                        wrapped = measureWrappedLines(measureCtx, run.text || '', maxPixels);
                        renderedBlocks.push({ fontPx, font, linePx, lines: wrapped });
                    }
                    totalPx += wrapped.length * linePx;
                }
                totalTextHeightMm = totalPx / pxPerMm + padYMm * 2;
                if (totalTextHeightMm <= maxHeightMm || scaleFactor <= minScale) {
                    if (totalTextHeightMm > maxHeightMm && Number.isFinite(maxHeightMm)) {
                        const maxLinesPx = Math.max(10, (maxHeightMm - padYMm * 2) * pxPerMm);
                        let accPx = 0;
                        const clampedBlocks = [];
                        for (const block of renderedBlocks) {
                            const kept = [];
                            for (const line of block.lines) {
                                if (accPx + block.linePx <= maxLinesPx || clampedBlocks.length === 0 && kept.length === 0) {
                                    kept.push(line);
                                    accPx += block.linePx;
                                } else {
                                    if (kept.length > 0) {
                                        const lastIdx = kept.length - 1;
                                        kept[lastIdx] = kept[lastIdx].length > 1 ? `${kept[lastIdx].slice(0, -1)}…` : '…';
                                    }
                                    break;
                                }
                            }
                            if (kept.length > 0) clampedBlocks.push({ ...block, lines: kept });
                            if (accPx + block.linePx > maxLinesPx) break;
                        }
                        if (clampedBlocks.length > 0) {
                            renderedBlocks = clampedBlocks;
                            totalTextHeightMm = accPx / pxPerMm + padYMm * 2;
                        }
                    }
                    break;
                }
                scaleFactor = Math.max(minScale, scaleFactor - 0.08);
            }

            return {
                renderedBlocks,
                innerW,
                padXMm,
                padYMm,
                requiredHeightMm: totalTextHeightMm
            };
        };

        /**
         * 將排版後的文字塊繪製至透明 Canvas 並嵌入 PDF 區域
         */
        const drawTextLayoutInRect = (textLayout, xMm, yMm, wMm, hMm, align = 'center', valign = 'center') => {
            const blocks = textLayout.renderedBlocks;
            if (!blocks.some(b => b.prefixBold || b.lines.some(line => line.length > 0))) return;

            const innerPxW = Math.max(8, Math.ceil(textLayout.innerW * pxPerMm));
            const totalLinesPx = blocks.reduce((sum, b) => sum + b.lines.length * b.linePx, 0);
            const canvasH = Math.max(8, Math.ceil(totalLinesPx + 4));
            const canvas = document.createElement('canvas');
            canvas.width = innerPxW;
            canvas.height = canvasH;
            const ctx = canvas.getContext('2d');
            ctx.fillStyle = '#000000';
            ctx.textBaseline = 'middle';

            let cursorY = 2;
            for (const block of blocks) {
                for (let li = 0; li < block.lines.length; li++) {
                    const lineText = block.lines[li];
                    const centerY = cursorY + block.linePx / 2;
                    if (li === 0 && block.prefixBold) {
                        ctx.font = block.prefixFont;
                        ctx.fillText(block.prefixBold, 0, centerY);
                        ctx.font = block.font;
                        ctx.fillText(lineText, block.prefixW, centerY);
                    } else {
                        ctx.font = block.font;
                        const textW = ctx.measureText(lineText).width;
                        let drawX = 0;
                        if (align === 'center') drawX = Math.max(0, (innerPxW - textW) / 2);
                        else if (align === 'right') drawX = Math.max(0, innerPxW - textW);
                        ctx.fillText(lineText, drawX, centerY);
                    }
                    cursorY += block.linePx;
                }
            }

            const drawnHMm = canvasH / pxPerMm;
            const drawX = xMm + textLayout.padXMm;
            const availHMm = Math.max(1, hMm - textLayout.padYMm * 2);
            const fitH = Math.min(drawnHMm, availHMm);
            const drawY = valign === 'top'
                ? (yMm + textLayout.padYMm)
                : (yMm + Math.max(textLayout.padYMm, (hMm - fitH) / 2));

            pdf.addImage(canvas.toDataURL('image/png'), 'PNG', drawX, drawY, textLayout.innerW, fitH, undefined, 'FAST');
        };

        /**
         * 繪製無框線的單行頁首／頁尾文字（沿用標楷體置中）
         */
        const drawHeaderFooterText = (text, xMm, yMm, wMm, fontPt, bold = false) => {
            const layoutObj = layoutCellRuns([{ text, fontPt, bold }], wMm, 0, 0);
            drawTextLayoutInRect(layoutObj, xMm, yMm, wMm, layoutObj.requiredHeightMm, 'center', 'center');
        };

        /**
         * 依最大寬高置中繪製照片
         */
        const drawPhotoInRect = (dataUrl, xMm, yMm, wMm, hMm, maxImgWMm, maxImgHMm) => {
            const props = pdf.getImageProperties(dataUrl);
            const boundW = Math.min(wMm - 2, maxImgWMm || (wMm - 2));
            const boundH = Math.min(hMm - 2, maxImgHMm || (hMm - 2));
            const ratio = Math.min(boundW / props.width, boundH / props.height);
            const imgW = props.width * ratio;
            const imgH = props.height * ratio;
            const imgX = xMm + (wMm - imgW) / 2;
            const imgY = yMm + (hMm - imgH) / 2;
            pdf.addImage(dataUrl, 'JPEG', imgX, imgY, imgW, imgH, undefined, 'FAST');
        };

        /**
         * 統一繪製帶黑色格線的表格（先算各列高度，再統一描繪文字、照片與儲存格框線）
         * @returns {number} 表格總高度 (mm)
         */
        const renderTable = (startX, startY, colWidthsMm, rows, maxTableHeightMm = Infinity) => {
            const computedRows = rows.map(row => {
                let cursorCol = 0;
                const preparedCells = row.cells.map(cell => {
                    const span = cell.colSpan || 1;
                    let cellW = 0;
                    for (let c = 0; c < span; c++) cellW += colWidthsMm[cursorCol + c] || 0;
                    const xOffset = colWidthsMm.slice(0, cursorCol).reduce((s, v) => s + v, 0);
                    cursorCol += span;

                    if (cell.photoDataUrl !== undefined) {
                        return { ...cell, xOffset, cellW, kind: 'photo' };
                    }
                    const runs = cell.runs || [{ text: cell.text ?? '', fontPt: cell.fontPt || 11, bold: Boolean(cell.bold) }];
                    const textLayout = layoutCellRuns(
                        runs,
                        cellW,
                        cell.padXMm ?? 1.4,
                        cell.padYMm ?? 1.0,
                        cell.maxHeightMm ?? Infinity
                    );
                    return { ...cell, xOffset, cellW, kind: 'text', textLayout };
                });

                let rowH = row.minHeightMm || 6;
                if (row.exactHeightMm) {
                    rowH = row.exactHeightMm;
                } else {
                    for (const pc of preparedCells) {
                        if (pc.kind === 'text' && pc.textLayout.requiredHeightMm > rowH) {
                            rowH = pc.textLayout.requiredHeightMm;
                        }
                    }
                }
                return { ...row, preparedCells, rowH };
            });

            // 若整張表有預設照片列吸收差額，避免文字長高時把表底推出頁面
            const totalTextRowsH = computedRows.reduce((s, r) => s + (r.isPhotoRow ? 0 : r.rowH), 0);
            const photoRow = computedRows.find(r => r.isPhotoRow);
            if (photoRow && Number.isFinite(maxTableHeightMm)) {
                const maxAvailForPhoto = Math.max(25, maxTableHeightMm - totalTextRowsH);
                if (photoRow.rowH > maxAvailForPhoto) {
                    photoRow.rowH = maxAvailForPhoto;
                }
            }

            let curY = startY;
            pdf.setDrawColor(0, 0, 0);
            pdf.setLineWidth(0.25);

            for (const row of computedRows) {
                for (const pc of row.preparedCells) {
                    const cx = startX + pc.xOffset;
                    const cy = curY;
                    const cw = pc.cellW;
                    const ch = row.rowH;

                    if (pc.kind === 'photo') {
                        if (pc.photoDataUrl) {
                            drawPhotoInRect(pc.photoDataUrl, cx, cy, cw, ch, pc.maxImgWMm, pc.maxImgHMm);
                        } else {
                            const emptyLayout = layoutCellRuns([{ text: '（無照片）', fontPt: 10, bold: false }], cw, 1.4, 1.0);
                            drawTextLayoutInRect(emptyLayout, cx, cy, cw, ch, 'center', 'center');
                        }
                    } else {
                        drawTextLayoutInRect(pc.textLayout, cx, cy, cw, ch, pc.align || 'center', pc.valign || 'center');
                    }

                    pdf.rect(cx, cy, cw, ch, 'S');
                }
                curY += row.rowH;
            }

            return curY - startY;
        };

        // 預先壓縮並轉換所有照片為 Data URL
        const preparedPhotos = [];
        setProgress('正在壓縮照片並建立 PDF 清冊…', 0, photos.length);
        for (let index = 0; index < photos.length; index++) {
            const photo = photos[index];
            const processed = await processImageToBlob(photo, 1600, 0.88);
            const dataUrl = await blobToDataUrl(processed.blob);
            throwIfCancelled();
            preparedPhotos.push({ item: photo, dataUrl });
            setProgress('正在壓縮照片並建立 PDF 清冊…', index + 1, photos.length);
            if ((index + 1) % 3 === 0) await yieldToBrowser();
        }

        // 依三大版型分頁渲染
        for (let start = 0; start < preparedPhotos.length; start += perPage) {
            if (start > 0) pdf.addPage();
            const pageIndex = Math.floor(start / perPage) + 1;

            if (layout === 'up_down_2') {
                const marginLeft = 1797 * DXA_TO_MM;
                const marginTop = 1077 * DXA_TO_MM;
                const marginBottom = 907 * DXA_TO_MM;
                const tableW = 8302 * DXA_TO_MM;
                const cols = [553, 554, 731, 1559, 709, 1701, 142, 709, 1644].map(d => d * DXA_TO_MM);
                const maxImgW = 535 * PX_TO_MM;
                const maxImgH = 365 * PX_TO_MM;
                const gapMm = 100 * DXA_TO_MM;
                const maxSingleTableH = (pageHeight - marginTop - marginBottom - gapMm - 4) / 2;

                drawHeaderFooterText(`${dept}蒐證照片`, marginLeft, 10.5, tableW, 18, true);

                let curY = marginTop;
                for (let slot = 0; slot < 2; slot++) {
                    const p = preparedPhotos[start + slot];
                    if (!p) break;
                    const tableH = renderTable(marginLeft, curY, cols, [
                        {
                            minHeightMm: 420 * DXA_TO_MM,
                            cells: [
                                { colSpan: 2, text: '案由', fontPt: 12, bold: true },
                                { colSpan: 2, text: title, fontPt: 11, maxHeightMm: 14 },
                                { text: '日期', fontPt: 12, bold: true },
                                { colSpan: 2, text: p.item.date || date, fontPt: 10 },
                                { text: '時間', fontPt: 12, bold: true },
                                { text: p.item.time || '', fontPt: 10 }
                            ]
                        },
                        {
                            isPhotoRow: true,
                            exactHeightMm: 5726 * DXA_TO_MM,
                            cells: [
                                { colSpan: 9, photoDataUrl: p.dataUrl, maxImgWMm: maxImgW, maxImgHMm: maxImgH }
                            ]
                        },
                        {
                            minHeightMm: 375 * DXA_TO_MM,
                            cells: [
                                { text: '編號', fontPt: 11, bold: true, padXMm: 0.35, maxHeightMm: 6.5 },
                                { text: p.item.seq || '', fontPt: 11, bold: true, padXMm: 0.35, maxHeightMm: 6.5 },
                                { text: '地點：', fontPt: 11, bold: true, padXMm: 0.35, maxHeightMm: 6.5 },
                                { colSpan: 3, text: p.item.location || loc, fontPt: 11, padXMm: 0.8 },
                                { colSpan: 2, text: '製作人', fontPt: 11, bold: true, padXMm: 0.5, maxHeightMm: 6.5 },
                                { text: officer, fontPt: 11, padXMm: 0.8 }
                            ]
                        },
                        {
                            minHeightMm: 664 * DXA_TO_MM,
                            cells: [
                                { colSpan: 2, text: '說明', fontPt: 11, bold: true },
                                { colSpan: 7, text: p.item.desc || '現場照片紀錄。', fontPt: 10, maxHeightMm: 22 }
                            ]
                        }
                    ], maxSingleTableH);
                    curY += tableH + gapMm;
                }

                drawHeaderFooterText(`第 ${pageIndex} 頁 / 共 ${totalPages} 頁`, marginLeft, pageHeight - 10, tableW, 10, false);
            } else if (layout === 'left_right_2') {
                const marginLeft = 1020 * DXA_TO_MM;
                const marginTop = 1020 * DXA_TO_MM;
                const marginBottom = 1020 * DXA_TO_MM;
                const tableW = 9864 * DXA_TO_MM;
                const cols = [1000, 3932, 1000, 3932].map(d => d * DXA_TO_MM);
                const maxImgW = 315 * PX_TO_MM;
                const maxImgH = 690 * PX_TO_MM;
                const p1 = preparedPhotos[start];
                const p2 = preparedPhotos[start + 1] || null;
                const maxTableH = pageHeight - marginTop - marginBottom - 4;

                drawHeaderFooterText(`${dept}蒐證照片`, marginLeft, 9.5, tableW, 18, true);

                renderTable(marginLeft, marginTop, cols, [
                    {
                        minHeightMm: 500 * DXA_TO_MM,
                        cells: [
                            { colSpan: 4, text: title, fontPt: 16, bold: true, maxHeightMm: 16 }
                        ]
                    },
                    {
                        minHeightMm: 620 * DXA_TO_MM,
                        cells: [
                            {
                                colSpan: 2,
                                align: 'left',
                                runs: [
                                    { prefixBold: '蒐證時間：', text: `${p1.item.date || date} ${p1.item.time || ''}`.trim(), fontPt: 10 },
                                    { prefixBold: '蒐證地點：', text: p1.item.location || loc, fontPt: 10 }
                                ]
                            },
                            {
                                colSpan: 2,
                                align: 'left',
                                runs: [
                                    { prefixBold: '蒐證單位：', text: officer, fontPt: 10 }
                                ]
                            }
                        ]
                    },
                    {
                        isPhotoRow: true,
                        exactHeightMm: 10642 * DXA_TO_MM,
                        cells: [
                            { colSpan: 2, photoDataUrl: p1.dataUrl, maxImgWMm: maxImgW, maxImgHMm: maxImgH },
                            { colSpan: 2, photoDataUrl: p2 ? p2.dataUrl : null, maxImgWMm: maxImgW, maxImgHMm: maxImgH }
                        ]
                    },
                    {
                        minHeightMm: 380 * DXA_TO_MM,
                        cells: [
                            { text: '編號', fontPt: 11, bold: true },
                            { text: p1.item.seq || '', fontPt: 11 },
                            { text: '編號', fontPt: 11, bold: true },
                            { text: p2 ? (p2.item.seq || '') : '', fontPt: 11 }
                        ]
                    },
                    {
                        minHeightMm: 1200 * DXA_TO_MM,
                        cells: [
                            { text: '說明', fontPt: 11, bold: true },
                            { text: p1.item.desc || '無說明', fontPt: 10, align: 'left', valign: 'top', maxHeightMm: 36 },
                            { text: '說明', fontPt: 11, bold: true },
                            { text: p2 ? (p2.item.desc || '無說明') : '', fontPt: 10, align: 'left', valign: 'top', maxHeightMm: 36 }
                        ]
                    }
                ], maxTableH);

                drawHeaderFooterText(`第 ${pageIndex} 頁 / 共 ${totalPages} 頁`, marginLeft, pageHeight - 10, tableW, 10, false);
            } else {
                const marginLeft = 510 * DXA_TO_MM;
                const tableTop = 13.5;
                const marginBottom = 680 * DXA_TO_MM;
                const tableW = 15648 * DXA_TO_MM;
                const cols = [1266, 3950, 1295, 3921, 1182, 4034].map(d => d * DXA_TO_MM);
                const maxImgW = 335 * PX_TO_MM;
                const maxImgH = 535 * PX_TO_MM;
                const p1 = preparedPhotos[start];
                const p2 = preparedPhotos[start + 1] || null;
                const p3 = preparedPhotos[start + 2] || null;
                const maxTableH = pageHeight - tableTop - marginBottom - 2;

                drawHeaderFooterText(`${dept}${title}蒐證照片`, marginLeft, 4.5, tableW, 18, true);

                renderTable(marginLeft, tableTop, cols, [
                    {
                        minHeightMm: 420 * DXA_TO_MM,
                        cells: [
                            { text: '蒐證時間', fontPt: 12, bold: true },
                            { text: `${p1.item.date || date} ${p1.item.time || ''}`.trim(), fontPt: 12 },
                            { text: '蒐證地點', fontPt: 12, bold: true },
                            { text: p1.item.location || loc, fontPt: 12 },
                            { text: '製作人', fontPt: 12, bold: true },
                            { text: officer, fontPt: 12 }
                        ]
                    },
                    {
                        minHeightMm: 380 * DXA_TO_MM,
                        cells: [
                            { text: '照片編號', fontPt: 12, bold: true },
                            { text: p1.item.seq || '', fontPt: 12 },
                            { text: '照片編號', fontPt: 12, bold: true },
                            { text: p2 ? (p2.item.seq || '') : '', fontPt: 12 },
                            { text: '照片編號', fontPt: 12, bold: true },
                            { text: p3 ? (p3.item.seq || '') : '', fontPt: 12 }
                        ]
                    },
                    {
                        isPhotoRow: true,
                        exactHeightMm: 8254 * DXA_TO_MM,
                        cells: [
                            { colSpan: 2, photoDataUrl: p1.dataUrl, maxImgWMm: maxImgW, maxImgHMm: maxImgH },
                            { colSpan: 2, photoDataUrl: p2 ? p2.dataUrl : null, maxImgWMm: maxImgW, maxImgHMm: maxImgH },
                            { colSpan: 2, photoDataUrl: p3 ? p3.dataUrl : null, maxImgWMm: maxImgW, maxImgHMm: maxImgH }
                        ]
                    },
                    {
                        minHeightMm: 962 * DXA_TO_MM,
                        cells: [
                            { text: '照片說明', fontPt: 12, bold: true },
                            { text: p1.item.desc || '無', fontPt: 10, align: 'left', valign: 'top', maxHeightMm: 30 },
                            { text: '照片說明', fontPt: 12, bold: true },
                            { text: p2 ? (p2.item.desc || '無') : '', fontPt: 10, align: 'left', valign: 'top', maxHeightMm: 30 },
                            { text: '照片說明', fontPt: 12, bold: true },
                            { text: p3 ? (p3.item.desc || '無') : '', fontPt: 10, align: 'left', valign: 'top', maxHeightMm: 30 }
                        ]
                    }
                ], maxTableH);

                drawHeaderFooterText(`第 ${pageIndex} 頁 / 共 ${totalPages} 頁`, marginLeft, pageHeight - 7, tableW, 10, false);
            }
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
