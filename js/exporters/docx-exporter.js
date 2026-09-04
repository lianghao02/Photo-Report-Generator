/**
 * docx-exporter.js - Word (.docx) 清冊匯出模組
 * 封裝三大公務排版（上下兩張 8302 dxa、左右兩張 9864 dxa、橫式三張 15648 dxa）、
 * 表格 OpenXML 建立、標楷體字型、等比縮放、固定 5 點行距與跨頁 Header/Footer。
 */
(function(root, factory) {
    if (typeof module === 'object' && typeof module.exports === 'object') {
        module.exports = factory();
    } else {
        root.PhotoReportDocxExporter = factory();
    }
})(typeof self !== 'undefined' ? self : this, function() {
    'use strict';

    function getDocxLib() {
        if (typeof docx !== 'undefined') return docx;
        if (typeof window !== 'undefined' && window.docx) return window.docx;
        return null;
    }

    async function exportDocx({
        photos = [],
        caseTitle = '現場照片清冊',
        caseDate = '',
        defaultLocation = '',
        deptName = '臺南市政府警察局新化分局',
        globalOfficer = '巡官梁家豪',
        layout = 'up_down_2',
        processImageToBlob,
        setProgress = () => {},
        clearProgress = () => {},
        throwIfCancelled = () => {},
        yieldToBrowser = () => Promise.resolve(),
        docxLib = getDocxLib(),
        saveAs = (typeof window !== 'undefined' && window.saveAs ? window.saveAs : (typeof saveAs !== 'undefined' ? saveAs : null))
    } = {}) {
        if (!photos.length) {
            if (typeof alert === 'function') alert('請先載入照片再匯出！');
            return;
        }

        const lib = docxLib || getDocxLib();
        if (!lib) {
            if (typeof alert === 'function') alert('docx.js 函式庫尚未載入完成，請確認網路連線是否正常！');
            return;
        }

        const title = (caseTitle || '').trim() || '現場照片清冊';
        const date = (caseDate || '').trim() || '';
        const loc = (defaultLocation || '').trim() || '';
        const dept = (deptName || '').trim() || '臺南市政府警察局新化分局';
        const officer = (globalOfficer || '').trim() || '巡官梁家豪';

        const {
            Document, Packer, Paragraph, Table, TableRow, TableCell,
            TextRun, WidthType, AlignmentType, HeightRule,
            ImageRun, BorderStyle, Header, Footer, PageNumber, PageOrientation, VerticalAlign,
            LineRuleType
        } = lib;

        const borderSingle = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
        const bordersAllSingle = { top: borderSingle, bottom: borderSingle, left: borderSingle, right: borderSingle };

        const processedPhotos = [];
        setProgress('正在壓縮照片並建立 Word 報表…', 0, photos.length);
        for (let i = 0; i < photos.length; i++) {
            const item = photos[i];
            const processed = await processImageToBlob(item, 1600, 0.88);
            const arrayBuffer = await processed.blob.arrayBuffer();
            processedPhotos.push({
                data: new Uint8Array(arrayBuffer),
                w: processed.width,
                h: processed.height,
                item: item
            });
            setProgress('正在壓縮照片並建立 Word 報表…', i + 1, photos.length);
            throwIfCancelled();
            if ((i + 1) % 5 === 0) await yieldToBrowser();
        }

        let docSections = [];

                // -------------------------------------------------------------
                // 1. 經典版型：上下兩張 (A4 直式，100% 精準對齊 上下兩張.docx 範本)
                // -------------------------------------------------------------
                if (layout === 'up_down_2') {
                    const children = [];
                    // 圖片儲存格可用範圍：約 535 × 365px；以較先觸頂的一邊決定等比例縮放
                    const MAX_IMG_W = 535;
                    const MAX_IMG_H = 365;

                    for (let i = 0; i < processedPhotos.length; i++) {
                        const p = processedPhotos[i];
                        let imgW = p.w;
                        let imgH = p.h;
                        const ratio = Math.min(MAX_IMG_W / imgW, MAX_IMG_H / imgH);
                        imgW = Math.round(imgW * ratio);
                        imgH = Math.round(imgH * ratio);

                        // 構建與 上下兩張.docx 完全相同的 4 列 9 欄 OpenXML 表格
                        const table = new Table({
                            width: { size: 8302, type: WidthType.DXA },
                            columnWidths: [553, 554, 731, 1559, 709, 1701, 142, 709, 1644],
                            cantSplit: true,
                            borders: bordersAllSingle,
                            rows: [
                                // Row 1: 案由 | [案由名稱] | 日期 | [採證日期] | 時間 | [時間]
                                new TableRow({
                                    height: { value: 420, rule: HeightRule.ATLEAST },
                                    children: [
                                        new TableCell({ width: { size: 1107, type: WidthType.DXA }, columnSpan: 2, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "案由", bold: true, font: "標楷體", size: 24 })] })] }),
                                        new TableCell({ width: { size: 2290, type: WidthType.DXA }, columnSpan: 2, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: title, font: "標楷體", size: 22 })] })] }),
                                        new TableCell({ width: { size: 709, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "日期", bold: true, font: "標楷體", size: 24 })] })] }),
                                        new TableCell({ width: { size: 1843, type: WidthType.DXA }, columnSpan: 2, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: p.item.date || date, font: "標楷體", size: 20 })] })] }),
                                        new TableCell({ width: { size: 709, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "時間", bold: true, font: "標楷體", size: 24 })] })] }),
                                        new TableCell({ width: { size: 1644, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: p.item.time, font: "標楷體", size: 20 })] })] }),
                                    ]
                                }),
                                // Row 2: 照片主圖 (全寬 8302 dxa，高度 5726 dxa，精確置中)
                                new TableRow({
                                    height: { value: 5726, rule: HeightRule.EXACT },
                                    children: [
                                        new TableCell({
                                            width: { size: 8302, type: WidthType.DXA },
                                            columnSpan: 9,
                                            borders: bordersAllSingle,
                                            verticalAlign: VerticalAlign.CENTER,
                                            children: [
                                                new Paragraph({
                                                    alignment: AlignmentType.CENTER,
                                                    children: [
                                                        new ImageRun({
                                                            data: p.data,
                                                            transformation: { width: imgW, height: imgH },
                                                        })
                                                    ]
                                                })
                                            ]
                                        })
                                    ]
                                }),
                                // Row 3: 編號 | [編號] | 地點： | [地點] | 製作人 | [製作人]
                                new TableRow({
                                    height: { value: 375, rule: HeightRule.ATLEAST },
                                    children: [
                                        new TableCell({ width: { size: 553, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "編號", bold: true, font: "標楷體", size: 22 })] })] }),
                                        new TableCell({ width: { size: 554, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: p.item.seq, bold: true, font: "標楷體", size: 22 })] })] }),
                                        new TableCell({ width: { size: 731, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "地點：", bold: true, font: "標楷體", size: 22 })] })] }),
                                        new TableCell({ width: { size: 3969, type: WidthType.DXA }, columnSpan: 3, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: p.item.location || loc, font: "標楷體", size: 22 })] })] }),
                                        new TableCell({ width: { size: 851, type: WidthType.DXA }, columnSpan: 2, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "製作人", bold: true, font: "標楷體", size: 22 })] })] }),
                                        new TableCell({ width: { size: 1644, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: officer, font: "標楷體", size: 22 })] })] }),
                                    ]
                                }),
                                // Row 4: 說明 | [現場跡證說明]
                                new TableRow({
                                    height: { value: 664, rule: HeightRule.ATLEAST },
                                    children: [
                                        new TableCell({ width: { size: 1107, type: WidthType.DXA }, columnSpan: 2, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "說明", bold: true, font: "標楷體", size: 22 })] })] }),
                                        new TableCell({
                                            width: { size: 7195, type: WidthType.DXA },
                                            columnSpan: 7,
                                            borders: bordersAllSingle,
                                            verticalAlign: VerticalAlign.CENTER,
                                            children: [
                                                new Paragraph({
                                                    alignment: AlignmentType.CENTER,
                                                    children: [new TextRun({ text: p.item.desc || "現場照片紀錄。", font: "標楷體", size: 20 })]
                                                })
                                            ]
                                        })
                                    ]
                                })
                            ]
                        });

                        children.push(table);

                        // 上下兩張表格之間：精確設定固定行高 5 點 (line: 100, lineRule: exact)，前後段距離 0
                        if (i % 2 === 0 && i < processedPhotos.length - 1) {
                            children.push(new Paragraph({
                                spacing: { line: 100, lineRule: LineRuleType.EXACT, before: 0, after: 0 }
                            }));
                        } else if (i % 2 === 1 && i < processedPhotos.length - 1) {
                            children.push(new Paragraph({
                                pageBreakBefore: true,
                                spacing: { line: 20, lineRule: LineRuleType.EXACT, before: 0, after: 0 }
                            }));
                        }
                    }

                    // 頁首 Header: 置中 標楷體 18pt 粗體
                    // 頁尾 Footer: 置中 標楷體 10pt (第 X 頁 / 共 Y 頁)
                    docSections.push({
                        properties: {
                            page: {
                                size: { width: 11906, height: 16838, orientation: PageOrientation.PORTRAIT },
                                margin: { top: 1077, right: 1797, bottom: 907, left: 1797, header: 850, footer: 567 }
                            }
                        },
                        headers: {
                            default: new Header({
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [new TextRun({ text: `${dept}蒐證照片`, bold: true, font: "標楷體", size: 36 })]
                                    })
                                ]
                            })
                        },
                        footers: {
                            default: new Footer({
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                            new TextRun({ text: "第 ", font: "標楷體", size: 20 }),
                                            new TextRun({ children: [PageNumber.CURRENT], font: "標楷體", size: 20 }),
                                            new TextRun({ text: " 頁 / 共 ", font: "標楷體", size: 20 }),
                                            new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "標楷體", size: 20 }),
                                            new TextRun({ text: " 頁", font: "標楷體", size: 20 }),
                                        ]
                                    })
                                ]
                            })
                        },
                        children: children
                    });
                }

                // -------------------------------------------------------------
                // 2. 對照版型：左右兩張 (A4 直式雙欄，100% 精準對齊 左右兩張.docx 範本)
                // -------------------------------------------------------------
                else if (layout === 'left_right_2') {
                    const children = [];
                    // 每欄圖片儲存格可用範圍：約 315 × 690px；以較先觸頂的一邊決定等比例縮放
                    const MAX_IMG_W = 315;
                    const MAX_IMG_H = 690;

                    for (let i = 0; i < processedPhotos.length; i += 2) {
                        const p1 = processedPhotos[i];
                        const p2 = processedPhotos[i + 1] || null;

                        let w1 = p1.w, h1 = p1.h;
                        let r1 = Math.min(MAX_IMG_W / w1, MAX_IMG_H / h1);
                        w1 = Math.round(w1 * r1);
                        h1 = Math.round(h1 * r1);

                        let w2 = p2 ? p2.w : 1, h2 = p2 ? p2.h : 1;
                        if (p2) {
                            let r2 = Math.min(MAX_IMG_W / w2, MAX_IMG_H / h2);
                            w2 = Math.round(w2 * r2);
                            h2 = Math.round(h2 * r2);
                        }

                        const table = new Table({
                            width: { size: 9864, type: WidthType.DXA },
                            columnWidths: [1000, 3932, 1000, 3932],
                            cantSplit: true,
                            borders: bordersAllSingle,
                            rows: [
                                // Row 1: 案由 (全寬 9864 dxa, 32pt/16pt 粗體置中)
                                new TableRow({
                                    height: { value: 500, rule: HeightRule.ATLEAST },
                                    children: [
                                        new TableCell({
                                            width: { size: 9864, type: WidthType.DXA },
                                            columnSpan: 4,
                                            borders: bordersAllSingle,
                                            verticalAlign: VerticalAlign.CENTER,
                                            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: title, bold: true, font: "標楷體", size: 32 })] })]
                                        })
                                    ]
                                }),
                                // Row 2: 蒐證時間/地點 (左) | 蒐證單位 (右)
                                new TableRow({
                                    height: { value: 400, rule: HeightRule.ATLEAST },
                                    children: [
                                        new TableCell({
                                            width: { size: 4932, type: WidthType.DXA },
                                            columnSpan: 2,
                                            borders: bordersAllSingle,
                                            verticalAlign: VerticalAlign.CENTER,
                                            children: [
                                                new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: "蒐證時間：", bold: true, font: "標楷體", size: 22 }), new TextRun({ text: `${p1.item.date || date} ${p1.item.time}`, font: "標楷體", size: 20 })] }),
                                                new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: "蒐證地點：", bold: true, font: "標楷體", size: 22 }), new TextRun({ text: p1.item.location || loc, font: "標楷體", size: 20 })] })
                                            ]
                                        }),
                                        new TableCell({
                                            width: { size: 4932, type: WidthType.DXA },
                                            columnSpan: 2,
                                            borders: bordersAllSingle,
                                            verticalAlign: VerticalAlign.CENTER,
                                            children: [
                                                new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: "蒐證單位：", bold: true, font: "標楷體", size: 22 }), new TextRun({ text: officer, font: "標楷體", size: 20 })] })
                                            ]
                                        })
                                    ]
                                }),
                                // Row 3: 左右雙照片 (各 4932 dxa，高度 10642 dxa，精確置中)
                                new TableRow({
                                    height: { value: 10642, rule: HeightRule.EXACT },
                                    children: [
                                        new TableCell({
                                            width: { size: 4932, type: WidthType.DXA },
                                            columnSpan: 2,
                                            borders: bordersAllSingle,
                                            verticalAlign: VerticalAlign.CENTER,
                                            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: p1.data, transformation: { width: w1, height: h1 } })] })]
                                        }),
                                        new TableCell({
                                            width: { size: 4932, type: WidthType.DXA },
                                            columnSpan: 2,
                                            borders: bordersAllSingle,
                                            verticalAlign: VerticalAlign.CENTER,
                                            children: [
                                                p2 ? new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: p2.data, transformation: { width: w2, height: h2 } })] })
                                                   : new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "（無照片）", font: "標楷體", size: 20 })] })
                                            ]
                                        })
                                    ]
                                }),
                                // Row 4: 編號 | [編號1] | 編號 | [編號2]
                                new TableRow({
                                    height: { value: 380, rule: HeightRule.ATLEAST },
                                    children: [
                                        new TableCell({ width: { size: 1000, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "編號", bold: true, font: "標楷體", size: 22 })] })] }),
                                        new TableCell({ width: { size: 3932, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: p1.item.seq, font: "標楷體", size: 22 })] })] }),
                                        new TableCell({ width: { size: 1000, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "編號", bold: true, font: "標楷體", size: 22 })] })] }),
                                        new TableCell({ width: { size: 3932, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: p2 ? p2.item.seq : "", font: "標楷體", size: 22 })] })] }),
                                    ]
                                }),
                                // Row 5: 說明 | [說明1] | 說明 | [說明2]
                                new TableRow({
                                    height: { value: 1200, rule: HeightRule.ATLEAST },
                                    children: [
                                        new TableCell({ width: { size: 1000, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "說明", bold: true, font: "標楷體", size: 22 })] })] }),
                                        new TableCell({ width: { size: 3932, type: WidthType.DXA }, borders: bordersAllSingle, children: [new Paragraph({ children: [new TextRun({ text: p1.item.desc || "無說明", font: "標楷體", size: 20 })] })] }),
                                        new TableCell({ width: { size: 1000, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "說明", bold: true, font: "標楷體", size: 22 })] })] }),
                                        new TableCell({ width: { size: 3932, type: WidthType.DXA }, borders: bordersAllSingle, children: [new Paragraph({ children: [new TextRun({ text: p2 ? (p2.item.desc || "無說明") : "", font: "標楷體", size: 20 })] })] }),
                                    ]
                                })
                            ]
                        });

                        children.push(table);

                        if (i + 2 < processedPhotos.length) {
                            children.push(new Paragraph({ pageBreakBefore: true }));
                        }
                    }

                    docSections.push({
                        properties: {
                            page: {
                                size: { width: 11906, height: 16838, orientation: PageOrientation.PORTRAIT },
                                margin: { top: 1020, right: 1020, bottom: 1020, left: 1020, header: 720, footer: 720 }
                            }
                        },
                        headers: {
                            default: new Header({
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [new TextRun({ text: `${dept}蒐證照片`, bold: true, font: "標楷體", size: 36 })]
                                    })
                                ]
                            })
                        },
                        footers: {
                            default: new Footer({
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                            new TextRun({ text: "第 ", font: "標楷體", size: 20 }),
                                            new TextRun({ children: [PageNumber.CURRENT], font: "標楷體", size: 20 }),
                                            new TextRun({ text: " 頁 / 共 ", font: "標楷體", size: 20 }),
                                            new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "標楷體", size: 20 }),
                                            new TextRun({ text: " 頁", font: "標楷體", size: 20 }),
                                        ]
                                    })
                                ]
                            })
                        },
                        children: children
                    });
                }

                // -------------------------------------------------------------
                // 3. 廣角版型：橫式三張 (A4 橫式三欄，100% 精準對齊 橫式三張.docx 範本)
                // -------------------------------------------------------------
                else if (layout === 'landscape_3') {
                    const children = [];
                    // 每欄圖片儲存格可用範圍：約 335 × 535px；以較先觸頂的一邊決定等比例縮放
                    const MAX_IMG_W = 335;
                    const MAX_IMG_H = 535;

                    for (let i = 0; i < processedPhotos.length; i += 3) {
                        const p1 = processedPhotos[i];
                        const p2 = processedPhotos[i + 1] || null;
                        const p3 = processedPhotos[i + 2] || null;

                        const calcSize = (p) => {
                            if (!p) return { w: 1, h: 1 };
                            let w = p.w, h = p.h;
                            let r = Math.min(MAX_IMG_W / w, MAX_IMG_H / h);
                            return { w: Math.round(w * r), h: Math.round(h * r) };
                        };

                        const s1 = calcSize(p1);
                        const s2 = calcSize(p2);
                        const s3 = calcSize(p3);

                        const table = new Table({
                            width: { size: 15648, type: WidthType.DXA },
                            columnWidths: [1266, 3950, 1295, 3921, 1182, 4034],
                            cantSplit: true,
                            borders: bordersAllSingle,
                            rows: [
                                // Row 1: 蒐證時間 | [時間] | 蒐證地點 | [地點] | 製作人 | [製作人]
                                new TableRow({
                                    height: { value: 420, rule: HeightRule.ATLEAST },
                                    children: [
                                        new TableCell({ width: { size: 1266, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: "蒐證時間", bold: true, font: "標楷體", size: 24 })] })] }),
                                        new TableCell({ width: { size: 3950, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: `${p1.item.date || date} ${p1.item.time}`, font: "標楷體", size: 24 })] })] }),
                                        new TableCell({ width: { size: 1295, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: "蒐證地點", bold: true, font: "標楷體", size: 24 })] })] }),
                                        new TableCell({ width: { size: 3921, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: p1.item.location || loc, font: "標楷體", size: 24 })] })] }),
                                        new TableCell({ width: { size: 1182, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: "製作人", bold: true, font: "標楷體", size: 24 })] })] }),
                                        new TableCell({ width: { size: 4034, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.TOP, children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: officer, font: "標楷體", size: 24 })] })] }),
                                    ]
                                }),
                                // Row 2: 照片編號 | [編號1] | 照片編號 | [編號2] | 照片編號 | [編號3]
                                new TableRow({
                                    height: { value: 380, rule: HeightRule.ATLEAST },
                                    children: [
                                        new TableCell({ width: { size: 1266, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: "照片編號", bold: true, font: "標楷體", size: 24 })] })] }),
                                        new TableCell({ width: { size: 3950, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: p1.item.seq, font: "標楷體", size: 24 })] })] }),
                                        new TableCell({ width: { size: 1295, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: "照片編號", bold: true, font: "標楷體", size: 24 })] })] }),
                                        new TableCell({ width: { size: 3921, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: p2 ? p2.item.seq : "", font: "標楷體", size: 24 })] })] }),
                                        new TableCell({ width: { size: 1182, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: "照片編號", bold: true, font: "標楷體", size: 24 })] })] }),
                                        new TableCell({ width: { size: 4034, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: p3 ? p3.item.seq : "", font: "標楷體", size: 24 })] })] }),
                                    ]
                                }),
                                // Row 3: 並排三照片 (各 5216 dxa，高度 8254 dxa，精確置中)
                                new TableRow({
                                    height: { value: 8254, rule: HeightRule.EXACT },
                                    children: [
                                        new TableCell({ width: { size: 5216, type: WidthType.DXA }, columnSpan: 2, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: p1.data, transformation: { width: s1.w, height: s1.h } })] })] }),
                                        new TableCell({ width: { size: 5216, type: WidthType.DXA }, columnSpan: 2, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [p2 ? new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: p2.data, transformation: { width: s2.w, height: s2.h } })] }) : new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "（無照片）", font: "標楷體", size: 20 })] })] }),
                                        new TableCell({ width: { size: 5216, type: WidthType.DXA }, columnSpan: 2, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [p3 ? new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: p3.data, transformation: { width: s3.w, height: s3.h } })] }) : new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "（無照片）", font: "標楷體", size: 20 })] })] }),
                                    ]
                                }),
                                // Row 4: 照片說明 | [說明1] | 照片說明 | [說明2] | 照片說明 | [說明3]
                                new TableRow({
                                    height: { value: 962, rule: HeightRule.ATLEAST },
                                    children: [
                                        new TableCell({ width: { size: 1266, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "照片說明", bold: true, font: "標楷體", size: 24 })] })] }),
                                        new TableCell({ width: { size: 3950, type: WidthType.DXA }, borders: bordersAllSingle, children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: p1.item.desc || "無", font: "標楷體", size: 20 })] })] }),
                                        new TableCell({ width: { size: 1295, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "照片說明", bold: true, font: "標楷體", size: 24 })] })] }),
                                        new TableCell({ width: { size: 3921, type: WidthType.DXA }, borders: bordersAllSingle, children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: p2 ? (p2.item.desc || "無") : "", font: "標楷體", size: 20 })] })] }),
                                        new TableCell({ width: { size: 1182, type: WidthType.DXA }, borders: bordersAllSingle, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "照片說明", bold: true, font: "標楷體", size: 24 })] })] }),
                                        new TableCell({ width: { size: 4034, type: WidthType.DXA }, borders: bordersAllSingle, children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: p3 ? (p3.item.desc || "無") : "", font: "標楷體", size: 20 })] })] }),
                                    ]
                                })
                            ]
                        });

                        children.push(table);

                        if (i + 3 < processedPhotos.length) {
                            children.push(new Paragraph({ pageBreakBefore: true }));
                        }
                    }

                    // 橫式三張 Header: 置中 標楷體 18pt 粗體
                    // 橫式三張 Footer: 置中 標楷體 10pt (第 X 頁 / 共 Y 頁)
                    docSections.push({
                        properties: {
                            page: {
                                size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE },
                                margin: { top: 454, right: 510, bottom: 680, left: 510, header: 283, footer: 283 }
                            }
                        },
                        headers: {
                            default: new Header({
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [new TextRun({ text: `${dept}${title}蒐證照片`, bold: true, font: "標楷體", size: 36 })]
                                    })
                                ]
                            })
                        },
                        footers: {
                            default: new Footer({
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                            new TextRun({ text: "第 ", font: "標楷體", size: 20 }),
                                            new TextRun({ children: [PageNumber.CURRENT], font: "標楷體", size: 20 }),
                                            new TextRun({ text: " 頁 / 共 ", font: "標楷體", size: 20 }),
                                            new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "標楷體", size: 20 }),
                                            new TextRun({ text: " 頁", font: "標楷體", size: 20 }),
                                        ]
                                    })
                                ]
                            })
                        },
                        children: children
                    });
                }
        const doc = new Document({
            styles: {
                default: {
                    document: {
                        run: {
                            font: "標楷體",
                            size: 22,
                            color: "000000"
                        }
                    }
                }
            },
            sections: docSections
        });

        const blob = await Packer.toBlob(doc);
        const safeName = `${date || '報告'}_${title}_現場照片清冊.docx`.replace(/[\\/:*?"<>|]/g, '_');
        if (typeof saveAs === 'function') {
            saveAs(blob, safeName);
        }
        clearProgress();
    }

    return {
        getDocxLib: getDocxLib,
        exportDocx: exportDocx
    };
});