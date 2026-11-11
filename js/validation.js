/**
 * validation.js - 日期與時間格式合法性驗證模組
 * 遵循專案既有民國日期格式（7碼純數字或斜線分隔）與時間範圍（4碼/6碼/冒號分隔）
 */
(function(root, factory) {
    if (typeof module === 'object' && typeof module.exports === 'object') {
        module.exports = factory();
    } else {
        root.PhotoReportValidation = factory();
    }
})(typeof self !== 'undefined' ? self : this, function() {
    'use strict';

    /**
     * 驗證民國日期合法性
     * @param {string} raw - 原始日期字串，支援 113/08/26 或 1130826
     * @returns {boolean}
     */
    function isValidMinguoDate(raw) {
        if (!raw || typeof raw !== 'string') return false;
        const digits = raw.replace(/[^0-9]/g, '');
        if (digits.length !== 7) return false;
        const year = parseInt(digits.slice(0, 3), 10);
        const month = parseInt(digits.slice(3, 5), 10);
        const day = parseInt(digits.slice(5, 7), 10);

        if (year < 1 || year > 999) return false;
        if (month < 1 || month > 12) return false;

        const daysInMonth = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
        const adYear = year + 1911;
        const isLeap = (adYear % 4 === 0 && adYear % 100 !== 0) || (adYear % 400 === 0);
        const maxDay = (month === 2 && isLeap) ? 29 : daysInMonth[month - 1];

        return day >= 1 && day <= maxDay;
    }

    /**
     * 驗證時間格式與數值範圍
     * @param {string} raw - 原始時間字串，支援 14:30、14:30:15、1430、143015
     * @returns {boolean}
     */
    function isValidTimeFormat(raw) {
        if (!raw || typeof raw !== 'string') return false;
        const trimmed = raw.trim();
        let h, m, s = 0;
        if (/^\d{2}:\d{2}$/.test(trimmed)) {
            [h, m] = trimmed.split(':').map(Number);
        } else if (/^\d{2}:\d{2}:\d{2}$/.test(trimmed)) {
            [h, m, s] = trimmed.split(':').map(Number);
        } else {
            const digits = trimmed.replace(/[^0-9]/g, '');
            if (digits.length === 4) {
                h = parseInt(digits.slice(0, 2), 10);
                m = parseInt(digits.slice(2, 4), 10);
            } else if (digits.length === 6) {
                h = parseInt(digits.slice(0, 2), 10);
                m = parseInt(digits.slice(2, 4), 10);
                s = parseInt(digits.slice(4, 6), 10);
            } else {
                return false;
            }
        }
        return h >= 0 && h <= 23 && m >= 0 && m <= 59 && s >= 0 && s <= 59;
    }

    const MAX_LOCATION_FULLWIDTH = 17;
    const MAX_DESC_FULLWIDTH = 68;
    const DESC_LINE_FULLWIDTH = 34;
    const MAX_DESC_LINES = 2;

    /**
     * 判斷單一字元（碼位）是否為半形字元（佔 0.5 全形 = 1 半形單位）
     * @param {string} ch
     * @returns {boolean}
     */
    function isHalfWidthChar(ch) {
        if (!ch) return true;
        const code = ch.codePointAt(0);
        if (code === undefined) return true;
        if ((code >= 0x0020 && code <= 0x007E) || code === 0x0009) return true;
        if (code >= 0xFF61 && code <= 0xFF9F) return true;
        return false;
    }

    /**
     * 計算字串佔用的半形單位數（半形=1，全形=2，換行符號不計寬）
     * @param {string} text
     * @returns {number}
     */
    function getTextHalfWidthUnits(text) {
        if (text === null || text === undefined) return 0;
        let units = 0;
        for (const ch of String(text)) {
            if (ch === '\r' || ch === '\n') continue;
            units += isHalfWidthChar(ch) ? 1 : 2;
        }
        return units;
    }

    /**
     * 計算字串佔用的全形字數（兩個半形算 1 個全形佔位）
     * @param {string} text
     * @returns {number}
     */
    function getTextFullWidthLength(text) {
        return getTextHalfWidthUnits(text) / 2;
    }

    /**
     * 格式化全形佔位顯示數字（整數或 .5）
     * @param {number} value
     * @returns {string}
     */
    function formatFullWidthCount(value) {
        const rounded = Math.round((Number(value) || 0) * 2) / 2;
        return Number.isInteger(rounded) ? String(rounded) : rounded.toFixed(1);
    }

    /**
     * 預估說明文字在報表儲存格內的視覺行數（單行約 34 全形 = 68 半形單位）
     * @param {string} text
     * @param {number} [lineFullWidth=34]
     * @returns {number}
     */
    function estimateDescVisualLines(text, lineFullWidth = DESC_LINE_FULLWIDTH) {
        const normalized = String(text ?? '').replace(/\r\n?/g, '\n');
        if (!normalized.length) return 0;
        const lineMaxUnits = Math.max(2, Math.round((Number(lineFullWidth) || DESC_LINE_FULLWIDTH) * 2));
        const segments = normalized.split('\n');
        let totalLines = 0;
        for (let i = 0; i < segments.length; i++) {
            const segUnits = getTextHalfWidthUnits(segments[i]);
            totalLines += Math.max(1, Math.ceil(segUnits / lineMaxUnits));
        }
        return totalLines;
    }

    /**
     * 依全形上限與最大視覺行數安全截斷字串（保留 UTF-16 Surrogate Pair 完整）
     * @param {string} text
     * @param {number} maxFullWidth
     * @param {Object} [options]
     * @param {boolean} [options.singleLine=false] - 單行欄位（如地點）將移除換行
     * @param {number} [options.maxLines=Infinity] - 最大視覺行數（如說明限制 2 行）
     * @param {number} [options.lineFullWidth=34] - 每行全形容量
     * @returns {{ text: string, halfWidthUnits: number, fullWidthLength: number, visualLines: number, truncated: boolean }}
     */
    function truncateToFullWidth(text, maxFullWidth, options = {}) {
        const {
            singleLine = false,
            maxLines = Infinity,
            lineFullWidth = DESC_LINE_FULLWIDTH
        } = options;
        const maxUnits = Math.max(0, Math.round((Number(maxFullWidth) || 0) * 2));
        const raw = String(text ?? '');
        let normalized = raw.replace(/\r\n?/g, '\n');
        let truncated = false;
        if (singleLine && normalized.includes('\n')) {
            normalized = normalized.replace(/\n+/g, '');
            truncated = true;
        }
        let result = '';
        let usedUnits = 0;
        for (const ch of normalized) {
            if (ch === '\n') {
                const candidate = result + '\n';
                if (Number.isFinite(maxLines) && estimateDescVisualLines(candidate, lineFullWidth) > maxLines) {
                    truncated = true;
                    continue;
                }
                result = candidate;
                continue;
            }
            const charUnits = isHalfWidthChar(ch) ? 1 : 2;
            if (usedUnits + charUnits > maxUnits) {
                truncated = true;
                break;
            }
            const candidate = result + ch;
            if (Number.isFinite(maxLines) && estimateDescVisualLines(candidate, lineFullWidth) > maxLines) {
                truncated = true;
                break;
            }
            result = candidate;
            usedUnits += charUnits;
        }
        return {
            text: result,
            halfWidthUnits: usedUnits,
            fullWidthLength: usedUnits / 2,
            visualLines: Number.isFinite(maxLines) ? estimateDescVisualLines(result, lineFullWidth) : (result ? 1 : 0),
            truncated: truncated
        };
    }

    /**
     * 檢查單張照片的有效地點與說明是否超過排版建議上限
     * @param {string} effectiveLocation
     * @param {string} desc
     * @returns {{ locationFullWidth: number, descFullWidth: number, descLines: number, locationOver: boolean, descOver: boolean, isOver: boolean }}
     */
    function inspectTextLimits(effectiveLocation, desc) {
        const locStr = String(effectiveLocation ?? '').trim();
        const descStr = String(desc ?? '').trim();
        const locationFullWidth = getTextFullWidthLength(locStr);
        const descFullWidth = getTextFullWidthLength(descStr);
        const descLines = estimateDescVisualLines(descStr, DESC_LINE_FULLWIDTH);
        const locationOver = locationFullWidth > MAX_LOCATION_FULLWIDTH;
        const descOver = descFullWidth > MAX_DESC_FULLWIDTH || descLines > MAX_DESC_LINES;
        return {
            locationFullWidth,
            descFullWidth,
            descLines,
            locationOver,
            descOver,
            isOver: locationOver || descOver
        };
    }

    return {
        MAX_LOCATION_FULLWIDTH: MAX_LOCATION_FULLWIDTH,
        MAX_DESC_FULLWIDTH: MAX_DESC_FULLWIDTH,
        DESC_LINE_FULLWIDTH: DESC_LINE_FULLWIDTH,
        MAX_DESC_LINES: MAX_DESC_LINES,
        isValidMinguoDate: isValidMinguoDate,
        isValidTimeFormat: isValidTimeFormat,
        isHalfWidthChar: isHalfWidthChar,
        getTextHalfWidthUnits: getTextHalfWidthUnits,
        getTextFullWidthLength: getTextFullWidthLength,
        formatFullWidthCount: formatFullWidthCount,
        estimateDescVisualLines: estimateDescVisualLines,
        truncateToFullWidth: truncateToFullWidth,
        inspectTextLimits: inspectTextLimits
    };
});
