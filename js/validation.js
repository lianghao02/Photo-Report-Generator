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

    return {
        isValidMinguoDate: isValidMinguoDate,
        isValidTimeFormat: isValidTimeFormat
    };
});
