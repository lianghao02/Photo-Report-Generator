/**
 * audit.js - 照片資料完整度稽核與同名照片計算模組
 * 純邏輯計算，不依賴 DOM、不操作 App State、不修改傳入陣列
 */
(function(root, factory) {
    if (typeof module === 'object' && typeof module.exports === 'object') {
        const validation = require('./validation');
        module.exports = factory(validation);
    } else {
        root.PhotoReportAudit = factory(root.PhotoReportValidation);
    }
})(typeof self !== 'undefined' ? self : this, function(validation) {
    'use strict';

    const isValidMinguoDate = validation?.isValidMinguoDate || function() { return false; };
    const isValidTimeFormat = validation?.isValidTimeFormat || function() { return false; };

    /**
     * 計算同名照片檔名集合
     * @param {Array<{name: string}>} photos
     * @returns {Set<string>}
     */
    function buildDuplicateNameSet(photos) {
        if (!Array.isArray(photos)) return new Set();
        const nameCount = new Map();
        photos.forEach(p => {
            const n = p?.name || '';
            if (!n) return;
            nameCount.set(n, (nameCount.get(n) || 0) + 1);
        });
        const dupSet = new Set();
        nameCount.forEach((count, name) => {
            if (count > 1) dupSet.add(name);
        });
        return dupSet;
    }

    /**
     * 稽核照片資料完整度
     * @param {Array<Object>} photos - 照片物件陣列
     * @param {string} [defaultLocation=''] - 全域預設地點
     * @returns {Object} 稽核統計結果物件
     */
    function auditPhotosCompleteness(photos, defaultLocation = '') {
        if (!Array.isArray(photos)) {
            return {
                total: 0,
                missingLocation: 0,
                missingDesc: 0,
                invalidDateTime: 0,
                duplicatePhotos: 0,
                issuesCount: 0,
                firstIssueFilter: null
            };
        }

        const trimmedLocation = String(defaultLocation || '').trim();
        let missingLocation = 0;
        let missingDesc = 0;
        let invalidDateTime = 0;
        let duplicatePhotos = 0;

        const dupSet = buildDuplicateNameSet(photos);

        photos.forEach(photo => {
            if (!photo) return;
            if (!photo.location?.trim() && !trimmedLocation) missingLocation++;
            if (!photo.desc?.trim()) missingDesc++;

            const d = photo.date?.trim() || '';
            const t = photo.time?.trim() || '';
            const dateInvalid = d ? !isValidMinguoDate(d) : false;
            const timeInvalid = !isValidTimeFormat(t);
            if (dateInvalid || timeInvalid) invalidDateTime++;

            if (dupSet.has(photo.name)) duplicatePhotos++;
        });

        const issuesCount = missingLocation + missingDesc + invalidDateTime + duplicatePhotos;
        let firstIssueFilter = null;
        if (missingDesc > 0) firstIssueFilter = 'missingDesc';
        else if (missingLocation > 0) firstIssueFilter = 'missingLocation';
        else if (invalidDateTime > 0) firstIssueFilter = 'invalidDateTime';
        else if (duplicatePhotos > 0) firstIssueFilter = 'duplicatePhotos';

        return {
            total: photos.length,
            missingLocation,
            missingDesc,
            invalidDateTime,
            duplicatePhotos,
            issuesCount,
            firstIssueFilter
        };
    }

    return {
        buildDuplicateNameSet: buildDuplicateNameSet,
        auditPhotosCompleteness: auditPhotosCompleteness
    };
});
