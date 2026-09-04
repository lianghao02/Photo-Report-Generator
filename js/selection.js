/**
 * selection.js - 照片選取、過濾與鍵盤導航索引計算模組
 * 純邏輯計算，不依賴 DOM、不處理指標拖曳事件
 */
(function(root, factory) {
    if (typeof module === 'object' && typeof module.exports === 'object') {
        const validation = require('./validation');
        const audit = require('./audit');
        module.exports = factory(validation, audit);
    } else {
        root.PhotoReportSelection = factory(root.PhotoReportValidation, root.PhotoReportAudit);
    }
})(typeof self !== 'undefined' ? self : this, function(validation, audit) {
    'use strict';

    const isValidMinguoDate = validation?.isValidMinguoDate || function() { return false; };
    const isValidTimeFormat = validation?.isValidTimeFormat || function() { return false; };
    const buildDuplicateNameSet = audit?.buildDuplicateNameSet || function() { return new Set(); };

    /**
     * 計算可見照片索引清單（依 activeFilter）
     * @param {Array<Object>} photos - 照片物件陣列
     * @param {string} [activeFilter='all'] - 目前篩選分類
     * @param {string} [defaultLocation=''] - 全域預設地點
     * @param {Set<string>} [cachedDupSet=null] - 預先計算好的重複檔名集合（可選）
     * @returns {Array<{photo: Object, index: number}>}
     */
    function getVisiblePhotoIndices(photos, activeFilter = 'all', defaultLocation = '', cachedDupSet = null) {
        if (!Array.isArray(photos)) return [];
        const all = photos.map((photo, index) => ({ photo, index }));
        if (activeFilter === 'all') return all;

        const globalLocation = String(defaultLocation || '').trim();
        let dupSet = cachedDupSet;

        return all.filter(({ photo }) => {
            if (!photo) return false;
            switch (activeFilter) {
                case 'missingLocation':
                    return !photo.location?.trim() && !globalLocation;
                case 'missingDesc':
                    return !photo.desc?.trim();
                case 'invalidDateTime': {
                    const d = photo.date?.trim() || '';
                    const t = photo.time?.trim() || '';
                    const dateInvalid = d ? !isValidMinguoDate(d) : false;
                    const timeInvalid = !isValidTimeFormat(t);
                    return dateInvalid || timeInvalid;
                }
                case 'duplicatePhotos': {
                    if (!dupSet) dupSet = buildDuplicateNameSet(photos);
                    return dupSet.has(photo.name);
                }
                default:
                    return true;
            }
        });
    }

    /**
     * 取得已選取之照片索引陣列；若未選取任何照片，則回傳目前焦點照片索引
     * @param {Array<Object>} photos
     * @param {number} currentIndex
     * @returns {Array<number>}
     */
    function getSelectedIndicesOrCurrent(photos, currentIndex) {
        if (!Array.isArray(photos)) return [];
        const selectedIndices = photos
            .map((photo, index) => (photo && photo.selected ? index : -1))
            .filter(index => index !== -1);
        if (selectedIndices.length) return selectedIndices;
        return photos[currentIndex] ? [currentIndex] : [];
    }

    /**
     * 計算鍵盤導航後的目標索引
     * @param {Object} options
     * @param {number} options.totalPhotos - 照片總數
     * @param {number} options.currentIndex - 目前索引
     * @param {number} options.delta - 移動增量（例如 -1, +1, -3, +3）
     * @param {string} [options.activeFilter='all'] - 目前篩選分類
     * @param {Array<{index: number}>} [options.visibleEntries=[]] - 可見項目陣列
     * @returns {number} 目標索引
     */
    function computeKeyboardNavIndex({ totalPhotos, currentIndex, delta, activeFilter = 'all', visibleEntries = [] }) {
        if (!totalPhotos || !delta) return currentIndex;
        if (activeFilter !== 'all') {
            if (!visibleEntries.length) return currentIndex;
            const currentPos = visibleEntries.findIndex(({ index }) => index === currentIndex);
            const pos = currentPos < 0 ? 0 : currentPos;
            const nextPos = Math.max(0, Math.min(pos + delta, visibleEntries.length - 1));
            return visibleEntries[nextPos].index;
        } else {
            return Math.max(0, Math.min(currentIndex + delta, totalPhotos - 1));
        }
    }

    return {
        getVisiblePhotoIndices: getVisiblePhotoIndices,
        getSelectedIndicesOrCurrent: getSelectedIndicesOrCurrent,
        computeKeyboardNavIndex: computeKeyboardNavIndex
    };
});
