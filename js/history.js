/**
 * history.js - 專案快照、簽名比對與 Undo/Redo 歷史管理器模組
 * 遵循鐵律：僅保存純資料屬性與案件資訊，UI 視圖狀態（activeFilter, zoom, scroll 等）不進歷史
 */
(function(root, factory) {
    if (typeof module === 'object' && typeof module.exports === 'object') {
        module.exports = factory();
    } else {
        root.PhotoReportHistory = factory();
    }
})(typeof self !== 'undefined' ? self : this, function() {
    'use strict';

    /**
     * 計算歷史快照特徵簽名
     * 僅提取 uid, rotation, seq, date, time, location, desc, selected, stageX, stageY
     * @param {Object} state - 快照物件
     * @returns {string}
     */
    function historySignature(state) {
        if (!state) return '';
        const photos = Array.isArray(state.photos) ? state.photos : [];
        return JSON.stringify({
            caseData: state.caseData || {},
            currentIndex: state.currentIndex ?? 0,
            lastSelectedIndex: state.lastSelectedIndex ?? null,
            photos: photos.map(({ uid, rotation, seq, date, time, location, desc, selected, stageX, stageY }) =>
                ({ uid, rotation, seq, date, time, location, desc, selected, stageX, stageY }))
        });
    }

    /**
     * 計算專案未存盤簽名（不含 selected/currentIndex/lastSelectedIndex 等編輯器暫存狀態）
     * @param {Object} caseData
     * @param {Array<Object>} photos
     * @returns {string}
     */
    function projectSignature(caseData, photos) {
        const photoList = Array.isArray(photos) ? photos : [];
        return JSON.stringify({
            caseData: caseData || {},
            photos: photoList.map(({ uid, rotation, seq, date, time, location, desc, stageX, stageY }) =>
                ({ uid, rotation, seq, date, time, location, desc, stageX, stageY }))
        });
    }

    /**
     * 歷史管理器類別
     */
    class HistoryManager {
        constructor(limit = 50) {
            this.history = [];
            this.historyIndex = -1;
            this.historyLimit = limit;
            this.historyTimer = null;
            this.isRestoringHistory = false;
        }

        canUndo() {
            return this.historyIndex > 0;
        }

        canRedo() {
            return this.historyIndex < this.history.length - 1;
        }

        record(nextState) {
            if (this.isRestoringHistory) return false;
            const current = this.history[this.historyIndex];
            if (current && historySignature(current) === historySignature(nextState)) {
                return false;
            }
            this.history.splice(this.historyIndex + 1);
            this.history.push(nextState);
            if (this.history.length > this.historyLimit) {
                this.history.shift();
            }
            this.historyIndex = this.history.length - 1;
            return true;
        }

        getCurrentState() {
            return this.history[this.historyIndex] || null;
        }

        undo() {
            if (!this.canUndo()) return null;
            this.historyIndex--;
            return this.history[this.historyIndex];
        }

        redo() {
            if (!this.canRedo()) return null;
            this.historyIndex++;
            return this.history[this.historyIndex];
        }

        clear() {
            this.history = [];
            this.historyIndex = -1;
            if (this.historyTimer) {
                clearTimeout(this.historyTimer);
                this.historyTimer = null;
            }
        }
    }

    return {
        historySignature: historySignature,
        projectSignature: projectSignature,
        HistoryManager: HistoryManager
    };
});
