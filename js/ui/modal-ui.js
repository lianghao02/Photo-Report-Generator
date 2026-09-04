/**
 * modal-ui.js - 通用對話框 Modal UI 控制器模組
 * 負責 Modal 開啟、關閉與匯出前稽核確認對話框資料注入
 */
(function(root, factory) {
    if (typeof module === 'object' && typeof module.exports === 'object') {
        module.exports = factory();
    } else {
        root.PhotoReportModalUi = factory();
    }
})(typeof self !== 'undefined' ? self : this, function() {
    'use strict';

    /**
     * 開啟指定 ID 之對話框
     * @param {string} id - Modal 元素 ID
     */
    function openModal(id) {
        const el = typeof id === 'string' ? document.getElementById(id) : id;
        if (el) el.classList.remove('hidden');
    }

    /**
     * 關閉指定 ID 之對話框
     * @param {string} id - Modal 元素 ID
     */
    function closeModal(id) {
        const el = typeof id === 'string' ? document.getElementById(id) : id;
        if (el) el.classList.add('hidden');
    }

    /**
     * 渲染並開啟匯出前完整度確認 Modal
     * @param {Object} options
     * @param {HTMLElement} options.modal - 對話框容器
     * @param {HTMLElement} options.titleEl - 標題元素
     * @param {HTMLElement} options.listEl - 缺失項目清單容器
     * @param {string} options.exportTitle - 匯出清冊名稱（如 Word 清冊）
     * @param {Object} options.audit - 稽核結果物件
     */
    function showExportAuditPrompt({ modal, titleEl, listEl, exportTitle, audit }) {
        if (!modal) return;
        if (listEl && audit) {
            const items = [];
            if (audit.missingDesc > 0) items.push(`未填說明：${audit.missingDesc} 張`);
            if (audit.missingLocation > 0) items.push(`未填地點：${audit.missingLocation} 張`);
            if (audit.invalidDateTime > 0) items.push(`時間異常：${audit.invalidDateTime} 張`);
            if (audit.duplicatePhotos > 0) items.push(`同名照片：${audit.duplicatePhotos} 張`);
            listEl.innerHTML = items.map(t => `<li>${t}</li>`).join('');
        }
        if (titleEl) {
            titleEl.textContent = `匯出前確認（${exportTitle}）`;
        }
        openModal(modal);
    }

    return {
        openModal: openModal,
        closeModal: closeModal,
        showExportAuditPrompt: showExportAuditPrompt
    };
});
