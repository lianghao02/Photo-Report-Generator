/**
 * audit-ui.js - 照片完整度篩選列 UI 控制器模組
 * 負責篩選列按鈕樣式、即時徽章更新與篩選提示文字渲染
 */
(function(root, factory) {
    if (typeof module === 'object' && typeof module.exports === 'object') {
        module.exports = factory();
    } else {
        root.PhotoReportAuditUi = factory();
    }
})(typeof self !== 'undefined' ? self : this, function() {
    'use strict';

    /**
     * 更新完整度篩選列 DOM
     * @param {Object} options
     * @param {HTMLElement} options.filterBar - 篩選列容器
     * @param {number} options.totalPhotos - 照片總數
     * @param {Object} options.audit - 稽核結果物件
     * @param {string} options.activeFilter - 目前篩選分類
     * @param {number} options.visibleCount - 目前可見照片數量
     * @param {Object} options.badges - 徽章元素物件 { all, missingLocation, missingDesc, invalidDateTime, duplicatePhotos }
     * @param {HTMLElement} options.filterNoticeText - 提示文字元素
     */
    function renderAuditBar({
        filterBar,
        totalPhotos,
        audit,
        activeFilter,
        visibleCount,
        badges = {},
        filterNoticeText
    }) {
        if (!filterBar) return;
        if (totalPhotos === 0) {
            filterBar.classList.add('hidden');
            return;
        }
        filterBar.classList.remove('hidden');

        if (badges.all) badges.all.textContent = totalPhotos;
        if (badges.missingLocation) {
            badges.missingLocation.textContent = audit.missingLocation;
            badges.missingLocation.classList.toggle('has-warning', audit.missingLocation > 0);
        }
        if (badges.missingDesc) {
            badges.missingDesc.textContent = audit.missingDesc;
            badges.missingDesc.classList.toggle('has-warning', audit.missingDesc > 0);
        }
        if (badges.invalidDateTime) {
            badges.invalidDateTime.textContent = audit.invalidDateTime;
            badges.invalidDateTime.classList.toggle('has-warning', audit.invalidDateTime > 0);
        }
        if (badges.duplicatePhotos) {
            badges.duplicatePhotos.textContent = audit.duplicatePhotos;
            badges.duplicatePhotos.classList.toggle('has-warning', audit.duplicatePhotos > 0);
        }

        // 更新按鈕 active 樣式
        const filterIds = {
            'all': 'filterBtn_all',
            'missingLocation': 'filterBtn_missingLocation',
            'missingDesc': 'filterBtn_missingDesc',
            'invalidDateTime': 'filterBtn_invalidDateTime',
            'duplicatePhotos': 'filterBtn_duplicatePhotos',
        };
        Object.entries(filterIds).forEach(([key, id]) => {
            const btn = document.getElementById(id);
            if (!btn) return;
            btn.classList.toggle('active', activeFilter === key);
        });

        // 篩選提示文字
        if (filterNoticeText) {
            if (activeFilter === 'all') {
                filterNoticeText.classList.add('hidden');
            } else {
                const labels = {
                    missingLocation: '未填地點',
                    missingDesc: '未填說明',
                    invalidDateTime: '時間異常',
                    duplicatePhotos: '同名照片',
                };
                const label = labels[activeFilter] || '';
                filterNoticeText.innerHTML = `<i class="fa-solid fa-filter text-[#5B5CE2]"></i> <span>篩選中：${label}（共 ${visibleCount} 張）</span>`;
                filterNoticeText.classList.remove('hidden');
            }
        }
    }

    return {
        renderAuditBar: renderAuditBar
    };
});
