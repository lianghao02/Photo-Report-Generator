/**
 * config.js - 全域預設參數集中管理 (Config-First)
 * 供主程式與 Word / PDF / Excel 匯出模組共用；使用者於介面刻意清空欄位時，匯出層將尊重空白不強行補回。
 */
(function(root, factory) {
    if (typeof module === 'object' && typeof module.exports === 'object') {
        module.exports = factory();
    } else {
        root.PhotoReportConfig = factory();
    }
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
    'use strict';

    return Object.freeze({
        DEFAULT_CASE_TITLE: '詐欺案',
        DEFAULT_REPORT_FALLBACK_TITLE: '現場照片清冊',
        DEFAULT_DEPT_NAME: '',
        DEFAULT_OFFICER_NAME: '',
        RECOMMENDED_MAX_PHOTOS: 500,
        RECOMMENDED_MAX_BYTES: 2 * 1024 * 1024 * 1024,
        AUTO_DRAFT_MAX_PHOTOS: 120,
        AUTO_DRAFT_MAX_BYTES: 200 * 1024 * 1024,
        GITHUB_RELEASES_URL: 'https://github.com/lianghao02/Photo-Report-Generator/releases',
        GITHUB_API_LATEST_RELEASE_URL: 'https://api.github.com/repos/lianghao02/Photo-Report-Generator/releases/latest'
    });
});
