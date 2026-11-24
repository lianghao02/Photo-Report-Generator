/**
 * auto-draft-store.js - IndexedDB 自動暫存草稿與案件匯出區控制代碼儲存模組
 * 封裝瀏覽器本機資料庫開啟、讀寫與清除，供 PhotoReportApp 委派使用。
 */
(function(root, factory) {
    if (typeof module === 'object' && typeof module.exports === 'object') {
        module.exports = factory();
    } else {
        root.PhotoReportAutoDraftStore = factory();
    }
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
    'use strict';

    function openDb({
        dbName = 'PhotoReportAutoDraftDB',
        draftStoreName = 'drafts',
        settingsStoreName = 'settings'
    } = {}) {
        return new Promise((resolve, reject) => {
            if (typeof indexedDB === 'undefined') {
                reject(new Error('瀏覽器不支援 IndexedDB'));
                return;
            }
            const request = indexedDB.open(dbName, 2);
            request.onupgradeneeded = () => {
                const db = request.result;
                if (!db.objectStoreNames.contains(draftStoreName)) {
                    db.createObjectStore(draftStoreName);
                }
                if (!db.objectStoreNames.contains(settingsStoreName)) {
                    db.createObjectStore(settingsStoreName);
                }
            };
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error || new Error('無法開啟自動暫存資料庫'));
        });
    }

    async function readKey(options, storeName, key) {
        const db = await openDb(options);
        try {
            return await new Promise((resolve, reject) => {
                const tx = db.transaction(storeName, 'readonly');
                const req = tx.objectStore(storeName).get(key);
                req.onsuccess = () => resolve(req.result || null);
                req.onerror = () => reject(req.error);
            });
        } finally {
            db.close();
        }
    }

    async function writeKey(options, storeName, key, value) {
        const db = await openDb(options);
        try {
            await new Promise((resolve, reject) => {
                const tx = db.transaction(storeName, 'readwrite');
                const store = tx.objectStore(storeName);
                if (value !== null && value !== undefined) store.put(value, key);
                else store.delete(key);
                tx.oncomplete = () => resolve();
                tx.onerror = () => reject(tx.error);
            });
        } finally {
            db.close();
        }
    }

    return {
        openDb,
        readKey,
        writeKey
    };
});
