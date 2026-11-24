/**
 * export-directory.js - 案件匯出區命名、權限檢查與寫檔防呆模組
 * 純粹管理檔名／子資料夾名稱清洗、File System Access 授權與寫入衝突防呆。
 */
(function(root, factory) {
    if (typeof module === 'object' && typeof module.exports === 'object') {
        module.exports = factory();
    } else {
        root.PhotoReportExportDirectory = factory();
    }
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
    'use strict';

    function sanitizeFilename(rawName, fallback = '現場照片清冊') {
        const cleaned = String(rawName ?? '')
            .replace(/[\\/:*?"<>|\u0000-\u001F]/g, '_')
            .replace(/\s+/g, ' ')
            .trim();
        return cleaned || fallback;
    }

    function cleanFilenameSegment(rawValue, fallback = '', maxChars = 0) {
        let cleaned = String(rawValue ?? '')
            .replace(/[\\/:*?"<>|\u0000-\u001F\s]+/g, '_')
            .replace(/_+/g, '_')
            .replace(/^_+|_+$/g, '');
        if (maxChars > 0 && cleaned) {
            cleaned = Array.from(cleaned).slice(0, maxChars).join('').replace(/^_+|_+$/g, '');
        }
        return cleaned || fallback;
    }

    function buildCaseFolderName(caseData = {}) {
        const compactDate = String(caseData?.caseDate || '').replace(/[\\/:*?"<>|\s-]+/g, '');
        const cleanSubject = cleanFilenameSegment(caseData?.caseSubject, '', 15);
        const cleanTitle = cleanFilenameSegment(caseData?.caseTitle, '現場照片清冊');
        return [compactDate, cleanSubject, cleanTitle].filter(Boolean).join('_') || '現場照片清冊';
    }

    function buildProjectSuggestedFilename(caseDataOrTitle, now = new Date(), options = {}) {
        const caseTitle = typeof caseDataOrTitle === 'object' && caseDataOrTitle !== null
            ? caseDataOrTitle.caseTitle
            : caseDataOrTitle;
        const caseSubject = typeof caseDataOrTitle === 'object' && caseDataOrTitle !== null
            ? caseDataOrTitle.caseSubject
            : '';
        const cleanSubject = cleanFilenameSegment(caseSubject, '', 15);
        const cleanTitle = cleanFilenameSegment(caseTitle, '現場照片清冊');
        const baseName = [cleanSubject, cleanTitle].filter(Boolean).join('_') || '現場照片清冊';
        if (options?.fixedForFolder) {
            return `${baseName}.prp`;
        }
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const day = String(now.getDate()).padStart(2, '0');
        const hours = String(now.getHours()).padStart(2, '0');
        const minutes = String(now.getMinutes()).padStart(2, '0');
        return `${baseName}_${month}${day}_${hours}${minutes}.prp`;
    }

    function buildTimeSuffixedFilename(filename, now = new Date()) {
        const stamp = [
            String(now.getHours()).padStart(2, '0'),
            String(now.getMinutes()).padStart(2, '0'),
            String(now.getSeconds()).padStart(2, '0')
        ].join('');
        const dotIdx = filename.lastIndexOf('.');
        return dotIdx > 0
            ? `${filename.slice(0, dotIdx)}_${stamp}${filename.slice(dotIdx)}`
            : `${filename}_${stamp}`;
    }

    async function ensureDirectoryWritable(handle) {
        if (!handle) return false;
        if (typeof handle.queryPermission === 'function') {
            let status = await handle.queryPermission({ mode: 'readwrite' });
            if (status !== 'granted' && typeof handle.requestPermission === 'function') {
                status = await handle.requestPermission({ mode: 'readwrite' });
            }
            return status === 'granted';
        }
        return true;
    }

    async function prepareDirectoryPermission(handle, confirmFn = (typeof confirm === 'function' ? confirm : null)) {
        if (!handle) return 'download';
        let granted = false;
        try {
            granted = await ensureDirectoryWritable(handle);
        } catch {
            granted = false;
        }
        if (granted) return 'directory';
        const folderLabel = handle.name || '已選資料夾';
        const fallbackConfirm = typeof confirmFn === 'function'
            ? confirmFn(`瀏覽器目前沒有「${folderLabel}」的寫入權限（若剛重新開啟網頁，請按「取消」並點一下底部「存到：...」重新選取授權）。\n\n按「確定」＝本次先存進瀏覽器「下載」資料夾\n按「取消」＝先不匯出，去重新授權資料夾`)
            : false;
        return fallbackConfirm ? 'fallback' : 'cancel';
    }

    async function writeBlobToCaseDirectory({
        directoryHandle,
        folderName,
        filename,
        blob,
        sessionKnownExportPaths,
        confirmFn = (typeof confirm === 'function' ? confirm : null)
    }) {
        if (!directoryHandle) return { written: false };
        let granted = false;
        try {
            granted = await ensureDirectoryWritable(directoryHandle);
        } catch {
            granted = false;
        }
        if (!granted) {
            const fallbackConfirm = typeof confirmFn === 'function'
                ? confirmFn('尚未取得案件匯出區資料夾的寫入權限，要改用瀏覽器下載區儲存嗎？')
                : true;
            return { written: false, fallbackToDownload: fallbackConfirm };
        }

        const caseDirHandle = await directoryHandle.getDirectoryHandle(folderName, { create: true });
        let finalName = filename;
        const initialPathKey = `${folderName}/${finalName}`;
        if (!sessionKnownExportPaths || !sessionKnownExportPaths.has(initialPathKey)) {
            let exists = false;
            try {
                await caseDirHandle.getFileHandle(finalName, { create: false });
                exists = true;
            } catch (err) {
                if (err?.name !== 'NotFoundError') {
                    throw new Error(`無法檢查資料夾「${folderName}」內的「${finalName}」，可能權限不足或檔案正在被其他程式開啟（${err?.message || err?.name || '讀取異常'}）`);
                }
            }
            if (exists) {
                const overwrite = typeof confirmFn === 'function'
                    ? confirmFn(`資料夾「${folderName}」內已有同名檔案「${finalName}」，是否覆寫為最新版？\n（按「取消」會加上時間尾碼另存新檔，不會覆寫舊檔）`)
                    : true;
                if (!overwrite) {
                    finalName = buildTimeSuffixedFilename(finalName);
                }
            }
        }

        const fileHandle = await caseDirHandle.getFileHandle(finalName, { create: true });
        const writable = await fileHandle.createWritable();
        try {
            await writable.write(blob);
            await writable.close();
        } catch (error) {
            if (typeof writable.abort === 'function') {
                try { await writable.abort(); } catch {}
            }
            throw error;
        }
        if (sessionKnownExportPaths) {
            sessionKnownExportPaths.add(`${folderName}/${finalName}`);
        }
        return { written: true, folderName, fileName: finalName, fileHandle };
    }

    return {
        sanitizeFilename,
        cleanFilenameSegment,
        buildCaseFolderName,
        buildProjectSuggestedFilename,
        buildTimeSuffixedFilename,
        ensureDirectoryWritable,
        prepareDirectoryPermission,
        writeBlobToCaseDirectory
    };
});
