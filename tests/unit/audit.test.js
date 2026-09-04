const assert = require('assert');
const { buildDuplicateNameSet, auditPhotosCompleteness } = require('../../js/audit');
const { loadAppMethods } = require('./app-bridge');

const app = loadAppMethods();

console.log('--- 測試 audit.test.js (包含獨立模組與 App 委派) ---');

// 1. 同名照片集合計算
const samplePhotoNames = [
    { name: 'A.jpg' },
    { name: 'B.jpg' },
    { name: 'A.jpg' },
    { name: 'C.jpg' },
    { name: 'D.jpg' },
    { name: 'B.jpg' },
    { name: 'E.jpg' }
];

// 測試獨立模組 buildDuplicateNameSet
const dupSetModule = buildDuplicateNameSet(samplePhotoNames);
assert.strictEqual(dupSetModule instanceof Set, true, '獨立模組 dupSet 應為 Set 物件');
assert.strictEqual(dupSetModule.has('A.jpg'), true, 'A.jpg 出現2次應在 dupSet 內');
assert.strictEqual(dupSetModule.has('B.jpg'), true, 'B.jpg 出現2次應在 dupSet 內');
assert.strictEqual(dupSetModule.has('C.jpg'), false, 'C.jpg 僅出現1次不應在 dupSet 內');
assert.strictEqual(dupSetModule.size, 2, '應只有 2 組重複檔名');

// 測試 App 委派 _buildDupSet
app.photos = samplePhotoNames;
app._buildDupSet();
assert.strictEqual(app._dupSet instanceof Set, true, 'App._dupSet 應為 Set 物件');
assert.strictEqual(app._dupSet.has('A.jpg'), true, 'A.jpg 應在 App._dupSet 內');
assert.strictEqual(app._dupSet.has('B.jpg'), true, 'B.jpg 應在 App._dupSet 內');
assert.strictEqual(app._dupSet.size, 2, 'App._dupSet 大小應為 2');

// 2. 完整度稽核 auditPhotosCompleteness
global.document = {
    getElementById: (id) => {
        if (id === 'defaultLocation') return { value: '' };
        return null;
    }
};

const samplePhotos = [
    {
        uid: 'p1',
        name: 'IMG_01.jpg',
        date: '113/08/26',
        time: '14:30',
        location: '現場',
        desc: '正面外觀'
    }, // 完整無缺失
    {
        uid: 'p2',
        name: 'IMG_02.jpg',
        date: '113/08/26',
        time: '14:31',
        location: '',
        desc: '客廳內部'
    }, // 缺失地點
    {
        uid: 'p3',
        name: 'IMG_03.jpg',
        date: '113/08/26',
        time: '14:32',
        location: '臥室',
        desc: ''
    }, // 缺失說明
    {
        uid: 'p4',
        name: 'IMG_04.jpg',
        date: '114/02/29', // 平年非法2月29日
        time: '14:33',
        location: '陽台',
        desc: '地面痕跡'
    }, // 日期異常
    {
        uid: 'p5',
        name: 'IMG_05.jpg',
        date: '113/08/26',
        time: '24:00', // 非法時間
        location: '玄關',
        desc: '門鎖受損'
    }, // 時間異常
    {
        uid: 'p6',
        name: 'IMG_01.jpg', // 與 p1 同名
        date: '113/08/26',
        time: '14:35',
        location: '現場',
        desc: '特寫鏡頭'
    } // 同名照片
];

const originalPhotosJson = JSON.stringify(samplePhotos);

// 測試獨立模組 auditPhotosCompleteness
const auditModuleResult = auditPhotosCompleteness(samplePhotos, '');
assert.strictEqual(auditModuleResult.total, 6, '總張數應為 6');
assert.strictEqual(auditModuleResult.missingLocation, 1, '未填地點應為 1 (p2)');
assert.strictEqual(auditModuleResult.missingDesc, 1, '未填說明應為 1 (p3)');
assert.strictEqual(auditModuleResult.invalidDateTime, 2, '時間日期異常應為 2 (p4, p5)');
assert.strictEqual(auditModuleResult.duplicatePhotos, 2, '同名照片張數應為 2 (p1, p6)');
assert.strictEqual(auditModuleResult.issuesCount, 6, '問題總數應為 1+1+2+2 = 6');
assert.strictEqual(auditModuleResult.firstIssueFilter, 'missingDesc', '首要篩選應為未填說明');

// 測試 App 委派 auditPhotosCompleteness
app.photos = samplePhotos;
const auditAppResult = app.auditPhotosCompleteness();
assert.deepStrictEqual(auditAppResult, auditModuleResult, 'App 委派結果應與獨立模組計算結果完全一致');

// 驗證唯讀性保證
assert.strictEqual(
    JSON.stringify(samplePhotos),
    originalPhotosJson,
    'audit 計算前後，傳入之 photos 陣列絕對不得被修改（唯讀性）'
);

console.log('✅ audit.test.js 全部斷言通過！');
