const assert = require('assert');
const { loadAppMethods } = require('./app-bridge');

const app = loadAppMethods();

console.log('--- 測試 audit.test.js ---');

// 1. 同名照片集合計算 _buildDupSet
app.photos = [
    { name: 'A.jpg' },
    { name: 'B.jpg' },
    { name: 'A.jpg' },
    { name: 'C.jpg' },
    { name: 'D.jpg' },
    { name: 'B.jpg' },
    { name: 'E.jpg' }
];

app._buildDupSet();
assert.strictEqual(app._dupSet instanceof Set, true, '_dupSet 應為 Set 物件');
assert.strictEqual(app._dupSet.has('A.jpg'), true, 'A.jpg 出現2次應在 _dupSet 內');
assert.strictEqual(app._dupSet.has('B.jpg'), true, 'B.jpg 出現2次應在 _dupSet 內');
assert.strictEqual(app._dupSet.has('C.jpg'), false, 'C.jpg 僅出現1次不應在 _dupSet 內');
assert.strictEqual(app._dupSet.has('D.jpg'), false, 'D.jpg 僅出現1次不應在 _dupSet 內');
assert.strictEqual(app._dupSet.has('E.jpg'), false, 'E.jpg 僅出現1次不應在 _dupSet 內');
assert.strictEqual(app._dupSet.size, 2, '應只有 2 組重複檔名');

// 2. 完整度稽核 auditPhotosCompleteness
// 建立含有各種缺失狀態的測試照片陣列
// 模擬全域預設地點為空白
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

// 保留原始深度拷貝以驗證唯讀性
const originalPhotosJson = JSON.stringify(samplePhotos);

app.photos = samplePhotos;
const auditResult = app.auditPhotosCompleteness();

assert.strictEqual(auditResult.total, 6, '總張數應為 6');
assert.strictEqual(auditResult.missingLocation, 1, '未填地點應為 1 (p2)');
assert.strictEqual(auditResult.missingDesc, 1, '未填說明應為 1 (p3)');
assert.strictEqual(auditResult.invalidDateTime, 2, '時間日期異常應為 2 (p4, p5)');
assert.strictEqual(auditResult.duplicatePhotos, 2, '同名照片張數應為 2 (p1, p6)');
assert.strictEqual(auditResult.issuesCount, 6, '問題總數應為 1+1+2+2 = 6');
assert.strictEqual(auditResult.firstIssueFilter, 'missingDesc', '優先切換之首個問題篩選應為 missingDesc');

// 驗證當全域有設定清冊地點時，照片未填個別地點不算 missingLocation
global.document = {
    getElementById: (id) => {
        if (id === 'defaultLocation') return { value: '新化分局' };
        return null;
    }
};
const auditWithGlobalLoc = app.auditPhotosCompleteness();
assert.strictEqual(auditWithGlobalLoc.missingLocation, 0, '有全域預設地點時，未填個別地點應視為已繼承');

// 3. 唯讀性驗證：確認 audit 執行後 photos 內容未被更動
assert.strictEqual(JSON.stringify(samplePhotos), originalPhotosJson, '稽核演算法必須為純讀取，不得修改 photos 陣列或物件屬性');

console.log('✅ audit.test.js 全部斷言通過！');
