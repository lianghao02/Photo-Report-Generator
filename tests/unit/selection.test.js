const assert = require('assert');
const { getVisiblePhotoIndices, getSelectedIndicesOrCurrent, computeKeyboardNavIndex } = require('../../js/selection');

console.log('--- 測試 selection.test.js ---');

const testPhotos = [
    { name: 'A.jpg', location: '客廳', desc: '全景', date: '113/08/26', time: '14:00', selected: false },
    { name: 'B.jpg', location: '', desc: '特寫', date: '113/08/26', time: '14:05', selected: true },
    { name: 'A.jpg', location: '臥室', desc: '', date: '113/08/26', time: '14:10', selected: false },
    { name: 'C.jpg', location: '陽台', desc: '欄杆', date: '114/02/29', time: '14:15', selected: false } // 日期異常
];

// 1. 測試 getVisiblePhotoIndices 各篩選模式
// all 模式
const visibleAll = getVisiblePhotoIndices(testPhotos, 'all');
assert.strictEqual(visibleAll.length, 4, 'all 模式應包含 4 張照片');

// missingLocation 模式（預設地點空白）
const visibleMissingLoc = getVisiblePhotoIndices(testPhotos, 'missingLocation', '');
assert.strictEqual(visibleMissingLoc.length, 1, 'missingLocation 應為 1 張 (B.jpg)');
assert.strictEqual(visibleMissingLoc[0].index, 1);

// missingLocation 模式（預設地點有值）
const visibleMissingLocWithDefault = getVisiblePhotoIndices(testPhotos, 'missingLocation', '現場');
assert.strictEqual(visibleMissingLocWithDefault.length, 0, '有預設地點時不應視為 missingLocation');

// missingDesc 模式
const visibleMissingDesc = getVisiblePhotoIndices(testPhotos, 'missingDesc');
assert.strictEqual(visibleMissingDesc.length, 1, 'missingDesc 應為 1 張');
assert.strictEqual(visibleMissingDesc[0].index, 2);

// invalidDateTime 模式
const visibleInvalidDt = getVisiblePhotoIndices(testPhotos, 'invalidDateTime');
assert.strictEqual(visibleInvalidDt.length, 1, 'invalidDateTime 應為 1 張 (C.jpg 平年 2/29)');
assert.strictEqual(visibleInvalidDt[0].index, 3);

// duplicatePhotos 模式
const visibleDup = getVisiblePhotoIndices(testPhotos, 'duplicatePhotos');
assert.strictEqual(visibleDup.length, 2, 'duplicatePhotos 應為 2 張 (A.jpg 出現兩次)');
assert.deepStrictEqual(visibleDup.map(v => v.index), [0, 2]);

// 2. 測試 getSelectedIndicesOrCurrent
// 當有選取時回傳已選索引
const selected = getSelectedIndicesOrCurrent(testPhotos, 0);
assert.deepStrictEqual(selected, [1], '應回傳已勾選照片索引 [1]');

// 當無選取時回傳目前焦點索引
const unselectedPhotos = testPhotos.map(p => ({ ...p, selected: false }));
const currentFallback = getSelectedIndicesOrCurrent(unselectedPhotos, 2);
assert.deepStrictEqual(currentFallback, [2], '無勾選時應回傳目前焦點照片 [2]');

// 3. 測試 computeKeyboardNavIndex
// all 模式導航
assert.strictEqual(computeKeyboardNavIndex({ totalPhotos: 4, currentIndex: 1, delta: 1, activeFilter: 'all' }), 2, '全部模式下 +1 應移至 2');
assert.strictEqual(computeKeyboardNavIndex({ totalPhotos: 4, currentIndex: 3, delta: 1, activeFilter: 'all' }), 3, '全部模式下已在尾端應停在 3');
assert.strictEqual(computeKeyboardNavIndex({ totalPhotos: 4, currentIndex: 0, delta: -1, activeFilter: 'all' }), 0, '全部模式下已在首端應停在 0');

// 篩選模式導航（可見照片為索引 0 與 2）
const filteredEntries = [{ index: 0 }, { index: 2 }];
assert.strictEqual(
    computeKeyboardNavIndex({ totalPhotos: 4, currentIndex: 0, delta: 1, activeFilter: 'duplicatePhotos', visibleEntries: filteredEntries }),
    2,
    '篩選模式下焦點在 0，+1 應跳至下一張可見照片索引 2'
);
assert.strictEqual(
    computeKeyboardNavIndex({ totalPhotos: 4, currentIndex: 2, delta: -1, activeFilter: 'duplicatePhotos', visibleEntries: filteredEntries }),
    0,
    '篩選模式下焦點在 2，-1 應跳回上一張可見照片索引 0'
);

console.log('✅ selection.test.js 全部斷言通過！');
