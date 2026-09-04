const assert = require('assert');
const { loadAppMethods } = require('./app-bridge');

const app = loadAppMethods();

console.log('--- 測試 history.test.js ---');

// 1. 建立基礎狀態
const baseState = {
    label: '編輯照片',
    currentIndex: 0,
    lastSelectedIndex: 0,
    caseData: {
        caseNumber: '113-警-001',
        caseName: '測試案件',
        investigator: '調查員A'
    },
    photos: [
        {
            uid: 'photo-1',
            rotation: 0,
            seq: 1,
            date: '113/08/26',
            time: '14:30',
            location: '客廳',
            desc: '全景',
            selected: true,
            stageX: 100,
            stageY: 200,
            // UI 狀態屬性（不應納入 historySignature）
            previewUrl: 'blob:http://localhost/xxx',
            isDragging: false,
            tempFilter: 'all',
            uiHovered: true
        }
    ]
};

// 2. 測試：UI 屬性變更時，簽名必須完全一致（不應產生多餘歷史步驟）
const stateWithUiChange = JSON.parse(JSON.stringify(baseState));
stateWithUiChange.photos[0].previewUrl = 'blob:http://localhost/yyy';
stateWithUiChange.photos[0].isDragging = true;
stateWithUiChange.photos[0].tempFilter = 'missing-location';
stateWithUiChange.photos[0].uiHovered = false;

const sig1 = app.historySignature(baseState);
const sig2 = app.historySignature(stateWithUiChange);

assert.strictEqual(
    sig1,
    sig2,
    '僅更動 UI 暫存屬性 (previewUrl, isDragging, tempFilter, uiHovered) 時，historySignature 必須維持完全一致'
);

// 3. 測試：純資料屬性變更時，簽名必須改變
const dataFieldsToTest = [
    { field: 'rotation', value: 90 },
    { field: 'seq', value: 2 },
    { field: 'date', value: '113/08/27' },
    { field: 'time', value: '15:00' },
    { field: 'location', value: '臥室' },
    { field: 'desc', value: '特寫' },
    { field: 'selected', value: false },
    { field: 'stageX', value: 150 },
    { field: 'stageY', value: 250 }
];

dataFieldsToTest.forEach(({ field, value }) => {
    const modifiedState = JSON.parse(JSON.stringify(baseState));
    modifiedState.photos[0][field] = value;
    const modifiedSig = app.historySignature(modifiedState);
    assert.notStrictEqual(
        sig1,
        modifiedSig,
        `當核心資料欄位 [${field}] 改變時，historySignature 必須改變`
    );
});

// 4. 測試：caseData 案件資訊改變時，簽名必須改變
const stateWithCaseDataChange = JSON.parse(JSON.stringify(baseState));
stateWithCaseDataChange.caseData.caseNumber = '113-警-002';
assert.notStrictEqual(
    sig1,
    app.historySignature(stateWithCaseDataChange),
    '當 caseData 改變時，historySignature 必須改變'
);

// 5. 測試：currentIndex 與 lastSelectedIndex 改變時，簽名必須改變
const stateWithIndexChange = JSON.parse(JSON.stringify(baseState));
stateWithIndexChange.currentIndex = 1;
assert.notStrictEqual(
    sig1,
    app.historySignature(stateWithIndexChange),
    '當 currentIndex 改變時，historySignature 必須改變'
);

const stateWithLastSelectedChange = JSON.parse(JSON.stringify(baseState));
stateWithLastSelectedChange.lastSelectedIndex = 1;
assert.notStrictEqual(
    sig1,
    app.historySignature(stateWithLastSelectedChange),
    '當 lastSelectedIndex 改變時，historySignature 必須改變'
);

// 6. 測試：新增或刪除照片時，簽名必須改變
const stateWithMorePhotos = JSON.parse(JSON.stringify(baseState));
stateWithMorePhotos.photos.push({
    uid: 'photo-2',
    rotation: 0,
    seq: 2,
    date: '113/08/26',
    time: '14:35',
    location: '門口',
    desc: '門牌',
    selected: false,
    stageX: 0,
    stageY: 0
});
assert.notStrictEqual(
    sig1,
    app.historySignature(stateWithMorePhotos),
    '當照片清單長度改變時，historySignature 必須改變'
);

console.log('✅ history.test.js 全部斷言通過！');
