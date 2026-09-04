const assert = require('assert');
const { historySignature, projectSignature, HistoryManager } = require('../../js/history');
const { loadAppMethods } = require('./app-bridge');

const app = loadAppMethods();

console.log('--- 測試 history.test.js (包含獨立模組與 App 委派) ---');

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

const sig1Module = historySignature(baseState);
const sig2Module = historySignature(stateWithUiChange);
assert.strictEqual(
    sig1Module,
    sig2Module,
    '獨立模組：僅更動 UI 暫存屬性時，historySignature 必須維持完全一致'
);

const sig1App = app.historySignature(baseState);
const sig2App = app.historySignature(stateWithUiChange);
assert.strictEqual(
    sig1App,
    sig2App,
    'App 委派：僅更動 UI 暫存屬性時，historySignature 必須維持完全一致'
);
assert.strictEqual(sig1App, sig1Module, 'App 委派簽名應與模組導出函式計算結果完全一致');

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
    const modifiedSig = historySignature(modifiedState);
    assert.notStrictEqual(
        sig1Module,
        modifiedSig,
        `當核心資料欄位 [${field}] 改變時，historySignature 必須改變`
    );
});

// 4. 測試：caseData 案件資訊改變時，簽名必須改變
const stateWithCaseDataChange = JSON.parse(JSON.stringify(baseState));
stateWithCaseDataChange.caseData.caseNumber = '113-警-002';
assert.notStrictEqual(
    sig1Module,
    historySignature(stateWithCaseDataChange),
    '當 caseData 改變時，historySignature 必須改變'
);

// 5. 測試：currentIndex 與 lastSelectedIndex 改變時，簽名必須改變
const stateWithIndexChange = JSON.parse(JSON.stringify(baseState));
stateWithIndexChange.currentIndex = 1;
assert.notStrictEqual(
    sig1Module,
    historySignature(stateWithIndexChange),
    '當 currentIndex 改變時，historySignature 必須改變'
);

const stateWithLastSelectedChange = JSON.parse(JSON.stringify(baseState));
stateWithLastSelectedChange.lastSelectedIndex = 1;
assert.notStrictEqual(
    sig1Module,
    historySignature(stateWithLastSelectedChange),
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
    sig1Module,
    historySignature(stateWithMorePhotos),
    '當照片清單長度改變時，historySignature 必須改變'
);

// 7. 測試 HistoryManager 類別狀態推移與 Undo/Redo
const manager = new HistoryManager(3);
assert.strictEqual(manager.canUndo(), false, '初始無步驟時不可 Undo');
assert.strictEqual(manager.canRedo(), false, '初始無步驟時不可 Redo');

manager.record(baseState);
assert.strictEqual(manager.canUndo(), false, '只有 1 個初始步驟時不可 Undo');
assert.strictEqual(manager.canRedo(), false, '只有 1 個步驟時不可 Redo');

// 重複相同狀態不應推入
const duplicateRecorded = manager.record(stateWithUiChange);
assert.strictEqual(duplicateRecorded, false, '相同歷史簽名狀態不應重複推入堆疊');

// 推入第 2 個不同狀態
manager.record(stateWithCaseDataChange);
assert.strictEqual(manager.canUndo(), true, '有 2 個不同步驟時應可 Undo');
assert.strictEqual(manager.canRedo(), false, '處於最新步驟時不可 Redo');

// 執行 Undo
const undoneState = manager.undo();
assert.strictEqual(historySignature(undoneState), sig1Module, 'Undo 應回到初始狀態');
assert.strictEqual(manager.canUndo(), false, '回到初始步驟時不可再 Undo');
assert.strictEqual(manager.canRedo(), true, 'Undo 後應可 Redo');

// 執行 Redo
const redoneState = manager.redo();
assert.strictEqual(historySignature(redoneState), historySignature(stateWithCaseDataChange), 'Redo 應回到最新狀態');
assert.strictEqual(manager.canRedo(), false, 'Redo 後無後續步驟，不可再 Redo');

console.log('✅ history.test.js 全部斷言通過！');
