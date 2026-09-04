const { execSync } = require('child_process');
const path = require('path');

const tests = [
    'tests/unit/validation.test.js',
    'tests/unit/audit.test.js',
    'tests/unit/history.test.js'
];

console.log('========================================');
console.log('開始執行 Phase 0A 單元測試套件');
console.log('========================================\n');

let passedCount = 0;

for (const testFile of tests) {
    try {
        const fullPath = path.resolve(__dirname, '..', testFile);
        const output = execSync(`node "${fullPath}"`, { encoding: 'utf8' });
        process.stdout.write(output);
        passedCount++;
    } catch (error) {
        console.error(`❌ 測試失敗: ${testFile}`);
        console.error(error.stdout || error.message);
        process.exit(1);
    }
}

console.log('\n========================================');
console.log(`🎉 測試結果: 全部通過 (${passedCount}/${tests.length})`);
console.log('========================================');
