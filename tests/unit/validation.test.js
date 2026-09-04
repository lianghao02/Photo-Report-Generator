const assert = require('assert');
const { loadAppMethods } = require('./app-bridge');

const app = loadAppMethods();

console.log('--- 測試 validation.test.js ---');

// 1. 民國日期合法性
// 正常合法日期
assert.strictEqual(app.isValidMinguoDate('113/08/26'), true, '標準民國日期應合法');
assert.strictEqual(app.isValidMinguoDate('1130826'), true, '7碼純數字應合法');
assert.strictEqual(app.isValidMinguoDate('080/01/01'), true, '早期民國年應合法');

// 閏年與平年測試
// 民國 113 年 = 西元 2024 年（閏年）
assert.strictEqual(app.isValidMinguoDate('113/02/29'), true, '民國113年2月29日（閏年）應合法');
// 民國 114 年 = 西元 2025 年（平年）
assert.strictEqual(app.isValidMinguoDate('114/02/29'), false, '民國114年2月29日（平年）應判定非法');
assert.strictEqual(app.isValidMinguoDate('114/02/28'), true, '民國114年2月28日應合法');

// 大小月測試
assert.strictEqual(app.isValidMinguoDate('113/04/30'), true, '4月30日應合法');
assert.strictEqual(app.isValidMinguoDate('113/04/31'), false, '4月無31日應判定非法');
assert.strictEqual(app.isValidMinguoDate('113/07/31'), true, '7月31日應合法');

// 非法月份與日數
assert.strictEqual(app.isValidMinguoDate('113/00/01'), false, '月份為0應判定非法');
assert.strictEqual(app.isValidMinguoDate('113/13/01'), false, '月份為13應判定非法');
assert.strictEqual(app.isValidMinguoDate('113/05/00'), false, '日數為0應判定非法');
assert.strictEqual(app.isValidMinguoDate('113/05/32'), false, '日數超過31應判定非法');

// 位數與空值
assert.strictEqual(app.isValidMinguoDate(''), false, '空字串應判定非法');
assert.strictEqual(app.isValidMinguoDate(null), false, 'null 應判定非法');
assert.strictEqual(app.isValidMinguoDate('11308'), false, '不足7碼應判定非法');
assert.strictEqual(app.isValidMinguoDate('11308261'), false, '超過7碼應判定非法');

// 2. 時間格式與數值範圍
// 合法時間格式
assert.strictEqual(app.isValidTimeFormat('14:30'), true, '4碼標準時間應合法');
assert.strictEqual(app.isValidTimeFormat('14:30:15'), true, '6碼標準時間應合法');
assert.strictEqual(app.isValidTimeFormat('1430'), true, '4碼純數字應合法');
assert.strictEqual(app.isValidTimeFormat('143015'), true, '6碼純數字應合法');
assert.strictEqual(app.isValidTimeFormat('00:00'), true, '午夜00:00應合法');
assert.strictEqual(app.isValidTimeFormat('23:59:59'), true, '極限時間23:59:59應合法');

// 非法數值
assert.strictEqual(app.isValidTimeFormat('24:00'), false, '小時24應判定非法');
assert.strictEqual(app.isValidTimeFormat('12:60'), false, '分鐘60應判定非法');
assert.strictEqual(app.isValidTimeFormat('12:30:60'), false, '秒數60應判定非法');
assert.strictEqual(app.isValidTimeFormat('12'), false, '不足4碼應判定非法');
assert.strictEqual(app.isValidTimeFormat('12345'), false, '5碼數字應判定非法');
assert.strictEqual(app.isValidTimeFormat(''), false, '空字串應判定非法');
assert.strictEqual(app.isValidTimeFormat(null), false, 'null 應判定非法');

console.log('✅ validation.test.js 全部斷言通過！');
