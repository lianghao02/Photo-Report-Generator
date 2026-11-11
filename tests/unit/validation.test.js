const assert = require('assert');
const {
    MAX_LOCATION_FULLWIDTH,
    MAX_DESC_FULLWIDTH,
    DESC_LINE_FULLWIDTH,
    MAX_DESC_LINES,
    isValidMinguoDate,
    isValidTimeFormat,
    isHalfWidthChar,
    getTextHalfWidthUnits,
    getTextFullWidthLength,
    formatFullWidthCount,
    estimateDescVisualLines,
    truncateToFullWidth,
    inspectTextLimits
} = require('../../js/validation');
const { loadAppMethods } = require('./app-bridge');

const app = loadAppMethods();

console.log('--- 測試 validation.test.js (包含獨立模組與 App 委派) ---');

assert.strictEqual(MAX_LOCATION_FULLWIDTH, 17, '地點上限應為 17 全形');
assert.strictEqual(MAX_DESC_FULLWIDTH, 68, '說明上限應為 68 全形');
assert.strictEqual(DESC_LINE_FULLWIDTH, 34, '說明單行容量應為 34 全形');
assert.strictEqual(MAX_DESC_LINES, 2, '說明最大行數應為 2 行');

[
    {
        name: 'PhotoReportValidation (獨立模組)',
        validator: {
            isValidMinguoDate,
            isValidTimeFormat,
            getTextHalfWidthUnits,
            getTextFullWidthLength,
            formatFullWidthCount,
            estimateDescVisualLines,
            truncateToFullWidth,
            inspectTextLimits
        }
    },
    {
        name: 'PhotoReportApp (宿主委派方法)',
        validator: {
            isValidMinguoDate: (...args) => app.isValidMinguoDate(...args),
            isValidTimeFormat: (...args) => app.isValidTimeFormat(...args),
            getTextHalfWidthUnits: (...args) => app.getTextHalfWidthUnits(...args),
            getTextFullWidthLength: (...args) => app.getTextFullWidthLength(...args),
            formatFullWidthCount: (...args) => app.formatFullWidthCount(...args),
            estimateDescVisualLines: (...args) => app.estimateDescVisualLines(...args),
            truncateToFullWidth: (...args) => app.truncateToFullWidth(...args),
            inspectTextLimits: (...args) => app.inspectPhotoTextLimits(...args)
        }
    }
].forEach(({ name, validator }) => {
    console.log(`正在檢驗: ${name}`);

    // 1. 民國日期合法性
    // 正常合法日期
    assert.strictEqual(validator.isValidMinguoDate('113/08/26'), true, '標準民國日期應合法');
    assert.strictEqual(validator.isValidMinguoDate('1130826'), true, '7碼純數字應合法');
    assert.strictEqual(validator.isValidMinguoDate('080/01/01'), true, '早期民國年應合法');

    // 閏年與平年測試
    // 民國 113 年 = 西元 2024 年（閏年）
    assert.strictEqual(validator.isValidMinguoDate('113/02/29'), true, '民國113年2月29日（閏年）應合法');
    // 民國 114 年 = 西元 2025 年（平年）
    assert.strictEqual(validator.isValidMinguoDate('114/02/29'), false, '民國114年2月29日（平年）應判定非法');
    assert.strictEqual(validator.isValidMinguoDate('114/02/28'), true, '民國114年2月28日應合法');

    // 大小月測試
    assert.strictEqual(validator.isValidMinguoDate('113/04/30'), true, '4月30日應合法');
    assert.strictEqual(validator.isValidMinguoDate('113/04/31'), false, '4月無31日應判定非法');
    assert.strictEqual(validator.isValidMinguoDate('113/07/31'), true, '7月31日應合法');

    // 非法月份與日數
    assert.strictEqual(validator.isValidMinguoDate('113/00/01'), false, '月份為0應判定非法');
    assert.strictEqual(validator.isValidMinguoDate('113/13/01'), false, '月份為13應判定非法');
    assert.strictEqual(validator.isValidMinguoDate('113/05/00'), false, '日數為0應判定非法');
    assert.strictEqual(validator.isValidMinguoDate('113/05/32'), false, '日數超過31應判定非法');

    // 位數與空值
    assert.strictEqual(validator.isValidMinguoDate(''), false, '空字串應判定非法');
    assert.strictEqual(validator.isValidMinguoDate(null), false, 'null 應判定非法');
    assert.strictEqual(validator.isValidMinguoDate('11308'), false, '不足7碼應判定非法');
    assert.strictEqual(validator.isValidMinguoDate('11308261'), false, '超過7碼應判定非法');

    // 2. 時間格式與數值範圍
    // 合法時間格式
    assert.strictEqual(validator.isValidTimeFormat('14:30'), true, '4碼標準時間應合法');
    assert.strictEqual(validator.isValidTimeFormat('14:30:15'), true, '6碼標準時間應合法');
    assert.strictEqual(validator.isValidTimeFormat('1430'), true, '4碼純數字應合法');
    assert.strictEqual(validator.isValidTimeFormat('143015'), true, '6碼純數字應合法');
    assert.strictEqual(validator.isValidTimeFormat('00:00'), true, '午夜00:00應合法');
    assert.strictEqual(validator.isValidTimeFormat('23:59:59'), true, '極限時間23:59:59應合法');

    // 非法數值
    assert.strictEqual(validator.isValidTimeFormat('24:00'), false, '小時24應判定非法');
    assert.strictEqual(validator.isValidTimeFormat('12:60'), false, '分鐘60應判定非法');
    assert.strictEqual(validator.isValidTimeFormat('12:30:60'), false, '秒數60應判定非法');
    assert.strictEqual(validator.isValidTimeFormat('12'), false, '不足4碼應判定非法');
    assert.strictEqual(validator.isValidTimeFormat('12345'), false, '5碼數字應判定非法');
    assert.strictEqual(validator.isValidTimeFormat(''), false, '空字串應判定非法');
    assert.strictEqual(validator.isValidTimeFormat(null), false, 'null 應判定非法');

    // 3. 全形／半形佔位與排版上限（2 半形＝1 全形，地點上限 17 全形，說明上限 68 全形・最多 2 行）
    assert.strictEqual(validator.getTextHalfWidthUnits('()'), 2, '兩個半形括號 () 應佔 2 半形單位');
    assert.strictEqual(validator.getTextFullWidthLength('()'), 1, '兩個半形括號 () 應計為 1 個全形佔位');
    assert.strictEqual(validator.getTextFullWidthLength('議員支持 4 號(李偉智)'), 10.5, '中英數空白與半形括號混合計算應為 10.5 全形');
    assert.strictEqual(validator.formatFullWidthCount(10.5), '10.5', '小數 0.5 全形應顯示為 10.5');
    assert.strictEqual(validator.formatFullWidthCount(7), '7', '整數全形不應有多餘小數點');
    assert.strictEqual(validator.getTextFullWidthLength('關廟尚好食餐廳'), 7, '純中文應以每字 1 全形計算');
    assert.strictEqual(validator.getTextFullWidthLength('（全形括號）'), 6, '全形標點與全形括號應計為 1 全形');
    assert.strictEqual(validator.getTextFullWidthLength('ｶﾅ'), 1, '半形片假名 2 字應算 1 全形');
    assert.strictEqual(validator.getTextFullWidthLength('🚓📍'), 2, 'Emoji 應以 Unicode 碼位計為 1 全形而非拆成 2 個 UTF-16 code unit');
    assert.strictEqual(validator.getTextFullWidthLength('第一行\n第二行'), 6, '換行符號不應額外計算寬度單位');
    assert.strictEqual(validator.getTextFullWidthLength(null), 0, 'null 應安全回傳 0');

    // 使用者回饋實際超長案例檢驗
    const userOverflowExample = '議員李偉智、里長參選人賴延銓，由長壽會幹部陪同逐桌敬酒，過程中請托拜票，議員支持 4 號(李偉智)、里長支持 3 號(賴延銓)，總幹事洪金河邊講，邊比手勢';
    const exampleCheck = validator.inspectTextLimits('關廟尚好食餐廳', userOverflowExample);
    assert.strictEqual(exampleCheck.descOver, true, '使用者回饋會變成 3 行的說明實際測得應判定超標');
    assert.ok(exampleCheck.descLines >= 3 || exampleCheck.descFullWidth > 68, '實際超標說明應觸發 3 行或逾 68 全形告警');

    // 4. 視覺行數估算
    assert.strictEqual(validator.estimateDescVisualLines(''), 0, '空白說明應為 0 行');
    assert.strictEqual(validator.estimateDescVisualLines('一'.repeat(34)), 1, '34 全形應為 1 行');
    assert.strictEqual(validator.estimateDescVisualLines('一'.repeat(34) + 'a'), 2, '34.5 全形應折為 2 行');
    assert.strictEqual(validator.estimateDescVisualLines('()'.repeat(68)), 2, '136 半形（68 全形）應剛好為 2 行');
    assert.strictEqual(validator.estimateDescVisualLines('()'.repeat(68) + 'a'), 3, '68.5 全形應折為 3 行');
    assert.strictEqual(validator.estimateDescVisualLines('第一行\n第二行'), 2, '單次換行應為 2 行');
    assert.strictEqual(validator.estimateDescVisualLines('第1行\n第2行\n第3行'), 3, '兩次換行應計為 3 行');
    assert.strictEqual(validator.estimateDescVisualLines('a'.repeat(69) + '\n第2段'), 3, '第 1 段折為 2 行後再換行應計為 3 行');

    // 5. 全形上限與行數安全截斷
    const locTruncChinese = validator.truncateToFullWidth('台'.repeat(20), 17, { singleLine: true });
    assert.strictEqual(locTruncChinese.text, '台'.repeat(17), '地點純中文超過 17 字應截斷為 17 字');
    assert.strictEqual(locTruncChinese.fullWidthLength, 17);
    assert.strictEqual(locTruncChinese.truncated, true);

    const locTruncHalf = validator.truncateToFullWidth('A'.repeat(40), 17, { singleLine: true });
    assert.strictEqual(locTruncHalf.text, 'A'.repeat(34), '地點純半形應最多容納 34 字（17 全形）');
    assert.strictEqual(locTruncHalf.fullWidthLength, 17);

    const locTruncBoundary = validator.truncateToFullWidth('中'.repeat(16) + 'a中', 17, { singleLine: true });
    assert.strictEqual(locTruncBoundary.text, '中'.repeat(16) + 'a', '剩 0.5 全形時不得塞入 1 個全形字');
    assert.strictEqual(locTruncBoundary.fullWidthLength, 16.5);

    const locTruncMultiline = validator.truncateToFullWidth('第一\r\n第二', 17, { singleLine: true });
    assert.strictEqual(locTruncMultiline.text, '第一第二', '單行地點貼入換行應自動移除換行符號');
    assert.strictEqual(locTruncMultiline.truncated, true);

    const emojiBoundary = validator.truncateToFullWidth('甲'.repeat(16) + '🚓乙', 17, { singleLine: true });
    assert.strictEqual(emojiBoundary.text, '甲'.repeat(16) + '🚓', 'Emoji 落在邊界時應完整保留或整顆略過，不得拆出半個 surrogate');
    assert.strictEqual(/[\uD800-\uDFFF]/.test(emojiBoundary.text.replace(/\p{Emoji_Presentation}|\p{Extended_Pictographic}/gu, '')), false);

    const descTrunc68 = validator.truncateToFullWidth('()'.repeat(70), 68, { maxLines: 2, lineFullWidth: 34 });
    assert.strictEqual(descTrunc68.text, '()'.repeat(68), '說明貼入 70 個半形括號對（70 全形）應截斷為 68 全形（136 半形）');
    assert.strictEqual(descTrunc68.visualLines, 2);
    assert.strictEqual(descTrunc68.truncated, true);

    const descTruncExtraNewline = validator.truncateToFullWidth('第一行\n第二行\n第三行', 68, { maxLines: 2, lineFullWidth: 34 });
    assert.strictEqual(descTruncExtraNewline.text, '第一行\n第二行第三行', '貼入多餘換行超過 2 行時，應略過第 2 個換行符號而保留後方文字');
    assert.strictEqual(descTruncExtraNewline.visualLines, 2);
    assert.strictEqual(descTruncExtraNewline.truncated, true);

    const descTruncLine2Cap = validator.truncateToFullWidth('短句\n' + '字'.repeat(40), 68, { maxLines: 2, lineFullWidth: 34 });
    assert.strictEqual(descTruncLine2Cap.text, '短句\n' + '字'.repeat(34), '第 1 行結束換行後，第 2 行應限制在 34 全形內避免變第 3 行');
    assert.strictEqual(descTruncLine2Cap.visualLines, 2);
    assert.strictEqual(descTruncLine2Cap.truncated, true);

    // 6. inspectTextLimits 綜合判斷
    assert.deepStrictEqual(
        validator.inspectTextLimits('甲'.repeat(17), '字'.repeat(68)),
        {
            locationFullWidth: 17,
            descFullWidth: 68,
            descLines: 2,
            locationOver: false,
            descOver: false,
            isOver: false
        },
        '剛好等於 17 全形地點與 68 全形說明不應判定超標'
    );
    assert.strictEqual(validator.inspectTextLimits('甲'.repeat(17) + '1', '').isOver, true, '地點 17.5 全形應判定超標');
    assert.strictEqual(validator.inspectTextLimits('現場', '一\n二\n三').descOver, true, '說明 3 行即使未滿 68 全形也應判定超標');
});

assert.strictEqual(isHalfWidthChar('('), true, '半形左括號應為半形字元');
assert.strictEqual(isHalfWidthChar(')'), true, '半形右括號應為半形字元');
assert.strictEqual(isHalfWidthChar('（'), false, '全形左括號應為全形字元');
assert.strictEqual(isHalfWidthChar('中'), false, '中文字應為全形字元');

console.log('✅ validation.test.js 全部斷言通過！');
