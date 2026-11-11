const fs = require('fs');
const path = require('path');

let cachedMethods = null;

function loadAppMethods() {
    if (cachedMethods) return cachedMethods;
    // 主程式已由 index.html 拆到 js/app.js；兩個檔案都讀，方法在哪一個都找得到
    const sources = ['../../index.html', '../../js/app.js']
        .map((p) => path.resolve(__dirname, p))
        .filter((p) => fs.existsSync(p));
    const html = sources.map((p) => fs.readFileSync(p, 'utf8')).join('\n');

    function extractMethod(name) {
        const regex = new RegExp('(?:^|\\n)[ \\t]*(?:async\\s+)?' + name + '\\s*\\([^)]*\\)\\s*\\{');
        const match = regex.exec(html);
        if (!match) throw new Error(`Method ${name} not found in index.html or js/app.js`);
        const start = match[0].startsWith('\n') ? match.index + 1 : match.index;
        const bodyBraceIndex = match.index + match[0].length - 1;
        let braceCount = 0;
        let inBraces = false;
        for (let i = bodyBraceIndex; i < html.length; i++) {
            if (html[i] === '{') {
                braceCount++;
                inBraces = true;
            } else if (html[i] === '}') {
                braceCount--;
                if (inBraces && braceCount === 0) {
                    return html.slice(start, i + 1);
                }
            }
        }
        throw new Error(`Unmatched braces for method ${name}`);
    }

    const classCode = `
    class MockApp {
        ${extractMethod('isValidMinguoDate')}
        ${extractMethod('isValidTimeFormat')}
        ${extractMethod('getTextHalfWidthUnits')}
        ${extractMethod('getTextFullWidthLength')}
        ${extractMethod('formatFullWidthCount')}
        ${extractMethod('estimateDescVisualLines')}
        ${extractMethod('truncateToFullWidth')}
        ${extractMethod('inspectPhotoTextLimits')}
        ${extractMethod('_buildDupSet')}
        ${extractMethod('auditPhotosCompleteness')}
        ${extractMethod('historySignature')}
    }
    return new MockApp();
    `;

    const appInstance = new Function(classCode)();
    cachedMethods = appInstance;
    return cachedMethods;
}

module.exports = { loadAppMethods };
