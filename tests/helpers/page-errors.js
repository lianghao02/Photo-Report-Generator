const assert = require('node:assert/strict');
function trackPageErrors(page) {
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    return () => assert.deepEqual(errors, [], '頁面出現未處理例外');
}
module.exports = { trackPageErrors };
