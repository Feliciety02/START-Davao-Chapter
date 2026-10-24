const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const context = vm.createContext({ console });
for (const file of fs.readdirSync(root).filter(name => name.endsWith('.gs')).sort()) {
  vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), context, { filename: file });
}

const calls = [];
const scriptLock = {
  tryLock: () => { calls.push('try'); return true; },
  releaseLock: () => { calls.push('release'); }
};
const documentLock = { tryLock: () => true, releaseLock: () => {} };
context.LockService = {
  getDocumentLock: () => documentLock,
  getScriptLock: () => scriptLock
};
assert.equal(context.docLock_(), documentLock);

// Apps Script returns null for a document lock when called from a web app.
context.LockService.getDocumentLock = () => null;
assert.equal(context.docLock_(), scriptLock);

const delegatedLock = { tryLock: () => true, releaseLock: () => {} };
context.connectTracker({ lock: () => delegatedLock });
assert.equal(context.docLock_(), delegatedLock);
context.connectTracker({ lock: () => null });
assert.equal(context.docLock_(), scriptLock);
context.connectTracker(null);

// An Edit scholar request must acquire and release that fallback lock.
const old = new Array(12).fill('');
old[4] = 'pminoza@addu.edu.ph';
const scholars = {
  getLastRow: () => 2,
  getMaxColumns: () => 12,
  getRange: () => ({ getValues: () => [old] })
};
const ss = { getSheetByName: name => name === context.TAB.scholars ? scholars : null };
const result = context.webScholarUpdate_(ss, 'rr@example.com', {
  row: 2, email: old[4], name: 'Miñoza, Precious', contact: ''
});
assert.equal(result.ok, false);
assert.equal(result.msg, 'Enter a name and a contact email.');
assert.deepEqual(calls, ['try', 'release']);
console.log('Web app lock checks passed');
