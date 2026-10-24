const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const c = vm.createContext({ console });
for (const file of fs.readdirSync(root).filter(name => name.endsWith('.gs')).sort())
  vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), c, { filename: file });

let maxRows = 1, lastRow = 1, saved = null, released = false;
const headings = ['Timestamp', 'Email Address', c.CONSENT_TITLE, 'Week due', 'Submitted at'];
const sheet = {
  getLastRow: () => lastRow,
  getMaxRows: () => maxRows,
  insertRowsAfter: (at, count) => { assert.equal(at, 1); maxRows += count; },
  getRange: (row, col, height, width) => ({
    getValues: () => [headings.slice(col - 1, col - 1 + width)],
    setValues: rows => { assert.ok(row <= maxRows); saved = rows[0]; lastRow = row; }
  })
};
c.Logger = { log() {} };
c.webSafe_ = fn => fn();
c.webSession_ = () => ({ e: 'ana@example.com' });
c.webSheet_ = () => ({});
c.webRoles_ = () => ({ email: 'ana@example.com', member: { name: 'Ana', school: 'USeP' }, leads: [] });
c.webCheckinSpec_ = () => ({ open: true, late: false, week: '2026-10-25', q: [
  { k: 'vibe', title: 'How are you?', type: 'one', choices: ['Good'] }
] });
c.memberTab_ = () => sheet;
c.memberCols_ = () => ({ email: 1, name: -1, school: -1, role: -1, vibe: -1,
  events: -1, learning: -1, help: -1, tea: -1, width: headings.length });
c.ensureEmailCol_ = c.memberCols_;
c.plainHeading_ = x => x;
c.ensureCheckinMetadata_ = () => ({ week: 3, submitted: 4, width: headings.length });
c.webAnswerRow_ = () => -1;
c.docLock_ = () => ({ tryLock: () => true, releaseLock: () => { released = true; } });
c.SpreadsheetApp = { flush() {} };
c.onMemberCheckin_ = () => {};
c.markAnswersSeen_ = () => {};
c.webForget_ = () => { throw new Error('cache unavailable'); };
c.webMember_ = () => { throw new Error('page refresh unavailable'); };

const result = c.webSubmitCheckin('token', { vibe: 'Good', consent: true });
assert.equal(result.ok, true);
assert.equal(result.saved, true, 'a failed page refresh must not turn a saved answer into an error');
assert.equal(maxRows, 2, 'a full answer sheet grows before the next row is written');
assert.equal(saved[1], 'ana@example.com');
assert.equal(saved[3], '2026-10-25');
assert.equal(released, true);
c.webAnswerRow_ = () => 2;
const retry = c.webSubmitCheckin('token', { vibe: 'Good', consent: true });
assert.equal(retry.ok, true);
assert.equal(retry.already, true, 'a retry acknowledges the previously saved answer');
assert.equal(lastRow, 2, 'a retry never writes a second answer');
console.log('check-in submission tests passed');
