const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', '05_Portals.gs'), 'utf8');
const start = source.indexOf('var REWARD_TAB =');
const end = source.indexOf('/* ---------- confirmed attendance', start);
assert.ok(start > 0 && end > start);

class Sheet {
  constructor(rows = []) { this.rows = rows; }
  getLastRow() { return this.rows.length; }
  getRange(row, col, height, width) {
    const sheet = this;
    return {
      getValues() { return sheet.rows.slice(row - 1, row - 1 + height).map(r => r.slice(col - 1, col - 1 + width)); },
      setValues(values) { values.forEach((v, i) => {
        while (sheet.rows.length < row + i) sheet.rows.push([]);
        v.forEach((x, j) => { sheet.rows[row + i - 1][col + j - 1] = x; });
      }); return this; },
      setFontWeight() { return this; }
    };
  }
  setFrozenRows() {}
  appendRow(row) { this.rows.push(row); }
}
const tabs = {
  Rewards: new Sheet([['Reward', 'XP cost', 'Details', 'Active (Y/N)'], ['Notebook', 70, 'START notebook', 'Y']]),
  'Reward requests': new Sheet([['Requested at', 'Request ID', 'Scholar email', 'School', 'Reward', 'XP cost', 'Status', 'Decided at', 'Decision note']])
};
const ss = { getSheetByName: name => tabs[name], insertSheet: name => (tabs[name] = new Sheet()) };
let earned = 100;
let sequence = 0;
const who = { member: { primary: 'scholar@example.com', school: 'AdDU' }, email: 'scholar@example.com' };
const context = vm.createContext({
  Date, Number, String, RegExp, Utilities: { getUuid: () => `id-${++sequence}` },
  webSafe_: fn => fn(), webSession_: () => ({ e: who.email }), webSheet_: () => ss,
  webRoles_: () => who, emailOf_: x => String(x || '').toLowerCase(),
  scholarByEmail_: () => ({ 'scholar@example.com': { status: 'Active' } }),
  docLock_: () => ({ tryLock: () => true, releaseLock() {} }), webBumpData_: () => {},
  webGame_: () => ({ xp: earned }), webMyAnswers_: () => [], webMyDaily_: () => ({}),
  webFmt_: () => 'Oct 25, 2026'
});
vm.runInContext(source.slice(start, end), context);

assert.equal(context.webRewardRequest('token', '2').ok, true);
assert.equal(context.rewardState_(ss, who.email).available, 30);
assert.equal(context.webRewardRequest('token', '2').ok, false, 'pending XP must be reserved');
let pending = context.rewardPending_(ss);
assert.equal(pending.length, 1);
assert.equal(context.rewardDecide_(ss, pending[0].id, 'Declined').ok, true);
assert.equal(context.rewardState_(ss, who.email).available, 100, 'decline releases XP');

assert.equal(context.webRewardRequest('token', '2').ok, true);
pending = context.rewardPending_(ss);
tabs.Rewards.rows[1][1] = 90;
assert.equal(context.rewardDecide_(ss, pending[0].id, 'Approved').ok, true);
assert.equal(context.rewardState_(ss, who.email).available, 30, 'approval spends original price');
assert.equal(context.rewardDecide_(ss, pending[0].id, 'Approved').ok, false, 'cannot approve twice');
earned = 50;
assert.equal(context.rewardState_(ss, who.email).available, 0);
assert.match(source, /if \(!who\.admin \|\| who\.suspended\)/, 'admin RPC guards decisions');
console.log('Reward reservation and approval checks passed');
