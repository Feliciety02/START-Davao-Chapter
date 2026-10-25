const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const c = vm.createContext({ Date, console });
for (const file of fs.readdirSync(root).filter(name => name.endsWith('.gs')).sort())
  vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), c, { filename: file });

const old = new Date(); old.setFullYear(old.getFullYear() - 3);
const recent = new Date(); recent.setDate(recent.getDate() - 3);
const day = d => d.toISOString().slice(0, 10);
const tab = initial => {
  let rows = initial.map(r => r.slice());
  return {
    getLastRow: () => rows.length + 1,
    getRange(_at, _col, count, width) {
      return {
        getValues: () => rows.slice(0, count).map(r => r.slice(0, width)),
        clearContent: () => { rows = []; },
        setValues: values => { rows = values.map(r => r.slice()); }
      };
    },
    rows: () => rows
  };
};
const daily = tab([
  [day(old), 'former@example.com', 'Former'],
  [day(old), 'active@example.com', 'Active'],
  [day(recent), 'former@example.com', 'Recent']
]);
const lead = tab([
  [old, '', '', 'lead@example.com', '', 'former@example.com'],
  [old, '', '', 'lead@example.com', '', 'active@example.com'],
  [recent, '', '', 'former@example.com', '', 'active@example.com']
]);
const rewards = tab([
  [old, 'old-1', 'former@example.com', 'AdDU', 'Notebook', 70, 'Approved'],
  [old, 'old-2', 'active@example.com', 'AdDU', 'Notebook', 70, 'Approved'],
  [recent, 'new-1', 'former@example.com', 'AdDU', 'Notebook', 70, 'Pending']
]);
const sheets = { 'Daily games': daily, 'Lead actions': lead, 'Reward requests': rewards };
c.memberTab_ = () => null;
c.scholarByEmail_ = () => ({
  'active@example.com': { status: 'Active' }, 'lead@example.com': { status: 'Active' },
  'former@example.com': { status: 'Graduated' }
});
c.docLock_ = () => ({ tryLock: () => true, releaseLock() {} });
const removed = c.cleanOldAnswers_({ getSheetByName: name => sheets[name] || null });
assert.equal(removed, 3);
assert.equal(daily.rows().length, 2);
assert.equal(lead.rows().length, 2);
assert.equal(rewards.rows().length, 2);
assert.equal(daily.rows()[0][1], 'active@example.com');
assert.equal(lead.rows()[0][5], 'active@example.com');
console.log('Activity retention checks passed');
