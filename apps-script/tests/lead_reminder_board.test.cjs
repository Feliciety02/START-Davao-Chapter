const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(path.join(root, '05_Portals.gs'), 'utf8');
const start = source.indexOf('function leadBoard_(');
const end = source.indexOf('// One Lead action.', start);
assert.ok(start > 0 && end > start);
const ticks = { 'LEAD_DONE|AdDU': ['remind'], 'LEAD_DONE|DNSC': [] };
const context = vm.createContext({
  Date, memberWeekKey_: () => 'this-week',
  leadRows_: () => [{ week: 'this-week', school: 'AdDU', action: 'contacted' },
    { week: 'last-week', school: 'DNSC', action: 'contacted' }],
  doneList_: key => ticks[key] || [], leadNote_: () => null
});
vm.runInContext(source.slice(start, end), context);
const schools = [
  { name: 'AdDU', lead: 'Ana', leadDone: true },
  { name: 'DNSC', lead: 'Ben', leadDone: false }
];
let board = context.leadBoard_(null, schools);
assert.equal(board.find(r => r.school === 'AdDU').reminderSent, true);
assert.equal(board.find(r => r.school === 'DNSC').reminderSent, false);
ticks['LEAD_DONE|AdDU'] = [];
ticks['LEAD_DONE|DNSC'] = ['last'];
board = context.leadBoard_(null, schools);
assert.equal(board.find(r => r.school === 'AdDU').reminderSent, false, 'undo clears the mark');
assert.equal(board.find(r => r.school === 'DNSC').reminderSent, true, 'last-day reminder counts');
assert.match(fs.readFileSync(path.join(root, 'Web_App.html'), 'utf8'), /<th>Reminder sent<\/th>/);
console.log('Campus Lead reminder board checks passed');
