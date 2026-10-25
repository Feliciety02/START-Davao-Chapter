const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const c = vm.createContext({ console });
for (const file of fs.readdirSync(root).filter(name => name.endsWith('.gs')).sort()) {
  vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), c, { filename: file });
}

const props = new Map();
c.docProps_ = () => ({ getProperty: k => props.get(k) || null, setProperty: (k, v) => props.set(k, v) });
c.webHash_ = x => x;
c.me_ = () => 'owner@example.com';
c.automationOwner_ = () => 'owner@example.com';
c.getSettings_ = () => ({ recipients: 'summary@example.com', rrName: 'RR' });
c.scholarByEmail_ = () => ({});
c.schoolRows_ = () => [];
assert.equal(c.webRoles_({}, 'summary@example.com').admin, false);
assert.equal(c.webRoles_({}, 'owner@example.com').admin, true);

// A stale edit cannot overwrite a school that someone else renamed.
const rows = Array.from({ length: c.CONFIG.maxSchools }, () => ['', '', '', '', '', '']);
rows[0][0] = 'New name';
const sh = { getRange: () => ({ getValues: () => rows }) };
c.docLock_ = () => ({ tryLock: () => true, releaseLock: () => {} });
const ss = { getSheetByName: name => name === c.TAB.schools ? sh : null };
const stale = c.webSchoolSave_(ss, 'owner@example.com', {
  row: 2, original: 'Old name', name: 'Another name', province: 'Davao del Sur', lead: '',
  info: c.SCHOOL_INFO_HEADERS.map(() => ''), inactive: false
});
assert.equal(stale.ok, false);
assert.match(stale.msg, /changed/i);

const writes = [], renamed = [];
rows[0][0] = 'Old name';
const editing = { getRange: (r, col, nr, nc) => ({
  getValues: () => rows,
  setValues: values => { writes.push([r, col, values]); return this; },
  setValue: value => { writes.push([r, col, value]); return this; }
}) };
const editSheet = { getSheetByName: name => name === c.TAB.schools ? editing : null };
c.scholarRows_ = () => [];
c.findScholar_ = () => null;
c.renameSchool_ = (sheet, from, to) => renamed.push([from, to]);
c.syncRoles_ = () => {};
c.updateDropdown_ = () => {};
c.calPeopleChanged_ = () => {};
c.webForget_ = () => {};
c.SpreadsheetApp = { flush: () => {} };
c.Logger = { log: () => {} };
const saved = c.webSchoolSave_(editSheet, 'owner@example.com', {
  row: 2, original: 'Old name', name: 'New name', province: 'Davao del Sur', lead: '',
  info: c.SCHOOL_INFO_HEADERS.map(() => ''), inactive: false
});
assert.equal(saved.ok, true);
assert.deepEqual(renamed, [['Old name', 'New name']]);
assert.equal(writes.some(w => w[0] === 2 && w[1] === 1 && w[2][0][0] === 'New name'), true);

// Derived reminder day is server enforced, even if an old client submits it.
c.webSettingRecords_ = () => c.SETTINGS.map(s => ({ key: s[0], readonly: s[0] === 'remDay' }));
const remDay = c.webSettingSave_({ getSheetByName: () => ({}) }, 'owner@example.com', { key: 'remDay', value: 'Monday' });
assert.equal(remDay.ok, false);
assert.match(remDay.msg, /deadline/i);

console.log('Website database admin checks passed');
