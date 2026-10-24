const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const c = vm.createContext({ Date });
for (const file of fs.readdirSync(root).filter(name => name.endsWith('.gs')).sort())
  vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), c, { filename: file });

const sequence = [];
const sheets = new Map(['Settings', 'Schools', 'Old masterlist', 'Data change log',
  'Weekly roster', 'Scholars', 'Check-in answers', 'Custom old tab'].map(name => [name, { getName: () => name }]));
const schoolRange = { clearContent: () => sequence.push('clear campus leads') };
sheets.set('Schools', { getName: () => 'Schools', getRange: () => schoolRange });
const ss = {
  getId: () => 'tracker-id', getName: () => 'START',
  getSheets: () => [...sheets.values()], getSheetByName: name => sheets.get(name),
  deleteSheet: sheet => { sequence.push('delete ' + sheet.getName()); sheets.delete(sheet.getName()); }
};
const props = new Map([['WEB_A_ana', 'old account'], ['WEB_SUSP_ana', 'suspended'], ['SETUP_DONE', 'v61'],
  ['SETUP_PROGRESS', 'old run'], ['TABLE_LOOK', 'old look'], ['region', 'XI']]);
c.docProps_ = () => ({
  getProperties: () => Object.fromEntries(props), getProperty: key => props.get(key),
  deleteProperty: key => props.delete(key), setProperty: (key, value) => props.set(key, value)
});
c.docLock_ = () => ({ tryLock: () => { sequence.push('lock'); return true; }, releaseLock: () => sequence.push('unlock') });
c.DriveApp = { getFileById: () => ({
  makeCopy: () => { sequence.push('backup'); return { getUrl: () => 'https://backup.test', moveTo() {} }; },
  getParents: () => ({ hasNext: () => false })
}) };
c.Utilities = { formatDate: () => '2026-10-24 1200' };
c.SpreadsheetApp = { flush: () => sequence.push('flush') };
c.removeTriggers_ = () => {};
c.CacheService = { getScriptCache: () => ({ removeAll() {} }) };
c.setAutoApprove_ = on => { assert.equal(on, false); props.set('AUTO_APPROVE', '0'); };
c.getOrCreate_ = (_ss, name) => { sheets.set(name, { getName: () => name }); };
c.joinTab_ = () => { sheets.set('Join requests', { getName: () => 'Join requests' }); };
c.refreshJoinViews_ = () => {};
c.setupFull = () => { sequence.push('rebuild'); };

const result = c.freshStart_(ss);
assert.equal(result.backup, 'https://backup.test');
assert.equal(result.dropped, 6);
assert.deepEqual([...sheets.keys()].sort(), ['Events', 'Join requests', 'Schools', 'Settings'].sort());
assert.ok(sequence.indexOf('backup') < sequence.findIndex(x => x.startsWith('delete ')), 'backup precedes deletion');
assert.ok(sequence.indexOf('unlock') < sequence.indexOf('rebuild'), 'rebuild runs outside the write lock');
assert.ok(sequence.includes('clear campus leads'));
assert.equal(props.get('region'), 'XI');
assert.equal(props.has('WEB_A_ana'), false);
assert.equal(props.has('WEB_SUSP_ana'), false);
assert.equal(props.has('SETUP_DONE'), false);
assert.equal(props.has('SETUP_PROGRESS'), false);
assert.equal(props.has('FRESH_LIST_DONE'), true);
assert.equal(props.get('AUTO_APPROVE'), '0');

sheets.set('Custom after reset', { getName: () => 'Custom after reset' });
c.DriveApp.getFileById = () => ({ makeCopy: () => { throw new Error('backup failed'); } });
assert.throws(() => c.freshStart_(ss), /backup failed/);
assert.ok(sheets.has('Custom after reset'), 'backup failure leaves every tab intact');

const moved = [];
ss.setActiveSheet = sheet => { ss.active = sheet; };
ss.moveActiveSheet = position => moved.push([ss.active.getName(), position]);
c.orderTabs_(ss);
assert.deepEqual(moved.map(x => x[0]), ['Settings', 'Schools', 'Join requests', 'Events']);
assert.deepEqual(moved.map(x => x[1]), [1, 2, 3, 4], 'positions must be contiguous even with missing tabs');

console.log('fresh start tab reset checks passed');
