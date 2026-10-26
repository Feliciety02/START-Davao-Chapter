const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const context = vm.createContext({ console });
for (const file of fs.readdirSync(root).filter(name => name.endsWith('.gs')).sort()) {
  vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), context, { filename: file });
}

// Older styled trackers prefix database headings with icons; schema repair must accept them.
let headingWrites = 0;
const decoratedHeaders = {
  getMaxColumns: () => 21,
  getRange: () => ({ getValues: () => [['🔒 Request ID', '🔒 Updated at', '🔒 Updated by']], setValue: () => { headingWrites++; } }),
  getName: () => 'Join requests'
};
context.ensureHeaderBlock_(decoratedHeaders, 19, ['Request ID', 'Updated at', 'Updated by']);
assert.equal(headingWrites, 0);

// Auto-approval covers valid sign-ups from listed schools, even without a masterlist row.
const old = new Array(11).fill('');
old[0] = 'Dela Cruz, Ana'; old[1] = 'USeP'; old[2] = 'BS Biology'; old[3] = '2023';
old[4] = 'ana@example.com'; old[5] = 'Active';
context.oldMasterRows_ = () => [{ r: old, at: 2 }];
context.docProps_ = () => ({ getProperty: () => context.autoOn ? '1' : '0' });
context.schoolRows_ = () => [{ name: 'USeP' }];
const request = new Array(18).fill('');
request[1] = old[0]; request[2] = old[1]; request[3] = old[4];
request[7] = 'Computer Science'; request[10] = 'Eligible'; request[15] = old[3];
assert.equal(context.autoJoinEligible_({}, request, []), false);
context.autoOn = true;
assert.equal(context.autoJoinEligible_({}, request, []), true);
context.oldMasterRows_ = () => [];
assert.equal(context.autoJoinEligible_({}, request, []), true, 'a new eligible scholar is approved without an old record');
request[3] = 'new@example.com';
assert.equal(context.autoJoinEligible_({}, request, []), true, 'a different email is allowed for a new scholar');
request[3] = old[4];
assert.equal(context.autoJoinEligible_({}, request, [[old[0], old[1], '', '', old[4], 'Suspended']]), false);
assert.equal(context.autoJoinEligible_({}, request, [[old[0], old[1], '', '', old[4], 'Active']]), true);
assert.equal(context.autoJoinEligible_({}, request, [[old[0], old[1], '', '', 'other@example.com', 'Active']]), false, 'existing account email conflicts need review');
assert.equal(context.autoJoinEligible_({}, request, [['Another Name', old[1], '', '', old[4], 'Active']]), false, 'a reused email needs review');
request[7] = 'Not a listed program';
assert.equal(context.autoJoinEligible_({}, request, []), false);
request[7] = 'Computer Science'; request[10] = 'Needs review (school not on the list)';
assert.equal(context.autoJoinEligible_({}, request, []), false);
request[10] = 'Eligible'; request[2] = 'Unlisted school';
assert.equal(context.autoJoinEligible_({}, request, []), false);
request[2] = old[1];

// The Sheets checkbox controls the website switch. A matching legacy property is only a fallback
// until setup adds the Settings row.
const settingRow = context.SETTINGS.findIndex(s => s[0] === 'autoApprove') + 2;
let sheetToggle = false, savedToggle = '1';
context.webSheet_ = () => ({ getSheetByName: name => name === context.TAB.settings ? {
  getRange: (row, col) => ({
    getValue: () => col === 1 ? context.SETTINGS[settingRow - 2][1] : sheetToggle,
    setValue: value => { sheetToggle = value; }
  })
} : null });
context.docProps_ = () => ({ getProperty: () => savedToggle, setProperty: (_, value) => { savedToggle = value; } });
assert.equal(context.autoApproveOn_(), false);
assert.equal(context.setAutoApprove_(true), true);
assert.equal(sheetToggle, true);
assert.equal(savedToggle, '1');
assert.equal(context.autoApproveOn_(), true);
context.setAutoApprove_(false);
assert.equal(sheetToggle, false);
assert.equal(savedToggle, '0');
context.setAutoApprove_(true);

// A strict legacy match can still enrich an automatic approval without deciding eligibility.
const other = old.slice(); other[4] = 'other@example.com'; other[2] = 'BS Mathematics';
context.oldMasterRows_ = () => [{ r: other, at: 2 }, { r: old, at: 3 }];
request[7] = old[2];
assert.equal(context.autoJoinMasterMatch_({}, request, []).at, 3);

// A historical snapshot keeps its original denominator when the current roster grows.
context.scholarByEmail_ = () => ({
  'ana@example.com': { primary: 'ana@example.com', status: 'Active', school: 'USeP' },
  'ben@example.com': { primary: 'ben@example.com', status: 'Active', school: 'USeP' }
});
context.memberWeek_ = () => ({ due: new Date('2026-10-25T12:00:00Z') });
context.memberWeekKey_ = d => d.getUTCDate() <= 18 ? '2026-10-18' : '2026-10-25';
context.ymd_ = d => d.toISOString().slice(0, 10);
context.webFmt_ = d => d.toISOString().slice(5, 10);
context.memberCols_ = () => ({ email: 1, width: 2 });
const answers = [
  [vm.runInContext("new Date('2026-10-17T12:00:00Z')", context), 'ana@example.com'],
  [vm.runInContext("new Date('2026-10-17T12:00:00Z')", context), 'ben@example.com'],
  [vm.runInContext("new Date('2026-10-24T12:00:00Z')", context), 'ana@example.com']
];
context.memberTab_ = () => ({ getLastRow: () => 4, getRange: () => ({ getValues: () => answers }) });
const snap = {
  getLastRow: () => 4,
  getRange: () => ({ getDisplayValues: () => [
    ['2026-10-18', 'id-ana', 'ana@example.com', 'USeP'],
    ['2026-10-25', 'id-ana', 'ana@example.com', 'USeP'],
    ['2026-10-25', 'id-ben', 'ben@example.com', 'USeP']
  ] })
};
const ss = { getSheetByName: name => name === context.WEEKLY_ROSTER_TAB_ ? snap : null };
const rates = context.weeklyRates_(ss, 2);
assert.equal(rates[0].members, 1);
assert.equal(rates[0].answered, 1);
assert.equal(rates[0].rate, 100);
assert.equal(rates[1].members, 2);
assert.equal(rates[1].answered, 1);
assert.equal(rates[1].rate, 50);
assert.equal(rates[0].estimated, false);

// Repairing existing scholars should write metadata in one batch and preserve prior values.
const scholarData = [
  ['Ana', 'USeP', '', '', 'ana@example.com', 'Active'].concat(new Array(12).fill('')),
  ['Ben', 'USeP', '', '', 'ben@example.com', 'Active'].concat(new Array(7).fill(''), ['existing-id', '', 'old-date', 'old-editor', ''])
];
let metadataWrites = 0, writtenMetadata;
const scholarSheet = {
  getMaxColumns: () => 18,
  getLastRow: () => 3,
  getRange: (row, col, count, width) => ({
    getValues: () => row === 1 ? [context.SCHOLAR_META_COLS_] : scholarData,
    setValues: values => { if (row === 2 && col === 14 && count === 2 && width === 5) { metadataWrites++; writtenMetadata = values; } }
  })
};
context.dataLogTab_ = () => {};
context.weeklyRosterTab_ = () => {};
context.docProps_ = () => ({ getProperties: () => ({}) });
context.webAcctKey_ = email => email;
context.Utilities = { getUuid: () => 'new-id' };
context.ensureDataFlowSchema_({ getSheetByName: name => name === context.TAB.scholars ? scholarSheet : null });
assert.equal(metadataWrites, 1);
assert.equal(writtenMetadata[0][0], 'new-id');
assert.equal(writtenMetadata[1][0], 'existing-id');
assert.equal(writtenMetadata[1][2], 'old-date');
console.log('Data flow checks passed');
