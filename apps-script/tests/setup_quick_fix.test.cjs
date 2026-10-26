const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', '02_Menu_Setup.gs'), 'utf8');
const start = source.indexOf('function schoolsNeedRepair_(');
const end = source.indexOf('// The full build', start);
assert.ok(start > 0 && end > start);
const fullStepsStart = source.indexOf('function setupSteps_(');
const fullStepsEnd = source.indexOf('var SETUP_VERSION', fullStepsStart);
const stepsContext = vm.createContext({ DB_MODE: true });
vm.runInContext(source.slice(fullStepsStart, fullStepsEnd), stepsContext);
const fullSteps = stepsContext.setupSteps_({}).map(step => step[0]);
assert.ok(fullSteps.includes('Data flow schema'));
assert.ok(fullSteps.includes('Protections and automatic updates'));
assert.equal(fullSteps.includes('Theme'), false);
assert.equal(fullSteps.includes('Same table look on every tab'), false);
assert.equal(fullSteps.includes('Database protections'), false);
assert.ok(fullSteps.includes('Database headings'));
assert.ok(fullSteps.includes('Campus Lead check-in formulas'));
assert.equal(fullSteps.includes('Extras: check-in open/closed, email checks, leaderboard'), false);
assert.equal(fullSteps.includes('Form school list'), false);

const headers = ['', 'Keep this column', 'Old header'];
let headerWrites = 0;
const headerSheet = {
  getMaxColumns: () => 3,
  getRange: () => ({ getValues: () => [headers.slice()], setValues: rows => { headers.splice(0, headers.length, ...rows[0]); headerWrites++; } })
};
stepsContext.TAB = { schools: 'Schools', scholars: 'Scholars' };
stepsContext.columnSpecs_ = () => ({ Schools: [[1, 'edit', 'School'], [3, 'auto', 'Active scholars']] });
stepsContext.ensureDatabaseHeaders_({ getSheetByName: name => name === 'Schools' ? headerSheet : null });
assert.deepEqual(headers, ['School', 'Keep this column', 'Active scholars']);
stepsContext.ensureDatabaseHeaders_({ getSheetByName: name => name === 'Schools' ? headerSheet : null });
assert.equal(headerWrites, 1);

const props = new Map([['DB_TIDY', '2'], ['REGION_APPLIED', 'XI|Davao'], ['FRESH_LIST_DONE', '1'], ['SCHEDULE_SYNC', 'current']]);
const calls = [];
let message = '';
let missingFormula = false;
let failSchoolRepair = false;
let owner = false;
let handlers = ['weeklySummary', 'autoReminder'];
const school = {
  getMaxRows: () => 101,
  getMaxColumns: () => 17,
  getRange: (...args) => ({
    getFormula: () => '=ARRAYFORMULA(Scholars!A:A)',
    getFormulas: () => args[1] === 17
      ? Array.from({ length: 100 }, () => ['=Q'])
      : Array.from({ length: 100 }, (_, i) => [i === 0 && missingFormula ? '' : '=C', '', '', '', '=G', '=H'])
  })
};
const sheets = new Map([['Schools', school], ['Weekly Check-in', {}], ['Check-in answers', {}]]);
const spreadsheet = { toast() {}, getSheetByName: name => sheets.get(name) };
const context = vm.createContext({
  DB_MODE: true, WEB_CHECKIN: true, TAB: { schools: 'Schools', scholars: 'Scholars', checkins: 'Weekly Check-in', mcheckins: 'Check-in answers' },
  CONFIG: { maxSchools: 100 }, OLD_MEMBER_TAB: 'Member check-ins',
  SpreadsheetApp: { getActive: () => spreadsheet },
  docProps_: () => ({ getProperty: key => props.get(key), setProperty: (key, value) => props.set(key, value) }),
  outOfTime_: () => false, migrateTabs_: () => calls.push('migrate'), tidyToDatabase_: () => 'done',
  regionProfile_: () => ({ code: 'XI', label: 'Davao' }), mergeUmMainSchool_: () => false,
  buildSchools_: () => { calls.push('buildSchools'); if (failSchoolRepair) throw new Error('formula failure'); },
  ensureDataFlowSchema_: () => calls.push('schema'), captureWeeklyRoster_: () => calls.push('roster'),
  syncRoles_: () => calls.push('roles'), claimIfFree_: () => owner,
  installSubmitTrigger_: () => calls.push('submitTrigger'), installSyncTrigger_: () => calls.push('syncTrigger'),
  installExtrasTrigger_: () => calls.push('extrasTrigger'), scheduleKey_: () => 'current',
  getSettings_: () => ({ autoWeekly: true }), applySchedule_: () => calls.push('schedule'),
  ScriptApp: { getProjectTriggers: () => handlers.map(name => ({ getHandlerFunction: () => name })) },
  protect_: () => calls.push('protect'), notify_: value => { message = value; },
  Logger: { log() {} }
});
vm.runInContext(source.slice(start, end), context);

context.quickFix_();
assert.equal(calls.includes('buildSchools'), false);
assert.deepEqual(calls, ['migrate', 'schema', 'roster', 'roles', 'protect']);
assert.match(message, /Tracker check complete/);

missingFormula = true;
calls.length = 0;
context.quickFix_();
assert.equal(calls.filter(name => name === 'buildSchools').length, 1);

failSchoolRepair = true;
context.quickFix_();
assert.match(message, /Some checks failed:[\s\S]*School formulas/);
assert.match(message, /formula failure/);

failSchoolRepair = false;
missingFormula = false;
owner = true;
calls.length = 0;
context.quickFix_();
assert.equal(calls.includes('schedule'), false);
handlers = ['weeklySummary'];
calls.length = 0;
context.quickFix_();
assert.equal(calls.filter(name => name === 'schedule').length, 1);

// An already protected database should not remove and recreate protections on every click.
const protectStart = source.indexOf('function protect_(ss)');
const protectEnd = source.indexOf('/* ---- Form look', protectStart);
let removed = 0;
const makeProtection = sheet => ({
  getDescription: () => 'START protection',
  getRange: () => ({ getSheet: () => sheet }),
  getProtectionType: () => 'SHEET',
  isWarningOnly: () => false,
  canDomainEdit: () => false,
  getEditors: () => [{ getEmail: () => 'owner@example.com' }],
  remove: () => { removed++; }
});
const protectedSheets = ['Settings', 'Schools'].map(name => ({ getName: () => name }));
const protections = protectedSheets.map(makeProtection);
const protectContext = vm.createContext({
  DB_MODE: true, PROTECT_TAG: 'START protection', TAB: { settings: 'Settings', schools: 'Schools' },
  SpreadsheetApp: { ProtectionType: { RANGE: 'RANGE', SHEET: 'SHEET' } },
  me_: () => 'owner@example.com', automationOwner_: () => 'owner@example.com'
});
vm.runInContext(source.slice(protectStart, protectEnd), protectContext);
protectContext.protect_({
  getProtections: type => type === 'SHEET' ? protections : [],
  getSheetByName: name => protectedSheets.find(sheet => sheet.getName() === name),
});
assert.equal(removed, 0);
console.log('Quick setup repair checks passed');
