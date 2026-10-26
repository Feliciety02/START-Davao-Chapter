const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', '02_Menu_Setup.gs'), 'utf8');
const start = source.indexOf('function safe_(');
const end = source.indexOf('function requireSetup_(', start);
assert.ok(start > 0 && end > start);

let status = 'NOT_REQUIRED';
let alert = '';
let dialog = '';
const context = vm.createContext({
  checkTrackerCopy_: () => false,
  Logger: { log() {} },
  ScriptApp: {
    AuthMode: { FULL: 'FULL' }, AuthorizationStatus: { REQUIRED: 'REQUIRED' },
    getAuthorizationInfo: () => ({ getAuthorizationStatus: () => status,
      getAuthorizationUrl: () => 'https://script.google.com/authorize' })
  },
  SpreadsheetApp: { getUi: () => ({ ButtonSet: { OK: 'OK' }, alert: (_title, message) => { alert = message; } }) },
  dlg_: (_title, html) => { dialog = html; },
  esc_: value => String(value)
});
vm.runInContext(source.slice(start, end), context);
const fail = () => { throw new Error('You do not have permission to access this Drive folder.'); };
context.safe_(fail);
assert.match(alert, /access this Drive folder/);
assert.doesNotMatch(alert, /Click the menu item again/);
status = 'REQUIRED';
context.safe_(fail);
assert.match(dialog, /https:\/\/script.google.com\/authorize/);
assert.match(dialog, /Open Google permission page/);
console.log('Menu authorization diagnostics checks passed');
