const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', '06_Calendar.gs'), 'utf8');
const start = source.indexOf('var CALENDAR_SCOPE_ =');
const end = source.indexOf('/* ---------- menu ---------- */', start);
assert.ok(start > 0 && end > start);
let status = 'REQUIRED';
let opened = 0;
let dialog = '';
const context = vm.createContext({
  ScriptApp: { AuthMode: { FULL: 'FULL' }, AuthorizationStatus: { REQUIRED: 'REQUIRED' },
    getAuthorizationInfo: (_mode, scopes) => {
      assert.equal(scopes[0], 'https://www.googleapis.com/auth/calendar');
      return { getAuthorizationStatus: () => status,
        getAuthorizationUrl: () => 'https://script.google.com/calendar-consent' };
    } },
  CalendarApp: { getDefaultCalendar: () => { opened++; } },
  dlg_: (_title, html) => { dialog = html; }, esc_: x => String(x)
});
vm.runInContext(source.slice(start, end), context);
assert.equal(context.calendarAuthorized_(), false);
assert.equal(opened, 0, 'do not access Calendar before consent');
assert.match(dialog, /calendar-consent/);
status = 'NOT_REQUIRED';
assert.equal(context.calendarAuthorized_(), true);
assert.equal(opened, 1);

const setup = source.slice(source.indexOf('function setupCalendars()'), source.indexOf('function stopCalendarUpdates()'));
assert.ok(setup.indexOf('calendarAuthorized_()') < setup.indexOf('ensureOwner_('), 'authorize before changing owner');
assert.ok(setup.indexOf('runCalendarSync_(') < setup.indexOf("setProperty('CAL_ON'"), 'enable automatic updates after first sync');
console.log('Calendar authorization checks passed');
