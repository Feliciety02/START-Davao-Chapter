const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const html = fs.readFileSync(path.resolve(__dirname, '../Web_App.html'), 'utf8');
const slice = (from, to) => html.slice(html.indexOf(from), html.indexOf(to, html.indexOf(from)));

// A declined request must not inherit the pending request's review instructions.
const elements = new Map();
const element = id => {
  if (!elements.has(id)) elements.set(id, { value: 'ana@example.com', textContent: '', style: {}, classList: { toggle() {} }, focus() {} });
  return elements.get(id);
};
const auth = vm.createContext({
  AUTH: { purpose: 'new' }, S: { email: 'ana@example.com' }, REGION: 'XI',
  $: element, document: { querySelector: () => ({ classList: { toggle() {} } }) },
  setTimeout() {}, say() {}, busy() {}, keep() {}, renderJoinStage() {}, showCodeCooldown() {},
  call: () => Promise.resolve({ ok: true, state: 'declined' })
});
vm.runInContext(slice('function go(step)', 'function say(t, ok)'), auth);
vm.runInContext(slice('function checkRegistration(fallback)', 'function begin(purpose, again)'), auth);
(async () => {
  await auth.checkRegistration();
  assert.equal(element('h').textContent, 'Registration not approved');
  assert.match(element('pending-note').textContent, /Contact your Regional Representative/);
  assert.doesNotMatch(element('pending-note').textContent, /reviewing|check for a decision/i);
  auth.go('pending');
  assert.match(element('pending-note').textContent, /reviewing your request/);

  // A selected calendar event appears once, even when it is also upcoming.
  const calendar = vm.createContext({
    Date, START_CAL_MONTH: new Date(2026, 9, 1).getTime(), START_CAL_DAY: '2026-10-25',
    esc: String, eventCards: list => list.map(e => '[' + e.title + ']').join(''),
    calButtons: () => '', card: () => ''
  });
  vm.runInContext(slice('function startCalDate(ev)', 'function eventsView(d)'), calendar);
  const view = calendar.startCalendarView([
    { title: 'Selected event', date: '2026-10-25' },
    { title: 'Later event', date: '2026-10-26' }
  ]);
  assert.equal(view.match(/\[Selected event\]/g).length, 1);
  assert.equal(view.match(/\[Later event\]/g).length, 1);
  assert.match(view, /Other upcoming START events/);
  const empty = calendar.startCalendarView([]);
  assert.match(empty, /No START events scheduled for this day/);
  assert.doesNotMatch(empty, /No upcoming events available/);
  console.log('Dashboard copy and registration status checks passed');
})().catch(error => { console.error(error); process.exitCode = 1; });
