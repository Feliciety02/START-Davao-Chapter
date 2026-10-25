const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const c = vm.createContext({ console });
for (const file of fs.readdirSync(root).filter(name => name.endsWith('.gs')).sort()) {
  vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), c, { filename: file });
}
c.Logger = { log: () => {} };

assert.equal(c.webValidDate_('2026-02-30'), null);
assert.equal(c.webValidDate_('2026-02-28') instanceof Date, false); // cross-context Date
assert.equal(c.webValidDate_('2026-02-28').getDate(), 28);

c.schoolRows_ = () => [{ name: 'School A' }];
c.docLock_ = () => ({ tryLock: () => true, releaseLock: () => {} });
c.webBumpData_ = () => {};
c.renderCalView_ = () => {};
let outreachSaved = null;
c.saveOutreach = (arg, ss) => { outreachSaved = [arg, ss]; };
const ss = { getSheetByName: () => null };
const baseOutreach = { school: 'School A', date: '2026-10-25', person: 'Pat', channel: 'Email',
  outcome: c.OUTREACH_OUTCOMES[0], next: 'Follow up', follow: '' };
assert.equal(c.webOutreachSave_(ss, 'rr@example.com', { ...baseOutreach, next: '=IMPORTDATA("x")' }).ok, false);
assert.equal(c.webOutreachSave_(ss, 'rr@example.com', { ...baseOutreach, date: '2026-02-30' }).ok, false);
assert.equal(c.webOutreachSave_(ss, 'rr@example.com', baseOutreach).ok, true);
assert.equal(outreachSaved[1], ss);

let eventSaved = null;
c.addEventFromBox = (arg, sheet) => { eventSaved = [arg, sheet]; return 'Event saved'; };
const baseEvent = { title: 'Meeting', date: '2026-10-25', start: '09:00', end: '10:00',
  type: c.EV_TYPES[0], who: c.WHO.all, rep: c.EV_REPEATS[0], until: '', place: '', details: '', invite: false };
assert.equal(c.webEventSave_(ss, 'rr@example.com', { ...baseEvent, who: 'Unknown school' }).ok, false);
assert.equal(c.webEventSave_(ss, 'rr@example.com', { ...baseEvent, end: '08:00' }).ok, false);
assert.equal(c.webEventSave_(ss, 'rr@example.com', baseEvent).ok, true);
assert.equal(eventSaved[1], ss);

const history = Array.from({ length: 55 }, (_, i) => ({ row: i + 2, status: 'Added' }));
c.webSession_ = () => ({ e: 'rr@example.com' });
c.webSheet_ = () => ss;
c.webRoles_ = () => ({ admin: true, suspended: false });
c.webJoins_ = () => history;
const page = c.webJoinHistory('token', 'Approved', 20);
assert.equal(page.ok, true);
assert.equal(page.rows.length, 20);
assert.equal(page.rows[0].row, 36);
assert.equal(page.total, 55);

c.PLACE_PICS = [{ name: 'A', src: 'data:image/webp;base64,YQ==' }];
c.placesOn_ = () => true;
c.webRoles_ = () => ({ member: { school: 'School A' }, email: 'member@example.com', admin: false });
c.webMyAnswers_ = () => [];
c.webGame_ = () => ({ cards: [0] });
const postcards = c.webPostcards('token');
assert.equal(postcards.ok, true);
assert.equal(postcards.cards[0].got, true);
assert.equal('src' in postcards.cards[0], false);

const html = fs.readFileSync(path.join(root, 'Web_App.html'), 'utf8');
const start = html.indexOf('var ROSTER_UI =');
const end = html.indexOf('function openScholarEditor(', start);
assert.ok(start >= 0 && end > start);
const front = vm.createContext({
  S: { data: { admin: null } },
  esc: value => String(value == null ? '' : value).replace(/[&<>"']/g, ''),
  schoolBadge: () => '', ico: () => '',
});
vm.runInContext(html.slice(start, end), front);
const scholars = Array.from({ length: 125 }, (_, i) => ({ row: i + 2, name: 'Scholar ' + String(i).padStart(3, '0'),
  email: 's' + i + '@example.com', school: 'School A', status: 'Active', verified: true }));
const admin = { scholars, scholarChoices: { schools: ['School A'] }, memberCheckins: false };
assert.equal((front.scholarManagementCard(admin).match(/data-rrsch/g) || []).length, 50);
front.ROSTER_UI.page = 2;
assert.equal((front.scholarManagementCard(admin).match(/data-rrsch/g) || []).length, 25);
front.ROSTER_UI.search = 'Scholar 100'; front.ROSTER_UI.page = 0;
assert.equal((front.scholarManagementCard(admin).match(/data-rrsch/g) || []).length, 1);

console.log('Website performance and workflow checks passed');
