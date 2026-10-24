const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const c = vm.createContext({ console });
for (const file of fs.readdirSync(root).filter(name => name.endsWith('.gs')).sort())
  vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), c, { filename: file });

const schools = [
  { name: 'AdDU', inactive: false },
  { name: 'UIC', inactive: true },
  { name: 'UM', inactive: false },
  { name: 'UM Main', inactive: true }
];
const ss = { getUrl: () => 'https://example.test/tracker', getSheetByName: () => ({ getRange: () => ({ getValues: () => [] }) }) };
c.schoolRows_ = () => schools;
c.webSheet_ = () => ss;
c.webSafe_ = fn => fn();
c.joinPlaces_ = () => ({});

const publicNames = c.webJoinSchools().schools.map(s => s.name);
assert.deepEqual([...publicNames].sort(), ['AdDU', 'UIC', 'UM Main'].sort(), 'registration lists one main campus');
assert.ok(publicNames.includes('UIC'), 'UIC appears even when marked inactive last term');
assert.equal(c.joinSchoolList_(ss).find(s => s.name === 'UIC').label, 'University of the Immaculate Conception (UIC)');
assert.equal(c.joinSchoolList_(ss).find(s => s.name === 'UM Main').label, 'University of Mindanao — UM Main');

c.webSession_ = () => ({ e: 'rr@example.com' });
c.webRoles_ = () => ({ admin: true });
c.scholarByEmail_ = () => ({});
assert.deepEqual([...c.webViewList('token').schools.map(s => s.name)].sort(), [...publicNames].sort(), 'RR preview uses the same schools');

c.focalData_ = () => [];
c.checkinOn_ = () => false;
c.webScholarRoster_ = () => [];
c.webEvents_ = () => [];
c.reminderWeek_ = () => ({ deadline: new Date() });
c.webFmt_ = () => '';
c.webJoins_ = () => [];
c.docProps_ = () => ({ getProperty: () => null });
c.missedSchools_ = () => [];
c.dataFlowQuality_ = () => ({ count: 0, issues: [] });
assert.deepEqual([...c.webAdmin_(ss, {}).scholarChoices.schools].sort(), [...publicNames].sort(), 'RR roster and editor use the same schools');

const html = fs.readFileSync(path.join(root, 'Web_App.html'), 'utf8');
const start = html.indexOf('function joinCard(');
const end = html.indexOf('// RR history:', start);
const ui = vm.createContext({
  S: { data: { admin: { scholarChoices: { schools: publicNames } } } },
  esc: value => String(value),
  ico: () => '',
  schoolBadge: () => ''
});
vm.runInContext(html.slice(start, end), ui);
const card = ui.joinCard([{ row: 2, name: 'Test', email: 'test@example.com', typed: true, schoolLabel: 'Other', elig: 'Needs review' }], false);
assert.match(card, /<option>UIC<\/option>/, 'join review matching includes UIC');
assert.match(card, /<option>UM Main<\/option>/, 'join review matching includes UM Main');

console.log('School dropdown checks passed');
