const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
function startContext() {
  const c = vm.createContext({ console });
  for (const file of fs.readdirSync(root).filter(name => name.endsWith('.gs')).sort())
    vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), c, { filename: file });
  c.Logger = { log() {} };
  c.webSafe_ = fn => fn();
  return c;
}
function properties() {
  const values = new Map();
  return {
    getProperty: key => values.has(key) ? values.get(key) : null,
    setProperty: (key, value) => { values.set(key, String(value)); },
    deleteProperty: key => { values.delete(key); },
    getProperties: () => Object.fromEntries(values)
  };
}

// A new approved scholar has no password yet. The email check only opens setup;
// the code is sent after they confirm the registered address.
{
  const c = startContext(), props = properties();
  let active = true, hasPassword = false, codeCalls = 0, joinState = 'waiting';
  c.webSheet_ = () => ({});
  c.primaryEmail_ = (_, email) => email;
  c.webMoveAccount_ = () => {};
  c.webRoles_ = (_, email) => ({ email, member: active ? {} : null, leads: [] });
  c.webAccount_ = () => hasPassword ? { h: 'saved hash' } : null;
  c.webMailCode_ = () => { codeCalls++; return { ok: true, retryAfter: 60 }; };
  c.scholarByEmail_ = () => ({});
  c.joinStateOf_ = () => joinState;
  c.docProps_ = () => props;
  c.webHash_ = s => s;
  assert.equal(c.webStart('new@example.com', '').mode, 'setup');
  assert.equal(codeCalls, 0);
  assert.equal(c.webStart('new@example.com', 'setup').mode, 'code');
  assert.equal(codeCalls, 1);
  hasPassword = true;
  assert.equal(c.webStart('new@example.com', '').mode, 'password');
  assert.equal(c.webStart('new@example.com', 'reset').purpose, 'reset');
  assert.equal(c.webStart('new@example.com', 'setup').purpose, 'reset', 'a saved password can be replaced only after email verification');
  active = false;
  assert.equal(c.webStart('new@example.com', '').mode, 'pending');
  assert.equal(c.webJoinStatus('new@example.com').state, 'pending');
  joinState = 'added';
  assert.equal(c.webStart('new@example.com', '').mode, 'access-check');
  assert.equal(c.webJoinStatus('new@example.com').state, 'access-check');
  active = true; hasPassword = false;
  assert.equal(c.webJoinStatus('new@example.com').state, 'setup');
  hasPassword = true;
  assert.equal(c.webJoinStatus('new@example.com').state, 'password');
}

// Registration is one Sheet row. If a reply is lost, the status endpoint gives
// the saved state; a second submit does not create a second request.
{
  const c = startContext(), props = properties(), rows = [];
  let active = false, auto = true, approvals = 0, auditFails = true;
  const sheet = {
    getLastRow: () => rows.length + 1,
    getMaxRows: () => Math.max(rows.length + 1, 2),
    getRange: (row, col, height, width) => ({
      getValues: () => rows.slice(row - 2, row - 2 + height).map(r => r.slice(col - 1, col - 1 + width)),
      setValues: data => { rows[row - 2] = data[0].slice(); }
    })
  };
  c.webSheet_ = () => ({});
  c.joinTab_ = () => sheet;
  c.schoolRows_ = () => [{ name: 'USeP' }];
  c.joinPlaces_ = () => ({ 'Davao del Sur': ['Davao City'] });
  c.webRoles_ = () => ({ member: active ? {} : null, leads: [] });
  c.autoApproveOn_ = () => auto;
  c.webJoinAction_ = () => { approvals++; active = true; rows[0][4] = 'Added automatically Oct 24'; return { ok: true }; };
  c.markJoinChange_ = () => { if (auditFails) throw new Error('audit unavailable'); };
  c.refreshJoinViews_ = () => {};
  c.docProps_ = () => props;
  c.docLock_ = () => ({ tryLock: () => true, releaseLock() {} });
  c.webHash_ = s => s;
  c.Utilities = { formatDate: () => '20261024' };
  const details = { last: 'Rivera', first: 'Ana', mi: '', program: 'Computer Science', year: c.JOIN.years[0],
    awarded: '2024', kind: c.JOIN.kinds[0], city: 'Davao City', province: 'Davao del Sur' };
  const send = () => c.webJoinRequest('ana@example.com', '', 'USeP', '', details);
  const first = send();
  assert.equal(first.ok, true);
  assert.equal(first.auto, true);
  assert.equal(first.hasPassword, false);
  assert.equal(rows.length, 1);
  assert.equal(approvals, 1);
  assert.equal(c.webJoinStatus('ana@example.com').state, 'setup');
  assert.equal(send().onList, true);
  assert.equal(rows.length, 1);
  active = false; auto = false; auditFails = false;
  assert.equal(send().accessCheck, true, 'an approved row never creates another join request');
  assert.equal(rows.length, 1);
  rows.length = 0;
  props.deleteProperty('WEB_JD_20261024');
  const pending = send();
  assert.equal(pending.ok, true);
  assert.equal(pending.pending, true);
  assert.equal(pending.auto, undefined);
  assert.equal(c.webJoinStatus('ana@example.com').state, 'pending');
  assert.equal(send().pending, true);
  assert.equal(rows.length, 1);
}

// The real approval action creates a Scholar row when the old masterlist is empty.
{
  const c = startContext(), request = new Array(18).fill(''), scholars = [];
  request[1] = 'Rivera, Ana'; request[2] = 'USeP'; request[3] = 'ana@example.com';
  request[4] = 'New'; request[7] = 'Computer Science'; request[8] = c.JOIN.years[0];
  request[10] = 'Eligible'; request[15] = '2024';
  const joinSheet = {
    getLastRow: () => 2,
    getRange: (row, col) => ({ getValues: () => [request], setValue: value => { request[col - 1] = value; } }),
    getMaxColumns: () => 18
  };
  const scholarSheet = {
    getLastRow: () => scholars.length + 1,
    getMaxColumns: () => 13,
    getMaxRows: () => 2,
    getRange: () => ({ getValues: () => scholars, setValues: rows => { scholars.push(...rows); } })
  };
  c.joinTab_ = () => joinSheet;
  c.schoolRows_ = () => [{ name: 'USeP' }];
  c.oldMasterRows_ = () => [];
  c.autoApproveOn_ = () => true;
  c.docLock_ = () => ({ tryLock: () => true, releaseLock() {} });
  c.Utilities = { formatDate: () => 'Oct 24' };
  c.SpreadsheetApp = { flush() {} };
  c.refreshJoinViews_ = () => {};
  c.calPeopleChanged_ = () => {};
  const ss = { getSheetByName: name => name === c.TAB.scholars ? scholarSheet : null };
  const approved = c.webJoinAction_(ss, 'join-add', { row: 2, auto: true });
  assert.equal(approved.ok, true, 'an eligible new scholar is approved without a masterlist');
  assert.equal(scholars.length, 1);
  assert.equal(scholars[0][4], 'ana@example.com');
  assert.equal(scholars[0][5], 'Active');
  assert.match(request[4], /^Added automatically/);
}

// Codes are rate limited, single-use and time limited. A failed mail send must
// restore the previous code state so the scholar can retry.
{
  const c = startContext(), props = properties();
  let mailFails = false, sends = 0, uuid = 1, approved = true;
  c.docProps_ = () => props;
  c.docLock_ = () => ({ tryLock: () => true, releaseLock() {} });
  c.webHash_ = s => String(s);
  c.webEmail_ = (_, __, code) => ({ subject: 'Code', html: '', text: code, from: 'START' });
  c.webSheet_ = () => ({});
  c.primaryEmail_ = (_, email) => email;
  c.webRoles_ = () => ({ member: approved ? {} : null, leads: [] });
  c.Utilities = { formatDate: () => '20261024', getUuid: () => String(uuid++).padStart(12, '0') };
  c.MailApp = { getRemainingDailyQuota: () => 100, sendEmail: () => { sends++; if (mailFails) throw new Error('mail unavailable'); } };
  const email = 'ana@example.com';
  const sent = c.webMailCode_(email, 'new', { name: 'Ana' });
  assert.equal(sent.ok, true);
  assert.equal(c.webMailCode_(email, 'new', { name: 'Ana' }).recent, true);
  assert.equal(sends, 1);
  const rec = JSON.parse(props.getProperty('WEB_C_' + email));
  const code = String(100000 + parseInt('000000000001', 16) % 900000);
  assert.equal(c.webCheckCode(email, '000000').ok, false);
  approved = false;
  assert.match(c.webCheckCode(email, code).msg, /approved/i, 'pending membership cannot use a code');
  approved = true;
  const checked = c.webCheckCode(email, code);
  assert.equal(checked.ok, true);
  assert.match(c.webCheckCode(email, code).msg, /already used/i);
  c.webPwHash_ = (password, salt) => password + '|' + salt;
  c.webPwProblem_ = () => '';
  c.webEndSessions_ = () => {};
  c.webNewSession_ = () => 'authenticated-session';
  c.markVerifiedEmail_ = () => {};
  const created = c.webSetPassword(checked.ticket, 'goodpass123', true);
  assert.equal(created.ok, true);
  assert.equal(created.token, 'authenticated-session');
  assert.equal(c.webSetPassword(checked.ticket, 'goodpass123', true).ok, false, 'setup ticket works once');
  assert.equal(c.webPasswordSignIn(email, 'goodpass123', true).ok, true);
  assert.equal(c.webPasswordSignIn(email, 'wrongpass', true).ok, false);
  rec.h = c.webHash_(email + '|' + code); rec.exp = Date.now() - 1;
  props.setProperty('WEB_C_' + email, JSON.stringify(rec));
  assert.match(c.webCheckCode(email, code).msg, /expired/i);
  mailFails = true;
  props.deleteProperty('WEB_C_' + email);
  assert.equal(c.webMailCode_(email, 'new', { name: 'Ana' }).ok, false);
  assert.equal(props.getProperty('WEB_C_' + email), null);
}

console.log('Registration, approval, and code flow checks passed');
