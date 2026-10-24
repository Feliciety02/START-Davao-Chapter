const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const html = fs.readFileSync(path.resolve(__dirname, '../Web_App.html'), 'utf8');
const start = html.indexOf('function joinSend()');
const end = html.indexOf('function password()', start);
assert.ok(start > 0 && end > start);

const fields = Object.fromEntries(['em', 'pw', 'cd', 'np', 'np2', 'jln', 'jfn', 'jmi', 'js', 'jp', 'jy', 'jya', 'jc', 'jprov', 'jcity']
  .map(id => [id, { value: '' }]));
Object.assign(fields, {
  em: { value: 'ana@example.com' }, pw: { value: 'existing-password' }, np: { value: 'new-good-password' }, np2: { value: 'new-good-password' },
  jln: { value: 'Rivera' }, jfn: { value: 'Ana' },
  js: { value: 'USeP' }, jp: { value: 'Computer Science' }, jy: { value: '1st year' },
  jya: { value: '2024' }, jc: { value: 'RA 7687' }, jprov: { value: 'Davao del Sur' },
  jcity: { value: 'Davao City' }, 'pending-note': {}, joinedt: {}
});
const calls = [], screens = [], responses = [];
const c = vm.createContext({
  AUTH: { joinStage: 0, mode: 'create', purpose: 'new' }, S: { email: 'ana@example.com' },
  $: id => fields[id],
  renderJoinStage() {}, joinValidate: () => true,
  busy: value => { c.AUTH.busy = value; },
  say: (message, good) => { c.message = message; c.good = good; },
  go: screen => { c.AUTH.step = screen; screens.push(screen); },
  keep() {},
  esc: value => String(value),
  call: (...args) => { calls.push(args); return responses.shift()(); }
});
vm.runInContext(html.slice(start, end), c);
const authEnd = html.indexOf('function signedIn(', end);
vm.runInContext(html.slice(end, authEnd), c);
c.keepMe = () => true;
c.signedIn = token => { c.signedInToken = token; };
const settle = () => new Promise(resolve => setImmediate(resolve));

(async () => {
  responses.push(() => Promise.resolve({ ok: true, auto: true, hasPassword: false }));
  c.joinSend(); await settle();
  assert.equal(screens.at(-1), 'welcome');
  assert.deepEqual(calls.map(x => x[0]), ['webJoinRequest'], 'automatic approval does not request a second code');

  responses.push(() => Promise.resolve({ ok: true, pending: true, msg: 'Request saved.' }));
  c.joinSend(); await settle();
  assert.equal(screens.at(-1), 'pending');

  responses.push(() => Promise.resolve({ ok: false, msg: 'Something went wrong on our side.' }));
  responses.push(() => Promise.resolve({ ok: true, state: 'setup' }));
  c.joinSend(); await settle();
  assert.equal(calls.at(-1)[0], 'webJoinStatus');
  assert.equal(screens.at(-1), 'welcome', 'a saved approval is recovered after an unexpected response');

  responses.push(() => Promise.reject(new Error('connection dropped')));
  responses.push(() => Promise.resolve({ ok: true, state: 'pending' }));
  c.joinSend(); await settle();
  assert.equal(screens.at(-1), 'pending', 'a saved pending request is recovered after a dropped response');

  responses.push(() => Promise.resolve({ ok: true, mode: 'setup' }));
  c.begin(''); await settle();
  assert.equal(screens.at(-1), 'welcome');
  responses.push(() => Promise.resolve({ ok: true, mode: 'password' }));
  c.begin(''); await settle();
  assert.equal(screens.at(-1), 'signin');
  responses.push(() => Promise.resolve({ ok: true, mode: 'code', purpose: 'new', retryAfter: 60 }));
  c.begin('setup'); await settle();
  assert.equal(screens.at(-1), 'code');
  assert.ok(c.AUTH.codeUntil > Date.now());

  responses.push(() => Promise.resolve({ ok: true, token: 'first-session' }));
  c.AUTH.ticket = 'verified-ticket'; c.newPassword(); await settle();
  assert.equal(screens.at(-1), 'ready');
  assert.equal(c.S.token, 'first-session');
  responses.push(() => Promise.resolve({ ok: true, token: 'returning-session' }));
  c.password(); await settle();
  assert.equal(c.signedInToken, 'returning-session', 'existing password still signs in directly');
  responses.push(() => Promise.reject(new Error('response lost after save')));
  responses.push(() => Promise.resolve({ ok: true, state: 'password' }));
  c.AUTH.purpose = 'new'; c.newPassword(); await settle();
  assert.equal(screens.at(-1), 'ready');
  assert.equal(c.AUTH.readyNeedsSignIn, true, 'lost setup response is recovered from saved account state');
  console.log('Frontend registration state checks passed');
})().catch(error => { console.error(error); process.exitCode = 1; });
