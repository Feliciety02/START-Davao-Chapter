const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'Web_App.html'), 'utf8');
const settle = () => new Promise(resolve => setImmediate(resolve));

(async () => {
  const alerts = [], responses = [];
  let pending = { a: { vibe: 'Good' } }, drops = 0;
  const offline = vm.createContext({
    S: { token: 'token', data: { me: { email: 'ana@example.com' } } },
    pendingGet: () => pending,
    dropPending: () => { drops++; pending = null; },
    call: () => responses.shift()(),
    banner: (message, tone) => alerts.push([message, tone]),
    draw() {}, load() {}
  });
  vm.runInContext(html.slice(html.indexOf('var SENDING = false;'), html.indexOf('function banner(', html.indexOf('var SENDING = false;'))), offline);
  responses.push(() => Promise.resolve({ ok: false, msg: 'START is busy' }));
  offline.sendPending(); await settle();
  assert.equal(drops, 0, 'a rejected retry keeps the saved answer');
  assert.ok(pending);
  assert.match(alerts.at(-1)[0], /still on this device/);
  responses.push(() => Promise.resolve({ ok: true, member: {}, already: false }));
  offline.sendPending(); await settle();
  assert.equal(drops, 1, 'a confirmed save clears the queued answer');

  const accountCalls = [], accountResponses = [], accountAlerts = [];
  const account = vm.createContext({
    S: { token: 'token' },
    call: (...args) => { accountCalls.push(args); return accountResponses.shift()(); },
    keep: (...args) => accountCalls.push(args),
    $: () => null,
    showSignIn: message => { account.message = message; },
    banner: message => accountAlerts.push(message)
  });
  vm.runInContext(html.slice(html.indexOf('function signOutAll()'), html.indexOf('function signOut()', html.indexOf('function signOutAll()'))), account);
  accountResponses.push(() => Promise.reject(new Error('offline')));
  account.signOutAll(); await settle();
  assert.equal(account.S.token, 'token', 'a failed remote sign-out does not claim all devices were signed out');
  assert.equal(account.message, undefined);
  assert.match(accountAlerts.at(-1), /Could not reach START/);
  accountResponses.push(() => Promise.resolve({ ok: true }));
  account.signOutAll(); await settle();
  assert.equal(account.S.token, null);
  assert.match(account.message, /Signed out on all your devices/);

  const server = vm.createContext({ console });
  for (const file of fs.readdirSync(root).filter(name => name.endsWith('.gs')).sort())
    vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), server, { filename: file });
  server.webSafe_ = fn => fn();
  server.webSession_ = () => null;
  assert.equal(server.webSignOutAll('expired').ok, false, 'an expired session cannot report all-device sign-out');

  assert.match(html, /if \(fresh === true\) VA\.list = null;/, 'fresh RR loads refresh preview school choices');
  assert.match(html, /class="card mehead me-profile-card"/, 'the Me card uses its responsive layout rules');
  console.log('Account and offline consistency checks passed');
})().catch(error => { console.error(error); process.exitCode = 1; });
