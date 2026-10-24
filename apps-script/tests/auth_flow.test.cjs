const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');

const html = fs.readFileSync(path.resolve(__dirname, '../Web_App.html'), 'utf8');
const start = html.indexOf('var JOIN_STAGES =');
const end = html.indexOf('function begin(', start);
assert.ok(start >= 0 && end > start, 'registration flow should be present');

const fields = Object.fromEntries(['em', 'jln', 'jfn', 'jmi', 'js', 'jp', 'jy', 'jya', 'jc', 'jprov', 'jcity']
  .map(id => [id, { value: '', focus() { this.focused = true; } }]));
fields.em.value = 'scholar@example.com';
fields.js.options = [{}, {}];
const sections = [0, 1, 2].map(n => ({
  hidden: n !== 0,
  getAttribute() { return String(n); }
}));
let sent;
const elements = {
  ...fields,
  joincount: {}, joinlabel: {}, next: {}, alt: { style: {} },
  joinprogress: { style: { setProperty() {} } }
};
const context = vm.createContext({
  AUTH: { joinStage: 0 }, S: { email: 'scholar@example.com' },
  $: id => elements[id],
  document: {
    querySelectorAll: () => sections,
    querySelector: () => fields.js
  },
  say: message => { context.lastError = message; },
  busy: () => {},
  call: (...args) => { sent = args; return new Promise(() => {}); }
});
vm.runInContext(html.slice(start, end), context);

context.joinNext();
assert.equal(context.AUTH.joinStage, 0, 'empty name must hold on the first step');
assert.equal(context.AUTH.bad, 'f-jln');

Object.assign(fields.jln, { value: 'Rivera' });
Object.assign(fields.jfn, { value: 'Ana' });
context.joinNext();
assert.equal(context.AUTH.joinStage, 1);
assert.equal(sections[1].hidden, false);
assert.equal(sections[0].hidden, true);
assert.equal(elements.alt.textContent, 'Back');

context.joinNext();
assert.equal(context.AUTH.joinStage, 1, 'school is required before proceeding');
assert.equal(context.AUTH.bad, 'f-js');
Object.assign(fields.js, { value: 'USeP' });
Object.assign(fields.jp, { value: 'BS Biology' });
Object.assign(fields.jy, { value: '2' });
context.joinNext();
assert.equal(context.AUTH.joinStage, 2);

Object.assign(fields.jya, { value: '2025' });
Object.assign(fields.jc, { value: 'RA 7687' });
Object.assign(fields.jprov, { value: 'Davao del Sur' });
context.joinNext();
assert.equal(context.AUTH.bad, 'f-jcity', 'city is required on the last step');
assert.equal(sent, undefined);
fields.jcity.value = 'Davao City';
context.joinNext();
assert.equal(sent[0], 'webJoinRequest');
assert.equal(sent[1], 'scholar@example.com');
assert.equal(sent[2], 'Rivera, Ana');
assert.equal(sent[3], 'USeP');
assert.equal(sent[5].city, 'Davao City');

elements.alt.onclick();
assert.equal(context.AUTH.joinStage, 1, 'Back must return to the previous section');
assert.equal(fields.js.value, 'USeP', 'Back must preserve entered values');
console.log('Auth registration flow checks passed');
