const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const html = fs.readFileSync(path.join(__dirname, '..', 'Web_App.html'), 'utf8');
const start = html.indexOf('var SCHOOL_LOGO_IDS = ');
const end = html.indexOf('function schoolLine(', start);
assert.ok(start >= 0 && end > start);
const context = vm.createContext({
  localStorage: { getItem: () => null },
  esc: value => String(value)
});
vm.runInContext(html.slice(start, end), context);

const imageId = name => {
  const markup = context.schoolBadge(name, 40);
  const match = markup.match(/thumbnail\?id=([^&"]+)/);
  assert.ok(match, `${name} should use a supplied logo`);
  return match[1];
};
const main = imageId('UM');
assert.equal(imageId('UM Main'), main);
assert.equal(imageId('The University of Mindanao'), main);
assert.equal(imageId('University of Mindanao — UM Main'), main);
assert.notEqual(imageId('UM Digos'), main);
assert.notEqual(imageId('UM Tagum'), main);
assert.notEqual(imageId('UM Digos'), imageId('UM Tagum'));
assert.equal(imageId('Ateneo de Davao University'), imageId('AdDU'));
assert.equal(imageId('USeP Obrero'), imageId('USeP'));
assert.equal(imageId('University of the Philippines Mindanao'), imageId('UP Mindanao'));
assert.match(context.schoolBadge('UM Main', 40), /alt="UM Main logo"/);
console.log('School badge logo checks passed');
