const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const c = vm.createContext({ console });
for (const file of fs.readdirSync(root).filter(name => name.endsWith('.gs')).sort())
  vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), c, { filename: file });

assert.equal(c.REGION_SCHOOLS.XI.filter(s => s[0] === 'UM Main').length, 1);
assert.equal(c.REGION_SCHOOLS.XI.filter(s => s[0] === 'UM').length, 0);
assert.equal(c.REGION_SCHOOLS.XI.find(s => s[0] === 'UIC')[2], 'University of the Immaculate Conception');

const cells = Array.from({ length: c.CONFIG.maxSchools }, () => Array(c.LINK_COL).fill(''));
cells[0][0] = 'UM'; cells[0][3] = 'Same Lead'; cells[0][1] = 'Davao City';
cells[1][0] = 'UM Main'; cells[1][3] = 'Same Lead'; cells[1][8] = 'The University of Mindanao';
cells[2][0] = 'UIC'; cells[2][8] = 'University of the Immaculate Conception';
const sheet = {
  getRange(row, col, count = 1, width = 1) {
    return {
      getValues: () => cells.slice(row - 2, row - 2 + count).map(r => r.slice(col - 1, col - 1 + width)),
      setValue(v) { cells[row - 2][col - 1] = v; return this; },
      clearContent() { for (let i = 0; i < width; i++) cells[row - 2][col - 1 + i] = ''; return this; }
    };
  }
};
const ss = { getSheetByName: name => name === c.TAB.schools ? sheet : null };
const renamed = [];
c.renameSchool_ = (book, from, to) => renamed.push([from, to]);
assert.equal(c.mergeUmMainSchool_(ss), true);
assert.equal(cells[0][0], '');
assert.equal(cells[1][0], 'UM Main');
assert.equal(cells[1][1], 'Davao City');
assert.equal(cells[2][0], 'UIC');
assert.deepEqual(renamed, [['UM', 'UM Main']]);
assert.equal(c.mergeUmMainSchool_(ss), false, 'merge is safe to repeat');

cells[0][0] = 'UM'; cells[0][3] = 'Different Lead';
assert.throws(() => c.mergeUmMainSchool_(ss), /different Campus Leads/);
assert.equal(cells[0][0], 'UM', 'conflicting lead retains the old row');

console.log('UM Main migration checks passed');
