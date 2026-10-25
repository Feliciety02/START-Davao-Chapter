const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '../..');
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const script = html.slice(html.indexOf('<script>') + 8, html.indexOf('</script>'));
new vm.Script(script, { filename: 'index.html' });
new vm.Script(fs.readFileSync(path.join(root, 'config.js'), 'utf8'), { filename: 'config.js' });
assert.match(html, /\$\('splashstar'\)\.innerHTML = star\(84, 'bob'\)/);
assert.match(html, /<div class="brand">START<small>Region 11<\/small><\/div>/);
assert.match(html, /return '<svg class="star '/);

const web = fs.readFileSync(path.join(root, 'apps-script/Web_App.html'), 'utf8');
assert.match(web, /logo-s\.7f4f7462\.png/);
assert.match(web, /class="region-pill"/);
assert.doesNotMatch(web, /function starMascot\(/);

const manifest = JSON.parse(fs.readFileSync(path.join(root, 'manifest.webmanifest'), 'utf8'));
for (const icon of manifest.icons) {
  const name = icon.src.split('?')[0];
  const png = fs.readFileSync(path.join(root, name));
  assert.equal(png.subarray(1, 4).toString(), 'PNG');
  const side = Number(icon.sizes.split('x')[0]);
  assert.equal(png.readUInt32BE(16), side);
  assert.equal(png.readUInt32BE(20), side);
}
console.log('App logo and icon checks passed');
