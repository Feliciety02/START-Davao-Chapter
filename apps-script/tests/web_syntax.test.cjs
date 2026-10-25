const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');

const html = fs.readFileSync(path.resolve(__dirname, '../Web_App.html'), 'utf8');
assert.doesNotMatch(html, /<\?/, 'Web_App.html must remain valid editor JavaScript');
const start = html.indexOf('<script>') + '<script>'.length;
const end = html.indexOf('</script>', start);
if (start < '<script>'.length || end < 0) throw new Error('Web app script block not found');
const script = html.slice(start, end).replace(/<\?!=[\s\S]*?\?>/g, 'null');
new vm.Script(script, { filename: 'Web_App.html' });

// Registration choices must not overwrite the fallback postcard scenes.
assert.match(script, /JOIN_PLACES\s*=\s*r\.places\s*\|\|\s*\{\}/);
const artStart = script.indexOf('function eagle(');
const artEnd = script.indexOf('function startLogo(', artStart);
assert.ok(artStart >= 0 && artEnd > artStart);
const art = vm.createContext({ PLACE: null, JOIN_PLACES: { Davao: ['Davao City'] } });
vm.runInContext(script.slice(artStart, artEnd), art);
assert.match(art.scene(), /<svg class="scene"/);
console.log('Web app script syntax OK');
