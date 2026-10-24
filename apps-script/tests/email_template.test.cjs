const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const source = fs.readFileSync(path.resolve(__dirname, '../Email_Reminder.html'), 'utf8');
for (const style of source.matchAll(/style="([^"]*)"/gs)) {
  assert.doesNotMatch(style[1], /<\?/, 'Template code inside CSS produces editor errors');
}
let code = 'let output = ""; with (data) {\n';
let pos = 0;
for (const match of source.matchAll(/<\?(!?=)?([\s\S]*?)\?>/g)) {
  code += `output += ${JSON.stringify(source.slice(pos, match.index))};\n`;
  code += match[1] ? `output += String(${match[2].trim()});\n` : `${match[2]}\n`;
  pos = match.index + match[0].length;
}
code += `output += ${JSON.stringify(source.slice(pos))};\n} return output;`;
const render = new Function('data', code);
const html = render({
  bg: '', hero: '', bar: '', sparkle: '', star: '', logo: 'https://example.com/logo.png', logoW: 375, logoH: 120,
  badge: 'Check-in', badgeColor: '#B3261E', title1: 'Weekly', title2: 'Check-in',
  schoolLabel: 'School', school: 'USeP', weekLabel: 'Week', week: 'Oct 19', deadlineLabel: 'Due', deadline: 'Oct 25',
  greeting: 'Hello', intro: ['Please check in.'], extraHtml: '', outro: ['Thank you.'],
  buttonText: 'Open START', link: 'https://example.com/start', buttonNote: '', portal: '',
  signature: ['START Team'], footer: 'Sent by START', region: 'Region XI',
});
assert.doesNotMatch(html, /<\?[^>]*\?>/);
assert.match(html, /chip-count-3/);
assert.match(html, /<font color="#B3261E">Oct 25<\/font>/);
assert.match(html, /width="375" height="120"/);
assert.match(html, /https:\/\/example\.com\/start/);
console.log('Email template render check passed');
