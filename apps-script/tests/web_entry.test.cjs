const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const files = fs.readdirSync(root).filter(name => name.endsWith('.gs') && name !== '07_Look.gs').sort();
let servedHtml = '';
const context = vm.createContext({
  console,
  Logger: { log() {} },
  ScriptApp: { getService: () => ({ getUrl: () => '' }) },
  HtmlService: {
    createTemplateFromFile(name) {
      if (name !== 'Web_App') throw new Error('Email template unavailable');
      return { getRawContent: () => fs.readFileSync(path.join(root, 'Web_App.html'), 'utf8') };
    },
    createHtmlOutput(html) {
      servedHtml = html;
      return {
        setTitle() { return this; },
        addMetaTag() { return this; },
        setXFrameOptionsMode() { return this; },
      };
    },
    XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' },
  },
});
for (const file of files) {
  vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), context, { filename: file });
}

// Look artwork is optional for opening the page and rendering a plain email.
context.webSheet_ = () => { throw new Error('No sheet in this isolated test'); };
context.webPics_ = () => ({});
context.docProps_ = () => ({ getProperty: () => '' });
assert.throws(() => context.requireLook_(), /07_Look\.gs is empty or incomplete/);
assert.ok(context.doGet({ parameter: {} }));
assert.match(servedHtml, /var START_BOOT = \{/);
assert.doesNotMatch(servedHtml, /__START_BOOT__/);
context.webSheet_ = () => ({});
context.getSettings_ = () => ({ region: '</script><script>alert(1)</script>' });
assert.ok(context.doGet({ parameter: {} }));
assert.match(servedHtml, /\\u003c\/script>/);
const client = servedHtml.slice(servedHtml.indexOf('<script>') + 8, servedHtml.indexOf('</script>', servedHtml.indexOf('<script>')));
new vm.Script(client, { filename: 'served Web_App.html' });
assert.ok(context.doGet({ parameter: { app: 'https://feliciety02.github.io/START-Davao-Chapter/' } }));
assert.match(servedHtml, /"appOrigin":"https:\/\/feliciety02\.github\.io"/);
assert.match(context.renderEmail_({ greeting: 'Hello' }), /Hello/);

// Every direct client RPC must have a corresponding Apps Script function.
const html = fs.readFileSync(path.join(root, 'Web_App.html'), 'utf8');
const rpcs = [...html.matchAll(/\bcall\('([A-Za-z][A-Za-z0-9_]*)'/g)].map(match => match[1]);
const missing = [...new Set(rpcs)].filter(rpc => typeof context[rpc] !== 'function');
assert.deepEqual(missing, ['webLogos'], 'Unexpected missing server functions outside 07_Look.gs');

console.log('Web entry and RPC checks passed (webLogos awaits 07_Look.gs)');
