const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const crypto = require('node:crypto');

const root = path.resolve(__dirname, '..');
const c = vm.createContext({ Date, console });
for (const file of fs.readdirSync(root).filter(name => name.endsWith('.gs')).sort())
  vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), c, { filename: file });

let secret = '';
c.docProps_ = () => ({ getProperty: () => secret, setProperty: (_, value) => { secret = value; } });
c.Utilities = { DigestAlgorithm: { SHA_256: 'sha256' }, getUuid: () => 'private-random-secret',
  computeDigest: (_kind, input) => [...crypto.createHash('sha256').update(input).digest()] };
const firstSeed = [...c.miniSeed_('ana@example.com', '2026-10-25')];
assert.equal(secret, 'private-random-secret');
assert.deepEqual([...c.miniSeed_('ana@example.com', '2026-10-25')], firstSeed);
secret = 'another-private-secret';
assert.notDeepEqual([...c.miniSeed_('ana@example.com', '2026-10-25')], firstSeed,
  'a public email and date do not determine the game without the private tracker secret');

const game = {
  id: 'today-abc', kind: 'odd', xp: 10, each: 3, perfect: 1, secs: 20,
  rounds: [
    { title: 'First', choices: ['A', 'B', 'C', 'D'], odd: 2, why: 'First answer' },
    { title: 'Second', choices: ['E', 'F', 'G', 'H'], odd: 1, why: 'Second answer' },
    { title: 'Third', choices: ['I', 'J', 'K', 'L'], odd: 3, why: 'Third answer' }
  ]
};
const publicGame = c.miniPublic_(game, 0, new Date());
assert.equal(publicGame.rounds[0].title, 'First');
assert.equal(publicGame.rounds[1], null);
assert.doesNotMatch(JSON.stringify(publicGame), /"odd"\s*:|First answer|Second answer|Third answer/);

const day = '2026-10-25';
let row = [day, 'ana@example.com', 'Ana', 'USeP', 'odd', 'Started', 0, 0,
  new Date(), '', game.id, 0, 0, new Date(), 'prompt'];
const sh = { getRange(_row, col, _height, width) {
  return {
    getValues: () => [row.slice(col - 1, col - 1 + width)],
    setValues: values => { values[0].forEach((value, i) => { row[col - 1 + i] = value; }); }
  };
} };
c.webSafe_ = fn => fn();
c.webSession_ = () => ({ e: 'ana@example.com' });
c.webSheet_ = () => ({});
c.webRoles_ = () => ({ email: 'ana@example.com', member: { name: 'Ana' } });
c.dailyTab_ = () => sh;
c.dailyRowFor_ = () => 2;
c.docLock_ = () => ({ tryLock: () => true, releaseLock() {} });
c.webMiniFor_ = () => game;
c.miniProgram_ = () => '';
c.dayKey_ = () => day;
c.webForget_ = () => {};
c.webMember_ = () => ({ name: 'Ana' });

row[13] = new Date(Date.now() - 21000);
const expired = c.webDailyPick('token', { day, id: game.id, round: 0, pick: 2 });
assert.equal(expired.ok, true);
assert.equal(expired.timeUp, true);
assert.equal(expired.correct, false);
assert.equal(row[12], 0, 'a late correct answer earns nothing');
assert.equal(c.webDailyPick('token', { day, id: game.id, round: 0, pick: 2 }).correct, false, 'the same pick can retry after a lost reply');
assert.equal(c.webDailyPick('token', { day, id: game.id, round: 0, pick: 1 }).ok, false, 'a different second pick cannot score');

row[11] = 1; row[13] = new Date(); row[14] = 'prompt';
const correct = c.webDailyPick('token', { day, id: game.id, round: 1, pick: 1 });
assert.equal(correct.correct, true);
assert.equal(row[12], 1);
const next = c.webDailyNext('token', { day, id: game.id, round: 1 });
assert.equal(next.prompt.title, 'Third');
assert.equal(next.prompt.odd, undefined);
assert.equal(row[11], 2);
assert.equal(c.webDailyNext('token', { day, id: game.id, round: 1 }).prompt.title, 'Third', 'a lost next-round reply can retry');

row[12] = 2; row[14] = 'answered';
const finished = c.webDailyFinish('token', { day, id: game.id, picks: ['C', 'F', 'L'], answer: 'anything' });
assert.equal(finished.xp, 6, 'XP comes from server recorded results, not supplied answers');
assert.equal(finished.stars, 2);

console.log('Daily game answer secrecy and server timer checks passed');
