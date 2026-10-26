/* Stable row identifiers and an append-only change log. Existing A:M and A:R data stays in place. */
var SCHOLAR_META_COLS_ = ['Scholar ID', 'Status since', 'Updated at', 'Updated by', 'Email verified at']; // N:R
var JOIN_META_COLS_ = ['Request ID', 'Updated at', 'Updated by'];             // S:U
var CHECKIN_META_COLS_ = ['Week due', 'Submitted at'];
var DATA_AUDIT_TAB_ = 'Data change log';
var WEEKLY_ROSTER_TAB_ = 'Weekly roster';

function ensureHeaderBlock_(sh, first, labels) {
  const last = first + labels.length - 1;
  if (sh.getMaxColumns() < last) sh.insertColumnsAfter(sh.getMaxColumns(), last - sh.getMaxColumns());
  const cells = sh.getRange(1, first, 1, labels.length).getValues()[0];
  cells.forEach((v, i) => { if (!String(v || '').trim()) sh.getRange(1, first + i).setValue(labels[i]);
    else if (plainHeading_(v) !== labels[i]) throw new Error(sh.getName() + ' column ' + (first + i) + ' already contains ' + v); });
}
function ensureDataFlowSchema_(ss) {
  const sc = ss.getSheetByName(TAB.scholars);
  if (sc) {
    ensureHeaderBlock_(sc, 14, SCHOLAR_META_COLS_);
    if (sc.getLastRow() > 1) {
      const n = sc.getLastRow() - 1, rows = sc.getRange(2, 1, n, 18).getValues();
      const now = new Date(), meta = rows.map(r => r.slice(13, 18));
      let changed = false, accounts = null;
      if (rows.some(r => String(r[0] || '').trim() && emailOf_(r[4]) && !r[17])) {
        try { accounts = docProps_().getProperties(); } catch (e) { /* fall back to individual lookup */ }
      }
      rows.forEach((r, i) => {
        const email = emailOf_(r[4]);
        if (!String(r[0] || '').trim() || !email) return;
        const m = meta[i];
        if (!m[0]) { m[0] = Utilities.getUuid(); changed = true; }
        // Existing status dates cannot be reconstructed. Keep them blank until the next change.
        if (!m[2]) { m[2] = now; changed = true; }
        if (!m[3]) { m[3] = 'schema migration'; changed = true; }
        if (!m[4]) {
          let account = null;
          if (accounts) { try { account = JSON.parse(accounts[webAcctKey_(email)] || 'null'); } catch (e) { /* invalid old account */ } }
          else account = webAccount_(email);
          if (account && account.h) { m[4] = account.at ? new Date(account.at) : now; changed = true; }
        }
      });
      if (changed) sc.getRange(2, 14, n, 5).setValues(meta);
    }
  }
  const join = ss.getSheetByName(JOIN_TAB);
  if (join) {
    ensureHeaderBlock_(join, 19, JOIN_META_COLS_);
    if (join.getLastRow() > 1) {
      const n = join.getLastRow() - 1, rows = join.getRange(2, 1, n, 19).getValues();
      let changed = false;
      const ids = rows.map(r => {
        if (String(r[3] || '').trim() && !r[18]) { changed = true; return [Utilities.getUuid()]; }
        return [r[18]];
      });
      if (changed) join.getRange(2, 19, n, 1).setValues(ids);
    }
  }
  const ck = ss.getSheetByName(TAB.mcheckins);
  if (ck) ensureCheckinMetadata_(ck);
  dataLogTab_(ss);
  weeklyRosterTab_(ss);
}
function ensureCheckinMetadata_(ck) {
  const h = ck.getRange(1, 1, 1, Math.max(ck.getLastColumn(), 1)).getValues()[0].map(plainHeading_);
  CHECKIN_META_COLS_.forEach(label => {
    if (h.indexOf(label) !== -1) return;
    const at = h.length + 1;
    if (ck.getMaxColumns() < at) ck.insertColumnAfter(ck.getMaxColumns());
    ck.getRange(1, at).setValue(label); h.push(label);
  });
  return { week: h.indexOf(CHECKIN_META_COLS_[0]), submitted: h.indexOf(CHECKIN_META_COLS_[1]), width: h.length };
}
function dataLogTab_(ss) {
  let sh = ss.getSheetByName(DATA_AUDIT_TAB_);
  if (!sh) sh = ss.insertSheet(DATA_AUDIT_TAB_);
  ensureHeaderBlock_(sh, 1, ['At', 'Actor', 'Record type', 'Record ID', 'Action', 'Details']);
  sh.setFrozenRows(1);
  return sh;
}
function logDataChange_(ss, actor, type, id, action, details) {
  dataLogTab_(ss).appendRow([new Date(), String(actor || ''), type, String(id || ''), action, String(details || '').slice(0, 500)]);
}
function markScholarChange_(ss, sh, row, actor, action, details, oldStatus, newStatus) {
  ensureHeaderBlock_(sh, 14, SCHOLAR_META_COLS_);
  const idCell = sh.getRange(row, 14), id = String(idCell.getValue() || Utilities.getUuid());
  idCell.setValue(id);
  if (oldStatus !== newStatus) sh.getRange(row, 15).setValue(new Date());
  sh.getRange(row, 16, 1, 2).setValues([[new Date(), actor]]);
  logDataChange_(ss, actor, 'Scholar', id, action, details);
  return id;
}
function markVerifiedEmail_(ss, email) {
  const sh = ss.getSheetByName(TAB.scholars);
  if (!sh || sh.getLastRow() < 2) return;
  ensureHeaderBlock_(sh, 14, SCHOLAR_META_COLS_);
  const rows = sh.getRange(2, 5, sh.getLastRow() - 1, 1).getValues();
  const i = rows.findIndex(r => emailOf_(r[0]) === emailOf_(email));
  if (i < 0) return;
  const cell = sh.getRange(i + 2, 18);
  if (!cell.getValue()) cell.setValue(new Date());
}
function markJoinChange_(ss, sh, row, actor, action, details) {
  ensureHeaderBlock_(sh, 19, JOIN_META_COLS_);
  const idCell = sh.getRange(row, 19), id = String(idCell.getValue() || Utilities.getUuid());
  sh.getRange(row, 19, 1, 3).setValues([[id, new Date(), actor]]);
  logDataChange_(ss, actor, 'Join request', id, action, details);
  return id;
}
function weeklyRosterTab_(ss) {
  let sh = ss.getSheetByName(WEEKLY_ROSTER_TAB_);
  if (!sh) sh = ss.insertSheet(WEEKLY_ROSTER_TAB_);
  ensureHeaderBlock_(sh, 1, ['Week due', 'Scholar ID', 'Email', 'School', 'Captured at']);
  sh.setFrozenRows(1);
  if (sh.getRange('A1').getNumberFormat() !== '@') sh.getRange('A:A').setNumberFormat('@');
  return sh;
}
// Capture the denominator once per week. Historical weeks before this migration remain estimates.
function captureWeeklyRoster_(ss, now) {
  const sh = weeklyRosterTab_(ss), key = memberWeekKey_(now || new Date());
  if (sh.getLastRow() > 1 && sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().some(r => r[0] === key)) return 0;
  const sc = ss.getSheetByName(TAB.scholars);
  if (!sc || sc.getLastRow() < 2) return 0;
  const rows = sc.getRange(2, 1, sc.getLastRow() - 1, Math.min(sc.getMaxColumns(), 17)).getValues();
  const stamp = new Date(), values = rows.filter(r => String(r[5] || '').trim() === 'Active' && emailOf_(r[4]))
    .map(r => [key, String(r[13] || ''), emailOf_(r[4]), String(r[1] || '').trim(), stamp]);
  // A zero-member week still needs a marker so future reads do not reinterpret it.
  if (!values.length) values.push([key, '', '', '', stamp]);
  sh.getRange(sh.getLastRow() + 1, 1, values.length, 5).setValues(values);
  return values.length;
}

// RR-only data quality queue. Read live cells so it also catches direct spreadsheet edits.
function dataFlowQuality_(ss) {
  const issues = [], schoolNames = {};
  const schools = ss.getSheetByName(TAB.schools);
  if (schools && schools.getLastRow() > 1) {
    const n = schools.getLastRow() - 1, names = schools.getRange(2, 1, n, 1).getDisplayValues();
    const formulas = schools.getRange(2, 3, n, 1).getFormulas();
    names.forEach((r, i) => { if (!r[0]) return; schoolNames[r[0]] = true;
      if (!formulas[i][0]) issues.push('Schools row ' + (i + 2) + ' (' + r[0] + '): active-member count formula is missing. Run Setup to restore it.'); });
  }
  const sc = ss.getSheetByName(TAB.scholars), seen = {};
  if (sc && sc.getLastRow() > 1) sc.getRange(2, 1, sc.getLastRow() - 1, 12).getDisplayValues().forEach((r, i) => {
    if (!r[0]) return;
    const missing = [];
    if (!schoolNames[r[1]]) missing.push('school not in Schools');
    if (!r[2]) missing.push('program');
    if (!r[3]) missing.push('award year');
    const email = emailOf_(r[4]); if (!email) missing.push('email');
    if (!r[5]) missing.push('status');
    if (!r[6]) missing.push('year level');
    if (!r[11]) missing.push('scholarship type');
    if (email) { if (seen[email]) missing.push('duplicate email (row ' + seen[email] + ')'); else seen[email] = i + 2; }
    if (missing.length) issues.push('Scholars row ' + (i + 2) + ' (' + r[0] + '): ' + missing.join(', ') + '.');
  });
  const ev = ss.getSheetByName(TAB.events);
  if (ev && ev.getLastRow() > 1 && ev.getLastColumn() >= 12) {
    const errors = ev.getRange(2, 12, ev.getLastRow() - 1, 1).getDisplayValues().filter(r => /permission|authoriz|error|failed/i.test(r[0]));
    if (errors.length) issues.push(errors.length + ' event(s) failed Google Calendar sync. The Apps Script owner must authorize Calendar access, then retry calendar setup.');
  }
  return { count: issues.length, issues: issues.slice(0, 40) };
}
