/* =====================================================================
 *  START SCHOLAR TRACKER  ·  05_Portals
 *  ---------------------------------------------------------------------
 *  Inside: The website (members, Leads, RR) with passwords. The old Campus Lead portal code stays but is off.
 *  To customise: Website pages and the password rules (WEB_PW).
 *
 *  The code has 9 files + Email_Reminder and Web_App (HTML). Keep these exact names:
 *    00_Config, 01_Region, 02_Menu_Setup, 03_Forms, 04_Emails, 05_Portals, 06_Calendar, 07_Look, 08_Tools
 *  To update: open each file, select all (Ctrl+A), paste the new version, Save.
 *  Then in the Google Sheet: START Tracker › ① Set up / fix the tracker.
 *  You never need to run anything from here: use the START Tracker menu in the Sheet.
 * ===================================================================== */

// The ▶ Run button runs the FIRST function of the open file: this one, so it is always safe.
function about_05_Portals() {
  Logger.log('START Tracker' + (typeof START_MODULES === 'string' ? ' ' + START_MODULES : '') + ', file 05_Portals. Nothing to run here: open the Google Sheet and use the START Tracker menu.');
}


/* ---------- sheet-only functions are never run from the website ----------
 * The website runs as you, and a web page can ask for any function whose name has no "_" at the end. Every tool that
 * belongs to the sheet (send an announcement, reset a password, remove a person, upload pictures, fresh start…) starts
 * with this check: it runs only for you or another editor of the tracker, never for a website visitor. */
function sheetUser_() {
  if (typeof Session === 'undefined' || typeof Session.getActiveUser !== 'function') return '';   // test stand-ins only (Apps Script always has it)
  let u = '';
  try { u = emailOf_(Session.getActiveUser().getEmail()); } catch (e) { u = ''; }
  if (u) {
    if (u === me_() || u === automationOwner_()) return u;
    try { if (SpreadsheetApp.getActive().getEditors().some(x => emailOf_(x.getEmail()) === u)) return u; } catch (e) { /* not in the sheet */ }
  }
  throw new Error('This works only from the START Tracker sheet, for its editors.');
}

// ━━━━━━━━━━━━━━━━━━━━  FocalPortal  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  🧑‍🏫 START CAMPUS LEADS
 *  1. "START Campus Leads" tab (for the RR): every START Campus Lead's progress in one table.
 *  2. START Campus Lead PORTALS: each START Campus Lead gets their OWN private Google Sheet,
 *     shared ONLY with their email (typed in Scholars › Contact). It shows only their school.
 *     Why a separate file: in Google Sheets anyone who can open a file can see every
 *     tab in it, so per-person tabs inside this tracker would expose all scholars' data.
 *  3. Connected names: Schools > START Campus Lead is a dropdown of Scholars names; picking one
 *     fills the email from Scholars, and the Scholars tab highlights START Campus Leads.
 * ================================================================ */

var FOCAL_HEAD = ['School', 'START Campus Lead', 'Email', 'This week', 'Check-ins this term', 'Streak (weeks)',
                  'Last check-in', 'Portal', 'Portal status', 'Portal updated'];
var PORTAL_FOLDER = 'START Campus Lead Portals';

/* ---------- data: one record per school ---------- */

function focalData_(ss) {
  const now = new Date();
  const weekStart = mondayOf_(now);
  const due = reminderWeek_(now).deadline;
  const ts = termStart_(ss);
  const termStart = ts instanceof Date ? ts : null;

  const schools = schoolRows_(ss).map(s => Object.assign(s, { checkins: [], scholars: [] }));
  const by = {};
  schools.forEach(s => { by[s.name] = s; });

  const ck = ss.getSheetByName(TAB.checkins);
  if (ck && ck.getLastRow() > 1) {
    const c = ckCols_(ck);
    ck.getRange(2, 1, ck.getLastRow() - 1, c.width).getValues().forEach(r => {
      const s = by[String(r[c.school]).trim()];
      if (!s || !(r[0] instanceof Date)) return;
      s.checkins.push({ date: r[0], by: c.name >= 0 ? r[c.name] : '', events: c.events >= 0 ? r[c.events] : '',
                        dc: c.dc >= 0 ? r[c.dc] : '', help: c.help >= 0 ? r[c.help] : '',
                        concern: c.concern >= 0 ? r[c.concern] : '', vibe: c.vibe >= 0 ? r[c.vibe] : '' });
    });
  }
  const sc = ss.getSheetByName(TAB.scholars);
  if (sc && sc.getLastRow() > 1) {
    sc.getRange(2, 1, sc.getLastRow() - 1, SCHOLAR_COLS).getValues().forEach(r => {
      const s = by[String(r[1]).trim()];
      if (s && String(r[0]).trim()) s.scholars.push({ name: r[0], program: r[2], awarded: r[3], status: r[5],
                                                       level: r[6], campus: r[7] });
    });
  }

  schools.forEach(s => {
    s.checkins.sort((a, b) => b.date - a.date);
    s.last = s.checkins.length ? s.checkins[0].date : null;
    s.thisWeek = s.checkins.some(x => x.date >= weekStart);
    s.term = termStart ? s.checkins.filter(x => x.date >= termStart).length : s.checkins.length;
    // streak: consecutive weeks with a check-in (this week only counts once it is in)
    const weeks = {};
    s.checkins.forEach(x => { weeks[mondayOf_(x.date).getTime()] = true; });
    let w = new Date(weekStart);
    if (!s.thisWeek) w.setDate(w.getDate() - 7);
    let streak = 0;
    while (weeks[w.getTime()]) { streak++; w.setDate(w.getDate() - 7); }
    s.streak = streak;
    s.active = s.scholars.filter(x => ['Graduated', 'Dropped'].indexOf(String(x.status)) === -1).length;
    s.due = due;
  });
  return schools;
}

/* ---------- store of portal files (only valid for THIS tracker) ---------- */

function portalStore_() {
  let s = {};
  try { s = JSON.parse(docProps_().getProperty('PORTALS') || '{}'); } catch (e) { s = {}; }
  const id = SpreadsheetApp.getActive().getId();
  if (s.master !== id) s = { master: id, map: {} };     // copied tracker: never reuse the original's portals
  s.map = s.map || {};
  return s;
}
// A school's portal link, only while it is shared with the CURRENT START Campus Lead's email ('' otherwise)
function portalUrl_(store, s) {
  if (DB_MODE) return '';                                   // database mode: not used (see 00_Config)
  const p = store.map[s.name];
  return p && p.id && p.email && p.email === s.email ? `https://docs.google.com/spreadsheets/d/${p.id}/edit` : '';
}

// One name for the portal everywhere (file, emails, links)
var PORTAL_NAME = 'START Campus Lead Portal';
function portalTitle_(school) {      // the region is in the name: a Lead can tell which RR's portal it is
  let reg = ''; try { reg = regionProfile_().label; } catch (e) { reg = ''; }
  return PORTAL_NAME + ' – ' + school + (reg ? ' · ' + reg : '');
}

function savePortalStore_(s) { docProps_().setProperty('PORTALS', JSON.stringify(s)); }

/* ---------- the "START Campus Leads" tab (RR view) ---------- */

function buildFocalTab_(ss, data, store) {
  if (DB_MODE) return;                                   // database mode: not used (see 00_Config)
  data = data || focalData_(ss);
  store = store || portalStore_();
  const sh = getOrCreate_(ss, TAB.focal);
  const n = Math.max(data.length, 1), W = FOCAL_HEAD.length;
  if (sh.getMaxRows() < CONFIG.maxSchools + 2) sh.insertRowsAfter(sh.getMaxRows(), CONFIG.maxSchools + 2 - sh.getMaxRows());
  if (sh.getMaxColumns() < W) sh.insertColumnsAfter(sh.getMaxColumns(), W - sh.getMaxColumns());
  sh.getRange(1, 1, sh.getMaxRows(), W).clearContent().setBackground(null);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);

  const rows = data.map(s => {
    const p = store.map[s.name] || {};
    let status = '—';
    if (!s.focal) status = 'No START Campus Lead yet';
    else if (!s.email) status = 'Add an email in Scholars (Contact)';
    else if (p.id && p.email === s.email) status = 'Shared with ' + s.email;
    else status = 'Not created yet (🛠 Setup tools › 🔐 Create / update Campus Lead portals)';
    return [s.name, s.focal || '— not assigned', s.email || '',
            !s.focal ? '—' : (s.thisWeek ? '✅ Submitted' : '⏳ Not yet'),
            s.term, s.streak ? '🔥 ' + s.streak : '0', s.last || '',
            '', status, p.updated ? new Date(p.updated) : ''];
  });
  if (rows.length) {
    sh.getRange(2, 1, rows.length, W).setValues(rows);
    rows.forEach((r, i) => {
      const p = store.map[data[i].name];
      if (p && p.id) sh.getRange(2 + i, 8).setFormula(`=HYPERLINK("https://docs.google.com/spreadsheets/d/${p.id}/edit","Open portal ↗")`);
      // (the RR, as owner, can always open it; START Campus Leads get portalUrl_ only while it is shared with them)
    });
  }
  bodyRows_(sh.getRange(2, 1, n, W));
  sh.getRange(2, 1, n, 1).setFontWeight('bold');
  sh.getRange(2, 2, n, 1).setFontWeight('bold').setFontColor(THEME.navy);
  sh.getRange(2, 3, n, 1).setFontColor(THEME.muted);
  sh.getRange(2, 4, n, 4).setHorizontalAlignment('center');
  sh.getRange(2, 5, n, 1).setFontWeight('bold').setFontColor(THEME.blue);
  sh.getRange(2, 7, n, 1).setNumberFormat('mmm d');
  sh.getRange(2, 8, n, 1).setHorizontalAlignment('center');
  sh.getRange(2, 9, n, 1).setFontColor(THEME.muted).setFontSize(9);
  sh.getRange(2, 10, n, 1).setNumberFormat('mmm d, h:mm AM/PM').setFontColor(THEME.muted).setFontSize(9);
  markColumns_(sh, FOCAL_HEAD.map((h, i) => [i + 1, 'auto', h, '', h === 'Email' ? 'link' : h === 'Portal status' ? true : undefined]), n, false);   // whole tab is automatic; fits widths

  const D = sh.getRange(2, 4, n, 1), B = sh.getRange(2, 2, n, 1), I = sh.getRange(2, 9, n, 1);
  sh.setConditionalFormatRules([
    chipRule_(D, '✅', CHIP.green, false), chipRule_(D, '⏳', CHIP.yellow, false),
    chipRule_(B, '— not assigned', CHIP.red),
    chipRule_(I, 'Shared', CHIP.green, false), chipRule_(I, 'Add an email', CHIP.yellow, false),
    chipRule_(I, 'No START Campus Lead', CHIP.red, false), chipRule_(I, 'Not created', CHIP.blue, false),
  ]);
  sh.getRange('A1').setNote('This tab refreshes itself (① Set up / fix the tracker, daily portal update, and after each check-in).\n' +
    'Names come from Schools › START Campus Lead. Change them there, not here.');
  sh.setHiddenGridlines(true);
  sh.setTabColor(THEME.tabs.focal);
  placeTutorial_(sh);
  placeTabBanner_(sh, W);
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).filter(p => p.getDescription() === PROTECT_TAG).forEach(p => p.remove());
  sh.protect().setDescription(PROTECT_TAG).setWarningOnly(true);
  try { docProps_().setProperty('FOCAL_SYNC', focalKey_(ss)); } catch (e) {}
}

/* ---------- portals ---------- */

function refreshPortalsMenu() {
  sheetUser_();                                   // sheet only: never from the website
  requireSetup_();
  if (!ensureOwner_('create and update the portals')) return;
  const r = refreshPortals_();
  installPortalTrigger_();
  notify_(`🧑‍🏫 Portals updated\n\n` +
    `• ${r.ok} up to date\n` +
    (r.newShares.length ? `• New (welcome sent): ${r.newShares.join(', ')}\n` : '') +
    (r.replaced.length ? `• New email, new portal: ${r.replaced.join(', ')}\n` : '') +
    (r.noEmail.length ? `• No email in Scholars: ${r.noEmail.join(', ')}\n` : '') +
    (r.failed.length ? `• Not shared: ${r.failed.join('; ')}\n` : '') +
    (r.later.length ? `• ⏳ Not done yet: ${r.later.join(', ')} (click again)\n` : '') +
    `\nPrivate to each Lead. Updates every morning and after each check-in.`);
  SpreadsheetApp.getActive().setActiveSheet(SpreadsheetApp.getActive().getSheetByName(TAB.focal));
}

function refreshPortalsAuto(e) { if (!ownerGuard_(e)) return; try { refreshPortals_(); } catch (e) { Logger.log('Portal refresh: ' + e.message); } }

function installPortalTrigger_() {
  if (DB_MODE) { removeTriggers_(['refreshPortalsAuto']); return; }                                   // database mode: not used (see 00_Config)
  docProps_().setProperty('PORTALS_ON', '1');
  if (ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'refreshPortalsAuto')) return;
  ScriptApp.newTrigger('refreshPortalsAuto').timeBased().everyDays(1).atHour(6).inTimezone(CONFIG.timezone).create();
}

function stopPortalUpdates() {
  sheetUser_();                                   // sheet only: never from the website
  docProps_().deleteProperty('PORTALS_ON');
  const n = removeTriggers_(['refreshPortalsAuto']);
  notify_(n ? 'Daily portal updates are OFF. Portals stay shared but will no longer refresh by themselves.'
            : 'Daily portal updates were not on.');
}

// onlySchool: refresh just one school's portal (used right after a check-in)
function refreshPortals_(onlySchool) {
  if (DB_MODE) return { ok: 0, newShares: [], noEmail: [], failed: [], replaced: [], later: [] };                                   // database mode: not used (see 00_Config)
  const ss = SpreadsheetApp.getActive();
  const data = focalData_(ss);
  const store = portalStore_();
  const st = getSettings_();
  const rr = { name: st.rrName || 'your Regional Representative', email: Session.getEffectiveUser().getEmail(),
               region: st.region || '' };
  const res = { ok: 0, newShares: [], noEmail: [], failed: [], replaced: [], later: [] };
  let folder = null;

  // oldest first, so a run that hits Google's time limit continues with the rest next time
  const age = s => (store.map[s.name] || {}).updated || 0;
  data.slice().sort((a, b) => age(a) - age(b)).forEach(s => {
    if (onlySchool && s.name !== onlySchool) return;
    if (!onlySchool && outOfTime_(60 * 1000)) { if (s.focal && s.email) res.later.push(s.name); return; }
    const p = store.map[s.name] || {};
    // Email changed (e.g. swapped in Scholars) or Campus Lead removed: the old portal is closed for good
    // (old account loses access, file goes to Drive's trash) and a NEW portal is made for the new email below.
    if (p.id && p.email && p.email !== s.email) {
      try { const old = DriveApp.getFileById(p.id); old.removeViewer(p.email); old.setTrashed(true); } catch (e) { /* already gone */ }
      res.replaced.push(s.name);
      p.id = ''; p.email = '';
    }
    if (!s.focal || !s.email) { if (s.focal) res.noEmail.push(s.name); store.map[s.name] = p; return; }
    try {
      let file = null;
      if (p.id) { try { file = DriveApp.getFileById(p.id); if (file.isTrashed()) file = null; } catch (e) { file = null; } }
      if (!file) {
        folder = folder || portalFolder_(ss);
        const created = SpreadsheetApp.create(portalTitle_(s.name));
        file = DriveApp.getFileById(created.getId());
        file.moveTo(folder);
        p.id = created.getId();
      }
      if (file.getName() !== portalTitle_(s.name)) file.setName(portalTitle_(s.name));   // older portals get the current name
      file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);   // never "anyone with the link"
      try {                                                                  // nobody else: only you and this Lead
        const keep = [String(rr.email || '').toLowerCase(), String(s.email || '').toLowerCase()];
        file.getEditors().forEach(u => { if (keep.indexOf(String(u.getEmail()).toLowerCase()) === -1) file.removeEditor(u); });
        file.getViewers().forEach(u => { if (keep.indexOf(String(u.getEmail()).toLowerCase()) === -1) file.removeViewer(u); });
      } catch (e) { Logger.log('Portal sharing tidy-up (' + s.name + '): ' + e.message); }
      if (p.email !== s.email) {
        file.addViewer(s.email);
        p.email = s.email;
        sendPortalWelcome_(s, p.id, rr);
        res.newShares.push(s.name);
      }
      writePortal_(SpreadsheetApp.openById(p.id), s, rr);
      p.updated = Date.now();
      res.ok++;
    } catch (e) {
      res.failed.push(`${s.name} (${e.message})`);
    }
    store.map[s.name] = p;
  });
  savePortalStore_(store);
  buildFocalTab_(ss, data, store);
  return res;
}

function portalFolder_(ss) {
  const props = docProps_();
  const id = props.getProperty('PORTAL_FOLDER');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  const parents = DriveApp.getFileById(ss.getId()).getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const f = parent.createFolder(PORTAL_FOLDER);
  props.setProperty('PORTAL_FOLDER', f.getId());
  return f;
}

/* ---------- what a START Campus Lead sees ---------- */

// The short "how to use this page" text on the right of every portal
// The steps panel on the right of a portal: [text, style] lines (styles as in the tab panels)
function portalGuideLines_(s, rr) {
  const st = getSettings_();
  const members = typeof memberForm_ === 'function' && memberForm_() && st.portalMembers !== false;   // section hidden by the RR: no member steps
  const L = [['📖  HOW TO USE THIS PAGE', 'head']];
  L.push([`🗓  Opens ${String(st.mOpenDay).slice(0, 3)} · due ${st.dueDay}`, 'label']);
  L.push(['1   Answer your check-in (yellow button)', 'step']);
  if (members) {
    L.push(['2   Message ONLY members marked ⏳', 'step']);
    L.push(['3   🆘 Help? Message them today', 'step']);
  }
  L.push(['Everything else updates by itself.', 'note']);
  L.push([`❓  ${rr.name}  ·  ${rr.email}`, 'tip']);
  return L;
}
function portalGuide_(s, rr) { return portalGuideLines_(s, rr).map(l => l[0]).join('\n'); }

// Draws the panel in column col from row 2. Rows that are very tall (the copy-message box) are left as panel space.
function drawPortalGuide_(sh, lines, lastRow, col) {
  col = col || 1;
  const S = SpreadsheetApp.BorderStyle;
  let r = 2;
  const cells = [];
  lines.forEach(l => {
    while (r < lastRow && sh.getRowHeight(r) > 60) { cells.push([r, null]); r++; }
    cells.push([r, l]); r++;
  });
  const first = 2, last = r - 1;
  sh.getRange(first, col, last - first + 1, 1).setFontFamily(THEME.font).setFontSize(10).setFontColor(THEME.navy)
    .setWrap(false).setVerticalAlignment('middle').setHorizontalAlignment('left').setBackground('#FFFFFF');
  cells.forEach(([row, l]) => {
    const c = sh.getRange(row, col);
    if (!l) { c.setBackground(TUTORIAL_LOOK.step); return; }
    c.setValue(l[0]);
    switch (l[1]) {
      case 'head':  header_(c, THEME.navy); c.setHorizontalAlignment('left').setWrap(false); break;
      case 'label': c.setBackground('#EEF5FD').setFontWeight('bold').setFontColor(THEME.blue).setFontSize(9); break;
      case 'step':  c.setBackground(TUTORIAL_LOOK.step).setBorder(null, true, null, null, null, null, THEME.blue, S.SOLID_THICK); break;
      case 'sub':   c.setBackground(TUTORIAL_LOOK.step).setFontColor(THEME.muted).setFontSize(9)
                     .setBorder(null, true, null, null, null, null, THEME.blue, S.SOLID_THICK); break;
      case 'note':  c.setBackground(TUTORIAL_LOOK.from).setFontColor(THEME.muted).setFontSize(9); break;
      case 'tip':   c.setBackground(TUTORIAL_LOOK.tip).setFontWeight('bold').setFontSize(9)
                     .setBorder(null, true, null, null, null, null, THEME.sun, S.SOLID_THICK); break;
    }
  });
  sh.getRange(first, col, last - first + 1, 1).setBorder(true, true, true, true, null, null, TABLE_LINE, S.SOLID);
  return last;
}

// How wide a text is in Poppins (pixels). Google's own "fit to data" under-measures Poppins and emoji,
// so the portal measures text itself (character classes, with a little safety margin).
function textWidth_(t, pt, bold) {
  const em = (pt || 10) * 4 / 3;
  let w = 0;
  for (const ch of String(t)) {
    const c = ch.codePointAt(0);
    w += c > 0x2000 ? 1.35 : /[A-Z]/.test(ch) ? 0.69 : /[a-z]/.test(ch) ? 0.6 : /[0-9]/.test(ch) ? 0.61 : ch === ' ' ? 0.28 : 0.52;
  }
  return Math.ceil(w * em * (bold ? 1.06 : 1) * 1.08);
}

// Column widths from the real values: each column as wide as its longest value (+ padding),
// long free text (concerns, notes) capped and wrapped, small columns kept readable.
// Merged cells (titles, the copy message, the button) are not measured: they simply span the table.
function fitPortalColumns_(sh, c1, n, firstDataRow, lastRow) {
  const PAD = 30, MIN = 110, MAX = 320, MIN_TABLE = 980;
  const rows = lastRow - firstDataRow + 1;
  const block = sh.getRange(firstDataRow, c1, rows, n);
  const merged = block.getMergedRanges().map(m => [m.getRow(), m.getColumn(), m.getNumRows(), m.getNumColumns()]);
  const inMerge = (r, c) => merged.some(m => r >= m[0] && r < m[0] + m[2] && c >= m[1] && c < m[1] + m[3]);
  const vals = block.getDisplayValues(), weights = block.getFontWeights(), sizes = block.getFontSizes();
  const widths = [];
  for (let j = 0; j < n; j++) {
    let w = 0;
    for (let i = 0; i < rows; i++) {
      if (!vals[i][j] || inMerge(firstDataRow + i, c1 + j)) continue;
      String(vals[i][j]).split('\n').forEach(line => { w = Math.max(w, textWidth_(line, sizes[i][j], weights[i][j] === 'bold')); });
    }
    widths.push(Math.max(w + PAD, MIN));
  }
  block.setWrap(false);
  widths.forEach((w, j) => {
    if (w > MAX) { widths[j] = MAX; sh.getRange(firstDataRow, c1 + j, rows, 1).setWrap(true); }
  });
  merged.forEach(m => sh.getRange(m[0], m[1], m[2], m[3]).setWrap(true));   // titles, copy message, notes keep wrapping
  // the three cards sit on column pairs: each pair needs room for its big number
  for (let j = 0; j + 1 < n; j += 2) if (widths[j] + widths[j + 1] < 300) widths[j + 1] = 300 - widths[j];
  // a roomy page: a narrow table is widened evenly (looks like the tracker's own tabs)
  const sum = widths.reduce((a, b) => a + b, 0);
  if (sum < MIN_TABLE) { const add = Math.floor((MIN_TABLE - sum) / n); widths.forEach((w, j) => { widths[j] = w + add; }); }
  widths.forEach((w, j) => sh.setColumnWidth(c1 + j, w));
}

function writePortal_(pss, s, rr) {
  pss.setSpreadsheetTimeZone(CONFIG.timezone);
  const sheets = pss.getSheets();
  const sh = sheets[0];
  sheets.slice(1).forEach(x => pss.deleteSheet(x));
  sh.setName('My dashboard');
  sh.getImages().forEach(i => i.remove());
  sh.clear(); sh.clearConditionalFormatRules();
  const W = [170, 150, 150, 150, 150, 210], NC = W.length;
  if (sh.getMaxColumns() > NC) sh.deleteColumns(NC + 1, sh.getMaxColumns() - NC);
  if (sh.getMaxRows() < 80) sh.insertRowsAfter(sh.getMaxRows(), 80 - sh.getMaxRows());
  W.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(1, 1, sh.getMaxRows(), NC).breakApart().setFontFamily(THEME.font).setFontColor(THEME.navy)
    .setVerticalAlignment('middle').setFontSize(10);
  const f = (d, p) => Utilities.formatDate(d, CONFIG.timezone, p);
  const merge = (r, c, w) => sh.getRange(r, c, 1, w).merge();

  // ---------- 1. Hello + 3 cards: the whole week at a glance ----------
  const pst = getSettings_();                                 // this RR's own portal choices (Settings tab)
  const first = (typeof firstName_ === 'function' ? firstName_(s.focal) : String(displayName_(s.focal)).split(' ')[0]) || 'there';
  merge(2, 1, NC).setValue(`Hi ${first}! 👋`).setFontSize(20).setFontWeight('bold');
  merge(3, 1, NC).setValue(`${s.name}${rr.region ? ' · ' + rr.region : ''} · START Campus Lead · updated ${f(new Date(), 'MMM d, h:mm a')}`)
    .setFontColor(THEME.muted).setFontSize(9);
  sh.setRowHeight(2, 40);
  const note = String(pst.portalNote || '').trim();
  if (note) {
    merge(4, 1, NC).setValue('📣  ' + note + (rr.name && rr.name !== 'your Regional Representative' ? '   — ' + rr.name : ''))
      .setFontWeight('bold').setFontColor(THEME.navy).setBackground(THEME.yellow).setWrap(true);
    sh.setRowHeight(4, 34);
  }
  const P = portalMembers_(s, pst);                          // members, answers, what to do (23_Members data)
  const card = (c1, label, value, cap, chip) => {
    merge(5, c1, 2).setValue(label).setFontSize(9).setFontWeight('bold').setFontColor(THEME.blue);
    merge(6, c1, 2).setValue(value).setFontSize(20).setFontWeight('bold').setHorizontalAlignment('left');
    merge(7, c1, 2).setValue(cap).setFontSize(9).setFontColor(THEME.muted);
    sh.getRange(5, c1, 3, 2).setBackground(chip ? chip.bg : THEME.band)
      .setBorder(true, true, true, true, null, null, TABLE_LINE, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(5, c1, 3, 1).setBorder(null, true, null, null, null, null, THEME.blue, SpreadsheetApp.BorderStyle.SOLID_THICK);
  };
  card(1, '📝  YOUR CHECK-IN', s.thisWeek ? '✅ Done' : '⏳ Not yet',
       s.thisWeek ? 'Thank you! See you next week.' : `Due ${f(s.due, 'EEEE, MMM d')}`, s.thisWeek ? CHIP.green : CHIP.yellow);
  if (P.answers) card(3, '👥  MEMBERS ANSWERED', `${P.answered} of ${P.members.length}`,
       P.members.length && P.answered === P.members.length ? 'Everyone is in 🎉' : `${P.members.length - P.answered} still to go`,
       P.members.length && P.answered === P.members.length ? CHIP.green : null);
  else card(3, '🎓  MY SCHOLARS', `${s.active} active`, 'On the START masterlist');
  card(5, '🔥  STREAK', `${s.streak} week${s.streak === 1 ? '' : 's'}`, `${s.term} check-in${s.term === 1 ? '' : 's'} this term`);
  [[5, 22], [6, 42], [7, 24]].forEach(([r, h]) => sh.setRowHeight(r, h));

  // ---------- 2. The one button ----------
  merge(9, 1, NC).setFormula(`=HYPERLINK("${(s.link || '').replace(/"/g, '')}","${s.thisWeek ? '✅  Check-in done · open the form again' : '📝   ANSWER THIS WEEK\'S CHECK-IN   →'}")`)
    .setBackground(s.thisWeek ? CHIP.green.bg : THEME.sun).setFontWeight('bold').setFontSize(s.thisWeek ? 11 : 13).setHorizontalAlignment('center')
    .setBorder(true, true, true, true, null, null, s.thisWeek ? CHIP.green.fg : THEME.blue, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(9, s.thisWeek ? 34 : 44);

  // ---------- 3. To do this week: only what needs the Lead, nothing else ----------
  let r = 11;
  merge(r, 1, NC).setValue('✅  TO DO THIS WEEK'); header_(sh.getRange(r, 1, 1, NC)); r++;
  const todo = [];
  if (!s.thisWeek) todo.push([`Answer your check-in by ${f(s.due, 'EEEE, MMM d')}, 11:59 PM (yellow button above).`, CHIP.yellow]);
  P.help.forEach(m => todo.push([`🆘 ${displayName_(m.name)} asked for help: message them today.`, CHIP.red]));
  if (P.answers && P.toMessage.length) todo.push([`Message the ${P.toMessage.length} member${P.toMessage.length === 1 ? '' : 's'} marked ⏳ below. Copy this:`, CHIP.yellow, P.message]);
  if (P.noEmail.length) todo.push([`${P.noEmail.length} member${P.noEmail.length === 1 ? ' has' : 's have'} no email yet (📭 below): ask them to email ${rr.name} (${rr.email}) the Google email they use.`, CHIP.grey]);
  if (!todo.length) todo.push(['🎉 Nothing to do: you and your members are all in. Salamat!', CHIP.green]);
  todo.forEach(([t, chip, copy]) => {
    merge(r, 1, NC).setValue(t).setWrap(true).setBackground(chip.bg).setFontColor(chip.fg).setFontWeight('bold');
    sh.getRange(r, 1).setBorder(null, true, null, null, null, null, chip.fg, SpreadsheetApp.BorderStyle.SOLID_THICK);
    sh.setRowHeight(r, 28); r++;
    if (copy) {
      merge(r, 1, NC).setValue(copy).setWrap(true).setFontSize(9).setBackground('#FFFBE6').setVerticalAlignment('top')
        .setBorder(true, true, true, true, null, null, '#F5E27A', SpreadsheetApp.BorderStyle.DASHED);
      sh.setRowHeight(r, Math.min(150, 18 * (String(copy).split('\n').length + 1))); r++;
    }
  });
  r++;

  // ---------- 4. ONE table of members (answers + who they are), most urgent first ----------
  const firstDataRow = r + 1;
  if (P.cols.length) {
    merge(r, 1, NC).setValue(`👥  MY MEMBERS  ·  ${P.members.length}` + (P.answers ? `  ·  ${P.answered} answered this week (due ${f(P.due, 'EEE, MMM d')})` : ''));
    header_(sh.getRange(r, 1, 1, NC)); r++;
    const last = P.cols.length - 1;
    sh.getRange(r, 1, 1, P.cols.length).setValues([P.cols.map(c => c[0])])
      .setFontWeight('bold').setFontSize(9).setFontColor(THEME.muted).setBackground(THEME.band);
    if (P.cols.length < NC) merge(r, P.cols.length, NC - P.cols.length + 1);
    r++;
    if (sh.getMaxRows() < r + P.rows.length + 8) sh.insertRowsAfter(sh.getMaxRows(), r + P.rows.length + 8 - sh.getMaxRows());
    if (P.rows.length) {
      sh.getRange(r, 1, P.rows.length, P.cols.length).setValues(P.rows.map(x => x.cells));
      sh.getRange(r, 1, P.rows.length, 1).setFontWeight('bold');
      P.rows.forEach((x, i) => {
        if (P.cols.length < NC) merge(r + i, P.cols.length, NC - P.cols.length + 1);
        if (x.chip) sh.getRange(r + i, 2).setBackground(x.chip.bg).setFontColor(x.chip.fg).setFontWeight('bold');
        if (x.me) sh.getRange(r + i, 1).setBackground('#FFF8D6');
      });
      sh.getRange(r, last + 1, P.rows.length, NC - last).setWrap(true);
    } else merge(r, 1, NC).setValue('No active members listed for your school yet.').setFontColor(THEME.muted);
    const n = Math.max(P.rows.length, 1);
    sh.getRange(r - 1, 1, n + 1, NC).setBorder(true, true, true, true, true, true, TABLE_LINE, SpreadsheetApp.BorderStyle.SOLID);
    r += n + 1;
  }

  // ---------- 5. Your last 4 check-ins (short) ----------
  merge(r, 1, NC).setValue('📅  YOUR LAST CHECK-INS'); header_(sh.getRange(r, 1, 1, NC)); r++;
  sh.getRange(r, 1, 1, 3).setValues([['Date', 'Events this week', 'Needs help?']])
    .setFontWeight('bold').setFontSize(9).setFontColor(THEME.muted).setBackground(THEME.band);
  merge(r, 3, NC - 2); r++;
  const hist = s.checkins.slice(0, 4);
  if (hist.length) {
    hist.forEach((x, i) => {
      sh.getRange(r + i, 1, 1, 3).setValues([[x.date, x.events || '—', x.help || '—']]);
      merge(r + i, 3, NC - 2);
      const chip = x.help === HELP_CHOICES[2] ? CHIP.red : x.help === HELP_CHOICES[1] ? CHIP.yellow : null;
      if (chip) sh.getRange(r + i, 3).setBackground(chip.bg).setFontColor(chip.fg).setFontWeight('bold');
    });
    sh.getRange(r, 1, hist.length, 1).setNumberFormat('mmm d, yyyy').setFontColor(THEME.muted);
  } else merge(r, 1, NC).setValue('No check-ins yet: the yellow button above starts your streak.').setFontColor(THEME.muted);
  const hn = Math.max(hist.length, 1);
  sh.getRange(r - 1, 1, hn + 1, NC).setBorder(true, true, true, true, true, true, TABLE_LINE, SpreadsheetApp.BorderStyle.SOLID);
  r += hn + 1;

  // ---------- 6. One privacy line ----------
  merge(r, 1, NC).setValue(`🔒 Private: only you and ${rr.name} can open this page. Please don't share, download or screenshot it (Data Privacy Act). ` +
    'It updates every morning and right after each check-in.').setFontColor(THEME.muted).setFontSize(9).setWrap(true);
  sh.setRowHeight(r, 34);

  // 📖 steps panel on the RIGHT (like every tab): tables first, a thin gap column, then the panel
  const C1 = 1, GAP = NC + 1, GC = NC + 2;
  sh.insertColumnsAfter(NC, 2);
  sh.setColumnWidth(GAP, 22);
  fitPortalColumns_(sh, C1, NC, firstDataRow, r);            // measured from the members table down
  drawPortalGuide_(sh, portalGuideLines_(s, rr), r, GC);
  const gl = portalGuideLines_(s, rr);
  const gw = Math.max.apply(null, gl.map(l => textWidth_(l[0], l[1] === 'step' || l[1] === 'head' ? 10 : 9, l[1] !== 'step')));
  sh.setColumnWidth(GC, gw + 34);                           // the panel is exactly as wide as its longest line
  // header: the portal banner covers the whole table, the "HOW TO USE" picture covers the whole panel, same height
  const wide = picture_('portal-wide');
  if (wide) {
    layoutHeader_(sh, { lastCol: NC, panelCol: GC, blob: wide, alt: 'START banner', desc: PICTURES['portal-wide'], under: 6, minPanel: gw + 34 });
  } else {                                                    // older pictures: the previous layout
    let tableW = 0;
    for (let c = 1; c <= NC; c++) tableW += sh.getColumnWidth(c);
    placeBanner_(sh, 'portal', tableW, 6);
    try {
      const g = picture_('guide');
      if (g) {
        const gpw = sh.getColumnWidth(GC), gph = Math.round(gpw / 3.4);
        sh.insertImage(g, GC, 1, 0, Math.max(0, sh.getRowHeight(1) - 6 - gph)).setWidth(gpw).setHeight(gph)
          .setAltTextTitle('START banner (how to use)').setAltTextDescription(PICTURES.guide);
      }
    } catch (e) { Logger.log('Portal guide picture: ' + e.message); }
  }
  sh.setHiddenGridlines(true);
  sh.setTabColor(THEME.tabs.portal);
}

// Welcome email: same deck design as the reminders (Email_Reminder template, logos, pill button)
function sendPortalWelcome_(s, fileId, rr) {
  const url = `https://docs.google.com/spreadsheets/d/${fileId}/edit`;
  const greet = `Dear ${displayName_(s.focal)},`;
  const sign = rrSignature_({ rrName: rr.name !== 'your Regional Representative' ? rr.name : '', region: rr.region });
  const intro = [
    'Good day!',
    `Thank you for serving as the START Campus Lead of ${s.name}. You now have your own ${PORTAL_NAME}: ` +
      'a private page that shows your school\'s weekly check-ins, your streak, and your school\'s scholars.',
  ];
  const outro = [
    `The portal is private: only your Google account (${s.email}) can open it. ` +
      'It updates every morning and right after each check-in. You can also find it in Google Drive under "Shared with me".',
    'Each week you will also receive a reminder with your school\'s check-in form. It takes about two (2) minutes.',
    'Thank you for your continued support of the START program.',
  ];
  const html = renderEmail_({
    badge: 'Welcome', badgeColor: THEME.blue, accent: THEME.navy, title1: 'Your', title2: 'Portal',
    greeting: greet, intro: intro, outro: outro,
    link: url, buttonText: 'Open my ' + PORTAL_NAME, buttonNote: 'Private to your Google account',
    school: s.name, weekLabel: 'YOUR ROLE', week: 'START Campus Lead',
    deadlineLabel: 'UPDATES', deadline: 'Every morning',
    signature: sign,
    region: rr.region ? rr.region + (sign.length > 2 ? ' · ' + sign[0] : '') : (sign.length > 2 ? sign[0] : ''),
    footer: 'This message was sent through the START Tracker on behalf of your Regional Representative. ' +
            'You may reply to this email directly.',
  });
  const text = [greet, ''].concat(intro, ['', PORTAL_NAME + ': ' + url, ''], outro, ['', 'Respectfully,'], sign).join('\n');
  try {
    sendMail_({ to: s.email, subject: `START · Your ${PORTAL_NAME} (${s.name})`, htmlBody: html, body: text,
                        inlineImages: logoInlineImages_(),
                        name: sign.length > 2 ? `${sign[0]} (START ${rr.region || ''})`.replace(/\s+\)/, ')') : 'START Tracker' });
  } catch (e) { Logger.log('Welcome email not sent: ' + e.message); }
}

/* ---------- connected names: picking a START Campus Lead fills their email ---------- */

// Simple trigger: runs on every edit by a person.
//  Schools › START Campus Lead (D): the name is matched to the Scholars spelling (assignFocal_)
//  Schools › Contact (E) is a formula from Scholars: typing there is undone, with a message
//  Emails are typed ONLY in Scholars › Contact (email); every other tab reads them from there
//  Renaming ONE cell carries the new name everywhere:
//    Schools › School  -> Scholars, Outreach log, Weekly Check-in history, portal list (the form follows via onSheetChange)
//    Scholars › Name   -> Schools › START Campus Lead
function onEdit(e) {
  try {                                                                   // 🗓 Calendar view (22_Calendar): sheet-only, safe here
    const t = e && e.range ? e.range.getSheet().getName() : '';
    if ([TAB.calview, TAB.events, TAB.outreach, TAB.settings].indexOf(t) !== -1) renderCalView_(e.source);
  } catch (err) { /* view not built yet */ }
  try {
    const rng = e && e.range;
    if (!rng || rng.getLastRow() < 2) return;
    const sh = rng.getSheet(), tab = sh.getName(), ss = sh.getParent();
    const c1 = rng.getColumn(), c2 = rng.getLastColumn(), covers = c => c1 <= c && c <= c2;
    const r1 = Math.max(rng.getRow(), 2), n = rng.getLastRow() - r1 + 1;
    const renamed = rng.getNumRows() === 1 && rng.getNumColumns() === 1 && e.oldValue !== undefined &&
                    String(e.oldValue).trim() && String(e.value || '').trim() && String(e.oldValue).trim() !== String(e.value).trim();
    if (renamed && tab === TAB.schools && c1 === 1) renameSchool_(ss, String(e.oldValue).trim(), String(e.value).trim());
    if (renamed && tab === TAB.scholars && c1 === 1) {
      replaceExact_(ss.getSheetByName(TAB.schools).getRange(2, 4, CONFIG.maxSchools, 1), String(e.oldValue).trim(), String(e.value).trim());
    }
    if (tab === TAB.schools) {
      if (covers(5)) {                                                   // emails live in Scholars only
        setSchoolEmailFormula_(sh, false);
        ss.toast('Emails are typed in the Scholars tab (Contact column). Schools shows them automatically.', 'START Tracker', 8);
      }
      if (!covers(4)) return;
      const scholars = scholarRows_(ss);
      sh.getRange(r1, 4, n, 1).getValues().forEach((v, i) => assignFocal_(sh, r1 + i, v[0], scholars));
      syncRoles_(ss);                                                    // Scholars › Role follows
    } else if (tab === TAB.scholars) {
      if (covers(10)) applyRoles_(ss, sh, r1, n);                        // Role picked in Scholars
      if (covers(1) || covers(10)) syncRoles_(ss);                       // new names get "Member"; one Lead per school
    }
  } catch (err) { /* never block typing */ }
}

// Whole-cell find & replace inside one range
/* ---------- Role in Scholars (J): Member / ⭐ START Campus Lead ---------- */
// Schools › START Campus Lead (D) is THE record (one cell per school = one Campus Lead per school).
// Scholars › Role mirrors it, and picking a role there writes it back to Schools.
var ROLES = ['Member', '⭐ START Campus Lead'];

// Writes every scholar's Role from Schools › D (Campus Lead if they lead a school, otherwise Member)
function syncRoles_(ss) {
  const sc = ss.getSheetByName(TAB.scholars);
  if (!sc || sc.getLastRow() < 2) return;
  const n = sc.getLastRow() - 1;
  const leads = {};
  schoolRows_(ss).forEach(s => { if (s.focal) leads[s.focal] = true; });
  const want = sc.getRange(2, 1, n, 1).getValues()
    .map(r => [String(r[0]).trim() ? (leads[String(r[0]).trim()] ? ROLES[1] : ROLES[0]) : '']);
  const J = sc.getRange(2, 10, n, 1);
  if (JSON.stringify(J.getValues()) !== JSON.stringify(want)) J.setValues(want);
}

// A Role was picked in Scholars: make that scholar the Campus Lead of THEIR school (column B),
// or remove them as Campus Lead. The previous Lead of that school becomes a Member automatically.
function applyRoles_(ss, sh, r1, n) {
  const sch = ss.getSheetByName(TAB.schools), schools = schoolRows_(ss), scholars = scholarRows_(ss);
  sh.getRange(r1, 1, n, 10).getValues().forEach(x => {
    const nm = String(x[0]).trim(), school = String(x[1]).trim(), lead = String(x[9]).trim() === ROLES[1];
    if (!nm) return;
    if (lead) {
      const target = schools.find(s => s.name === school);
      if (!target) {
        ss.toast(`${nm}: the school "${school}" is not in the Schools tab, so they can't be its Campus Lead yet.`, 'START Tracker', 8);
        return;
      }
      if (target.focal === nm) return;
      const before = target.focal;
      sch.getRange(target.row, 4).setValue(nm);
      assignFocal_(sch, target.row, nm, scholars, 'Picked in Scholars › Role on ' +
        Utilities.formatDate(new Date(), CONFIG.timezone, 'MMMM d, yyyy') + '.');
      target.focal = nm;
      ss.toast(`${nm} is now the START Campus Lead of ${school}` + (before ? ` (${before} is now a Member).` : '.'), 'START Tracker', 8);
    } else {
      schools.filter(s => s.focal === nm).forEach(s => {
        sch.getRange(s.row, 4).clearContent().clearNote();
        s.focal = '';
        ss.toast(`${nm} is no longer the START Campus Lead of ${s.name}.`, 'START Tracker', 6);
      });
    }
  });
}

function replaceExact_(range, from, to) {
  if (range) range.createTextFinder(from).matchEntireCell(true).matchCase(true).replaceAllWith(to);
}

// A school was renamed in the Schools tab: carry the new name to every place that stores it
function renameSchool_(ss, from, to) {
  const col = (tab, c) => { const s = ss.getSheetByName(tab); return s && s.getLastRow() > 1 ? s.getRange(2, c, s.getLastRow() - 1, 1) : null; };
  replaceExact_(col(TAB.scholars, 2), from, to);
  replaceExact_(col(TAB.outreach, 2), from, to);
  const mc = ss.getSheetByName(TAB.mcheckins);
  if (mc && mc.getLastRow() > 1) { const c = memberCols_(mc); if (c.school >= 0) replaceExact_(col(TAB.mcheckins, c.school + 1), from, to); }
  replaceExact_(col(OLD_MASTER_TAB, 2), from, to);
  replaceExact_(col(JOIN_TAB, 3), from, to);
  replaceExact_(col(WEEKLY_ROSTER_TAB_, 4), from, to);
  replaceExact_(col(TAB.events, EV.who + 1), from, to);
  if (ss.getSheetByName(JOIN_TAB)) refreshJoinViews_(ss);
  const ck = ss.getSheetByName(TAB.checkins);
  if (ck && ck.getLastRow() > 1) { const c = ckCols_(ck); if (c.school >= 0) replaceExact_(col(TAB.checkins, c.school + 1), from, to); }
  const store = portalStore_();
  if (store.map[from] && !store.map[to]) { store.map[to] = store.map[from]; delete store.map[from]; savePortalStore_(store); }
  const calendars = calStore_();
  if (calendars.schools && calendars.schools[from] && !calendars.schools[to]) {
    calendars.schools[to] = calendars.schools[from];
    delete calendars.schools[from];
    saveCalStore_(calendars);
  }
}

// THE one place a START Campus Lead is written into Schools › START Campus Lead (typing, check-ins, Outreach log).
// A typed name like "Hannah Faye Bayno" becomes the Scholars spelling "Bayno, Hannah Faye P." so the same
// name is used everywhere (yellow highlight, Role, portal), and Schools › Contact finds their email.
function assignFocal_(sh, row, typed, scholars, note) {
  typed = String(typed || '').trim();
  const hit = typed ? findScholar_(scholars || scholarRows_(sh.getParent()), typed) : null;
  if (typed) {
    const name = hit ? String(hit[0]).trim() : typed;
    const cell = sh.getRange(row, 4);
    if (name !== typed) cell.setValue(name);
    const why = note || (name !== typed ? 'Matched to the Scholars list (you typed "' + typed + '").' : '');
    if (why) cell.setNote(why);
  }
}

// Finds a scholar by exact name, or by the same words ignoring order, commas and middle initials
function findScholar_(scholars, typed) {
  const key = n => displayName_(n).toLowerCase().replace(/[.,]/g, ' ').split(/\s+/)
    .filter(w => w.length > 1).sort().join(' ');
  const exact = scholars.find(r => String(r[0]).trim() === typed);
  if (exact) return exact;
  const k = key(typed);
  const matches = scholars.filter(r => key(r[0]) === k);
  return matches.length === 1 ? matches[0] : null;      // only when it is unambiguous
}


// ━━━━━━━━━━━━━━━━━━━━  Web  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  🌐 WEBSITE  (one per region: members, START Campus Leads and the RR, same address for everyone)
 *   • An Apps Script "web app" made from this tracker. The spreadsheet stays the database.
 *   • Sign-in: type your email -> a 6-digit code arrives by email -> signed in for 30 days on that device.
 *     No password, works with any Gmail or school account.
 *   • WHO SEES WHAT is decided here on the server, on every page load, from the tracker itself:
 *       member   = Active scholar with that email in Scholars › Contact             -> only their own check-ins
 *       Lead     = START Campus Lead of a school (Schools tab, email in Scholars)    -> their school's members
 *       RR       = this tracker's owner, the automatic-email account, or Settings › "Send weekly summary to"
 *                                                                                    -> the whole region
 *     Remove someone from the tracker and they lose access on their next page load.
 *   • Deploy once (Apps Script › Deploy › New deployment › Web app):
 *       Execute as: Me   ·   Who has access: Anyone
 * ================================================================ */
var WEB = { codeMins: 15, sessionDays: 30, shortHours: 12, maxTries: 5, perHour: 5,
            codesPerDay: 40,      // sign-up / reset codes per day, all scholars together
            mailReserve: 25,      // codes stop when fewer emails than this are left today (reminders come first)
            lateDays: 0,          // a missed check-in is closed for good after 11:59 PM of its due day (was 2 days, marked late)
            cacheSecs: 300 };     // a page is kept ready for 5 minutes; your own changes clear it at once (faster on busy nights)

// The page itself
function doGet(e) {
  if (e && e.parameter && e.parameter.ping) {                       // 🩺 Check my tracker: which code version is live?
    return ContentService.createTextOutput(JSON.stringify({ v: CODE_VERSION })).setMimeType(ContentService.MimeType.JSON);
  }
  // The page's start-up facts (region, pictures) are kept for 10 minutes, so opening START does not wait for the sheet each time
  let boot = null, bc = null;
  try { bc = CacheService.getScriptCache(); const hit = bc.get('WEB_BOOT_V1'); if (hit) boot = JSON.parse(hit); } catch (err) { boot = null; }
  if (!boot) {
    let st = {};
    try { webSheet_(); st = getSettings_(); } catch (err) { Logger.log('Website: ' + err.message); }
    // Google's own answer to "what is my address?" can be an old, broken one: only a fallback until the RR saves the real one (🌐)
    try { const u = ScriptApp.getService().getUrl(), p = docProps_(); if (u && !p.getProperty('WEB_URL_SET') && p.getProperty('WEB_URL') !== u) p.setProperty('WEB_URL', u); } catch (err) { /* not deployed */ }
    boot = {
      region: String(st.region || ''), regionName: String(st.regionName || ''), pics: webPics_(),
      mark: typeof START_MARK === 'undefined' ? '' : START_MARK,
      place: (() => { try { return placeOfWeek_(); } catch (err) { return null; } })(),
      appOrigin: ''
    };
    try { const j = JSON.stringify(boot); if (bc && j.length < 95000) bc.put('WEB_BOOT_V1', j, 600); } catch (err) { /* no cache */ }
  }
  const st = { region: boot.region };
  // Opened inside the START app (GitHub Pages)? Only the app address saved by the RR may show this page in a frame.
  const app = webOrigin_(e && e.parameter && e.parameter.app);
  const inApp = !!app && webAppOrigins_().indexOf(app) !== -1;
  boot.appOrigin = inApp ? app : '';
  // Keep Web_App.html valid JavaScript in local editors. Escape '<' so a setting cannot close the script tag.
  const raw = HtmlService.createTemplateFromFile('Web_App').getRawContent();
  const marker = '"__START_BOOT__"';
  if (raw.indexOf(marker) === -1) throw new Error('Web_App bootstrap marker is missing.');
  const json = JSON.stringify(boot).replace(/</g, '\\u003c').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
  const out = HtmlService.createHtmlOutput(raw.replace(marker, json))
    .setTitle('START Scholar Portal' + (st.region ? ' · ' + st.region : ''))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
  return inApp ? out.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) : out;
}

// "https://name.github.io/start-app/" -> "https://name.github.io" ('' if it is not a web address)
function webOrigin_(u) {
  const m = String(u || '').trim().match(/^(https:\/\/[a-z0-9.-]+(?::\d+)?)(?:[\/?#]|$)/i);
  return m ? m[1].toLowerCase() : '';
}
// The START app addresses allowed to show this website (saved from 🌐 Website in the menu)
function webAppOrigins_() {
  const saved = String(docProps_().getProperty('WEB_APP_ORIGIN') || '').split(/[\s,]+/).filter(Boolean);
  if (saved.indexOf('https://feliciety02.github.io') === -1) saved.push('https://feliciety02.github.io');
  return saved;
}
// Menu box: save the START app address
// The address people should open: the START app if saved, else the website
function webHome_() { try { const p = docProps_(); return p.getProperty('WEB_APP_URL') || 'https://feliciety02.github.io/START-Davao-Chapter/'; } catch (e) { return 'https://feliciety02.github.io/START-Davao-Chapter/'; } }
// The website's real address: the one the RR pasted from Deploy › Manage deployments (🌐 Website), else Google's guess
function webSiteUrl_() { try { const p = docProps_(); return p.getProperty('WEB_URL_SET') || p.getProperty('WEB_URL') || ''; } catch (e) { return ''; } }
// 🌐 Website › Website address: checked (it must open and answer as START) before it is saved
function saveWebsiteAddress(url) {
  sheetUser_();                                   // sheet only: never from the website
  url = String(url || '').trim().replace(/\?.*$/, '');
  if (!/^https:\/\/script\.google\.com\/(a\/[^/]+\/)?macros\/s\/[A-Za-z0-9_-]{20,}\/exec$/.test(url))
    return '⚠️ Paste the "Web app URL" from Deploy › Manage deployments (it ends in /exec).';
  let v = '';
  try { const r = UrlFetchApp.fetch(url + '?ping=1', { muteHttpExceptions: true, followRedirects: true }); v = JSON.parse(r.getContentText()).v || ''; } catch (e) { v = ''; }
  if (!v) return '⚠️ That address does not open START. Copy the Web app URL again from Deploy › Manage deployments, and check Who has access: Anyone.';
  docProps_().setProperty('WEB_URL_SET', url); docProps_().setProperty('WEB_URL', url);
  return v === CODE_VERSION ? '✅ Saved. Emails and this menu now use this address.' : '✅ Saved. (It runs an older code version: Deploy › Manage deployments › ✏️ › New version.)';
}

function saveAppAddress(url) {
  sheetUser_();                                   // sheet only: never from the website
  const o = webOrigin_(url);
  if (!o && String(url || '').trim()) return 'Use the full address, starting with https://';
  if (!o) { docProps_().deleteProperty('WEB_APP_ORIGIN'); docProps_().deleteProperty('WEB_APP_URL'); return 'Removed.'; }
  url = String(url).trim();
  if (!/\/$/.test(url) && !/\.[a-z]+$/i.test(url)) url += '/';
  // email buttons open this address: it must be the START app itself (e.g. https://name.github.io/start-app/), not just the site
  try {
    const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    const page = r.getResponseCode() === 200 ? r.getContentText() : '';
    if (!/START_APP|config\.js/.test(page))
      return '⚠️ That address is not the START app. Use the full address of the app, e.g. ' + o + '/start-app/';
  } catch (e) { return '⚠️ Could not open that address. Check it in your browser, then try again.'; }
  docProps_().setProperty('WEB_APP_ORIGIN', o);
  docProps_().setProperty('WEB_APP_URL', url);
  return 'Saved ✓ Email buttons open the app.';
}

// The same public decoration pictures as the emails (logos, sparkles, star). Only saved links are used, so the page
// opens fast; anything missing falls back to plain colours and text.
function webPics_() {
  const out = { logo: '', sparkle: '', star: '' };
  try {
    const p = docProps_(), at = (id, size) => id ? 'https://lh3.googleusercontent.com/d/' + id + size : '';
    out.logo = at(p.getProperty('EMAIL_LOGO_ID'), '=w375');
    out.sparkle = at(p.getProperty('EMAIL_PIC_email-sparkle'), '=s0');
    out.star = at(p.getProperty('EMAIL_PIC_email-star'), '=s0');
  } catch (e) { /* plain look */ }
  return out;
}

// The tracker spreadsheet (a website visit has no "open" spreadsheet: use this tracker's own id)
function webSheet_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActive(); } catch (e) { ss = null; }
  if (ss) return ss;
  const id = docProps_().getProperty(TRACKER_KEY);
  if (!id) throw new Error('The tracker is not set up yet. The RR must open the Google Sheet once.');
  ss = SpreadsheetApp.openById(id);
  SpreadsheetApp.setActiveSpreadsheet(ss);
  return ss;
}

function webHash_(s) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8));
}
function webFmt_(d, p) { return d instanceof Date ? Utilities.formatDate(d, CONFIG.timezone, p) : ''; }
function webSafe_(fn) {
  try { return fn(); }
  catch (e) {
    Logger.log('Website error: ' + (e && e.stack || e));
    return { ok: false, msg: 'Something went wrong on our side. Please try again in a minute.' };
  }
}

// Timings contain no account or scholar details. They appear in Apps Script executions.
function webTimed_(label, fn) {
  const started = Date.now();
  try { return webSafe_(fn); }
  finally { if (typeof Logger !== 'undefined') Logger.log('Website timing ' + label + ': ' + (Date.now() - started) + ' ms'); }
}

// Who is this email in THIS tracker? (checked again on every request)
function webRoles_(ss, email) {
  email = emailOf_(email);
  const r = { email, name: '', member: null, leads: [], admin: false };
  if (!email) return r;
  r.suspended = docProps_().getProperty('WEB_SUSP_' + webHash_(email)) === '1';
  let st = {}; try { st = getSettings_(); } catch (e) { st = {}; }
  // Summary recipients receive a report, not control over scholar and school records.
  const admins = [me_(), automationOwner_()].filter(Boolean);
  r.admin = admins.indexOf(email) !== -1;
  const p = scholarByEmail_(ss)[email];
  if (p) r.suspended = p.status === 'Suspended';
  if (p && p.status === 'Active') { r.member = p; r.name = p.name; }
  schoolRows_(ss).forEach(s => {
    if (s.focal && s.email === email && !s.inactive) { r.leads.push(s.name); r.name = r.name || s.focal; }
  });
  if (!r.name && r.admin) r.name = String(st.rrName || '').trim();
  return r;
}
function webAllowed_(who) { return !!(who && !who.suspended && (who.admin || who.member || who.leads.length)); }

/* ---------- sign-in: password, with email codes ONLY for a new account or a forgotten password ----------
 *  Passwords are never stored: only a salted, many-times-hashed copy, in this tracker's own hidden storage.
 *  Too many wrong passwords lock that email for 15 minutes. A new password signs out every other device.
 */
var WEB_PW = { min: 8, rounds: 800, maxFails: 5, lockMins: 15 };

// An account made with an email that is now an "other" email moves to the main email (once), so nobody starts over
function webMoveAccount_(from, to) {
  if (!from || !to || from === to) return;
  const a = webAccount_(from);
  if (!a || !a.h || webAccount_(to)) return;
  webSaveAccount_(to, a);
  docProps_().deleteProperty(webAcctKey_(from));
}
function webAcctKey_(email) { return 'WEB_A_' + webHash_(emailOf_(email)); }
function webAccount_(email) { try { return JSON.parse(docProps_().getProperty(webAcctKey_(email)) || 'null'); } catch (e) { return null; } }
function webSaveAccount_(email, a) { docProps_().setProperty(webAcctKey_(email), JSON.stringify(a)); }
function webPwHash_(pw, salt) {
  let h = salt + '|' + pw;
  for (let i = 0; i < WEB_PW.rounds; i++) h = webHash_(h + '|' + salt);
  return h;
}
// Passwords people often pick (and ones easy to guess for START): refused
var WEB_PW_COMMON = ['password', 'passw0rd', 'qwerty', 'qwertyui', 'asdfghjk', 'iloveyou', 'welcome', 'letmein', 'monkey', 'dragon',
  'sunshine', 'princess', 'football', 'baseball', 'superman', 'batman', 'trustno1', 'abc123', 'abcdefg', 'admin', 'login', 'master',
  'start', 'dost', 'dostsei', 'scholar', 'scholarship', 'davao', 'mindanao', 'philippines', 'pilipinas', 'datacamp', 'mahalkita', 'iloveu'];
function webPwProblem_(pw, email) {
  pw = String(pw || '');
  if (pw.length < WEB_PW.min) return `Use ${WEB_PW.min} or more characters`;
  if (pw.length > 200) return 'That password is too long';
  const low = pw.toLowerCase(), em = String(email || '').toLowerCase(), user = em.split('@')[0];
  if (low === em || (user.length >= 4 && low.indexOf(user) !== -1)) return 'Don\'t use your email in your password';
  if (/^(.)\1+$/.test(pw)) return 'Choose a password that is harder to guess';
  const core = low.replace(/[^a-z0-9]/g, '').replace(/[0-9]+$/, '').replace(/^[0-9]+/, '');      // "Davao2026!" -> "davao"
  if (WEB_PW_COMMON.indexOf(core) !== -1 || WEB_PW_COMMON.indexOf(low) !== -1) return 'That password is too common. Add a few words only you would think of';
  const seq = '01234567890abcdefghijklmnopqrstuvwxyzqwertyuiopasdfghjklzxcvbnm';
  if (/^[0-9]+$/.test(pw) && pw.length < 12) return 'Use letters too, or 12 or more numbers';
  if (pw.length < 12 && (seq.indexOf(low) !== -1 || seq.split('').reverse().join('').indexOf(low) !== -1)) return 'Choose a password that is harder to guess';
  if (new Set(low).size < 4) return 'Use more different characters';
  return '';
}
function webNewSession_(email, ver, shared) {
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const life = shared ? WEB.shortHours * 3600000 : WEB.sessionDays * 86400000;     // not their own device: 12 hours
  docProps_().setProperty('WEB_S_' + webHash_(token), JSON.stringify({ e: email, v: ver || 0, exp: Date.now() + life }));
  webTidy_(docProps_());
  return token;
}

// Step 1: the email. A first-time scholar confirms the address before requesting a code.
function webStart(email, purpose) {
  return webSafe_(() => {
    const ss = webSheet_();
    const typed = emailOf_(email);
    if (!typed) return { ok: false, msg: 'Enter an email address' };
    email = primaryEmail_(ss, typed);                               // an "other" email signs in as the scholar's main email
    webMoveAccount_(typed, email);
    const who = webRoles_(ss, email), acct = webAccount_(email);
    // Only accounts in the Scholars roster can receive verification codes; eligible new requests enter the roster automatically.
    if (!webAllowed_(who)) return webNotListed_(ss, email);
    if (purpose !== 'reset' && purpose !== 'setup' && acct && acct.h) return { ok: true, mode: 'password' };
    if ((!acct || !acct.h) && purpose !== 'setup' && purpose !== 'reset')
      return { ok: true, mode: 'setup', to: email === typed ? '' : webMask_(email) };
    // the code fits the account: no password yet = create it; a password = reset it (whichever button they pressed)
    const kind = acct && acct.h ? 'reset' : 'new';
    const sent = webMailCode_(email, kind, who);                   // codes go ONLY to the main email
    if (!sent.ok) return sent;
    return { ok: true, mode: 'code', purpose: kind, recent: !!sent.recent, retryAfter: sent.retryAfter || 60,
             to: email === typed ? '' : webMask_(email) };
  });
}

// "a•••z@gmail.com": shows which inbox to check without spelling out the whole address
function webMask_(email) {
  const at = String(email).indexOf('@'); if (at < 1) return '';
  const u = email.slice(0, at);
  return (u.length <= 2 ? u.charAt(0) + '•' : u.charAt(0) + '•••' + u.charAt(u.length - 1)) + email.slice(at);
}

// Not on the list (yet): new scholars go straight to "Ask to join"; a waiting or closed request gets a plain answer.
function webNotListed_(ss, email) {
  const p = scholarByEmail_(ss)[email];
  if (p) return { ok: false, msg: 'Your START access is turned off right now (Status: ' + (p.status || 'blank') + '). Please contact your Campus Lead or Regional Representative.' };
  const j = joinStateOf_(ss, email);
  if (j === 'waiting') return { ok: true, mode: 'pending', msg: 'Your Regional Representative will review your details. Once approved, return here to set up your password.' };
  if (j === 'sorry') return { ok: true, mode: 'joined', sorry: true, msg: JOIN_SORRY };
  if (j === 'added') return { ok: true, mode: 'access-check', msg: 'Your join request was approved, but your account access needs a tracker check. Please contact your Regional Representative. You do not need to register again.' };
  return { ok: true, mode: 'join', note: j === 'declined'
    ? 'Your last request wasn\'t approved. Check your details and send again, or contact your Campus Lead or Regional Representative.'
    : 'This email isn\'t on the START list yet. New scholar? Fill this in once, and you\'ll be added to START.' };
}

function webMailCode_(email, kind, who) {
  const lock = docLock_();
  if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try sending the code again.' };
  try {
    const props = docProps_(), key = 'WEB_C_' + webHash_(email), now = Date.now();
    const previous = props.getProperty(key);
    let rec = {}; try { rec = JSON.parse(previous || '{}'); } catch (e) { rec = {}; }
    rec.sent = (Array.isArray(rec.sent) ? rec.sent : []).filter(t => now - t < 3600000);
    if (rec.sent.length && now - rec.sent[rec.sent.length - 1] < 60000 && rec.k === kind && rec.h && rec.exp > now)
      return { ok: true, recent: true, retryAfter: Math.ceil((60000 - (now - rec.sent[rec.sent.length - 1])) / 1000) };
    if (rec.sent.length >= WEB.perHour) return { ok: false, msg: 'You asked for ' + WEB.perHour + ' codes in the last hour. Please use the newest code, or try again in an hour.' };
    // The daily email limit is shared with reminders.
    const dayKey = 'WEB_D_' + Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyyMMdd');
    const previousDay = props.getProperty(dayKey), today = Number(previousDay || 0);
    if (today >= WEB.codesPerDay || MailApp.getRemainingDailyQuota() <= WEB.mailReserve)
      return { ok: false, msg: 'START has reached today\'s email limit, so no code could be sent. Please try again tomorrow, or ask your Regional Representative.' };
    const code = String(100000 + parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 12), 16) % 900000);
    let m;
    try { m = webEmail_(kind, who, code); }
    catch (e) { Logger.log('Sign-in email template: ' + e.message); return { ok: false, msg: 'START could not prepare the email. Please ask your Regional Representative to check the tracker.' }; }
    Object.assign(rec, { h: webHash_(email + '|' + code), k: kind, exp: now + WEB.codeMins * 60000, tries: 0 });
    rec.sent.push(now);
    const restore = () => {
      try { if (previous === null) props.deleteProperty(key); else props.setProperty(key, previous); } catch (e) { Logger.log('Code rollback: ' + e.message); }
      try { if (previousDay === null) props.deleteProperty(dayKey); else props.setProperty(dayKey, previousDay); } catch (e) { Logger.log('Code count rollback: ' + e.message); }
    };
    try { props.setProperty(key, JSON.stringify(rec)); props.setProperty(dayKey, String(today + 1)); }
    catch (e) { restore(); Logger.log('Sign-in code storage: ' + e.message); return { ok: false, msg: 'START could not save the code. Please try again later.' }; }
    try { MailApp.sendEmail({ to: email, subject: m.subject, htmlBody: m.html, body: m.text, name: m.from }); }
    catch (e) {
      Logger.log('Sign-in code email failed: ' + e.message);
      restore();
      return { ok: false, msg: 'We couldn\'t send the email just now. Please try again in a few minutes.' };
    }
    return { ok: true, retryAfter: 60 };
  } finally { lock.releaseLock(); }
}

// Step 2 (new account / reset): check the code. Gives a 15-minute pass to choose a password.
function webCheckCode(email, code) {
  return webSafe_(() => {
    const ss = webSheet_();
    email = primaryEmail_(ss, email);
    const lock = docLock_();
    if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try that code again.' };
    try {
    const props = docProps_(), key = 'WEB_C_' + webHash_(email);
    let rec = {}; try { rec = JSON.parse(props.getProperty(key) || '{}'); } catch (e) { rec = {}; }
    if (!email || !rec.h) return { ok: false, msg: 'This code was already used or is no longer valid. Send a new one.' };
    if (Date.now() > rec.exp) return { ok: false, msg: 'This code has expired. Send a new one.' };
    if ((rec.tries || 0) >= WEB.maxTries) return { ok: false, msg: 'Too many wrong tries. Send a new code.' };
    if (webHash_(email + '|' + String(code || '').replace(/\D/g, '')) !== rec.h) {
      rec.tries = (rec.tries || 0) + 1; props.setProperty(key, JSON.stringify(rec));
      const left = WEB.maxTries - rec.tries;
      return { ok: false, msg: left > 0 ? `Wrong code. ${left} ${left === 1 ? 'try' : 'tries'} left.` : 'Too many wrong tries. Send a new code.' };
    }
    if (!webAllowed_(webRoles_(ss, email)))
      return { ok: false, msg: 'Your membership must be approved before password setup. Contact your Regional Representative.' };
    const ticket = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    const tkey = 'WEB_T_' + webHash_(ticket);
    try { props.setProperty(tkey, JSON.stringify({ e: email, k: rec.k || 'new', exp: Date.now() + 15 * 60000 })); }
    catch (e) { Logger.log('Password setup ticket: ' + e.message); return { ok: false, msg: 'START could not continue password setup. Please try the code again.' }; }
    rec.h = '';                                               // a code works once
    try { props.setProperty(key, JSON.stringify(rec)); }
    catch (e) { try { props.deleteProperty(tkey); } catch (cleanup) {} return { ok: false, msg: 'START could not finish checking the code. Please try again.' }; }
    return { ok: true, ticket, purpose: rec.k || 'new' };
    } finally { lock.releaseLock(); }
  });
}

// Step 3 (new account / reset): save the password, sign in here, sign out everywhere else
function webSetPassword(ticket, pw, keep) {
  return webSafe_(() => {
    const ss = webSheet_(), lock = docLock_();
    if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try saving your password again.' };
    try {
    const props = docProps_(), tkey = 'WEB_T_' + webHash_(ticket || '');
    let t = null; try { t = JSON.parse(props.getProperty(tkey) || 'null'); } catch (e) { t = null; }
    if (!t || t.exp < Date.now()) return { ok: false, restart: true, msg: 'That took too long. Start again.' };
    const bad = webPwProblem_(pw, t.e);
    if (bad) return { ok: false, msg: bad };
    const who = webRoles_(ss, t.e);
    if (!webAllowed_(who)) return { ok: false, restart: true, msg: 'Your START membership is not approved yet.' };
    const old = webAccount_(t.e) || {};
    if (t.k === 'new' && old.h) return { ok: false, restart: true, msg: 'A password is already set for this email. Sign in or use Forgot password.' };
    const salt = Utilities.getUuid().replace(/-/g, '');
    const a = { s: salt, h: webPwHash_(pw, salt), v: (old.v || 0) + 1, at: Date.now(), fails: 0, lock: 0 };
    webSaveAccount_(t.e, a);
    try { markVerifiedEmail_(ss, t.e); } catch (e) { Logger.log('Verified email: ' + e.message); }
    try { props.deleteProperty(tkey); } catch (e) { Logger.log('Password ticket cleanup: ' + e.message); }
    try { webEndSessions_(t.e); } catch (e) { Logger.log('Old session cleanup: ' + e.message); }
    try { const m = webEmail_(old.h ? 'changed' : 'created', who); MailApp.sendEmail({ to: t.e, subject: m.subject, htmlBody: m.html, body: m.text, name: m.from }); }
    catch (e) { Logger.log('Password email: ' + e.message); }
    try { return { ok: true, token: webNewSession_(t.e, a.v, keep === false) }; }
    catch (e) { Logger.log('New session: ' + e.message); return { ok: true, signIn: true, msg: 'Password saved. Please sign in with it.' }; }
    } finally { lock.releaseLock(); }
  });
}

// Everyday sign-in: email + password
function webPasswordSignIn(email, pw, keep) {
  return webSafe_(() => {
    const ss = webSheet_();
    const typed = emailOf_(email);
    email = primaryEmail_(ss, typed);
    webMoveAccount_(typed, email);
    const a = webAccount_(email);
    if (!a || !a.h) return { ok: false, msg: 'Wrong email or password' };
    if (a.lock && a.lock > Date.now()) return { ok: false, msg: `Too many tries. Wait ${Math.ceil((a.lock - Date.now()) / 60000)} min or reset your password.` };
    if (webPwHash_(String(pw || ''), a.s) !== a.h) {
      a.fails = (a.fails || 0) + 1;
      if (a.fails >= WEB_PW.maxFails) { a.lock = Date.now() + WEB_PW.lockMins * 60000; a.fails = 0; }
      webSaveAccount_(email, a);
      return { ok: false, msg: a.lock > Date.now() ? `Too many tries. Wait ${WEB_PW.lockMins} min or reset your password.` : 'Wrong password. Try again or reset it.' };
    }
    if (!webAllowed_(webRoles_(ss, email))) return { ok: false, msg: 'This email is no longer on your region\'s START list.' };
    if (a.fails || a.lock) { a.fails = 0; a.lock = 0; webSaveAccount_(email, a); }
    return { ok: true, token: webNewSession_(email, a.v, keep === false) };
  });
}

// Signed in: change the password (needs the current one)
function webChangePassword(token, oldPw, newPw) {
  return webSafe_(() => {
    const s = webSession_(token);
    if (!s) return { ok: false, signedOut: true, msg: 'Please sign in again' };
    const a = webAccount_(s.e);
    if (!a || webPwHash_(String(oldPw || ''), a.s) !== a.h) return { ok: false, msg: 'Current password is wrong' };
    const bad = webPwProblem_(newPw, s.e);
    if (bad) return { ok: false, msg: bad };
    const salt = Utilities.getUuid().replace(/-/g, '');
    Object.assign(a, { s: salt, h: webPwHash_(newPw, salt), v: (a.v || 0) + 1, at: Date.now(), fails: 0, lock: 0 });
    webSaveAccount_(s.e, a);
    webEndSessions_(s.e);
    try { const m = webEmail_('changed', webRoles_(webSheet_(), s.e)); MailApp.sendEmail({ to: s.e, subject: m.subject, htmlBody: m.html, body: m.text, name: m.from }); }
    catch (e) { Logger.log('Password email: ' + e.message); }
    return { ok: true, token: webNewSession_(s.e, a.v) };
  });
}

function webSignOut(token) {
  return webSafe_(() => { if (token) docProps_().deleteProperty('WEB_S_' + webHash_(token)); return { ok: true }; });
}
// Account menu › Sign out of all devices (lab computers, a lost phone)
function webSignOutAll(token) {
  return webSafe_(() => {
    const s = webSession_(token);
    if (!s) return { ok: false, signedOut: true, msg: 'This session expired. Sign in again before signing out other devices.' };
    webEndSessions_(s.e); webForget_(s.e);
    return { ok: true };
  });
}

// A session is valid only while it matches the account's password version (a new password ends old sessions)
function webSession_(token) {
  if (!token || String(token).length < 40) return null;
  let s = null; try { s = JSON.parse(docProps_().getProperty('WEB_S_' + webHash_(token)) || 'null'); } catch (e) { s = null; }
  if (!s || s.exp <= Date.now()) return null;
  if (docProps_().getProperty('WEB_SUSP_' + webHash_(s.e)) === '1') return null;
  const a = webAccount_(s.e);
  if (!a || !a.h || (s.v || 0) !== (a.v || 0)) return null;
  return s;
}
function webEndSessions_(email) {
  const props = docProps_(), all = props.getProperties();
  Object.keys(all).forEach(k => {
    if (k.indexOf('WEB_S_') !== 0) return;
    try { if (JSON.parse(all[k]).e === email) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); }
  });
}

// A person's cached page is dropped (after a check-in, sign-out everywhere)
function webCacheKey_(email) { return 'WL_' + (docProps_().getProperty('WEB_DATA_V') || '0') + '_' + webHash_(emailOf_(email)); }
function webBumpData_() {
  const props = docProps_();
  props.setProperty('WEB_DATA_V', String(Number(props.getProperty('WEB_DATA_V') || 0) + 1));
}
// Cached pages are zipped, so even the RR's big page fits (Google allows 100 KB per cached item)
function webCachePack_(json) {
  if (json.length < 60000) return 'j:' + json;
  const z = 'z:' + Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(json, 'application/json')).getBytes());
  return z.length < 95000 ? z : '';
}
function webCacheRead_(v) {
  if (!v) return null;
  if (v.slice(0, 2) === 'j:') return JSON.parse(v.slice(2));
  if (v.slice(0, 2) === 'z:') return JSON.parse(Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(v.slice(2)), 'application/x-gzip')).getDataAsString());
  return JSON.parse(v);                                                // an older plain entry
}
function webForget_(email) { try { CacheService.getScriptCache().remove(webCacheKey_(email)); } catch (e) { /* no cache */ } }

// Every morning: accounts of people no longer on the list (graduated, dropped, changed email) are removed
function webCleanAccounts_(ss) {
  ss = ss || webSheet_();
  const keep = {};
  const add = e => { e = emailOf_(e); if (e) keep['WEB_A_' + webHash_(e)] = true; };
  [me_(), automationOwner_()].forEach(add);
  const by = scholarByEmail_(ss);
  Object.keys(by).forEach(e => { if (by[e].status === 'Active') add(e); });
  schoolRows_(ss).forEach(s => { if (s.focal && s.email && !s.inactive) add(s.email); });
  if (Object.keys(keep).length < 2) return 0;                          // the lists could not be read: never wipe everyone
  const props = docProps_(), all = props.getProperties();
  let n = 0;
  Object.keys(all).forEach(k => { if (k.indexOf('WEB_A_') === 0 && !keep[k]) { props.deleteProperty(k); n++; } });
  if (n) Logger.log('Website: removed ' + n + ' account(s) no longer on the list');
  return n;
}

// Old codes, passes and sessions are removed so the tracker's storage stays small (accounts stay)
function webTidy_(props) {
  try {
    const all = props.getProperties(), now = Date.now();
    Object.keys(all).forEach(k => {
      if (/^WEB_D_/.test(k)) { if (k !== 'WEB_D_' + Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyyMMdd')) props.deleteProperty(k); return; }
      if (/^WEB_JD_/.test(k)) { if (k !== 'WEB_JD_' + Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyyMMdd')) props.deleteProperty(k); return; }   // join requests: today's count only
      if (/^WEB_J_/.test(k)) { if (now - Number(all[k] || 0) > 86400000) props.deleteProperty(k); return; }                                                // one-a-day marks older than a day
      if (!/^WEB_[CST]_/.test(k)) return;
      let v = null; try { v = JSON.parse(all[k]); } catch (e) { v = null; }
      const dead = !v || (k.indexOf('WEB_C_') === 0 ? (!v.h || v.exp < now) && !(v.sent || []).some(t => now - t < 3600000) : v.exp < now);
      if (dead) props.deleteProperty(k);
    });
  } catch (e) { Logger.log('Website tidy-up: ' + e.message); }
}

// Every sign-in email, in the same design as all other START emails
function webEmail_(kind, who, code) {
  let st = {}; try { st = getSettings_(); } catch (e) { st = {}; }
  const first = (typeof firstName_ === 'function' ? firstName_(who && who.name) : '') || 'there';
  const url = docProps_().getProperty('WEB_APP_URL') || webSiteUrl_();
  const big = c => `<div style="margin:6px 0 4px 0;text-align:center;font-size:34px;font-weight:800;letter-spacing:10px;color:#12203A;` +
                   `background:#FFF6BF;border:1px solid #F5E27A;border-radius:14px;padding:14px 0;">${c}</div>` +
                   `<p style="margin:6px 0 0 0;text-align:center;font-size:12px;color:#4b5b77;">Works once, for ${WEB.codeMins} minutes.</p>`;
  const K = {
    new:     { subject: `START · Your code: ${code}`, badge: 'New account', title: ['Your', 'code'],
               intro: ['Your START membership is approved. Enter this code to create your password.'], extra: code ? big(code) : '',
               outro: ['Didn\'t ask for this? You can ignore it. Never share this code: START will never ask you for it.'] },
    reset:   { subject: `START · Your reset code: ${code}`, badge: 'Password reset', title: ['Reset', 'code'],
               intro: ['Enter this code to set a new password.'], extra: code ? big(code) : '',
               outro: ['Didn\'t ask for this? Ignore it. Your password stays the same. Never share this code: START will never ask you for it.'] },
    created: { subject: 'START · Your password is set', badge: 'Account', title: ['Password', 'set'],
               intro: ['From now on, sign in with your email and password.'], extra: '',
               outro: ['Didn\'t do this? Reply to this email.'] },
    changed: { subject: 'START · Your password was changed', badge: 'Security', title: ['Password', 'changed'],
               intro: ['Your password was changed and your other devices were signed out.'], extra: '',
               outro: ['Didn\'t do this? Click "Forgot password?" on the sign-in page, then reply to this email.'] },
  }[kind];
  const button = !!url;                                           // also guide newly approved scholars back to START
  const html = designedEmail_({
    badge: K.badge, badgeColor: kind === 'changed' ? THEME.red : THEME.blue, title: K.title,
    greeting: `Hi ${first},`, intro: K.intro, extraHtml: K.extra, outro: K.outro,
    link: button ? url : '', buttonText: button ? 'Open START' : '', buttonNote: '',
    chips: [['', ''], ['', ''], ['', '']], signature: rrSignature_(st),
  });
  const text = [`Hi ${first},`, '', K.intro.join(' '), code ? '\n' + code + '\n' : '', K.outro.join(' '), button ? '\n' + url : '',
                '', 'Respectfully,'].concat(rrSignature_(st)).join('\n');
  return { subject: K.subject, html, text, from: 'START' + (st.region ? ' ' + st.region : '') };
}

// RR (menu box): forget someone's password so they set a new one with an email code
// RR, in the sheet only (never on the website): why does this email get no sign-in code? Plain answers, with the fix.
function signInCheck(email) {
  sheetUser_();                                   // sheet only: never from the website
  const ss = SpreadsheetApp.getActive(), typed = emailOf_(email);
  if (!typed) return 'Type an email address.';
  const main = primaryEmail_(ss, typed), who = webRoles_(ss, main), acct = webAccount_(main), lines = [];
  if (main !== typed) lines.push('• ' + typed + ' is an "other email" of ' + main + ': the code goes to ' + main + '.');
  if (!webAllowed_(who)) {
    const p = scholarByEmail_(ss)[main];
    const j = joinStateOf_(ss, main);
    lines.push('✗ Not on the START list yet, so no code is sent. The website tells them this and opens "Ask to join START".');
    lines.push(p ? '  They are in Scholars but their Status is "' + (p.status || 'blank') + '": set it to Active (the website tells them their access is off).'
      : j === 'waiting' ? '  Their join request is waiting on your Region page: approve it, and they get a "You\'re in" email.'
      : j === 'sorry' ? '  Their join request was marked Not eligible. If that is wrong, approve it on your Region page.'
      : j === 'declined' ? '  You declined their join request. They can send a new one.'
      : '  Fix: they fill in "Ask to join START" on the website, or you add them in Scholars (Contact email = ' + main + ', Status = Active).');
    if (main === me_() || main === automationOwner_()) lines.push('  (This is the tracker\'s own account: it should be allowed. Run ① Set up / fix the tracker.)');
    return lines.join('\n');
  }
  lines.push('✓ On the START list as ' + [who.admin ? 'Regional Representative' : '', who.leads.length ? 'Campus Lead (' + who.leads.join(', ') + ')' : '', who.member ? 'member (' + who.member.school + ')' : ''].filter(Boolean).join(' + ') + '.');
  if (acct && acct.h) lines.push('• They already have a password: they should press Sign in (or "Forgot password?" to get a code). Or use Reset above.');
  else lines.push('• No password yet: "Create account" sends them a code.');
  const props = docProps_(), today = Number(props.getProperty('WEB_D_' + Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyyMMdd')) || 0);
  let quota = 0; try { quota = MailApp.getRemainingDailyQuota(); } catch (e) { quota = 0; }
  let rec = {}; try { rec = JSON.parse(props.getProperty('WEB_C_' + webHash_(main)) || '{}'); } catch (e) { rec = {}; }
  const lastHour = (rec.sent || []).filter(t => Date.now() - t < 3600000).length;
  lines.push('• Emails left today: ' + quota + ' (codes stop at ' + WEB.mailReserve + ' so reminders still go out) · codes sent today: ' + today + ' of ' + WEB.codesPerDay + '.');
  if (quota <= WEB.mailReserve || today >= WEB.codesPerDay) lines.push('✗ No more codes today (daily email limit). They can try tomorrow, or you can Reset their password above.');
  if (lastHour >= WEB.perHour) lines.push('✗ ' + lastHour + ' codes in the last hour for this email: wait an hour.');
  if (lastHour && lastHour < WEB.perHour) lines.push('• A code was sent to ' + main + ' in the last hour: check Spam and Promotions, and search for "START".');
  return lines.join('\n');
}

function resetWebPassword(email) {
  sheetUser_();                                   // sheet only: never from the website
  return resetWebPassword_(email);
}
// the reset itself: from the sheet (above) or from the RR's website page (webAdminAction checks the RR's sign-in)
function resetWebPassword_(email) {
  email = primaryEmail_(webSheet_(), email);
  if (!email) return 'Type an email';
  const a = webAccount_(email);
  if (!a || !a.h) return 'No account for ' + email;
  docProps_().deleteProperty(webAcctKey_(email));
  webEndSessions_(email);
  return `Done. ${email} sets a new password next time.`;
}
function webAccountCount_() {
  try { return Object.keys(docProps_().getProperties()).filter(k => k.indexOf('WEB_A_') === 0).length; } catch (e) { return 0; }
}

/* ---------- what the page shows ---------- */

function webLoad(token, fresh) {
  return webTimed_('load' + (fresh ? ' fresh' : ''), () => {
    const s = webSession_(token);
    if (!s) return { ok: false, signedOut: true };
    const ss = webSheet_(), who = webRoles_(ss, s.e);
    if (!webAllowed_(who)) {
      docProps_().deleteProperty('WEB_S_' + webHash_(token));
      return { ok: false, signedOut: true, msg: 'This email is no longer on your region\'s START list. Please contact your Regional Representative.' };
    }
    const cache = (() => { try { return CacheService.getScriptCache(); } catch (e) { return null; } })(), ck = webCacheKey_(who.email);
    if (!fresh && cache) { try { const hit = webCacheRead_(cache.get(ck)); if (hit) return hit; } catch (e) { /* read it fresh */ } }
    const out = webBuild_(ss, who);
    const json = JSON.stringify(out);
    if (cache) { try { const v = webCachePack_(json); if (v) cache.put(ck, v, WEB.cacheSecs); } catch (e) { /* too big: no cache */ } }
    return JSON.parse(json);
  });
}

// One person's whole page (used for them, and by the RR's "View as")
function webBuild_(ss, who) {
    const st = getSettings_();
    const out = { ok: true, me: { email: who.email, name: displayName_(who.name || who.email),
                                 first: (typeof firstName_ === 'function' ? firstName_(who.name) : '') || 'there' },
                  region: st.region || '', regionName: st.regionName || '', rr: String(st.rrName || '').trim(), rrEmail: me_(),
                  note: String(st.portalNote || '').trim(), updated: webFmt_(new Date(), 'MMM d, h:mm a'), cals: webCalendars_(who), version: CODE_VERSION };
    if (who.member) { try { out.member = webMember_(ss, who); } catch (e) { Logger.log('Website member: ' + e.message); } }
    if (who.leads.length) { try { out.lead = webLead_(ss, who.leads, st); } catch (e) { Logger.log('Website Lead: ' + e.message); } }
    if (who.admin) { try { out.admin = webAdmin_(ss, st); } catch (e) { Logger.log('Website RR: ' + e.message); } }
    // Full event feed for monthly navigation (home cards intentionally remain six).
    try { out.calendarEvents = webCalendarEvents_(ss, who); } catch (e) { Logger.log('Website calendar: ' + e.message); out.calendarEvents = []; }
    try { out.logoV = schoolLogoVersion_(ss); } catch (e) { out.logoV = ''; }
    try {
      const r = webRanksFor_(webLeaderboard_(ss, (who.member ? [who.member.school] : []).concat(who.leads)), who.admin);
      out.ranks = r.ranks; out.ranksHidden = r.hidden;
    } catch (e) { Logger.log('Leaderboard: ' + e.message); }
    return out;
}

/* ---------- RR only: "View as" — see the website exactly as a Campus Lead or a member sees it (read-only) ---------- */
// The choices: every school (with its Lead) and every active scholar
function webViewList(token) {
  return webSafe_(() => {
    const s = webSession_(token);
    if (!s) return { ok: false, signedOut: true };
    const ss = webSheet_();
    if (!webRoles_(ss, s.e).admin) return { ok: false, msg: 'Only the Regional Representative can do this.' };
    const people = {};
    Object.entries(scholarByEmail_(ss)).forEach(([e, p]) => { if (p.status === 'Active' && p.primary === e) people[e] = { name: displayName_(p.name), school: p.school, email: e }; });
    return { ok: true, schools: joinSchoolList_(ss).map(x => { const s = schoolRows_(ss).find(y => y.name === x.name); return { name: x.name, lead: s && s.focal ? displayName_(s.focal) : '' }; }),
             people: Object.values(people).sort((a, b) => a.school.localeCompare(b.school) || a.name.localeCompare(b.name)) };
  });
}
// role 'lead' + school name, or role 'member' + email. Nothing is saved; the RR stays signed in as themselves.
function webViewAs(token, role, key) {
  return webSafe_(() => {
    const s = webSession_(token);
    if (!s) return { ok: false, signedOut: true };
    const ss = webSheet_();
    if (!webRoles_(ss, s.e).admin) return { ok: false, msg: 'Only the Regional Representative can do this.' };
    let who = null, label = '';
    // The RR uses every role as herself: her own Gmail and name, no other account and nobody else's record.
    // Member: her own Scholars record if she has one, else a member record in her name. Campus Lead: of her own school,
    // else the first school with a Campus Lead. Practice only: nothing she does as Lead or member is saved.
    const rr = webRoles_(ss, s.e), rrName = rr.name || 'Regional Representative';
    const rows = schoolRows_(ss), mine = scholarByEmail_(ss)[emailOf_(s.e)];
    const sc = (mine && rows.find(x => x.name === mine.school)) || rows.find(x => x.focal) || rows[0];
    if (!sc) return { ok: false, msg: 'Add a school in the Schools tab first.' };
    const member = mine && mine.status === 'Active' ? mine : { name: rrName, school: sc.name, status: 'Active', primary: emailOf_(s.e), program: '' };
    const asYou = { role, school: sc.name, member: '__test__', memberName: displayName_(rrName) };
    let out = null;
    if (role === 'lead') {
      who = { email: s.e, name: rrName, member: null, leads: [sc.name], admin: false };
      label = 'Campus Lead · ' + sc.name;
      out = webBuild_(ss, who);
      out.viewAs = Object.assign(asYou, { key: '', label, member: '' });
    } else if (role === 'member') {
      who = { email: s.e, name: rrName, member, leads: [], admin: false };
      label = 'Member · ' + sc.name;
      out = webBuild_(ss, who);
      out.viewAs = Object.assign(asYou, { key: '', label });
    } else if (role === 'all') {
      who = { email: s.e, name: rrName, member, leads: [sc.name], admin: true };
      label = 'All three · ' + sc.name;
      out = webBuild_(ss, who);
      out.viewAs = Object.assign(asYou, { key: '', label });
    }
    if (out) return JSON.parse(JSON.stringify(out));
    return { ok: false, msg: 'Choose a role.' };
  });
}

function webEvents_(ss, school, forLeads) {
  return upcomingEvents_(ss, 21, school).filter(e => forLeads || e.who !== WHO.leads).slice(0, 6)
    .map(e => ({ when: eventWhen_(e), date: Utilities.formatDate(e.date, CONFIG.timezone, 'yyyy-MM-dd'), title: e.title, place: e.place || '', type: e.type || '', add: gcalAddLink_(e) }));
}
// Calendar view: expand repeating Events rows across a full year, not just the
// next 21 days / six items used by home widgets. Include the current month.
function webCalendarEvents_(ss, who) {
  const today = new Date(), from = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const until = new Date(today.getFullYear(), today.getMonth() + 13, 0);
  const sh = ss.getSheetByName(TAB.events);
  if (!sh || sh.getLastRow() < 2) return [];
  const out = [], school = who.member && who.member.school;
  const schools = (who.leads || []).concat(school ? [school] : []);
  sh.getRange(2, 1, sh.getLastRow() - 1, EV_EDIT).getValues().forEach(r => {
    const title = String(r[EV.title] || '').trim();
    const audience = String(r[EV.who] || '').trim() || WHO.all;
    if (!title) return;
    // All members = scholars, Campus Leads and RR. Private = RR only.
    const visible = who.admin || audience === WHO.all ||
      (audience === WHO.leads && (who.leads || []).length > 0) ||
      schools.indexOf(audience) !== -1;
    if (!visible) return;
    occurrences_(r, from, until).forEach(date => {
      const e = {date: date, start:r[EV.start], end:r[EV.end], title:title,
        details:String(r[EV.details] || ''), type:String(r[EV.type] || ''),
        place:String(r[EV.place] || ''), who:audience};
      out.push({date:Utilities.formatDate(date, CONFIG.timezone, 'yyyy-MM-dd'),
        when:eventWhen_(e), title:title, place:e.place, type:e.type,
        details:e.details, audience:audience, add:gcalAddLink_(e)});
    });
  });
  out.sort((a,b) => a.date.localeCompare(b.date) || a.title.localeCompare(b.title));
  return out.slice(0, 500);
}
// "Add to Google Calendar" link for one event (works for anyone, no sharing needed)
function gcalAddLink_(e) {
  const pad = n => String(n).padStart(2, '0');
  const day = d => d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate());
  const s = minutesOf_(e.start), en = minutesOf_(e.end);
  let dates;
  if (s === null) { const next = new Date(e.date.getTime() + 86400000); dates = day(e.date) + '/' + day(next); }
  else {
    const t = m => pad(Math.floor(m / 60)) + pad(m % 60) + '00';
    dates = day(e.date) + 'T' + t(s) + '/' + day(e.date) + 'T' + t(en !== null && en > s ? en : Math.min(s + 60, 23 * 60 + 59));
  }
  const q = { action: 'TEMPLATE', text: 'START: ' + e.title, dates, ctz: CONFIG.timezone, location: e.place || '', details: e.details || '' };
  return 'https://calendar.google.com/calendar/render?' + Object.keys(q).filter(k => q[k]).map(k => k + '=' + encodeURIComponent(q[k])).join('&');
}
// The shared START calendars this person can add (only once the calendars are set up)
function webCalendars_(who) {
  try {
    if (!calOn_()) return [];
    const c = calStore_(), out = [];
    if (c.region) out.push({ name: CAL_NAMES.region, url: calLink_(c.region) });
    if (who.member && c.schools && c.schools[who.member.school]) out.push({ name: CAL_NAMES.school(who.member.school), url: calLink_(c.schools[who.member.school]) });
    if ((who.leads.length || who.admin) && c.leads) out.push({ name: CAL_NAMES.leads, url: calLink_(c.leads) });
    if (who.admin && c.mine) out.push({ name: CAL_NAMES.mine, url: calLink_(c.mine) });
    return out;
  } catch (e) { return []; }
}
function webTone_(chip) {
  return chip === CHIP.red ? 'red' : chip === CHIP.yellow ? 'yellow' : chip === CHIP.green ? 'green' : chip === CHIP.grey ? 'grey' : '';
}
function webCell_(v) { return v instanceof Date ? webFmt_(v, 'MMM d') : String(v == null ? '' : v); }

// A member: only their own week, streak, history and their school's events
function webMember_(ss, who) {
  const school = who.member.school;
  let d = null; try { if (checkinOn_()) d = memberData_(ss); } catch (e) { d = null; }
  const me = d ? ((d.bySchool[school] || { members: [] }).members.find(m => m.email === who.email) || {}) : {};
  const history = [];
  const sh = memberTab_(ss);
  if (sh && sh.getLastRow() > 1) {
    const c = memberCols_(sh);
    if (c.email >= 0) sh.getRange(2, 1, sh.getLastRow() - 1, c.width).getValues().forEach(v => {
      if (v[0] instanceof Date && emailOf_(v[c.email]) === who.email)
        history.push({ t: v[0].getTime(), date: webFmt_(v[0], 'EEE, MMM d'), vibe: c.vibe >= 0 ? String(v[c.vibe]) : '',
                       help: c.help >= 0 ? String(v[c.help]) : '' });
    });
  }
  history.sort((a, b) => b.t - a.t);
  const sc = schoolRows_(ss).find(x => x.name === school) || {};
  let url = ''; try { url = WEB_CHECKIN ? '' : (memberForm_() ? memberForm_().getPublishedUrl() : checkinUrl_()); } catch (e) { url = ''; }
  let spec = null;
  return { school, lead: sc.focal ? displayName_(sc.focal) : '', done: !!me.now, streak: me.streak || 0,
           vibe: me.now ? me.now.vibe : '', help: me.now ? me.now.help : '',
           due: d ? dueLabel_(d.week.due) : '', last: d ? isDueDay_(d.week.due) : false, near: d ? memberDeadlineNear_(d.week) : false,
           formUrl: url, open: !!d, checkin: (spec = webCheckinSpec_(ss, who.email)), game: (() => { try { return webGame_(ss, webMyAnswers_(ss, who.email), spec, webMyDaily_(ss, who.email)); } catch (e) { Logger.log('Game: ' + e.message); return null; } })(), history: history.slice(0, 5).map(h => ({ date: h.date, vibe: h.vibe, help: h.help })),
           events: webEvents_(ss, school, false), leadNote: leadNote_(school) };
}

/* ---------- the weekly check-in, answered on the website ----------
 * Same questions as the Google Form (which stays as a backup), same week rules (events page only in weeks
 * with events, DataCamp only while it runs), and the answer goes in the Check-in answers tab, so
 * streaks, alerts, the Campus Lead's view and reports all work the same. */
function webCheckinSpec_(ss, email) {
  try {
    if (!checkinOn_()) return null;
    const now = new Date(), st = getSettings_();
    let w = memberWeek_(now), open = !extraOn_('autoClose') || checkinWindowOpen_(now), late = false, lateUntil = '';
    if (!open) {                                                    // just missed it? It can still be answered for a few days, marked late
      const last = startOfDay_(now);
      for (let i = 0; i < 7 && DAYS[(last.getDay() + 6) % 7] !== st.mDueDay; i++) last.setDate(last.getDate() - 1);
      const days = Math.round((startOfDay_(now) - last) / 86400000);
      if (days >= 1 && days <= WEB.lateDays && email && webAnswerRow_(ss, email, ymd_(last)) < 0) {
        w = memberWeek_(last); open = true; late = true;
        lateUntil = webFmt_(new Date(last.getTime() + WEB.lateDays * 86400000), 'EEEE');
      }
    }
    const events = memberEventChoices_(ss, w);
    const q = [{ k: 'vibe', title: MQ.vibe, type: 'one', choices: M_VIBES }];
    if (events.length) q.push({ k: 'events', title: MQ.events, type: 'many', choices: events.concat([M_MISSED]), none: M_MISSED });
    if (learningOn_(now)) q.push({ k: 'learning', title: MQ.learning, type: 'one', choices: M_LEARN });
    q.push({ k: 'help', title: MQ.help, type: 'one', choices: M_HELP });
    q.push({ k: 'tea', title: MQ.tea, type: 'text' });
    return { open, late, lateUntil, opens: st.mOpenDay, week: ymd_(w.due), dueAt: w.due.getTime(),
             due: dueLabel_(w.due), last: !late && isDueDay_(w.due), q, consent: CONSENT_AGREE, privacy: privacyNote_() };
  } catch (e) { Logger.log('Website check-in: ' + e.message); return null; }
}

var LATE_HEAD = 'Answered late on';
// Row number of this person's answer for a week (key = due date yyyy-MM-dd), or -1
function webAnswerRow_(ss, email, weekKey, sh, c) {
  sh = sh || memberTab_(ss);
  if (!sh || sh.getLastRow() < 2) return -1;
  c = c || memberCols_(sh);
  if (c.email < 0) return -1;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, c.width).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    const v = vals[i];
    if (v[0] instanceof Date && emailOf_(v[c.email]) === email && memberWeekKey_(v[0]) === weekKey) return i + 2;
  }
  return -1;
}

/* ---------- 🧩 daily brain breaks (website Journey) ----------
 * Every day has one thing to do. The check-in day (Settings › "Weekly check-in opens") is the weekly check-in, worth the
 * most XP. Every other day is ONE random brain break: a word scramble (one hard word, one try, against the clock),
 * odd one out or a number pattern (three hard rounds, timed). Questions follow the scholar's own program (Scholars › C),
 * with Davao Region rounds mixed in. Days cannot be replayed: a game counts only on its own day, opening it starts it,
 * and leaving before the end scores 0. The page shows the game; THIS file checks every answer and decides the XP.
 * Saved in the "Daily games" tab. XP from games and check-ins also funds reward requests. */
var MINI = { xpScramble: 10, xpRound: 3, perfect: 1, rounds: 3, hearts: 1, programRounds: 2, secsScramble: 50, secsRound: 25, slackMs: 4000 };   // +5 s each (generous)   // slackMs: time for a tap to travel to START on slow data
// Number patterns (logic, for every program): [question, four choices, index of the answer, why]
var MINI_PATTERNS = [
  ['What comes next? 2, 4, 8, 16, …', ['24', '32', '30', '64'], 1, 'Each number doubles.'],
  ['What comes next? 3, 7, 11, 15, …', ['18', '19', '20', '22'], 1, 'Add 4 each time.'],
  ['What comes next? 1, 4, 9, 16, …', ['24', '25', '26', '36'], 1, 'These are square numbers.'],
  ['What comes next? 2, 6, 12, 20, …', ['28', '30', '32', '40'], 1, 'Add 4, then 6, then 8, then 10.'],
  ['What comes next? 81, 27, 9, 3, …', ['0', '1', '2', '6'], 1, 'Divide by 3 each time.'],
  ['What comes next? 5, 10, 20, 35, …', ['45', '50', '55', '60'], 2, 'Add 5, 10, 15, then 20.'],
  ['What comes next? 1, 1, 2, 3, 5, …', ['6', '7', '8', '10'], 2, 'Each number is the sum of the two before it.'],
  ['What comes next? 100, 90, 81, 73, …', ['64', '65', '66', '67'], 2, 'Subtract 10, 9, 8, then 7.'],
  ['What comes next? 2, 3, 5, 8, 12, …', ['15', '16', '17', '18'], 2, 'Add 1, 2, 3, 4, then 5.'],
  ['What comes next? 4, 7, 13, 22, …', ['31', '32', '34', '36'], 2, 'Add 3, 6, 9, then 12.'],
  ['What comes next? 2, 3, 5, 7, 11, …', ['12', '13', '15', '17'], 1, 'These are prime numbers.'],
  ['What comes next? 1, 8, 27, 64, …', ['100', '125', '128', '81'], 1, 'These are cubes: 1³, 2³, 3³, 4³, 5³.'],
];
var MINI_DAVAO = {
  words: [['KADAYAWAN', 'Davao City\'s thanksgiving festival, every August'], ['HAMIGUITAN', 'Mountain range in Davao Oriental, a UNESCO World Heritage Site'],
          ['MANGOSTEEN', 'Purple "queen of fruits" grown in Davao'], ['NABUNTURAN', 'Capital town of Davao de Oro'],
          ['MALAGOS', 'Davao City garden resort known for its chocolate'], ['MINDANAO', 'The big island Davao is on']],
  odd: [['Which is NOT a Davao Region city?', ['Tagum', 'Panabo', 'Mati', 'Valencia'], 3, 'Valencia City is in Bukidnon. The others are Davao Region cities.'],
        ['Which mountain is NOT in Davao Region?', ['Mt. Apo', 'Mt. Hamiguitan', 'Mt. Talomo', 'Mt. Kitanglad'], 3, 'Mt. Kitanglad is in Bukidnon.'],
        ['Which island is NOT part of Samal or nearby?', ['Talikud', 'Malipano', 'Siargao', 'Samal'], 2, 'Siargao is in Surigao del Norte.'],
        ['Which is NOT a Davao Region provincial capital?', ['Mati', 'Nabunturan', 'Malita', 'Kidapawan'], 3, 'Kidapawan is the capital of Cotabato (Region XII).'],
        ['Which people are NOT native to Davao Region?', ['Bagobo', 'Mandaya', 'Mansaka', 'Ifugao'], 3, 'The Ifugao are from the Cordillera.'],
        ['Which place is NOT in Davao?', ['Eden Nature Park', 'Chocolate Hills', 'Philippine Eagle Center', 'People\'s Park'], 1, 'The Chocolate Hills are in Bohol.'],
        ['Which is NOT a fruit?', ['Marang', 'Pomelo', 'Ampalaya', 'Durian'], 2, 'Ampalaya (bitter gourd) is eaten as a vegetable.']],
};
// Program questions, shared between programs that study the same things
var MQ_ = {
  code: { words: [['ALGORITHM', 'Step-by-step recipe a computer follows'], ['RECURSION', 'A function that calls itself'], ['COMPILER', 'Turns source code into machine code']],
          odd: [['Which is NOT a sorting algorithm?', ['Merge sort', 'Quicksort', 'Heap sort', 'Binary search'], 3, 'Binary search finds an item in a sorted list; it does not sort.'],
                ['Which is NOT a data structure?', ['Stack', 'Queue', 'Linked list', 'Compiler'], 3, 'A compiler is a program that translates code.'],
                ['Which is NOT a programming language?', ['Python', 'Java', 'Rust', 'Linux'], 3, 'Linux is an operating system kernel.']] },
  net: { words: [['FIREWALL', 'Filters the traffic going in and out of a network'], ['BANDWIDTH', 'How much data a link can carry each second'], ['ENCRYPTION', 'Scrambling data so only the right person can read it']],
         odd: [['Which port is NOT usually for websites?', ['80', '443', '8080', '22'], 3, 'Port 22 is for SSH (remote login).'],
               ['Which is NOT a web browser?', ['Chrome', 'Firefox', 'Safari', 'Apache'], 3, 'Apache is a web server.'],
               ['Which is NOT a social engineering attack?', ['Phishing', 'Pretexting', 'Baiting', 'Defragmenting'], 3, 'Defragmenting tidies a hard drive.'],
               ['Which is NOT an operating system?', ['Windows', 'Linux', 'macOS', 'MySQL'], 3, 'MySQL is a database system.']] },
  sys: { words: [['DATABASE', 'An organised store of records'], ['FLOWCHART', 'A diagram of steps and decisions'], ['STAKEHOLDER', 'Anyone affected by a new system']],
         odd: [['Which is NOT a phase of the SDLC?', ['Planning', 'Analysis', 'Design', 'Marketing'], 3, 'The SDLC covers planning, analysis, design, building, testing and upkeep.'],
               ['Which is NOT a database system?', ['MySQL', 'PostgreSQL', 'Oracle Database', 'Photoshop'], 3, 'Photoshop edits images.'],
               ['Which is NOT a systems analysis diagram?', ['ERD', 'DFD', 'Use case diagram', 'CPU'], 3, 'A CPU is computer hardware.']] },
  biz: { words: [['DASHBOARD', 'One screen of key numbers for managers'], ['FORECAST', 'An estimate of future sales or demand'], ['INVENTORY', 'The stock a business keeps on hand']],
         odd: [['Which is NOT a business (ERP) system?', ['SAP', 'Odoo', 'Oracle NetSuite', 'Figma'], 3, 'Figma is a design tool.'],
               ['Which is NOT one of the lean wastes?', ['Waiting', 'Overproduction', 'Defects', 'Ergonomics'], 3, 'Ergonomics fits work to people; it is not a waste.'],
               ['Which is NOT a quality method?', ['Six Sigma', 'Kaizen', 'Lean', 'Kerberos'], 3, 'Kerberos is a network login protocol.']] },
  teach: { words: [['PEDAGOGY', 'The method and practice of teaching'], ['CURRICULUM', 'Everything a course or school sets out to teach'], ['ASSESSMENT', 'Checking what learners know']],
           odd: [['Which is NOT a domain in Bloom\'s taxonomy?', ['Cognitive', 'Affective', 'Psychomotor', 'Binary'], 3, 'Bloom\'s three domains are cognitive, affective and psychomotor.'],
                 ['Which is NOT an input device?', ['Keyboard', 'Mouse', 'Scanner', 'Monitor'], 3, 'A monitor is an output device.']] },
  hw: { words: [['FIRMWARE', 'Software built into a piece of hardware'], ['REGISTER', 'Tiny, very fast memory inside a CPU'], ['TRANSISTOR', 'The tiny switch inside every chip']],
        odd: [['Which is NOT a logic gate?', ['AND', 'NOR', 'XOR', 'JUMP'], 3, 'JUMP is a CPU instruction, not a gate.'],
              ['Which is NOT a kind of memory?', ['RAM', 'ROM', 'Cache', 'ALU'], 3, 'The ALU does the arithmetic.'],
              ['Which is NOT a microcontroller board?', ['Arduino Uno', 'ESP32', 'Raspberry Pi Pico', 'Ubuntu'], 3, 'Ubuntu is an operating system.']] },
  ece: { words: [['CAPACITOR', 'Stores charge between two plates'], ['FREQUENCY', 'Cycles per second, measured in hertz'], ['OSCILLATOR', 'A circuit that makes a repeating signal']],
         odd: [['Which is NOT a passive component?', ['Resistor', 'Capacitor', 'Inductor', 'Transistor'], 3, 'A transistor is an active component.'],
               ['Which is NOT an electrical unit?', ['Ohm', 'Farad', 'Henry', 'Kelvin'], 3, 'Kelvin measures temperature.'],
               ['Which is NOT a type of modulation?', ['AM', 'FM', 'PSK', 'USB'], 3, 'USB is a port and cable standard.']] },
  mech: { words: [['ACTUATOR', 'Turns a control signal into motion'], ['ROBOTICS', 'Designing machines that sense and move'], ['KINEMATICS', 'The study of motion without its forces']],
          odd: [['Which is NOT an actuator?', ['DC motor', 'Servo', 'Solenoid', 'Thermistor'], 3, 'A thermistor is a sensor.'],
                ['Which is NOT a sensor?', ['Ultrasonic sensor', 'Gyroscope', 'Encoder', 'Gearbox'], 3, 'A gearbox passes on motion; it does not sense.'],
                ['Which is NOT a PID controller term?', ['Proportional', 'Integral', 'Derivative', 'Digital'], 3, 'PID stands for proportional, integral, derivative.']] },
  math: { words: [['DERIVATIVE', 'The rate of change of a function'], ['INTEGRAL', 'The area under a curve'], ['EIGENVALUE', 'The lambda in Av = λv']],
          odd: [['Which is NOT a prime number?', ['2', '17', '21', '29'], 2, '21 = 3 × 7.'],
                ['Which is NOT a root-finding method?', ['Newton-Raphson', 'Bisection', 'Secant', 'Dijkstra'], 3, 'Dijkstra\'s algorithm finds shortest paths in a graph.'],
                ['Which is NOT a conic section?', ['Ellipse', 'Parabola', 'Hyperbola', 'Cycloid'], 3, 'A cycloid is traced by a point on a rolling circle.']] },
  stat: { words: [['REGRESSION', 'Fitting a line to predict one variable from another'], ['VARIANCE', 'The average squared distance from the mean'], ['HYPOTHESIS', 'A claim that is tested with data']],
          odd: [['Which is NOT an "average"?', ['Mean', 'Variance', 'Median', 'Mode'], 1, 'Variance measures spread, not the centre.'],
                ['Which is NOT a sampling method?', ['Stratified', 'Cluster', 'Systematic', 'Regression'], 3, 'Regression models relationships; the others choose a sample.'],
                ['Which is NOT a measure of spread?', ['Range', 'Standard deviation', 'Interquartile range', 'Median'], 3, 'The median is a measure of the centre.'],
                ['Which is NOT a probability distribution?', ['Normal', 'Poisson', 'Binomial', 'Histogram'], 3, 'A histogram is a chart of data.']] },
};
// Each of the 13 START programs (START_PROGRAMS) -> the question sets it draws from
var MINI_PROGRAMS = {
  'Applied Mathematics': ['math', 'stat'],
  'Technology and Livelihood Education with Specialization in Information and Communication Technology': ['teach', 'net'],
  'Computer Engineering': ['hw', 'code'],
  'Computer Science': ['code', 'math'],
  'Electronics and Communication Engineering': ['ece', 'hw'],
  'Industrial Management Engineering – Information Technology': ['biz', 'sys'],
  'Information and Communications Technology': ['net', 'code'],
  'Information Systems': ['sys', 'biz'],
  'Information Technology': ['net', 'sys'],
  'Information Technology Systems': ['sys', 'net'],
  'Management Information Systems': ['biz', 'sys'],
  'Statistics': ['stat', 'math'],
  'Mechatronics Engineering': ['mech', 'hw'],
};

// MEMORY PAIRS (icons only): two cards with the same picture make a pair. A board = 6 pairs: 2 icons from each of the
// scholar's two topics and 2 from Davao (6 Davao icons when the program is not known). The pictures are drawn in Web_App.
var MEMORY = { pairs: 6, perTopic: 2, previewSecs: 5, secs: 80, slackMs: 4000, xpPair: 1, tidyBonus: 4, okBonus: 2 };
var MEMI_ = { code: ['braces', 'window', 'hash'], net: ['wifi', 'globe', 'cloud'], sys: ['db', 'lock', 'shield'], biz: ['bars', 'briefcase', 'target'],
  teach: ['book', 'bulb', 'cap'], hw: ['cpu', 'battery', 'monitor'], ece: ['tower', 'plug', 'magnet'], mech: ['gear', 'robot', 'nut'],
  math: ['calc', 'infinity', 'cube'], stat: ['bell', 'dice', 'scale'], davao: ['durian', 'mountain', 'eagle', 'orchid', 'island', 'sunny'] };
// Stars: 3 = every pair in at most pairs+2 moves, 2 = every pair, 1 = half or more. XP follows the stars (DAILY_XP).
function memScore_(found, moves) {
  const P = MEMORY.pairs, all = found >= P;
  const bonus = all ? (moves <= P + 2 ? MEMORY.tidyBonus : moves <= P + 5 ? MEMORY.okBonus : 0) : 0;
  const stars = all ? (moves <= P + 2 ? 3 : 2) : found >= P / 2 ? 1 : 0; void bonus;
  return ['Memory pairs · ' + found + ' of ' + P + ' pairs · ' + moves + ' moves', dailyXp_(stars), stars];
}
function memPairOf_(game, cid) { const c = game.cards[cid]; return c ? c.pair : -1; }

// A private tracker secret keeps the answer sequence unpredictable even if the source is shared.
// The calls that first create a game and later check it run under the document lock.
function miniSeed_(email, week) {
  const props = docProps_();
  let secret = props.getProperty('MINI_SECRET');
  if (!secret) { secret = Utilities.getUuid(); props.setProperty('MINI_SECRET', secret); }
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    'mini|' + secret + '|' + emailOf_(email) + '|' + week).map(b => b & 255);
}
// The scholar's degree program from Scholars (column C), '' when not found
function miniProgram_(ss, email) {
  try {
    const sh = ss.getSheetByName(TAB.scholars);
    if (!sh || sh.getLastRow() < 2) return '';
    const row = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().find(r => emailOf_(r[4]) === emailOf_(email));
    return row ? String(row[2] || '').trim() : '';
  } catch (e) { return ''; }
}
// The questions for a program: { words, odd } (Davao Region when the program is not one of the 13)
function miniBank_(program) {
  const sets = MINI_PROGRAMS[String(program || '').trim()];
  if (!sets) return { words: MINI_DAVAO.words, odd: [], own: false };
  return { words: [].concat.apply([], sets.map(k => MQ_[k].words)), odd: [].concat.apply([], sets.map(k => MQ_[k].odd)), own: true };
}
// The full game stays on the server. Only miniPublic_ sends the current prompt to a browser.
function webMiniFor_(email, week, program) {
  if (!email || !week) return null;
  const b = miniSeed_(email, week), bank = miniBank_(program);
  const id = week + '-' + b[1] + b[2] + (bank.own ? '-' + Object.keys(MINI_PROGRAMS).indexOf(String(program).trim()) : '');
  const shuffle = (list, at) => {                          // a fixed shuffle from the seed
    const a = list.slice();
    for (let i = a.length - 1; i > 0; i--) { const j = b[(at + i) % b.length] % (i + 1); const t = a[i]; a[i] = a[j]; a[j] = t; }
    return a;
  };
  const kind = b[0] % 4 === 3 ? 'memory' : ['scramble', 'odd', 'pattern'][b[0] % 3];
  if (kind === 'memory') {
    const sets = MINI_PROGRAMS[String(program || '').trim()];
    let icons = [];
    if (sets) sets.forEach((k, n) => { icons = icons.concat(shuffle(MEMI_[k], 5 + n * 4).slice(0, MEMORY.perTopic).filter(x => icons.indexOf(x) === -1)); });
    icons = icons.concat(shuffle(MEMI_.davao, 9).filter(x => icons.indexOf(x) === -1).slice(0, MEMORY.pairs - icons.length));
    const cards = [];
    icons.forEach((ic, i) => { cards.push({ kind: 'icon', icon: ic, pair: i }); cards.push({ kind: 'icon', icon: ic, pair: i }); });
    const order = shuffle(cards.map((c, i) => i), 13);                // card ids say nothing about the pairs
    return { id, kind, cards: order.map(i => cards[i]), pairs: MEMORY.pairs, xp: MEMORY.pairs * MEMORY.xpPair + MEMORY.tidyBonus,
             secs: MEMORY.secs, preview: MEMORY.previewSecs, topic: sets ? String(program) : 'Davao Region' };
  }
  if (kind === 'scramble') {
    const w = bank.words[b[3] % bank.words.length], word = w[0];
    let letters = shuffle(word.split(''), 4);
    if (letters.join('') === word) letters = letters.slice(1).concat(letters[0]);
    return { id, kind: 'scramble', letters, clue: w[1], answer: word, hearts: MINI.hearts, xp: MINI.xpScramble, secs: MINI.secsScramble, topic: bank.own ? String(program) : 'Davao Region' };
  }
  const pack = (o, r, topic) => { const idx = shuffle([0, 1, 2, 3], 8 + r * 3);
    return { title: o[0], choices: idx.map(i => o[1][i]), odd: idx.indexOf(o[2]), why: o[3], topic }; };
  const points = { xp: MINI.xpRound * MINI.rounds + MINI.perfect, each: MINI.xpRound, perfect: MINI.perfect, secs: MINI.secsRound };
  if (kind === 'pattern') {
    const rounds = shuffle(MINI_PATTERNS.map((x, i) => i), 6).slice(0, MINI.rounds).map((i, r) => pack(MINI_PATTERNS[i], r, 'Logic'));
    return Object.assign({ id, kind, rounds }, points);
  }
  // odd one out: two rounds about their program, one about Davao (three Davao rounds when the program is not known)
  const own = shuffle(bank.odd.map((x, i) => i), 6).slice(0, bank.own ? MINI.programRounds : 0).map(i => bank.odd[i]);
  const davao = shuffle(MINI_DAVAO.odd.map((x, i) => i), 11).slice(0, MINI.rounds - own.length).map(i => MINI_DAVAO.odd[i]);
  const rounds = own.concat(davao).map((o, r) => pack(o, r, r < own.length ? String(program) : 'Davao Region'));
  return Object.assign({ id, kind, rounds }, points);
}
// What the scholar sent -> [text for the sheet, XP, stars 0-3]. Anything that does not match this game scores 0.
function miniResult_(game, sent) {
  if (!game) return ['', 0, 0];
  const name = { scramble: 'Word scramble', odd: 'Odd one out', pattern: 'Number pattern' }[game.kind] || 'Brain break';
  if (!sent) return [name + ' · not played', 0, 0];
  if (sent.id !== game.id || sent.skip) return [name + ' · skipped', 0, 0];
  if (game.kind === 'scramble') {
    const tries = Math.max(1, Math.floor(Number(sent.tries) || 1));
    const ok = String(sent.answer || '').toUpperCase().replace(/[^A-Z]/g, '') === game.answer && tries <= MINI.hearts;
    return ok ? [name + ' · ' + game.answer + ' · solved', MINI.xpScramble, 3] : [name + ' · ' + game.answer + ' · not solved', 0, 0];
  }
  const picks = Array.isArray(sent.picks) ? sent.picks : [];
  const right = game.rounds.filter((r, i) => String(picks[i]) === r.choices[r.odd]).length;
  return [name + ' · ' + right + ' of ' + game.rounds.length, right * MINI.xpRound + (right === game.rounds.length ? MINI.perfect : 0), right];
}

var DAILY_TAB = 'Daily games';
var DAILY_HEAD = ['Day', 'Email', 'Name', 'School', 'Game', 'Result', 'XP', 'Stars', 'Started', 'Finished', 'Game ID',
                  'Round', 'Correct', 'Prompt at', 'Round state', 'Last pick', 'Was correct', 'Timed out'];
var DAILY_GRACE_MIN = 20;                      // a game must be finished within this many minutes of opening it
function dailyTab_(ss) {
  let sh = ss.getSheetByName(DAILY_TAB);
  if (!sh) { sh = ss.insertSheet(DAILY_TAB); sh.setFrozenRows(1); sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@'); }
  if (sh.getMaxColumns() < DAILY_HEAD.length) sh.insertColumnsAfter(sh.getMaxColumns(), DAILY_HEAD.length - sh.getMaxColumns());
  if (sh.getRange(1, DAILY_HEAD.length).getValue() !== DAILY_HEAD[DAILY_HEAD.length - 1])
    sh.getRange(1, 1, 1, DAILY_HEAD.length).setValues([DAILY_HEAD]).setFontWeight('bold');
  return sh;
}
function miniCount_(game) { return game.kind === 'scramble' ? 1 : game.kind === 'memory' ? game.pairs : game.rounds.length; }
// The phone checks a tap at once, so it gets each answer in a coded form (not readable at a glance).
// The server still checks every answer and its time again, and only the server's result gives XP.
function miniSeal_(key, text) {
  key = String(key); text = String(text);
  let out = '';
  for (let i = 0; i < text.length; i++) {
    const k = (key.charCodeAt(i % key.length) * 31 + i * 17 + 7) & 0xffff;
    out += ('000' + ((text.charCodeAt(i) ^ k) & 0xffff).toString(16)).slice(-4);
  }
  return out;
}
function miniPrompt_(game, round) {
  if (game.kind === 'scramble') return { letters: game.letters.slice(), clue: game.clue, length: game.answer.length, s: miniSeal_(game.id, game.answer) };
  const r = game.rounds[round];
  return { title: r.title, choices: r.choices.slice(), topic: r.topic, s: miniSeal_(game.id + ':' + round, JSON.stringify({ o: r.odd, w: r.why || '' })) };
}
function miniPublic_(game, at, promptAt) {
  if (game.kind === 'memory') return { id: game.id, kind: 'memory', xp: game.xp, secs: game.secs, preview: game.preview, pairs: game.pairs, topic: game.topic,
    cards: game.cards.map((c, i) => ({ cid: i, kind: 'icon', icon: c.icon })), remainingMs: (game.preview + game.secs) * 1000,
    deadline: promptAt.getTime() + (game.preview + game.secs) * 1000 };
  const out = { id: game.id, kind: game.kind, xp: game.xp, secs: game.secs,
    each: game.each, perfect: game.perfect, hearts: game.hearts,
    deadline: promptAt.getTime() + game.secs * 1000, remainingMs: game.secs * 1000 };
  if (game.kind === 'scramble') Object.assign(out, miniPrompt_(game, 0));
  else out.rounds = game.rounds.map((r, i) => miniPrompt_(game, i));     // every round now, so "Next round" opens at once
  return out;
}
// what the phone shows for a catch-up: up to 4 XP (rounds: 1 each, +1 for all three)
function catchupPublic_(pub) { return Object.assign(pub, { catchup: true, xp: catchupXp_(pub.xp), each: pub.kind === 'scramble' || pub.kind === 'memory' ? pub.each : 1, perfect: pub.kind === 'scramble' || pub.kind === 'memory' ? 0 : 1 }); }
function miniScore_(game, right) {
  const name = { scramble: 'Word scramble', odd: 'Odd one out', pattern: 'Number pattern', memory: 'Memory pairs' }[game.kind] || 'Brain break';
  if (game.kind === 'memory') return [name, 0, 0];
  if (game.kind === 'scramble') return [name + ' · ' + (right ? 'solved' : 'not solved'), right ? MINI.xpScramble : 0, right ? 3 : 0];
  return [name + ' · ' + right + ' of ' + miniCount_(game), right * MINI.xpRound + (right === miniCount_(game) ? MINI.perfect : 0), right];
}
function miniFeedback_(game, round, correct, right, late) {
  const item = game.kind === 'scramble' ? null : game.rounds[round];
  return { ok: true, round, correct, correctTotal: right, timeUp: late,
    answer: item ? '' : game.answer, correctIndex: item ? item.odd : -1, why: item ? item.why : '',
    last: round === miniCount_(game) - 1,
    score: round === miniCount_(game) - 1 ? miniScore_(game, right)[1] : 0 };
}
// One scholar's daily games: { byDay: { 'yyyy-MM-dd': { row, kind, result, xp, stars, started, finished, id } }, total, wins }
function webMyDaily_(ss, email) {
  const out = { byDay: {}, total: 0, wins: 0, email: emailOf_(email) };
  const sh = ss.getSheetByName(DAILY_TAB);
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, DAILY_HEAD.length).getDisplayValues().forEach((r, i) => {
    if (emailOf_(r[1]) !== emailOf_(email) || !r[0]) return;
    const xp = Math.max(0, Math.min(MINI.xpScramble, Number(r[6]) || 0)), stars = Number(r[7]) || 0;
    out.byDay[r[0]] = { row: i + 2, kind: r[4], result: r[5], xp, stars, started: r[8], finished: r[9], id: r[10], catchup: isCatchupRow_(r) };
    if (r[9]) { out.total += xp; if (stars >= 2) out.wins++; }
  });
  return out;
}
function dayKey_(d) { return ymd_(d || new Date()); }
// Catch-up: a brain-break day that was never opened can still be played later (on the Journey map), for up to 4 XP (DAILY_XP.catchup).
// It never mends the day streak. Check-in days can never be caught up: a missed or overdue check-in stays missed.
var CATCHUP = { share: 0.3, label: 'catch-up' };
function catchupXp_(xp) { xp = Number(xp) || 0; let st = 0; DAILY_XP.today.forEach((v, i) => { if (xp >= v && v > 0) st = i; }); return DAILY_XP.catchup[st]; }
function isCatchupRow_(r) { return new RegExp(CATCHUP.label).test(String(r[5] || '')); }
// a past game day that may be caught up: shown on the map, not a check-in day, not before the term, never opened
function catchupDayOk_(ss, email, key, st) {
  if (!/^\d{4}-\d\d-\d\d$/.test(String(key || ''))) return 'Pick a day on your Journey map.';
  const today = startOfDay_(new Date()), d = new Date(key + 'T12:00:00'), oldest = new Date(today); oldest.setDate(oldest.getDate() - GAME.mapBack);
  if (!(key < dayKey_(today))) return 'That day is not over yet.';
  if (startOfDay_(d) < oldest) return 'That day is too far back to catch up.';
  if (isCheckinDay_(d, st)) return 'Check-in days cannot be caught up.';
  let ts = null; try { ts = termStart_(ss); } catch (e) { ts = null; }
  if (ts instanceof Date && startOfDay_(d) < startOfDay_(ts)) return 'That day is before this term.';
  return '';
}
// a game row can still be answered: today's, or a catch-up opened today
function dailyPlayable_(key, r) { return key === dayKey_() || (isCatchupRow_(r) && r[8] instanceof Date && dayKey_(r[8]) === dayKey_()); }
// The check-in deadline is 11:59 PM of the due day: "Sunday, Oct 25, 11:59 PM", or "today, 11:59 PM" on the day itself
function isDueDay_(d) { return d instanceof Date && ymd_(d) === ymd_(new Date()); }
function dueLabel_(d) { return d instanceof Date ? (isDueDay_(d) ? 'today' : webFmt_(d, 'EEEE, MMM d')) + ', 11:59 PM' : ''; }

// "Done" ticks on to-do items: kept for the current week only (Campus Lead per school, RR for the region)
function doneList_(key, week) {
  try { const v = JSON.parse(docProps_().getProperty(key) || 'null'); return v && v.w === week ? v.k || [] : []; } catch (e) { return []; }
}
function doneToggle_(key, week, item, undo) {
  const list = doneList_(key, week).filter(k => k !== item);
  if (!undo) list.push(item);
  docProps_().setProperty(key, JSON.stringify({ w: week, k: list.slice(-60) }));
  return list;
}
// Brain-break stars per school this week (Monday to today): a tie-breaker on the school leaderboard
function dailyStarsBySchool_(ss) {
  const sh = ss.getSheetByName(DAILY_TAB), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  const mon = dayKey_(mondayOf_(new Date()));
  sh.getRange(2, 1, sh.getLastRow() - 1, DAILY_HEAD.length).getDisplayValues().forEach(r => {
    if (r[0] >= mon && r[9]) out[r[3]] = (out[r[3]] || 0) + (Number(r[7]) || 0); });
  return out;
}
function isCheckinDay_(d, st) { return DAYS[(d.getDay() + 6) % 7] === (st || getSettings_()).mOpenDay; }
// The Journey map: the last days, today and the next few. One tile a day: the check-in day, or a brain break.
function webDayMap_(ss, byWeek, daily, spec) {
  let program = '';
  const st = getSettings_(), today = startOfDay_(new Date()), todayKey = dayKey_(today), out = [];
  let ts = null; try { ts = termStart_(ss); } catch (e) { ts = null; }
  const weekDue = d => { const x = new Date(d); for (let i = 0; i < 7 && DAYS[(x.getDay() + 6) % 7] !== st.mDueDay; i++) x.setDate(x.getDate() + 1); return dayKey_(x); };
  for (let i = -GAME.mapBack; i <= GAME.mapAhead; i++) {
    const d = new Date(today); d.setDate(d.getDate() + i);
    if (ts instanceof Date && d < startOfDay_(ts)) continue;
    const key = dayKey_(d), label = webFmt_(d, 'EEE'), date = webFmt_(d, 'MMM d');
    if (isCheckinDay_(d, st)) {
      const wk = weekDue(d), r = byWeek[wk];
      const open = !!(spec && spec.open && spec.week === wk && !r && key <= todayKey);
      out.push({ key, label, date, kind: 'checkin', state: r ? 'done' : open ? 'today' : key > todayKey ? 'locked' : 'missed',
                 stars: r ? (r.late ? 2 : 3) : 0, xp: r ? GAME.xpCheckin + (r.late ? 0 : GAME.xpOnTime) + GAME.xpEvent * r.ev : 0, max: GAME.xpCheckin + GAME.xpOnTime, today: key === todayKey });
    } else {
      const g = daily.byDay[key], done = !!(g && g.finished);
      let game = g && /^(scramble|odd|pattern|memory)$/.test(g.kind) ? g.kind : '';
      if (!game && daily.email) { try { if (!program) program = miniProgram_(ss, daily.email); const m = webMiniFor_(daily.email, key, program); game = m ? m.kind : ''; } catch (e) { game = ''; } }
      const past = key < todayKey, canCatch = past && !g;              // never opened: it can be caught up (a game opened and left stays missed)
      out.push({ key, label, date, kind: 'game', state: done ? 'done' : key === todayKey ? (g ? 'missed' : 'today') : key > todayKey ? 'locked' : canCatch ? 'catchup' : 'missed',
                 stars: done ? g.stars : 0, xp: done ? g.xp : 0, max: canCatch ? catchupXp_(MINI.xpScramble) : MINI.xpScramble, today: key === todayKey,
                 caught: !!(done && g.catchup), game });
    }
  }
  return out;
}
// Days in a row with something done (the brain break, or the check-in on check-in day), and this week's recap
function webDayStreak_(byWeek, daily) {
  const st = getSettings_(), today = startOfDay_(new Date()), todayKey = dayKey_(today);
  const weekDue = d => { const x = new Date(d); for (let i = 0; i < 7 && DAYS[(x.getDay() + 6) % 7] !== st.mDueDay; i++) x.setDate(x.getDate() + 1); return dayKey_(x); };
  const doneOn = d => isCheckinDay_(d, st) ? !!byWeek[weekDue(d)] : !!(daily.byDay[dayKey_(d)] && daily.byDay[dayKey_(d)].finished && !daily.byDay[dayKey_(d)].catchup);   // catch-ups never mend the streak
  let streak = 0; const d = new Date(today);
  if (!doneOn(d)) d.setDate(d.getDate() - 1);                     // today not played yet: the streak is still alive
  for (let i = 0; i < 366 && doneOn(d); i++) { streak++; d.setDate(d.getDate() - 1); }
  const mon = mondayOf_(today), recap = { xp: 0, stars: 0, days: 0, of: Math.round((today - mon) / 86400000) + 1 };
  for (let x = new Date(mon); x <= today; x.setDate(x.getDate() + 1)) {
    const key = dayKey_(x), g = daily.byDay[key];
    if (isCheckinDay_(x, st)) { const r = byWeek[weekDue(x)]; if (r) { recap.days++; recap.stars += r.late ? 2 : 3; recap.xp += GAME.xpCheckin + (r.late ? 0 : GAME.xpOnTime) + GAME.xpEvent * r.ev; } }
    else if (g && g.finished) { recap.days++; recap.stars += g.stars; recap.xp += g.xp; }
  }
  return { streak, recap, today: todayKey };
}
function dailyRowFor_(sh, email, key) {
  if (sh.getLastRow() < 2) return -1;
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues();
  for (let i = v.length - 1; i >= 0; i--) if (v[i][0] === key && emailOf_(v[i][1]) === email) return i + 2;
  return -1;
}
// Open today's brain break. Opening it starts it: a game opened before and not finished is lost (no replays).
function webDailyStart(token, sent) {
  return webSafe_(() => {
    const s = webSession_(token);
    if (!s) return { ok: false, signedOut: true };
    const ss = webSheet_(), who = webRoles_(ss, s.e);
    if (!who.member) return { ok: false, msg: 'Daily brain breaks are for active START scholars.' };
    const now = new Date(), today = dayKey_(now), want = sent && sent.day ? String(sent.day) : today, catchup = want !== today;
    if (catchup) { const why = catchupDayOk_(ss, who.email, want); if (why) return { ok: false, msg: why }; }
    else if (isCheckinDay_(startOfDay_(now))) return { ok: false, checkin: true, msg: 'Today is check-in day: your check-in is today\'s challenge.' };
    const key = want;
    const sh = dailyTab_(ss), lock = docLock_();
    if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try again.' };
    try {
      const row = dailyRowFor_(sh, who.email, key);
      if (row > 0) {
        const r = sh.getRange(row, 1, 1, DAILY_HEAD.length).getDisplayValues()[0];
        if (!r[9]) { sh.getRange(row, 6, 1, 5).setValues([[r[4] + ' · left before the end' + (isCatchupRow_(r) ? ' · ' + CATCHUP.label : ''), 0, 0, r[8], new Date()]]); }
        return { ok: true, over: true, msg: catchup ? 'That day was already played.' : 'Today\'s brain break is already used. A new one opens tomorrow.' };
      }
      const game = webMiniFor_(who.email, key, miniProgram_(ss, who.email)), started = new Date();
      const mem = game.kind === 'memory';                              // memory: moves, pairs found, when play starts (after the preview), matched cards
      sh.appendRow([key, who.email, who.member.name || who.name || '', who.member.school || '', game.kind, 'Started' + (catchup ? ' · ' + CATCHUP.label : ''), 0, 0,
                    started, '', game.id, 0, 0, mem ? new Date(started.getTime() + game.preview * 1000) : started, mem ? 'memory' : 'prompt', '', '', '']);
      sh.getRange(sh.getLastRow(), 1).setNumberFormat('@').setValue(key);
      return { ok: true, day: key, game: catchup ? catchupPublic_(miniPublic_(game, 0, started)) : miniPublic_(game, 0, started) };
    } finally { lock.releaseLock(); }
  });
}
// Each pick is checked once, under a lock, against the prompt time stored in the sheet.
function webDailyPick(token, sent) {
  return webSafe_(() => {
    const s = webSession_(token);
    if (!s) return { ok: false, signedOut: true };
    const ss = webSheet_(), who = webRoles_(ss, s.e);
    if (!who.member) return { ok: false, msg: 'Daily brain breaks are for active START scholars.' };
    sent = sent || {};
    const key = String(sent.day || ''), sh = dailyTab_(ss), lock = docLock_();
    if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try again.' };
    try {
      const row = dailyRowFor_(sh, who.email, key);
      if (row < 0) return { ok: false, msg: 'Open today\'s brain break first.' };
      const r = sh.getRange(row, 1, 1, DAILY_HEAD.length).getValues()[0];
      if (!dailyPlayable_(key, r)) return { ok: false, msg: 'This brain break has closed.' };
      const cu = x => isCatchupRow_(r) && x && x.ok ? Object.assign(x, { score: catchupXp_(x.score), catchup: true }) : x;
      const game = webMiniFor_(who.email, key, miniProgram_(ss, who.email));
      const round = Number(r[11]), at = r[13], now = new Date();
      if (r[9] || sent.id !== r[10] || r[10] !== game.id ||
          Number(sent.round) !== round || round < 0 || round >= miniCount_(game))
        return { ok: false, msg: 'This round has already been answered.' };
      if (r[14] === 'answered' && String(r[15]) === String(sent.pick))
        return cu(miniFeedback_(game, round, r[16] === true, Number(r[12] || 0), r[17] === true));
      if (r[14] !== 'prompt') return { ok: false, msg: 'This round has already been answered.' };
      const late = !(at instanceof Date) || now.getTime() - at.getTime() > game.secs * 1000 + MINI.slackMs ||
        !(r[8] instanceof Date) || now.getTime() - r[8].getTime() > DAILY_GRACE_MIN * 60000;
      let correct = false;
      if (game.kind === 'scramble') {
        correct = !late && String(sent.pick || '').toUpperCase().replace(/[^A-Z]/g, '') === game.answer;
      } else {
        const item = game.rounds[round];
        correct = !late && String(sent.pick).trim() !== '' && Number.isInteger(Number(sent.pick)) && Number(sent.pick) === item.odd;
      }
      const right = Number(r[12] || 0) + (correct ? 1 : 0);
      sh.getRange(row, 12, 1, 7).setValues([[round, right, at, 'answered', String(sent.pick), correct, late]]);
      return cu(miniFeedback_(game, round, correct, right, late));
    } finally { lock.releaseLock(); }
  });
}
function webDailyNext(token, sent) {
  return webSafe_(() => {
    const s = webSession_(token);
    if (!s) return { ok: false, signedOut: true };
    const ss = webSheet_(), who = webRoles_(ss, s.e);
    if (!who.member) return { ok: false, msg: 'Daily brain breaks are for active START scholars.' };
    sent = sent || {};
    const key = String(sent.day || ''), sh = dailyTab_(ss), lock = docLock_();
    if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try again.' };
    try {
      const row = dailyRowFor_(sh, who.email, key);
      if (row < 0) return { ok: false, msg: 'Open today\'s brain break first.' };
      const r = sh.getRange(row, 1, 1, DAILY_HEAD.length).getValues()[0];
      if (!dailyPlayable_(key, r)) return { ok: false, msg: 'This brain break has closed.' };
      const cu = x => isCatchupRow_(r) && x && x.ok ? Object.assign(x, { score: catchupXp_(x.score), catchup: true }) : x;
      const game = webMiniFor_(who.email, key, miniProgram_(ss, who.email));
      const round = Number(r[11]);
      if (r[9] || sent.id !== r[10] || r[10] !== game.id || round >= miniCount_(game))
        return { ok: false, msg: 'The next round is not ready.' };
      if (r[14] === 'prompt' && Number(sent.round) === round - 1) {
        if (!(r[13] instanceof Date)) return { ok: false, msg: 'The next round is not ready.' };
        return { ok: true, round, prompt: miniPrompt_(game, round),
          remainingMs: Math.max(0, r[13].getTime() + game.secs * 1000 - Date.now()) };
      }
      if (r[14] !== 'answered' || Number(sent.round) !== round || round >= miniCount_(game) - 1)
        return { ok: false, msg: 'The next round is not ready.' };
      const next = round + 1, now = new Date();
      sh.getRange(row, 12, 1, 7).setValues([[next, Number(r[12] || 0), now, 'prompt', '', '', '']]);
      return { ok: true, round: next, prompt: miniPrompt_(game, next), remainingMs: game.secs * 1000 };
    } finally { lock.releaseLock(); }
  });
}
// MEMORY PAIRS: one attempt (two cards). The server knows the pairs, counts the moves and keeps the time limit.
// sent = { day, id, a, b } (card ids). Reply: { match, moves, found, done, timeUp, score }
function memTry_(game, st, a, b, now) {
  // st: { moves, found, playAt (ms), matched: [card ids] } -> updated in place
  if (!(a >= 0 && b >= 0 && a < game.cards.length && b < game.cards.length) || a === b) return { ok: false, msg: 'Pick two different cards.' };
  if (st.matched.indexOf(a) !== -1 || st.matched.indexOf(b) !== -1) return { ok: false, msg: 'That card is already matched.' };
  if (now > st.playAt + game.secs * 1000 + MEMORY.slackMs) { st.over = true; return { ok: true, timeUp: true, done: true, match: false, moves: st.moves, found: st.found, score: memScore_(st.found, st.moves)[1] }; }
  st.moves++;
  const ca = game.cards[a], cb = game.cards[b], match = ca.pair === cb.pair;
  if (match) { st.found++; st.matched.push(a, b); }
  const done = st.found >= game.pairs;
  return { ok: true, match, moves: st.moves, found: st.found, done, timeUp: false, score: done ? memScore_(st.found, st.moves)[1] : 0 };
}
function webDailyMatch(token, sent) {
  return webSafe_(() => {
    const s = webSession_(token);
    if (!s) return { ok: false, signedOut: true };
    const ss = webSheet_(), who = webRoles_(ss, s.e);
    if (!who.member) return { ok: false, msg: 'Daily brain breaks are for active START scholars.' };
    sent = sent || {};
    const key = String(sent.day || ''), sh = dailyTab_(ss), lock = docLock_();
    if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try again.' };
    try {
      const row = dailyRowFor_(sh, who.email, key);
      if (row < 0) return { ok: false, msg: 'Open today\'s brain break first.' };
      const r = sh.getRange(row, 1, 1, DAILY_HEAD.length).getValues()[0];
      if (!dailyPlayable_(key, r)) return { ok: false, msg: 'This brain break has closed.' };
      const game = webMiniFor_(who.email, key, miniProgram_(ss, who.email));
      if (r[9] || game.kind !== 'memory' || sent.id !== r[10] || r[10] !== game.id || r[14] !== 'memory') return { ok: false, msg: 'This game is over.' };
      const st = { moves: Number(r[11] || 0), found: Number(r[12] || 0), playAt: r[13] instanceof Date ? r[13].getTime() : 0,
                   matched: String(r[15] || '').split(',').filter(x => x !== '').map(Number) };
      if (!(r[8] instanceof Date) || Date.now() - r[8].getTime() > DAILY_GRACE_MIN * 60000) return { ok: true, timeUp: true, done: true, match: false, moves: st.moves, found: st.found };
      const out = memTry_(game, st, Math.floor(Number(sent.a)), Math.floor(Number(sent.b)), Date.now());
      if (!out.ok) return out;
      sh.getRange(row, 12, 1, 7).setValues([[st.moves, st.found, r[13], st.over ? 'over' : 'memory', st.matched.join(','), out.match, !!out.timeUp]]);
      if (isCatchupRow_(r) && out.score) { out.score = catchupXp_(out.score); out.catchup = true; }
      return out;
    } finally { lock.releaseLock(); }
  });
}

// RR test mode ("All three"): today's brain break for the test member, played for practice. Same game, same checks and
// timer, but the progress lives in a 30-minute cache only: no row in Daily games, no XP, no streak. op: start · pick · next · finish
function webDailyTest(token, op, sent) {
  return webSafe_(() => {
    const s = webSession_(token);
    if (!s) return { ok: false, signedOut: true };
    const ss = webSheet_();
    if (!webRoles_(ss, s.e).admin) return { ok: false, msg: 'Only the Regional Representative can test.' };
    sent = sent || {};
    const as = sent.as === '__test__' ? emailOf_(s.e) : emailOf_(sent.as), p = as ? scholarByEmail_(ss)[as] : null;
    if (sent.as !== '__test__' && (!p || p.status !== 'Active')) return { ok: false, msg: 'Choose an active member to test with.' };
    const now = Date.now(), today = dayKey_();
    const cache = CacheService.getUserCache ? CacheService.getUserCache() : CacheService.getScriptCache();
    const ck = 'DTEST_' + webHash_(s.e + '|' + as);
    const load = () => { try { return JSON.parse(cache.get(ck) || 'null'); } catch (e) { return null; } };
    const save = st => cache.put(ck, JSON.stringify(st), 1800);
    if (op === 'start') {
      const day = sent.day && /^\d{4}-\d\d-\d\d$/.test(String(sent.day)) && String(sent.day) < today ? String(sent.day) : today;   // a catch-up day, or today
      const g0 = webMiniFor_(as, day, miniProgram_(ss, as)), pub = miniPublic_(g0, 0, new Date(now));
      save({ id: g0.id, day, catchup: day !== today, round: 0, right: 0, at: now, started: now, state: g0.kind === 'memory' ? 'memory' : 'prompt',
             moves: 0, found: 0, playAt: now + (g0.preview || 0) * 1000, matched: [] });
      return { ok: true, day, test: true, game: day !== today ? catchupPublic_(pub) : pub };
    }
    const st = load();
    const key = st && st.day ? st.day : today, game = webMiniFor_(as, key, miniProgram_(ss, as));
    const cu = x => st && st.catchup && x && x.ok ? Object.assign(x, { score: catchupXp_(x.score), catchup: true }) : x;
    if (!st || st.id !== game.id || sent.id !== game.id) return { ok: false, msg: 'Practice game expired. Open it again.' };
    if (op === 'pick') {
      const round = st.round;
      if (Number(sent.round) !== round) return { ok: false, msg: 'This round has already been answered.' };
      if (st.state === 'answered' && String(st.pick) === String(sent.pick)) return cu(miniFeedback_(game, round, st.correct, st.right, st.late));
      if (st.state !== 'prompt') return { ok: false, msg: 'This round has already been answered.' };
      const late = now - st.at > game.secs * 1000 + MINI.slackMs || now - st.started > DAILY_GRACE_MIN * 60000;
      const correct = !late && (game.kind === 'scramble'
        ? String(sent.pick || '').toUpperCase().replace(/[^A-Z]/g, '') === game.answer
        : String(sent.pick).trim() !== '' && Number(sent.pick) === game.rounds[round].odd);
      Object.assign(st, { state: 'answered', pick: String(sent.pick), correct, late, right: st.right + (correct ? 1 : 0) }); save(st);
      return cu(miniFeedback_(game, round, correct, st.right, late));
    }
    if (op === 'match') {
      if (game.kind !== 'memory' || st.state !== 'memory') return { ok: false, msg: 'This game is over.' };
      const out = memTry_(game, st, Math.floor(Number(sent.a)), Math.floor(Number(sent.b)), now);
      if (out.ok) { if (st.over) st.state = 'over'; save(st); }
      return cu(out);
    }
    if (op === 'next') {
      if (st.state !== 'answered' || Number(sent.round) !== st.round || st.round >= miniCount_(game) - 1) return { ok: false, msg: 'The next round is not ready.' };
      Object.assign(st, { round: st.round + 1, at: now, state: 'prompt' }); save(st);
      return { ok: true, round: st.round, prompt: miniPrompt_(game, st.round), remainingMs: game.secs * 1000 };
    }
    if (op === 'finish') {
      const done = !sent.skip && st.state === 'answered' && st.round === miniCount_(game) - 1;
      const res = game.kind === 'memory' ? (sent.skip ? ['', 0, 0] : memScore_(st.found, st.moves)) : done ? miniScore_(game, st.right) : [miniScore_(game, 0)[0] + ' · left', 0, 0];
      cache.remove(ck);
      return { ok: true, test: true, xp: st.catchup ? catchupXp_(res[1]) : res[1], stars: res[2] };
    }
    return { ok: false, msg: 'Unknown step.' };
  });
}

// Finish only from the server's recorded round results. Client answers never award XP.
function webDailyFinish(token, sent) {
  return webSafe_(() => {
    const s = webSession_(token);
    if (!s) return { ok: false, signedOut: true };
    const ss = webSheet_(), who = webRoles_(ss, s.e);
    if (!who.member) return { ok: false, msg: 'Daily brain breaks are for active START scholars.' };
    sent = sent || {};
    const key = String(sent.day || ''), sh = dailyTab_(ss), lock = docLock_();
    if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try again.' };
    let res;
    try {
      const row = dailyRowFor_(sh, who.email, key);
      if (row < 0) return { ok: false, msg: 'This game was not started. Open today\'s brain break from your Journey.' };
      const r = sh.getRange(row, 1, 1, DAILY_HEAD.length).getValues()[0];
      if (r[9]) return { ok: true, already: true, member: JSON.parse(JSON.stringify(webMember_(ss, who))) };
      const game = webMiniFor_(who.email, key, miniProgram_(ss, who.email));
      const complete = dailyPlayable_(key, r) && !sent.skip && sent.id === r[10] && r[10] === game.id &&
        r[14] === 'answered' && Number(r[11]) === miniCount_(game) - 1 &&
        r[8] instanceof Date && r[13] instanceof Date && r[13].getTime() - r[8].getTime() <= DAILY_GRACE_MIN * 60000;
      res = complete ? miniScore_(game, Number(r[12] || 0)) : [miniScore_(game, 0)[0] + ' · left before the end', 0, 0];
      if (game.kind === 'memory') {                                    // pairs found and moves, as counted on the server
        const inTime = r[8] instanceof Date && Date.now() - r[8].getTime() <= DAILY_GRACE_MIN * 60000;
        res = !sent.skip && sent.id === r[10] && r[10] === game.id && inTime ? memScore_(Number(r[12] || 0), Number(r[11] || 0)) : ['Memory pairs · left before the end', 0, 0];
      }
      if (isCatchupRow_(r)) res = [res[0] + ' · ' + CATCHUP.label, catchupXp_(res[1]), res[2]];
      sh.getRange(row, 6, 1, 5).setValues([[res[0], res[1], res[2], r[8], new Date()]]);
    } finally { lock.releaseLock(); }
    // only this scholar's page changes (the school ranks catch up within 5 minutes): everyone else keeps their fast page
    try { webForget_(who.email); } catch (e) { /* cache */ }
    return { ok: true, xp: res[1], stars: res[2], member: JSON.parse(JSON.stringify(webMember_(ss, who))) };
  });
}

function webSubmitCheckin(token, a) {
  return webSafe_(() => {
    const s = webSession_(token);
    if (!s) return { ok: false, signedOut: true };
    const ss = webSheet_(), who = webRoles_(ss, s.e);
    if (!who.member) return { ok: false, msg: 'Only active START scholars on the list can check in.' };
    const spec = webCheckinSpec_(ss, who.email);
    if (!spec) return { ok: false, msg: 'The weekly check-in is not set up yet.' };
    if (!spec.open) return { ok: false, msg: `This week's check-in is closed. The next one opens on ${spec.opens}.` };
    a = a || {};
    const ans = {};
    for (const q of spec.q) {
      if (q.type === 'one') {
        ans[q.k] = q.choices.indexOf(String(a[q.k] || '')) !== -1 ? String(a[q.k]) : '';
        if (!ans[q.k]) return { ok: false, msg: 'Please answer: ' + q.title };
      } else if (q.type === 'many') {
        const v = (Array.isArray(a[q.k]) ? a[q.k] : []).map(String).filter(x => q.choices.indexOf(x) !== -1);
        if (!v.length) return { ok: false, msg: 'Please answer: ' + q.title };
        ans[q.k] = v.indexOf(q.none) !== -1 ? q.none : v.join(', ');
      } else ans[q.k] = String(a[q.k] || '').trim().slice(0, 1000);
    }
    if (a.consent !== true) return { ok: false, msg: 'Please tick the data privacy consent.' };
    let sh = memberTab_(ss);
    if (!sh) return { ok: false, msg: 'The weekly check-in is not set up yet.' };
    const lock = docLock_();
    if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try again.' };
    let row;
    try {
      if (memberCols_(sh).email < 0) { try { if (undoFormSwap_(ss)) sh = memberTab_(ss); } catch (e) { /* add a column instead */ } }
      const c = ensureEmailCol_(sh);                                            // a tab without one gets an email column
      const head = sh.getRange(1, 1, 1, c.width).getValues()[0].map(x => plainHeading_(x));
      if (c.email < 0) return { ok: false, msg: 'START could not save this right now. Please try again in a minute.' };
      let lateCol = head.indexOf(LATE_HEAD);
      if (spec.late && lateCol < 0) {                                            // "Answered late on" column, added the first time
        lateCol = c.width;
        if (sh.getMaxColumns() <= lateCol) sh.insertColumnsAfter(sh.getMaxColumns(), 1);
        sh.getRange(1, lateCol + 1).setValue(LATE_HEAD);
      }
      const meta = ensureCheckinMetadata_(sh);
      const width = Math.max(c.width, lateCol + 1, meta.width);
      const v = new Array(width).fill('');
      // a late answer counts for the week it belongs to (timestamp = that due day, 11:59 PM); the real time is kept beside it
      v[0] = spec.late ? new Date(spec.dueAt + (23 * 60 + 59) * 60000) : new Date();
      v[c.email] = who.email;
      if (c.name >= 0) v[c.name] = who.member.name || who.name || '';
      if (c.school >= 0) v[c.school] = who.member.school || '';
      if (c.role >= 0) v[c.role] = (who.leads || []).indexOf(who.member.school) !== -1 ? 'Campus Lead' : 'Member';
      ['vibe', 'events', 'learning', 'help', 'tea'].forEach(k => { if (c[k] >= 0 && ans[k] !== undefined) v[c[k]] = ans[k]; });
      const consentCol = head.indexOf(CONSENT_TITLE);
      if (consentCol >= 0) v[consentCol] = CONSENT_AGREE;
      if (spec.late) v[lateCol] = new Date();
      v[meta.week] = spec.week;
      v[meta.submitted] = new Date();
      // First submission is final. The check runs inside the document lock to prevent races.
      const again = webAnswerRow_(ss, who.email, spec.week, sh, c);
      if (again > 0) return { ok: true, saved: true, already: true }; // a lost reply must not ask for a second answer
      row = Math.max(sh.getLastRow(), 1) + 1;
      if (sh.getMaxRows() < row) sh.insertRowsAfter(sh.getMaxRows(), row - sh.getMaxRows());
      sh.getRange(row, 1, 1, v.length).setValues([v]);
      SpreadsheetApp.flush();
      MEMBER_CACHE_ = null;
    } finally { lock.releaseLock(); }          // the answer is saved: the next scholar can save theirs right away
    // follow-ups (help alerts, Campus Lead status, ranks) run AFTER the save, outside the lock, so a Friday rush does not queue up
    try { onMemberCheckin_({ range: sh.getRange(row, 1) }); } catch (e) { Logger.log('Website check-in follow-up: ' + e.message); }
    try { markAnswersSeen_(ss); } catch (e) { /* the hourly check catches up */ }
    MEMBER_CACHE_ = null;
    try { webForget_(who.email); } catch (e) { Logger.log('Check-in cache refresh: ' + e.message); }
    try { return { ok: true, member: JSON.parse(JSON.stringify(webMember_(ss, who))) }; }
    catch (e) { Logger.log('Check-in page refresh: ' + e.message); return { ok: true, saved: true }; }
  });
}

// A START Campus Lead: the same things as their portal sheet, for each school they lead
function webLead_(ss, schools, st) {
  const data = focalData_(ss);
  // One properties read for all member account badges on the Lead page.
  const accountProps = docProps_().getProperties();
  const link = (() => { try { return checkinUrl_(); } catch (e) { return ''; } })();
  const rrName = String(st.rrName || '').trim() || 'your Regional Representative';
  const tools = leadToolsData_(ss, schools);
  return schools.map(name => {
    const s = data.find(x => x.name === name);
    if (!s) return null;
    const P = portalMembers_(s, st);
    const todo = [], week = memberWeekKey_(new Date()), done = doneList_('LEAD_DONE|' + name, week);
    const lastDay = isDueDay_(P.due instanceof Date ? P.due : s.due);
    if (!s.thisWeek) todo.push({ tone: lastDay ? 'red' : 'yellow', text: `Answer your check-in by ${dueLabel_(s.due)}.` });
    P.help.forEach(m => todo.push({ tone: 'red', key: 'help|' + (m.email || m.name), text: `${displayName_(m.name)} asked for help: message them today.` }));
    if (P.answers && P.toMessage.length) {
      const n = P.toMessage.length, names = P.toMessage.slice(0, 8).map(m => displayName_(m.name)).join(', ') + (n > 8 ? ` and ${n - 8} more` : '');
      todo.push(lastDay
        ? { tone: 'red', key: 'last', copy: P.message, last: true,
            text: `Last day! ${n} member${n === 1 ? ' has' : 's have'} not answered yet (${names}). The check-in closes tonight at 11:59 PM. Send this last reminder:` }
        : { tone: 'yellow', key: 'remind', copy: P.message,
            text: `${n} member${n === 1 ? ' has' : 's have'} not answered yet (${names}). Remind them before ${dueLabel_(P.due instanceof Date ? P.due : s.due)}. You can copy this:` });
    }
    if (P.noEmail.length) todo.push({ tone: 'grey', key: 'noemail',
      text: `${P.noEmail.length} member${P.noEmail.length === 1 ? ' has' : 's have'} no email on file: ask them to email ${rrName} (${me_()}) the Google email they use.` });
    const doneCount = todo.filter(t => t.key && done.indexOf(t.key) !== -1).length;
    for (let i = todo.length - 1; i >= 0; i--) if (todo[i].key && done.indexOf(todo[i].key) !== -1) todo.splice(i, 1);
    if (!todo.length) todo.push({ tone: 'green', text: doneCount ? 'All done for this week. Salamat!' : 'Nothing to do: you and your members are all in. Salamat!' });
    return {
      school: s.name, done: !!s.thisWeek, due: dueLabel_(s.due), lastDay, doneCount, streak: s.streak, term: s.term, active: s.active,
      answers: P.answers, answered: P.answered, members: P.members.length, link: link || s.link || '', todo,
      cols: P.cols.map(c => c[0]), rows: P.rows.map(x => ({ cells: x.cells.map(webCell_), tone: webTone_(x.chip), me: !!x.me, email: x.email || '',
        account: x.email ? (/"h"\s*:/.test(accountProps[webAcctKey_(x.email)] || '') ? 'ready' : 'setup') : '', contacted: x.email ? (tools.contacted[name + '|' + x.email] || '') : '' })),
      joins: tools.joins.filter(j => j.school === name), attend: (() => { try { return leadRecentEvents_(ss, name, tools.att); } catch (e) { Logger.log('Lead events: ' + e.message); return []; } })(),
      note: leadNote_(name),
      history: s.checkins.slice(0, 4).map(x => ({ date: webFmt_(x.date, 'EEE, MMM d'), events: String(x.events || '—'), help: String(x.help || '—') })),
      events: webEvents_(ss, s.name, true),
      weekly: P.answers ? (() => { try { return weeklyRates_(ss, 10, s.name); } catch (e) { return []; } })() : [],
    };
  }).filter(Boolean);
}

// The RR: the whole region at a glance
// RR-only roster, read from the Scholars sheet; never return password or session data.
function webScholarRoster_(ss, md) {
  const sh = ss.getSheetByName(TAB.scholars);
  if (!sh || sh.getLastRow() < 2) return [];
  const width = Math.min(sh.getMaxColumns(), 18);
  const leads = schoolRows_(ss);
  const leadKeys = new Set(leads.filter(s => s.focal && s.email).map(s => s.focal + '\n' + s.email));
  const members = new Map();
  if (md) Object.keys(md.bySchool).forEach(school => {
    const byEmail = new Map();
    md.bySchool[school].members.forEach(m => { if (m.email && !byEmail.has(m.email)) byEmail.set(m.email, m); });
    members.set(school, byEmail);
  });
  return sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues()
    .map((r, i) => ({ row: i + 2, name: String(r[0] || '').trim(), school: String(r[1] || '').trim(),
      program: String(r[2] || '').trim(), awarded: String(r[3] || '').trim(), contact: String(r[4] || '').trim(),
      email: emailOf_(r[4]), status: String(r[5] || '').trim(), yearLevel: String(r[6] || '').trim(),
      campus: String(r[7] || '').trim(), hometown: String(r[8] || '').trim(), role: String(r[9] || '').trim(),
      otherEmails: String(r[10] || '').trim(), scholarship: String(r[11] || '').trim(), verified: !!r[17] }))
    .map(p => {
      const byEmail = members.get(p.school), member = byEmail && byEmail.get(p.email);
      p.role = leadKeys.has(p.name + '\n' + p.email) ? ROLES[1] : ROLES[0];
      p.checkin = member ? (member.now ? 'Submitted' : 'Pending') : '';
      p.needsHelp = !!(member && member.now && memberRank_(member) === 0);
      return p;
    })
    .filter(p => p.name && p.email);
}
// Suspending preserves all scholar records and history, but immediately revokes existing sessions.
function webScholarStatus_(ss, actor, arg) {
  arg = arg || {};
  const row = Number(arg.row), to = String(arg.status || '');
  if (!Number.isInteger(row) || row < 2 || ['Active', 'Suspended'].indexOf(to) === -1)
    return { ok: false, msg: 'Choose a valid scholar and action.' };
  const sh = ss.getSheetByName(TAB.scholars);
  if (!sh || row > sh.getLastRow()) return { ok: false, msg: 'Scholar not found. Reload the list.' };
  const lock = docLock_();
  if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try again.' };
  let email, old, aliases = [];
  try {
    const rec = sh.getRange(row, 1, 1, Math.min(sh.getMaxColumns(), 11)).getValues()[0];
    email = emailOf_(rec[4]); old = String(rec[5] || '').trim();
    aliases = altEmails_(rec[10]);
    if (!email || email !== emailOf_(arg.email)) return { ok: false, msg: 'The scholar list changed. Reload and try again.' };
    if (email === emailOf_(actor)) return { ok: false, msg: 'You cannot suspend your own RR account.' };
    if (['Active', 'Suspended'].indexOf(old) === -1) return { ok: false, msg: 'This account has status ' + (old || 'blank') + '. Change it in Scholars if needed.' };
    if (old === to) return { ok: true, msg: 'This scholar is already ' + to.toLowerCase() + '.' };
    sh.getRange(row, 6).setValue(to);
    markScholarChange_(ss, sh, row, actor, 'status', old + ' -> ' + to, old, to);
    const key = 'WEB_SUSP_' + webHash_(email);
    if (to === 'Suspended') docProps_().setProperty(key, '1'); else docProps_().deleteProperty(key);
  } finally { lock.releaseLock(); }
  [email].concat(aliases).forEach(e => { webEndSessions_(e); webForget_(e); });
  webForget_(actor);
  MEMBER_CACHE_ = null;
  try { CacheService.getScriptCache().remove('WEB_LEADERBOARD'); } catch (e) { /* no cache */ }
  try { calPeopleChanged_(ss); } catch (e) { Logger.log('Calendar sharing: ' + e.message); }
  Logger.log('RR scholar status: ' + email + ' ' + old + ' -> ' + to + ' by ' + actor);
  return { ok: true, msg: to === 'Suspended' ? 'Scholar suspended. Their sessions have been ended; records are retained.' : 'Scholar reactivated. They can sign in again.' };
}

// Save one existing scholar after checking the row and every sign-in email under a document lock.
function webScholarUpdate_(ss, actor, arg) {
  arg = arg || {};
  const row = Number(arg.row), sh = ss.getSheetByName(TAB.scholars);
  if (!sh || !Number.isInteger(row) || row < 2 || row > sh.getLastRow())
    return { ok: false, msg: 'Scholar not found. Reload the list.' };
  const clean = (v, n) => String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, n);
  const lock = docLock_();
  if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try again.' };
  let oldEmail = '', newEmail = '', displacedEmail = '', oldAliases = [], newAliases = [], accessChanged = false;
  try {
    const old = sh.getRange(row, 1, 1, Math.min(sh.getMaxColumns(), 12)).getValues()[0];
    oldEmail = emailOf_(old[4]);
    if (!oldEmail || oldEmail !== emailOf_(arg.email)) return { ok: false, msg: 'The scholar list changed. Reload and try again.' };
    const name = clean(arg.name, 80), school = clean(arg.school, 80), program = clean(arg.program, 180);
    const awarded = clean(arg.awarded, 4), contact = clean(arg.contact, 180), status = clean(arg.status, 30);
    const yearLevel = clean(arg.yearLevel, 40), campus = clean(arg.campus, 100), hometown = clean(arg.hometown, 120);
    const scholarship = clean(arg.scholarship, 80), otherEmails = clean(arg.otherEmails, 300);
    const role = clean(arg.role, 40);
    newEmail = emailOf_(contact);
    if (name.length < 2 || !newEmail) return { ok: false, msg: 'Enter a name and a contact email.' };
    const schools = schoolRows_(ss), target = schools.find(s => s.name === school);
    if (!target) return { ok: false, msg: 'Choose a school on the Schools list.' };
    if (target.inactive && role === ROLES[1]) return { ok: false, msg: 'Campus Leads must belong to an active school.' };
    if (program && START_PROGRAMS.list.indexOf(program) === -1 && program !== String(old[2] || '').trim())
      return { ok: false, msg: 'Choose a START degree program.' };
    if (awarded && (!/^\d{4}$/.test(awarded) || Number(awarded) < 2015 || Number(awarded) > new Date().getFullYear() + 1))
      return { ok: false, msg: 'Enter a valid scholarship award year.' };
    if (['Active', 'On leave', 'Graduated', 'Dropped', 'Suspended'].indexOf(status) === -1)
      return { ok: false, msg: 'Choose a valid status.' };
    if (yearLevel && ['1st Year', '2nd Year', '3rd Year', '4th Year', '5th Year'].indexOf(yearLevel) === -1 && yearLevel !== String(old[6] || '').trim())
      return { ok: false, msg: 'Choose a year level.' };
    if (scholarship && JOIN.kinds.indexOf(scholarship) === -1 && scholarship !== String(old[11] || '').trim())
      return { ok: false, msg: 'Choose a DOST scholarship.' };
    if (ROLES.indexOf(role) === -1) return { ok: false, msg: 'Choose Member or Campus Lead.' };
    if (role === ROLES[1] && status !== 'Active' && status !== 'Suspended')
      return { ok: false, msg: 'Set the scholar to Active before assigning them as Campus Lead.' };
    if (oldEmail === emailOf_(actor) && (newEmail !== oldEmail || status !== 'Active'))
      return { ok: false, msg: 'You cannot change your own RR email or access status here.' };
    const aliases = otherEmails ? otherEmails.split(/[,;\s]+/).filter(Boolean).map(x => x.toLowerCase()) : [];
    if (aliases.some(x => emailOf_(x) !== x) || new Set([newEmail].concat(aliases)).size !== aliases.length + 1)
      return { ok: false, msg: 'Other emails must be distinct, valid email addresses.' };
    const all = sh.getRange(2, 1, sh.getLastRow() - 1, Math.min(sh.getMaxColumns(), 11)).getValues();
    const wanted = [newEmail].concat(aliases);
    if (all.some((r, i) => i + 2 !== row && [emailOf_(r[4])].concat(altEmails_(r[10])).some(e => e && wanted.indexOf(e) !== -1)))
      return { ok: false, msg: 'One of these emails already belongs to another scholar.' };
    const oldName = String(old[0] || '').trim();
    oldAliases = altEmails_(old[10]); newAliases = aliases;
    const led = schools.filter(s => s.focal === oldName && s.email === oldEmail);
    accessChanged = newEmail !== oldEmail || status !== String(old[5] || '').trim() ||
      role !== (led.length ? ROLES[1] : ROLES[0]) ||
      (role === ROLES[1] && (name !== oldName || school !== String(old[1] || '').trim())) ||
      oldAliases.slice().sort().join(',') !== newAliases.slice().sort().join(',');
    if (role === ROLES[1] && target.focal && target.email !== oldEmail) displacedEmail = target.email;
    sh.getRange(row, 1, 1, 9).setValues([[name, school, program, awarded ? Number(awarded) : '', contact,
      status, yearLevel, campus, hometown]]);
    if (sh.getMaxColumns() < 12) sh.insertColumnsAfter(sh.getMaxColumns(), 12 - sh.getMaxColumns());
    if (!sh.getRange(1, ALT_EMAIL_COL).getValue()) sh.getRange(1, ALT_EMAIL_COL).setValue(ALT_EMAIL_HEAD);
    if (!sh.getRange(1, SCHOLAR_KIND_COL).getValue()) sh.getRange(1, SCHOLAR_KIND_COL).setValue('Scholarship (DOST)');
    sh.getRange(row, ALT_EMAIL_COL).setValue(aliases.join(', '));
    sh.getRange(row, SCHOLAR_KIND_COL).setValue(scholarship);
    if (newEmail !== oldEmail && sh.getMaxColumns() >= 18) sh.getRange(row, 18).clearContent();
    const changed = ['Name', 'School', 'Program', 'Award year', 'Contact', 'Status', 'Year level', 'Campus', 'Hometown']
      .filter((label, i) => String(old[i] || '').trim() !== String([name, school, program, awarded, contact, status, yearLevel, campus, hometown][i] || '').trim());
    if (String(old[10] || '').trim() !== aliases.join(', ')) changed.push('Other emails');
    if (String(old[11] || '').trim() !== scholarship) changed.push('Scholarship');
    if (String(old[9] || '').trim() !== role) changed.push('Role');
    markScholarChange_(ss, sh, row, actor, 'edit', changed.join(', ') || 'No fields changed', String(old[5] || '').trim(), status);
    const schoolsSheet = ss.getSheetByName(TAB.schools);
    led.forEach(s => { if (role !== ROLES[1] || s.name !== school) schoolsSheet.getRange(s.row, 4).clearContent(); });
    if (role === ROLES[1]) schoolsSheet.getRange(target.row, 4).setValue(name);
    SpreadsheetApp.flush();
    syncRoles_(ss);
    const props = docProps_();
    if (newEmail !== oldEmail) props.deleteProperty('WEB_SUSP_' + webHash_(oldEmail));
    const key = 'WEB_SUSP_' + webHash_(newEmail);
    if (status === 'Suspended') props.setProperty(key, '1'); else props.deleteProperty(key);
    MEMBER_CACHE_ = null;
  } finally { lock.releaseLock(); }
  const emails = [...new Set([oldEmail, newEmail, displacedEmail].concat(oldAliases, newAliases).filter(Boolean))];
  emails.forEach(e => { if (accessChanged || e === displacedEmail) webEndSessions_(e); webForget_(e); });
  webForget_(actor);
  try { CacheService.getScriptCache().remove('WEB_LEADERBOARD'); } catch (e) { /* no cache */ }
  try { calPeopleChanged_(ss); } catch (e) { Logger.log('Calendar sharing: ' + e.message); }
  Logger.log('RR updated scholar row ' + row + ' by ' + actor);
  return { ok: true, msg: 'Scholar details saved.' + (newEmail !== oldEmail ? ' The new email must set up its START password.' : '') };
}

// Editable database fields for the RR website. Formula columns are deliberately excluded.
function webSchoolRecords_(ss) {
  const sh = ss.getSheetByName(TAB.schools);
  if (!sh) return [];
  return sh.getRange(2, 1, CONFIG.maxSchools, 15).getValues().map((v, i) => ({
    row: i + 2, name: String(v[0] || '').trim(), province: String(v[1] || '').trim(),
    lead: String(v[3] || '').trim(), inactive: String(v[5] || '').trim().toUpperCase() === 'Y',
    info: v.slice(8, 15).map(x => String(x || '').trim()),
  })).filter(s => s.name);
}

function webSettingRecords_(ss) {
  const sh = ss.getSheetByName(TAB.settings);
  if (!sh) return [];
  const vals = sh.getRange(2, 2, SETTINGS.length, 1).getValues();
  const derived = ['remDay', 'fuDay', 'mDueDay'];
  return SETTINGS.map((s, i) => {
    const key = s[0], raw = vals[i] ? vals[i][0] : '', date = raw instanceof Date;
    return { key, label: s[1], help: s[3],
      value: date ? webFmt_(raw, 'yyyy-MM-dd') : raw == null ? '' : raw,
      type: CHECKBOX_SETTINGS.indexOf(key) !== -1 ? 'check' : date || ['dFocal', 'dReeng', 'dcStart', 'dcEnd', 'termStart'].indexOf(key) !== -1 ? 'date' :
            ['quietDays', 'tFocal', 'tReeng', 'tRate'].indexOf(key) !== -1 ? 'number' : 'text',
      options: key === 'region' ? regionChoices_() : ['day', 'dueDay', 'mOpenDay'].indexOf(key) !== -1 ? DAYS :
               ['hour', 'remHour'].indexOf(key) !== -1 ? HOURS : [],
      readonly: DB_MODE && derived.indexOf(key) !== -1 };
  });
}

function webOutreachRecords_(ss) {
  const sh = ss.getSheetByName(TAB.outreach);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, Math.min(sh.getLastRow() - 1, 100), OUTREACH_HEAD.length).getValues()
    .map((r, i) => ({ row: i + 2, date: webFmt_(r[0], 'yyyy-MM-dd'), school: String(r[1] || ''),
      person: String(r[2] || ''), channel: String(r[3] || ''), outcome: String(r[4] || ''),
      next: String(r[5] || ''), follow: webFmt_(r[6], 'yyyy-MM-dd') }))
    .filter(r => r.school);
}

function webEventRecords_(ss) {
  const sh = ss.getSheetByName(TAB.events);
  if (!sh || sh.getLastRow() < 2) return [];
  const webTime = v => { const n = minutesOf_(v); return n === null ? '' :
    String(Math.floor(n / 60)).padStart(2, '0') + ':' + String(n % 60).padStart(2, '0'); };
  return sh.getRange(2, 1, sh.getLastRow() - 1, EV_EDIT).getValues().map((r, i) => ({
    row: i + 2, date: webFmt_(r[EV.date], 'yyyy-MM-dd'), start: webTime(r[EV.start]),
    end: webTime(r[EV.end]), title: String(r[EV.title] || ''), type: String(r[EV.type] || ''),
    who: String(r[EV.who] || ''), rep: String(r[EV.rep] || ''), until: webFmt_(r[EV.until], 'yyyy-MM-dd'),
    place: String(r[EV.place] || ''), details: String(r[EV.details] || ''), invite: r[EV.invite] === true,
  })).filter(r => r.title).sort((a, b) => b.date.localeCompare(a.date)).slice(0, 100);
}

function webValidDate_(value) {
  if (!value) return '';
  const s = String(value), m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const d = new Date(+m[1], +m[2] - 1, +m[3]);
  return d.getFullYear() === +m[1] && d.getMonth() === +m[2] - 1 && d.getDate() === +m[3] ? d : null;
}

function webOutreachSave_(ss, actor, arg) {
  arg = arg || {};
  const school = String(arg.school || '').trim(), date = webValidDate_(arg.date), follow = webValidDate_(arg.follow);
  if (!schoolRows_(ss).some(s => s.name === school) || !date || follow === null ||
      OUTREACH_CHANNELS.indexOf(arg.channel) < 0 || OUTREACH_OUTCOMES.indexOf(arg.outcome) < 0 ||
      [arg.person, arg.next].some(x => String(x || '').length > 500 || /^\s*=/.test(String(x || ''))))
    return { ok: false, msg: 'Check the school, date, channel, outcome and notes.' };
  const lock = docLock_();
  if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try again.' };
  try { saveOutreach(arg, ss); } finally { lock.releaseLock(); }
  webBumpData_();
  try { renderCalView_(ss); } catch (e) { Logger.log('Outreach calendar view: ' + e.message); }
  Logger.log('RR saved outreach by ' + actor);
  return { ok: true, msg: 'Contact recorded.' };
}

function webEventSave_(ss, actor, arg) {
  arg = arg || {};
  const title = String(arg.title || '').trim(), date = webValidDate_(arg.date), until = webValidDate_(arg.until);
  const schoolNames = schoolRows_(ss).map(s => s.name);
  const text = [title, arg.place, arg.details];
  if (!title || title.length > 150 || !date || until === null ||
      EV_TYPES.indexOf(arg.type) < 0 || [WHO.all, WHO.leads, WHO.me].concat(schoolNames).indexOf(arg.who) < 0 ||
      EV_REPEATS.indexOf(arg.rep) < 0 || arg.rep !== 'Never' && (!until || until < date) ||
      text.some(x => String(x || '').length > 500 || /^\s*=/.test(String(x || ''))))
    return { ok: false, msg: 'Check the event title, dates, type, audience and details.' };
  const start = String(arg.start || ''), end = String(arg.end || '');
  if (start && minutesOf_(start) === null || end && (!start || minutesOf_(end) === null || minutesOf_(end) <= minutesOf_(start)))
    return { ok: false, msg: 'Choose a valid start and later end time, or leave both blank.' };
  const sh = ss.getSheetByName(TAB.events);
  const row = Number(arg.row || 0);
  if (row && (!sh || !Number.isInteger(row) || row < 2 || row > sh.getLastRow()))
    return { ok: false, msg: 'Event not found. Reload and try again.' };
  const lock = docLock_();
  if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try again.' };
  let message = 'Event saved.';
  try {
    if (!row) message = addEventFromBox(arg, ss);
    else {
      const old = sh.getRange(row, EV.title + 1).getValue();
      if (String(old || '').trim() !== String(arg.original || '').trim())
        return { ok: false, msg: 'This event changed. Reload and try again.' };
      const fmt = t => t ? Utilities.formatDate(new Date(2000, 0, 1, Math.floor(minutesOf_(t) / 60), minutesOf_(t) % 60), CONFIG.timezone, 'h:mm a') : '';
      sh.getRange(row, 1, 1, EV_EDIT).setValues([[
        date, fmt(start), fmt(end), title, arg.type, arg.who, arg.rep, arg.rep === 'Never' ? '' : until,
        String(arg.place || '').trim(), String(arg.details || '').trim(), arg.invite === true,
      ]]);
      try { renderCalView_(ss); } catch (e) { Logger.log('Event calendar view: ' + e.message); }
    }
  } finally { lock.releaseLock(); }
  if (row) try { eventsChanged_(ss); } catch (e) { Logger.log('Event sync: ' + e.message); message = 'Event saved. Calendar sync will retry.'; }
  webBumpData_();
  Logger.log('RR saved event by ' + actor);
  return { ok: true, msg: message };
}

function webEventDelete_(ss, actor, arg) {
  const sh = ss.getSheetByName(TAB.events), row = Number(arg && arg.row);
  if (!sh || !Number.isInteger(row) || row < 2 || row > sh.getLastRow())
    return { ok: false, msg: 'Event not found. Reload and try again.' };
  const lock = docLock_();
  if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try again.' };
  try {
    if (String(sh.getRange(row, EV.title + 1).getValue() || '').trim() !== String(arg.original || '').trim())
      return { ok: false, msg: 'This event changed. Reload and try again.' };
    sh.getRange(row, 1, 1, EV_EDIT).clearContent();
    try { renderCalView_(ss); } catch (e) { Logger.log('Event calendar view: ' + e.message); }
  } finally { lock.releaseLock(); }
  try { eventsChanged_(ss); } catch (e) { Logger.log('Event removal sync: ' + e.message); }
  webBumpData_();
  Logger.log('RR removed event by ' + actor);
  return { ok: true, msg: 'Event removed.' };
}

// RR editing lists are loaded when their page opens, so the overview stays small and quick.
function webAdminPage(token, page) {
  return webTimed_('admin ' + String(page || '').replace(/[^a-z]/gi, ''), () => {
    const session = webSession_(token);
    if (!session) return { ok: false, signedOut: true };
    const ss = webSheet_(), who = webRoles_(ss, session.e);
    if (!who.admin || who.suspended) return { ok: false, msg: 'Only an active Regional Representative can view this page.' };
    if (page === 'schools') return { ok: true, schoolRecords: webSchoolRecords_(ss), schoolInfoHeaders: SCHOOL_INFO_HEADERS };
    if (page === 'settings') return { ok: true, settings: webSettingRecords_(ss) };
    if (page === 'outreach') return { ok: true, outreach: webOutreachRecords_(ss),
      outreachChannels: OUTREACH_CHANNELS, outreachOutcomes: OUTREACH_OUTCOMES };
    if (page === 'events') return { ok: true, eventRecords: webEventRecords_(ss),
      eventTypes: EV_TYPES, eventRepeats: EV_REPEATS, calendarEvents: webCalendarEvents_(ss, who) };
    if (page === 'diagnostics') return { ok: true, diagnostics: healthRows_().slice(0, 100) };
    if (page === 'rewards') return { ok: true, rewardPending: rewardPending_(ss), rewardCatalog: rewardCatalog_(ss) };
    if (page === 'scholars') {
      let md = null;
      try { if (checkinOn_()) md = memberData_(ss); } catch (e) { Logger.log('Scholar check-ins: ' + e.message); }
      let joins = [];
      try { joins = webJoins_(ss); } catch (e) { Logger.log('Join requests: ' + e.message); }
      let dataQuality = { count: 0, issues: [] };
      try { dataQuality = dataFlowQuality_(ss); } catch (e) { Logger.log('Data quality: ' + e.message); }
      const approved = joins.filter(j => /^Added/.test(j.status)).sort((a, b) => b.row - a.row);
      const declined = joins.filter(j => /^Declined/.test(j.status)).sort((a, b) => b.row - a.row);
      const checks = (() => { try { return leadJoinChecks_(ss); } catch (e) { return {}; } })();
      joins.forEach(j => { if (checks[j.row]) j.leadCheck = checks[j.row]; });
      return { ok: true, scholars: webScholarRoster_(ss, md), dataQuality,
        joins: joins.filter(j => !/^Added|^Declined/.test(j.status)),
        joinHistory: approved.slice(0, 20).concat(declined.slice(0, 20)),
        joinHistoryTotal: { Approved: approved.length, Declined: declined.length } };
    }
    return { ok: false, msg: 'Unknown website page.' };
  });
}

function webJoinHistory(token, kind, offset) {
  return webTimed_('join history', () => {
    const session = webSession_(token);
    if (!session) return { ok: false, signedOut: true };
    const ss = webSheet_(), who = webRoles_(ss, session.e);
    if (!who.admin || who.suspended) return { ok: false, msg: 'Only an active Regional Representative can view this page.' };
    if (['Approved', 'Declined'].indexOf(kind) < 0 || !Number.isInteger(offset) || offset < 0)
      return { ok: false, msg: 'Invalid history page.' };
    const prefix = kind === 'Approved' ? /^Added/ : /^Declined/;
    const rows = webJoins_(ss).filter(j => prefix.test(j.status)).sort((a, b) => b.row - a.row);
    return { ok: true, rows: rows.slice(offset, offset + 20), total: rows.length };
  });
}

function webAdmin_(ss, st) {
  const data = focalData_(ss).filter(s => !s.inactive);
  let md = null; try { if (checkinOn_()) md = memberData_(ss); } catch (e) { md = null; }
  const schools = data.map(s => {
    const m = (md && md.bySchool[s.name]) || { members: [], answered: 0 };
    return { name: s.name, lead: s.focal ? displayName_(s.focal) : '', leadDone: !!s.thisWeek, streak: s.streak,
             members: md ? m.members.length : s.active, answered: m.answered,
             help: m.members.filter(x => memberRank_(x) === 0).length, last: webFmt_(s.last, 'MMM d') };
  });
  const sum = k => schools.reduce((a, s) => a + (Number(s[k]) || 0), 0);
  const help = md ? md.helpRecent.slice().sort((a, b) => b.date - a.date).slice(0, 15)
    .map(h => ({ name: displayName_(h.name), school: h.school || '', when: webFmt_(h.date, 'EEE, MMM d'),
                 what: [h.help, h.vibe, h.tea].filter(Boolean).join(' · ') })) : [];
  let form = ''; try { form = WEB_CHECKIN ? '' : checkinUrl_(); } catch (e) { form = ''; }
  return {
    due: md ? dueLabel_(md.week.due) : dueLabel_(reminderWeek_(new Date()).deadline),
    attnDone: doneList_('RR_DONE', memberWeekKey_(new Date())),
    totals: { schools: schools.length, leads: schools.filter(s => s.lead).length, leadsDone: schools.filter(s => s.lead && s.leadDone).length,
              members: sum('members'), answered: sum('answered'), help: sum('help') },
    memberCheckins: !!md, schools, help, events: webEvents_(ss, undefined, true), tracker: ss.getUrl(), form,
    scholarChoices: { schools: joinSchoolList_(ss).map(s => s.name), programs: START_PROGRAMS.list, scholarships: JOIN.kinds },
    autoApprove: autoApproveOn_(),
    leadBoard: (() => { try { return leadBoard_(ss, schools); } catch (e) { Logger.log('Lead board: ' + e.message); return []; } })(),
    rewardPending: (() => { try { return rewardPending_(ss); } catch (e) { Logger.log('Rewards: ' + e.message); return []; } })(),
    weekly: md ? (() => { try { return weeklyRates_(ss, 10); } catch (e) { Logger.log('Weekly rate: ' + e.message); return []; } })() : [],
    // what used to come by email: quiet schools, the tracker check, the monthly report
    quiet: (() => { try { return missedSchools_(ss).slice(0, 12).map(q => ({ name: q.name, lead: q.focal ? displayName_(q.focal) : '', weeks: q.missed, msg: missedMessage_(q, st) })); } catch (e) { return []; } })(),
    check: (() => { try { const c = JSON.parse(docProps_().getProperty('LAST_CHECK') || 'null'); return c ? { when: webFmt_(new Date(c.at), 'EEE, MMM d'), rows: c.rows } : null; } catch (e) { return null; } })(),
    report: (() => { try { const r = JSON.parse(docProps_().getProperty('LAST_REPORT') || 'null'); return r && Date.now() - r.at < 45 * 86400000 ? { label: r.label, url: r.url } : null; } catch (e) { return null; } })(),
  };
}

/* ---------- menu: 🌐 Website ---------- */

function showWebsite() {
  sheetUser_();                                   // sheet only: never from the website
  const fixed = docProps_().getProperty('WEB_URL_SET') || '';
  let url = fixed;
  if (!url) { try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { url = ''; } }
  if (!url) url = docProps_().getProperty('WEB_URL') || '';
  if (!url) {
    notify_('🌐 Switch on the website (once)\n\n' +
      '1. Extensions › Apps Script › Deploy › New deployment\n' +
      '2. ⚙️ › Web app\n' +
      '3. Execute as: Me · Who has access: Anyone › Deploy\n' +
      '4. Come back and click this again');
    return;
  }
  if (!fixed) docProps_().setProperty('WEB_URL', url);
  const st = getSettings_();
  const app = webAppOrigins_()[0] || '';
  const appUrl = app ? (docProps_().getProperty('WEB_APP_URL') || '') : '';   // saved before the full address was kept: re-save once
  const share = appUrl || url;
  const msg = `Hi START scholars! Our START ${st.region || ''} portal:\n${share}\n\n` +
    'Sign in with the email you gave START. ' + (appUrl ? 'Tip: add it to your home screen.' : '');
  const linkRow = (icon, label, u) => `<div class="box row" style="margin-bottom:8px"><b style="min-width:72px">${icon} ${label}</b>` +
    `<a href="${esc_(u)}" target="_blank" style="flex:1;word-break:break-all;font-size:12px">${esc_(u)}</a>` +
    `<button class="link" onclick="copy('${esc_(u)}', this)">Copy</button></div>`;
  const html = `
    ${fixed ? '' : '<div class="box" style="background:#FFF6BF;margin-bottom:10px"><b>Paste your website address once (below).</b> ' +
      'Google sometimes reports an old address that does not open. Copy the <b>Web app URL</b> from Deploy › Manage deployments.</div>'}
    ${fixed ? linkRow('🌐', 'Website', url) : ''}
    ${appUrl && fixed ? linkRow('📱', 'App', appUrl) : ''}
    ${fixed ? `<label>Message for group chats</label>
    <textarea id="m" class="copy" style="height:110px">${esc_(msg)}</textarea>
    <div style="margin-top:8px"><button class="btn" onclick="copy(document.getElementById('m').value, this)">Copy message</button></div>` : ''}
    <label style="margin-top:18px">Website address (Deploy › Manage deployments › Web app URL)</label>
    <div class="row"><input id="web" style="flex:1" value="${esc_(fixed)}" placeholder="https://script.google.com/macros/s/…/exec">
      <button class="btn dark" id="svw" onclick="saveWeb()">Save</button></div>
    <div class="msg ok" id="webmsg"></div>
    <label style="margin-top:18px">App address</label>
    <div class="row"><input id="app" style="flex:1" value="${esc_(appUrl || app)}" placeholder="https://name.github.io/start-app/">
      <button class="btn dark" id="sv" onclick="save()">Save</button></div>
    <div class="msg ok" id="appmsg"></div>
    <label style="margin-top:18px">Accounts · ${webAccountCount_()} with a password</label>
    <div class="row"><input id="rp" style="flex:1" placeholder="Forgot it? Their email here"><button class="btn dark" id="rpb" onclick="resetPw()">Reset</button></div>
    <div class="msg ok" id="rpmsg"></div>
    <label style="margin-top:18px">Someone gets no sign-in code? Check their email</label>
    <div class="row"><input id="ck" style="flex:1" placeholder="Their email here"><button class="btn dark" id="ckb" onclick="checkMail()">Check</button></div>
    <div class="msg ok" id="ckmsg" style="white-space:pre-line"></div>
    <p class="muted" style="margin-top:12px">Changed the code? Deploy › Manage deployments › ✏️ › New version.</p>
    <details style="margin-top:8px"><summary class="muted" style="cursor:pointer">Handing over to a new RR</summary>
      <ol class="muted" style="margin:6px 0 0 18px;padding:0">
        <li>Share this tracker with them as Editor (File › Share).</li>
        <li>They open Extensions › Apps Script › Deploy › New deployment › Web app (Execute as: Me · Anyone).</li>
        <li>They paste the new address in the START app's config.js on GitHub, then save the app address here.</li>
        <li>In the sheet: ⏰ Automatic emails › 🔑 Make my account send them.</li>
        <li>Settings: their name and email. Everyone keeps their password.</li>
      </ol></details>
    <script>
      function copy(t, b) { var x = document.createElement('textarea'); x.value = t; document.body.appendChild(x); x.select();
        try { navigator.clipboard.writeText(t); } catch (e) {} try { document.execCommand('copy'); } catch (e) {}
        document.body.removeChild(x); var o = b.textContent; b.textContent = 'Copied ✓'; setTimeout(function () { b.textContent = o; }, 1800); }
      function checkMail() { var b = document.getElementById('ckb'); b.disabled = true; document.getElementById('ckmsg').textContent = 'Checking…';
        google.script.run.withSuccessHandler(function (r) { document.getElementById('ckmsg').textContent = r; b.disabled = false; })
          .withFailureHandler(function (e) { document.getElementById('ckmsg').textContent = 'Problem: ' + e.message; b.disabled = false; })
          .signInCheck(document.getElementById('ck').value); }
      function resetPw() { var b = document.getElementById('rpb'); b.disabled = true;
        google.script.run.withSuccessHandler(function (r) { document.getElementById('rpmsg').textContent = r; b.disabled = false; })
          .resetWebPassword(document.getElementById('rp').value); }
      function saveWeb() { var b = document.getElementById('svw'); b.disabled = true; document.getElementById('webmsg').textContent = 'Checking the address…';
        google.script.run.withSuccessHandler(function (r) { document.getElementById('webmsg').textContent = r; b.disabled = false;
            if (/^✅/.test(r)) setTimeout(function () { google.script.run.showWebsite(); }, 900); })     // reopen with the right links
          .saveWebsiteAddress(document.getElementById('web').value); }
      function save() { var b = document.getElementById('sv'); b.disabled = true;
        google.script.run.withSuccessHandler(function (r) { document.getElementById('appmsg').textContent = r; b.disabled = false; })
          .saveAppAddress(document.getElementById('app').value); }
    </script>`;
  dlg_('🌐 START website', html, 480, (appUrl ? 690 : 640) + (fixed ? 0 : 60));
}


// ━━━━━━━━━━━━━━━━━━━━  Game  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  🎮 GAME (website): XP, levels, badges, streak saver, the weekly journey, postcards, the weekly goal,
 *  the school leaderboard. Everything is worked out from the Check-in answers tab: nothing new is stored.
 * ================================================================ */
var GAME = { rankMin: 3, xpCheckin: 40, xpOnTime: 10, xpEvent: 10, journeyWeeks: 12, mapBack: 10, mapAhead: 2 };   // check-in 50 XP (+10 per START event); a brain break up to 10
/* ---------- the XP standard (one table for every game) ----------
 *   Weekly check-in ........ 50 XP, +10 for each START event joined (the biggest reward: it is what START needs most)
 *   Today's brain break ..... by stars: 3★ 10 · 2★ 6 · 1★ 3 · 0★ 0   (same for every game)
 *   Catch-up (an overdue day) by stars: 3★ 4  · 2★ 2 · 1★ 1 · 0★ 0   (no day-streak credit)
 * Stars: word scramble solved 3; odd one out / number pattern 1 per right round; memory pairs 3 all found in few moves,
 * 2 all found, 1 half or more. */
var DAILY_XP = { today: [0, 3, 6, 10], catchup: [0, 1, 2, 4] };
function dailyXp_(stars, catchup) { const t = catchup ? DAILY_XP.catchup : DAILY_XP.today; return t[Math.max(0, Math.min(3, Math.floor(Number(stars) || 0)))]; }
var LEVELS_DAVAO = ['Bagong Agila', 'Batang Apo', 'Lakad Samal', 'Agila sa Eden', 'Bantay Dabaw', 'Agila ng Apo'];
var LEVELS_ANY = ['Rookie', 'Explorer', 'Achiever', 'Trailblazer', 'Leader', 'Legend'];
var LEVEL_XP = [0, 100, 250, 450, 700, 1000];

// mine = [{ t: Date, ev: events joined, late: true/false }] for one scholar
function webGame_(ss, mine, spec, daily) {
  daily = daily || { byDay: {}, total: 0, wins: 0 };
  const davao = typeof placesOn_ === 'function' && placesOn_();
  const byWeek = {};
  mine.forEach(r => { const k = memberWeekKey_(r.t); if (!byWeek[k] || r.t > byWeek[k].t) byWeek[k] = Object.assign({ k }, r); });
  const keys = Object.keys(byWeek).sort();
  const nowW = memberWeek_(new Date()), nowKey = ymd_(nowW.due);
  const weekBefore = k => { const d = new Date(k + 'T12:00:00'); d.setDate(d.getDate() - 7); return ymd_(d); };
  let ts = null; try { ts = termStart_(ss); } catch (e) { ts = null; }
  const termKey = ts instanceof Date ? ymd_(memberWeek_(ts).due) : '';
  // streak, with one "streak saver" per term: a single missed week inside this term does not break it
  let k = byWeek[nowKey] ? nowKey : weekBefore(nowKey), streak = 0, saved = '';
  for (let i = 0; i < 104; i++) {
    if (byWeek[k]) streak++;
    else if (!saved && termKey && k >= termKey && byWeek[weekBefore(k)]) saved = k;
    else break;
    k = weekBefore(k);
  }
  // best run (no saver), totals, XP
  let best = 0, run = 0, prev = '', events = 0, onTime = 0, xp = daily.total, puzzles = daily.wins;
  keys.forEach(x => {
    run = prev && weekBefore(x) === prev ? run + 1 : 1; prev = x; best = Math.max(best, run);
    const r = byWeek[x]; events += r.ev; if (!r.late) onTime++;
    xp += GAME.xpCheckin + (r.late ? 0 : GAME.xpOnTime) + GAME.xpEvent * r.ev;
  });
  const names = davao ? LEVELS_DAVAO : LEVELS_ANY;
  let lv = 0; LEVEL_XP.forEach((need, i) => { if (xp >= need) lv = i; });
  const nextXp = LEVEL_XP[lv + 1] || null;
  const level = { n: lv + 1, title: names[lv], next: nextXp ? names[lv + 1] : '', xp, from: LEVEL_XP[lv], to: nextXp,
                  pct: nextXp ? Math.round(100 * (xp - LEVEL_XP[lv]) / (nextXp - LEVEL_XP[lv])) : 100 };
  // postcards: the place of every week answered (Region XI)
  const cards = {};
  if (davao) keys.forEach(x => { const d = new Date(x + 'T12:00:00'); d.setDate(d.getDate() - 6); cards[placeIndexForDate_(d)] = true; });
  const nCards = Object.keys(cards).length, allCards = davao ? PLACE_PICS.length : 0;
  const runBest = Math.max(best, streak);
  // [id, name, how to get it, now, target, icon, tier]
  const badges = [
    ['first', 'First check-in', 'Answer your first weekly check-in', keys.length, 1, 'check', 'bronze'],
    ['s3', '3 in a row', 'Check in 3 weeks in a row', runBest, 3, 'flame', 'bronze'],
    ['s5', '5 in a row', 'Check in 5 weeks in a row', runBest, 5, 'flame', 'silver'],
    ['s10', '10 in a row', 'Check in 10 weeks in a row', runBest, 10, 'flame', 'gold'],
    ['ev1', 'Event goer', 'Join a START event', events, 1, 'events', 'bronze'],
    ['ev5', 'Regular', 'Join 5 START events', events, 5, 'events', 'silver'],
    ['time5', 'Always on time', 'Check in on time 5 times', onTime, 5, 'clock', 'silver'],
    ['mini5', 'Puzzle pro', 'Win 5 daily brain breaks', puzzles, 5, 'bolt', 'silver'],
    davao ? ['cards', 'Postcard collector', 'Collect all ' + allCards + ' place postcards', nCards, allCards, 'star', 'gold']
          : ['ten', '10 check-ins', 'Answer 10 weekly check-ins', keys.length, 10, 'star', 'gold'],
  ].map(b => ({ id: b[0], name: b[1], desc: b[2], got: b[3] >= b[4], now: Math.min(b[3], b[4]), need: b[4], icon: b[5], tier: b[6] }));
  // the journey: the last weeks of this term as a path (done, late, missed, this week, next week)
  const journey = [];
  let j = ymd_(new Date(nowW.due.getTime() + 7 * 86400000));
  for (let i = 0; i < GAME.journeyWeeks; i++) {
    if (termKey && j < termKey) break;
    const due = new Date(j + 'T12:00:00'), mon = new Date(due.getTime() - 6 * 86400000);
    const r = byWeek[j];
    journey.push({ key: j, label: webFmt_(mon, 'MMM d'), state: j > nowKey ? 'next' : j === nowKey ? (r ? 'done' : 'now') : r ? (r.late ? 'late' : 'done') : (j === saved ? 'saved' : 'missed'),
                   place: davao ? PLACE_PICS[placeIndexForDate_(mon)].name : '' });
    j = weekBefore(j);
  }
  // this week's goal: the check-in plus each START event of the week
  const evThisWeek = spec && spec.q ? ((spec.q.find(q => q.k === 'events') || { choices: [] }).choices.length - (spec.q.some(q => q.k === 'events') ? 1 : 0)) : 0;
  const me = byWeek[nowKey];
  const goal = { total: 1 + Math.max(0, evThisWeek), done: (me ? 1 : 0) + (me ? Math.min(me.ev, Math.max(0, evThisWeek)) : 0) };
  const days = webDayMap_(ss, byWeek, daily, spec);
  const extra = webDayStreak_(byWeek, daily);
  return { xp, level, streak, best, puzzles, days, dayStreak: extra.streak, recap: extra.recap, saver: { used: !!saved, available: !!termKey && !saved }, events, checkins: keys.length,
           badges, journey: journey.reverse(), cards: Object.keys(cards).map(Number), nCards, allCards, goal };
}

/* ---------- XP rewards ----------
 * The RR edits Rewards in the spreadsheet. Requests keep a price snapshot and reserve XP
 * while pending, so two requests cannot spend the same points. Only an RR can decide.
 */
var REWARD_TAB = 'Rewards';
var REWARD_REQUEST_TAB = 'Reward requests';
var REWARD_HEAD = ['Reward', 'XP cost', 'Details', 'Active (Y/N)'];
var REWARD_REQUEST_HEAD = ['Requested at', 'Request ID', 'Scholar email', 'School', 'Reward', 'XP cost', 'Status', 'Decided at', 'Decision note'];
function ensureRewardTabs_(ss) {
  [[REWARD_TAB, REWARD_HEAD], [REWARD_REQUEST_TAB, REWARD_REQUEST_HEAD]].forEach(pair => {
    let sh = ss.getSheetByName(pair[0]);
    if (!sh) sh = ss.insertSheet(pair[0]);
    sh.getRange(1, 1, 1, pair[1].length).setValues([pair[1]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
}
function rewardCatalog_(ss) {
  const sh = ss.getSheetByName(REWARD_TAB);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, REWARD_HEAD.length).getValues().map((r, i) => ({
    id: String(i + 2), name: String(r[0] || '').trim().slice(0, 100), cost: Number(r[1]),
    details: String(r[2] || '').trim().slice(0, 300), active: /^(y|yes|true|1)$/i.test(String(r[3] || '').trim())
  })).filter(r => r.name && Number.isInteger(r.cost) && r.cost > 0 && r.active);
}
function rewardRequests_(ss) {
  const sh = ss.getSheetByName(REWARD_REQUEST_TAB);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, REWARD_REQUEST_HEAD.length).getValues().map((r, i) => ({
    row: i + 2, at: r[0], id: String(r[1]), email: emailOf_(r[2]), school: String(r[3]),
    name: String(r[4]), cost: Number(r[5]) || 0, status: String(r[6]), decidedAt: r[7], note: String(r[8] || '')
  }));
}
function rewardEarned_(ss, email) {
  return webGame_(ss, webMyAnswers_(ss, email), null, webMyDaily_(ss, email)).xp;
}
function rewardState_(ss, email) {
  const mine = rewardRequests_(ss).filter(r => r.email === email);
  const earned = rewardEarned_(ss, email);
  const reserved = mine.filter(r => r.status === 'Pending' || r.status === 'Approved').reduce((n, r) => n + r.cost, 0);
  return { earned, available: Math.max(0, earned - reserved), rewards: rewardCatalog_(ss),
    requests: mine.slice(-10).reverse().map(r => ({ id: r.id, name: r.name, cost: r.cost, status: r.status,
      when: webFmt_(r.at, 'MMM d, yyyy'), note: r.note })) };
}
function webRewards(token) {
  return webSafe_(() => {
    const session = webSession_(token);
    if (!session) return { ok: false, signedOut: true };
    const ss = webSheet_(), who = webRoles_(ss, session.e);
    if (!who.member || who.suspended) return { ok: false, msg: 'Only an active scholar can redeem XP.' };
    return { ok: true, state: rewardState_(ss, who.member.primary || who.email) };
  });
}
function webRewardRequest(token, rewardId) {
  return webSafe_(() => {
    const session = webSession_(token);
    if (!session) return { ok: false, signedOut: true };
    const ss = webSheet_(), who = webRoles_(ss, session.e);
    if (!who.member || who.suspended) return { ok: false, msg: 'Only an active scholar can redeem XP.' };
    const lock = docLock_();
    if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try again.' };
    try {
      const email = who.member.primary || who.email;
      const reward = rewardCatalog_(ss).find(r => r.id === String(rewardId));
      if (!reward) return { ok: false, msg: 'That reward is no longer available.' };
      const state = rewardState_(ss, email);
      if (state.available < reward.cost) return { ok: false, msg: 'You need ' + (reward.cost - state.available) + ' more XP for this reward.' };
      ensureRewardTabs_(ss);
      ss.getSheetByName(REWARD_REQUEST_TAB).appendRow([new Date(), Utilities.getUuid(), email, who.member.school || '', reward.name,
        reward.cost, 'Pending', '', '']);
      webBumpData_();
      return { ok: true, msg: 'Request sent to your Regional Representative.', state: rewardState_(ss, email) };
    } finally { lock.releaseLock(); }
  });
}
function rewardPending_(ss) {
  return rewardRequests_(ss).filter(r => r.status === 'Pending').map(r => ({ id: r.id, email: r.email, school: r.school,
    name: r.name, cost: r.cost, when: webFmt_(r.at, 'MMM d, yyyy') }));
}
function rewardDecide_(ss, id, decision, note) {
  if (decision !== 'Approved' && decision !== 'Declined') return { ok: false, msg: 'Choose Approve or Decline.' };
  const lock = docLock_();
  if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try again.' };
  try {
    const req = rewardRequests_(ss).find(r => r.id === String(id));
    if (!req || req.status !== 'Pending') return { ok: false, msg: 'This request is no longer pending.' };
    if (decision === 'Approved') {
      const scholar = scholarByEmail_(ss)[req.email];
      if (!scholar || scholar.status !== 'Active') return { ok: false, msg: 'This scholar is no longer active. Decline the request instead.' };
      const state = rewardState_(ss, req.email);
      // Pending requests already reserve their cost. A changed XP total must still cover all reservations.
      if (state.earned < rewardRequests_(ss).filter(r => r.email === req.email &&
          (r.status === 'Pending' || r.status === 'Approved')).reduce((n, r) => n + r.cost, 0))
        return { ok: false, msg: 'This scholar no longer has enough XP. Decline the request instead.' };
    }
    const sh = ss.getSheetByName(REWARD_REQUEST_TAB);
    sh.getRange(req.row, 7, 1, 3).setValues([[decision, new Date(), String(note || '').trim().slice(0, 300)]]);
    webBumpData_();
    return { ok: true, msg: decision === 'Approved' ? 'Reward approved. Please arrange delivery with the scholar.' : 'Request declined; XP is available again.' };
  } finally { lock.releaseLock(); }
}

/* ---------- confirmed attendance (optional): the "Event attendance" tab ----------
 * Paste who really came to an event (e.g. from Google Meet or a sign-in sheet): Event title · Date · Email.
 * For an event listed there, only those people get event XP and badges. Events not listed: what scholars ticked counts. */
var ATTEND_TAB = 'Event attendance';
var ATTEND_HEAD = ['Event title', 'Date', 'Email'];
function attendKey_(title, date) {
  const t = String(title || '').trim().toLowerCase().replace(/\s*\([^)]*\)\s*$/, '');
  const d = date instanceof Date ? Utilities.formatDate(date, CONFIG.timezone, 'MMM d') : String(date || '').replace(/^[A-Za-z]{3},\s*/, '').trim();
  return t + '|' + d.toLowerCase();
}
// { "title|mmm d": { email: true } } for every event someone listed
function attendanceMap_(ss) {
  const sh = ss.getSheetByName(ATTEND_TAB), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(r => {
    const e = emailOf_(r[2]); if (!String(r[0]).trim() || !e) return;
    const k = attendKey_(r[0], r[1]); (out[k] = out[k] || {})[e] = true;
  });
  return out;
}
// How many of the ticked events count: confirmed ones need the person on the list
function countedEvents_(ticked, email, att) {
  return splitEvents_(ticked).filter(x => x && !missedPick_(x)).filter(label => {
    const i = label.lastIndexOf(' · '), k = i < 0 ? null : attendKey_(label.slice(0, i), label.slice(i + 3));
    return !k || !att[k] || !!att[k][email];
  }).length;
}
function ensureAttendanceTab_(ss) {
  let sh = ss.getSheetByName(ATTEND_TAB);
  if (sh) return sh;
  sh = ss.insertSheet(ATTEND_TAB);
  sh.getRange(1, 1, 1, 3).setValues([ATTEND_HEAD]);
  sh.setFrozenRows(1);
  sh.getRange('B:B').setNumberFormat('mmm d, yyyy');
  try { sh.getRange(1, 1).setNote('Optional. Paste who attended an event (Event title as in the Events tab · Date · Email). ' +
    'For a listed event, only these people get event XP on the website. Unlisted events count what scholars ticked.'); } catch (e) { /* no note */ }
  return sh;
}

// One scholar's answers in the Check-in answers tab: [{ t, ev, late }]
function webMyAnswers_(ss, email) {
  const sh = memberTab_(ss), out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  const c = memberCols_(sh);
  if (c.email < 0) return out;
  const head = sh.getRange(1, 1, 1, c.width).getValues()[0].map(x => plainHeading_(x));
  const lateCol = head.indexOf(typeof LATE_HEAD === 'string' ? LATE_HEAD : 'Answered late on');
  let att = {}; try { att = attendanceMap_(ss); } catch (e) { att = {}; }
  sh.getRange(2, 1, sh.getLastRow() - 1, c.width).getValues().forEach(v => {
    if (!(v[0] instanceof Date) || emailOf_(v[c.email]) !== email) return;
    const ev = c.events >= 0 ? countedEvents_(v[c.events], email, att) : 0;
    out.push({ t: v[0], ev, late: lateCol >= 0 && !!v[lateCol] });
  });
  return out;
}

/* ---------- 🧑‍🏫 Campus Lead tools (website) ----------
 * Campus Leads take part of the follow-up work: they mark who they contacted, check join requests for their school,
 * take event attendance and post a short note to their members. Everything is kept in the "Lead actions" tab
 * (attendance also goes to "Event attendance", which already decides event XP). A Lead acts only for their own school. */
var LEAD_TAB = 'Lead actions';
var LEAD_HEAD = ['At', 'Week', 'School', 'Lead email', 'Action', 'About (email)', 'Value', 'Note'];
function leadTab_(ss) {
  let sh = ss.getSheetByName(LEAD_TAB);
  if (!sh) { sh = ss.insertSheet(LEAD_TAB); sh.getRange(1, 1, 1, LEAD_HEAD.length).setValues([LEAD_HEAD]).setFontWeight('bold'); sh.setFrozenRows(1); sh.getRange(1, 2, sh.getMaxRows(), 1).setNumberFormat('@'); }
  return sh;
}
function leadRows_(ss) {
  const sh = ss.getSheetByName(LEAD_TAB);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, LEAD_HEAD.length).getValues().map((r, i) => ({ row: i + 2, at: r[0], week: String(r[1]), school: String(r[2]),
    lead: emailOf_(r[3]), action: String(r[4]), about: emailOf_(r[5]) || String(r[5]), value: String(r[6]), note: String(r[7]) }));
}
// What the Lead pages need, read once: this week's "contacted" marks, pending join requests, attendance
function leadToolsData_(ss, schools) {
  const week = memberWeekKey_(new Date()), rows = leadRows_(ss), contacted = {};
  rows.filter(r => r.week === week && r.action === 'contacted').forEach(r => { contacted[r.school + '|' + r.about] = webFmt_(r.at, 'EEE'); });
  let joins = [];
  try {
    const checks = leadJoinChecks_(ss, rows);
    joins = webJoins_(ss).filter(j => !/^Added|^Declined/.test(j.status) && schools.indexOf(j.school) !== -1)
      .map(j => ({ row: j.row, school: j.school, name: j.name, program: j.program, year: j.year, awarded: j.awarded, when: j.when, check: checks[j.row] || null }));
  } catch (e) { Logger.log('Lead joins: ' + e.message); }
  let att = {}; try { att = attendanceMap_(ss); } catch (e) { att = {}; }
  return { contacted, joins, att };
}
// The latest Lead check for each join request row: { row: { value: 'yes'|'no', by, at, note } }
function leadJoinChecks_(ss, rows) {
  const out = {};
  (rows || leadRows_(ss)).filter(r => r.action === 'join-check').forEach(r => {
    out[Number(r.note.split('|')[0])] = { value: r.value, by: r.lead, at: webFmt_(r.at, 'EEE, MMM d'), note: r.note.split('|').slice(1).join('|') }; });
  return out;
}
// Events in the last 3 weeks for this school's members (all-member events and the school's own): who was marked present
function leadRecentEvents_(ss, school, att) {
  const sh = ss.getSheetByName(TAB.events);
  if (!sh || sh.getLastRow() < 2) return [];
  const today = startOfDay_(new Date()), from = new Date(today.getTime() - 21 * 86400000), to = new Date(today.getTime() + 86400000), out = [];
  sh.getRange(2, 1, sh.getLastRow() - 1, EV_EDIT).getValues().forEach(r => {
    const title = String(r[EV.title] || '').trim(), who = String(r[EV.who] || '').trim() || WHO.all;
    if (!title || !(who === WHO.all || who === school)) return;
    occurrences_(r, from, to).forEach(d => {
      const key = attendKey_(title, d);
      out.push({ key, title, date: ymd_(d), when: webFmt_(d, 'EEE, MMM d'), present: Object.keys((att || {})[key] || {}) });
    });
  });
  return out.sort((a, b) => b.date.localeCompare(a.date)).slice(0, 8);
}
function leadNote_(school) {
  try { const n = JSON.parse(docProps_().getProperty('LEAD_NOTE|' + school) || 'null');
    return n && n.t && Date.now() - n.at < 21 * 86400000 ? { text: n.t, by: n.by, when: webFmt_(new Date(n.at), 'EEE, MMM d') } : null; } catch (e) { return null; }
}
// Active members of a school, by email (from the check-in data, or the Scholars tab)
function schoolMemberEmails_(ss, school) {
  const out = {};
  try { if (checkinOn_()) (memberData_(ss).bySchool[school] || { members: [] }).members.forEach(m => { if (m.email) out[emailOf_(m.email)] = m.name; }); } catch (e) { /* Scholars below */ }
  if (!Object.keys(out).length) scholarRows_(ss).forEach(r => { if (String(r[1]).trim() === school && emailOf_(r[4])) out[emailOf_(r[4])] = r[0]; });
  return out;
}
// The RR's view: what each Campus Lead did this week
function leadBoard_(ss, schools) {
  const week = memberWeekKey_(new Date()), rows = leadRows_(ss).filter(r => r.week === week);
  return schools.filter(s => s.lead).map(s => {
    const mine = rows.filter(r => r.school === s.name);
    const reminderSent = doneList_('LEAD_DONE|' + s.name, week).some(k => k === 'remind' || k === 'last');
    return { school: s.name, lead: s.lead, checkedIn: !!s.leadDone, reminderSent,
             contacted: mine.filter(r => r.action === 'contacted').length,
             joinChecks: mine.filter(r => r.action === 'join-check').length, attendance: mine.filter(r => r.action === 'attendance').length,
             note: !!leadNote_(s.name), help: s.help || 0 };
  }).sort((a, b) => (b.contacted + b.joinChecks + b.attendance) - (a.contacted + a.joinChecks + a.attendance) || a.school.localeCompare(b.school));
}
// One Lead action. act: contacted · join-check · attendance · note. arg.school must be one of the Lead's schools.
function webLeadAction(token, act, arg) {
  return webSafe_(() => {
    const s = webSession_(token);
    if (!s) return { ok: false, signedOut: true };
    const ss = webSheet_(), who = webRoles_(ss, s.e);
    arg = arg || {};
    const school = String(arg.school || '');
    if (!(who.leads || []).length || who.leads.indexOf(school) === -1) return { ok: false, msg: 'Only this school\'s Campus Lead can do that.' };
    const week = memberWeekKey_(new Date()), sh = leadTab_(ss), lock = docLock_();
    if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try again.' };
    let msg = 'Saved.';
    try {
      const add = (action, about, value, note) => sh.appendRow([new Date(), week, school, who.email, action, about, value, String(note || '').slice(0, 300)]);
      if (act === 'contacted') {
        const email = emailOf_(arg.email);
        if (!email || !schoolMemberEmails_(ss, school)[email]) return { ok: false, msg: 'That member is not on your school\'s list.' };
        const old = leadRows_(ss).filter(r => r.week === week && r.school === school && r.action === 'contacted' && r.about === email);
        if (old.length) { old.reverse().forEach(r => sh.deleteRow(r.row)); msg = 'Unmarked.'; }      // tap again to undo
        else { add('contacted', email, arg.kind === 'help' ? 'help' : 'reminder'); msg = 'Marked as contacted.'; }
      } else if (act === 'join-check') {
        const row = Number(arg.row), value = arg.value === 'yes' ? 'yes' : arg.value === 'no' ? 'no' : '';
        const j = webJoins_(ss).find(x => x.row === row);
        if (!value || !j || j.school !== school || /^Added|^Declined/.test(j.status)) return { ok: false, msg: 'This request is no longer waiting for your school.' };
        add('join-check', j.email, value, row + '|' + String(arg.note || '').slice(0, 200));
        msg = value === 'yes' ? 'Thanks! Your RR will see that you confirmed them.' : 'Thanks! Your RR will see your note.';
      } else if (act === 'attendance') {
        const ev = leadRecentEvents_(ss, school, {}).find(e => e.title === String(arg.title || '') && e.date === String(arg.date || ''));
        if (!ev) return { ok: false, msg: 'Pick an event from the last three weeks.' };
        const members = schoolMemberEmails_(ss, school), chosen = (Array.isArray(arg.emails) ? arg.emails : []).map(emailOf_).filter(e => members[e]);
        const at = ensureAttendanceTab_(ss);
        if (at.getLastRow() > 1) {                                                     // replace this school's list for that event
          const vals = at.getRange(2, 1, at.getLastRow() - 1, 3).getValues();
          for (let i = vals.length - 1; i >= 0; i--) if (attendKey_(vals[i][0], vals[i][1]) === ev.key && members[emailOf_(vals[i][2])]) at.deleteRow(i + 2);
        }
        const d = new Date(ev.date + 'T12:00:00');
        if (chosen.length) at.getRange(at.getLastRow() + 1, 1, chosen.length, 3).setValues(chosen.map(e => [ev.title, d, e]));
        add('attendance', '', chosen.length, ev.title + ' · ' + ev.when);
        msg = chosen.length + ' marked present for ' + ev.title + '.';
      } else if (act === 'todo-done') {
        const key = String(arg.key || '').slice(0, 120);
        if (!/^(help\|.+|last|remind|noemail)$/.test(key)) return { ok: false, msg: 'That to-do item is gone. Reload the page.' };
        doneToggle_('LEAD_DONE|' + school, week, key, !!arg.undo);
        add('done', key.indexOf('help|') === 0 ? key.slice(5) : '', arg.undo ? 'undone' : key.split('|')[0]);
        msg = arg.undo ? 'Back on your to-do list.' : 'Done! Removed from your to-do list.';
      } else if (act === 'note') {
        const text = String(arg.text || '').replace(/\s+/g, ' ').trim().slice(0, 300);
        if (text) docProps_().setProperty('LEAD_NOTE|' + school, JSON.stringify({ t: text, by: displayName_(who.name || who.email), at: Date.now() }));
        else docProps_().deleteProperty('LEAD_NOTE|' + school);
        add('note', '', text ? 'posted' : 'cleared', text);
        msg = text ? 'Your note is on your members\' Home page.' : 'Note removed.';
      } else return { ok: false, msg: 'Unknown action.' };
    } finally { lock.releaseLock(); }
    webBumpData_();
    try { webForget_(who.email); } catch (e) { /* cache */ }
    return { ok: true, msg, lead: JSON.parse(JSON.stringify(webLead_(ss, who.leads, getSettings_()))) };
  });
}

// Schools ranked: this week's members answered (%), then the school's streak
function webLeaderboard_(ss, mySchools) {
  const data = focalData_(ss).filter(s => !s.inactive);
  let md = null; try { if (checkinOn_()) md = memberData_(ss); } catch (e) { md = null; }
  const stars = (() => { try { return dailyStarsBySchool_(ss); } catch (e) { return {}; } })();
  const all = data.map(s => {
    const m = (md && md.bySchool[s.name]) || { members: [], answered: 0 };
    const n = md ? m.members.length : 0;
    return { name: s.name, rate: n ? Math.round(100 * m.answered / n) : 0, answered: m.answered || 0, members: n, streak: s.streak || 0, stars: stars[s.name] || 0,
             me: (mySchools || []).indexOf(s.name) !== -1 };
  });
  // fair: a school needs at least GAME.rankMin members to be ranked (1 of 1 is not "100%"); ties go to more answers, then the streak
  const ranked = all.filter(s => s.members >= GAME.rankMin)
    .sort((a, b) => (b.rate - a.rate) || (b.answered - a.answered) || (b.stars - a.stars) || (b.streak - a.streak) || a.name.localeCompare(b.name))
    .map((s, i) => Object.assign(s, { rank: i + 1 }));
  const small = all.filter(s => s.members < GAME.rankMin).sort((a, b) => a.name.localeCompare(b.name))
    .map(s => Object.assign(s, { rank: null, small: true }));
  return ranked.concat(small);
}

// What one person sees: the RR sees every school; everyone else the top 3 and their own school(s) (no one is shamed at the bottom)
function webRanksFor_(list, admin) {
  if (admin) return { ranks: list, hidden: 0 };
  const keep = list.filter(s => (s.rank && s.rank <= 3) || s.me);
  return { ranks: keep, hidden: list.length - keep.length };
}

// Keep the postcard list tiny. A visible, unlocked card asks for its picture separately.
function webPostcards(token, asEmail) {
  return webSafe_(() => {
    const s = webSession_(token);
    if (!s) return { ok: false, signedOut: true };
    const ss = webSheet_();
    let who = webRoles_(ss, s.e);
    if (asEmail && who.admin) { who = webRoles_(ss, asEmail); s.e = who.email; }        // RR "View as" a member
    if (!who.member || !placesOn_()) return { ok: true, cards: [] };
    const got = {}; webGame_(ss, webMyAnswers_(ss, s.e), null).cards.forEach(i => { got[i] = true; });
    return { ok: true, cards: PLACE_PICS.map((p, i) => ({ index: i, name: p.name, got: !!got[i] })) };
  });
}

// These are decorative pictures only. Publish a copy once, then return its small URL.
// If public sharing is unavailable, return just this requested image as a fallback.
function webPostcardImage(token, index, asEmail) {
  return webSafe_(() => {
    const session = webSession_(token);
    if (!session) return { ok: false, signedOut: true };
    const ss = webSheet_(), actor = webRoles_(ss, session.e);
    const who = asEmail && actor.admin ? webRoles_(ss, asEmail) : actor;
    const i = Number(index);
    if (!who.member || !placesOn_() || !Number.isInteger(i) || i < 0 || i >= PLACE_PICS.length)
      return { ok: false, msg: 'Postcard unavailable.' };
    const got = webGame_(ss, webMyAnswers_(ss, who.email), null).cards;
    if (got.indexOf(i) === -1) return { ok: false, msg: 'Postcard locked.' };
    const pic = PLACE_PICS[i], key = 'WEB_POSTCARD_' + i, props = docProps_();
    let id = props.getProperty(key);
    if (id) try {
      const file = DriveApp.getFileById(id);
      if (!file.isTrashed() && file.getSharingAccess() === DriveApp.Access.ANYONE_WITH_LINK)
        return { ok: true, src: 'https://lh3.googleusercontent.com/d/' + id + '=w960' };
    } catch (e) { id = ''; }
    try {
      const raw = String(pic.src || '').match(/^data:image\/webp;base64,(.+)$/);
      if (!raw) throw new Error('Picture source missing');
      const file = DriveApp.createFile(Utilities.newBlob(Utilities.base64Decode(raw[1]), 'image/webp',
        'START postcard ' + i + '.webp'));
      try { file.moveTo(pictureFolder_(true)); } catch (e) { /* keep in My Drive */ }
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
      catch (e) { file.setTrashed(true); throw e; }
      props.setProperty(key, file.getId());
      return { ok: true, src: 'https://lh3.googleusercontent.com/d/' + file.getId() + '=w960' };
    } catch (e) {
      Logger.log('Postcard picture link: ' + e.message);
      return { ok: true, src: pic.src || '' };
    }
  });
}

// RR edits to Schools use the existing row as an optimistic concurrency check, then propagate renames.
function webSchoolSave_(ss, actor, arg) {
  arg = arg || {};
  const sh = ss.getSheetByName(TAB.schools), row = arg.row == null || arg.row === '' ? 0 : Number(arg.row);
  if (!sh) return { ok: false, msg: 'The Schools database is missing. Run setup first.' };
  const clean = (v, n) => String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, n);
  const name = clean(arg.name, 80), province = clean(arg.province, 100), lead = clean(arg.lead, 100);
  const info = Array.isArray(arg.info) && arg.info.length === SCHOOL_INFO_HEADERS.length ? arg.info.map(v => clean(v, 250)) : null;
  if (!name || !info || [name, province, lead].concat(info).some(v => /^=/.test(v)))
    return { ok: false, msg: 'Enter a school name and valid details.' };
  if (!Number.isInteger(row) || row < 0 || row === 1 || row > CONFIG.maxSchools + 1)
    return { ok: false, msg: 'School row is invalid. Reload and try again.' };
  const lock = docLock_();
  if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try again.' };
  let oldName = '', oldLead = '', oldEmail = '';
  try {
    const values = sh.getRange(2, 1, CONFIG.maxSchools, 6).getValues();
    const index = row ? row - 2 : values.findIndex(v => !String(v[0] || '').trim());
    if (index < 0) return { ok: false, msg: 'The Schools database is full.' };
    const target = index + 2, old = values[index];
    oldName = String(old[0] || '').trim(); oldLead = String(old[3] || '').trim(); oldEmail = emailOf_(old[4]);
    if ((row && oldName !== String(arg.original || '').trim()) || (!row && oldName))
      return { ok: false, msg: 'The school list changed. Reload and try again.' };
    if (values.some((v, i) => i !== index && String(v[0] || '').trim().toLowerCase() === name.toLowerCase()))
      return { ok: false, msg: 'That school is already listed.' };
    if (lead && values.some((v, i) => i !== index && String(v[3] || '').trim().toLowerCase() === lead.toLowerCase()))
      return { ok: false, msg: 'That Campus Lead is already assigned to another school.' };
    const scholar = lead ? findScholar_(scholarRows_(ss), lead) : null;
    if (scholar && String(scholar[1] || '').trim() !== name && String(scholar[1] || '').trim() !== oldName)
      return { ok: false, msg: 'That scholar belongs to another school.' };
    sh.getRange(target, 1, 1, 2).setValues([[name, province]]);
    sh.getRange(target, 4).setValue(scholar ? String(scholar[0]).trim() : lead);
    sh.getRange(target, 6).setValue(arg.inactive === true ? 'Y' : 'N');
    sh.getRange(target, 9, 1, SCHOOL_INFO_HEADERS.length).setValues([info]);
    if (oldName && oldName !== name) renameSchool_(ss, oldName, name);
    SpreadsheetApp.flush();
    syncRoles_(ss);
  } finally { lock.releaseLock(); }
  webBumpData_();
  if (oldEmail) webForget_(oldEmail);
  try { updateDropdown_(); } catch (e) { Logger.log('School form list: ' + e.message); }
  try { calPeopleChanged_(ss); } catch (e) { Logger.log('Calendar sharing: ' + e.message); }
  Logger.log('RR saved school ' + name + (oldName ? ' (was ' + oldName + ')' : ' (new)') + ' by ' + actor);
  return { ok: true, msg: 'School saved.' };
}

function webSettingSave_(ss, actor, arg) {
  arg = arg || {};
  const index = SETTINGS.findIndex(s => s[0] === arg.key), sh = ss.getSheetByName(TAB.settings);
  if (!sh || index < 0) return { ok: false, msg: 'Setting not found. Reload and try again.' };
  const key = SETTINGS[index][0], spec = webSettingRecords_(ss)[index];
  if (spec.readonly) return { ok: false, msg: 'This value follows the check-in deadline automatically.' };
  let value = arg.value;
  if (spec.type === 'check') value = value === true;
  else if (spec.type === 'number') {
    value = Number(value);
    if (!Number.isInteger(value) || value < 1 || value > 365 || key === 'tRate' && value > 100)
      return { ok: false, msg: 'Enter a whole number in the allowed range.' };
  } else if (spec.type === 'date') {
    value = String(value || '').trim();
    if (value && !/^\d{4}-\d{2}-\d{2}$/.test(value)) return { ok: false, msg: 'Choose a valid date.' };
    if (value) {
      const d = new Date(value + 'T00:00:00+08:00');
      if (isNaN(d.getTime()) || webFmt_(d, 'yyyy-MM-dd') !== value) return { ok: false, msg: 'Choose a valid date.' };
      value = d;
    }
  } else {
    value = String(value == null ? '' : value).trim();
    if (value.length > 500 || /^=/.test(value)) return { ok: false, msg: 'That setting value is invalid.' };
    if (spec.options.length && spec.options.indexOf(value) === -1) return { ok: false, msg: 'Choose a listed value.' };
    if (key === 'replyTo' && value && emailOf_(value) !== value.toLowerCase()) return { ok: false, msg: 'Enter one valid reply address.' };
    if (key === 'recipients' && value && value.split(/[,;\s]+/).some(e => emailOf_(e) !== e.toLowerCase()))
      return { ok: false, msg: 'Enter valid summary email addresses, separated by commas.' };
  }
  const lock = docLock_();
  if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try again.' };
  try {
    const cell = sh.getRange(index + 2, 2), old = cell.getValue();
    const previous = old instanceof Date ? webFmt_(old, 'yyyy-MM-dd') : old == null ? '' : old;
    if (String(previous) !== String(arg.previous == null ? '' : arg.previous))
      return { ok: false, msg: 'This setting changed. Reload and try again.' };
    cell.setValue(value);
    if (key === 'autoApprove') setAutoApprove_(value);
    if (key === 'region') applyRegion_(ss);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  webBumpData_();
  if (['day', 'hour', 'remHour', 'dueDay', 'autoWeekly'].indexOf(key) !== -1) {
    try { if (claimIfFree_()) { applySchedule_(); docProps_().setProperty('SCHEDULE_SYNC', scheduleKey_()); } }
    catch (e) { Logger.log('Website schedule update: ' + e.message); return { ok: true, msg: 'Setting saved. Run setup to refresh the email schedule.' }; }
  }
  if (key === 'region') { try { updateDropdown_(); } catch (e) { Logger.log('Region form list: ' + e.message); } }
  Logger.log('RR saved setting ' + key + ' by ' + actor);
  if (key === 'autoApprove' && value === true) {
    const added = (() => { try { return approveWaitingJoins_(ss); } catch (e) { Logger.log('Auto-approve waiting: ' + e.message); return 0; } })();
    return { ok: true, msg: 'Setting saved.' + (added ? ' ' + added + ' waiting request' + (added === 1 ? ' was' : 's were') + ' approved now.' : '') };
  }
  return { ok: true, msg: 'Setting saved.' };
}

// RR on the website: update database records or send a check-in reminder.
function webAdminAction(token, action, arg) {
  return webSafe_(() => {
    const s = webSession_(token);
    if (!s) return { ok: false, signedOut: true };
    const ss = webSheet_(), who = webRoles_(ss, s.e);
    if (!who.admin || who.suspended) return { ok: false, msg: 'Only an active Regional Representative can do this.' };
    if (action === 'school-save') return webSchoolSave_(ss, s.e, arg);
    if (action === 'school-import') {
      if (!regionProfile_().schools.length) return { ok: false, msg: 'Choose a region in website Settings first.' };
      const lock = docLock_();
      if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try again.' };
      let count = 0;
      try { mergeUmMainSchool_(ss); count = writeSchoolList_(ss); syncRoles_(ss); }
      finally { lock.releaseLock(); }
      webBumpData_();
      try { updateDropdown_(); } catch (e) { Logger.log('School import form list: ' + e.message); }
      Logger.log('RR imported region schools by ' + s.e);
      return { ok: true, msg: count + ' region schools loaded.' };
    }
    if (action === 'attention-done') {
      const key = String((arg && arg.key) || '').slice(0, 200);
      if (!/^(quiet|check|report)\|/.test(key)) return { ok: false, msg: 'That item is gone. Reload the page.' };
      doneToggle_('RR_DONE', memberWeekKey_(new Date()), key, !!(arg && arg.undo));
      webForget_(s.e);
      return { ok: true, msg: arg && arg.undo ? 'Back on the list.' : 'Done! Removed from Needs your attention.' };
    }
    if (action === 'setting-save') return webSettingSave_(ss, s.e, arg);
    if (action === 'outreach-save') return webOutreachSave_(ss, s.e, arg);
    if (action === 'event-save') return webEventSave_(ss, s.e, arg);
    if (action === 'event-delete') return webEventDelete_(ss, s.e, arg);
    if (action === 'reward-decide') return rewardDecide_(ss, arg && arg.id, arg && arg.decision, arg && arg.note);
    if (action === 'scholar-status') return webScholarStatus_(ss, s.e, arg);
    if (action === 'scholar-update') return webScholarUpdate_(ss, s.e, arg);
    if (action === 'resetpw') return { ok: true, msg: resetWebPassword_(arg) };
    if (action === 'auto-approve') { const on = setAutoApprove_(arg === 'on');
      const added = on ? approveWaitingJoins_(ss) : 0;
      try { webForget_(s.e); } catch (e) { Logger.log('Auto-approve cache refresh: ' + e.message); }
      return { ok: true, msg: on ? 'Auto-approve is ON: everyone who registers from a listed school is added at once.' + (added ? ' ' + added + ' waiting request' + (added === 1 ? ' was' : 's were') + ' approved now.' : '') +
        ' Only clashes with an existing record wait for you.' : 'Auto-approve is OFF: every sign-up waits for you.' }; }
    if (action === 'join-add' || action === 'join-decline' || action === 'join-clarify') {
      webForget_(s.e); return webJoinAction_(ss, action, Object.assign({}, arg || {}, { actor: s.e, auto: false })); }
    if (action === 'remind' || action === 'followup') {
      const r = sendReminders_(action === 'followup' ? 'followup' : 'reminder');
      const m = r.members || { sent: 0, waiting: 0 };
      webForget_(s.e);
      return { ok: true, msg: `Sent to ${r.sent} Campus Lead${r.sent === 1 ? '' : 's'} and ${m.sent} member${m.sent === 1 ? '' : 's'}.` +
        (r.done ? ` ${r.done} already had it this week.` : '') + (r.waiting + m.waiting ? ` ${r.waiting + m.waiting} go out tomorrow morning (daily email limit).` : '') };
    }
    return { ok: false, msg: 'Unknown action' };
  });
}


/* ---------- weekly check-in rate: % of active members who answered, week by week ----------
 * [{ key, label, answered, members, rate, now }] oldest first. school = one school (Campus Lead page), none = the region. */
function weeklyRates_(ss, n, school) {
  const roster = {}, all = scholarByEmail_(ss);                              // read Scholars once
  Object.entries(all).forEach(([e, p]) => {
    if (p.status === 'Active' && p.primary === e && (!school || p.school === school)) roster[e] = true;
  });
  const snapshots = {};
  const snap = ss.getSheetByName(WEEKLY_ROSTER_TAB_);
  if (snap && snap.getLastRow() > 1) snap.getRange(2, 1, snap.getLastRow() - 1, 4).getDisplayValues().forEach(r => {
    const key = r[0], email = emailOf_(r[2]);
    if (!key) return;
    if (!snapshots[key]) snapshots[key] = {};
    if (email && (!school || r[3] === school)) snapshots[key][email] = true;
  });
  const byWeek = {};
  const sh = memberTab_(ss);
  if (sh && sh.getLastRow() > 1) {
    const c = memberCols_(sh);
    if (c.email >= 0) sh.getRange(2, 1, sh.getLastRow() - 1, c.width).getValues().forEach(v => {
      if (!(v[0] instanceof Date)) return;
      const raw = emailOf_(v[c.email]), e = (all[raw] && all[raw].primary) || raw;     // an "other" email counts for its scholar
      const k = memberWeekKey_(v[0]); (byWeek[k] = byWeek[k] || []).push([raw, e]);
    });
  }
  const out = [], due = new Date(memberWeek_(new Date()).due);
  for (let i = 0; i < (n || 10); i++) {
    const k = ymd_(due), mon = new Date(due.getTime() - 6 * 86400000), membersForWeek = snapshots[k] || roster;
    const answered = {};
    (byWeek[k] || []).forEach(pair => { if (membersForWeek[pair[0]] || membersForWeek[pair[1]]) answered[pair[1]] = true; });
    const members = Object.keys(membersForWeek).length, a = Object.keys(answered).length;
    out.unshift({ key: k, label: webFmt_(mon, 'MMM d'), answered: a, members,
      rate: members ? Math.round(100 * a / members) : 0, now: i === 0, estimated: !snapshots[k] });
    due.setDate(due.getDate() - 7);
  }
  return out;
}


// ━━━━━━━━━━━━━━━━━━━━  Join  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  JOIN REQUESTS  ·  someone who is not on the START list yet can ask to join from the sign-in page
 *  ----------------------------------------------------------------
 *  • The sign-in page never says whether an email is on the list (strangers can't test emails).
 *    Instead, the code screen offers "Not getting a code? Ask to join START".
 *  • Requests land in the "Join requests" tab and on the RR's website page (Add / Decline).
 *    The RR gets ONE summary email a day when there are new requests (saves the daily email limit).
 *  • Add: the person's row in Scholars is completed (email filled in, or made Active again), or a new row is
 *    added. They get a short "You're in" email and can create their password right away.
 *  Check the ONSSE masterlist before adding: only DOST-SEI scholars belong in START.
 * ================================================================ */

var JOIN_TAB = 'Join requests';
var JOIN_VIEWS = { pending: 'Join - Pending', approved: 'Join - Approved', declined: 'Join - Declined' };

// new columns go at the END so older tabs keep every value in its place
var JOIN_HEAD = ['Requested on', 'Name', 'School', 'Email', 'Status', 'Note', 'School (typed)', 'Program', 'Year level', 'Scholarship', 'Eligibility', 'RR note',
                 'Last name', 'First name', 'Middle initial', 'Year awarded (DOST)', 'Campus / Branch', 'Hometown'];
var JOIN = { perDay: 30, other: 'My school is not listed', otherProgram: 'My program is not listed',
             years: ['1st year', '2nd year', '3rd year', '4th year', '5th year', 'Junior level (JLSS)'],
             // DOST-SEI undergraduate scholarships: RA 7687 (S&T Scholarship Act of 1994), Merit (formerly NSDB/NSTA, RA 2067),
             // RA 10612 (Fast-Tracked S&T Scholarship Act of 2013, future science and math teachers). JLSS = Junior Level Science
             // Scholarship, given in 3rd year under one of the three.
             kinds: ['RA 7687', 'Merit', 'RA 10612', 'JLSS · RA 7687', 'JLSS · Merit', 'JLSS · RA 10612', 'Not sure'],
             pending: /^(New|Not eligible|Asked for details)/ };
var JOIN_SORRY = 'Thank you for your interest in START! We appreciate your willingness to become part of our scholar community. ' +
  'Based on the program information you provided, we are unable to confirm your eligibility for registration at this time. ' +
  'If you believe your degree program should be included, please contact your Campus Lead or Regional Representative for verification. ' +
  'Thank you for your understanding and interest!';

function joinTab_(ss, create) {
  let sh = ss.getSheetByName(JOIN_TAB);
  if (!sh && !create) return null;
  if (!sh) {
    sh = ss.insertSheet(JOIN_TAB);
    sh.getRange(1, 1, 1, JOIN_HEAD.length).setValues([JOIN_HEAD]);
    sh.setFrozenRows(1);
    try { sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm'); } catch (e) { /* plain dates */ }
    try { sh.getRange(1, 1, 1, JOIN_HEAD.length).protect().setDescription(PROTECT_TAG + ' heading').setWarningOnly(true); } catch (e) { /* fine */ }
    return sh;
  }
  try {                                                         // an older tab: the new headings are added at the end
    const w = sh.getLastColumn();
    if (w < JOIN_HEAD.length) {
      if (sh.getMaxColumns() < JOIN_HEAD.length) sh.insertColumnsAfter(sh.getMaxColumns(), JOIN_HEAD.length - sh.getMaxColumns());
      sh.getRange(1, w + 1, 1, JOIN_HEAD.length - w).setValues([JOIN_HEAD.slice(w)]);
    }
  } catch (e) { Logger.log('Join tab headings: ' + e.message); }
  return sh;
}

// The master Join requests tab is the only source of truth. These three sheets are
// read-only status views and can be rebuilt safely without deleting application history.
function refreshJoinViews_(ss) {
  const master = joinTab_(ss, false);
  const rows = master && master.getLastRow() > 1
    ? master.getRange(2, 1, master.getLastRow() - 1, JOIN_HEAD.length).getValues() : [];
  const buckets = { pending: [], approved: [], declined: [] };
  rows.forEach((r, i) => {
    const status = String(r[4] || '').trim();
    const key = /^Added/.test(status) ? 'approved' : /^Declined/.test(status) ? 'declined' : 'pending';
    buckets[key].push([i + 2].concat(r));
  });
  Object.keys(JOIN_VIEWS).forEach(key => {
    let tab = ss.getSheetByName(JOIN_VIEWS[key]);
    if (!tab) tab = ss.insertSheet(JOIN_VIEWS[key]);
    const width = JOIN_HEAD.length + 1;
    if (tab.getMaxColumns() < width) tab.insertColumnsAfter(tab.getMaxColumns(), width - tab.getMaxColumns());
    if (tab.getMaxRows() < buckets[key].length + 1) tab.insertRowsAfter(tab.getMaxRows(), buckets[key].length + 1 - tab.getMaxRows());
    tab.getRange(1, 1, Math.max(1, tab.getLastRow()), width).clearContent();
    tab.getRange(1, 1, 1, width).setValues([['Master row'].concat(JOIN_HEAD)]).setFontWeight('bold');
    if (buckets[key].length) tab.getRange(2, 1, buckets[key].length, width).setValues(buckets[key]);
    tab.setFrozenRows(1);
  });
  return { pending: buckets.pending.length, approved: buckets.approved.length, declined: buckets.declined.length };
}

// Run this once to create the history sheets for existing applications.
function setupJoinHistory() {
  sheetUser_();                                   // sheet only: never from the website
  const totals = refreshJoinViews_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getActiveSpreadsheet().toast('Join history updated: ' + JSON.stringify(totals), 'START Tracker', 7);
}

// What a school is called on the sign-up page. The University of Mindanao's main campus is spelled out, so it is never
// mixed up with UM Digos or UM Tagum.
function schoolLabel_(name, official) {
  name = String(name || '').trim(); official = String(official || '').trim();
  if (name === 'UIC') return 'University of the Immaculate Conception (UIC)';
  if (/^UM( Main)?$/i.test(name) || /^(The )?University of Mindanao$/i.test(official)) return 'University of Mindanao — UM Main';
  const um = name.match(/^UM\s+(.+)$/i);                                  // UM Digos, UM Tagum… each its own school
  if (um) return 'University of Mindanao — UM ' + um[1].replace(/\s*College$/i, '');
  if (!official || official.toLowerCase() === name.toLowerCase()) return name;
  return official.toLowerCase().indexOf(name.toLowerCase()) !== -1 ? official : official + ' (' + name + ')';
}
function joinSchoolList_(ss) {
  const off = {};
  try { ss.getSheetByName(TAB.schools).getRange(2, 1, CONFIG.maxSchools, 9).getValues().forEach(r => { off[String(r[0]).trim()] = r[8]; }); } catch (e) { /* names only */ }
  const schools = schoolRows_(ss);
  const list = schools.filter(s => s.name !== 'UM' || !schools.some(x => x.name === 'UM Main'))
    .map(s => ({ name: s.name, label: schoolLabel_(s.name, off[s.name]) }));
  const counts = {};
  list.forEach(s => { counts[s.label] = (counts[s.label] || 0) + 1; });
  return list.map(s => counts[s.label] > 1 ? { name: s.name, label: s.label + ' (' + s.name + ')' } : s)
    .sort((a, b) => a.label.localeCompare(b.label));
}

// Hometown choices: only this tracker's region (Region XI: the five Davao provinces and their cities / municipalities).
// A tracker with no region set shows the whole country.
function joinPlaces_() {
  let want = []; try { want = regionProfile_().provinces || []; } catch (e) { want = []; }
  const out = {};
  want.forEach(p => { if (PH_PLACES[p]) out[p] = PH_PLACES[p]; });
  return Object.keys(out).length ? out : PH_PLACES;
}

// Campus / Branch for Scholars › H, from the school picked (UM Main, UM Digos, UM Tagum…); blank for one-campus schools
function campusOf_(school) {
  const n = String(school || '').trim();
  if (/^UM( Main)?$/i.test(n)) return 'Main';
  const m = n.match(/^UM\s+(.+)$/i) || n.match(/\b(?:Obrero|Mintal|Tagum|Mabini|Digos|Bislig|Panabo|Mati)\b/i);
  return m ? (m[1] || m[0]).replace(/\s*College$/i, '') : '';
}

// Eligible when the program matches the 13 START-DOST courses exactly.
function joinEligibility_(program, typedProgram, schoolListed) {
  if (START_PROGRAMS.list.indexOf(String(program || '').trim()) === -1)
    return { result: 'Not eligible', why: 'program not on the START membership eligibility list' };
  return schoolListed
    ? { result: 'Eligible', why: '' }
    : { result: 'Needs review', why: 'school not on the list' };
}

// Sign-in page (not signed in): the schools and programs for the join form. School names are public.
function webJoinSchools() {
  return webSafe_(() => ({ ok: true, schools: joinSchoolList_(webSheet_()), other: '', programs: START_PROGRAMS.list, otherPrograms: [],
                           places: joinPlaces_(), nearProvinces: [],
                           otherProgram: JOIN.otherProgram, years: JOIN.years, kinds: JOIN.kinds,
                           awarded: (() => { const y = []; for (let k = new Date().getFullYear(); k >= 2015; k--) y.push(String(k)); return y; })() }));
}

// Sign-in page: "Ask to join START". The answer is the same whether or not the email is already on the list.
// d = { school, schoolTyped, program, programTyped, year, kind, note } (older pages send a plain school name + note)
function webJoinRequest(email, name, school, note, d) {
  return webSafe_(() => {
    const ss = webSheet_();
    d = d && typeof d === 'object' ? d : {};
    const clip = (x, n) => String(x || '').replace(/\s+/g, ' ').trim().slice(0, n);
    email = emailOf_(email);
    name = clip(name, 80);
    school = clip(school, 80);
    note = clip(note, 200);
    const schoolTyped = clip(d.schoolTyped, 120), program = clip(d.program, 140), programTyped = clip(d.programTyped, 140);
    // the Scholars details: "Last Name, First Name M." like the ONSSE masterlist, year awarded, campus, hometown
    const last = clip(d.last, 60), first = clip(d.first, 60), mi = clip(d.mi, 3).replace(/[^A-Za-zÑñ]/g, '').slice(0, 1).toUpperCase();   // one letter only
    const awarded = /^\d{4}$/.test(String(d.awarded || '')) && Number(d.awarded) >= 2010 && Number(d.awarded) <= new Date().getFullYear() + 1 ? String(d.awarded) : '';
    const campus = d.campus !== undefined ? clip(d.campus, 60) : campusOf_(school), city = clip(d.city, 60), province = clip(d.province, 60);
    const hometown = city && province ? city + ', ' + province : '';
    if (last || first) name = (last + ', ' + first + (mi ? ' ' + mi + '.' : '')).trim();
    const year = JOIN.years.indexOf(d.year) !== -1 ? d.year : '', kind = JOIN.kinds.indexOf(d.kind) !== -1 ? d.kind : '';
    if (!email) return { ok: false, msg: 'Enter your email address' };
    if (name.length < 3 || !/[A-Za-zÀ-ÿñÑ]/.test(name)) return { ok: false, msg: 'Enter your full name' };
    const schools = schoolRows_(ss).map(s => s.name);
    if (schools.indexOf(school) === -1 && school !== JOIN.other) return { ok: false, msg: 'Choose your school' };
    const newForm = true;  // Program is required for all new sign-ups, including requests from older website versions.
    if (newForm && school === JOIN.other && schoolTyped.length < 4) return { ok: false, msg: 'Please enter the complete name of your school.' };
    if (newForm) {
      if (START_PROGRAMS.list.indexOf(program) === -1) return { ok: false, msg: 'Choose one of the 13 START-eligible degree programs' };
      if (!year) return { ok: false, msg: 'Choose your year level' };
      if (d.province !== undefined && d.city !== undefined && d.kind !== undefined && d.campus === undefined) {   // the all-dropdown form
        const places = joinPlaces_();
        if (!places[d.province]) return { ok: false, msg: 'Choose your province' };
        if (places[d.province].indexOf(d.city) === -1) return { ok: false, msg: 'Choose your city or municipality' };
        if (JOIN.kinds.indexOf(d.kind) === -1) return { ok: false, msg: 'Choose your DOST scholarship' };
        if (school === JOIN.other) return { ok: false, msg: 'Choose your school' };
        if (START_PROGRAMS.list.indexOf(program) === -1) return { ok: false, msg: 'Choose one of the 13 START-eligible degree programs' };
      }
      if (d.last !== undefined) {                                     // the full sign-up form (older pages sent one name box)
        const okName = x => x.length >= 2 && /[A-Za-zÀ-ÿñÑ]/.test(x);
        if (!okName(last)) return { ok: false, msg: 'Enter your last name' };
        if (!okName(first)) return { ok: false, msg: 'Enter your first name' };
        if (!awarded) return { ok: false, msg: 'Choose the year your DOST scholarship was awarded' };
        if (!city || !province) return { ok: false, msg: 'Enter your hometown (city or municipality, and province)' };
      }
    }
    const el = newForm ? joinEligibility_(program, programTyped, school !== JOIN.other) : { result: 'Needs review', why: 'no program given' };
    const done = el.result === 'Not eligible' ? { ok: true, sorry: true, msg: JOIN_SORRY }
               : { ok: true, pending: true, msg: 'Your Regional Representative will review your details. Once approved, return here to set up your password.' };
    if (webAllowed_(webRoles_(ss, email))) return { ok: true, onList: true, hasPassword: !!(webAccount_(email) || {}).h, msg: 'Your START membership is already approved.' };
    const prior = joinStateOf_(ss, email);
    if (prior === 'waiting') {
      if (autoApproveOn_()) {
        try {
          const jsh = joinTab_(ss, false), jr = jsh.getRange(2, 1, jsh.getLastRow() - 1, 5).getValues();
          let at = -1; for (let q = jr.length - 1; q >= 0; q--) if (emailOf_(jr[q][3]) === email && /^(New|Asked for details)/.test(String(jr[q][4]).trim())) { at = q + 2; break; }
          const res = at > 0 ? webJoinAction_(ss, 'join-add', { row: at, auto: true }) : null;
          if (res && res.ok) { MEMBER_CACHE_ = null; return { ok: true, auto: true, hasPassword: !!(webAccount_(email) || {}).h, msg: 'Your START membership is approved. Confirm your email to set up your password.' }; }
        } catch (e) { Logger.log('Auto-approve on retry: ' + e.message); }
      }
      return { ok: true, pending: true, msg: 'Your Regional Representative is still reviewing this registration.' };
    }
    if (prior === 'added') return { ok: true, accessCheck: true, msg: 'Your request was approved, but account access needs a tracker check. Contact your Regional Representative. You do not need to register again.' };
    const props = docProps_(), now = Date.now();
    // one request per email a day, and a cap for the whole site (stops floods)
    const k = 'WEB_J_' + webHash_(email), lastAsk = Number(props.getProperty(k) || 0);
    // Duplicate requests are matched below; do not pretend a request was saved before writing it.
    const dayKey = 'WEB_JD_' + Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyyMMdd'), today = Number(props.getProperty(dayKey) || 0);
    if (today >= JOIN.perDay) return { ok: false, msg: 'START has received many requests today. Please try again tomorrow.' };
    const lock = docLock_();
    if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try again.' };
    let savedRow = -1;
    try {
      const sh = joinTab_(ss, true);
      const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, JOIN_HEAD.length).getValues() : [];
      const i = rows.findIndex(r => emailOf_(r[3]) === email && JOIN.pending.test(String(r[4]).trim()));   // one open request per email
      const v = [new Date(), name, school, email, el.result === 'Not eligible' ? 'Not eligible' : 'New', note,
                 school === JOIN.other ? schoolTyped : '', program,
                 year, kind, el.result + (el.why ? ' (' + el.why + ')' : ''), i >= 0 ? rows[i][11] : '', last, first, mi, awarded, campus, hometown];
      savedRow = i >= 0 ? i + 2 : sh.getLastRow() + 1;
      if (sh.getMaxRows() < savedRow) sh.insertRowsAfter(sh.getMaxRows(), savedRow - sh.getMaxRows());
      sh.getRange(savedRow, 1, 1, v.length).setValues([v]);                   // asked again: the newer details win
      try { markJoinChange_(ss, sh, savedRow, email, i >= 0 ? 'request updated' : 'request created', el.result); }
      catch (e) { Logger.log('Join audit: ' + e.message); }
      try { props.setProperty(k, String(now)); if (savedRow > 1) props.setProperty(dayKey, String(today + 1)); }
      catch (e) { Logger.log('Join rate counter: ' + e.message); }
      try { refreshJoinViews_(ss); } catch (e) { Logger.log('Join views: ' + e.message); }
    } finally { lock.releaseLock(); }
    // Valid applications from a listed school are activated immediately when the switch is on.
    // The applicant must still prove control of their email before choosing a password.
    if (autoApproveOn_() && el.result === 'Eligible' && school !== JOIN.other && savedRow > 1) {
      try {
        const r = webJoinAction_(ss, 'join-add', { row: savedRow, auto: true });
        if (r && r.ok) {
          MEMBER_CACHE_ = null;
          if (webAllowed_(webRoles_(ss, email)))
            return { ok: true, auto: true, hasPassword: !!(webAccount_(email) || {}).h, msg: 'Your START membership is approved. Confirm your email to set up your password.' };
          return { ok: true, accessCheck: true, msg: 'Your request was approved, but account access needs a tracker check. Contact your Regional Representative. You do not need to register again.' };
        }
      } catch (e) { Logger.log('Auto-approve: ' + e.message); }                  // not added: it stays in the list for the RR
      // A later notification or status-view error must not make a completed approval look pending.
      MEMBER_CACHE_ = null;
      try { if (webAllowed_(webRoles_(ss, email))) return { ok: true, auto: true, hasPassword: !!(webAccount_(email) || {}).h, msg: 'Your START membership is approved. Confirm your email to set up your password.' }; }
      catch (e) { Logger.log('Auto-approve status: ' + e.message); }
    }
    // Requests that cannot be added safely remain for RR review.
    return done;
  });
}

// Read-only recovery after a mobile connection drops while registration is being saved.
function webJoinStatus(email) {
  return webSafe_(() => {
    const ss = webSheet_(), main = primaryEmail_(ss, email);
    if (!main) return { ok: false, msg: 'Enter a valid email address.' };
    const who = webRoles_(ss, main), account = webAccount_(main);
    if (webAllowed_(who)) return { ok: true, state: account && account.h ? 'password' : 'setup' };
    const state = joinStateOf_(ss, main);
    return { ok: true, state: state === 'waiting' ? 'pending' : state === 'added' ? 'access-check' : state === 'sorry' ? 'noteligible' : state === 'declined' ? 'declined' : 'notfound' };
  });
}

// Open requests for the RR's website page, with what "Add" will do
function webJoins_(ss) {
  const sh = joinTab_(ss, false);
  if (!sh || sh.getLastRow() < 2) return [];
  const sc = ss.getSheetByName(TAB.scholars);
  const roster = sc && sc.getLastRow() > 1 ? sc.getRange(2, 1, sc.getLastRow() - 1, 6).getValues() : [];
  const labels = {}; joinSchoolList_(ss).forEach(x => { labels[x.name] = x.label; });
  const old = oldMasterRows_(ss);
  const w = Math.min(Math.max(sh.getLastColumn(), 6), JOIN_HEAD.length);
  return sh.getRange(2, 1, sh.getLastRow() - 1, w).getValues().map((r, i) => ({ r: r.concat(new Array(JOIN_HEAD.length - r.length).fill('')), row: i + 2 }))
    .map(x => {
      const typed = String(x.r[2]).trim() === JOIN.other;
      const oldHit = oldMasterMatch_(old, x.r[1], typed ? '' : x.r[2]);
      const m = typed ? { plan: 'School typed by the applicant: match it to a school in the list, or add the school first' } : joinMatch_(roster, x.r[1], x.r[2], x.r[3]);
      const el = String(x.r[10] || 'Needs review (no program given)');
      return { row: x.row, name: String(x.r[1]), school: typed ? String(x.r[6]) : String(x.r[2]), schoolLabel: typed ? String(x.r[6]) : (labels[x.r[2]] || String(x.r[2])),
               typed, email: emailOf_(x.r[3]), note: String(x.r[5] || ''), program: String(x.r[7] || ''), year: String(x.r[8] || ''), kind: String(x.r[9] || ''),
               elig: el.split(' (')[0], why: (el.match(/\((.*)\)$/) || [])[1] || '', status: String(x.r[4]).trim(), rrNote: String(x.r[11] || ''),
               awarded: String(x.r[15] || ''), campus: String(x.r[16] || ''), hometown: String(x.r[17] || ''),
               old: oldHit ? 'On the old masterlist: ' + oldHit.text : '', when: webFmt_(x.r[0], 'EEE, MMM d'), plan: m.plan };
    });
}

// Which Scholars row this person is: by email first, then by name at the same school (masterlist pasted without emails)
function joinMatch_(roster, name, school, email) {
  const norm = s => String(s || '').toLowerCase().replace(/[^a-zñ ]/g, ' ').replace(/\s+/g, ' ').trim();
  const words = s => norm(s).split(' ').filter(w => w.length > 1).sort().join(' ');
  email = emailOf_(email);
  let i = roster.findIndex(r => emailOf_(r[4]) === email);
  if (i >= 0) return { i, plan: String(roster[i][5]).trim() === 'Active' ? 'Already active' : 'Existing status: ' + (String(roster[i][5]).trim() || 'blank') + '. Review in Scholars before approval' };
  i = roster.findIndex(r => String(r[1]).trim() === String(school).trim() && !emailOf_(r[4]) && words(r[0]) && words(r[0]) === words(name));
  if (i >= 0) return { i, plan: 'Adds this email to ' + String(roster[i][0]).trim() + ' (already in Scholars)' };
  // already in Scholars with another main email: this one becomes an "other email" (sign-in only); START keeps emailing the main one
  i = roster.findIndex(r => String(r[1]).trim() === String(school).trim() && emailOf_(r[4]) && words(r[0]) && words(r[0]) === words(name));
  if (i >= 0) return { i, alt: true, plan: 'Adds this as another sign-in email for ' + String(roster[i][0]).trim() + ' (emails still go to ' + emailOf_(roster[i][4]) + ')' };
  return { i: -1, plan: 'Adds a new row in Scholars with their program, year level and scholarship (check the ONSSE masterlist first)' };
}

// A strict old-masterlist match can enrich a new automatic approval and mark that old row signed up.
// It is optional for approval. A recorded email must match before using old-row details.
function autoJoinMasterMatch_(ss, request, roster) {
  if (!autoApproveOn_() || String(request[2] || '').trim() === JOIN.other) return null;
  const norm = v => String(v || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim().replace(/\s+/g, ' ');
  const words = v => norm(v).split(' ').filter(Boolean).sort().join(' ');
  const school = String(request[2] || '').trim(), email = emailOf_(request[3]);
  if (joinMatch_(roster, request[1], school, email).i >= 0) return null; // never reactivate or attach an alias automatically
  const hits = oldMasterRows_(ss).filter(o => !String(o.r[10] || '').trim() &&
    email && String(o.r[1] || '').trim() === school && words(o.r[0]) === words(request[1]) &&
    norm(o.r[2]) === norm(request[7]) && String(o.r[3] || '').trim() === String(request[15] || '').trim() &&
    String(o.r[5] || '').trim() === 'Active');
  if (hits.length !== 1 || (emailOf_(hits[0].r[4]) && emailOf_(hits[0].r[4]) !== email)) return null;
  const hit = hits[0];
  return { at: hit.at, row: hit.r, text: [hit.r[0], hit.r[1], hit.r[2], hit.r[3]].filter(Boolean).join(' · ') };
}
// Auto-approve: anyone who registers with a complete request (name, email, one of the 13 programs, a school on the list).
// The old masterlist plays no part. Only a clash with an existing record waits for the RR: the same email on a
// non-Active scholar, or the same name at the same school under another email (that could be someone else's account).
function autoJoinEligible_(ss, request, roster) {
  const school = String(request[2] || '').trim(), email = emailOf_(request[3]);
  if (!autoApproveOn_() || !email || !String(request[1] || '').trim() || /^Not eligible/i.test(String(request[10] || '').trim()) ||   // older "Needs review" labels (masterlist era) no longer hold anyone back
     
      START_PROGRAMS.list.indexOf(String(request[7] || '').trim()) === -1 ||
      schoolRows_(ss).map(s => s.name).indexOf(school) === -1) return false;
  const m = joinMatch_(roster, request[1], school, email);
  if (m.i < 0) return true;
  const existing = roster[m.i];
  const sameName = v => displayName_(v).toLowerCase().replace(/[^a-z\u00f1 ]/g, ' ').replace(/\s+/g, ' ').trim();
  return !m.alt && String(existing[5] || '').trim() === 'Active' &&
    String(existing[1] || '').trim() === school && sameName(existing[0]) === sameName(request[1]);
}
// Approve every request still waiting (used when auto-approve is switched on). Returns how many were added.
function approveWaitingJoins_(ss) {
  const sh = joinTab_(ss, false);
  if (!sh || sh.getLastRow() < 2 || !autoApproveOn_()) return 0;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  let n = 0;
  rows.forEach((r, i) => {
    if (!/^(New|Asked for details)/.test(String(r[4]).trim()) || outOfTime_(30000)) return;
    try { const res = webJoinAction_(ss, 'join-add', { row: i + 2, auto: true }); if (res && res.ok) n++; }
    catch (e) { Logger.log('Auto-approve waiting row ' + (i + 2) + ': ' + e.message); }
  });
  if (n) MEMBER_CACHE_ = null;
  return n;
}

// RR on the website: Add or Decline a request
function webJoinAction_(ss, action, arg) {
  const sh = joinTab_(ss, false);
  const o = arg && typeof arg === 'object' ? arg : { row: arg };
  const row = Number(o.row);
  if (!sh || !(row >= 2) || row > sh.getLastRow()) return { ok: false, msg: 'That request is gone. Refresh the page.' };
  const lock = docLock_();
  if (!lock.tryLock(20000)) return { ok: false, msg: 'START is busy. Please try again.' };
  let r, email, school, name, addedRow = '';
  try {
    r = sh.getRange(row, 1, 1, JOIN_HEAD.length).getValues()[0];
    if (!JOIN.pending.test(String(r[4]).trim())) return { ok: false, msg: 'Already handled. Refresh the page.' };
    email = emailOf_(r[3]); name = String(r[1]).trim(); school = String(r[2]).trim();
    const stamp = Utilities.formatDate(new Date(), CONFIG.timezone, 'MMM d');
    if (action === 'join-decline') { sh.getRange(row, 5).setValue('Declined ' + stamp);
      try { markJoinChange_(ss, sh, row, String(o.actor || 'RR'), 'declined', name); } catch (e) { Logger.log('Join audit: ' + e.message); }
      try { refreshJoinViews_(ss); } catch (e) { Logger.log('Join views: ' + e.message); }
      return { ok: true, msg: name + ' was not added.' }; }
    if (action === 'join-clarify') {                                     // one short email to the applicant; the request stays open
      const q = String(o.note || '').replace(/\s+/g, ' ').trim().slice(0, 300);
      if (q.length < 5) return { ok: false, msg: 'Type what you need to know first.' };
      if (mailRoom_() < 1) return { ok: false, msg: 'No emails left today (daily limit). Try again tomorrow.' };
      const m = joinAskEmail_(name, q);
      try { MailApp.sendEmail({ to: email, subject: m.subject, htmlBody: m.html, body: m.text, name: m.from, replyTo: me_() }); }
      catch (e) { return { ok: false, msg: 'Not sent: ' + e.message }; }
      sh.getRange(row, 5).setValue('Asked for details ' + stamp);
      if (sh.getMaxColumns() >= 12) sh.getRange(row, 12).setValue(q);
      try { markJoinChange_(ss, sh, row, String(o.actor || 'RR'), 'clarification requested', 'Question sent'); } catch (e) { Logger.log('Join audit: ' + e.message); }
      try { refreshJoinViews_(ss); } catch (e) { Logger.log('Join views: ' + e.message); }
      return { ok: true, msg: 'Asked ' + name + '. Their reply comes to your email.' };
    }
    if (o.school) {                                                      // the RR matched a typed school to one on the list
      if (schoolRows_(ss).map(s => s.name).indexOf(String(o.school)) === -1) return { ok: false, msg: 'Choose a school from the list.' };
      school = String(o.school); sh.getRange(row, 3).setValue(school);
    }
    if (school === JOIN.other || schoolRows_(ss).map(s => s.name).indexOf(school) === -1)
      return { ok: false, msg: 'Match ' + name + '\'s school to one on the list first (or add the school in the Schools tab).' };
    const sc = ss.getSheetByName(TAB.scholars);
    const n = Math.max(sc.getLastRow() - 1, 0), roster = n ? sc.getRange(2, 1, n, 6).getValues() : [];
    if (o.auto && !autoJoinEligible_(ss, r, roster)) return { ok: false, msg: 'This request needs Regional Representative review.' };
    const autoHit = null;                                                 // automatic approval never relies on the old masterlist
    const m = joinMatch_(roster, name, school, email);
    if (m.i >= 0 && String(roster[m.i][5] || '').trim() !== 'Active')
      return { ok: false, msg: 'This scholar is not Active. Verify the record and change its status in Scholars before approving this request.' };
    // what the applicant gave: program (without "Not listed: "), year level (Scholars spelling), scholarship
    const program = String(r[7] || '').replace(/^Not listed:\s*/i, '').trim();
    const yl = String(r[8] || '').trim(), year = /^\d(st|nd|rd|th) year$/i.test(yl) ? yl.replace(/year$/i, 'Year') : '';
    const kind = String(r[9] || '').trim() === 'Not sure' ? '' : String(r[9] || '').trim();
    const from = 'Join request · ' + stamp;
    const oldHit = o.auto ? autoHit : oldMasterMatch_(oldMasterRows_(ss), name, school), ov = oldHit ? oldHit.row : [];
    const pick = (mine, k) => String(mine || '').trim() || String(ov[k] || '').trim();      // the request first, else the old masterlist
    const awarded = pick(r[15], 3), campus = pick(r[16], 7), hometown = pick(r[17], 8);
    const extras = c => {                                            // fills only EMPTY cells: your masterlist always wins
      if (sc.getMaxColumns() < 13) sc.insertColumnsAfter(sc.getMaxColumns(), 13 - sc.getMaxColumns());
      const cur = sc.getRange(c, 1, 1, 13).getValues()[0];
      [[3, program], [4, awarded], [7, year], [8, campus], [9, hometown], [SCHOLAR_KIND_COL, kind]].forEach(([col, v]) => { if (v && !String(cur[col - 1]).trim()) sc.getRange(c, col).setValue(v); });
    };
    let at = -1;
    if (m.alt) {                                                    // another sign-in email for someone already listed
      const cell = sc.getRange(m.i + 2, ALT_EMAIL_COL), have = altEmails_(cell.getValue());
      if (have.indexOf(email) === -1) cell.setValue(have.concat([email]).join(', '));
      if (String(roster[m.i][5]).trim() !== 'Active') sc.getRange(m.i + 2, 6).setValue('Active');
      email = emailOf_(roster[m.i][4]);                             // the "You're in" email goes to the main email
      at = m.i + 2; extras(at);
    } else if (m.i >= 0) {
      at = m.i + 2;
      if (!emailOf_(roster[m.i][4])) sc.getRange(at, 5).setValue(email);
      if (String(roster[m.i][5]).trim() !== 'Active') sc.getRange(at, 6).setValue('Active');
      extras(at);
    } else {                                                        // a new scholar: one full row, A:M
      at = n + 2;
      if (sc.getMaxColumns() < 13) sc.insertColumnsAfter(sc.getMaxColumns(), 13 - sc.getMaxColumns());
      if (sc.getMaxRows() < at) sc.insertRowsAfter(sc.getMaxRows(), at - sc.getMaxRows());
      sc.getRange(at, 1, 1, 13).setValues([[name, school, pick(program, 2), awarded, email, 'Active', year || pick('', 6), campus, hometown, 'Member', '', kind, from]]);
    }
    sh.getRange(row, 5).setValue((o.auto ? 'Added automatically ' : 'Added ') + stamp); addedRow = at;
    try { markScholarChange_(ss, sc, at, o.auto ? 'automatic approval' : String(o.actor || 'RR'), 'join approved', 'Join request row ' + row,
      m.i >= 0 ? String(roster[m.i][5] || '').trim() : '', 'Active'); } catch (e) { Logger.log('Scholar audit: ' + e.message); }
    try { markJoinChange_(ss, sh, row, o.auto ? 'automatic approval' : String(o.actor || 'RR'), 'approved', 'Scholar ID at row ' + at); }
    catch (e) { Logger.log('Join audit: ' + e.message); }
    if (oldHit) try { oldMasterTab_(ss).getRange(oldHit.at, OLD_MASTER_HEAD.length).setValue('Signed up · Scholars row ' + at + ' · ' + stamp); } catch (e) { /* note only */ }
    if (sh.getMaxColumns() >= 12) sh.getRange(row, 12).setValue('Scholars row ' + at);   // where they are now
    SpreadsheetApp.flush();
    try { refreshJoinViews_(ss); } catch (e) { Logger.log('Join views: ' + e.message); }
  } finally { lock.releaseLock(); }
  MEMBER_CACHE_ = null;
  try { calPeopleChanged_(ss); } catch (e) { Logger.log('Calendar sharing: ' + e.message); }
  if (o.auto) return { ok: true, approved: true, msg: name + ' was approved (Scholars row ' + addedRow + '). Password setup starts after they confirm their email on the website.' };
  // Approval is complete before any password code is sent.
  // A new applicant gets a code. An existing account gets an approval/sign-in email.
  let mailed = false, mailIssue = '';
  const hasPassword = !!(webAccount_(email) || {}).h;
  if (!hasPassword) {
    try { const sent = webMailCode_(email, 'new', { name: name });
      mailed = !!sent.ok;
      if (!mailed) mailIssue = sent.msg || 'Could not send verification code';
    } catch (e) { mailIssue = 'Could not prepare the verification email'; Logger.log('Join code: ' + e.message); }
  } else {
    try {
      if (mailRoom_() > 0) {
        const m = joinWelcomeEmail_(name, email);
        MailApp.sendEmail({ to: email, subject: m.subject, htmlBody: m.html, body: m.text, name: m.from });
        mailed = true;
      } else mailIssue = 'Email quota reached';
    } catch (e) { mailIssue = e.message; Logger.log('Join approval email: ' + e.message); }
  }
  return { ok: true, msg: name + ' was approved (Scholars row ' + addedRow + '). ' +
    (mailed ? (hasPassword ? 'Approval email sent.' : 'Password verification code sent.') :
      'Email not sent: ' + mailIssue + '. Ask the scholar to open START, enter their email, then select Set Up Password to request a code.') };

}

function joinAskEmail_(name, question) {
  let st = {}; try { st = getSettings_(); } catch (e) { st = {}; }
  const first = (typeof firstName_ === 'function' ? firstName_(name) : '') || 'there';
  const intro = ['Thank you for asking to join START. Before we can add you, we need a little more information:', question, 'Please reply to this email.'];
  const html = designedEmail_({ badge: 'Join request', badgeColor: THEME.blue, title: ['A quick', 'question'], greeting: `Hi ${first},`, intro,
    extraHtml: '', outro: [], link: '', buttonText: '', buttonNote: '', chips: [['', ''], ['', ''], ['', '']], signature: rrSignature_(st) });
  const text = [`Hi ${first},`, '', intro[0], '', question, '', intro[2], '', 'Respectfully,'].concat(rrSignature_(st)).join('\n');
  return { subject: 'START · About your join request', html, text, from: 'START' + (st.region ? ' ' + st.region : '') };
}

function joinWelcomeEmail_(name, email) {
  let st = {}; try { st = getSettings_(); } catch (e) { st = {}; }
  const first = (typeof firstName_ === 'function' ? firstName_(name) : '') || 'there', url = webHome_();
  const hasPassword = !!(webAccount_(email) || {}).h;
  const intro = [hasPassword
    ? 'Your START membership has been approved! Open START and sign in with ' + email + ' and the password you already created.'
    : 'Your START membership has been approved! Open START, enter ' + email + ', then select Set Up Password to receive a verification code. You can request another code if yours has expired.'];
  const html = designedEmail_({ badge: 'Welcome', badgeColor: THEME.blue, title: ['You\'re', 'in'], greeting: `Hi ${first},`, intro,
    extraHtml: '', outro: ['Each week you answer a short check-in there. See you inside!'],
    link: url, buttonText: url ? 'Open START' : '', buttonNote: '', chips: [['', ''], ['', ''], ['', '']], signature: rrSignature_(st) });
  const text = [`Hi ${first},`, '', intro[0], url ? '\n' + url : '', '', 'Respectfully,'].concat(rrSignature_(st)).join('\n');
  return { subject: 'START · You\'re in', html, text, from: 'START' + (st.region ? ' ' + st.region : '') };
}

// Once a day (daily jobs): one email to the RR when there are new requests
function joinDigest_(ss) {
  const sh = joinTab_(ss, false);
  if (!sh || sh.getLastRow() < 2) return 0;
  const fresh = sh.getRange(2, 1, sh.getLastRow() - 1, Math.min(Math.max(sh.getLastColumn(), 6), JOIN_HEAD.length)).getValues().filter(r => String(r[4]).trim() === 'New');
  if (!fresh.length) return 0;
  const me = me_(), day = Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd');
  if (!RR_EMAILS) return fresh.length;                                        // the "Asked to join" card on the RR's START page
  if (alreadySent_('joinDigest', day, me) || mailRoom_() < 1) return 0;
  const list = fresh.slice(0, 20).map(r => `${String(r[1])} · ${String(r[2])} · ${emailOf_(r[3])}`);
  const url = webHome_();
  const html = designedEmail_({ badge: 'Join requests', badgeColor: THEME.blue, title: ['Asked to', 'join'], greeting: '', signature: [],
    intro: [`${fresh.length} ${fresh.length === 1 ? 'person' : 'people'} asked to join START. Check the ONSSE masterlist, then Add or Decline on your START page.`],
    extraHtml: '<ul>' + list.map(x => '<li>' + esc_(x) + '</li>').join('') + '</ul>', outro: [],
    link: url || ss.getUrl(), buttonText: url ? 'Open START' : 'Open the tracker', buttonNote: '', chips: [['', ''], ['', ''], ['', '']] });
  sendMail_({ to: me, name: 'START Tracker', subject: `START · ${fresh.length} join request${fresh.length === 1 ? '' : 's'}`, htmlBody: html,
              body: list.join('\n') + (url ? '\n\n' + url : '') });
  markSent_('joinDigest', day, me);
  return fresh.length;
}

// Privacy: handled requests (Added / Declined) are removed after 90 days; the person is in Scholars, or was not added
function joinTidy_(ss) {
  const sh = joinTab_(ss, false);
  if (!sh || sh.getLastRow() < 2) return 0;
  const old = new Date(Date.now() - 90 * 86400000);
  let n = 0;
  for (let r = sh.getLastRow(); r >= 2; r--) {
    const v = sh.getRange(r, 1, 1, 5).getValues()[0];
    if (v[0] instanceof Date && v[0] < old && !JOIN.pending.test(String(v[4]).trim())) { sh.deleteRow(r); n++; }
    else if (v[0] instanceof Date && v[0] < old && String(v[4]).trim() === 'Not eligible') { sh.deleteRow(r); n++; }   // privacy: 90 days
  }
  return n;
}


/* ---------- a FRESH Scholars list: only scholars who signed up (with an email and full details) ----------
 * Done once by ① Set up / fix the tracker: Scholars rows WITHOUT an email (old ONSSE masterlist rows) are moved to the
 * "Old masterlist" tab. Nothing is deleted. Campus Leads stay (the Schools tab picks them from Scholars).
 * When someone from the old masterlist signs up and you approve them, their new row is made in Scholars and the old row
 * is marked "Signed up". The old row also fills any detail the sign-up left blank. */
var OLD_MASTER_TAB = 'Old masterlist';
var OLD_MASTER_HEAD = ['Name', 'School', 'Program', 'Year awarded (DOST)', 'Contact (email)', 'Status', 'Year level', 'Campus / Branch',
                       'Hometown (City/Municipality, Province)', 'Moved on', 'Signed up?'];
function oldMasterTab_(ss, create) {
  let sh = ss.getSheetByName(OLD_MASTER_TAB);
  if (sh || !create) return sh;
  sh = ss.insertSheet(OLD_MASTER_TAB);
  sh.getRange(1, 1, 1, OLD_MASTER_HEAD.length).setValues([OLD_MASTER_HEAD]);
  sh.setFrozenRows(1);
  try { sh.protect().setDescription(PROTECT_TAG + ' old masterlist').setWarningOnly(true); } catch (e) { /* fine */ }
  return sh;
}
function oldMasterRows_(ss) {
  const sh = oldMasterTab_(ss, false);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, OLD_MASTER_HEAD.length).getValues().map((r, i) => ({ r, at: i + 2 }));
}
// The old masterlist row for this name (any word order, initials ignored), preferably at the same school
function oldMasterMatch_(old, name, school) {
  const words = s => String(s || '').toLowerCase().replace(/[^a-zñ ]/g, ' ').split(/\s+/).filter(w => w.length > 1).sort().join(' ');
  const w = words(name); if (!w || !old.length) return null;
  const hits = old.filter(o => words(o.r[0]) === w && !String(o.r[10]).trim());
  const h = hits.find(o => String(o.r[1]).trim() === String(school || '').trim()) || hits[0];
  return h ? { at: h.at, row: h.r, text: [String(h.r[0]).trim(), String(h.r[1]).trim(), String(h.r[2]).trim(), String(h.r[3]).trim()].filter(Boolean).join(' · ') } : null;
}
// Once: Scholars rows with no email (main or other) and not a Campus Lead -> Old masterlist. Returns how many moved.
function freshScholarList_(ss, force) {
  const props = docProps_();
  if (!force && props.getProperty('FRESH_LIST_DONE')) return 0;
  const sc = ss.getSheetByName(TAB.scholars);
  if (!sc || sc.getLastRow() < 2) { props.setProperty('FRESH_LIST_DONE', String(Date.now())); return 0; }
  const lock = docLock_();
  if (!lock.tryLock(30000)) return 0;
  try {
    const n = sc.getLastRow() - 1, w = Math.max(Math.min(sc.getLastColumn(), 13), 11);
    const rows = sc.getRange(2, 1, n, w).getValues();
    const leads = {}; schoolRows_(ss).forEach(s => { if (s.focal) leads[s.focal] = true; });
    const move = [];
    rows.forEach((r, i) => {
      const nm = String(r[0]).trim();
      if (!nm || emailOf_(r[4]) || altEmails_(r[10]).length || leads[nm]) return;
      move.push(i);
    });
    if (!move.length) { props.setProperty('FRESH_LIST_DONE', String(Date.now())); return 0; }
    const old = oldMasterTab_(ss, true), stamp = Utilities.formatDate(new Date(), CONFIG.timezone, 'MMM d, yyyy');
    const out = move.map(i => rows[i].slice(0, 9).concat([stamp, '']));
    const at = Math.max(old.getLastRow(), 1) + 1;
    if (old.getMaxRows() < at + out.length) old.insertRowsAfter(old.getMaxRows(), at + out.length - old.getMaxRows());
    old.getRange(at, 1, out.length, OLD_MASTER_HEAD.length).setValues(out);
    SpreadsheetApp.flush();
    if (old.getLastRow() < at + out.length - 1) throw new Error('copy check failed: nothing was removed from Scholars');
    // remove from Scholars from the bottom up, in runs (fast), only after the copy is safe
    const runs = []; move.slice().reverse().forEach(i => { const last = runs[runs.length - 1];
      if (last && last.start === i + 1) { last.start = i; last.n++; } else runs.push({ start: i, n: 1 }); });
    runs.forEach(x => sc.deleteRows(x.start + 2, x.n));
    props.setProperty('FRESH_LIST_DONE', String(Date.now()));
    MEMBER_CACHE_ = null;
    try { ss.toast(out.length + ' scholars without an email were moved to "' + OLD_MASTER_TAB + '". Scholars now lists only scholars with an account email.', 'START Tracker', 12); } catch (e) { /* no screen */ }
    return out.length;
  } finally { lock.releaseLock(); }
}


/* ---------- 🧹 FRESH START: recreate the database tabs from the current schema ----------
 * Settings and the school list are the seed for new sign-ups. Every other tab is
 * removed after a complete private backup, including retired and unknown tabs. */
var FRESH_KEEP = () => [TAB.settings, TAB.schools];
var FRESH_PROPS = /^(WEB_A_|WEB_S_|WEB_SUSP_|WEB_C_|WEB_T_|WEB_D_|WEB_J_|WEB_JD_|SENT_LEDGER$|MAIL_QUEUE$|ANSWERS_SEEN$|ANSWERS_BACKUP$|LAST_CHECK$|LAST_REPORT$|PORTALS$|PORTALS_ON$|FOCAL_SYNC$|MINI_SECRET$)/;

function freshStart() {
  sheetUser_();                                   // sheet only: never from the website
  requireSetup_();
  requireLook_();
  const ui = SpreadsheetApp.getUi(), ss = SpreadsheetApp.getActive();
  if (!ensureOwner_('start fresh')) return;
  const drop = ss.getSheets().map(s => s.getName()).filter(n => FRESH_KEEP().indexOf(n) === -1);
  const msg = 'This empties the tracker so you can start again from the website.\n\n' +
    '• A full backup copy is saved first (private, next to this file).\n' +
    (drop.length ? '• Delete and rebuild ' + drop.length + ' tabs, including old tabs and all history.\n' : '') +
    '• Keep: Settings and the Schools list; Campus Lead assignments are cleared.\n' +
    '• Old masterlist is deleted; automatic approval needs a new masterlist.\n' +
    '• Website accounts are deleted: everyone, you too, creates their account again.\n\n' +
    'Type FRESH START to continue.';
  const r = ui.prompt('🧹 Fresh start', msg, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK || r.getResponseText().trim().toUpperCase() !== 'FRESH START') { notify_('Nothing was changed.'); return; }
  const out = freshStart_(ss);
  notify_((out.pending ? '⏳ Fresh start tabs recreated. Click ① Set up / fix the tracker again to finish setup.' : '🧹 Fresh start done.') +
          '\n\nBackup: ' + out.backup + '\n\nDeleted ' + out.dropped + ' old/data tab(s), recreated clean tabs, removed ' + out.props + ' website records.\n\n' +
          'Next: open the website and create your account again. New scholars use "Ask to join START"; approve them on your Region page.');
}

function freshStart_(ss) {
  const lock = docLock_();
  if (!lock.tryLock(30000)) throw new Error('START is busy (someone is saving a check-in). Try again in a minute.');
  let dropped = 0, props = 0, copy;
  try {
    if (!FRESH_KEEP().every(n => ss.getSheetByName(n)))
      throw new Error('Settings and Schools must exist before a fresh start. Run ① Set up / fix the tracker first.');
    // The backup is made while writes are locked, so no check-in can fall between the copy and deletion.
    const stamp = Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd HHmm');
    const file = DriveApp.getFileById(ss.getId());
    copy = file.makeCopy(ss.getName() + ' (backup before fresh start ' + stamp + ')');
    try { const p = file.getParents(); if (p.hasNext()) copy.moveTo(p.next()); } catch (e) { /* stays in My Drive */ }
    if (WEB_CHECKIN) retireForms_(ss); // prevent linked old Forms from recreating response tabs
    ss.getSheets().map(s => s.getName()).forEach(n => {
      if (FRESH_KEEP().indexOf(n) !== -1) return;
      const sh = ss.getSheetByName(n);
      if (sh) { ss.deleteSheet(sh); dropped++; }
    });
    const sch = ss.getSheetByName(TAB.schools);
    sch.getRange(2, 4, CONFIG.maxSchools, 1).clearContent();
    // Accounts, codes, sessions, sent-email ledger and waiting emails belong to the old data.
    const dp = docProps_(), all = dp.getProperties();
    Object.keys(all).forEach(k => { if (FRESH_PROPS.test(k)) { dp.deleteProperty(k); props++; } });
    setAutoApprove_(false); // a fresh database must not inherit automatic approval from the previous roster
    dp.setProperty('FRESH_LIST_DONE', String(Date.now()));
    ['TABLE_LOOK', 'SETUP_DONE', SETUP_KEY, DB_TIDY_KEY].forEach(k => dp.deleteProperty(k));
    removeTriggers_(['sendWaitingEmails', 'refreshPortalsAuto']);
    try { CacheService.getScriptCache().removeAll(['SLOGO_VER', 'WEB_LEADERBOARD']); } catch (e) { /* fine */ }
    MEMBER_CACHE_ = null;
    getOrCreate_(ss, TAB.events); // setup only formats Events when its tab exists
    joinTab_(ss, true);
    refreshJoinViews_(ss);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  // Setup resumes safely on the next menu click if Google reaches its execution limit.
  try { setupFull(); }
  catch (e) {
    throw new Error('The old tabs were deleted and the backup is safe at ' + copy.getUrl() +
      '. Run ① Set up / fix the tracker to finish rebuilding. Setup error: ' + e.message);
  }
  return { backup: copy.getUrl(), dropped, props, pending: !!docProps_().getProperty(SETUP_KEY) };
}

/* ---------- Scholars tab = the sign-up form (same fields, same choices) ----------
 * Dropdowns match the website exactly, so a row typed by hand and a row made by an approved sign-up look the same:
 *   C Program (13 START-DOST eligible programs) · G Year level · I Hometown (this region's
 *   "City, Province") · L Scholarship (DOST). Typing something else is allowed but shows a warning triangle. */
function alignScholarsTab_(ss) {
  const sh = ss.getSheetByName(TAB.scholars);
  if (!sh) return;
  const rows = Math.max(sh.getMaxRows() - 1, 1);
  const dv = (col, list, help) => sh.getRange(2, col, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(list.slice(0, 500), true).setAllowInvalid(true).setHelpText(help).build());
  try { dv(3, START_PROGRAMS.list, 'Only the 13 START-DOST eligible courses listed in the sign-up form.'); } catch (e) { Logger.log('Program list: ' + e.message); }
  try { dv(6, ['Active', 'On leave', 'Graduated', 'Dropped', 'Suspended'], 'Scholar account status.'); } catch (e) { /* fine */ }
  try { dv(7, ['1st Year', '2nd Year', '3rd Year', '4th Year', '5th Year'], 'Year level.'); } catch (e) { /* fine */ }
  try {
    const places = joinPlaces_(), list = [];
    Object.keys(places).sort().forEach(p => places[p].forEach(c => list.push(c + ', ' + p)));
    if (list.length && list.length <= 500) dv(9, list, 'City / Municipality, Province (same as the sign-up form).');
  } catch (e) { Logger.log('Hometown list: ' + e.message); }
  try { if (sh.getMaxColumns() >= SCHOLAR_KIND_COL) dv(SCHOLAR_KIND_COL, JOIN.kinds, 'DOST scholarship (same as the sign-up form).'); } catch (e) { /* fine */ }
  try {                                                                   // formats
    sh.getRange(2, 4, rows, 1).setNumberFormat('0');                      // year awarded: 2024, never 2,024
    if (!DB_MODE) sh.getRange(2, 1, rows, 13).setVerticalAlignment('middle');
  } catch (e) { /* fine */ }
}


/* ---------- auto-approve (Settings and the RR's Scholars page) ----------
 * ON: valid sign-ups from listed schools are added automatically, even without an old
 * masterlist match. Record conflicts wait for review. OFF: every request waits. */
// Where this email's join request stands: 'waiting' | 'sorry' (not eligible) | 'declined' | 'added' | '' (none). Newest row wins.
function joinStateOf_(ss, email) {
  email = emailOf_(email);
  const sh = joinTab_(ss, false);
  if (!email || !sh || sh.getLastRow() < 2) return '';
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    if (emailOf_(rows[i][3]) !== email) continue;
    const s = String(rows[i][4]).trim();
    if (/^Not eligible/.test(s)) return 'sorry';
    if (/^(New|Asked for details)/.test(s)) return 'waiting';
    if (/^Declined/.test(s)) return 'declined';
    if (/^Added/.test(s)) return 'added';
    return '';
  }
  return '';
}

// The Settings checkbox is the source of truth, so an RR can use either the Sheet or the website.
// Older trackers without the new row retain their website switch until setup adds the row.
function autoApproveCell_() {
  const sh = webSheet_().getSheetByName(TAB.settings);
  const i = SETTINGS.findIndex(s => s[0] === 'autoApprove');
  return sh && i >= 0 && sh.getRange(i + 2, 1).getValue() === SETTINGS[i][1] ? sh.getRange(i + 2, 2) : null;
}
function autoApproveOn_() {
  try { const cell = autoApproveCell_(); if (cell) return cell.getValue() === true; } catch (e) { /* older tracker: use the saved switch */ }
  try { return docProps_().getProperty('AUTO_APPROVE') === '1'; } catch (e) { return false; }
}
function setAutoApprove_(on) {
  const cell = autoApproveCell_();
  if (cell) cell.setValue(!!on);
  try { docProps_().setProperty('AUTO_APPROVE', on ? '1' : '0'); }
  catch (e) { if (!cell) throw e; Logger.log('Auto-approve backup setting: ' + e.message); }
  return !!on;
}