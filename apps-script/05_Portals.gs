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


// Guard for sheet-only administrative actions. These actions must not be
// callable through the public website (which runs as a web app).
function sheetUser_() {
  try {
    const ui = SpreadsheetApp.getUi();
    if (!ui) throw new Error('No spreadsheet interface');
    return true;
  } catch (err) {
    throw new Error('Open the START Tracker spreadsheet and use its START Tracker menu for this action.');
  }
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
  if (!s.thisWeek) todo.push([`Answer your check-in by ${f(s.due, 'EEEE, MMM d')} (yellow button above).`, CHIP.yellow]);
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
  const ck = ss.getSheetByName(TAB.checkins);
  if (ck && ck.getLastRow() > 1) { const c = ckCols_(ck); if (c.school >= 0) replaceExact_(col(TAB.checkins, c.school + 1), from, to); }
  const store = portalStore_();
  if (store.map[from] && !store.map[to]) { store.map[to] = store.map[from]; delete store.map[from]; savePortalStore_(store); }
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
            lateDays: 2,          // a missed check-in can still be answered this many days after its due day (marked late)
            cacheSecs: 60 };      // a page is re-read from the sheet at most once a minute (faster on busy nights)

// The page itself
function doGet(e) {
  if (e && e.parameter && e.parameter.ping) {                       // 🩺 Check my tracker: which code version is live?
    return ContentService.createTextOutput(JSON.stringify({ v: CODE_VERSION })).setMimeType(ContentService.MimeType.JSON);
  }
  let st = {};
  try { webSheet_(); st = getSettings_(); } catch (err) { Logger.log('Website: ' + err.message); }
  // Google's own answer to "what is my address?" can be an old, broken one: only a fallback until the RR saves the real one (🌐)
  try { const u = ScriptApp.getService().getUrl(); if (u && !docProps_().getProperty('WEB_URL_SET')) docProps_().setProperty('WEB_URL', u); } catch (err) { /* not deployed */ }
  const boot = {
    region: String(st.region || ''), regionName: String(st.regionName || ''), pics: webPics_(),
    mark: typeof START_MARK === 'undefined' ? '' : START_MARK,
    place: (() => { try { return placeOfWeek_(); } catch (err) { return null; } })(),
    appOrigin: ''
  };
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

// Who is this email in THIS tracker? (checked again on every request)
function webRoles_(ss, email) {
  email = emailOf_(email);
  const r = { email, name: '', member: null, leads: [], admin: false };
  if (!email) return r;
  r.suspended = docProps_().getProperty('WEB_SUSP_' + webHash_(email)) === '1';
  let st = {}; try { st = getSettings_(); } catch (e) { st = {}; }
  const admins = [me_(), automationOwner_()].concat(String(st.recipients || '').split(/[,;\s]+/).map(emailOf_)).filter(Boolean);
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

// Step 1: the email. Tells the page what comes next: the password box, or a code (new account).
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
    if (purpose !== 'reset' && acct && acct.h) return { ok: true, mode: 'password' };
    // the code fits the account: no password yet = create it; a password = reset it (whichever button they pressed)
    const kind = acct && acct.h ? 'reset' : 'new';
    const sent = webMailCode_(email, kind, who);                   // codes go ONLY to the main email
    if (!sent.ok) return sent;
    return { ok: true, mode: 'code', purpose: kind, recent: !!sent.recent, to: email === typed ? '' : webMask_(email) };
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
  if (j === 'waiting') return { ok: true, mode: 'joined', msg: 'Your request is pending RR review. Once approved, we will email you a verification code so you can create your password and sign in.' };
  if (j === 'sorry') return { ok: true, mode: 'joined', sorry: true, msg: JOIN_SORRY };
  return { ok: true, mode: 'join', note: j === 'declined'
    ? 'Your last request wasn\'t approved. Check your details and send again, or contact your Campus Lead or Regional Representative.'
    : 'This email isn\'t on the START list yet. New scholar? Fill this in once, and you\'ll be added to START.' };
}

function webMailCode_(email, kind, who) {
  const props = docProps_(), key = 'WEB_C_' + webHash_(email), now = Date.now();
  let rec = {}; try { rec = JSON.parse(props.getProperty(key) || '{}'); } catch (e) { rec = {}; }
  rec.sent = (rec.sent || []).filter(t => now - t < 3600000);
  if (rec.sent.length && now - rec.sent[rec.sent.length - 1] < 60000 && rec.k === kind && rec.h) return { ok: true, recent: true };   // one went out less than a minute ago
  if (rec.sent.length >= WEB.perHour) return { ok: false, msg: 'You asked for ' + WEB.perHour + ' codes in the last hour. Please use the newest code, or try again in an hour.' };
  // the daily email limit is shared with the reminders: sign-up codes stop early so the reminders still go out
  const dayKey = 'WEB_D_' + Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyyMMdd');
  const today = Number(props.getProperty(dayKey) || 0);
  if (today >= WEB.codesPerDay || MailApp.getRemainingDailyQuota() <= WEB.mailReserve)
    return { ok: false, msg: 'START has reached today\'s email limit, so no code could be sent. Please try again tomorrow, or ask your Regional Representative.' };
  const code = String(100000 + parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 12), 16) % 900000);
  const m = webEmail_(kind, who, code);
  // send FIRST, then remember it: a failed email never blocks the next try or counts against the limits
  try { MailApp.sendEmail({ to: email, subject: m.subject, htmlBody: m.html, body: m.text, name: m.from }); }
  catch (e) {
    Logger.log('Sign-in code email failed: ' + (e && e.message));
    return { ok: false, msg: 'We couldn\'t send the email just now. Please try again in a few minutes.' };
  }
  props.setProperty(dayKey, String(today + 1));
  Object.assign(rec, { h: webHash_(email + '|' + code), k: kind, exp: now + WEB.codeMins * 60000, tries: 0 });
  rec.sent.push(now);
  props.setProperty(key, JSON.stringify(rec));
  return { ok: true };
}

// Step 2 (new account / reset): check the code. Gives a 15-minute pass to choose a password.
function webCheckCode(email, code) {
  return webSafe_(() => {
    const ss = webSheet_();
    email = primaryEmail_(ss, email);
    const props = docProps_(), key = 'WEB_C_' + webHash_(email);
    let rec = {}; try { rec = JSON.parse(props.getProperty(key) || '{}'); } catch (e) { rec = {}; }
    if (!email || !rec.h || Date.now() > rec.exp) return { ok: false, msg: 'This code has expired. Send a new one.' };
    if ((rec.tries || 0) >= WEB.maxTries) return { ok: false, msg: 'Too many wrong tries. Send a new code.' };
    if (webHash_(email + '|' + String(code || '').replace(/\D/g, '')) !== rec.h) {
      rec.tries = (rec.tries || 0) + 1; props.setProperty(key, JSON.stringify(rec));
      const left = WEB.maxTries - rec.tries;
      return { ok: false, msg: left > 0 ? `Wrong code. ${left} ${left === 1 ? 'try' : 'tries'} left.` : 'Too many wrong tries. Send a new code.' };
    }
    rec.h = ''; props.setProperty(key, JSON.stringify(rec));                    // a code works once
    if (!webAllowed_(webRoles_(ss, email)))
      return { ok: false, msg: 'Your membership must be approved before password setup. Contact your Regional Representative.' };
    const ticket = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    props.setProperty('WEB_T_' + webHash_(ticket), JSON.stringify({ e: email, k: rec.k || 'new', exp: Date.now() + 15 * 60000 }));
    return { ok: true, ticket, purpose: rec.k || 'new' };
  });
}

// Step 3 (new account / reset): save the password, sign in here, sign out everywhere else
function webSetPassword(ticket, pw, keep) {
  return webSafe_(() => {
    const ss = webSheet_(), props = docProps_(), tkey = 'WEB_T_' + webHash_(ticket || '');
    let t = null; try { t = JSON.parse(props.getProperty(tkey) || 'null'); } catch (e) { t = null; }
    if (!t || t.exp < Date.now()) return { ok: false, restart: true, msg: 'That took too long. Start again.' };
    const bad = webPwProblem_(pw, t.e);
    if (bad) return { ok: false, msg: bad };
    const who = webRoles_(ss, t.e);
    if (!webAllowed_(who)) return { ok: false, restart: true, msg: 'Your START membership is not approved yet.' };
    const old = webAccount_(t.e) || {};
    const salt = Utilities.getUuid().replace(/-/g, '');
    const a = { s: salt, h: webPwHash_(pw, salt), v: (old.v || 0) + 1, at: Date.now(), fails: 0, lock: 0 };
    webSaveAccount_(t.e, a);
    try { markVerifiedEmail_(ss, t.e); } catch (e) { Logger.log('Verified email: ' + e.message); }
    props.deleteProperty(tkey);
    webEndSessions_(t.e);
    try { const m = webEmail_(old.h ? 'changed' : 'created', who); MailApp.sendEmail({ to: t.e, subject: m.subject, htmlBody: m.html, body: m.text, name: m.from }); }
    catch (e) { Logger.log('Password email: ' + e.message); }
    return { ok: true, token: webNewSession_(t.e, a.v, keep === false) };
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
    if (s) { webEndSessions_(s.e); webForget_(s.e); }
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
function webForget_(email) { try { CacheService.getScriptCache().remove('WL_' + webHash_(emailOf_(email))); } catch (e) { /* no cache */ } }

// Every morning: accounts of people no longer on the list (graduated, dropped, changed email) are removed
function webCleanAccounts_(ss) {
  ss = ss || webSheet_();
  const st = getSettings_(), keep = {};
  const add = e => { e = emailOf_(e); if (e) keep['WEB_A_' + webHash_(e)] = true; };
  [me_(), automationOwner_()].concat(String(st.recipients || '').split(/[,;\s]+/)).forEach(add);
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
               intro: ['Enter this code to create your START password. If you applied to join, your request still needs approval from the Regional Representative before you can sign in.'], extra: code ? big(code) : '',
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
  const button = !code && url;                                   // code emails: the code is all they need
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
  return webSafe_(() => {
    const s = webSession_(token);
    if (!s) return { ok: false, signedOut: true };
    const ss = webSheet_(), who = webRoles_(ss, s.e);
    if (!webAllowed_(who)) {
      docProps_().deleteProperty('WEB_S_' + webHash_(token));
      return { ok: false, signedOut: true, msg: 'This email is no longer on your region\'s START list. Please contact your Regional Representative.' };
    }
    const cache = (() => { try { return CacheService.getScriptCache(); } catch (e) { return null; } })(), ck = 'WL_' + webHash_(who.email);
    if (!fresh && cache) { try { const hit = cache.get(ck); if (hit) return JSON.parse(hit); } catch (e) { /* read it fresh */ } }
    const out = webBuild_(ss, who);
    const json = JSON.stringify(out);
    if (cache && json.length < 90000) { try { cache.put(ck, json, WEB.cacheSecs); } catch (e) { /* too big: no cache */ } }
    return JSON.parse(json);
  });
}

// One person's whole page (used for them, and by the RR's "View as")
function webBuild_(ss, who) {
    const st = getSettings_();
    const out = { ok: true, me: { email: who.email, name: displayName_(who.name || who.email),
                                 first: (typeof firstName_ === 'function' ? firstName_(who.name) : '') || 'there' },
                  region: st.region || '', regionName: st.regionName || '', rr: String(st.rrName || '').trim(), rrEmail: me_(),
                  note: String(st.portalNote || '').trim(), updated: webFmt_(new Date(), 'MMM d, h:mm a'), cals: webCalendars_(who) };
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
    return { ok: true, schools: schoolRows_(ss).filter(x => !x.inactive).map(x => ({ name: x.name, lead: x.focal ? displayName_(x.focal) : '' })),
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
    if (role === 'lead') {
      const sc = schoolRows_(ss).find(x => x.name === key);
      if (!sc) return { ok: false, msg: 'That school is not in the Schools tab.' };
      const p = sc.email ? scholarByEmail_(ss)[sc.email] : null;
      who = { email: sc.email || '', name: sc.focal || ('Campus Lead of ' + sc.name), member: p && p.status === 'Active' ? p : null, leads: [sc.name], admin: false };
      label = (sc.focal ? displayName_(sc.focal) : 'No Campus Lead yet') + ' · Campus Lead, ' + sc.name;
    } else if (role === 'member') {
      who = webRoles_(ss, key);
      if (!who.member) return { ok: false, msg: 'That scholar is not Active in Scholars.' };
      who.admin = false;
      label = displayName_(who.name) + ' · ' + (who.leads.length ? 'Campus Lead and member' : 'Member') + ', ' + who.member.school;
    } else return { ok: false, msg: 'Choose a role.' };
    const out = webBuild_(ss, who);
    out.viewAs = { role, key, label };
    return JSON.parse(JSON.stringify(out));
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
           due: d ? webFmt_(d.week.due, 'EEEE, MMM d') : '', near: d ? memberDeadlineNear_(d.week) : false,
           formUrl: url, open: !!d, checkin: (spec = webCheckinSpec_(ss, who.email)), game: (() => { try { return webGame_(ss, webMyAnswers_(ss, who.email), spec); } catch (e) { Logger.log('Game: ' + e.message); return null; } })(), history: history.slice(0, 5).map(h => ({ date: h.date, vibe: h.vibe, help: h.help })),
           events: webEvents_(ss, school, false) };
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
             due: webFmt_(w.due, 'EEEE, MMM d'), q, consent: CONSENT_AGREE, privacy: privacyNote_() };
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
      if (again > 0) return { ok: false, msg: 'Your check-in is already submitted. Responses cannot be edited.' };
      row = Math.max(sh.getLastRow(), 1) + 1;
      sh.getRange(row, 1, 1, v.length).setValues([v]);
      SpreadsheetApp.flush();
      MEMBER_CACHE_ = null;
    } finally { lock.releaseLock(); }          // the answer is saved: the next scholar can save theirs right away
    // follow-ups (help alerts, Campus Lead status, ranks) run AFTER the save, outside the lock, so a Friday rush does not queue up
    try { onMemberCheckin_({ range: sh.getRange(row, 1) }); } catch (e) { Logger.log('Website check-in follow-up: ' + e.message); }
    try { markAnswersSeen_(ss); } catch (e) { /* the hourly check catches up */ }
    MEMBER_CACHE_ = null; webForget_(who.email);
    return { ok: true, member: JSON.parse(JSON.stringify(webMember_(ss, who))) };
  });
}

// A START Campus Lead: the same things as their portal sheet, for each school they lead
function webLead_(ss, schools, st) {
  const data = focalData_(ss);
  const link = (() => { try { return checkinUrl_(); } catch (e) { return ''; } })();
  const rrName = String(st.rrName || '').trim() || 'your Regional Representative';
  return schools.map(name => {
    const s = data.find(x => x.name === name);
    if (!s) return null;
    const P = portalMembers_(s, st);
    const todo = [];
    if (!s.thisWeek) todo.push({ tone: 'yellow', text: `Answer your check-in by ${webFmt_(s.due, 'EEEE, MMM d')}.` });
    P.help.forEach(m => todo.push({ tone: 'red', text: `${displayName_(m.name)} asked for help: message them today.` }));
    if (P.answers && P.toMessage.length) todo.push({ tone: 'yellow', copy: P.message,
      text: `Message the ${P.toMessage.length} member${P.toMessage.length === 1 ? ' who has' : 's who have'} not answered yet. You can copy this:` });
    if (P.noEmail.length) todo.push({ tone: 'grey',
      text: `${P.noEmail.length} member${P.noEmail.length === 1 ? ' has' : 's have'} no email on file: ask them to email ${rrName} (${me_()}) the Google email they use.` });
    if (!todo.length) todo.push({ tone: 'green', text: 'Nothing to do: you and your members are all in. Salamat!' });
    return {
      school: s.name, done: !!s.thisWeek, due: webFmt_(s.due, 'EEEE, MMM d'), streak: s.streak, term: s.term, active: s.active,
      answers: P.answers, answered: P.answered, members: P.members.length, link: link || s.link || '', todo,
      cols: P.cols.map(c => c[0]), rows: P.rows.map(x => ({ cells: x.cells.map(webCell_), tone: webTone_(x.chip), me: !!x.me })),
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
  return sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues()
    .map((r, i) => ({ row: i + 2, name: String(r[0] || '').trim(), school: String(r[1] || '').trim(),
      program: String(r[2] || '').trim(), awarded: String(r[3] || '').trim(), contact: String(r[4] || '').trim(),
      email: emailOf_(r[4]), status: String(r[5] || '').trim(), yearLevel: String(r[6] || '').trim(),
      campus: String(r[7] || '').trim(), hometown: String(r[8] || '').trim(), role: String(r[9] || '').trim(),
      otherEmails: String(r[10] || '').trim(), scholarship: String(r[11] || '').trim(), verified: !!r[17] }))
    .map(p => {
      const member = md && md.bySchool[p.school] && md.bySchool[p.school].members.find(m => m.email === p.email);
      p.role = leads.some(s => s.focal === p.name && s.email === p.email) ? ROLES[1] : ROLES[0];
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
    const schools = schoolRows_(ss), target = schools.find(s => s.name === school && (!s.inactive || school === String(old[1] || '').trim()));
    if (!target) return { ok: false, msg: 'Choose a school on the active Schools list.' };
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
    due: md ? webFmt_(md.week.due, 'EEEE, MMM d') : webFmt_(reminderWeek_(new Date()).deadline, 'EEEE, MMM d'),
    totals: { schools: schools.length, leads: schools.filter(s => s.lead).length, leadsDone: schools.filter(s => s.lead && s.leadDone).length,
              members: sum('members'), answered: sum('answered'), help: sum('help') },
    memberCheckins: !!md, schools, scholars: webScholarRoster_(ss, md), help, events: webEvents_(ss, undefined, true), tracker: ss.getUrl(), form,
    dataQuality: (() => { try { return dataFlowQuality_(ss); } catch (e) { Logger.log('Data quality: ' + e.message); return { count: 0, issues: [] }; } })(),
    scholarChoices: { schools: schools.map(s => s.name), programs: START_PROGRAMS.list, scholarships: JOIN.kinds },
    joins: (() => { try { return webJoins_(ss).filter(j => !/^Added|^Declined/.test(j.status)); } catch (e) { Logger.log('Join requests: ' + e.message); return []; } })(),
    joinHistory: (() => { try { return webJoins_(ss).filter(j => /^Added|^Declined/.test(j.status)); } catch (e) { Logger.log('Join history: ' + e.message); return []; } })(),
    autoApprove: autoApproveOn_(),
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
var GAME = { rankMin: 3, xpCheckin: 10, xpOnTime: 5, xpEvent: 5, journeyWeeks: 12 };
var LEVELS_DAVAO = ['Bagong Agila', 'Batang Apo', 'Lakad Samal', 'Agila sa Eden', 'Bantay Dabaw', 'Agila ng Apo'];
var LEVELS_ANY = ['Rookie', 'Explorer', 'Achiever', 'Trailblazer', 'Leader', 'Legend'];
var LEVEL_XP = [0, 50, 120, 220, 350, 520];

// mine = [{ t: Date, ev: events joined, late: true/false }] for one scholar
function webGame_(ss, mine, spec) {
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
  let best = 0, run = 0, prev = '', events = 0, onTime = 0, xp = 0;
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
  return { xp, level, streak, best, saver: { used: !!saved, available: !!termKey && !saved }, events, checkins: keys.length,
           badges, journey: journey.reverse(), cards: Object.keys(cards).map(Number), nCards, allCards, goal };
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

// Schools ranked: this week's members answered (%), then the school's streak
function webLeaderboard_(ss, mySchools) {
  const data = focalData_(ss).filter(s => !s.inactive);
  let md = null; try { if (checkinOn_()) md = memberData_(ss); } catch (e) { md = null; }
  const all = data.map(s => {
    const m = (md && md.bySchool[s.name]) || { members: [], answered: 0 };
    const n = md ? m.members.length : 0;
    return { name: s.name, rate: n ? Math.round(100 * m.answered / n) : 0, answered: m.answered || 0, members: n, streak: s.streak || 0,
             me: (mySchools || []).indexOf(s.name) !== -1 };
  });
  // fair: a school needs at least GAME.rankMin members to be ranked (1 of 1 is not "100%"); ties go to more answers, then the streak
  const ranked = all.filter(s => s.members >= GAME.rankMin)
    .sort((a, b) => (b.rate - a.rate) || (b.answered - a.answered) || (b.streak - a.streak) || a.name.localeCompare(b.name))
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

// Postcards page: every place, with the picture only for the ones collected
function webPostcards(token, asEmail) {
  return webSafe_(() => {
    const s = webSession_(token);
    if (!s) return { ok: false, signedOut: true };
    const ss = webSheet_();
    let who = webRoles_(ss, s.e);
    if (asEmail && who.admin) { who = webRoles_(ss, asEmail); s.e = who.email; }        // RR "View as" a member
    if (!who.member || !placesOn_()) return { ok: true, cards: [] };
    const got = {}; webGame_(ss, webMyAnswers_(ss, s.e), null).cards.forEach(i => { got[i] = true; });
    return { ok: true, cards: PLACE_PICS.map((p, i) => ({ name: p.name, got: !!got[i], src: got[i] ? p.src : '' })) };
  });
}

// RR on the website: send this week's reminder or follow-up now, or reset someone's password
function webAdminAction(token, action, arg) {
  return webSafe_(() => {
    const s = webSession_(token);
    if (!s) return { ok: false, signedOut: true };
    const ss = webSheet_(), who = webRoles_(ss, s.e);
    if (!who.admin) return { ok: false, msg: 'Only the Regional Representative can do this.' };
    if (action === 'scholar-status') return webScholarStatus_(ss, s.e, arg);
    if (action === 'scholar-update') return webScholarUpdate_(ss, s.e, arg);
    if (action === 'resetpw') return { ok: true, msg: resetWebPassword_(arg) };
    if (action === 'auto-approve') { webForget_(s.e); const on = setAutoApprove_(arg === 'on');
      return { ok: true, msg: on ? 'Auto-approve is ON: only requests matching one old masterlist record by email, name, school, program and award year are added at once. Others wait for review.' : 'Auto-approve is OFF: every sign-up waits for you.' }; }
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
  const totals = refreshJoinViews_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getActiveSpreadsheet().toast('Join history updated: ' + JSON.stringify(totals), 'START Tracker', 7);
}

// What a school is called on the sign-up page. The University of Mindanao's main campus is spelled out, so it is never
// mixed up with UM Digos or UM Tagum.
function schoolLabel_(name, official) {
  name = String(name || '').trim(); official = String(official || '').trim();
  if (/^UM( Main)?$/i.test(name) || /^(The )?University of Mindanao$/i.test(official)) return 'University of Mindanao — UM Main';
  const um = name.match(/^UM\s+(.+)$/i);                                  // UM Digos, UM Tagum… each its own school
  if (um) return 'University of Mindanao — UM ' + um[1].replace(/\s*College$/i, '');
  if (!official || official.toLowerCase() === name.toLowerCase()) return name;
  return official.toLowerCase().indexOf(name.toLowerCase()) !== -1 ? official : official + ' (' + name + ')';
}
function joinSchoolList_(ss) {
  const off = {};
  try { ss.getSheetByName(TAB.schools).getRange(2, 1, CONFIG.maxSchools, 9).getValues().forEach(r => { off[String(r[0]).trim()] = r[8]; }); } catch (e) { /* names only */ }
  return schoolRows_(ss).filter(s => !s.inactive).map(s => ({ name: s.name, label: schoolLabel_(s.name, off[s.name]) }))
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
    const last = clip(d.last, 60), first = clip(d.first, 60), mi = clip(d.mi, 3).replace(/[^A-Za-zÑñ]/g, '').slice(0, 2).toUpperCase();
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
               : { ok: true, msg: 'Your request has been sent to your Regional Representative. Once approved, START will email you a verification code to create your password and sign in.' };
    if (webAllowed_(webRoles_(ss, email))) return { ok: true, onList: true, msg: 'You\'re already on the START list. We\'ll send your code now.' };
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
      sh.getRange(savedRow, 1, 1, v.length).setValues([v]);                   // asked again: the newer details win
      markJoinChange_(ss, sh, savedRow, email, i >= 0 ? 'request updated' : 'request created', el.result);
      props.setProperty(k, String(now));
      if (savedRow > 1) props.setProperty(dayKey, String(today + 1));
    } finally { lock.releaseLock(); }
    try { refreshJoinViews_(ss); } catch (err) { Logger.log('Join views: ' + err.message); }
    // Eligible applications from a recognized school are activated immediately.
    // The applicant must still prove control of their email before choosing a password.
    if (autoApproveOn_() && el.result === 'Eligible' && school !== JOIN.other && savedRow > 1) {
      try {
        const r = webJoinAction_(ss, 'join-add', { row: savedRow, auto: true });
        if (r && r.ok) return { ok: true, auto: true, msg: 'Your START details were accepted. Check your email for a verification code to create your password.' };   // the page sends the code next
      } catch (e) { Logger.log('Auto-approve: ' + e.message); }                  // not added: it stays in the list for the RR
    }
    // Requests that do not meet all listed eligibility conditions remain for RR review.
    return done;
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

// Automatic approval requires a single matching pre-existing masterlist record.
// The applicant's self-reported school and program alone are insufficient evidence.
function autoJoinEligible_(ss, request, roster) {
  if (!autoApproveOn_() || String(request[2] || '').trim() === JOIN.other) return false;
  const norm = v => String(v || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim().replace(/\s+/g, ' ');
  const words = v => norm(v).split(' ').filter(Boolean).sort().join(' ');
  const school = String(request[2] || '').trim(), email = emailOf_(request[3]);
  if (joinMatch_(roster, request[1], school, email).i >= 0) return false; // never reactivate or attach an alias automatically
  const hits = oldMasterRows_(ss).filter(o => !String(o.r[10] || '').trim() &&
    email && emailOf_(o.r[4]) === email && String(o.r[1] || '').trim() === school && words(o.r[0]) === words(request[1]) &&
    norm(o.r[2]) === norm(request[7]) && String(o.r[3] || '').trim() === String(request[15] || '').trim() &&
    String(o.r[5] || '').trim() === 'Active');
  return hits.length === 1;
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
      markJoinChange_(ss, sh, row, String(o.actor || 'RR'), 'declined', name); refreshJoinViews_(ss); return { ok: true, msg: name + ' was not added.' }; }
    if (action === 'join-clarify') {                                     // one short email to the applicant; the request stays open
      const q = String(o.note || '').replace(/\s+/g, ' ').trim().slice(0, 300);
      if (q.length < 5) return { ok: false, msg: 'Type what you need to know first.' };
      if (mailRoom_() < 1) return { ok: false, msg: 'No emails left today (daily limit). Try again tomorrow.' };
      const m = joinAskEmail_(name, q);
      try { MailApp.sendEmail({ to: email, subject: m.subject, htmlBody: m.html, body: m.text, name: m.from, replyTo: me_() }); }
      catch (e) { return { ok: false, msg: 'Not sent: ' + e.message }; }
      sh.getRange(row, 5).setValue('Asked for details ' + stamp);
      if (sh.getMaxColumns() >= 12) sh.getRange(row, 12).setValue(q);
      markJoinChange_(ss, sh, row, String(o.actor || 'RR'), 'clarification requested', 'Question sent');
      refreshJoinViews_(ss);
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
    const m = joinMatch_(roster, name, school, email);
    if (m.i >= 0 && String(roster[m.i][5] || '').trim() !== 'Active')
      return { ok: false, msg: 'This scholar is not Active. Verify the record and change its status in Scholars before approving this request.' };
    // what the applicant gave: program (without "Not listed: "), year level (Scholars spelling), scholarship
    const program = String(r[7] || '').replace(/^Not listed:\s*/i, '').trim();
    const yl = String(r[8] || '').trim(), year = /^\d(st|nd|rd|th) year$/i.test(yl) ? yl.replace(/year$/i, 'Year') : '';
    const kind = String(r[9] || '').trim() === 'Not sure' ? '' : String(r[9] || '').trim();
    const from = 'Join request · ' + stamp;
    const oldHit = oldMasterMatch_(oldMasterRows_(ss), name, school), ov = oldHit ? oldHit.row : [];
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
    markScholarChange_(ss, sc, at, o.auto ? 'automatic approval' : String(o.actor || 'RR'), 'join approved', 'Join request row ' + row,
      m.i >= 0 ? String(roster[m.i][5] || '').trim() : '', 'Active');
    markJoinChange_(ss, sh, row, o.auto ? 'automatic approval' : String(o.actor || 'RR'), 'approved', 'Scholar ID at row ' + at);
    if (oldHit) try { oldMasterTab_(ss).getRange(oldHit.at, OLD_MASTER_HEAD.length).setValue('Signed up · Scholars row ' + at + ' · ' + stamp); } catch (e) { /* note only */ }
    if (sh.getMaxColumns() >= 12) sh.getRange(row, 12).setValue('Scholars row ' + at);   // where they are now
    SpreadsheetApp.flush();
    refreshJoinViews_(ss);
  } finally { lock.releaseLock(); }
  MEMBER_CACHE_ = null;
  try { calPeopleChanged_(ss); } catch (e) { Logger.log('Calendar sharing: ' + e.message); }
  // Approval is complete before any password code is sent.
  // A new applicant gets a code. An existing account gets an approval/sign-in email.
  let mailed = false, mailIssue = '';
  const hasPassword = !!(webAccount_(email) || {}).h;
  if (!hasPassword) {
    const sent = webMailCode_(email, 'new', { name: name });
    mailed = !!sent.ok;
    if (!mailed) mailIssue = sent.msg || 'Could not send verification code';
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
      'Email not sent: ' + mailIssue + '. Ask the scholar to use Create Account to request a new code.') };

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
    : 'Your START membership has been approved! A verification code is being sent separately. Open START, select Create Account and enter ' + email + ' to verify your email and set your password. You can request another code if yours has expired.'];
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


/* ---------- 🧹 FRESH START: an empty tracker that matches the website (asks first, backs up first) ----------
 *  1. A full private copy of the tracker is saved next to it ("…backup before fresh start…").
 *  2. Tabs the website does not use are removed: Start here, Dashboard, ChartData, Links, Calendar view,
 *     START Campus Leads, Install steps, Old masterlist, old Google Form tabs.
 *  3. Every row of data is cleared: Scholars, Check-in answers, Campus Lead check-ins, Event attendance,
 *     Events, Join requests, Reminder log, Outreach log. The Schools list and Settings stay (the sign-up form needs them);
 *     only Schools › START Campus Lead is emptied.
 *  4. Website accounts, sign-in codes and sessions, the "already emailed" list and waiting emails are deleted:
 *     everyone (you too) creates their account again; new scholars sign up and you approve them.
 *  5. ① runs: every tab is rebuilt with its headings, dropdowns, formats and look. */
var FRESH_DROP = ['Start here', 'Dashboard', 'ChartData', 'Links', 'Calendar view', 'START Campus Leads', 'Install steps', 'Old masterlist',
                  'Google Form answers (not used)', 'Member check-ins', 'Archive', 'Not recognised'];
var FRESH_CLEAR = () => [TAB.scholars, TAB.mcheckins, TAB.checkins, ATTEND_TAB, TAB.events, JOIN_TAB, TAB.log, TAB.outreach];
var FRESH_PROPS = /^(WEB_A_|WEB_S_|WEB_C_|WEB_T_|WEB_D_|WEB_J_|WEB_JD_|SENT_LEDGER$|MAIL_QUEUE$|ANSWERS_SEEN$|ANSWERS_BACKUP$|LAST_CHECK$|LAST_REPORT$|PORTALS$|PORTALS_ON$|FOCAL_SYNC$)/;

function freshStart() {
  sheetUser_();                                   // sheet only: never from the website
  requireSetup_();
  const ui = SpreadsheetApp.getUi(), ss = SpreadsheetApp.getActive();
  if (!ensureOwner_('start fresh')) return;
  const drop = ss.getSheets().map(s => s.getName()).filter(n => FRESH_DROP.indexOf(n) !== -1 || /^Form Responses \d+$/.test(n) || / \(old/.test(n));
  const msg = 'This empties the tracker so you can start again from the website.\n\n' +
    '• A full backup copy is saved first (private, next to this file).\n' +
    (drop.length ? '• Removed tabs: ' + drop.join(', ') + '\n' : '') +
    '• Cleared: every scholar, check-in, event, attendance, join request and email log.\n' +
    '• Kept: the Schools list and Settings.\n' +
    '• Website accounts are deleted: everyone, you too, creates their account again.\n\n' +
    'Type FRESH START to continue.';
  const r = ui.prompt('🧹 Fresh start', msg, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK || r.getResponseText().trim().toUpperCase() !== 'FRESH START') { notify_('Nothing was changed.'); return; }
  const out = freshStart_(ss, drop);
  notify_('🧹 Fresh start done.\n\nBackup: ' + out.backup + '\n\nRemoved ' + out.dropped + ' tab(s), cleared ' + out.cleared + ' tab(s), ' + out.props + ' website records.\n\n' +
          'Next: open the website and create your account again (Create account, with this email). New scholars use "Ask to join START"; approve them on your Region page.');
}

function freshStart_(ss, drop) {
  // 1. backup (the whole file; private: only you)
  const stamp = Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd HHmm');
  const file = DriveApp.getFileById(ss.getId()), copy = file.makeCopy(ss.getName() + ' (backup before fresh start ' + stamp + ')');
  try { const p = file.getParents(); if (p.hasNext()) copy.moveTo(p.next()); } catch (e) { /* stays in My Drive */ }
  const lock = docLock_();
  if (!lock.tryLock(30000)) throw new Error('START is busy (someone is saving a check-in). Try again in a minute.');
  let dropped = 0, cleared = 0, props = 0;
  try {
    // 2. tabs the website does not use
    (drop || []).forEach(n => { const sh = ss.getSheetByName(n); if (sh && ss.getSheets().length > 1) { ss.deleteSheet(sh); dropped++; } });
    // 3. data rows (headings stay)
    FRESH_CLEAR().forEach(n => {
      const sh = ss.getSheetByName(n);
      if (!sh || sh.getLastRow() < 2) return;
      sh.getRange(2, 1, sh.getMaxRows() - 1, sh.getMaxColumns()).clearContent();
      if (sh.getMaxRows() > 501) sh.deleteRows(502, sh.getMaxRows() - 501);
      cleared++;
    });
    const sch = ss.getSheetByName(TAB.schools);
    if (sch) sch.getRange(2, 4, CONFIG.maxSchools, 1).clearContent();      // START Campus Lead (people) – the list of schools stays
    // the two answer tabs get exactly the headings the website writes (no leftover "Column 10" from old forms)
    [[TAB.mcheckins, MEMBER_HEAD_()], [TAB.checkins, LEAD_HEAD_()]].forEach(([n, head]) => {
      const sh = ss.getSheetByName(n); if (!sh) return;
      sh.getRange(1, 1, 1, sh.getMaxColumns()).clearContent();
      sh.getRange(1, 1, 1, head.length).setValues([head]);
      if (sh.getMaxColumns() > head.length + 1) sh.deleteColumns(head.length + 2, sh.getMaxColumns() - head.length - 1);
      sh.setFrozenRows(1);
    });
    // 4. website accounts, codes, sessions, the sent-email list, waiting emails
    const dp = docProps_(), all = dp.getProperties();
    Object.keys(all).forEach(k => { if (FRESH_PROPS.test(k)) { dp.deleteProperty(k); props++; } });
    dp.setProperty('FRESH_LIST_DONE', String(Date.now()));                 // the Scholars list is new: nothing to move
    dp.deleteProperty('TABLE_LOOK');                                       // every table gets its look again
    removeTriggers_(['sendWaitingEmails', 'refreshPortalsAuto']);
    try { CacheService.getScriptCache().removeAll(['SLOGO_VER', 'WEB_LEADERBOARD']); } catch (e) { /* fine */ }
    MEMBER_CACHE_ = null;
    SpreadsheetApp.flush();
    refreshJoinViews_(ss);
  } finally { lock.releaseLock(); }
  // 5. rebuild every tab's headings, dropdowns and look
  try { setup(); } catch (e) { Logger.log('Setup after fresh start: ' + e.message); }
  return { backup: copy.getUrl(), dropped, cleared, props };
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
    sh.getRange(2, 1, rows, 13).setVerticalAlignment('middle');
  } catch (e) { /* fine */ }
}


/* ---------- auto-approve (the RR's switch on the Region page) ----------
 * ON: a sign-up with a START program and a school from the list is added to Scholars at once and gets the "You're in" email.
 * "Needs review" and "Not eligible" always wait for the RR. OFF (the default): every request waits. */
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

function autoApproveOn_() { try { return docProps_().getProperty('AUTO_APPROVE') === '1'; } catch (e) { return false; } }
function setAutoApprove_(on) { docProps_().setProperty('AUTO_APPROVE', on ? '1' : '0'); return on; }
