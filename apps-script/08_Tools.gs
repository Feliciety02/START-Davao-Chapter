/* =====================================================================
 *  START SCHOLAR TRACKER  ·  08_Tools
 *  ---------------------------------------------------------------------
 *  Inside: Today panel, messages, 🩺 check, links, outreach log, reports, backups, extras.
 *  To customise: Rarely anything.
 *
 *  The code has 9 files + Email_Reminder and Web_App (HTML). Keep these exact names:
 *    00_Config, 01_Region, 02_Menu_Setup, 03_Forms, 04_Emails, 05_Portals, 06_Calendar, 07_Look, 08_Tools
 *  To update: open each file, select all (Ctrl+A), paste the new version, Save.
 *  Then in the Google Sheet: START Tracker › ① Set up / fix the tracker.
 *  You never need to run anything from here: use the START Tracker menu in the Sheet.
 * ===================================================================== */

// The ▶ Run button runs the FIRST function of the open file: this one, so it is always safe.
function about_08_Tools() {
  Logger.log('START Tracker' + (typeof START_MODULES === 'string' ? ' ' + START_MODULES : '') + ', file 08_Tools. Nothing to run here: open the Google Sheet and use the START Tracker menu.');
}


// ━━━━━━━━━━━━━━━━━━━━  Links  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  🔗 LINK BANK
 *  Top part (rows 5-49): every tracker link, kept up to date by itself (forms, calendars, folders, tabs).
 *  Bottom part (row 53 down): your own links. Add as many as you like.
 *  ① refreshes only the top part; your rows are never touched (older trackers: they are moved down once).
 * ================================================================ */

var LINKS = { autoFirst: 5, autoRows: 45, userTitle: 51, userHead: 52, userFirst: 53, userRows: 200 };
var LINK_CATEGORIES = ['Forms', 'Sheets & Drive', 'Meetings', 'Reports', 'Social media', 'Training', 'ONSSE / DOST-SEI', 'Other'];

function buildLinks_(ss) {
  if (DB_MODE) return;                                   // database mode: not used (see 00_Config)
  const sh = getOrCreate_(ss, TAB.links);
  // Older trackers had YOUR LINKS higher up: keep them, and put them back below the bigger top part
  const moved = [];
  const colC = sh.getRange(1, 3, Math.min(sh.getMaxRows(), 80), 1).getValues().map(r => String(r[0]));
  const oldHead = colC.findIndex(v => v === 'Link (paste here)') + 1;
  if (oldHead && oldHead !== LINKS.userHead) {
    sh.getRange(oldHead + 1, 1, LINKS.userRows, 6).getValues()
      .forEach(r => { if (String(r[0]).trim() || String(r[2]).trim()) moved.push([r[0], r[2], r[3], r[4], r[5]]); });
    sh.getRange(oldHead - 1, 1, LINKS.userRows + 2, 6).breakApart().clear();
  }
  if (sh.getMaxColumns() < 6) sh.insertColumnsAfter(sh.getMaxColumns(), 6 - sh.getMaxColumns());
  const need = LINKS.userFirst + LINKS.userRows;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  [240, 80, 260, 130, 300, 100].forEach((w, i) => sh.setColumnWidth(i + 1, w));   // C: links wrap instead of stretching
  sh.getRange(1, 1, need, 6).setFontFamily(THEME.font).setFontColor(THEME.navy).setVerticalAlignment('middle');

  // Title
  sh.getRange('A1:F1').merge().setValue('');                // the picture banner is the title (19_Banners)
  sh.getRange('A2:F2').merge()
    .setValue('Every important link in one place. The top part updates by itself; add your own links in the bottom part.')
    .setFontColor(THEME.muted).setFontSize(10);
  sh.setRowHeight(1, 44);

  // Auto section
  sh.getRange('A3:F3').merge().setValue('🔒  TRACKER LINKS (kept up to date automatically, do not edit)');
  header_(sh.getRange('A3:F3'));
  sh.getRange('A4:F4').setValues([['Name', 'Open', 'Link (copy this)', 'Category', 'What it is for', 'Updated']])
    .setFontWeight('bold').setFontSize(9).setFontColor(THEME.muted).setBackground(THEME.band);

  writeAutoLinks_(ss, sh, true);

  // User section (never cleared)
  sh.getRange(LINKS.userTitle, 1, 1, 6).merge().setValue('✏️  YOUR LINKS (add your own below: name, link, category, notes)');
  header_(sh.getRange(LINKS.userTitle, 1, 1, 6), THEME.blue);
  sh.getRange(LINKS.userHead, 1, 1, 6).setValues([['Name', 'Open', 'Link (paste here)', 'Category', 'Notes', 'Added on']])
    .setFontWeight('bold').setFontSize(9).setFontColor(THEME.muted).setBackground(THEME.band);
  const openF = [];
  for (let r = LINKS.userFirst; r < LINKS.userFirst + LINKS.userRows; r++) {
    openF.push([`=IF(C${r}="","",HYPERLINK(C${r},"Open ↗"))`]);
  }
  sh.getRange(LINKS.userFirst, 2, LINKS.userRows, 1).setFormulas(openF);
  const user = sh.getRange(LINKS.userFirst, 1, LINKS.userRows, 6);
  user.setFontSize(10).setBorder(true, true, true, true, null, true, TABLE_LINE, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(LINKS.userFirst, 4, LINKS.userRows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(LINK_CATEGORIES, true).setAllowInvalid(true).build());
  sh.getRange(LINKS.userFirst, 6, LINKS.userRows, 1).setNumberFormat('mmm d, yyyy').setFontColor(THEME.muted);
  sh.getRange(LINKS.userFirst, 1, LINKS.userRows, 1).setFontWeight('bold');
  sh.getRange(LINKS.userFirst, 3, LINKS.userRows, 1).setFontColor(THEME.muted).setFontSize(9);
  sh.getBandings().forEach(b => b.remove());
  user.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false)
      .setFirstRowColor('#FFFFFF').setSecondRowColor(THEME.band);

  if (moved.length) {                                             // your old links, back in place
    sh.getRange(LINKS.userFirst, 1, moved.length, 1).setValues(moved.map(m => [m[0]]));
    sh.getRange(LINKS.userFirst, 3, moved.length, 4).setValues(moved.map(m => [m[1], m[2], m[3], m[4]]));
  }
  sh.getRange(LINKS.autoFirst, 2, LINKS.autoRows, 1).setHorizontalAlignment('center');
  sh.getRange(LINKS.userFirst, 2, LINKS.userRows, 1).setHorizontalAlignment('center');
  [3, 5].forEach(c => sh.getRange(LINKS.autoFirst, c, need - LINKS.autoFirst + 1, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP));
  fitRows_(sh, LINKS.autoFirst, need - LINKS.autoFirst + 1);
  sh.setFrozenRows(0);
  sh.setHiddenGridlines(true);
  sh.setTabColor(THEME.tabs.links);
  placeTutorial_(sh);
  placeTabBanner_(sh, 6);

  // Warn before anyone types over the automatic part
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).filter(p => p.getDescription() === PROTECT_TAG)
    .forEach(p => p.remove());
  sh.getRange(1, 1, LINKS.userHead, 6).protect().setDescription(PROTECT_TAG).setWarningOnly(true);
}

// Writes the top part (all tracker links). Cheap: also run by the quick ① and after setting up calendars / members.
function writeAutoLinks_(ss, sh, building) {
  if (DB_MODE) return;                                   // database mode: not used (see 00_Config)
  sh = sh || ss.getSheetByName(TAB.links);
  if (!sh) return;
  // safety: if YOUR LINKS are not where this version expects them (older layout), rebuild the tab properly
  // first; it moves your links down, so they are never overwritten
  if (!building && String(sh.getRange(LINKS.userHead, 3).getValue()) !== 'Link (paste here)') return buildLinks_(ss);
  const auto = autoLinks_(ss);
  const block = sh.getRange(LINKS.autoFirst, 1, LINKS.autoRows, 6);
  block.breakApart().clearContent().setFontWeight('normal').setBackground('#F7FAFE');
  const stamp = new Date(), rows = [], open = [];
  auto.forEach((l, i) => {
    const r = LINKS.autoFirst + i;
    rows.push(l.section ? [l.section, '', '', '', '', ''] : [l.name, '', l.url, l.cat, l.note, stamp]);
    open.push([l.section ? '' : `=IF(C${r}="","",HYPERLINK(C${r},"Open ↗"))`]);
  });
  if (rows.length) {
    sh.getRange(LINKS.autoFirst, 1, rows.length, 6).setValues(rows);
    sh.getRange(LINKS.autoFirst, 2, rows.length, 1).setFormulas(open);
  }
  block.setFontSize(10).setFontColor(THEME.navy)
       .setBorder(true, true, true, true, null, true, TABLE_LINE, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(LINKS.autoFirst, 1, LINKS.autoRows, 1).setFontWeight('bold');
  sh.getRange(LINKS.autoFirst, 3, LINKS.autoRows, 1).setFontColor(THEME.muted).setFontSize(9);
  sh.getRange(LINKS.autoFirst, 5, LINKS.autoRows, 1).setFontColor(THEME.muted).setWrap(true);
  sh.getRange(LINKS.autoFirst, 6, LINKS.autoRows, 1).setNumberFormat('mmm d').setFontColor(THEME.muted);
  auto.forEach((l, i) => {                                         // small section bars: Forms, Calendars, Folders …
    if (!l.section) return;
    sh.getRange(LINKS.autoFirst + i, 1, 1, 6).merge().setBackground(THEME.band).setFontColor(THEME.blue)
      .setFontSize(9).setFontWeight('bold');
  });
}

function autoLinks_(ss) {
  const out = [];
  const add = (name, url, cat, note) => { if (url) out.push({ name, url, cat, note }); };
  const section = t => out.push({ section: t });
  const props = docProps_();
  const folderUrl = id => { try { return id ? DriveApp.getFolderById(id).getUrl() : ''; } catch (e) { return ''; } };
  const summary = f => f.getEditUrl().replace(/\/edit.*$/, '/edit#responses');

  section('📝  FORMS');
  let form = null;
  try { form = getForm_(); } catch (e) { /* not set up yet */ }
  if (form) {
    if (typeof oneFormOn_ === 'function' && oneFormOn_()) add('Old Campus Lead form (closed)', form.getPublishedUrl(), 'Forms',
      'Closed: it tells anyone who opens it to use the weekly check-in below.');
    else add('Campus Lead weekly check-in (share with Leads)', form.getPublishedUrl(), 'Forms',
        'General link. Each school\'s own link (school already selected) is in Schools, column P. Reminder emails use it.');
    add('Campus Lead form: edit questions & theme', form.getEditUrl(), 'Forms', 'Owner only. Do not rename the questions.');
    add('Campus Lead form: answers summary', summary(form), 'Forms', 'Charts of all answers inside Google Forms.');
  }
  let mform = null;
  try { mform = typeof memberForm_ === 'function' ? memberForm_() : null; } catch (e) { mform = null; }
  if (!mform && props.getProperty('MFORM_ID')) { try { mform = FormApp.openById(props.getProperty('MFORM_ID')); } catch (e) { mform = null; } }
  if (!mform) {
    out.push({ name: '👥 Member check-in (Leads share it in their GC)', url: '', cat: 'Forms',
               note: 'Not made yet: click START Tracker › 🛠 Setup tools › 📝 Set up / update member check-in. The link appears here by itself.' });
  } else {
    add('📝 Weekly check-in: everyone, Campus Leads too (emailed every week to welcomed members)', mform.getPublishedUrl(), 'Forms',
        `Weekly vibe check for all members. Opens ${getSettings_().mOpenDay}, due ${getSettings_().mDueDay}.`);
    add('Member form: edit questions & theme', mform.getEditUrl(), 'Forms', 'Owner only. Questions change by themselves each week.');
    add('Member form: answers summary', summary(mform), 'Forms', 'Charts of all member answers.');
  }

  let c = null;
  try { c = typeof calStore_ === 'function' ? calStore_() : null; } catch (e) { c = null; }
  if (c && (c.region || c.leads || c.mine)) {
    section('📅  CALENDARS (click to add to your Google Calendar)');
    add(CAL_NAMES.region + ' calendar', c.region && calLink_(c.region), 'Meetings', 'All active members can see it. Share this link in member GCs.');
    add('START Campus Leads calendar', c.leads && calLink_(c.leads), 'Meetings', 'Leads only: weekly check-in days and Lead meetings.');
    add('My START work (private)', c.mine && calLink_(c.mine), 'Meetings', 'Only you: follow-ups and target deadlines.');
    Object.keys(c.schools || {}).sort().forEach(n =>
      add(`START – ${n} calendar`, calLink_(c.schools[n]), 'Meetings', `${n}'s members can see it; its Campus Lead can add events.`));
  }

  section('📁  FOLDERS IN YOUR DRIVE');
  add('Campus Lead portals + end-of-term reports', folderUrl(props.getProperty('PORTAL_FOLDER')), 'Sheets & Drive',
      'Each Lead\'s private portal, the forms and your 📄 reports. Never share the folder itself.');
  try {
    const d = typeof memePicFolder_ === 'function' ? memePicFolder_() : null;
    if (d) add('Meme pictures for the forms', d.getUrl(), 'Sheets & Drive',
      'vibe-1.jpg, events-1.jpg, learning-1.jpg, help-1.jpg … (lead-… for the Campus Lead form). They rotate weekly.');
  } catch (e) { /* no folder yet */ }
  add('Tracker pictures (logos, banners)', folderUrl(props.getProperty('PICTURE_FOLDER')), 'Sheets & Drive',
      'Added once with 🛠 Setup tools › 🖼 Add the tracker pictures.');

  section('📊  THIS TRACKER');
  add('This tracker (Google Sheet)', ss.getUrl(), 'Sheets & Drive', 'Bookmark this. Never share it: it has every scholar\'s data.');
  const tab = (name, label, note) => { const t = ss.getSheetByName(name); if (t) add(label, ss.getUrl() + '#gid=' + t.getSheetId(), 'Sheets & Drive', note); };
  tab(TAB.schools, 'Schools tab (per-school form links)', 'Column P has each school\'s pre-selected form link.');
  tab(TAB.focal, 'START Campus Leads tab (portal links)', 'Each Lead\'s private portal link and status.');
  tab(TAB.events, 'Events tab', 'Add events (or 📅 Events & calendar › Add an event).');
  tab(TAB.calview, 'Calendar view tab', 'The month at a glance.');
  try {
    add('Apps Script code (behind the tracker)', 'https://script.google.com/d/' + ScriptApp.getScriptId() + '/edit',
        'Sheets & Drive', 'Only for updating the code. Day-to-day use is the START Tracker menu.');
  } catch (e) { /* ignore */ }

  section('🌐  OTHER');
  add('Google Calendar', 'https://calendar.google.com', 'Meetings', 'See every START calendar together.');
  add('DOST-SEI website', 'https://www.sei.dost.gov.ph', 'ONSSE / DOST-SEI', 'Official announcements and scholarship information.');

  // drop empty sections (e.g. no calendars yet) and keep within the box
  const clean = out.filter((l, i) => !l.section || (out[i + 1] && !out[i + 1].section));
  return clean.slice(0, LINKS.autoRows);
}


// ━━━━━━━━━━━━━━━━━━━━  Outreach  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  📞 OUTREACH LOG
 *  Every time you contact a school (message, call, visit), log it here.
 *  - START Tracker > 📞 Record a call or message opens a quick form (newest entries go on top)
 *  - Schools tab column Q and the Dashboard show "Last contacted" per school
 * ================================================================ */

var OUTREACH_HEAD = ['Date', 'School', 'Person contacted', 'Channel', 'Outcome', 'Next step', 'Follow up on'];
var OUTREACH_CHANNELS = ['Messenger', 'Group chat', 'Text / SMS', 'Call', 'Email', 'Visit', 'Meeting', 'Other'];
var OUTREACH_OUTCOMES = ['START Campus Lead named', 'Reached – will follow up', 'Reached – not interested',
                         'No response yet', 'Wrong / old contact', 'Other'];
var OUTREACH_ROWS = 1000;

function buildOutreach_(ss) {
  const sh = getOrCreate_(ss, TAB.outreach);
  const need = OUTREACH_ROWS + 1;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  if (sh.getMaxColumns() < OUTREACH_HEAD.length) sh.insertColumnsAfter(sh.getMaxColumns(), OUTREACH_HEAD.length - sh.getMaxColumns());

  const body = (c) => sh.getRange(2, c, OUTREACH_ROWS, 1);
  [1, 7].forEach(c => body(c).setNumberFormat('mmm d, yyyy')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build()));
  body(2).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(TAB.schools).getRange(2, 1, CONFIG.maxSchools, 1), true)
    .setAllowInvalid(true).build());
  body(4).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(OUTREACH_CHANNELS, true).setAllowInvalid(true).build());
  body(5).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(OUTREACH_OUTCOMES, true).setAllowInvalid(true).build());

  if (DB_MODE) {
    sh.getRange(1, 1, 1, OUTREACH_HEAD.length).setValues([OUTREACH_HEAD]);
    sh.setFrozenRows(1);
    return;
  }

  bodyRows_(sh.getRange(2, 1, OUTREACH_ROWS, OUTREACH_HEAD.length));   // repair rows that copied the header
  tableLook_(sh, OUTREACH_ROWS, OUTREACH_HEAD.length, 26, THEME.tabs.outreach);
  band_(sh.getRange(2, 1, OUTREACH_ROWS, OUTREACH_HEAD.length));
  placeTutorial_(sh);
  body(2).setFontWeight('bold').setFontColor(THEME.blue);
  body(1).setFontColor(THEME.muted);
  markColumns_(sh, OUTREACH_HEAD.map((h, i) => [i + 1, 'edit', h, i === 0
    ? 'Newest contacts at the top. Use START Tracker > 📞 Record a call or message, or type directly.' : i === 6
    ? 'Red date = follow-up due today or overdue.' : '', i === 5]), OUTREACH_ROWS, false);

  const E = body(5), G = body(7);
  sh.setConditionalFormatRules([
    chipRule_(E, OUTREACH_OUTCOMES[0], CHIP.green), chipRule_(E, OUTREACH_OUTCOMES[1], CHIP.blue),
    chipRule_(E, OUTREACH_OUTCOMES[2], CHIP.grey), chipRule_(E, OUTREACH_OUTCOMES[3], CHIP.yellow),
    chipRule_(E, OUTREACH_OUTCOMES[4], CHIP.red),
    // follow-up date due today or overdue
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($G2<>"",$G2<=TODAY())')
      .setBackground(CHIP.red.bg).setFontColor(CHIP.red.fg).setBold(true).setRanges([G]).build(),
  ]);
  placeTabBanner_(sh, OUTREACH_HEAD.length);
}

/* ---- 📞 quick-entry box ---- */

function showOutreachDialog() {
  requireSetup_();
  const ss = SpreadsheetApp.getActive();
  const schools = schoolNames_(ss);
  if (!schools.length) throw new Error('Add your schools in the Schools tab first.');
  const focal = {};
  schoolRows_(ss).forEach(s => { focal[s.name] = s.focal; });
  const today = Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd');
  const data = { schools, focal, channels: OUTREACH_CHANNELS, outcomes: OUTREACH_OUTCOMES, today };

  const html = HtmlService.createHtmlOutput(DLG_CSS + `
<style>.chip{border:1px solid #D6E4F5;background:#EAF3FC;border-radius:999px;padding:6px 12px;cursor:pointer;font-size:12px;font-weight:600}
  .chip.on{background:#FFDD00;border-color:#FFDD00}.half{display:flex;gap:10px}.half>div{flex:1}</style>
<div class="dbar"></div>
<label style="margin-top:0">School</label><select id="school"></select>
<div class="half"><div><label>Date</label><input type="date" id="date"></div>
  <div><label>Talked to</label><input id="person"></div></div>
<label>How</label><div class="chips" id="channels"></div>
<label>Result</label><div class="chips" id="outcomes"></div>
<label>Next step</label><textarea id="next" style="height:54px" placeholder="optional"></textarea>
<label>Follow up on</label><input type="date" id="follow">
<div style="margin-top:14px"><button class="btn" id="save">Save</button></div><div class="msg" id="msg"></div>
<script>
  var D = ${JSON.stringify(data).replace(/</g, '\\u003c')};
  var pick = { channel: '', outcome: '' };
  var s = document.getElementById('school');
  D.schools.forEach(function (n) { var o = document.createElement('option'); o.value = n; o.textContent = n; s.appendChild(o); });
  document.getElementById('date').value = D.today;
  function fillPerson() { var p = document.getElementById('person'); if (!p.dataset.typed) p.value = D.focal[s.value] || ''; }
  s.onchange = fillPerson; fillPerson();
  document.getElementById('person').oninput = function () { this.dataset.typed = '1'; };
  function chips(id, list, key) {
    var box = document.getElementById(id);
    list.forEach(function (v) {
      var c = document.createElement('span'); c.className = 'chip'; c.textContent = v;
      c.onclick = function () { [].forEach.call(box.children, function (x) { x.classList.remove('on'); });
                                c.classList.add('on'); pick[key] = v; };
      box.appendChild(c);
    });
  }
  chips('channels', D.channels, 'channel'); chips('outcomes', D.outcomes, 'outcome');
  document.getElementById('save').onclick = function () {
    var m = document.getElementById('msg');
    if (!pick.channel || !pick.outcome) { m.className = 'msg err'; m.textContent = 'Pick how and the result.'; return; }
    m.className = 'msg muted'; m.textContent = 'Saving…';
    google.script.run
      .withSuccessHandler(function () {
        m.className = 'msg ok'; m.textContent = '✓ Saved for ' + s.value;
        document.getElementById('next').value = ''; document.getElementById('follow').value = '';
      })
      .withFailureHandler(function (e) { m.className = 'msg err'; m.textContent = 'Not saved: ' + e.message; })
      .saveOutreach({ school: s.value, date: document.getElementById('date').value,
                      person: document.getElementById('person').value, channel: pick.channel, outcome: pick.outcome,
                      next: document.getElementById('next').value, follow: document.getElementById('follow').value });
  };
</script>`).setWidth(500).setHeight(650);
  ui_().showModalDialog(html, '📞 Record a call or message');
}

// Called from the dialog (must not end with "_" so the dialog can reach it)
function saveOutreach(d, spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActive();
  if (!d || !d.school) throw new Error('No school selected.');
  let sh = ss.getSheetByName(TAB.outreach);
  if (!sh) { buildOutreach_(ss); sh = ss.getSheetByName(TAB.outreach); }
  const toDate = s => { if (!s) return ''; const p = String(s).split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); };
  sh.insertRowAfter(1);
  bodyRows_(sh.getRange(2, 1, 1, OUTREACH_HEAD.length)).setWrap(true);
  sh.getRange(2, 1).setFontColor(THEME.muted);
  sh.getRange(2, 2).setFontWeight('bold').setFontColor(THEME.blue);
  sh.getRange(2, 1, 1, OUTREACH_HEAD.length).setValues([[
    toDate(d.date) || new Date(), d.school, String(d.person || '').trim(), d.channel, d.outcome,
    String(d.next || '').trim(), toDate(d.follow)]]);
  fitRows_(sh, 2, 1);
  // "START Campus Lead named" also fills the Schools tab if it is still blank
  if (d.outcome === OUTREACH_OUTCOMES[0] && String(d.person || '').trim()) {
    const row = schoolRows_(ss).find(s => s.name === d.school);
    if (row && !row.focal) {
      assignFocal_(ss.getSheetByName(TAB.schools), row.row, d.person, null,
        'Entered from the Outreach log on ' + Utilities.formatDate(new Date(), CONFIG.timezone, 'MMMM d, yyyy') + '.');
      try { syncFocal_(ss); } catch (e) { Logger.log('Focal sync after outreach: ' + e.message); }
    }
  }
  return true;
}


// ━━━━━━━━━━━━━━━━━━━━  Template  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  🧳 CLEAN TEMPLATE FOR OTHER RRs
 *  Makes a copy of this tracker with the code, design and school list,
 *  but WITHOUT scholar data, check-ins, member check-ins, outreach notes, contacts or your form.
 *  Your own tracker is not changed.
 * ================================================================ */

function makeCleanTemplate() {
  requireSetup_();
  const ui = ui_();
  const ok = ui.alert('Make a blank copy for other RRs?',
    'This creates a NEW copy of the tracker for other Regional Representatives.\n\n' +
    'Kept: the code, design, menus and Start here guide.\n' +
    'Removed: scholars, schools, check-ins, outreach notes, START Campus Leads and contacts, reminder log, ' +
    'your region, settings and private links. The new RR picks THEIR region when they click ①. The copy gets no form until its new owner clicks ①.\n\n' +
    'Also send them START_Tracker_pictures.zip: they add it once with 🖼 Add the tracker pictures.\n\n' +
    'Your own tracker is NOT changed. Continue?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActive();
  ss.toast('Making a clean copy… about 30 seconds.', 'START Tracker', 30);
  const myFormId = docProps_().getProperty('FORM_ID');
  const file = DriveApp.getFileById(ss.getId());
  const name = 'START Tracker – TEMPLATE (no data)';
  const parents = file.getParents();
  const copyFile = parents.hasNext() ? file.makeCopy(name, parents.next()) : file.makeCopy(name);
  Utilities.sleep(3000);
  const t = SpreadsheetApp.openById(copyFile.getId());

  // 1. Detach any form Google copied along with the file (never touches YOUR form)
  t.getSheets().forEach(sh => {
    const url = sh.getFormUrl();
    if (!url) return;
    let detached = false;
    try {
      const f = FormApp.openByUrl(url);
      if (f.getId() !== myFormId && f.getDestinationId() === t.getId()) {
        f.removeDestination();
        DriveApp.getFileById(f.getId()).setTrashed(true);
        detached = true;
      }
    } catch (e) { Logger.log('Copied form not removed: ' + e.message); }
    if (!detached || /^(Member check-ins|Check-in answers)/.test(sh.getName())) { t.deleteSheet(sh); return; }   // member answers never travel
    // keep an EMPTY Weekly Check-in tab so the template's formulas don't show #REF! before ①
    const w = ckCols_(sh).width;
    if (sh.getMaxRows() > 1) sh.getRange(2, 1, sh.getMaxRows() - 1, w).clearContent().clearNote();
    if (sh.getName() !== TAB.checkins) sh.setName(TAB.checkins);
  });

  // website check-in: the answer tabs are plain tables, so empty them here (answers never travel)
  if (WEB_CHECKIN) [TAB.mcheckins, TAB.checkins].forEach(n => {
    const sh = t.getSheetByName(n);
    if (sh && sh.getMaxRows() > 1) sh.getRange(2, 1, sh.getMaxRows() - 1, sh.getMaxColumns()).clearContent().clearNote();
  });

  // 2. Remove personal data
  const clear = (tab, row, col, rows, cols) => {
    const sh = t.getSheetByName(tab);
    if (!sh) return;
    const n = rows || Math.max(sh.getMaxRows() - row + 1, 1);
    sh.getRange(row, col, n, cols).clearContent();
  };
  clear(TAB.scholars, 2, 1, null, 13);                  // A:M (Role, other emails, scholarship, added from): no personal data in a copy
  clear(TAB.schools, 2, 1, CONFIG.maxSchools, 4);       // school list + START Campus Lead (another region has other schools; E is a formula)
  clear(TAB.schools, 2, 9, CONFIG.maxSchools, 7);       // school details I–O
  clear(TAB.schools, 2, 6, CONFIG.maxSchools, 1);       // inactive
  const tsch = t.getSheetByName(TAB.schools);           // notes there can hold names ("you typed …")
  if (tsch) tsch.getRange(2, 4, CONFIG.maxSchools, 1).clearNote();
  clear(TAB.schools, 2, LINK_COL, CONFIG.maxSchools, 1); // per-school links pointed to YOUR form
  clear(TAB.outreach, 2, 1, null, OUTREACH_HEAD.length);
  [TAB.log, TAB.focal, TAB.checkins + ' (old)', 'Archive', 'Not recognised', 'Old masterlist', 'Join requests', 'Event attendance', 'Google Form answers (not used)']
    .forEach(n => { const s = t.getSheetByName(n); if (s) t.deleteSheet(s); });   // personal data never travels in a copy

  const links = t.getSheetByName(TAB.links);
  if (links) {
    links.getRange(LINKS.autoFirst, 1, LINKS.autoRows, 6).clearContent();
    links.getRange(LINKS.userFirst, 1, LINKS.userRows, 1).clearContent();
    links.getRange(LINKS.userFirst, 3, LINKS.userRows, 4).clearContent();
  }
  const evs = t.getSheetByName(TAB.events);                    // events and their calendar links stay with you
  if (evs && evs.getMaxRows() > 1) evs.getRange(2, 1, evs.getMaxRows() - 1, EV_HEAD.length).clearContent();
  // notes can hold names ("same email as …"), the leaderboard and the month view hold school results and events
  const tsc = t.getSheetByName(TAB.scholars);
  if (tsc && tsc.getMaxRows() > 1) { tsc.getRange(2, 1, tsc.getMaxRows() - 1, tsc.getMaxColumns()).clearNote(); tsc.getRange(2, 5, tsc.getMaxRows() - 1, 1).setBackground(null); }   // orange email checks
  const tdash = t.getSheetByName(TAB.dashboard);
  if (tdash) {
    const colA = tdash.getRange(1, 1, tdash.getMaxRows(), 1).getValues().map(r => String(r[0]));
    const top = colA.findIndex(v => v.indexOf(LEADERBOARD_MARK) === 0) + 1;
    if (top > 0) tdash.getRange(top, 1, Math.min(70, tdash.getMaxRows() - top + 1), 6).breakApart().clear();
  }
  const tcv = t.getSheetByName(TAB.calview);
  if (tcv && tcv.getMaxRows() >= CV.firstWeek) tcv.getRange(CV.firstWeek, 1, tcv.getMaxRows() - CV.firstWeek + 1, tcv.getMaxColumns()).clearContent();
  const set = t.getSheetByName(TAB.settings);
  if (set) ['region', 'rrName', 'recipients', 'replyTo'].forEach(k => set.getRange(settingCell_(k).split('!')[1]).clearContent());
  const dash = t.getSheetByName(TAB.dashboard);
  if (dash) dash.getRange('B3').setValue(new Date());
  if (set) try { set.getRange(settingCell_('termStart').split('!')[1]).setValue(new Date()); } catch (e) { /* older copy */ }

  // 3. Ready-to-share link
  const copyLink = t.getUrl().replace(/\/edit.*$/, '/copy');
  addUserLink_(ss, 'Clean template for other RRs (share this)', copyLink, 'Sheets & Drive',
               'Opens "Make a copy" of the empty template. Safe to share: it has no scholar data.');
  notify_('✅ Blank copy ready: "' + name + '"\n\n' +
          'Share this link with other RRs (no scholar data):\n' + copyLink + '\n\n' +
          '');
}

// Adds (or updates) a row in the "YOUR LINKS" part of the link bank
function addUserLink_(ss, name, url, cat, note) {
  if (DB_MODE) return;                                   // database mode: not used (see 00_Config)
  const sh = ss.getSheetByName(TAB.links);
  if (!sh) return;
  const vals = sh.getRange(LINKS.userFirst, 1, LINKS.userRows, 1).getValues().map(r => String(r[0]).trim());
  let i = vals.indexOf(name);
  if (i === -1) i = vals.indexOf('');
  if (i === -1) return;
  const r = LINKS.userFirst + i;
  sh.getRange(r, 1).setValue(name);
  sh.getRange(r, 3, 1, 4).setValues([[url, cat, note, new Date()]]);
}


// ━━━━━━━━━━━━━━━━━━━━  Tools  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  🧰 TOOLS
 *  🗓 Today          a side panel with only what needs you today
 *  📣 Write a message  to Campus Leads, members or everyone (same design as the other emails)
 *  🩺 Check          finds old code files, missing settings, emails that are off, privacy risks
 * ================================================================ */

/* ================================================================
 *  🗓 TODAY  (side panel on the right of the sheet)
 * ================================================================ */

function showToday() {
  requireSetup_();
  const html = HtmlService.createHtmlOutput(todayHtml_()).setTitle('🗓 Today');
  ui_().showSidebar(html);
}

// What needs the RR today. Every item can jump to the cell it is about.
function todayData() {
  const ss = SpreadsheetApp.getActive();
  const st = getSettings_();
  const now = new Date(), today = startOfDay_(now);
  const data = focalData_(ss);                    // schools + their check-ins (one read)
  const w = reminderWeek_(now);
  const out = { date: Utilities.formatDate(now, CONFIG.timezone, 'EEEE, MMM d'), groups: [] };
  const add = (title, tone, items, empty) => out.groups.push({ title, tone, items, empty });

  // 1. help requests in the last 7 days
  const since = new Date(now.getTime() - 7 * 86400000), help = [];
  data.forEach(s => s.checkins.forEach(c => {
    const sos = typeof L_VIBES !== 'undefined' && samePick_(c.vibe, L_VIBES[3]);          // 🆘 in the Lead's kamustahan
    if (c.date >= since && (String(c.help).indexOf('Yes') === 0 || sos)) {
      help.push({ text: `${s.name}: ${sos ? c.vibe + ' · ' : ''}${c.help}${c.concern ? ' · ' + String(c.concern).slice(0, 80) : ''}`,
                  tab: TAB.schools, a1: 'A' + s.row });
    }
  }));
  add('Help requests (last 7 days)', 'red', help, 'None 🎉');
  try {                                                         // 👥 members who asked for help (23_Members)
    if (checkinOn_()) add('Members asking for help (last 7 days)', 'red', memberData_(ss).helpRecent.map(h =>
      ({ text: `${h.school} · ${h.name} · ${h.help}${samePick_(h.vibe, M_VIBES[3]) ? ' · ' + h.vibe : ''}${h.tea ? ' · ' + String(h.tea).slice(0, 60) : ''}`,
         tab: TAB.mcheckins, a1: 'A1' })), 'None 🎉');
  } catch (e) { /* member check-in not set up */ }
  try {                                                         // answers waiting because their email is not in Scholars
    const un = ss.getSheetByName(UNKNOWN_TAB);
    if (un && un.getLastRow() > 1) add('Answers waiting: email not in Scholars', 'yellow',
      un.getRange(2, 1, un.getLastRow() - 1, Math.min(un.getLastColumn(), 2)).getValues()
        .map((r, i) => ({ text: `${r[1] || '(no email)'} · answered ${r[0] instanceof Date ? Utilities.formatDate(r[0], CONFIG.timezone, 'MMM d') : ''} · add this email in Scholars`,
                          tab: UNKNOWN_TAB, a1: 'A' + (i + 2) })), '');
  } catch (e) { /* nothing waiting */ }

  try {                                                         // 📨 members with an email who have no welcome email yet (sent only by you)
    const calReady = calOn_() && !!calStore_().region;
    const nw = welcomeList_(ss).filter(p => !alreadySent_('welcome', 'v1', p.email) || (calReady && !alreadySent_('welcomeCal', 'v1', p.email)));
    if (nw.length) add(`Not welcomed yet (${nw.length}): send with ✉️ Send emails now › 📨 Send welcome emails`, 'blue',
      nw.slice(0, 15).map(p => ({ text: `${p.name} · ${p.school}${p.isLead ? ' · ⭐ Lead' : ''} · ${p.email}`, tab: TAB.scholars, a1: 'E1' })), '');
  } catch (e) { /* calendar or members not set up */ }

  // 2. schools with a Campus Lead that have not checked in this week
  const quiet = {};                                // schools quiet 2+ weeks in a row (21_Reports)
  try { missedSchools_(ss).forEach(q => { quiet[q.name] = q.missed; }); } catch (e) { /* no check-ins yet */ }
  const missing = data.filter(s => s.focal && !s.thisWeek)
    .sort((a, b) => (quiet[b.name] || 0) - (quiet[a.name] || 0))
    .map(s => ({ text: `${quiet[s.name] ? '🔔 quiet ' + quiet[s.name] + ' weeks · ' : ''}${s.name} · ${displayName_(s.focal)}${s.last ? ' · last ' + Utilities.formatDate(s.last, CONFIG.timezone, 'MMM d') : ' · never'}`,
                 tab: TAB.schools, a1: 'A' + s.row }));
  add(`Not checked in this week (due ${Utilities.formatDate(w.deadline, CONFIG.timezone, 'EEE, MMM d')})`, 'yellow', missing, 'Everyone has answered 🎉');

  // 3. follow-ups due from the Outreach log
  const due = [], ol = ss.getSheetByName(TAB.outreach);
  if (ol && ol.getLastRow() > 1) {
    ol.getRange(2, 1, ol.getLastRow() - 1, OUTREACH_HEAD.length).getValues().forEach((r, i) => {
      const d = r[6];
      if (d instanceof Date && startOfDay_(d) <= today) {
        due.push({ text: `${r[1]} · ${Utilities.formatDate(d, CONFIG.timezone, 'MMM d')}${r[5] ? ' · ' + String(r[5]).slice(0, 60) : ''}`,
                   tab: TAB.outreach, a1: 'A' + (i + 2) });
      }
    });
  }
  add('Follow-ups due (Outreach log)', 'yellow', due, 'Nothing due');

  // 3b. events in the next 7 days (Events tab)
  try {
    if (ss.getSheetByName(TAB.events)) add('Coming up (next 7 days)', 'blue', upcomingEvents_(ss, 7).map(e =>
      ({ text: `${eventWhen_(e)} · ${(EV_ICON[e.type] || '📌')} ${e.title} · ${e.who}`, tab: TAB.events, a1: 'A' + e.row })), 'No events this week');
  } catch (e) { /* calendar not set up */ }

  // 4. missing Campus Leads / emails
  const fix = [];
  data.filter(s => !s.focal).forEach(s => fix.push({ text: `${s.name}: no START Campus Lead yet`, tab: TAB.schools, a1: 'D' + s.row }));
  const sc = ss.getSheetByName(TAB.scholars);
  data.filter(s => s.focal && !s.email).forEach(s => fix.push({ text: `${s.name}: add ${displayName_(s.focal)}'s email in Scholars`,
    tab: TAB.scholars, a1: 'E' + (sc ? scholarRowNumber_(sc, s.focal) : 2) }));
  add('To complete', 'blue', fix, 'All schools have a Campus Lead with an email');

  // 5. switches
  out.status = [
    st.autoWeekly === true ? '🚀 Weekly emails: ON' : '⏸ Weekly emails: OFF (⏰ Automatic emails › 🚀)',
    portalsOn_() ? '🔐 Portals: ON' : '🔐 Portals: not created yet (🛠 Setup tools › 🔐)',
    automationOwner_() ? '📧 Automatic emails sent from: ' + automationOwner_() : '📧 Email account: not chosen yet (⏰ Automatic emails › 🚀)',
  ].concat(waitingCount_() ? [`⏳ ${waitingCount_()} email(s) waiting for tomorrow morning (daily limit)`] : []);
  return out;
}

// Called from the side panel: show a cell (name must not end with "_")
function goToCell(tab, a1) {
  const sh = SpreadsheetApp.getActive().getSheetByName(tab);
  if (!sh) return;
  sh.activate();
  sh.getRange(a1).activate();
}

function todayHtml_() {
  return DLG_CSS + `
<style>
  body{padding:14px}.date{color:#4B5B77;font-size:12px;margin:0 0 6px}
  .g{margin:16px 0 6px;font-weight:800;font-size:11px;letter-spacing:1px;text-transform:uppercase;display:flex;align-items:center;gap:6px}
  .g::before{content:"";width:4px;height:14px;border-radius:4px;background:#FFDD00}
  .g span{margin-left:auto;background:#12203A;color:#FFDD00;border-radius:999px;padding:0 8px;font-size:10px}
  .it{padding:8px 10px;border-radius:10px;margin:5px 0;cursor:pointer;line-height:1.35;border:1px solid transparent}
  .it:hover{border-color:#0099FF}
  .red{background:#FADCD9}.yellow{background:#FFF6BF}.blue{background:#EAF3FC}
  .none{color:#4B5B77;font-size:12px;padding:2px 2px}
  .st{background:#12203A;color:#B9C3D6;border-radius:12px;padding:10px 12px;margin-top:16px;font-size:11px;line-height:1.7}
</style>
<div class="dbar"></div><p class="lead" style="margin:0">Today</p><div class="date" id="d">Loading…</div><div id="list"></div><div class="st" id="st"></div>
<p style="margin-top:10px"><button class="link" onclick="load()">↻ Refresh</button></p>
<script>
  function esc(s){return String(s).replace(/[&<>"]/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
  function draw(D){
    document.getElementById('d').textContent = D.date;
    var h = '';
    D.groups.forEach(function(g){
      h += '<div class="g">' + esc(g.title) + '<span>' + g.items.length + '</span></div>';
      if (!g.items.length) h += '<div class="none">' + esc(g.empty) + '</div>';
      g.items.forEach(function(it, i){
        h += '<div class="it ' + g.tone + '" data-tab="' + esc(it.tab) + '" data-a1="' + esc(it.a1) + '">' + esc(it.text) + '</div>';
      });
    });
    document.getElementById('list').innerHTML = h;
    document.getElementById('st').innerHTML = D.status.map(esc).join('<br>');
    [].forEach.call(document.querySelectorAll('.it'), function(el){
      el.onclick = function(){ google.script.run.goToCell(el.getAttribute('data-tab'), el.getAttribute('data-a1')); };
    });
  }
  function load(){
    document.getElementById('d').textContent = 'Loading…';
    google.script.run.withSuccessHandler(draw)
      .withFailureHandler(function(e){ document.getElementById('d').textContent = 'Could not load: ' + e.message; })
      .todayData();
  }
  load();
</script>`;
}

/* ================================================================
 *  📣 WRITE A MESSAGE  to Campus Leads, members or everyone (same design as every email)
 * ================================================================ */

function showAnnouncement() {
  requireSetup_();
  const t = reminderTargets_();
  const members = announceMembers_();
  const data = { leads: t.ready.length, members: members.length, skipped: t.skipped.map(s => s.name) };
  dlg_('📣 Write a message', `
<label style="margin-top:0">To</label>
<div class="chips" id="to"><button data-v="leads" class="on">Campus Leads</button><button data-v="members">Members</button><button data-v="all">Everyone</button></div>
<div class="muted" id="who" style="margin-top:6px"></div>
<label>Subject</label><input id="subject" placeholder="START General Assembly, Oct 25">
<label>Message</label><textarea id="body" style="height:130px" placeholder="Each person is greeted by name."></textarea>
<div class="row" style="gap:10px"><div style="flex:1"><label>Link</label><input id="link" placeholder="optional"></div>
  <div style="flex:1"><label>Button text</label><input id="btn" placeholder="optional"></div></div>
<div class="row" style="margin-top:14px"><button class="btn" id="send">Send</button><button class="link" id="test">Send a test to me</button></div>
<div class="msg" id="msg"></div>
<script>
  var D = ${JSON.stringify(data).replace(/</g, '\\u003c')}, TO = 'leads';
  function val(id){ return document.getElementById(id).value; }
  function count(){ return (TO !== 'members' ? D.leads : 0) + (TO !== 'leads' ? D.members : 0); }
  function who(){
    var p = [];
    if (TO !== 'members') p.push(D.leads + ' Lead' + (D.leads === 1 ? '' : 's'));
    if (TO !== 'leads') p.push(D.members + ' member' + (D.members === 1 ? '' : 's'));
    document.getElementById('who').textContent = p.join(' + ') + (TO !== 'members' && D.skipped.length ? ' · no Lead email: ' + D.skipped.join(', ') : '');
    document.getElementById('send').textContent = 'Send to ' + count();
  }
  [].forEach.call(document.querySelectorAll('#to button'), function(b){ b.onclick = function(){
    [].forEach.call(document.querySelectorAll('#to button'), function(x){ x.classList.remove('on'); }); b.classList.add('on'); TO = b.getAttribute('data-v'); who(); }; });
  who();
  function payload(){ return { to: TO, subject: val('subject'), body: val('body'), link: val('link'), button: val('btn') }; }
  function run(fn, label){
    var m = document.getElementById('msg');
    if (!val('subject').trim() || !val('body').trim()) { m.className = 'msg err'; m.textContent = 'Add a subject and a message.'; return; }
    m.className = 'msg muted'; m.textContent = label;
    google.script.run.withSuccessHandler(function(r){ m.className = 'msg ok'; m.textContent = r; })
      .withFailureHandler(function(e){ m.className = 'msg err'; m.textContent = 'Not sent: ' + e.message; })[fn](payload());
  }
  document.getElementById('test').onclick = function(){ run('sendAnnouncementTest', 'Sending a test…'); };
  document.getElementById('send').onclick = function(){
    var m = document.getElementById('msg');
    if (!count()) { m.className = 'msg err'; m.textContent = 'Nobody to send to yet.'; return; }
    if (!confirm('Send to ' + count() + ' people now?')) return;
    run('sendAnnouncement', 'Sending…');
  };
</script>`, 540, 540);
}

// Members for a message: every active scholar with an email who is not a Campus Lead
function announceMembers_() {
  try { return welcomeList_(SpreadsheetApp.getActive()).filter(p => !p.isLead); } catch (e) { return []; }
}

function cleanAnnouncement_(d) {
  d = d || {};
  const subject = String(d.subject || '').trim(), body = String(d.body || '').trim();
  if (!subject || !body) throw new Error('Please write a subject and a message.');
  let link = String(d.link || '').trim();
  if (link && !/^https?:\/\//i.test(link)) link = 'https://' + link;
  return { subject: /^START\b/i.test(subject) ? subject : 'START · ' + subject, paras: body.split(/\n\s*\n/).map(p => p.trim()).filter(Boolean), link, button: String(d.button || '').trim() };
}

function announcementEmail_(a, s) {
  const html = designedEmail_({
    badge: 'Announcement', title: ['START', 'Announcement'],
    greeting: s.focal ? `Dear ${displayName_(s.focal)},` : `Dear START Campus Lead of ${s.name},`,
    intro: a.paras,
    link: a.link || (oneFormOn_() ? checkinLink_() : s.link), buttonText: a.link ? (a.button || 'Open link') : 'Answer the check-in',
    buttonNote: '',
    chips: [['', ''], ['', ''], ['', '']], outro: [],
  });
  const text = [s.focal ? `Dear ${displayName_(s.focal)},` : `Dear START Campus Lead of ${s.name},`, '']
    .concat(a.paras.join('\n\n'), '', (a.link ? (a.button || 'Link') + ': ' + a.link : ''), '',
            'Respectfully,', rrSignature_(getSettings_()).join('\n')).join('\n');
  return { html, text };
}

// "Malasarte, Fe Anne L." -> "Fe Anne"   ·   "Fe Anne Malasarte" -> "Fe"
function firstName_(n) {
  n = String(n || '').trim();
  if (!n) return '';
  const after = n.indexOf(',') !== -1 ? n.split(',')[1].trim() : n.split(/\s+/)[0];
  return after.replace(/(\s+[A-Z]\.?)+$/, '').replace(/\s+(Jr\.?|Sr\.?|II|III|IV)$/i, '').trim();
}

function memberAnnouncementEmail_(a, p) {
  const hi = 'Hi ' + (firstName_(p.name) || 'START scholar') + ',';
  const html = designedEmail_({
    badge: 'Announcement', title: ['START', 'Announcement'],
    greeting: hi, intro: a.paras,
    link: a.link || checkinLink_(), buttonText: a.link ? (a.button || 'Open link') : 'Answer the check-in',
    buttonNote: '',
    chips: [['', ''], ['', ''], ['', '']], outro: [],
  });
  const text = [hi, ''].concat(a.paras.join('\n\n'), '', (a.link ? (a.button || 'Link') + ': ' + a.link : ''), '',
            'Respectfully,', rrSignature_(getSettings_()).join('\n')).join('\n');
  return { html, text };
}

// Called from the dialog
function sendAnnouncementTest(d) {
  const a = cleanAnnouncement_(d), to = (d && d.to) || 'leads';
  const me = Session.getEffectiveUser().getEmail();
  const sent = [];
  if (to !== 'members') {
    const first = reminderTargets_().ready[0] || { name: 'Sample School', focal: 'Sample START Campus Lead', link: checkinUrl_() };
    const m = announcementEmail_(a, first);
    sendMail_({ to: me, subject: '[TEST · Campus Lead version] ' + a.subject, htmlBody: m.html, body: m.text,
                name: 'START Tracker', inlineImages: logoInlineImages_() });
    sent.push('the Campus Lead version');
  }
  if (to !== 'leads') {
    const p = announceMembers_()[0] || { name: 'Sample Scholar', school: 'Sample School' };
    const m = memberAnnouncementEmail_(a, p);
    sendMail_({ to: me, subject: '[TEST · member version] ' + a.subject, htmlBody: m.html, body: m.text,
                name: 'START Tracker', inlineImages: logoInlineImages_() });
    sent.push('the member version');
  }
  return `Test sent to ${me} (${sent.join(' and ')}). Nothing went to anyone else.`;
}

// Called from the dialog
function sendAnnouncement(d) {
  const a = cleanAnnouncement_(d), to = (d && d.to) || 'leads';
  const r = to !== 'members' ? sendAnnouncementTo_(a, null) : { sent: 0, waiting: 0, done: 0 };
  const m = to !== 'leads' ? sendMemberAnnouncementTo_(a, null) : { sent: 0, waiting: 0, done: 0 };
  const parts = [];
  if (to !== 'members') parts.push(`${r.sent} Campus Lead(s)`);
  if (to !== 'leads') parts.push(`${m.sent} member(s)`);
  return `Sent to ${parts.join(' and ')}.` +
    (r.waiting + m.waiting ? ` ${r.waiting + m.waiting} more will go out tomorrow morning (Google's daily email limit).` : '') +
    (r.done + m.done ? ` ${r.done + m.done} already had this message, so they were skipped.` : '') +
    ' Every email is in the Reminder log.';
}

// Members (not Leads). onlyEmails: only these (used by the waiting list the next morning)
function sendMemberAnnouncementTo_(a, onlyEmails) {
  const ss = SpreadsheetApp.getActive(), st = getSettings_(), stamp = new Date();
  const from = st.rrName ? `${st.rrName} (START ${st.region || ''})`.replace(/\s+\)/, ')') : 'START Tracker';
  const aid = announceId_(a), log = [], waiting = [];
  let room = mailRoom_(), sent = 0, done = 0;
  announceMembers_().filter(p => !onlyEmails || onlyEmails.indexOf(p.email) !== -1).forEach(p => {
    if (alreadySent_('announce', aid, p.email)) { done++; return; }   // ✅ same message already received
    if (room < 1) { waiting.push(p.email); return; }
    try {
      const m = memberAnnouncementEmail_(a, p);
      sendMail_({ to: p.email, subject: a.subject, htmlBody: m.html, body: m.text, name: from });
      markSent_('announce', aid, p.email);
      log.push([stamp, 'Message to members', p.school, p.email, 'Sent']); sent++; room--;
    } catch (e) { log.push([stamp, 'Message to members', p.school, p.email, 'Not sent: ' + e.message]); }
  });
  if (waiting.length) {
    const kept = queueEmails_({ kind: 'mannounce', schools: waiting, a: a });
    waiting.forEach(e => log.push([stamp, 'Message to members', '', e, kept ? WAIT_NOTE : 'Not sent: daily email limit reached']));
  }
  if (log.length) { try { writeLog_(ss, log); } catch (e) { Logger.log('Reminder log: ' + e.message); } }
  return { sent, waiting: waiting.length, done };
}

// onlyNames: only these schools (used by the waiting list the next morning)
function sendAnnouncementTo_(a, onlyNames) {
  const ss = SpreadsheetApp.getActive(), st = getSettings_();
  const t = reminderTargets_(), logos = logoInlineImages_(), stamp = new Date();
  const from = st.rrName ? `${st.rrName} (START ${st.region || ''})`.replace(/\s+\)/, ')') : 'START Tracker';
  const log = [], waiting = [];
  let sample = null, room = mailRoom_(), done = 0;
  const aid = announceId_(a);
  t.ready.filter(s => !onlyNames || onlyNames.indexOf(s.name) !== -1).forEach(s => {
    if (alreadySent_('announce', aid, s.email)) { done++; return; }   // ✅ same announcement already received
    if (room < 1) { waiting.push(s); return; }
    try {
      const m = announcementEmail_(a, s);
      sendMail_({ to: s.email, subject: a.subject, htmlBody: m.html, body: m.text, name: from, inlineImages: logos });
      markSent_('announce', aid, s.email);
      log.push([stamp, 'Announcement', s.name, s.email, 'Sent']); room--;
      sample = sample || { subject: a.subject, text: m.text };
    } catch (e) { log.push([stamp, 'Announcement', s.name, s.email, 'Not sent: ' + e.message]); }
  });
  if (waiting.length) {
    const kept = queueEmails_({ kind: 'announce', schools: waiting.map(s => s.name), a: a });
    waiting.forEach(s => log.push([stamp, 'Announcement', s.name, s.email, kept ? WAIT_NOTE : 'Not sent: daily email limit reached']));
  }
  if (!onlyNames) t.skipped.forEach(s => log.push([stamp, 'Announcement', s.name, s.contact || '(blank)', 'Skipped: no email in Scholars']));
  writeLog_(ss, log);
  try { reportToMe_('Announcement', { monday: stamp }, log, sample, logos); } catch (e) { Logger.log('Report not sent: ' + e.message); }
  return { sent: log.filter(r => r[4] === 'Sent').length, waiting: waiting.length, done };
}

/* ================================================================
 *  🩺 CHECK MY TRACKER
 *  Each check: [✅ ok | ⚠️ worth fixing | ❌ will break something, what was found, how to fix]
 * ================================================================ */

// The code files (Apps Script) and what is inside each. Keep these names.
var FILE_GROUPS = [
  ['00_Config',     ['00_Config', '02_Settings', '09_Helpers'],                   'privacyNote_'],
  ['01_Region',     ['05_Region'],                                                'applyRegion_'],
  ['02_Menu_Setup', ['01_Menu', '03_StartHere', '04_Setup'],                      'setupSteps_'],
  ['03_Forms',      ['06_Form', '23_Members'],                                    'setupMemberCheckin'],
  ['04_Emails',     ['07_WeeklyEmail', '10_Reminders', '20_Safety'],              'memberWeeklyTargets_'],
  ['05_Portals',    ['08_FocalPortal'],                                           'applyRoles_'],
  ['06_Calendar',   ['22_Calendar'],                                              'setupCalendars'],
  ['07_Look',       ['11_Logos', '12_Style', '13_Dashboard', '17_Tutorials', '19_Banners'], 'emailPics_'],
  ['08_Tools',      ['14_Links', '15_Outreach', '16_Template', '18_Tools', '21_Reports', '24_Extras'], 'sendMemberAnnouncementTo_'],
];
// One function that only exists in the CURRENT version of each file. Missing = that file is old or empty.
var FILE_MARKERS = FILE_GROUPS.map(g => [g[0], g[2]]);
// Older messages name the old part ("02_Settings"): say which file it is in now ("00_Config")
function fileNow_(text) {
  let t = String(text || '');
  FILE_GROUPS.forEach(g => g[1].forEach(old => { if (old !== g[0]) t = t.split(old).join(g[0]); }));
  return t;
}

function healthCheck() { return showHealth_(healthRows_()); }

// Every check as rows [icon, what, fix] (also used by the Monday self-check email)
function healthRows_() {
  const rows = [];
  // one-file version: "replace file X" becomes "delete the old files" (an old file left in Apps Script is the usual cause)
  const both = typeof START_ONE_FILE !== 'undefined' && typeof START_MODULES !== 'undefined';   // old one-file left next to the new files
  const oneFile = typeof START_ONE_FILE !== 'undefined' && !both;
  const fixText = f => oneFile && /^Replace |^Add an HTML|Replace these files/.test(String(f || ''))
    ? 'In Apps Script keep ONLY 2 files: START_Tracker and Email_Reminder. Delete every other file, paste the newest START_Tracker, Save.' : f;
  const ok = (t) => rows.push(['✅', t, '']), warn = (t, fix) => rows.push(['⚠️', fileNow_(t), fileNow_(fixText(fix))]),
        bad = (t, fix) => rows.push(['❌', fileNow_(t), fileNow_(fixText(fix))]);
  // looks the name up in the shared (global) scope that every .gs file uses
  const fnExists = n => { try { return new Function('return typeof ' + n)() === 'function'; } catch (e) { return false; } };

  // 1. code files
  if (both)
    bad('Two versions of the code are in Apps Script (START_Tracker AND the separate files)',
        'In Apps Script delete the file START_Tracker (⋮ › Delete). Keep 00_Config … 08_Tools and Email_Reminder.');
  const missing = FILE_MARKERS.filter(f => !fnExists(f[1])).map(f => f[0]);
  if (missing.length) bad('Old or empty code file(s): ' + missing.join(', '), 'Replace these files in Apps Script with the latest versions, then save.');
  else if (libMode_()) ok(`Code comes from the shared START library (version ${typeof START_MODULES === 'undefined' ? '?' : START_MODULES})`);
  else ok(oneFile ? `Code is up to date (one file, ${START_ONE_FILE})` : `All ${FILE_MARKERS.length} code files are the current version`);
  if (libMode_() && stubVersion_() < STUB_NEEDED)
    bad('The small "START" file in Apps Script is an older version (some menu items or automatic steps may not work)',
        'In Apps Script open the file START, select all, paste the newest START file from the library owner, Save.');
  try {
    const r = regionProfile_();
    if (r.code) ok(`Region: ${r.label} (${r.name})` + (r.schools.length ? `, ready-made school list available` : ''));
    else warn(r.label ? `Region "${r.label}" is not on the list of Philippine regions` : 'No region chosen yet',
              'Settings tab › Region: pick your region from the drop-down.');
  } catch (e) { /* 05_Region missing: reported above */ }
  if (typeof SETTINGS !== 'undefined' && !SETTINGS.some(s => s[0] === 'autoWeekly')) bad('02_Settings is an old version (no 🚀 row)', 'Replace 02_Settings.');
  // a few files changed without getting a new function: look inside them for the newest code
  const has = (fn, text) => { try { return String(new Function('return ' + fn)()).indexOf(text) !== -1; } catch (e) { return false; } };
  if (fnExists('applySchedule_') && !has('applySchedule_', 'autoWeekly')) bad('06_Form is an old version', 'Replace 06_Form.');
  if (fnExists('onSheetChange') && !has('onSheetChange', 'getActiveSheet')) bad('06_Form is an old version (Scholars email changes are missed)', 'Replace 06_Form.');
  if (typeof PICTURES === 'undefined' || !fnExists('emailLogoSrc_')) bad('11_Logos is an old version (the big one with pictures inside)', 'Replace 11_Logos with the new small one.');
  else {
    if (fnExists('placeTabBanner_') && has('placeTabBanner_', 'TAB_BANNER_B64')) bad('19_Banners is the old big version', 'Replace 19_Banners with the new small one.');
    [['buildStartHere_', '03_StartHere'], ['buildDashboard_', '13_Dashboard'], ['writePortal_', '08_FocalPortal']].forEach(([fn, file]) => {
      if (fnExists(fn) && has(fn, '_BANNER_B64')) bad(file + ' is an old version (made for the old big 11_Logos)', 'Replace ' + file + '.');
    });
    const miss = missingPictures_();
    if (miss.length) warn(`${miss.length} of ${Object.keys(PICTURES).length} tracker pictures not found: ` + miss.map(k => PICTURES[k] || k).slice(0, 8).join(', ') + (miss.length > 8 ? ' …' : ''),
      '🛠 Setup tools › 🖼 Connect the pictures folder (paste the folder link; subfolders are searched too) or 🖼 Add the tracker pictures. Then click ①. ' +
      'The folder must be in, or shared with, the account that runs the tracker (' + (me_() || 'you') + ').');
    else if (emailLogoSrc_() === LOGO_STRIP_CID) bad('Email logos are attached, not linked (your Google account did not allow the public logo picture), so Gmail may hide them', 'Emails still work. Ask your IT office to allow "anyone with the link" files if you want logos to always show.');
    else ok('All tracker pictures are in your Drive');
    try {
      const memes = typeof memePicFiles_ === 'function' ? memePicFiles_().filter(x => /^(lead-)?(vibe|events|learning|help)-/i.test(x.getName())).length : 0;
      if (!memes) warn('No meme pictures found for the forms', '🛠 Setup tools › 🖼 Connect the pictures folder and paste the link of the folder with vibe-1.jpg, help-1.jpg … (a subfolder is fine).');
      else ok(`Form pictures: ${memes} memes found`);
    } catch (e) { /* no meme folder */ }
  }
  if (fnExists('renderEmail_') && !has('renderEmail_', 'emailLogoSrc_')) bad('10_Reminders is an old version (email logos)', 'Replace 10_Reminders.');
  if (fnExists('setupCalendars') && typeof EV_EDIT === 'undefined') bad('22_Calendar is an old version (Events tab, member check-ins)', 'Replace 22_Calendar.');
  if (typeof Q !== 'undefined' && Q.events === 'Events joined this week') bad('00_Config is an old version (old Campus Lead questions)', 'Replace 00_Config, then click ①.');
  if ((typeof TAB !== 'undefined' && !TAB.mcheckins) || (typeof OLD_CONFIG_ !== 'undefined' && OLD_CONFIG_)) bad('00_Config is an old version (member check-ins would be mixed with Lead check-ins)', 'Replace 00_Config, then click ①.');
  if (fnExists('makeCleanTemplate') && !has('makeCleanTemplate', 'Member check-ins')) bad('16_Template is an old version (the clean copy could keep member answers)', 'Replace 16_Template.');
  if (typeof SETTINGS !== 'undefined' && !SETTINGS.some(s => s[0] === 'replyTo')) bad('02_Settings is an old version (no reply-to row)', 'Replace 02_Settings, then click ①.');
  if (fnExists('buildForm_') && !has('buildForm_', 'CONSENT_TITLE')) bad('04_Setup is an old version (no privacy consent in the form)', 'Replace 04_Setup, then click ①.');
  if (fnExists('setup') && !has('setup', 'SETUP_KEY')) bad('04_Setup is an old version (① can run out of time)', 'Replace 04_Setup.');
  if (fnExists('buildScholars_') && !has('buildScholars_', 'ROLES')) bad('04_Setup is an old version (Role is not a dropdown)', 'Replace 04_Setup, then click ①.');
  if (fnExists('syncContact_')) bad('08_FocalPortal is an old version (emails are still copied)', 'Replace 08_FocalPortal.');
  if (fnExists('buildSchools_') && !has('buildSchools_', 'setSchoolEmailFormula_')) bad('04_Setup is an old version (emails not read from Scholars)', 'Replace 04_Setup, then click ①.');
  if (fnExists('columnSpecs_') && has('columnSpecs_', "[10, 'auto', 'Role'")) bad('12_Style is an old version (Role column)', 'Replace 12_Style, then click ①.');
  try {
    const raw = HtmlService.createTemplateFromFile('Email_Reminder').getRawContent();
    if (raw.indexOf('extraHtml') === -1 || raw.indexOf('title1') === -1 || raw.indexOf('logoW') === -1 || raw.indexOf('if (logo)') === -1) bad('Email_Reminder.html is an old version (logos break in Gmail, not phone-friendly)', 'Replace Email_Reminder.html.');
    else ok('Email design file is current');
  } catch (e) { bad('Email_Reminder.html is missing', 'Add an HTML file named Email_Reminder and paste the latest version.'); }

  const ss = SpreadsheetApp.getActive();
  const props = docProps_();
  // 2. setup, tabs and form
  if (!props.getProperty(WEB_CHECKIN ? 'SETUP_DONE' : 'FORM_ID')) { bad('The tracker is not set up yet', 'Click ① Set up / fix the tracker.'); return rows; }
  if (WEB_CHECKIN) {
    if (ss.getSheetByName(TAB.mcheckins) && ss.getSheetByName(TAB.checkins)) ok('Check-in tabs are there'); else bad('A check-in tab is missing', 'Click ① Set up / fix the tracker.');
    webVersionRow_(ok, warn, bad);
  } else { try { getForm_(); ok('Check-in form is connected'); } catch (e) { bad('The check-in form cannot be opened', 'Click ① Set up / fix the tracker.'); } }
  const tabs = (DB_MODE ? [] : [TAB.start, TAB.dashboard, TAB.links]).concat([TAB.schools, TAB.scholars, TAB.checkins, TAB.settings, TAB.outreach]);
  const noTab = tabs.filter(t => !ss.getSheetByName(t));
  if (noTab.length) bad('Missing tab(s): ' + noTab.join(', '), 'Click ① Set up / fix the tracker (tabs are rebuilt, data is kept).');
  else ok('All main tabs are there');
  // email budget: can one day's limit cover the Friday check-in emails?
  try {
    const cap = dailyMailCap_(), keep = MAIL_RESERVE + (WEB_CHECKIN ? WEB.mailReserve : 0), need = welcomeList_(ss).length;
    if (need > cap - keep) warn(`About ${need} scholars get the weekly email, but this account can send about ${cap - keep} a day (Google's limit is ${cap}, ${keep} kept for sign-in codes and urgent alerts)`,
      'The rest go out the next morning by themselves. To send them all on one day, run the tracker from a Google Workspace account (about 1,500 a day). The website and app always show the check-in.');
    else ok(`Email budget is fine (${need} scholars, about ${cap - keep} emails a day)`);
  } catch (e) { /* not countable now */ }

  // 3. automatic updates
  const handlers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  if (handlers.indexOf('onSheetChange') === -1) bad('Automatic sync is off (form list, Campus Leads)', 'Click ① Set up / fix the tracker.');
  else ok('Automatic sync is on');
  if (typeof WEB_CHECKIN !== 'undefined' && WEB_CHECKIN) {           // website check-in: saved and processed the moment it is sent
    if (handlers.indexOf('catchUpAnswers') === -1) warn('The hourly safety net for missed answers is off', 'Click ① Set up / fix the tracker.');
    else ok('New check-ins are processed as soon as scholars submit them on the website');
  } else if (handlers.indexOf('onCheckin') === -1) {
    if (automationOwner_() && automationOwner_() !== me_())
      warn(`New check-ins are processed by ${automationOwner_()}, not by you`, 'Fine if that account is still signed in to Google. If the portals stop updating, click ⏰ Automatic emails › 🔑 Make my account send the automatic emails.');
    else bad('New check-ins are not being processed (portals and the Campus Leads tab will not update)', 'Click ① Set up / fix the tracker.');
  } else if (handlers.indexOf('catchUpAnswers') === -1) warn('The hourly safety net for missed answers is off', 'Click ① Set up / fix the tracker.');
  else ok('New check-ins are processed automatically');

  // 4. settings and weekly emails
  const st = getSettings_();
  if (!String(st.rrName).trim()) warn('Your name is empty in Settings', 'Type it in Settings (it signs every email).');
  const on = handlers.indexOf('weeklySummary') !== -1 && handlers.indexOf('autoReminder') !== -1;
  if (st.autoWeekly === true && !on) bad('🚀 is ticked but the weekly emails are not scheduled', 'Click ⏰ Automatic emails › 🚀 Turn ON automatic weekly emails.');
  else if (st.autoWeekly === true) ok(`Weekly emails ON: reminders ${st.remDay} ${HOURS[st.remHour24]}, summary ${st.day} ${HOURS[st.hour24]}`);
  else warn('Weekly emails are OFF', 'Click ⏰ Automatic emails › 🚀 Turn ON when you are ready.');
  const quota = MailApp.getRemainingDailyQuota(), need = schoolRows_(ss).filter(s => s.email).length;
  if (quota - MAIL_RESERVE < need) warn(`${quota} emails left today (Google limit); a full reminder needs ${need}`,
    'Nothing to do: extra emails wait and go out tomorrow morning by themselves.');
  else ok(`${quota} emails left today (Google's daily limit)`);
  const owner = automationOwner_();
  if (!owner) warn('No email account chosen yet', 'Click ⏰ Automatic emails › 🚀 (or ① Set up) from the account that should send the automatic emails.');
  else if (owner === me_()) ok(`Automatic emails are sent from your account (${owner})`);
  else warn(`Automatic emails are sent from ${owner}, not from you`,
    'That is fine. Emails you send from the menu use YOUR daily limit. To move them to your account: ⏰ Automatic emails › 🔑 Make my account send the automatic emails.');
  if (fnExists('calOn_') && calOn_()) {
    if (!calServiceOn_()) bad('Calendar service is off (the calendars cannot update)', 'Apps Script › Services › + › Google Calendar API › Add. Then ' + MENU_CAL + '.');
    else {
      let L = {};
      try { L = JSON.parse(docProps_().getProperty('CAL_LAST') || '{}'); } catch (e) { L = {}; }
      if (L.blocked) bad('Your Google account blocks sharing calendars outside your school', 'Make a personal Gmail the email account (⏰ Automatic emails › 🔑), then ' + MENU_CAL + '.');
      else if (L.failed) warn(`Calendar could not be shared with ${L.failed} person(s)`, 'Check their emails in Scholars › Contact (E). It tries again every morning.');
      else ok('Shared calendars are ON and up to date' + (L.waiting ? ` (${L.waiting} school calendar(s) still being made, over the next days)` : ''));
    }
  }
  if (fnExists('memberForm_') && docProps_().getProperty('MFORM_ID')) {
    if (!memberForm_()) bad('The member check-in form is missing', '🛠 Setup tools › 📝 Set up / update member check-in.');
    else if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'memberDaily') && automationOwner_() === me_())
      warn('The member form does not update by itself', '🛠 Setup tools › 📝 Set up / update member check-in.');
    else ok('Member check-in is ON (updates every morning)');
  }
  const wait = waitingCount_();
  if (wait) warn(`${wait} email(s) waiting for tomorrow morning (daily limit)`, 'They send by themselves. Listed in the Reminder log as "Waiting".');

  // 5. data
  const schools = schoolRows_(ss);
  if (!schools.length) warn('No schools yet', 'Add them in the Schools tab (or 🛠 Setup tools › 🏫 Fill in my region\'s schools).');
  const noLead = schools.filter(s => !s.focal), noMail = schools.filter(s => s.focal && !s.email);
  if (noLead.length) warn(`${noLead.length} school(s) without a START Campus Lead`, noLead.map(s => s.name).join(', '));
  if (noMail.length) warn(`${noMail.length} Campus Lead(s) without an email (no reminders, no portal)`,
    'Type their email in Scholars › Contact (email): ' + noMail.map(s => displayName_(s.focal) + ' (' + s.name + ')').join(', '));
  if (schools.length && !noLead.length && !noMail.length) ok('Every school has a Campus Lead with an email');
  const names = schools.map(s => s.name);
  const odd = [...new Set(scholarRows_(ss).map(r => String(r[1]).trim()).filter(n => n && names.indexOf(n) === -1))];
  if (odd.length) warn('Scholars with a school name not in the Schools tab: ' + odd.join(', '), 'Fix the spelling in Scholars, column B (it must match exactly).');
  else ok('Every scholar\'s school matches the Schools tab');

  // 6. privacy: the tracker and every portal
  try {
    const probs = privacyProblems_();
    probs.forEach(p => (p.level === 'bad' ? bad : warn)(p.text, p.fix));
    if (!probs.length) ok('Private: only you can open the tracker, and each portal only by its own Campus Lead');
  } catch (e) { warn('Could not check sharing', 'Open Share and make sure General access is Restricted.'); }

  // ✨ extras
  try {
    if (automationOwner_() === me_() && !ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'extrasDaily'))
      warn('Daily extras are not scheduled (open/close, backups, monthly report, self-check)', 'Click ① Set up / fix the tracker.');
    const waiting = unrecognisedCount_(SpreadsheetApp.getActive());
    if (waiting) warn(`${waiting} answer(s) are waiting in "Not recognised" (their email is not in Scholars)`, 'Open that tab and add each email to Scholars › Contact (E): the answers then count by themselves.');
    const badEmails = checkScholarEmails_(SpreadsheetApp.getActive());
    if (badEmails.length) warn(`${badEmails.length} email(s) in Scholars look wrong (orange cells)`, 'Scholars › Contact: ' + badEmails.slice(0, 3).map(x => x.name + ': ' + x.why).join(' · '));
  } catch (e) { /* extras file missing: reported above */ }
  return rows;
}

function showHealth_(rows) {
  const n = { '❌': 0, '⚠️': 0 };
  rows.forEach(r => { if (n[r[0]] !== undefined) n[r[0]]++; });
  const head = n['❌'] ? `${n['❌']} problem(s) to fix` : n['⚠️'] ? `All working · ${n['⚠️']} thing(s) worth a look` : 'Everything looks good 🎉';
  const order = { '❌': 0, '⚠️': 1, '✅': 2 };
  const tone = { '❌': 'background:#FADCD9;border-color:#F2B8B2', '⚠️': 'background:#FFF6BF;border-color:#F5E27A', '✅': '' };
  const bad = rows.filter(r => r[0] !== '✅'), good = rows.filter(r => r[0] === '✅');
  const html = `<p class="lead">${esc_(head)}</p>` +
    bad.sort((a, b) => order[a[0]] - order[b[0]]).map(r =>
      `<div class="box" style="${tone[r[0]]}">${r[0]} <b>${esc_(r[1])}</b>${r[2] ? `<div class="muted" style="margin-top:3px">→ ${esc_(r[2])}</div>` : ''}</div>`).join('') +
    (good.length ? `<details${bad.length ? '' : ' open'}><summary class="muted" style="cursor:pointer;margin:6px 0">✅ ${good.length} working</summary>` +
      `<ul class="dots muted">${good.map(r => `<li>${esc_(r[1])}</li>`).join('')}</ul></details>` : '');
  dlg_('🩺 Check my tracker', html, 520, Math.min(600, 160 + bad.length * 70 + (bad.length ? 30 : good.length * 20)));
  return rows;
}

/* ================================================================
 *  🧪 SEND ME A TEST OF EVERY EMAIL
 *  Every email the tracker can send, built by the real code with real data, all sent to YOU only,
 *  marked [TEST 1] … [TEST 11]. Nothing goes to Campus Leads or members, nothing is logged or marked done.
 * ================================================================ */
function sendAllTests() {
  requireSetup_();
  const ss = SpreadsheetApp.getActive(), st = getSettings_(), me = Session.getEffectiveUser().getEmail();
  const N = 17;
  if (MailApp.getRemainingDailyQuota() < N + MAIL_RESERVE) {
    return notify_(`Google's daily email limit has only ${MailApp.getRemainingDailyQuota()} left today. Please try again tomorrow.`);
  }
  const ui = ui_();
  if (ui.alert('🧪 Send me a test of every email?',
      `${N} test emails will be sent ONLY to you (${me}), marked [TEST 1] … [TEST ${N}].\n\n` +
      'Nobody else receives anything, and nothing is marked as sent.\n\nContinue?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const schools = schoolRows_(ss), store = portalStore_();
  const s = schools.find(x => x.focal) || schools[0] || { name: 'Sample School', focal: 'Dela Cruz, Juan', email: me, link: '' };
  const school = Object.assign({ link: s.link || checkinUrl_() || ss.getUrl() }, s);
  school.portal = portalUrl_(store, school);
  const rr = { name: st.rrName || 'your Regional Representative', email: me, region: st.region || '' };
  const done = [], failed = [];
  const run = (label, fn) => { try { fn(); done.push(label); } catch (e) { failed.push(label + ': ' + e.message); } };
  TEST_MODE_ = me; TEST_COUNT_ = 0;
  try {
    run('Weekly check-in reminder (to Campus Leads, Friday)', () => {
      const m = composeEmail_('reminder', school, reminderWeek_(new Date()));
      sendMail_({ to: me, subject: m.subject, htmlBody: m.html, body: m.text, name: m.from });
    });
    run('Follow-up (to Leads who have not answered, Saturday)', () => {
      const sat = new Date(reminderWeek_(new Date()).deadline); sat.setDate(sat.getDate() - 1);   // as it looks on follow-up day
      const m = composeEmail_('followup', school, followUpWeek_(sat));
      sendMail_({ to: me, subject: m.subject, htmlBody: m.html, body: m.text, name: m.from });
    });
    run('Copy report (to you, after each send)', () => reportToMe_('Reminder', reminderWeek_(new Date()),
      [[new Date(), 'Reminder', school.name, school.email || '(sample)', 'Sent']],
      composeEmail_('reminder', school, reminderWeek_(new Date())), null));
    if (!DB_MODE) run('Portal welcome (to a new Campus Lead)', () => {
      const p = store.map[school.name];
      sendPortalWelcome_(Object.assign({}, school, { email: school.email || me }), p && p.id ? p.id : ss.getId(), rr);
    });
    run('Weekly summary (to you, Monday)' + (st.missedAlert === true ? ' + quiet schools list' : ''), () => weeklySummary());
    if (st.missedAlert !== true) run('Quiet schools 2+ weeks (to you)', () => sendMissedAlert_(false));   // else it comes with the summary
    run('Urgent alert from a Campus Lead check-in (to you)', () =>
      leadUrgentAlert_(ss, school.name, displayName_(school.focal) || 'Sample Lead', HELP_CHOICES[2], 'This is a test of the urgent alert.', false));
    run('Urgent alert from a member (to the Lead and you)', () =>
      memberUrgentAlert_(ss, { name: 'Sample Member', school: school.name }, 'This is a test of the member alert.', M_VIBES[3]));
    const mp = { name: 'Sample Member', email: me, school: school.name, isLead: false, lead: school.focal || '' };
    run('📨 Welcome email (to a member, from 👥 › 📨)', () =>
      sendMail_({ to: me, subject: `START · Welcome, ${school.name}`, htmlBody: welcomeHtml_(mp, {}), body: 'Welcome to START' }));
    run('📨 Welcome email (to a Campus Lead)', () =>
      sendMail_({ to: me, subject: `START · Welcome, ${school.name}`, htmlBody: welcomeHtml_(Object.assign({}, mp, { isLead: true, portal: school.portal || '' }), {}), body: 'Welcome' }));
    run('📅 "Your START calendar is ready" (to members welcomed before the calendar)', () =>
      sendMail_({ to: me, subject: 'START · Your calendar is ready', htmlBody: welcomeHtml_(mp, { calendarOnly: true }), body: 'Calendar' }));
    run('Weekly check-in (to welcomed members, Friday)', () => {
      const m = composeMemberEmail_('reminder', mp, reminderWeek_(new Date()));
      sendMail_({ to: me, subject: m.subject, htmlBody: m.html, body: m.text });
    });
    run('Follow-up (to members who have not answered, Saturday)', () => {
      const m = composeMemberEmail_('followup', mp, reminderWeek_(new Date()));
      sendMail_({ to: me, subject: m.subject, htmlBody: m.html, body: m.text });
    });
    run('Event tomorrow (to Campus Leads, the day before)', () => {
      const t = new Date(); t.setDate(t.getDate() + 1);
      sendMail_({ to: me, subject: 'START · Tomorrow: Sample START event', body: 'Tomorrow',
        htmlBody: eventTomorrowHtml_(school.focal, [['Sample START event', Utilities.formatDate(t, CONFIG.timezone, 'EEE, MMM d'), 'Google Meet']], 'https://calendar.google.com') });
    });
    run('🩺 Self-check (to you, only when something is wrong)', () =>
      sendMail_({ to: me, subject: 'START · Tracker check: 1 thing(s) to fix', body: 'Self-check',
        htmlBody: designedEmail_({ badge: 'SELF-CHECK', badgeColor: '#B3261E', title: ['Tracker', 'Check'], greeting: '', signature: [],
          intro: ['Something in the tracker needs you. (Sample.)'],
          extraHtml: emailTable_(['', 'What', 'How to fix'], [['❌', 'Sample problem', 'Click ① Set up / fix the tracker']]),
          link: ss.getUrl(), buttonText: 'Open the tracker' }) }));
    run('📣 Announcement (to all Campus Leads)', () => {
      const a = cleanAnnouncement_({ subject: 'Sample announcement', body: 'Good day! This is how your announcements look.\n\nSee you at the next START meeting.' });
      const m = announcementEmail_(a, school);
      sendMail_({ to: me, subject: a.subject, htmlBody: m.html, body: m.text, name: 'START Tracker' });
    });
  } finally { TEST_MODE_ = null; }
  notify_(`🧪 ${TEST_COUNT_} test emails sent to ${me}\n\n` + done.map(d => '• ' + d).join('\n') +
    (failed.length ? '\n\nCould not build:\n' + failed.map(f => '• ⚠️ ' + f).join('\n') : ''));
}


// 🩺 Is the website running the code pasted here? (forgetting Deploy › Manage deployments › New version is easy)
function webVersionRow_(ok, warn, bad) {
  const url = webSiteUrl_();
  if (url && !docProps_().getProperty('WEB_URL_SET')) warn('The website address is not saved yet', '🌐 Website › paste the Web app URL from Deploy › Manage deployments, then Save (Google sometimes reports an old address).');
  if (!url) { warn('The website is not deployed yet', 'Deploy › New deployment › Web app (Execute as: Me · Who has access: Anyone).'); return; }
  try {
    const r = UrlFetchApp.fetch(url + (url.indexOf('?') === -1 ? '?' : '&') + 'ping=1', { muteHttpExceptions: true, followRedirects: true });
    let v = ''; try { v = JSON.parse(r.getContentText()).v || ''; } catch (e) { v = ''; }
    if (v === CODE_VERSION) ok('The website runs this code (' + CODE_VERSION + ')');
    else bad('The website runs older code' + (v ? ' (' + v + ')' : ''), 'Apps Script › Deploy › Manage deployments › ✏️ › Version: New version › Deploy.');
  } catch (e) { warn('Could not reach the website to check its version', 'Open the website once, then check again.'); }
}

/* ---------- 🗑 Data Privacy Act request: remove one person's data everywhere ----------
 * Scholars row, website account and sign-ins, their weekly answers (Member check-ins, Not recognised),
 * their Campus Lead check-ins and Outreach notes under their name. Shows the counts first and asks. */
function removePersonData() {
  requireSetup_();
  const ui = SpreadsheetApp.getUi(), ss = SpreadsheetApp.getActive();
  const r = ui.prompt('Remove a person\'s data', 'Their email (as in Scholars › Contact):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const email = emailOf_(r.getResponseText());
  if (!email) { notify_('That is not an email address.'); return; }
  const found = personRows_(ss, email);
  const total = found.reduce((n, f) => n + f.rows.length, 0) + (found.acct ? 1 : 0);
  if (!total) { notify_(`Nothing found for ${email}.`); return; }
  const list = found.filter(f => f.rows.length).map(f => `• ${f.tab}: ${f.rows.length} row${f.rows.length === 1 ? '' : 's'}`)
    .concat(found.acct ? ['• Website account and sign-ins'] : []);
  if (ui.alert(`Remove ${found.name || email}?`, list.join('\n') + '\n\nThis cannot be undone. Make a copy first if DOST-SEI needs a record.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  found.forEach(f => { const sh = ss.getSheetByName(f.tab); f.rows.slice().sort((a, b) => b - a).forEach(n => sh.deleteRow(n)); });
  try { docProps_().deleteProperty(webAcctKey_(email)); webEndSessions_(email); webForget_(email); } catch (e) { /* no website */ }
  notify_(`🗑 Removed\n\n${list.join('\n')}\n\nTheir calendar access ends with the next morning update.`);
}
// Where one person appears: [{ tab, rows: [row numbers] }] plus .name and .acct
function personRows_(ss, email) {
  const out = [], by = scholarByEmail_(ss), name = by[email] ? by[email].name : '';
  const scan = (tab, test) => {
    const sh = ss.getSheetByName(tab);
    if (!sh || sh.getLastRow() < 2) return;
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues(), rows = [];
    vals.forEach((v, i) => { if (test(v, sh)) rows.push(i + 2); });
    if (rows.length) out.push({ tab, rows });
  };
  const hasEmail = v => v.some(x => emailOf_(x) === email);
  scan(TAB.scholars, v => emailOf_(v[4]) === email);
  scan(TAB.mcheckins, hasEmail);
  scan(UNKNOWN_TAB, hasEmail);
  if (name) {
    const same = x => String(x).trim().toLowerCase() === name.toLowerCase();
    scan(TAB.checkins, v => v.some(same));
    scan(TAB.outreach, v => same(v[2]));
  }
  out.name = name;
  try { out.acct = !!webAccount_(email); } catch (e) { out.acct = false; }
  return out;
}


// ━━━━━━━━━━━━━━━━━━━━  Reports  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  📄 REPORTS
 *   1. Quiet schools alert   schools that missed 2+ weekly check-ins in a row, with a
 *                            ready-to-copy Messenger message for each START Campus Lead.
 *                            Sent with the weekly summary (Settings › "Alert me…"). It also lists them in 🗓 Today.
 *   2. End-of-term report    one click makes a Google Doc (private, in your Drive) for DOST-SEI:
 *                            targets, every school's check-ins, outreach, urgent requests.
 * ================================================================ */

var MISSED_MIN = 2;          // weeks in a row without a check-in before a school is listed
var MISSED_LOOKBACK = 12;    // never counts back more than this many weeks

// [{ name, focal, email, link, missed, last }] for schools with a Campus Lead that missed MISSED_MIN+ weeks in a row.
// Only weeks that are over count (this week still has time), and never weeks before tracking started.
function missedSchools_(ss) {
  const data = focalData_(ss);
  const thisMonday = mondayOf_(new Date());
  let first = null;
  data.forEach(s => s.checkins.forEach(c => { if (!first || c.date < first) first = c.date; }));
  if (!first) return [];                                   // no check-ins yet: nothing to compare with
  let floor = mondayOf_(first);
  const ts = termStart_(ss);
  if (ts instanceof Date && mondayOf_(ts) > floor) floor = mondayOf_(ts);
  const general = checkinUrl_();

  const out = [];
  data.filter(s => s.focal).forEach(s => {
    const weeks = {};
    s.checkins.forEach(c => { weeks[mondayOf_(c.date).getTime()] = true; });
    let w = new Date(thisMonday), missed = 0;
    w.setDate(w.getDate() - 7);
    while (missed < MISSED_LOOKBACK && w >= floor && !weeks[w.getTime()]) { missed++; w.setDate(w.getDate() - 7); }
    if (missed >= MISSED_MIN) out.push({ name: s.name, focal: s.focal, email: s.email, link: s.link || general, missed, last: s.last });
  });
  return out.sort((a, b) => b.missed - a.missed);
}

// The friendly nudge for Messenger / group chat
function missedMessage_(s, st) {
  const who = displayName_(s.focal) || 'there';
  const me = st.rrName ? `${st.rrName}, your START Regional Representative${st.region ? ' (' + st.region + ')' : ''}` : 'your START Regional Representative';
  return `Hi ${who}! This is ${me}. We haven't received ${s.name}'s Campus Lead check-in for the past ${s.missed} weeks. ` +
         `Could you please fill it in this week? It only takes about 2 minutes: ${s.link}\n\n` +
         'If anything is keeping you busy or you need help from START, just reply here. Thank you so much!';
}

// Email to you listing quiet schools. quiet = return without sending when there is nothing to report.
function sendMissedAlert_(quiet) {
  const ss = SpreadsheetApp.getActive(), st = getSettings_();
  let list = missedSchools_(ss);
  if (!list.length && TEST_MODE_) {                                // 🧪 test: show what it looks like with a sample school
    const s = schoolRows_(ss).find(x => x.focal) || { name: 'Sample School', focal: 'Dela Cruz, Juan', link: '' };
    list = [{ name: s.name, focal: s.focal, email: s.email, link: s.link, missed: 2, last: null }];
  }
  if (!list.length) return 0;
  const fmt = d => d ? Utilities.formatDate(d, CONFIG.timezone, 'MMM d, yyyy') : 'never';
  const box = t => `<div style="border:1px dashed #9AA7BA;border-radius:10px;padding:12px 14px;margin:6px 0 14px 0;` +
                   `background:#F6F9FD;font-size:13px;line-height:1.6;color:#33415c;">${esc_(t).replace(/\n/g, '<br>')}</div>`;
  const extra = emailTable_(['School', 'START Campus Lead', 'Weeks missed', 'Last check-in'],
                            list.map(s => [s.name, displayName_(s.focal), s.missed, fmt(s.last)])) +
    emailSection_('Ready-to-copy messages (Messenger or group chat)') +
    list.map(s => `<div style="font-size:12px;font-weight:800;color:#12203A;margin-top:10px;">${esc_(s.name)}</div>` +
                  box(missedMessage_(s, st))).join('');
  const html = designedEmail_({
    badge: 'Needs a nudge', badgeColor: '#C2410C', title: ['Quiet', 'Schools'], greeting: '',
    intro: [`${list.length} school(s) missed ${MISSED_MIN}+ weeks in a row. A personal message works best: ready-to-copy ones are below.`],
    extraHtml: extra,
    link: ss.getUrl(), buttonText: 'Open the tracker', buttonNote: '',
    chips: [['QUIET SCHOOLS', list.length], ['LONGEST', list[0].missed + ' weeks'], ['CHECKED', fmt(new Date())]],
    outro: ['To stop these: untick "Alert me: schools quiet 2+ weeks" in Settings.'],
    signature: [],
  });
  if (!RR_EMAILS && !TEST_MODE_) return list.length;                         // quiet schools are on the RR's START page instead
  const to = String(st.recipients || '').trim() || Session.getEffectiveUser().getEmail();
  sendMail_({ to, name: 'START Tracker', htmlBody: html,
              subject: `START · ${list.length} school(s) quiet for ${MISSED_MIN}+ weeks${st.region ? ' (' + st.region + ')' : ''}` });
  return list.length;
}

// Send the list now (kept for running by hand from the editor)
function sendMissedAlertNow() {
  requireSetup_();
  const n = sendMissedAlert_(false);
  notify_(n ? `📬 Sent: ${n} quiet school(s), each with a ready-to-copy Messenger message.`
            : `🎉 No school has missed ${MISSED_MIN} weeks in a row. Nothing to send.`);
}

/* ---------- 📄 End-of-term report (Google Doc) ---------- */

// opts (optional, from the monthly report): { from, to, label, silent } -> returns the Doc's link
function makeTermReport(opts) {
  opts = opts && opts.from ? opts : null;
  if (!opts) requireSetup_();
  const ss = SpreadsheetApp.getActive(), st = getSettings_();
  const now = opts ? new Date(opts.to.getTime() - 1) : new Date();
  const ts = termStart_(ss);
  const from = opts ? opts.from : ts instanceof Date ? ts : null;
  const fmt = d => Utilities.formatDate(d, CONFIG.timezone, 'MMMM d, yyyy');
  const inTerm = d => d instanceof Date && (!from || d >= from) && (!opts || d < opts.to);
  const data = focalData_(ss);

  // ---- numbers ----
  const withLead = data.filter(s => s.focal);
  const weeksInTerm = from ? Math.max(1, Math.ceil((mondayOf_(now) - mondayOf_(from)) / (7 * 86400000)) + 1) : null;
  const reeng = data.filter(s => s.inactive && s.last && inTerm(s.last));
  const all = [];
  data.forEach(s => s.checkins.filter(c => inTerm(c.date)).forEach(c => all.push(Object.assign({ school: s.name }, c))));
  const urgent = all.filter(c => c.help === HELP_CHOICES[2]).sort((a, b) => a.date - b.date);
  const help = all.filter(c => c.help === HELP_CHOICES[1] || (String(c.concern || '').trim() && c.help !== HELP_CHOICES[2])).sort((a, b) => a.date - b.date);
  const evAns = all.filter(c => String(c.events || '').trim() && isNaN(Number(c.events)));
  const evGood = evAns.filter(c => EVENT_GO_CHOICES.slice(0, 2).indexOf(String(c.events)) !== -1).length;
  const weeksAnswered = s => new Set(s.checkins.filter(c => inTerm(c.date)).map(c => mondayOf_(c.date).getTime())).size;

  const out = ss.getSheetByName(TAB.outreach);
  const contacts = out && out.getLastRow() > 1
    ? out.getRange(2, 1, out.getLastRow() - 1, OUTREACH_HEAD.length).getValues().filter(r => inTerm(r[0]) && String(r[1]).trim()) : [];
  const byChannel = {};
  contacts.forEach(r => { const k = String(r[3] || 'Other'); byChannel[k] = (byChannel[k] || 0) + 1; });
  const named = contacts.filter(r => r[4] === OUTREACH_OUTCOMES[0]).length;

  const log = ss.getSheetByName(TAB.log);
  const sentRows = log && log.getLastRow() > 1 ? log.getRange(2, 1, log.getLastRow() - 1, 5).getValues().filter(r => inTerm(r[0]) && r[4] === 'Sent') : [];
  const sentBy = k => sentRows.filter(r => r[1] === k).length;

  // ---- document (private: in your Drive, next to the portals) ----
  const title = (opts ? `START ${st.region || ''} – ${opts.label}`
                      : `START ${st.region || ''} – End-of-term report (${Utilities.formatDate(now, CONFIG.timezone, 'MMM d, yyyy')})`).replace(/\s+/g, ' ');
  const doc = DocumentApp.create(title);
  try { DriveApp.getFileById(doc.getId()).moveTo(portalFolder_(ss)); } catch (e) { /* stays in My Drive */ }
  const body = doc.getBody();
  body.setMarginTop(48).setMarginBottom(48).setMarginLeft(56).setMarginRight(56);
  const A = DocumentApp.Attribute;
  const style = (el, o) => { const s = {}; s[A.FONT_FAMILY] = 'Poppins'; Object.keys(o || {}).forEach(k => { s[k] = o[k]; }); el.setAttributes(s); return el; };
  const para = (t, o) => style(body.appendParagraph(t), Object.assign({ [A.FONT_SIZE]: 11, [A.FOREGROUND_COLOR]: '#12203A' }, o));
  const heading = t => { const p = body.appendParagraph(t).setHeading(DocumentApp.ParagraphHeading.HEADING2);
                         style(p, { [A.FONT_SIZE]: 14, [A.BOLD]: true, [A.FOREGROUND_COLOR]: '#0A6ED1' }); p.setSpacingBefore(14); return p; };
  const table = (head, rows) => {
    const t = body.appendTable([head].concat(rows.length ? rows : [head.map((h, i) => i ? '' : 'None')]).map(r => r.map(x => String(x === null || x === undefined ? '' : x))));
    t.setBorderColor('#CFE0F3');
    for (let r = 0; r < t.getNumRows(); r++) {
      const row = t.getRow(r);
      for (let c = 0; c < row.getNumCells(); c++) {
        const cell = row.getCell(c);
        cell.setPaddingTop(4).setPaddingBottom(4);
        style(cell, { [A.FONT_SIZE]: 9, [A.FOREGROUND_COLOR]: r ? '#33415c' : '#FFFFFF', [A.BOLD]: !r });
        cell.setBackgroundColor(r ? (r % 2 ? '#FFFFFF' : '#F6F9FD') : '#12203A');
      }
    }
    return t;
  };

  style(body.getParagraphs()[0].setText('Department of Science and Technology – Science Education Institute'),
        { [A.FONT_SIZE]: 9, [A.FOREGROUND_COLOR]: '#5b6b84' });
  para('SCHOLARS TRANSFORMING ADVANCEMENT AND RESEARCH FOR TECHNOLOGY (START – DOST)', { [A.FONT_SIZE]: 9, [A.BOLD]: true });
  const h = body.appendParagraph(opts ? 'MONTHLY REPORT' : 'END-OF-TERM REPORT').setHeading(DocumentApp.ParagraphHeading.TITLE);
  style(h, { [A.FONT_SIZE]: 26, [A.BOLD]: true, [A.FOREGROUND_COLOR]: '#12203A' });
  para(`${st.region || ''}${st.rrName ? ' · Prepared by ' + st.rrName + ', Regional Representative' : ''}`, { [A.BOLD]: true, [A.FOREGROUND_COLOR]: '#0099FF' });
  para(`Period covered: ${from ? fmt(from) : 'start of tracking'} – ${fmt(now)}`, { [A.FONT_SIZE]: 10, [A.FOREGROUND_COLOR]: '#5b6b84' });

  heading('1. Targets');
  const due = d => d instanceof Date ? fmt(d) : '';
  table(['Target', 'Achieved', 'Goal', 'Deadline', 'Status'], [
    ['START Campus Leads named', withLead.length, st.tFocal, due(st.dFocal), withLead.length >= st.tFocal ? 'Met' : 'In progress'],
    ['Schools re-engaged', reeng.length, st.tReeng, due(st.dReeng), reeng.length >= st.tReeng ? 'Met' : 'In progress'],
  ]);
  para('');
  para(`Schools tracked: ${data.length} · With a START Campus Lead: ${withLead.length} · ` +
       `Weekly check-ins received: ${all.length} · Event weeks that went well: ${evGood} of ${evAns.length}` +
       (weeksInTerm ? ` · Weeks in the period: ${weeksInTerm}` : ''), { [A.FONT_SIZE]: 10 });
  if (reeng.length) para('Re-engaged this term: ' + reeng.map(s => s.name).join(', '), { [A.FONT_SIZE]: 10 });

  heading('2. Weekly check-ins per school');
  table(['School', 'START Campus Lead', 'Active scholars', 'Check-ins', weeksInTerm ? 'Weeks answered' : 'Weeks answered (all)', 'Last check-in'],
    data.map(s => [s.name, s.focal ? displayName_(s.focal) : '—', s.active, s.checkins.filter(c => inTerm(c.date)).length,
                   weeksInTerm ? `${weeksAnswered(s)} / ${weeksInTerm} (${Math.round(100 * weeksAnswered(s) / weeksInTerm)}%)` : weeksAnswered(s),
                   s.last ? fmt(s.last) : 'never']));

  heading('3. Outreach');
  para(`Contacts recorded: ${contacts.length} · START Campus Leads named through outreach: ${named}` +
       (Object.keys(byChannel).length ? ' · By channel: ' + Object.keys(byChannel).map(k => `${k} ${byChannel[k]}`).join(', ') : ''), { [A.FONT_SIZE]: 10 });
  para(`Emails sent by the tracker: ${sentBy('Reminder')} reminders, ${sentBy('Follow-up')} follow-ups, ${sentBy('Announcement')} announcements.`, { [A.FONT_SIZE]: 10 });

  heading('4. Urgent requests');
  table(['Date', 'School', 'Concern'], urgent.map(c => [fmt(c.date), c.school, String(c.concern || 'No details given').slice(0, 300)]));

  heading('5. Other concerns and help requests');
  table(['Date', 'School', 'Concern'], help.map(c => [fmt(c.date), c.school, String(c.concern || '(asked for help, no details)').slice(0, 300)]));

  heading('6. Notes and next steps');
  para('(Add your highlights, challenges and plans for next term here.)', { [A.ITALIC]: true, [A.FOREGROUND_COLOR]: '#5b6b84' });
  para('');
  para('Prepared with the START Tracker. Contains scholar names: share only with DOST-SEI / ONSSE (Data Privacy Act of 2012).',
       { [A.FONT_SIZE]: 8, [A.FOREGROUND_COLOR]: '#5b6b84' });
  doc.saveAndClose();

  const url = doc.getUrl();
  if (opts && opts.silent) return url;
  dlgDone_('📄 End-of-term report', { head: 'Your report is ready', text: 'Private, in your Drive. Add your notes in section 6.',
    link: url, button: 'Open the report', note: 'Then File › Download › PDF or Word.' });
}


// ━━━━━━━━━━━━━━━━━━━━  Extras  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  ✨ EXTRAS: the things that keep the tracker healthy by themselves
 *   1. Missing emails      each portal lists members with no email + a ready message (23_Members)
 *   2. Email checks        Scholars › Contact: same email twice, or another school's email domain
 *   3. Open / close        the weekly check-in closes after the deadline and opens again on its day
 *   4. Monthly report      on the 1st: a short Google Doc for DOST-SEI, emailed to you as a link
 *   5. Leaderboard         Dashboard: schools ranked by streak and response rate
 *   6. Attendance lists    menu: one Google Sheet with who attended each START event
 *   7. Event reminders     the day before an event, its Campus Leads get a short email
 *   8. Self-check email    every Monday, 🩺 runs and emails you ONLY if something is wrong
 *   9. Weekly backup       every Sunday, a private copy (values) in "START Tracker backups"; last 8 kept
 *  10. Archive             menu: Graduated / Dropped scholars move to the Archive tab (history kept)
 *  Runs every morning (extrasDaily, the tracker's email account only). Each job has an ON/OFF tick in Settings.
 * ================================================================ */

var BACKUP_FOLDER = 'START Tracker backups';
var BACKUPS_KEPT = 8;
var ARCHIVE_TAB = 'Archive';
var LEADERBOARD_MARK = '🏆  SCHOOL LEADERBOARD';

// A Settings tick that is ON unless you untick it (blank = ON, so it works before the Settings tab is rebuilt)
function extraOn_(key) { try { return getSettings_()[key] !== false; } catch (e) { return true; } }

function installExtrasTrigger_() {
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'extrasDaily')) {
    ScriptApp.newTrigger('extrasDaily').timeBased().everyDays(1).atHour(5).inTimezone(CONFIG.timezone).create();
  }
}

// Every morning at about 5 AM
function extrasDaily(e) {
  if (!ownerGuard_(e)) return;
  const ss = SpreadsheetApp.getActive(), now = new Date(), day = DAYS[(now.getDay() + 6) % 7];
  const run = (name, fn) => { try { fn(); } catch (err) { Logger.log(name + ': ' + err.message); } };
  run('Open/close', () => applyFormWindow_());
  run('Website accounts', () => webCleanAccounts_(ss));
  run('Email checks', () => checkScholarEmails_(ss));
  run('Not recognised', () => adoptRecognised_(ss));
  run('Leaderboard', () => writeLeaderboard_(ss));
  run('Event reminders', () => {
    const st = getSettings_();
    const inReminder = st.autoWeekly === true && day === st.remDay && (() => { try { return calOn_(); } catch (e) { return false; } })();
    if (extraOn_('eventRemind') && !inReminder) sendEventReminders_(ss);   // reminder day: the reminder already lists the events
  });
  run('Old answers removed (privacy)', () => { if (now.getDate() === 2) { const n = cleanOldAnswers_(ss) + joinTidy_(ss); if (n) Logger.log(n + ' old rows removed'); } });
  run('Join requests', () => joinDigest_(ss));
  run('Monthly report', () => { if (now.getDate() === 1 && extraOn_('monthlyReport')) monthlyReportAuto_(); });
  run('Self-check', () => { if (!RR_EMAILS) selfCheckEmail_(); else if (day === 'Monday' && extraOn_('selfCheck') && getSettings_().autoWeekly !== true) selfCheckEmail_(); });   // no emails: checked daily for the RR's START page   // weekly emails ON: it is inside the weekly summary
  run('Backup', () => { if (day === 'Sunday' && extraOn_('backups')) weeklyBackup_(ss); });
}

/* ---------- 3. the weekly check-in opens and closes by itself ---------- */

// true while this week's check-in should take answers: from the opening day through the deadline day
function checkinWindowOpen_(now) {
  const st = getSettings_();
  const d = (now.getDay() + 6) % 7, o = DAYS.indexOf(st.mOpenDay), due = DAYS.indexOf(st.mDueDay);
  if (o === -1 || due === -1) return true;
  return o <= due ? (d >= o && d <= due) : (d >= o || d <= due);
}

function applyFormWindow_(now) {
  const form = typeof memberForm_ === 'function' ? memberForm_() : null;
  if (!form) return '';
  const st = getSettings_();
  const open = !extraOn_('autoClose') || checkinWindowOpen_(now || new Date());
  if (open && !form.isAcceptingResponses()) retry_(() => form.setAcceptingResponses(true));
  if (!open && form.isAcceptingResponses()) {
    retry_(() => form.setCustomClosedFormMessage(`This week's START check-in is closed. The next one opens on ${st.mOpenDay}. See you then!`));
    retry_(() => form.setAcceptingResponses(false));
  }
  return open ? 'open' : 'closed';
}

/* ---------- 2. email checks in Scholars › Contact ---------- */

// "student.usep.edu.ph" -> "usep.edu.ph"; "gmail.com" stays
function baseDomain_(d) {
  const p = String(d || '').toLowerCase().replace(/^www\./, '').split('.').filter(Boolean);
  if (p.length >= 3 && /^(edu|gov|com|org)$/.test(p[p.length - 2]) && p[p.length - 1].length === 2) return p.slice(-3).join('.');
  return p.slice(-2).join('.');
}

// Marks problem emails in Scholars (orange + a note) and returns [{row, name, email, why}]
function checkScholarEmails_(ss) {
  const sc = ss.getSheetByName(TAB.scholars), sch = ss.getSheetByName(TAB.schools);
  if (!sc || sc.getLastRow() < 2) return [];
  const n = sc.getLastRow() - 1;
  const rows = sc.getRange(2, 1, n, 5).getValues();
  const domainOf = {};                                            // school -> its website's domain
  if (sch) sch.getRange(2, 1, CONFIG.maxSchools, 12).getValues().forEach(r => {
    const name = String(r[0]).trim(), site = String(r[11]).trim().replace(/^https?:\/\//, '').split('/')[0];
    if (name && site) domainOf[name] = baseDomain_(site);
  });
  const schoolOfDomain = {};
  Object.keys(domainOf).forEach(s => { (schoolOfDomain[domainOf[s]] = schoolOfDomain[domainOf[s]] || []).push(s); });
  const seen = {}, out = [];
  rows.forEach((r, i) => { const e = emailOf_(r[4]); if (e) (seen[e] = seen[e] || []).push(i); });
  const notes = sc.getRange(2, 5, n, 1).getNotes();
  rows.forEach((r, i) => {
    const e = emailOf_(r[4]), school = String(r[1]).trim();
    let why = '';
    if (e && seen[e].length > 1) why = 'Same email as row ' + seen[e].filter(j => j !== i).map(j => j + 2).join(', ');
    else if (e) {
      const owners = schoolOfDomain[baseDomain_(e.split('@')[1])] || [];
      if (owners.length && owners.indexOf(school) === -1) why = `This looks like a ${owners.join('/')} email, but the scholar is in ${school || '(no school)'}`;
    }
    const cell = sc.getRange(i + 2, 5), old = String(notes[i][0] || '');
    if (why) {
      if (old !== '⚠ ' + why) cell.setNote('⚠ ' + why).setBackground('#FFE2B8');
      out.push({ row: i + 2, name: String(r[0]).trim(), email: e, why });
    } else if (old.indexOf('⚠ ') === 0) cell.clearNote().setBackground(null);
  });
  return out;
}

/* ---------- 5. leaderboard on the Dashboard ---------- */

function leaderboardRows_(ss) {
  const data = focalData_(ss).filter(s => s.focal);
  const ts = termStart_(ss);
  const start = ts instanceof Date ? mondayOf_(ts) : null;
  const weeks = start ? Math.max(1, Math.floor((mondayOf_(new Date()) - start) / (7 * 86400000)) + 1) : null;
  return data.map(s => {
    const answered = new Set(s.checkins.filter(c => !start || c.date >= start).map(c => mondayOf_(c.date).getTime())).size;
    const rate = weeks ? Math.min(100, Math.round(100 * answered / weeks)) : null;
    return { name: s.name, lead: displayName_(s.focal), streak: s.streak, term: s.term, rate, week: s.thisWeek };
  }).sort((a, b) => (b.streak - a.streak) || ((b.rate || 0) - (a.rate || 0)) || a.name.localeCompare(b.name));
}

// Its own block at the bottom of the Dashboard (below the charts); rewritten every morning
function writeLeaderboard_(ss) {
  if (DB_MODE) return;                                   // database mode: not used (see 00_Config)
  const sh = ss.getSheetByName(TAB.dashboard);
  if (!sh) return;
  const W = 6;
  const colA = sh.getRange(1, 1, sh.getMaxRows(), 1).getValues().map(r => String(r[0]));
  let top = colA.findIndex(v => v.indexOf(LEADERBOARD_MARK) === 0) + 1;
  if (top > 0) sh.getRange(top, 1, Math.min(70, sh.getMaxRows() - top + 1), W).breakApart().clear();
  else {
    let bottom = sh.getLastRow();
    sh.getCharts().forEach(c => { try { bottom = Math.max(bottom, c.getContainerInfo().getAnchorRow() + 18); } catch (e) { /* ignore */ } });
    top = bottom + 3;
  }
  const rows = leaderboardRows_(ss);
  if (sh.getMaxRows() < top + rows.length + 4) sh.insertRowsAfter(sh.getMaxRows(), top + rows.length + 4 - sh.getMaxRows());
  const head = sh.getRange(top, 1, 1, W).merge().setValue(LEADERBOARD_MARK + '  ·  ranked by streak, then response rate this term');
  header_(head, THEME.navy); head.setHorizontalAlignment('left').setFontSize(11);
  sh.setRowHeight(top, 30);
  sh.getRange(top + 1, 1, 1, W).setValues([['#', 'School', 'START Campus Lead', '🔥 Streak (weeks)', 'Response rate this term', 'This week']])
    .setFontWeight('bold').setFontSize(9).setFontColor(THEME.muted).setBackground(THEME.band).setFontFamily(THEME.font);
  const medal = i => ['🥇', '🥈', '🥉'][i] || String(i + 1);
  const body = rows.map((r, i) => [medal(i), r.name, r.lead, r.streak, r.rate === null ? '—' : r.rate + '%', r.week ? '✅ In' : '⏳ Not yet']);
  if (body.length) {
    const rg = sh.getRange(top + 2, 1, body.length, W).setValues(body).setFontFamily(THEME.font).setFontSize(10).setFontColor(THEME.navy);
    rg.setBorder(true, true, true, true, true, true, TABLE_LINE, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(top + 2, 1, body.length, 1).setHorizontalAlignment('center');
    sh.getRange(top + 2, 2, body.length, 1).setFontWeight('bold');
    sh.getRange(top + 2, 4, body.length, 3).setHorizontalAlignment('center');
    body.forEach((b, i) => { if (i < 3) sh.getRange(top + 2 + i, 1, 1, W).setBackground(['#FFF3B0', '#EEF2F7', '#FBE7D6'][i]); });
  } else {
    sh.getRange(top + 2, 1, 1, W).merge().setValue('No START Campus Leads yet.').setFontColor(THEME.muted);
  }
  sh.getRange(top + 3 + Math.max(body.length, 1), 1, 1, W).merge()
    .setValue('💡 Screenshot this for the Campus Leads group chat: a little friendly competition helps.')
    .setFontColor(THEME.muted).setFontSize(9).setFontFamily(THEME.font);
}

/* ---------- 7. the day before an event: a short email to its Campus Leads ---------- */

function sendEventReminders_(ss) {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const tomorrow = new Date(today.getTime() + 86400000);
  const list = upcomingEvents_(ss, 2).filter(e => e.date >= tomorrow && e.date < new Date(tomorrow.getTime() + 86400000) && e.who !== WHO.me);
  if (!list.length) return 0;
  const leads = schoolRows_(ss).filter(s => s.focal && s.email);
  const per = {};                                                    // email -> { school, focal, events: [] }
  list.forEach(e => leads.forEach(s => {
    if (e.who === WHO.all || e.who === WHO.leads || e.who === s.name) {
      (per[s.email] = per[s.email] || { school: s.name, focal: s.focal, events: [] }).events.push(e);
    }
  }));
  const day = Utilities.formatDate(tomorrow, CONFIG.timezone, 'yyyy-MM-dd');
  let sent = 0;
  Object.keys(per).forEach(email => {
    const p = per[email];
    if (alreadySent_('eventRemind', day, email) || mailRoom_() < 1) return;
    let calUrl = 'https://calendar.google.com';
    try { if (calOn_() && calStore_().leads) calUrl = calLink_(calStore_().leads); } catch (e) { /* calendar off */ }
    const rows = p.events.map(e => [e.title, eventWhen_(e), e.place || '—']);
    const html = eventTomorrowHtml_(p.focal, rows, calUrl);
    sendMail_({ to: email, subject: `START · Tomorrow: ${p.events.map(e => e.title).join(', ')}`, htmlBody: html,
                body: `Hi ${displayName_(p.focal)},\n\nA reminder for tomorrow:\n\n` +
                      p.events.map(e => `• ${e.title}: ${eventWhen_(e)}${e.place ? ', ' + e.place : ''}`).join('\n') +
                      `\n\nCalendar: ${calUrl}\n\nRespectfully,\n` + rrSignature_(getSettings_()).join('\n') });
    markSent_('eventRemind', day, email);
    sent++;
  });
  return sent;
}

// "Event tomorrow" email to a Campus Lead (also used by 🧪 test emails)
function eventTomorrowHtml_(focal, rows, calUrl) {
  return designedEmail_({
    badge: 'Tomorrow', title: ['Event', 'Tomorrow'], greeting: `Hi ${displayName_(focal).split(' ')[0] || 'there'},`,
    intro: [`${rows.length === 1 ? 'A START event is' : 'START events are'} happening tomorrow. Kindly remind your members.`],
    extraHtml: emailTable_(['Event', 'When', 'Where / link'], rows),
    link: calUrl, buttonText: 'Open the START calendar', chips: [['', ''], ['', ''], ['', '']],
  });
}

/* ---------- 4. monthly report (on the 1st, for last month) ---------- */

function monthlyReportAuto_(force) {
  const now = new Date();
  const from = new Date(now.getFullYear(), now.getMonth() - 1, 1), to = new Date(now.getFullYear(), now.getMonth(), 1);
  const label = Utilities.formatDate(from, CONFIG.timezone, 'MMMM yyyy');
  const me = Session.getEffectiveUser().getEmail();
  if (!force && alreadySent_('monthlyReport', label, me)) return '';
  const url = makeTermReport({ from, to, label: 'Monthly report – ' + label, silent: true });
  docProps_().setProperty('LAST_REPORT', JSON.stringify({ label, url, at: Date.now() }));   // shown on the RR's START page
  if (!RR_EMAILS) { markSent_('monthlyReport', label, me); return url; }
  sendMail_({ to: me, subject: `START · Monthly report, ${label}`,
    htmlBody: designedEmail_({ badge: 'MONTHLY', title: ['Monthly', 'Report'], greeting: '', signature: [],
      intro: [`Your ${label} report is ready in your Drive. Add your notes, then File › Download › PDF.`],
      link: url, buttonText: 'Open the report', chips: [['', ''], ['', ''], ['', '']] }),
    body: `Your START report for ${label} is ready: ${url}` });
  markSent_('monthlyReport', label, me);
  return url;
}
function monthlyReportNow() {
  requireSetup_();
  const url = monthlyReportAuto_(true);
  notify_('📄 Last month\'s report is ready' + (RR_EMAILS ? ' (also emailed to you)' : ' (also on your START page)') + ':\n' + url);
}

/* ---------- 8. self-check email (Mondays, only when something is wrong) ---------- */

function selfCheckEmail_() {
  const rows = healthRows_().filter(r => r[0] === '❌' || r[0] === '⚠️');
  docProps_().setProperty('LAST_CHECK', JSON.stringify({ at: Date.now(), rows: rows.slice(0, 12).map(r => [r[0], String(r[1]).slice(0, 160), String(r[2] || '').slice(0, 200)]) }));
  if (!RR_EMAILS) return rows.filter(r => r[0] === '❌').length;               // shown on the RR's START page instead
  if (!rows.some(r => r[0] === '❌')) return 0;                    // only real problems trigger an email
  const me = Session.getEffectiveUser().getEmail();
  const week = Utilities.formatDate(mondayOf_(new Date()), CONFIG.timezone, 'yyyy-MM-dd');
  if (alreadySent_('selfCheck', week, me)) return 0;
  sendMail_({ to: me, subject: `START · Tracker check: ${rows.filter(r => r[0] === '❌').length} thing(s) to fix`,
    htmlBody: designedEmail_({ badge: 'SELF-CHECK', badgeColor: '#B3261E', title: ['Tracker', 'Check'], greeting: '', signature: [],
      intro: ['Something in the tracker needs you. Everything else is fine.'],
      extraHtml: emailTable_(['', 'What', 'How to fix'], rows.map(r => [r[0], r[1], r[2] || ''])),
      link: SpreadsheetApp.getActive().getUrl(), buttonText: 'Open the tracker',
      outro: ['Then click START Tracker › 🩺 Check my tracker to confirm.'] }),
    body: rows.map(r => `${r[0]} ${r[1]} → ${r[2] || ''}`).join('\n') });
  markSent_('selfCheck', week, me);
  return 1;
}

/* ---------- 9. weekly backup (private, values only, last 8 kept) ---------- */

function backupFolder_(ss) {
  const props = docProps_();
  const id = props.getProperty('BACKUP_FOLDER');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { /* gone */ } }
  const parents = DriveApp.getFileById(ss.getId()).getParents();
  const f = (parents.hasNext() ? parents.next() : DriveApp.getRootFolder()).createFolder(BACKUP_FOLDER);
  props.setProperty('BACKUP_FOLDER', f.getId());
  return f;
}

function weeklyBackup_(ss) {
  const folder = backupFolder_(ss);
  const stamp = Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd');
  const name = `START Tracker backup – ${stamp}`;
  if (folder.getFilesByName(name).hasNext()) return '';
  const dest = SpreadsheetApp.create(name);
  DriveApp.getFileById(dest.getId()).moveTo(folder);
  ss.getSheets().forEach(sh => {
    try {
      const copy = sh.copyTo(dest).setName(sh.getName());
      const rg = copy.getDataRange();
      rg.setValues(rg.getValues());                                  // a snapshot: plain values, no formulas or form links
    } catch (e) { Logger.log('Backup of ' + sh.getName() + ': ' + e.message); }
  });
  const first = dest.getSheets()[0];
  if (dest.getSheets().length > 1 && /^Sheet1$|^Sheet 1$/.test(first.getName())) dest.deleteSheet(first);
  // keep the newest few
  const all = [], it = folder.getFiles();
  while (it.hasNext()) { const f = it.next(); if (!f.isTrashed() && f.getName().indexOf('START Tracker backup – ') === 0) all.push(f); }
  all.sort((a, b) => b.getName().localeCompare(a.getName())).slice(BACKUPS_KEPT).forEach(f => f.setTrashed(true));
  return dest.getUrl();
}
function backupNow() {
  requireSetup_();
  const url = weeklyBackup_(SpreadsheetApp.getActive());
  notify_(url ? '💾 Backup saved (private, in "' + BACKUP_FOLDER + '"):\n' + url : 'Today\'s backup already exists in "' + BACKUP_FOLDER + '".');
}

/* ---------- 6. attendance list per event (from the weekly check-in answers) ---------- */

function makeAttendanceList() {
  requireSetup_();
  const ss = SpreadsheetApp.getActive();
  const sh = memberTab_(ss);
  if (!sh || sh.getLastRow() < 2) { notify_('No weekly check-in answers yet.'); return; }
  const c = memberCols_(sh);
  if (c.events < 0) { notify_('The weekly check-in has no events question yet.'); return; }
  const who = scholarByEmail_(ss), by = {};
  const since = new Date(Date.now() - 90 * 86400000);
  sh.getRange(2, 1, sh.getLastRow() - 1, c.width).getValues().forEach(v => {
    if (!(v[0] instanceof Date) || v[0] < since) return;
    const email = c.email >= 0 ? emailOf_(v[c.email]) : '', p = who[email] || {};
    splitEvents_(v[c.events]).filter(x => x && !missedPick_(x)).forEach(ev => {
      (by[ev] = by[ev] || {})[email || Math.random()] = [p.name || '(not in Scholars)', p.school || '', email, v[0]];
    });
  });
  const events = Object.keys(by);
  if (!events.length) { notify_('Nobody has ticked an event in the last 90 days yet.'); return; }
  const out = SpreadsheetApp.create(`START attendance – ${Utilities.formatDate(new Date(), CONFIG.timezone, 'MMM d, yyyy')}`);
  try { DriveApp.getFileById(out.getId()).moveTo(portalFolder_(ss)); } catch (e) { /* stays in My Drive */ }
  const first = out.getSheets()[0];
  events.slice(0, 40).forEach((ev, i) => {
    const t = i === 0 ? first : out.insertSheet();
    t.setName(String(i + 1) + '. ' + ev.replace(/[\[\]\*\?\/\\:]/g, ' ').slice(0, 90));
    const rows = Object.values(by[ev]).sort((a, b) => a[1].localeCompare(b[1]) || a[0].localeCompare(b[0]));
    t.getRange(1, 1, 1, 5).merge().setValue(`${ev}  ·  ${rows.length} attended (from the weekly check-in)`);
    header_(t.getRange(1, 1, 1, 5), THEME.navy);
    t.getRange(2, 1, 1, 5).setValues([['#', 'Name', 'School', 'Email', 'Answered on']]).setFontWeight('bold').setBackground(THEME.band);
    t.getRange(3, 1, rows.length, 5).setValues(rows.map((r, j) => [j + 1].concat(r)));
    t.getRange(3, 5, rows.length, 1).setNumberFormat('mmm d, yyyy');
    t.getRange(1, 1, rows.length + 2, 5).setFontFamily(THEME.font);
    [40, 230, 120, 260, 120].forEach((w, k) => t.setColumnWidth(k + 1, w));
    t.setFrozenRows(2);
  });
  dlgDone_('📋 Attendance lists', { head: `Attendance for ${events.length} event${events.length === 1 ? '' : 's'}`, text: 'Last 90 days, one tab per event.',
    link: out.getUrl(), button: 'Open the list', note: 'Has names and emails: share only with DOST-SEI.' });
}

/* ---------- 10. archive Graduated / Dropped scholars ---------- */

function archiveScholars() {
  requireSetup_();
  const ss = SpreadsheetApp.getActive(), sc = ss.getSheetByName(TAB.scholars);
  if (!sc || sc.getLastRow() < 2) { notify_('No scholars yet.'); return; }
  const W = SCHOLAR_COLS + 1;                                       // + Role
  const vals = sc.getRange(2, 1, sc.getLastRow() - 1, W).getValues();
  const pick = vals.map((r, i) => ({ r, row: i + 2 })).filter(x => ['Graduated', 'Dropped'].indexOf(String(x.r[5]).trim()) !== -1 && String(x.r[0]).trim());
  if (!pick.length) { notify_('Nobody is marked Graduated or Dropped in Scholars › Status.'); return; }
  const ui = ui_();
  const ok = ui.alert(`Move ${pick.length} scholar(s) to the Archive tab?`,
    pick.slice(0, 12).map(x => `• ${x.r[0]} (${x.r[1]}, ${x.r[5]})`).join('\n') + (pick.length > 12 ? `\n… and ${pick.length - 12} more` : '') +
    '\n\nTheir past check-ins stay. They stop seeing the START calendars and are no longer counted.', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  let ar = ss.getSheetByName(ARCHIVE_TAB);
  if (!ar) {
    ar = ss.insertSheet(ARCHIVE_TAB);
    const head = sc.getRange(1, 1, 1, W).getValues()[0].map(h => String(h).replace(/^[✏️🔒\s]+/u, '')).concat(['Archived on']);
    ar.getRange(1, 1, 1, head.length).setValues([head]);
    header_(ar.getRange(1, 1, 1, head.length), THEME.slate || THEME.navy);
    ar.setFrozenRows(1).setTabColor('#9AA7BA');
  }
  const at = Math.max(ar.getLastRow(), 1) + 1;
  ar.getRange(at, 1, pick.length, W + 1).setValues(pick.map(x => x.r.concat([new Date()])))
    .setFontFamily(THEME.font).setFontColor(THEME.navy);
  ar.getRange(at, W + 1, pick.length, 1).setNumberFormat('mmm d, yyyy');
  pick.slice().reverse().forEach(x => sc.deleteRow(x.row));
  try { syncRoles_(ss); } catch (e) { /* roles refresh on the next edit */ }
  notify_(`🗄 ${pick.length} scholar(s) moved to the "${ARCHIVE_TAB}" tab.` +
    (pick.some(x => String(x.r[SCHOLAR_COLS]).indexOf('⭐') !== -1) ? '\n\n⚠ One of them was a START Campus Lead: pick a new Lead for that school.' : ''));
}

/* ---------- 11. 📨 welcome email: weekly check-in + "add the START calendar", one email per person ---------- */
// NEVER automatic: only when you click ✉️ Send emails now › 📨 Send welcome emails. Each person gets it once (the button
// can resend to everyone). If calendars are turned on later, the button sends people already welcomed one short
// "your START calendar is ready" email. Google's own "calendar shared with you" email is off: this is the one email.
// 🗓 Today reminds you when members with an email have not been welcomed yet.

// The calendar part of a welcome (HTML), for a member or a Campus Lead; '' when the calendars are off
function welcomeCalendarHtml_(school, isLead) {
  let c = null;
  try { if (calOn_()) c = calStore_(); } catch (e) { c = null; }
  if (!c || !c.region) return '';
  const btn = (id, label) => id ? `<a href="${calLink_(id)}" style="display:inline-block;margin:4px 6px 4px 0;padding:9px 14px;border-radius:999px;` +
    `background:#0099FF;color:#ffffff;font-weight:700;text-decoration:none;font-size:13px;">Add: ${esc_(label)}</a>` : '';
  return emailSection_('1 · Add the START calendar') +
    '<p style="margin:0 0 6px 0;font-size:13px;color:#33415c;">Tap a button while signed in to this Google account, then <b>Add</b>.</p>' +
    btn(c.region, CAL_NAMES.region) + btn(c.schools && c.schools[school], CAL_NAMES.school(school)) +
    (isLead ? btn(c.leads, CAL_NAMES.leads) : '');
}

function welcomeHtml_(p, opts) {
  const st = getSettings_(), link = checkinLink_();
  const cal = welcomeCalendarHtml_(p.school, p.isLead);
  const n = cal ? 2 : 1;
  const steps = (opts.calendarOnly ? '' : '') + cal +
    (opts.calendarOnly ? '' :
      emailSection_(`${n} · Answer the weekly check-in`) +
      `<p style="margin:0;font-size:13px;color:#33415c;">We email it every <b>${esc_(st.mOpenDay)}</b>. Sign in with <b>${esc_(p.email)}</b>.` +
      (p.isLead ? ' Your answer also counts as your school\'s check-in.' : '') + '</p>');
  return designedEmail_({
    badge: opts.calendarOnly ? 'CALENDAR' : 'WELCOME', title: opts.calendarOnly ? ['Your START', 'Calendar'] : ['Welcome to', 'START'],
    greeting: `Hi ${displayName_(p.name).split(' ')[0] || 'there'},`,
    intro: [opts.calendarOnly ? 'Your START calendar is ready. Add it once to see every START event in your Google Calendar.'
                              : `Welcome to START! You're on the list for ${p.school}${p.isLead ? ' as its ⭐ Campus Lead' : ''}.`],
    extraHtml: steps + (p.portal && !DB_MODE ? emailSection_('Your Campus Lead portal') +
      `<p style="margin:0;font-size:13px;"><a href="${p.portal}" style="color:#0099FF;font-weight:700;">Open your private portal ↗</a></p>` : ''),
    link: opts.calendarOnly ? (cal ? calLink_(calStore_().region) : link) : link,
    buttonText: opts.calendarOnly ? 'Add the START calendar' : 'Open the weekly check-in',
    chips: [['SCHOOL', p.school], ['OPENS', st.mOpenDay], ['DUE', st.dueDay]],
    outro: [privacyNote_().split('. ')[0] + '.', 'START will never ask for your password or sign-in code.', 'Questions? Reply to this email.'],
  });
}

// Who should get a welcome: active scholars with an email (Campus Leads marked)
function welcomeList_(ss) {
  const sc = ss.getSheetByName(TAB.scholars);
  if (!sc || sc.getLastRow() < 2) return [];
  const schools = {}; schoolRows_(ss).forEach(s => { schools[s.name] = s; });
  const store = (() => { try { return portalStore_().map; } catch (e) { return {}; } })();
  const seen = {};
  return sc.getRange(2, 1, sc.getLastRow() - 1, 6).getValues().map(r => {
    const email = emailOf_(r[4]), school = String(r[1]).trim(), s = schools[school] || {};
    if (!email || seen[email] || ['Graduated', 'Dropped'].indexOf(String(r[5]).trim()) !== -1) return null;
    seen[email] = 1;
    const isLead = !!s.email && s.email === email, p = store[school] || {};
    return { name: String(r[0]).trim(), email, school, isLead, lead: s.focal || '',
             portal: isLead && p.id && p.email === email ? `https://docs.google.com/spreadsheets/d/${p.id}/edit` : '' };
  }).filter(Boolean);
}

// Sends what is due. opts: { resend: true } = everyone again (menu). Returns { sent, calendar, waiting }.
function sendWelcomeEmails_(ss, opts) {
  opts = opts || {};
  const calReady = (() => { try { return calOn_() && !!calStore_().region; } catch (e) { return false; } })();
  const out = { sent: 0, calendar: 0, waiting: 0 };
  welcomeList_(ss).forEach(p => {
    const first = opts.resend || !alreadySent_('welcome', 'v1', p.email);
    const calOnly = !first && calReady && !alreadySent_('welcomeCal', 'v1', p.email);
    if (!first && !calOnly) return;
    if (mailRoom_() < 1) { out.waiting++; return; }               // daily email limit: the rest go tomorrow morning
    const html = welcomeHtml_(p, { calendarOnly: calOnly });
    sendMail_({ to: p.email, subject: calOnly ? 'START · Your calendar is ready' : `START · Welcome, ${p.school}`, htmlBody: html,
                body: `Hi ${displayName_(p.name)},\n\n` +
                      (calOnly ? 'Your START calendar is ready.' : `Welcome to START, ${p.school}! Each week you answer a short check-in.\n\nWeekly check-in: ${checkinLink_()}\nSign in with ${p.email}.`) +
                      (calReady ? `\n\nAdd the START calendar: ${calLink_(calStore_().region)}` : '') +
                      '\n\nRespectfully,\n' + rrSignature_(getSettings_()).join('\n') });
    if (first) { markSent_('welcome', 'v1', p.email); out.sent++; }
    if (calReady) { markSent_('welcomeCal', 'v1', p.email); if (calOnly) out.calendar++; }
  });
  return out;
}

// ✉️ Send emails now › 📨 Send welcome emails
function sendWelcomeNow() {
  requireSetup_();
  const ss = SpreadsheetApp.getActive(), ui = ui_();
  const list = welcomeList_(ss);
  if (!list.length) { notify_('Nobody to welcome yet: add emails in Scholars › Contact (E) for active scholars.'); return; }
  const calReady = (() => { try { return calOn_() && !!calStore_().region; } catch (e) { return false; } })();
  const due = list.filter(p => !alreadySent_('welcome', 'v1', p.email) || (calReady && !alreadySent_('welcomeCal', 'v1', p.email)));
  let resend = false;
  if (!due.length) {
    if (ui.alert('Everyone has their welcome email already.',
      `Send it AGAIN to all ${list.length} member(s) with an email? (e.g. they lost it, or the calendar links changed)`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    resend = true;
  } else if (ui.alert(`Send the welcome email to ${due.length} member(s)?`,
      due.slice(0, 12).map(p => `• ${p.name} (${p.school}${p.isLead ? ', ⭐ Lead' : ''}) – ${p.email}`).join('\n') + (due.length > 12 ? `\n… and ${due.length - 12} more` : '') +
      (calReady ? '\n\nIncludes the check-in link and calendar buttons.' : '\n\nIncludes the check-in link.'),
      ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const r = sendWelcomeEmails_(ss, { resend });
  notify_(`📨 ${r.sent} welcome email${r.sent === 1 ? '' : 's'} sent` + (r.calendar ? ` + ${r.calendar} calendar email${r.calendar === 1 ? '' : 's'}` : '') +
          (r.waiting ? `\n\n⏳ ${r.waiting} go out tomorrow morning (daily limit).` : ''));
}

/* ---------- 12. weekly check-in email to WELCOMED members (Friday) + follow-up (Saturday) ---------- */
// Rides on the Campus Lead reminder/follow-up (same 🚀 switch, same days, same design). Only people you welcomed
// with 📨 get it; Campus Leads are left out (they get the Lead reminder). Over the daily limit: the rest wait in
// the queue and go out the next morning. Every send is marked done: nobody gets the same email twice.

function composeMemberEmail_(type, p, w) {
  const st = getSettings_(), f = d => Utilities.formatDate(d, CONFIG.timezone, 'EEE, MMM d');
  const first = displayName_(p.name).split(' ')[0] || 'there';
  let events = '';
  try { events = eventsForEmail_(p.school); } catch (e) { /* calendar off */ }
  const html = designedEmail_({
    badge: type === 'followup' ? 'FOLLOW-UP' : 'THIS WEEK', badgeColor: type === 'followup' ? '#E0B800' : THEME.blue,
    title: ['Weekly', 'Check-in'], greeting: `Hi ${first},`,
    intro: [type === 'followup' ? 'We haven\'t received your check-in this week yet. It takes about a minute.'
                                : 'This week\'s START check-in is open. It takes about a minute.'],
    extraHtml: events,
    link: checkinLink_(), buttonText: 'Answer the check-in', buttonNote: `Sign in with ${p.email}`,
    chips: [['SCHOOL', p.school], ['WEEK OF', Utilities.formatDate(w.monday, CONFIG.timezone, 'MMM d')], ['DUE', f(w.deadline)]],
    outro: ['Need help? Choose "Yes" in the form and we\'ll reach out.'],
  });
  return { subject: `START · ${type === 'followup' ? 'Reminder: w' : 'W'}eekly check-in, due ${f(w.deadline)}`, html,
           text: `Hi ${first},\n\nThis week's START check-in: ${checkinLink_()} (due ${f(w.deadline)}). Sign in with ${p.email}.\n\nRespectfully,\n` + rrSignature_(st).join('\n') };
}

// type 'reminder' | 'followup'; onlyEmails: from the waiting queue. Returns { sent, waiting }
// Welcomed members (not Leads) who have not answered yet (answered = { email: 1 })
function memberWeeklyList_(ss, answered) {
  return welcomeList_(ss).filter(p => !p.isLead && alreadySent_('welcome', 'v1', p.email) && !(answered || {})[p.email]);
}

// For the "Send now" confirmation: the members who would get this week's email right now
function memberWeeklyTargets_(type) {
  try {
    if (!oneFormOn_()) return [];
    const ss = SpreadsheetApp.getActive();
    const w = type === 'followup' ? followUpWeek_(new Date()) : reminderWeek_(new Date());
    let answered = {};
    if (type === 'followup' || remindTwoDays_()) {
      try { Object.values(memberData_(ss).bySchool).forEach(s => s.members.forEach(m => { if (m.now && m.email) answered[m.email] = 1; })); }
      catch (e) { answered = {}; }
    }
    return memberWeeklyList_(ss, answered).filter(p => !alreadySent_('m' + type, ymd_(w.deadline), p.email));
  } catch (e) { return []; }
}

function sendMemberWeekly_(type, w, onlyEmails) {
  const ss = SpreadsheetApp.getActive();
  if (!oneFormOn_()) return { sent: 0, waiting: 0 };
  if (remindTwoDays_() && daysToDeadline_(w.deadline) < 0) return { sent: 0, waiting: 0 };   // never after the deadline
  let answered = {};
  if (type === 'followup' || remindTwoDays_()) {
    try { Object.values(memberData_(ss).bySchool).forEach(s => s.members.forEach(m => { if (m.now && m.email) answered[m.email] = 1; })); }
    catch (e) { if (remindTwoDays_()) { Logger.log('Member reminders skipped (answers unreadable): ' + e.message); return { sent: 0, waiting: 0 }; } answered = {}; }
  }
  const period = ymd_(w.deadline), stamp = new Date(),
        label = (type === 'followup' ? 'Member follow-up' : 'Member check-in') + ' · due ' + Utilities.formatDate(w.deadline, CONFIG.timezone, 'MMM d');
  const list = memberWeeklyList_(ss, answered).filter(p => !onlyEmails || onlyEmails.indexOf(p.email) !== -1);
  const log = [], waiting = [];
  let room = mailRoom_(), sent = 0;
  list.forEach(p => {
    if (alreadySent_('m' + type, period, p.email)) return;          // ✅ never twice
    if (room < 1) { waiting.push(p.email); return; }
    try {
      const m = composeMemberEmail_(type, p, w);
      sendMail_({ to: p.email, subject: m.subject, htmlBody: m.html, body: m.text });
      markSent_('m' + type, period, p.email);
      log.push([stamp, label, p.school, p.email, 'Sent']); sent++; room--;
    } catch (e) { log.push([stamp, label, p.school, p.email, 'Not sent: ' + e.message]); }
  });
  if (waiting.length) {
    const kept = queueEmails_({ kind: 'm' + type, schools: waiting, week: { monday: w.monday.getTime(), deadline: w.deadline.getTime() } });
    waiting.forEach(e => log.push([stamp, label, '', e, kept ? WAIT_NOTE : 'Not sent: daily email limit reached']));
  }
  if (log.length) { try { writeLog_(ss, log); } catch (e) { Logger.log('Reminder log: ' + e.message); } }
  return { sent, waiting: waiting.length };
}
