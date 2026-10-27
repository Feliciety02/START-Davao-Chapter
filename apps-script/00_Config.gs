/* =====================================================================
 *  START SCHOLAR TRACKER  ·  00_Config   (version 2026-10-27f)
 *  ---------------------------------------------------------------------
 *  Inside: Settings, colours, tab names, form questions, shared helpers.
 *  To customise: The Settings tab rows, colours (THEME) and the form question titles (Q).
 *
 *  The code has 9 files + Email_Reminder and Web_App (HTML). Keep these exact names:
 *    00_Config, 01_Region, 02_Menu_Setup, 03_Forms, 04_Emails, 05_Portals, 06_Calendar, 07_Look, 08_Tools
 *  To update: open each file, select all (Ctrl+A), paste the new version, Save.
 *  Then in the Google Sheet: START Tracker › ① Set up / fix the tracker.
 *  You never need to run anything from here: use the START Tracker menu in the Sheet.
 * ===================================================================== */
var START_MODULES = '2026-10-27f';

// The ▶ Run button runs the FIRST function of the open file: this one, so it is always safe.
function about_00_Config() {
  Logger.log('START Tracker' + (typeof START_MODULES === 'string' ? ' ' + START_MODULES : '') + ', file 00_Config. Nothing to run here: open the Google Sheet and use the START Tracker menu.');
}


// ━━━━━━━━━━━━━━━━━━━━  Config  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  CONFIG  ·  colours, tab names, form questions, privacy notice
 * ================================================================ */

var CONFIG = {
  timezone: 'Asia/Manila',
  maxSchools: 50,                // Schools tab formulas cover rows 2-51
};

var TAB = {
  start: 'Start here', dashboard: 'Dashboard', schools: 'Schools',
  scholars: 'Scholars', checkins: 'Campus Lead check-ins', settings: 'Settings', log: 'Reminder log', links: 'Links', outreach: 'Outreach log', focal: 'START Campus Leads', events: 'Events', calview: 'Calendar view', mcheckins: 'Check-in answers',
};
var NO_SCHOOLS = '(no schools added yet)';
// The sheet is a plain DATABASE: no banners, "how to use" panels, Start here, Links, Dashboard, Calendar view,
// START Campus Leads tab, Reminder log or portal files. The website shows all of that. false = the old full look.
var DB_MODE = true;
// The weekly check-in is answered on the START website only (no Google Forms). Answers go straight into the
// Check-in answers tab; a Campus Lead's answer is also copied to Campus Lead check-ins. false = the old Google Forms.
var WEB_CHECKIN = DB_MODE;
// Emails TO the RR (weekly summary, copies of what was sent, quiet schools, tracker check, monthly report, join requests,
// urgent alerts): false = none; all of it is on the RR's START page instead (saves the daily email limit).
// Urgent help from a member still emails that school's Campus Lead (and the RR only when the school has no Lead email).
var RR_EMAILS = false;

// START-DOST Membership Eligibility: the 13 courses shown in the supplied official image.
// Registration only accepts these exact dropdown values. Existing historical scholar records are not deleted.
var START_PROGRAMS = {
  source: 'START-DOST Membership Eligibility', checked: '2026-10-23',
  list: [
    'Applied Mathematics',
    'Technology and Livelihood Education with Specialization in Information and Communication Technology',
    'Computer Engineering',
    'Computer Science',
    'Electronics and Communication Engineering',
    'Industrial Management Engineering – Information Technology',
    'Information and Communications Technology',
    'Information Systems',
    'Information Technology',
    'Information Technology Systems',
    'Management Information Systems',
    'Statistics',
    'Mechatronics Engineering'
  ]
};
// Code version: the website reports it, so 🩺 Check my tracker can tell when "New version" was not deployed
var CODE_VERSION = '2026-10-27f';
// Menu paths named in messages (the short database menu groups them differently)
var MENU_CAL = DB_MODE ? '🛠 Setup tools › 📅 Set up / update calendars' : '📅 Events & calendar › Set up / update calendars';
var MENU_REPORT = DB_MODE ? '🛠 Setup tools › 📄 End-of-term report' : '📄 Reports & backups › Make an end-of-term report';

// Look & feel used everywhere
var THEME = {        // matches the StaRRt IGNITE deck: navy, electric blue, sunshine yellow, Poppins
  font: 'Poppins',
  navy: '#12203A', blue: '#0099FF', sun: '#FFDD00',
  slate: '#2F4A73', muted: '#5b6b84',
  band: '#EAF3FC', yellow: '#FFF6BF', amber: '#E0B800',
  green: '#1E8E3E', greenBg: '#D7F2DF', red: '#B3261E', redBg: '#FADCD9', orangeBg: '#FFE9B8',
  // every tab colour in one place (blue = you type, grey/slate = automatic, yellow = guides & settings)
  tabs: { start: '#FFDD00', dashboard: '#12203A', schools: '#0099FF', scholars: '#0099FF', outreach: '#0099FF',
          checkins: '#9AA7BA', log: '#2F4A73', focal: '#FFDD00', links: '#FFDD00', settings: '#FFDD00', portal: '#0099FF', events: '#0099FF', calview: '#FFDD00', mcheckins: '#9AA7BA' },
};
var PROTECT_TAG = 'START auto';

// Form question titles. The code finds answers by these titles, so don't rename them in the form.
var Q = {
  school:  'School',
  name:    'Your name',
  events:  'How did this week\'s START events go?',
  dc:      'How are your members doing on DataCamp?',
  help:    'Do you need help from START this week?',
  concern: 'Anything START should know? (optional)',
};
// Data Privacy Act of 2012 (RA 10173): the same promise everywhere personal data is collected or shown
// The notice names the tracker's region: use privacyNote_() in texts (PRIVACY_NOTE_FORM is the region-free fallback)
var PRIVACY_NOTE_FORM =
  'DATA PRIVACY NOTICE: START collects your name and your school\'s weekly activity only to monitor and support ' +
  'the START program in your region. It is kept in a private tracker, seen only by your Regional Representative and ' +
  'shared only with DOST-SEI / ONSSE, in line with the Data Privacy Act of 2012 (RA 10173). ' +
  'Answers are deleted 2 years after you leave START. By submitting this form, you agree to this. You may ask to see, correct or remove your information anytime by replying to any START email.';
var PRIVACY_NOTE_PORTAL =
  '🔒 Data privacy: this page contains personal information of your school\'s scholars (Data Privacy Act of 2012). ' +
  'Please do not share it, download it, or take screenshots of it for others. It is shared only with your Google account.';
var CONSENT_TITLE = 'Data privacy consent';   // required tick box at the end of both check-in forms
var CONSENT_AGREE = 'I agree that START may collect and use the information in this form for the START program, ' +
                    'in line with the Data Privacy Act of 2012 (RA 10173).';
var HELP_CHOICES = ['No, all good', 'Yes, but not urgent', 'Yes, urgent'];
// Campus Lead form choices (a Lead reports how things went, not head counts)
var EVENT_GO_CHOICES = ['🙌 Great, members showed up', '🙂 Okay, a few joined', '😬 Hardly anyone joined', '⏭ We sat this one out'];
var DC_CHOICES = ['🚀 Most are on track', '🙂 About half', '🐢 Only a few', '🤷 Not sure yet'];
var OLD_Q = { events: 'Events joined this week', dc: 'Active on DataCamp' };   // older forms: answers still read

function privacyNote_() {
  let label = '';
  try { label = regionProfile_().label; } catch (e) { label = ''; }
  return label ? PRIVACY_NOTE_FORM.replace('in your region', 'in ' + label) : PRIVACY_NOTE_FORM;
}


// ━━━━━━━━━━━━━━━━━━━━  Settings  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  SETTINGS TAB  (scholars change settings here, not in the code)
 * ================================================================ */

var SETTINGS = [
  // key, label, default, help
  ['region',     'Region',                              '',
   'Pick your region from the drop-down. Form titles, emails, the privacy notice, the calendar name, the province list and the ready-made school list all follow it.'],
  ['rrName',     'Your name (Regional Representative)', '',
   'Shown at the end of the message to START Campus Leads.'],
  ['recipients', 'Send weekly summary to',              '',
   'Email address(es), separated by commas. Leave blank to send only to yourself.'],
  ['day',        'Weekly email day',                    'Monday',
   'Day you get the weekly summary. Applies by itself.'],
  ['hour',       'Weekly email time',                   '7 AM',
   'Applies by itself.'],
  ['quietDays',  'Mark a school red after (days)',      14,
   'A school with no check-in for this many days turns red.'],
  ['tFocal',     'Target: START Campus Leads named',         6,   'Your KPI target.'],
  ['tReeng',     'Target: schools re-engaged',          2,   'Your KPI target.'],
  ['tRate',      'Target: weekly response rate (%)',    80,  'Share of schools (with a START Campus Lead) that answer each week.'],
  ['remDay',     'Automatic reminder day',              'Friday',
   'Day the check-in reminder is emailed to every START Campus Lead. Applies by itself.'],
  ['remHour',    'Automatic reminder time',             '8 AM',
   'Applies by itself.'],
  ['dueDay',     'Weekly check-in deadline (day): everyone',  'Sunday',
   'ONE deadline for everyone (members and Campus Leads): written in the reminder; the form closes the day after.'],
  ['fuDay',      'Follow-up day',                       'Saturday',
   'Day a follow-up goes ONLY to schools that have not answered yet (e.g. Saturday, the day before a Sunday deadline). Choose "No follow-up" to skip.'],
  ['dFocal',     'Deadline: START Campus Leads target',      new Date(2026, 10, 30),
   'Date the START Campus Lead target should be reached.'],
  ['dReeng',     'Deadline: re-engaged schools target', new Date(2027, 2, 31),
   'Date the re-engaged-schools target should be reached.'],
  ['autoWeekly', '🚀 Automatic weekly emails (tick = ON)', '',
   'Tick to turn ON: the weekly summary to you, and the reminder + follow-up to every START Campus Lead, ' +
   'on the days above. After each send you get a copy and a report. Untick to stop. Same as menu ⏰ Automatic emails › 🚀 Turn ON.'],
  // new rows go at the END so older trackers keep every value in its place
  ['replyTo',    'Replies go to (email)',               '',
   'When a Campus Lead presses Reply, the answer goes here. Leave blank to use your own account.'],
  ['missedAlert', 'Alert me: schools quiet 2+ weeks (tick = ON)', true,
   'With the weekly summary, you get a short list of schools that missed 2 or more check-ins in a row, ' +
   'with a ready-to-copy Messenger message for each Campus Lead.'],
  ['mOpenDay',   'Weekly check-in opens (day)',        'Friday',
   'Day the member check-in form switches to the new week (with that week\'s events). Campus Leads share it that day.'],
  ['mDueDay',    'Weekly check-in due (day): follows the deadline above', 'Sunday',
   'Same day as the deadline above (one deadline for everyone). Change the deadline row, not this one.'],
  ['dcStart',    'DataCamp starts on (date)',          '',
   'From this date the member check-in also asks about learning progress. Leave blank until DataCamp starts.'],
  ['dcEnd',      'DataCamp ends on (date)',            '',
   'Optional. After this date the learning question is no longer asked.'],
  ['autoClose',  'Close the check-in between weeks (tick = ON)', true,
   'The weekly check-in stops taking answers the day after its deadline and opens again on its opening day.'],
  ['eventRemind', 'Remind Campus Leads the day before an event (tick = ON)', true,
   'A short email to the Campus Leads of each event in the Events tab, the day before.'],
  ['monthlyReport', 'Monthly report on the 1st (tick = ON)', true,
   'On the 1st of each month you get a link to a ready Google Doc about last month (for DOST-SEI).'],
  ['selfCheck',  'Weekly self-check email (tick = ON)', true,
   'Every Monday the tracker checks itself and emails you ONLY if something needs fixing.'],
  ['backups',    'Weekly private backup (tick = ON)', true,
   'Every Sunday a private copy is saved in your Drive folder "START Tracker backups" (the last 8 are kept).'],
  // 🔐 YOUR Campus Lead portals: each RR decides what their own Leads see (applies at the next portal update)
  ['portalNote', 'Website: note at the top of every page (optional)', '',
   'e.g. "General Assembly on Oct 25, see you!". Leave blank for none.'],
  ['portalMembers', 'Website: Leads see their members\' weekly answers (tick = ON)', true,
   'Untick to hide members\' answers from the Campus Lead pages.'],
  ['portalScholars', 'Website: Leads see their school\'s scholar list (tick = ON)', true,
   'Untick to hide the scholar list (names, year, program) from the Campus Lead pages.'],
  ['termStart',  'Term start date',                    '',
   'Counts "this term" (streaks, end-of-term report) from this date.'],
  ['autoApprove', 'Auto-approve verified join requests (tick = ON)', false,
   'Everyone who registers from a listed school is approved at once (no masterlist check). Turning it on also approves requests already waiting. Only clashes with an existing record wait for you. The applicant confirms their email before setting a password.'],
];
var CHECKBOX_SETTINGS = ['autoWeekly', 'missedAlert', 'autoClose', 'eventRemind', 'monthlyReport', 'selfCheck', 'backups', 'portalMembers', 'portalScholars', 'autoApprove'];
var DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
var HOURS = Array.from({ length: 24 }, (_, h) =>
  (h % 12 || 12) + (h < 12 ? ' AM' : ' PM'));

function settingCell_(key) {          // e.g. "Settings!$B$7"
  const i = SETTINGS.findIndex(s => s[0] === key);
  // an unknown key means 02_Settings is older than the other files: stop instead of writing into the wrong cell
  if (i === -1) throw new Error('The code files are out of date (no "' + key + '" setting). Replace every code file with the latest version.');
  return `${TAB.settings}!$B$${i + 2}`;
}

function buildSettings_(ss) {
  const sh = getOrCreate_(ss, TAB.settings);
  // A and C are written in one go; B keeps what the RR typed (defaults only fill empty cells)
  const N = SETTINGS.length;
  const cur = sh.getRange(2, 2, N, 1).getValues();
  const props = docProps_();
  sh.getRange(2, 1, N, 1).setValues(SETTINGS.map(x => [x[1]]));
  sh.getRange(2, 3, N, 1).setValues(SETTINGS.map(x => [x[3]]));
  if (!DB_MODE) {
    sh.getRange(2, 1, N, 1).setFontWeight('bold');
    sh.getRange(2, 3, N, 1).setFontColor(THEME.muted).setWrap(true);
  }
  // the 🚀 switch starts ticked if weekly emails were already turned on with an older version
  const wasOn = ScriptApp.getProjectTriggers().some(t => ['weeklySummary', 'autoReminder'].indexOf(t.getHandlerFunction()) !== -1);
  const values = sh.getRange(2, 2, N, 1).setValues(SETTINGS.map((x, i) => [cur[i][0] === '' ?
    (x[0] === 'autoWeekly' ? wasOn : x[0] === 'autoApprove' ? props.getProperty('AUTO_APPROVE') === '1' : x[2]) : cur[i][0]]));
  if (!DB_MODE) values.setBackground(THEME.yellow);
  // One time: move the Campus Lead emails from the old Mon / Wed / Thu to Fri (reminder) / Sun (deadline) / Sat (follow-up)
  if (!props.getProperty('DAYS_V2')) {
    const cell = k => sh.getRange(settingCell_(k).split('!')[1]);
    if (cell('remDay').getValue() === 'Monday' && cell('dueDay').getValue() === 'Wednesday' && cell('fuDay').getValue() === 'Thursday') {
      cell('remDay').setValue('Friday'); cell('dueDay').setValue('Sunday'); cell('fuDay').setValue('Saturday');
      try { if (claimIfFree_()) applySchedule_(); } catch (e) { Logger.log('Schedule not moved: ' + e.message); }   // emails move with it
    }
    props.setProperty('DAYS_V2', '1');
  }
  CHECKBOX_SETTINGS.forEach(k => {
    const cell = sh.getRange(settingCell_(k).split('!')[1]).insertCheckboxes();
    if (!DB_MODE) cell.setHorizontalAlignment('center');
  });
  const dv = (key, rule) => sh.getRange(settingCell_(key).split('!')[1]).setDataValidation(rule);
  const list = items => SpreadsheetApp.newDataValidation().requireValueInList(items, true).setAllowInvalid(false).build();
  dv('replyTo', SpreadsheetApp.newDataValidation().requireFormulaSatisfied(
       `=OR(${settingCell_('replyTo').split('!')[1].replace(/\$/g, '')}="",REGEXMATCH(${settingCell_('replyTo').split('!')[1].replace(/\$/g, '')}&"","@"))`)
     .setAllowInvalid(false).setHelpText('One email address, or leave blank.').build());
  ['day', 'dueDay', 'mOpenDay'].concat(remindTwoDays_() ? [] : ['remDay']).forEach(k => dv(k, list(DAYS)));
  // one deadline: the second row only mirrors the first (grey, not typed in)
  const memberDue = sh.getRange(settingCell_('mDueDay').split('!')[1]).clearDataValidations()
    .setFormula('=' + settingCell_('dueDay').split('!')[1]);
  if (!DB_MODE) memberDue.setBackground('#EEF2F7').setFontColor(THEME.muted);
  if (remindTwoDays_()) {                        // set by the deadline: shown grey, with the rule in a note
    const reminderDay = sh.getRange(settingCell_('remDay').split('!')[1]).clearDataValidations()
      .setFormula('=INDEX({"Friday","Saturday","Sunday","Monday","Tuesday","Wednesday","Thursday"},MATCH(' + settingCell_('dueDay').split('!')[1] +
        ',{"Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"},0))')
      .setNote('Set by the deadline: ONE reminder, 2 days before it, only to scholars and Campus Leads who have not answered yet.');
    const followupDay = sh.getRange(settingCell_('fuDay').split('!')[1]).clearDataValidations().setValue('No follow-up')
      .setNote('Off: the website check-in sends one reminder only (2 days before the deadline).');
    if (!DB_MODE) {
      reminderDay.setBackground('#EEF2F7').setFontColor(THEME.muted);
      followupDay.setBackground('#EEF2F7').setFontColor(THEME.muted);
    }
  }
  ['hour', 'remHour'].forEach(k => dv(k, list(HOURS)));
  if (!remindTwoDays_()) dv('fuDay', list(DAYS.concat(['No follow-up'])));
  ['dFocal', 'dReeng', 'dcStart', 'dcEnd', 'termStart'].forEach(k => {
    dv(k, SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
      .setHelpText('Pick a date (double-click the cell for a calendar).').build());
    sh.getRange(settingCell_(k).split('!')[1]).setNumberFormat('mmmm d, yyyy');
  });
  ['quietDays', 'tFocal', 'tReeng', 'tRate'].forEach(k => dv(k, SpreadsheetApp.newDataValidation()
    .requireNumberBetween(1, 365).setAllowInvalid(false).setHelpText('Enter a whole number.').build()));
  sh.setFrozenRows(1);
  if (DB_MODE) {
    sh.getRange(1, 1, 1, 3).setValues([['Setting', 'Your value', 'What it means']]);
  } else {
    sh.getRange(1, 1, N + 1, 3).setFontFamily(THEME.font).setVerticalAlignment('middle');
    sh.getRange(2, 2, N, 1).setFontWeight('bold').setHorizontalAlignment('left')
      .setBorder(true, true, true, true, false, true, THEME.amber, SpreadsheetApp.BorderStyle.SOLID);
    markColumns_(sh, [[1, 'auto', 'Setting', 'Fixed labels.'], [2, 'edit', 'Your value (yellow cells)', 'Change these anytime.'],
                      [3, 'auto', 'What it means', 'Fixed explanations.', true]], N, false);
    sh.setHiddenGridlines(true);
    sh.setTabColor(THEME.tabs.settings);
    placeTutorial_(sh);
    placeTabBanner_(sh, 3);
  }
}

function getSettings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(TAB.settings);
  const out = {};
  const vals = sh ? sh.getRange(2, 2, SETTINGS.length, 1).getValues() : [];   // one read for all settings
  SETTINGS.forEach((s, i) => {
    const v = vals[i] ? vals[i][0] : '';
    out[s[0]] = (v === '' || v === null) ? s[2] : v;
  });
  const h24 = (v, dflt) => { const i = HOURS.indexOf(String(v)); return i === -1 ? dflt : i; };
  out.hour24 = h24(out.hour, 7);
  out.remHour24 = h24(out.remHour, 8);
  const dayDefault = { day: 'Monday', remDay: 'Friday', dueDay: 'Sunday', mOpenDay: 'Friday', mDueDay: 'Sunday' };
  Object.keys(dayDefault).forEach(k => { if (DAYS.indexOf(String(out[k])) === -1) out[k] = dayDefault[k]; });
  if (DAYS.indexOf(String(out.fuDay)) === -1 && out.fuDay !== 'No follow-up') out.fuDay = 'Saturday';
  out.quietDays = Number(out.quietDays) || 14;
  out.regionRaw = out.region;                         // what the cell says, e.g. "Region XI – Davao Region"
  const reg = regionProfile_(out);
  out.region = reg.label;                            // short name used in texts: "Region XI"
  out.regionName = reg.name;                         // "Davao Region"
  out.mDueDay = out.dueDay;                          // ONE deadline for everyone (the row below it just shows the same day)
  if (remindTwoDays_()) {                            // website check-in: ONE reminder, 2 calendar days before the deadline, no follow-up
    out.remDay = DAYS[(DAYS.indexOf(out.dueDay) + 5) % 7];
    out.fuDay = 'No follow-up';
  }
  return out;
}
// Website check-in: the only check-in email is one reminder, exactly 2 calendar days before the deadline (Asia/Manila),
// and only to people who have not answered that week. Nothing on other days, nothing after the deadline.
function remindTwoDays_() { return typeof WEB_CHECKIN !== 'undefined' && !!WEB_CHECKIN; }
// Calendar days from today (Manila) to the deadline day
function daysToDeadline_(deadline, now) {
  const tz = CONFIG.timezone, k = d => Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  return Math.round((new Date(k(deadline) + 'T00:00:00Z') - new Date(k(now || new Date()) + 'T00:00:00Z')) / 86400000);
}


// ━━━━━━━━━━━━━━━━━━━━  Helpers  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  HELPERS
 * ================================================================ */

// ── Library mode ──────────────────────────────────────────────────
// The tracker's code can live in ONE shared Apps Script library. Each tracker then has only a small
// "START" file that calls connectTracker(...) first, handing over THIS spreadsheet's own settings
// store and lock. That keeps every region's tracker separate even though they share the code.
// Without the library (all code inside the tracker) nothing is handed over and Google's own is used.
var LIB_CTX_ = null;
var STUB_NEEDED = 3;                    // raise when the small START file gets new functions
function connectTracker(ctx) { LIB_CTX_ = ctx || null; return typeof START_MODULES === 'undefined' ? '' : START_MODULES; }
function docProps_() { return LIB_CTX_ && LIB_CTX_.props ? LIB_CTX_.props : PropertiesService.getDocumentProperties(); }
// Web apps have no containing-document lock, even when they open a spreadsheet by ID.
// Use the shared script lock there so website writes still run one at a time.
function docLock_() {
  return (LIB_CTX_ && LIB_CTX_.lock && LIB_CTX_.lock()) || LockService.getDocumentLock() || LockService.getScriptLock();
}
function libMode_() { return !!(LIB_CTX_ && LIB_CTX_.props); }
function stubVersion_() { return LIB_CTX_ && LIB_CTX_.stub ? Number(LIB_CTX_.stub) : 0; }

// Finds each answer's column in Weekly Check-in by its question title (safe even if questions were added later).
function ckCols_(ck) {
  const last = Math.max(ck ? ck.getLastColumn() : 0, 2);
  const all = ck ? ck.getRange(1, 1, 1, last).getValues()[0].map(x => String(x).trim()) : [];
  // form columns end at the first blank heading or the first of the tracker's extra columns
  const stop = all.findIndex((x, i) => i > 0 && (x === '' || CK_EXTRA.indexOf(x) !== -1));
  const w = Math.max(stop === -1 ? last : stop, 2);
  const h = all.slice(0, w);
  const f = t => h.indexOf(t);
  const school = f(Q.school);
  if (school === -1) throw new Error('The ' + TAB.checkins + ' tab has no "School" column. Click ① Set up / fix the tracker again.');
  const either = (t, old) => { const i = typeof colOf_ === 'function' ? colOf_(h, t) : f(t); return i !== -1 || !old ? i : f(old); };   // newest wording first, then older ones
  const O = typeof OLD_Q === 'undefined' ? {} : OLD_Q;
  return { school, name: f(Q.name), events: either(Q.events, O.events), dc: either(Q.dc, O.dc), help: f(Q.help), concern: either(Q.concern),
           vibe: typeof LQ_VIBE === 'undefined' ? -1 : either(LQ_VIBE), shared: typeof LQ_SHARED === 'undefined' ? -1 : either(LQ_SHARED), width: w };
}

function colLetter_(n) {      // 1 -> A, 27 -> AA
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s;
}

function getForm_() {
  const id = docProps_().getProperty('FORM_ID');
  if (!id) throw new Error('SETUP_FIRST');
  const form = retry_(() => FormApp.openById(id));
  // In a copied tracker the saved form may belong to the original: treat that as "not set up yet"
  const dest = destId_(form);
  if (dest && dest !== SpreadsheetApp.getActive().getId()) throw new Error('SETUP_FIRST');
  return form;
}

function retry_(fn, tries) {
  tries = tries || 5;
  for (let i = 1; ; i++) {
    try { return fn(); }
    catch (e) {
      if (i >= tries) throw e;
      Logger.log('Retrying after: ' + e.message);
      Utilities.sleep(2000 * i);
    }
  }
}

function addOnce_(form, title, addFn) {
  const has = () => retry_(() => form.getItems()).some(it => it.getTitle() === title);
  if (has()) return;
  try { addFn(); }
  catch (e) {
    Utilities.sleep(3000);
    if (!has()) retry_(addFn);
  }
}

// A tab name as it must be written inside a formula: 'Weekly Check-in'!A:A
function ref_(tab) { return "'" + String(tab).replace(/'/g, "''") + "'"; }

// Tabs renamed in newer versions (e.g. Checkins -> Weekly Check-in). Formulas follow a renamed tab by themselves.
function migrateTabs_(ss) {
  // older name -> current name (kept here so this works even if 00_Config is an older copy)
  [['Checkins', TAB.checkins], ['Weekly Check-in', TAB.checkins], ['Checkins (old)', TAB.checkins + ' (old)'], ['Weekly Check-in (old)', TAB.checkins + ' (old)'], ['Focal persons', TAB.focal]].forEach(([from, to]) => {
    if (from === to) return;
    const s = ss.getSheetByName(from);
    if (s && !ss.getSheetByName(to)) s.setName(to);
  });
}

function getOrCreate_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function header_(range, color) {
  range.setBackground(color || THEME.navy).setFontColor('#ffffff').setFontWeight('bold')
    .setFontFamily(THEME.font).setVerticalAlignment('middle').setWrap(true)
    .setBorder(null, null, true, null, null, null, THEME.sun, SpreadsheetApp.BorderStyle.SOLID_THICK);  // deck yellow accent
}

// Rows inserted right under a header copy the header's navy look; this turns them back into normal body rows.
function bodyRows_(range) {
  range.setBackground(null).setFontColor(THEME.navy).setFontWeight('normal').setFontStyle('normal')
       .setFontFamily(THEME.font).setFontSize(10).setWrap(false)
       .setBorder(true, true, true, true, true, true, TABLE_LINE, SpreadsheetApp.BorderStyle.SOLID);
  return range;
}

function band_(range, first, second) {
  range.getSheet().getBandings().forEach(b => b.remove());
  range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false)
    .setFirstRowColor(first || '#ffffff').setSecondRowColor(second || THEME.band);
}

/* ---------- the ONE reader for the Schools tab (every module uses this) ---------- */

// Every school row that has a name: { row, name, focal, contact, email, inactive, link }
function schoolRows_(ss) {
  const sh = ss.getSheetByName(TAB.schools);
  if (!sh) return [];
  return sh.getRange(2, 1, CONFIG.maxSchools, LINK_COL).getValues().map((r, i) => ({
    row: i + 2, name: String(r[0]).trim(), focal: String(r[3]).trim(), contact: String(r[4]).trim(),
    email: emailOf_(r[4]), inactive: String(r[5]).trim().toUpperCase() === 'Y', link: String(r[LINK_COL - 1]).trim(),
  })).filter(s => s.name);
}

// Scholars A:E for every row with a name (name, school, program, awarded, contact)
function scholarRows_(ss) {
  const sc = ss.getSheetByName(TAB.scholars);
  if (!sc || sc.getLastRow() < 2) return [];
  return sc.getRange(2, 1, sc.getLastRow() - 1, 5).getValues().filter(r => String(r[0]).trim());
}

// First email address found in a cell ("" if none). A phone number next to it is ignored.
function emailOf_(v) {
  const m = String(v || '').match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/);
  return m ? m[0].toLowerCase() : '';
}

function orderTabs_(ss) {
  const names = DB_MODE
    ? [TAB.settings, TAB.schools, TAB.scholars, 'Join requests', 'Join - Pending', 'Join - Approved', 'Join - Declined',
       TAB.mcheckins, 'Daily games', 'Lead actions', 'Rewards', 'Reward requests', TAB.checkins, 'Event attendance', TAB.events, TAB.outreach, 'Data change log', 'Weekly roster']
    : [TAB.start, TAB.dashboard, TAB.links, TAB.schools, TAB.focal, TAB.outreach, TAB.calview,
       TAB.events, TAB.scholars, TAB.checkins, TAB.mcheckins, TAB.log, TAB.settings];
  names.filter(name => ss.getSheetByName(name)).forEach((name, i) => {
    ss.setActiveSheet(ss.getSheetByName(name));
    ss.moveActiveSheet(i + 1);
  });
}

// Tabs that are not used in database mode (their information is on the website)
function removedTabs_() { return [TAB.start, TAB.links, TAB.dashboard, TAB.calview, TAB.focal, TAB.log]; }
// Term start date: Settings › Term start date (older trackers: Dashboard › B3)
function termStart_(ss) {
  try { const v = getSettings_().termStart; if (v instanceof Date) return v; } catch (e) { /* no setting yet */ }
  try { const d = (ss || SpreadsheetApp.getActive()).getSheetByName(TAB.dashboard); const v = d && d.getRange('B3').getValue(); if (v instanceof Date) return v; } catch (e) {}
  return null;
}

/* ---------- pop-ups: one design for every box (same look as the emails and the website) ---------- */
var DLG_CSS = '<style>' +
  '*{box-sizing:border-box}body{margin:0;font-family:Poppins,"Segoe UI",Arial,sans-serif;color:#12203A;font-size:13px;line-height:1.55}' +
  '.dbar{height:5px;border-radius:5px;background:linear-gradient(90deg,#0099FF 0%,#0099FF 55%,#FFDD00 100%);margin:0 0 14px}' +
  '.lead{font-size:15px;font-weight:800;margin:0 0 4px;line-height:1.35}.sub{color:#4B5B77;font-size:12px;margin:0 0 12px}' +
  '.muted{color:#4B5B77;font-size:12px}.center{text-align:center}.big{font-size:38px;line-height:1;margin:6px 0 10px}' +
  'p{margin:0 0 10px}b{font-weight:700}a{color:#0077CC;font-weight:600}' +
  '.btn{display:inline-block;background:#FFDD00 linear-gradient(90deg,#FFDD00 0%,#FFF3A0 50%,#7CCBFF 100%);color:#12203A;font:800 12px Poppins,Arial,sans-serif;' +
    'letter-spacing:.8px;text-transform:uppercase;padding:11px 22px;border-radius:999px;border:0;border-bottom:3px solid #0099FF;text-decoration:none;cursor:pointer}' +
  '.btn.dark{background:#12203A;color:#FFDD00;border-bottom-color:#0099FF}.btn[disabled]{opacity:.6;cursor:default}' +
  '.link{background:none;border:0;padding:0;color:#0077CC;font:600 12px Poppins,Arial,sans-serif;cursor:pointer}' +
  'label{display:block;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:#4B5B77;margin:12px 0 5px}' +
  'input,select,textarea{width:100%;font:13px Poppins,Arial,sans-serif;color:#12203A;padding:9px 11px;background:#EAF3FC;border:1px solid #D6E4F5;border-radius:10px;outline:none}' +
  'input:focus,select:focus,textarea:focus{border-color:#0099FF;background:#fff}textarea{resize:vertical;line-height:1.45}' +
  'textarea.copy{background:#FFF6BF;border-color:#F5E27A}' +
  '.row{display:flex;gap:8px;align-items:center;flex-wrap:wrap}.right{margin-left:auto}' +
  '.box{background:#EAF3FC;border:1px solid #D6E4F5;border-radius:12px;padding:10px 12px;margin:0 0 10px}' +
  '.steps{list-style:none;counter-reset:s;padding:0;margin:0 0 10px}.steps li{counter-increment:s;position:relative;padding:4px 0 6px 32px}' +
  '.steps li::before{content:counter(s);position:absolute;left:0;top:3px;width:22px;height:22px;border-radius:50%;background:#12203A;color:#FFDD00;' +
    'font-weight:800;font-size:11px;display:flex;align-items:center;justify-content:center}' +
  'ul.dots{padding-left:18px;margin:0 0 10px}ul.dots li{margin:2px 0}' +
  '.chips{display:flex;gap:6px;flex-wrap:wrap}.chips button{background:#EAF3FC;border:1px solid #D6E4F5;border-radius:999px;padding:6px 12px;' +
    'font:600 12px Poppins,Arial,sans-serif;color:#12203A;cursor:pointer}.chips button.on{background:#FFDD00;border-color:#FFDD00}' +
  '.msg{font-size:12px;margin-top:8px;min-height:16px}.ok{color:#1E7A3C}.err{color:#B3261E}' +
  '</style>';

// A pop-up in the START style. body = HTML; w/h in pixels.
function dlg_(title, body, w, h) {
  ui_().showModalDialog(HtmlService.createHtmlOutput(DLG_CSS + '<div class="dbar"></div>' + body).setWidth(w || 440).setHeight(h || 300), title);
}
// "Done" box: big icon, one line, an optional button, an optional small note
function dlgDone_(title, o) {
  dlg_(title, '<div class="center"><div class="big">' + (o.icon || '✅') + '</div><p class="lead">' + esc_(o.head) + '</p>' +
    (o.text ? '<p class="sub">' + esc_(o.text) + '</p>' : '') +
    (o.link ? '<p><a class="btn" href="' + esc_(o.link) + '" target="_blank">' + esc_(o.button || 'Open') + ' ↗</a></p>' : '') +
    (o.note ? '<p class="muted">' + esc_(o.note) + '</p>' : '') + '</div>', o.w || 400, o.h || (o.link ? 270 : 220));
}
// Copy box: one line, the text to copy, a Copy button
function dlgCopy_(title, lead, text, extra, h) {
  dlg_(title, '<p class="lead">' + esc_(lead) + '</p>' + (extra || '') +
    '<textarea id="cp" class="copy" style="height:' + (h || 210) + 'px">' + esc_(text) + '</textarea>' +
    '<div class="row" style="margin-top:10px"><button class="btn" id="cb" onclick="var t=document.getElementById(\'cp\');t.select();' +
    'try{navigator.clipboard.writeText(t.value)}catch(e){}try{document.execCommand(\'copy\')}catch(e){}this.textContent=\'Copied ✓\'">Copy message</button></div>',
    460, (h || 210) + 170);
}

// Plain text -> the START pop-up: first line = headline, "• " lines = list, "1. " lines = steps, menu paths in bold.
function notifyHtml_(msg) {
  const fmt = t => esc_(t).replace(/((?:START Tracker|🛠|⏰|✉️|📋|📅|📄|🌐|①|🩺|🔐)[^.\n;]*?›[^.,;\n]*)/g, '<b>$1</b>');
  const blocks = String(msg || '').trim().split(/\n\s*\n/);
  let html = '';
  blocks.forEach((b, i) => {
    const lines = b.split('\n').map(l => l.trim()).filter(Boolean);
    if (!lines.length) return;
    if (i === 0) { html += '<p class="lead">' + fmt(lines.shift()) + '</p>'; if (!lines.length) return; }
    if (lines.every(l => /^[•\-✅⚠️❌⏳]\s?/.test(l) || /^[•\-]/.test(l))) {
      html += '<ul class="dots">' + lines.map(l => '<li>' + fmt(l.replace(/^[•\-]\s*/, '')) + '</li>').join('') + '</ul>';
    } else if (lines.every(l => /^\d+[.)]\s/.test(l))) {
      html += '<ol class="steps">' + lines.map(l => '<li>' + fmt(l.replace(/^\d+[.)]\s*/, '')) + '</li>').join('') + '</ol>';
    } else html += lines.map(l => '<p>' + fmt(l) + '</p>').join('');
  });
  return html;
}
function notify_(msg) {
  Logger.log(msg);
  try {
    SpreadsheetApp.getUi();
    const n = String(msg || '').split('\n').length, len = String(msg || '').length;
    dlg_('START Tracker', notifyHtml_(msg) + '<div style="text-align:right;margin-top:6px"><button class="btn dark" onclick="google.script.host.close()">OK</button></div>',
         460, Math.min(560, 150 + n * 20 + Math.floor(len / 60) * 12));
    return;
  } catch (e) { /* no pop-ups here: try the plain box, then a corner note */ }
  try { SpreadsheetApp.getUi().alert('START Tracker', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { try { SpreadsheetApp.getActive().toast(String(msg).split('\n')[0], 'START Tracker', 10); } catch (e2) {} }
}

function esc_(v) {
  return String(v === undefined || v === null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// Dialog boxes only work when a function is started from the START Tracker menu in the Sheet.
// Run from the Apps Script editor or a trigger, this explains that instead of a technical error.
function ui_() {
  try { return SpreadsheetApp.getUi(); }
  catch (e) {
    throw new Error('Please run this from the Google Sheet: START Tracker menu (not from the Apps Script editor). ' +
                    'Nothing was changed.');
  }
}

// "Malasarte, Fe Anne L." -> "Fe Anne L. Malasarte"  (names typed normally are left as they are)
function displayName_(n) {
  n = String(n || '').trim();
  const m = n.match(/^([^,]+),\s*(.+)$/);
  if (!m) return n;
  let first = m[2].trim(), suffix = '';
  const sx = first.match(/\s+(Jr\.?|Sr\.?|II|III|IV|V)$/i);       // keep suffixes at the very end
  if (sx) { suffix = ' ' + sx[1]; first = first.slice(0, sx.index); }
  return (first + ' ' + m[1].trim() + suffix).replace(/\s+/g, ' ').trim();
}

/* ---------- ONE tracker = ONE RR ----------
 * Each RR has their own tracker (🧳 blank copy) and so their own forms, portals, calendars and emails.
 * If a tracker is copied (File › Make a copy), the copy must never keep using the original RR's forms,
 * portals, calendars, folders or email list: they are forgotten here and ① makes new ones for the new owner. */
var TRACKER_KEY = 'TRACKER_ID';
function checkTrackerCopy_() {
  try {
    const ss = SpreadsheetApp.getActive(), props = docProps_();
    const id = ss.getId(), was = props.getProperty(TRACKER_KEY);
    if (!was) { props.setProperty(TRACKER_KEY, id); return false; }      // first run: this tracker is stamped
    if (was === id) return false;
    props.deleteAllProperties();                                          // the original RR's links, ledger, portals…
    props.setProperty(TRACKER_KEY, id);
    props.setProperty('COPIED_AT', String(Date.now()));
    const set = ss.getSheetByName(TAB.settings);                           // the original RR's own details
    if (set) ['rrName', 'recipients', 'replyTo', 'portalNote'].forEach(k => {
      try { set.getRange(settingCell_(k).split('!')[1]).clearContent(); } catch (e) { /* older Settings tab */ }
    });
    return true;
  } catch (e) { Logger.log('Tracker copy check: ' + e.message); return false; }
}
var COPY_NOTICE = 'This tracker is a COPY of another RR\'s tracker.\n\n' +
  'So it never touches their forms, portals, calendars or emails, those links were removed from this copy, ' +
  'and the previous RR\'s name and email list were cleared in Settings.\n\n' +
  'Next: fill in YOUR name in Settings, then click ① Set up / fix the tracker. It makes your own form, portals and calendars.\n\n' +
  'Tip: to give a tracker to another RR, use 🛠 Setup tools › 🧳 Make a blank copy (it has no scholar data).';

/* ---------- Google's time limit (6 minutes per click / per automatic run) ----------
 * Long jobs check outOfTime_() and stop safely: what is left is done on the next click or the next morning. */
var RUN_STARTED_ = Date.now();
var RUN_BUDGET_MS = 4.5 * 60 * 1000;
function outOfTime_(spareMs) { return Date.now() - RUN_STARTED_ > RUN_BUDGET_MS - (spareMs || 0); }