/* =====================================================================
 *  START SCHOLAR TRACKER  ·  06_Calendar
 *  ---------------------------------------------------------------------
 *  Inside: Events tab, shared START calendars, Calendar view.
 *  To customise: Ready-made event types.
 *
 *  The code has 9 files + Email_Reminder and Web_App (HTML). Keep these exact names:
 *    00_Config, 01_Region, 02_Menu_Setup, 03_Forms, 04_Emails, 05_Portals, 06_Calendar, 07_Look, 08_Tools
 *  To update: open each file, select all (Ctrl+A), paste the new version, Save.
 *  Then in the Google Sheet: START Tracker › ① Set up / fix the tracker.
 *  You never need to run anything from here: use the START Tracker menu in the Sheet.
 * ===================================================================== */

// The ▶ Run button runs the FIRST function of the open file: this one, so it is always safe.
function about_06_Calendar() {
  Logger.log('START Tracker' + (typeof START_MODULES === 'string' ? ' ' + START_MODULES : '') + ', file 06_Calendar. Nothing to run here: open the Google Sheet and use the START Tracker menu.');
}


// ━━━━━━━━━━━━━━━━━━━━  Calendar  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  📅 SHARED CALENDARS
 *   START <region>       all ACTIVE scholars with an email (view)      START events for everyone
 *   START Campus Leads   every Campus Lead (view)                      weekly check-in dates, Lead meetings
 *   START – <school>     that school's active scholars (view),         the school's own reminders/events,
 *                        its Campus Lead (EDIT)                        added by the Lead in Google Calendar
 *   My START work        only you                                      Outreach follow-ups, target deadlines
 *
 *  You type events in the "Events" tab; they go to the right calendar by themselves.
 *  Everything belongs to the tracker's email account (20_Safety). Needs Google's "Calendar" service
 *  turned on once in Apps Script (Services › + › Google Calendar API). Sharing follows the Scholars tab
 *  every morning: new active scholars get access, graduated/dropped scholars lose it.
 *  Calendar shares and invites come from Google Calendar itself, not from the tracker's daily email limit.
 * ================================================================ */

var EV_HEAD = ['Date', 'Start', 'End', 'Title', 'Type', 'Who', 'Repeats', 'Until', 'Place or link', 'Details', 'Invite Leads?',
               'On calendar', 'Leads\' answers', 'Members joined', 'ID'];
var EV = { date: 0, start: 1, end: 2, title: 3, type: 4, who: 5, rep: 6, until: 7, place: 8, details: 9, invite: 10,
           status: 11, rsvp: 12, joined: 13, id: 14 };              // 0-based positions in a row (column = position + 1)
var EV_ROWS = 300;
var EV_EDIT = 11;                                 // A–K you type · L–N automatic · O hidden
var EV_REPEATS = ['Never', 'Every week', 'Every 2 weeks', 'Every month'];
var EV_TYPES = ['Meeting', 'Webinar', 'Activity', 'Deadline', 'Other'];
var EV_ICON = { Meeting: '👥', Webinar: '💻', Activity: '🚀', Deadline: '⏰', Other: '📌' };
var WHO = { all: 'All members', leads: 'Campus Leads', me: 'Only me' };
var CAL_NAMES = { get region() { const l = (() => { try { return regionProfile_().label; } catch (e) { return ''; } })(); return 'START ' + (l || 'Scholars'); },
                  leads: 'START Campus Leads', mine: 'My START work', school: n => 'START – ' + n };
var CAL_NEW_PER_RUN = 5;                          // Google limits how many calendars can be made per day
var CAL_BUDGET_MS = 4 * 60 * 1000;

/* ---------- where things are stored ---------- */

function calStore_() {
  try { return Object.assign({ schools: {} }, JSON.parse(docProps_().getProperty('CALENDARS') || '{}')); }
  catch (e) { return { schools: {} }; }
}
function saveCalStore_(c) { docProps_().setProperty('CALENDARS', JSON.stringify(c)); }
function calOn_() { return docProps_().getProperty('CAL_ON') === '1'; }
// The Google Calendar API service, whatever name it was given when it was added (usually "Calendar")
var CAL_API_ = null;
function calApi_() {
  if (CAL_API_) return CAL_API_;
  try { if (typeof Calendar !== 'undefined' && Calendar && Calendar.Events && Calendar.Acl) return (CAL_API_ = Calendar); } catch (e) { /* not added */ }
  const g = (function () { return this; })() || {};
  const name = Object.getOwnPropertyNames(g).find(k => {
    try { const v = g[k]; return v && typeof v === 'object' && v.Events && v.Acl && v.Calendars && v.CalendarList; } catch (e) { return false; }
  });
  if (!name) throw new Error('CAL_SERVICE_OFF');
  return (CAL_API_ = g[name]);
}
function calServiceOn_() { try { return !!calApi_(); } catch (e) { return false; } }
function calLink_(id) { return 'https://calendar.google.com/calendar/u/0/r?cid=' + encodeURIComponent(id); }

/* ---------- menu ---------- */

// 📅 Events & calendar › Set up / update calendars
function setupCalendars() {
  requireSetup_();
  if (!calServiceOn_()) return showCalendarServiceHelp_();
  if (!ensureOwner_('run the shared calendars')) return;
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(TAB.events)) buildEvents_(ss);
  docProps_().setProperty('CAL_ON', '1');
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'refreshCalendarsAuto')) {
    ScriptApp.newTrigger('refreshCalendarsAuto').timeBased().everyDays(1).atHour(6).inTimezone(CONFIG.timezone).create();
  }
  const r = runCalendarSync_(ss);
  try { writeAutoLinks_(ss); } catch (e) { Logger.log('Links: ' + e.message); }   // calendar links in the Links tab
  const c = calStore_();
  notify_('📅 Calendars updated\n\n' +
    `• ${r.calendars} calendar${r.calendars === 1 ? '' : 's'} ready` + (r.waitingCalendars ? ` · ${r.waitingCalendars} more over the next days` : '') + '\n' +
    `• Shared with ${r.members} members, ${r.leads} Leads` + (r.shareFailed.length ? `\n• Not shared: ${r.shareFailed.slice(0, 5).join(', ')}${r.shareFailed.length > 5 ? '…' : ''}` : '') + '\n' +
    `• ${r.events} event${r.events === 1 ? '' : 's'}` + (r.eventProblems ? ` · ${r.eventProblems} to fix (Events tab)` : '') +
    (r.blocked ? '\n\n⚠️ Your school account can\'t share outside the school. Use a Gmail: ⏰ Automatic emails › 🔑.' : '') +
    '\n\nNext: ✉️ Send emails now › 📨 Send welcome emails (members get the calendar link).');
  ss.setActiveSheet(ss.getSheetByName(TAB.events));
}

function stopCalendarUpdates() {
  docProps_().deleteProperty('CAL_ON');
  const n = removeTriggers_(['refreshCalendarsAuto']);
  notify_(n ? 'Daily calendar updates are OFF. The calendars stay as they are; nothing new is added or shared.'
            : 'Daily calendar updates were not on.');
}

function refreshCalendarsAuto(e) {
  if (!ownerGuard_(e) || !calOn_() || !calServiceOn_()) return;
  try { runCalendarSync_(SpreadsheetApp.getActive()); } catch (err) { Logger.log('Calendar refresh: ' + err.message); }
}

function showCalendarServiceHelp_() {
  dlg_('📅 Calendar not connected', `<p class="lead">Add Google's Calendar service once</p>
    <ol class="steps">
      <li><b>Extensions › Apps Script</b> (from this sheet)</li>
      <li><b>Services</b> › <b>+</b> › <b>Google Calendar API</b> › <b>Add</b></li>
      <li>Save, reload this sheet</li>
      <li><b>${MENU_CAL}</b> › <b>Allow</b></li>
    </ol>`, 420, 300);
}

/* ---------- the whole update (menu, every morning) ---------- */

function runCalendarSync_(ss) {
  const started = Date.now(), timeLeft = () => Date.now() - started < CAL_BUDGET_MS;
  const res = { calendars: 0, waitingCalendars: 0, members: 0, leads: 0, shareFailed: [], blocked: false, events: 0, eventProblems: 0 };
  const people = calPeople_(ss);
  const c = ensureCalendars_(people, res);
  saveCalStore_(c);
  res.calendars = ['region', 'leads', 'mine'].filter(k => c[k]).length + Object.keys(c.schools).length;
  try { syncCheckinSeries_(c); } catch (e) { Logger.log('Check-in dates: ' + e.message); }
  try { syncMemberSeries_(c); } catch (e) { Logger.log('Member check-in dates: ' + e.message); }
  try { syncMine_(ss, c); } catch (e) { Logger.log('My START work: ' + e.message); }
  const ev = syncEvents_(ss, c, people);
  try { renderCalView_(ss); } catch (e) { /* view not built yet */ }
  res.events = ev.ok; res.eventProblems = ev.problems;
  // sharing last: it is the slowest part, and it simply continues tomorrow if time runs out
  shareAll_(c, people, res, timeLeft);
  docProps_().setProperty('CAL_LAST', JSON.stringify(
    { at: Date.now(), failed: res.shareFailed.length, blocked: res.blocked, waiting: res.waitingCalendars }));
  return res;
}

// Who should see what: from Scholars (ACTIVE + email) and Schools (Campus Leads)
function calPeople_(ss) {
  const out = { members: {}, bySchool: {}, leads: {}, leadOf: {}, schools: [] };
  const sh = ss.getSheetByName(TAB.scholars);
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().forEach(r => {
      const email = emailOf_(r[4]), school = String(r[1]).trim();
      if (!email || String(r[5]).trim() !== 'Active') return;
      out.members[email] = true;
      (out.bySchool[school] = out.bySchool[school] || {})[email] = true;
    });
  }
  schoolRows_(ss).forEach(s => {
    out.schools.push(s.name);
    if (s.focal && s.email) { out.leads[s.email] = true; out.leadOf[s.name] = s.email; }
  });
  return out;
}

// Makes any calendar that is missing (school calendars only for schools with members or a Lead)
function ensureCalendars_(people, res) {
  const c = calStore_();
  const exists = id => { if (!id) return false; try { calApi_().Calendars.get(id); return true; } catch (e) { return false; } };
  const make = (name, desc) => calApi_().Calendars.insert({ summary: name, description: desc, timeZone: CONFIG.timezone }).id;
  const regionDesc = 'START events for all active scholars' + (regionProfile_().label ? ' in ' + regionProfile_().label : '') + '. Managed by the START Tracker.';
  if (!exists(c.region)) c.region = make(CAL_NAMES.region, regionDesc);
  else {                                                       // region changed in Settings: the calendar is renamed, same people
    try { const cal = calApi_().Calendars.get(c.region); if (cal.summary !== CAL_NAMES.region) calApi_().Calendars.patch({ summary: CAL_NAMES.region, description: regionDesc }, c.region); }
    catch (e) { Logger.log('Region calendar name: ' + e.message); }
  }
  if (!exists(c.leads)) c.leads = make(CAL_NAMES.leads, 'Weekly check-in dates and meetings for START Campus Leads. Managed by the START Tracker.');
  if (!exists(c.mine)) c.mine = make(CAL_NAMES.mine, 'Private: your follow-ups and target deadlines from the START Tracker.');
  let made = 0;
  people.schools.forEach(name => {
    const needed = people.leadOf[name] || Object.keys(people.bySchool[name] || {}).length;
    if (!needed) return;
    if (exists(c.schools[name])) return;
    if (made >= CAL_NEW_PER_RUN) { res.waitingCalendars++; return; }
    try {
      c.schools[name] = make(CAL_NAMES.school(name), `START reminders and events for ${name}. The school's Campus Lead can add events here.`);
      made++;
    } catch (e) { res.waitingCalendars++; Logger.log('School calendar ' + name + ': ' + e.message); }
  });
  return c;
}

/* ---------- sharing (follows Scholars and Schools) ---------- */

function shareAll_(c, people, res, timeLeft) {
  const me = me_();
  const want = {};                                                    // calendarId -> { email: role }
  const add = (id, email, role) => { if (id && email && email !== me) (want[id] = want[id] || {})[email] = role; };
  Object.keys(people.members).forEach(e => add(c.region, e, 'reader'));
  Object.keys(people.leads).forEach(e => { add(c.region, e, 'reader'); add(c.leads, e, 'reader'); });
  Object.keys(c.schools).forEach(name => {
    const id = c.schools[name];
    Object.keys(people.bySchool[name] || {}).forEach(e => add(id, e, 'reader'));
    if (people.leadOf[name]) add(id, people.leadOf[name], 'writer');   // the Lead edits their school's calendar
    want[id] = want[id] || {};
  });
  want[c.region] = want[c.region] || {}; want[c.leads] = want[c.leads] || {};
  res.members = Object.keys(people.members).length;
  res.leads = Object.keys(people.leads).length;

  Object.keys(want).forEach(id => {
    if (!timeLeft()) { res.incomplete = true; return; }
    let rules = [];
    try { rules = (calApi_().Acl.list(id).items || []).filter(r => r.scope && r.scope.type === 'user' && r.role !== 'owner'); }
    catch (e) { res.incomplete = true; Logger.log('Sharing list: ' + e.message); return; }
    const have = {};
    rules.forEach(r => { have[String(r.scope.value).toLowerCase()] = r; });
    Object.keys(want[id]).forEach(email => {
      if (!timeLeft()) { res.incomplete = true; return; }
      const role = want[id][email], r = have[email];
      try {
        if (!r) calApi_().Acl.insert({ role, scope: { type: 'user', value: email } }, id,
                                { sendNotifications: false });   // no Google email: your welcome email (👥 › 📨) carries the calendar link
        else if (r.role !== role) calApi_().Acl.patch({ role }, id, r.id);
      } catch (e) {
        res.shareFailed.push(email);
        if (/domain|outside|not allowed|forbidden/i.test(e.message)) res.blocked = true;
      }
    });
    Object.keys(have).forEach(email => {                               // no longer active / no longer the Lead
      if (want[id][email]) return;
      if (!timeLeft()) { res.incomplete = true; return; }
      try { calApi_().Acl.remove(id, have[email].id); } catch (e) { res.incomplete = true; Logger.log('Unshare ' + email + ': ' + e.message); }
    });
  });
  res.shareFailed = [...new Set(res.shareFailed)];
}

/* ---------- weekly check-in dates (Campus Leads calendar) ---------- */

function syncCheckinSeries_(c) {
  const st = getSettings_(), props = docProps_();
  const key = [c.leads, st.remDay, st.dueDay, st.fuDay, st.autoWeekly].join('|');
  let old = {};
  try { old = JSON.parse(props.getProperty('CAL_SERIES') || '{}'); } catch (e) { old = {}; }
  if (old.key === key) return;
  (old.ids || []).forEach(id => { try { calApi_().Events.remove(old.cal || c.leads, id); } catch (e) { /* already gone */ } });
  const ids = [];
  if (st.autoWeekly === true) {
    const form = checkinUrl_();
    const series = [[st.remDay, '📨 START weekly check-in: reminder arrives today', 'Your school\'s check-in form arrives by email today.'],
                    [st.dueDay, '⏰ START weekly check-in due today', 'Please send your school\'s weekly check-in today. It takes about 2 minutes.']];
    if (st.fuDay !== 'No follow-up') series.push([st.fuDay, '🔁 START follow-up for schools that have not answered', 'A follow-up email goes only to schools that have not answered yet.']);
    series.forEach(([day, title, desc]) => {
      const first = nextWeekday_(day), next = new Date(first.getTime() + 86400000);
      const ev = calApi_().Events.insert({
        summary: title, description: desc + (form ? '\n\nCheck-in form: ' + form : ''), transparency: 'transparent',
        start: { date: ymd_(first) }, end: { date: ymd_(next) },
        recurrence: ['RRULE:FREQ=WEEKLY;BYDAY=' + day.slice(0, 2).toUpperCase()],
      }, c.leads);
      ids.push(ev.id);
    });
  }
  props.setProperty('CAL_SERIES', JSON.stringify({ key, cal: c.leads, ids }));
}

// Member check-in days on the START <region> calendar (all members see them), once the member form exists
function syncMemberSeries_(c) {
  const st = getSettings_(), props = docProps_();
  const on = WEB_CHECKIN ? checkinOn_() : !!props.getProperty('MFORM_ID');
  const key = [c.region, st.mOpenDay, st.mDueDay, on].join('|');
  let old = {};
  try { old = JSON.parse(props.getProperty('CAL_MSERIES') || '{}'); } catch (e) { old = {}; }
  if (old.key === key) return;
  (old.ids || []).forEach(id => { try { calApi_().Events.remove(old.cal || c.region, id); } catch (e) { /* already gone */ } });
  const ids = [];
  if (on) {
    const link = (() => { try { return WEB_CHECKIN ? checkinLink_() : memberForm_().getPublishedUrl(); } catch (e) { return ''; } })();
    [[st.mOpenDay, '✨ START member check-in is open', 'Your weekly vibe check (30 seconds). Your Campus Lead shares the link.'],
     [st.mDueDay, '⏰ START member check-in due today', 'Last day for this week\'s vibe check.']].forEach(([day, title, desc]) => {
      const first = nextWeekday_(day), next = new Date(first.getTime() + 86400000);
      ids.push(calApi_().Events.insert({ summary: title, description: desc + (link ? '\n\nForm: ' + link : ''), transparency: 'transparent',
        start: { date: ymd_(first) }, end: { date: ymd_(next) },
        recurrence: ['RRULE:FREQ=WEEKLY;BYDAY=' + day.slice(0, 2).toUpperCase()] }, c.region).id);
    });
  }
  props.setProperty('CAL_MSERIES', JSON.stringify({ key, cal: c.region, ids }));
}

/* ---------- My START work: follow-ups and target deadlines (private) ---------- */

function syncMine_(ss, c) {
  const st = getSettings_(), props = docProps_();
  const want = {};
  const put = (date, title, desc) => { if (date instanceof Date) want[ymd_(date) + '|' + title] = { date, title, desc }; };
  const ol = ss.getSheetByName(TAB.outreach);
  if (ol && ol.getLastRow() > 1) {
    ol.getRange(2, 1, ol.getLastRow() - 1, OUTREACH_HEAD.length).getValues().forEach(r => {
      if (String(r[1]).trim()) put(r[6], `📞 Follow up: ${r[1]}`, [r[5] ? 'Next step: ' + r[5] : '', r[2] ? 'Contact: ' + r[2] : ''].filter(Boolean).join('\n'));
    });
  }
  put(st.dFocal, `🎯 Deadline: ${st.tFocal} START Campus Leads named`, 'Target from Settings.');
  put(st.dReeng, `🎯 Deadline: ${st.tReeng} schools re-engaged`, 'Target from Settings.');
  let have = {};
  try { have = JSON.parse(props.getProperty('CAL_MINE') || '{}'); } catch (e) { have = {}; }
  Object.keys(have).forEach(k => {
    if (want[k]) return;
    try { calApi_().Events.remove(c.mine, have[k]); } catch (e) { /* already gone */ }
    delete have[k];
  });
  Object.keys(want).forEach(k => {
    if (have[k]) return;
    const w = want[k];
    try {
      have[k] = calApi_().Events.insert({ summary: w.title, description: w.desc, start: { date: ymd_(w.date) },
        end: { date: ymd_(new Date(w.date.getTime() + 86400000)) },
        reminders: { useDefault: false, overrides: [{ method: 'popup', minutes: 540 }] } }, c.mine).id;
    } catch (e) { Logger.log('My START work: ' + e.message); }
  });
  props.setProperty('CAL_MINE', JSON.stringify(have));
}

/* ---------- Events tab → calendars ---------- */

function buildEvents_(ss) {
  const sh = getOrCreate_(ss, TAB.events);
  const N = EV_HEAD.length;
  if (sh.getMaxRows() < EV_ROWS + 1) sh.insertRowsAfter(sh.getMaxRows(), EV_ROWS + 1 - sh.getMaxRows());
  if (sh.getMaxColumns() < N) sh.insertColumnsAfter(sh.getMaxColumns(), N - sh.getMaxColumns());
  sh.setFrozenRows(1);
  const body = c => sh.getRange(2, c, EV_ROWS, 1);
  sh.getRange(1, 1, 1, N).setValues([EV_HEAD]);
  bodyRows_(sh.getRange(2, 1, EV_ROWS, N));
  const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).setHelpText('Pick a date (double-click for a calendar).').build();
  [EV.date + 1, EV.until + 1].forEach(c => body(c).setNumberFormat('ddd, mmm d, yyyy').setDataValidation(dateRule));
  [EV.start + 1, EV.end + 1].forEach(c => body(c).setNumberFormat('h:mm am/pm'));
  body(EV.type + 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(EV_TYPES, true).setAllowInvalid(true).build());
  body(EV.rep + 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(EV_REPEATS, true).setAllowInvalid(true).build());
  body(EV.invite + 1).insertCheckboxes();
  tableLook_(sh, EV_ROWS, N - 1, 26, THEME.tabs.events);
  band_(sh.getRange(2, 1, EV_ROWS, N - 1));
  body(EV.title + 1).setFontWeight('bold');
  markColumns_(sh, [
    [1, 'edit', 'Date', 'Day of the event (the first one, if it repeats). Clear the whole row to remove it.' + (DB_MODE ? '' : ' Easiest: menu 📅 Events & calendar › Add an event.')],
    [2, 'edit', 'Start', 'e.g. 9:00 AM. Leave Start and End blank for an all-day event.'],
    [3, 'edit', 'End', 'e.g. 11:00 AM. Blank = 1 hour after Start.'],
    [4, 'edit', 'Title', 'What everyone sees on the calendar. No personal data (no student names, phone numbers).', true],
    [5, 'edit', 'Type'],
    [6, 'edit', 'Who', 'All members = all scholars, Campus Leads and RR · Campus Leads = Leads and RR · a school = that school and RR · Only me = RR only.'],
    [7, 'edit', 'Repeats', 'Never, every week, every 2 weeks or every month: ONE row for the whole series.'],
    [8, 'edit', 'Until', 'Last date of a repeating event. Blank = keeps repeating.'],
    [9, 'edit', 'Place or link', 'Venue, or the Zoom / Meet link.', 'link'],
    [10, 'edit', 'Details', 'Shown in the event description.', true],
    [11, 'edit', 'Invite Leads?', 'Tick to send Campus Leads a calendar invite (Yes / No / Maybe). Their answers show in "Leads\' answers".'],
    [12, 'auto', 'On calendar'], [13, 'auto', 'Leads\' answers'],
    [14, 'auto', 'Members joined', 'How many members ticked this event in their weekly member check-in (all dates of a repeating event).'],
  ], EV_ROWS, false);
  sh.hideColumns(N);
  refreshWhoList_(ss, sh);
  const st = body(EV.status + 1);
  sh.setConditionalFormatRules([
    chipRule_(st, '✅', CHIP.green, false), chipRule_(st, '⚠️', CHIP.yellow, false), chipRule_(st, '⏳', CHIP.blue, false),
    chipRule_(st, '📝', CHIP.grey, false),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($A2<>"",IF($H2<>"",$H2,IF(OR($G2="",$G2="Never"),$A2,TODAY()))<TODAY())')
      .setFontColor(THEME.muted).setRanges([sh.getRange(2, 1, EV_ROWS, EV_EDIT)]).build(),   // finished events greyed
  ]);
  sh.getRange(2, EV.status + 1, EV_ROWS, 3).protect().setDescription(PROTECT_TAG).setWarningOnly(true);
  placeTabBanner_(sh, EV_EDIT);
  placeTutorial_(sh);
  return sh;
}

// The "Who" dropdown: the 3 groups + every school
function refreshWhoList_(ss, sh) {
  sh = sh || ss.getSheetByName(TAB.events);
  if (!sh) return;
  sh.getRange(2, EV.who + 1, EV_ROWS, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList([WHO.all, WHO.leads, WHO.me].concat(schoolNames_(ss)), true).setAllowInvalid(true).build());
}

// Which calendar a "Who" value goes to: { cal, guests: [Lead emails to invite] } or { problem }
function calTarget_(who, c, people) {
  const allLeads = Object.keys(people.leads);
  if (who === WHO.all || who === '') return { key: 'region', cal: c.region, guests: allLeads };
  if (who === WHO.leads) return { key: 'leads', cal: c.leads, guests: allLeads };
  if (who === WHO.me) return { key: 'mine', cal: c.mine, guests: [] };
  if (people.schools.indexOf(who) === -1) return { problem: '⚠️ "Who" must be a group or a school from the list' };
  if (!c.schools[who]) return { problem: '⏳ This school\'s calendar is not ready yet (needs members with an email or a Lead)' };
  return { key: 'school:' + who, cal: c.schools[who], guests: people.leadOf[who] ? [people.leadOf[who]] : [] };
}

// "9:00 AM", a time cell, or blank -> minutes after midnight (null = blank)
function minutesOf_(v) {
  if (v instanceof Date) return v.getHours() * 60 + v.getMinutes();
  const m = String(v || '').trim().match(/^(\d{1,2})(?::(\d{2}))?\s*([ap])?\.?\s*m?\.?$/i);
  if (!m) return null;
  let h = Number(m[1]);
  if (m[3]) h = h % 12 + (m[3].toLowerCase() === 'p' ? 12 : 0);   // 12 AM = 0:00, 12 PM = 12:00
  if (h > 23 || Number(m[2] || 0) > 59) return null;
  return h * 60 + Number(m[2] || 0);
}

// The repeat rule for Google Calendar ('' = one-time)
function recurrence_(r, allDay) {
  const rep = String(r[EV.rep] || 'Never');
  const base = { 'Every week': 'FREQ=WEEKLY', 'Every 2 weeks': 'FREQ=WEEKLY;INTERVAL=2', 'Every month': 'FREQ=MONTHLY' }[rep];
  if (!base) return '';
  const u = r[EV.until];
  const until = u instanceof Date
    ? ';UNTIL=' + (allDay ? Utilities.formatDate(u, CONFIG.timezone, 'yyyyMMdd')
                          : Utilities.formatDate(new Date(u.getFullYear(), u.getMonth(), u.getDate(), 23, 59, 59), 'UTC', "yyyyMMdd'T'HHmmss'Z'"))
    : '';
  return 'RRULE:' + base + until;
}

function eventResource_(r, guests) {
  const date = r[EV.date], type = EV_TYPES.indexOf(String(r[EV.type])) !== -1 ? String(r[EV.type]) : 'Other';
  const s = minutesOf_(r[EV.start]), e0 = minutesOf_(r[EV.end]);
  const at = mins => { const d = new Date(date); d.setHours(0, 0, 0, 0); d.setMinutes(mins); return d; };
  let start, end;
  if (s === null) {
    const next = new Date(date); next.setDate(next.getDate() + 1);
    start = { date: ymd_(date) }; end = { date: ymd_(next) };
  } else {
    const e = e0 === null || e0 <= s ? s + 60 : e0;
    start = { dateTime: Utilities.formatDate(at(s), CONFIG.timezone, "yyyy-MM-dd'T'HH:mm:ssXXX"), timeZone: CONFIG.timezone };
    end = { dateTime: Utilities.formatDate(at(e), CONFIG.timezone, "yyyy-MM-dd'T'HH:mm:ssXXX"), timeZone: CONFIG.timezone };
  }
  const rule = recurrence_(r, s === null);
  const res = {
    summary: `${EV_ICON[type]} ${String(r[EV.title]).trim()}`,
    location: String(r[EV.place] || '').trim(),
    description: [String(r[EV.details] || '').trim(), `${type} · START ${getSettings_().region || ''}`.trim()].filter(Boolean).join('\n\n'),
    start, end,
    attendees: guests.map(email => ({ email })),
    guestsCanSeeOtherGuests: false, guestsCanInviteOthers: false,
    reminders: { useDefault: false, overrides: [{ method: 'popup', minutes: 1440 }, { method: 'popup', minutes: 60 }] },
  };
  res.recurrence = rule ? [rule] : [];
  return res;
}

function rsvpText_(attendees) {
  const n = { accepted: 0, declined: 0, tentative: 0, needsAction: 0 };
  (attendees || []).forEach(a => { n[a.responseStatus] = (n[a.responseStatus] || 0) + 1; });
  if (!(attendees || []).length) return '';
  return `${n.accepted} Yes · ${n.declined} No · ${n.tentative} Maybe · ${n.needsAction} no answer`;
}

// Puts every row of the Events tab on its calendar. Only changed rows call Google.
function syncEvents_(ss, c, people) {
  const out = { ok: 0, problems: 0 };
  const sh = ss.getSheetByName(TAB.events);
  if (!sh) return out;
  c = c || calStore_();
  people = people || calPeople_(ss);
  refreshWhoList_(ss, sh);
  const last = Math.max(sh.getLastRow(), 2);
  const rows = sh.getRange(2, 1, last - 1, EV_HEAD.length).getValues();
  const status = [], ids = [];
  rows.forEach(r => {
    let saved = {};
    try { saved = JSON.parse(r[EV.id] || '{}'); } catch (e) { saved = {}; }
    const blank = !String(r[EV.title]).trim() && !(r[EV.date] instanceof Date);
    const drop = () => { if (saved.id) { try { calApi_().Events.remove(saved.cal, saved.id, { sendUpdates: 'all' }); } catch (e) { /* gone */ } } };
    if (blank) { drop(); status.push(['', '']); ids.push(['']); return; }
    if (!(r[EV.date] instanceof Date) || !String(r[EV.title]).trim()) {
      out.problems++; status.push(['⚠️ Add a date and a title', '']); ids.push([r[EV.id]]); return;
    }
    if (r[EV.until] instanceof Date && (String(r[EV.rep] || 'Never') === 'Never')) {
      out.problems++; status.push(['⚠️ "Until" is filled but "Repeats" is empty: pick how often it repeats (or clear Until)', '']); ids.push([r[EV.id]]); return;
    }
    if (r[EV.until] instanceof Date && r[EV.until] < r[EV.date]) {
      out.problems++; status.push(['⚠️ "Until" is before the date', '']); ids.push([r[EV.id]]); return;
    }
    const t = calTarget_(String(r[EV.who]).trim(), c, people);
    if (t.problem) { out.problems++; status.push([t.problem, '']); ids.push([r[EV.id]]); return; }
    const guests = r[EV.invite] === true && t.key !== 'mine' ? t.guests : [];
    const res = eventResource_(r, guests);
    const fp = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify([res, t.cal])));
    try {
      let ev;
      if (saved.id && saved.cal === t.cal && saved.fp === fp) {
        ev = calApi_().Events.get(t.cal, saved.id);                            // unchanged: just read the answers
        if (ev.status === 'cancelled') throw new Error('deleted in Google Calendar');
      } else if (saved.id && saved.cal === t.cal) {
        ev = calApi_().Events.patch(res, t.cal, saved.id, { sendUpdates: guests.length ? 'all' : 'none' });
      } else {
        drop();                                                               // moved to another calendar
        ev = calApi_().Events.insert(res, t.cal, { sendUpdates: guests.length ? 'all' : 'none' });
      }
      out.ok++;
      const where = t.key === 'region' ? 'All members' : t.key === 'leads' ? 'Campus Leads' : t.key === 'mine' ? 'Only you' : t.key.slice(7);
      const rp = String(r[EV.rep] || 'Never') !== 'Never' ? ` · ${String(r[EV.rep]).toLowerCase()}` : '';
      status.push([`✅ On "${where}" calendar${rp}`, guests.length ? rsvpText_(ev.attendees) : '']);
      ids.push([JSON.stringify({ cal: t.cal, id: ev.id, fp })]);
    } catch (e) {
      if (/deleted/.test(e.message)) {                                        // removed in Calendar: put it back next run
        status.push(['⏳ Was deleted in Google Calendar: it comes back on the next update', '']); ids.push(['']);
      } else { out.problems++; status.push(['⚠️ ' + String(e.message).slice(0, 80), '']); ids.push([r[EV.id]]); }
    }
  });
  sh.getRange(2, EV.status + 1, rows.length, 2).setValues(status);
  sh.getRange(2, EV.id + 1, rows.length, 1).setValues(ids);
  if (out.problems) docProps_().deleteProperty('EVENTS_SYNC');
  else docProps_().setProperty('EVENTS_SYNC', eventsKey_(sh));
  return out;
}

// Fingerprint of what you typed (A–K): the sheet-change check only calls Google when it changed
function eventsKey_(sh) {
  if (!sh) return '';
  const last = Math.max(sh.getLastRow(), 2);
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,
    JSON.stringify(sh.getRange(2, 1, last - 1, EV_EDIT).getValues())));
}

// Called by onSheetChange (06_Form) after any edit: updates the calendars when the Events tab changed
function eventsChanged_(ss) {
  if (!calOn_() || !calServiceOn_()) return '';
  const sh = ss.getSheetByName(TAB.events);
  if (!sh || docProps_().getProperty('EVENTS_SYNC') === eventsKey_(sh)) return '';
  const lock = docLock_();                   // two quick edits never make the same event twice
  if (!lock.tryLock(15000)) return '';
  let r;
  try {
    if (docProps_().getProperty('EVENTS_SYNC') === eventsKey_(sh)) return '';
    r = syncEvents_(ss);
  } finally { lock.releaseLock(); }
  return `📅 Calendar updated (${r.ok} event(s)${r.problems ? ', ' + r.problems + ' to fix' : ''}).`;
}

/* ---------- shown elsewhere: Today, reminder emails ---------- */

// Every date a row happens on between `from` and `to` (repeats expanded)
function occurrences_(r, from, to) {
  const first = r[EV.date];
  if (!(first instanceof Date)) return [];
  const rep = String(r[EV.rep] || 'Never');
  const last = r[EV.until] instanceof Date && r[EV.until] < to ? r[EV.until] : to;
  const out = [], day0 = new Date(first.getFullYear(), first.getMonth(), first.getDate());
  if (!{ 'Every week': 1, 'Every 2 weeks': 1, 'Every month': 1 }[rep]) return day0 >= from && day0 < to ? [day0] : [];
  for (let i = 0; i < 400; i++) {
    let d;
    if (rep === 'Every month') {
      d = new Date(day0.getFullYear(), day0.getMonth() + i, day0.getDate());
      if (d.getDate() !== day0.getDate()) continue;                     // e.g. the 31st in a short month
    } else d = new Date(day0.getFullYear(), day0.getMonth(), day0.getDate() + i * (rep === 'Every week' ? 7 : 14));
    if (d > last || d >= to) break;
    if (d >= from) out.push(d);
  }
  return out;
}

// Events from the Events tab between today and `days` ahead (repeats included).
// school = only what that school's Lead sees.
function upcomingEvents_(ss, days, school) {
  const sh = ss.getSheetByName(TAB.events);
  if (!sh || sh.getLastRow() < 2) return [];
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const until = new Date(today.getTime() + days * 86400000);
  const out = [];
  sh.getRange(2, 1, sh.getLastRow() - 1, EV_EDIT).getValues().forEach((r, i) => {
    const title = String(r[EV.title]).trim(), who = String(r[EV.who]).trim() || WHO.all;
    if (!title) return;
    if (school !== undefined && !(who === WHO.all || who === WHO.leads || who === school)) return;
    occurrences_(r, today, until).forEach(d => out.push({ date: d, start: r[EV.start], end: r[EV.end], details: String(r[EV.details] || '').trim(), title, type: String(r[EV.type]), who,
                                                         place: String(r[EV.place]).trim(), row: i + 2 }));
  });
  return out.sort((a, b) => a.date - b.date || (minutesOf_(a.start) || 0) - (minutesOf_(b.start) || 0));
}

function eventWhen_(e) {
  const m = minutesOf_(e.start);
  return Utilities.formatDate(e.date, CONFIG.timezone, 'EEE, MMM d') +
         (m === null ? '' : ' · ' + Utilities.formatDate(new Date(2000, 0, 1, Math.floor(m / 60), m % 60), CONFIG.timezone, 'h:mm a'));
}

// The small "Upcoming START events" block for a Campus Lead's reminder email ('' when nothing is coming up)
function eventsForEmail_(school) {
  if (!calOn_()) return '';
  const list = upcomingEvents_(SpreadsheetApp.getActive(), 14, school);
  const c = calStore_();
  const links = [[c.leads, 'Campus Leads calendar'], [c.schools[school], school + ' calendar'], [c.region, 'All-members calendar']]
    .filter(x => x[0]).map(x => `<a href="${calLink_(x[0])}" style="color:#0099FF;font-weight:700;text-decoration:none;">${esc_(x[1])}</a>`);
  if (!list.length && !links.length) return '';
  return (list.length ? emailSection_('Upcoming START events') +
           emailTable_(['When', 'Event', 'Where'], list.slice(0, 6).map(e => [eventWhen_(e), (EV_ICON[e.type] || '📌') + ' ' + e.title, e.place])) : '') +
         (links.length ? `<p style="margin:10px 0 0 0;font-size:12px;color:#4b5b77;">📅 Add to your Google Calendar: ${links.join(' · ')}</p>` : '');
}

/* ---------- helpers ---------- */

function ymd_(d) { return Utilities.formatDate(d, CONFIG.timezone, 'yyyy-MM-dd'); }

// The next date (today or later) that falls on `day` ('Monday' …)
function nextWeekday_(day) {
  const want = DAYS.indexOf(day), d = new Date();
  d.setHours(0, 0, 0, 0);
  while (((d.getDay() + 6) % 7) !== want) d.setDate(d.getDate() + 1);
  return d;
}

/* ================================================================
 *  📅 ADD AN EVENT  (menu box: no typing in rows)
 *  Pick a ready-made START event or start blank, adjust, Save. It becomes one row in the Events tab
 *  and goes on the calendar right away (when the calendars are on).
 * ================================================================ */

// Ready-made events. They are only starting points: everything can be changed in the box.
// days = how many days from today the date starts at; start/end blank = all day.
var EV_PRESETS = [
  { name: 'Monthly START Campus Leads meeting', type: 'Meeting', who: WHO.leads, rep: 'Every month', start: '18:00', end: '19:00',
    invite: true, days: 7, details: 'Monthly meeting: updates from each school, plans for the coming month, concerns.' },
  { name: 'Campus Leads check-in call (every 2 weeks)', type: 'Meeting', who: WHO.leads, rep: 'Every 2 weeks', start: '19:00', end: '19:30',
    invite: true, days: 3, details: 'Short call to follow up on weekly check-ins and help requests.' },
  { name: 'Orientation for new START members', type: 'Activity', who: WHO.all, rep: 'Never', start: '09:00', end: '12:00',
    invite: true, days: 14, details: 'Welcome and orientation for new START members.' },
  { name: 'START webinar', type: 'Webinar', who: WHO.all, rep: 'Never', start: '14:00', end: '16:00',
    invite: true, days: 10, details: 'Topic, speaker and registration link here.' },
  { name: 'Deadline for Campus Leads', type: 'Deadline', who: WHO.leads, rep: 'Never', start: '', end: '',
    invite: true, days: 7, details: 'What is due, and where to send it.' },
  { name: 'School visit', type: 'Meeting', who: '', rep: 'Never', start: '09:00', end: '11:00',
    invite: true, days: 7, details: 'Visit plan and who to meet.' },
  { name: 'My report to DOST-SEI due', type: 'Deadline', who: WHO.me, rep: 'Never', start: '', end: '',
    invite: false, days: 30, details: 'Tip: ' + MENU_REPORT + ' creates the document.' },
];

function showAddEvent() {
  requireSetup_();
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(TAB.events)) buildEvents_(ss);
  const data = JSON.stringify({ presets: EV_PRESETS, types: EV_TYPES, repeats: EV_REPEATS,
                                who: [WHO.all, WHO.leads, WHO.me].concat(schoolNames_(ss)), today: ymd_(new Date()) });
  const html = `<style>.half{display:flex;gap:8px}.half>div{flex:1}.tick{display:flex;gap:8px;align-items:center;font-size:12px;margin-top:12px}
    .tick input{width:auto}</style>
  <label style="margin-top:0">Start from (optional)</label><select id="preset"><option value="">Blank event</option></select>
  <label>Title</label><input id="title">
  <div class="half"><div><label>Type</label><select id="type"></select></div><div><label>For</label><select id="who"></select></div></div>
  <div class="half"><div><label>Date</label><input id="date" type="date"></div>
    <div><label>Start</label><input id="start" type="time"></div><div><label>End</label><input id="end" type="time"></div></div>
  <div class="muted" style="margin-top:4px">No times = all day</div>
  <div class="half"><div><label>Repeats</label><select id="rep"></select></div><div><label>Until</label><input id="until" type="date"></div></div>
  <label>Place or link</label><input id="place">
  <label>Details</label><textarea id="details" rows="2"></textarea>
  <label class="tick"><input id="invite" type="checkbox"> Invite Campus Leads</label>
  <div style="margin-top:14px"><button class="btn" id="save">Save event</button></div><div class="msg" id="msg"></div>
  <script>
    const D = ${data}, $ = id => document.getElementById(id);
    const fill = (id, list) => list.forEach(v => { const o = document.createElement('option'); o.value = o.textContent = v; $(id).appendChild(o); });
    D.presets.forEach((p, i) => { const o = document.createElement('option'); o.value = i; o.textContent = p.name; $('preset').appendChild(o); });
    fill('type', D.types); fill('who', D.who); fill('rep', D.repeats);
    const plus = n => { const d = new Date(D.today + 'T00:00:00'); d.setDate(d.getDate() + n);
      return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };
    $('date').value = plus(1);
    $('preset').onchange = () => {
      const p = D.presets[$('preset').value]; if (!p) return;
      $('title').value = p.name; $('type').value = p.type; if (p.who) $('who').value = p.who; $('rep').value = p.rep;
      $('start').value = p.start; $('end').value = p.end; $('details').value = p.details; $('invite').checked = p.invite;
      $('date').value = plus(p.days); $('until').value = p.rep === 'Never' ? '' : plus(p.days + 180);
    };
    $('save').onclick = () => {
      const v = { title: $('title').value, type: $('type').value, who: $('who').value, date: $('date').value, start: $('start').value,
                  end: $('end').value, rep: $('rep').value, until: $('until').value, place: $('place').value,
                  details: $('details').value, invite: $('invite').checked };
      if (!v.title.trim() || !v.date) { $('msg').className = 'msg err'; $('msg').textContent = 'Add a title and a date.'; return; }
      $('save').disabled = true; $('msg').textContent = 'Saving…';
      google.script.run.withSuccessHandler(m => { $('msg').textContent = m; $('save').disabled = false; $('title').value = ''; $('preset').value = ''; })
        .withFailureHandler(e => { $('msg').textContent = '⚠️ ' + e.message; $('save').disabled = false; }).addEventFromBox(v);
    };
  </script>`;
  dlg_('📅 Add an event', html, 480, 690);
}

// Called by the box: one new row in the Events tab, then onto the calendar
function addEventFromBox(v, spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(TAB.events) || buildEvents_(ss);
  const day = s => { const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : ''; };
  const time = t => { const m = minutesOf_(t); return m === null ? '' : Utilities.formatDate(new Date(2000, 0, 1, Math.floor(m / 60), m % 60), CONFIG.timezone, 'h:mm a'); };
  const date = day(v.date);
  if (!date || !String(v.title || '').trim()) throw new Error('Please add a title and a date.');
  const row = new Array(EV_EDIT).fill('');
  row[EV.date] = date; row[EV.start] = time(v.start); row[EV.end] = v.start ? time(v.end) : '';
  row[EV.title] = String(v.title).trim().slice(0, 150); row[EV.type] = EV_TYPES.indexOf(v.type) !== -1 ? v.type : 'Other';
  row[EV.who] = v.who || WHO.all; row[EV.rep] = EV_REPEATS.indexOf(v.rep) !== -1 ? v.rep : 'Never';
  row[EV.until] = row[EV.rep] === 'Never' ? '' : day(v.until);
  row[EV.place] = String(v.place || '').trim(); row[EV.details] = String(v.details || '').trim(); row[EV.invite] = v.invite === true;
  // first empty row (no title and no date)
  const vals = sh.getRange(2, 1, EV_ROWS, EV.title + 1).getValues();
  let at = vals.findIndex(r => !(r[EV.date] instanceof Date) && !String(r[EV.title]).trim());
  if (at === -1) { sh.insertRowsAfter(sh.getMaxRows(), 1); at = sh.getMaxRows() - 2; }
  sh.getRange(at + 2, 1, 1, EV_EDIT).setValues([row]);
  try { renderCalView_(ss); } catch (e) { Logger.log('Calendar view: ' + e.message); }
  const when = Utilities.formatDate(date, CONFIG.timezone, 'EEE, MMM d') + (row[EV.rep] !== 'Never' ? ' · ' + row[EV.rep].toLowerCase() : '');
  if (!calOn_() || !calServiceOn_()) {
    sh.getRange(at + 2, EV.status + 1).setValue('📝 Saved: share it with ' + MENU_CAL);
    return `✅ Saved in the Events tab (${when}).\nTo put it on the shared calendars: ${MENU_CAL}.`;
  }
  const r = syncEvents_(ss);
  const st = sh.getRange(at + 2, EV.status + 1).getValue();
  return `${String(st).indexOf('✅') === 0 ? '✅' : '⚠️'} ${row[EV.title]} (${when})\n${st}` +
         (r.problems ? `\n(${r.problems} other row(s) in the Events tab need a fix)` : '') + '\n\nAdd another, or close this box.';
}

/* ================================================================
 *  🗓 CALENDAR VIEW  (a month on one page, fills in by itself)
 *  Shows: your events (repeats included), weekly check-in dates, Outreach follow-ups, target deadlines.
 *  Pick the month in B2. Uses only the sheet (no Google Calendar needed), so it works even before 📅 is set up.
 * ================================================================ */

var CV = { titleRow: 1, pickRow: 2, headRow: 3, firstWeek: 4, weeks: 6, colW: 150, rowH: 104, maxLines: 5 };
var CV_MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

function monthLabel_(d) { return CV_MONTHS[d.getMonth()] + ' ' + d.getFullYear(); }
function monthChoices_() {
  const now = new Date(), out = [];
  for (let i = -3; i <= 12; i++) out.push(monthLabel_(new Date(now.getFullYear(), now.getMonth() + i, 1)));
  return out;
}

function buildCalView_(ss) {
  if (DB_MODE) return;                                   // database mode: not used (see 00_Config)
  const sh = getOrCreate_(ss, TAB.calview);
  const cols = 7;
  if (sh.getMaxColumns() < cols + 4) sh.insertColumnsAfter(sh.getMaxColumns(), cols + 4 - sh.getMaxColumns());
  const rows = CV.firstWeek + CV.weeks + 2;
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  sh.setHiddenGridlines(true).setTabColor(THEME.tabs.calview).setFrozenRows(CV.headRow);
  for (let c = 1; c <= cols; c++) sh.setColumnWidth(c, CV.colW);
  sh.getRange(1, 1, rows, cols).setFontFamily(THEME.font).setFontColor(THEME.navy);
  const t = sh.getRange(CV.titleRow, 1, 1, cols).merge().setValue('🗓  CALENDAR VIEW  ·  fills in by itself from Events, Settings and the Outreach log');
  header_(t, THEME.navy); t.setFontSize(12).setHorizontalAlignment('left'); sh.setRowHeight(CV.titleRow, 36);
  sh.getRange(CV.pickRow, 1).setValue('Month ▸').setFontWeight('bold').setHorizontalAlignment('right');
  const pick = sh.getRange(CV.pickRow, 2, 1, 2).merge();
  pick.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(monthChoices_(), true).setAllowInvalid(false).build())
      .setBackground(THEME.yellow).setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  if (!String(pick.getValue()).trim()) pick.setValue(monthLabel_(new Date()));
  sh.getRange(CV.pickRow, 4, 1, 4).merge()
    .setValue('👥💻🚀⏰📌 your events   ·   📨 ⏰ 🔁 check-in days   ·   📞 follow-ups   ·   🎯 deadlines')
    .setFontSize(9).setFontColor(THEME.muted).setHorizontalAlignment('left');
  sh.setRowHeight(CV.pickRow, 30);
  const head = sh.getRange(CV.headRow, 1, 1, cols).setValues([['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']]);
  header_(head, THEME.blue); head.setHorizontalAlignment('center');
  for (let w = 0; w < CV.weeks; w++) sh.setRowHeight(CV.firstWeek + w, CV.rowH);
  sh.getRange(CV.firstWeek, 1, CV.weeks, cols).setVerticalAlignment('top').setWrap(true).setFontSize(9)
    .setBorder(true, true, true, true, true, true, TABLE_LINE, SpreadsheetApp.BorderStyle.SOLID);
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).filter(p => p.getDescription() === PROTECT_TAG).forEach(p => p.remove());
  sh.getRange(CV.firstWeek, 1, CV.weeks, cols).protect().setDescription(PROTECT_TAG).setWarningOnly(true);
  placeTabBanner_(sh, 7);
  placeTutorial_(sh);
  renderCalView_(ss, sh);
  return sh;
}

// What happens on each day of the month: { 'yyyy-mm-dd': [ 'text', … ] }
function calViewItems_(ss, from, to) {
  const st = getSettings_(), items = {};
  const add = (d, text, order) => { const k = ymd_(d); (items[k] = items[k] || []).push({ text, order }); };
  let joined = {};
  try { if (checkinOn_()) joined = memberData_(ss).joined; } catch (e) { joined = {}; }   // 👥 member check-ins
  const ev = ss.getSheetByName(TAB.events);
  if (ev && ev.getLastRow() > 1) {
    ev.getRange(2, 1, ev.getLastRow() - 1, EV_EDIT).getValues().forEach(r => {
      const title = String(r[EV.title]).trim();
      if (!title) return;
      const m = minutesOf_(r[EV.start]);
      const when = m === null ? '' : Utilities.formatDate(new Date(2000, 0, 1, Math.floor(m / 60), m % 60), CONFIG.timezone, 'h:mma').toLowerCase() + ' ';
      const who = String(r[EV.who]).trim(), tag = who === WHO.me ? ' (me)' : who && who !== WHO.all && who !== WHO.leads ? ` (${who})` : '';
      occurrences_(r, from, to).forEach(d => {
        const n = joined[evLabel_(title, d, who || WHO.all)];
        add(d, `${EV_ICON[r[EV.type]] || '📌'} ${when}${title}${tag}${n ? ' · ' + n + ' joined' : ''}`, m === null ? 0 : m);
      });
    });
  }
  if (st.autoWeekly === true) {                                   // weekly check-in days (same as the Campus Leads calendar)
    const marks = [[st.remDay, '📨 Check-in reminder'], [st.dueDay, '⏰ Check-in due']];
    if (st.fuDay !== 'No follow-up') marks.push([st.fuDay, '🔁 Follow-up']);
    for (let d = new Date(from); d < to; d.setDate(d.getDate() + 1)) {
      const name = DAYS[(d.getDay() + 6) % 7];
      marks.forEach(([day, text]) => { if (day === name) add(new Date(d), text, -1); });
    }
  }
  if (docProps_().getProperty('MFORM_ID')) {   // 👥 member check-in days
    for (let d = new Date(from); d < to; d.setDate(d.getDate() + 1)) {
      const name = DAYS[(d.getDay() + 6) % 7];
      if (name === st.mOpenDay) add(new Date(d), '✨ Member check-in opens', -1);
      if (name === st.mDueDay) add(new Date(d), '⏰ Member check-in due', -1);
    }
  }
  const ol = ss.getSheetByName(TAB.outreach);
  if (ol && ol.getLastRow() > 1) {
    ol.getRange(2, 1, ol.getLastRow() - 1, OUTREACH_HEAD.length).getValues().forEach(r => {
      if (r[6] instanceof Date && r[6] >= from && r[6] < to && String(r[1]).trim()) add(r[6], `📞 Follow up: ${r[1]}`, 2000);
    });
  }
  [[st.dFocal, `🎯 ${st.tFocal} Campus Leads named`], [st.dReeng, `🎯 ${st.tReeng} schools re-engaged`]]
    .forEach(([d, text]) => { if (d instanceof Date && d >= from && d < to) add(d, text, 2001); });
  return items;
}

// Draws the month picked in B2. Cheap (sheet only), so it runs after every relevant edit.
function renderCalView_(ss, sh) {
  if (DB_MODE) return;                                   // database mode: not used (see 00_Config)
  sh = sh || ss.getSheetByName(TAB.calview);
  if (!sh) return;
  const label = String(sh.getRange(CV.pickRow, 2).getValue()).trim();
  const m = label.match(/^(\w+)\s+(\d{4})$/);
  const month = m && CV_MONTHS.indexOf(m[1]) !== -1 ? new Date(+m[2], CV_MONTHS.indexOf(m[1]), 1) : new Date(new Date().getFullYear(), new Date().getMonth(), 1);
  const gridStart = new Date(month); gridStart.setDate(1 - ((month.getDay() + 6) % 7));        // the Monday on or before the 1st
  const gridEnd = new Date(gridStart); gridEnd.setDate(gridStart.getDate() + 7 * CV.weeks);
  const items = calViewItems_(ss, gridStart, gridEnd);
  const today = ymd_(new Date());
  const rich = [], bg = [];
  for (let w = 0; w < CV.weeks; w++) {
    const rr = [], bb = [];
    for (let c = 0; c < 7; c++) {
      const d = new Date(gridStart); d.setDate(gridStart.getDate() + w * 7 + c);
      const k = ymd_(d), inMonth = d.getMonth() === month.getMonth();
      const list = (items[k] || []).sort((a, b) => a.order - b.order).map(x => x.text);
      const shown = list.slice(0, CV.maxLines).concat(list.length > CV.maxLines ? [`+${list.length - CV.maxLines} more`] : []);
      const top = String(d.getDate()) + (k === today ? '  · TODAY' : '');
      const text = [top].concat(shown).join('\n');
      const b = SpreadsheetApp.newRichTextValue().setText(text)
        .setTextStyle(0, top.length, SpreadsheetApp.newTextStyle().setBold(true).setFontSize(11)
          .setForegroundColor(inMonth ? (k === today ? THEME.blue : THEME.navy) : '#9AA7BA').build());
      if (!inMonth && text.length > top.length) b.setTextStyle(top.length, text.length, SpreadsheetApp.newTextStyle().setForegroundColor('#9AA7BA').build());
      rr.push(b.build());
      bb.push(k === today ? '#FFF6BF' : !inMonth ? '#F4F6F9' : c >= 5 ? '#F6F9FD' : '#FFFFFF');
    }
    rich.push(rr); bg.push(bb);
  }
  const grid = sh.getRange(CV.firstWeek, 1, CV.weeks, 7);
  grid.setRichTextValues(rich).setBackgrounds(bg);
}

// Called by onSheetChange: when members, their emails or Campus Leads change, sharing follows right away
// (not only the next morning). Sharing only; at most about a minute per change.
function calPeopleChanged_(ss) {
  if (!calOn_() || !calServiceOn_()) return '';
  const p = calPeople_(ss), props = docProps_();
  const key = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify([p.bySchool, p.leadOf])));
  if (props.getProperty('CAL_PEOPLE') === key) return '';
  const t0 = Date.now(), res = { shareFailed: [], blocked: false };
  shareAll_(calStore_(), p, res, () => Date.now() - t0 < 60000);
  if (!res.shareFailed.length && !res.blocked && !res.incomplete) props.setProperty('CAL_PEOPLE', key);
  return '📅 Calendar sharing updated' + (res.shareFailed.length ? ` (${res.shareFailed.length} could not be shared)` : '') + '.';
}
