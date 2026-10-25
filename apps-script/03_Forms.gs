/* =====================================================================
 *  START SCHOLAR TRACKER  ·  03_Forms
 *  ---------------------------------------------------------------------
 *  Inside: The weekly check-in form for everyone: questions, memes, school list, answers.
 *  To customise: Member questions and choices.
 *
 *  The code has 9 files + Email_Reminder and Web_App (HTML). Keep these exact names:
 *    00_Config, 01_Region, 02_Menu_Setup, 03_Forms, 04_Emails, 05_Portals, 06_Calendar, 07_Look, 08_Tools
 *  To update: open each file, select all (Ctrl+A), paste the new version, Save.
 *  Then in the Google Sheet: START Tracker › ① Set up / fix the tracker.
 *  You never need to run anything from here: use the START Tracker menu in the Sheet.
 * ===================================================================== */

// The ▶ Run button runs the FIRST function of the open file: this one, so it is always safe.
function about_03_Forms() {
  Logger.log('START Tracker' + (typeof START_MODULES === 'string' ? ' ' + START_MODULES : '') + ', file 03_Forms. Nothing to run here: open the Google Sheet and use the START Tracker menu.');
}


// ━━━━━━━━━━━━━━━━━━━━  Form  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  FORM SCHOOL LIST  (updates by itself; ① also refreshes it)
 * ================================================================ */

function refreshSchoolDropdown() {
  sheetUser_();                                   // sheet only: never from the website
  requireSetup_();
  const n = updateDropdown_();
  notify_(n ? `The check-in form has been updated. It now lists ${n} school(s).`
            : 'No schools have been entered in the Schools tab. Please add the schools first, then click ① Set up / fix the tracker.');
}

function updateDropdown_() {
  const ss = SpreadsheetApp.getActive();
  if (WEB_CHECKIN) { docProps_().setProperty('FORM_SYNC', syncKey_(ss)); return schoolNames_(ss).length; }   // no form to update
  const names = schoolNames_(ss);
  const form = getForm_();
  const st = getSettings_();
  retry_(() => form.setTitle('START Campus Lead Check-in' + (st.region ? ' – ' + st.region : '')));
  const item = retry_(() => form.getItems(FormApp.ItemType.LIST)).find(i => i.getTitle() === Q.school);
  if (!item) throw new Error('The "School" question in the form could not be found. It may have been renamed or deleted. Please restore its title to "School".');
  retry_(() => item.asListItem().setChoiceValues(names.length ? names : [NO_SCHOOLS]));
  writeSchoolLinks_(ss, form, item.asListItem());
  docProps_().setProperty('FORM_SYNC', syncKey_(ss));
  return names.length;
}

/* ---- Automatic form sync ----
 * Runs by itself whenever the sheet changes (typing, pasting, deleting rows).
 * It only touches the form when the school list or region actually changed. */

function installSyncTrigger_(ss) {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'onSheetChange');
  if (!exists) ScriptApp.newTrigger('onSheetChange').forSpreadsheet(ss).onChange().create();
}

function syncKey_(ss, rows) {
  return schoolNames_(ss, rows).join('|') + '#' + getSettings_().region;
}

// Runs by itself after any change in the sheet (typing, pasting, deleting rows).
// 1) school list or region changed      -> update the form and the per-school links
// 2) a START Campus Lead or their email changed -> START Campus Leads tab + portals (syncFocal_)
// 3) a schedule setting changed           -> move the email triggers that are on, and the red/green scale
function onSheetChange(e) {
  if (!ownerGuard_(e)) return;
  const ss = SpreadsheetApp.getActive();
  const props = docProps_();
  if (!props.getProperty('FORM_ID') && !(WEB_CHECKIN && props.getProperty('SETUP_DONE'))) return;
  Utilities.sleep(1500);                                // let onEdit finish (name/email matching) before reading
  let rows = schoolRows_(ss);                           // Schools tab read once for every check
  let fk = focalKey_(ss, rows);
  let tab = '';
  try { tab = e.source.getActiveSheet().getName(); } catch (err) { /* no event (run by hand) */ }
  if (tab === TAB.scholars && props.getProperty('FOCAL_SYNC') === fk) {   // an email changed in Scholars reaches Schools a moment later
    Utilities.sleep(3000);
    rows = schoolRows_(ss); fk = focalKey_(ss, rows);
  }
  if (tab === TAB.scholars) {                                                          // an email was added: answers waiting in "Not recognised" now count
    try { const n = adoptRecognised_(ss); if (n) ss.toast(`${n} waiting answer(s) now counted (their email is in Scholars).`, 'START Tracker', 6); }
    catch (err) { Logger.log('Not recognised: ' + err.message); }
  }
  if (tab === TAB.scholars) {                                                          // ✨ same email twice / another school's email
    try { const bad = checkScholarEmails_(ss); if (bad.length) ss.toast(`${bad.length} email(s) in Scholars look wrong: see the orange cells`, 'START Tracker', 6); }
    catch (err) { Logger.log('Email checks: ' + err.message); }
  }
  try { const said = eventsChanged_(ss); if (said) ss.toast(said, 'START Tracker', 5); }   // 📅 Events tab (22_Calendar)
  catch (err) { Logger.log('Calendar sync: ' + err.message); }
  try { const said = calPeopleChanged_(ss); if (said) ss.toast(said, 'START Tracker', 5); }   // 📅 members / Leads changed
  catch (err) { Logger.log('Calendar sync: ' + err.message); }
  if (tab === TAB.settings) {                                                          // 🗺 region changed: drop-downs + names follow
    try {
      const r = regionProfile_();
      if (props.getProperty('REGION_APPLIED') !== r.code + '|' + r.label) { applyRegion_(ss); ss.toast(`Region set to ${r.label || '(none)'}: forms, emails and lists follow it.`, 'START Tracker', 6); }
    } catch (err) { Logger.log('Region: ' + err.message); }
  }
  if (tab === TAB.events || tab === TAB.settings) {                                     // 👥 member form follows events / dates
    try { refreshMemberForm_(ss); writeMemberJoined_(ss); } catch (err) { Logger.log('Member form: ' + err.message); }
    try { refreshLeadForm_(ss); } catch (err) { Logger.log('Campus Lead form: ' + err.message); }
  }
  const formChanged = props.getProperty('FORM_SYNC') !== syncKey_(ss, rows);
  const sk = scheduleKey_();
  if (!formChanged && props.getProperty('FOCAL_SYNC') === fk && props.getProperty('SCHEDULE_SYNC') === sk) return;
  const lock = docLock_();
  if (!lock.tryLock(5000)) return;
  try {
    if (formChanged) {
      const n = updateDropdown_();
      ss.toast(`The check-in form has been updated (${n} school(s)).`, 'START Tracker', 5);
    }
    const msg = syncFocal_(ss, fk);
    if (msg) ss.toast(msg, 'START Tracker', 6);
    if (props.getProperty('SCHEDULE_SYNC') !== sk) {
      props.setProperty('SCHEDULE_SYNC', sk);
      const said = applySchedule_();
      if (calOn_() && calServiceOn_()) { try { syncCheckinSeries_(calStore_()); } catch (err) { Logger.log('Check-in dates: ' + err.message); } }
      try { polishSchools_(ss); } catch (e) { /* heat scale uses "Mark a school red after" */ }
      if (said) ss.toast(said, 'START Tracker', 6);
    }
  } catch (e) {
    Logger.log('Auto sync failed: ' + e.message);
  } finally { lock.releaseLock(); }
}

// THE one routine for "who are the START Campus Leads" changes, whoever made them (typing, a check-in,
// the Outreach log). Refreshes the START Campus Leads tab and, when portals are on, shares/unshares portals.
// Returns a short message, or '' when nothing changed.
function syncFocal_(ss, fk) {
  const props = docProps_();
  fk = fk || focalKey_(ss);
  if (props.getProperty('FOCAL_SYNC') === fk) return '';
  props.setProperty('FOCAL_SYNC', fk);
  try { syncRoles_(ss); } catch (e) { Logger.log('Roles not updated: ' + e.message); }
  if (portalsOn_()) {
    const r = refreshPortals_();
    return 'START Campus Leads updated' + (r.newShares.length ? '. Portal shared with: ' + r.newShares.join(', ') : '.');
  }
  buildFocalTab_(ss);
  return 'START Campus Leads tab updated.';
}

// Settings that decide WHEN things happen
function scheduleKey_() {
  const st = getSettings_();
  return [st.autoWeekly, st.day, st.hour24, st.remDay, st.remHour24, st.fuDay, st.quietDays].join('|');
}

// THE switch for automatic weekly emails (Settings › 🚀, menu ⏰ Automatic emails). Returns a short message.
function applySchedule_() {
  const st = getSettings_();
  if (st.autoWeekly === true) {
    scheduleWeekly_(); scheduleReminders_();
    return `🚀 Weekly emails ON: summary ${st.day} ${HOURS[st.hour24]}, reminders ${st.remDay} ${HOURS[st.remHour24]}.`;
  }
  if (st.autoWeekly === false) {
    return removeTriggers_(['weeklySummary', 'autoReminder', 'autoFollowUp']) ? '⏸ Weekly emails OFF.' : '';
  }
  return '';                                              // switch not created yet (run ①)
}

// Fingerprint of who is START Campus Lead (and their contact) for every school
function focalKey_(ss, rows) {
  return (rows || schoolRows_(ss)).map(s => [s.name, s.focal, s.contact].join('~')).join('|');
}

// Portals are "on" once the RR has created them (the daily update trigger exists)
function portalsOn_() {
  return docProps_().getProperty('PORTALS_ON') === '1' ||
         ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'refreshPortalsAuto');
}


// Column P: a form link per school with the school already picked, so START Campus Leads can't choose the wrong one.
// With the one weekly form (member check-in on), every school simply gets that form's link (people sign in, so no school to pick).
function writeSchoolLinks_(ss, form, listItem) {
  const sh = ss.getSheetByName(TAB.schools);
  ensureInfoHeaders_(sh);
  const vals = sh.getRange(2, 1, CONFIG.maxSchools, 1).getValues().map(r => String(r[0]).trim());
  const one = oneFormOn_() ? checkinUrl_() : '';
  const out = vals.map(n => {
    if (!n) return [''];
    if (one) return [one];
    try { return [retry_(() => form.createResponse().withItemResponse(listItem.createResponse(n)).toPrefilledUrl())]; }
    catch (e) { return ['']; }
  });
  sh.getRange(2, LINK_COL, CONFIG.maxSchools, 1).setValues(out).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  fitRows_(sh, 2, CONFIG.maxSchools);
}

function schoolNames_(ss, rows) {
  return [...new Set((rows || schoolRows_(ss)).map(s => s.name))].sort();
}

/* ================================================================
 *  📋 MESSAGE FOR START CAMPUS LEADS
 * ================================================================ */

function showFocalMessage() {
  sheetUser_();                                   // sheet only: never from the website
  requireSetup_();
  const st = getSettings_();
  const general = checkinLink_();
  const links = {}, store = portalStore_();
  schoolRows_(SpreadsheetApp.getActive()).forEach(s => {
    links[s.name] = { link: s.link, focal: displayName_(s.focal), portal: portalUrl_(store, s) };
  });

  // Week covered and deadline: same rule as the automatic reminders (deadline day is in Settings)
  const w = reminderWeek_(new Date());
  const monday = w.monday, deadline = w.deadline;
  const f = d => Utilities.formatDate(d, CONFIG.timezone, 'MMMM d, yyyy');
  const fd = d => Utilities.formatDate(d, CONFIG.timezone, 'EEEE, MMMM d, yyyy');

  const data = {
    region: st.region || '', rs: regionShort_(st.region), rr: st.rrName || '', general: general, schools: links,
    week: f(monday), deadline: fd(deadline),
  };
  const html = HtmlService.createHtmlOutput(DLG_CSS +
    `<div class="dbar"></div>
       <label style="margin-top:0">School</label>
       <select id="s"></select>
       <label>Message</label>
       <textarea id="t" class="copy" style="height:250px"></textarea>
       <div class="row" style="margin-top:10px"><button class="btn" id="b">Copy message</button>
         <span class="muted">Pick "All schools" for the group chat.</span></div>
     <script>
       var D = ${JSON.stringify(data).replace(/</g, '\\u003c')};
       var s = document.getElementById("s"), t = document.getElementById("t"), b = document.getElementById("b");
       function opt(v, label) { var o = document.createElement("option"); o.value = v; o.textContent = label; s.appendChild(o); }
       opt("", "All schools");
       Object.keys(D.schools).sort().forEach(function (k) {
         opt(k, k + (D.schools[k].focal ? " – " + D.schools[k].focal : ""));
       });
       function render() {
         var k = s.value, info = D.schools[k] || {};
         var link = (k && info.link) ? info.link : D.general;
         var greet = (k && info.focal) ? "Dear " + info.focal + "," : (k ? "Dear START Campus Lead of " + k + "," : "Dear START Campus Leads,");
         var scope = k ? " for " + k : (D.region ? " for the schools of " + D.region : "");
         var lines = [
           greet,
           "",
           "Good day! This is a reminder to accomplish the START Weekly Check-in" + scope +
             " for the week of " + D.week + ". The form will take approximately two (2) minutes to complete.",
           "",
           "Form link: " + link,
           "",
         ];
         if (k && info.portal) lines.push("Your START Campus Lead Portal (private to your Google account): " + info.portal, "");
         lines = lines.concat([
           "Kindly submit your response on or before " + D.deadline + ". Should you have any concerns or require " +
             "assistance, please indicate them in the form or contact me directly.",
           "",
           "Thank you for your continued support of the START program.",
           "",
           "Respectfully,"
         ]);
         if (D.rr) lines.push(D.rr);
         lines.push("START " + (D.rs ? D.rs + " " : "") + "Regional Representative");
         t.value = lines.join("\\n");
         b.innerText = "Copy message";
       }
       s.onchange = render; render();
       b.onclick = function () {
         t.select();
         try { navigator.clipboard.writeText(t.value); } catch (e) {}
         try { document.execCommand("copy"); } catch (e) {}
         b.innerText = "Copied ✓";
       };
     </script>`).setWidth(580).setHeight(500);
  ui_().showModalDialog(html, '📋 Message for Campus Leads');
}

/* ================================================================
 *  NEW CHECK-IN ARRIVES  (runs by itself on every form submission)
 *  - fills in the START Campus Lead's name if the Schools tab doesn't have one yet
 *  - emails the RR immediately when a school says "Yes, urgent"
 * ================================================================ */

function installSubmitTrigger_(ss) {
  const has = n => ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === n);
  if (!has('onCheckin') && !WEB_CHECKIN) ScriptApp.newTrigger('onCheckin').forSpreadsheet(ss).onFormSubmit().create();
  if (!has('catchUpAnswers')) ScriptApp.newTrigger('catchUpAnswers').timeBased().everyHours(1).create();   // safety net
  try { if (!docProps_().getProperty(SEEN_KEY)) markAnswersSeen_(ss); } catch (e) { /* first run */ }
}

/* ---------- safety net: answers Google's "on form submit" missed ----------
 * Google sometimes skips an on-submit run (busy, time limit, another account). Every hour this looks for
 * answers that arrived since the last look and updates those schools: START Campus Leads tab, portals,
 * members joined. Quick when nothing is new. */
var SEEN_KEY = 'ANSWERS_SEEN';
function answerSheets_(ss) {
  const out = [];
  try { const m = memberTab_(ss); if (m) out.push(['m', m]); } catch (e) { /* no member form */ }
  const l = ss.getSheetByName(TAB.checkins); if (l) out.push(['l', l]);
  return out;
}
function markAnswersSeen_(ss) {
  const seen = {}; answerSheets_(ss).forEach(([k, sh]) => { seen[k] = sh.getLastRow(); });
  docProps_().setProperty(SEEN_KEY, JSON.stringify(seen));
}
function catchUpAnswers(e) {
  if (!ownerGuard_(e)) return;
  const ss = SpreadsheetApp.getActive(), props = docProps_();
  let seen = {}; try { seen = JSON.parse(props.getProperty(SEEN_KEY) || '{}'); } catch (err) { seen = {}; }
  const schools = {}, byEmail = (() => { try { return scholarByEmail_(ss); } catch (err) { return {}; } })();
  let any = false;
  answerSheets_(ss).forEach(([k, sh]) => {
    const last = sh.getLastRow(), from = Math.max((seen[k] || 1) + 1, 2);
    seen[k] = last;
    if (last < from) return;
    any = true;
    if (k === 'm') {
      const c = memberCols_(sh);
      if (c.email < 0) return;
      sh.getRange(from, c.email + 1, last - from + 1, 1).getValues()
        .forEach(r => { const w = byEmail[emailOf_(r[0])]; if (w) schools[w.school] = 1; });
    } else {
      const col = ckCols_(sh).school;
      sh.getRange(from, col + 1, last - from + 1, 1).getValues().forEach(r => { if (String(r[0]).trim()) schools[String(r[0]).trim()] = 1; });
    }
  });
  props.setProperty(SEEN_KEY, JSON.stringify(seen));
  if (!any) return;
  try { MEMBER_CACHE_ = null; writeMemberJoined_(ss); } catch (err) { Logger.log('Members joined: ' + err.message); }
  const store = portalStore_();
  Object.keys(schools).forEach(name => {
    if (outOfTime_(60 * 1000)) return;
    const p = store.map[name];
    if (p && p.id && p.email) { try { refreshPortals_(name); } catch (err) { Logger.log('Catch-up portal ' + name + ': ' + err.message); } }
  });
  try { buildFocalTab_(ss); } catch (err) { Logger.log('Catch-up Leads tab: ' + err.message); }
}

function onCheckin(e) {
  if (!ownerGuard_(e)) return;                          // only the email account handles answers (no double alerts)
  try { if (e && e.range && isMemberSheet_(e.range.getSheet())) { memberTab_(SpreadsheetApp.getActive()); return onMemberCheckin_(e); } }   // 👥 23_Members
  catch (err) { Logger.log('Member check-in: ' + err.message); return; }
  const nv = (e && e.namedValues) || {};
  const get = t => (nv[t] && nv[t][0] ? String(nv[t][0]).trim() : '');
  const sos = typeof LQ_VIBE !== 'undefined' && samePick_(get(LQ_VIBE) || get('Kamustahan: how are you holding up as Campus Lead?'), L_VIBES[3]);      // 🆘 in the kamustahan
  leadCheckinDone_(get(Q.school), get(Q.name), get(Q.help), get(Q.concern), sos);
  try { markAnswersSeen_(SpreadsheetApp.getActive()); } catch (err) { /* the hourly check will catch up */ }
}

// After a Campus Lead's check-in (from the old Lead form, or the one weekly form): Leads tab, portal, urgent alert
function leadCheckinDone_(school, name, help, concern, sos) {
  if (!school) return;
  const ss = SpreadsheetApp.getActive();
  const row = schoolRows_(ss).find(s => s.name === school);
  if (row && name && !row.focal) {
    assignFocal_(ss.getSheetByName(TAB.schools), row.row, name, null, 'Entered automatically from the check-in submitted on ' +
      Utilities.formatDate(new Date(), CONFIG.timezone, 'MMMM d, yyyy') + '. Kindly verify this entry.');
  }
  // Keep the START Campus Leads tab and this school's private portal up to date
  try {
    if (!syncFocal_(ss)) {                              // START Campus Leads unchanged: refresh just this school
      const p = portalStore_().map[school];
      if (p && p.id && p.email) refreshPortals_(school);
      else buildFocalTab_(ss);
    }
  } catch (err) { Logger.log('Focal update after check-in: ' + err.message); }

  if (help === HELP_CHOICES[2] || sos) leadUrgentAlert_(ss, school, name, help, concern, sos);
}

// The urgent alert to you (from a Campus Lead check-in). Also used by 🧪 Send me a test of every email.
function leadUrgentAlert_(ss, school, name, help, concern, sos) {
    if (!RR_EMAILS && !TEST_MODE_) return;                                   // on the RR's START page (Asked for help) instead
    const st = getSettings_();
    const to = String(st.recipients || '').trim() || Session.getEffectiveUser().getEmail();
    const when = Utilities.formatDate(new Date(), CONFIG.timezone, 'MMM d, yyyy · h:mm a');
    const html = designedEmail_({
      badge: 'Urgent', badgeColor: THEME.red, title: ['Urgent', 'Request'], greeting: '',
      intro: [sos && help !== HELP_CHOICES[2]
        ? `${school}'s Campus Lead answered "${L_VIBES[3]}" this week. Please check in with them.`
        : `${school}'s check-in asks for urgent help.`],
      extraHtml: emailSection_('Their concern') +
        emailAlert_(esc_(concern || 'No details were provided.').replace(/\n/g, '<br>'), 'red'),
      link: ss.getUrl(), buttonText: 'Open the tracker', buttonNote: '',
      chips: [['SCHOOL', school], ['SUBMITTED BY', name || 'Not provided'], ['RECEIVED', when]],
      outro: [],
      signature: [],
      footer: 'This is an automated notification from the START Tracker.',
    });
    sendMail_({
      to, name: 'START Tracker', inlineImages: logoInlineImages_(),
      subject: `START · Urgent: help request from ${school}`,
      htmlBody: html,
    });
}


/* ---------- ONE weekly form for everyone (members and Campus Leads) ---------- */
// On as soon as the member check-in exists: Campus Leads answer the same form as their members.
function oneFormOn_() { if (WEB_CHECKIN) return true; try { return typeof memberForm_ === 'function' && !!memberForm_(); } catch (e) { return false; } }
// Is the weekly member check-in running? (website check-in: always, once its tab exists)
function checkinOn_() {
  if (WEB_CHECKIN) { try { const ss = SpreadsheetApp.getActive(); return !!(ss.getSheetByName(TAB.mcheckins) || ss.getSheetByName('Member check-ins')); } catch (e) { return false; } }   // old tab name: until ① moves it
  try { return typeof memberForm_ === 'function' && !!memberForm_(); } catch (e) { return false; }
}

// The link everyone uses this week (the one weekly form; the older Campus Lead form only before members are set up)
// Where people answer the weekly check-in: the START website (database mode), else the Google Form
function checkinLink_() { const w = DB_MODE && typeof webHome_ === 'function' ? webHome_() : ''; return w || checkinUrl_(); }

function checkinUrl_() {
  if (WEB_CHECKIN) { try { return webHome_() || ''; } catch (e) { return ''; } }
  if (oneFormOn_()) { try { return memberForm_().getPublishedUrl(); } catch (e) { /* fall back below */ } }
  try { return getForm_().getPublishedUrl(); } catch (e) { return ''; }
}


// ━━━━━━━━━━━━━━━━━━━━  Members  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  👥 MEMBER CHECK-INS  (every active scholar, weekly, short and fun)
 *   • ONE form for all members. They sign in with Google; the email is matched to Scholars › Contact (E).
 *     Emails that are not in Scholars are ignored (the answer is removed from the tracker).
 *   • The form changes by itself (daily check, and after you edit Events or Settings):
 *       no START events this week  -> short version: vibe + need help
 *       START events this week     -> also "which events did you join?" (this week's events listed)
 *       DataCamp running           -> also the learning question (Settings › DataCamp dates)
 *   • Meme pictures: put them in the Drive folder "START member check-in pictures" named
 *       vibe-1.jpg, vibe-2.jpg … events-1.jpg … learning-1.jpg … help-1.jpg …  (rotate every week)
 *   • Campus Leads share the link (ready-made group-chat message in their portal) and see their members there.
 *   • "Yes, urgent" emails that school's Campus Lead and you right away. No other emails.
 * ================================================================ */

var MQ = {
  vibe: 'How was your week?',
  events: 'Which START events did you attend this week?',
  learning: 'How is your DataCamp progress this week?',
  help: 'Do you need help from START?',
  tea: 'Comments or concerns (optional)',
};
var M_PAGES = { events: 'START events', learning: 'DataCamp', help: 'Support' };
// the first emoji of each choice never changes (older answers are still understood)
var M_VIBES = ['🔥 Very good', '🙂 Good', '😵‍💫 Challenging', '🆘 Difficult, I need support'];
var M_LEARN = ['📈 Good progress', '🐢 Some progress', '😴 No progress'];
var M_HELP = ['👍 No', '🙋 Yes, not urgent', '🚨 Yes, urgent'];
var M_MISSED = 'None this week';
// The events a scholar ticked, one per item. They are saved joined by ", ", but each event name ends with its date
// ("START Kapehan · Sat, Oct 24"), so a comma followed by a date is part of the name, not a new event.
function splitEvents_(text) {
  return String(text || '').split(/, (?![A-Z][a-z]{2} \d{1,2}(?:, |$))/).map(x => x.trim()).filter(Boolean);
}
// "did not attend" answer, new or old wording ("😢 I did not attend any")
function missedPick_(x) { return samePick_(x, M_MISSED) || samePick_(x, '😢'); }

// Same answer? Compares the leading emoji, so answers given with older wording still count.
function samePick_(answer, choice) {
  const a = String(answer || '').trim();
  return !!a && a.split(' ')[0] === String(choice).split(' ')[0];
}

// Older question and section titles -> current ones. Renaming keeps the same answer column in the sheet.
var FORM_RENAMES = {
  'Vibe check! How\'s your week?': 'How was your week?',
  'You pulled up to…?': 'Which START events did you attend this week?',
  'Learning grind status?': 'How is your DataCamp progress this week?',
  'Need backup from START?': 'Do you need help from START?',
  'Spill the tea: what\'s up? (optional)': 'Comments or concerns (optional)',
  'Anything you want to tell us? (optional)': 'Comments or concerns (optional)',
  '🎉 This week\'s START events': 'START events', '📚 Learning check': 'DataCamp', '🙋 Almost done!': 'Support',
  '🎉 START events': 'START events', '📚 DataCamp': 'DataCamp', '🙋 Last step': 'Support', '🙋 About you': 'About you',
  'Kamustahan: how are you holding up as Campus Lead?': 'How are you doing as Campus Lead this week?',
  'Did you post this week\'s member check-in in your school group chat?': 'Did you share the member check-in link with your members?',
  '🎉 Event check!': 'START events', '📚 DataCamp check': 'DataCamp', '✨ Kamustahan (Campus Lead edition)': 'About you',
  'Concerns for START': 'Anything START should know? (optional)',
};
function renameFormItems_(form) {
  const items = form.getItems(), titles = items.map(x => x.getTitle());
  items.forEach(it => {
    const to = FORM_RENAMES[it.getTitle()];
    if (to && titles.indexOf(to) === -1) { try { it.setTitle(to); titles.push(to); } catch (e) { Logger.log('Rename: ' + e.message); } }
  });
}
// A sheet column by its current title, or by an older title of the same question
function colOf_(h, title) {
  let i = h.indexOf(title);
  if (i !== -1) return i;
  Object.keys(FORM_RENAMES).forEach(o => { if (i === -1 && FORM_RENAMES[o] === title) i = h.indexOf(o); });
  return i;
}
// Safety net: an older 00_Config has no Member check-ins tab name (Google then says "Invalid argument: name")
var OLD_CONFIG_ = typeof TAB !== 'undefined' && !TAB.mcheckins;   // 🩺 reports it
if (OLD_CONFIG_) TAB.mcheckins = 'Check-in answers';
if (typeof THEME !== 'undefined' && THEME.tabs && !THEME.tabs.mcheckins) THEME.tabs.mcheckins = '#9AA7BA';
var M_PIC_FOLDER = 'START member check-in pictures';
var M_PIC_SLOTS = ['vibe', 'events', 'learning', 'help'];
var MEMBER_CACHE_ = null;

// The meme folder can have any of these names (the zip unpacks as START_member_checkin_pictures).
var M_PIC_FOLDER_NAMES = [M_PIC_FOLDER, 'START_member_checkin_pictures', 'START member checkin pictures'];

// The meme folder with the most pictures in it (an empty folder of the same name is ignored)
function memePicFolder_() {
  const props = docProps_();
  const id = props.getProperty('MEME_FOLDER_ID');   // set with 🛠 › Connect the pictures folder
  if (id) { try { const d = DriveApp.getFolderById(id); if (!d.isTrashed()) return d; } catch (e) { Logger.log('Meme folder link: ' + e.message); } }
  // one folder for everything: memes merged into the tracker pictures folder (banners + memes together)
  try {
    const pf = pictureFolder_(false);
    if (pf && hasMemes_(pf)) return pf;
  } catch (e) { /* no pictures folder */ }
  let best = null, bestN = -1;
  M_PIC_FOLDER_NAMES.forEach(name => {
    const it = DriveApp.getFoldersByName(name);
    while (it.hasNext()) {
      const d = it.next();
      if (d.isTrashed()) continue;
      let n = 0; const fs = d.getFiles();
      while (fs.hasNext()) { const f = fs.next(); if (!f.isTrashed() && /^image\//.test(f.getMimeType())) n++; }
      if (n > bestN) { best = d; bestN = n; }
    }
  });
  return best;
}

// true when a folder holds at least one meme (vibe-…, events-…, learning-…, help-…, lead-…)
function hasMemes_(d) {
  const re = /^(lead-)?(vibe|events|learning|help)-/i;
  return withSubfolders_(d).some(x => { const fs = x.getFiles(); while (fs.hasNext()) { const f = fs.next(); if (!f.isTrashed() && re.test(f.getName())) return true; } return false; });
}

// All meme pictures, sorted by name. Two files with the same name? The newest one wins (your replacement).
function memePicFiles_() {
  const d = memePicFolder_();
  if (!d) return [];
  const byName = {};
  const files = [];
  withSubfolders_(d).forEach(x => { const it = x.getFiles(); while (it.hasNext()) files.push(it.next()); });   // memes may sit in a subfolder ("Gform Images")
  while (files.length) {
    const f = files.shift();
    if (f.isTrashed() || !/^image\//.test(f.getMimeType())) continue;
    const k = f.getName().toLowerCase().trim();
    if (!byName[k] || f.getLastUpdated() > byName[k].getLastUpdated()) byName[k] = f;
  }
  return Object.keys(byName).map(k => byName[k])
    .sort((a, b) => a.getName().localeCompare(b.getName(), undefined, { numeric: true }));
}

/* ---------- the member week: opens on Settings › "Member check-in opens", due on "… due" ---------- */

// { open, due, from, to }: the coming due date (today or later), its open day, and the 7-day window it covers
function memberWeek_(now) {
  const st = getSettings_();
  const due = startOfDay_(now || new Date());
  for (let i = 0; i < 7 && DAYS[(due.getDay() + 6) % 7] !== st.mDueDay; i++) due.setDate(due.getDate() + 1);
  const open = new Date(due);
  for (let i = 0; i < 7 && DAYS[(open.getDay() + 6) % 7] !== st.mOpenDay; i++) open.setDate(open.getDate() - 1);
  const from = new Date(due); from.setDate(due.getDate() - 6);
  const to = new Date(due); to.setDate(due.getDate() + 1);
  return { open, due, from, to };
}
function memberWeekKey_(d) { return ymd_(memberWeek_(d).due); }

function learningOn_(now) {
  const st = getSettings_(), d = startOfDay_(now || new Date());
  return st.dcStart instanceof Date && d >= startOfDay_(st.dcStart) && !(st.dcEnd instanceof Date && d > startOfDay_(st.dcEnd));
}

// This week's events members can join: "All members" ones + every school's own (labelled with the school)
function memberEventChoices_(ss, w) {
  const sh = ss.getSheetByName(TAB.events);
  if (!sh || sh.getLastRow() < 2) return [];
  const out = [];
  sh.getRange(2, 1, sh.getLastRow() - 1, EV_EDIT).getValues().forEach(r => {
    const title = String(r[EV.title]).trim(), who = String(r[EV.who]).trim() || WHO.all;
    if (!title || who === WHO.leads || who === WHO.me) return;
    occurrences_(r, w.from, w.to).forEach(d => out.push({ label: evLabel_(title, d, who), date: d }));
  });
  return out.sort((a, b) => a.date - b.date).map(x => x.label);
}
function evLabel_(title, d, who) {
  return `${title}${who && who !== WHO.all ? ' (' + who + ')' : ''} · ${Utilities.formatDate(d, CONFIG.timezone, 'EEE, MMM d')}`;
}

/* ---------- the form ---------- */

function memberForm_() {
  const id = docProps_().getProperty('MFORM_ID');
  if (!id) return null;
  try { const f = FormApp.openById(id); return destId_(f) === SpreadsheetApp.getActive().getId() ? f : null; }
  catch (e) { return null; }
}

// Menu 🛠 Setup tools › 📝 Set up / update member check-in
function setupMemberCheckin() {
  sheetUser_();                                   // sheet only: never from the website
  requireSetup_();
  if (!ensureOwner_('run the member check-in')) return;
  const ss = SpreadsheetApp.getActive();
  if (WEB_CHECKIN) {                                                  // no Google Form: the check-in is on the website
    ensureCheckinTabs_(ss); retireForms_(ss); styleMemberTab_(ss);
    notify_('✅ The weekly check-in is on the START website\n\nScholars sign in and answer it there. Answers land in the ' + TAB.mcheckins + ' tab.');
    return;
  }
  let form = memberForm_();
  if (!form) form = createMemberForm_(ss);
  docProps_().deleteProperty('MFORM_KEY');
  const said = refreshMemberForm_(ss, form);
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'memberDaily')) {
    ScriptApp.newTrigger('memberDaily').timeBased().everyDays(1).atHour(6).inTimezone(CONFIG.timezone).create();
  }
  try { styleMemberTab_(ss); } catch (e) { Logger.log('Member tab: ' + e.message); }
  try { writeAutoLinks_(ss); } catch (e) { Logger.log('Links: ' + e.message); }   // member form + pictures folder links
  let folderNote = '';
  try {                                                           // make the pictures folder next to the tracker, once
    if (!memePicFolder_()) {
      const parents = DriveApp.getFileById(ss.getId()).getParents();
      (parents.hasNext() ? parents.next() : DriveApp.getRootFolder()).createFolder(M_PIC_FOLDER);
      folderNote = ' (folder just made next to this tracker)';
    }
  } catch (e) { Logger.log('Member pictures folder: ' + e.message); }
  const pics = memberPictures_();
  const nPics = Object.keys(pics).reduce((n, k) => n + pics[k].length, 0);
  notify_('👥 Member check-in is ready\n\n' +
    `• ${said}\n• Form: ${form.getPublishedUrl()}\n• Meme pictures: ${nPics}` +
    (nPics ? '' : '\n\nNo pictures? 🛠 Setup tools › 🖼 Connect the pictures folder.') +
    '\n\nNext: 🛠 Setup tools › 🔐 Create / update Campus Lead portals.');
}

function createMemberForm_(ss) {
  const st = getSettings_();
  const form = FormApp.create(`START Member Check-in${st.region ? ' – ' + st.region : ''}`);
  try { form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED); }   // must sign in with Google: the email is real
  catch (e) { retry_(() => form.setCollectEmail(true)); }
  form.setProgressBar(true).setShowLinkToRespondAgain(false);
  form.addMultipleChoiceItem().setTitle(MQ.vibe).setChoiceValues(M_VIBES).setRequired(true);
  form.addPageBreakItem().setTitle(M_PAGES.events);
  form.addCheckboxItem().setTitle(MQ.events).setChoiceValues([M_MISSED]).setRequired(true);
  form.addPageBreakItem().setTitle(M_PAGES.learning);
  form.addMultipleChoiceItem().setTitle(MQ.learning).setChoiceValues(M_LEARN).setRequired(true);
  form.addPageBreakItem().setTitle(M_PAGES.help);
  form.addMultipleChoiceItem().setTitle(MQ.help).setChoiceValues(M_HELP).setRequired(true);
  form.addParagraphTextItem().setTitle(MQ.tea);
  form.addCheckboxItem().setTitle(CONSENT_TITLE).setChoiceValues([CONSENT_AGREE]).setRequired(true);
  docProps_().setProperty('MFORM_ID', form.getId());
  try { DriveApp.getFileById(form.getId()).moveTo(portalFolder_(ss)); } catch (e) { /* stays in My Drive */ }
  retry_(() => form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId()));
  for (let i = 0; i < 10; i++) {                                 // "Form Responses N" -> Member check-ins
    SpreadsheetApp.flush();
    const fresh = SpreadsheetApp.openById(ss.getId());
    const sh = fresh.getSheets().find(s => (s.getFormUrl() || '').indexOf(form.getId()) !== -1);
    if (sh) {
      const old = fresh.getSheetByName(TAB.mcheckins);
      if (old && old.getSheetId() !== sh.getSheetId()) old.setName(TAB.mcheckins + ' (old)');
      sh.setName(TAB.mcheckins);
      break;
    }
    Utilities.sleep(1500);
  }
  return form;
}

// The "MEMBER CHECK-IN" picture at the very top of the member form (added once; a newer picture replaces it)
function memberFormBanner_(form) {
  const props = docProps_();
  if (props.getProperty('MFORM_BANNER') === PICTURES.mform) return;          // already the current picture
  const blob = typeof picture_ === 'function' ? picture_('mform') : null;
  if (!blob) return;                                                         // picture not added yet
  let memes = [];
  try { memes = Object.values(JSON.parse(props.getProperty('MFORM_IMG') || '{}')); } catch (e) { memes = []; }
  let banner = formBannerItem_(form, 'MFORM_BANNER_ID', memes);
  if (banner) retry_(() => banner.asImageItem().setImage(blob));
  else banner = form.getItemById(retry_(() => form.addImageItem().setImage(blob).setTitle('')
                  .setAlignment(FormApp.Alignment.CENTER).setWidth(740)).getId());
  props.setProperty('MFORM_BANNER_ID', String(banner.getId()));
  if (banner.getIndex() !== 0) retry_(() => form.moveItem(banner, 0));
  props.setProperty('MFORM_BANNER', PICTURES.mform);
}

// Copies a Campus Lead's weekly answer into the Campus Lead check-ins tab (one row per school per week), so the
// school's status, follow-up emails, Dashboard and portal keep working exactly as before. Returns the copied answers.
function mirrorLeadCheckin_(ss, who, v, c) {
  const ck = ss.getSheetByName(TAB.checkins);
  const pick = (val, from, to) => { const i = from.findIndex(x => samePick_(val, x)); return i === -1 ? String(val || '') : to[i]; };
  const out = {
    help: c.help >= 0 ? pick(v[c.help], M_HELP, HELP_CHOICES) : HELP_CHOICES[0],
    concern: c.tea >= 0 ? String(v[c.tea] || '') : '',
    vibe: c.vibe >= 0 ? String(v[c.vibe] || '') : '',
  };
  if (!ck) return out;
  const k = ckCols_(ck);
  const heads = ck.getRange(1, 1, 1, k.width).getValues()[0].map(x => plainHeading_(x));
  const row = new Array(k.width).fill('');
  row[0] = v[0] instanceof Date ? v[0] : new Date();
  row[k.school] = who.school;
  if (k.name >= 0) row[k.name] = who.name;
  if (k.help >= 0) row[k.help] = out.help;
  if (k.concern >= 0) row[k.concern] = out.concern;
  if (k.vibe >= 0) row[k.vibe] = pick(out.vibe, M_VIBES, L_VIBES);
  const ci = heads.indexOf(CONSENT_TITLE);
  if (ci >= 0) row[ci] = CONSENT_AGREE;
  // finding the free row and writing it happen together (two Leads answering at the same moment never overwrite each other)
  let lk = null, mine = false;
  try { lk = docLock_(); mine = !(typeof lk.hasLock === 'function' && lk.hasLock()) && lk.tryLock(20000); } catch (e) { lk = null; }
  try {
  const ts = ck.getRange(2, 1, Math.max(ck.getMaxRows() - 1, 1), k.width).getValues();
  let last = ts.length; while (last > 0 && ts[last - 1][0] === '') last--;
  let at = last + 2;
  let wk = ''; try { wk = memberWeekKey_(row[0]); } catch (e) { wk = ''; }   // answered again this week: replace, not add
  for (let i = last - 1; wk && i >= 0; i--) {
    const r = ts[i];
    if (r[0] instanceof Date && String(r[k.school]).trim() === who.school && (k.name < 0 || String(r[k.name]).trim() === who.name) && memberWeekKey_(r[0]) === wk) { at = i + 2; break; }
  }
  if (at > ck.getMaxRows()) ck.insertRowsAfter(ck.getMaxRows(), 20);
  ck.getRange(at, 1, 1, k.width).setValues([row]);
  } finally { if (mine && lk) lk.releaseLock(); }
  return out;
}

// Puts this week's version in place. Only touches the form when something changed. Returns a short message.
function refreshMemberForm_(ss, form) {
  form = form || memberForm_();
  if (!form) return '';
  try { memberFormBanner_(form); } catch (e) { Logger.log('Member form banner: ' + e.message); }
  try { renameFormItems_(form); } catch (e) { Logger.log('Member form wording: ' + e.message); }
  const st = getSettings_(), w = memberWeek_(new Date());
  const events = memberEventChoices_(ss, w), learn = learningOn_();
  const pics = memberPictures_(), weekNo = Math.floor(w.due.getTime() / (7 * 86400000));
  const pick = {};
  M_PIC_SLOTS.forEach(k => { const list = pics[k] || []; pick[k] = list.length ? list[weekNo % list.length] : null; });
  const key = JSON.stringify(['v3', ymd_(w.due), events, learn, M_PIC_SLOTS.map(k => pick[k] && pick[k].getId()), st.region]);
  const props = docProps_();
  const said = `This week (due ${Utilities.formatDate(w.due, CONFIG.timezone, 'EEE, MMM d')}): ` +
    (events.length ? `${events.length} event(s) to tick` : 'no START events, so the short version') +
    (learn ? ' + the learning question.' : '.');
  if (props.getProperty('MFORM_KEY') === key) return said;

  const items = () => form.getItems();
  const find = t => items().find(x => x.getTitle() === t);
  form.setTitle(`START Weekly Check-in${st.region ? ' – ' + st.region : ''} · due ${Utilities.formatDate(w.due, CONFIG.timezone, 'EEE, MMM d')}`);
  form.setDescription(`Weekly check-in for all START scholars, Campus Leads included (${events.length || learn ? 'about 1 minute' : '30 seconds'}).\n` +
    'Sign in with the email you gave START.\n\n' +
    'Only your Campus Lead and Regional Representative see your answers.\n\n' + privacyNote_());
  form.setConfirmationMessage('Thank you. Your check-in for this week is recorded. If you asked for urgent support, your Campus Lead and Regional Representative have been notified.');
  // this week's events
  [[MQ.vibe, M_VIBES], [MQ.learning, M_LEARN], [MQ.help, M_HELP]].forEach(([t, c]) => {      // current, plain choices
    const it = find(t);
    if (it) { try { it.asMultipleChoiceItem().setChoiceValues(c); } catch (e) { Logger.log('Choices: ' + e.message); } }
  });
  const evItem = find(MQ.events);
  if (evItem) evItem.asCheckboxItem().setChoiceValues(events.concat([M_MISSED]));
  // skip pages that are not needed this week
  const pb = t => { const x = find(t); return x ? x.asPageBreakItem() : null; };
  const pEv = pb(M_PAGES.events), pLe = pb(M_PAGES.learning), pHe = pb(M_PAGES.help);
  const NEXT = FormApp.PageNavigationType.CONTINUE;
  if (pEv && pLe && pHe) {
    pEv.setGoToPage(events.length ? NEXT : (learn ? pLe : pHe));    // after the vibe page
    pLe.setGoToPage(learn ? NEXT : pHe);                            // after the events page
  }
  // meme pictures above each question (rotate every week; none = no picture)
  let img = {};
  try { img = JSON.parse(props.getProperty('MFORM_IMG') || '{}'); } catch (e) { img = {}; }
  const qTitle = { vibe: MQ.vibe, events: MQ.events, learning: MQ.learning, help: MQ.help };
  M_PIC_SLOTS.forEach(k => {
    let it = null;
    if (img[k]) { try { it = form.getItemById(img[k]); } catch (e) { it = null; } }
    if (!pick[k]) { if (it) form.deleteItem(it); delete img[k]; return; }
    if (!it) { img[k] = form.addImageItem().getId(); it = form.getItemById(img[k]); }   // a plain Item, like the saved ones
    it.asImageItem().setImage(pick[k].getBlob()).setTitle('').setAlignment(FormApp.Alignment.CENTER).setWidth(480);
    const q = find(qTitle[k]);
    if (q) form.moveItem(it, q.getIndex() > it.getIndex() ? q.getIndex() - 1 : q.getIndex());
  });
  props.setProperty('MFORM_IMG', JSON.stringify(img));
  // clean out what does not belong (empty picture slots from earlier tries, old or doubled sections)
  const memeIds = Object.values(img), mBanner = formBannerItem_(form, 'MFORM_BANNER_ID', memeIds);
  const removed = cleanFormItems_(form, { titles: Object.values(MQ).concat([CONSENT_TITLE]), pages: Object.values(M_PAGES),
    images: memeIds.concat(mBanner ? [mBanner.getId()] : []) });
  if (removed) Logger.log('Member form: removed ' + removed + ' leftover item(s)');
  if (mBanner && mBanner.getIndex() !== 0) form.moveItem(mBanner, 0);
  props.setProperty('MFORM_KEY', key);
  return said;
}

// { vibe: [files], events: [...], … } from the pictures folder (sorted by name). Empty when there is no folder.
function memberPictures_() {
  const out = {};
  try {
    const all = memePicFiles_();
    if (!all.length) return out;
    M_PIC_SLOTS.forEach(k => { out[k] = all.filter(f => f.getName().toLowerCase().indexOf(k + '-') === 0); });
  } catch (e) { Logger.log('Member pictures: ' + e.message); }
  return out;
}

// Every morning (email account only): new week, new events, pictures rotate
function memberDaily(e) {
  if (!ownerGuard_(e)) return;
  try { refreshMemberForm_(SpreadsheetApp.getActive()); } catch (err) { Logger.log('Member form: ' + err.message); }
  try { refreshLeadForm_(SpreadsheetApp.getActive()); } catch (err) { Logger.log('Campus Lead form: ' + err.message); }
}

/* ---------- answers ---------- */

// Column positions in the Member check-ins tab, found by heading
function memberCols_(sh) {
  const h = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(x => plainHeading_(x));
  const f = t => h.indexOf(t);
  return { ts: 0, email: emailCol_(h), name: f('Name'), school: f('School'), role: f('Role'), vibe: colOf_(h, MQ.vibe), events: colOf_(h, MQ.events),
           learning: colOf_(h, MQ.learning), help: colOf_(h, MQ.help), tea: colOf_(h, MQ.tea), width: h.length };
}

// The column with the scholar's email: "Email Address" (Google Forms), "Email", "Gmail", "Email address (Gmail)"…
// but never "Other emails" or a Lead's email column.
function emailCol_(h) {
  let i = h.findIndex(x => /^e-?mail( address)?$/i.test(String(x).trim()));
  if (i < 0) i = h.findIndex(x => /e-?mail|gmail/i.test(x) && !/other|lead|alt|backup|contact person/i.test(x));
  return i;
}
// Older versions followed a still-linked Google Form: its tab took the name "Check-in answers" and the real tab became
// "Check-in answers (old …)". Put the real one back (the form tab keeps its answers under a clear name). Nothing is deleted.
function undoFormSwap_(ss) {
  if (typeof WEB_CHECKIN === 'undefined' || !WEB_CHECKIN) return false;
  const cur = ss.getSheetByName(TAB.mcheckins);
  if (!cur) return false;
  let linked = ''; try { linked = cur.getFormUrl() || ''; } catch (e) { linked = ''; }
  if (!linked && emailCol_(cur.getRange(1, 1, 1, Math.max(cur.getLastColumn(), 1)).getValues()[0].map(x => plainHeading_(x))) >= 0) return false;
  const real = ss.getSheets().find(x => x.getName().indexOf(TAB.mcheckins + ' (old') === 0 &&
    emailCol_(x.getRange(1, 1, 1, Math.max(x.getLastColumn(), 1)).getValues()[0].map(y => plainHeading_(y))) >= 0);
  if (!real) return false;
  cur.setName('Google Form answers (not used)');
  real.setName(TAB.mcheckins);
  MEMBER_CACHE_ = null;
  return true;
}

// Older or hand-made answer tabs without an email column get one (at the end), so check-ins never stop
function ensureEmailCol_(sh) {
  const c = memberCols_(sh);
  if (c.email >= 0) return c;
  const head = sh.getRange(1, 1, 1, c.width).getValues()[0].map(x => plainHeading_(x));
  if (!head.filter(Boolean).length) {                                    // an empty tab: the full set of headings
    const full = MEMBER_HEAD_();
    if (sh.getMaxColumns() < full.length) sh.insertColumnsAfter(sh.getMaxColumns(), full.length - sh.getMaxColumns());
    sh.getRange(1, 1, 1, full.length).setValues([full]); sh.setFrozenRows(1);
  } else {
    const at = c.width + 1;
    if (sh.getMaxColumns() < at) sh.insertColumnsAfter(sh.getMaxColumns(), at - sh.getMaxColumns());
    sh.getRange(1, at).setValue('Email Address');
  }
  SpreadsheetApp.flush();
  Logger.log('Added an email column to ' + sh.getName());
  return memberCols_(sh);
}

// email -> { name, school, status } from Scholars
// Every email a scholar can sign in with -> that scholar. "primary" = their main email (Contact, column E): the ONLY one
// START ever emails. Other emails (column K) are for signing in only, and count as the same person (one account, one history).
var ALT_EMAIL_COL = 11;                       // K: Other emails (sign-in only)
var ALT_EMAIL_HEAD = 'Other emails (sign-in only)';
var SCHOLAR_KIND_COL = 12, SCHOLAR_FROM_COL = 13;            // L: Scholarship (DOST) · M: Added from (filled when the RR approves a join request)
var SCHOLAR_EXTRA_HEAD = [[SCHOLAR_KIND_COL, 'Scholarship (DOST)', 'RA 7687, Merit, RA 10612 or JLSS. Filled from the join request when you approve one; type or fix it anytime.'],
                          [SCHOLAR_FROM_COL, 'Added from', 'How this scholar got on the list: blank = your ONSSE masterlist; "Join request · date" = approved on the START website.']];
function scholarByEmail_(ss) {
  const out = {}, sh = ss.getSheetByName(TAB.scholars);
  if (!sh || sh.getLastRow() < 2) { leadRows_(ss, out).forEach(r => { out[r[4]] = { name: r[0], school: r[1], status: 'Active', primary: r[4], leadOnly: true }; }); return out; }
  const w = Math.min(Math.max(sh.getLastColumn(), 6), ALT_EMAIL_COL);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, w).getValues();
  rows.forEach(r => {                                            // main emails first: they always win
    const e = emailOf_(r[4]);
    if (e && !out[e]) out[e] = { name: String(r[0]).trim(), school: String(r[1]).trim(), status: String(r[5]).trim(), primary: e };
  });
  if (w >= ALT_EMAIL_COL) rows.forEach(r => {
    const e = emailOf_(r[4]); if (!e) return;                    // no main email: nothing to send to, so no sign-in either
    altEmails_(r[ALT_EMAIL_COL - 1]).forEach(a => { if (!out[a]) out[a] = Object.assign({}, out[e], { primary: e }); });
  });
  leadRows_(ss, out).forEach(r => { out[r[4]] = { name: r[0], school: r[1], status: 'Active', primary: r[4], leadOnly: true }; });
  return out;
}
// Campus Leads answer the weekly check-in too. A Lead who is not (yet) Active in Scholars still counts as a scholar of
// their school for the check-in (pop-up, reminders, streak), as rows shaped like Scholars A:F. Website check-in only.
function leadRows_(ss, known) {
  if (typeof WEB_CHECKIN === 'undefined' || !WEB_CHECKIN) return [];
  const seen = {};
  if (known) Object.keys(known).forEach(e => { seen[e] = true; });
  else {
    const sh = ss.getSheetByName(TAB.scholars);
    if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, Math.min(Math.max(sh.getLastColumn(), 6), ALT_EMAIL_COL)).getValues().forEach(r => {
      [emailOf_(r[4])].concat(r.length >= ALT_EMAIL_COL ? altEmails_(r[ALT_EMAIL_COL - 1]) : []).forEach(e => { if (e) seen[e] = true; });
    });
  }
  return schoolRows_(ss).filter(x => x.focal && x.email && !x.inactive && !seen[x.email])
    .map(x => { seen[x.email] = true; return [x.focal, x.name, '', '', x.email, 'Active']; });
}
function altEmails_(cell) {
  return (String(cell || '').match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g) || []).map(x => x.toLowerCase());
}
// The scholar's main email for any of their emails (unknown emails come back as they are)
function primaryEmail_(ss, email) {
  email = emailOf_(email);
  if (!email) return '';
  try { const p = scholarByEmail_(ss)[email]; return p && p.primary ? p.primary : email; } catch (e) { return email; }
}

// Called from onCheckin (06_Form) when the answer came from the member form
function onMemberCheckin_(e) {
  const ss = SpreadsheetApp.getActive();
  const sh = e.range.getSheet(), row = e.range.getRow();
  const c = memberCols_(sh);
  const v = sh.getRange(row, 1, 1, c.width).getValues()[0];
  const email = c.email >= 0 ? emailOf_(v[c.email]) : '';
  const who = scholarByEmail_(ss)[email];
  if (!who) {                                                    // not in Scholars (yet): kept aside, never lost
    try { parkUnrecognised_(ss, sh, row, c); } catch (err) { Logger.log('Not recognised: ' + err.message); }
    return;
  }
  MEMBER_CACHE_ = null;
  try { writeMemberJoined_(ss); } catch (err) { Logger.log('Members joined: ' + err.message); }
  // A Campus Lead answering the one weekly form: it is also their school's check-in (status, follow-ups, portal)
  const sRow = schoolRows_(ss).find(s => s.name === who.school);
  if (sRow && sRow.email && sRow.email === email && oneFormOn_()) {
    try {
      const m = mirrorLeadCheckin_(ss, who, v, c);
      leadCheckinDone_(who.school, who.name, m.help, m.concern, samePick_(m.vibe, M_VIBES[3]));
    } catch (err) { Logger.log('Campus Lead check-in (one form): ' + err.message); }
    try { renderCalView_(ss); } catch (err) { /* no calendar view */ }
    return;
  }
  if (c.help >= 0 && samePick_(v[c.help], M_HELP[2])) {
    try { memberUrgentAlert_(ss, who, c.tea >= 0 ? String(v[c.tea]) : '', c.vibe >= 0 ? String(v[c.vibe]) : ''); }
    catch (err) { Logger.log('Member alert: ' + err.message); }
  }
  const p = (() => { try { return portalStore_().map[who.school]; } catch (err) { return null; } })();
  if (p && p.id && p.email) { try { refreshPortals_(who.school); } catch (err) { Logger.log('Portal after member check-in: ' + err.message); } }
  else { try { buildFocalTab_(ss); } catch (err) { Logger.log('Leads tab after member check-in: ' + err.message); } }
  try { renderCalView_(ss); } catch (err) { /* no calendar view */ }
  try { markAnswersSeen_(ss); } catch (err) { /* the hourly check will catch up */ }
}

function memberUrgentAlert_(ss, who, tea, vibe) {
  const lead = schoolRows_(ss).find(s => s.name === who.school);
  const me = Session.getEffectiveUser().getEmail();
  // the school's Campus Lead always; the RR by email only when RR_EMAILS is on or the school has no Lead email (the RR page shows it too)
  const to = (RR_EMAILS || !(lead && lead.email) ? [me] : []).concat(lead && lead.email ? [lead.email] : []).join(',');
  const html = designedEmail_({
    badge: 'Urgent', badgeColor: '#B3261E', title: ['Member', 'Needs Help'], greeting: '',
    intro: [`${displayName_(who.name)} asked for urgent help in this week's check-in. Please reach out today.`],
    extraHtml: emailAlert_(`<b>Vibe:</b> ${esc_(vibe || '—')}<br><b>What's up:</b> ${esc_(tea || '(no details given)')}`, 'red'),
    link: (DB_MODE && webHome_()) || ss.getUrl(), buttonText: DB_MODE && webHome_() ? 'Open START' : 'Open the tracker', buttonNote: '',
    chips: [['SCHOOL', who.school], ['MEMBER', displayName_(who.name)], ['WHEN', Utilities.formatDate(new Date(), CONFIG.timezone, 'MMM d, h:mm a')]],
    footer: 'Sent to the school\'s START Campus Lead and the Regional Representative only.',
  });
  sendMail_({ to, name: 'START Tracker', htmlBody: html, subject: `START · Urgent: a member needs help (${who.school})` });
}

// Everything about members, read once per run: { bySchool: { UM: { members: [...], answered, due } }, joined: { label: n } }
function memberData_(ss) {
  if (MEMBER_CACHE_) return MEMBER_CACHE_;
  const w = memberWeek_(new Date()), thisKey = ymd_(w.due);
  const out = { week: w, bySchool: {}, joined: {}, helpRecent: [] };
  const sc = ss.getSheetByName(TAB.scholars);
  const roster = (sc && sc.getLastRow() > 1 ? sc.getRange(2, 1, sc.getLastRow() - 1, 6).getValues() : []).concat(leadRows_(ss));
  const answers = {};                                             // email -> { weekKey: latest answer }
  const sh = memberTab_(ss);
  if (sh && sh.getLastRow() > 1) {
    const c = memberCols_(sh);
    const since = new Date(Date.now() - 7 * 86400000);
    sh.getRange(2, 1, sh.getLastRow() - 1, c.width).getValues().forEach(v => {
      if (!(v[0] instanceof Date) || c.email < 0) return;
      const email = emailOf_(v[c.email]); if (!email) return;
      const a = { date: v[0], vibe: c.vibe >= 0 ? String(v[c.vibe]) : '', events: c.events >= 0 ? String(v[c.events]) : '',
                  learning: c.learning >= 0 ? String(v[c.learning]) : '', help: c.help >= 0 ? String(v[c.help]) : '',
                  tea: c.tea >= 0 ? String(v[c.tea]) : '' };
      const k = memberWeekKey_(v[0]);
      (answers[email] = answers[email] || {})[k] = a;
      splitEvents_(a.events).filter(x => x && !missedPick_(x)).forEach(l => { out.joined[l] = (out.joined[l] || 0) + 1; });
      if (a.date >= since && (samePick_(a.help, M_HELP[1]) || samePick_(a.help, M_HELP[2]) || samePick_(a.vibe, M_VIBES[3]))) out.helpRecent.push(Object.assign({ email }, a));
    });
  }
  const weekKeys = n => { const ks = []; const d = new Date(w.due); for (let i = 0; i < n; i++) { ks.push(ymd_(d)); d.setDate(d.getDate() - 7); } return ks; };
  const past = weekKeys(26);
  roster.forEach(r => {
    const name = String(r[0]).trim(), school = String(r[1]).trim();
    if (!name || String(r[5]).trim() !== 'Active') return;
    const email = emailOf_(r[4]), mine = (email && answers[email]) || {};
    let streak = 0;
    for (let i = mine[thisKey] ? 0 : 1; i < past.length && mine[past[i]]; i++) streak++;   // this week may still be open
    const s = out.bySchool[school] = out.bySchool[school] || { members: [], answered: 0 };
    const now = mine[thisKey] || null;
    if (now) s.answered++;
    s.members.push({ name, email, now, streak });
  });
  Object.keys(out.bySchool).forEach(k => out.bySchool[k].members.sort((a, b) => memberRank_(a) - memberRank_(b) || a.name.localeCompare(b.name)));
  const people = scholarByEmail_(ss);
  out.helpRecent.forEach(h => { const p = people[h.email]; h.name = p ? p.name : h.email; h.school = p ? p.school : ''; });
  return (MEMBER_CACHE_ = out);
}
function memberRank_(m) {
  if (m.now && (samePick_(m.now.help, M_HELP[2]) || samePick_(m.now.vibe, M_VIBES[3]))) return 0;    // needs help: top
  if (m.now && (samePick_(m.now.help, M_HELP[1]) || samePick_(m.now.vibe, M_VIBES[2]))) return 1;
  return m.now ? 3 : 2;                                                               // not answered before answered
}

// The deadline is "near" from the day before it (e.g. Saturday for a Sunday deadline) until it passes
function memberDeadlineNear_(w) {
  const today = startOfDay_(new Date()), eve = new Date(w.due); eve.setDate(eve.getDate() - 1);
  return today >= eve && today <= w.due;
}

// Events tab › "Members joined" (per row, all dates of a repeating event added up)
function writeMemberJoined_(ss) {
  const sh = ss.getSheetByName(TAB.events);
  if (!sh || sh.getLastRow() < 2) return;
  const joined = memberData_(ss).joined;
  const far = new Date(2000, 0, 1), soon = new Date(Date.now() + 400 * 86400000);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, EV_EDIT).getValues();
  sh.getRange(2, EV.joined + 1, rows.length, 1).setValues(rows.map(r => {
    const title = String(r[EV.title]).trim(), who = String(r[EV.who]).trim() || WHO.all;
    if (!title || who === WHO.leads || who === WHO.me) return [''];
    const n = occurrences_(r, far, soon).reduce((t, d) => t + (joined[evLabel_(title, d, who)] || 0), 0);
    return [n ? n : ''];
  }));
}

// The group-chat message a Campus Lead copies (school = their school, or blank for everyone)
function memberMessage_(ss, school) {
  const form = WEB_CHECKIN ? null : memberForm_();
  if (!form && !(WEB_CHECKIN && checkinOn_())) return '';
  const link = form ? form.getPublishedUrl() : checkinLink_();
  const w = memberWeek_(new Date()), events = memberEventChoices_(ss, w)
    .filter(l => !/\(([^)]+)\) · /.test(l) || (school && l.indexOf('(' + school + ') · ') !== -1));
  // a short PERSONAL nudge for Messenger: only for members who have not answered yet (everyone else got it by email)
  return `Hi! Friendly reminder: your START weekly check-in isn't in yet (1 minute lang!)\n` +
    `${link}\n` +
    `Due: ${Utilities.formatDate(w.due, CONFIG.timezone, 'EEEE, MMM d')}. Sign in with the email you gave START. ` +
    `(It's also in your email inbox.)` +
    (events.length ? `\n\nThis week's START events:\n${events.map(l => '• ' + l).join('\n')}` : '') +
    '\n\nSalamat!';
}

// Menu 📋 Copy a message for Messenger › For members
function showMemberMessage() {
  sheetUser_();                                   // sheet only: never from the website
  requireSetup_();
  const ss = SpreadsheetApp.getActive();
  const msg = memberMessage_(ss, '');
  if (!msg) return notify_('Set up the member check-in first: 🛠 Setup tools › 📝 Set up / update member check-in.');
  dlgCopy_('👥 Message for members', 'For members who haven\'t answered yet', msg, '', 220);
}

// The Member check-ins tab: tidy heading and the "About this tab" panel
// The tab that receives member answers. Google first calls it "Form Responses N": renamed here so it always is "Member check-ins".
function memberTab_(ss) {
  // website check-in: ALWAYS the Check-in answers tab. An old Google Form that is still linked is never followed
  // (its tab has different columns, often no email) and never renames the real tab.
  if (typeof WEB_CHECKIN !== 'undefined' && WEB_CHECKIN) return ss.getSheetByName(TAB.mcheckins) || ss.getSheetByName(OLD_MEMBER_TAB);
  let linked = null;
  try {
    const id = docProps_().getProperty('MFORM_ID');
    linked = id ? ss.getSheets().find(s => (s.getFormUrl() || '').indexOf(id) !== -1) : null;
  } catch (e) { linked = null; }
  if (!linked) return ss.getSheetByName(TAB.mcheckins) || ss.getSheetByName(OLD_MEMBER_TAB);   // the old tab until ① moves it
  if (linked.getName() !== TAB.mcheckins) {
    const old = ss.getSheetByName(TAB.mcheckins);
    if (old) old.setName(TAB.mcheckins + ' (old ' + Utilities.formatDate(new Date(), CONFIG.timezone, 'MMM d') + ')');
    linked.setName(TAB.mcheckins);
  }
  return linked;
}

// True for the tab of the member form (by its name, or by its form link when Google named it "Form Responses N")
function isMemberSheet_(sh) {
  if (!sh) return false;
  if (sh.getName() === TAB.mcheckins || sh.getName() === OLD_MEMBER_TAB) return true;
  const id = docProps_().getProperty('MFORM_ID');
  return !!id && (sh.getFormUrl() || '').indexOf(id) !== -1;
}

function styleMemberTab_(ss) {
  if (DB_MODE) { styleNewerTables_(ss); return; }             // database mode: the same look as every other table (12_Style)
  const sh = memberTab_(ss);
  if (!sh) return;
  const w = Math.max(sh.getLastColumn(), 1);
  header_(sh.getRange(1, 1, 1, w), THEME.navy);
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.setTabColor(THEME.tabs.mcheckins);
  placeTabBanner_(sh, Math.max(formColCount_(sh), 4));
  placeTutorial_(sh);
}

/* ---------- 🔐 the member section of a Campus Lead's portal (called by writePortal_) ---------- */

// Writes the section starting at row r; returns the next free row
// Everything the portal needs about a school's members, in ONE list (answers + who they are), most urgent first.
// cols/rows follow the RR's portal choices: answers (portalMembers) and scholar details (portalScholars).
function portalMembers_(s, pst) {
  const ss = SpreadsheetApp.getActive();
  const form = (() => { try { return checkinOn_(); } catch (e) { return false; } })();
  const answers = !!form && pst.portalMembers !== false, details = pst.portalScholars !== false;
  const out = { answers, members: [], answered: 0, help: [], toMessage: [], noEmail: [], message: '', cols: [], rows: [], due: s.due };
  let info = null, near = false;
  if (form) {
    const d = memberData_(ss); info = d.bySchool[s.name] || { members: [], answered: 0 };
    out.due = d.week.due; near = memberDeadlineNear_(d.week);
  }
  const bySch = {}; (s.scholars || []).forEach(x => { bySch[String(x.name).trim()] = x; });
  const list = info ? info.members
    : (s.scholars || []).filter(x => ['Graduated', 'Dropped'].indexOf(String(x.status)) === -1).map(x => ({ name: x.name, email: 'x', now: null }));
  out.members = list;
  out.answered = list.filter(m => m.now).length;
  if (answers) {
    out.help = list.filter(m => typeof memberRank_ === 'function' && memberRank_(m) === 0);
    out.toMessage = list.filter(m => !m.now && m.email);
    out.noEmail = list.filter(m => !m.email && !m.now);
    if (out.toMessage.length) out.message = memberMessage_(ss, s.name);
  }
  if (!answers && !details) return out;
  out.cols = [['Member']];
  if (answers) out.cols.push(['This week'], ['Vibe'], ['Help · what\'s up']);
  if (details) out.cols.push(['Year level'], ['Program']);
  const sorted = list.slice().sort((a, b) => (answers ? memberRank_(a) - memberRank_(b) : 0) || String(a.name).localeCompare(String(b.name)));
  out.rows = sorted.map(m => {
    const x = bySch[String(m.name).trim()] || {};
    const cells = [m.name];
    let chip = null;
    if (answers) {
      const rank = memberRank_(m);
      const status = rank === 0 ? '🆘 Needs help' : m.now ? '✅ Answered' + (m.streak > 1 ? ` · 🔥${m.streak}` : '') :
                     !m.email ? '📭 No email yet' : near ? '🔴 Not yet' : '⏳ Not yet';
      chip = rank === 0 ? CHIP.red : m.now ? (rank === 1 ? CHIP.yellow : CHIP.green) : !m.email ? CHIP.grey : near ? CHIP.red : CHIP.yellow;
      cells.push(status, m.now ? m.now.vibe || '' : '', m.now ? [m.now.help, m.now.tea].filter(Boolean).join(' · ') : '');
    }
    if (details) cells.push(x.level || '', x.program || '');
    return { cells, chip, email: m.email && m.email !== 'x' ? m.email : '', me: String(m.name).trim() === String(s.focal).trim() };
  });
  return out;
}

function writeMemberSection_(sh, r, s, NC, merge) {
  const ss = SpreadsheetApp.getActive();
  if (!checkinOn_()) return r;
  const d = memberData_(ss), info = d.bySchool[s.name] || { members: [], answered: 0 };
  const due = Utilities.formatDate(d.week.due, CONFIG.timezone, 'EEE, MMM d');
  const near = memberDeadlineNear_(d.week), late = info.members.filter(m => !m.now).length;
  merge(r, 1, NC).setValue(`👥  MY MEMBERS THIS WEEK  ·  ${info.answered} of ${info.members.length} responded  ·  due ${due}` +
    (near && late ? `  ·  🔴 ${late} to message` : ''));
  header_(sh.getRange(r, 1, 1, NC));
  r++;
  merge(r, 1, NC).setValue(`📋 Members get the check-in by email every ${getSettings_().mOpenDay}. Message ONLY those marked ⏳ / 🔴 below (Messenger), with this:\n\n` + memberMessage_(ss, s.name))
    .setWrap(true).setFontSize(9).setBackground('#FFF6BF').setVerticalAlignment('top');
  sh.setRowHeight(r, 150);
  r++;
  // ✨ members with no email can't answer or get the calendar: list them + a message the Lead can send
  const noEmail = info.members.filter(m => !m.email);
  if (noEmail.length) {
    const st = getSettings_();
    let rrMail = ''; try { rrMail = Session.getEffectiveUser().getEmail(); } catch (e) { /* not available */ }
    merge(r, 1, NC).setValue(`📭 ${noEmail.length} member(s) have NO email on file yet, so they can't answer the check-in or see the START calendar: ` +
      noEmail.map(m => displayName_(m.name)).join(', ') + '\n\nSend them: "Hi! Please email ' + (st.rrName || 'our START Regional Representative') +
      `${rrMail ? ' (' + rrMail + ')' : ''} the Google email you use, so START can add you to the weekly check-in and the calendar. Salamat!"`)
      .setWrap(true).setFontSize(9).setBackground('#FFE2B8').setVerticalAlignment('top');
    sh.setRowHeight(r, 78);
    r++;
  }
  sh.getRange(r, 1, 1, NC).setValues([['Member', 'This week (responded?)', 'Vibe', 'Events joined', 'Learning', 'Help & what\'s up']])
    .setFontWeight('bold').setFontSize(9).setFontColor(THEME.muted).setBackground(THEME.band);
  r++;
  const rows = info.members.map(m => [
    m.name, (m.now ? '✅ Responded' : near ? '🔴 Not yet: message them!' : '⏳ Not yet') + (m.streak ? ` · 🔥${m.streak}` : ''),
    m.now ? m.now.vibe : '', m.now ? splitEvents_(m.now.events).filter(x => x && !missedPick_(x)).length || (m.now.events ? '0' : '') : '',
    m.now ? m.now.learning : '', m.now ? [m.now.help, m.now.tea].filter(Boolean).join(' · ') : (m.email ? '' : '📭 No email yet: tell your RR')]);
  if (sh.getMaxRows() < r + rows.length + 8) sh.insertRowsAfter(sh.getMaxRows(), r + rows.length + 8 - sh.getMaxRows());
  if (rows.length) {
    sh.getRange(r, 1, rows.length, NC).setValues(rows).setWrap(true);
    sh.getRange(r, 1, rows.length, 1).setFontWeight('bold');
    info.members.forEach((m, i) => {
      const rank = memberRank_(m);
      const chip = rank === 0 ? CHIP.red : rank === 1 ? CHIP.yellow : m.now ? CHIP.green : near ? CHIP.red : CHIP.grey;
      if (chip) sh.getRange(r + i, 2).setBackground(chip.bg).setFontColor(chip.fg).setFontWeight('bold');
      if (rank === 0) sh.getRange(r + i, 6).setBackground(CHIP.red.bg);
    });
  } else {
    merge(r, 1, NC).setValue('No active members listed for your school yet.').setFontColor(THEME.muted);
  }
  const n = Math.max(rows.length, 1);
  sh.getRange(r - 1, 1, n + 1, NC).setBorder(true, true, true, true, true, true, TABLE_LINE, SpreadsheetApp.BorderStyle.SOLID);
  return r + n + 1;
}

/* ---------- for your weekly summary and 🗓 Today ---------- */

function memberSummaryHtml_(ss) {
  if (!checkinOn_()) return '';
  const d = memberData_(ss);
  const rows = Object.keys(d.bySchool).sort().map(k => {
    const s = d.bySchool[k], need = s.members.filter(m => memberRank_(m) <= 1).length;
    return [k, `${s.answered} / ${s.members.length}`, need ? need + ' 🙋' : ''];
  });
  const total = rows.reduce((t, r) => t + Number(r[1].split(' / ')[1]), 0);
  const answered = Object.keys(d.bySchool).reduce((t, k) => t + d.bySchool[k].answered, 0);
  return emailSection_(`Member check-ins: ${answered} of ${total} answered (due ${Utilities.formatDate(d.week.due, CONFIG.timezone, 'MMM d')})`) +
         emailTable_(['School', 'Answered', 'Need help'], rows);
}

/* ================================================================
 *  🧑‍🏫 THE CAMPUS LEAD FORM, same idea: memes + it changes with the calendar
 *   Pages:  school & name  →  🎉 events (only weeks with START events)  →  📚 DataCamp (only while it runs)
 *           →  ✨ kamustahan + need help + concerns
 *   The question titles stay the same (your Weekly Check-in history keeps working); the fun is in the
 *   page titles, the new kamustahan question and the pictures.
 *   Pictures: same folder as the members. Files named lead-vibe-1.jpg, lead-events-1.jpg … are used first;
 *   without them the member pictures (vibe-1.jpg …) are shared.
 * ================================================================ */

var LQ_VIBE = 'How are you doing as Campus Lead this week?';
var L_VIBES = ['🔥 Very good', '🙂 Good', '😵‍💫 Overloaded', '🆘 I need support'];
var L_PAGES = { events: 'START events', learning: 'DataCamp', vibe: 'About you' };
// Only while the member check-in is on: did the Lead pass the link to the members?
var LQ_SHARED = 'Did you share the member check-in link with your members?';
var L_SHARED = ['✅ Yes', '⏳ Not yet, I will today', '🚫 No (tell us why below)'];

// The form's banner picture: the one remembered under propKey, else the picture at the very top
function formBannerItem_(form, propKey, memeIds) {
  const props = docProps_();
  const id = props.getProperty(propKey);
  if (id) { try { const it = form.getItemById(Number(id)); if (it && it.getType() === FormApp.ItemType.IMAGE) return it; } catch (e) { /* gone */ } }
  const imgs = form.getItems().filter(x => x.getType() === FormApp.ItemType.IMAGE && memeIds.indexOf(x.getId()) === -1);
  const top = imgs.find(x => x.getIndex() === 0) || imgs[0] || null;
  if (top) props.setProperty(propKey, String(top.getId()));
  return top;
}

// Removes everything the tracker did not put in the form: empty picture slots, old or doubled sections,
// old questions and blank items. keep = { titles: [...], pages: [...], images: [ids] }. Returns how many were removed.
function cleanFormItems_(form, keep) {
  const seen = {};
  let removed = 0;
  form.getItems().forEach(it => {
    const type = it.getType(), title = String(it.getTitle() || '').trim();
    let ok;
    if (type === FormApp.ItemType.IMAGE) ok = keep.images.indexOf(it.getId()) !== -1;
    else if (type === FormApp.ItemType.PAGE_BREAK) ok = keep.pages.indexOf(title) !== -1 && !seen['p:' + title];
    else if (type === FormApp.ItemType.SECTION_HEADER) ok = false;
    else ok = !!title && keep.titles.indexOf(title) !== -1 && !seen['q:' + title];
    if (type === FormApp.ItemType.PAGE_BREAK) seen['p:' + title] = true; else seen['q:' + title] = true;
    if (!ok) { try { retry_(() => form.deleteItem(it)); removed++; } catch (e) { Logger.log('Not removed: ' + title + ' ' + e.message); } }
  });
  return removed;
}

// This week's events a Campus Lead should know about (everything except your private ones)
function leadEventChoices_(ss, w) {
  const sh = ss.getSheetByName(TAB.events);
  if (!sh || sh.getLastRow() < 2) return [];
  const out = [];
  sh.getRange(2, 1, sh.getLastRow() - 1, EV_EDIT).getValues().forEach(r => {
    const title = String(r[EV.title]).trim(), who = String(r[EV.who]).trim() || WHO.all;
    if (!title || who === WHO.me) return;
    occurrences_(r, w.from, w.to).forEach(d => out.push({ label: evLabel_(title, d, who === WHO.leads ? '' : who), date: d }));
  });
  return out.sort((a, b) => a.date - b.date).map(x => x.label);
}

// Pictures for the Lead form: lead-<slot>-N first, else the member ones
function leadPictures_() {
  const out = {};
  try {
    const all = memePicFiles_();
    if (!all.length) return out;
    M_PIC_SLOTS.forEach(k => {
      const lead = all.filter(f => f.getName().toLowerCase().indexOf('lead-' + k + '-') === 0);
      out[k] = lead.length ? lead : all.filter(f => f.getName().toLowerCase().indexOf(k + '-') === 0);
    });
  } catch (e) { Logger.log('Lead pictures: ' + e.message); }
  return out;
}

// Ids of the meme picture items in the Lead form (so the banner code never mistakes one for the banner)
function leadMemeIds_() {
  try { return Object.values(JSON.parse(docProps_().getProperty('LFORM_IMG') || '{}')); }
  catch (e) { return []; }
}

// Puts this week's version of the Campus Lead form in place. Only touches the form when something changed.
function refreshLeadForm_(ss, form) {
  try { form = form || getForm_(); } catch (e) { return ''; }
  if (!form) return '';
  if (oneFormOn_()) return closeLeadForm_(form);                  // everyone uses the one weekly form now
  const props = docProps_();
  const rw = reminderWeek_(new Date());
  const w = { from: startOfDay_(rw.monday), to: new Date(startOfDay_(rw.monday).getTime() + 7 * 86400000) };
  const events = leadEventChoices_(ss, w), learn = learningOn_();
  const pics = leadPictures_(), weekNo = Math.floor(w.from.getTime() / (7 * 86400000));
  const pick = {};
  M_PIC_SLOTS.forEach(k => { const list = pics[k] || []; pick[k] = list.length ? list[weekNo % list.length] : null; });
  try { renameFormItems_(form); } catch (e) { Logger.log('Campus Lead form wording: ' + e.message); }
  const membersOn = !!memberForm_();
  const key = JSON.stringify(['v3', ymd_(w.from), events, learn, membersOn, M_PIC_SLOTS.map(k => pick[k] && pick[k].getId())]);
  const said = `Campus Lead form (week of ${Utilities.formatDate(w.from, CONFIG.timezone, 'MMM d')}): ` +
    (events.length ? `${events.length} event(s) listed` : 'no START events, so no event page') + (learn ? ' + DataCamp page.' : '.');
  if (props.getProperty('LFORM_KEY') === key) return said;

  const items = () => retry_(() => form.getItems());
  const find = t => items().find(x => x.getTitle() === t);
  // the pages and the kamustahan question (added once)
  [L_PAGES.events, L_PAGES.learning, L_PAGES.vibe].forEach(t => { if (!find(t)) retry_(() => form.addPageBreakItem().setTitle(t)); });
  addOnce_(form, LQ_VIBE, () => form.addMultipleChoiceItem().setTitle(LQ_VIBE).setChoiceValues(L_VIBES).setRequired(true)
    .setHelpText('🆘 lets your Regional Representative know to check in with you.'));
  if (membersOn) addOnce_(form, LQ_SHARED, () => form.addMultipleChoiceItem().setTitle(LQ_SHARED).setChoiceValues(L_SHARED).setRequired(true)
    .setHelpText('The ready-to-copy message is in your portal.'));
  // keep the choices current (older forms had a number box and other wording)
  [[Q.events, EVENT_GO_CHOICES], [Q.dc, DC_CHOICES], [LQ_VIBE, L_VIBES], [LQ_SHARED, L_SHARED]].forEach(([t, c]) => {
    const it = find(t);
    if (it && it.getType() === FormApp.ItemType.MULTIPLE_CHOICE) retry_(() => it.asMultipleChoiceItem().setChoiceValues(c));
  });
  const pb = t => find(t).asPageBreakItem();
  pb(L_PAGES.events).setHelpText(events.length
    ? 'This week\'s START events:\n' + events.map(l => '• ' + l).join('\n') : '');
  pb(L_PAGES.learning).setHelpText('');
  pb(L_PAGES.vibe).setHelpText('Last page. 💙');

  // meme pictures (rotate every week; none = no picture)
  let img = {};
  try { img = JSON.parse(props.getProperty('LFORM_IMG') || '{}'); } catch (e) { img = {}; }
  M_PIC_SLOTS.forEach(k => {
    let it = null;
    if (img[k]) { try { it = form.getItemById(img[k]); } catch (e) { it = null; } }
    if (!pick[k]) { if (it) form.deleteItem(it); delete img[k]; return; }
    if (!it) { img[k] = retry_(() => form.addImageItem()).getId(); it = form.getItemById(img[k]); }
    it.asImageItem().setImage(pick[k].getBlob()).setTitle('').setAlignment(FormApp.Alignment.CENTER).setWidth(480);
  });
  props.setProperty('LFORM_IMG', JSON.stringify(img));

  // clean out what does not belong (empty picture slots, old questions, doubled sections)
  const byId = id => { try { return id ? form.getItemById(id) : null; } catch (e) { return null; } };
  const memes = Object.values(img);
  const banner = formBannerItem_(form, 'LFORM_BANNER_ID', memes);
  const removed = cleanFormItems_(form, {
    titles: [Q.school, Q.name, Q.events, Q.dc, LQ_VIBE, Q.help, Q.concern, CONSENT_TITLE].concat(membersOn ? [LQ_SHARED] : []),
    pages: [L_PAGES.events, L_PAGES.learning, L_PAGES.vibe],
    images: memes.concat(banner ? [banner.getId()] : []) });
  if (removed) Logger.log('Campus Lead form: removed ' + removed + ' leftover item(s)');

  // the full order, top to bottom
  const order = [banner, find(Q.school), find(Q.name),
    find(L_PAGES.events), byId(img.events), find(Q.events),
    find(L_PAGES.learning), byId(img.learning), find(Q.dc),
    find(L_PAGES.vibe), byId(img.vibe), find(LQ_VIBE), membersOn ? find(LQ_SHARED) : null,
    byId(img.help), find(Q.help), find(Q.concern), find(CONSENT_TITLE)]
    .filter(Boolean);
  order.forEach((it, i) => { if (it.getIndex() !== i) retry_(() => form.moveItem(it, i)); });

  // skip the pages that are not needed this week
  const NEXT = FormApp.PageNavigationType.CONTINUE;
  pb(L_PAGES.events).setGoToPage(events.length ? NEXT : (learn ? pb(L_PAGES.learning) : pb(L_PAGES.vibe)));
  pb(L_PAGES.learning).setGoToPage(learn ? NEXT : pb(L_PAGES.vibe));
  props.setProperty('LFORM_KEY', key);
  return said;
}


/* ---------- 🖼 Connect the pictures folder: banners + memes, one folder or two (paste its Drive link) ---------- */
function connectMemeFolder() {
  sheetUser_();                                   // sheet only: never from the website
  requireSetup_();
  const ui = ui_();
  const r = ui.prompt('🖼 Connect the pictures folder',
    'Open your pictures folder in Google Drive: banners (start-tracker-….jpg) and/or memes (vibe-1.jpg, help-1.jpg, lead-vibe-1.jpg …).\n' +
    'One folder with both is fine. Copy the link from the address bar and paste it here:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const m = String(r.getResponseText()).match(/folders\/([-\w]{20,})/) || String(r.getResponseText()).match(/([-\w]{25,})/);
  if (!m) { notify_('That does not look like a Google Drive folder link.\n\nIt should look like:\nhttps://drive.google.com/drive/folders/1AbC…'); return; }
  let d;
  try { d = DriveApp.getFolderById(m[1]); d.getName(); }
  catch (e) {
    notify_('❌ This Google account cannot open that folder.\n\n' +
      'In Google Drive, right-click the folder › Share › add ' + Session.getEffectiveUser().getEmail() + ' as Viewer' +
      (automationOwner_() && automationOwner_() !== String(Session.getEffectiveUser().getEmail()).toLowerCase() ? ' (and ' + automationOwner_() + ', the tracker\'s email account)' : '') +
      '. Then try again.');
    return;
  }
  const props = docProps_();
  // what is inside: banners (start-tracker-…) and/or memes
  let banners = 0;
  const known = Object.values(PICTURES).map(n => n.toLowerCase());
  withSubfolders_(d).forEach(x => {
    const fsAll = x.getFiles();
    while (fsAll.hasNext()) { const f = fsAll.next(); if (!f.isTrashed() && known.indexOf(f.getName().toLowerCase().replace(/\s*\(\d+\)(\.\w+)$/, '$1')) !== -1) banners++; }
  });
  const memes = hasMemes_(d);
  if (!banners && !memes) {
    notify_('That folder has no tracker pictures (start-tracker-….jpg) and no memes (vibe-1.jpg …).\n\nCheck the folder and the file names, then try again.');
    return;
  }
  if (memes) props.setProperty('MEME_FOLDER_ID', d.getId());
  if (banners) {                                                       // banners are looked up here first from now on
    props.setProperty('PICTURE_FOLDER', d.getId());
    props.deleteProperty('PICTURE_IDS'); props.deleteProperty('EMAIL_LOGO_ID');
    Object.keys(PICTURES).forEach(k => props.deleteProperty('EMAIL_PIC_' + k));
    PICTURE_CACHE_ = {}; EMAIL_PICS_CACHE_ = null;
    props.deleteProperty('SETUP_DONE');                                // next ① redraws every header with these pictures
  }
  props.deleteProperty('MFORM_KEY'); props.deleteProperty('LFORM_KEY');       // put the pictures in right now
  const ss = SpreadsheetApp.getActive();
  const errs = [];
  try { if (memberForm_()) refreshMemberForm_(ss); } catch (e) { errs.push('Member form: ' + e.message); }
  try { refreshLeadForm_(ss); } catch (e) { errs.push('Campus Lead form: ' + e.message); }
  try { writeAutoLinks_(ss); } catch (e) { /* links tab refreshes later */ }
  const mp = memberPictures_(), lp = leadPictures_();
  const line = (p, pre) => M_PIC_SLOTS.map(k => `${pre}${k}: ${(p[k] || []).length}`).join(' · ');
  notify_((errs.length ? '⚠ Connected, but a form was not updated\n\n' + errs.join('\n') + '\n\n' : '✅ Pictures folder connected\n\n') +
    `• Banners: ${banners} of ${Object.keys(PICTURES).length}\n` +
    '• Member form: ' + line(mp, '') + '\n' +
    '• Lead form: ' + line(lp, 'lead-') +
    (banners ? '\n\nNext: click ① Set up / fix the tracker to place the banners.' : '') +
    (automationOwner_() && automationOwner_() !== String(Session.getEffectiveUser().getEmail()).toLowerCase()
      ? '\n\nAlso share the folder with ' + automationOwner_() + ' (Viewer).' : ''));
}

// The older Campus Lead form stops taking answers; anyone opening an old link is sent to the one weekly form
function closeLeadForm_(form) {
  const url = checkinUrl_();
  try {
    const msg = 'This form has moved. Campus Leads now answer the same weekly check-in as their members:\n' + url;
    if (form.isAcceptingResponses() || docProps_().getProperty('LFORM_CLOSED') !== url) {
      retry_(() => form.setCustomClosedFormMessage(msg));
      retry_(() => form.setAcceptingResponses(false));
      docProps_().setProperty('LFORM_CLOSED', url);
    }
  } catch (e) { Logger.log('Campus Lead form not closed: ' + e.message); }
  return 'One weekly form for everyone (Campus Leads included). The older Campus Lead form is closed and points to it.';
}

/* ---------- answers from emails that are not in Scholars (yet): kept, not lost ---------- */
var UNKNOWN_TAB = 'Not recognised';

// Moves the answer to the "Not recognised" tab (with the reason) so the weekly tab only has known members
function parkUnrecognised_(ss, sh, row, c) {
  const v = sh.getRange(row, 1, 1, c.width).getValues()[0];
  const heads = sh.getRange(1, 1, 1, c.width).getValues()[0].map(x => plainHeading_(x));
  let un = ss.getSheetByName(UNKNOWN_TAB);
  if (!un) {
    un = ss.insertSheet(UNKNOWN_TAB);
    un.getRange(1, 1, 1, heads.length + 1).setValues([heads.concat(['What to do'])]);
    header_(un.getRange(1, 1, 1, heads.length + 1), THEME.navy);
    un.setFrozenRows(1).setTabColor('#E0B800');
  }
  const email = c.email >= 0 ? String(v[c.email]).trim() : '';
  un.appendRow(v.concat([`Add ${email || 'this email'} to Scholars › Contact (E) for the right scholar: this answer then counts by itself.`]));
  sh.deleteRow(row);
  try { ss.toast(`An answer from ${email} is in "${UNKNOWN_TAB}": that email is not in Scholars yet.`, 'START Tracker', 8); } catch (e) { /* no UI */ }
}

// Answers in "Not recognised" whose email is now in Scholars go back to the weekly tab and count normally.
// Runs after Scholars edits and every morning. Returns how many were brought back.
function adoptRecognised_(ss) {
  const un = ss.getSheetByName(UNKNOWN_TAB), mt = memberTab_(ss);
  if (!un || !mt || un.getLastRow() < 2) return 0;
  const who = scholarByEmail_(ss);
  const c = memberCols_(mt);
  const uh = un.getRange(1, 1, 1, un.getLastColumn()).getValues()[0].map(x => plainHeading_(x));
  const ue = uh.findIndex(x => /^email( address)?$/i.test(x));
  if (ue < 0) return 0;
  const rows = un.getRange(2, 1, un.getLastRow() - 1, uh.length).getValues();
  const back = [];
  rows.forEach((r, i) => { if (who[emailOf_(r[ue])]) back.push(i + 2); });
  back.slice().reverse().forEach(rn => {
    const r = un.getRange(rn, 1, 1, uh.length).getValues()[0];
    const line = new Array(c.width).fill('');
    const mh = mt.getRange(1, 1, 1, c.width).getValues()[0].map(x => plainHeading_(x));
    mh.forEach((h, j) => { const k = uh.indexOf(h); if (k !== -1) line[j] = r[k]; });
    mt.appendRow(line);
    un.deleteRow(rn);
    try { onMemberCheckin_({ range: { getSheet: () => mt, getRow: () => mt.getLastRow() } }); }
    catch (e) { Logger.log('Brought back, not processed: ' + e.message); }
  });
  if (un.getLastRow() < 2) { try { ss.deleteSheet(un); } catch (e) { /* keep it */ } }   // nothing left: the tab goes away
  return back.length;
}

function unrecognisedCount_(ss) {
  const un = ss.getSheetByName(UNKNOWN_TAB);
  return un ? Math.max(un.getLastRow() - 1, 0) : 0;
}


/* ---------- website check-in (no Google Forms) ---------- */

// The Check-in answers tab: one row per answer, the same questions as the website check-in
var MEMBER_HEAD_ = () => ['Timestamp', 'Email Address', 'Name', 'School', 'Role', MQ.vibe, MQ.events, MQ.learning, MQ.help, MQ.tea, CONSENT_TITLE, LATE_HEAD];
var LEAD_HEAD_ = () => ['Timestamp', Q.school, Q.name, Q.events, Q.dc, Q.help, Q.concern, LQ_VIBE, CONSENT_TITLE];

// The two answer tabs as plain tables (a new tracker gets them empty; an older one keeps its rows)
function ensureCheckinTabs_(ss) {
  try { moveToAnswersTab_(ss); } catch (e) { Logger.log('New answers tab: ' + e.message); }
  try { ensureAttendanceTab_(ss); } catch (e) { Logger.log('Event attendance tab: ' + e.message); }
  try { guardTabs_(ss); } catch (e) { Logger.log('Tab protection: ' + e.message); }
  try {                                                                     // Scholars › K: Other emails (sign-in only)
    const sc = ss.getSheetByName(TAB.scholars);
    if (sc && !String(sc.getRange(1, ALT_EMAIL_COL).getValue()).trim()) {
      sc.getRange(1, ALT_EMAIL_COL).setValue(ALT_EMAIL_HEAD).setNote('Optional. Other emails this scholar may sign in with (comma between them). ' +
        'START never emails these: every email goes to Contact (email), column E.');
      try { sc.getRange(1, 10).copyTo(sc.getRange(1, ALT_EMAIL_COL), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); } catch (e) { /* plain heading */ }
    }
    SCHOLAR_EXTRA_HEAD.forEach(([c, h, note]) => {                       // L, M: added once, after K
      if (sc && !String(sc.getRange(1, c).getValue()).trim()) {
        if (sc.getMaxColumns() < c) sc.insertColumnsAfter(sc.getMaxColumns(), c - sc.getMaxColumns());
        sc.getRange(1, c).setValue(h).setNote(note);
        try { sc.getRange(1, 10).copyTo(sc.getRange(1, c), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); } catch (e) { /* plain heading */ }
      }
    });
  } catch (e) { Logger.log('Other emails column: ' + e.message); }
  try { undoFormSwap_(ss); } catch (e) { Logger.log('Answers tab swap: ' + e.message); }
  try { alignScholarsTab_(ss); } catch (e) { Logger.log('Scholars dropdowns: ' + e.message); }
  [[TAB.mcheckins, MEMBER_HEAD_()], [TAB.checkins, LEAD_HEAD_()]].forEach(([name, head]) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const have = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => plainHeading_(x)) : [];
    if (!have.filter(Boolean).length) {
      if (sh.getMaxColumns() < head.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length - sh.getMaxColumns());
      sh.getRange(1, 1, 1, head.length).setValues([head]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('yyyy-mm-dd hh:mm');
    } else if (name === TAB.mcheckins) {                                           // an older tab: add any missing question at the end
      const miss = head.filter(h => colOf_(have, h) === -1 && !(h === 'Email Address' && emailCol_(have) >= 0));
      if (miss.length) {
        const at = have.length + 1;
        if (sh.getMaxColumns() < at + miss.length - 1) sh.insertColumnsAfter(sh.getMaxColumns(), at + miss.length - 1 - sh.getMaxColumns());
        sh.getRange(1, at, 1, miss.length).setValues([miss]);
      }
    }
  });
}

// The old Google Forms: closed (with a pointer to the website) and unlinked from the tabs. Nothing is deleted.
// Returns how many forms were retired this time.
function retireForms_(ss) {
  const props = docProps_();
  let old = []; try { old = JSON.parse(props.getProperty('OLD_FORMS') || '[]'); } catch (e) { old = []; }
  let n = 0;
  ['FORM_ID', 'MFORM_ID'].forEach(k => {
    const id = props.getProperty(k);
    if (!id) return;
    try {
      const f = FormApp.openById(id);
      const link = (() => { try { return checkinLink_(); } catch (e) { return ''; } })();
      retry_(() => f.setCustomClosedFormMessage('This form is closed. The START weekly check-in is now on the START website' + (link ? ':\n' + link : '.')));
      retry_(() => f.setAcceptingResponses(false));
      try { f.removeDestination(); } catch (e) { /* already unlinked */ }
    } catch (e) { Logger.log('Old form ' + k + ': ' + e.message); }
    old.push(id); props.deleteProperty(k); n++;
  });
  if (n) {
    props.setProperty('OLD_FORMS', JSON.stringify(old));
    ['MFORM_KEY', 'LFORM_KEY', 'MFORM_IMG', 'LFORM_IMG', 'FORM_SYNC'].forEach(k => props.deleteProperty(k));
    removeTriggers_(['onCheckin', 'memberDaily']);
  }
  return n;
}


/* ---------- the new Check-in answers tab (replaces the old Google Form tab "Member check-ins") ----------
 * Done once by ①: every old answer is copied into the new tab (same weeks, streaks, XP and history), the old tab
 * is saved in a private backup file next to the tracker, then removed from the tracker. Nothing is lost. */
var OLD_MEMBER_TAB = 'Member check-ins';

function moveToAnswersTab_(ss) {
  if (TAB.mcheckins === OLD_MEMBER_TAB) return null;
  const old = ss.getSheetByName(OLD_MEMBER_TAB);
  if (!old) return null;
  const lock = docLock_();
  if (!lock.tryLock(30000)) return null;                     // a check-in is being saved: next ① does it
  try {
    const head = MEMBER_HEAD_();
    let nw = ss.getSheetByName(TAB.mcheckins);
    if (!nw) nw = ss.insertSheet(TAB.mcheckins, old.getIndex());
    nw.clear();                                               // a half-finished earlier try starts over
    if (nw.getMaxColumns() < head.length) nw.insertColumnsAfter(nw.getMaxColumns(), head.length - nw.getMaxColumns());
    nw.getRange(1, 1, 1, head.length).setValues([head]);
    nw.setFrozenRows(1);
    // old answers, by heading
    const c = memberCols_(old), oh = old.getRange(1, 1, 1, c.width).getValues()[0].map(x => plainHeading_(x));
    const consent = oh.indexOf(CONSENT_TITLE), late = oh.indexOf(LATE_HEAD);
    const people = scholarByEmail_(ss), leads = {};
    schoolRows_(ss).forEach(s => { if (s.email) leads[s.email] = s.name; });
    const src = old.getLastRow() > 1 ? old.getRange(2, 1, old.getLastRow() - 1, c.width).getValues() : [];
    const rows = src.filter(v => v[0] instanceof Date || String(v[0]).trim()).map(v => {
      const email = c.email >= 0 ? emailOf_(v[c.email]) : '', p = people[email] || {};
      const g = k => (c[k] >= 0 ? v[c[k]] : '');
      return [v[0], c.email >= 0 ? v[c.email] : '', p.name || '', p.school || '', leads[email] ? 'Campus Lead' : (p.name ? 'Member' : ''),
              g('vibe'), g('events'), g('learning'), g('help'), g('tea'), consent >= 0 ? v[consent] : '', late >= 0 ? v[late] : ''];
    });
    if (rows.length) {
      if (nw.getMaxRows() < rows.length + 1) nw.insertRowsAfter(nw.getMaxRows(), rows.length + 1 - nw.getMaxRows());
      nw.getRange(2, 1, rows.length, head.length).setValues(rows);
    }
    nw.getRange(1, 1, nw.getMaxRows(), 1).setNumberFormat('yyyy-mm-dd hh:mm');
    nw.getRange(1, head.length, nw.getMaxRows(), 1).setNumberFormat('yyyy-mm-dd hh:mm');
    SpreadsheetApp.flush();
    if (nw.getLastRow() - 1 !== rows.length) throw new Error('copy check failed: the old tab is kept');
    // private backup of the old tab, then it is removed from the tracker
    const stamp = Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd');
    const bk = SpreadsheetApp.create('START – old Member check-ins tab (backup ' + stamp + ')');
    old.copyTo(bk).setName(OLD_MEMBER_TAB);
    bk.getSheets().filter(x => x.getName() !== OLD_MEMBER_TAB).forEach(x => bk.deleteSheet(x));
    try {
      const file = DriveApp.getFileById(bk.getId()), parents = DriveApp.getFileById(ss.getId()).getParents();
      if (parents.hasNext()) file.moveTo(parents.next());
    } catch (e) { Logger.log('Backup stays in My Drive: ' + e.message); }
    ss.deleteSheet(old);
    MEMBER_CACHE_ = null;
    docProps_().setProperty('ANSWERS_BACKUP', bk.getUrl());
    try { const p = docProps_(), d = JSON.parse(p.getProperty('TABLE_LOOK') || '{}'); delete d[TAB.mcheckins]; p.setProperty('TABLE_LOOK', JSON.stringify(d)); } catch (e) { /* styled next time */ }
    try { ss.toast(rows.length + ' answers moved to the new "' + TAB.mcheckins + '" tab. The old tab is backed up in Drive.', 'START Tracker', 10); } catch (e) { /* no screen */ }
    return { moved: rows.length, backup: bk.getUrl() };
  } finally { lock.releaseLock(); }
}


/* ---------- headings that must not change (the tracker finds every column by its heading) ----------
 * Warning-only: Google asks "are you sure?" before anyone edits them. The website and the code can still write. */
function guardTabs_(ss) {
  const tagged = sh => sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).concat(sh.getProtections(SpreadsheetApp.ProtectionType.RANGE))
    .some(p => p.getDescription() === PROTECT_TAG + ' heading' || p.getDescription() === PROTECT_TAG + ' answers');
  const ans = ss.getSheetByName(TAB.mcheckins);
  if (ans && !tagged(ans)) ans.protect().setDescription(PROTECT_TAG + ' answers').setWarningOnly(true);   // written by the website only
  [[TAB.scholars, 11], [TAB.schools, 17], [TAB.checkins, 0], [ATTEND_TAB, 3], [TAB.events, 0]].forEach(([n, w]) => {
    const sh = ss.getSheetByName(n);
    if (!sh || tagged(sh)) return;
    sh.getRange(1, 1, 1, Math.max(w || sh.getLastColumn(), 1)).protect().setDescription(PROTECT_TAG + ' heading').setWarningOnly(true);
  });
}

/* ---------- data retention (Data Privacy Act): answers are not kept forever ----------
 * Once a month: check-in answers older than RETAIN_YEARS from people who are no longer Active scholars are deleted,
 * and "Not recognised" answers older than 90 days. Active scholars keep their history (streaks, XP). */
var RETAIN_YEARS = 2;
function cleanOldAnswers_(ss) {
  const sh = memberTab_(ss);
  let removed = 0;
  if (sh && sh.getLastRow() > 1) {
    const c = memberCols_(sh), active = {};
    Object.entries(scholarByEmail_(ss)).forEach(([e, p]) => { if (p.status === 'Active') active[e] = true; });
    const cut = new Date(); cut.setFullYear(cut.getFullYear() - RETAIN_YEARS);
    const lock = docLock_();
    if (!lock.tryLock(30000)) return 0;
    try {
      const n = sh.getLastRow() - 1, vals = sh.getRange(2, 1, n, c.width).getValues();
      const keep = vals.filter(v => !(v[0] instanceof Date && v[0] < cut && !active[emailOf_(c.email >= 0 ? v[c.email] : '')]));
      removed = vals.length - keep.length;
      if (removed) {
        sh.getRange(2, 1, n, c.width).clearContent();
        if (keep.length) sh.getRange(2, 1, keep.length, c.width).setValues(keep);
        MEMBER_CACHE_ = null;
      }
    } finally { lock.releaseLock(); }
  }
  const un = ss.getSheetByName(typeof UNKNOWN_TAB === 'string' ? UNKNOWN_TAB : '');
  if (un && un.getLastRow() > 1) {
    const old = new Date(Date.now() - 90 * 86400000);
    for (let r = un.getLastRow(); r >= 2; r--) { const t = un.getRange(r, 1).getValue(); if (t instanceof Date && t < old) { un.deleteRow(r); removed++; } }
  }
  return removed;
}