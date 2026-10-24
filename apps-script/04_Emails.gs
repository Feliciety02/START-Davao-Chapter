/* =====================================================================
 *  START SCHOLAR TRACKER  ·  04_Emails
 *  ---------------------------------------------------------------------
 *  Inside: Every email: weekly summary, reminders, follow-ups, daily limit, privacy checks.
 *  To customise: Email wording.
 *
 *  The code has 9 files + Email_Reminder and Web_App (HTML). Keep these exact names:
 *    00_Config, 01_Region, 02_Menu_Setup, 03_Forms, 04_Emails, 05_Portals, 06_Calendar, 07_Look, 08_Tools
 *  To update: open each file, select all (Ctrl+A), paste the new version, Save.
 *  Then in the Google Sheet: START Tracker › ① Set up / fix the tracker.
 *  You never need to run anything from here: use the START Tracker menu in the Sheet.
 * ===================================================================== */

// The ▶ Run button runs the FIRST function of the open file: this one, so it is always safe.
function about_04_Emails() {
  Logger.log('START Tracker' + (typeof START_MODULES === 'string' ? ' ' + START_MODULES : '') + ', file 04_Emails. Nothing to run here: open the Google Sheet and use the START Tracker menu.');
}


// ━━━━━━━━━━━━━━━━━━━━  WeeklyEmail  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  📧 WEEKLY SUMMARY EMAIL
 * ================================================================ */

function weeklySummary(e) {
  if (!RR_EMAILS) return;                                                     // on the RR's START page instead
  if (e && e.triggerUid && !(ownerGuard_(e) && weeklyStillOn_())) return;   // automatic run: email account only
  const sumWeek = ymd_(mondayOf_(new Date()));
  if (e && e.triggerUid && alreadySent_('summary', sumWeek, 'me')) return;     // ✅ this week's summary already went out
  const ss = SpreadsheetApp.getActive();
  const st = getSettings_();
  try { updateDropdown_(); } catch (e) { /* the email still matters */ }

  const now = new Date();
  const since = new Date(now.getTime() - 7 * 86400000);
  const fmt = d => Utilities.formatDate(d, CONFIG.timezone, 'MMM d');

  const schools = schoolRows_(ss);

  const ck = ss.getSheetByName(TAB.checkins);
  const c = ck ? ckCols_(ck) : null;
  const rows = ck && ck.getLastRow() > 1
    ? ck.getRange(2, 1, ck.getLastRow() - 1, c.width).getValues()
        .map(r => [r[0], r[c.school], c.events >= 0 ? r[c.events] : '', c.dc >= 0 ? r[c.dc] : '',
                   c.concern >= 0 ? r[c.concern] : '', c.name >= 0 ? r[c.name] : '', c.help >= 0 ? r[c.help] : ''])
    : [];
  const lastBySchool = {}, recentBySchool = {};
  rows.forEach(r => {
    if (!(r[0] instanceof Date)) return;
    const s = String(r[1]).trim();
    if (!lastBySchool[s] || r[0] > lastBySchool[s]) lastBySchool[s] = r[0];
    if (r[0] >= since && (!recentBySchool[s] || r[0] >= recentBySchool[s][0])) recentBySchool[s] = r;
  });
  const recent = rows.filter(r => r[0] instanceof Date && r[0] >= since);

  const withFocal = schools.filter(s => s.focal);
  const replied = schools.filter(s => recentBySchool[s.name]);
  const noReply = withFocal.filter(s => !recentBySchool[s.name]);
  const noFocal = schools.filter(s => !s.focal);
  const termStart = termStart_(ss);
  const reengaged = schools.filter(s => s.inactive && lastBySchool[s.name] &&
    (!(termStart instanceof Date) || lastBySchool[s.name] >= termStart)).length;
  const repliedFocal = withFocal.filter(s => recentBySchool[s.name]).length;
  const rate = withFocal.length ? Math.round(100 * repliedFocal / withFocal.length) : 0;
  const evAns = recent.filter(r => String(r[2]).trim() && isNaN(Number(r[2])));                 // "How did the events go?"
  const evGood = evAns.filter(r => typeof EVENT_GO_CHOICES !== 'undefined' && EVENT_GO_CHOICES.slice(0, 2).indexOf(String(r[2])) !== -1);
  const concerns = recent.filter(r => String(r[4]).trim() || String(r[6]).indexOf('Yes') === 0);
  const urgent = recent.filter(r => r[6] === HELP_CHOICES[2]);
  const daysQuiet = s => lastBySchool[s.name]
    ? Math.floor((now - lastBySchool[s.name]) / 86400000) + ' days ago' : 'never';

  const link = checkinUrl_();
  const dueText = (d, met) => {
    if (met) return 'Reached ✅';
    if (!(d instanceof Date)) return '';
    const left = Math.ceil((d - now) / 86400000);
    return fmt(d) + (left < 0 ? ' (passed)' : ` (${left} days left)`);
  };

  let privacy = '';
  try {
    const probs = privacyProblems_().filter(p => p.level === 'bad');
    if (probs.length) privacy = emailAlert_('<b>🔒 Privacy: please fix now.</b><br>' + probs.map(p => esc_(p.text + '. ' + p.fix)).join('<br>'), 'red');
  } catch (e) { /* sharing could not be checked */ }

  let health = '';                                              // 🩺 the weekly self-check rides along here (24_Extras)
  try {
    const bad = typeof healthRows_ === 'function' ? healthRows_().filter(r => r[0] === '❌') : [];
    if (bad.length) health = emailAlert_('<b>🩺 The tracker needs a fix:</b><br>' + bad.map(r => esc_(r[1] + (r[2] ? ' → ' + r[2] : ''))).join('<br>'), 'red');
  } catch (e) { /* check could not run */ }

  const extra = privacy + health +
    (urgent.length ? emailAlert_(`<b>⚠️ ${urgent.length} urgent request(s) this week:</b> ${urgent.map(r => esc_(r[1])).join(', ')}`, 'red') : '') +
    emailSection_('Targets') +
    emailTable_(['Target', 'Now', 'Goal', 'Deadline'], [
      ['START Campus Leads named', withFocal.length, st.tFocal, dueText(st.dFocal, withFocal.length >= st.tFocal)],
      ['Schools re-engaged', reengaged, st.tReeng, dueText(st.dReeng, reengaged >= st.tReeng)],
      ['Response rate (last 7 days)', `${rate}% (${repliedFocal}/${withFocal.length})`, st.tRate + '%', ''],
      ['Events that went well (Campus Leads\' view)', evAns.length ? `${evGood.length} of ${evAns.length} schools` : 'No event this week', '', ''],
    ]) +
    emailSection_(`Follow up: no reply this week (${noReply.length})`) +
    emailTable_(['School', 'START Campus Lead', 'Contact', 'Last check-in'], noReply.map(s => [s.name, s.focal, s.contact, daysQuiet(s)])) +
    emailSection_(`Replied this week (${replied.length})`) +
    emailTable_(['School', 'Answered by', 'Events', 'DataCamp'], replied.map(s => {
      const r = recentBySchool[s.name]; return [s.name, r[5], r[2], r[3]];
    })) +
    emailSection_(`Concerns and help requests (${concerns.length})`) +
    emailTable_(['School', 'Date', 'Needs help?', 'Concern'], concerns.map(r => [r[1], fmt(r[0]), r[6], r[4]])) +
    emailSection_(`No START Campus Lead yet (${noFocal.length})`) +
    emailTable_(['School', 'Last check-in'], noFocal.map(s => [s.name, daysQuiet(s)])) +
    (() => { try { return memberSummaryHtml_(ss); } catch (e) { return ''; } })() +   // 👥 member check-ins
    emailSection_('Check-in form for START Campus Leads') +
    `<p style="margin:0;font-size:13px;"><a href="${link}" style="color:#0099FF;word-break:break-all;">${esc_(link)}</a></p>`;

  const html = designedEmail_({
    badge: 'Weekly summary', title: ['Weekly', 'Summary'], greeting: '',
    intro: [],
    extraHtml: extra,
    link: (DB_MODE && webHome_()) || ss.getUrl(), buttonText: DB_MODE && webHome_() ? 'Open the website' : 'Open the tracker',
    buttonNote: '',
    chips: [['REGION', st.region || '—'], ['WEEK', `${fmt(since)} – ${fmt(now)}`], ['REPLIED', `${repliedFocal} / ${withFocal.length}`]],
    outro: [`Sent every ${st.day}. To stop: START Tracker › ⏰ Automatic emails › ⏸ Turn OFF.`],
    signature: [],
  });

  const to = String(st.recipients || '').trim() || Session.getEffectiveUser().getEmail();
  sendMail_({
    to,
    subject: `START · Weekly summary${st.region ? ' (' + st.region + ')' : ''} ${fmt(now)}: ` +
             `${repliedFocal}/${withFocal.length} schools replied`,
    htmlBody: html, name: 'START Tracker', inlineImages: logoInlineImages_(),
  });
  markSent_('summary', sumWeek, 'me');
  try { ss.toast('Summary sent to ' + to, 'START Tracker', 8); } catch (e) { /* trigger run */ }
  if (st.missedAlert === true) {          // Settings › "Alert me: schools quiet 2+ weeks"
    try { sendMissedAlert_(true); } catch (e) { Logger.log('Quiet schools alert not sent: ' + e.message); }
  }
}

// Creates (or moves) the weekly summary trigger. Called only by applySchedule_ (the 🚀 switch).
function scheduleWeekly_() {
  const st = getSettings_();
  removeTriggers_(['weeklySummary']);
  ScriptApp.newTrigger('weeklySummary').timeBased()
    .onWeekDay(ScriptApp.WeekDay[st.day.toUpperCase()]).atHour(st.hour24)
    .inTimezone(CONFIG.timezone).create();
  return st;
}


// ━━━━━━━━━━━━━━━━━━━━  Reminders  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  📧 AUTOMATIC REMINDERS TO START CAMPUS LEADS
 *  - Reminder day: every school whose START Campus Lead has an email (in Scholars) gets the formal
 *    reminder with its own form link (school already selected).
 *  - Follow-up day: only schools that have NOT answered since the start of that week.
 *  - Every email (and every school skipped) is listed in the "Reminder log" tab.
 *  Emails are sent from the Google account that turned reminders on.
 * ================================================================ */

/* ---------- Triggers (runs by itself) ---------- */

function autoReminder(e) {
  if (!ownerGuard_(e) || !weeklyStillOn_()) return;
  if (remindTwoDays_() && daysToDeadline_(reminderWeek_(new Date()).deadline) !== 2) return;   // only 2 days before the deadline
  sendReminders_('reminder');
}
function autoFollowUp(e) {
  if (remindTwoDays_()) { removeTriggers_(['autoFollowUp']); return; }                        // website check-in: one reminder only
  if (ownerGuard_(e) && weeklyStillOn_()) sendReminders_('followup');
}

// The 🚀 switch can be unticked from any account: the schedule checks it before every send and stops itself.
function weeklyStillOn_() {
  if (getSettings_().autoWeekly === true) return true;
  removeTriggers_(['weeklySummary', 'autoReminder', 'autoFollowUp']);
  return false;
}

// ⏰ Automatic emails › 🚀 one click: weekly summary + reminders + follow-ups ON (ticks the 🚀 switch in Settings)
function turnOnWeekly() { setWeekly_(true); }
function turnOffWeekly() { setWeekly_(false); }

function setWeekly_(on) {
  requireSetup_();
  if (on && !ensureOwner_('turn on the weekly emails')) return;
  const ss = SpreadsheetApp.getActive();
  ss.getRange(settingCell_('autoWeekly')).setValue(on);
  applySchedule_();
  docProps_().setProperty('SCHEDULE_SYNC', scheduleKey_());
  if (!on) return notify_('⏸ Automatic weekly emails are OFF.\n\nNothing more will be sent until you turn them on again ' +
                          '(⏰ Automatic emails › 🚀 Turn ON, or tick 🚀 in Settings).');
  const st = getSettings_(), t = reminderTargets_();
  let msg = '🚀 Automatic weekly emails are ON.\n\n' +
    `• Weekly summary to you: every ${st.day} at about ${HOURS[st.hour24]}.\n` +
    `• Reminder to every START Campus Lead: every ${st.remDay} at about ${HOURS[st.remHour24]} (deadline: ${st.dueDay}).\n` +
    (st.fuDay !== 'No follow-up'
      ? `• Follow-up to schools that have not answered: every ${st.fuDay} at about ${HOURS[st.remHour24]}.\n`
      : '• Follow-ups: off (see Settings).\n') +
    '• After each send, you get a copy of the email and a report of who received it.\n\n' +
    `START Campus Leads who will receive it: ${t.ready.length}.`;
  if (t.skipped.length) msg += `\nNo email in Scholars › Contact yet (${t.skipped.length}): ${t.skipped.map(s => s.name).join(', ')}.`;
  msg += '\n\nChange days and times in Settings: they apply by themselves.';
  notify_(msg);
}

// Creates (or moves) the reminder and follow-up triggers. Called only by applySchedule_ (the 🚀 switch).
function scheduleReminders_() {
  const st = getSettings_();
  removeTriggers_(['autoReminder', 'autoFollowUp']);
  ScriptApp.newTrigger('autoReminder').timeBased()
    .onWeekDay(ScriptApp.WeekDay[String(st.remDay).toUpperCase()]).atHour(st.remHour24)
    .inTimezone(CONFIG.timezone).create();
  if (st.fuDay !== 'No follow-up') {
    ScriptApp.newTrigger('autoFollowUp').timeBased()
      .onWeekDay(ScriptApp.WeekDay[String(st.fuDay).toUpperCase()]).atHour(st.remHour24)
      .inTimezone(CONFIG.timezone).create();
  }
  return st;
}

function removeTriggers_(names) {
  let n = 0;
  ScriptApp.getProjectTriggers().filter(t => names.indexOf(t.getHandlerFunction()) !== -1)
    .forEach(t => { ScriptApp.deleteTrigger(t); n++; });
  return n;
}

/* ---------- Manual sends from the menu ---------- */

function confirmAndRun_(type) {
  requireSetup_();
  const ui = ui_();
  const t = reminderTargets_();
  const answered = type === 'followup' || remindTwoDays_()
    ? answeredSince_(SpreadsheetApp.getActive(), (type === 'followup' ? followUpWeek_(new Date()) : reminderWeek_(new Date())).monday) : {};
  const list = t.ready.filter(s => !answered[s.name]);
  const mem = typeof memberWeeklyTargets_ === 'function' ? memberWeeklyTargets_(type) : [];   // 👥 welcomed members get it too
  if (!list.length && !mem.length) {
    notify_(type === 'followup'
      ? 'There is no one to follow up: everyone with an email address has already answered this week.'
      : 'No email addresses were found.\n\nPlease type each START Campus Lead\'s email in the Scholars tab, Contact (email) column.');
    return;
  }
  const label = type === 'followup' ? 'follow-up' : 'check-in reminder';
  const total = list.length + mem.length;
  const short = (arr, max) => arr.slice(0, max).join('\n') + (arr.length > max ? `\n  … and ${arr.length - max} more` : '');
  const ok = ui.alert('Send the ' + label + ' now?',
    `${total} people:\n\n` +
    (list.length ? `Campus Leads (${list.length}):\n` + short(list.map(s => `  – ${s.name} (${s.email})`), 15) + '\n\n' : '') +
    (mem.length ? `Members (${mem.length}):\n` +
                  short(mem.map(p => `  – ${p.name || p.email} · ${p.school}`), 15) + '\n\n' : '') +
    (t.skipped.length ? `No email: ${t.skipped.map(s => s.name).join(', ')}\n\n` : '') +
    (total > mailRoom_() ? `⏳ ${mailRoom_()} go today, the rest tomorrow morning (daily limit).` : ''), ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  const r = sendReminders_(type);
  const m = r.members || { sent: 0, waiting: 0 };
  notify_(`✅ Sent to ${r.sent} Lead${r.sent === 1 ? '' : 's'} and ${m.sent} member${m.sent === 1 ? '' : 's'}` +
    (r.done || r.waiting + m.waiting || r.failed ? '\n' : '') +
    (r.done ? `\n• ${r.done} already had it this week` : '') +
    (r.waiting + m.waiting ? `\n• ⏳ ${r.waiting + m.waiting} go out tomorrow morning (daily limit)` : '') +
    (r.failed ? `\n• ${r.failed} failed (see Reminder log)` : ''));
}

function sendTestReminder() {
  requireSetup_();
  const me = Session.getEffectiveUser().getEmail();
  const first = schoolRows_(SpreadsheetApp.getActive()).sort((a, b) => a.name.localeCompare(b.name))[0];
  const school = {
    name: first ? first.name : 'Sample School',
    focal: (first && first.focal) || 'Sample START Campus Lead',
    link: (first && first.link) || checkinUrl_(),
  };
  const logos = logoInlineImages_();
  const w = reminderWeek_(new Date());
  const m = composeEmail_('reminder', school, w);
  sendMail_({ to: me, subject: '[TEST] ' + m.subject, htmlBody: m.html, body: m.text, name: m.from, inlineImages: logos });
  const f = composeEmail_('followup', school, followUpWeek_(new Date()));
  sendMail_({ to: me, subject: '[TEST] ' + f.subject, htmlBody: f.html, body: f.text, name: f.from, inlineImages: logos });
  notify_(`📧 2 test emails sent to ${me}\n\n• Weekly reminder\n• Follow-up\n\nNobody else got them.`);
}

/* ---------- Core ---------- */

// onlyNames: send only to these schools (used by the waiting list). week: keep the original week.
function sendReminders_(type, onlyNames, week) {
  const ss = SpreadsheetApp.getActive();
  const t = reminderTargets_();
  const w = week || (type === 'followup' ? followUpWeek_(new Date()) : reminderWeek_(new Date()));
  let sent = 0, failed = 0;
  const log = [];
  const stamp = new Date();
  const label = type === 'followup' ? 'Follow-up' : 'Reminder';
  // website check-in: the reminder too goes ONLY to Leads who have not answered; never after the deadline
  if (remindTwoDays_() && daysToDeadline_(w.deadline) < 0) {
    writeLog_(ss, [[new Date(), label + ' · due ' + Utilities.formatDate(w.deadline, CONFIG.timezone, 'MMM d'), '', '', 'Not sent: the deadline has passed']]);
    return { sent: 0, failed: 0, waiting: 0, done: 0, members: { sent: 0, waiting: 0 } };
  }
  const answered = type === 'followup' || remindTwoDays_() ? answeredSince_(ss, w.monday) : {};   // Checkins read once, not per school
  const logos = logoInlineImages_();
  const store = portalStore_();
  const list = t.ready.filter(s => !answered[s.name] && (!onlyNames || onlyNames.indexOf(s.name) !== -1));
  list.forEach(s => { s.portal = portalUrl_(store, s); });

  let room = mailRoom_(), sample = null, done = 0;
  const waiting = [], period = ymd_(w.deadline);
  const logLabel = label + ' · due ' + Utilities.formatDate(w.deadline, CONFIG.timezone, 'MMM d');   // the log shows the week
  list.forEach(s => {
    if (alreadySent_(type, period, s.email)) { done++; return; }      // ✅ already got this week's email: never twice
    if (room < 1) { waiting.push(s); return; }
    try {
      const m = composeEmail_(type, s, w);
      sendMail_({ to: s.email, subject: m.subject, htmlBody: m.html, body: m.text, name: m.from, inlineImages: logos });
      markSent_(type, period, s.email);
      log.push([stamp, logLabel, s.name, s.email, 'Sent']); sent++; room--;
      sample = sample || m;
    } catch (e) {
      log.push([stamp, logLabel, s.name, s.email, 'Not sent: ' + e.message]); failed++;
    }
  });
  if (waiting.length) {
    const kept = queueEmails_({ kind: type, schools: waiting.map(s => s.name),
                                week: { monday: w.monday.getTime(), deadline: w.deadline.getTime() } });
    waiting.forEach(s => log.push([stamp, logLabel, s.name, s.email, kept ? WAIT_NOTE : 'Not sent: daily email limit reached']));
    if (!kept) failed += waiting.length;
  }
  if (!onlyNames) t.skipped.forEach(s => log.push([stamp, label, s.name, s.contact || '(blank)', 'Skipped: no email in Scholars']));
  writeLog_(ss, log);
  try { reportToMe_(label, w, log, sample, logos); } catch (e) { Logger.log('Report not sent: ' + e.message); }
  let members = { sent: 0, waiting: 0 };
  if (!onlyNames && typeof sendMemberWeekly_ === 'function') {     // 👥 welcomed members get theirs right after (24_Extras)
    try { members = sendMemberWeekly_(type, w) || members; } catch (e) { Logger.log('Member weekly emails: ' + e.message); }
  }
  return { sent, failed, waiting: waiting.length, done, members };
}

// After every reminder / follow-up run: who got it + a copy of the message, sent to you (the tracker's owner)
function reportToMe_(label, w, log, sample, logos) {
  const me = Session.getEffectiveUser().getEmail();
  const week = Utilities.formatDate(w.monday, CONFIG.timezone, 'MMM d, yyyy');
  const sent = log.filter(r => r[4] === 'Sent').length, waitN = log.filter(r => r[4] === WAIT_NOTE).length, skipped = log.length - sent - waitN;
  const extra = (log.length ? emailSection_('Who received it') + emailTable_(['School', 'Sent to', 'Status'], log.map(r => [r[2], r[3], r[4]]))
                            : emailAlert_('Nothing to send: every school with an email has already answered this week. 🎉')) +
    (sample ? emailSection_('The message they received') +
      `<div style="border:1px solid #CFE0F3;border-radius:10px;padding:12px 14px;background:#ffffff;font-size:13px;line-height:1.6;color:#33415c;">` +
      `<b>${esc_(sample.subject)}</b><br><br>${esc_(sample.text).replace(/\n/g, '<br>')}</div>` : '');
  const html = designedEmail_({
    badge: 'Copy for you', title: ['Sent', 'Report'], greeting: 'Hi!',
    intro: [`Your ${label.toLowerCase()}s went out. Who got them and the message are below.`]
      .concat(waitN ? [`⏳ ${waitN} wait for tomorrow morning (Google's daily email limit). They go out by themselves.`] : []),
    extraHtml: extra,
    link: SpreadsheetApp.getActive().getUrl(), buttonText: 'Open the tracker', buttonNote: '',
    chips: [['SENT', sent], ['WEEK OF', week], ['SKIPPED / NOT SENT', skipped]],
    signature: [],
  });
  if (!RR_EMAILS) return;                                                     // the Reminder log and the RR's START page show it
  sendMail_({ to: me, subject: `START · Copy: ${label}s sent (${sent}), week of ${week}`,
                      htmlBody: html, name: 'START Tracker', inlineImages: logos || logoInlineImages_() });
}

// Schools with a usable email in the Contact column (column E)
function reminderTargets_() {
  const general = checkinUrl_();
  const ready = [], skipped = [];
  schoolRows_(SpreadsheetApp.getActive()).forEach(s => {
    s.link = s.link || general;
    (s.email ? ready : skipped).push(s);
  });
  return { ready, skipped };
}

// { "School name": true } for every school that checked in on/after the given date
function answeredSince_(ss, since) {
  const out = {};
  const ck = ss.getSheetByName(TAB.checkins);
  const c = ck && ck.getLastRow() > 1 ? ckCols_(ck) : null;
  (c ? ck.getRange(2, 1, ck.getLastRow() - 1, c.width).getValues() : [])
    .forEach(r => { if (r[0] instanceof Date && r[0] >= since) out[String(r[c.school]).trim()] = true; });
  if (remindTwoDays_()) {                    // website check-in: a Lead's own answer in Check-in answers counts too
    try {
      const lead = {}; schoolRows_(ss).forEach(x => { if (x.email) lead[x.email] = x.name; });
      const sh = memberTab_(ss);
      if (sh && sh.getLastRow() > 1) { const m = memberCols_(sh);
        if (m.email >= 0) sh.getRange(2, 1, sh.getLastRow() - 1, m.width).getValues().forEach(v => {
          const e = emailOf_(v[m.email]); if (v[0] instanceof Date && v[0] >= since && lead[e]) out[lead[e]] = true; }); }
    } catch (e) { Logger.log('Lead answers: ' + e.message); }
  }
  return out;
}

/* ---------- Dates ---------- */

function startOfDay_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function dayIndex_(name) { return (DAYS.indexOf(name) + 1) % 7; }          // Monday -> 1 ... Sunday -> 0
function mondayOf_(d) { const m = startOfDay_(d); m.setDate(m.getDate() - ((m.getDay() + 6) % 7)); return m; }

// Reminder: the next deadline on/after today, and the Monday of that deadline's week
function reminderWeek_(now) {
  const due = startOfDay_(now), want = dayIndex_(getSettings_().dueDay);
  while (due.getDay() !== want) due.setDate(due.getDate() + 1);
  return { monday: mondayOf_(due), deadline: due };
}

// Follow-up: the week of the most recent reminder (works whether the deadline is before the follow-up,
// e.g. Wed deadline → Thu follow-up, or after it, e.g. Sat follow-up → Sun deadline)
function followUpWeek_(now) {
  const rem = startOfDay_(now), want = dayIndex_(getSettings_().remDay);
  while (rem.getDay() !== want) rem.setDate(rem.getDate() - 1);
  return reminderWeek_(rem);
}

/* ---------- Email text (formal) ---------- */

function composeEmail_(type, s, w) {
  const st = getSettings_();
  const f = d => Utilities.formatDate(d, CONFIG.timezone, 'MMMM d, yyyy');
  const fd = d => Utilities.formatDate(d, CONFIG.timezone, 'EEEE, MMMM d, yyyy');
  const greet = s.focal ? `Dear ${displayName_(s.focal)},` : `Dear START Campus Lead of ${s.name},`;
  const sign = rrSignature_(st);
  const from = st.rrName ? `${st.rrName} (START ${st.region || ''})`.replace(/\s+\)/, ')') : 'START Tracker';

  let subject, paras;
  if (type === 'followup') {
    subject = `START · Reminder: weekly check-in, ${s.name}, week of ${f(w.monday)}`;
    paras = [
      `We have not received ${s.name}'s check-in for this week yet. ` +
        (w.deadline >= startOfDay_(new Date()) ? `Kindly answer it by ${fd(w.deadline)}. It takes about 2 minutes.`
                                               : `It was due ${fd(w.deadline)}. Kindly answer it as soon as you can.`),
      'Already sent it? Please disregard this email.',
    ];
  } else {
    subject = `START · Weekly check-in, ${s.name}, week of ${f(w.monday)}`;
    paras = [
      `Kindly answer ${s.name}'s START check-in for this week. It takes about 2 minutes.`,
      'Any concerns? Write them in the form or reply to this email.',
    ];
  }

  const html = renderEmail_({
    badge: type === 'followup' ? 'Follow-up' : 'Reminder',
    badgeColor: type === 'followup' ? '#C2410C' : THEME.blue,
    accent: THEME.navy,
    greeting: greet,
    intro: paras.slice(0, 1),
    outro: paras.slice(1),
    link: oneFormOn_() ? checkinLink_() : s.link,
    portal: s.portal || '',
    extraHtml: (() => { try { return eventsForEmail_(s.name); } catch (e) { return ''; } })(),   // 📅 upcoming events
    buttonText: 'Answer the check-in',
    school: s.name,
    week: Utilities.formatDate(w.monday, CONFIG.timezone, 'MMM d, yyyy'),
    deadlineLabel: type === 'followup' && w.deadline < startOfDay_(new Date()) ? 'WAS DUE' : 'DUE ON',
    deadline: Utilities.formatDate(w.deadline, CONFIG.timezone, 'EEE, MMM d'),
    signature: sign,
    region: st.region ? st.region + (st.rrName ? ' · ' + st.rrName : '') : (st.rrName || ''),
    footer: 'This message was sent through the START Tracker on behalf of your Regional Representative. ' +
            'You may reply to this email directly.',
  });
  const text = [greet, ''].concat(paras.slice(0, 1).join('\n\n'), '', 'Form link: ' + (oneFormOn_() ? checkinLink_() : s.link), '',
    s.portal ? 'Your ' + PORTAL_NAME + ': ' + s.portal + '\n' : '',
    paras.slice(1).join('\n\n'), '', 'Respectfully,', sign.join('\n')).join('\n');
  return { subject, html, text, from };
}

/* ---------- one design for every email ---------- */

// Builds any tracker email in the Email_Reminder design.
// o: badge, badgeColor, greeting, intro[], extraHtml, outro[], link, buttonText, buttonNote,
//    chips [[label, value] x3], signature[] (default: the RR), footer
function designedEmail_(o) {
  const st = getSettings_();
  const sign = o.signature || rrSignature_(st);
  const c = o.chips || [['', ''], ['', ''], ['', '']];
  return renderEmail_({
    badge: o.badge, badgeColor: o.badgeColor || THEME.blue,
    title1: (o.title || ['Weekly', 'Check-in'])[0], title2: (o.title || ['Weekly', 'Check-in'])[1],
    greeting: o.greeting || '', intro: o.intro || [], extraHtml: o.extraHtml || '', outro: o.outro || [],
    link: o.link, buttonText: o.buttonText, buttonNote: o.buttonNote || '',
    schoolLabel: c[0][0], school: String(c[0][1]), weekLabel: c[1][0], week: String(c[1][1]),
    deadlineLabel: c[2][0], deadline: String(c[2][1]),
    signature: sign, region: st.region ? st.region + (st.rrName ? ' · ' + st.rrName : '') : (st.rrName || ''),
    footer: o.footer || 'This message was sent automatically by the START Tracker.',
  });
}

// The one sign-off on every email:  Respectfully, / Fe Anne Malasarte / START R11 Regional Representative
function rrSignature_(st) {
  st = st || {};
  return [String(st.rrName || '').trim(), ['START', regionShort_(st.region), 'Regional Representative'].filter(Boolean).join(' ')].filter(Boolean);
}
// "Region XI" -> "R11", "Region IV-A" -> "R4A"; NCR, CAR, BARMM… stay as they are
function regionShort_(region) {
  const r = String(region || '').trim();
  const m = r.match(/^Region\s+([IVXL]+)\s*-?\s*([AB])?$/i);
  if (!m) return r;
  const v = { I: 1, V: 5, X: 10, L: 50 };
  let n = 0; const s = m[1].toUpperCase();
  for (let i = 0; i < s.length; i++) n += v[s[i]] < (v[s[i + 1]] || 0) ? -v[s[i]] : v[s[i]];
  return 'R' + n + (m[2] ? m[2].toUpperCase() : '');
}

// Deck-style pieces for the extra content (tables, section titles, alert box)
function emailSection_(title) {
  return `<div style="margin:20px 0 8px 0;padding-left:10px;border-left:4px solid #FFDD00;font-size:12px;font-weight:800;` +
         `letter-spacing:1px;text-transform:uppercase;color:#12203A;">${esc_(title)}</div>`;
}
function emailTable_(head, rows) {
  if (!rows.length) return '<p style="margin:0;color:#5b6b84;font-size:13px;">None.</p>';
  const th = 'padding:8px 10px;text-align:left;background:#12203A;color:#ffffff;font-size:11px;letter-spacing:.5px;text-transform:uppercase;';
  const td = (i, j) => `padding:8px 10px;border-bottom:1px solid #CFE0F3;font-size:13px;${i % 2 ? 'background:#F6F9FD;' : ''}` +
                       (j === 0 ? 'font-weight:700;' : 'color:#33415c;');
  return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border-radius:10px;overflow:hidden;">` +
    `<tr>${head.map(h => `<th style="${th}">${esc_(h)}</th>`).join('')}</tr>` +
    rows.map((r, i) => `<tr>${r.map((c, j) => `<td style="${td(i, j)}">${esc_(c)}</td>`).join('')}</tr>`).join('') + '</table>';
}
function emailAlert_(html, colour) {
  return `<div style="background:${colour === 'red' ? '#FADCD9' : '#FFF6BF'};border-left:5px solid ${colour === 'red' ? '#B3261E' : '#E0B800'};` +
         `border-radius:10px;padding:12px 14px;margin:8px 0;font-size:14px;color:#12203A;">${html}</div>`;
}

// Fills in the Email_Reminder HTML file (edit that file to change the design).
// Falls back to a simple layout if the file is missing.
function renderEmail_(data) {
  let logo = '';
  try { if (typeof emailLogoSrc_ === 'function') logo = emailLogoSrc_(); } catch (e) { /* plain email */ }
  const logoSize = typeof LOGO_STRIP_SIZE === 'undefined' ? { w: 0, h: 0 } : LOGO_STRIP_SIZE;
  data = Object.assign({ title1: 'Weekly', title2: 'Check-in', schoolLabel: 'SCHOOL', weekLabel: 'WEEK OF', buttonNote: '', portal: '',
                        extraHtml: '', greeting: '', intro: [], outro: [], signature: [], badgeColor: THEME.blue,
                        accent: THEME.navy, logo: logo, logoW: logoSize.w, logoH: logoSize.h,
                        bg: '', hero: '', bar: '', sparkle: '', star: '',
                        footer: 'This message was sent by the START Tracker.' }, data);   // defaults
  try { if (typeof emailPics_ === 'function') Object.assign(data, emailPics_()); } catch (e) { /* plain colours */ }
  try {
    const t = HtmlService.createTemplateFromFile('Email_Reminder');
    Object.keys(data).forEach(k => { t[k] = data[k]; });
    const out = t.evaluate().getContent();
    if (data.buttonText && data.link && out.indexOf(data.link) === -1) throw new Error('Email_Reminder template is empty or missing the link');
    return out;
  } catch (e) {
    Logger.log('Email_Reminder template not used: ' + e.message);
    return `<div style="font-family:Poppins,'Segoe UI',Arial,sans-serif;font-size:14px;line-height:1.5;max-width:600px">
      <p>${esc_(data.greeting)}</p>${data.intro.map(p => `<p>${esc_(p)}</p>`).join('')}${data.extraHtml}
      ${data.buttonText && data.link ? `<p><a href="${data.link}" style="background:${data.accent};color:#fff;padding:10px 20px;border-radius:4px;
         text-decoration:none;font-weight:bold;display:inline-block">${esc_(data.buttonText)}</a></p>` : ''}
      ${data.portal ? `<p><a href="${data.portal}">Open my ${PORTAL_NAME}</a> (private to your Google account)</p>` : ''}
      ${data.outro.map(p => `<p>${esc_(p)}</p>`).join('')}
      <p>Respectfully,<br>${data.signature.map(esc_).join('<br>')}</p>
      </div>`;
  }
}

/* ---------- Log tab ---------- */

var LOG_ROWS = 1000;

// Adds rows to the Reminder log. Rows are appended, then the list is sorted newest-first
// (sorting never shifts the summary panel or copies header colours).
function writeLog_(ss, rows) {
  if (DB_MODE) return;                                   // database mode: not used (see 00_Config)
  if (!rows.length || TEST_MODE_) return;                         // tests are not written in the Reminder log
  const sh = ss.getSheetByName(TAB.log) || buildLog_(ss);
  const colA = sh.getRange(2, 1, Math.max(sh.getMaxRows() - 1, 1), 1).getValues();
  let last = 1;
  colA.forEach((r, i) => { if (r[0] !== '' && r[0] !== null) last = i + 2; });
  const need = last + rows.length;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  logBodyStyle_(sh, last + 1, rows.length).setValues(rows);
  sh.getRange(2, 1, need - 1, 5).sort({ column: 1, ascending: false });
  fitRows_(sh, 2, need - 1);
}

// Builds / restyles the Reminder log tab (safe to run again; keeps all entries)
function buildLog_(ss) {
  if (DB_MODE) return null;                                   // database mode: not used (see 00_Config)
  const sh = getOrCreate_(ss, TAB.log);
  if (sh.getMaxRows() < LOG_ROWS + 1) sh.insertRowsAfter(sh.getMaxRows(), LOG_ROWS + 1 - sh.getMaxRows());
  if (sh.getMaxColumns() < 8) sh.insertColumnsAfter(sh.getMaxColumns(), 8 - sh.getMaxColumns());
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 38);
  [[6, 24], [7, 230], [8, 190]].forEach(([c, w]) => sh.setColumnWidth(c, w));   // A:E are fitted by markColumns_
  logBodyStyle_(sh, 2, LOG_ROWS);
  markColumns_(sh, [[1, 'auto', 'Date sent'], [2, 'auto', 'Type'], [3, 'auto', 'School'], [4, 'auto', 'Sent to', '', 'link'],
                    [5, 'auto', 'Status', '', true]], LOG_ROWS, false);

  // Chips
  const B = sh.getRange(2, 2, LOG_ROWS, 1), E = sh.getRange(2, 5, LOG_ROWS, 1);
  sh.setConditionalFormatRules([
    chipRule_(E, 'Sent', CHIP.green), chipRule_(E, 'Skipped', CHIP.grey, false), chipRule_(E, 'Not sent', CHIP.red, false),
    chipRule_(B, 'Reminder', CHIP.blue), chipRule_(B, 'Follow-up', CHIP.yellow), chipRule_(B, 'Announcement', CHIP.green),
  ]);

  // Summary panel (G:H) - formulas use INDIRECT so sorting the log never breaks them
  const A_ = 'INDIRECT("A2:A' + (LOG_ROWS + 1) + '")', E_ = 'INDIRECT("E2:E' + (LOG_ROWS + 1) + '")';
  const L = CONFIG.maxSchools + 1, S = TAB.schools;
  sh.getRange('G1:H1').merge().setValue('📬  REMINDER SUMMARY');
  header_(sh.getRange('G1:H1'));
  const kpi = [
    ['Last emails sent',
     `=IF(COUNTIF(${E_},"Sent")=0,"Not yet",TEXT(MAXIFS(${A_},${E_},"Sent"),"mmm d · h:mm AM/PM"))`],
    ['Sent in the last 7 days', `=COUNTIFS(${A_},">="&(NOW()-7),${E_},"Sent")`],
    ['Skipped in the last run (no email)', `=IF(COUNT(${A_})=0,0,COUNTIFS(${A_},MAX(${A_}),${E_},"Skipped*"))`],
    ['Not sent in the last 7 days', `=COUNTIFS(${A_},">="&(NOW()-7),${E_},"Not sent*")`],
    ['Total emails sent', `=COUNTIF(${E_},"Sent")`],
  ];
  kpi.forEach((k, i) => {
    sh.getRange(2 + i, 7).setValue(k[0]);
    sh.getRange(2 + i, 8).setFormula(k[1]);
  });
  const panel = sh.getRange(2, 7, kpi.length, 2);
  panel.setBackground(THEME.band).setFontFamily(THEME.font).setFontColor(THEME.navy).setFontSize(10)
       .setBorder(true, true, true, true, null, true, TABLE_LINE, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(2, 8, kpi.length, 1).setFontWeight('bold').setFontSize(12).setHorizontalAlignment('right');
  sh.getRange(2, 7, kpi.length, 1).setBorder(null, true, null, null, null, null, THEME.blue, SpreadsheetApp.BorderStyle.SOLID_THICK);

  sh.getRange('G8:H8').merge().setValue('📭  SCHOOLS WITHOUT AN EMAIL');
  header_(sh.getRange('G8:H8'));
  sh.getRange('G9:H9').setValues([['School', 'START Campus Lead']])
    .setFontWeight('bold').setFontSize(9).setFontColor(THEME.muted).setBackground(THEME.band);
  sh.getRange('G10').setFormula(
    `=IFERROR(FILTER({${S}!A2:A${L},IF(${S}!D2:D${L}="","— none yet",${S}!D2:D${L})},` +
    `${S}!A2:A${L}<>"",NOT(REGEXMATCH(${S}!E2:E${L}&"","@"))),"All schools have an email 🎉")`);
  const list = sh.getRange(10, 7, CONFIG.maxSchools, 2);
  list.setBackground(null).setFontFamily(THEME.font).setFontSize(10).setFontColor(THEME.navy).setFontWeight('normal');
  sh.getRange(10, 7, CONFIG.maxSchools, 1).setFontWeight('bold').setFontColor(THEME.red);
  sh.getRange(10, 7, CONFIG.maxSchools + 1, 2).setBorder(null, true, true, true, null, null, TABLE_LINE, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('G9').setNote('These schools are skipped by automatic reminders.\n' +
    'Fix: type the START Campus Lead\'s email in the Scholars tab, Contact (email) column.');

  sh.setHiddenGridlines(true);
  sh.setTabColor(THEME.tabs.log);
  placeTutorial_(sh);
  placeTabBanner_(sh, 5);
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).filter(p => p.getDescription() === PROTECT_TAG).forEach(p => p.remove());
  sh.protect().setDescription(PROTECT_TAG).setWarningOnly(true);
  return sh;
}

function logBodyStyle_(sh, row, n) {
  const r = bodyRows_(sh.getRange(row, 1, n, 5));
  sh.getRange(row, 1, n, 1).setNumberFormat('mmm d, yyyy · h:mm AM/PM').setFontColor(THEME.muted);
  sh.getRange(row, 2, n, 1).setHorizontalAlignment('center');
  sh.getRange(row, 3, n, 1).setFontWeight('bold').setFontColor(THEME.blue);
  sh.getRange(row, 4, n, 1).setFontColor(THEME.muted).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sh.getRange(row, 5, n, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  return r;
}


// ━━━━━━━━━━━━━━━━━━━━  Safety  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  🛡 SAFETY: every email goes through sendMail_ here
 *   1. Reply-to   replies always reach the address in Settings › "Replies go to"
 *   2. Daily limit Google allows a limited number of emails per day (about 100 for a
 *                 personal Gmail, more for school accounts). Big sends (reminders,
 *                 follow-ups, announcements) only use what is left, keep a few for
 *                 urgent alerts, and the rest WAIT and go out the next morning by themselves.
 *   3. Privacy    privacyProblems_() checks that the tracker and every portal are private.
 *   5. Sent once     every reminder, follow-up, announcement and weekly summary is marked "done" the moment
 *                 it goes out. Running the same send again (by hand, after a pause, or the waiting list)
 *                 skips everyone already done, so nobody gets the same email twice.
 *   4. ONE email account  Only one Google account runs the automatic emails, portals and form
 *                 handling (the "email account"). Collaborators can use the tracker, but their
 *                 account never sends automatic emails, so nothing is ever sent twice.
 * ================================================================ */

var MAIL_RESERVE = 5;                 // always kept for urgent alerts and your copy reports
var QUEUE_KEY = 'MAIL_QUEUE';
var QUEUE_HOUR = 7;                   // waiting emails go out the next morning at about this hour
var WAIT_NOTE = 'Waiting: daily email limit reached. Sends tomorrow morning by itself.';

// The ONE way the tracker sends email. Adds Reply-to and the logo (when the logo is attached).
// 🧪 Test mode: every email goes to ONE address (you), marked [TEST n], and nothing is marked done or logged.
var TEST_MODE_ = null, TEST_COUNT_ = 0;

function sendMail_(o) {
  if (TEST_MODE_) {
    o.to = TEST_MODE_; delete o.cc; delete o.bcc;
    o.subject = `[TEST ${++TEST_COUNT_}] ` + String(o.subject || '').replace(/^\[TEST\]\s*/, '');
  }
  // Plain-text version (for phones' previews and mail apps that block pictures): built from the designed email
  // whenever none was written, or only a one-liner was.
  if (o.htmlBody && String(o.body || '').length < 80) o.body = htmlToText_(o.htmlBody) || o.body || o.subject;
  const r = replyTo_();
  if (r && !o.replyTo) o.replyTo = r;
  if (!o.inlineImages) o.inlineImages = logoInlineImages_();
  if (!Object.keys(o.inlineImages).length) delete o.inlineImages;
  MailApp.sendEmail(o);
}

// The designed email as plain text: headings and paragraphs on their own lines, buttons/links as "Label: link".
function htmlToText_(html) {
  let t = String(html || '')
    .replace(/<(style|script|head|title)[^>]*>[\s\S]*?<\/\1>/gi, '')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<img[^>]*>/gi, '')
    .replace(/<a\s[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/gi, (m, href, label) => {
      const l = label.replace(/<[^>]+>/g, '').trim();
      return /^mailto:/i.test(href) || !l || l === href ? (l || href) : `${l}: ${href}`;
    })
    .replace(/<(p|div|table|tr|h[1-6])[\s>]/gi, m => '\n' + m)
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|h[1-6]|tr|li|table)>/gi, '\n')
    .replace(/<\/t[dh]>/gi, '  ')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#39;|&rsquo;/g, "'").replace(/&middot;/g, '·').replace(/&ndash;/g, '–').replace(/&mdash;/g, '—');
  t = t.split('\n').map(l => l.replace(/[ \t]+/g, ' ').trim()).join('\n').replace(/\n{3,}/g, '\n\n').trim();
  return t;
}

// Settings › "Replies go to" (blank = your own account, Google's normal behaviour)
function replyTo_() {
  try { return emailOf_(getSettings_().replyTo); } catch (e) { return ''; }
}

// Emails a big send may still use today (keeps MAIL_RESERVE for alerts)
// How many bulk emails (reminders, follow-ups, announcements, welcomes) may go now. Always leaves room for
// urgent help alerts AND for website sign-in codes, so scholars can still sign in after Friday's reminders.
function mailRoom_() {
  const codes = (typeof WEB_CHECKIN !== 'undefined' && WEB_CHECKIN && typeof WEB !== 'undefined') ? WEB.mailReserve : 0;
  return Math.max(0, MailApp.getRemainingDailyQuota() - MAIL_RESERVE - codes);
}

// Google's daily email limit for this account: about 100 for a free Gmail account, 1,500 for Google Workspace
function dailyMailCap_() {
  let me = ''; try { me = Session.getEffectiveUser().getEmail(); } catch (e) { me = ''; }
  return /@(gmail|googlemail)\.com$/i.test(me) ? 100 : 1500;
}

/* ---------- waiting list (sent the next morning) ---------- */

function mailQueue_() {
  try { return JSON.parse(docProps_().getProperty(QUEUE_KEY) || '[]'); } catch (e) { return []; }
}

// job = { kind: 'reminder' | 'followup' | 'announce', schools: [names], week?: {monday, deadline}, a?: announcement }
function queueEmails_(job) {
  if (TEST_MODE_) return true;
  if (!job.schools.length) return true;
  const q = mailQueue_();
  const same = q.find(j => j.kind === job.kind && JSON.stringify(j.week || null) === JSON.stringify(job.week || null) &&
                           JSON.stringify(j.a || null) === JSON.stringify(job.a || null));
  if (same) same.schools = [...new Set(same.schools.concat(job.schools))];
  else q.push(job);
  const text = JSON.stringify(q);
  if (text.length > 8500) return false;                  // too big to keep (very long announcement): report as not sent
  docProps_().setProperty(QUEUE_KEY, text);
  scheduleQueue_();
  return true;
}

// One trigger, tomorrow morning. Replaced (never stacked) each time.
function scheduleQueue_() {
  removeTriggers_(['sendWaitingEmails']);
  const at = new Date();
  at.setDate(at.getDate() + 1);
  at.setHours(QUEUE_HOUR, Math.floor(Math.random() * 20), 0, 0);
  ScriptApp.newTrigger('sendWaitingEmails').timeBased().at(at).create();
}

// Runs by itself the morning after the limit was reached (also safe to run by hand).
function sendWaitingEmails() {
  const q = mailQueue_();
  docProps_().deleteProperty(QUEUE_KEY);
  removeTriggers_(['sendWaitingEmails']);
  q.forEach(job => {
    try {
      const wk = job.week ? { monday: new Date(job.week.monday), deadline: new Date(job.week.deadline) } : null;
      if (job.kind === 'announce') sendAnnouncementTo_(job.a, job.schools);
      else if (job.kind === 'mannounce') sendMemberAnnouncementTo_(job.a, job.schools);          // 👥 message to members (emails)
      else if (/^m(reminder|followup)$/.test(job.kind)) sendMemberWeekly_(job.kind.slice(1), wk, job.schools);   // 👥 members (emails)
      else sendReminders_(job.kind, job.schools, job.week ? { monday: new Date(job.week.monday), deadline: new Date(job.week.deadline) } : null);
    } catch (e) { Logger.log('Waiting emails (' + job.kind + '): ' + e.message); }
  });
}

// For 🩺 and the Today panel: how many emails are waiting
function waitingCount_() {
  return mailQueue_().reduce((n, j) => n + j.schools.length, 0);
}

/* ---------- privacy ---------- */

// Problems found with who can open the tracker and the portals. [] = all private.
function privacyProblems_() {
  const out = [];
  const ss = SpreadsheetApp.getActive();
  const me = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  const people = f => f.getEditors().concat(f.getViewers()).map(u => String(u.getEmail() || '').toLowerCase())
                       .filter(e => e && e !== me);
  const f = DriveApp.getFileById(ss.getId());
  const access = f.getSharingAccess();
  if (access !== DriveApp.Access.PRIVATE) {
    out.push({ level: 'bad', text: 'The tracker can be opened by link (' + access + ')',
               fix: 'Share › General access › Restricted. It holds every scholar\'s personal data.' });
  }
  const others = [...new Set(people(f))];
  if (others.length) out.push({ level: 'warn', text: 'The tracker is shared with: ' + others.join(', '),
                                fix: 'Remove anyone who is not ONSSE / START staff (Share › People with access). Never Campus Leads: they use their portal.' });
  const store = portalStore_();
  const leads = {};
  schoolRows_(ss).forEach(s => { leads[s.name] = s.email; });
  Object.keys(store.map).forEach(name => {
    const p = store.map[name];
    if (!p || !p.id) return;
    try {
      const pf = DriveApp.getFileById(p.id);
      if (pf.isTrashed()) return;
      const extra = people(pf).filter(e => e !== String(p.email || '').toLowerCase());
      if (pf.getSharingAccess() !== DriveApp.Access.PRIVATE || extra.length) {
        out.push({ level: 'bad', text: `${name}'s portal is open to more than its Campus Lead` + (extra.length ? ' (' + extra.join(', ') + ')' : ''),
                   fix: 'Click 🛠 Setup tools › 🔐 Create / update Campus Lead portals: it makes every portal private again.' });
      }
    } catch (e) { /* portal deleted: 🔐 Portals makes a new one */ }
  });
  return out;
}

/* ---------- 4. one email account ---------- */
// Google runs automatic schedules ("triggers") per account, and each account can only see and remove its own.
// If two accounts each turned on the weekly emails, Campus Leads would get every email twice.
// So the tracker remembers ONE email account. Every automatic run first checks that it is that account;
// a schedule left behind in any other account switches itself off the first time it runs.

var OWNER_KEY = 'AUTOMATION_OWNER';

function me_() { return String(Session.getEffectiveUser().getEmail() || '').toLowerCase(); }
function automationOwner_() { return String(docProps_().getProperty(OWNER_KEY) || '').toLowerCase(); }

// true when this account is (or just became, because none was set) the email account
function claimIfFree_() {
  const owner = automationOwner_(), me = me_();
  if (!me) return false;
  if (!owner) { docProps_().setProperty(OWNER_KEY, me); return true; }
  return owner === me;
}

// For menu actions that turn on automatic sending. Asks before another account takes over.
function ensureOwner_(what) {
  if (claimIfFree_()) return true;
  const owner = automationOwner_(), ui = ui_();
  const ok = ui.alert('Another account sends the automatic emails',
    `The automatic emails and updates are sent from ${owner}.\n\n` +
    `To ${what} from your account (${me_()}) instead, choose YES. From then on:\n` +
    '• automatic emails come from your account and use YOUR daily email limit;\n' +
    `• the schedule in ${owner} switches itself off, so nothing is sent twice.\n\n` +
    `Choose NO to leave it with ${owner}.`, ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return false;
  takeOverAutomation_();
  return true;
}

// Moves everything automatic to this account: form answers, sheet sync, weekly emails, daily portal update.
function takeOverAutomation_() {
  const ss = SpreadsheetApp.getActive(), props = docProps_();
  props.setProperty(OWNER_KEY, me_());
  installSubmitTrigger_(ss);
  installSyncTrigger_(ss);
  applySchedule_();
  if (props.getProperty('PORTALS_ON') === '1') { try { installPortalTrigger_(); } catch (e) { Logger.log('Portal schedule: ' + e.message); } }
  if (props.getProperty('MFORM_ID') && !ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'memberDaily')) {
    ScriptApp.newTrigger('memberDaily').timeBased().everyDays(1).atHour(6).inTimezone(CONFIG.timezone).create();
  }
  try { installExtrasTrigger_(); } catch (e) { Logger.log('Extras schedule: ' + e.message); }
  if (props.getProperty('CAL_ON') === '1' && !ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'refreshCalendarsAuto')) {
    ScriptApp.newTrigger('refreshCalendarsAuto').timeBased().everyDays(1).atHour(6).inTimezone(CONFIG.timezone).create();   // 📅 new calendars are made in this account
  }
}

// Menu ⏰ Automatic emails: "Make my account send the automatic emails"
function takeOverAutomationMenu() {
  requireSetup_();
  const owner = automationOwner_();
  if (owner === me_()) return notify_(`✅ Your account (${owner}) already sends the automatic emails.`);
  if (!ensureOwner_('send them')) return;
  notify_(`✅ Done. Automatic emails, form answers and portal updates now run from ${me_()}.` +
          (owner ? `\n\nThe old schedule in ${owner} switches itself off the next time it runs.` : ''));
}

// First line of every AUTOMATIC run (pass the event e). false = this account is not the email account:
// its own leftover schedule is removed and nothing is done. Menu clicks (no e) are never blocked.
function ownerGuard_(e) {
  if (!e || !e.triggerUid) return true;
  if (claimIfFree_()) return true;
  try {
    ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() !== 'sendWaitingEmails') ScriptApp.deleteTrigger(t); });
  } catch (err) { /* nothing to remove */ }
  Logger.log('Automatic run skipped: ' + me_() + ' is not the email account (' + automationOwner_() + ')');
  return false;
}

/* ---------- 5. sent once: the "done" list ---------- */
// key = what | which week (or which announcement) | to whom   ->   when it was sent. Kept 10 weeks.

var LEDGER_KEY = 'SENT_LEDGER';
var LEDGER_DAYS = 70;

function ledger_() {
  try { return JSON.parse(docProps_().getProperty(LEDGER_KEY) || '{}'); } catch (e) { return {}; }
}
function sentKey_(what, period, who) { return [what, period, String(who || '').toLowerCase()].join('|'); }
function alreadySent_(what, period, who) { return !TEST_MODE_ && !!ledger_()[sentKey_(what, period, who)]; }

// Marked right after each email, so even a run that Google stops halfway never repeats the ones already sent
function markSent_(what, period, who) {
  if (TEST_MODE_) return;                                         // tests never count as sent
  const L = ledger_(), now = Date.now(), old = now - LEDGER_DAYS * 86400000;
  Object.keys(L).forEach(k => { if (L[k] < old) delete L[k]; });
  L[sentKey_(what, period, who)] = now;
  let text = JSON.stringify(L);
  while (text.length > 8500) {                                   // stay inside Google's size limit: drop the oldest
    const oldest = Object.keys(L).sort((a, b) => L[a] - L[b])[0];
    delete L[oldest]; text = JSON.stringify(L);
  }
  docProps_().setProperty(LEDGER_KEY, text);
}

// Short fingerprint of an announcement (same subject + text = same announcement)
function announceId_(a) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(a.subject) + '\n' + String(a.body || (a.paras || []).join('\n\n')) + '\n' + String(a.link || ''))).slice(0, 10);
}
