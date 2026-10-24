/* =====================================================================
 *  START SCHOLAR TRACKER  ·  02_Menu_Setup
 *  ---------------------------------------------------------------------
 *  Inside: The START Tracker menu and ① Set up / fix the tracker (incl. the one-time switch to a plain database).
 *  To customise: Menu names and the Start here guide text.
 *
 *  The code has 9 files + Email_Reminder and Web_App (HTML). Keep these exact names:
 *    00_Config, 01_Region, 02_Menu_Setup, 03_Forms, 04_Emails, 05_Portals, 06_Calendar, 07_Look, 08_Tools
 *  To update: open each file, select all (Ctrl+A), paste the new version, Save.
 *  Then in the Google Sheet: START Tracker › ① Set up / fix the tracker.
 *  You never need to run anything from here: use the START Tracker menu in the Sheet.
 * ===================================================================== */

// The ▶ Run button runs the FIRST function of the open file: this one, so it is always safe.
function about_02_Menu_Setup() {
  Logger.log('START Tracker' + (typeof START_MODULES === 'string' ? ' ' + START_MODULES : '') + ', file 02_Menu_Setup. Nothing to run here: open the Google Sheet and use the START Tracker menu.');
}


// ━━━━━━━━━━━━━━━━━━━━  Menu  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  MENU  (every item goes through safe_ so errors are plain English)
 * ================================================================ */

function onOpen() {
  let ui;
  try { ui = SpreadsheetApp.getUi(); }
  catch (e) {                                       // run from the Apps Script editor: nothing to do here
    Logger.log('onOpen only works when the Google Sheet opens. Close this, reload the Sheet, and use the START Tracker menu.');
    return;
  }
  let copied = false;
  try { copied = checkTrackerCopy_(); } catch (e) { /* view-only users */ }
  try { migrateTabs_(SpreadsheetApp.getActive()); } catch (e) { /* view-only users */ }
  // Grouped by what you want to DO. Things that happen by themselves have no menu item.
  if (DB_MODE) { dbMenu_(ui); if (copied) { try { ui.alert('START Tracker', COPY_NOTICE, ui.ButtonSet.OK); } catch (e) { /* shown later */ } } return; }
  ui.createMenu('START Tracker')
    .addItem('🗓 Today: what needs me', 'menuToday')
    .addItem('① Set up / fix the tracker', 'menuSetup')
    .addSeparator()
    .addSubMenu(ui.createMenu('✉️ Send emails now')
      .addItem('📨 Send welcome emails to new members', 'menuWelcome')
      .addItem('🔔 Send this week\'s check-in reminder', 'menuRemindNow')
      .addItem('🔁 Send a follow-up to those who haven\'t answered', 'menuFollowUpNow')
      .addItem('📣 Write a message (Leads, members or everyone)', 'menuAnnounce')
      .addSeparator()
      .addItem('🧪 Send me a test of every email', 'menuTestAll'))
    .addSubMenu(ui.createMenu('⏰ Automatic emails')
      .addItem('🚀 Turn ON automatic weekly emails', 'menuWeeklyOn')
      .addItem('⏸ Turn OFF automatic weekly emails', 'menuWeeklyOff')
      .addSeparator()
      .addItem('📬 Send me this week\'s summary now', 'menuSendNow')
      .addItem('🔑 Make my account send the automatic emails', 'menuTakeOver'))
    .addSubMenu(ui.createMenu('📋 Copy a message for Messenger')
      .addItem('For Campus Leads: this week\'s reminder', 'menuFocalMessage')
      .addItem('For members: the check-in link', 'menuMemberMsg'))
    .addSubMenu(ui.createMenu('📅 Events & calendar')
      .addItem('Add an event', 'menuAddEvent')
      .addItem('📋 Attendance list for an event', 'menuAttendance')
      .addSeparator()
      .addItem('Set up / update calendars', 'menuCalendar')
      .addItem('⏸ Stop daily calendar updates', 'menuCalendarOff'))
    .addItem('📞 Record a call or message', 'menuOutreach')
    .addItem('🌐 Website for members, Leads & you', 'menuWebsite')
    .addSubMenu(ui.createMenu('📄 Reports & backups')
      .addItem('📄 Make an end-of-term report', 'menuTermReport')
      .addItem('📄 Last month\'s report now', 'menuMonthly')
      .addItem('💾 Back up the tracker now', 'menuBackup')
      .addItem('🗄 Archive graduated / dropped scholars', 'menuArchive'))
    .addSeparator()
    .addItem('📋 Refresh join request history', 'setupJoinHistory')
    .addItem('🩺 Check my tracker', 'menuCheck')
    .addSubMenu(ui.createMenu('🛠 Setup tools')
      .addItem('🔐 Create / update Campus Lead portals', 'menuPortals')
      .addItem('⏸ Stop daily portal updates', 'menuPortalsOff')
      .addSeparator()
      .addItem('📝 Set up / update member check-in', 'menuMembers')
      .addItem('🖼 Connect the pictures folder (banners + memes)', 'menuMemeFolder')
      .addItem('🖼 Add the tracker pictures', 'menuPictures')
      .addSeparator()
      .addItem('🏫 Fill in my region\'s schools', 'menuLoadSchools')
      .addItem('Full rebuild (all tabs, takes minutes)', 'menuFullRebuild')
      .addItem('🧳 Make a blank copy for other RRs', 'menuTemplate'))
    .addToUi();
  // Google gives opening the sheet only 30 seconds: nothing slow here. The guide is rebuilt by ① (quick check) instead.
  if (!DB_MODE) try {
    const st = SpreadsheetApp.getActive().getSheetByName(TAB.start);
    if (!st || String(st.getRange(st.getLastRow(), 1).getValue()) !== GUIDE_VERSION)
      SpreadsheetApp.getActive().toast('New version: click START Tracker › ① Set up / fix the tracker once.', 'START Tracker', 10);
  } catch (e) { /* view-only users */ }
  if (copied) { try { ui.alert('START Tracker', COPY_NOTICE, ui.ButtonSet.OK); } catch (e) { /* simple trigger: shown on the next menu click */ } }
}

// Database mode: the sheet only holds the data. Viewing, check-ins and reports are on the website.
function dbMenu_(ui) {
  ui.createMenu('START Tracker')
    .addItem('① Set up / fix the tracker', 'menuSetup')
    .addItem('🌐 Website for members, Leads & you', 'menuWebsite')
    .addSeparator()
    .addSubMenu(ui.createMenu('✉️ Send emails now')
      .addItem('🔔 This week\'s check-in reminder', 'menuRemindNow')
      .addItem('🔁 Follow-up to those who haven\'t answered', 'menuFollowUpNow')
      .addItem('📨 Welcome emails to new members', 'menuWelcome')
      .addItem('📣 Write a message', 'menuAnnounce')
      .addSeparator()
      .addItem('🧪 Send me a test of every email', 'menuTestAll'))
    .addSubMenu(ui.createMenu('⏰ Automatic emails')
      .addItem('🚀 Turn ON', 'menuWeeklyOn')
      .addItem('⏸ Turn OFF', 'menuWeeklyOff')
      .addItem('🔑 Make my account send them', 'menuTakeOver'))
    .addSeparator()
    .addItem('📋 Refresh join request history', 'setupJoinHistory')
    .addItem('🩺 Check my tracker', 'menuCheck')
    .addSubMenu(ui.createMenu('🛠 Setup tools')
      .addItem('📅 Set up / update calendars', 'menuCalendar')
      .addItem('🖼 Connect the pictures folder', 'menuMemeFolder')
      .addItem('🏫 Fill in my region\'s schools', 'menuLoadSchools')
      .addItem('🏫 School logos (website)', 'menuSchoolLogos')
      .addSeparator()
      .addItem('📄 End-of-term report', 'menuTermReport')
      .addItem('🧳 Blank copy for other RRs', 'menuTemplate')
      .addItem('🗑 Remove a person\'s data (privacy request)', 'menuRemovePerson')
      .addItem('Full rebuild (takes minutes)', 'menuFullRebuild'))
    .addItem('🧹 Fresh start: remove all data (backup first)', 'menuFreshStart')
    .addToUi();
}

function menuSetup()           { safe_(setup); }
function menuFreshStart()      { safe_(freshStart); }
function menuFullRebuild()     { safe_(() => { requireLook_(); docProps_().deleteProperty('SETUP_DONE'); setupFull(); }); }
function menuPictures()        { safe_(showPictureUpload); }
function menuTakeOver()        { safe_(takeOverAutomationMenu); }
function menuCalendar()        { safe_(setupCalendars); }
function menuAddEvent()        { safe_(showAddEvent); }
function menuMembers()         { safe_(setupMemberCheckin); }
function menuMemberMsg()       { safe_(showMemberMessage); }
function menuMemeFolder()      { safe_(connectMemeFolder); }
function menuAttendance()      { safe_(makeAttendanceList); }
function menuMonthly()         { safe_(monthlyReportNow); }
function menuBackup()          { safe_(backupNow); }
function menuArchive()         { safe_(archiveScholars); }
function menuWelcome()         { safe_(sendWelcomeNow); }
function menuCalendarOff()     { safe_(stopCalendarUpdates); }
function menuTermReport()      { safe_(makeTermReport); }
function menuToday()           { safe_(showToday); }
function menuAnnounce()        { safe_(showAnnouncement); }
function menuCheck()           { safe_(healthCheck); }
function menuLoadSchools()     { safe_(loadSchoolList); }
function menuWeeklyOn()        { safe_(turnOnWeekly); }
function menuWeeklyOff()       { safe_(turnOffWeekly); }
function menuFocalMessage()    { safe_(showFocalMessage); }
function menuSendNow()         { safe_(() => { requireSetup_(); weeklySummary(); }); }
function menuTestReminder()    { safe_(sendTestReminder); }
function menuTestAll()         { safe_(sendAllTests); }
function menuRemindNow()       { safe_(() => confirmAndRun_('reminder')); }
function menuFollowUpNow()     { safe_(() => confirmAndRun_('followup')); }
function menuOutreach()        { safe_(showOutreachDialog); }
function menuPortals()         { safe_(refreshPortalsMenu); }
function menuPortalsOff()      { safe_(stopPortalUpdates); }
function menuTemplate()        { safe_(makeCleanTemplate); }
function menuRemovePerson()    { safe_(removePersonData); }
function menuWebsite()         { safe_(showWebsite); }

function safe_(fn) {
  try {
    if (checkTrackerCopy_()) { SpreadsheetApp.getUi().alert('START Tracker', COPY_NOTICE, SpreadsheetApp.getUi().ButtonSet.OK); return; }
    fn();
  }
  catch (e) {
    const m = String(e && e.message || e);
    let msg;
    if (/timed out|Service unavailable|try again/i.test(m)) {
      msg = 'Google is slow right now.\n\nWait one minute, then click the same menu item again. It is safe to repeat.';
    } else if (/SETUP_FIRST/.test(m)) {
      msg = 'The tracker is not set up yet.\n\nClick START Tracker › ① Set up / fix the tracker first.';
    } else if (/authoriz|permission/i.test(m)) {
      msg = 'Google needs your permission first.\n\nClick the menu item again and choose Allow. ' +
            'See the "Start here" tab, step 2, if you get stuck.';
    } else {
      msg = 'Something went wrong:\n\n' + m +
            '\n\nTry again in a minute. If it keeps happening, take a screenshot of this box and send it to your Regional Representative (RR).';
    }
    SpreadsheetApp.getUi().alert('START Tracker', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function requireSetup_() {
  if (!docProps_().getProperty(WEB_CHECKIN ? 'SETUP_DONE' : 'FORM_ID')) throw new Error('SETUP_FIRST');
}


// ━━━━━━━━━━━━━━━━━━━━  StartHere  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  START HERE TAB  (plain-language guide inside the sheet)
 * ================================================================ */

var GUIDE_VERSION = 'Guide version 30';
var GUIDE_WIDTH = 980;            // the guide column and the banner use the same width, so they line up

function buildStartHere_(ss) {
  if (DB_MODE) return;                                   // database mode: not used (see 00_Config)
  const old = ss.getSheetByName(TAB.start);
  if (old) {
    if (String(old.getRange(old.getLastRow(), 1).getValue()) === GUIDE_VERSION) return;
    try { old.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove()); } catch (e) {}
    ss.deleteSheet(old);                       // older guide: rebuild it with the new steps
  }
  const sh = ss.insertSheet(TAB.start, 0);
  const lines = [
    ['START Tracker: Start here', 'title'],
    ['Everything is in the "START Tracker" menu at the top. Each tab has its own short steps on the right →', 'note'],
    ['', ''],
    ['FIRST TIME  ·  10 minutes', 'head'],
    ['1.  Click ① Set up / fix the tracker, allow access (see "Permission screen" below) and pick your REGION when asked. Once: 🛠 Setup tools › 🖼 Add the tracker pictures, then ① again.', ''],
    ['2.  Settings: fill in the yellow cells. Region is a drop-down: forms, emails, calendars and the province list follow it.', ''],
    ['3.  Schools: add your schools (🛠 Setup tools › 🏫 Fill in my region\'s schools, if your region has a ready-made list).', ''],
    ['4.  Scholars: paste the ONSSE masterlist at cell A2.', ''],
    ['5.  Pick each school\'s START Campus Lead: Scholars › Role (J) › ⭐ START Campus Lead, or Schools › column D. One per school.', ''],
    ['6.  Click ⏰ Automatic emails › 🚀 Turn ON, then 🛠 Setup tools › 🔐 Create / update Campus Lead portals.', ''],
    ['7.  Optional: 📅 Events & calendar › Set up / update calendars (shared START calendars).', ''],
    ['8.  Optional: 🛠 Setup tools › 📝 Set up / update member check-in (weekly vibe check for every member).', ''],
    ['', ''],
    ['EVERY WEEK  ·  5 minutes', 'head'],
    ['•  Open 🗓 Today (menu): it lists who to follow up, help requests and missing emails. Click an item to go to it.', ''],
    ['•  Read your weekly summary email. Contact the schools under "No reply".', ''],
    ['•  After every call or message to a school: click 📞 Record a call or message.', ''],
    ['•  Check the Dashboard. Red schools need you.', ''],
    ['•  End of term: 📄 Reports & backups › Make an end-of-term report (a ready Google Doc for DOST-SEI).', ''],
    ['', ''],
    ['HOW THE TABS CONNECT', 'head'],
    ['✏️ You type in only 4 tabs:   Settings  ·  Schools  ·  Scholars  ·  Outreach log', ''],
    ['🔒 Everything else fills in by itself from them:   Dashboard  ·  START Campus Leads  ·  Campus Lead check-ins  ·  Check-in answers  ·  Reminder log  ·  portals  ·  emails', ''],
    ['So if something looks wrong in a 🔒 tab, fix it in the ✏️ tab it comes from. Each tab\'s "About this tab" panel on the right shows ← where it gets data and → what it updates.', 'note'],
    ['', ''],
    ['COLOURS', 'head'],
    ['✏️ blue heading = you type   ·   🔒 navy heading = automatic   ·   yellow cell = a setting', ''],
    ['Green school = checked in within 7 days   ·   Red school = quiet too long', ''],
    ['Orange scholar row = duplicate name   ·   Yellow ⭐ row = START Campus Lead', ''],
    ['', ''],
    ['RULES', 'head'],
    ['•  Never share this file: it has every scholar\'s personal data. START Campus Leads get a private portal (🛠 Setup tools › 🔐).', ''],
    ['•  Don\'t rename tabs, move columns or change the form\'s question titles.', ''],
    ['•  Rename a school or scholar one cell at a time; the new name is carried everywhere.', ''],
    ['', ''],
    ['MORE DETAIL', 'head'],
    ['Permission screen', 'sub'],
    ['Continue › choose your account › Advanced › Go to … (unsafe) › Allow. This is normal: it is your own copy.', 'note'],
    ['Emails', 'sub'],
    ['⏰ Automatic emails › 🚀 Turn ON (or tick 🚀 in Settings) turns ON, every week: the summary to you + the reminder and follow-up to START Campus Leads. ' +
     'After each send you get a copy and a report.  ·  ✉️ Send emails now: welcome, reminder, follow-up, a message, or a test  ·  ⏰ Automatic emails: on / off  ·  urgent alert to you, automatic  ·  ' +
     '🛠 Setup tools › 🔐 Create / update Campus Lead portals also emails each Lead their portal link. Reminders and follow-ups are listed in "Reminder log".', 'note'],
    ['START Campus Leads & portals', 'sub'],
    ['Emails are typed in ONE place only: Scholars › Contact (email). Schools, reminders, portals and announcements read them from there. ' +
     'Each portal is private to that email and updates every morning and after each check-in. ' +
     'Its link is in every reminder and in their Drive › Shared with me.', 'note'],
    ['Deadlines & term', 'sub'],
    ['Targets and dates are in Settings. Cards turn yellow 30 days and red 14 days before a deadline. ' +
     'Set the term start date (Dashboard, B3) each term.', 'note'],
    ['Other RRs', 'sub'],
    ['🧳 Make a blank copy for other RRs: a copy with no scholar data, safe to share.', 'note'],
    ['Form colours (optional)', 'sub'],
    ['Open the form › palette icon › colour #12203A, font Poppins.', 'note'],
    ['Daily email limit & replies', 'sub'],
    ['Google allows a limited number of emails per day. If a send is bigger than what is left, the rest wait and go out the next ' +
     'morning by themselves ("Waiting" in Reminder log). Replies from Campus Leads go to Settings › "Replies go to" (blank = you).', 'note'],
    ['Quiet schools & privacy', 'sub'],
    ['With the weekly summary you get a list of schools quiet for 2+ weeks, each with a ready-to-copy Messenger message (Settings to turn off). ' +
     'The summary also warns you in red if this tracker or a portal is ever shared with someone it should not be.', 'note'],
    ['Shared calendars', 'sub'],
    ['📅 Events & calendar makes 4 kinds of calendars: START <your region> (all active scholars), START Campus Leads (check-in dates), START – <school> (its Lead can add events) ' +
     'and My START work (only you: follow-ups and deadlines). Add events with 📅 Events & calendar › Add an event (ready-made START events, repeats); see the month in the Calendar view tab. Sharing follows Scholars every morning. ' +
     'One-time: Apps Script › Services › + › Google Calendar API.', 'note'],
    ['The weekly check-in (one form for everyone)', 'sub'],
    ['📝 ONE short form for everyone, Campus Leads included (sign in with Google; emails not in Scholars are ignored). A Campus Lead\'s answer also counts as their school\'s check-in (Campus Lead check-ins tab). The older Campus Lead form is closed and points to it. 📨 Welcome each member once with ✉️ Send emails now › 📨 Send welcome emails (check-in link + "add the START calendar"). From then on they get the check-in by email every week (with 🚀 ON) and a follow-up if they have not answered; Campus Leads only Messenger the ones still missing. It shows only what is needed that week: ' +
     'events of the week (from Events), the learning question while DataCamp runs (Settings). Meme pictures: Drive folder "START member check-in pictures" ' +
     '(vibe-1.jpg, events-1.jpg, learning-1.jpg, help-1.jpg …). Campus Leads share it and see their members in their portal.', 'note'],
    ['Real time & sending once', 'sub'],
    ['Edits to names, emails, schools, Campus Leads and events update everything by themselves as you type: no need to click ①. ' +
     '① is now a quick check (seconds); the full rebuild only runs after a code update (or 🛠 Setup tools › Full rebuild). ' +
     'Every email is marked done when sent, so running a send again never emails the same person twice.', 'note'],
    ['Collaborators & the email account', 'sub'],
    ['Only ONE Google account sends the automatic emails (shown in 🩺 and 🗓 Today), so nothing is ever sent twice. Collaborators can use the menu; ' +
     'what they send uses their own daily limit. To move the automatic emails to another account: ⏰ Automatic emails › 🔑 Make my account send the automatic emails.', 'note'],
    ['Announcements', 'sub'],
    ['✉️ Send emails now › 📣 Write a message: one message to Campus Leads, members or everyone, in the tracker\'s email design. Send a test to yourself first.', 'note'],
    ['Runs by itself (tick ON/OFF in Settings)', 'sub'],
    ['✨ The weekly check-in closes after its deadline and opens again on its day · Campus Leads get an email the day before an event · ' +
     'on the 1st: last month\'s report (a Google Doc link) · Mondays: a self-check email ONLY if something is wrong · ' +
     'Sundays: a private backup in "START Tracker backups" (last 8 kept) · emails in Scholars that look wrong turn orange · ' +
     'the Dashboard ends with a school leaderboard. Menu 📄 Reports & backups: reports, backup now, archive graduated / dropped scholars. Attendance lists: 📅 Events & calendar.', 'note'],
    ['Something wrong?', 'sub'],
    ['Click 🩺 Check my tracker: it lists what is wrong and how to fix it. Then ① Set up / fix the tracker: it is safe and keeps your data. Still stuck? Send a screenshot of the error to your Regional Representative (RR).', 'note'],
    ['', ''],
    [GUIDE_VERSION, 'note'],
  ];
  sh.getRange(1, 1, lines.length, 1).setValues(lines.map(l => [l[0]]));
  lines.forEach((l, i) => {
    const c = sh.getRange(i + 1, 1);
    c.setFontFamily(THEME.font).setFontSize(11);
    if (l[1] === 'title') { c.setFontSize(22).setFontWeight('bold').setFontColor(THEME.navy); sh.setRowHeight(i + 1, 48); }
    if (l[1] === 'head') { c.setFontSize(12).setFontWeight('bold').setFontColor('#ffffff').setBackground(THEME.navy)
      .setBorder(null, null, true, null, null, null, THEME.sun, SpreadsheetApp.BorderStyle.SOLID_THICK); sh.setRowHeight(i + 1, 32); }
    if (l[1] === 'note') c.setFontColor(THEME.muted).setFontStyle('italic');
    if (l[1] === 'sub') c.setFontWeight('bold').setFontColor(THEME.blue);
  });
  sh.setColumnWidth(1, GUIDE_WIDTH);
  sh.getRange(1, 1, lines.length, 1).setWrap(true).setVerticalAlignment('middle');
  sh.setHiddenGridlines(true);
  sh.setTabColor(THEME.tabs.start);
  placeBanner_(sh, 'start', GUIDE_WIDTH);
  if (sh.getMaxColumns() > 1) sh.deleteColumns(2, sh.getMaxColumns() - 1);
}


// ━━━━━━━━━━━━━━━━━━━━  Setup  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  ① SETUP
 * ================================================================ */

// ① Set up / fix the tracker. Google stops any menu action after 6 minutes, so setup runs in steps
// and saves its place. If it gets close to the limit it stops safely and asks you to click ① again;
// the next click continues from the step where it stopped (nothing is done twice).
var SETUP_BUDGET_MS = 3.5 * 60 * 1000;   // stop well before Google's 6-minute limit (a long step still has room)
var SETUP_KEY = 'SETUP_PROGRESS';

// [name, work, optional]  optional = a finishing touch: if it fails, setup keeps going
function setupSteps_(ss) {
  const NOT_IN_DB = ['Start here tab', 'Dashboard tab', 'Links tab', 'Reminder log tab', 'START Campus Leads tab', 'Calendar view tab', 'Tab header pictures'];
  return [
    ['Start here tab',        () => { buildStartHere_(ss); applyTheme_(ss); }],
    ['Theme',                 () => applyTheme_(ss), true],
    ['Settings tab',          () => { buildSettings_(ss); docProps_().setProperty('SCHEDULE_SYNC', scheduleKey_());
                                      getOrCreate_(ss, TAB.scholars); getOrCreate_(ss, TAB.schools); }],
    ['Your region',           () => { askRegionIfMissing_(); applyRegion_(ss); }, true],   // asked once (05_Region)
    ['Outreach log tab',      () => buildOutreach_(ss)],
    ['Campus Lead check-in form',  () => { if (WEB_CHECKIN) ensureCheckinTabs_(ss); else buildForm_(ss); }],
    ['Scholars tab',          () => buildScholars_(ss)],
    ['Schools tab',           () => buildSchools_(ss)],
    ['Data flow schema',      () => { ensureDataFlowSchema_(ss); captureWeeklyRoster_(ss); }],
    ['Dashboard tab',         () => buildDashboard_(ss)],
    ['Links tab',             () => buildLinks_(ss)],
    ['Protections and automatic updates', () => {
      protect_(ss);
      if (claimIfFree_()) {                                                  // only the email account (20_Safety)
        installSubmitTrigger_(ss); installSyncTrigger_(ss);
        if (!WEB_CHECKIN && !ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'memberDaily')) {   // forms follow the calendar daily
          ScriptApp.newTrigger('memberDaily').timeBased().everyDays(1).atHour(6).inTimezone(CONFIG.timezone).create();
        }
        installExtrasTrigger_();                                             // ✨ 24_Extras: open/close, backups, reports…
      }
    }],
    ['Extras: check-in open/closed, email checks, leaderboard', () => {
      try { applyFormWindow_(); } catch (e) { Logger.log('Open/close: ' + e.message); }
      try { checkScholarEmails_(ss); } catch (e) { Logger.log('Email checks: ' + e.message); }
      try { writeLeaderboard_(ss); } catch (e) { Logger.log('Leaderboard: ' + e.message); }
    }, true],
    ['Scholars table look',   () => polishScholars_(ss), true],
    ['Schools table look',    () => polishSchools_(ss), true],
    ['Campus Lead check-ins table look', () => polishCheckins_(ss), true],
    ['Reminder log tab',      () => buildLog_(ss), true],
    ['Headings and banners',  () => applyColumnKinds_(ss), true],
    ['START Campus Leads tab', () => buildFocalTab_(ss), true],
    ['Form school list',      () => updateDropdown_(), true],
    ['Events tab',            () => { if (ss.getSheetByName(TAB.events)) buildEvents_(ss); }, true],
    ['Calendar view tab',     () => buildCalView_(ss), true],
    ['Check-in answers tab',  () => { if (checkinOn_()) styleMemberTab_(ss); }, true],
    ['Tab header pictures',   () => repairHeaders_(ss), true],                       // any tab still bare (e.g. pictures added later)
    ['Plain database tabs',   () => { if (DB_MODE) tidyToDatabase_(ss); }, true],
    ['Same table look on every tab', () => styleNewerTables_(ss, true), true],
  ].filter(x => !DB_MODE || NOT_IN_DB.indexOf(x[0]) === -1);
}

// ① is QUICK once the tracker is built: everything you edit (names, emails, Campus Leads, schools, roles, events)
// already updates by itself as you type. The full rebuild (all tabs, headings, banners, the form) only runs
// the first time, after you replace code files with a newer version, or from 🛠 Setup tools › Full rebuild.
var SETUP_VERSION = 'v61';
var CORE_TABS_ = () => DB_MODE ? [TAB.settings, TAB.schools, TAB.scholars, TAB.checkins] : [TAB.settings, TAB.schools, TAB.scholars, TAB.dashboard, TAB.checkins];

function requireLook_() {
  if (typeof tableLook_ !== 'function' || typeof markColumns_ !== 'function' || typeof emailPics_ !== 'function')
    throw new Error('07_Look.gs is empty or incomplete. Add its full source and save before running setup.');
}

function setup() {
  requireLook_();
  const ss = SpreadsheetApp.getActive(), props = docProps_();
  const built = props.getProperty('SETUP_DONE') === SETUP_VERSION && !props.getProperty(SETUP_KEY) &&
                CORE_TABS_().every(t => ss.getSheetByName(t)) && (WEB_CHECKIN || props.getProperty('FORM_ID'));
  return built ? quickFix_() : setupFull();
}

// Fast repair: seconds, not minutes. Only fixes what can drift; nothing is rebuilt or restyled.
function quickFix_() {
  const t0 = Date.now(), ss = SpreadsheetApp.getActive(), props = docProps_();
  ss.toast('Quick check… a few seconds.', 'START Tracker', 15);
  const done = [], later = [];
  const step = (label, fn) => {
    if (outOfTime_(45 * 1000)) { later.push(label); return; }                 // Google's 6-minute limit: the rest next click
    try { if (fn() !== false) done.push(label); } catch (e) { Logger.log(label + ': ' + e.message); }
  };
  migrateTabs_(ss);
  if (DB_MODE && tidyToDatabase_(ss) === 'later') return;
  if (!DB_MODE) step('Start here guide', () => { buildStartHere_(ss); });
  step('Region (Settings drop-down, provinces)', () => { askRegionIfMissing_(); applyRegion_(ss); });
  step('Emails in Schools read from Scholars', () => setSchoolEmailFormula_(ss.getSheetByName(TAB.schools)));
  step('Scholar counts and database fields', () => { buildSchools_(ss); ensureDataFlowSchema_(ss); captureWeeklyRoster_(ss); });
  step('Roles (one Campus Lead per school)', () => syncRoles_(ss));
  if (WEB_CHECKIN) step('Check-in tabs and old Google Forms', () => { ensureCheckinTabs_(ss); return retireForms_(ss) > 0; });
  if (WEB_CHECKIN) step('Fresh Scholars list (no-email rows -> Old masterlist)', () => freshScholarList_(ss) > 0);
  else step('Form school list', () => { if (props.getProperty('FORM_SYNC') === syncKey_(ss)) return false; updateDropdown_(); });
  if (!DB_MODE) step('START Campus Leads tab and portals', () => { if (!syncFocal_(ss)) buildFocalTab_(ss); });
  step('Automatic emails and form handling', () => {
    if (!claimIfFree_()) return false;                               // another account runs them (20_Safety)
    installSubmitTrigger_(ss); installSyncTrigger_(ss); applySchedule_();
    if (!WEB_CHECKIN && !ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'memberDaily')) {
      ScriptApp.newTrigger('memberDaily').timeBased().everyDays(1).atHour(6).inTimezone(CONFIG.timezone).create();
    }
    props.setProperty('SCHEDULE_SYNC', scheduleKey_());
  });
  if (!WEB_CHECKIN) step('Campus Lead and member forms (this week\'s version)', () => { refreshLeadForm_(ss); if (memberForm_()) refreshMemberForm_(ss); });
  step('Campus Lead check-ins: old questions hidden', () => { const sh = ss.getSheetByName(TAB.checkins); if (!sh) return false; hideOldQuestionCols_(sh); });
  step('Same table look on every tab', () => styleNewerTables_(ss) > 0);
  step('Check-in answers tab (name and look)', () => { if (!checkinOn_()) return false; styleMemberTab_(ss); });
  step('Daily jobs, check-in open/closed, email checks', () => {
    if (automationOwner_() === me_()) installExtrasTrigger_();
    applyFormWindow_(); checkScholarEmails_(ss); writeLeaderboard_(ss);
  });
  if (!DB_MODE) {
    step('Calendar view', () => renderCalView_(ss));
    step('Links tab (top part)', () => writeAutoLinks_(ss));
    step('Tab header pictures (any missing ones)', () => repairHeaders_(ss) > 0);
    step('"How to use" panels on every tab', () => { Object.keys(tutorials_()).forEach(n => { const sh = ss.getSheetByName(n); if (sh) placeTutorial_(sh); }); });
  }
  const secs = Math.round((Date.now() - t0) / 1000);
  if (later.length) {
    notify_(`⏳ Part done\n\nGoogle allows a few minutes per click. Click ① again to finish (safe to repeat).\n\n• Still to do: ` + later.join(', '));
    return;
  }
  notify_(`✅ All good\n\n• ` + done.join('\n• '));
}

// The full build, in steps that continue where they stopped (Google allows 6 minutes per click)
function setupFull() {
  requireLook_();
  const started = Date.now();
  const ss = SpreadsheetApp.getActive();
  const props = docProps_();
  ss.setSpreadsheetTimeZone(CONFIG.timezone);
  migrateTabs_(ss);

  const steps = setupSteps_(ss);
  let saved = null;
  try { saved = JSON.parse(props.getProperty(SETUP_KEY) || 'null'); } catch (e) { saved = null; }
  const resume = saved && Date.now() - saved.at < 6 * 3600 * 1000 && saved.next > 0 && saved.next < steps.length;
  let from = resume ? saved.next : 0;
  // Google stopped the last click INSIDE this step (it was too slow). Twice in a row: a finishing touch is skipped,
  // so ① never loops; ① or 🩺 can redo it later.
  let tries = resume && saved.started ? (saved.tries || 0) + 1 : 0;
  if (tries >= 2 && steps[from] && steps[from][2]) {
    Logger.log('Setup step skipped (too slow twice): ' + steps[from][0]);
    from++; tries = 0;
  }
  ss.toast(resume ? `Continuing setup from step ${from + 1} of ${steps.length}…` : 'Setting up… please wait.', 'START Tracker', 30);

  for (let i = from; i < steps.length; i++) {
    if (i > from && Date.now() - started > SETUP_BUDGET_MS) {
      props.setProperty(SETUP_KEY, JSON.stringify({ next: i, at: Date.now() }));
      notify_(`⏳ Setup ${Math.round(i / steps.length * 100)}% done\n\nClick ① Set up / fix the tracker again to continue.`);
      return;
    }
    // saved BEFORE each step: if Google stops this step, the next click starts here, not from the beginning
    props.setProperty(SETUP_KEY, JSON.stringify({ next: i, at: Date.now(), started: true, tries: i === from ? tries : 0 }));
    ss.toast(`Step ${i + 1} of ${steps.length}: ${steps[i][0]}…`, 'START Tracker', 60);
    try { steps[i][1](); }
    catch (e) {
      if (!steps[i][2]) throw e;                                  // main steps: show the error (safe_ explains it)
      Logger.log(`Setup step "${steps[i][0]}": ${e.message}`);   // finishing touches: keep going
    }
  }
  props.deleteProperty(SETUP_KEY);
  props.setProperty('SETUP_DONE', SETUP_VERSION);               // from now on ① is the quick check

  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  orderTabs_(ss);
  ss.setActiveSheet(ss.getSheetByName(TAB.settings));

  notify_('✅ Your tracker is ready\n\n' +
          '1. Fill in the yellow cells here in Settings\n' +
          '2. Schools tab: add schools (or 🛠 Setup tools › 🏫 Fill in my region\'s schools)\n' +
          '3. Scholars tab: paste the masterlist from A2\n' +
          '4. Schools tab: pick each Campus Lead\n' +
          '5. ⏰ Automatic emails › 🚀 Turn ON');
}

function buildScholars_(ss) {
  const sh = getOrCreate_(ss, TAB.scholars);
  sh.setFrozenRows(1);                         // headings come from columnSpecs_ (12_Style)
  const rows = Math.max(sh.getMaxRows() - 1, 500);
  if (sh.getMaxRows() < rows + 1) sh.insertRowsAfter(sh.getMaxRows(), rows + 1 - sh.getMaxRows());

  sh.getRange(2, 6, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'On leave', 'Graduated', 'Dropped', 'Suspended'], true)
    .setAllowInvalid(false).build());

  // Scholars › C: the same 13 START-DOST eligible programs shown on the website.
  // Keep existing cell values intact: historical entries outside the list remain visible
  // until the administrator reviews them, but cannot be selected for new records.
  sh.getRange(2, 3, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(START_PROGRAMS.list, true)
    .setAllowInvalid(false)
    .setHelpText('Select one of the 13 START-DOST eligible programs. Existing entries are not automatically changed.').build());

  const schoolList = ss.getSheetByName(TAB.schools).getRange(2, 1, CONFIG.maxSchools, 1);
  sh.getRange(2, 2, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(schoolList, true).setAllowInvalid(false)
    .setHelpText('Spell the school exactly as in the Schools tab (e.g. always "USeP").').build());

  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($A2<>"",COUNTIF($A$2:$A,$A2)>1)')
      .setBackground(THEME.orangeBg)
      .setRanges([sh.getRange(2, 1, rows, SCHOLAR_COLS)]).build(),
  ]);
  // Column D = year the DOST scholarship was awarded (batch), e.g. 2023
  const thisYear = new Date().getFullYear();
  const years = [];
  for (let y = thisYear + 1; y >= 2015; y--) years.push(String(y));
  sh.getRange(2, 4, rows, 1).setNumberFormat('0').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(years, true).setAllowInvalid(true)
    .setHelpText('Year the DOST scholarship was granted, e.g. 2023.').build());

  sh.getRange(2, 5, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied('=OR(E2="",REGEXMATCH(E2&"","@"))').setAllowInvalid(true)
    .setHelpText('Email address (used when this scholar becomes a START Campus Lead).').build());
  sh.getRange(2, 7, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['1st Year', '2nd Year', '3rd Year', '4th Year', '5th Year'], true)
    .setAllowInvalid(true).build()).setHorizontalAlignment('center');

  // column widths are fitted to the headings and content by markColumns_ (12_Style)
  placeTutorial_(sh);
  const f = sh.getFilter();
  if (f && f.getRange().getNumColumns() !== SCHOLAR_FILTER_COLS_) f.remove();
  addFilter_(sh, SCHOLAR_FILTER_COLS_);

  // J: Role dropdown (Member / ⭐ START Campus Lead). It is kept in step with Schools › START Campus Lead (D):
  // one Campus Lead per school, picked in either place. See ROLES and syncRoles_ in 08_FocalPortal.
  const J = sh.getRange(2, 10, rows, 1);
  if (String(sh.getRange('J2').getFormula()).indexOf('ARRAYFORMULA') !== -1) J.clearContent();   // older versions used a formula
  J.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ROLES, true).setAllowInvalid(false)
    .setHelpText('Member, or ⭐ START Campus Lead (one per school; picking a new one makes the old one a Member).').build());
  syncRoles_(ss);
}

var SCHOLAR_COLS = 9;   // Name, School, Program, Year awarded, Contact, Status, Year level, Campus, Hometown
var SCHOLAR_FILTER_COLS_ = 13; // Show role, other emails, scholarship and source without changing A:I business logic.

function addFilter_(sh, cols) {
  const f = sh.getFilter();
  if (f && f.getRange().getNumColumns() === cols) return;
  if (f) f.remove();                                        // older filter of another width
  try { sh.getRange(1, 1, sh.getMaxRows(), cols).createFilter(); } catch (e) { /* filter already there */ }
}

function buildForm_(ss) {
  // Google Forms often times out right after a form is created, so each step retries,
  // and the form ID is saved immediately. Re-running setup continues where it stopped.
  const props = docProps_();
  const st = getSettings_();
  const title = 'START Campus Lead Check-in' + (st.region ? ' – ' + st.region : '');
  let form = null;

  const id = props.getProperty('FORM_ID');
  if (id) {
    try { form = FormApp.openById(id); if (destId_(form) && destId_(form) !== ss.getId()) form = null; }
    catch (e) { form = null; }
  }
  if (!form) {                      // e.g. a copied file: adopt the form already linked to this sheet
    const linked = ss.getSheets().find(s => s.getFormUrl() && !isMemberSheet_(s));   // never the member form
    if (linked) {
      try {
        const f = FormApp.openByUrl(linked.getFormUrl());
        if (destId_(f) === ss.getId()) form = f;
      } catch (e) { /* ignore */ }
    }
  }
  if (!form) {
    form = retry_(() => FormApp.create(title));
    Utilities.sleep(3000);
  }
  props.setProperty('FORM_ID', form.getId());

  retry_(() => form.setTitle(title));


  try { if (typeof renameFormItems_ === 'function') renameFormItems_(form); } catch (e) { Logger.log('Form wording: ' + e.message); }
  addOnce_(form, Q.school, () => form.addListItem().setTitle(Q.school)
      .setHelpText(FORM_HELP.school)
      .setChoiceValues(schoolNames_(ss).length ? schoolNames_(ss) : [NO_SCHOOLS])
      .setRequired(true));
  addOnce_(form, Q.name, () => form.addTextItem().setTitle(Q.name)
      .setHelpText(FORM_HELP.name)
      .setRequired(true));
  addOnce_(form, Q.events, () => form.addMultipleChoiceItem().setTitle(Q.events)
      .setHelpText(FORM_HELP.events)
      .setChoiceValues(EVENT_GO_CHOICES)
      .setRequired(true));
  addOnce_(form, Q.dc, () => form.addMultipleChoiceItem().setTitle(Q.dc)
      .setHelpText(FORM_HELP.dc)
      .setChoiceValues(DC_CHOICES)
      .setRequired(true));
  addOnce_(form, Q.help, () => form.addMultipleChoiceItem().setTitle(Q.help)
      .setHelpText(FORM_HELP.help)
      .setChoiceValues(HELP_CHOICES)
      .setRequired(true));
  addOnce_(form, Q.concern, () => form.addParagraphTextItem().setTitle(Q.concern)
      .setHelpText(FORM_HELP.concern));
  // Data Privacy Act: a required "I agree" tick box, always the last question
  addOnce_(form, CONSENT_TITLE, () => form.addCheckboxItem().setTitle(CONSENT_TITLE)
      .setChoiceValues([CONSENT_AGREE])
      .setRequired(true));

  // Keep the questions in a sensible order (forms made by older versions get the new questions at the end)
  Object.values(Q).forEach((t, i) => {
    const it = retry_(() => form.getItems()).find(x => x.getTitle() === t);
    if (it && it.getIndex() !== i) retry_(() => form.moveItem(it, i));
  });
  const consent = retry_(() => form.getItems()).find(x => x.getTitle() === CONSENT_TITLE);
  const lastIx = retry_(() => form.getItems()).length - 1;
  if (consent && consent.getIndex() !== lastIx) retry_(() => form.moveItem(consent, lastIx));
  styleForm_(form, st);
  docProps_().deleteProperty('LFORM_KEY');   // the order above was reset: lay out the pages again
  try { refreshLeadForm_(ss, form); } catch (e) { Logger.log('Campus Lead form pages: ' + e.message); }

  if (destId_(form) !== ss.getId()) {
    retry_(() => form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId()));
  }

  for (let i = 0; i < 10; i++) {          // rename "Form Responses 1" to the Weekly Check-in tab
    SpreadsheetApp.flush();
    const fresh = SpreadsheetApp.openById(ss.getId());
    const sh = fresh.getSheets().find(s => (s.getFormUrl() || '').indexOf(form.getId()) !== -1);
    if (sh) {
      const old = fresh.getSheetByName(TAB.checkins);
      if (old && old.getSheetId() !== sh.getSheetId()) {
        if (old.getLastRow() <= 1 && !old.getFormUrl()) fresh.deleteSheet(old);   // empty (e.g. from the template)
        else old.setName(TAB.checkins + ' (old)');
      }
      sh.setName(TAB.checkins);
      sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
      return;
    }
    Utilities.sleep(1500);
  }
  throw new Error('The form was made, but Google is slow to create its answers tab. Wait a minute and click ① again.');
}

function destId_(form) {
  try { return form.getDestinationId(); } catch (e) { return ''; }
}

function buildSchools_(ss) {
  const sh = getOrCreate_(ss, TAB.schools);
  sh.setFrozenRows(1);                         // headings come from columnSpecs_ (12_Style)
  const n = CONFIG.maxSchools, last = n + 1;
  if (sh.getMaxRows() < last) sh.insertRowsAfter(sh.getMaxRows(), last - sh.getMaxRows());

  const C = ref_(TAB.checkins), S = TAB.scholars;
  let SC = 'B';                                             // School column of the Weekly Check-in tab
  try { SC = colLetter_(ckCols_(ss.getSheetByName(TAB.checkins)).school + 1); } catch (e) { /* new tracker: B */ }
  const count = [], lastCk = [], days = [];
  for (let r = 2; r <= last; r++) {
    count.push([`=IF(A${r}="","",COUNTIFS(${S}!B:B,A${r},${S}!F:F,"Active"))`]);
    lastCk.push([`=IF(A${r}="","",IF(COUNTIF(${C}!${SC}:${SC},A${r})=0,"",MAXIFS(${C}!A:A,${C}!${SC}:${SC},A${r})))`]);
    days.push([`=IF(G${r}="","",TODAY()-INT(G${r}))`]);
  }
  sh.getRange(2, 3, n, 1).setFormulas(count).setHorizontalAlignment('center');
  sh.getRange(2, 7, n, 1).setFormulas(lastCk).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, 8, n, 1).setFormulas(days).setNumberFormat('0');
  sh.getRange(2, 7, n, 2).setHorizontalAlignment('center');
  sh.getRange(2, 6, n, 1).setHorizontalAlignment('center');

  // D: START Campus Lead = pick from the Scholars list so the same name is used everywhere
  sh.getRange(2, 4, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(TAB.scholars).getRange('A2:A501'), true).setAllowInvalid(true)
    .setHelpText('Pick the START Campus Lead from your Scholars list. Their email comes from Scholars › Contact (email).').build());
  sh.getRange(2, 6, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Y', 'N'], true).setAllowInvalid(false).build());
  // E: the Campus Lead's email, read from Scholars (the ONE place emails are typed)
  setSchoolEmailFormula_(sh, true);

  const area = sh.getRange(2, 1, n, 1);       // status colour on the School cell only
  const q = `INDIRECT("${settingCell_('quietDays').replace(/\$/g, '')}")`;
  sh.setConditionalFormatRules([
    // a Campus Lead without an email in Scholars: the Contact cell turns red
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($D2<>"",$E2="")')
      .setBackground(THEME.redBg).setRanges([sh.getRange(2, 5, n, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($A2<>"",OR($G2="",AND(ISNUMBER($H2),$H2>${q})))`)
      .setBackground(THEME.redBg).setRanges([area]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($A2<>"",ISNUMBER($H2),$H2<=7)')
      .setBackground(THEME.greenBg).setRanges([area]).build(),
  ]);

  // School details I–O (from the ready-made school list) overlap with School/Province/Contact, so they are folded away (+ opens them)
  if (sh.getColumnGroupDepth(9) === 0) {
    sh.getRange(1, 9, 1, SCHOOL_INFO_HEADERS.length).shiftColumnGroupDepth(1);
    try { sh.getColumnGroup(9, 1).collapse(); } catch (e) { /* stays open; still fine */ }
  }
  // Q: last time you contacted this school (from the Outreach log)
  const O = ref_(TAB.outreach), lastC = [];
  for (let r = 2; r <= last; r++) {
    lastC.push([`=IF(A${r}="","",IF(COUNTIF(${O}!B:B,A${r})=0,"",MAXIFS(${O}!A:A,${O}!B:B,A${r})))`]);
  }
  sh.getRange(2, 17, n, 1).setFormulas(lastC).setNumberFormat('mmm d, yyyy').setHorizontalAlignment('center');
  addFilter_(sh, 17);
  placeTutorial_(sh);
}

// Schools › Contact (E) is ONE formula: the START Campus Lead's email from Scholars › Contact (email).
// migrate: the first time, emails that were typed in Schools are moved to Scholars so none are lost.
function setSchoolEmailFormula_(sh, migrate) {
  const ss = sh.getParent(), n = CONFIG.maxSchools, E = sh.getRange(2, 5, n, 1);
  const already = String(sh.getRange('E2').getFormula()).indexOf(TAB.scholars) !== -1;
  if (migrate && !already) {
    const sc = ss.getSheetByName(TAB.scholars), scholars = scholarRows_(ss);
    sh.getRange(2, 4, n, 2).getValues().forEach(r => {
      const lead = String(r[0]).trim(), mail = emailOf_(r[1]);
      if (!lead || !mail || !sc) return;
      const i = scholars.findIndex(x => String(x[0]).trim() === lead);
      if (i !== -1 && !emailOf_(scholars[i][4])) sc.getRange(scholarRowNumber_(sc, lead), 5).setValue(mail);
    });
  }
  E.clearContent().clearNote().clearDataValidations();
  sh.getRange('E2').setFormula(`=ARRAYFORMULA(IF(D2:D${n + 1}="","",IFERROR(VLOOKUP(D2:D${n + 1},` +
    `${ref_(TAB.scholars)}!$A$2:$E,5,FALSE)&"","")))`);
}

// Sheet row of a scholar by exact name (2 if not found, never used then)
function scholarRowNumber_(sc, name) {
  const names = sc.getRange(2, 1, Math.max(sc.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0]).trim());
  return names.indexOf(name) + 2;
}

// Warning-only protection: Google shows "are you sure?" before anyone types over formulas.
function protect_(ss) {
  ss.getProtections(SpreadsheetApp.ProtectionType.RANGE).concat(ss.getProtections(SpreadsheetApp.ProtectionType.SHEET))
    .filter(p => p.getDescription() === PROTECT_TAG).forEach(p => p.remove());
  const L = CONFIG.maxSchools + 1, sch = ss.getSheetByName(TAB.schools);
  [sch.getRange(`C1:C${L}`), sch.getRange(`E1:E${L}`), sch.getRange(`G1:H${L}`), sch.getRange('A1:H1'), sch.getRange(`P1:Q${L}`),
   ss.getSheetByName(TAB.scholars).getRange('A1:J1'),
   ss.getSheetByName(TAB.settings).getRange('A1:A' + (SETTINGS.length + 1)), ss.getSheetByName(TAB.settings).getRange('C1:C' + (SETTINGS.length + 1))]
    .forEach(r => r.protect().setDescription(PROTECT_TAG).setWarningOnly(true));
  [TAB.checkins, TAB.start].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s) s.protect().setDescription(PROTECT_TAG).setWarningOnly(true);
  });
  if (ss.getSheetByName(TAB.dashboard)) {
    const d = ss.getSheetByName(TAB.dashboard).protect().setDescription(PROTECT_TAG).setWarningOnly(true);
    d.setUnprotectedRanges([ss.getSheetByName(TAB.dashboard).getRange('B3')]);
  }
}

/* ---- Form look & wording (StaRRt IGNITE style, formal tone) ----
 * Google does not let scripts change a form's theme colour or font; see the Start here tab for the
 * 30-second manual step. Everything else is set here and re-applied each time ① runs. */
var FORM_HELP = {
  school:  'Already picked if you used your school\'s link. Not listed? Tell your Regional Representative.',
  name:    'Your full name.',
  events:  'Your overall impression. Members report their own attendance.',
  dc:      'Your best guess is fine.',
  help:    '"Yes, urgent" alerts your Regional Representative right away.',
  concern: 'Problems, requests or updates. Leave blank if none.',
};

function styleForm_(form, st) {
  retry_(() => form.setDescription(
    'Department of Science and Technology – Science Education Institute\n' +
    'Scholars Transforming Advancement and Research for Technology (START – DOST)\n\n' +
    'Good day! This weekly check-in is to be accomplished by the designated START Campus Lead of each school' +
    (st.region ? ' in ' + st.region : '') + '. It will take approximately two (2) minutes to complete. ' +
    'Your responses help START monitor scholar participation and respond promptly to concerns.\n\n' +
    privacyNote_()));
  retry_(() => form.setConfirmationMessage(
    'Thank you! Your weekly check-in has been recorded. ' +
    'If you indicated an urgent concern, your Regional Representative has been notified and will contact you shortly.'));
  retry_(() => form.setProgressBar(false));
  retry_(() => form.setShowLinkToRespondAgain(false));

  const items = retry_(() => form.getItems());
  Object.keys(FORM_HELP).forEach(k => {
    const it = items.find(x => x.getTitle() === Q[k]);
    if (!it) return;
    const t = it.getType();
    try {
      if (t === FormApp.ItemType.LIST) it.asListItem().setHelpText(FORM_HELP[k]);
      else if (t === FormApp.ItemType.TEXT) it.asTextItem().setHelpText(FORM_HELP[k]);
      else if (t === FormApp.ItemType.MULTIPLE_CHOICE) it.asMultipleChoiceItem().setHelpText(FORM_HELP[k]);
      else if (t === FormApp.ItemType.PARAGRAPH_TEXT) it.asParagraphTextItem().setHelpText(FORM_HELP[k]);
    } catch (e) { Logger.log('Help text not set for ' + k + ': ' + e.message); }
  });

  // Banner image at the very top (added once; replaced if it was deleted)
  const memes = leadMemeIds_();                       // meme pictures (23_Members) are not the banner
  let banner = typeof formBannerItem_ === 'function' ? formBannerItem_(form, 'LFORM_BANNER_ID', memes)
             : items.find(x => x.getType() === FormApp.ItemType.IMAGE && memes.indexOf(x.getId()) === -1);
  if (!banner && formBannerBlob_()) {                // needs the pictures (🖼 menu); added on the next ① otherwise
    try {
      banner = retry_(() => form.addImageItem().setImage(formBannerBlob_())
        .setTitle('').setAlignment(FormApp.Alignment.CENTER).setWidth(740));
      docProps_().setProperty('LFORM_BANNER_ID', String(banner.getId()));
    } catch (e) { Logger.log('Banner not added: ' + e.message); }
  }
  // a newer banner picture (e.g. the "START Campus Lead report" one) replaces the old one in place
  const props = docProps_();
  const fb = formBannerBlob_();
  if (banner && fb && props.getProperty('FORM_BANNER') !== fb.getName()) {
    try { retry_(() => banner.asImageItem().setImage(fb)); props.setProperty('FORM_BANNER', fb.getName()); }
    catch (e) { Logger.log('Banner not replaced: ' + e.message); }
  }
  if (banner && banner.getIndex() !== 0) retry_(() => form.moveItem(banner, 0));
}

/* ---------- database mode: one-time tidy (asks first) ----------
 * Removes the tabs the website now shows, closes the old Campus Lead portal files (they go to Drive's trash,
 * shared with nobody), and takes the pictures, "how to use" panels and tab colours off the tabs that stay. */
var DB_TIDY_KEY = 'DB_TIDY';
function tidyToDatabase_(ss) {
  const props = docProps_();
  if (props.getProperty(DB_TIDY_KEY) === '2') return 'done';
  const gone = removedTabs_().filter(n => ss.getSheetByName(n));
  let store = {}; try { store = JSON.parse(props.getProperty('PORTALS') || '{}') || {}; } catch (e) { store = {}; }
  const portals = Object.keys(store.map || {}).filter(k => store.map[k] && store.map[k].id);
  if (gone.length || portals.length) {
    let ui = null; try { ui = SpreadsheetApp.getUi(); } catch (e) { ui = null; }
    if (!ui) return 'later';
    const ans = ui.alert('Make the sheet a plain database?',
      [gone.length ? 'Remove these tabs (they\'re on the website now):\n• ' + gone.join('\n• ') : '',
       portals.length ? `Close ${portals.length} Campus Lead portal file${portals.length === 1 ? '' : 's'} (the website replaces them).` : '']
        .filter(Boolean).join('\n\n'), ui.ButtonSet.YES_NO);
    if (ans !== ui.Button.YES) return 'skipped';
  }
  try {                                                                       // keep the term start date
    if (!(getSettings_().termStart instanceof Date)) {
      const d = ss.getSheetByName(TAB.dashboard), v = d && d.getRange('B3').getValue();
      if (v instanceof Date) ss.getRange(settingCell_('termStart')).setValue(v);
    }
  } catch (e) { Logger.log('Term start: ' + e.message); }
  portals.forEach(k => {
    const p = store.map[k];
    try { const f = DriveApp.getFileById(p.id); if (p.email) { try { f.removeViewer(p.email); } catch (e) { /* already gone */ } } f.setTrashed(true); }
    catch (e) { Logger.log('Portal ' + k + ': ' + e.message); }
  });
  props.deleteProperty('PORTALS'); props.deleteProperty('PORTALS_ON');
  removeTriggers_(['refreshPortalsAuto']);
  gone.forEach(n => { const sh = ss.getSheetByName(n); if (sh && ss.getSheets().length > 1) ss.deleteSheet(sh); });
  ss.getSheets().forEach(sh => plainTab_(sh));
  props.setProperty(DB_TIDY_KEY, '2');
  return 'done';
}

// One tab back to a plain table: no pictures, no "how to use" panel, normal header row, no tab colour
function plainTab_(sh) {
  try { sh.getImages().forEach(i => i.remove()); } catch (e) { /* none */ }
  try {
    const t = typeof tutorials_ === 'function' ? tutorials_()[sh.getName()] : null;
    if (t) {
      const col = tutorialCol_(sh, t);
      if (col <= sh.getMaxColumns() && String(sh.getRange(t.row, col).getValue()).indexOf(TUTORIAL_MARK) === 0) {   // only the panel itself
        const rows = Math.min(60, sh.getMaxRows() - t.row + 1);
        sh.getRange(t.row, col, rows, 1).breakApart().clear();
        sh.setColumnWidth(col, 100);
        if (col > 1) sh.getRange(t.row, col - 1, rows, 1).clearFormat();                           // the thin gap column
      }
    }
  } catch (e) { Logger.log('Panel (' + sh.getName() + '): ' + e.message); }
  try { sh.setRowHeight(1, 32); sh.getRange(1, 1, 1, sh.getMaxColumns()).setVerticalAlignment('middle'); } catch (e) { /* ok */ }
  try { sh.setTabColor(null); } catch (e) { /* ok */ }
}
