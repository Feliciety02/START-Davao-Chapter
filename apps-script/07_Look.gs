/* =====================================================================
 *  START SCHOLAR TRACKER  ·  07_Look
 *  ---------------------------------------------------------------------
 *  Inside: Pictures, banners, table styles, Dashboard, "how to use" panels.
 *  To customise: Picture file names and panel texts.
 *
 *  The code has 9 files + Email_Reminder and Web_App (HTML). Keep these exact names:
 *    00_Config, 01_Region, 02_Menu_Setup, 03_Forms, 04_Emails, 05_Portals, 06_Calendar, 07_Look, 08_Tools
 *  To update: open each file, select all (Ctrl+A), paste the new version, Save.
 *  Then in the Google Sheet: START Tracker › ① Set up / fix the tracker.
 *  You never need to run anything from here: use the START Tracker menu in the Sheet.
 * ===================================================================== */

// The ▶ Run button runs the FIRST function of the open file: this one, so it is always safe.
function about_07_Look() {
  Logger.log('START Tracker' + (typeof START_MODULES === 'string' ? ' ' + START_MODULES : '') + ', file 07_Look. Nothing to run here: open the Google Sheet and use the START Tracker menu.');
}


// ━━━━━━━━━━━━━━━━━━━━  Logos  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  🖼 PICTURES  (logos and banners, taken from the StaRRt IGNITE deck)
 *  The pictures are NOT stored in the code any more (that made this file ~850 KB, easy to break
 *  when pasting, and loaded on every run). They live in your Google Drive, in the folder
 *  "START Tracker pictures" next to this tracker. Add them once with
 *  START Tracker › 🖼 Add the tracker pictures (one time) and pick all the files from the zip.
 *  If a picture is missing, nothing breaks: that banner is simply skipped and 🩺 tells you.
 * ================================================================ */

// key -> file name (the names in the pictures zip)
var PICTURES = {
  logos: 'start-tracker-logos.png',          // the 3 logos side by side (emails)
  form: 'start-tracker-form-v3.jpg',         // Campus Lead check-in form header ("CAMPUS LEAD CHECK-IN")
  mform: 'start-tracker-member-form.jpg',    // member check-in form header
  start: 'start-tracker-start.jpg', dashboard: 'start-tracker-dashboard.jpg', portal: 'start-tracker-portal-v2.jpg',
  'tab-schools': 'start-tracker-tab-schools.jpg',   'tab-scholars': 'start-tracker-tab-scholars.jpg',
  'tab-focal': 'start-tracker-tab-focal.jpg',       'tab-outreach': 'start-tracker-tab-outreach.jpg',
  'tab-checkins': 'start-tracker-tab-checkins-v2.jpg', 'tab-log': 'start-tracker-tab-log.jpg',
  'tab-settings': 'start-tracker-tab-settings.jpg', 'tab-links': 'start-tracker-tab-links.jpg',
  'tab-events': 'start-tracker-tab-events.jpg',     'tab-calview': 'start-tracker-tab-calview.jpg',
  'tab-mcheckins': 'start-tracker-tab-mcheckins-v2.jpg',
  guide: 'start-tracker-guide.jpg',
  'portal-wide': 'start-tracker-portal-wide.jpg',   // the Campus Lead portal's header (same 10 : 1 shape as the tabs)
  // email decorations (shown on phones too: the Gmail apps drop CSS gradients, but show pictures)
  'email-bg': 'start-tracker-email-bg.jpg', 'email-hero': 'start-tracker-email-hero.jpg', 'email-bar': 'start-tracker-email-bar.png',
  'email-sparkle': 'start-tracker-email-sparkle.gif', 'email-star': 'start-tracker-email-star.gif',          // "HOW TO USE" picture above each tab's steps panel
};
var PICTURE_FOLDER_NAME = 'START Tracker pictures';
var PICTURE_CACHE_ = {};

// The pictures folder (made next to the tracker the first time). Private: only you.
function pictureFolder_(create) {
  const props = docProps_();
  const id = props.getProperty('PICTURE_FOLDER');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { /* gone */ } }
  let found = null;                                                     // not linked yet: a folder with the usual name (or with _ for spaces)
  try { [PICTURE_FOLDER_NAME, PICTURE_FOLDER_NAME.replace(/ /g, '_')].forEach(n => { if (found) return; const it = DriveApp.getFoldersByName(n); while (!found && it.hasNext()) { const f = it.next(); if (!f.isTrashed()) found = f; } }); }
  catch (e) { /* Drive search not available */ }
  if (found) { props.setProperty('PICTURE_FOLDER', found.getId()); return found; }
  if (!create) return null;
  const parents = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const f = parent.createFolder(PICTURE_FOLDER_NAME);
  props.setProperty('PICTURE_FOLDER', f.getId());
  return f;
}

// A folder and every folder inside it (3 levels), so pictures can be sorted into subfolders
function withSubfolders_(folder, depth) {
  const out = [folder];
  if ((depth || 0) >= 3) return out;
  try { const it = folder.getFolders(); while (it.hasNext()) { const d = it.next(); if (!d.isTrashed()) withSubfolders_(d, (depth || 0) + 1).forEach(x => out.push(x)); } }
  catch (e) { /* no access to a subfolder */ }
  return out;
}

// The Drive file for a picture: the pictures folder first, then anywhere in your Drive (by file name).
// Until the newer pictures are added, the older ones are used (so nothing disappears)
var PICTURE_FALLBACK = {
  form: ['start-tracker-form-v2.jpg', 'start-tracker-form.jpg'], portal: ['start-tracker-portal.jpg'],
  'tab-checkins': ['start-tracker-tab-checkins.jpg'], 'tab-mcheckins': ['start-tracker-tab-mcheckins.jpg'],
};

function pictureFile_(key, older) {
  const name = older ? (PICTURE_FALLBACK[key] || [])[older - 1] : PICTURES[key];
  if (!name) return null;
  const props = docProps_();
  let ids = {};
  try { ids = JSON.parse(props.getProperty('PICTURE_IDS') || '{}'); } catch (e) { ids = {}; }
  if (ids[key]) { try { const f = DriveApp.getFileById(ids[key]); if (!f.isTrashed() && f.getName() === name) return f; } catch (e) { /* gone */ } }   // a renamed (newer) picture is looked up again
  let found = null;
  const folder = pictureFolder_(false);
  const look = it => { while (!found && it.hasNext()) { const f = it.next(); if (!f.isTrashed()) found = f; } };
  if (folder) withSubfolders_(folder).forEach(d => { if (!found) look(d.getFilesByName(name)); });   // subfolders too ("Spreadsheet Header" …)
  if (!found && folder) {                                                // "start-tracker-logos (1).png" style copies inside the folder
    const dot = name.lastIndexOf('.'), base = name.slice(0, dot).toLowerCase(), ext = name.slice(dot).toLowerCase();
    withSubfolders_(folder).forEach(d => {
      const it = d.getFiles();
      while (!found && it.hasNext()) {
        const f = it.next(), n = f.getName().toLowerCase();
        if (!f.isTrashed() && n.endsWith(ext) && n.slice(0, -ext.length).replace(/\s*\(\d+\)$/, '') === base) found = f;
      }
    });
  }
  if (!found) look(DriveApp.getFilesByName(name));
  if (!found) {                                                       // Drive renamed a re-upload, e.g. "start-tracker-email-bg (1).jpg"
    const dot = name.lastIndexOf('.'), base = name.slice(0, dot), ext = name.slice(dot);
    const same = new RegExp('^' + base.replace(/[-.]/g, '\\$&') + '( \\(\\d+\\))?\\' + ext + '$', 'i');
    try {
      const it = DriveApp.searchFiles("title contains '" + base + "' and trashed = false");
      while (!found && it.hasNext()) { const f = it.next(); if (same.test(f.getName())) found = f; }
    } catch (e) { /* search not available */ }
  }
  if (!found && importPicturesZip_()) return pictureFile_(key, older);    // the zip itself was uploaded to Drive
  if (!found && PICTURE_FALLBACK[key] && (older || 0) < PICTURE_FALLBACK[key].length) return pictureFile_(key, (older || 0) + 1);
  if (found && !older) { ids[key] = found.getId(); props.setProperty('PICTURE_IDS', JSON.stringify(ids)); }
  return found;
}

// If you uploaded START_Tracker_pictures.zip to Drive WITHOUT unzipping it, the tracker unzips it once
// into the pictures folder. Tried at most once per run. true = something new was added.
var ZIP_TRIED_ = false;
function importPicturesZip_() {
  if (ZIP_TRIED_) return false;
  ZIP_TRIED_ = true;
  try {
    const it = DriveApp.searchFiles("(title contains 'START_Tracker_pictures' or title contains 'START_email_pictures') and trashed = false");
    while (it.hasNext()) {
      const z = it.next();
      if (!/zip/i.test(z.getMimeType()) && !/\.zip$/i.test(z.getName())) continue;
      let added = 0;
      Utilities.unzip(z.getBlob().setContentType('application/zip')).forEach(b => {
        const name = String(b.getName()).split('/').pop();
        if (Object.values(PICTURES).indexOf(name) === -1) return;
        savePicture(name, /\.png$/i.test(name) ? 'image/png' : /\.gif$/i.test(name) ? 'image/gif' : 'image/jpeg', Utilities.base64Encode(b.getBytes()));
        added++;
      });
      if (added) return true;
    }
  } catch (e) { Logger.log('Pictures zip not read: ' + e.message); }
  return false;
}

// The picture as a blob, or null when it has not been added yet
function picture_(key) {
  if (PICTURE_CACHE_[key] !== undefined) return PICTURE_CACHE_[key];
  let blob = null;
  try { const f = pictureFile_(key); if (f) blob = f.getBlob().setName(f.getName()); }   // the real file name (an older picture keeps its own)
  catch (e) { Logger.log('Picture ' + key + ': ' + e.message); }
  return (PICTURE_CACHE_[key] = blob);
}

function missingPictures_() {
  return Object.keys(PICTURES).filter(k => { try { const f = pictureFile_(k); return !f || f.getName().toLowerCase().replace(/\s*\(\d+\)(\.\w+)$/, '$1') !== PICTURES[k].toLowerCase(); } catch (e) { return true; } });
}

/* ---- 🖼 one-time upload (menu) ---- */

function showPictureUpload() {
  const names = JSON.stringify(Object.values(PICTURES));
  const html = `<p class="lead">Add the tracker pictures</p>
    <ol class="steps"><li>Unzip <b>START_Tracker_pictures.zip</b></li><li>Pick all ${Object.keys(PICTURES).length} pictures below (Ctrl+A)</li></ol>
    <input type="file" id="f" accept="image/png,image/jpeg,image/gif" multiple>
    <div id="s" class="msg" style="white-space:pre-line"></div>
    <script>
      const want = ${names};
      document.getElementById('f').onchange = async e => {
        const files = [...e.target.files].filter(f => want.includes(f.name));
        const s = document.getElementById('s');
        if (!files.length) { s.textContent = 'These are not tracker pictures.'; return; }
        let done = 0;
        for (const f of files) {
          s.textContent = 'Saving ' + (done + 1) + ' of ' + files.length + ': ' + f.name + '…';
          const data = await new Promise(r => { const fr = new FileReader(); fr.onload = () => r(String(fr.result).split(',')[1]); fr.readAsDataURL(f); });
          await new Promise((ok, bad) => google.script.run.withSuccessHandler(ok).withFailureHandler(bad).savePicture(f.name, f.type, data))
            .catch(err => { s.textContent = 'Could not save ' + f.name + ': ' + err.message; throw err; });
          done++;
        }
        const left = want.length - (await new Promise(r => google.script.run.withSuccessHandler(r).pictureCount()));
        s.textContent = '✅ Saved ' + done + '.' + (left ? ' ' + left + ' still missing.' : ' All in. Now click ① Set up / fix the tracker.');
      };
    </script>`;
  dlg_('🖼 Pictures', html, 440, 300);
}

// Called by the upload dialog. Only the 13 known file names are accepted.
function savePicture(name, type, base64) {
  const key = Object.keys(PICTURES).find(k => PICTURES[k] === name);
  if (!key) throw new Error('Not a tracker picture: ' + name);
  if (['image/png', 'image/jpeg', 'image/gif'].indexOf(type) === -1) throw new Error('Only PNG, JPG or GIF pictures');
  const folder = pictureFolder_(true);
  const old = folder.getFilesByName(name);
  while (old.hasNext()) old.next().setTrashed(true);                     // replace, never duplicate
  const f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(base64), type, name));
  const props = docProps_();
  let ids = {};
  try { ids = JSON.parse(props.getProperty('PICTURE_IDS') || '{}'); } catch (e) { ids = {}; }
  ids[key] = f.getId();
  props.setProperty('PICTURE_IDS', JSON.stringify(ids));
  if (key === 'logos') props.deleteProperty('EMAIL_LOGO_ID');           // the email logo is made again from the new file
  if (/^email-/.test(key)) props.deleteProperty('EMAIL_PIC_' + key);     // an email decoration: published again
  PICTURE_CACHE_ = {}; LOGO_SRC_CACHE_ = null; EMAIL_PICS_CACHE_ = null;
  return true;
}

function pictureCount() { return Object.keys(PICTURES).length - missingPictures_().length; }

/* ---- 📧 email logo ---- */
// Gmail often hides pictures attached inside an email ("cid:"), so the tracker keeps ONE public copy
// of the logos picture in your Drive (LOGOS ONLY, no personal data) and every email links to it.
// If your Google account does not allow public files, emails fall back to the attached picture.
// The approved public file is this one picture only: never share anything else this way.
var LOGO_STRIP_CID = 'cid:logoStrip';
var LOGO_STRIP_SIZE = { w: 125, h: 40 };   // shown size in the email (the picture is 3x for sharp screens)
var LOGO_STRIP_NAME = 'START email logos (public picture, no personal data)';
var LOGO_SRC_CACHE_ = null;

// The picture address used in emails: the public Drive copy, 'cid:logoStrip' (attached), or '' (no picture yet).
function emailLogoSrc_() {
  if (LOGO_SRC_CACHE_ !== null) return LOGO_SRC_CACHE_;
  const link = id => 'https://lh3.googleusercontent.com/d/' + id + '=w375';
  try {
    const props = docProps_();
    const id = props.getProperty('EMAIL_LOGO_ID');
    if (id) {
      try {
        const old = DriveApp.getFileById(id);
        if (!old.isTrashed() && old.getSharingAccess() === DriveApp.Access.ANYONE_WITH_LINK) return (LOGO_SRC_CACHE_ = link(id));
        if (!old.isTrashed()) old.setTrashed(true);   // replaced below
      } catch (e) { /* deleted: make a new one */ }
    }
    const blob = picture_('logos');
    if (!blob) return (LOGO_SRC_CACHE_ = '');          // pictures not added yet: emails show the names as text
    const blocked = Number(props.getProperty('EMAIL_LOGO_BLOCKED') || 0);
    if (Date.now() - blocked < 24 * 3600 * 1000) throw new Error('public files not allowed (tries again tomorrow)');
    const f = DriveApp.createFile(blob.copyBlob().setName('start-logos.png')).setName(LOGO_STRIP_NAME)
      .setDescription('Used by the START Tracker emails. Contains only the DOST, START and DOST-SEI logos.');
    try { f.moveTo(pictureFolder_(true)); } catch (e) { /* stays in My Drive */ }
    try {
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);   // logos only, never scholar data
    } catch (e) {
      f.setTrashed(true);                                                         // no leftover private copies
      props.setProperty('EMAIL_LOGO_BLOCKED', String(Date.now()));
      throw e;
    }
    props.deleteProperty('EMAIL_LOGO_BLOCKED');
    props.setProperty('EMAIL_LOGO_ID', f.getId());
    return (LOGO_SRC_CACHE_ = link(f.getId()));
  } catch (e) {
    Logger.log('Email logo: using the attached picture (' + e.message + ')');
    return (LOGO_SRC_CACHE_ = picture_('logos') ? LOGO_STRIP_CID : '');
  }
}

// For MailApp's inlineImages: only needed when the fallback (attached picture) is in use.
function logoInlineImages_() {
  return emailLogoSrc_() === LOGO_STRIP_CID ? { logoStrip: picture_('logos') } : {};
}

/* ---- banners ---- */

function formBannerBlob_() { return picture_('form'); }

// Older copies of 03_StartHere, 08_FocalPortal and 13_Dashboard still pass these names:
// they now simply mean "the picture with this key", so ① keeps working while files are being replaced.
var START_BANNER_B64 = 'start', DASH_BANNER_B64 = 'dashboard', PORTAL_BANNER_B64 = 'portal';

// Puts a banner in the top-left of a sheet and makes row 1 tall enough for it.
// Removes older banners first, so re-running ① never stacks images. Missing picture: keeps what is there.
function placeBanner_(sh, key, widthPx, ratio) {
  if (DB_MODE) return;                                   // database mode: not used (see 00_Config)
  try {
    const blob = PICTURES[key] ? picture_(key) : null;               // an unknown key (very old file) is skipped
    if (!blob) { Logger.log('Banner skipped (picture "' + key + '" not added yet)'); return; }
    sh.getImages().filter(i => /^START (banner|logo)/.test(String(i.getAltTextTitle()))).forEach(i => i.remove());
    const img = sh.insertImage(blob, 1, 1, 0, 0);
    const h = Math.round(widthPx / (ratio || 4));      // 4:1 banners; the Dashboard one is 6:1
    img.setWidth(widthPx).setHeight(h).setAltTextTitle('START banner').setAltTextDescription('START – DOST-SEI');
    sh.setRowHeight(1, h + 6);
    sh.getRange('A1').setValue('');
  } catch (e) { Logger.log('Banner not placed: ' + e.message); }
}

// Spreadsheet theme: Poppins + deck colours (used by charts, links and new formatting)
function applyTheme_(ss) {
  try {
    const theme = ss.getSpreadsheetTheme() || ss.resetSpreadsheetTheme();
    const c = hex => SpreadsheetApp.newColor().setRgbColor(hex).build();
    const T = SpreadsheetApp.ThemeColorType;
    theme.setFontFamily(THEME.font);
    [[T.TEXT, THEME.navy], [T.BACKGROUND, '#FFFFFF'], [T.ACCENT1, THEME.blue], [T.ACCENT2, THEME.sun],
     [T.ACCENT3, THEME.navy], [T.ACCENT4, '#7CC8FF'], [T.ACCENT5, '#FFE866'], [T.ACCENT6, THEME.slate],
     [T.HYPERLINK, THEME.blue]].forEach(([k, v]) => theme.setConcreteColor(k, c(v)));
  } catch (e) { Logger.log('Theme not applied: ' + e.message); }
}

/* ---------- email decorations: published like the logo (pictures only, never scholar data) ---------- */
var EMAIL_PICS_CACHE_ = null;

// { bg, hero, bar, sparkle, star } public links ('' = not added yet / not allowed: the email uses plain colours)
function emailPics_() {
  if (EMAIL_PICS_CACHE_) return EMAIL_PICS_CACHE_;
  const out = { bg: '', hero: '', bar: '', sparkle: '', star: '' };
  const map = { bg: 'email-bg', hero: 'email-hero', bar: 'email-bar', sparkle: 'email-sparkle', star: 'email-star' };
  const props = docProps_();
  if (Date.now() - Number(props.getProperty('EMAIL_LOGO_BLOCKED') || 0) < 24 * 3600 * 1000) return (EMAIL_PICS_CACHE_ = out);   // public files not allowed
  Object.keys(map).forEach(k => {
    const key = map[k], gif = /\.gif$/i.test(PICTURES[key]);
    const link = id => 'https://lh3.googleusercontent.com/d/' + id + (gif ? '=s0' : '=w1200');   // =s0 keeps the GIF animated
    try {
      const id = props.getProperty('EMAIL_PIC_' + key);
      if (id) {
        try { const f = DriveApp.getFileById(id); if (!f.isTrashed() && f.getSharingAccess() === DriveApp.Access.ANYONE_WITH_LINK) { out[k] = link(id); return; } }
        catch (e) { /* gone: published again below */ }
      }
      const blob = picture_(key);
      if (!blob) return;                                                 // not added yet
      const f = DriveApp.createFile(blob.copyBlob().setName(PICTURES[key])).setName('START email · ' + PICTURES[key])
        .setDescription('Decoration used by the START Tracker emails. Contains no personal data.');
      try { f.moveTo(pictureFolder_(true)); } catch (e) { /* stays in My Drive */ }
      try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
      catch (e) { f.setTrashed(true); props.setProperty('EMAIL_LOGO_BLOCKED', String(Date.now())); return; }
      props.setProperty('EMAIL_PIC_' + key, f.getId());
      out[k] = link(f.getId());
    } catch (e) { Logger.log('Email picture ' + key + ': ' + e.message); }
  });
  return (EMAIL_PICS_CACHE_ = out);
}


// ━━━━━━━━━━━━━━━━━━━━  Style  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  TABLE & TEXT POLISH  (StaRRt IGNITE look)
 *  Runs at the end of ① Set up / fix the tracker. Safe to re-run: it rebuilds its own rules.
 *  - navy Poppins text, light-blue table lines instead of grey gridlines, roomier rows
 *  - "chip" colours for Status, Year level, Type, Inactive, and check-in answers
 *  - Days since check-in as a green → yellow → red heat scale
 * ================================================================ */

var CHIP = {
  green:  { bg: '#D7F2DF', fg: '#14532D' },
  blue:   { bg: '#D6ECFF', fg: '#0B4A80' },
  yellow: { bg: '#FFF3B0', fg: '#5C4A00' },
  red:    { bg: '#FADCD9', fg: '#8A1C14' },
  grey:   { bg: '#E7EBF1', fg: '#3F4A5C' },
  navy:   { bg: '#12203A', fg: '#FFFFFF' },
};
var TABLE_LINE = '#CFE0F3';

function polishTables_(ss) {
  try { polishScholars_(ss); } catch (e) { Logger.log('Scholars polish: ' + e.message); }
  try { polishSchools_(ss); }  catch (e) { Logger.log('Schools polish: ' + e.message); }
  try { polishCheckins_(ss); } catch (e) { Logger.log('Checkins polish: ' + e.message); }
  try { buildLog_(ss); }       catch (e) { Logger.log('Reminder log: ' + e.message); }
  try { applyColumnKinds_(ss); } catch (e) { Logger.log('Column kinds: ' + e.message); }   // headings last, so nothing repaints them
}

/* ================================================================
 *  ✏️ EDITABLE vs 🔒 AUTOMATIC — one visual rule for every table
 *   ✏️ blue heading + white cells  = you type here
 *   🔒 navy heading + grey cells   = fills in by itself (protected with a warning)
 *   yellow cell                    = a setting
 * ================================================================ */

var KIND = {
  edit: { icon: '✏️ ', head: '#0A6ED1', note: 'You can type here.' },
  auto: { icon: '🔒 ', head: '#12203A', note: 'Fills in by itself. Please don\'t type here.' },
  autoBg: '#EEF2F7',
};

// THE one place that names every column of Schools and Scholars: [column, kind, heading, note (optional)].
// Build code never writes these headings itself, so a heading can't be spelled two ways.
function columnSpecs_() {
  const infoFit = { 'Official name': true, 'Website': 'link', 'Public email': 'link', 'DOST-SEI basis': true };
  const info = SCHOOL_INFO_HEADERS.map((h, i) => [9 + i, 'edit', h,
    i === 0 ? 'School details from the ready-made school list (columns I–O). Folded away: click + above the column to open.' : '',
    infoFit[h]]);
  return {
    [TAB.schools]: [
      [1, 'edit', 'School', 'One school per row. Use the same short name everywhere (it appears in the form).\n' +
                            'School cell: green = checked in within 7 days, red = quiet for too long (see Settings).\n' +
                            '✏️ blue headings = you type here.  🔒 navy headings = fills in by itself.'],
      [2, 'edit', 'Province'],
      [3, 'auto', 'Active scholars', 'Scholars from this school in the Scholars tab, not counting Graduated or Dropped.'],
      [4, 'edit', 'START Campus Lead', 'Pick from the Scholars list. A non-scholar can still be typed in.'],
      [5, 'auto', 'Contact (email)', 'The START Campus Lead\'s email, read from Scholars › Contact (email). ' +
                                     'Change it there: emails are typed in one place only. Red = the Lead has no email in Scholars.', 'link'],
      [6, 'edit', 'Inactive last term (Y/N)'],
      [7, 'auto', 'Last check-in'], [8, 'auto', 'Days since check-in'],
    ].concat(info).concat([
      [LINK_COL, 'auto', 'Form link for this school', 'Send each START Campus Lead their school\'s link: the school is already selected.', 'link'],
      [17, 'auto', 'Last contacted', 'From the Outreach log tab.'],
    ]),
    [TAB.scholars]: [[1, 'edit', 'Name', 'Paste the ONSSE masterlist starting at A2.\nOrange rows = same name twice. ' +
                                         'Yellow rows = START Campus Leads.\nUse the filter icons in this row to sort or filter.'], [2, 'edit', 'School', 'Spell it exactly as in the Schools tab.'], [3, 'edit', 'Program'],
                     [4, 'edit', 'Year awarded (DOST)', 'The year the DOST scholarship was granted (batch), e.g. 2023.'],
                     [5, 'edit', 'Contact (email)', 'THE place to type emails. Schools, reminders, portals and announcements all use it.', 'link'], [6, 'edit', 'Status'], [7, 'edit', 'Year level'],
                     [8, 'edit', 'Campus / Branch'], [9, 'edit', 'Hometown (City/Municipality, Province)', '', true],
                     [10, 'edit', 'Role', 'Member or ⭐ START Campus Lead. One Campus Lead per school (their school in column B): ' +
                                          'picking a new one makes the old one a Member. Same as Schools › START Campus Lead (D).'],
                     [11, 'edit', 'Other emails (sign-in only)', 'Optional. Other emails this scholar may sign in to the START website with ' +
                                          '(separate several with a comma). START never emails these: every email goes to Contact (email) in column E.'],
                     [12, 'edit', 'Scholarship (DOST)', 'RA 7687, Merit, RA 10612 or JLSS. Filled from the join request when you approve one.'],
                     [13, 'auto', 'Added from', 'Blank = your ONSSE masterlist. "Join request · date" = approved on the START website.']],
  };
}

// Strips an old ✏️/🔒 so headings never get the icon twice
function plainHeading_(v) { return String(v || '').replace(/^\s*(✏️|✏|🔒)\s*/u, '').trim(); }
function isMarked_(v) { return /^(✏️|🔒)/u.test(String(v || '')); }

// Paints headings (text, colour, note) and grey automatic cells. Consecutive columns of the same kind
// are painted in one go. keepText: only colour the headings (Weekly Check-in: the form needs its own titles).
function markColumns_(sh, spec, rows, keepText) {
  const runs = [];
  spec.forEach(sp => {
    const last = runs[runs.length - 1];
    if (last && last.kind === sp[1] && last.to === sp[0] - 1) { last.to = sp[0]; last.items.push(sp); }
    else runs.push({ kind: sp[1], from: sp[0], to: sp[0], items: [sp] });
  });
  runs.forEach(run => {
    const k = KIND[run.kind], w = run.to - run.from + 1;
    const head = sh.getRange(1, run.from, 1, w);
    if (!keepText) {
      const old = head.getValues()[0];
      head.setValues([run.items.map((sp, i) => k.icon + (sp[2] || plainHeading_(old[i])))]);
    }
    header_(head, k.head);
    head.setNotes([run.items.map(sp => sp[3] || (run.kind === 'auto' ? k.note : ''))]);
    if (run.kind === 'auto' && rows > 0) sh.getRange(2, run.from, rows, w).setBackground(KIND.autoBg);
  });
  fitColumns_(sh, spec);
}

/* ---------- flexible tables: every column fits its content, rows grow to fit ---------- */
// spec items: [column, kind, heading, note, fit]
//   fit = (nothing) short values: the column is as wide as its longest entry (up to FIT.max)
//   fit = true      long text (concerns, notes): wider column, text wraps onto more lines
//   fit = 'link'    links and emails: a narrow column, the link wraps instead of stretching the table
var FIT = { min: 64, max: 300, wide: 420, link: 240, pad: 22 };

// Rough pixel width of a heading in bold Poppins 10 (emoji/arrows count wider)
function textPx_(s) {
  let w = 0;
  for (const ch of String(s || '')) w += /[\u2190-\u2BFF\u{1F000}-\u{1FFFF}]/u.test(ch) ? 20 : /[A-Z@MWmw]/.test(ch) ? 9.2 : 7.4;
  return w;
}

function fitColumns_(sh, spec) {
  if (!spec.length) return;
  const cols = spec.map(s => s[0]).sort((a, b) => a - b), last = cols[cols.length - 1];
  const heads = sh.getRange(1, 1, 1, last).getDisplayValues()[0];
  // one auto-resize call per run of neighbouring columns
  let start = cols[0], prev = cols[0];
  const flush = () => { sh.getRange(1, start, 1, prev - start + 1).setWrap(false); sh.autoResizeColumns(start, prev - start + 1); };
  cols.slice(1).forEach(c => { if (c !== prev + 1) { flush(); start = c; } prev = c; });
  flush();
  const body = c => sh.getRange(2, c, Math.max(sh.getMaxRows() - 1, 1), 1);
  spec.forEach(sp => {
    const c = sp[0], fit = sp[4];
    const cap = fit === 'link' ? FIT.link : fit ? FIT.wide : FIT.max;
    const head = textPx_(heads[c - 1]) + FIT.pad + 10;                        // the heading always fits on one line
    const content = sh.getColumnWidth(c) + FIT.pad;                           // longest entry (after auto-resize)
    sh.setColumnWidth(c, Math.round(Math.max(Math.min(content, cap), head, FIT.min)));
    body(c).setWrapStrategy(fit ? SpreadsheetApp.WrapStrategy.WRAP : SpreadsheetApp.WrapStrategy.CLIP);
  });
  fitRows_(sh);
}

// Rows grow or shrink to fit what is in them (no fixed heights), header row stays 40 px
function fitRows_(sh, from, n) {
  from = from || 2;
  n = n || Math.max(sh.getMaxRows() - from + 1, 1);
  try { sh.autoResizeRows(from, n); } catch (e) { Logger.log('Rows not resized: ' + e.message); }
}

function applyColumnKinds_(ss) {
  const specs = columnSpecs_();
  Object.keys(specs).forEach(tab => {
    const sh = ss.getSheetByName(tab);
    if (!sh) return;
    const rows = tab === TAB.schools ? CONFIG.maxSchools : Math.min(sh.getMaxRows() - 1, 500);
    markColumns_(sh, specs[tab], rows, false);
  });
  // Weekly Check-in: every column is automatic (form answers + the extra columns). Form titles must keep their wording.
  const ck = ss.getSheetByName(TAB.checkins);
  if (ck && ck.getLastColumn() > 1) {
    const c = ckCols_(ck);
    markColumns_(ck, Array.from({ length: c.width + CK_EXTRA.length }, (_, i) => [i + 1, 'auto', '',
      i === 0 ? 'Answers from the check-in form. They arrive by themselves; please don\'t edit here.'
              : i >= c.width ? 'Worked out by the tracker from the answer and the Schools tab.' : '',
      i === c.concern]), 0, true);
    placeTabBanner_(ck, c.width + CK_EXTRA.length);
  }
  [[TAB.schools, 8], [TAB.scholars, SCHOLAR_COLS + 1]].forEach(([t, n]) => { const s = ss.getSheetByName(t); if (s) placeTabBanner_(s, n); });
}

/* ---------- helpers ---------- */

function chipRule_(range, text, chip, exact) {
  const b = SpreadsheetApp.newConditionalFormatRule();
  (exact === false ? b.whenTextStartsWith(text) : b.whenTextEqualTo(text));
  return b.setBackground(chip.bg).setFontColor(chip.fg).setBold(true).setRanges([range]).build();
}

// Body look shared by every table. Headings are painted by markColumns_ only.
function tableLook_(sh, rows, cols, rowHeight, tabColor) {
  sh.getRange(1, 1, rows + 1, cols).setFontColor(THEME.navy).setFontFamily(THEME.font).setFontSize(10)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, TABLE_LINE, SpreadsheetApp.BorderStyle.SOLID);
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 40);                       // body rows are sized to their content by fitRows_
  sh.setHiddenGridlines(true);
  if (tabColor) sh.setTabColor(tabColor);
}

/* ---------- Scholars ---------- */

function polishScholars_(ss) {
  const sh = ss.getSheetByName(TAB.scholars);
  if (!sh) return;
  const rows = Math.min(sh.getMaxRows() - 1, 500);
  tableLook_(sh, rows, SCHOLAR_COLS + 1, 28, THEME.tabs.scholars);   // A:I + Role
  band_(sh.getRange(2, 1, rows, SCHOLAR_COLS));
  sh.setFrozenColumns(1);
  sh.getRange(2, 1, rows, 1).setFontWeight('bold');                       // names
  sh.getRange(2, 2, rows, 1).setFontColor(THEME.blue).setFontWeight('bold').setHorizontalAlignment('center'); // school
  sh.getRange(2, 4, rows, 1).setHorizontalAlignment('center');            // year awarded
  sh.getRange(2, 6, rows, 2).setHorizontalAlignment('center');            // status, year level
  sh.getRange(2, 8, rows, 2).setFontColor(THEME.muted);                   // campus, hometown

  const F = sh.getRange(2, 6, rows, 1), G = sh.getRange(2, 7, rows, 1);
  const J = sh.getRange(2, 10, rows, 1), row = sh.getRange(2, 1, rows, 8);   // A:H (the duplicate-name rule uses A:I)
  bodyRows_(J).setHorizontalAlignment('left');
  const L = CONFIG.maxSchools + 1;
  // START Campus Leads: whole row turns soft yellow and the Role cell becomes a navy chip
  const isFocal = `=AND($A2<>"",COUNTIF(INDIRECT("${TAB.schools}!D2:D${L}"),$A2)>0)`;
  const focalRules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('⭐').setBackground(THEME.navy)
      .setFontColor('#FFFFFF').setBold(true).setRanges([J]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(isFocal).setBackground('#FFF3B0')
      .setBold(true).setRanges([row]).build(),
  ];
  const rules = sh.getConditionalFormatRules().filter(r => !isOurChip_(r, [F, G]) && !isOurChip_(r, [J]) && !isOurChip_(r, [row]));
  sh.setConditionalFormatRules(focalRules.concat(rules).concat([
    chipRule_(F, 'Active', CHIP.green), chipRule_(F, 'On leave', CHIP.yellow),
    chipRule_(F, 'Graduated', CHIP.blue), chipRule_(F, 'Dropped', CHIP.grey),
    chipRule_(G, '1st Year', CHIP.yellow), chipRule_(G, '2nd Year', CHIP.blue),
    chipRule_(G, '3rd Year', CHIP.green), chipRule_(G, '4th Year', CHIP.navy), chipRule_(G, '5th Year', CHIP.navy),
  ]));
}

// True if a rule only touches the given chip columns (so re-running doesn't stack duplicates)
function isOurChip_(rule, ranges) {
  const a1 = ranges.map(r => r.getA1Notation());
  return rule.getRanges().every(r => a1.indexOf(r.getA1Notation()) !== -1);
}

/* ---------- Schools ---------- */

function polishSchools_(ss) {
  const sh = ss.getSheetByName(TAB.schools);
  if (!sh) return;
  const n = CONFIG.maxSchools;
  tableLook_(sh, n, 17, 30, THEME.tabs.schools);                   // A:Q
  sh.setFrozenColumns(1);
  sh.getRange(2, 9, n, LINK_COL - 8).setFontColor(THEME.muted);
  sh.getRange(1, 1, n + 1, 8).setBorder(null, null, null, true, null, null, THEME.slate, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(2, 1, n, 1).setFontWeight('bold').setFontColor(THEME.navy);
  sh.getRange(2, 3, n, 1).setFontWeight('bold').setFontColor(THEME.blue);
  sh.getRange(2, 11, n, 1).setHorizontalAlignment('center');

  const H = sh.getRange(2, 8, n, 1), F = sh.getRange(2, 6, n, 1), K = sh.getRange(2, 11, n, 1);
  const heat = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#BFE9CC', SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMidpointWithValue('#FFF3B0', SpreadsheetApp.InterpolationType.NUMBER, '7')
    .setGradientMaxpointWithValue('#F7B4AE', SpreadsheetApp.InterpolationType.NUMBER, String(getSettings_().quietDays))
    .setRanges([H]).build();
  const own = [H, F, K];
  const keep = sh.getConditionalFormatRules().filter(r => !isOurChip_(r, own));
  // Specific chips first so they show inside red/green rows
  sh.setConditionalFormatRules([
    heat,
    chipRule_(F, 'Y', CHIP.red), chipRule_(F, 'N', CHIP.grey),
    chipRule_(K, 'SUC', CHIP.blue, false), chipRule_(K, 'Private', CHIP.yellow),
  ].concat(keep));
}

/* ---------- Weekly Check-in ---------- */

// Extra columns the tracker adds after the form's answers (headings are part of each formula, in row 1,
// so new answers are covered automatically). The form never writes into them.
var CK_EXTRA = ['Week of', 'On time?', 'START Campus Lead on record', 'Active scholars'];

function polishCheckins_(ss) {
  const sh = ss.getSheetByName(TAB.checkins);
  if (!sh || sh.getLastColumn() < 2) return;
  const c = ckCols_(sh);
  hideOldQuestionCols_(sh, c);
  const X = CK_EXTRA.length, W = c.width + X;
  if (sh.getMaxColumns() < W) sh.insertColumnsAfter(sh.getMaxColumns(), W - sh.getMaxColumns());
  const rows = Math.max(sh.getMaxRows() - 1, 1), body = Math.min(rows, 1000);

  // extra columns
  const L = colLetter_(c.school + 1), S = TAB.schools, M = CONFIG.maxSchools + 1;
  const days = '{"' + DAYS.join('","') + '"}';
  // the array formulas need empty cells below them (an older steps panel may have sat here)
  if (sh.getMaxRows() > 1) sh.getRange(2, c.width + 1, sh.getMaxRows() - 1, X).clearContent();
  const weekOf = 'INT(A2:A)-WEEKDAY(A2:A,3)';
  sh.getRange(1, c.width + 1, 1, X).setFormulas([[
    `={"${CK_EXTRA[0]}";ARRAYFORMULA(IF(A2:A="","",${weekOf}))}`,
    `={"${CK_EXTRA[1]}";ARRAYFORMULA(IF(A2:A="","",IF(INT(A2:A)<=${weekOf}+MATCH(${settingCell_('dueDay')},${days},0)-1,"✅ On time","⏰ Late")))}`,
    `={"${CK_EXTRA[2]}";ARRAYFORMULA(IF(${L}2:${L}="","",IFERROR(IF(VLOOKUP(${L}2:${L},${S}!$A$2:$D$${M},4,FALSE)="",` +
      `"— none",VLOOKUP(${L}2:${L},${S}!$A$2:$D$${M},4,FALSE)),"— none")))}`,
    `={"${CK_EXTRA[3]}";ARRAYFORMULA(IF(${L}2:${L}="","",IFERROR(VLOOKUP(${L}2:${L},${S}!$A$2:$C$${M},3,FALSE),"")))}`,
  ]]);

  if (DB_MODE) return; // The database needs the computed columns, not the table styling.

  // look
  tableLook_(sh, body, W, 28, THEME.tabs.checkins);
  band_(sh.getRange(2, 1, rows, W), '#F6F8FB', KIND.autoBg);          // grey bands = automatic
  sh.setFrozenColumns(Math.min(c.school + 1, 2));
  const col = i => sh.getRange(2, i + 1, rows, 1);
  col(0).setNumberFormat('mmm d, yyyy  h:mm AM/PM').setFontColor(THEME.muted);
  col(c.school).setFontWeight('bold').setFontColor(THEME.blue);
  if (c.name >= 0) col(c.name).setFontWeight('bold');
  col(c.width).setNumberFormat('mmm d').setHorizontalAlignment('center');                  // Week of
  col(c.width + 1).setHorizontalAlignment('center');                                         // On time?
  col(c.width + 3).setHorizontalAlignment('center').setFontWeight('bold').setFontColor(THEME.blue); // Active scholars
  sh.getRange(1, c.width + 1, 1, X).setBorder(null, true, null, null, null, null, THEME.sun, SpreadsheetApp.BorderStyle.SOLID_THICK);

  const rules = [];
  if (c.help >= 0) {
    const R = col(c.help);
    rules.push(chipRule_(R, HELP_CHOICES[2], CHIP.red), chipRule_(R, HELP_CHOICES[1], CHIP.yellow),
               chipRule_(R, HELP_CHOICES[0], CHIP.green));
  }
  if (c.dc >= 0) {
    const R = col(c.dc);
    rules.push(chipRule_(R, 'Most', CHIP.green), chipRule_(R, 'Some', CHIP.blue),                    // older wording
               chipRule_(R, 'Few or none', CHIP.yellow), chipRule_(R, 'Not sure', CHIP.grey));
    const D = typeof DC_CHOICES === 'undefined' ? [] : DC_CHOICES;
    [CHIP.green, CHIP.blue, CHIP.yellow, CHIP.grey].forEach((k, i) => { if (D[i]) rules.push(chipRule_(R, D[i], k)); });
  }
  if (c.events >= 0 && typeof EVENT_GO_CHOICES !== 'undefined') {
    const R = col(c.events);
    [CHIP.green, CHIP.blue, CHIP.yellow, CHIP.grey].forEach((k, i) => rules.push(chipRule_(R, EVENT_GO_CHOICES[i], k)));
  }
  if (c.shared >= 0) {
    const R = col(c.shared);
    [CHIP.green, CHIP.yellow, CHIP.red].forEach((k, i) => rules.push(chipRule_(R, L_SHARED[i], k)));
  }
  rules.push(chipRule_(col(c.width + 1), '✅', CHIP.green, false), chipRule_(col(c.width + 1), '⏰', CHIP.yellow, false),
             chipRule_(col(c.width + 2), '— none', CHIP.red));
  sh.setConditionalFormatRules(rules);
  placeTutorial_(sh);
}



// Questions no longer in the Campus Lead form keep their old answers, but their columns are hidden (nothing is deleted)
function hideOldQuestionCols_(sh, c) {
  try {
    c = c || ckCols_(sh);
    const now = ['Timestamp', 'Email Address', Q.school, Q.name, Q.events, Q.dc, Q.help, Q.concern, CONSENT_TITLE]
      .concat(typeof LQ_VIBE === 'undefined' ? [] : [LQ_VIBE, LQ_SHARED]);
    const heads = sh.getRange(1, 1, 1, c.width).getValues()[0].map(x => String(x).trim());
    heads.forEach((h, j) => {
      if (now.indexOf(h) === -1) { if (!sh.isColumnHiddenByUser(j + 1)) sh.hideColumns(j + 1); }
      else if (sh.isColumnHiddenByUser(j + 1)) sh.showColumns(j + 1);
    });
  } catch (e) { Logger.log('Old question columns: ' + e.message); }
}


/* ---------- ONE look for the newer tables (same as Schools, Scholars, Outreach log) ----------
 * Poppins, navy text, hidden gridlines with thin blue lines, 40 px header with the yellow underline,
 * ✏️ blue headings = you type here, 🔒 navy headings = fills in by itself, white / light-blue rows, a coloured tab.
 * kinds: { heading: 'edit' | 'auto' }, anything not listed uses def. */
var TABLE_LOOK_V = 't1';
function newerTables_() {
  return [
    { name: TAB.mcheckins, def: 'auto', color: THEME.tabs.mcheckins, wide: [MQ.events, MQ.tea], when: ['Timestamp', LATE_HEAD] },
    { name: ATTEND_TAB, def: 'edit', color: THEME.tabs.events, wide: [], date: ['Date'] },
    { name: JOIN_TAB, def: 'auto', color: THEME.tabs.settings, wide: ['Note'], when: ['Requested on'] },
    { name: typeof UNKNOWN_TAB === 'string' ? UNKNOWN_TAB : '', def: 'auto', color: THEME.tabs.checkins, wide: [MQ.events, MQ.tea], when: ['Timestamp'] },
  ].filter(t => t.name);
}
function tableStyle_(sh, t) {
  const raw = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  let W = raw.length; while (W > 0 && !String(raw[W - 1]).trim()) W--;
  if (!W) return;
  const heads = raw.slice(0, W).map(plainHeading_);
  const rows = Math.max(Math.min(sh.getMaxRows() - 1, Math.max(sh.getLastRow() + 200, 300), 3000), 1);
  if (sh.getMaxRows() < rows + 1) sh.insertRowsAfter(sh.getMaxRows(), rows + 1 - sh.getMaxRows());
  tableLook_(sh, rows, W, 28, t.color);
  band_(sh.getRange(2, 1, rows, W));
  const spec = heads.map((h, i) => {
    const kind = (t.kinds && t.kinds[h]) || t.def;
    const fit = /email/i.test(h) ? 'link' : (t.wide || []).indexOf(h) !== -1 ? true : undefined;
    return [i + 1, kind, h, i === 0 ? (kind === 'auto' ? KIND.auto.note : KIND.edit.note) : '', fit];
  });
  markColumns_(sh, spec, 0, false);                    // headings get ✏️ / 🔒 and their colour (the code ignores the marks)
  heads.forEach((h, i) => {
    const col = sh.getRange(2, i + 1, rows, 1);
    if ((t.when || []).indexOf(h) !== -1) col.setNumberFormat('mmm d, yyyy h:mm am/pm');
    else if ((t.date || []).indexOf(h) !== -1) col.setNumberFormat('mmm d, yyyy');
  });
  sh.getRange(2, 1, rows, 1).setFontWeight('bold');
  sh.setFrozenRows(1);
}
// Every newer table, once per look version (and whenever one is new). Older tabs keep their own builders, which use the same pieces.
function styleNewerTables_(ss, force) {
  const props = docProps_(); let done = {};
  try { done = JSON.parse(props.getProperty('TABLE_LOOK') || '{}'); } catch (e) { done = {}; }
  let n = 0;
  newerTables_().forEach(t => {
    const sh = ss.getSheetByName(t.name);
    if (!sh || (!force && done[t.name] === TABLE_LOOK_V)) return;
    try { tableStyle_(sh, t); done[t.name] = TABLE_LOOK_V; n++; } catch (e) { Logger.log('Table look (' + t.name + '): ' + e.message); }
  });
  props.setProperty('TABLE_LOOK', JSON.stringify(done));
  // every tab coloured by the same rule: blue = you type, grey = automatic, yellow = settings and requests
  const colors = { [TAB.schools]: THEME.tabs.schools, [TAB.scholars]: THEME.tabs.scholars, [TAB.outreach]: THEME.tabs.outreach,
                   [TAB.events]: THEME.tabs.events, [TAB.settings]: THEME.tabs.settings, [TAB.checkins]: THEME.tabs.checkins };
  Object.keys(colors).forEach(n2 => { const sh = ss.getSheetByName(n2); if (sh) try { sh.setTabColor(colors[n2]); } catch (e) { /* fine */ } });
  return n;
}


// ━━━━━━━━━━━━━━━━━━━━  Dashboard  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  DASHBOARD  (visual, at-a-glance)
 *  Row 1      wide deck banner
 *  Rows 2-3   region · last update · term start date (B3 is read by other code; keep it there)
 *  Rows 5-8   four KPI cards with progress bars
 *  Rows 10+   left: School status board A:G (most urgent first, incl. last contacted)
 *             right: help requests, then charts
 *  Charts read from a hidden "ChartData" tab that fills itself with formulas.
 * ================================================================ */

var DASH = {
  widths: [140, 150, 160, 75, 95, 60, 105, 26, 100, 115, 125, 115, 115],   // A..M
  boardCols: 7,                                                              // A..G
  right: 9,                                                                  // right panel starts at I
  boardRows: 40,                                                         // status board space (schools)
  helpRows: 6,
};
var CHART_TAB = 'ChartData';

function buildDashboard_(ss) {
  if (DB_MODE) return;                                   // database mode: not used (see 00_Config)
  const sh = getOrCreate_(ss, TAB.dashboard);
  const keepDate = sh.getRange('B3').getValue();
  sh.getCharts().forEach(c => sh.removeChart(c));
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.getBandings().forEach(b => b.remove());
  const NC = DASH.widths.length, HC = NC + 2;                // + hidden helper columns
  if (sh.getMaxColumns() < HC + 1) sh.insertColumnsAfter(sh.getMaxColumns(), HC + 1 - sh.getMaxColumns());
  sh.showColumns(1, sh.getMaxColumns());
  if (sh.getMaxRows() < 70) sh.insertRowsAfter(sh.getMaxRows(), 70 - sh.getMaxRows());
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart()
    .setFontFamily(THEME.font).setFontColor(THEME.navy).setVerticalAlignment('middle').setBackground(null);
  DASH.widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.hideColumns(NC + 1, 2);                                  // helpers live in hidden columns N:O

  const data = buildChartData_(ss);
  const L = CONFIG.maxSchools + 1, S = TAB.schools;
  const col = c => `${S}!$${c}$2:$${c}$${L}`;
  const q = settingCell_('quietDays');
  const totalW = DASH.widths.reduce((a, b) => a + b, 0);

  /* ---- Rows 2-3 ---- */
  sh.getRange(2, 1, 1, NC).merge().setFormula(
    `="Live overview  ·  "&IF(${settingCell_('region')}="","",${settingCell_('region')}&"  ·  ")&` +
    `"Updated "&TEXT(NOW(),"mmm d, yyyy h:mm AM/PM")`).setFontColor(THEME.muted).setFontSize(10);
  sh.getRange('A3').setValue('Term start date').setFontWeight('bold').setFontSize(10);
  sh.getRange('B3').setValue(keepDate instanceof Date ? keepDate : new Date())
    .setNumberFormat('mmm d, yyyy').setBackground(THEME.yellow).setHorizontalAlignment('center').setFontWeight('bold')
    .setBorder(true, true, true, true, false, false, THEME.amber, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('C3:F3').merge().setValue('← first day of this term (used for "re-engaged")')
    .setFontColor(THEME.muted).setFontStyle('italic').setFontSize(9);
  const sw = settingCell_('autoWeekly');
  sh.getRange('I3:M3').merge().setFormula(`=IF(${sw}=TRUE,"🚀  Automatic weekly emails: ON","⏸  Automatic weekly emails: OFF  ·  menu ⏰ Automatic emails › 🚀 to turn on")`)
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');

  /* ---- Rows 5-8: KPI cards ---- */
  const kpiBar = (v, t) =>
    `=SPARKLINE({MIN(${v},${t}),MAX(${t}-(${v}),0)},{"charttype","bar";"color1",IF(${v}>=${t},"${THEME.green}","${THEME.blue}");"color2","#D6E4F5"})`;
  // Same rule everywhere (Dashboard, weekly email, START Campus Leads tab): a school row with a name AND a START Campus Lead
  const focal = `COUNTIFS(${col('A')},"<>",${col('D')},"<>")`, tF = settingCell_('tFocal');
  const reeng = `COUNTIFS(${col('F')},"Y",${col('G')},">="&$B$3)`, tR = settingCell_('tReeng');
  const rate = `IFERROR(COUNTIFS(${col('A')},"<>",${col('D')},"<>",${col('G')},">="&(NOW()-7))/${focal},0)`;
  const tP = `(${settingCell_('tRate')}/100)`;
  const quiet = `(COUNTIFS(${col('A')},"<>",${col('G')},"")+COUNTIFS(${col('H')},">"&${q}))`;
  const checks7 = `COUNTIFS(${ref_(TAB.checkins)}!A:A,">="&(NOW()-7))`;

  // Hidden helpers (column NC+2) so the cards can colour themselves by deadline
  const HX = colLetter_(NC + 2);
  const helpers = [
    `=${focal}`, `=${tF}`, `=IF(ISNUMBER(${settingCell_('dFocal')}),${settingCell_('dFocal')}-TODAY(),999)`,
    `=${reeng}`, `=${tR}`, `=IF(ISNUMBER(${settingCell_('dReeng')}),${settingCell_('dReeng')}-TODAY(),999)`,
  ];
  helpers.forEach((f, i) => sh.getRange(5 + i, NC + 2).setFormula(f));
  const H = r => `$${HX}$${r}`;                    // 5 focal, 6 target, 7 days left; 8 reeng, 9 target, 10 days left
  const countdown = (v, t, d, dateCell, label) =>
    `=IF(${v}>=${t},"✅ Target reached",(${t}-${v})&" ${label} to go  ·  "&IF(${d}=999,"set a deadline in Settings",` +
    `IF(${d}<0,"⚠️ deadline passed ("&TEXT(${dateCell},"mmm d")&")",${d}&" days left (by "&TEXT(${dateCell},"mmm d")&")")))`;

  const cards = [
    { c1: 1, c2: 2, label: '🧑‍🏫  START CAMPUS LEADS NAMED',
      value: `=${focal}&" / "&${tF}`, bar: kpiBar(focal, tF),
      cap: countdown(H(5), H(6), H(7), settingCell_('dFocal'), 'more'),
      warn: [H(5), H(6), H(7)] },
    { c1: 3, c2: 5, label: '🔁  SCHOOLS RE-ENGAGED THIS TERM',
      value: `=${reeng}&" / "&${tR}`, bar: kpiBar(reeng, tR),
      cap: countdown(H(8), H(9), H(10), settingCell_('dReeng'), 'more'),
      warn: [H(8), H(9), H(10)] },
    { c1: 6, c2: 9, label: '📈  WEEKLY RESPONSE RATE',
      value: `=TEXT(${rate},"0%")&"  of  "&TEXT(${tP},"0%")`, bar: kpiBar(rate, tP),
      cap: `="Schools with a START Campus Lead that answered in the last 7 days"` },
    { c1: 10, c2: NC, label: '🎓  SCHOLARS & SCHOOLS',
      value: `=COUNTA(${TAB.scholars}!A2:A)&" scholars"`,
      bar: `=COUNTA(${col('A')})&" schools  ·  "&${quiet}&" quiet"`,
      cap: `="📝 "&${checks7}&" check-in(s) in the last 7 days"` },
  ];
  cards.forEach(k => {
    const w = k.c2 - k.c1 + 1;
    const r = (row) => sh.getRange(row, k.c1, 1, w).merge();
    r(5).setValue(k.label).setFontSize(9).setFontWeight('bold').setFontColor(THEME.blue);
    r(6).setFormula(k.value).setFontSize(24).setFontWeight('bold');
    r(7).setFormula(k.bar).setFontSize(10).setFontWeight('bold');
    r(8).setFormula(k.cap).setFontSize(9).setFontColor(THEME.muted).setWrap(true);
    const box = sh.getRange(5, k.c1, 4, w);
    box.setBackground(THEME.band).setHorizontalAlignment('left')
       .setBorder(true, true, true, true, null, null, TABLE_LINE, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(5, k.c1, 4, 1).setBorder(null, true, null, null, null, null, THEME.blue, SpreadsheetApp.BorderStyle.SOLID_THICK);
  });
  [[5, 26], [6, 48], [7, 24], [8, 34]].forEach(([r, h]) => sh.setRowHeight(r, h));
  // Card colours: amber when the deadline is within 30 days and the target is not met, red within 14 days / passed
  const cardRules = [];
  cards.filter(k => k.warn).forEach(k => {
    const box = sh.getRange(5, k.c1, 4, k.c2 - k.c1 + 1);
    const [v, t, d] = k.warn;
    cardRules.push(
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(${v}<${t},${d}<=14)`)
        .setBackground(CHIP.red.bg).setRanges([box]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(${v}<${t},${d}<=30)`)
        .setBackground(CHIP.yellow.bg).setRanges([box]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=${v}>=${t}`)
        .setBackground(CHIP.green.bg).setRanges([box]).build());
  });
  sh.setRowHeight(4, 12); sh.setRowHeight(9, 16);

  /* ---- Left: School status board ---- */
  const bTop = 10, bHead = 11, bFirst = 12, bLast = bFirst + DASH.boardRows - 1;
  const BC = DASH.boardCols;
  sh.getRange(bTop, 1, 1, BC).merge().setValue('🏫  SCHOOL STATUS BOARD  ·  most urgent first');
  header_(sh.getRange(bTop, 1, 1, BC));
  sh.getRange(bHead, 1, 1, BC).setValues([['School', 'Status', 'START Campus Lead', 'Scholars', 'Last check-in', 'Days', 'Last contacted']])
    .setFontWeight('bold').setFontSize(9).setFontColor(THEME.muted).setBackground(THEME.band);
  const status = `IF(${col('G')}="","🔴 No check-in yet",IF(${col('H')}>${q},"🔴 Quiet",IF(${col('H')}<=7,"🟢 Checked in","🟡 Due soon")))`;
  const rank = `IF(${col('G')}="",0,IF(${col('H')}>${q},1,IF(${col('H')}<=7,3,2)))`;
  // Rank is added as an 8th column so we can sort on it, then CHOOSECOLS drops it again
  sh.getRange(bFirst, 1).setFormula(
    `=IFERROR(ARRAYFORMULA(CHOOSECOLS(SORT(FILTER({${col('A')},${status},IF(${col('D')}="","— none yet",${col('D')}),` +
    `${col('C')},${col('G')},${col('H')},${col('Q')},${rank}},${col('A')}<>""),8,TRUE,4,FALSE),1,2,3,4,5,6,7)),` +
    `"Add schools in the Schools tab")`);
  const board = sh.getRange(bFirst, 1, DASH.boardRows, BC);
  board.setFontSize(10);
  sh.getRange(bFirst, 1, DASH.boardRows, 1).setFontWeight('bold');
  sh.getRange(bFirst, 4, DASH.boardRows, 1).setHorizontalAlignment('center').setFontColor(THEME.blue).setFontWeight('bold');
  sh.getRange(bFirst, 5, DASH.boardRows, 1).setNumberFormat('mmm d').setHorizontalAlignment('center');
  sh.getRange(bFirst, 6, DASH.boardRows, 1).setNumberFormat('0').setHorizontalAlignment('center');
  sh.getRange(bFirst, 7, DASH.boardRows, 1).setNumberFormat('mmm d').setHorizontalAlignment('center').setFontColor(THEME.muted);
  sh.getRange(bHead, 1, DASH.boardRows + 1, BC)
    .setBorder(true, true, true, true, null, true, TABLE_LINE, SpreadsheetApp.BorderStyle.SOLID);

  /* ---- Right: help requests ---- */
  const hFirst = 12;
  const RC = DASH.right;
  sh.getRange(bTop, RC, 1, 5).merge().setValue('💬  HELP REQUESTS & CONCERNS  ·  last 14 days');
  header_(sh.getRange(bTop, RC, 1, 5));
  sh.getRange(bHead, RC, 1, 4).setValues([['Date', 'School', 'Needs help?', 'Concern']])
    .setFontWeight('bold').setFontSize(9).setFontColor(THEME.muted).setBackground(THEME.band);
  sh.getRange(bHead, RC + 4).setBackground(THEME.band);
  const helpRange = sh.getRange(hFirst, RC, DASH.helpRows, 5);
  try {
    const c = ckCols_(ss.getSheetByName(TAB.checkins));
    if (c.help < 0 || c.concern < 0) throw new Error('old form');
    const C = ref_(TAB.checkins), Lt = x => colLetter_(x + 1);
    const ts = `${C}!A2:A`, sc = `${C}!${Lt(c.school)}2:${Lt(c.school)}`;
    const hp = `${C}!${Lt(c.help)}2:${Lt(c.help)}`, cn = `${C}!${Lt(c.concern)}2:${Lt(c.concern)}`;
    const realConcern = `(${cn}<>"")*NOT(REGEXMATCH(LOWER(TRIM(${cn})),"^(none|nothing|n/?a|wala|no|no concerns?|-+|\\.+)\\.?$"))`;
    sh.getRange(hFirst, RC).setFormula(
      `=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({${ts},${sc},${hp},${cn}},${ts}>=NOW()-14,` +
      `(LEFT(${hp},3)="Yes")+${realConcern}),1,FALSE),${DASH.helpRows},4),"None in the last 14 days 🎉")`);
  } catch (e) {
    sh.getRange(hFirst, RC).setValue('(Click ① Set up / fix the tracker again to show this list.)');
  }
  helpRange.setFontSize(10).setWrap(false);
  sh.getRange(hFirst, RC, DASH.helpRows, 1).setNumberFormat('mmm d').setFontColor(THEME.muted);
  sh.getRange(hFirst, RC + 1, DASH.helpRows, 1).setFontWeight('bold').setFontColor(THEME.blue);
  sh.getRange(bHead, RC, DASH.helpRows + 1, 5)
    .setBorder(true, true, true, true, null, true, TABLE_LINE, SpreadsheetApp.BorderStyle.SOLID);

  /* ---- Row heights for the lower area ---- */
  sh.setRowHeight(bTop, 32);
  sh.setRowHeights(bHead, 60 - bHead + 1, 26);

  /* ---- Chips ---- */
  const B = sh.getRange(bFirst, 2, DASH.boardRows, 1), J = sh.getRange(hFirst, RC + 2, DASH.helpRows, 1);
  sh.setConditionalFormatRules(cardRules.concat([
    chipRule_(B, '🟢', CHIP.green, false), chipRule_(B, '🟡', CHIP.yellow, false), chipRule_(B, '🔴', CHIP.red, false),
    chipRule_(J, HELP_CHOICES[2], CHIP.red), chipRule_(J, HELP_CHOICES[1], CHIP.yellow),
    chipRule_(J, HELP_CHOICES[0], CHIP.green),
  ]));

  /* ---- Charts (right column, then under the board) ---- */
  const chartTitle = (row, c1, w, text) => {
    sh.getRange(row, c1, 1, w).merge().setValue(text);
    header_(sh.getRange(row, c1, 1, w));
  };
  const rightW = DASH.widths.slice(RC - 1).reduce((a, b) => a + b, 0);
  const leftW = DASH.widths.slice(0, BC).reduce((a, b) => a + b, 0);
  const base = (b, w, h) => b
    .setOption('width', w).setOption('height', h)
    .setOption('fontName', THEME.font).setOption('legend', { position: 'none' })
    .setOption('backgroundColor', { fill: '#FFFFFF' })
    .setOption('chartArea', { left: '22%', top: 12, width: '72%', height: '80%' })
    .setOption('hAxis', { textStyle: { color: THEME.navy, fontSize: 10 }, gridlines: { color: '#E3ECF7' } })
    .setOption('vAxis', { textStyle: { color: THEME.navy, fontSize: 10 }, gridlines: { color: '#E3ECF7' } })
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.SHOW_BOTH);

  const r1 = bHead + DASH.helpRows + 2;                     // under help requests
  chartTitle(r1, RC, 5, '🎓  ACTIVE SCHOLARS PER SCHOOL');
  sh.insertChart(base(sh.newChart().setChartType(Charts.ChartType.BAR)
      .addRange(data.getRange('A1:B' + (CONFIG.maxSchools + 1))), rightW, 330)
    .setOption('colors', [THEME.blue]).setOption('dataOpacity', 0.9)
    .setPosition(r1 + 1, RC, 0, 4).build());

  const r2 = r1 + 15;
  chartTitle(r2, RC, 5, '📚  SCHOLARS BY YEAR LEVEL');
  sh.insertChart(base(sh.newChart().setChartType(Charts.ChartType.PIE)
      .addRange(data.getRange('G1:H7')), rightW, 260)
    .setOption('pieHole', 0.55).setOption('legend', { position: 'right', textStyle: { color: THEME.navy, fontSize: 10 } })
    .setOption('chartArea', { left: 10, top: 10, width: '90%', height: '90%' })
    .setOption('colors', [THEME.sun, THEME.blue, THEME.green, THEME.navy, THEME.slate, '#C9D3E0'])
    .setOption('pieSliceText', 'value')
    .setPosition(r2 + 1, RC, 0, 4).build());

  const r3 = bLast + 2;
  chartTitle(r3, 1, BC, '📅  CHECK-INS PER WEEK  ·  last 8 weeks');
  sh.insertChart(base(sh.newChart().setChartType(Charts.ChartType.COLUMN)
      .addRange(data.getRange('D1:E9')), leftW, 260)
    .setOption('colors', [THEME.blue])
    .setOption('chartArea', { left: 40, top: 12, width: '90%', height: '75%' })
    .setPosition(r3 + 1, 1, 0, 4).build());

  /* ---- Banner last (sets row 1 height) ---- */
  placeBanner_(sh, 'dashboard', totalW, 6);
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(0);
  sh.setTabColor(THEME.tabs.dashboard);
  placeTutorial_(sh);
}

/* Hidden helper tab with live formulas the charts read from */
function buildChartData_(ss) {
  const sh = getOrCreate_(ss, CHART_TAB);
  sh.clear();
  const L = CONFIG.maxSchools + 1, S = TAB.schools, C = ref_(TAB.checkins);
  sh.getRange('A1:B1').setValues([['School', 'Active scholars']]);
  sh.getRange('A2').setFormula(`=IFERROR(SORT(FILTER({${S}!A2:A${L},${S}!C2:C${L}},${S}!A2:A${L}<>""),2,FALSE),)`);

  sh.getRange('C1:E1').setValues([['Week start', 'Week of', 'Check-ins']]);
  sh.getRange('C9').setFormula('=TODAY()-WEEKDAY(TODAY(),3)');
  for (let r = 8; r >= 2; r--) sh.getRange('C' + r).setFormula(`=C${r + 1}-7`);
  for (let r = 2; r <= 9; r++) {
    sh.getRange('D' + r).setFormula(`=TEXT(C${r},"mmm d")`);
    sh.getRange('E' + r).setFormula(`=COUNTIFS(${C}!$A:$A,">="&C${r},${C}!$A:$A,"<"&(C${r}+7))`);
  }

  sh.getRange('G1:H1').setValues([['Year level', 'Scholars']]);
  ['1st Year', '2nd Year', '3rd Year', '4th Year', '5th Year'].forEach((y, i) => {
    sh.getRange(2 + i, 7).setValue(y);
    sh.getRange(2 + i, 8).setFormula(`=COUNTIF(${TAB.scholars}!G:G,G${2 + i})`);
  });
  sh.getRange('G7').setValue('Not set');
  sh.getRange('H7').setFormula(`=MAX(COUNTA(${TAB.scholars}!A2:A)-SUM(H2:H6),0)`);

  sh.getRange('J1').setValue('This tab feeds the Dashboard charts. It fills itself in; please leave it as is.');
  sh.hideSheet();
  return sh;
}


// ━━━━━━━━━━━━━━━━━━━━  Tutorials  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  📖 "ABOUT THIS TAB" PANELS  (on the right of every tab)
 *  Each panel answers four questions in a few short lines:
 *    1. Do I type here?         ✏️ you edit  ·  🔒 automatic  ·  ✏️🔒 both
 *    2. Why does it matter?
 *    3. What do I do?           numbered steps
 *    4. How is it connected?    ← where it gets data from  ·  → what it updates
 *  Placed one empty column after each table. Re-drawn by ① (and when a tab rebuilds itself).
 * ================================================================ */

var TUTORIAL_MARK = '📖';
var TUTORIAL_WIDTH = 380;

// role: 'edit' (you type here), 'auto' (fills in by itself), 'mix' (some columns each)
// from: tabs this one reads from (←)  ·  to: tabs/things it updates (→). Keep every line short.
function tutorials_() {
  const st = (() => { try { return getSettings_(); } catch (e) { return {}; } })();
  const d = k => st[k] || '';
  return {
    [TAB.settings]: { col: 5, row: 1, role: 'edit', title: '⚙️ SETTINGS', roleText: '✏️  YOU EDIT the yellow cells only',
      why: 'Your control panel: names, targets, days and switches.',
      when: '🕒 Once at the start of the term · change anytime',
      steps: ['Type your region and your name',
              'Set your targets and their deadlines',
              `Email days: reminder ${d('remDay')}, follow-up ${d('fuDay')}, due ${d('dueDay')}`,
              'Tick 🚀 to turn ON the automatic weekly emails',
              'DataCamp dates: the forms ask about learning only then'],
      from: [],
      to: ['Both forms, every email, the Dashboard', 'Calendars and the Calendar view'],
      tip: 'Changes apply by themselves: no button needed' },

    [TAB.schools]: { col: 19, row: 1, role: 'mix', title: '🏫 SCHOOLS', roleText: '✏️  Blue columns: you type  ·  🔒 grey: automatic',
      why: 'The list of schools. Every other tab reads from it.',
      when: '🕒 When a school joins, or its Campus Lead changes',
      steps: ['One school per row: short name (A), province (B)',
              'Campus Lead (D): pick a name (or set ⭐ in Scholars)',
              'Email (E) fills in by itself from Scholars',
              'Need a school\'s own form link? Copy it from column P'],
      from: ['Scholars: emails, number of active scholars', 'Campus Lead check-ins and Outreach log: dates'],
      to: ['The form\'s school list', 'Campus Leads tab, portals, reminders'],
      tip: 'Rename a school here once: it changes everywhere' },

    [TAB.scholars]: { col: 12, row: 1, role: 'edit', title: '🎓 SCHOLARS', roleText: '✏️  YOU EDIT every column (A–J)',
      why: 'The ONE place for each scholar\'s info and email.',
      when: '🕒 When the masterlist changes · emails anytime',
      steps: ['Paste the ONSSE masterlist at cell A2',
              'School (B) must match the Schools tab',
              'Type every email in column E (only here!)',
              'Status (F): Active, Graduated, On leave, Dropped',
              'Role (J): Member or ⭐ START Campus Lead (1 per school)'],
      from: ['Schools: a Lead picked there shows ⭐ here'],
      to: ['Schools, portals, calendars, member check-in', 'Who can answer the member form (by email)'],
      tip: 'Orange row = same name twice · yellow ⭐ = Campus Lead' },

    [TAB.outreach]: { col: 9, row: 1, role: 'edit', title: '📞 OUTREACH LOG', roleText: '✏️  YOU EDIT · your contact diary',
      why: 'A record of every message, call and visit to a school.',
      when: '🕒 After every call or message',
      steps: ['Menu › 📞 Record a call or message (fastest)',
              'Or type a row: date, school, channel, outcome',
              'Add a "Follow up on" date to be reminded'],
      from: ['Schools: the school list'],
      to: ['Schools: last contacted', 'My START work calendar: follow-ups', '🗓 Today: follow-ups due'],
      tip: 'Red date = follow-up due today or late' },

    [TAB.links]: { col: 8, row: 1, role: 'mix', title: '🔗 LINKS', roleText: '🔒 Top: automatic  ·  ✏️ Bottom: yours',
      why: 'Every important link in one place.',
      when: '🕒 Anytime you need a link',
      steps: ['Top part: every form, calendar and folder link',
              'Copy a link from column C to share it',
              'Bottom part: add your own (name, link, category)',
              'Click "Open ↗" to visit a link'],
      from: ['The forms and this tracker'],
      to: ['Nothing: it is for your reference'],
      tip: 'Share links from here, never this tracker itself' },

    [TAB.dashboard]: { col: 17, row: 3, role: 'auto', title: '📊 DASHBOARD', roleText: '🔒  AUTOMATIC · just read it (except B3)',
      why: 'Your progress at a glance, for reports and planning.',
      when: '🕒 Every week, and before reporting',
      steps: ['Read the 4 cards: progress and days left',
              'Status board: most urgent school at the top',
              'Each new term: set the start date in B3 (yellow)'],
      from: ['Schools · Scholars · Settings', 'Campus Lead check-ins · Outreach log'],
      to: ['Nothing: fix the source tab instead'],
      tip: 'Wrong number? Fix it in Schools or Scholars' },

    [TAB.focal]: { col: 12, row: 1, role: 'auto', title: '🧑‍🏫 START CAMPUS LEADS', roleText: '🔒  AUTOMATIC · just read it',
      why: 'Who leads each school, and who checked in this week.',
      when: '🕒 Every week, after the deadline',
      steps: ['✅ = checked in this week · ⏳ = not yet',
              'To change a Lead: Schools (D) or Scholars (J)',
              'Once: 🛠 Setup tools › 🔐 Create / update Campus Lead portals'],
      from: ['Schools and Scholars: names and emails', 'Campus Lead check-ins: answers'],
      to: ['Each Lead\'s private portal + welcome email'],
      tip: 'Updates by itself every morning and after each check-in' },

    [TAB.checkins]: { col: null, row: 1, role: 'auto', title: '🧑‍🏫 CAMPUS LEAD CHECK-INS', roleText: '🔒  AUTOMATIC · Campus Lead answers',
      why: 'Each Campus Lead\'s weekly answer, copied here by itself: it decides the school\'s status and follow-ups.',
      when: `🕒 Fills in ${d('remDay')} to ${d('dueDay')} each week`,
      steps: ['A Campus Lead answers the weekly check-in: a row appears here',
              'One row per school per week; nothing to type',
              'Last columns are worked out for you; old questions are hidden',
              'Use the filter icons to see one school'],
      from: ['The weekly check-in (Campus Leads\' answers)', 'Older rows: the old Campus Lead form'],
      to: ['Schools red/green, Dashboard, portals', 'Follow-up emails (only to those who did not answer)'],
      tip: '"Yes, urgent" or 🆘 emails you right away' },

    [TAB.events]: { col: 17, row: 1, role: 'mix', title: '📅 EVENTS', roleText: '✏️  YOU EDIT A–K  ·  🔒 L–N fill in',
      why: 'One row = one event on the shared START calendars.',
      when: '🕒 Whenever you plan an event',
      steps: ['Easiest: menu 📅 Events & calendar › Add an event',
              'Pick "Who": All members, Campus Leads, a school, Only me',
              'Repeats + Until: ONE row for a whole series',
              'Tick "Invite Leads?" for Yes / No / Maybe',
              'Clear the whole row to remove the event'],
      from: ['Scholars and Schools: who sees it'],
      to: ['Google Calendars + invites', 'Both forms (this week\'s events)', 'Calendar view, emails, 🗓 Today'],
      tip: 'Members can see titles: no personal data in them' },

    [TAB.calview]: { col: 9, row: 1, role: 'mix', title: '🗓 CALENDAR VIEW', roleText: '🔒 AUTOMATIC  ·  ✏️ only the month (B2)',
      why: 'The whole month at a glance, like a wall calendar.',
      when: '🕒 Anytime you plan',
      steps: ['Pick the month in the yellow cell (B2)',
              'Add events: menu 📅 Events & calendar › Add an event',
              'Yellow = today · grey = other months'],
      from: ['Events · Settings (check-in days) · Outreach log'],
      to: ['Nothing: it is a picture of your plans'],
      tip: 'To change an event, edit its row in Events' },

    [TAB.mcheckins]: { col: sh => Math.max(formColCount_(sh), 6) + 2, row: 1, role: 'auto', title: '👥 MEMBER CHECK-INS', roleText: '🔒  AUTOMATIC · member answers',
      why: 'Everyone\'s weekly check-in (members AND Campus Leads), one form for all.',
      when: `🕒 Opens ${d('mOpenDay')}, due ${d('mDueDay')}`,
      steps: ['Once: 🛠 Setup tools › 📝 Set up / update member check-in',
              'Campus Leads post the link (message is in their portal)',
              'Add meme pictures to the Drive folder (rotate weekly)'],
      from: ['The member form', 'Scholars: emails (who is who)', 'Events: this week\'s events'],
      to: ['Each Lead\'s portal: My members', 'Events: Members joined', 'Your summary and 🗓 Today'],
      tip: 'Emails not in Scholars are ignored automatically' },

    [TAB.log]: { col: 10, row: 1, role: 'auto', title: '📬 REMINDER LOG', roleText: '🔒  AUTOMATIC · email history',
      why: 'Every reminder, follow-up and announcement, and to whom.',
      when: '🕒 Check after each send day',
      steps: ['Turn on: ⏰ Automatic emails › 🚀 (or tick 🚀 in Settings)',
              '"Skipped" = add their email in Scholars (E)',
              '"Waiting" = daily limit: goes out tomorrow by itself'],
      from: ['Scholars: emails', 'Settings: email days and times'],
      to: ['Nothing: it is a record'],
      tip: 'Nobody gets the same email twice: each send is marked done' },
  };
}

// Draws one tab's panel. Each tab's own build step calls this last. Cheap when nothing changed,
// so it is safe to call often (the START Campus Leads tab rebuilds after every check-in).
function placeTutorial_(sh) {
  if (DB_MODE) return;                                   // database mode: not used (see 00_Config)
  try {
    const t = tutorials_()[sh.getName()];
    if (!t) return;
    drawTutorial_(sh, t);
    if (t.row === 1 && typeof placeGuideBanner_ === 'function') placeGuideBanner_(sh, tutorialCol_(sh, t));
  } catch (e) { Logger.log('Steps panel (' + sh.getName() + '): ' + e.message); }
}

// One panel = a list of [text, style] lines
function tutorialLines_(t) {
  const L = [[TUTORIAL_MARK + '  ' + (t.title || 'THIS TAB') + '  ·  HOW TO USE', 'head'], [t.roleText, 'role'], [t.why, 'why']]
    .concat(t.when ? [[t.when, 'when']] : []).concat([['WHAT TO DO', 'label']]);
  t.steps.forEach((x, i) => L.push([(i + 1) + '   ' + x, 'step']));
  if (t.from.length || t.to.length) L.push(['CONNECTED TABS   ·   ← gets data from   ·   → updates', 'label']);
  t.from.forEach(x => L.push(['←  ' + x, 'from']));
  t.to.forEach(x => L.push(['→  ' + x, 'to']));
  L.push(['💡  ' + t.tip, 'tip']);
  return L;
}

var TUTORIAL_LOOK = {
  role: { edit: ['#0A6ED1', '#FFFFFF'], auto: ['#12203A', '#FFFFFF'], mix: ['#FFF6BF', '#12203A'] },
  why: '#FFFFFF', label: '#FFFFFF', step: '#F2F8FF', from: '#EEF2F7', to: '#E8F6EC', tip: '#FFF6BF',
};

function tutorialCol_(sh, t) {
  return typeof t.col === 'function' ? t.col(sh)
       : t.col || (ckCols_(sh).width + CK_EXTRA.length + 2);         // Weekly Check-in: after the extra columns
}

function drawTutorial_(sh, t) {
  const col = tutorialCol_(sh, t);
  const lines = tutorialLines_(t);
  const want = lines.map(l => [l[0]]);
  const rows = want.length;
  if (sh.getMaxColumns() < col) sh.insertColumnsAfter(sh.getMaxColumns(), col - sh.getMaxColumns());
  if (sh.getMaxRows() < t.row + rows) sh.insertRowsAfter(sh.getMaxRows(), t.row + rows - sh.getMaxRows());
  const box = sh.getRange(t.row, col, rows, 1);
  if (JSON.stringify(box.getValues()) === JSON.stringify(want)) return;   // already there

  // remove an older panel (another column, or a different length)
  const head = sh.getRange(t.row, 1, 1, sh.getMaxColumns()).getValues()[0];
  head.forEach((v, i) => {
    if (String(v).indexOf(TUTORIAL_MARK) === 0) sh.getRange(t.row, i + 1, Math.min(rows + 8, sh.getMaxRows() - t.row + 1), 1).clear().clearNote();
  });

  if (sh.getColumnWidth(col - 1) < 24) sh.setColumnWidth(col - 1, 24);                 // never shrinks a header layout
  if (sh.getColumnWidth(col) < TUTORIAL_WIDTH) sh.setColumnWidth(col, TUTORIAL_WIDTH);
  box.setValues(want).setFontFamily(THEME.font).setFontSize(10).setFontColor(THEME.navy)
     .setVerticalAlignment('middle').setWrap(false).setHorizontalAlignment('left');
  const at = i => sh.getRange(t.row + i, col);
  const S = SpreadsheetApp.BorderStyle;
  lines.forEach((l, i) => {
    const c = at(i);
    switch (l[1]) {
      case 'head':  header_(c, THEME.navy); break;
      case 'role': { const k = TUTORIAL_LOOK.role[t.role]; c.setBackground(k[0]).setFontColor(k[1]).setFontWeight('bold'); break; }
      case 'why':   c.setBackground(TUTORIAL_LOOK.why).setFontStyle('italic'); break;
      case 'when':  c.setBackground('#EEF5FD').setFontWeight('bold').setFontColor(THEME.blue); break;
      case 'label': c.setBackground(TUTORIAL_LOOK.label).setFontSize(8).setFontWeight('bold').setFontColor(THEME.muted); break;
      case 'step':  c.setBackground(TUTORIAL_LOOK.step).setBorder(null, true, null, null, null, null, THEME.blue, S.SOLID_THICK); break;
      case 'from':  c.setBackground(TUTORIAL_LOOK.from).setFontColor(THEME.muted)
                     .setBorder(null, true, null, null, null, null, '#9AA7BA', S.SOLID_THICK); break;
      case 'to':    c.setBackground(TUTORIAL_LOOK.to).setFontColor('#14532D')
                     .setBorder(null, true, null, null, null, null, THEME.green, S.SOLID_THICK); break;
      case 'tip':   c.setBackground(TUTORIAL_LOOK.tip).setFontWeight('bold').setFontSize(9)
                     .setBorder(null, true, null, null, null, null, THEME.sun, S.SOLID_THICK); break;
    }
  });
  at(rows - 1).setBorder(null, null, true, null, null, null, TABLE_LINE, S.SOLID);
  sh.getRange(t.row, col, rows, 1).setBorder(true, null, null, true, null, null, TABLE_LINE, S.SOLID);
}

// How many answer columns a form tab has (row 1 headings up to the first gap; the steps panel is not counted)
function formColCount_(sh) {
  const h = sh.getRange(1, 1, 1, sh.getMaxColumns()).getValues()[0];
  let n = 0;
  while (n < h.length && String(h[n]).trim() && String(h[n]).indexOf(TUTORIAL_MARK) !== 0) n++;
  return n;
}


// ━━━━━━━━━━━━━━━━━━━━  Banners  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  🖼 TAB BANNERS  (deck-style picture header at the top of every table tab)
 *  The picture sits in the top part of row 1; the column headings stay in row 1, just below it,
 *  so no table moves and the form can still write its headings in row 1 of Weekly Check-in.
 *  The pictures come from the "START Tracker pictures" folder in your Drive (see 11_Logos).
 * ================================================================ */

var TAB_BANNER_ALT = 'START tab banner';
var TAB_BANNER_RATIO = 10;          // pictures are 10 : 1
var TAB_BANNER_MAXW = 1600;         // never wider than this (px)
var TAB_BANNER_MINW = 1000;         // and at least this wide, so the header is big (it may reach past a short table)

// Which picture goes on which tab, and whether row 1 also holds column headings (heads = true)
function tabBanners_() {
  return {
    [TAB.schools]:  { key: 'schools',  heads: true },  [TAB.scholars]: { key: 'scholars', heads: true },
    [TAB.focal]:    { key: 'focal',    heads: true },  [TAB.outreach]: { key: 'outreach', heads: true },
    [TAB.checkins]: { key: 'checkins', heads: true },  [TAB.log]:      { key: 'log',      heads: true },
    [TAB.settings]: { key: 'settings', heads: true },  [TAB.links]:    { key: 'links',    heads: false },
    [TAB.events]:   { key: 'events',   heads: true },  [TAB.calview]:  { key: 'calview',  heads: true },
    [TAB.mcheckins]:{ key: 'mcheckins', heads: true },
  };
}

// The header row as ONE picture strip: the tab banner covers the whole table area and the "HOW TO USE"
// picture covers the whole steps panel, both exactly the same height. Pictures keep their shape: the
// columns are sized around them (a wider panel, or a wider gap column when the table is narrow).
var HEADER_MIN_H = 112, HEADER_MAX_H = 160, GUIDE_RATIO = 3.4;

// o = { lastCol: last table column, panelCol: steps panel column (0 = none), blob, file, alt, desc, under, minPanel }
function layoutHeader_(sh, o) {
  if (DB_MODE) return;                                   // database mode: not used (see 00_Config)
  let room = 0;
  for (let c = 1; c <= o.lastCol; c++) if (!sh.isColumnHiddenByUser(c)) room += sh.getColumnWidth(c);
  let Z = Math.round(Math.min(room, HEADER_MAX_H * TAB_BANNER_RATIO) / TAB_BANNER_RATIO);
  if (o.panelCol) Z = Math.max(Z, Math.ceil((o.minPanel || TUTORIAL_WIDTH) / GUIDE_RATIO));
  Z = Math.min(Math.max(Z, o.panelCol ? HEADER_MIN_H : 60), HEADER_MAX_H);
  const bw = Z * TAB_BANNER_RATIO;                                   // the banner's own shape (10 : 1)
  if (o.panelCol) {
    const gap = o.panelCol - 1;
    let between = 0;                                                  // columns between the table and the gap (none usually)
    for (let c = o.lastCol + 1; c < gap; c++) if (!sh.isColumnHiddenByUser(c)) between += sh.getColumnWidth(c);
    const gw = Math.max(24, bw + 14 - room - between);               // a narrow table: the gap widens so nothing overlaps
    sh.setColumnWidth(gap, gw);
    sh.getRange(1, gap).setBackground(gw > 24 ? THEME.navy : null);   // the heading bar runs on under the banner
    sh.setColumnWidth(o.panelCol, Math.round(Z * GUIDE_RATIO));       // the panel is exactly as wide as its picture
  }
  sh.setRowHeight(1, Z + o.under);
  const same = (img, w, h, d) => img && Math.abs(img.getWidth() - w) < 3 && Math.abs(img.getHeight() - h) < 3 && String(img.getAltTextDescription()) === d;
  const imgs = sh.getImages();
  // the banner over the table
  const oldB = imgs.filter(i => String(i.getAltTextTitle()) === o.alt);
  if (o.blob && !(oldB.length === 1 && same(oldB[0], bw, Z, o.desc))) {
    oldB.forEach(i => i.remove());
    sh.insertImage(o.blob, 1, 1, 0, 0).setWidth(bw).setHeight(Z).setAltTextTitle(o.alt).setAltTextDescription(o.desc);
  } else if (oldB.length === 1 && !same(oldB[0], bw, Z, o.desc)) { oldB[0].setWidth(bw).setHeight(Z); }
  // the "HOW TO USE" picture over the panel
  if (o.panelCol) {
    const pw = Math.round(Z * GUIDE_RATIO);
    const oldG = imgs.filter(i => String(i.getAltTextTitle()) === GUIDE_BANNER_ALT);
    const okG = oldG.length === 1 && oldG[0].getAnchorCell().getColumn() === o.panelCol && same(oldG[0], pw, Z, PICTURES.guide);
    if (!okG) {
      const g = picture_('guide');
      if (g) {
        oldG.forEach(i => i.remove());
        sh.insertImage(g, o.panelCol, 1, 0, 0).setWidth(pw).setHeight(Z).setAltTextTitle(GUIDE_BANNER_ALT).setAltTextDescription(PICTURES.guide);
      }
    }
  }
  return { Z, bw };
}

// Puts the tab's banner (and the panel's picture) at the top. cols = table columns when the tab has no panel.
function placeTabBanner_(sh, cols) {
  if (DB_MODE) return;                                   // database mode: not used (see 00_Config)
  try {
    const b = tabBanners_()[sh.getName()];
    if (!b) return;
    let file = PICTURES['tab-' + b.key];
    try { const f = pictureFile_('tab-' + b.key); if (f) file = f.getName(); } catch (e) { /* not added yet */ }   // an older picture until the new one is added
    let panelCol = 0;
    try { const t = tutorials_()[sh.getName()]; if (t && t.row === 1) panelCol = tutorialCol_(sh, t); } catch (e) { /* no panel */ }
    const desc = 'START – ' + sh.getName() + ' · ' + file;
    const have = sh.getImages().filter(i => String(i.getAltTextTitle()) === TAB_BANNER_ALT);
    const blob = have.length === 1 && String(have[0].getAltTextDescription()) === desc ? null : picture_('tab-' + b.key);
    if (!blob && !have.length) return;                               // picture not added yet: nothing to place
    layoutHeader_(sh, { lastCol: panelCol ? panelCol - 2 : Math.min(cols || 6, sh.getMaxColumns()), panelCol,
                        blob, alt: TAB_BANNER_ALT, desc, under: b.heads ? 42 : 6, minPanel: TUTORIAL_WIDTH });
    // headings sit at the bottom of row 1, under the pictures
    sh.getRange(1, 1, 1, sh.getMaxColumns()).setVerticalAlignment('bottom');
  } catch (e) { Logger.log('Tab banner (' + sh.getName() + '): ' + e.message); }
}

/* ---------- 📖 "HOW TO USE" picture above each tab's steps panel ---------- */
var GUIDE_BANNER_ALT = 'START guide banner';

// Called after a steps panel is drawn: lays out both header pictures again (same height, full width)
function placeGuideBanner_(sh, col) {
  if (DB_MODE) return;                                   // database mode: not used (see 00_Config)
  try {
    if (!sh.getImages().some(i => String(i.getAltTextTitle()) === TAB_BANNER_ALT)) {   // no tab banner: no panel picture either
      sh.getImages().filter(i => String(i.getAltTextTitle()) === GUIDE_BANNER_ALT).forEach(i => i.remove());
      return;
    }
    placeTabBanner_(sh);
  } catch (e) { Logger.log('Guide banner (' + sh.getName() + '): ' + e.message); }
}

/* ---------- 🖼 any tab still without its header picture gets it now ----------
 * Headers are drawn while tabs are built. A tab built before the pictures were added (or while Google's
 * time limit stopped setup) stays bare; this fills only the missing ones, a few per click, quick when
 * all are there. Called by ① (quick check and full setup). Returns how many were added. */
function repairHeaders_(ss) {
  if (DB_MODE) return 0;                                   // database mode: not used (see 00_Config)
  let added = 0;
  const has = (sh, re) => sh.getImages().some(i => re.test(String(i.getAltTextTitle())));
  const tabs = tabBanners_();
  Object.keys(tabs).forEach(name => {
    if (outOfTime_(60 * 1000)) return;
    const sh = ss.getSheetByName(name);
    if (!sh || has(sh, /^START tab banner$/)) return;
    if (!picture_('tab-' + tabs[name].key)) return;                  // picture not added: nothing to place
    placeTabBanner_(sh);
    if (has(sh, /^START tab banner$/)) added++;
  });
  [[TAB.start, 'start', () => GUIDE_WIDTH, 4], [TAB.dashboard, 'dashboard', () => DASH.widths.reduce((a, b) => a + b, 0), 6]].forEach(([name, key, w, ratio]) => {
    if (outOfTime_(60 * 1000)) return;
    const sh = ss.getSheetByName(name);
    if (!sh || has(sh, /^START banner$/) || !picture_(key)) return;
    placeBanner_(sh, key, w(), ratio);
    if (has(sh, /^START banner$/)) added++;
  });
  return added;
}


// ━━━━━━━━━━━━━━━━━━━━  Places  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  DAVAO PLACES (Region XI website): one picture a week, Monday to Sunday
 *  Pictures: START colours, made for this site. Order = the weekly rotation.
 * ================================================================ */
var PLACE_PICS = [
  { name: "Mt. Apo", src: 'data:image/webp;base64,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' },
  { name: "Samal Island", src: 'data:image/webp;base64,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' },
  { name: "Philippine Eagle Center", src: 'data:image/webp;base64,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' },
  { name: "Eden Nature Park", src: 'data:image/webp;base64,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' },
  { name: "Davao City at night", src: 'data:image/webp;base64,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' },
  { name: "San Pedro Cathedral", src: 'data:image/webp;base64,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' }
];

// Which place belongs to a date's week (switches every Monday, Philippine time): an index into PLACE_PICS
function placeIndexForDate_(date) {
  const ymd = Utilities.formatDate(date || new Date(), CONFIG.timezone, 'yyyy-MM-dd').split('-').map(Number);
  const day = new Date(Date.UTC(ymd[0], ymd[1] - 1, ymd[2]));
  const monday = day.getTime() - ((day.getUTCDay() + 6) % 7) * 86400000;
  return Math.round(monday / 604800000) % PLACE_PICS.length;
}
function placesOn_() { try { return regionProfile_().code === 'XI'; } catch (e) { return false; } }
// This week's place. Only for Region XI; other regions get none.
function placeOfWeek_() {
  try { return placesOn_() ? PLACE_PICS[placeIndexForDate_(new Date())] : null; }
  catch (e) { Logger.log('Place picture: ' + e.message); return null; }
}

// The START mark (the middle of the logo picture), built in so the website logo never depends on a Drive picture.
var START_MARK = 'data:image/png;base64,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';


// ━━━━━━━━━━━━━━━━━━━━  SchoolLogos  ━━━━━━━━━━━━━━━━━━━━
/* ================================================================
 *  SCHOOL LOGOS  ·  each school's logo on the website (ranks, home, school pages, RR page)
 *  ----------------------------------------------------------------
 *  Put each logo in the "School logos" folder (inside "START Tracker pictures"), named after the school
 *  as written in the Schools tab, e.g.  UM.png,  AdDU.jpg,  USeP.png  (the official name also works).
 *  PNG with a see-through background looks best. No logo yet = a START-coloured initials badge.
 *  🛠 Setup tools › 🏫 School logos  downloads each logo from the school's official website (Schools › Website),
 *  shows which schools still need one, and makes one zip of all the logos.
 * ================================================================ */

var SCHOOL_LOGO_FOLDER = 'School logos';
// Preferred user-supplied school-logo folder; fallback to the existing tracker folder if inaccessible.
var START_SCHOOL_LOGO_DRIVE_ID = '1HMB5NsTB8O5T7s1mqJYNpa_9yuh6WtKj';
var SCHOOL_LOGO_MAX = 40000;          // bytes per logo sent to the website (bigger files are made small first)
var SCHOOL_LOGO_PX = 160;             // size of the small copy: the biggest logo on the website is 72 px, so 160 px is sharp on any phone (about 8–20 KB)

// How many pictures a folder holds (stops counting at 3: enough to know it is in use)
function logoPics_(f) {
  let n = 0;
  try { const it = f.getFiles(); while (it.hasNext() && n < 3) { const x = it.next(); if (!x.isTrashed() && /^image\//.test(x.getMimeType())) n++; } } catch (e) { /* no access */ }
  return n;
}
// Two folders called "School logos" (one empty, one with the logos)? Use the one WITH pictures.
function bestLogoFolder_() {
  let best = null, bestN = 0;
  try { const it = DriveApp.getFoldersByName(SCHOOL_LOGO_FOLDER);
    while (it.hasNext()) { const x = it.next(); if (x.isTrashed()) continue; const n = logoPics_(x); if (n > bestN) { best = x; bestN = n; } } }
  catch (e) { /* Drive search not available */ }
  return best;
}
function schoolLogoFolder_(create) {
  const props = docProps_();
  // A named folder link takes precedence over folders of the same name elsewhere.
  const preferred = props.getProperty('SCHOOL_LOGO_PICKED') || START_SCHOOL_LOGO_DRIVE_ID;
  if (preferred) {
    try { const f = DriveApp.getFolderById(preferred); if (!f.isTrashed()) { props.setProperty('SCHOOL_LOGO_FOLDER', f.getId()); return f; } }
    catch (e) { Logger.log('School logos folder not accessible: ' + e.message); }
  }
  const id = props.getProperty('SCHOOL_LOGO_FOLDER');
  if (id) { try { const f = DriveApp.getFolderById(id);
    if (!f.isTrashed()) {
      if (logoPics_(f) || props.getProperty('SCHOOL_LOGO_PICKED') === id) return f;   // has logos, or you picked it yourself
      const b = bestLogoFolder_();                                                       // empty: another "School logos" folder has them
      if (b && b.getId() !== id) { props.setProperty('SCHOOL_LOGO_FOLDER', b.getId()); return b; }
      return f;
    } } catch (e) { /* gone */ } }
  const parent = pictureFolder_(!!create);
  if (!parent) {                                                   // pictures folder not linked: a "School logos" folder of yours anywhere in Drive
    try { const it = DriveApp.getFoldersByName(SCHOOL_LOGO_FOLDER); while (it.hasNext()) { const x = it.next(); if (!x.isTrashed()) { props.setProperty('SCHOOL_LOGO_FOLDER', x.getId()); return x; } } }
    catch (e) { /* Drive search not available */ }
    return null;
  }
  const it = parent.getFoldersByName(SCHOOL_LOGO_FOLDER);
  let f = null;
  while (!f && it.hasNext()) { const x = it.next(); if (!x.isTrashed()) f = x; }
  if (!f && create) f = parent.createFolder(SCHOOL_LOGO_FOLDER);
  if (f) props.setProperty('SCHOOL_LOGO_FOLDER', f.getId());
  return f;
}

// The short school code is the permanent database key.  Logos are matched to exact files.
// UM is the MAIN campus; UM Digos and UM Tagum each require their own PNG.
// Do not share a generic UM image across independent campuses.
function logoOwner_(name, official) { return String(name || '').trim(); }

// Drive folder (as shown in the provided screenshot): exact files per school.
var START_SCHOOL_LOGO_NAMES = {
  'USeP': ['usep'],
  'UP Mindanao': ['upmindanao'],
  'DNSC': ['dnsc'],
  'DOrSU': ['dorsu'],
  'DSSC': ['dssc'],
  'DdOSC': ['ddosc'],
  'SPAMAST': ['spamast'],
  'AdDU': ['addu'],
  'UM': ['ummain', 'ummaincampus', 'um'], // UM = UM Main, NOT Tagum or Digos
  'UM Main': ['ummain', 'ummaincampus', 'um'],
  'UM Digos': ['umdigos'],
  'UM Tagum': ['umtagum'],
  'UIC': ['uic'],
  'CJC': ['cjc']
};

function logoKey_(s) { return String(s || '').toLowerCase().replace(/ñ/g, 'n').replace(/\.[a-z0-9]{2,4}$/, '').replace(/[^a-z0-9]/g, ''); }

// { school name (as in the Schools tab): Drive file }
function schoolLogoFiles_(ss) {
  // the "School logos" folder, and also logos dropped straight into the pictures folder (the School logos folder wins)
  const folders = [];
  try { const p = pictureFolder_(false); if (p) folders.push(p); } catch (e) { /* no pictures folder */ }
  const folder = schoolLogoFolder_(false); if (folder) folders.push(folder);
  if (!folders.length) return {};
  const files = {}, loose = {};
  folders.forEach(d => {
    const inLogoFolder = !!folder && d === folder;
    const it = d.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      if (!f.isTrashed() && /^image\//.test(f.getMimeType()) && !/^start-tracker-/i.test(f.getName())) {
        const k = logoKey_(f.getName().replace(/\s*\(\d+\)(\.[a-z0-9]+)$/i, '$1'));
        files[k] = f; if (inLogoFolder) loose[k] = true;                  // forgiving names only in the logo folder
      }
    }
  });
  const sh = ss.getSheetByName(TAB.schools), out = {};
  if (!sh) return out;
  const schools = sh.getRange(2, 1, CONFIG.maxSchools, 9).getValues().map(r => ({ name: String(r[0]).trim(), official: String(r[8] || '').trim() })).filter(x => x.name);
  return logoMatch_(schools, files, loose);
}

// Exact matching prevents incorrect school logos, including the three UM campuses.
// A file named um.png belongs ONLY to the UM main campus.
function logoMatch_(schools, files, loose) {
  const out = {};
  schools.forEach(s => {
    const aliases = START_SCHOOL_LOGO_NAMES[s.name];
    // Region XI files are intentionally strict. No broad "contains" tests.
    if (aliases) {
      for (const a of aliases) {
        for (const k of [a, a + 'logo', 'logo' + a]) {
          if (files[k]) { out[s.name] = files[k]; return; }
        }
      }
      return;
    }
    // Other regions: match exact school code or official name only.
    const aliasesOther = [logoKey_(s.name), logoKey_(s.official)].filter(Boolean);
    for (const a of aliasesOther) {
      for (const k of [a, a + 'logo', 'logo' + a]) {
        if (files[k]) { out[s.name] = files[k]; return; }
      }
    }
  });
  return out;
}
// Files in the logo folders that no school matched (shown in 🏫 School logos, with a hint)
function unmatchedLogos_(ss) {
  try {
    const used = {}; Object.values(schoolLogoFiles_(ss)).forEach(f => { used[f.getId()] = true; });
    const out = [];
    [schoolLogoFolder_(false)].filter(Boolean).forEach(d => { const it = d.getFiles(); while (it.hasNext()) { const f = it.next();
      if (!f.isTrashed() && /^image\//.test(f.getMimeType()) && !used[f.getId()]) out.push(f.getName()); } });
    return out;
  } catch (e) { return []; }
}
// 🏫 School logos › "Use this folder": link any Drive folder you can open (paste its link)
function setSchoolLogoFolder(link) {
  const m = String(link || '').match(/[-\w]{25,}/);
  if (!m) return { ok: false, msg: 'Paste the link of the folder (it looks like drive.google.com/drive/folders/…).' };
  let f;
  try { f = DriveApp.getFolderById(m[0]); f.getName(); } catch (e) {
    return { ok: false, msg: 'This Google account cannot open that folder. Share it with ' + (Session.getEffectiveUser().getEmail() || 'the tracker\'s account') + ' (Viewer is enough), then try again.' };
  }
  docProps_().setProperty('SCHOOL_LOGO_FOLDER', f.getId()); docProps_().setProperty('SCHOOL_LOGO_PICKED', f.getId());
  CacheService.getScriptCache().remove('SLOGO_VER');
  let n = 0; const it = f.getFiles(); while (it.hasNext()) { const x = it.next(); if (/^image\//.test(x.getMimeType())) n++; }
  return { ok: true, msg: 'Linked "' + f.getName() + '" (' + n + ' picture' + (n === 1 ? '' : 's') + '). Reopen this window to see which school each one matched.' };
}

// A big logo made small: Drive's own 256-pixel copy of the picture (sharp on phones, about 15-40 KB instead of 1-2 MB)
function schoolLogoSmall_(f) {
  try {
    const r = UrlFetchApp.fetch('https://drive.google.com/thumbnail?sz=w' + SCHOOL_LOGO_PX + '&id=' + f.getId(),
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true, followRedirects: true });
    const type = String(r.getHeaders()['Content-Type'] || r.getHeaders()['content-type'] || '').split(';')[0].trim();
    if (r.getResponseCode() === 200 && /^image\//.test(type) && r.getContent().length <= SCHOOL_LOGO_MAX) return r.getBlob().setContentType(type);
  } catch (e) { Logger.log('Small logo ' + f.getName() + ': ' + e.message); }
  return null;
}

// A logo as a data: link (no sharing needed). Kept for 6 hours; a changed file is read again.
function schoolLogoData_(f) {
  const key = 'SLOGO_' + SCHOOL_LOGO_PX + '_' + f.getId() + '_' + f.getLastUpdated().getTime();
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit) return hit;
  let blob = schoolLogoSmall_(f) || (f.getSize() <= SCHOOL_LOGO_MAX ? f.getBlob() : null);   // always the small copy first: big logos load fast too
  if (!blob) { try { blob = f.getThumbnail(); } catch (e) { blob = null; } }
  if (!blob || blob.getBytes().length > SCHOOL_LOGO_MAX * 4) { Logger.log('Logo too big to send: ' + f.getName()); return ''; }
  const uri = 'data:' + (blob.getContentType() || 'image/png') + ';base64,' + Utilities.base64Encode(blob.getBytes());
  try { cache.put(key, uri, 21600); } catch (e) { /* too big for the cache: still sent */ }
  return uri;
}

// A short fingerprint of the logo set, so each browser keeps its copy until a logo is added or changed
function schoolLogoVersion_(ss) {
  const cache = CacheService.getScriptCache(), hit = cache.get('SLOGO_VER');
  if (hit !== null && hit !== undefined) return hit;
  let v = '';
  try {
    const files = schoolLogoFiles_(ss);
    const sig = 'exact-region-xi-v2|' + SCHOOL_LOGO_PX + '|' + Object.keys(files).sort().map(n => n + ':' + files[n].getId() + ':' + files[n].getLastUpdated().getTime()).join('|');
    v = sig ? Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, sig)).slice(0, 12) : '';
  } catch (e) { Logger.log('School logos: ' + e.message); }
  cache.put('SLOGO_VER', v, 600);
  return v;
}

// Website: every school's logo (only for signed-in people; logos are public pictures anyway)
function webLogos(token) {
  return webSafe_(() => {
    const s = webSession_(token);
    if (!s) return { ok: false, signedOut: true };
    const ss = webSheet_(), files = schoolLogoFiles_(ss), logos = {};
    Object.keys(files).forEach(n => { try { const u = schoolLogoData_(files[n]); if (u) logos[n] = u; } catch (e) { Logger.log('Logo ' + n + ': ' + e.message); } });
    return { ok: true, v: schoolLogoVersion_(ss), logos };
  });
}

/* ---------- getting the logos from each school's official website ---------- */
var LOGO_AUTO_TAG = 'START auto logo: ';

// The school's website: Schools tab › Website (L), else the ready-made school list
function schoolWebsite_(name, typed) {
  let w = String(typed || '').trim();
  if (!w) { try { const r = regionProfile_().schools.find(x => logoKey_(x[0]) === logoKey_(name)); if (r) w = r[5]; } catch (e) { /* no list */ } }
  if (w && !/^https?:\/\//i.test(w)) w = 'https://' + w;
  return w;
}

// Picture links on a page that may be the school's logo, best first:
// site tile / app icon (usually the seal) > big icons > header pictures named "logo" or "seal" > share picture (only if it is a logo)
// > Google's large site icon > the small favicon
function logoCandidates_(html, site, names) {
  const abs = u => { try { u = String(u || '').trim().replace(/&amp;/g, '&'); if (!u || /^data:/i.test(u)) return ''; if (/^\/\//.test(u)) return 'https:' + u;
    if (/^https?:/i.test(u)) return u; const o = site.match(/^https?:\/\/[^/]+/i)[0]; return u[0] === '/' ? o + u : site.replace(/[^/]*$/, '') + u; } catch (e) { return ''; } };
  const attr = (tag, a) => { const m = tag.match(new RegExp('\\b' + a + '\\s*=\\s*("([^"]*)"|\'([^\']*)\'|([^\\s>]+))', 'i')); return m ? (m[2] || m[3] || m[4] || '') : ''; };
  const out = [], add = (u, why) => { u = abs(u); if (u && !out.some(x => x.url === u)) out.push({ url: u, why }); };
  const tags = re => (html.match(re) || []);
  tags(/<meta[^>]+>/gi).forEach(t => { if (/msapplication-TileImage/i.test(attr(t, 'name'))) add(attr(t, 'content'), 'site tile'); });
  tags(/<link[^>]+>/gi).forEach(t => { if (/apple-touch-icon/i.test(attr(t, 'rel'))) add(attr(t, 'href'), 'app icon'); });
  tags(/<link[^>]+>/gi).filter(t => /(^|\s)icon(\s|$)/i.test(attr(t, 'rel')) && (parseInt(attr(t, 'sizes'), 10) || 0) >= 128)
    .sort((a, b) => (parseInt(attr(b, 'sizes'), 10) || 0) - (parseInt(attr(a, 'sizes'), 10) || 0)).forEach(t => add(attr(t, 'href'), 'big icon'));
  const want = names.map(logoKey_).filter(Boolean);
  tags(/<img[^>]+>/gi).forEach(t => {
    const src = attr(t, 'src') || attr(t, 'data-src'), text = (src + ' ' + attr(t, 'alt') + ' ' + attr(t, 'class') + ' ' + attr(t, 'id')).toLowerCase();
    if (/logo|seal|crest|emblem|shield/.test(text) || want.indexOf(logoKey_(attr(t, 'alt'))) !== -1) add(src, 'logo on the page');
  });
  tags(/<meta[^>]+>/gi).forEach(t => { if (/og:image$/i.test(attr(t, 'property')) && /logo|favicon|seal|icon/i.test(attr(t, 'content'))) add(attr(t, 'content'), 'share picture'); });
  const host = (site.match(/^https?:\/\/([^/]+)/i) || [])[1];
  if (host) add('https://t3.gstatic.com/faviconV2?client=SOCIAL&type=FAVICON&fallback_opts=TYPE,SIZE,URL&size=256&url=https://' + host, "Google's site icon");
  tags(/<link[^>]+>/gi).forEach(t => { if (/(^|\s)icon(\s|$)/i.test(attr(t, 'rel'))) add(attr(t, 'href'), 'icon'); });
  if (host) add('https://' + host + '/favicon.ico', 'favicon');
  if (host) add('https://' + host + '/favicon.png', 'site seal');                  // UM campuses keep their seal here
  return out;
}

// Downloads the first candidate that really is a picture (not a 1-pixel placeholder, not huge)
function fetchLogo_(site, names) {
  const opt = { muteHttpExceptions: true, followRedirects: true, validateHttpsCertificates: false, headers: { 'User-Agent': 'Mozilla/5.0 (START Tracker; logo for the DOST-SEI START website)' } };
  let html = '';
  try { const r = UrlFetchApp.fetch(site, opt); if (r.getResponseCode() < 400) html = r.getContentText(); } catch (e) { html = ''; }
  // "www.digos.umindanao.edu.ph" does not exist but "digos.umindanao.edu.ph" does: try the address without "www." too
  if (!html && /^https?:\/\/www\.[^/]+\.[^/]+\.[^/]+/i.test(site)) {
    const bare = site.replace(/^(https?:\/\/)www\./i, 'https://');
    try { const r = UrlFetchApp.fetch(bare, opt); if (r.getResponseCode() < 400) { html = r.getContentText(); site = bare; } } catch (e) { /* keep trying the icons */ }
  }
  const tries = logoCandidates_(html, site, names).slice(0, 8);
  for (let i = 0; i < tries.length; i++) {
    try {
      const r = UrlFetchApp.fetch(tries[i].url, opt), type = String(r.getHeaders()['Content-Type'] || r.getHeaders()['content-type'] || '').split(';')[0].trim().toLowerCase();
      const n = r.getContent().length;
      if (r.getResponseCode() === 200 && /^image\//.test(type) && n >= 700 && n <= 4000000) return { blob: r.getBlob().setContentType(type), from: tries[i].url, why: tries[i].why, type };
    } catch (e) { /* next one */ }
  }
  return null;
}

// Gets every missing logo (refresh = also replace the ones fetched before; logos you added yourself are never replaced).
// Each campus with its own website (UM Digos, UM Tagum) is fetched from it; a campus without one shares the group's logo (UM.png).
function getSchoolLogos(refresh) {
  const ss = SpreadsheetApp.getActive(), folder = schoolLogoFolder_(true), have = schoolLogoFiles_(ss), sh = ss.getSheetByName(TAB.schools);
  const stop = Date.now() + 4.5 * 60 * 1000, rows = sh.getRange(2, 1, CONFIG.maxSchools, 12).getValues().filter(r => String(r[0]).trim());
  const groups = {}, order = [];
  const ownerSite = {}; rows.forEach(r => { const sc = String(r[0]).trim(); if (logoOwner_(sc, r[8]) === sc) ownerSite[sc] = schoolWebsite_(sc, r[11]); });
  rows.forEach(r => {
    const school = String(r[0]).trim(), shared = logoOwner_(school, r[8]), site0 = schoolWebsite_(school, r[11]);
    // a campus with its OWN website (UM Digos, UM Tagum) gets its own logo; one without falls back to the group's (UM)
    const owner = shared !== school && site0 && site0 !== ownerSite[shared] ? school : shared;
    if (!groups[owner]) { groups[owner] = { owner, schools: [], site: '', names: [owner] }; order.push(owner); }
    const g = groups[owner];
    g.schools.push(school); g.names.push(school, r[8]);
    const own = schoolWebsite_(school, r[11]);
    if (own && (!g.site || school === owner)) g.site = own;          // the main school's website first
  });
  const out = [];
  order.forEach(owner => {
    const g = groups[owner], old = have[g.schools[0]], auto = old && String(old.getDescription() || '').indexOf(LOGO_AUTO_TAG) === 0;
    const site = schoolWebsite_(owner, '') || g.site;
    let res;
    if (old && (!auto || !refresh)) res = { done: true, note: auto ? 'already fetched' : 'your own file (kept)' };
    else if (Date.now() > stop) res = { done: false, note: 'time is up: click again for the rest' };
    else if (!site) res = { done: false, note: 'no website in Schools › Website' };
    else {
      const got = fetchLogo_(site, g.names.filter(Boolean));
      if (!got) res = { done: false, note: 'nothing found on ' + site };
      else {
        const ext = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/webp': 'webp', 'image/svg+xml': 'svg', 'image/x-icon': 'ico', 'image/vnd.microsoft.icon': 'ico' }[got.type] || 'png';
        if (old && auto) old.setTrashed(true);
        folder.createFile(got.blob.setName(owner.replace(/[\\/:*?"<>|]/g, ' ').trim() + '.' + ext)).setDescription(LOGO_AUTO_TAG + got.from);
        res = { done: true, note: 'from the ' + got.why };
      }
    }
    g.schools.forEach(school => out.push({ name: school, done: res.done, note: school === owner ? res.note : 'uses the ' + owner + ' logo' + (res.done ? '' : ' (' + res.note + ')') }));
  });
  CacheService.getScriptCache().remove('SLOGO_VER');
  return out;
}

// 📦 One zip with every school's logo (a shared logo once, e.g. UM.png for all UM schools), saved next to the logos
function zipSchoolLogos() {
  const ss = SpreadsheetApp.getActive(), folder = schoolLogoFolder_(true), files = schoolLogoFiles_(ss), seen = {}, blobs = [];
  const officials = {}; try { ss.getSheetByName(TAB.schools).getRange(2, 1, CONFIG.maxSchools, 9).getValues().forEach(r => { officials[String(r[0]).trim()] = r[8]; }); } catch (e) { /* names only */ }
  Object.keys(files).forEach(n => {
    const f = files[n], owner = logoOwner_(n, officials[n]);
    if (seen[f.getId()]) return; seen[f.getId()] = true;
    const ext = (f.getName().match(/\.[a-z0-9]+$/i) || ['.png'])[0];
    const label = files[owner] && files[owner].getId() === f.getId() ? owner : n;     // a campus's own logo keeps the campus name
    blobs.push(f.getBlob().setName(label.replace(/[\\/:*?"<>|]/g, ' ') + ext));
  });
  if (!blobs.length) throw new Error('No logos yet. Get them first, or add them to the folder.');
  const zipName = 'START school logos (' + (regionProfile_().label || 'region') + ').zip';
  const old = folder.getFilesByName(zipName); while (old.hasNext()) old.next().setTrashed(true);
  const z = folder.createFile(Utilities.zip(blobs, zipName));
  return { url: z.getUrl(), name: zipName, count: blobs.length };
}

// 🛠 Setup tools › 🏫 School logos: get them from the school websites, see who has one, make a zip
function showSchoolLogos() {
  const ss = SpreadsheetApp.getActive(), folder = schoolLogoFolder_(true);
  CacheService.getScriptCache().remove('SLOGO_VER');
  const files = schoolLogoFiles_(ss);
  const names = schoolRows_(ss).map(s => s.name), unmatched = unmatchedLogos_(ss);
  const row = n => `<tr id="r-${logoKey_(n)}"><td class="st">${files[n] ? '✅' : '⬜'}</td><td><b>${esc_(n)}</b></td><td class="m">${files[n] ? esc_(files[n].getName()) : 'no logo yet'}</td></tr>`;
  const html = `<style>body{font:14px/1.5 Poppins,Arial,sans-serif;color:#12203A;margin:0;padding:4px 6px}
    .b{display:inline-block;background:#FFDD00;color:#12203A;font:700 13px Poppins,Arial,sans-serif;text-decoration:none;padding:9px 16px;border-radius:999px;border:0;border-bottom:3px solid #0099FF;cursor:pointer;margin:0 6px 6px 0}
    .b.l{background:#fff;border:1.5px solid #12203A}.b[disabled]{opacity:.5}
    table{border-collapse:collapse;width:100%;margin-top:8px}td{padding:6px 8px;border-bottom:1px solid #E7EBF1}.m{color:#7A8699;font-size:12.5px}#msg{margin:6px 0;font-weight:600}</style>
    <p>The tracker can download each school's logo from its <b>official website</b> (Schools tab › Website). Logos you put in the folder yourself are never replaced.</p>
    <button class="b" id="get" onclick="go(false)">🌐 Get logos from the school websites</button><button class="b l" onclick="go(true)">Fetch again</button>
    <button class="b l" onclick="zip()">📦 Make a zip of all logos</button>
    <a class="b l" href="${folder.getUrl()}" target="_blank">Open the folder</a>
    <p class="m" style="margin:4px 0 0">Logo folder in use: <b>${esc_(folder.getName())}</b>${(() => { try { return ' (owner: ' + esc_(folder.getOwner().getEmail()) + ')'; } catch (e) { return ''; } })()}</p>
    <div style="display:flex;gap:6px;margin:8px 0"><input id="fl" placeholder="Use another folder: paste its Drive link" style="flex:1;font:13px Poppins,Arial,sans-serif;padding:8px 12px;border:1.5px solid #CBD3DE;border-radius:999px">
      <button class="b l" style="margin:0" onclick="link()">Use this folder</button></div>
    ${unmatched.length ? '<p class="m">Not matched to a school (rename them like the school, e.g. <b>' + esc_(names[0] || 'USeP') + '.png</b>): ' + unmatched.map(esc_).join(', ') + '</p>' : ''}
    <div id="msg"></div>
    <table>${names.map(row).join('')}</table>
    <p class="m">Wrong or blurry logo? Put the right file in the folder, named after the school (e.g. ${esc_(names[0] || 'UM')}.png): your file wins.
    Names are forgiving: "USeP logo.png" or the school's full name also work. Each UM campus can have its own logo (UM Tagum.png); a campus without one uses UM.png. Logos are shown only to signed-in START members.</p>
    <script>
      const key = s => String(s||'').toLowerCase().replace(/ñ/g,'n').replace(/[^a-z0-9]/g,'');
      function busy(t){ document.getElementById('msg').textContent = t; document.querySelectorAll('button').forEach(b => b.disabled = !!t && t.slice(-1) === '…'); }
      function go(refresh){ busy('Visiting the school websites… (up to 4 minutes)');
        google.script.run.withSuccessHandler(list => { let ok = 0; list.forEach(x => { const r = document.getElementById('r-' + key(x.name)); if (x.done) ok++;
            if (r) { r.querySelector('.st').textContent = x.done ? '✅' : '⚠️'; r.querySelector('.m').textContent = x.note; } });
            busy(ok + ' of ' + list.length + ' schools have a logo. The website shows them within 10 minutes.'); })
          .withFailureHandler(e => busy('Problem: ' + e.message)).getSchoolLogos(refresh); }
      function link(){ busy('Linking…'); google.script.run.withSuccessHandler(r => busy(r.msg)).withFailureHandler(e => busy('Problem: ' + e.message)).setSchoolLogoFolder(document.getElementById('fl').value); }
      function zip(){ busy('Making the zip…');
        google.script.run.withSuccessHandler(z => { document.getElementById('msg').innerHTML = '📦 <a href="' + z.url + '" target="_blank">' + z.name + '</a> (' + z.count + ' logos) is in the folder.'; document.querySelectorAll('button').forEach(b => b.disabled = false); })
          .withFailureHandler(e => busy('Problem: ' + e.message)).zipSchoolLogos(); }
    </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(600).setHeight(620), '🏫 School logos');
}
function menuSchoolLogos() { safe_(showSchoolLogos); }
