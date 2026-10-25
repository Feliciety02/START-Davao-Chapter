# START Tracker Apps Script source

These files are local copies of the pasted Apps Script project. The `.gs` and `.html` filenames match the names used in the Apps Script editor.

`07_Look.gs` exists locally but is currently empty (0 bytes). Save the complete Look source into it before treating this folder as a complete project. It defines the spreadsheet table styling, picture helpers, school logos, and the `webLogos` server call. The web app entry now tolerates missing decorative images, but setup and some portal features still require this file.

Editing these files does not update the live Google Apps Script project. Changes must be copied into its editor or synced through a configured deployment workflow.

## Website database controls (2026-10-25)

The RR website now has a Schools editor (including the region school list) and a Settings page. School renames update linked records. The Sheets menu no longer duplicates the region list import or the automatic email switch; those controls are on the website. Setup restricts editing of the Schools and Settings tabs to the tracker and automation owners, while those tabs continue to store the data.

The RR overview loads without the scholar roster, join history, data quality scan, event calendar expansion, or editing records. Scholars, Schools, Outreach, Settings, Events, and Diagnostics fetch their data when opened. The Scholars roster shows 50 rows at a time; join history loads 20 at a time. Outreach and Events can be edited on the website. Postcard pictures load only when their cards become visible. The website logs server and browser request times for diagnosing slow pages. A page's first visit still depends on Google Sheets response time; local tests cannot measure live Apps Script latency.

The website header displays the [original START-DOST logo](https://www.start-dost.org/) as an image beside the yellow region pill. The app loader and install icons use the same mark with a REGION XI pill baked into the PNG. No circular frame is drawn around the mark. The logo source and icon build script are in `../icons/`.

Weekly summary recipients no longer receive website admin access merely because their address is in Settings. Admin access is limited to the tracker owner and automation owner. Existing sessions are checked against that rule on every request.

Sync all changed `.gs` files and `Web_App.html` to the bound Apps Script project, retaining its complete `07_Look.gs`. Run **Set up / fix the tracker** as the owner to apply protections, then redeploy the web app as a new version. Publish the root app files separately for the new installed icon. The local `07_Look.gs` is empty, so a live setup cannot be verified from this checkout. Local checks: `node --test apps-script/tests/*.test.cjs` from the workspace root.

## Data flow update (2026-10-24)

1. Make a backup of the live spreadsheet. In its bound Apps Script project, replace `02_Menu_Setup.gs`, `04_Emails.gs`, `05_Portals.gs`, `06_Calendar.gs`, `08_Tools.gs`, and `Web_App.html` with these versions. Add `09_Data_Flow.gs`. Keep the live project's complete `07_Look.gs` until the empty local file is replaced with its source.
2. Save the project and run `setup()` as the automation owner. Accept authorization prompts. If setup says it needs another click, run it again. Version `v61` adds stable Scholar and Join request IDs, status/update metadata, a data change log, check-in week/submission timestamps, and a weekly roster snapshot. It repairs school count formulas and the Scholars filter. Existing records remain in their original columns.
3. In the spreadsheet, use **Events & calendar → Set up / update calendars** as the owner and authorize the Advanced Calendar service. The current Events rows show `calendar.events.insert` permission errors, so the owner must grant access before they can sync. Failed events and sharing now retry daily.
4. Open the Regional Representative's **Scholars → Data quality** card. Fill missing scholar details and review any unmatched school or duplicate email. Existing historical weekly rates without snapshots are marked as estimates; snapshots preserve denominators from this release onward.
5. Redeploy the web app as a new version so the portal uses the updated server and HTML. Publish the app host separately for icon changes.

Local verification: `node tests/data_flow.test.cjs`, `node tests/web_syntax.test.cjs`, `node tests/web_entry.test.cjs`, and `node tests/email_template.test.cjs` from this folder. The web app configuration is inserted by `doGet()` into the `__START_BOOT__` marker, keeping `Web_App.html` valid JavaScript in local editors. The entry test checks the page and plain email with the Look module absent; a full integration check still needs that module.

## Registration and first-time login update (2026-10-24o)

Copy `00_Config.gs`, `05_Portals.gs`, and `Web_App.html` into the same live Apps Script project, keeping the live `07_Look.gs`. Run **START Tracker → ① Set up / fix the tracker** as the tracker owner so Settings gets the **Auto-approve verified join requests** checkbox. Then redeploy the web app as a **new version**. The website switch and Settings checkbox control the same value. When enabled, auto-approval adds every valid sign-up from a listed school, including scholars absent from the old masterlist. Conflicting Scholars records wait for RR review. Scholars confirm their email before setting a password.

An approved scholar without a password sees **Welcome to START!**, confirms the email, receives a code, creates a password, and reaches **You're All Set!**. Existing accounts with a password still use **Welcome Back!**. Pending requests can be checked without submitting the registration again. Verification codes remain limited in time and rate, and are consumed once.

Local checks: `node tests/checkin_submit.test.cjs`, `node tests/registration_flow.test.cjs`, `node tests/auth_states.test.cjs`, `node tests/auth_flow.test.cjs`, `node tests/data_flow.test.cjs`, `node tests/web_syntax.test.cjs`, and `node tests/web_entry.test.cjs`. The live owner still needs to test one real check-in, one new auto-approved application, one manually approved application, actual email delivery, and sign-in on a second device after deployment.

## Fresh spreadsheet tabs (2026-10-24)

After syncing the updated `00_Config.gs` and `05_Portals.gs` into the **bound Apps Script project** (keep the live project's complete `07_Look.gs`), open the tracker as its owner and choose **START Tracker → 🧹 Fresh start: remove all data (backup first)**. Type `FRESH START` in the sheet's prompt. The action makes a private full copy first, then deletes every tab except **Settings** and **Schools**, including unknown and retired tabs. It clears Campus Lead assignments, removes website accounts and sessions, and rebuilds clean data tabs with the current headings and formats. The school list and Settings remain so new scholars can register. The old masterlist is deleted and auto-approval is switched off until a new masterlist is supplied. If setup pauses at Google's execution limit, click **① Set up / fix the tracker** until it finishes. The backup link is shown in the result.

This local folder has no live spreadsheet connection, so editing source here does not run the reset in the Google Sheet. Verify the backup link and the resulting tab list in the live tracker after deployment. Local reset check: `node tests/fresh_start.test.cjs`.
