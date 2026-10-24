# START Tracker Apps Script source

These files are local copies of the pasted Apps Script project. The `.gs` and `.html` filenames match the names used in the Apps Script editor.

`07_Look.gs` exists locally but is currently empty (0 bytes). Save the complete Look source into it before treating this folder as a complete project. It defines the spreadsheet table styling, picture helpers, school logos, and the `webLogos` server call. The web app entry now tolerates missing decorative images, but setup and some portal features still require this file.

Editing these files does not update the live Google Apps Script project. Changes must be copied into its editor or synced through a configured deployment workflow.

## Data flow update (2026-10-24)

1. Make a backup of the live spreadsheet. In its bound Apps Script project, replace `02_Menu_Setup.gs`, `04_Emails.gs`, `05_Portals.gs`, `06_Calendar.gs`, `08_Tools.gs`, and `Web_App.html` with these versions. Add `09_Data_Flow.gs`. Keep the live project's complete `07_Look.gs` until the empty local file is replaced with its source.
2. Save the project and run `setup()` as the automation owner. Accept authorization prompts. If setup says it needs another click, run it again. Version `v61` adds stable Scholar and Join request IDs, status/update metadata, a data change log, check-in week/submission timestamps, and a weekly roster snapshot. It repairs school count formulas and the Scholars filter. Existing records remain in their original columns.
3. In the spreadsheet, use **Events & calendar → Set up / update calendars** as the owner and authorize the Advanced Calendar service. The current Events rows show `calendar.events.insert` permission errors, so the owner must grant access before they can sync. Failed events and sharing now retry daily.
4. Open the Regional Representative's **Scholars → Data quality** card. Fill missing scholar details and review any unmatched school or duplicate email. Existing historical weekly rates without snapshots are marked as estimates; snapshots preserve denominators from this release onward.
5. Redeploy the web app as a new version so the portal uses the updated server and HTML. The website loader icons in `../icons/` were also replaced with the supplied Region XI badge; publish the site host separately.

Local verification: `node tests/data_flow.test.cjs`, `node tests/web_syntax.test.cjs`, `node tests/web_entry.test.cjs`, and `node tests/email_template.test.cjs` from this folder. The web app configuration is inserted by `doGet()` into the `__START_BOOT__` marker, keeping `Web_App.html` valid JavaScript in local editors. The entry test checks the page and plain email with the Look module absent; a full integration check still needs that module.
