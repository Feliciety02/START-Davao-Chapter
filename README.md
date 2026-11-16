<p align="center">
  <img src="docs/readme-banner.png" width="1200" alt="START Scholar Portal banner showing real phone screens for Member, Campus Lead, and Regional Representative roles">
</p>

<h1 align="center">START Scholar Portal</h1>

<p align="center">One place for DOST-SEI START scholars, Campus Leads, and Regional Representatives in Davao Region. Check in each week, follow your Journey, and stay connected to your school.</p>

<p align="center"><code>Region XI · Davao</code> &nbsp; <code>Web / PWA</code> &nbsp; <code>GitHub Pages</code> &nbsp; <code>Google Sheets + Apps Script</code></p>

<p align="center">
  <a href="https://feliciety02.github.io/START-Davao-Chapter/"><strong>Open the app</strong></a> &nbsp;·&nbsp;
  <a href="#app-tour"><strong>App tour</strong></a> &nbsp;·&nbsp;
  <a href="#how-it-works"><strong>How it works</strong></a> &nbsp;·&nbsp;
  <a href="#install-start"><strong>Install START</strong></a> &nbsp;·&nbsp;
  <a href="#set-up-a-region"><strong>Set up a region</strong></a>
</p>

The installable app is hosted on GitHub Pages. It opens the START website from the region's Google Apps Script deployment, which uses the START Tracker spreadsheet. The app shell works on phones and computers without an app store. It caches the shell for weak connections and can keep a scholar's recent page and an unsent weekly check-in on their device.

## At a glance

<table>
  <tr>
    <td align="center" width="25%"><strong>01 · Weekly check-ins</strong><br><sub>One question at a time, with an offline answer option</sub></td>
    <td align="center" width="25%"><strong>02 · Journey</strong><br><sub>Brain breaks, XP, progress, and reward requests</sub></td>
    <td align="center" width="25%"><strong>03 · Campus Leads</strong><br><sub>School updates, events, and scholar follow-up</sub></td>
    <td align="center" width="25%"><strong>04 · Regional Representatives</strong><br><sub>Registration, schools, outreach, events, and settings</sub></td>
  </tr>
</table>

Scholars can sign in with their START account. New eligible scholars can register from the portal; requests that need review go to the Regional Representative. Campus Lead and Regional Representative pages use their assigned access in the tracker.

## App tour

These are **unedited 1080 × 2400 screenshots from the installed START web app**, captured through USB debugging on a realme RMX2151 on November 16, 2026. The Lead and Member screens use the Regional Representative account's practice view. Select an image to open it at full size. The [capture archive](docs/screenshots/portal-2026-11-16/README.md) records the device, method, and file hashes.

<table>
  <tr>
    <td align="center"><a href="docs/screenshots/portal-2026-11-16/01-member-journey.png"><img src="docs/screenshots/portal-2026-11-16/01-member-journey.png" width="190" alt="Member practice view showing Journey level, XP, and weekly streak"></a><br><strong>1 · Journey</strong><br><sub>XP, streak, and progress</sub></td>
    <td align="center"><a href="docs/screenshots/portal-2026-11-16/02-member-events.png"><img src="docs/screenshots/portal-2026-11-16/02-member-events.png" width="190" alt="Member practice view showing the START events calendar"></a><br><strong>2 · Events</strong><br><sub>START dates and schedules</sub></td>
    <td align="center"><a href="docs/screenshots/portal-2026-11-16/03-member-ranks.png"><img src="docs/screenshots/portal-2026-11-16/03-member-ranks.png" width="190" alt="Member practice view showing school rankings and eligibility"></a><br><strong>3 · Ranks</strong><br><sub>School check-in standings</sub></td>
  </tr>
  <tr>
    <td align="center"><a href="docs/screenshots/portal-2026-11-16/04-lead-school.png"><img src="docs/screenshots/portal-2026-11-16/04-lead-school.png" width="190" alt="Campus Lead practice view showing the school home page and weekly check-in"></a><br><strong>4 · Campus Lead</strong><br><sub>School and weekly check-in</sub></td>
    <td align="center"><a href="docs/screenshots/portal-2026-11-16/05-rr-schools.png"><img src="docs/screenshots/portal-2026-11-16/05-rr-schools.png" width="190" alt="Regional Representative screen for managing school records"></a><br><strong>5 · Schools</strong><br><sub>Regional school records</sub></td>
    <td align="center"><a href="docs/screenshots/portal-2026-11-16/06-rr-calendar.png"><img src="docs/screenshots/portal-2026-11-16/06-rr-calendar.png" width="190" alt="Regional Representative screen for managing events and the calendar"></a><br><strong>6 · Calendar</strong><br><sub>Regional event management</sub></td>
  </tr>
</table>

The app shell's [loading](docs/screenshots/01-loading.png), [region picker](docs/screenshots/02-region-picker.png), [installation hint](docs/screenshots/03-install-iphone.png), and [offline screen](docs/screenshots/04-offline.png) are also available.

## How it works

```mermaid
flowchart TD
    A["Phone or desktop browser"] --> B["START app shell on GitHub Pages"]
    B --> C["Region's Apps Script website"]
    C --> D["START Tracker Google Sheet"]
    B --> E["Device storage: sign-in, recent scholar page, pending check-in"]
    E -->|Reconnect and open START| C
    D --> C
    classDef device fill:#eaf3fc,stroke:#0099ff,color:#12203a
    classDef shared fill:#fff3b0,stroke:#d4b200,color:#12203a
    class A,B,E device
    class C,D shared
```

The tracker spreadsheet is the source of truth. GitHub Pages hosts static app files; account, school, event, and check-in data are handled by each region's Apps Script and tracker. The device may retain a sign-in, a recent scholar page, and an unsent check-in. **Sign out** removes the saved page and pending check-in from that device. An offline check-in is sent when the scholar is signed in and opens START with a connection.

The app shell can open without a connection after it has been cached. Live pages, brain breaks, and new data still need internet access.

## Install START

Open the [Region XI app](https://feliciety02.github.io/START-Davao-Chapter/) in a browser. The app shows an install reminder when the browser supports one; installation always needs the person's approval.

| Device | Steps |
| --- | --- |
| Android · Chrome | Tap **Install** in START, or use **⋮ → Install app**. |
| iPhone · Safari | Tap **Share → Add to Home Screen**. |
| Computer · Chrome or Edge | Use the install icon in the address bar. |

START is a web app. This repository does not contain an Android APK or an iOS App Store build.

## Explore the repository

| Area | Start here | What you will find |
| --- | --- | --- |
| Installable app | [index.html](index.html) · [config.js](config.js) | Loading, region selection, offline check-in, and the region website frame |
| Install assets | [manifest.webmanifest](manifest.webmanifest) · [sw.js](sw.js) · [icons/](icons/) | App identity, service worker cache, and generated icons |
| Tracker website and backend | [apps-script/](apps-script/) | Portal UI, registration, check-ins, roles, events, email, and spreadsheet setup |
| Local checks | [apps-script/tests/](apps-script/tests/) | Node tests for portal and tracker behavior |
| Screenshots | [phone capture archive](docs/screenshots/portal-2026-11-16/README.md) · [app shell states](docs/screenshots/) | Installed app screens and shell states |

## Set up a region

The checkout already names Region XI in [config.js](config.js). For a new region or a fork, set up the tracker website and app address together:

1. In the region's bound Apps Script project, deploy the START website as a **Web app** with **Who has access: Anyone**. A website that requires a Google sign-in cannot load inside the app frame.
2. Put the deployment's `/exec` URL in the `regions` list in [config.js](config.js). With more than one region, START asks the visitor to choose one and remembers the selection.
3. Publish the root app files, including `.nojekyll` and `icons/`, through GitHub Pages: **Settings → Pages → Deploy from a branch → main → / (root)**.
4. In each tracker, open **START Tracker → 🌐 Website for members, Leads & you**. Enter the published app URL under **📱 START app address** and save it. This allows the website to appear inside the app.

The [Apps Script guide](apps-script/README.md) covers tracker setup, schema changes, and deployment details. Editing local `.gs` or `Web_App.html` files does **not** change the live bound Apps Script project.

## Updating and checking

- **App shell, configuration, or icons:** publish the changed root files to GitHub Pages. Installed devices load the updated shell when they next open START online. If the icon remains old, reinstalling the home screen app may be necessary.
- **Tracker website or server:** copy or sync the changed files from [apps-script/](apps-script/) into the bound project, run **Set up / fix the tracker** as the owner when schema or protections changed, then deploy a **new web-app version**. Keep the live project's complete `07_Look.gs` when syncing.
- **Icons:** run `powershell -NoProfile -ExecutionPolicy Bypass -File ./icons/build-icons.ps1`, then publish the generated files in `icons/` and the manifest.
- **Local checks:** run `node --test apps-script/tests/*.test.cjs` from the repository root.

Local tests cannot verify live Google Sheets permissions, email delivery, or Apps Script response time. After a deployment, test registration, sign-in, one check-in, and the installed app on real Android and iPhone devices.
