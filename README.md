<p align="center">
  <img src="docs/readme-banner.svg" width="1200" alt="START Scholar Portal banner showing the star mascot, a weekly check-in, Journey progress, and school updates">
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

These screenshots show **the app shell's own screens**, including loading, installation, and offline states. The signed-in START website fills the screen after loading. The captures are illustrative and may differ from the latest local UI; select an image to open it at full size.

<table>
  <tr>
    <td align="center"><a href="docs/screenshots/01-loading.png"><img src="docs/screenshots/01-loading.png" width="190" alt="START loading screen on a phone"></a><br><strong>1 · Opening</strong><br><sub>Loads the selected region</sub></td>
    <td align="center"><a href="docs/screenshots/02-region-picker.png"><img src="docs/screenshots/02-region-picker.png" width="190" alt="Region selection screen"></a><br><strong>2 · Region picker</strong><br><sub>Appears when multiple regions are configured</sub></td>
    <td align="center"><a href="docs/screenshots/03-install-iphone.png"><img src="docs/screenshots/03-install-iphone.png" width="190" alt="Instructions for adding START to an iPhone home screen"></a><br><strong>3 · Install hint</strong><br><sub>Phone-specific home screen steps</sub></td>
  </tr>
  <tr>
    <td align="center"><a href="docs/screenshots/04-offline.png"><img src="docs/screenshots/04-offline.png" width="190" alt="START offline screen"></a><br><strong>4 · Offline</strong><br><sub>Recent scholar page and retry action</sub></td>
    <td align="center"><a href="docs/screenshots/05-not-set-up.png"><img src="docs/screenshots/05-not-set-up.png" width="190" alt="Message shown when no region website is configured"></a><br><strong>5 · Setup needed</strong><br><sub>Shown when the region URL is missing</sub></td>
    <td align="center"><a href="docs/screenshots/06-desktop.png"><img src="docs/screenshots/06-desktop.png" width="240" alt="START loading screen on a desktop browser"></a><br><strong>6 · Desktop</strong><br><sub>The same app shell on a larger screen</sub></td>
  </tr>
</table>

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
| Screenshots | [docs/screenshots/](docs/screenshots/) | App shell captures used in the tour |

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
