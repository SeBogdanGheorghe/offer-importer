# Offer Importer

Version: 1.10.14.10.2026

Browser-based Excel importer for creating separate upload-template workbooks from Secret Escapes offer forms, rate-update workbooks, and paired Rates/Allocation tabs.

## Where the tool lives now

The GitHub Pages site redirects to the Google Apps Script version, which adds a
**History** tab and archives every uploaded workbook to Drive as a Google Sheet:

https://script.google.com/a/macros/secretescapes.com/s/AKfycby2AuOyTSJm1hnUlAXLk_sVGGQ8kS9bQZLX1s-cvNUYxgOGqQMe5LbAyo8VYVYqumaEdQ/exec

It is restricted to `secretescapes.com` accounts, so it asks for a Google login.

The original browser-only tool is still here, unchanged and with no login, at
[`standalone.html`](standalone.html) —
https://sebogdangheorghe.github.io/offer-importer/standalone.html

## Use

1. Drop or choose the source workbook.
2. The tool starts automatically.
3. If the browser asks, allow multiple downloads.

All Excel processing runs locally in the browser. The selected Excel files are not uploaded to GitHub or any server.
The upload template is built into the browser tool, so no separate template file is required.
The success animation uses the vendored `canvas-confetti` library; its license is included in `vendor/`.
The Google Sheet link import uses public/shared Google Sheets only and does not store API keys or credentials.

## Supported Source Workbooks

- Submission offer forms.
- Older no-flight submission forms with `Departure Date` rates/allocation tabs.
- German offer forms with fields like `PRO ZIMMER Angebotsrate`, `Reise-Startdatum`, and `SHARED Allocation`.
- Rate-update workbooks with sheets like `Rates (1)` / `Allocation (1)` or `5 Nts Departures` / `5 Nts Summary`.
- `.xls`, `.xlsx`, and `.xlsm` workbooks with paired Rates and Allocation tabs, such as `Superior rates sheet` / `Superior alloc.tab`.
- Cabin-pair rate updates such as `Eastern - Inside Cabin` / `Eastern Inside -Allo`; empty cabin groups are skipped.
- Slash-separated airport codes such as `FCO/CIA`, which are split into separate upload rows.

## Layout

| Path | What it is |
|---|---|
| `index.html` | Redirect to the Apps Script web app |
| `standalone.html` | The original browser-only tool |
| `src/` | `importer.js` (Excel logic) and `browser.js` |
| `offer-importer.bundle.js` | Built bundle used by `standalone.html` |
| `apps-script/` | Source of the Apps Script web app, pushed with `clasp` |

The Apps Script version reuses `src/importer.js` unchanged — `apps-script/Importer.html`
is that file with only the ES module import/export lines removed, so a fix to the
Excel logic applies to both. See `apps-script/README.md`.

## Deployment

This repo is configured for GitHub Pages from the `main` branch root.

When changes are pushed to `main`, GitHub Pages serves the updated static files automatically.

The Apps Script app deploys separately from `apps-script/` with `clasp push` and
`clasp update-deployment`; see `apps-script/README.md` for the exact commands.

## Changes

See `CHANGELOG.md` for the short team-facing change list.
