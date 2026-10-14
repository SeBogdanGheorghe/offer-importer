# Offer Importer — Apps Script edition

Version `2.07.14.10.2026`

The GitHub Pages Offer Importer, rebuilt as a Google Apps Script web app with an
added **History** tab. Excel parsing and output generation are byte-for-byte the
same logic as the public tool — `Importer.html` is `src/importer.js` from the
repo with only the ES module import/export lines removed.

**Web app:** https://script.google.com/a/macros/secretescapes.com/s/AKfycby2AuOyTSJm1hnUlAXLk_sVGGQ8kS9bQZLX1s-cvNUYxgOGqQMe5LbAyo8VYVYqumaEdQ/exec

**Script project:** https://script.google.com/d/1f3HkYZ--ce-Mnf1qWA_pgxUTXxN87LWCIVvrCnRW_SyW1baPG3DcN9yK/edit

## Files

| File | What it is |
|---|---|
| `Code.gs` | Server: Drive archiving, History read/filter, Google Sheet export fetch |
| `Index.html` | Page shell — header, the two tabs, both panels |
| `Styles.html` | CSS (the public tool's stylesheet + tab/history rules) |
| `Importer.html` | `src/importer.js`, unmodified apart from module syntax |
| `App.html` | Client logic — upload, downloads, archiving, history UI |
| `selfcheck.js` | `node selfcheck.cjs` — checks search, 3PP detection, date ranges. Not pushed (`.claspignore`) |

## What happens on upload

1. The workbook is parsed in the browser and one Excel file per populated
   Rates/Allocation pair downloads immediately — unchanged behaviour.
2. A second status line shows **"File being converted to Google Sheet and sent
   to history"** while the workbook is uploaded to Drive, converted to Google
   Sheets format, and a History row is written.
3. Generated files also stay listed as clickable links, in case the browser
   blocks automatic multi-file downloads inside the Apps Script iframe.

## File type

`3PP form` when the workbook has **both** an `Offer Analysis` and an
`Offer Details` tab. Everything else is logged as `update form`. Rows written
before the rename still say `EF` in the spreadsheet; they are shown and
filtered as `3PP form`.

## Hotel names

The form lists hotels across one row — a `Hotel name` label on the left, then
one hotel per column. All of them are read off that row (however many there
are), stored comma separated in `Hotel_Names`, shown in the History table and
included in the search text.

Defaults to the last 7 days; 14 / 30 days or a custom date range are available,
plus a type filter and a search box. Search is token based and ignores
punctuation, so `Amsterdam & Berlin` finds
`RATE UPDATE_Amsterdam & Berlin with European Sleeper 13.10.2026`. It also
matches on the uploader's email.

## Storage

- Archived sheets: https://drive.google.com/drive/folders/1JKUqS-loYgZjvJULYY4xQmGnoIuVMrNt
- Log rows: https://docs.google.com/spreadsheets/d/1Iw4_A5SnVXxgxkTVrVkkcUWGOMyqMo0cSflhEvnumkw/edit

Columns: `Timestamp`, `User_Email`, `File_Type`, `File_Name`, `Sheet_URL`,
`Outputs`, `Hotel_Names`.
Headers are written automatically if the sheet is empty.

## Log spreadsheet link

Tapping the version pill three times (within 2 seconds) reveals a link to the
raw log spreadsheet. The URL is returned by the server only for the addresses
in `LOG_VIEWERS` in `Code.gs` — currently `b.gheorghe@secretescapes.com`. For
anyone else the taps do nothing and the URL never reaches the page.

## Access

The web app runs **as the person using it** and is open to anyone at
`secretescapes.com`. Each colleague authorises once on first open, and needs
edit access to the Drive folder and the log spreadsheet — their own email is
what lands in History.

## Deploying a change

Update the existing deployment in place so the shared URL never changes:

```bash
clasp push --force
clasp update-deployment -V <newVersion> -d "Offer Importer <version>" AKfycby2AuOyTSJm1hnUlAXLk_sVGGQ8kS9bQZLX1s-cvNUYxgOGqQMe5LbAyo8VYVYqumaEdQ
```

`clasp create-deployment` would mint a *new* URL, so only use it for a
deliberately separate link. `clasp list-deployments` shows the IDs, and the
`@HEAD` deployment always serves the latest push for testing.

## Notes

- Output file names are de-duplicated before download: two pairs can share a
  label (`5 Nts` twice), which would otherwise make two downloads collide on
  one name. Repeats get ` (2)`, ` (3)` before the extension.
- The download anchor sets `target="_self"` on purpose — the page's
  `<base target="_top">` would otherwise turn the click into a top-frame
  navigation and silently cancel the download inside the Apps Script iframe.
