# Offer Importer

Version: 0.99.05.06.26

Browser-based Excel importer for creating separate upload-template workbooks from Secret Escapes offer forms, rate-update workbooks, and paired Rates/Allocation tabs.

## Use

Open the GitHub Pages site, then:

1. Choose the Shared Template file.
2. Choose the source workbook.
3. Click **Generate Excel**.
4. If the browser asks, allow multiple downloads.

All Excel processing runs locally in the browser. The selected Excel files are not uploaded to GitHub or any server.

## Supported Source Workbooks

- Submission offer forms.
- Older no-flight submission forms with `Departure Date` rates/allocation tabs.
- German offer forms with fields like `PRO ZIMMER Angebotsrate`, `Reise-Startdatum`, and `SHARED Allocation`.
- Rate-update workbooks with sheets like `Rates (1)` / `Allocation (1)` or `5 Nts Departures` / `5 Nts Summary`.
- `.xls`, `.xlsx`, and `.xlsm` workbooks with paired Rates and Allocation tabs, such as `Superior rates sheet` / `Superior alloc.tab`.
- Slash-separated airport codes such as `FCO/CIA`, which are split into separate upload rows.

## Deployment

This repo is configured for GitHub Pages from the `main` branch root.

When changes are pushed to `main`, GitHub Pages serves the updated static files automatically.

## Changes

See `CHANGELOG.md` for the short team-facing change list.
