# Vietlott Auto Update Bot

This repository contains a small automation that fetches the latest Mega 6/45 result and appends it to an Excel workbook.

## Project status

The repository is intended to be run by GitHub Actions on a schedule and on push.

## Root cause found

The original workflow was using the old Vietlott JSON endpoint:

- https://vietlott.vn/api/front/get-result-mega645

That endpoint stopped returning the expected JSON payload. The site now serves HTML or a different page structure, so the script exited before writing to the Excel workbook. This is why the GitHub Action appeared to run without updating the Excel file.

## Fix implemented

The script now:

- uses the Minh Ngọc Vietlott page as the primary accessible source
- tries multiple Mega 6/45 URLs as fallbacks
- supports HTML parsing of the current Vietlott result/history pages
- avoids duplicate inserts when the same draw is already present in the workbook
- logs the actual response type and the fallback behavior

The workflow was also adjusted to only push when the Excel file actually changed.

## Current GitHub Actions issue

The latest scheduled run completed on September 13, 2026, but its logs show
that all configured Vietlott URLs returned `403 Forbidden` from the
GitHub-hosted runner. The previous script treated that fetch failure as a
normal result and exited with status 0, so GitHub displayed a successful run
even though no workbook update was possible. The requested Minh Ngọc page is
now the first source and its Mega 6/45 HTML block is parsed using the stable
`DT6X45_*` element IDs.

The script now raises an error when every source fails. This makes the Action
correctly show `failure` instead of silently reporting success. The next fix
should add and test a permitted alternate data source rather than writing
guessed lottery results.

The September 16 run successfully parsed draw `#01563`, but the workbook
already contained `#01563` with an old July date. The duplicate check has been
updated to repair that stale row when the draw ID matches but the date differs.

## Files

- `vietlott.py` — main scraper and Excel updater
- `.github/workflows/python-app.yml` — scheduled workflow and commit flow
- `Vietlott_Mega_645_Full_Results.xlsx` — generated workbook with the historical results
- `tests/test_vietlott.py` — parser regression checks

## How to continue from here

1. Push the latest changes to GitHub.
2. Trigger the workflow manually from the Actions tab or wait for the scheduled run.
3. If Vietlott changes the page layout again, inspect the HTML on the Mega 6/45 history page and update the parser regex patterns in `parse_mega645_html`.
4. Keep the workbook as the source of truth for duplicate detection.
5. Check the Action log for `403 Forbidden` if the workbook stops changing.

## Notes for the next session

- The old API contract is no longer reliable.
- The current fallback is the Vietlott Mega 6/45 history page, not the JSON endpoint.
- The workbook must be updated only when a new draw id appears.

## Local run

From the repo root:

```bash
python vietlott.py
```

If Python is not available on PATH in Windows, use:

```powershell
py -3 vietlott.py
```

## SQLite data store

`vietlott.db` is the incremental source of truth. A normal run stores or
updates the latest draw. To crawl the available history into SQLite once, run:

```powershell
py -3 vietlott.py --backfill
```

The database keeps every fetched draw by `draw_id`; rerunning the backfill is
safe because existing records are updated rather than duplicated.
