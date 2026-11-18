# Operations runbook

## Is it up?
- Open the app. The **Server status** panel at the bottom says whether LibreOffice is ready.
- Liveness: `GET <app-url>/_stcore/health` returns `ok`.
- Streamlit Cloud: **Manage app** (bottom-right of the app) shows status, logs and a **Reboot** button.

## Logs
- One JSON line per event on stdout: Streamlit Cloud → **Manage app → Logs**; self-hosted → the
  terminal/service output, plus `LOG_FILE` if set (rotated by size).
- Events: `app.start` (with `soffice` true/false), `generate.ok`, `generate.rejected`, `generate.failed`,
  `recalc.fallback`, `recalc.verify_failed`, `access.denied`.
- Fields: `rid` (the **Reference** shown to the user with errors), `status`, `error_class`, sizes, counts,
  durations (`ms_*`). Logs never contain item names, work centres, quantities or file names.
- To trace a user's error, search the logs for the Reference code they saw.

## Common problems
| Symptom (what the user sees) | Cause | Fix |
|---|---|---|
| Warning "could not be pre-calculated ... LibreOffice not found" | LibreOffice missing | Cloud: check `packages.txt` contains `libreoffice-calc-nogui`, then reboot the app. Self-hosted: install LibreOffice or set `SOFFICE_PATH`. The file is still valid; Excel calculates on open. |
| Warning "... timed out" | Server slow or overloaded | Retry. If frequent, raise `RECALC_TIMEOUT_S` (e.g. 120). |
| Warning "... recalculation busy" | Another generation held LibreOffice for the whole timeout | Retry; on a bigger server raise `RECALC_CONCURRENCY`. |
| "larger than the 10 MB limit" (413) | Upload too big | Remove extra sheets/images from the master, or raise `MAX_UPLOAD_MB` **and** `server.maxUploadSize`. |
| Upload box refuses the file before Generate | Over `server.maxUploadSize` or not `.xlsx` | As above; save as `.xlsx`. |
| "Too many generations in the last hour" (429) | Rate limit | Wait, or raise `RATE_LIMIT_PER_HOUR`. |
| "not a valid .xlsx" / "damaged" | Renamed file, corrupted or partial download | Open the master in Excel, **Save As → Excel Workbook (.xlsx)**, upload again. |
| "password-protected/encrypted or old .xls" | Encrypted or legacy file | Remove the password (File → Info → Protect Workbook) and save as `.xlsx`. |
| "contains macros" | `.xlsm` or VBA inside | Save as a macro-free `.xlsx`. |
| "links to other files" | External workbook links | Excel: Data → Edit Links → Break Link, save, upload. |
| "Could not find the header row" / "missing: ..." | Headers renamed/moved | Row 3 (or lower, within the first 10 rows) must have `Item Name`, `Work Center`, `Week Plan Qty` in A, B, C. |
| "header row is row 1/2" | Title rows deleted | Insert rows so the header is on row 3. |
| "Access denied (403)" | Client IP not in `ALLOWED_IPS` | Add the IP/CIDR, or check `TRUSTED_PROXY_HOPS` if behind a proxy. |
| "The app is misconfigured: X ..." | Bad setting | Fix the named variable (README table) and restart/reboot. |
| "Something went wrong ... Reference: abc123" | Unexpected bug | Find `abc123` in the logs (`error_class`), keep the master for reproduction. |

## Restart / upgrade
- Cloud: **Manage app → Reboot**. Upgrades: change code or pinned versions in `requirements.txt`,
  run `scripts/test.ps1` (or `make test`) locally, push; the app redeploys.
- Self-hosted: stop the service, `git pull`, `pip install -r requirements.txt`, run the tests, start it.
