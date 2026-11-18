# MDK Weekly Planning Generator

Upload the weekly production master (.xlsx), get back the day-wise planning workbook
(`<master name>_Planning.xlsx`): the master sheet unchanged, **Monday–Saturday** planning tabs,
**Consolidated**, and a hidden **_Lists** sheet. All planning logic lives in the downloaded file as
formulas, dropdowns and conditional formatting (no macros). The layout is the client-approved
template in `planning_template_reference.py`; `planner/builder.py` is a faithful port of it.

Nothing uploaded is stored: the file is processed in memory, LibreOffice temp files are deleted
when each generation finishes, and the result lives only in the browser session until the next upload.

## How it works
1. **Validate** (`planner/validate.py`, `planner/reader.py`): real .xlsx by content, no macros, no
   encryption, no external links, zip-bomb and size limits, header row found, warnings per master row.
2. **Build** (`planner/builder.py`): the approved template. Text from the master is always written
   as text, so names starting with `= + - @` can never become formulas.
3. **Recalculate** (`planner/recalc.py`): headless LibreOffice saves calculated values into the file so
   it shows numbers in previews (phone, Outlook, Google Sheets). If LibreOffice is missing, fails or
   times out, you still get the workbook (Excel calculates on open) and a warning.
4. **Verify** the recalculated file (no formula errors; dropdowns, rules, filters, freeze panes and the
   hidden sheet still there), then offer it for download.

## Hosting: Streamlit Community Cloud
The app is a Streamlit app (`streamlit_app.py`), chosen for Streamlit Community Cloud hosting.
**The app has no login.** Anyone with the URL can use it unless you restrict it:

1. Put this folder in a **private** GitHub repository. Real masters (`*.xlsx`) are git-ignored on
   purpose: they are business data.
2. On <https://share.streamlit.io>, choose **Create app**, select the repo/branch, main file
   `streamlit_app.py`, and under **Advanced settings** pick **Python 3.12**.
3. `packages.txt` makes the platform install LibreOffice (`libreoffice-calc-nogui`); `requirements.txt`
   installs the pinned Python packages. The first build takes a few minutes.
4. **Restrict access:** in the app's **Settings → Sharing**, make the app private and invite only the
   planner's e-mail address. This is the platform's own viewer list; the app itself has no accounts.
5. Optional settings go in **Settings → Secrets** as TOML, e.g. `RATE_LIMIT_PER_HOUR = "30"`
   (root-level secrets become environment variables).

To update: push to the branch; the app redeploys automatically. Logs are under **Manage app**.

Differences from a self-hosted server, because the platform controls them: HTTPS and security
headers are set by Streamlit Cloud; health is at `/_stcore/health`; the app sleeps after a period
without visits (the first visit then takes ~30 s to wake); about 1 GB RAM, so keep
`RECALC_CONCURRENCY=1`.

### Running it on your own machine or server instead
Same code, no changes. Install Python 3.12 and LibreOffice, then:

```powershell
# Windows
winget install -e --id Python.Python.3.12
winget install -e --id TheDocumentFoundation.LibreOffice
scripts\run.cmd                   # creates .venv on first run, serves http://localhost:8501
# (if PowerShell says "running scripts is disabled", use run.cmd rather than run.ps1)
```
```sh
# Debian/Ubuntu
sudo apt install python3.12 python3.12-venv libreoffice-calc-nogui
./scripts/run.sh                  # or: make install run
```
If you serve it to others this way, keep it on the internal network or behind a VPN/firewall and set
`ALLOWED_IPS`. To keep it running, use a service manager (systemd on Linux, NSSM or Task Scheduler
"At startup" on Windows) running `python -m streamlit run streamlit_app.py --server.port 8501` from
the project folder.

## Configuration
Environment variables, a `.env` file (copy `.env.example`), or Streamlit Cloud Secrets.
Invalid values stop the app with a message naming the setting.

| Variable | Default | Meaning |
|---|---|---|
| `MAX_UPLOAD_MB` | `10` | Largest accepted upload. Keep `server.maxUploadSize` in `.streamlit/config.toml` the same. |
| `MAX_UNCOMPRESSED_MB` | `100` | Largest total uncompressed size of the .xlsx (zip-bomb guard). |
| `MAX_ZIP_ENTRIES` | `1000` | Most internal parts an .xlsx may have. |
| `MAX_ROWS` | `20000` | Most data rows on the master sheet. |
| `MAX_COLS` | `50` | Most used columns on the master sheet. |
| `RECALC_ENABLED` | `true` | Pre-calculate values with LibreOffice. `false` = Excel calculates on open. |
| `SOFFICE_PATH` | *(auto)* | Full path to `soffice` if it is not on PATH / in the default folder. |
| `RECALC_TIMEOUT_S` | `60` | Hard time limit for one recalculation. |
| `RECALC_CONCURRENCY` | `1` | Recalculations allowed at the same time. |
| `RATE_LIMIT_PER_HOUR` | `60` | Generations per hour per client (`0` = unlimited). |
| `ALLOWED_IPS` | *(empty = all)* | Comma-separated IPs/CIDRs allowed to use the app. Safety net, not a login. |
| `TRUSTED_PROXY_HOPS` | `1` | Proxies in front of the app whose `X-Forwarded-For` is trusted. |
| `LOG_LEVEL` | `INFO` | `DEBUG`, `INFO`, `WARNING`, `ERROR`. |
| `LOG_FILE` | *(empty)* | Also write logs to this file, rotated. Empty = stdout only. |
| `LOG_MAX_BYTES` | `5242880` | Rotate the log file at this size. |
| `LOG_BACKUPS` | `5` | Rotated log files kept. |

## Development
```powershell
scripts\test.cmd            # Windows: ruff + mypy + pytest      (Linux: make test)
```
Tests that need the real masters look for `*Week45*.xlsx`, `*Week46*.xlsx`, `*Week47*.xlsx` in the
project folder (or `tests/fixtures/`) and skip if they are missing; LibreOffice tests skip if it is not
installed.

## Troubleshooting
See [OPERATIONS.md](OPERATIONS.md).

## Notes on the approved template (not changed)
- The master header must be on row 3 or lower: rows 1–2 of the day tabs hold the title and instructions.
- Headers must be in columns A, B, C (`Item Name`, `Work Center`, `Week Plan Qty`); the formulas use those columns.
- A text quantity counts as 0 in Consolidated, but the item's **Weekly plan** on a day tab shows `#VALUE!`
  when picked, until the master is corrected. The app warns with the row number.
- Duplicate (work centre, item) rows share one day-tab row whose weekly plan is their sum; both master
  rows show the same day quantities in Consolidated, so its TOTAL counts them twice. The app warns.
