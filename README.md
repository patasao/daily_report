# Daily Operation Report (DOR)

Automated ETL pipeline that extracts hotel operational data from Guestline and Mews PMS, transforms it into standardised metrics, refreshes dynamic Excel dashboards, and distributes reports by email to stakeholders.

Covers **28 properties** (27 Guestline + 1 Mews) and runs daily at 09:30 via Windows Task Scheduler.

---

## Table of Contents

- [Architecture Overview](#architecture-overview)
- [Project Structure](#project-structure)
- [Prerequisites](#prerequisites)
- [Configuration](#configuration)
- [Running the Pipeline](#running-the-pipeline)
- [Data Flow](#data-flow)
- [Module Reference](#module-reference)
- [Logging](#logging)
- [Troubleshooting](#troubleshooting)

---

## Architecture Overview

```
Guestline PMS ──Selenium──► CSV/PDF ──────────────────┐
                                                       ▼
                                           SharePoint (raw files)
                                                       │
Mews PMS ──── Excel on SharePoint ──────────────────► │
                                                       ▼
                                         transform_data.py (Pandas)
                                                       │
                                                       ▼
                                    SharePoint /Report Data/ALL PY TRANS/
                                                       │
                                                       ▼
                                    refresh_excel_and_send.py (win32com)
                                                       │
                                              ┌────────┴────────┐
                                              ▼                  ▼
                                     Excel on SharePoint   Email (Graph API)
```

**Concurrency model:**
- Guestline **extraction** runs sequentially (shared Chrome profile avoids 2FA conflicts)
- Mews processing and all transform/send steps run concurrently via `ThreadPoolExecutor` (3 workers)

---

## Project Structure

```
Daily_Operation_Report/
├── main.py                        # Orchestration entry point
├── config.py                      # All paths, credentials, site list, date config
├── scheduled_main.bat             # Windows Task Scheduler wrapper
├── test.ipynb                     # Development / ad-hoc testing notebook
│
├── process/
│   ├── extract_data.py            # Selenium automation for Guestline reports
│   ├── transform_data.py          # Pandas data cleaning and aggregation
│   └── refresh_excel_and_send.py  # Excel refresh and email dispatch
│
├── services/
│   ├── logger.py                  # Timestamped log file setup
│   ├── ms_auth.py                 # Azure AD token acquisition (MSAL)
│   ├── sharepoint.py              # SharePoint Graph API read/upload helpers
│   └── email.py                   # Microsoft Graph API email sender
│
├── chromedriver/
│   └── chromedriver.exe           # Selenium WebDriver binary
│
└── log/                           # Runtime logs (auto-created, one file per run)
```

---

## Prerequisites

| Requirement | Notes |
|---|---|
| Python 3.11+ | Tested on 3.11 |
| Google Chrome | Version must match `chromedriver.exe` |
| Microsoft Excel | Required for win32com automation |
| Windows OS | win32com and Task Scheduler are Windows-only |

**Install Python dependencies:**

```powershell
pip install selenium pyotp pandas numpy requests msal openpyxl xlrd pywin32 Pillow python-dotenv python-dateutil
```

---

## Configuration

Create a `.env` file in the `Daily_Operation_Report/` directory with the following variables:

```env
# SharePoint paths
sp_dor=<relative path to DOR root in SharePoint>
sp_dor_root=<Graph API folder path for DOR root>
sp_dor_root_test=<Graph API folder path for test environment>

# Local paths
dor_download_dir=<local folder where Guestline reports are downloaded>

# Azure AD / Microsoft Graph
tenant_id=<Azure AD tenant ID>
client_id=<App registration client ID>
client_secret=<App registration client secret>
sp_site_id=<SharePoint site ID in Graph API format>

# Guestline credentials
gl_user=<Guestline username>
gl_pass=<Guestline password>
TOTP_SECRET=<Base32 TOTP secret for 2FA>

# Email
e_mail=<sender / test email address>
```

All other settings (site list, date calculations, folder structures) are defined in `config.py`.

---

## Running the Pipeline

**Scheduled (recommended):** Windows Task Scheduler runs `scheduled_main.bat` daily at 09:30.

**Manual run:**

```powershell
python main.py
```

There are no command-line arguments. All configuration is read from `config.py` and `.env`.

---

## Data Flow

### 1. Extract (Guestline only)

`extract_data.py` uses Selenium to log in to each Guestline site (with TOTP 2FA) and download up to 8 report types per site:

| Report | Format |
|---|---|
| Revenue By Market Segment (Actual) | CSV |
| Revenue By Market Segment (Forecast) | CSV |
| Period End Archive (EOD) | CSV + PDF |
| Availability | CSV |
| Actual & Forecast | PDF |
| Company Detail | PDF |
| Daily New Reservations | PDF |
| Cancellations | PDF |

Downloaded files are uploaded to SharePoint under `/Report Data/{Country}/{SiteID}/`.

Mews data arrives as Excel files uploaded manually to SharePoint; no extraction step is needed.

### 2. Transform

`transform_data.py` reads raw CSVs/Excel files from SharePoint and produces four aggregated datasets per site:

| Dataset | Contents |
|---|---|
| Actual Forecast | Rooms, revenue, sleepers, ADR by market segment and date |
| EOD | Financial totals (nett, tax, gross) by section and code |
| Availability | Room availability by date |
| Changes | Day-over-day variance versus the previous day's forecast |

Output CSVs are written to `SharePoint:/Report Data/ALL PY TRANS/{Country}/{SiteID}/` grouped by year.

### 3. Refresh & Send

`refresh_excel_and_send.py` opens each site's `{SiteID} Display.xlsm` workbook in an isolated Excel process, refreshes all Power Query connections, updates date dropdowns, then:

1. Captures the summary table as an inline PNG via the clipboard
2. Copies the refreshed workbook to SharePoint with a date-stamped filename
3. Sends an HTML email (inline image + SharePoint links) to the site's recipient list

---

## Module Reference

### `main.py`

| Function | Purpose |
|---|---|
| `_process_guestline()` | Sequential extraction followed by concurrent transform + send for all Guestline sites |
| `_process_mews()` | Concurrent transform + send for Mews sites |
| `_retry_site()` | Re-runs failed sites from their first failed step (up to 2 retries) |
| `_send_summary()` | Sends a consolidated HTML summary email to stakeholders |

### `process/extract_data.py`

Class `Guestline(SeleniumExtract)` — one instance per site.

Key methods: `_login()`, `_generate_actual_report()`, `_generate_forecast_report()`, `_generate_eod_report()`, `_generate_avai_report()`, `_generate_actual_and_forecast_report()`, `_generate_daily_new_reservations()`, `_generate_cancellation_report()`, `_generate_company_detail_report()`, `_move_files()`, `_skip_dates_copy()`.

### `process/transform_data.py`

| Class | PMS |
|---|---|
| `GuestlineTrans(TransformData)` | Guestline CSV sources |
| `MewsTrans(TransformData)` | Mews Excel sources |

Key methods: `data_files_check()`, `process_actual_forecast()`, `process_eod()`, `process_availability()`, `process_changes()`.

### `process/refresh_excel_and_send.py`

Class `RefreshAndSend` — one instance per site.

Key methods: `_open_excel()`, `_get_dates()`, `_set_default()`, `_snapshot()`, `_draft_email()`, `_get_email_list()`, `_save_file()`, `_send_email()`.

### `services/`

| Module | Purpose |
|---|---|
| `logger.py` | Creates a timestamped `.log` file and returns a configured `logging.Logger` |
| `ms_auth.py` | Acquires and caches Azure AD access tokens via MSAL client-credentials flow |
| `sharepoint.py` | `readSP()` and `uploadSP()` wrappers around the SharePoint Graph API |
| `email.py` | Sends email via the Microsoft Graph API `/sendMail` endpoint |

---

## Logging

Each run writes one log file per module to `./log/`:

```
log/
├── main_log_at_DDMMYYYYHHMMSS.log
├── extract_data_log_at_DDMMYYYYHHMMSS.log
├── transform_data_log_at_DDMMYYYYHHMMSS.log
└── refresh_excel_and_send_log_at_DDMMYYYYHHMMSS.log
```

Log level is `INFO`. Each line includes a timestamp, module name, and message. Errors include full tracebacks.

---

## Troubleshooting

**Chrome/Selenium fails to start**
- Ensure `chromedriver.exe` version matches the installed Chrome version.
- Check that `dor_download_dir` exists and Chrome has write permission to it.

**2FA / login loop**
- Verify `TOTP_SECRET` in `.env` is the correct Base32 secret.
- If the Chrome profile is corrupted, delete the profile folder and re-authenticate manually once.

**Excel refresh hangs**
- Kill any orphaned `EXCEL.EXE` processes before re-running.
- Confirm SharePoint is mounted and the Power Query data source paths resolve.

**Missing transformed files**
- Run `transform_data.py` in isolation via `test.ipynb` for the affected site.
- Check that the raw CSVs were uploaded to SharePoint in the expected folder structure.

**Graph API / SharePoint 401 errors**
- Verify `tenant_id`, `client_id`, `client_secret`, and `sp_site_id` in `.env`.
- Check that the Azure AD app registration has the required Graph API permissions (`Sites.ReadWrite.All`, `Mail.Send`).
