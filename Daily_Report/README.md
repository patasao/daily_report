# Daily Operation Report (DOR) Pipeline

## Overview

An end-to-end automated pipeline that extracts, transforms, and distributes **28 hotel property reports** on a daily schedule — replacing a previously manual, multi-hour process. Reports cover room revenue actuals, rolling forecasts, end-of-day accounting, availability, reservations, and cancellations.

---

## Business Impact

| Metric | Before Automation | After Automation |
|--------|-------------------|------------------|
| Reports distributed daily | 28 (manually) | 28 (automated) |
| Manual effort per cycle | ~4 hours | Near-zero |
| Distribution method | Manual email | Microsoft Graph API |
| Data freshness | Dependent on availability | Available by morning each day |
| Sites covered | 28 hotel properties | 28 hotel properties |
| PMS integrations | 2 (manual) | 2 (automated) |

- Eliminates daily manual extraction from two hotel Property Management Systems (PMS)
- Each property's stakeholders receive a personalised email with an inline summary table and deep links to SharePoint, distributed via the Microsoft Graph API `sendMail` endpoint
- The retry logic handles transient failures per-site without re-running the full batch, reducing operator intervention

---

## Architecture

```
┌──────────────────────────────────────────────────────────────┐
│                         main.py                              │
│  ┌─────────────────────────┐  ┌─────────────────────────┐   │
│  │  On-Premise PMS Pipeline │  │  Cloud PMS Pipeline      │   │
│  │  (sequential extract,   │  │  (fully concurrent —     │   │
│  │   concurrent transform) │  │   no extract step)       │   │
│  └──────────┬──────────────┘  └──────────┬──────────────┘   │
│             │                            │                    │
│             ▼                            ▼                    │
│  ┌──────────────────┐        ┌──────────────────────┐        │
│  │  Extract         │        │  Transform           │        │
│  │  (Selenium +     │        │  (Pandas / NumPy)    │        │
│  │   TOTP 2FA)      │        └──────────────────────┘        │
│  └──────────────────┘                   │                    │
│             │                           ▼                    │
│             ▼                 ┌──────────────────────┐       │
│     SharePoint (Graph API)    │  Refresh & Send      │       │
│     CSV / PDF upload          │  (win32com + Graph   │       │
│                               │   API sendMail)      │       │
│                               └──────────────────────┘       │
└──────────────────────────────────────────────────────────────┘
```

---

## Pipeline Stages

### 1. Extract — `process/extract_data.py`
- Selenium automates login (including TOTP 2FA) to the on-premise PMS web portal
- A shared Chrome browser profile persists the session across all 27 sites in a single run, so the TOTP prompt only fires once per 30-second window
- Eight report types are downloaded per property: Actual Revenue, Forecast, Actual & Forecast (PDF), Availability, EOD (Period End), Company Detail, Daily New Reservations, and Cancellations
- Downloaded files are uploaded to SharePoint via Graph API and the local staging directory is cleaned up

### 2. Transform — `process/transform_data.py`
- Two concrete subclasses (`MewsTrans`, `PmsTrans`) handle the schema differences between the cloud PMS (wide-format Excel pivot) and the on-premise PMS (CSV exports)
- Both are normalised into a shared long-format schema: `(site_id, year, as_of_date, stay_date, market_segment, rooms, revenue_nett, sleepers, ADR)` — enabling a single Power Query connection to serve all downstream reports
- A `process_changes` step computes day-over-day pickup deltas by joining today's snapshot against yesterday's and the prior-year equivalent, used to power the Pickup and Variance tabs in Excel

### 3. Refresh & Send — `process/refresh_excel_and_send.py`
- Each property has an Excel workbook with live Power Query connections to SharePoint
- `win32com.DispatchEx` opens an isolated Excel instance per workbook (supports concurrent runs)
- After `RefreshAll()` and date-filter updates, a versioned copy is saved back to SharePoint
- The summary table from the `ReportConfig` sheet is captured as an inline PNG (via PIL + clipboard) and embedded directly in the HTML email body
- Emails are dispatched via the Microsoft Graph API `sendMail` endpoint; a date-guard prevents sending if the data has not been updated

---

## Technical Stack

| Component | Technology |
|-----------|------------|
| Web automation | Selenium 4, ChromeDriver |
| 2FA | pyotp (TOTP) |
| Data transformation | Pandas, NumPy |
| Excel automation | win32com (pywin32) |
| Cloud storage & email | Microsoft Graph API (SharePoint, Outlook) |
| Authentication | OAuth 2.0 client credentials via MSAL |
| Concurrency | `concurrent.futures.ThreadPoolExecutor` |
| Environment config | python-dotenv |

---

## Configuration

All credentials and environment-specific paths are loaded from a `.env` file. No secrets are committed to source control.

| Variable | Description |
|----------|-------------|
| `tenant_id` | Azure AD tenant ID |
| `client_id` | App registration client ID |
| `client_secret` | App registration client secret |
| `sp_site_id` | SharePoint site ID |
| `sp_dor_root` | SharePoint root path for DOR folder |
| `PMS_BASE_URL` | On-premise PMS portal URL |
| `pms_user` / `pms_pass` | PMS portal credentials |
| `TOTP_SECRET` | Base32 TOTP secret for 2FA |
| `SP_BASE_URL` | SharePoint domain for email folder links |
| `SP_SITE_PATH` | SharePoint site path component |
| `SP_REPORTS_FOLDER` | Reports folder name in SharePoint |
| `e_mail` | Sender / summary recipient email |
| `BACKDATE_SITE_ID` | Site ID with non-standard fiscal cut-off |
| `dor_download_dir` | Local staging directory for downloads |

---

## Data Standardisation Note

The two PMS platforms export revenue data in fundamentally different formats:

- **Cloud PMS**: Nights data as a wide date-column pivot table in Excel, joined to a separate Reservations sheet for market-segment codes
- **On-Premise PMS**: Separate CSV files per report type (Actual, Forecast), with different column schemas and date representations

Both are standardised into the same long-format table, enabling uniform reporting across all 28 properties from a single data model.

---

## Project Structure

```
Daily_Operation_Report/
├── main.py                          # Entry point — orchestrates all pipelines
├── config.py                        # Centralised config, paths, and date helpers
├── process/
│   ├── extract_data.py              # Selenium web automation (on-premise PMS)
│   ├── transform_data.py            # Pandas/NumPy data cleaning and normalisation
│   └── refresh_excel_and_send.py   # Excel refresh + Graph API email distribution
└── services/
    ├── ms_auth.py                   # MSAL token manager (OAuth 2.0 client credentials)
    ├── sharepoint.py                # Graph API SharePoint read/upload helpers
    ├── email.py                     # Graph API sendMail wrapper
    └── logger.py                    # Rotating file logger setup
```
