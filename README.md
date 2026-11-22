# Daily Report Automation

A portfolio adaptation of a daily operations reporting pipeline. It demonstrates a multi-source ETL workflow: scheduled property-management exports are collected, normalized into reporting tables, used to refresh Excel dashboards, and distributed through Microsoft Graph.

**Privacy:** every property identifier and every row in `demo_data/normalized_daily_metrics.csv` is synthetic. No private company identifiers, customer or employee data, tenant details, or live operational data are included. The labels `portal_pms` and `cloud_pms` describe the two generic connector workflows. Credentials and tenant-specific paths are read from a local `.env` file and are blank by default.

## What the pipeline demonstrates

- Selenium-based report extraction for a legacy PMS, including retry handling and two-factor authentication hooks.
- Processing of pre-exported Excel data from a cloud PMS.
- Pandas transformations that normalize revenue, room, availability, and end-of-day data.
- Concurrent per-site processing, step-aware retries, and run summaries.
- SharePoint file operations and Excel refresh/email delivery through Microsoft Graph.
- Process cleanup and timestamped logging for long-running Windows automation.

## Data flow

```text
PMS exports → extract/load → normalize and aggregate → SharePoint
                                           ↓
                               Excel dashboard refresh
                                           ↓
                                  email summary
```

`demo_data/normalized_daily_metrics.csv` is a small, fabricated example of the normalized output. The live connectors require authorized accounts, tenant configuration, and report templates that are not included here. The configured sample site IDs (`DEMO_PORTAL_001`, `DEMO_CLOUD_001`) and default `example.invalid` portal URL are intentionally non-operational.

## Project layout

```text
.
├── config.py                  # Environment-based configuration with fake site IDs
├── main.py                    # CLI entry point for either integration
├── pipeline.py                # Per-system orchestration, retries, and summaries
├── main_portal_pms.py         # Portal-based workflow entry point
├── main_cloud_pms.py         # Export-based workflow entry point
├── process/
│   ├── extract_data.py        # Browser-based report extraction
│   ├── transform_data.py      # Data normalization and aggregation
│   └── refresh_excel_and_send.py
├── services/
│   ├── ms_auth.py             # Microsoft identity token acquisition
│   ├── sharepoint.py          # Microsoft Graph file operations
│   ├── email.py               # Microsoft Graph email delivery
│   ├── logger.py
│   └── process_guard.py
└── demo_data/
    └── normalized_daily_metrics.csv
```

## Running the integration

The sample CSV is a static portfolio fixture. The integration entry points call external PMS, SharePoint, Excel, and email services; they require valid configuration and authorized accounts. After setting up an environment you control, run `python main.py portal_pms` or `python main.py cloud_pms`. A fresh clone has no credentials and uses fake site IDs.

## Local configuration

Copy `.env.example` to `.env` only when preparing a deployment against systems you are authorized to access. Keep real credentials, tenant IDs, site IDs, and recipient addresses out of source control. The included configuration defaults to test mode and contains only fake site IDs.

The Windows-specific integration path also requires Chrome, Excel, and the matching Python dependencies listed in `requirements.txt`. This showcase repository does not include a browser driver binary, company templates, scheduled-task definitions, or real report exports.

## Synthetic data fields

| Field | Meaning |
|---|---|
| `site_id` | Fictional sample site (`DEMO_PORTAL_001` or `DEMO_CLOUD_001`) |
| `as_of` | Date the report snapshot was produced |
| `stay_date` | Date represented by the metric row |
| `market_segment` | Synthetic segment label |
| `rooms_sold`, `revenue_net`, `sleepers` | Fabricated operating measures |
| `adr` | Fabricated average daily rate |

All values are illustrative and should not be interpreted as real business performance.
