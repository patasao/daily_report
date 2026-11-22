"""Configuration for the portfolio-safe daily reporting example.

All site identifiers are synthetic. Credentials and tenant-specific locations are
read from an optional local .env file and are intentionally blank by default.
"""
from pathlib import Path
import os
import calendar
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from dotenv import load_dotenv

load_dotenv()

# Test mode is the default so browser sessions stay visible and report refreshes
# are skipped unless a local operator explicitly changes this setting.
env_status = os.getenv("ENV_STATUS", "test").strip().lower()
project_root = Path(__file__).resolve().parent
home_dir = Path.home()

def _env(name: str) -> str:
    return os.getenv(name, "").strip()

paths = {
    "project_root": project_root,
    "home_dir": home_dir,
    "log_dir": project_root / "log",
    "sp_report_dir": home_dir / _env("SHAREPOINT_SYNC_DIR"),
    "download_dir": home_dir / os.getenv("REPORT_DOWNLOAD_DIR", "Downloads/daily-report-demo"),
}

# Generic placeholders only; populate locally with credentials for systems you own.
tenant_id = _env("AZURE_TENANT_ID")
credentials = {
    "tenant_id": tenant_id,
    "client_id": _env("AZURE_CLIENT_ID"),
    "client_secret": _env("AZURE_CLIENT_SECRET"),
    "sp_site_id": _env("SHAREPOINT_SITE_ID"),
    "sp_report_root": _env("SHAREPOINT_REPORT_ROOT_TEST") if env_status == "test" else _env("SHAREPOINT_REPORT_ROOT"),
    "authority": f"https://login.microsoftonline.com/{tenant_id}" if tenant_id else "",
    "scope": ["https://graph.microsoft.com/.default"],
    "graph_endpoint": "https://graph.microsoft.com/v1.0",
    "report_data_api_key": _env("REPORT_DATA_API_KEY"),
    "TOTP_SECRET": _env("PMS_TOTP_SECRET"),
    "sp_base_url": _env("SHAREPOINT_BASE_URL"),
    "sp_site_path": _env("SHAREPOINT_SITE_PATH"),
    "sp_reports_folder": _env("SHAREPOINT_REPORTS_FOLDER"),
}

sender_email = _env("REPORT_SENDER_EMAIL")
e_mail = {
    "test_recipients": _env("REPORT_TEST_RECIPIENTS"),
    "sender_email": sender_email,
    "e_mail_sig": "Daily Report Automation Demo",
}

pms_credentials = {
    "pms_user": _env("PMS_USERNAME"),
    "pms_pass": _env("PMS_PASSWORD"),
    # A non-routable example URL prevents accidental login from a fresh clone.
    "base_url": _env("PMS_BASE_URL") or "https://example.invalid/",
    "backdate_site_id": "DEMO_PORTAL_001",
}

portal_pms = {
    "pms_user": pms_credentials["pms_user"],
    "pms_pass": pms_credentials["pms_pass"],
    "base_url": pms_credentials["base_url"],
}

# Fictional sample sites. These IDs are not connected to real properties.
group_ids = {
    "portal_pms": [("DEMO", "DEMO_PORTAL_001")],
    "cloud_pms": [("DEMO", "DEMO_CLOUD_001")],
}

today_date = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
as_of_date = today_date - timedelta(days=1)
this_month = as_of_date.replace(day=1)
dates = {
    "today_date": today_date,
    "as_of_date": as_of_date,
    "this_month": this_month,
    "eo_lastmonth": this_month - timedelta(days=1),
    "next_month": this_month + relativedelta(months=1),
    "next_2_months": this_month + relativedelta(months=2),
    "next_3_months": this_month + relativedelta(months=3),
    "next_12_months": this_month + relativedelta(months=1, years=1),
    "first_day_of_year": datetime(as_of_date.year, 1, 1),
    "last_day_of_year": datetime(as_of_date.year, 12, 31),
    "first_day_of_next_year": datetime(as_of_date.year + 1, 1, 1),
}

date_str = {
    "today_date": today_date.strftime("%d%m%Y"),
    "as_of_date": as_of_date.strftime("%d%m%Y"),
    "this_month": this_month.strftime("%d%m%Y"),
    "eo_lastmonth": dates["eo_lastmonth"].strftime("%d%m%Y"),
    "next_month": dates["next_month"].strftime("%d%m%Y"),
    "next_2_months": dates["next_2_months"].strftime("%d%m%Y"),
    "next_3_months": dates["next_3_months"].strftime("%d%m%Y"),
    "next_12_months": dates["next_12_months"].strftime("%d%m%Y"),
    "first_day_of_year": dates["first_day_of_year"].strftime("%d%m%Y"),
    "last_day_of_year": dates["last_day_of_year"].strftime("%d%m%Y"),
    "first_day_of_next_year": dates["first_day_of_next_year"].strftime("%d%m%Y"),
    "year": str(as_of_date.year),
    "this_month_name": calendar.month_name[this_month.month],
    "next_month_name": calendar.month_name[dates["next_month"].month],
    "next_2_months_name": calendar.month_name[dates["next_2_months"].month],
    "next_3_months_name": calendar.month_name[dates["next_3_months"].month],
    "month_number": as_of_date.strftime("%m"),
    "month_name": as_of_date.strftime("%b"),
}
