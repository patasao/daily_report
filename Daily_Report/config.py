from pathlib import Path
import os
from dotenv import load_dotenv
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import calendar

load_dotenv()

# Set to "test" to run in debug mode (visible browser, test email recipients)
env_status = ""

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
project_root = Path(__file__).resolve().parent
home_dir = Path.home()

paths = {
    'project_root' : project_root,
    'home_dir'     : home_dir,
    'log_dir'      : project_root / "log",
    # Local SharePoint sync directory (OneDrive client or mapped drive)
    'sp_dor_dir'   : (
        home_dir / os.getenv('sp_dor', '') / "TEST"
        if env_status == "test"
        else home_dir / os.getenv('sp_dor', '')
    ),
    # Temporary staging directory for Selenium downloads before SharePoint upload
    'download_dir' : home_dir / os.getenv('dor_download_dir', 'Downloads/DOR'),
}

# ---------------------------------------------------------------------------
# Microsoft 365 / Graph API — all values sourced from environment variables
# ---------------------------------------------------------------------------
tenant_id = os.getenv('tenant_id')

credentials = {
    'tenant_id'           : tenant_id,
    'client_id'           : os.getenv('client_id'),
    'client_secret'       : os.getenv('client_secret'),
    'sp_site_id'          : os.getenv('sp_site_id'),
    'sp_dor_root'         : os.getenv('sp_dor_root_test', '') if env_status == "test" else os.getenv('sp_dor_root', ''),
    'authority'           : f"https://login.microsoftonline.com/{tenant_id}",
    'scope'               : ["https://graph.microsoft.com/.default"],
    'graph_endpoint'      : "https://graph.microsoft.com/v1.0",
    'report_data_api_key' : os.getenv("report_data_api_key"),
    'TOTP_SECRET'         : os.getenv("TOTP_SECRET"),
    # SharePoint deep-link components used when composing per-property folder links in emails
    'sp_base_url'         : os.getenv('SP_BASE_URL'),
    'sp_site_path'        : os.getenv('SP_SITE_PATH'),
    'sp_reports_folder'   : os.getenv('SP_REPORTS_FOLDER'),
}

# ---------------------------------------------------------------------------
# Email
# ---------------------------------------------------------------------------
e_mail = {
    'test_recipients' : os.getenv('e_mail'),
    'sender_email'    : os.getenv('e_mail'),
    'e_mail_sig'      : (project_root.parent / "email_signature.html").read_text(encoding='utf-8'),
}

# ---------------------------------------------------------------------------
# On-premise PMS credentials (web portal with username/password + TOTP 2FA)
# ---------------------------------------------------------------------------
pms_credentials = {
    'pms_user'         : os.getenv('pms_user'),
    'pms_pass'         : os.getenv('pms_pass'),
    'base_url'         : os.getenv('PMS_BASE_URL'),
    # Site that uses end-of-last-month as the actuals start date (non-standard fiscal cut-off)
    'backdate_site_id' : os.getenv('BACKDATE_SITE_ID', 'SITE_A001'),
}

# ---------------------------------------------------------------------------
# Hotel property site IDs grouped by PMS platform
#
#   pms       = legacy on-premise PMS (reports extracted via Selenium web automation + TOTP)
#   cloud_pms = cloud-native PMS (data arrives via pre-scheduled SharePoint exports)
# ---------------------------------------------------------------------------
group_ids = {
    'cloud_pms': [
        ("UK", "SITE_B001"),
    ],
    'pms': [
        ("UK", "SITE_A001"),  # non-standard date offset (see backdate_site_id)
        ("UK", "SITE_A002"),
        ("UK", "SITE_A003"),
        ("UK", "SITE_A004"),
        ("UK", "SITE_A005"),
        ("UK", "SITE_A006"),
        ("UK", "SITE_A007"),
        ("UK", "SITE_A008"),
        ("UK", "SITE_A009"),
        ("UK", "SITE_A010"),
        ("UK", "SITE_A011"),
        ("UK", "SITE_A012"),
        ("UK", "SITE_A013"),
        ("UK", "SITE_A014"),
        ("UK", "SITE_A015"),
        ("UK", "SITE_A016"),
        ("UK", "SITE_A017"),
        ("UK", "SITE_A018"),
        ("UK", "SITE_A019"),
        ("UK", "SITE_A020"),
        ("UK", "SITE_A021"),
        ("UK", "SITE_A022"),
        ("UK", "SITE_A023"),  # extended processing time
        ("UK", "SITE_A024"),  # extended processing time
        ("UK", "SITE_A025"),
        ("UK", "SITE_A026"),  # extended processing time
        ("UK", "SITE_A027"),
    ],
}

# ---------------------------------------------------------------------------
# Date calculations — all relative to yesterday (the standard DOR "as-of" date)
# ---------------------------------------------------------------------------
today_date = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
as_of_date = today_date - timedelta(days=1)
this_month = as_of_date.replace(day=1)

dates = {
    'today_date'            : today_date,
    'as_of_date'            : as_of_date,
    'this_month'            : this_month,
    'eo_lastmonth'          : this_month - timedelta(days=1),
    'next_month'            : this_month + relativedelta(months=1),
    'next_2_months'         : this_month + relativedelta(months=2),
    'next_3_months'         : this_month + relativedelta(months=3),
    'next_12_months'        : this_month + relativedelta(months=1, years=1),
    'first_day_of_year'     : datetime(as_of_date.year, 1, 1),
    'last_day_of_year'      : datetime(as_of_date.year, 12, 31),
    'first_day_of_next_year': datetime(as_of_date.year + 1, 1, 1),
}

# Pre-formatted strings used in report file naming and web-automation parameter entry
date_str = {
    'today_date'            : today_date.strftime('%d%m%Y'),
    'as_of_date'            : as_of_date.strftime('%d%m%Y'),
    'this_month'            : this_month.strftime('%d%m%Y'),
    'eo_lastmonth'          : dates['eo_lastmonth'].strftime('%d%m%Y'),
    'next_month'            : dates['next_month'].strftime('%d%m%Y'),
    'next_2_months'         : dates['next_2_months'].strftime('%d%m%Y'),
    'next_3_months'         : dates['next_3_months'].strftime('%d%m%Y'),
    'next_12_months'        : dates['next_12_months'].strftime('%d%m%Y'),
    'first_day_of_year'     : dates['first_day_of_year'].strftime('%d%m%Y'),
    'last_day_of_year'      : dates['last_day_of_year'].strftime('%d%m%Y'),
    'first_day_of_next_year': dates['first_day_of_next_year'].strftime('%d%m%Y'),
    'year'                  : datetime(as_of_date.year, 1, 1).strftime('%Y'),
    'this_month_name'       : calendar.month_name[this_month.month],
    'next_month_name'       : calendar.month_name[dates['next_month'].month],
    'next_2_months_name'    : calendar.month_name[dates['next_2_months'].month],
    'next_3_months_name'    : calendar.month_name[dates['next_3_months'].month],
    'month_number'          : as_of_date.strftime('%m'),
    'month_name'            : as_of_date.strftime('%b'),
}
