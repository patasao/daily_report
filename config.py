from pathlib import Path
import os
from dotenv import load_dotenv
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import calendar

load_dotenv()

env_status = "" # input "test" if you wanna test the script

# --- Path ---
project_root = Path(__file__).resolve().parent
home_dir = Path.home()

paths = {
    'project_root' : project_root,
    'home_dir' : home_dir,
    'log_dir' : project_root / "log",
    'sp_dor_dir' : (
        home_dir / os.getenv('sp_dor', '') / "TEST" 
        if env_status == "test" 
        else home_dir / os.getenv('sp_dor', '')
    ),
    'download_dir' : home_dir / os.getenv('dor_download_dir','Downloads/DOR'),
}

# --- Microsoft ---
tenant_id = os.getenv('tenant_id')

credentials = {
    'tenant_id' : tenant_id,
    'client_id' : os.getenv('client_id'),
    'client_secret' : os.getenv('client_secret'),
    'sp_site_id' : os.getenv('sp_site_id'),
    'sp_dor_root' : os.getenv('sp_dor_root_test', '') if env_status == "test" else os.getenv('sp_dor_root', ''),
    'authority' : f"https://login.microsoftonline.com/{tenant_id}",
    'scope' : ["https://graph.microsoft.com/.default"],
    'graph_endpoint' : "https://graph.microsoft.com/v1.0",
    'report_data_api_key' : os.getenv("report_data_api_key"),
    'TOTP_SECRET' : os.getenv("TOTP_SECRET")
}

# --- Email ---
e_mail = {
    'test_recipients': os.getenv('e_mail'),
    'sender_email' : os.getenv('e_mail'),
    'e_mail_sig' : (project_root.parent / "compass-signature.html").read_text(encoding='utf-8')
}

# --- Guestline ---
guestline = {
    'gl_user' : os.getenv('gl_user'),
    'gl_pass' : os.getenv('gl_pass')
}

# --- group_id ---
group_ids = {
    'mews' : [
        ("UK","ANGUSHOTEL")
        ],
    'guestline' : [
        ("UK","AROSFA"),
        ("UK","CPHEAST"),
        ("UK","CPHCHELT"),       
        ("UK","CPHALEX"),        
        ("UK","CPHBOS"),          
        ("UK","CPHCOLUM"),  
        ("UK","CPHCRAIG"), 
        ("UK","CPHCROWN"),
        ("UK","CPHDILKHU"),    
        ("UK","CPHESPLAN"),   
        ("UK","CPHHAND"),   
        ("UK","CPHHIGH"),
        ("UK","CPHIVY"),
        ("UK","CPHMIL"),
        ("UK","CPHNELSON"),         
        ("UK","CPHNORB"),
        ("UK","CPHPORT"),
        ("UK","CPHQDUND"), 
        ("UK","CPHROX"),         
        ("UK","CPHSEA"),
        ("UK","CPHSTRATH"),
        ("UK","CPHTRO"),       
        ("UK","BWSTATION"), #long        
        ("UK","CPHGOGL"), #long           
        ("UK","KEAVIL"),        
        ("UK","SYMQUEENS"), #long              
        ("UK","CPHSUITES"),
    ]
}

# --- Dates ---
today_date = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
as_of_date = today_date - timedelta(days=1)
this_month = as_of_date.replace(day=1)

dates = {
    'today_date' : today_date,
    'as_of_date' : as_of_date,
    'this_month' : this_month, 
    'eo_lastmonth' : this_month - timedelta(days=1),
    'next_month' : this_month + relativedelta(months=1),
    'next_2_months' : this_month + relativedelta(months=2),
    'next_3_months' : this_month + relativedelta(months=3),
    'next_12_months' : this_month + relativedelta(months=1, years=1),
    'first_day_of_year' : datetime(as_of_date.year, 1, 1),
    'last_day_of_year' : datetime(as_of_date.year, 12, 31),
    'first_day_of_next_year' : datetime(as_of_date.year + 1, 1, 1)
}

# --- Dates str ---
date_str = {
    'today_date' : today_date.strftime('%d%m%Y'),
    'as_of_date' : as_of_date.strftime('%d%m%Y'),
    'this_month' : this_month.strftime('%d%m%Y'),
    'eo_lastmonth' : dates['eo_lastmonth'].strftime('%d%m%Y'),
    'next_month' : dates['next_month'].strftime('%d%m%Y'),
    'next_2_months' : dates['next_2_months'].strftime('%d%m%Y'),
    'next_3_months' : dates['next_3_months'].strftime('%d%m%Y'),
    'next_12_months' : dates['next_12_months'].strftime('%d%m%Y'),
    'first_day_of_year' : dates['first_day_of_year'].strftime('%d%m%Y'),
    'last_day_of_year' : dates['last_day_of_year'].strftime('%d%m%Y'),
    'first_day_of_next_year' : dates['first_day_of_next_year'].strftime('%d%m%Y'),
    'year' : datetime(as_of_date.year, 1, 1).strftime('%Y'),
    'this_month_name' : calendar.month_name[this_month.month],
    'next_month_name' : calendar.month_name[dates['next_month'].month],
    'next_2_months_name' : calendar.month_name[dates['next_2_months'].month],
    'next_3_months_name' : calendar.month_name[dates['next_3_months'].month],
    'month_number' : as_of_date.strftime('%m'),
    'month_name' : as_of_date.strftime('%b')
}
