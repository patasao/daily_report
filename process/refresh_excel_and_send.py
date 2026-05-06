import os
import win32com.client as win32
from datetime import datetime, timedelta
import time
from PIL import ImageGrab
import base64
from pathlib import Path
from services.email import Email
from services.ms_auth import TokenManager
from services.sharepoint import ReadSP, UploadSP
import config
import pythoncom
from services.logger import get_logger
import threading
import uuid

logger = get_logger('refresh_excel_and_send', config.paths['log_dir'])
_clipboard_lock = threading.Lock()

class RefreshAndSend:
    """
    Handles opening, refreshing, saving, and emailing Excel-based DOR reports.
    Uses win32.DispatchEx to spawn isolated Excel instances per workbook,
    avoiding shared-state issues in concurrent runs.
    """
    def __init__(self):
        self.sp_dor_folder      = config.paths['sp_dor_dir']
        self.report_data_folder = Path("Report Data")
        self.sender_email       = config.e_mail['sender_email']
        self.signature          = config.e_mail['e_mail_sig']
        self._token_manager = TokenManager(
            client_id=config.credentials['client_id'],
            authority=config.credentials['authority'],
            client_secret=config.credentials['client_secret'],
            scopes=config.credentials['scope']
        )
        self._sp_site_id = config.credentials['sp_site_id']
        sp_email_path    = f"{config.credentials['sp_dor_root']}/Report Data/OTHER/dor_email_list.csv"
        self.email_list  = self._read_sp().read_csv(sp_email_path)

    def _read_sp(self) -> ReadSP:
        return ReadSP(self._token_manager.get_token(), self._sp_site_id)

    def _upload_sp(self) -> UploadSP:
        return UploadSP(self._token_manager.get_token(), self._sp_site_id)

    def _sp_path(self, relative_path: Path) -> str:
        return f"{config.credentials['sp_dor_root']}/{relative_path.as_posix()}"

    # --- Internal helpers ---
    @staticmethod
    def _open_excel() -> object:
        """Spawn a fresh, isolated Excel process via DispatchEx."""
        excel = win32.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        if config.env_status == "test":
            excel.Visible = True #Deactivate run in background to inspect the automation movement
        else: excel.Visible = False #Activate run in Background
        return excel

    @staticmethod
    def _get_dates(wb) -> dict:
        """Extract date values from the ReportConfig sheet."""
        as_of = wb.Sheets("ReportConfig").Range("D1").Value 
        return {
            'as_of': as_of,
            'as_of_dot': as_of.strftime('%d.%m.%y'),
            'first_day_of_month': datetime(as_of.year, as_of.month, 1).strftime('%m/%d/%Y'),
            'year': as_of.strftime('%Y'),
            'month_number': as_of.strftime('%m'),
            'month_name': as_of.strftime('%b')
        }

    @staticmethod
    def _set_default(wb, excel, dates: dict) -> None:
        """Push date values into the dropdown/filter cells on each sheet."""
        logger.info("Updating report dates.")
        wb.Sheets("Pickup").Range("F2").Value = dates['as_of']
        wb.Sheets("Pickup").Range("I2").Value = dates['first_day_of_month']
        wb.Sheets("Variance").Range("D2").Value = dates['first_day_of_month']
        wb.Sheets("MarketSegments").Range("D2").Value = dates['first_day_of_month']
        wb.Sheets("Cancellations").Range("C2").Value = dates['first_day_of_month']
        excel.CalculateUntilAsyncQueriesDone()
        wb.Sheets("DORSum").Activate()
        wb.Save()
        logger.info("Report updated with date %s", dates["as_of"].strftime("%d/%m/%Y"))

    @staticmethod
    def _snapshot(wb, sheet_name: str, cell_range: str, save_path: Path) -> Path:
        """Copy an Excel range as an image and save it to disk."""
        with _clipboard_lock:  
            wb.Sheets(sheet_name).Range(cell_range).CopyPicture(Appearance=1, Format=2) # 1=Screen, 2=Bitmap
            time.sleep(0.5)
            img = ImageGrab.grabclipboard()
            if img is None:
                raise RuntimeError(f"Clipboard grab failed for range {cell_range}")
            img.save(save_path)

    def _get_email_list(self, site_id: str) -> str:
        """Get email list to send"""
        df = self.email_list[self.email_list['site_id'] == site_id]
        email_list = df['email'].dropna().astype(str).tolist()
        logger.info("%s: email list loaded (%d recipients)", site_id, len(email_list))
        return ";".join(email_list)
    
    def _draft_email(self, wb, country: str, site_id: str, dates: dict) -> dict:
        """Compose the Graph API message payload with an inline table snapshot."""
        as_of_dot = dates["as_of_dot"]
        year = dates["year"]
        month_number = dates["month_number"]
        month_name = dates["month_name"]

        ws = wb.Sheets("ReportConfig")
        ht_name = ws.Range("D6").Value
        subject = ws.Range("D8").Value
        page_link = ws.Range("A103").Value
        folder_link = (
            f"https://compasshospitalitycom.sharepoint.com/:f:/r/sites/DOR/Shared%20Documents"
            f"/1_Namfon_DOR/{country}/{site_id}/{year}"
            f"/{month_number}%20{month_name}%20{year}%20({site_id})"
            f"/{site_id}%20{as_of_dot}?csf=1&web=1"
        )

        # Capture inline table image
        temp_img = Path(os.environ.get("TEMP", ".")) / f"table_{site_id}_{uuid.uuid4().hex}.png"
        self._snapshot(wb, "ReportConfig", "C10:I24", temp_img)
        with open(temp_img, "rb") as f:
            img_base64 = base64.b64encode(f.read()).decode("utf-8")

        if config.env_status == "test":
            raw_recipients = config.e_mail['test_recipients']
        else:
            raw_recipients = self._get_email_list(site_id)
        recipients = [
            {"emailAddress": {"address": e.strip()}}
            for e in raw_recipients.split(";")
            if e.strip()
        ]

        body = f"""
        <html><body>
            <p>Dear all,<br><br>
            The DOR for {ht_name} is now available on SharePoint.<br><br>
            DOR: <a href="{page_link}">{ht_name}</a><br>
            Supporting Documents: <a href="{folder_link}">{ht_name} {as_of_dot}</a><br><br>
            <img src="cid:table1"><br><br>
            Best Regards,<br>{self.signature}
            </p>
        </body></html>
        """

        return {
            "ht_name": ht_name,
            "page_link": page_link,
            "folder_link": folder_link,
            "message": {
                "subject": subject,
                "body": {"contentType": "HTML", "content": body},
                "toRecipients": recipients,
                "attachments": [
                    {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": "table.png",
                        "contentBytes": img_base64,
                        "contentId": "table1",
                        "isInline": True,
                    }
                ],
            },
            "saveToSentItems": "true",
        }

    def _data_files_check(self, country: str, site_id: str) -> bool:
        """Check required source files exist on SharePoint before processing."""
        logger.info(f"{site_id}: Checking last update data")

        as_of        = config.dates['as_of_date']
        target_dates = [as_of - timedelta(days=1), as_of]
        date_strings = [d.strftime('%d.%m.%y') for d in target_dates]
        data_folder  = self.report_data_folder / country / site_id
        mews_ids     = {site[1] for site in config.group_ids['mews']}

        checks = [("Reservation Report", "Reservation report", "xlsx")] if site_id in mews_ids else [
            ("EOD", "EOD", "csv"), ("Actual", "Actual", "csv"), ("Forecast", "Forecast", "csv")
        ]

        read_sp = self._read_sp()
        missing_files = [
            self._sp_path(data_folder / f"{site_id} {sub}" / f"{site_id} {suffix} {ds}.{ext}")
            for sub, suffix, ext in checks
            for ds in date_strings
            if not read_sp.get_file_info_from_path(
                self._sp_path(data_folder / f"{site_id} {sub}" / f"{site_id} {suffix} {ds}.{ext}")
            )
        ]

        if missing_files:
            for m in missing_files:
                logger.error(f"{site_id}: Missing file at {m} ---")
            logger.error(f"{site_id}: Some data files are missing -----")
            return False

        logger.info(f"{site_id}: All data files are completely collected.")
        return True

    # --- Core per-site logic ---
    def _save_file(self, wb, file_path: Path, country: str, site_id: str, dates: dict) -> bool:
        """Copy Excel to a temp path via SaveCopyAs, upload to SharePoint, then clean up."""
        try:
            logger.info(f"{site_id}: Saving Excel file...")
            month_year = f"{dates['month_number']} {dates['month_name']} {dates['year']}"
            new_name   = (
                file_path.name.replace("Display", dates["as_of_dot"])
                if "Display" in file_path.name
                else f"{file_path.stem}_{dates['as_of_dot']}{file_path.suffix}"
            )

            temp_dir   = config.paths['download_dir'] / "temp_dor"
            temp_dir.mkdir(parents=True, exist_ok=True)
            local_path = temp_dir / new_name
            wb.SaveCopyAs(str(local_path))
            logger.info(f"{site_id}: Saved copy to temp: {local_path}")

            sp_dest = (
                f"{config.credentials['sp_dor_root']}/{country}/{site_id}"
                f"/{dates['year']}/{month_year} ({site_id})"
                f"/{site_id} {dates['as_of_dot']}"
            )
            self._upload_sp().upload_file(str(local_path), sp_dest, new_name)
            logger.info(f"{site_id}: Uploaded {new_name} -> {sp_dest}")

            local_path.unlink(missing_ok=True)
            return True
        except Exception as exc:
            logger.error(f"{site_id}: Save failed: %s", exc)
            return False

    def _send_email(self, wb, country: str, site_id: str, dates: dict) -> None:
        """Build and dispatch the DOR email, guarded by date/link conditions."""
        logger.info(f"{site_id}: Sending email..")
        yesterday = (datetime.today() - timedelta(days=1)).strftime("%d/%m/%Y")
        last_update = dates["as_of"].strftime("%d/%m/%Y")

        payload = self._draft_email(wb, country, site_id, dates)

        if last_update == yesterday and payload["folder_link"] and payload["page_link"]:
            Email().send(self.sender_email, {k: v for k, v in payload.items()
                                           if k not in ("ht_name", "page_link", "folder_link")})
            logger.info(f"{site_id}: Email sent for %s at %s", site_id, datetime.now().strftime("%Y%m%d_%H%M%S"))
            logger.info("---")
        else:
            logger.error(f"{site_id}: Date mismatch for {payload['ht_name']}: as_of={last_update}, expected as_of={yesterday}. Check upstream data source.")
            raise ValueError(f"{site_id}: Date mismatch for {payload['ht_name']}: as_of={last_update}, expected as_of={yesterday}. Check upstream data source.")

    def run(self, country: str, site_id: str) -> tuple[str, str]:
        """
        Full pipeline for one site: check files -> open Excel -> refresh ->
        set dropdowns -> save -> email.

        Each call spawns its own Excel process (DispatchEx) so sites can
        run concurrently without collision.
        """
        pythoncom.CoInitialize() 

        check = self._data_files_check(country, site_id)
        if not check:
            return site_id, "Skipped (Missing data source files to process)"

        logger.info("%s: data complete — starting processing", site_id)
        dor_folder = self.sp_dor_folder / country / "ALLXLSM"
        if not dor_folder.is_dir():
            logger.error("DOR folder not found: %s", dor_folder)
            return site_id, "Fail (DOR folder missing)"

        dor_files = [p for p in dor_folder.iterdir() if site_id in p.name]
        if not dor_files:
            return site_id, "Skipped (No DOR File)"

        excel = self._open_excel()
        wb = None
        try:
            file_path = dor_files[0]
            logger.info("Processing: %s", file_path.name)

            wb = excel.Workbooks.Open(str(file_path))
            time.sleep(10)
            wb.RefreshAll() if config.env_status != "test" else None
            excel.CalculateUntilAsyncQueriesDone()

            dates = self._get_dates(wb)
            self._set_default(wb, excel, dates)
            if not self._save_file(wb, file_path, country, site_id, dates):
                raise RuntimeError(f"File save failed for {site_id}")
            time.sleep(10)
            self._send_email(wb, country, site_id, dates)
            return (site_id, "REFRESH AND SEND EMAIL SUCCESS")

        except Exception as exc:
            logger.error("Processing failed for %s: %s", site_id, exc, exc_info=True)
            return (site_id, f"REFRESH AND SEND EMAIL FAILED: ({type(exc).__name__})")

        finally:
            if wb:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
            try:
                excel.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()




