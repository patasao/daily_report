import os
import win32com.client as win32
from datetime import datetime, timedelta
import time
from PIL import Image, ImageOps
import pymupdf
import base64
from pathlib import Path
from services.email import Email
from services.ms_auth import TokenManager
from services.sharepoint import ReadSP, UploadSP
import config
import pythoncom
from services.logger import get_logger
import uuid
import win32process
from services.process_guard import ensure_gone, watch_process

logger = get_logger('refresh_excel_and_send', config.paths['log_dir'])
_SNAPSHOT_DPI = 200  # render above screen resolution, then shrink, so text stays sharp


def _warm_up_excel_gencache() -> None:
    """Build pywin32's early-bound Excel wrapper once, before any worker thread starts.
    Lazy per-CLSID builds are not thread-safe and raise spurious AttributeErrors."""
    pythoncom.CoInitialize()
    proc = None
    try:
        # DispatchEx first. EnsureDispatch("Excel.Application") on its own attaches to an
        # Excel the user already has open, then quits it out from under them.
        excel = win32.DispatchEx("Excel.Application")
        _, pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
        proc = watch_process(pid)
        excel = win32.gencache.EnsureDispatch(excel)
        excel.DisplayAlerts = False
        excel.Quit()
        del excel
    finally:
        ensure_gone(proc, label="gencache warm-up")
        pythoncom.CoUninitialize()


_warm_up_excel_gencache()


class RefreshAndSend:
    """
    Handles opening, refreshing, saving, and emailing Excel-based reports.
    Uses win32.DispatchEx to spawn isolated Excel instances per workbook,
    avoiding shared-state issues in concurrent runs.
    """
    def __init__(self):
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
        sp_email_path    = f"{config.credentials['sp_report_root']}/Report Data/OTHER/CONFIG/report_email_list.csv"
        self.email_list  = self._read_sp().read_csv(sp_email_path)

    def _read_sp(self) -> ReadSP:
        return ReadSP(self._token_manager.get_token(), self._sp_site_id)

    def _upload_sp(self) -> UploadSP:
        return UploadSP(self._token_manager.get_token(), self._sp_site_id)

    def _sp_path(self, relative_path: Path) -> str:
        return f"{config.credentials['sp_report_root']}/{relative_path.as_posix()}"

    @staticmethod
    def _supporting_docs_folder(country: str, site_id: str, dates: dict) -> Path:
        """Relative SharePoint path to this run's Supporting Docs output folder."""
        month_year = f"{dates['month_number']} {dates['month_name']} {dates['year']}"
        return Path(
            f"Supporting Docs/{country}/{site_id}/{dates['year']}"
            f"/{month_year} ({site_id})/{site_id} {dates['as_of_dot']}"
        )

    # --- Internal helpers ---
    @staticmethod
    def _open_excel() -> tuple:
        """Spawn a fresh, isolated Excel process plus a handle to its OS process.
        The handle lets us force-kill an Excel that later refuses to Quit."""
        excel = win32.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        if config.env_status == "test":
            excel.Visible = True #Deactivate run in background to inspect the automation movement
        else: excel.Visible = False #Activate run in Background
        _, pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
        return excel, watch_process(pid)

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
        wb.Sheets("Pickup").Range("E3").Value = dates['as_of']
        wb.Sheets("Cancellations").Range("AO7").Value = dates['as_of']
        wb.Sheets("Variance").Range("C43").Value = dates['as_of']
        wb.Sheets("Pickup").Range("H3").Value = dates['first_day_of_month']
        wb.Sheets("Variance").Range("D3").Value = dates['first_day_of_month']
        wb.Sheets("MarketSegments").Range("C3").Value = dates['first_day_of_month']
        wb.Sheets("Cancellations").Range("C3").Value = dates['first_day_of_month']

        excel.CalculateUntilAsyncQueriesDone()
        wb.Sheets("ReportSum").Activate()
        wb.Save()
        logger.info("Report updated with date %s", dates["as_of"].strftime("%d/%m/%Y"))

    @staticmethod
    def _snapshot(wb, sheet_name: str, cell_range: str, save_path: Path, attempts: int = 3) -> Path:
        """Save an Excel range as a PNG by exporting it to PDF and rendering that page.
        CopyPicture needs the clipboard, which is unusable while the laptop is locked."""
        ws = wb.Sheets(sheet_name)
        rng = ws.Range(cell_range)
        pdf_path = Path(save_path).with_suffix(".pdf")
        excel = wb.Application
        excel.PrintCommunication = False  # batch the page-setup changes into one printer call
        try:
            ps = ws.PageSetup
            ps.PrintArea = rng.Address
            ps.Orientation = 2  # xlLandscape
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = 1
            ps.LeftMargin = ps.RightMargin = ps.TopMargin = ps.BottomMargin = 0
            ps.HeaderMargin = ps.FooterMargin = 0
            ps.CenterHeader = ps.CenterFooter = ps.LeftHeader = ps.LeftFooter = ""
            ps.RightHeader = ps.RightFooter = ""
        finally:
            excel.PrintCommunication = True

        for attempt in range(1, attempts + 1):
            try:
                ws.ExportAsFixedFormat(Type=0, Filename=str(pdf_path), Quality=0,  # 0=PDF, standard quality
                                       IncludeDocProperties=False, IgnorePrintAreas=False,
                                       From=1, To=1, OpenAfterPublish=False)
                with pymupdf.open(pdf_path) as doc:
                    pix = doc[0].get_pixmap(dpi=_SNAPSHOT_DPI)
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                bbox = ImageOps.invert(img).getbbox()  # crop the white page around the table
                if bbox is None:
                    raise RuntimeError(f"Snapshot of {cell_range} came out blank")
                img = img.crop(bbox)
                # 120 px per 72 pt = the old screen copy at 125% display scaling, so emails keep their size.
                width = round(rng.Width * 120 / 72)
                img = img.resize((width, round(img.height * width / img.width)), Image.LANCZOS)
                img.save(save_path)
                return save_path
            except Exception as exc:
                if attempt == attempts:
                    raise
                logger.warning("Snapshot attempt %d/%d failed for %s: %s", attempt, attempts, cell_range, exc)
                time.sleep(2)
            finally:
                pdf_path.unlink(missing_ok=True)

    def _get_email_list(self, site_id: str) -> str:
        """Get email list to send"""
        df = self.email_list[self.email_list['site_id'] == site_id]
        email_list = df['email'].dropna().astype(str).tolist()
        logger.info("%s: email list loaded (%d recipients)", site_id, len(email_list))
        return ";".join(email_list)

    def _draft_email(self, wb, country: str, site_id: str, dates: dict) -> dict:
        """Compose the Graph API message payload with an inline table snapshot."""
        as_of_dot = dates["as_of_dot"]

        ws = wb.Sheets("ReportConfig")
        ht_name = ws.Range("D6").Value
        subject = ws.Range("D8").Value
        # page_link = ws.Range("A103").Value
        if country == "UK":
            page_link = ws.Range("A77").Value # New Path
        else: page_link = None

        folder_path = self._sp_path(self._supporting_docs_folder(country, site_id, dates))
        folder_info = self._read_sp().get_file_info_from_path(folder_path)
        folder_link = folder_info.get('webUrl') if folder_info else None

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

        if country == "UK":
            report_line = f'Report: <a href="{page_link}">{ht_name}</a><br>\n            ' if page_link else ''
            body = f"""
            <html><body>
                <p>Dear all,<br><br>
                The report for {ht_name} is now available on SharePoint.<br><br>
                {report_line}Supporting Documents: <a href="{folder_link}">{ht_name} {as_of_dot}</a><br><br>
                <img src="cid:table1"><br><br>
                Best Regards,<br>{self.signature}
                </p>
            </body></html>
            """
        else:
            body = f"""
            <html><body>
                <p>Dear all,<br><br>
                The report for {ht_name} is now available on SharePoint: <a href="{folder_link}">{ht_name} {as_of_dot}</a><br><br>
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
        data_folder  = self.report_data_folder / f"{country}_RAW" / f"{site_id}"
        cloud_pms_ids     = {site[1] for site in config.group_ids['cloud_pms']}

        checks = [("Actual Forecast", "Actual Forecast", "xlsx")] if site_id in cloud_pms_ids else [
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
            new_name   = (
                file_path.name.replace("Display", dates["as_of_dot"])
                if "Display" in file_path.name
                else f"{file_path.stem}_{dates['as_of_dot']}{file_path.suffix}"
            )

            temp_dir   = config.paths['download_dir'] / "temp_report"
            temp_dir.mkdir(parents=True, exist_ok=True)
            local_path = temp_dir / new_name
            wb.SaveCopyAs(str(local_path))
            logger.info(f"{site_id}: Saved copy to temp: {local_path}")

            sp_dest = self._sp_path(self._supporting_docs_folder(country, site_id, dates))
            self._upload_sp().upload_file(str(local_path), sp_dest, new_name)
            logger.info(f"{site_id}: Uploaded {new_name} -> {sp_dest}")

            local_path.unlink(missing_ok=True)
            return True
        except Exception as exc:
            logger.error(f"{site_id}: Save failed: %s", exc)
            return False

    def _send_email(self, wb, country: str, site_id: str, dates: dict) -> None:
        """Build and dispatch the Report email, guarded by date/link conditions."""
        logger.info(f"{site_id}: Sending email..")
        yesterday = (datetime.today() - timedelta(days=1)).strftime("%d/%m/%Y")
        last_update = dates["as_of"].strftime("%d/%m/%Y")

        payload = self._draft_email(wb, country, site_id, dates)

        if last_update == yesterday and payload["folder_link"] and (payload["page_link"] or country != "UK"):
            Email().send(self.sender_email, {k: v for k, v in payload.items()
                                           if k not in ("ht_name", "page_link", "folder_link")})
            logger.info(f"{site_id}: Email sent for %s at %s", site_id, datetime.now().strftime("%Y%m%d_%H%M%S"))
            logger.info("---")
        else:
            logger.error(f"{site_id}: Date mismatch for {payload['ht_name']}: as_of={last_update}, expected as_of={yesterday}. Check upstream data source.")
            raise ValueError(f"{site_id}: Date mismatch for {payload['ht_name']}: as_of={last_update}, expected as_of={yesterday}. Check upstream data source.")

    # --- Public entry point ---
    def run(self, country: str, site_id: str, refresh: bool = True,
            send_email: bool = True) -> tuple[str, str]:
        """
        Pipeline for one site: check files -> open Excel -> refresh -> set dropdowns -> save -> email.
        Skip a step with refresh=False or send_email=False. Each call spawns its own Excel process.
        """
        pythoncom.CoInitialize()
        label = "REFRESH AND SEND EMAIL" if send_email else "RESAVE"

        check = self._data_files_check(country, site_id)
        if not check:
            return site_id, "Skipped (Missing data source files to process)"

        logger.info("%s: data complete — starting processing", site_id)
        report_sp_folder = self._sp_path(Path("Report Template") / country)
        report_files = [
            f for f in self._read_sp().list_files_in_folder_by_ext(report_sp_folder, ['.xlsm', '.xlsx'])
            if site_id in f['name']
        ]
        if not report_files:
            return site_id, "Skipped (No report file)"

        excel, excel_proc = self._open_excel()
        wb = None
        local_path = None
        try:
            report_file = report_files[0]
            logger.info("Processing: %s", report_file['name'])

            temp_dir   = config.paths['download_dir'] / "temp_report_template"
            temp_dir.mkdir(parents=True, exist_ok=True)
            local_path = temp_dir / report_file['name']
            local_path.write_bytes(self._read_sp().get_file_bytes_by_id(report_file['id']))

            wb = excel.Workbooks.Open(str(local_path))
            time.sleep(10)
            if refresh and config.env_status != "test":
                wb.RefreshAll()
            excel.CalculateUntilAsyncQueriesDone()

            dates = self._get_dates(wb)
            self._set_default(wb, excel, dates)
            self._upload_sp().upload_file(str(local_path), report_sp_folder, preserve_identity=True)
            logger.info(f"{site_id}: Updated template on SharePoint -> {report_sp_folder}")

            if not self._save_file(wb, local_path, country, site_id, dates):
                raise RuntimeError(f"File save failed for {site_id}")
            if send_email:
                time.sleep(10)
                self._send_email(wb, country, site_id, dates)
            return (site_id, f"{label} SUCCESS")

        except Exception as exc:
            logger.error("Processing failed for %s: %s", site_id, exc, exc_info=True)
            return (site_id, f"{label} FAILED: ({type(exc).__name__})")

        finally:
            if wb:
                try:
                    wb.Close(SaveChanges=False)
                except Exception as exc:
                    logger.warning("%s: workbook Close failed: %s", site_id, exc)
            try:
                excel.Quit()
            except Exception as exc:
                logger.warning("%s: Excel.Quit rejected: %s", site_id, exc)
            # A rejected Quit leaves Excel holding the workbook open, which blocks every retry.
            ensure_gone(excel_proc, label=site_id)
            if local_path:
                Path(local_path).unlink(missing_ok=True)
            pythoncom.CoUninitialize()
