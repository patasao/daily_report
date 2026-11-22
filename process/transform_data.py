import config
from datetime import timedelta
import pandas as pd
from pathlib import Path
import numpy as np
from abc import ABC, abstractmethod
from services.logger import get_logger
from services.ms_auth import TokenManager
from services.sharepoint import ReadSP, UploadSP
import time
import tempfile
from io import StringIO
import os
import threading
from concurrent.futures import ThreadPoolExecutor

pd.set_option('future.no_silent_downcasting', True)

logger = get_logger('transform_data', config.paths['log_dir'])

class TransformData(ABC):
    # (report name, extension) of daily source files; each child sets its own
    _source_checks: list[tuple[str, str]] = []
    _config_lock = threading.Lock()
    _market_segment_cache: pd.DataFrame | None = None

    def __init__(self):
        self.report_data_folder = Path("Report Data")
        self._token_manager = TokenManager(
            client_id=config.credentials['client_id'],
            authority=config.credentials['authority'],
            client_secret=config.credentials['client_secret'],
            scopes=config.credentials['scope']
        )
        self._sp_site_id = config.credentials['sp_site_id']
        with TransformData._config_lock:
            if TransformData._market_segment_cache is None:
                TransformData._market_segment_cache = self._get_market_segment()
        self.market_segment = TransformData._market_segment_cache

    def _sp_read(self) -> ReadSP:
        return ReadSP(self._token_manager.get_token(), self._sp_site_id)

    def _sp_upload(self) -> UploadSP:
        return UploadSP(self._token_manager.get_token(), self._sp_site_id)

    def _sp_path(self, relative_path: Path) -> str:
        return f"{config.credentials['sp_report_root']}/{relative_path.as_posix()}"

    # --- helpers ---
    def _get_market_segment(self) -> pd.DataFrame:
        """Get Market Segment configuration for each site_id."""
        sp_path = self._sp_path(self.report_data_folder / "OTHER" / "CONFIG"/ "market_segment.csv")
        try:
            return self._sp_read().read_csv(sp_path)
        except Exception as e:
            logger.warning(f"Could not load MarketSegment file: {e}")
            return pd.DataFrame()

    @staticmethod
    def _date_key(col: str):
        """Cover key: the row's date in `col`, so only dates present in new data are replaced."""
        return lambda d: pd.to_datetime(d[col], format='mixed', errors='coerce').dt.normalize()

    def _save_grouped_csv(self, df: pd.DataFrame, report_name: str, country: str, site_id: str,
                          by_quarter: bool = False, cover_key=None):
        """Write one CSV per quarter or year. With cover_key, rows of the existing file whose
        key is absent from df are kept, so cleaned-up source never wipes old data."""
        sp_folder = self._sp_path(self.report_data_folder / f"{country}_TRANSFORMED" / site_id)
        upload_sp = self._sp_upload()
        filename  = ""
        try:
            if by_quarter:
                iter_groups = (
                    (f"{site_id} {report_name} {int(yr)} Q{int(q)}.csv", grp)
                    for (yr, q), grp in df.groupby(["year", "quarter"])
                )
            else:
                iter_groups = (
                    (f"{site_id} {report_name} {int(yr)}.csv", grp)
                    for yr, grp in df.groupby("year")
                )
            existing_ids = {}
            if cover_key is not None:
                # No fallback on failure: writing without the old rows would wipe them
                read_sp = self._sp_read()
                existing_ids = {f['name']: f['id'] for f in read_sp.list_files_in_folder(sp_folder)}
            for filename, group in iter_groups:
                if filename in existing_ids:
                    old = pd.read_csv(StringIO(read_sp.get_file_content_by_id(existing_ids[filename])))
                    old = old[~cover_key(old).isin(set(cover_key(group).dropna()))]
                    if not old.empty:
                        # Old dates come back as text; match the new dtype so the CSV keeps one date format
                        for col in group.select_dtypes('datetime').columns.intersection(old.columns):
                            old[col] = pd.to_datetime(old[col], format='mixed', errors='coerce')
                        logger.info(f"{site_id}: {filename} keeps {len(old)} rows not in current source")
                        group = pd.concat([old, group], ignore_index=True)
                fd, tmp  = tempfile.mkstemp(suffix='.csv')
                os.close(fd)
                tmp_path = Path(tmp)
                try:
                    group.to_csv(tmp_path, index=False, encoding="utf-8")
                    upload_sp.upload_file(str(tmp_path), sp_folder, filename)
                    logger.info(f"{site_id}: Uploaded {filename} -> {sp_folder}")
                finally:
                    tmp_path.unlink(missing_ok=True)
        except Exception as e:
            logger.error(f"{site_id}: Failed to save/upload {filename} - {e}")

    def _read_bulk_excel(self, site_id: str, folderpath: Path, sheet_name: str = 0, transform=None) -> pd.DataFrame:
        """Read all Excel files from a SharePoint folder using ReadSP.read_excel()."""
        sp_path = self._sp_path(folderpath)
        try:
            df = self._sp_read().read_excel(sp_path, sheet_name=sheet_name, transform=transform)
            if df.empty:
                logger.warning(f"{site_id}: No Excel files at {sp_path}")
            return df
        except Exception as e:
            logger.error(f"{site_id}: Failed to read Excel from {sp_path}: {e}")
            return pd.DataFrame()

    def _read_bulk_csv(self, site_id: str, folderpath: Path) -> pd.DataFrame:
        """Read all CSV files from a SharePoint folder using ReadSP.read_csv()."""
        sp_path = self._sp_path(folderpath)
        try:
            df = self._sp_read().read_csv(sp_path)
            if df.empty:
                logger.warning(f"{site_id}: No CSV files at {sp_path}")
                return pd.DataFrame()
            df.rename(columns=str, inplace=True)
            df.dropna(how='all', axis=1, inplace=True)
            return df
        except Exception as e:
            logger.error(f"{site_id}: Failed to read CSVs from {sp_path}: {e}")
            return pd.DataFrame()

    @staticmethod
    def _format_date(series: pd.Series) -> pd.Series:
        """Helper method to format date"""
        return pd.to_datetime(series,format='mixed',errors="coerce",dayfirst=True)

    # --- Processors ---
    def data_files_check(self, country: str, site_id: str, attempts: int = 3, wait_time: int = 30) -> bool:
        """Check required source files exist on SharePoint before processing."""
        logger.info(f"{site_id}: Checking last update data")

        as_of        = config.dates['as_of_date']
        target_dates = [as_of - timedelta(days=1), as_of]
        date_strings = [d.strftime('%d.%m.%y') for d in target_dates]
        data_folder  = self.report_data_folder / f"{country}_RAW" / site_id
        checks       = self._source_checks

        def find_missing():
            read_sp = self._sp_read()
            missing = []
            for name, ext in checks:
                for ds in date_strings:
                    local_path = data_folder / f"{site_id} {name}" / f"{site_id} {name} {ds}.{ext}"
                    if not read_sp.get_file_info_from_path(self._sp_path(local_path)):
                        missing.append(self._sp_path(local_path))
            return missing

        missing_files = find_missing()
        for attempt in range(1, attempts + 1):
            if not missing_files:
                break
            logger.warning(f"{site_id}: Some files missing (attempt {attempt}/{attempts}), retrying in {wait_time}s...")
            time.sleep(wait_time)
            missing_files = find_missing()

        if missing_files:
            for m in missing_files:
                logger.error(f"{site_id}: Missing file at {m} ---")
            logger.error(f"{site_id}: Some data files are missing -----")
            return False
        return True

    @abstractmethod
    def process_actual_forecast(self, country: str, site_id: str) -> pd.DataFrame:
        """This method should be implement on child class"""
        pass

    @abstractmethod
    def process_eod(self, country: str, site_id: str) -> pd.DataFrame:
        """This method should be implement on child class"""
        pass

    @abstractmethod
    def process_availability(self, country: str, site_id: str) -> pd.DataFrame:
        """This method should be implement on child class"""
        pass

    def run(self, country: str, site_id: str) -> tuple[str,str]:
        """Method to clean data for single site. Each step reads every source file and
        only replaces the dates those file names cover."""
        try:
            should_process = self.data_files_check(country, site_id)
            if not should_process:
                logger.error(f"{site_id}: Need update data source files")
                return (site_id, "TRANSFORM DATA FAILED - Need update data source files")
            else:
                self.process_actual_forecast(country, site_id)
                self.process_eod(country, site_id)
                self.process_availability(country, site_id)
                logger.info(f"{site_id}: All data transformed successfully.")
                return (site_id, "TRANSFORM DATA SUCCESS")
        except Exception as e:
            logger.error(f"{site_id}: Error to transformed data - {e}")
            return (site_id, f"TRANSFORM DATA FAILED - {e}")

class CloudPmsTrans(TransformData):
    _source_checks = [("Actual Forecast", "xlsx")]

    def __init__(self):
        super().__init__()

    @staticmethod
    def _consumed_dates(source_files: pd.Series) -> pd.DatetimeIndex:
        """Every day covered by the Accounting file names' consumed range; the end date is exclusive."""
        ranges = (source_files.dropna().drop_duplicates()
                  .str.extract(r'consumed_(\d{2}-\d{2}-\d{4})_\d{2}-\d{2}-\d{2}_-_(\d{2}-\d{2}-\d{4})')
                  .dropna())
        days = [pd.date_range(pd.to_datetime(start, format='%d-%m-%Y'),
                              pd.to_datetime(end, format='%d-%m-%Y') - pd.Timedelta(days=1), freq='D')
                for start, end in ranges.itertuples(index=False)]
        return pd.DatetimeIndex(sorted(set().union(*days)))

    def process_actual_forecast(self, country: str, site_id: str) -> pd.DataFrame:
        """Method to clean cloud_pms data to get actual and forecast."""
        logger.info(f"{site_id}: Starting (CloudPmsTrans) process_actual_forecast")

        datasource_folder = self.report_data_folder / f"{country}_RAW" / f"{site_id}" / f"{site_id} Actual Forecast"
        act_fc_raw = self._read_bulk_excel(site_id, datasource_folder)
        if act_fc_raw.empty:
            logger.warning(f"{country}-{site_id}: No CloudPms Actual Forecast data to save.")
            return pd.DataFrame()

        act_fc_raw['site_id'] = site_id
        act_fc_raw['as_of'] = self._format_date(act_fc_raw['source_file'].str.extract(r'(\d{2}\.\d{2}\.\d{2})')[0])
        act_fc_raw['year'] = act_fc_raw['as_of'].dt.year
        act_fc_raw['quarter'] = act_fc_raw['as_of'].dt.quarter
        num_cols = ['room_sold', 'revenue_net', 'sleepers']
        act_fc_raw[num_cols] = act_fc_raw[num_cols].apply(pd.to_numeric, errors='coerce').fillna(0)

        actfc_merge = act_fc_raw.merge(self.market_segment[['site_id', 'code', 'ms_report']],
                                       left_on=['site_id', 'segment'],
                                       right_on=['site_id', 'code'],
                                       how="left")
        # Blank or unmapped segment (e.g. refund rows) would be dropped by groupby
        actfc_merge['ms_report'] = actfc_merge['ms_report'].fillna("MSNA")

        act_fc = (actfc_merge
                  .groupby(['site_id', 'year', 'quarter', 'as_of', 'consumed_date', 'ms_report'], as_index=False)
                  .agg(room_sold=('room_sold', 'sum'),
                       revenue_net=('revenue_net', 'sum'),
                       sleepers=('sleepers', 'sum'))
                  .rename(columns={'consumed_date': 'stay_date'}))

        act_fc['adr_net'] = act_fc['revenue_net'].div(act_fc['room_sold'].where(act_fc['room_sold'] != 0)).fillna(0).round(2)
        act_fc = act_fc.infer_objects(copy=False)

        if not act_fc.empty:

            self._save_grouped_csv(act_fc, "MBI Actual Forecast", country, site_id, True, self._date_key('as_of'))
            logger.info(f"{site_id}: Cleaned and saved CloudPms Actual Forecast data successfully")
            return act_fc
        else:
            logger.warning(f"{country}-{site_id}: No CloudPms Actual Forecast data to save.")
            return pd.DataFrame()

    def process_eod(self, country: str, site_id: str) -> pd.DataFrame:
        """Method to clean data to get accounting (EOD), using every file in the Accounting Report folder."""
        logger.info(f"{site_id}: Starting (MewTrans) process_eod")

        datasource_folder = self.report_data_folder / f"{country}_RAW" / f"{site_id}" / f"{site_id} Accounting Report"

        def get_eod(tab_name):
            df_raw = self._read_bulk_excel(site_id, datasource_folder, tab_name)
            df_raw.columns = df_raw.columns.astype(str).str.lower()

            consumed_dates = self._format_date(df_raw['consumed'].astype(str))
            eod_date = (consumed_dates - pd.to_timedelta(
                (consumed_dates.dt.hour < 5).astype(int), unit='D'
            )).dt.normalize()

            for col in ['net', 'vat', 'amount']:
                df_raw[col] = pd.to_numeric(df_raw[col].astype(str).replace('-', '0'), errors='coerce').fillna(0)

            df = (df_raw
                  .assign(site_id=site_id, year=eod_date.dt.year, stay_date=eod_date)
                  .drop(columns=['code'], errors='ignore')
                  .rename(columns={'type': 'section', 'ledger account code': 'code', 'vat': 'tax', 'amount': 'gross'}))
            df['source'] = tab_name
            df['code'] = df['code'].astype(str).str.replace(r'\.0$', '', regex=True)
            return df

        with ThreadPoolExecutor(max_workers=2) as ex:
            items_f = ex.submit(get_eod, "Items")
            outlets_f = ex.submit(get_eod, "Outlets")
            items = items_f.result()
            outlets = outlets_f.result()
        account = pd.concat([items, outlets], ignore_index=True)
        if account.empty:
            logger.warning(f"{site_id}: No CloudPms Accounting data.")
            return pd.DataFrame()

        file_dates = self._consumed_dates(account['source_file'])
        if file_dates.empty:
            logger.warning(f"{site_id}: No consumed date range in CloudPms Accounting file names.")
            return pd.DataFrame()
        dummy_table = pd.DataFrame([
            {'site_id': site_id, 'year': d.year, 'stay_date': d, 'source': 'Items',
             'section': sec, 'code': code, 'net': 0, 'tax': 0, 'gross': 0}
            for d in file_dates
            for sec, code in [('Revenue', '4000'), ('Payment', '1034')]
        ])

        account = pd.concat([account, dummy_table], ignore_index=True)
        account['quarter'] = account['stay_date'].dt.quarter

        eod = (account
               .groupby(['site_id', 'year', 'quarter', 'stay_date', 'section', 'source', 'code'], as_index=False)
               .agg({'net': 'sum', 'tax': 'sum', 'gross': 'sum'})
               .infer_objects(copy=False))
        # Only days the file names cover; a 05:00 shift onto the day before the range is dropped
        eod = eod[eod['stay_date'].isin(file_dates)]

        if not eod.empty:
            self._save_grouped_csv(eod, "CloudPms EOD", country, site_id, True, self._date_key('stay_date'))
            logger.info(f"{site_id}: Cleaned and saved CloudPms Accounting (EOD) data successfully")
            return eod
        else:
            logger.warning(f"{site_id}: No CloudPms Accounting data to save.")
            return pd.DataFrame()

    def process_availability(self, country: str, site_id: str) -> pd.DataFrame:
        """Method to clean data to get availability."""
        logger.info(f"{site_id}: Start (MewTrans) process_availability")

        datasource_folder = self.report_data_folder / f"{country}_RAW" / f"{site_id}" / f"{site_id} Availability Report"
        avai_raw = self._read_bulk_excel(site_id, datasource_folder, "Availability")

        avai_raw = avai_raw.loc[avai_raw['Service'] == "Total"].copy()

        avai_long = avai_raw.melt(
            id_vars=['Total'],
            var_name='date_raw',
            value_name='room_avai'
        )

        formatted_dates = self._format_date(avai_long['date_raw'])

        avai = (avai_long
                .assign(
                    site_id=site_id,
                    stay_date=formatted_dates.dt.date,
                    year=formatted_dates.dt.year,
                )
                .dropna(subset=['stay_date'])
                .astype({'year': 'int32'})
                .groupby(['site_id', 'year', 'stay_date'], as_index=False)
                .agg({'room_avai': 'sum'})
                .infer_objects(copy=False))

        if not avai.empty:
            self._save_grouped_csv(avai, "CloudPms Availability", country, site_id, False, self._date_key('stay_date'))
            logger.info(f"{country}-{site_id}: Cleaned and saved CloudPms Availability data successfully")
            return avai
        else:
            logger.warning(f"{country}-{site_id}: No CloudPms Availability data to save")
            return pd.DataFrame()

class PortalPmsTrans(TransformData):
    _source_checks = [("EOD", "csv"), ("Actual", "csv"), ("Forecast", "csv")]

    def __init__(self):
        super().__init__()

    def process_actual_forecast(self, country: str, site_id: str) -> pd.DataFrame:
        """Method to clean portal_pms data to get actual and forecast."""
        def clean_actual_data(country: str, site_id: str) -> pd.DataFrame:
            """Clean actual data"""
            logger.info(f"{site_id}: Cleaning Actual data.")
            folderpath = self.report_data_folder / f"{country}_RAW" / f"{site_id}" / f"{site_id} Actual"

            df = self._read_bulk_csv(site_id, folderpath)
            if df.empty:
                return pd.DataFrame()

            df.columns = df.columns.astype(str).str.lower()
            df = df[(df['rooms'] != 0) | (df['revenuenett'] != 0)].copy()

            df['as_of'] = self._format_date(df['source_file'].str.extract(r'(\d{2}\.\d{2}\.\d{2})')[0]) #[0] return a Series
            df['date'] = self._format_date(df['date'])
            df['site_id'] = site_id

            df.infer_objects(copy=False)

            days_to_subtract = np.where(site_id != config.pms_credentials['backdate_site_id'], 1, 0)
            df['stay_date'] = (df['date'] - pd.to_timedelta(days_to_subtract, unit='D')).dt.date

            df = df.merge(
                self.market_segment,
                left_on=['site_id', 'marketsegment'],
                right_on=['site_id', 'code'],
                how="left"
            )

            df['ms_report'] = df['ms_report'].fillna("MSNA")
            df['year'] = df['as_of'].dt.year
            df['quarter'] = df['as_of'].dt.quarter

            actual = (df
                      .groupby(['site_id','year','quarter','as_of','stay_date', 'ms_report'], as_index=False)
                      .agg({
                          'rooms': "sum",
                          'revenuenett': "sum",
                          'sleepers': "sum"
                          })
                        .rename(columns={
                            'rooms': 'room_sold',
                            'revenuenett': 'revenue_net'
                        })
            )

            actual[['room_sold', 'revenue_net', 'sleepers']] = actual[['room_sold', 'revenue_net', 'sleepers']].apply(pd.to_numeric, errors='coerce').fillna(0)
            actual['adr_net'] = actual['revenue_net'].div(actual['room_sold'].where(actual['room_sold'] != 0)).fillna(0).round(2)

            return actual.infer_objects(copy=False)

        def clean_forecast_data(country: str, site_id: str) -> pd.DataFrame:
            """Clean forecast data"""
            logger.info(f"{site_id}: Cleaning Forecast data.")
            folderpath = self.report_data_folder / f"{country}_RAW" / f"{site_id}" / f"{site_id} Forecast"

            df = self._read_bulk_csv(site_id, folderpath)
            if df.empty:
                return pd.DataFrame()

            df.columns = df.columns.astype(str).str.lower()

            df['as_of'] = self._format_date(df['source_file'].str.extract(r'(\d{2}\.\d{2}\.\d{2})')[0])
            df['stay_date'] = self._format_date(df['date']).dt.date
            df['year'] = df['as_of'].dt.year
            df['quarter'] = self._format_date(df['as_of']).dt.quarter

            df = df.infer_objects(copy=False)

            df = df.merge(
                self.market_segment,
                left_on=['site', 'marketsegment'],
                right_on=['site_id', 'code'],
                how="left"
            )

            df['ms_report'] = df['ms_report'].fillna("MSNA")

            forecast = df.groupby(
                ['site', 'year', 'quarter', 'as_of', 'stay_date', 'ms_report'],
                as_index=False
            ).agg({
                'rmnights': "sum",
                'accom revenue total': "sum",
                'sleepernights': "sum"
            }).rename(columns={
                'site': "site_id",
                'rmnights': "room_sold",
                'accom revenue total': "revenue_net",
                'sleepernights': "sleepers"
            })

            forecast[['room_sold', 'revenue_net', 'sleepers']] = forecast[['room_sold', 'revenue_net', 'sleepers']].apply(pd.to_numeric, errors='coerce').fillna(0)
            forecast['adr_net'] = forecast['revenue_net'].div(forecast['room_sold'].where(forecast['room_sold'] != 0)).fillna(0).round(2)

            return forecast.infer_objects(copy=False)

        actual_data = clean_actual_data(country, site_id)
        forecast_data = clean_forecast_data(country, site_id)

        to_combine = [df for df in [actual_data, forecast_data] if not df.empty]

        if not to_combine:
            logger.warning(f"{site_id}: No data found for processing. Skipping.")
            return pd.DataFrame()

        actual_forecast = pd.concat(to_combine, ignore_index=True)

        actual_forecast = actual_forecast.infer_objects(copy=False)

        if not actual_forecast.empty:
            self._save_grouped_csv(actual_forecast, "Actual Forecast", country, site_id, True, self._date_key('as_of'))
            logger.info(f"{site_id}: Cleaned and save PortalPms Actual Forecast data successfully.")
            return actual_forecast
        else:
            logger.warning(f"{site_id}: No PortalPms Actual Forecast data to save.")
            return pd.DataFrame()

    def process_eod(self, country: str, site_id: str) -> pd.DataFrame:
        """Method to clean portal_pms data to get eod, using every file in the EOD folder."""
        logger.info(f"{site_id}: Starting (PortalPms) process_eod")

        folderpath = self.report_data_folder / f"{country}_RAW" / site_id / f"{site_id} EOD"

        df = self._read_bulk_csv(site_id, folderpath)
        if df.empty:
            return pd.DataFrame()

        df.columns = df.columns.astype(str).str.lower()

        df[['netttotal', 'taxtotal', 'grosstotal']] = (
            df[['netttotal', 'taxtotal', 'grosstotal']]
            .apply(pd.to_numeric, errors='coerce')
            .fillna(0)
        )

        df['stay_date'] = self._format_date(df['source_file'].str.extract(r'(\d{2}\.\d{2}\.\d{2})')[0])
        df = df.dropna(subset=['stay_date'])
        if df.empty:
            logger.error(f"{site_id}: No dated PortalPms EOD files found.")
            return pd.DataFrame()
        # One file per day, so a day with no file gets no dummy row and keeps its old data
        file_dates = pd.DatetimeIndex(sorted(df['stay_date'].unique()))

        filters = ["Sales", "PMS Payments - Non Ledger CC", "PMS Payments - Ledger CC"]
        df = df[df["section"].isin(filters)].copy()

        df['year'] = df['stay_date'].dt.year
        df['site_id'] = site_id

        df = df.rename(columns={'netttotal': 'net', 'taxtotal': 'tax', 'grosstotal': 'gross'})
        dummy_table = pd.DataFrame([
            {'site_id': site_id, 'year': d.year, 'stay_date': d,
             'section': sec, 'code': code, 'net': 0, 'tax': 0, 'gross': 0}
            for d in file_dates
            for sec, code in [('Sales', 'ROOMS'), ('PMS Payments - Non Ledger CC', 'CASH')]
        ])

        eod = pd.concat([df, dummy_table], ignore_index=True)
        eod['quarter'] = eod['stay_date'].dt.quarter
        eod = (
            eod
            .groupby(["site_id", "year", "quarter", "stay_date", "section", "code"], as_index=False)
            .agg({"net": "sum", "tax": "sum", "gross": "sum"})
            .infer_objects(copy=False)
        )

        if not eod.empty:
            self._save_grouped_csv(eod, "EOD", country, site_id, True, self._date_key('stay_date'))
            logger.info(f"{site_id}: Cleaned and save PortalPms EOD data successfully.")
            return eod
        else:
            logger.error(f"{site_id}: No PortalPms EOD data to save.")
            return pd.DataFrame()

    def process_availability(self, country: str, site_id: str) -> pd.DataFrame:
        """Method to clean portal_pms data to get availability."""
        logger.info(f"{site_id}: Starting (PortalPms) process_availability.")

        folderpath = self.report_data_folder / f"{country}_RAW" / site_id / f"{site_id} Availability"

        df = self._read_bulk_csv(site_id, folderpath)
        if df.empty:
            return pd.DataFrame()

        df.columns = df.columns.astype(str).str.lower()

        df['year'] = pd.to_numeric(df['source_file'].str.extract(r'(\d{4})', expand=False), errors='coerce')

        id_vars = ['site', 'day', 'source_file', 'year']
        month_cols = [c for c in df.columns if 'month' in c]

        avai_melt = df.melt(
            id_vars=id_vars,
            value_vars=month_cols,
            var_name="month",
            value_name="room_avai"
        )

        avai_melt["month"] = pd.to_numeric(avai_melt["month"].str.replace("month", "", regex=False), errors="coerce")
        avai_melt["site_id"] = site_id

        avai_melt["stay_date"] = pd.to_datetime(avai_melt[["year", "month", "day"]], errors="coerce")

        avai = avai_melt.dropna(subset=['stay_date', 'room_avai']).copy()
        avai = avai[['site_id', 'year', 'stay_date', 'room_avai']].infer_objects(copy=False)

        if not avai.empty:
            self._save_grouped_csv(avai, "Availability", country, site_id)
            logger.info(f"{site_id}: Cleaned and saved PortalPms Availability data successfully.")
            return avai
        else:
            logger.error(f"{site_id}: No PortalPms Availability data to save.")
            return pd.DataFrame()
