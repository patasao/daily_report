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
import os

pd.set_option('future.no_silent_downcasting', True)

logger = get_logger('transform_data', config.paths['log_dir'])


class TransformData(ABC):
    """
    Abstract base class for all PMS data transformers.
    Handles SharePoint I/O, market-segment lookups, and the shared
    'changes' calculation; concrete subclasses implement the
    source-specific parsing for each PMS platform.
    """
    def __init__(self):
        # SharePoint root folder for all raw report data
        self.report_data_folder = Path("Report Data")
        self._token_manager = TokenManager(
            client_id=config.credentials['client_id'],
            authority=config.credentials['authority'],
            client_secret=config.credentials['client_secret'],
            scopes=config.credentials['scope']
        )
        self._sp_site_id = config.credentials['sp_site_id']
        # Market-segment mapping loaded once per run and shared across all sites
        self.ms_list = self._get_ms_list()

    def _sp_read(self) -> ReadSP:
        return ReadSP(self._token_manager.get_token(), self._sp_site_id)

    def _sp_upload(self) -> UploadSP:
        return UploadSP(self._token_manager.get_token(), self._sp_site_id)

    def _sp_path(self, relative_path: Path) -> str:
        return f"{config.credentials['sp_dor_root']}/{relative_path.as_posix()}"

    # --- Helpers ---
    def _get_ms_list(self) -> pd.DataFrame:
        """Load the market-segment code → display name mapping from SharePoint."""
        sp_path = self._sp_path(self.report_data_folder / "OTHER" / "MarketSegment.csv")
        try:
            return self._sp_read().read_csv(sp_path)
        except Exception as e:
            logger.warning(f"Could not load MarketSegment file: {e}")
            return pd.DataFrame()

    def _save_grouped_csv(self, df: pd.DataFrame, report_name: str, country: str, site_id: str):
        """
        Splits the DataFrame by year, writes each year-group to a temp CSV,
        and uploads it to the 'ALL PY TRANS' folder on SharePoint.
        Splitting by year keeps individual files small and enables
        year-over-year comparisons in the downstream Excel reports.
        """
        sp_folder = self._sp_path(self.report_data_folder / "ALL PY TRANS" / country / site_id)
        upload_sp = self._sp_upload()
        filename  = ""
        try:
            for yr, group in df.groupby("year"):
                filename = f"{site_id} {report_name} {yr}.csv"
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

    def _read_bulk_excel(self, site_id: str, folderpath: Path, sheet_name: str = 0) -> pd.DataFrame:
        """Read and concatenate all Excel files from a SharePoint folder into one DataFrame."""
        sp_path = self._sp_path(folderpath)
        try:
            df = self._sp_read().read_excel(sp_path, sheet_name=sheet_name)
            if df.empty:
                logger.warning(f"{site_id}: No Excel files at {sp_path}")
            return df
        except Exception as e:
            logger.error(f"{site_id}: Failed to read Excel from {sp_path}: {e}")
            return pd.DataFrame()

    def _read_bulk_csv(self, site_id: str, folderpath: Path) -> pd.DataFrame:
        """Read and concatenate all CSV files from a SharePoint folder into one DataFrame."""
        sp_path = self._sp_path(folderpath)
        try:
            df = self._sp_read().read_csv(sp_path)
            if df.empty:
                logger.warning(f"{site_id}: No CSV files at {sp_path}")
                return pd.DataFrame()
            df.rename(columns=str, inplace=True)
            df.dropna(how='all', axis=1, inplace=True)
            # Standardise the source-file column name across both PMS exports
            df.rename(columns={'source_file': 'Source.Name'}, inplace=True)
            return df
        except Exception as e:
            logger.error(f"{site_id}: Failed to read CSVs from {sp_path}: {e}")
            return pd.DataFrame()

    @staticmethod
    def _format_date(series: pd.Series) -> pd.Series:
        """Parse a mixed-format date column, returning NaT for unparseable values."""
        return pd.to_datetime(series, format='mixed', errors="coerce", dayfirst=True)

    # --- Processors ---
    def data_files_check(self, country: str, site_id: str, attempts: int = 3, wait_time: int = 30) -> bool:
        """
        Confirms required source files exist on SharePoint before processing begins.
        Retries up to `attempts` times with a `wait_time` second pause to handle
        the case where the extract step is still uploading.
        """
        logger.info(f"{site_id}: Checking last update data")

        as_of        = config.dates['as_of_date']
        target_dates = [as_of - timedelta(days=1), as_of]
        date_strings = [d.strftime('%d.%m.%y') for d in target_dates]
        data_folder  = self.report_data_folder / country / site_id
        cloud_pms_ids = {site[1] for site in config.group_ids['cloud_pms']}

        checks = [("Reservation Report", "xlsx")] if site_id in cloud_pms_ids else [
            ("EOD", "csv"), ("Actual", "csv"), ("Forecast", "csv")
        ]

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
        pass

    @abstractmethod
    def process_eod(self, country: str, site_id: str) -> pd.DataFrame:
        pass

    @abstractmethod
    def process_availability(self, country: str, site_id: str) -> pd.DataFrame:
        pass

    def process_changes(self, country: str, site_id: str, actual_forecast: pd.DataFrame) -> pd.DataFrame:
        """
        Computes day-over-day pickup changes in rooms and revenue by market segment.
        Joins today's snapshot against yesterday's and the same date from the prior year,
        then retains only rows where either rooms or revenue moved — enabling the
        Pickup and Variance tabs in the downstream Excel report.
        """
        logger.info(f"{site_id}: Starting process_changes.")

        as_of_date = config.dates['as_of_date']
        prev_year  = as_of_date.year - 1

        changes_raw  = actual_forecast.copy()
        sp_prev_path = self._sp_path(
            self.report_data_folder / "ALL PY TRANS" / country / site_id
            / f"{site_id} Actual Forecast {prev_year}.csv"
        )
        prev_year_df = self._sp_read().read_csv(sp_prev_path)
        if not prev_year_df.empty:
            last_as_of = prev_year_df['as_of'].max()
            prev_data  = prev_year_df[prev_year_df['as_of'] == last_as_of]
            changes_raw = pd.concat([changes_raw, prev_data], ignore_index=True)

        changes_raw['as_of']    = self._format_date(changes_raw['as_of']).dt.date
        changes_raw['date_dor'] = self._format_date(changes_raw['date_dor']).dt.date

        grouped = (changes_raw
                   .groupby(['site_id', 'year', 'as_of', 'date_dor', 'ms_dor'], as_index=False)
                   .agg({'rooms': 'sum', 'revenuenett': 'sum'}))

        # Build a full cartesian grid so missing market-segment/date combos show as 0
        idx_cols  = ['as_of', 'date_dor', 'ms_dor']
        full_grid = pd.MultiIndex.from_product(
            [grouped[c].unique() for c in idx_cols], names=idx_cols
        ).to_frame(index=False)
        full_grid['site_id'] = site_id
        full_grid['year']    = self._format_date(full_grid['as_of']).dt.year

        df_today = full_grid.merge(grouped, on=['site_id', 'year', 'as_of', 'date_dor', 'ms_dor'], how="left").fillna(0)
        df_today.rename(columns={'revenuenett': 'rev_today', 'rooms': 'rooms_today'}, inplace=True)
        df_today['yesterday'] = df_today['as_of'] - timedelta(days=1)

        # Join yesterday's snapshot to calculate the daily delta
        final = df_today.merge(
            grouped,
            left_on=['site_id', 'year', 'yesterday', 'date_dor', 'ms_dor'],
            right_on=['site_id', 'year', 'as_of', 'date_dor', 'ms_dor'],
            how='left',
            suffixes=('', '_prev')
        ).fillna(0)

        final["rev_changes"]  = (final["rev_today"] - final["revenuenett"]).round(2)
        final["room_changes"] = (final["rooms_today"] - final["rooms"]).astype(int)

        mask    = (final["rev_changes"] != 0) | (final["room_changes"] != 0)
        changes = (final.loc[mask & (final['year'] != prev_year),
                             ["site_id", "year", "as_of", "date_dor", "ms_dor", "rev_changes", "room_changes"]]
                   .infer_objects(copy=False))

        if not changes.empty:
            self._save_grouped_csv(changes, "Changes", country, site_id)
            logger.info(f"{country}-{site_id}: Cleaned and Saved Changes data successfully")
            return changes
        else:
            logger.warning(f"{country}-{site_id}: No Changes data to save")
            return pd.DataFrame()

    def run(self, country: str, site_id: str) -> tuple[str, str]:
        """Orchestrates the full transform pipeline for a single site."""
        try:
            should_process = self.data_files_check(country, site_id)
            if not should_process:
                logger.error(f"{site_id}: Need update data source files")
                return (site_id, "TRANSFORM DATA FAILED - Need update data source files")
            else:
                actual_forecast = self.process_actual_forecast(country, site_id)
                self.process_eod(country, site_id)
                self.process_availability(country, site_id)
                self.process_changes(country, site_id, actual_forecast)
                logger.info(f"{site_id}: All data transformed successfully.")
                return (site_id, "TRANSFORM DATA SUCCESS")
        except Exception as e:
            logger.error(f"{site_id}: Error to transformed data - {e}")
            return (site_id, f"TRANSFORM DATA FAILED - {e}")


class MewsTrans(TransformData):
    """
    Transforms cloud PMS exports (wide-format Excel) into the unified long schema.
    The Nights sheet is melted from a date-column pivot table into per-night rows,
    then joined to the Reservations sheet to attach market-segment codes before
    grouping into the standard (site_id, year, as_of, date_dor, ms_dor) grain.
    """
    def __init__(self):
        super().__init__()

    def process_actual_forecast(self, country: str, site_id: str) -> pd.DataFrame:
        """
        Merges the Nights (revenue pivot) and Reservations (segment codes) sheets,
        standardises both into a long-format DataFrame, and saves by year.
        """
        logger.info(f"{site_id}: Starting (MewsTrans) process_actual_forecast")

        datasource_folder = self.report_data_folder / country / f"{site_id}" / f"{site_id} Reservation Report"

        # --- Nights sheet: wide pivot (date columns) → long format ---
        night_raw = self._read_bulk_excel(site_id, datasource_folder, "Nights")
        night_raw.columns = night_raw.columns.astype(str).str.lower()

        as_of_dates = self._format_date(night_raw['source.name'].str.extract(r'(\d{2}\.\d{2}\.\d{2})')[0])

        night = (night_raw
                 .assign(as_of=as_of_dates.dt.date, site_id=site_id, year=as_of_dates.dt.year)
                 .drop(columns=['customer', 'assigned space', 'identifier'], errors='ignore')
                 # Melt date columns into rows so each row = one reservation-night
                 .melt(id_vars=['site_id', 'as_of', 'year', 'number'], var_name="date_dor", value_name="revenuenett")
                 .dropna(subset=['revenuenett'])
                 .infer_objects(copy=False))

        night['number'] = night['number'].astype(str)

        # --- Reservations sheet: provides market-segment codes per booking ---
        reser_raw = self._read_bulk_excel(site_id, datasource_folder, sheet_name='Reservations')
        reser_raw.columns = reser_raw.columns.astype(str).str.lower()

        res_as_of = self._format_date(reser_raw['source.name'].str.extract(r'(\d{2}\.\d{2}\.\d{2})')[0])

        reser = (reser_raw
                 .assign(as_of=res_as_of.dt.date, site_id=site_id, year=res_as_of.dt.year)
                 [['site_id', 'as_of', 'year', 'number', 'segment', 'person count']]
                 .dropna(subset=['number'])
                 .infer_objects(copy=False))

        reser['number'] = reser['number'].astype(str)

        # --- Merge nights + segments, then attach display-name market-segment ---
        merge    = night.merge(reser, on=['site_id', 'as_of', 'year', 'number'], how='left')
        merge['date_dor'] = self._format_date(merge['date_dor']).dt.date

        ms_merge = merge.merge(self.ms_list,
                               left_on=['site_id', 'segment'],
                               right_on=['SiteID', 'CODE'],
                               how="left").rename(columns={'MarketSegment': 'ms_dor'})

        act_fc = (ms_merge
                  .groupby(['site_id', 'year', 'as_of', 'date_dor', 'ms_dor'], as_index=False)
                  .agg({'number': 'count', 'revenuenett': 'sum', 'person count': 'sum'})
                  .rename(columns={'number': 'rooms', 'person count': 'sleepers'}))

        act_fc[['rooms', 'revenuenett', 'sleepers']] = act_fc[['rooms', 'revenuenett', 'sleepers']].apply(pd.to_numeric, errors='coerce').fillna(0)
        act_fc['adrnett'] = act_fc['revenuenett'].div(act_fc['rooms'].where(act_fc['rooms'] != 0)).fillna(0).round(2)
        act_fc = act_fc.infer_objects(copy=False)

        if not act_fc.empty:
            self._save_grouped_csv(act_fc, "Mews Actual Forecast", country, site_id)
            logger.info(f"{site_id}: Cleaned and saved Mews Actual Forecast data successfully")
            return act_fc
        else:
            logger.warning(f"{country}-{site_id}: No Mews Actual Forecast data to save.")
            return pd.DataFrame()

    def process_eod(self, country: str, site_id: str) -> pd.DataFrame:
        """
        Cleans accounting (EOD) data from the cloud PMS.
        Adjusts for the 5 AM night-audit cut-off: transactions before 05:00
        are credited to the previous calendar day.
        """
        logger.info(f"{site_id}: Starting (MewTrans) process_eod")

        datasource_folder = self.report_data_folder / country / f"{site_id}" / f"{site_id} Accounting Report"

        def get_eod(tab_name):
            df_raw = self._read_bulk_excel(site_id, datasource_folder, tab_name)
            df_raw.columns = df_raw.columns.astype(str).str.lower()

            consumed_dates = self._format_date(df_raw['consumed'])
            # Transactions before 05:00 belong to the previous EOD date
            eod_date = (consumed_dates - pd.to_timedelta(
                (consumed_dates.dt.hour < 5).astype(int), unit='D'
            )).dt.date

            cols_to_fix = ['net', 'vat', 'amount']
            for col in cols_to_fix:
                df_raw[col] = pd.to_numeric(df_raw[col].replace('-', '0'), errors='coerce').fillna(0)

            df = (df_raw
                  .assign(date=eod_date, year=consumed_dates.dt.year, site_id=site_id)
                  .groupby(['site_id', 'year', 'date', 'type', 'ledger account code'], as_index=False)
                  .agg({'net': 'sum', 'vat': 'sum', 'amount': 'sum'})
                  .rename(columns={
                      'type': 'section',
                      'ledger account code': 'code',
                      'net': 'nett',
                      'vat': 'tax',
                      'amount': 'gross',
                  })
                  .infer_objects(copy=False))
            df['source'] = tab_name
            return df

        items   = get_eod("Items")
        outlets = get_eod("Outlets")
        account = pd.concat([items, outlets], ignore_index=True)

        if not account.empty:
            self._save_grouped_csv(account, "Mews EOD", country, site_id)
            logger.info(f"{site_id}: Cleaned and saved Mews Accounting (EOD) data successfully")
            return account
        else:
            logger.warning(f"{site_id}: No Mews Accounting data to save.")
            return pd.DataFrame()

    def process_availability(self, country: str, site_id: str) -> pd.DataFrame:
        """Melts the wide-format availability pivot into a per-date long format."""
        logger.info(f"{site_id}: Start (MewTrans) process_availability")

        datasource_folder = self.report_data_folder / country / f"{site_id}" / f"{site_id} Availability Report"
        avai_raw = self._read_bulk_excel(site_id, datasource_folder, "Availability")

        avai_raw = avai_raw.loc[avai_raw['Service'] == "Total"].copy()

        # Wide → long: each date column becomes a row
        avai_long = avai_raw.melt(
            id_vars=['Total'],
            var_name='date_raw',
            value_name='room_avai'
        )

        formatted_dates = self._format_date(avai_long['date_raw'])

        avai = (avai_long
                .assign(site_id=site_id, date=formatted_dates.dt.date, year=formatted_dates.dt.year)
                .dropna(subset=['date'])
                .astype({'year': 'int32'})
                .groupby(['site_id', 'year', 'date'], as_index=False)
                .agg({'room_avai': 'sum'})
                .infer_objects(copy=False))

        if not avai.empty:
            self._save_grouped_csv(avai, "Mews Availability", country, site_id)
            logger.info(f"{country}-{site_id}: Cleaned and saved Mews Availability data successfully")
            return avai
        else:
            logger.warning(f"{country}-{site_id}: No Mews Availability data to save")
            return pd.DataFrame()


class PmsTrans(TransformData):
    """
    Transforms on-premise PMS CSV exports into the unified long schema.
    Actual and Forecast reports arrive in different schemas from the PMS web portal
    and are each cleaned and then concatenated into one combined actual_forecast table.
    """
    def __init__(self):
        super().__init__()

    def process_actual_forecast(self, country: str, site_id: str) -> pd.DataFrame:
        """
        Cleans and merges the Actual and Forecast CSVs into the shared schema.
        Both sources are normalised to (site_id, year, as_of, date_dor, ms_dor,
        rooms, revenuenett, sleepers, adrnett) before concatenation, enabling
        a single Power Query connection to serve both actuals and the rolling forecast.
        """
        def clean_actual_data(country: str, site_id: str) -> pd.DataFrame:
            """Parse the on-premise PMS actual CSV and standardise to the common schema."""
            logger.info(f"{site_id}: Cleaning Actual data.")
            folderpath = self.report_data_folder / country / site_id / f"{site_id} Actual"

            df = self._read_bulk_csv(site_id, folderpath)
            if df.empty:
                return pd.DataFrame()

            df.columns = df.columns.astype(str).str.lower()
            df = df[(df['rooms'] != 0) | (df['revenuenett'] != 0)].copy()

            # Extract the as-of date from the source filename (e.g. "SITE_A002 Actual 05.05.25.csv")
            df['as_of']   = self._format_date(df['source.name'].str.extract(r'(\d{2}\.\d{2}\.\d{2})')[0])
            df['date']    = self._format_date(df['date'])
            df['site_id'] = site_id
            df.infer_objects(copy=False)

            # One site's data arrives one day ahead — subtract a day to align stay dates
            days_to_subtract = np.where(site_id != config.pms_credentials['backdate_site_id'], 1, 0)
            df['date_dor'] = (df['date'] - pd.to_timedelta(days_to_subtract, unit='D')).dt.date

            df = df.merge(
                self.ms_list[['SiteID', 'CODE', 'MarketSegment']],
                left_on=['site_id', 'marketsegment'],
                right_on=['SiteID', 'CODE'],
                how="left"
            )

            df['ms_dor'] = df['MarketSegment'].fillna("NoMS")
            df['year']   = df['as_of'].dt.year

            actual = df.groupby(
                ['site_id', 'year', 'as_of', 'date_dor', 'ms_dor'], as_index=False
            ).agg({'rooms': "sum", 'revenuenett': "sum", 'sleepers': "sum"})

            actual[['rooms', 'revenuenett', 'sleepers']] = actual[['rooms', 'revenuenett', 'sleepers']].apply(pd.to_numeric, errors='coerce').fillna(0)
            actual['adrnett'] = actual['revenuenett'].div(actual['rooms'].where(actual['rooms'] != 0)).fillna(0).round(2)

            return actual.infer_objects(copy=False)

        def clean_forecast_data(country: str, site_id: str) -> pd.DataFrame:
            """Parse the on-premise PMS forecast CSV and standardise to the common schema."""
            logger.info(f"{site_id}: Cleaning Forecast data.")
            folderpath = self.report_data_folder / country / site_id / f"{site_id} Forecast"

            df = self._read_bulk_csv(site_id, folderpath)
            if df.empty:
                return pd.DataFrame()

            df.columns = df.columns.astype(str).str.lower()

            df['as_of']    = self._format_date(df['source.name'].str.extract(r'(\d{2}\.\d{2}\.\d{2})')[0])
            df['date_dor'] = self._format_date(df['date']).dt.date
            df['year']     = df['as_of'].dt.year
            df = df.infer_objects(copy=False)

            df = df.merge(self.ms_list, left_on=['site', 'marketsegment'], right_on=['SiteID', 'CODE'], how="left")
            df['ms_dor'] = df['MarketSegment'].fillna("DIRECT")

            forecast = df.groupby(
                ['site', 'year', 'as_of', 'date_dor', 'ms_dor'], as_index=False
            ).agg({
                'rmnights': "sum",
                'accom revenue total': "sum",
                'sleepernights': "sum",
            }).rename(columns={
                'site': "site_id",
                'rmnights': "rooms",
                'accom revenue total': "revenuenett",
                'sleepernights': "sleepers",
            })

            forecast[['rooms', 'revenuenett', 'sleepers']] = forecast[['rooms', 'revenuenett', 'sleepers']].apply(pd.to_numeric, errors='coerce').fillna(0)
            forecast['adrnett'] = forecast['revenuenett'].div(forecast['rooms'].where(forecast['rooms'] != 0)).fillna(0).round(2)

            return forecast.infer_objects(copy=False)

        actual_data   = clean_actual_data(country, site_id)
        forecast_data = clean_forecast_data(country, site_id)

        to_combine = [df for df in [actual_data, forecast_data] if not df.empty]

        if not to_combine:
            logger.warning(f"{site_id}: No data found for processing. Skipping.")
            return pd.DataFrame()

        # Concatenate actuals and forecast into one unified table
        actual_forecast = pd.concat(to_combine, ignore_index=True).infer_objects(copy=False)

        if not actual_forecast.empty:
            self._save_grouped_csv(actual_forecast, "Actual Forecast", country, site_id)
            logger.info(f"{site_id}: Cleaned and saved PMS Actual Forecast data successfully.")
            return actual_forecast
        else:
            logger.warning(f"{site_id}: No PMS Actual Forecast data to save.")
            return pd.DataFrame()

    def process_eod(self, country: str, site_id: str) -> pd.DataFrame:
        """
        Cleans the on-premise PMS Period End (EOD) CSV.
        Filters to relevant ledger sections (Sales, PMS Payments) and groups
        to the (site_id, year, date, section, code) grain before upload.
        """
        logger.info(f"{site_id}: Starting (PmsTrans) process_eod")

        folderpath = self.report_data_folder / country / site_id / f"{site_id} EOD"

        df = self._read_bulk_csv(site_id, folderpath)
        if df.empty:
            return pd.DataFrame()

        df.columns = df.columns.astype(str).str.lower()

        df = df.fillna({
            'section': 'Sales',
            'code': 'ROOMS',
            'netttotal': 0,
            'taxtotal': 0,
            'grosstotal': 0,
        })

        # Retain only the ledger sections relevant to revenue reporting
        filters = ["Sales", "PMS Payments - Non Ledger CC", "PMS Payments - Ledger CC"]
        df = df[df["section"].isin(filters)].copy()

        df['date']    = self._format_date(df['source.name'].str.extract(r'(\d{2}\.\d{2}\.\d{2})')[0])
        df['year']    = df['date'].dt.year
        df['site_id'] = site_id
        df = df.infer_objects(copy=False)

        eod = (df
               .groupby(["site_id", "year", "date", "section", "code"], as_index=False)
               .agg({"netttotal": "sum", "taxtotal": "sum", "grosstotal": "sum"})
               .rename(columns={"netttotal": "nett", "taxtotal": "tax", "grosstotal": "gross"}))

        if not eod.empty:
            self._save_grouped_csv(eod, "EOD", country, site_id)
            logger.info(f"{site_id}: Cleaned and saved PMS EOD data successfully.")
            return eod
        else:
            logger.error(f"{site_id}: No PMS EOD data to save.")
            return pd.DataFrame()

    def process_availability(self, country: str, site_id: str) -> pd.DataFrame:
        """
        Cleans the on-premise PMS availability CSV.
        The raw report is a wide pivot (month columns × day rows); this method
        melts it into a long per-date format consistent with the cloud PMS output.
        """
        logger.info(f"{site_id}: Starting (PmsTrans) process_availability.")

        folderpath = self.report_data_folder / country / site_id / f"{site_id} Availability"

        df = self._read_bulk_csv(site_id, folderpath)
        if df.empty:
            return pd.DataFrame()

        df.columns = df.columns.astype(str).str.lower()
        df['year'] = pd.to_numeric(df['source.name'].str.extract(r'(\d{4})', expand=False), errors='coerce')

        id_vars    = ['site', 'day', 'source.name', 'year']
        month_cols = [c for c in df.columns if 'month' in c]

        # Melt month columns into rows to produce one row per date
        avai_melt = df.melt(
            id_vars=id_vars,
            value_vars=month_cols,
            var_name="month",
            value_name="room_avai"
        )

        avai_melt["month"]   = pd.to_numeric(avai_melt["month"].str.replace("month", "", regex=False), errors="coerce")
        avai_melt["site_id"] = site_id
        avai_melt["date"]    = pd.to_datetime(avai_melt[["year", "month", "day"]], errors="coerce")

        avai = avai_melt.dropna(subset=['date', 'room_avai']).copy()
        avai = avai[['site_id', 'year', 'date', 'room_avai']].infer_objects(copy=False)

        if not avai.empty:
            self._save_grouped_csv(avai, "Availability", country, site_id)
            logger.info(f"{site_id}: Cleaned and saved PMS Availability data successfully.")
            return avai
        else:
            logger.error(f"{site_id}: No PMS Availability data to save.")
            return pd.DataFrame()
