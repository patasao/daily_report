from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException, TimeoutException, ElementNotInteractableException, StaleElementReferenceException, ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains
import pyotp
import requests
import threading
import time
from datetime import datetime, timedelta
import functools
import shutil
import tempfile
import os
from pathlib import Path
import config
from typing import Type, Union, Tuple, Any, Callable
from abc import ABC
from services.logger import get_logger
from services.ms_auth import TokenManager
from services.sharepoint import ReadSP, UploadSP

ExceptionType = Union[Type[BaseException], Tuple[Type[BaseException], ...]]

logger = get_logger('extract_data', config.paths['log_dir'])

# Status text for a 507 upload error; pipeline.py matches on it to skip retries
STORAGE_FULL = "SHAREPOINT STORAGE FULL"

class SiteMismatchError(Exception):
    """Browser is on another site; retrying would only download that site's reports."""

class SeleniumExtract(ABC):
    def __init__(self):
        """Initializes the base setting."""
        self.waittime = 60
        self.check_interval = 5 # wait time to check/retry again
        self.max_retries = 3

    @staticmethod
    def retry(exceptions: ExceptionType, tries: int = 3, delay: int = 5, backoff: int = 2):
        """
        A decorator function that automatically retries execution of the wrapped function
        when one of the specified exceptions is raised.
        """
        def decorator_retry(func: Callable) -> Callable:
            @functools.wraps(func)
            def wrapper_retry(*args: Any, **kwargs: Any) -> Any:
                _tries, _delay = tries , delay

                # --- Find the extractor and driver from the call arguments ---
                owner = args[0] if args and isinstance(args[0], SeleniumExtract) else None
                driver = kwargs.get('driver')
                if not driver:
                    for arg in args:
                        if hasattr(arg, 'refresh'):
                            driver = arg
                            break

                while _tries > 1:
                    try:
                        return func(*args, **kwargs)

                    except exceptions as e:
                        logger.warning(f"Retry Loop: {func.__name__} failed with {type(e).__name__}. Retrying in {_delay}s...")

                        if driver and owner:
                            owner._save_screenshot(driver, func.__name__)
                            try:
                                logger.info(f"Decorator resetting page for {func.__name__}")
                                owner._reset_page(driver)
                                time.sleep(1) # Give the reload a second to breathe
                            except SiteMismatchError:
                                raise
                            except Exception as reset_err:
                                logger.error(f"Could not reset page: {reset_err}")

                        time.sleep(_delay)
                        _tries -= 1
                        _delay *= backoff

                # Final attempt: if it fails here, it raises the error to the caller
                try:
                    return func(*args, **kwargs)
                except exceptions:
                    if driver and owner:
                        owner._save_screenshot(driver, func.__name__)
                    raise

            return wrapper_retry
        return decorator_retry

    def _reset_page(self, driver: WebDriver):
        """Puts the browser back in a known state before a retry."""
        driver.refresh()

    def _save_screenshot(self, driver: WebDriver, name: str):
        """Saves what the browser shows to log/screenshots; never raises."""
        try:
            shot_dir = config.paths['log_dir'] / 'screenshots'
            shot_dir.mkdir(parents=True, exist_ok=True)
            shot_path = shot_dir / f"{name}_{datetime.now().strftime('%d%m%Y%H%M%S')}.png"
            driver.save_screenshot(str(shot_path))
            logger.info(f"Saved screenshot: {shot_path}")
        except Exception as e:
            logger.warning(f"Could not save screenshot for {name}: {e}")

    @staticmethod
    def init_driver(download_dir: Path, profile_dir: Path = None) -> WebDriver:
        """Innitializes and setting a Chrome WevDriver."""
        logger.info("Initializing WebDriver")

        if download_dir.exists():
            shutil.rmtree(download_dir)
            logger.info(f"Deleted existing download directory: {download_dir}")

        download_dir.mkdir(parents=True, exist_ok=True)
        logger.info(f"Created new download directory: {download_dir}")

        options = Options()
        if config.env_status == "test":
            None # Deactivate run in background to inspect the automation
        else: options.add_argument("--headless")  # Activate run in background
        options.add_argument("--log-level=3")   # Suppress warnings/errors
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        prefs = {"download.default_directory": str(download_dir)}
        options.add_experimental_option("prefs", prefs)

        if profile_dir:
            profile_dir.mkdir(parents=True, exist_ok=True)
            options.add_argument(f"--user-data-dir={profile_dir}")
            logger.info(f"Using Chrome profile: {profile_dir}")

        try:
            driver = webdriver.Chrome(options=options)
        except WebDriverException as e:
            logger.error(f"Failed to initialize WebDriver: {e}")
            raise
        logger.info("WebDriver initialized successfully")
        return driver

    def _wait_for_file(self, download_dir: Path, new_name: str = None) -> bool:
        """
        Waits for a file download to complete in the specified directory, identifies the
        newest downloaded file, confirms its stability (by checking file size over time)
        to ensure the download is finished, and then renames it.
        """
        start_time = time.time()

        # Capture initial state to find 'new' files only
        preexisting_files = set(p.name for p in download_dir.iterdir() if p.is_file())
        retry_count = 0

        while True:
            current_files = set(p.name for p in download_dir.iterdir() if p.is_file())
            new_files = current_files - preexisting_files

            if new_files:
                # Sort to get the most recent file
                new_files_sorted = sorted(
                    new_files,
                    key=lambda name: (download_dir / name).stat().st_mtime,
                    reverse=True
                )
                latest_file = new_files_sorted[0]
                latest_file_path = download_dir / latest_file

                # 1. Skip temporary browser extensions
                if latest_file_path.suffix in (".crdownload", ".part", ".tmp"):
                    logger.info(f"File '{latest_file}' is still downloading (temp extension detected)...")
                else:
                    # 2. Wait for file size to stabilize (ensures download is complete)
                    previous_size = -1
                    size_stable_count = 0

                    while size_stable_count < 3:
                        if time.time() - start_time > self.waittime:
                            logger.error(f"Timeout reached while waiting for '{latest_file}' to stabilize.")
                            raise WebDriverException(f"Timeout reached while waiting for {latest_file} to stabilize.")

                        current_size = latest_file_path.stat().st_size
                        if current_size == previous_size and current_size > 0:
                            size_stable_count += 1
                        else:
                            size_stable_count = 0
                            previous_size = current_size

                        logger.info(f"Stabilizing '{latest_file}': {current_size} bytes...")
                        time.sleep(1) # Short sleep for stability check

                    # 3. Handle Renaming Logic
                    if new_name:
                        new_file_path = download_dir / new_name
                        try:
                            if new_file_path.exists():
                                logger.warning(f"Target '{new_name}' exists. Deleting old version.")
                                new_file_path.unlink()

                            latest_file_path.rename(new_file_path)
                            logger.info(f"File renamed to: {new_name}")
                            return True
                        except PermissionError:
                            retry_count += 1
                            if retry_count > self.max_retries:
                                logger.error(f"File locked, could not rename to '{new_name}' after {self.max_retries} retries.")
                                raise WebDriverException(f"File locked, could not rename to '{new_name}' after {self.max_retries} retries.")
                            logger.warning(f"File locked. Retrying rename ({retry_count}/{self.max_retries})...")
                            time.sleep(self.check_interval)
                            continue
                    else:
                        # Keep original name
                        logger.info(f"Download complete. Keeping original name: {latest_file}")
                        return True

            # Global Timeout Check
            if time.time() - start_time > self.waittime:
                logger.error("File download timed out.")
                raise WebDriverException("File download timed out.")

            time.sleep(self.check_interval)

    def _retry_switch_to_frame(self, driver: WebDriver, frame_id: str):
        """Attempts repeatedly to switch the WebDriver's context into a specified iframe."""
        for attempt in range(self.max_retries):
            try:
                logger.info(f"Attempt {attempt + 1} to switch to frame {frame_id}")

                # Ensure frame is present and visible
                frame = WebDriverWait(driver, self.waittime).until(
                    EC.presence_of_element_located((By.ID, frame_id))
                )

                # Alternative 1: Scroll using JavaScript to specific coordinates
                frame_location = frame.location_once_scrolled_into_view  # Get frame location
                driver.execute_script(f"window.scrollTo({frame_location['x']}, {frame_location['y']});")
                time.sleep(1)  # Allow time for scroll effect

                # Alternative 2: Scroll the parent container if frame inside scrollable div
                parent_scroll = driver.execute_script("return arguments[0].parentNode;", frame)
                driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", parent_scroll)
                time.sleep(1)

                # Alternative 3: Move to element using Actions class
                ActionChains(driver).move_to_element(frame).perform()

                # Try switching to the frame after scrolling
                WebDriverWait(driver, self.waittime).until(EC.frame_to_be_available_and_switch_to_it((By.ID, frame_id)))
                logger.info(f"Successfully switched to frame {frame_id}")
                return  # Exit if successful
            except (TimeoutException, ElementNotInteractableException) as e:
                logger.warning(f"Failed to switch to frame {frame_id} on attempt {attempt + 1}: {e}")
                time.sleep(2)  # Wait before retrying

        logger.error(f"Failed to switch to frame '{frame_id}' after {self.max_retries} attempts.")
        raise WebDriverException(f"Failed to switch to frame {frame_id} after {self.max_retries} attempts.") # If all retries fail, raise an exception

    def _click_element(self, driver: WebDriver, by: str, value: str):
        """Clicks an element after scrolling it into view and waiting for it to be clickable."""
        end_time = time.time() + self.waittime

        while time.time() < end_time:
            try:
                element = WebDriverWait(driver, self.waittime).until(
                    EC.element_to_be_clickable((by, value))
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                element.click()
                return

            except (StaleElementReferenceException, ElementClickInterceptedException) as e:
                last_exception = e
                logger.debug(f"Click interrupted on {value}, retrying... ({e})")
                time.sleep(0.5)
                continue

            except WebDriverException as e:
                last_exception = e
                logger.warning(f"Standard click failed on {value}. Attempting JS Click. Error: {e}")
                try:
                    element = driver.find_element(by, value)
                    driver.execute_script("arguments[0].click();", element)
                    logger.info(f"JS Click successful on {value}")
                    return
                except Exception as js_e:
                    logger.warning(f"JS Click also failed: {js_e}")
                    time.sleep(0.5)

            except Exception as e:
                last_exception = e
                logger.warning(f"Unexpected error clicking {value}: {e}")
                time.sleep(0.5)

        last_exception = WebDriverException(f"Timed out clicking element: {value}")
        logger.error(f"Timed out clicking element: {value}")
        raise last_exception

    def _enter_date(self, driver: WebDriver, by: str, value: str, date_str: str):
        """Helper to clear date field and enter date value."""
        logger.debug(f"Entering date '{date_str}' into field '{value}'")
        try:
            date_field = WebDriverWait(driver, self.waittime).until(
                EC.element_to_be_clickable((by, value))
            )
            date_field.clear()
            date_field.send_keys(date_str)
            logger.debug(f"Date '{date_str}' entered successfully into field '{value}'")
        except Exception as e:
            logger.error(f"Failed to enter date '{date_str}' into field '{value}': {e}")
            raise

class PortalPms(SeleniumExtract):
    _totp_lock = threading.Lock()

    def __init__(self, backdate: int = 1):
        super().__init__()
        self.backdate = backdate
        self.pms_user = config.portal_pms['pms_user']
        self.pms_pass = config.portal_pms['pms_pass']
        self.profiles_dir = config.paths['home_dir'] / '.pms_profiles'
        self._token_manager = TokenManager(
            client_id=config.credentials['client_id'],
            authority=config.credentials['authority'],
            client_secret=config.credentials['client_secret'],
            scopes=config.credentials['scope']
        )
        self._sp_site_id = config.credentials['sp_site_id']
        self._home_url = None  # site landing page, saved after login
        self._site_id = None   # site logged in to, checked again after each page reset

    def _reset_page(self, driver: WebDriver):
        """Reloads the landing page saved at login, then checks it is still the same site."""
        driver.switch_to.default_content()
        if not self._home_url:
            driver.refresh()
            return
        driver.get(self._home_url)
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//button[@data-testid="ok-button"]'))).click()
            logger.info("Closed pop up after page reset")
        except TimeoutException:
            pass
        try:
            self._verify_site(driver, self._site_id)
        except WebDriverException as e:
            raise SiteMismatchError(f"{self._site_id}: could not confirm site after page reset - {e}") from e

    def _sp_upload(self) -> UploadSP:
        return UploadSP(self._token_manager.get_token(), self._sp_site_id)

    def _sp_path(self, relative_path: Path) -> str:
        return f"{config.credentials['sp_report_root']}/{relative_path.as_posix()}"

    # --- Web automation ---
    @SeleniumExtract.retry(WebDriverException, tries=5)
    def _login(self, driver: WebDriver, site_id: str):
        """Nevigate to login page, login if needed, and entry specified site_id."""
        logger.info(f"{site_id}: Starting login process")

        driver.get(config.portal_pms['base_url'])
        logger.info(f"{site_id}: Navigated to login page")

        driver.maximize_window()

        # Check if already logged in — username field won't be visible if session is active
        try:
            WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.NAME, "username"))).send_keys(self.pms_user)
            logger.info(f"{site_id}: Entered username")

            driver.find_element(By.NAME, "password").send_keys(self.pms_pass)
            logger.info(f"{site_id}: Entered password")

            driver.find_element(By.NAME, "submit").click()
            logger.info(f"{site_id}: Clicked submit button")
        except TimeoutException:
            logger.info(f"{site_id}: Session still active, skipping credential entry")

        # Handle 2FA if prompted for one-time code
        try:
            totp_input = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "code")))
            logger.info(f"{site_id}: 2FA prompt detected, waiting for TOTP lock")

            with PortalPms._totp_lock:
                TOTP_SECRET = config.credentials['TOTP_SECRET']
                totp = pyotp.TOTP(TOTP_SECRET)
                totp_input.send_keys(totp.now())
                logger.info(f"{site_id}: Entered TOTP code")

                driver.find_element(By.XPATH, "//span[text()='Remember this device for 30 days']").click()
                logger.info(f"{site_id}: Clicked 'Remember this device for 30 days'")

                driver.find_element(By.CSS_SELECTOR, "button[data-action-button-primary='true']").click()
                logger.info(f"{site_id}: Clicked Continue on 2FA page")

                # Wait until current TOTP window expires so next site gets a fresh unused code
                seconds_remaining = 30 - (int(time.time()) % 30)
                logger.info(f"{site_id}: Waiting {seconds_remaining}s for next TOTP window")
                time.sleep(seconds_remaining + 1)
        except TimeoutException:
            logger.info(f"{site_id}: No 2FA prompt, continuing")

        WebDriverWait(driver, self.waittime).until(EC.element_to_be_clickable((By.ID, "oc-navbar-appswitch"))).click()
        logger.info(f"{site_id}: Clicked app switch button to select site")

        site_input = WebDriverWait(driver, self.waittime).until(EC.visibility_of_element_located((By.XPATH, "//div/div/form/div/div/div/input")))
        site_input.clear()
        site_input.send_keys(site_id)
        logger.info(f"{site_id}: Entered site ID")

        time.sleep(1)  # wait for autocomplete dropdown to populate
        driver.find_element(By.XPATH, "//div/div/form/div/div/div/input").send_keys(Keys.DOWN)
        driver.find_element(By.XPATH, "//div/div/form/div/div/div/input").send_keys(Keys.ENTER)
        logger.info(f"{site_id}: Selected site ID from dropdown")

        WebDriverWait(driver, self.waittime).until(EC.element_to_be_clickable((By.XPATH, '//button[@type="submit" and contains(@class, "oc-nav-MuiButton-root")]'))).click()
        logger.info(f"{site_id}: Clicked submit button to go to site")

        try:
            WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.NAME, "username")))
            driver.find_element(By.NAME, "username").send_keys(self.pms_user)
            driver.find_element(By.NAME, "password").send_keys(self.pms_pass)
            logger.info(f"{site_id}: Re-login required, submitted credentials for site")
        except TimeoutException:
            logger.info(f"{site_id}: No credential form, continuing")

        try:
            WebDriverWait(driver, self.waittime).until(
                EC.element_to_be_clickable((By.XPATH, f"//span[@class='MuiButton-label' and text()='Log in to {site_id}']"))
            ).click()
            logger.info(f"{site_id}: Clicked 'Log in to {site_id}' button")
        except TimeoutException:
            logger.info(f"{site_id}: No 'Log in to {site_id}' button, continuing")

        try:
            WebDriverWait(driver, self.waittime).until(EC.element_to_be_clickable((By.XPATH, '//button[@data-testid="ok-button"]'))).click()
            logger.info(f"{site_id}: Clicked pop up")
        except:
            logger.info(f"{site_id}: No pop up to click")

        self._verify_site(driver, site_id)
        self._home_url = driver.current_url
        self._site_id = site_id
        logger.info(f"{site_id}: Login successful for site")

    def _verify_site(self, driver: WebDriver, site_id: str):
        """Check the active site matches site_id; re-switches up to max_retries times if not."""
        for attempt in range(self.max_retries):
            current_site = WebDriverWait(driver, self.waittime).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, 'span[data-testid="siteIdText"]'))
            ).text.strip()

            if current_site == site_id:
                logger.info(f"{site_id}: Site verified — active site is '{current_site}'")
                return

            logger.warning(
                f"{site_id}: Site mismatch! Active site is '{current_site}', expected '{site_id}'. "
                f"Re-switching (attempt {attempt + 1}/{self.max_retries})"
            )

            self._click_element(driver, By.CSS_SELECTOR, 'button[data-testid="siteInformation"]')
            logger.info(f"{site_id}: Clicked site switcher button")

            site_input = WebDriverWait(driver, self.waittime).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[placeholder="Search by Site Name or ID"]'))
            )
            site_input.clear()
            site_input.send_keys(site_id)
            logger.info(f"{site_id}: Entered site ID in search box")

            time.sleep(1)
            site_input.send_keys(Keys.ENTER)
            logger.info(f"{site_id}: Submitted site selection")

            time.sleep(2)

        raise WebDriverException(
            f"{site_id}: Site mismatch persists after {self.max_retries} re-switch attempts."
        )

    @SeleniumExtract.retry(WebDriverException)
    def _generate_actual_report(self, driver: WebDriver, site_id: str, download_dir: Path):
        """Generate actual report from downloading 'Revenue By' report as csv."""
        logger.info(f"{site_id}: Generating actual report.")

        driver.switch_to.default_content()

        self._click_element(driver, By.CSS_SELECTOR, 'button[aria-label="Menu"]')
        logger.info(f"{site_id}: Clicked on menu button")

        self._click_element(driver, By.XPATH, "//span[text()='Revenue']")
        logger.info(f"{site_id}: Clicked on Revenue button")

        self._click_element(driver, By.XPATH, "//span[contains(@class, 'oc-nav-MuiTypography-bodySmall') and normalize-space()='Revenue By']")
        logger.info(f"{site_id}: Clicked on Revenue By button")

        WebDriverWait(driver, self.waittime).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, "progress-indicator-background"))
        )

        self._retry_switch_to_frame(driver, "MainFrame")

        self._click_element(driver, By.ID, "winp_Type")
        dropdown = driver.find_element(By.ID, "winp_Type")
        dropdown.find_element(By.XPATH, "//option[. = 'Market Segment']").click()
        logger.info(f"{site_id}: Selected 'Market Segment' in Type dropdown")

        self._retry_switch_to_frame(driver, "frmMainRepSection")

        self._click_element(driver, By.ID, "ctl00_InputParameters_ps_Period_winp_PeriodDateType")
        dropdown = driver.find_element(By.ID, "ctl00_InputParameters_ps_Period_winp_PeriodDateType")
        dropdown.find_element(By.XPATH, "//option[. = 'Opened']").click()
        logger.info(f"{site_id}: Selected 'Opened' in Period Date Type dropdown")

        driver.find_element(By.ID, "ctl00_InputParameters_ps_Period_winp_PeriodFromDate").clear()
        if site_id == config.pms_credentials['backdate_site_id']:
            fromdate = config.date_str['eo_lastmonth']
        else:
            fromdate = config.date_str['this_month']
        self._enter_date(driver, By.ID, "ctl00_InputParameters_ps_Period_winp_PeriodFromDate", fromdate)
        logger.info(f"{site_id}: Entered 'From' date {fromdate}")

        self._click_element(driver, By.ID, "ctl00_InputParameters_winp_Calculation")
        dropdown = driver.find_element(By.ID, "ctl00_InputParameters_winp_Calculation")
        dropdown.find_element(By.XPATH, "//option[. = 'Accomodation Items Only']").click()
        logger.info(f"{site_id}: Selected 'Accommodation Items Only' in Calculation dropdown")

        self._click_element(driver, By.ID, "ctl00_InputParameters_winp_RevenueType")
        dropdown = driver.find_element(By.ID, "ctl00_InputParameters_winp_RevenueType")
        dropdown.find_element(By.XPATH, "//option[. = 'Bedroom']").click()
        logger.info(f"{site_id}: Selected 'Bedroom' in Revenue Type dropdown")

        self._click_element(driver, By.CSS_SELECTOR, ".popupArrow")
        logger.info(f"{site_id}: Clicked on popup arrow")

        self._click_element(driver, By.ID, "ctl00_ReportsView_wbtn_CSV")
        logger.info(f"{site_id}: Clicked on CSV download button")

        new_file_name = f"{site_id} Actual {config.dates['as_of_date'].strftime('%d.%m.%y')}.csv"
        self._wait_for_file(download_dir, new_file_name)
        logger.info(f"{site_id}: Downloaded Actual report and renamed file to {new_file_name} successfully.")

    @SeleniumExtract.retry(WebDriverException)
    def _generate_forecast_report(self, driver: WebDriver, site_id: str, download_dir: Path):
        """Generate forecast report by downloading 'Revenue By Market Segment' report as csv."""

        logger.info(f"{site_id}: Generating forecast report for site")

        driver.switch_to.default_content()

        self._click_element(driver, By.CSS_SELECTOR, 'button[aria-label="Menu"]')
        logger.info(f"{site_id}: Clicked on dropdown menu")

        self._click_element(driver, By.XPATH, "//span[normalize-space()='Forecast']")
        logger.info(f"{site_id}: Clicked on forecast menu")

        self._click_element(driver, By.XPATH, "//span[contains(@class, 'oc-nav-MuiTypography-bodySmall') and normalize-space()='Revenue By Market Segment']")
        logger.info(f"{site_id}: Selected forecast option")

        WebDriverWait(driver, self.waittime).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, "progress-indicator-background"))
        )

        WebDriverWait(driver, self.waittime).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "MainFrame"))
        )
        logger.info(f"{site_id}: Switched to report frame")

        self._enter_date(driver, By.ID, "ctl00_InputParameters_winp_From", config.date_str['today_date'])
        self._enter_date(driver, By.ID, "ctl00_InputParameters_winp_To", config.date_str['next_12_months'])
        logger.info(f"{site_id}: Entered From and To date")

        booking_type_dropdown = WebDriverWait(driver, self.waittime).until(
            EC.element_to_be_clickable((By.ID, "ctl00_InputParameters_winp_BookingType"))
        )
        booking_type_dropdown.click()
        self._click_element(driver, By.XPATH, "//option[. = 'Bedroom']")
        logger.info(f"{site_id}: Selected 'Bedroom' in Booking Type dropdown")

        self._click_element(driver, By.CSS_SELECTOR, ".inlineblock:nth-child(2)")
        logger.info(f"{site_id}: Selected another option")

        show_as_dropdown = WebDriverWait(driver, self.waittime).until(
            EC.element_to_be_clickable((By.ID, "ctl00_InputParameters_winp_ShowAs"))
        )
        show_as_dropdown.click()
        self._click_element(driver, By.XPATH, "//option[. = 'Nett']")
        logger.info(f"{site_id}: Selected 'Nett' in Show As dropdown")

        display_by_dropdown = WebDriverWait(driver, self.waittime).until(
            EC.element_to_be_clickable((By.ID, "ctl00_InputParameters_winp_DisplayBy"))
        )
        display_by_dropdown.click()
        self._click_element(driver, By.XPATH, "//option[. = 'Date By Market Segment']")
        logger.info(f"{site_id}: Selected 'Date By Market Segment' in Display By dropdown")

        time.sleep(5)

        self._click_element(driver, By.CSS_SELECTOR, ".popupArrow")
        logger.info(f"{site_id}: Clicked on popup arrow")

        self._click_element(driver, By.ID, "ctl00_ReportsView_wbtn_CSV")
        logger.info(f"{site_id}: Clicked on CSV download button")

        new_file_name = f"{site_id} Forecast {config.dates['as_of_date'].strftime('%d.%m.%y')}.csv"
        self._wait_for_file(download_dir, new_file_name)
        logger.info(f"{site_id}: Downloaded Forecast report and renamed file to {new_file_name} successfully.")

    @SeleniumExtract.retry(WebDriverException)
    def _generate_actual_and_forecast_report(self, driver: WebDriver, site_id: str, download_dir: Path):
        """Generate actual and forecast report by downloading Actual&Forecast as pdf (3 files)."""
        logger.info(f"{site_id}: Generating actual and forecast report for site")

        driver.switch_to.default_content()

        self._click_element(driver, By.CSS_SELECTOR, 'button[aria-label="Menu"]')
        logger.info(f"{site_id}: Clicked on dropdown menu")

        self._click_element(driver, By.XPATH, "//span[text()='Revenue']")
        logger.info(f"{site_id}: Clicked on Revenue button")

        self._click_element(driver, By.XPATH, "//span[contains(@class, 'oc-nav-MuiTypography-bodySmall') and normalize-space()='Actual & Forecast']")
        logger.info(f"{site_id}: Selected actual and forecast option")

        WebDriverWait(driver, self.waittime).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, "progress-indicator-background"))
        )

        WebDriverWait(driver, self.waittime).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "MainFrame"))
        )
        logger.info(f"{site_id}: Switched to report frame")

        WebDriverWait(driver, self.waittime).until(
            EC.presence_of_element_located((By.ID, "ctl00_InputParameters_winp_StartDate"))
        )
        logger.info(f"{site_id}: Report parameters form is ready for interaction.")

        def enter_dates_and_select_options(from_date, to_date, report_name):
            try:
                self._enter_date(driver, By.ID, "ctl00_InputParameters_winp_StartDate", from_date)
                logger.info(f"{site_id}: Entered start date {from_date}")

                self._enter_date(driver, By.ID, "ctl00_InputParameters_winp_EndDate", to_date)
                logger.info(f"{site_id}: Entered end date {to_date}")

                show_as_dropdown = WebDriverWait(driver, self.waittime).until(
                    EC.element_to_be_clickable((By.ID, "ctl00_InputParameters_winp_ShowAs"))
                )
                show_as_dropdown.click()
                self._click_element(driver, By.XPATH, "//option[. = 'Nett']")
                logger.info(f"{site_id}: Selected 'Nett' in ShowAs dropdown")

                revenue_dropdown = WebDriverWait(driver, self.waittime).until(
                    EC.element_to_be_clickable((By.ID, "ctl00_InputParameters_winp_Revenue"))
                )
                revenue_dropdown.click()
                self._click_element(driver, By.XPATH, "//option[. = 'Bedroom']")
                logger.info(f"{site_id}: Selected 'Bedroom' in Revenue dropdown")

                self._click_element(driver, By.ID, "ctl00_ReportsView_wbtn_View")
                self._click_element(driver, By.ID, "ctl00_ReportsView_wbtn_PDF")
                logger.info(f"{site_id}: Clicked on View and PDF buttons to download report")

                self._wait_for_file(download_dir)

            except WebDriverException as e:
                logger.error(f"{site_id}: Error while generating report {report_name}: {e}")
                raise

        enter_dates_and_select_options(config.date_str['this_month'], config.date_str['next_month'], config.date_str['this_month_name'])
        enter_dates_and_select_options(config.date_str['next_month'], config.date_str['next_2_months'], config.date_str['next_month_name'])
        enter_dates_and_select_options(config.date_str['next_2_months'], config.date_str['next_3_months'], config.date_str['next_2_months_name'])

        logger.info(f"{site_id}: Downloaded Actual&Forecast succesfully.")

    @SeleniumExtract.retry(WebDriverException)
    def _generate_daily_new_reservations(self, driver: WebDriver, site_id: str, download_dir: Path):
        """Generate Daily New Reservations by downloading 'Daily New Reservation' report as pdf."""
        logger.info(f"{site_id}: Generating daily new reservations report for site.")

        driver.switch_to.default_content()

        self._click_element(driver, By.CSS_SELECTOR, 'button[aria-label="Menu"]')
        logger.info(f"{site_id}: Clicked on dropdown menu")

        self._click_element(driver, By.XPATH, "//span[normalize-space()='Reservation']")
        logger.info(f"{site_id}: Clicked on Reservation button")

        self._click_element(driver,By.XPATH, "//span[contains(@class, 'oc-nav-MuiTypography-bodySmall') and normalize-space()='Daily New Reservations']")
        logger.info(f"{site_id}: Selected daily new reservations report option")

        WebDriverWait(driver, self.waittime).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, "progress-indicator-background"))
        )

        WebDriverWait(driver, self.waittime).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "MainFrame"))
        )
        logger.info(f"{site_id}: Switched to report frame")

        today_date = config.dates['today_date']
        backdate_str = (today_date - timedelta(days=self.backdate)).strftime('%d%m%Y')

        self._enter_date(driver, By.ID, "ctl00_InputParameters_winp_EndDate", config.date_str['today_date'])
        self._enter_date(driver, By.ID, "ctl00_InputParameters_winp_StartDate", backdate_str)

        booking_type_dropdown = WebDriverWait(driver, self.waittime).until(
            EC.element_to_be_clickable((By.ID, "ctl00_InputParameters_winp_BookingType"))
        )
        booking_type_dropdown.click()
        self._click_element(driver, By.XPATH, "//option[. = 'Bedroom']")
        logger.info(f"{site_id}: Selected 'Bedroom' in Booking Type dropdown")

        self._click_element(driver, By.CSS_SELECTOR, ".popupArrow")
        self._click_element(driver, By.ID, "ctl00_ReportsView_wbtn_PDF")
        logger.info(f"{site_id}: Clicked on PDF download button")

        self._wait_for_file(download_dir)
        logger.info(f"{site_id}: Downloaded Daily New Reservation successfully.")

    @SeleniumExtract.retry(WebDriverException)
    def _generate_cancellation_report(self, driver: WebDriver, site_id: str, download_dir: Path):
        """Generate cancellation report by downloading cancellation report as pdf."""
        logger.info(f"{site_id} Generating Cancellation report for site.")
        driver.switch_to.default_content()

        self._click_element(driver, By.CSS_SELECTOR, 'button[aria-label="Menu"]')
        logger.info(f"{site_id}: Clicked on dropdown menu")

        self._click_element(driver, By.XPATH, "//span[normalize-space()='Reservation']")
        logger.info(f"{site_id}: Clicked on Reservation button")

        self._click_element(driver,By.XPATH, "//span[contains(@class, 'oc-nav-MuiTypography-bodySmall') and normalize-space()='Cancellations']")
        logger.info(f"{site_id}: Selected cancellation report option")

        WebDriverWait(driver, self.waittime).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, "progress-indicator-background"))
        )

        WebDriverWait(driver, self.waittime).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "MainFrame"))
        )
        logger.info(f"{site_id}: Switched to report frame")

        today_date = config.dates['today_date']
        backdate_str = (today_date - timedelta(days=self.backdate)).strftime('%d%m%Y')

        self._enter_date(driver, By.ID, "ctl00_InputParameters_winp_StartDate", backdate_str)
        self._enter_date(driver, By.ID, "ctl00_InputParameters_winp_EndDate", config.date_str['as_of_date'])
        logger.info(f"{site_id}: Entered start and end dates")

        self._click_element(driver,By.ID, "ctl00_InputParameters_rad_CancelDate")
        logger.info(f"{site_id}: Selected cancel date radio button")

        self._click_element(driver, By.CSS_SELECTOR, ".popupArrow")
        self._click_element(driver, By.ID, "ctl00_ReportsView_wbtn_PDF")
        logger.info(f"{site_id}: Clicked on PDF download button")

        self._wait_for_file(download_dir)
        logger.info(f"{site_id}: Downloaded Cancellation report successfully.")

    @SeleniumExtract.retry(WebDriverException)
    def _generate_eod_report(self, driver: WebDriver, site_id: str, download_dir: Path):
        """Generate EOD report by downloading 'Period End' as csv and pdf."""
        logger.info(f"{site_id}: Generating EOD report for site.")

        time.sleep(3)
        driver.switch_to.default_content()

        self._click_element(driver, By.CSS_SELECTOR, 'button[aria-label="Menu"]')
        logger.info(f"{site_id}: Clicked on menu button")

        self._click_element(driver, By.XPATH, "//span[normalize-space()='Period End']")
        logger.info(f"{site_id}: Clicked on Period End button")

        self._click_element(driver, By.XPATH, "//span[contains(@class, 'oc-nav-MuiTypography-bodySmall') and normalize-space()='Period End Archive']")
        logger.info(f"{site_id}: Clicked on period end button using XPath")

        WebDriverWait(driver, self.waittime).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, "progress-indicator-background"))
        )

        time.sleep(3)
        WebDriverWait(driver, self.waittime).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "MainFrame")))
        logger.info(f"{site_id}: Switched to report frame")
        WebDriverWait(driver, self.waittime).until(EC.visibility_of_element_located((By.ID, "winp_PeriodType")))

        period_dropdown = WebDriverWait(driver, self.waittime).until(EC.element_to_be_clickable((By.ID, "winp_PeriodType")))
        period_dropdown.click()
        period_dropdown.find_element(By.XPATH, "//option[. = 'End of day']").click()
        logger.info(f"{site_id}: Selected 'End of day' in Period Type dropdown")

        self._click_element(driver, By.ID, "wbtn_Refresh")
        logger.info(f"{site_id}: Clicked on Refresh button")
        time.sleep(3)

        WebDriverWait(driver, self.waittime).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, "progress-indicator-background"))
        )

        self._click_element(driver, By.ID, "grd_Periods_ctl02_wbtn_DetailedExtract")
        logger.info(f"{site_id}: Clicked on Detailed Extract button to download CSV")

        new_file_name = f"{site_id} EOD {config.dates['as_of_date'].strftime('%d.%m.%y')}.csv"
        self._wait_for_file(download_dir, new_file_name)
        logger.info(f"{site_id}: Downloaded EOD.csv and renamed file to: {new_file_name}")

        self._click_element(driver, By.CSS_SELECTOR, "#grd_Periods_ctl02_rpt_View_wbtn_View > .popupArrow")
        self._click_element(driver, By.ID, "grd_Periods_ctl02_rpt_View_wbtn_PDF")
        logger.info(f"{site_id}: Clicked on View and PDF buttons to download PDF")

        self._wait_for_file(download_dir)
        logger.info(f"{site_id}: Downloaded EOD.pdf successfully.")

        if self.backdate < 1:
            driver.switch_to.default_content()
        else:
            ctlinitialvalue = 2  # ctl02
            for i in range(1, self.backdate):
                logger.info(f"Processing back date {i}")
                ctlvalue = ctlinitialvalue + i
                back_date_value = config.dates['as_of'] - timedelta(days=i)

                self._click_element(driver, By.ID, "wbtn_Refresh")
                logger.info(f"{site_id}: Clicked on Refresh button")
                time.sleep(2)

                WebDriverWait(driver, self.waittime).until(
                    EC.invisibility_of_element_located((By.CLASS_NAME, "progress-indicator-background"))
                )

                self._click_element(driver, By.ID, f"grd_Periods_ctl0{ctlvalue}_wbtn_DetailedExtract")
                logger.info(f"{site_id}: Clicked on Detailed Extract button for back date {i} to download CSV")
                new_file_name = f"{site_id}: EOD {back_date_value.strftime('%d.%m.%y')}.csv"

                self._wait_for_file(download_dir, new_file_name)
                logger.info(f"{site_id}: Downloaded and renamed file to: {new_file_name}")

                self._click_element(driver, By.CSS_SELECTOR, f"#grd_Periods_ctl0{ctlvalue}_rpt_View_wbtn_View > .popupArrow")
                self._click_element(driver, By.ID, f"grd_Periods_ctl0{ctlvalue}_rpt_View_wbtn_PDF")
                logger.info(f"{site_id}: Clicked on View and PDF buttons for back date {i} to download PDF")

                self._wait_for_file(download_dir)
                logger.info(f"{site_id}: Downloaded EOD {back_date_value} pdf")

        logger.info(f"{site_id}: EOD report generated successfully for site.")

    @SeleniumExtract.retry(WebDriverException)
    def _generate_avai_report(self, driver: WebDriver, site_id: str, download_dir: Path):
        """Generate Availability report by downloading 'Availability' as csv."""
        logger.info(f"{site_id}: Generating availability report for site")

        driver.switch_to.default_content()

        self._click_element(driver, By.CSS_SELECTOR, 'button[aria-label="Menu"]')
        logger.info(f"{site_id}: Clicked on dropdown menu")

        self._click_element(driver, By.XPATH, "//span[normalize-space()='Reservation']")
        logger.info(f"{site_id}: Clicked on Reservation button")

        self._click_element(driver, By.XPATH, "//a[@href='RLXWebForms/rep/wfrm_rep_Availability.aspx']")
        logger.info(f"{site_id}: Selected 'AV' option")

        WebDriverWait(driver, self.waittime).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, "progress-indicator-background"))
        )

        WebDriverWait(driver, self.waittime).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "MainFrame"))
        )
        logger.info(f"{site_id}: Switched to report frame")

        self._enter_date(driver, By.ID, "ctl00_InputParameters_winp_StartDate", config.date_str['first_day_of_year'])
        self._enter_date(driver, By.ID, "ctl00_InputParameters_winp_EndDate", config.date_str['first_day_of_next_year'])
        logger.info(f"{site_id}: Entered start and end dates")

        self._click_element(driver, By.CSS_SELECTOR, "span.popupArrow")
        self._click_element(driver, By.ID, "ctl00_ReportsView_wbtn_CSV")
        logger.info(f"{site_id}: Clicked on export button to download CSV")

        new_file_name = f"{site_id} Availability {config.dates['as_of_date'].strftime('%Y')}.csv"
        self._wait_for_file(download_dir, new_file_name)
        logger.info(f"{site_id}: Downloaded Availability and renamed file to: {new_file_name}")

    @SeleniumExtract.retry(WebDriverException)
    def _generate_company_detail_report(self, driver: WebDriver, site_id: str, download_dir: Path):
        """Generate Company Detail report by downloading 'Company' report as pdf."""
        logger.info(f"{site_id}: Generating company detail report for site")

        driver.switch_to.default_content()

        self._click_element(driver, By.CSS_SELECTOR, 'button[aria-label="Menu"]')
        logger.info(f"{site_id}: Clicked on dropdown menu")

        self._click_element(driver, By.XPATH, "//span[text()='Revenue']")
        logger.info(f"{site_id}: Clicked on Revenue button")

        self._click_element(driver, By.XPATH, "//a[@href='RLXWebforms/rep/wfrm_rep_Company.aspx']")
        logger.info(f"{site_id}: Selected company report option")

        WebDriverWait(driver, self.waittime).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, "progress-indicator-background"))
        )

        WebDriverWait(driver, self.waittime).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "MainFrame"))
        )
        logger.info(f"{site_id}: Switched to report frame")

        self._enter_date(driver, By.ID, "ctl00_InputParameters_winp_From", config.date_str['first_day_of_year'])
        self._enter_date(driver, By.ID, "ctl00_InputParameters_winp_To", config.date_str['last_day_of_year'])
        logger.info(f"{site_id}: Entered From and To dates")

        def select_dropdown_option(dropdown_id, option_text):
            self._click_element(driver, By.ID, dropdown_id)
            self._click_element(driver, By.XPATH, f"//*[@id='{dropdown_id}']/option[. = '{option_text}']")
            logger.info(f"{site_id}: Selected '{option_text}' in dropdown with ID '{dropdown_id}'")

        select_dropdown_option("ctl00_InputParameters_winp_BookingType", "Bedroom")
        select_dropdown_option("ctl00_InputParameters_winp_ShowAs", "Nett")
        select_dropdown_option("ctl00_InputParameters_winp_Type", "Forecast")

        self._click_element(driver, By.ID, "ctl00_InputParameters_winp_ShowDetailed")
        logger.info(f"{site_id}: Included show details")

        dropdown_button_xpath = "//div[@id='ctl00_InputParameters_msd_Include_wdd_Property']//td[@mkr='Button']"
        self._click_element(driver, By.XPATH, dropdown_button_xpath)
        self._click_element(driver, By.LINK_TEXT, "Room Nights")
        self._click_element(driver, By.LINK_TEXT, "Accom Revenue")
        self._click_element(driver, By.LINK_TEXT, "ARR")
        logger.info(f"{site_id}: Selected additional details: Room Nights, Accom Revenue, ARR")

        self._click_element(driver, By.ID, "ctl00_ReportsView_wbtn_View")
        time.sleep(1)
        self._click_element(driver, By.ID, "ctl00_ReportsView_wbtn_PDF")
        logger.info(f"{site_id}: Clicked on View and PDF buttons to download report")

        self._wait_for_file(download_dir)
        logger.info(f"{site_id}: Downloaded Company Detail reprot successfully.")

    # --- OS ---
    @staticmethod
    def _check_site_files(download_dir: Path, site_id: str):
        """Raises if any PDF is named for another site; PortalPms names PDFs '<SITE>_<Report>_<time>.pdf'."""
        wrong = [
            f.name for f in download_dir.iterdir()
            if f.is_file() and f.suffix.lower() == '.pdf' and f.name.split('_', 1)[0] != site_id
        ]
        if wrong:
            raise SiteMismatchError(f"{site_id}: downloaded files belong to another site, upload stopped: {wrong}")

    def _move_files(self, download_dir: Path, country: str, site_id: str):
        """Upload downloaded files to SharePoint via Graph API."""
        self._check_site_files(download_dir, site_id)
        logger.info(f"{site_id}: Starting move files to SharePoint.")
        sp_report_data = f"{config.credentials['sp_report_root']}/Report Data"
        upload_sp = self._sp_upload()

        file_types = {
            f"{site_id} Actual": "Actual",
            f"{site_id} Forecast": "Forecast",
            f"{site_id} Availability": "Availability",
            f"{site_id} EOD": "EOD"
        }

        ## --- CSVs ---
        for source_file in list(download_dir.iterdir()):
            if source_file.is_file() and source_file.suffix.lower() == '.csv':
                sp_dest = None
                for key, folder_suffix in file_types.items():
                    if key in source_file.name:
                        sp_dest = f"{sp_report_data}/{country}_RAW/{site_id}/{site_id} {folder_suffix}"
                        break
                if sp_dest:
                    upload_sp.upload_file(str(source_file), sp_dest)
                    source_file.unlink()
                    logger.info(f"{site_id}: Uploaded {source_file.name} -> {sp_dest}")
                else:
                    logger.info(f"{site_id}: Skipped '{source_file.name}' (no matching report type)")

        ## --- PDFs ---
        remaining = [f for f in download_dir.iterdir() if f.is_file()]
        if remaining:
            month_number = config.date_str['month_number']
            month_name   = config.date_str['month_name']
            year         = config.date_str['year']
            month_year_folder = f"{month_number} {month_name} {year} ({site_id})"
            sp_pdf_dest = (
                f"{config.credentials['sp_report_root']}/Supporting Docs/{country}/{site_id}"
                f"/{year}/{month_year_folder}/{download_dir.name}"
            )
            for item in remaining:
                upload_sp.upload_file(str(item), sp_pdf_dest)
                item.unlink()
                logger.info(f"{site_id}: Uploaded {item.name} -> {sp_pdf_dest}")

        if download_dir.exists() and not any(download_dir.iterdir()):
            download_dir.rmdir()
            logger.info(f"{site_id}: Local temp folder cleaned up.")

    def _skip_dates_copy(self, site_id: str, country: str):
        """Copy the latest dated CSV to fill missing backdate slots, directly on SharePoint."""
        logger.info(f"Checking if backdate > 1 for {site_id}, value: {self.backdate}")
        if self.backdate <= 1:
            logger.info(f"Skipping skip_dates_copy for {site_id}")
            return

        logger.info(f"Calling skip_dates_copy for {site_id}")
        sp_root = config.credentials['sp_report_root']
        folder_sp_paths = [
            f"{sp_root}/Report Data/{country}_RAW/{site_id}/{site_id} Actual",
            f"{sp_root}/Report Data/{country}_RAW/{site_id}/{site_id} Forecast",
        ]

        def process_folder(sp_folder_path: str):
            read_sp_client = ReadSP(self._token_manager.get_token(), self._sp_site_id)
            csv_files = read_sp_client.list_files_in_folder(sp_folder_path)
            if not csv_files:
                logger.warning(f"No files found in: {sp_folder_path}")
                return

            latest_date, latest_file = None, None
            for file_info in csv_files:
                try:
                    date_str  = file_info['name'].split(" ")[-1].replace(".csv", "")
                    file_date = datetime.strptime(date_str, "%d.%m.%y")
                    if latest_date is None or file_date > latest_date:
                        latest_date, latest_file = file_date, file_info
                except ValueError:
                    logger.warning(f"Skipping file {file_info['name']}: date format invalid.")
                    continue

            if not latest_file:
                logger.warning(f"No valid dated files in: {sp_folder_path}")
                return

            logger.info(f"Latest file: {latest_file['name']} ({latest_date.strftime('%d.%m.%y')})")
            content     = read_sp_client.get_file_content_by_id(latest_file['id'])
            file_prefix = " ".join(latest_file['name'].split(" ")[:-1])
            upload_sp   = self._sp_upload()

            for i in range(1, self.backdate):
                new_name = f"{file_prefix} {(latest_date + timedelta(days=i)).strftime('%d.%m.%y')}.csv"
                fd, tmp  = tempfile.mkstemp(suffix='.csv')
                os.close(fd)
                tmp_path = Path(tmp)
                try:
                    tmp_path.write_text(content, encoding='utf-8')
                    upload_sp.upload_file(str(tmp_path), sp_folder_path, new_name)
                    logger.info(f"Copied to SP: {new_name}")
                finally:
                    tmp_path.unlink(missing_ok=True)

        for sp_path in folder_sp_paths:
            process_folder(sp_path)

    # --- Main Processor Function ---
    def run(self, country: str, site_id: str) -> tuple[str,str]:
        """Automates the login and sequential downloading of multiple reports for a specific site."""
        report_date_value = config.dates['as_of_date'].strftime('%d.%m.%y')
        folder_name = f"{site_id} {report_date_value}"
        download_directory = config.paths['download_dir'] / folder_name
        profile_dir = self.profiles_dir / 'shared'
        driver = None

        try:
            driver = self.init_driver(download_directory, profile_dir)
            self._login(driver, site_id)
            self._generate_actual_report(driver, site_id, download_directory)
            self._generate_forecast_report(driver, site_id, download_directory)
            self._generate_avai_report(driver, site_id, download_directory)
            self._generate_eod_report(driver, site_id, download_directory)
            if country == "UK":
                self._generate_actual_and_forecast_report(driver, site_id, download_directory)
                self._generate_company_detail_report(driver, site_id, download_directory)
                self._generate_daily_new_reservations(driver, site_id, download_directory)
                self._generate_cancellation_report(driver, site_id, download_directory)
            self._skip_dates_copy(site_id, country)
            self._move_files(download_directory, country, site_id)
            logger.info(f"{site_id}: All report data retrived successfully.")

            return (site_id, "EXTRACT DATA SUCCESS")

        except Exception as e:
            # Not reported as FAIL, so the pipeline does not retry and wipe the downloaded files
            if isinstance(e, requests.exceptions.HTTPError) and e.response is not None and e.response.status_code == 507:
                logger.error(f"{site_id}: {STORAGE_FULL} - files kept in {download_directory}")
                print(f"{site_id}: {STORAGE_FULL} - files kept in {download_directory}")
                return (site_id, f"EXTRACT DATA SKIPPED - {STORAGE_FULL} (files kept in {download_directory})")
            logger.error(f"{site_id}: FAILED - {e}")
            print(f"{site_id}: FAILED - {e}")
            return (site_id, f"EXTRACT DATA FAILED - {e}")

        finally:
            if driver:
                driver.quit()
