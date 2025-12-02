from __future__ import annotations
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import WebDriverException, TimeoutException
import time
import requests
from typing import Optional


class IBreviaryScraper:
    """Thin Selenium wrapper for iBreviary navigation.

    This class owns the WebDriver lifecycle. Navigation routines can be
    added here over time to fully extract scraping concerns out of the
    main generator.
    """

    def __init__(self, base_url: str):
        self.base_url = base_url
        self.driver: webdriver.Chrome | None = None

    def init_driver(self, force_reinit: bool = False) -> webdriver.Chrome:
        """Initialize (or reinitialize) a headless Chrome WebDriver and return it.

        Adds stability flags and retries for environments where the driver
        may crash intermittently (observed stacktrace with Chrome headless).
        """
        if self.driver is not None and not force_reinit:
            return self.driver

        if self.driver is not None and force_reinit:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None

        chrome_options = Options()
        # Use the new headless mode explicitly; some versions need --headless=new
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-features=VizDisplayCompositor")
        chrome_options.add_argument("--log-level=3")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])  # reduce noise
        # Performance: block images/fonts to reduce load variation
        prefs = {"profile.managed_default_content_settings.images": 2,
                 "profile.managed_default_content_settings.stylesheets": 1,
                 "profile.managed_default_content_settings.cookies": 1}
        chrome_options.add_experimental_option("prefs", prefs)

        try:
            self.driver = webdriver.Chrome(options=chrome_options)
        except WebDriverException as e:
            raise RuntimeError(f"Failed to initialize ChromeDriver: {e}")
        return self.driver

    def quit(self) -> None:
        """Dispose the driver if present."""
        if self.driver:
            try:
                self.driver.quit()
            finally:
                self.driver = None

    def navigate_morning_prayer_html(self, target_date) -> Optional[str]:
        """Navigate to Morning Prayer for a given date and return page HTML.

        Strategy:
        1. Attempt Selenium navigation with up to 2 retries (driver re-init on crash).
        2. If Selenium repeatedly fails, attempt a lightweight requests fallback
           (will return current day content if date switching unsupported without UI).
        """
        attempt = 0
        last_error: Optional[str] = None
        while attempt < 2:
            try:
                driver = self.init_driver(force_reinit=(attempt > 0))
                wait = WebDriverWait(driver, 15)
                print(f"  -> [Attempt {attempt+1}] Navigating to iBreviary mobile site...")
                driver.get(self.base_url)
                # Try to dismiss any cookie/consent banners if present
                self._attempt_consent_dismiss(driver)

                # Open 'More' to set date
                more_link = self._robust_find_any(wait, [
                    (By.LINK_TEXT, "More"),
                    (By.PARTIAL_LINK_TEXT, "More"),
                    (By.XPATH, "//a[contains(@href, 'opzioni.php')]")
                ], description="'More' menu")
                if not more_link:
                    raise RuntimeError("Could not locate 'More' navigation link")
                more_link.click()

                # Date inputs
                print(f"  -> Setting date to {target_date.strftime('%d/%m/%Y')}...")
                day_field = wait.until(EC.presence_of_element_located((By.NAME, "giorno")))
                month_dropdown_el = wait.until(EC.presence_of_element_located((By.NAME, "mese")))
                year_field = wait.until(EC.presence_of_element_located((By.NAME, "anno")))
                day_field.clear(); day_field.send_keys(str(target_date.day))
                Select(month_dropdown_el).select_by_index(target_date.month - 1)
                year_field.clear(); year_field.send_keys(str(target_date.year))
                ok_button = wait.until(EC.element_to_be_clickable((By.NAME, "ok")))
                ok_button.click()

                # Breviary link
                breviary_link = self._robust_find_any(wait, [
                    (By.LINK_TEXT, "Breviary"),
                    (By.PARTIAL_LINK_TEXT, "Breviary"),
                    (By.XPATH, "//a[contains(@href,'breviario.php')]")
                ], description="'Breviary' link")
                if not breviary_link:
                    raise RuntimeError("Could not locate 'Breviary' link after date set")
                breviary_link.click()

                # Morning Prayer link variants
                morning_prayer_link = self._robust_find_any(wait, [
                    (By.PARTIAL_LINK_TEXT, "Morning Prayer"),
                    (By.PARTIAL_LINK_TEXT, "Lauds"),
                    (By.PARTIAL_LINK_TEXT, "Lodi"),
                ], description="Morning Prayer/Lauds/Lodi link")
                if not morning_prayer_link:
                    raise RuntimeError("Could not locate Morning Prayer link")
                morning_prayer_link.click()

                html = driver.page_source
                print(f"  Successfully navigated to Morning Prayer for {target_date.strftime('%B %d, %Y')}")
                return html
            except Exception as e:
                last_error = str(e)
                print(f"  Error during Selenium navigation attempt {attempt+1}: {e}")
                attempt += 1
        print(f"  WARNING: Selenium navigation failed after retries: {last_error}")
        # Fallback: try direct request (may not reflect requested past date)
        try:
            fallback_url = f"{self.base_url}breviario.php?s=lodi"
            resp = requests.get(fallback_url, timeout=15, headers={"User-Agent": "Mozilla/5.0"})
            if resp.status_code == 200 and len(resp.text) > 5000:
                print("  ✓ Using requests fallback for Morning Prayer (date control may be inaccurate).")
                return resp.text
            else:
                print(f"  WARNING: Fallback HTTP fetch unsuccessful (status {resp.status_code}, length {len(resp.text)})")
        except Exception as e:
            print(f"  WARNING: Fallback requests fetch failed: {e}")
        return None

    def navigate_readings_html(self) -> Optional[str]:
        """Navigate from current context to Readings page and return page HTML.

        Includes retry and direct HTTP fallback similar to Morning Prayer.
        """
        attempt = 0
        last_error: Optional[str] = None
        while attempt < 2:
            try:
                driver = self.init_driver(force_reinit=(attempt > 0))
                wait = WebDriverWait(driver, 15)
                print(f"  -> [Attempt {attempt+1}] Navigating to readings (tab + link)...")
                reading_tab = self._robust_find_any(wait, [
                    (By.LINK_TEXT, "Reading"),
                    (By.PARTIAL_LINK_TEXT, "Reading"),
                    (By.XPATH, "//a[contains(@href, 'letture.php')]")
                ], description="Reading tab")
                if not reading_tab:
                    raise RuntimeError("Could not locate Reading tab")
                reading_tab.click()
                readings_link = self._robust_find_any(wait, [
                    (By.PARTIAL_LINK_TEXT, "Readings"),
                    (By.PARTIAL_LINK_TEXT, "Letture"),
                    (By.PARTIAL_LINK_TEXT, "readings"),
                ], description="Readings link")
                if not readings_link:
                    raise RuntimeError("Could not locate Readings link")
                readings_link.click()
                html = driver.page_source
                print("  Successfully navigated to Readings page")
                return html
            except Exception as e:
                last_error = str(e)
                print(f"  Error navigating to Readings attempt {attempt+1}: {e}")
                attempt += 1
        print(f"  WARNING: Selenium readings navigation failed after retries: {last_error}")
        try:
            fallback_url = f"{self.base_url}letture.php?s=letture"
            resp = requests.get(fallback_url, timeout=15, headers={"User-Agent": "Mozilla/5.0"})
            if resp.status_code == 200 and len(resp.text) > 5000:
                print("  ✓ Using requests fallback for Readings page.")
                return resp.text
            else:
                print(f"  WARNING: Fallback Readings HTTP fetch unsuccessful (status {resp.status_code})")
        except Exception as e:
            print(f"  WARNING: Fallback readings requests fetch failed: {e}")
        return None

    # ----------------- helper utilities -----------------

    def _robust_find_any(self, wait: WebDriverWait, locator_variants, description: str):
        """Try a list of locator variants, return the first element found or None."""
        for by, value in locator_variants:
            try:
                elem = wait.until(EC.element_to_be_clickable((by, value)))
                return elem
            except TimeoutException:
                continue
            except Exception:
                continue
        print(f"  WARNING: Locator not found for {description}")
        return None

    def _attempt_consent_dismiss(self, driver: webdriver.Chrome):
        """Attempt to dismiss cookie/consent modals that can block clicks."""
        try:
            # Common button texts in multiple languages
            for txt in ["Accept", "Accetta", "OK", "Chiudi", "Close"]:
                buttons = driver.find_elements(By.XPATH, f"//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{txt.lower()}')]")
                for b in buttons:
                    try:
                        if b.is_displayed():
                            b.click()
                            print("  -> Dismissed consent/cookie banner")
                            return
                    except Exception:
                        continue
        except Exception:
            pass
