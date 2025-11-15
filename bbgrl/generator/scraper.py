from __future__ import annotations
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
import time


class IBreviaryScraper:
    """Thin Selenium wrapper for iBreviary navigation.

    This class owns the WebDriver lifecycle. Navigation routines can be
    added here over time to fully extract scraping concerns out of the
    main generator.
    """

    def __init__(self, base_url: str):
        self.base_url = base_url
        self.driver: webdriver.Chrome | None = None

    def init_driver(self) -> webdriver.Chrome:
        """Initialize a headless Chrome WebDriver and return it."""
        if self.driver is not None:
            return self.driver

        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")

        self.driver = webdriver.Chrome(options=chrome_options)
        return self.driver

    def quit(self) -> None:
        """Dispose the driver if present."""
        if self.driver:
            try:
                self.driver.quit()
            finally:
                self.driver = None

    def navigate_morning_prayer_html(self, target_date) -> str | None:
        """Navigate to Morning Prayer for a given date and return page HTML."""
        driver = self.init_driver()
        wait = WebDriverWait(driver, 10)

        try:
            print(f"  -> Navigating to iBreviary mobile site...")
            driver.get(self.base_url)
            time.sleep(3)

            # Click 'More'
            print(f"  -> Clicking 'More' menu...")
            more_link = None
            try:
                more_link = wait.until(EC.presence_of_element_located((By.LINK_TEXT, "More")))
            except Exception:
                pass
            if not more_link:
                try:
                    more_link = wait.until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "More")))
                except Exception:
                    pass
            if not more_link:
                try:
                    more_link = wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(@href, 'opzioni.php')]")))
                except Exception:
                    pass
            if not more_link:
                print("  WARNING: Could not find 'More' link")
                return None
            driver.execute_script("arguments[0].scrollIntoView(true);", more_link)
            time.sleep(0.5)
            more_link.click()
            time.sleep(2)

            # Set date fields
            print(f"  -> Setting date to {target_date.strftime('%d/%m/%Y')}...")
            day_field = driver.find_element(By.NAME, "giorno")
            day_field.clear()
            day_field.send_keys(str(target_date.day))
            month_dropdown = Select(driver.find_element(By.NAME, "mese"))
            month_dropdown.select_by_index(target_date.month - 1)
            year_field = driver.find_element(By.NAME, "anno")
            year_field.clear()
            year_field.send_keys(str(target_date.year))
            print(f"    Set day: {target_date.day}")
            print(f"    Set month: {target_date.month}")
            print(f"    Set year: {target_date.year}")
            print("    Clicking 'OK' button to apply date...")
            ok_button = driver.find_element(By.NAME, "ok")
            ok_button.click()
            time.sleep(2)

            # Click 'Breviary'
            print("  -> Clicking 'Breviary' link...")
            breviary_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Breviary")))
            breviary_link.click()
            time.sleep(2)

            # Click Morning Prayer
            print("  -> Clicking 'Morning Prayer' link...")
            morning_prayer_link = None
            for text in ["Morning Prayer", "Lauds", "Lodi"]:
                try:
                    morning_prayer_link = wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, text)))
                    break
                except Exception:
                    continue
            if not morning_prayer_link:
                print("  WARNING: Could not find 'Morning Prayer' link")
                return None
            morning_prayer_link.click()
            time.sleep(2)

            html = driver.page_source
            print(f"  Successfully navigated to Morning Prayer for {target_date.strftime('%B %d, %Y')}")
            return html
        except Exception as e:
            print(f"  Error during Selenium navigation: {e}")
            return None

    def navigate_readings_html(self) -> str | None:
        """Navigate from current context to Readings page and return page HTML."""
        driver = self.init_driver()
        wait = WebDriverWait(driver, 10)

        try:
            print("  -> Clicking 'Reading' tab...")
            reading_tab = None
            try:
                reading_tab = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Reading")))
            except Exception:
                pass
            if not reading_tab:
                try:
                    reading_tab = wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Reading")))
                except Exception:
                    pass
            if not reading_tab:
                try:
                    reading_tab = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, 'letture.php')]")))
                except Exception:
                    pass
            if not reading_tab:
                print("  WARNING: Could not find 'Reading' tab")
                return None
            driver.execute_script("arguments[0].scrollIntoView(true);", reading_tab)
            time.sleep(0.5)
            reading_tab.click()
            time.sleep(2)

            print("  -> Clicking 'Readings' link...")
            readings_link = None
            for text in ["Readings", "Letture", "readings"]:
                try:
                    readings_link = wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, text)))
                    break
                except Exception:
                    continue
            if not readings_link:
                print("  WARNING: Could not find 'Readings' link")
                return None
            driver.execute_script("arguments[0].scrollIntoView(true);", readings_link)
            time.sleep(0.5)
            readings_link.click()
            time.sleep(2)

            html = driver.page_source
            print("  Successfully navigated to Readings page")
            return html
        except Exception as e:
            print(f"  Error navigating to Readings: {e}")
            return None
