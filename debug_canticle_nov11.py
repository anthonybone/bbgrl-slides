from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import time
from datetime import datetime
import re

# Navigate to November 11, 2025
chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 10)

try:
    driver.get("https://www.ibreviary.com/m2/")
    time.sleep(3)
    
    # Navigate to date
    more_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "More")))
    more_link.click()
    time.sleep(2)
    
    day_input = driver.find_element(By.NAME, "giorno")
    day_input.clear()
    day_input.send_keys("11")
    
    month_select = driver.find_element(By.NAME, "mese")
    month_select.send_keys("11")
    
    year_input = driver.find_element(By.NAME, "anno")
    year_input.clear()
    year_input.send_keys("2025")
    
    ok_button = driver.find_element(By.CSS_SELECTOR, "input[type='button'][value='OK']")
    ok_button.click()
    time.sleep(2)
    
    breviary_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Breviary")))
    breviary_link.click()
    time.sleep(2)
    
    morning_prayer_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Morning Prayer")))
    morning_prayer_link.click()
    time.sleep(3)
    
    html_content = driver.page_source
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Find all rubrica spans
    print("All rubrica spans containing 'Canticle':")
    for span in soup.find_all('span', class_='rubrica'):
        span_text = span.get_text().strip()
        if 'Canticle' in span_text:
            print(f"  - '{span_text}'")
            # Get surrounding context
            parent_text = span.parent.get_text()[:200] if span.parent else ""
            print(f"    Context: {parent_text[:100]}...")
    
finally:
    driver.quit()
