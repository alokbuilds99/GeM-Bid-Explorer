"""
GeM Bid Scraper (v2 - listing page only)
=========================================
Scrapes bid-card info directly from the search results listing on
https://bidplus.gem.gov.in/all-bids - no need to open "View BID Results".

For each bid card it grabs:
    BID No, Status, Items, Quantity, Department Name And Address,
    Start Date, End Date

Loops through every page of results and saves everything to Excel.

INSTALL (if not already):
    pip install selenium webdriver-manager openpyxl pandas

RUN:
    python gem_scraper_v2.py
"""

import time
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, StaleElementReferenceException
)

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------
SEARCH_TERM = "micom relays"
BASE_URL = "https://bidplus.gem.gov.in/all-bids"
OUTPUT_FILE = "gem_bids_output.xlsx"
HEADLESS = False
WAIT_TIMEOUT = 25
MAX_PAGES = None  # None = scrape all pages, or set an int (e.g. 1) while testing

# Set this to your chromedriver.exe path to skip webdriver-manager entirely.
CHROMEDRIVER_PATH = None  # e.g. r"C:\Users\Alok Mishra\...\chromedriver.exe"


def make_driver():
    opts = Options()
    if HEADLESS:
        opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1400,1000")
    opts.add_argument("--disable-notifications")

    if CHROMEDRIVER_PATH:
        service = Service(CHROMEDRIVER_PATH)
    else:
        try:
            import os
            lock_file = os.path.join(os.path.expanduser("~"), ".wdm", ".wdm-lock-chromedriver-win64")
            if os.path.exists(lock_file):
                os.remove(lock_file)
        except Exception:
            pass
        from webdriver_manager.chrome import ChromeDriverManager
        service = Service(ChromeDriverManager().install())

    driver = webdriver.Chrome(service=service, options=opts)
    driver.implicitly_wait(3)
    return driver


def find_visible(driver, selectors, timeout=WAIT_TIMEOUT):
    """Try a list of (by, selector) tuples. Returns the FIRST element that is
    actually visible+enabled, skipping hidden duplicates."""
    end_time = time.time() + timeout
    last_err = None
    while time.time() < end_time:
        for by, selector in selectors:
            try:
                elements = driver.find_elements(by, selector)
                for el in elements:
                    try:
                        if el.is_displayed() and el.is_enabled():
                            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                            return el
                    except StaleElementReferenceException:
                        continue
            except Exception as e:
                last_err = e
        time.sleep(0.5)
    raise TimeoutException(f"No visible/enabled element found for selectors: {selectors}. Last error: {last_err}")


def apply_search_and_filters(driver):
    driver.get(BASE_URL)
    time.sleep(2)

    search_box = find_visible(driver, [
        (By.ID, "searchBid"),
        (By.CSS_SELECTOR, "input[placeholder*='Bid' i]"),
        (By.CSS_SELECTOR, "#bid-listing input.form-control"),
        (By.CSS_SELECTOR, "input.form-control"),
    ])
    search_box.click()
    search_box.clear()
    search_box.send_keys(SEARCH_TERM)

    try:
        search_btn = find_visible(driver, [
            (By.ID, "searchBidRA"),
            (By.CSS_SELECTOR, "span.search-icon"),
            (By.CSS_SELECTOR, "button.btn-search"),
            (By.XPATH, "//button[contains(@class,'search')]"),
        ], timeout=5)
        search_btn.click()
    except TimeoutException:
        search_box.submit()

    time.sleep(2)

    bid_ra_status_checkbox = find_visible(driver, [
        (By.XPATH, "//label[contains(., 'Bid/RA Status')]/preceding-sibling::input"),
        (By.XPATH, "//label[contains(., 'Bid/RA Status')]/parent::*//input"),
        (By.ID, "bidRaStatus"),
    ])
    if not bid_ra_status_checkbox.is_selected():
        bid_ra_status_checkbox.click()
    time.sleep(1)

    awarded_checkbox = find_visible(driver, [
        (By.XPATH, "//label[contains(., 'Bid /RA Awarded') or contains(., 'Bid/RA Awarded')]/preceding-sibling::input"),
        (By.XPATH, "//label[contains(., 'Bid /RA Awarded') or contains(., 'Bid/RA Awarded')]/parent::*//input"),
        (By.ID, "bidAwarded"),
    ])
    if not awarded_checkbox.is_selected():
        awarded_checkbox.click()

    time.sleep(2)


def scrape_listing_page(driver):
    """Parse every bid card visible on the CURRENT results page directly
    from its rendered text - no need to open View BID Results at all.
    Splits the page's full text on 'BID NO:' markers (one per card) and
    regex-extracts the fields out of each chunk."""
    body_text = driver.find_element(By.TAG_NAME, "body").text

    # Split into one chunk per card, each starting at a line beginning with "BID NO:"
    chunks = re.split(r"\n(?=BID NO:)", body_text)

    records = []
    for chunk in chunks:
        if "BID NO:" not in chunk:
            continue

        def grab(pattern, flags=0):
            m = re.search(pattern, chunk, flags)
            return m.group(1).strip() if m else ""

        bid_no = grab(r"BID NO:\s*([A-Za-z0-9/]+)")
        status = grab(r"Status:\s*([^\n]+)")
        items = grab(r"Items:\s*([^\n]+)")
        quantity = grab(r"Quantity:\s*([0-9]+)")
        department = grab(r"Department Name And Address:\s*(.+?)(?=\nStart Date:|\nEnd Date:|$)", re.S)
        start_date = grab(r"Start Date:\s*([^\n]+)")
        end_date = grab(r"End Date:\s*([^\n]+)")

        records.append({
            "BID No": bid_no,
            "Status": status,
            "Items": items,
            "Quantity": quantity,
            "Department Name And Address": " ".join(department.split()) if department else "",
            "Start Date": start_date,
            "End Date": end_date,
        })

    return records


def go_to_next_page(driver):
    try:
        next_link = driver.find_element(By.XPATH, "//a[contains(., 'Next') or contains(@aria-label, 'Next')]")
        classes = next_link.get_attribute("class") or ""
        if "disabled" in classes:
            return False
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", next_link)
        next_link.click()
        time.sleep(2)
        return True
    except NoSuchElementException:
        return False


def main():
    driver = make_driver()
    all_records = []

    try:
        apply_search_and_filters(driver)

        page_num = 1
        while True:
            print(f"--- Scraping results page {page_num} ---")
            records = scrape_listing_page(driver)
            print(f"Found {len(records)} bids on this page")
            all_records.extend(records)
            for r in records:
                print(f"  {r['BID No']} | {r['Status']} | {r['Start Date']} -> {r['End Date']}")

            if MAX_PAGES and page_num >= MAX_PAGES:
                break
            if not go_to_next_page(driver):
                break
            page_num += 1

    finally:
        driver.quit()

    df = pd.DataFrame(all_records)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\nDone. Saved {len(all_records)} records to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
