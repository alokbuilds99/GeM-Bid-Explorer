"""
GeM Bid Scraper - Desktop App (GUI)
====================================
A point-and-click desktop application for scraping bid-card info from
https://bidplus.gem.gov.in/all-bids for ANY search keyword - not just
"micom relays".

Features:
    - Type any keyword to search
    - Choose where to save the Excel output
    - Optional headless mode (browser hidden) or visible browser
    - Optional page limit (useful for quick tests)
    - Live log window showing progress
    - Stop button to cancel a running scrape
    - Remembers your last-used settings (settings.json next to the app)

This file can be run directly with:
    python gem_scraper_gui.py

...or compiled into a standalone .exe (see build_exe.py / README.md in
this folder) so a user with NO Python installed can just double-click
it and run it.

Requires (only needed to RUN FROM SOURCE - not needed if using the
compiled .exe):
    pip install selenium webdriver-manager openpyxl pandas
"""

import json
import os
import queue
import re
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# --------------------------------------------------------------------------
# Settings persistence (remembers last used values between runs)
# --------------------------------------------------------------------------
SETTINGS_FILE = os.path.join(
    os.path.dirname(os.path.abspath(sys.argv[0])), "gem_scraper_settings.json"
)

DEFAULT_SETTINGS = {
    "search_term": "micom relays",
    "output_file": os.path.join(os.path.expanduser("~"), "Desktop", "gem_bids_output.xlsx"),
    "headless": False,
    "max_pages": "",  # empty string = no limit
    "only_awarded": True,
}


def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            merged = dict(DEFAULT_SETTINGS)
            merged.update(data)
            return merged
        except Exception:
            pass
    return dict(DEFAULT_SETTINGS)


def save_settings(settings):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
    except Exception:
        pass


# --------------------------------------------------------------------------
# Scraper logic (adapted from gem_scraper_v2.py to report through a
# log_fn callback and respect a stop_event instead of hardcoded config)
# --------------------------------------------------------------------------
BASE_URL = "https://bidplus.gem.gov.in/all-bids"
WAIT_TIMEOUT = 25


class ScraperCancelled(Exception):
    pass


def make_driver(headless):
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options

    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1400,1000")
    opts.add_argument("--disable-notifications")
    opts.add_argument("--log-level=3")

    try:
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


def find_visible(driver, selectors, timeout=WAIT_TIMEOUT, stop_event=None):
    from selenium.common.exceptions import StaleElementReferenceException, TimeoutException

    end_time = time.time() + timeout
    last_err = None
    while time.time() < end_time:
        if stop_event is not None and stop_event.is_set():
            raise ScraperCancelled()
        for by, selector in selectors:
            try:
                elements = driver.find_elements(by, selector)
                for el in elements:
                    try:
                        if el.is_displayed() and el.is_enabled():
                            driver.execute_script(
                                "arguments[0].scrollIntoView({block:'center'});", el
                            )
                            return el
                    except StaleElementReferenceException:
                        continue
            except Exception as e:
                last_err = e
        time.sleep(0.5)
    raise TimeoutException(f"No visible/enabled element found for selectors: {selectors}. Last error: {last_err}")


def apply_search_and_filters(driver, search_term, only_awarded, stop_event, log_fn):
    from selenium.webdriver.common.by import By
    from selenium.common.exceptions import TimeoutException

    log_fn(f"Opening {BASE_URL} ...")
    driver.get(BASE_URL)
    time.sleep(2)

    search_box = find_visible(driver, [
        (By.ID, "searchBid"),
        (By.CSS_SELECTOR, "input[placeholder*='Bid' i]"),
        (By.CSS_SELECTOR, "#bid-listing input.form-control"),
        (By.CSS_SELECTOR, "input.form-control"),
    ], stop_event=stop_event)
    search_box.click()
    search_box.clear()
    search_box.send_keys(search_term)
    log_fn(f"Typed search term: '{search_term}'")

    try:
        search_btn = find_visible(driver, [
            (By.ID, "searchBidRA"),
            (By.CSS_SELECTOR, "span.search-icon"),
            (By.CSS_SELECTOR, "button.btn-search"),
            (By.XPATH, "//button[contains(@class,'search')]"),
        ], timeout=5, stop_event=stop_event)
        search_btn.click()
    except TimeoutException:
        search_box.submit()

    time.sleep(2)

    if only_awarded:
        bid_ra_status_checkbox = find_visible(driver, [
            (By.XPATH, "//label[contains(., 'Bid/RA Status')]/preceding-sibling::input"),
            (By.XPATH, "//label[contains(., 'Bid/RA Status')]/parent::*//input"),
            (By.ID, "bidRaStatus"),
        ], stop_event=stop_event)
        if not bid_ra_status_checkbox.is_selected():
            bid_ra_status_checkbox.click()
        time.sleep(1)

        awarded_checkbox = find_visible(driver, [
            (By.XPATH, "//label[contains(., 'Bid /RA Awarded') or contains(., 'Bid/RA Awarded')]/preceding-sibling::input"),
            (By.XPATH, "//label[contains(., 'Bid /RA Awarded') or contains(., 'Bid/RA Awarded')]/parent::*//input"),
            (By.ID, "bidAwarded"),
        ], stop_event=stop_event)
        if not awarded_checkbox.is_selected():
            awarded_checkbox.click()
        log_fn("Applied filter: Bid/RA Awarded only")

    time.sleep(2)


def scrape_listing_page(driver):
    from selenium.webdriver.common.by import By

    body_text = driver.find_element(By.TAG_NAME, "body").text
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
    from selenium.webdriver.common.by import By
    from selenium.common.exceptions import NoSuchElementException

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


def run_scrape(search_term, output_file, headless, max_pages, only_awarded, stop_event, log_fn, done_fn):
    """Runs in a background thread. Reports progress via log_fn(str) and
    calls done_fn(success: bool, message: str) when finished."""
    import pandas as pd

    all_records = []
    driver = None
    try:
        log_fn("Launching Chrome ...")
        driver = make_driver(headless)

        apply_search_and_filters(driver, search_term, only_awarded, stop_event, log_fn)

        page_num = 1
        while True:
            if stop_event.is_set():
                raise ScraperCancelled()

            log_fn(f"--- Scraping results page {page_num} ---")
            records = scrape_listing_page(driver)
            log_fn(f"Found {len(records)} bids on this page")
            all_records.extend(records)
            for r in records:
                log_fn(f"  {r['BID No']} | {r['Status']} | {r['Start Date']} -> {r['End Date']}")

            if max_pages and page_num >= max_pages:
                break
            if stop_event.is_set():
                raise ScraperCancelled()
            if not go_to_next_page(driver):
                break
            page_num += 1

        df = pd.DataFrame(all_records)
        df.to_excel(output_file, index=False)
        done_fn(True, f"Done. Saved {len(all_records)} records to:\n{output_file}")

    except ScraperCancelled:
        if all_records:
            try:
                df = pd.DataFrame(all_records)
                df.to_excel(output_file, index=False)
                done_fn(False, f"Stopped by user. Partial results ({len(all_records)} records) saved to:\n{output_file}")
            except Exception:
                done_fn(False, "Stopped by user. No results were saved.")
        else:
            done_fn(False, "Stopped by user. No results were saved.")
    except Exception as e:
        done_fn(False, f"Error: {e}")
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass


# --------------------------------------------------------------------------
# GUI
# --------------------------------------------------------------------------
class GemScraperApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("GeM Bid Scraper")
        self.geometry("720x560")
        self.minsize(640, 480)

        self.settings = load_settings()
        self.stop_event = threading.Event()
        self.worker_thread = None
        self.log_queue = queue.Queue()

        self._build_ui()
        self.after(100, self._poll_log_queue)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------------- UI layout ----------------
    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        frm = ttk.Frame(self)
        frm.pack(fill="x", **pad)

        ttk.Label(frm, text="Search keyword:").grid(row=0, column=0, sticky="w")
        self.search_var = tk.StringVar(value=self.settings["search_term"])
        ttk.Entry(frm, textvariable=self.search_var, width=45).grid(
            row=0, column=1, columnspan=2, sticky="we", padx=(6, 0)
        )

        ttk.Label(frm, text="Save results to:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.output_var = tk.StringVar(value=self.settings["output_file"])
        ttk.Entry(frm, textvariable=self.output_var, width=45).grid(
            row=1, column=1, sticky="we", padx=(6, 0), pady=(8, 0)
        )
        ttk.Button(frm, text="Browse...", command=self._browse_output).grid(
            row=1, column=2, sticky="w", padx=(6, 0), pady=(8, 0)
        )

        opts = ttk.Frame(self)
        opts.pack(fill="x", **pad)

        self.headless_var = tk.BooleanVar(value=self.settings["headless"])
        ttk.Checkbutton(opts, text="Run in background (hide browser window)",
                        variable=self.headless_var).grid(row=0, column=0, sticky="w")

        self.only_awarded_var = tk.BooleanVar(value=self.settings["only_awarded"])
        ttk.Checkbutton(opts, text="Only Bid/RA Awarded results",
                        variable=self.only_awarded_var).grid(row=1, column=0, sticky="w")

        ttk.Label(opts, text="Limit pages (blank = all):").grid(row=0, column=1, sticky="e", padx=(20, 6))
        self.max_pages_var = tk.StringVar(value=self.settings["max_pages"])
        ttk.Entry(opts, textvariable=self.max_pages_var, width=8).grid(row=0, column=2, sticky="w")

        frm.columnconfigure(1, weight=1)
        opts.columnconfigure(1, weight=1)

        btns = ttk.Frame(self)
        btns.pack(fill="x", **pad)
        self.run_btn = ttk.Button(btns, text="Start Scraping", command=self._start)
        self.run_btn.pack(side="left")
        self.stop_btn = ttk.Button(btns, text="Stop", command=self._stop, state="disabled")
        self.stop_btn.pack(side="left", padx=(8, 0))

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(self, textvariable=self.status_var, foreground="#555").pack(
            fill="x", padx=10, anchor="w"
        )

        log_frame = ttk.LabelFrame(self, text="Log")
        log_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.log_text = tk.Text(log_frame, wrap="word", state="disabled", font=("Consolas", 9))
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scrollbar.set)

    # ---------------- Helpers ----------------
    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel file", "*.xlsx")],
            initialfile=os.path.basename(self.output_var.get() or "gem_bids_output.xlsx"),
        )
        if path:
            self.output_var.set(path)

    def _log(self, msg):
        # Thread-safe: called from the worker thread, just queues the message
        self.log_queue.put(msg)

    def _poll_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_text.configure(state="normal")
                self.log_text.insert("end", msg + "\n")
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(150, self._poll_log_queue)

    def _set_running_ui(self, running):
        self.run_btn.configure(state="disabled" if running else "normal")
        self.stop_btn.configure(state="normal" if running else "disabled")

    def _start(self):
        search_term = self.search_var.get().strip()
        output_file = self.output_var.get().strip()

        if not search_term:
            messagebox.showwarning("Missing keyword", "Please enter a search keyword.")
            return
        if not output_file:
            messagebox.showwarning("Missing output file", "Please choose where to save the Excel file.")
            return

        max_pages_raw = self.max_pages_var.get().strip()
        max_pages = None
        if max_pages_raw:
            try:
                max_pages = int(max_pages_raw)
                if max_pages <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("Invalid page limit", "Page limit must be a positive whole number, or left blank.")
                return

        # Persist settings for next time
        self.settings.update({
            "search_term": search_term,
            "output_file": output_file,
            "headless": self.headless_var.get(),
            "max_pages": max_pages_raw,
            "only_awarded": self.only_awarded_var.get(),
        })
        save_settings(self.settings)

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        self.stop_event = threading.Event()
        self.status_var.set(f"Scraping '{search_term}' ...")
        self._set_running_ui(True)

        self.worker_thread = threading.Thread(
            target=run_scrape,
            args=(
                search_term,
                output_file,
                self.headless_var.get(),
                max_pages,
                self.only_awarded_var.get(),
                self.stop_event,
                self._log,
                self._on_done,
            ),
            daemon=True,
        )
        self.worker_thread.start()

    def _stop(self):
        self.stop_event.set()
        self.status_var.set("Stopping... (finishing current page)")
        self.stop_btn.configure(state="disabled")

    def _on_done(self, success, message):
        # Called from the worker thread -> marshal back onto the UI thread
        def update():
            self._set_running_ui(False)
            self.status_var.set(message.splitlines()[0])
            self._log(message)
            if success:
                messagebox.showinfo("Finished", message)
            else:
                messagebox.showwarning("Stopped / Error", message)
        self.after(0, update)

    def _on_close(self):
        if self.worker_thread and self.worker_thread.is_alive():
            if not messagebox.askyesno("Scrape in progress", "A scrape is still running. Quit anyway?"):
                return
            self.stop_event.set()
        self.destroy()


if __name__ == "__main__":
    app = GemScraperApp()
    app.mainloop()
