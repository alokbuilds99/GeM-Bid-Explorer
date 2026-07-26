# GeM Bid Scraper

A tool that scrapes bid listings from the [GeM Bidplus
portal](https://bidplus.gem.gov.in/all-bids) for any keyword you choose,
and exports the results (Bid No, Status, Items, Quantity, Department,
Start/End Date) to an Excel file.

This project has two parts:

1. **[`core/`](core/)** — the scraping engine itself (a command-line
   Python script). Start here if you want to understand *how the
   scraping works* or want to run/modify it as a developer.
2. **[`app/`](app/)** — a desktop GUI built on top of the core engine,
   packaged into a standalone `.exe` so anyone can run it with a
   double-click, no coding or command line required.

---

## Part 1 — Core Scraper (`core/gem_scraper_v2_cli.py`)

This is the engine. It uses **Selenium** to drive a real Chrome browser
against the GeM Bidplus site and **pandas** to export results to Excel.

### How it works

1. **Opens the search page** and types in the configured keyword
   (`SEARCH_TERM` in the script).
2. **Applies filters** — ticks the "Bid/RA Status" and "Bid/RA Awarded"
   checkboxes so only awarded bids show up.
3. **Reads the listing page directly**, rather than opening each bid's
   individual detail page. Each bid card's text on the results page
   already contains everything needed (`BID NO:`, `Status:`, `Items:`,
   `Quantity:`, `Department Name And Address:`, `Start Date:`,
   `End Date:`), so the script:
   - grabs the full visible text of the page (`driver.find_element(...).text`)
   - splits it into one chunk per bid using `BID NO:` as the delimiter
     (`re.split(r"\n(?=BID NO:)", body_text)`)
   - regex-extracts each field out of every chunk
4. **Paginates automatically** — clicks "Next" until there are no more
   pages (or until `MAX_PAGES` is hit, if you set a limit for testing).
5. **Saves everything to Excel** via `pandas.DataFrame.to_excel(...)`.

This "read the listing page, don't open every bid" approach is what
makes it fast — it avoids one extra page load per bid.

### Run it

```bash
pip install -r requirements.txt
python core/gem_scraper_v2_cli.py
```

Edit the constants at the top of the file to change the keyword,
output filename, or headless/visible browser mode:

```python
SEARCH_TERM = "micom relays"
OUTPUT_FILE = "gem_bids_output.xlsx"
HEADLESS = False
MAX_PAGES = None  # or an int, e.g. 1, while testing
```

---

## Part 2 — Desktop App (`app/gem_scraper_gui.py`)

This wraps the same core logic in a **Tkinter GUI**, so non-technical
users can run it without touching any code.

### Features

- 🔍 Search **any keyword** — typed into a text box, not hardcoded
- 🖥️ Simple desktop window — no command line needed to use it
- 📄 Scrapes every page of results automatically (optional page limit)
- 🙈 Optional headless mode (run without a visible browser window)
- ✅ Toggle "Bid/RA Awarded only" filter on/off
- ⏹️ Stop button — cancel mid-scrape, partial results are still saved
- 📊 Exports directly to `.xlsx`
- 💾 Remembers your last-used settings between runs
- 📦 Ships as a standalone `.exe` — end users don't need Python installed

### Run it from source

```bash
pip install -r requirements.txt
python app/gem_scraper_gui.py
```

A window opens — type a keyword, choose where to save the Excel file,
and click **Start Scraping**.

### Build a standalone `.exe`

For end users who shouldn't need Python installed at all:

```bash
pip install pyinstaller
python app/build_exe.py
```

This produces `dist/GemBidScraper.exe` — a single file you can share
directly. The target machine just needs Google Chrome installed
(the matching chromedriver auto-downloads on first run).

> Pre-built binaries for this project (if available) are attached to
> the [Releases](../../releases) page instead of being committed to
> the repo.

---

## Project structure

```
gem-bid-scraper/
├── core/
│   └── gem_scraper_v2_cli.py   # Part 1: scraping engine (CLI)
├── app/
│   ├── gem_scraper_gui.py      # Part 2: desktop GUI app
│   └── build_exe.py            # Compiles the GUI into a standalone .exe
├── requirements.txt
├── LICENSE
└── README.md
```

## Requirements

- Python 3.10+
- Google Chrome installed
- Dependencies in `requirements.txt` (Selenium, webdriver-manager,
  pandas, openpyxl; PyInstaller only needed if building the `.exe`)

## Disclaimer

This tool automates browsing of a publicly accessible government
portal for personal research/data-collection purposes. Use responsibly
and in line with the [GeM portal's terms of use](https://gem.gov.in).
Site structure changes on GeM's end may require updating the selectors
in either script.

## License

MIT — see [LICENSE](LICENSE).
