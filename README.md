# 🚀 GeM Bid Explorer

> 🔎 **A powerful desktop tool to search, filter, and export GeM Bid listings to Excel in just a few clicks.**

**GeM Bid Explorer** automates the collection of bid information from the **[GeM Bidplus Portal](https://bidplus.gem.gov.in/all-bids)**. Simply enter any keyword, and the tool extracts bid details such as **Bid Number, Status, Items, Quantity, Department, Start & End Dates**, then exports everything into a neatly formatted **Excel (.xlsx)** file.

Whether you're a **supplier**, **sales professional**, **tender analyst**, or **business development team**, GeM Bid Explorer helps you save hours of manual searching. ⚡

---

# ✨ What's Included?

The project is divided into two parts:

## 🧠 1. `core/` — Scraping Engine

The heart of the project.

Designed for developers who want to understand, customize, or extend the scraping logic.

📂 **Location:** `core/`

---

## 🖥️ 2. `app/` — Desktop Application

A user-friendly Windows application built on top of the scraping engine.

No coding.
No command line.
Just launch the application and start scraping.

📂 **Location:** `app/`

---

# 🧠 Part 1 — Core Scraper (`core/gem_scraper_v2_cli.py`)

The scraper uses:

* 🐍 Python
* 🌐 Selenium WebDriver
* 📊 Pandas
* 📄 OpenPyXL

to automate Chrome and collect bid information directly from the GeM Bidplus portal.

---

## ⚙️ How It Works

### 🔍 Step 1 — Search

The script opens the GeM Bidplus portal and searches for your chosen keyword.

Example:

```python
SEARCH_TERM = "micom relays"
```

---

### ✅ Step 2 — Apply Filters

Automatically enables:

* ✔️ Bid/RA Status
* ✔️ Bid/RA Awarded

so only awarded bids are displayed.

---

### 📑 Step 3 — Extract Data

Instead of opening every individual bid page (which is much slower), the scraper reads the information already available on the search results page.

For each bid it extracts:

* 🆔 Bid Number
* 📌 Status
* 📦 Item Name
* 🔢 Quantity
* 🏢 Department Name
* 📅 Start Date
* ⏰ End Date

using efficient text parsing and regular expressions.

⚡ This approach dramatically reduces page loads, making the scraper significantly faster.

---

### 🔄 Step 4 — Automatic Pagination

The scraper continues clicking **Next** until:

* 📄 All pages are processed, or
* 🎯 A custom `MAX_PAGES` limit is reached.

---

### 📊 Step 5 — Export to Excel

All collected information is automatically saved into an Excel spreadsheet using Pandas.

Example output:

```
gem_bids_output.xlsx
```

---

## ▶️ Run the Core Scraper

Install dependencies:

```bash
pip install -r requirements.txt
```

Run:

```bash
python core/gem_scraper_v2_cli.py
```

---

### ⚙️ Customize

Simply edit the variables near the top of the script:

```python
SEARCH_TERM = "micom relays"
OUTPUT_FILE = "gem_bids_output.xlsx"
HEADLESS = False
MAX_PAGES = None
```

---

# 🖥️ Part 2 — Desktop Application (`app/gem_scraper_gui.py`)

Built with **Tkinter**, the desktop application allows anyone to use the scraper without writing a single line of code.

Launch it, enter a keyword, choose where to save the file, and click **Start Scraping**.

That's it. 🎉

---

# ✨ Features

* 🔍 Search **any keyword**
* 🖥️ Simple and intuitive desktop interface
* 📄 Automatically scrapes all result pages
* 👀 Optional visible browser mode
* 🙈 Optional headless mode
* ✅ Enable or disable **Awarded Only** filter
* ⏹️ Stop scraping anytime (partial results are safely saved)
* 📊 Export directly to Excel (.xlsx)
* 💾 Automatically remembers previous settings
* 📦 Share as a standalone Windows `.exe`
* ⚡ No Python knowledge required for end users

---

## ▶️ Run From Source

Install dependencies:

```bash
pip install -r requirements.txt
```

Run:

```bash
python app/gem_scraper_gui.py
```

A desktop window will open where you can:

* 🔍 Enter a keyword
* 📂 Choose the Excel save location
* 🚀 Click **Start Scraping**

---

# 📦 Build a Standalone Windows Executable

Want to share the application with others?

Install PyInstaller:

```bash
pip install pyinstaller
```

Build:

```bash
python app/build_exe.py
```

The compiled executable will be available as:

```
dist/GemBidExplorer.exe
```

Users only need:

* ✅ Google Chrome installed
* ✅ Internet connection

ChromeDriver is automatically downloaded during the first run.

> 📥 Pre-built releases (when available) are published under the **Releases** section of this repository instead of being committed to source control.

---

# 📂 Project Structure

```text
gem-bid-explorer/
├── core/
│   └── gem_scraper_v2_cli.py      # 🧠 Core scraping engine
│
├── app/
│   ├── gem_scraper_gui.py         # 🖥️ Desktop GUI
│   └── build_exe.py               # 📦 Build executable
│
├── requirements.txt
├── LICENSE
└── README.md
```

---

# 📋 Requirements

* 🐍 Python 3.10+
* 🌐 Google Chrome
* 📦 Packages listed in `requirements.txt`

  * Selenium
  * webdriver-manager
  * pandas
  * openpyxl
* 🛠️ PyInstaller (only required when building the executable)

---

# ⚠️ Disclaimer

This project automates interactions with the **publicly accessible GeM Bidplus portal** for research, analysis, and data collection purposes.

Please use this tool responsibly and ensure your usage complies with the **GeM Portal Terms of Use**.

As the GeM website evolves, future interface or structural changes may require updates to the scraper.

---

# 📜 License

Released under the **MIT License**.

See the [LICENSE](LICENSE) file for complete details.

---

