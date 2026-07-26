"""
Build script - compiles app/gem_scraper_gui.py into a single standalone
.exe that anyone can double-click, with no Python installation required.

HOW TO USE:
    1. On a Windows machine, install Python 3.10+ (https://python.org)
    2. Open a terminal in the REPO ROOT (the folder containing this
       file's parent, i.e. one level above app/) and run:
           pip install -r requirements.txt
           pip install pyinstaller
    3. Then run this script from the repo root:
           python app/build_exe.py
    4. Your app will appear at:  dist/GemBidScraper.exe
       Copy that single .exe to any Windows PC and double-click to run -
       no Python or pip needed on the target machine.

NOTE: Google Chrome must be installed on whichever machine actually RUNS
the .exe (webdriver-manager auto-downloads the matching chromedriver on
first run, but Chrome itself must already be present).
"""

import os
import PyInstaller.__main__

THIS_DIR = os.path.dirname(os.path.abspath(__file__))
GUI_SCRIPT = os.path.join(THIS_DIR, "gem_scraper_gui.py")

PyInstaller.__main__.run([
    GUI_SCRIPT,
    "--name=GemBidScraper",
    "--onefile",
    "--windowed",          # no console window behind the GUI
    "--noconfirm",
    "--clean",
])

print("\nBuild complete. Find your app at: dist/GemBidScraper.exe")
