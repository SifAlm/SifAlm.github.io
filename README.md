# Tamam Intelligence – Justlife Competitive Intelligence Scraper

This repository provides a single Google Colab cell that installs all dependencies, launches Playwright Chromium, crawls the Justlife UAE website, enumerates service configuration states, and exports boardroom-ready analytics.

## Usage

1. Open a fresh Google Colab notebook.
2. Copy the full contents of [`colab_cell.py`](./colab_cell.py).
3. Paste it into a single Python cell in Colab and run it.

The cell will:

- install Playwright, Chromium, and data-science dependencies;
- crawl service pages under `https://www.justlife.com/en-AE` with throttled Playwright automation;
- detect service models, enumerate option combinations (hours, packages, unit counts, etc.), and capture price breakdowns;
- gather add-ons, availability slots, offers, and selector diagnostics;
- generate CSV exports and a fully formatted `Justlife_CRM.xlsx` workbook with KPIs, conditional formatting, and charts;
- emit `NOTES.txt` summarizing detections, blockers, and QA flags;
- attempt to auto-download the Excel workbook to the local machine.

Environment variables can tweak behaviour when pasted into Colab:

- `BASE` – Base URL (default `https://www.justlife.com/en-AE`).
- `MARKET` – Market slug (default `dubai`).
- `DAYS` – Availability capture horizon (default `14`).
- `MAX_COMBOS` – Maximum option combinations per service (default `60`).
- `HEADLESS` – Set to `0` to run the browser non-headless.

All generated outputs are placed in a `justlife_intel/` directory within the Colab runtime and listed in the cell logs.
