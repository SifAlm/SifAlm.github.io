# Tamam Intelligence – Justlife Competitive Intelligence Scraper

This repository provides a single Google Colab cell that installs all dependencies, launches Playwright Chromium, maps the Justlife UAE site, and exports structured intelligence outputs.

## Usage

1. Open a fresh Google Colab notebook.
2. Copy the full contents of [`colab_cell.py`](./colab_cell.py).
3. Paste it into a single Python cell in Colab and run it.

The cell will:

- install Playwright, Chromium, and data-science dependencies;
- crawl service pages under `https://www.justlife.com/en-AE` with throttled Playwright automation;
- detect service models, enumerate option combinations (hours, packages, unit counts, etc.), and capture price breakdowns;
- gather add-ons, availability slots, offers, and selector diagnostics;
- generate CSV exports and a styled `Justlife_CRM.xlsx` workbook with filters, freeze panes, and conditional formatting;
- emit `NOTES.txt` summarizing detections, blockers, and QA flags;
- save annotated screenshots in `justlife_intel/shots/` for QA traceability.

Environment variables can tweak behaviour when pasted into Colab:

- `BASE` – Base URL (default `https://www.justlife.com/en-AE`).
- `MARKET` – Market slug (default `dubai`).
- `DAYS` – Availability capture horizon (default `14`).
- `MAX_COMBOS` – Maximum option combinations per service (default `60`).
- `HEADLESS` – Set to `0` to run the browser non-headless.

All generated outputs are placed in a `justlife_intel/` directory within the Colab runtime and listed in the cell logs.
