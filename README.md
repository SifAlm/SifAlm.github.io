This repository ships a single Google Colab-ready cell that installs all dependencies, launches Playwright Chromium, navigates Justlife's booking flows, enumerates pricing combinations, and exports a board-grade intelligence workbook.

## Quick start

1. Open a fresh Google Colab notebook.
2. Copy the entire contents of [`colab_cell.py`](./colab_cell.py).
3. Paste the script into **one** Python cell and run it.

The cell will automatically:

- install Playwright, Chromium, and data-science dependencies;
- crawl `https://www.justlife.com/en-AE` with polite throttling, harvesting homepage, footer, and sitemap links;
- drive "Book now" CTAs to reach live checkout widgets, enforce Dubai as the active city, and close blocking modals;
- detect option groups (hours × pros, packages, unit quantities, etc.), click combinations, and wait for price updates before recording totals, VAT, and fees;
- enumerate package sections with `Add +` buttons, hours/professionals selectors, and bundle cards, logging selector evidence along the way;
- capture add-ons, availability slots (next 14 days), and offer badges where present;
- save CSV exports plus a styled `Justlife_CRM.xlsx` workbook (freeze panes, auto filters, conditional formats);
- emit `NOTES.txt` highlighting detections, blockers, and QA warnings;
- drop annotated screenshots in `justlife_intel/shots/` for downstream QA.

### Workbook layout

`Justlife_CRM.xlsx` includes the following sheets so pricing, availability, and QA can be audited quickly:

- **Services** – service × city roll-up with min/max/avg totals and coverage stats.
- **Combinations** – one row per enumerated option set with fee/VAT breakdowns and screenshot references.
- **AddOns** – upsell catalogue with pricing flags.
- **Availability** – date/slot capture for the configured window (default 14 days).
- **Offers** – captured offer badges and text.
- **Errors** – failed URLs with stage, message, and screenshot path.
- **Meta** – run KPIs, runtime, and version fingerprint.

### Configuration knobs

The following environment variables can be set inside the Colab cell before running the script to adjust behaviour:

- `HEADLESS` – set to `0` to watch the browser.
- `DAYS_AHEAD` – number of calendar days to enumerate for availability (default `14`).
- `MAX_COMBOS` – per-service cap on combination enumeration (default `60`).
- `BASE_URL` – override the Justlife base URL (default `https://www.justlife.com/en-AE`).

All generated artifacts live in the `justlife_intel/` folder of the Colab runtime and are listed in the run summary printed at the end of execution.
