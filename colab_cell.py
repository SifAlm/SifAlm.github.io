import asyncio
import importlib
import json
import math
import os
import random
import re
import subprocess
import sys
from collections import defaultdict
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import urljoin, urlparse


###############################################################################
# Dependency bootstrap                                                         #
###############################################################################


def ensure_dependencies() -> None:
    """Install runtime dependencies required for the scraper."""

    packages = {
        "playwright": "playwright",
        "pandas": "pandas",
        "openpyxl": "openpyxl",
        "bs4": "beautifulsoup4",
        "lxml": "lxml",
        "tqdm": "tqdm",
        "fake_useragent": "fake-useragent",
        "tenacity": "tenacity",
        "requests": "requests",
        "numpy": "numpy",
        "PIL": "pillow",
    }

    missing: List[str] = []
    for module_name, package in packages.items():
        try:
            importlib.import_module(module_name)
        except ImportError:
            missing.append(package)

    if missing:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", *missing])

    subprocess.check_call(
        [sys.executable, "-m", "playwright", "install", "--with-deps", "chromium"]
    )


ensure_dependencies()

import contextlib
import time

import nest_asyncio
import numpy as np
import pandas as pd
import requests
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
from playwright.async_api import (
    Browser,
    BrowserContext,
    Locator,
    Page,
    TimeoutError as PlaywrightTimeout,
    async_playwright,
)
nest_asyncio.apply()


###############################################################################
# Global configuration                                                         #
###############################################################################

BASE_URL = os.environ.get("BASE_URL", "https://www.justlife.com/en-AE")
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", "justlife_intel"))
SHOTS_DIR = OUTPUT_DIR / "shots"
CHECKPOINT_PATH = OUTPUT_DIR / "run.json"
NOTES_PATH = OUTPUT_DIR / "NOTES.txt"
EXCEL_PATH = OUTPUT_DIR / "Justlife_CRM.xlsx"

OUTPUT_DIR.mkdir(exist_ok=True)
SHOTS_DIR.mkdir(exist_ok=True)

TARGET_SERVICE_COUNT = 120
MAX_COMBOS_PER_SERVICE = int(os.environ.get("MAX_COMBOS", "60"))
DAYS_AHEAD = int(os.environ.get("DAYS_AHEAD", "14"))
HEADLESS = os.environ.get("HEADLESS", "1") not in {"0", "false", "False"}
CITY_DEFAULT = os.environ.get("CITY", "dubai")
USER_AGENT = UserAgent().chrome
RANDOM_SLEEP_RANGE = (0.35, 0.95)

CTA_TEXT = re.compile(
    r"(book now|book|schedule|checkout|continue|select date|select time|get started)",
    re.IGNORECASE,
)
SERVICE_KEYWORDS = re.compile(
    r"(clean|maid|deep|move|villa|apartment|sofa|carpet|mattress|curtain|ac|duct|coil|pest|laundry|salon|spa|wax|brow|lash|nail|men|women|pet|plumb|electric|handyman|doctor|nurse|lab|physio|iv|oxygen|pack|therapy|clinic|massage|furniture)",
    re.IGNORECASE,
)
EXCLUDE_KEYWORDS = re.compile(
    r"(blog|privacy|terms|faq|career|about|login|signup|sitemap|support|polic|terms|gift|voucher)",
    re.IGNORECASE,
)
CITY_KEYWORDS = re.compile(r"/(dubai|abu-dhabi|sharjah|ajman)(/|$)")

PRICE_PATTERNS = [
    "[data-testid*='grand']",
    "[data-testid*='total']",
    "[data-testid*='price']",
    "[class*='grand']",
    "[class*='total']",
    "[class*='price']",
    "text=/AED/",
]

OPTION_SECTION_HINTS = [
    "hour",
    "duration",
    "cleaner",
    "professional",
    "maid",
    "bed",
    "bath",
    "room",
    "unit",
    "piece",
    "item",
    "package",
    "combo",
    "bundle",
    "area",
    "sqft",
    "material",
]

ADDON_HINTS = ["add-on", "addon", "extra", "upgrade", "upsell"]
SLOT_HINTS = ["slot", "time", "am", "pm"]


###############################################################################
# Data containers                                                              #
###############################################################################


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")


@dataclass
class ServiceSummary:
    service_id: str
    service_name: str
    category: str
    city: str
    service_url: str
    checkout_url: str
    pricing_model: str
    min_price_aed: Optional[float]
    max_price_aed: Optional[float]
    avg_price_aed: Optional[float]
    combos_count: int
    addons_count: int
    availability_days: int
    has_offers: bool
    extracted_at: str


@dataclass
class CombinationRow:
    service_id: str
    city: str
    combination_key: str
    dimensions_json: str
    base_price_aed: Optional[float]
    fees_aed: Optional[float]
    vat_aed: Optional[float]
    total_aed: Optional[float]
    currency: str
    discount_label: Optional[str]
    extraction_path: str
    screenshot_path: Optional[str]
    extracted_at: str


@dataclass
class AddonRow:
    service_id: str
    city: str
    addon_name: str
    addon_description: str
    addon_price_aed: Optional[float]
    is_required: bool
    is_recommended: bool
    appears_with_combination_key: Optional[str]
    extracted_at: str


@dataclass
class AvailabilityRow:
    service_id: str
    city: str
    date: str
    slot_label: str
    status: str
    scraped_from_combination_key: Optional[str]
    extracted_at: str


@dataclass
class OfferRow:
    service_id: str
    city: str
    offer_badge: str
    offer_text: str
    offer_terms: str
    applicable_to_combination_key: Optional[str]
    extracted_at: str


@dataclass
class ErrorRow:
    url: str
    stage: str
    message: str
    selector_used: str
    screenshot_path: Optional[str]
    ts: str


@dataclass
class MetaRow:
    total_discovered: int
    total_bookable: int
    processed: int
    skipped: int
    combos: int
    addons: int
    availability: int
    offers: int
    errors: int
    runtime_minutes: float
    timestamp: str
    code_version: str = "colab_cell_v2"


###############################################################################
# Utility helpers                                                              #
###############################################################################


def random_sleep() -> float:
    return random.uniform(*RANDOM_SLEEP_RANGE)


def canonicalize(url: str) -> str:
    if not url:
        return ""
    joined = urljoin(BASE_URL, url)
    parsed = urlparse(joined)
    path = re.sub(r"/en-AE/(en-AE/)+", "/en-AE/", parsed.path or "/")
    path = re.sub(r"/+", "/", path)
    if not path.startswith("/"):
        path = f"/{path}"
    normalized = f"{parsed.scheme or 'https'}://{parsed.netloc or urlparse(BASE_URL).netloc}{path}"
    if parsed.query:
        normalized = f"{normalized}?{parsed.query}"
    return normalized.rstrip("/")


def looks_like_service(url: str) -> bool:
    if not url:
        return False
    if EXCLUDE_KEYWORDS.search(url):
        return False
    return "/en-AE" in url and SERVICE_KEYWORDS.search(url) is not None


def service_city_from_url(url: str) -> str:
    match = CITY_KEYWORDS.search(url)
    if match:
        return match.group(1)
    return CITY_DEFAULT


def parse_price(value: str) -> Optional[float]:
    if not value:
        return None
    cleaned = re.sub(r"[^0-9.,]", "", value)
    cleaned = cleaned.replace(",", "")
    if not cleaned:
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


###############################################################################
# Discovery                                                                    #
###############################################################################


async def extract_links(page: Page) -> List[Tuple[str, str]]:
    anchors = await page.eval_on_selector_all(
        "a[href]",
        "(nodes) => nodes.map(a => [a.getAttribute('href') || '', (a.textContent || '').trim()])",
    )
    return [(canonicalize(href), text) for href, text in anchors if href]


async def discover_services(context: BrowserContext) -> List[Dict[str, Any]]:
    sources: Dict[str, List[str]] = defaultdict(list)
    page = await context.new_page()
    await page.goto(BASE_URL, wait_until="networkidle", timeout=60000)
    await handle_cookie_banner(page)
    links = await extract_links(page)
    for href, text in links:
        if looks_like_service(href):
            sources["homepage"].append(href)
    footer_links = await page.eval_on_selector_all(
        "footer a[href]",
        "(nodes) => nodes.map(a => a.getAttribute('href'))",
    )
    for href in footer_links:
        normalized = canonicalize(href)
        if looks_like_service(normalized):
            sources["footer"].append(normalized)
    await page.close()

    sitemap_urls: List[str] = []
    with contextlib.suppress(Exception):
        resp = requests.get(urljoin(BASE_URL, "/sitemap.xml"), timeout=15)
        if resp.ok:
            soup = BeautifulSoup(resp.text, "xml")
            for loc in soup.find_all("loc"):
                loc_url = canonicalize(loc.text.strip())
                if looks_like_service(loc_url):
                    sitemap_urls.append(loc_url)
    for href in sitemap_urls:
        sources["sitemap"].append(href)

    seed_slugs = [
        "house-cleaning",
        "deep-cleaning",
        "move-in-out-cleaning-services",
        "maid-service",
        "sofa-cleaning",
        "curtain-cleaning",
        "mattress-cleaning",
        "ac-cleaning",
        "pest-control",
        "laundry-and-dry-cleaning",
        "handyman-and-maintenance",
        "carpet-cleaning",
        "salon-services-at-home",
        "mens-salon",
        "mens-grooming",
        "womens-salon",
        "lab-tests-at-home",
        "nurse-care-at-home",
        "physiotherapy-at-home",
        "iv-therapy-at-home",
        "doctor-on-call",
    ]
    for slug in seed_slugs:
        for city in ["", "dubai", "abu-dhabi", "sharjah", "ajman"]:
            parts = ["/en-AE"]
            if city:
                parts.append(city)
            parts.append(slug)
            url = canonicalize("/".join(parts))
            sources["seed"].append(url)
            sources["seed"].append(f"{url}/checkout/flex")

    all_urls: Dict[str, Dict[str, Any]] = {}
    for source, hrefs in sources.items():
        for href in hrefs:
            if not looks_like_service(href):
                continue
            city = service_city_from_url(href)
            if href not in all_urls:
                all_urls[href] = {"url": href, "source": set(), "city": city}
            all_urls[href]["source"].add(source)

    for value in all_urls.values():
        value["source"] = sorted(value["source"])

    with CHECKPOINT_PATH.open("w", encoding="utf-8") as fh:
        json.dump({"discovered": list(all_urls.values())}, fh, indent=2)

    return list(all_urls.values())


###############################################################################
# Page interactions                                                            #
###############################################################################


async def handle_cookie_banner(page: Page) -> None:
    for pattern in ["accept", "agree", "ok"]:
        with contextlib.suppress(PlaywrightTimeout):
            button = page.get_by_role("button", name=re.compile(pattern, re.I))
            await button.first.click(timeout=1500)
            await page.wait_for_timeout(int(random_sleep() * 1000))
            break


async def ensure_location(page: Page, city: str) -> None:
    await page.wait_for_timeout(400)
    for _ in range(2):
        with contextlib.suppress(Exception):
            chip = page.locator("[data-testid*='location'], button", has_text=re.compile(city, re.I)).first
            await chip.click(timeout=1500)
            await page.wait_for_timeout(350)
        locator = page.locator("input[placeholder*='Search']")
        if await locator.count():
            await locator.first.fill("Downtown Dubai")
            await page.wait_for_timeout(350)
            with contextlib.suppress(Exception):
                await locator.first.press("Enter")
            with contextlib.suppress(Exception):
                suggestion = page.get_by_role("option").first
                await suggestion.click(timeout=2000)
            with contextlib.suppress(Exception):
                confirm = page.get_by_role("button", name=re.compile("Select|Save|Continue", re.I)).first
                await confirm.click(timeout=2000)
            await page.wait_for_timeout(800)
            break
        else:
            break


async def click_book_cta(page: Page) -> Optional[str]:
    buttons = page.locator("a, button").filter(has_text=CTA_TEXT)
    if await buttons.count():
        element = buttons.first
        href = await element.get_attribute("href")
        with contextlib.suppress(Exception):
            await element.click()
        return canonicalize(href) if href else None
    return None


async def ensure_checkout(page: Page, url: str) -> Tuple[str, bool]:
    current = page.url
    if "checkout" in current:
        return current, True
    checkout_href = None
    with contextlib.suppress(Exception):
        checkout_href = await click_book_cta(page)
    if checkout_href and "checkout" in checkout_href:
        await page.wait_for_timeout(500)
        if checkout_href != page.url:
            await page.goto(checkout_href, wait_until="networkidle", timeout=60000)
        return page.url, True
    anchors = await extract_links(page)
    for href, _ in anchors:
        if "checkout" in href:
            await page.goto(href, wait_until="networkidle", timeout=60000)
            return page.url, True
    return page.url, "checkout" in page.url


async def wait_for_widget(page: Page) -> bool:
    probes = [
        "[data-testid*='booking']",
        "[data-testid*='checkout']",
        "[class*='booking']",
        "[class*='checkout']",
        "text=/Step 1/i",
    ]
    for selector in probes:
        with contextlib.suppress(PlaywrightTimeout):
            await page.locator(selector).first.wait_for(timeout=4000)
            return True
    return False


async def handle_interruptions(page: Page) -> None:
    await handle_cookie_banner(page)
    for label in ["Time Slot Expired", "Session expired"]:
        with contextlib.suppress(PlaywrightTimeout):
            modal = page.get_by_text(label, exact=False)
            if await modal.count():
                button = page.get_by_role("button", name=re.compile("Select another slot|Close", re.I))
                with contextlib.suppress(Exception):
                    await button.first.click(timeout=1500)
                    await page.wait_for_timeout(400)


async def capture_screenshot(page: Page, service_id: str, label: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", f"{service_id}-{label}".lower()).strip("-")
    path = SHOTS_DIR / f"{slug}.png"
    await page.screenshot(path=path, full_page=True)
    return str(path.relative_to(OUTPUT_DIR))


###############################################################################
# Option discovery & enumeration                                                #
###############################################################################


async def collect_clickable_options(page: Page) -> Dict[str, List[Dict[str, Any]]]:
    controls: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    script = """
        (body, hints) => {
            const lower = (s) => (s || '').toLowerCase();
            const result = {};
            const add = (group, node) => {
                if (!node) return;
                const text = (node.innerText || '').trim();
                if (!text) return;
                if (!result[group]) result[group] = [];
                result[group].push({
                    text,
                    selector: node.getAttribute('data-testid') || node.getAttribute('name') || node.tagName,
                });
            };
            body.querySelectorAll('section, div, form').forEach(section => {
                const heading = section.querySelector('h1, h2, h3, h4, label, span');
                if (!heading) return;
                const headingText = lower(heading.innerText || '');
                for (const [group, tokens] of Object.entries(hints)) {
                    if (tokens.some(token => headingText.includes(token))) {
                        section.querySelectorAll('button, [role=\"button\"], option, select option').forEach(btn => add(group, btn));
                    }
                }
            });
            return result;
        }
    """
    hints = {
        "hours": ["hour", "duration"],
        "pros": ["cleaner", "professional", "maid"],
        "bedrooms": ["bed", "room"],
        "bathrooms": ["bath", "toilet"],
        "packages": ["package", "combo", "bundle", "plan"],
        "units": ["unit", "piece", "item"],
        "materials": ["material"],
    }
    elements = await page.eval_on_selector_all('body', script, hints)
    for group, values in (elements or {}).items():
        if isinstance(values, list):
            controls[group].extend(values)
    return controls

async def click_option(page: Page, option_text: str) -> None:
    locator = page.locator("button, [role='button'], option")
    candidate = locator.filter(has_text=re.compile(re.escape(option_text), re.I))
    if await candidate.count():
        with contextlib.suppress(Exception):
            await candidate.first.click()
            await page.wait_for_timeout(int(random_sleep() * 1000))


async def read_price_summary(page: Page) -> Tuple[Optional[float], Optional[float], Optional[float], Optional[float], str, Optional[str]]:
    text_chunks: List[str] = []
    for selector in PRICE_PATTERNS:
        locator = page.locator(selector)
        if await locator.count():
            text = await locator.first.inner_text()
            if text and "AED" in text:
                text_chunks.append(text)
    combined = " ".join(dict.fromkeys(text_chunks))
    currency = "AED" if "AED" in combined else ""
    numbers = [parse_price(part) for part in re.findall(r"AED?\s*[0-9,.]+", combined)]
    grand_total = numbers[-1] if numbers else parse_price(combined)
    vat = None
    fees = None
    base = None
    vat_match = re.search(r"VAT[^0-9]*([0-9,.]+)", combined, re.I)
    if vat_match:
        vat = parse_price(vat_match.group(1))
    fee_match = re.search(r"fee[^0-9]*([0-9,.]+)", combined, re.I)
    if fee_match:
        fees = parse_price(fee_match.group(1))
    base_match = re.search(r"base[^0-9]*([0-9,.]+)", combined, re.I)
    if base_match:
        base = parse_price(base_match.group(1))
    discount_label = None
    badge_locator = page.locator("[class*='offer'], [data-testid*='offer'], [class*='discount']")
    if await badge_locator.count():
        discount_label = (await badge_locator.first.inner_text() or "").strip()
    return base, fees, vat, grand_total, currency or "AED", discount_label


async def enumerate_combinations(
    page: Page,
    service_id: str,
    city: str,
    checkout_url: str,
    service_name: str,
) -> Tuple[List[CombinationRow], Dict[str, Dict[str, Any]]]:
    rows: List[CombinationRow] = []
    applied_controls: Dict[str, List[str]] = defaultdict(list)
    controls = await collect_clickable_options(page)
    if not controls:
        base, fees, vat, total, currency, badge = await read_price_summary(page)
        screenshot = await capture_screenshot(page, service_id, "default")
        dimensions = {}
        rows.append(
            CombinationRow(
                service_id=service_id,
                city=city,
                combination_key=json.dumps(dimensions, sort_keys=True),
                dimensions_json=json.dumps(dimensions, sort_keys=True),
                base_price_aed=base,
                fees_aed=fees,
                vat_aed=vat,
                total_aed=total,
                currency=currency,
                discount_label=badge,
                extraction_path=checkout_url,
                screenshot_path=screenshot,
                extracted_at=utc_now(),
            )
        )
        return rows, applied_controls

    option_space: Dict[str, List[str]] = {}
    for group, values in controls.items():
        texts = []
        for item in values:
            text = item.get("text", "").strip()
            if text and text not in texts:
                texts.append(text)
        if texts:
            option_space[group] = texts[:6]

    permutations: List[Dict[str, str]] = []
    keys = list(option_space.keys())

    def backtrack(idx: int, current: Dict[str, str]) -> None:
        if len(permutations) >= MAX_COMBOS_PER_SERVICE:
            return
        if idx == len(keys):
            permutations.append(current.copy())
            return
        key = keys[idx]
        for value in option_space[key][:5]:
            current[key] = value
            backtrack(idx + 1, current)
        current.pop(key, None)

    backtrack(0, {})

    seen_totals: Dict[str, float] = {}
    for combo in permutations or [{}]:
        for key, value in combo.items():
            await click_option(page, value)
            applied_controls[key].append(value)
            await handle_interruptions(page)
        await page.wait_for_timeout(500)
        base, fees, vat, total, currency, badge = await read_price_summary(page)
        dim_json = json.dumps(combo, sort_keys=True)
        if total is not None and math.isfinite(total):
            previous = seen_totals.get(dim_json)
            if previous is not None and abs(previous - total) < 0.01:
                continue
            seen_totals[dim_json] = total
        screenshot = await capture_screenshot(page, service_id, slugify(dim_json)[:60] or "combo")
        rows.append(
            CombinationRow(
                service_id=service_id,
                city=city,
                combination_key=dim_json,
                dimensions_json=dim_json,
                base_price_aed=base,
                fees_aed=fees,
                vat_aed=vat,
                total_aed=total,
                currency=currency,
                discount_label=badge,
                extraction_path=checkout_url,
                screenshot_path=screenshot,
                extracted_at=utc_now(),
            )
        )
    return rows, applied_controls


async def scrape_addons(page: Page, service_id: str, city: str) -> List[AddonRow]:
    addons: List[AddonRow] = []
    sections = page.locator("text=/Add-?on|Extras|Upgrades|Optional/i").locator("xpath=ancestor::*[1]")
    if await sections.count() == 0:
        sections = page.locator("[class*='addon'], [data-testid*='addon']")
    count = await sections.count()
    for idx in range(count):
        section = sections.nth(idx)
        cards = section.locator("[class*='card'], article, li")
        if not await cards.count():
            cards = section.locator("*")
        inner_count = min(await cards.count(), 6)
        for j in range(inner_count):
            card = cards.nth(j)
            try:
                name = await card.locator("h2, h3, h4, strong").first.inner_text()
            except Exception:
                name = ""
            try:
                description = await card.inner_text()
            except Exception:
                description = ""
            price_match = re.search(r"AED\s*[0-9,.]+", description or "")
            price = parse_price(price_match.group(0)) if price_match else None
            addons.append(
                AddonRow(
                    service_id=service_id,
                    city=city,
                    addon_name=(name or description[:30] or "Addon").strip(),
                    addon_description=(description or "").strip(),
                    addon_price_aed=price,
                    is_required="required" in (description or "").lower(),
                    is_recommended="recommended" in (description or "").lower(),
                    appears_with_combination_key=None,
                    extracted_at=utc_now(),
                )
            )
    return addons


async def scrape_availability(page: Page, service_id: str, city: str) -> List[AvailabilityRow]:
    records: List[AvailabilityRow] = []
    date_triggers = [
        "[data-testid*='date']",
        "button:has-text('Date')",
        "input[type='date']",
    ]
    opened = False
    for selector in date_triggers:
        trigger = page.locator(selector)
        if await trigger.count():
            with contextlib.suppress(Exception):
                await trigger.first.click()
                await page.wait_for_timeout(400)
                opened = True
                break
    if not opened:
        return records
    calendar = page.locator("[class*='calendar'], [data-testid*='calendar']")
    for day_offset in range(DAYS_AHEAD):
        day_selector = f"[data-day-index='{day_offset}'], button:has-text('{day_offset + 1}')"
        day = calendar.locator(day_selector)
        if await day.count():
            with contextlib.suppress(Exception):
                await day.first.click()
                await page.wait_for_timeout(250)
        slots = page.locator("[class*='slot'], [data-testid*='slot'], [class*='time']")
        for idx in range(min(await slots.count(), 12)):
            slot = slots.nth(idx)
            try:
                text = (await slot.inner_text()).strip()
            except Exception:
                text = ""
            if not text:
                continue
            status = "available"
            class_name = await slot.get_attribute("class") or ""
            if "disable" in class_name.lower():
                status = "unavailable"
            records.append(
                AvailabilityRow(
                    service_id=service_id,
                    city=city,
                    date=(datetime.now(timezone.utc) + pd.Timedelta(days=day_offset)).date().isoformat(),
                    slot_label=text,
                    status=status,
                    scraped_from_combination_key=None,
                    extracted_at=utc_now(),
                )
            )
        with contextlib.suppress(Exception):
            next_btn = page.get_by_role("button", name=re.compile("Next", re.I))
            if await next_btn.count():
                await next_btn.click()
                await page.wait_for_timeout(300)
    with contextlib.suppress(Exception):
        close_btn = page.get_by_role("button", name=re.compile("Close|Done", re.I))
        if await close_btn.count():
            await close_btn.click()
    return records


async def scrape_offers(page: Page, service_id: str, city: str) -> List[OfferRow]:
    offers: List[OfferRow] = []
    locator = page.locator("[class*='offer'], [data-testid*='offer'], [class*='discount']")
    for idx in range(min(await locator.count(), 5)):
        node = locator.nth(idx)
        text = await node.inner_text()
        if text:
            offers.append(
                OfferRow(
                    service_id=service_id,
                    city=city,
                    offer_badge=text.strip(),
                    offer_text=text.strip(),
                    offer_terms="",
                    applicable_to_combination_key=None,
                    extracted_at=utc_now(),
                )
            )
    return offers


###############################################################################
# Processing                                                                   #
###############################################################################


def slugify(value: str) -> str:
    value = re.sub(r"[^a-zA-Z0-9-_]+", "-", value.strip().lower())
    return re.sub(r"-+", "-", value).strip("-") or "item"


async def process_service(
    context: BrowserContext,
    candidate: Dict[str, Any],
    service_idx: int,
    totals: Dict[str, int],
    services: List[ServiceSummary],
    combinations: List[CombinationRow],
    addons: List[AddonRow],
    availability: List[AvailabilityRow],
    offers: List[OfferRow],
    errors: List[ErrorRow],
) -> None:
    url = candidate["url"]
    city = candidate.get("city", CITY_DEFAULT)
    service_id = slugify(f"{service_idx}-{city}-{url.split('/')[-1]}")
    page = await context.new_page()
    try:
        await page.goto(url, wait_until="networkidle", timeout=60000)
    except Exception as exc:
        errors.append(
            ErrorRow(
                url=url,
                stage="discovery",
                message=str(exc),
                selector_used="goto",
                screenshot_path=None,
                ts=utc_now(),
            )
        )
        await page.close()
        totals["skipped"] += 1
        return

    await handle_cookie_banner(page)
    await handle_interruptions(page)

    checkout_url, at_checkout = await ensure_checkout(page, url)
    if not at_checkout:
        errors.append(
            ErrorRow(
                url=url,
                stage="checkout",
                message="Booking widget not detected",
                selector_used="cta",
                screenshot_path=await capture_screenshot(page, service_id, "no-widget"),
                ts=utc_now(),
            )
        )
        await page.close()
        totals["skipped"] += 1
        return

    await ensure_location(page, city)
    await handle_interruptions(page)
    widget_ready = await wait_for_widget(page)
    if not widget_ready:
        errors.append(
            ErrorRow(
                url=checkout_url,
                stage="checkout",
                message="Widget probe timeout",
                selector_used="widget",
                screenshot_path=await capture_screenshot(page, service_id, "widget-timeout"),
                ts=utc_now(),
            )
        )
        await page.close()
        totals["skipped"] += 1
        return

    service_name = await page.title()
    category = candidate.get("source", ["unknown"])
    landing_shot = await capture_screenshot(page, service_id, "landing")

    combo_rows, control_map = await enumerate_combinations(
        page=page,
        service_id=service_id,
        city=city,
        checkout_url=checkout_url,
        service_name=service_name,
    )

    addons_rows = await scrape_addons(page, service_id, city)
    availability_rows = await scrape_availability(page, service_id, city)
    offer_rows = await scrape_offers(page, service_id, city)

    combinations.extend(combo_rows)
    addons.extend(addons_rows)
    availability.extend(availability_rows)
    offers.extend(offer_rows)

    totals["processed"] += 1
    totals["combos"] += len(combo_rows)
    totals["addons"] += len(addons_rows)
    totals["availability"] += len(availability_rows)
    totals["offers"] += len(offer_rows)

    price_values = [row.total_aed for row in combo_rows if row.total_aed is not None]
    summary = ServiceSummary(
        service_id=service_id,
        service_name=service_name or url.split("/")[-1].replace("-", " "),
        category=",".join(category) if isinstance(category, list) else str(category),
        city=city,
        service_url=url,
        checkout_url=checkout_url,
        pricing_model=",".join(control_map.keys()) or "UNKNOWN",
        min_price_aed=min(price_values) if price_values else None,
        max_price_aed=max(price_values) if price_values else None,
        avg_price_aed=float(np.mean(price_values)) if price_values else None,
        combos_count=len(combo_rows),
        addons_count=len(addons_rows),
        availability_days=len({row.date for row in availability_rows}),
        has_offers=bool(offer_rows),
        extracted_at=utc_now(),
    )
    services.append(summary)

    await page.close()


###############################################################################
# Output                                                                       #
###############################################################################


def dataframe(records: Iterable[Any], columns: List[str]) -> pd.DataFrame:
    rows = [asdict(record) for record in records]
    if not rows:
        return pd.DataFrame(columns=columns)
    return pd.DataFrame(rows, columns=columns)


def autosize_columns(ws) -> None:
    for column_cells in ws.columns:
        length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = min(max(length + 2, 12), 42)


def style_sheet(ws) -> None:
    from openpyxl.styles import Font, PatternFill

    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = PatternFill(start_color="FFEEF3", end_color="FFEEF3", fill_type="solid")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    autosize_columns(ws)


def build_excel(
    services: List[ServiceSummary],
    combinations: List[CombinationRow],
    addons: List[AddonRow],
    availability: List[AvailabilityRow],
    offers: List[OfferRow],
    errors: List[ErrorRow],
    meta: MetaRow,
) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.formatting.rule import ColorScaleRule

    wb = Workbook()
    wb.remove(wb.active)

    sheets = {
        "Services": dataframe(
            services,
            [
                "service_id",
                "service_name",
                "category",
                "city",
                "service_url",
                "checkout_url",
                "pricing_model",
                "min_price_aed",
                "max_price_aed",
                "avg_price_aed",
                "combos_count",
                "addons_count",
                "availability_days",
                "has_offers",
                "extracted_at",
            ],
        ),
        "Combinations": dataframe(
            combinations,
            [
                "service_id",
                "city",
                "combination_key",
                "dimensions_json",
                "base_price_aed",
                "fees_aed",
                "vat_aed",
                "total_aed",
                "currency",
                "discount_label",
                "extraction_path",
                "screenshot_path",
                "extracted_at",
            ],
        ),
        "AddOns": dataframe(
            addons,
            [
                "service_id",
                "city",
                "addon_name",
                "addon_description",
                "addon_price_aed",
                "is_required",
                "is_recommended",
                "appears_with_combination_key",
                "extracted_at",
            ],
        ),
        "Availability": dataframe(
            availability,
            [
                "service_id",
                "city",
                "date",
                "slot_label",
                "status",
                "scraped_from_combination_key",
                "extracted_at",
            ],
        ),
        "Offers": dataframe(
            offers,
            [
                "service_id",
                "city",
                "offer_badge",
                "offer_text",
                "offer_terms",
                "applicable_to_combination_key",
                "extracted_at",
            ],
        ),
        "Errors": dataframe(
            errors,
            [
                "url",
                "stage",
                "message",
                "selector_used",
                "screenshot_path",
                "ts",
            ],
        ),
        "Meta": pd.DataFrame([asdict(meta)]),
    }

    for name, df in sheets.items():
        ws = wb.create_sheet(name)
        if df.empty:
            ws.append(df.columns.tolist())
        else:
            ws.append(df.columns.tolist())
            for row in df.itertuples(index=False):
                ws.append(list(row))
        style_sheet(ws)
        if name == "Combinations" and ws.max_row > 1:
            color_rule = ColorScaleRule(start_type="min", start_color="FFF5F5", mid_type="percentile", mid_value=50, mid_color="FFF8CB", end_type="max", end_color="FF65B2")
            ws.conditional_formatting.add(f"H2:H{ws.max_row}", color_rule)

    wb.save(EXCEL_PATH)


###############################################################################
# NOTES                                                                        #
###############################################################################


def write_notes(discovered: List[Dict[str, Any]], totals: Dict[str, int], errors: List[ErrorRow]) -> None:
    lines = ["Justlife scrape run notes", f"Timestamp: {utc_now()}"]
    summary = {
        "discovered": len(discovered),
        **totals,
    }
    for key, value in summary.items():
        lines.append(f"{key}: {value}")
    if errors:
        lines.append("Errors observed:")
        for err in errors[:10]:
            lines.append(f"- [{err.stage}] {err.url}: {err.message}")
    NOTES_PATH.write_text("\n".join(lines), encoding="utf-8")


###############################################################################
# Main runner                                                                  #
###############################################################################


async def async_main() -> None:
    start_time = time.time()
    totals = defaultdict(int)

    async with async_playwright() as p:
        browser: Browser = await p.chromium.launch(headless=HEADLESS)
        context = await browser.new_context(
            user_agent=USER_AGENT,
            locale="en-AE",
            viewport={"width": 1440, "height": 900},
        )
        context.set_default_timeout(15000)

        discovered = await discover_services(context)
        totals["discovered"] = len(discovered)

        services: List[ServiceSummary] = []
        combos: List[CombinationRow] = []
        addons: List[AddonRow] = []
        availability: List[AvailabilityRow] = []
        offers: List[OfferRow] = []
        errors: List[ErrorRow] = []

        for idx, candidate in enumerate(discovered):
            try:
                await process_service(
                    context=context,
                    candidate=candidate,
                    service_idx=idx,
                    totals=totals,
                    services=services,
                    combinations=combos,
                    addons=addons,
                    availability=availability,
                    offers=offers,
                    errors=errors,
                )
            except KeyboardInterrupt:
                errors.append(
                    ErrorRow(
                        url=candidate["url"],
                        stage="interrupt",
                        message="Keyboard interrupt",
                        selector_used="run",
                        screenshot_path=None,
                        ts=utc_now(),
                    )
                )
                break
            except Exception as exc:
                errors.append(
                    ErrorRow(
                        url=candidate["url"],
                        stage="processing",
                        message=str(exc),
                        selector_used="process_service",
                        screenshot_path=None,
                        ts=utc_now(),
                    )
                )
                totals["skipped"] += 1

        await browser.close()

    runtime_minutes = (time.time() - start_time) / 60.0
    meta = MetaRow(
        total_discovered=totals.get("discovered", 0),
        total_bookable=totals.get("processed", 0),
        processed=totals.get("processed", 0),
        skipped=totals.get("skipped", 0),
        combos=len(combos),
        addons=len(addons),
        availability=len(availability),
        offers=len(offers),
        errors=len(errors),
        runtime_minutes=runtime_minutes,
        timestamp=utc_now(),
    )

    build_excel(services, combos, addons, availability, offers, errors, meta)
    write_notes(discovered, totals, errors)

    print("Run summary\n============")
    print(f"Discovered URLs: {totals.get('discovered', 0)}")
    print(f"Processed: {totals.get('processed', 0)}  Skipped: {totals.get('skipped', 0)}")
    print(f"Combinations: {len(combos)}  Add-ons: {len(addons)}  Availability: {len(availability)}  Offers: {len(offers)}")
    print(f"Workbook: {EXCEL_PATH}")
    print(f"Notes: {NOTES_PATH}")
    print(f"Screenshots: {SHOTS_DIR}")


def main() -> None:
    try:
        asyncio.run(async_main())
    except RuntimeError:
        loop = asyncio.get_event_loop()
        loop.run_until_complete(async_main())


if __name__ == "__main__":
    main()
