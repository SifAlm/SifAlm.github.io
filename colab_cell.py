import importlib
import json
import os
import random
import re
import subprocess
import sys
from collections import defaultdict
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import urljoin


def ensure_dependencies() -> None:
    """Install all runtime dependencies required by the scraper."""

    packages = {
        "playwright": "playwright",
        "pandas": "pandas",
        "openpyxl": "openpyxl",
        "bs4": "beautifulsoup4",
        "lxml": "lxml",
        "tqdm": "tqdm",
        "fake_useragent": "fake_useragent",
        "nest_asyncio": "nest_asyncio",
        "requests": "requests",
        "PIL": "pillow",
    }
    missing: List[str] = []
    for module, package in packages.items():
        try:
            importlib.import_module(module)
        except ImportError:
            missing.append(package)
    if missing:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", *missing])
    subprocess.check_call([sys.executable, "-m", "playwright", "install", "--with-deps", "chromium"])


ensure_dependencies()

import asyncio
import contextlib
import io
import time

import nest_asyncio
import pandas as pd
from fake_useragent import UserAgent
from PIL import Image, ImageDraw
from playwright.async_api import TimeoutError as PlaywrightTimeout, async_playwright
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

nest_asyncio.apply()


BASE_URL = "https://www.justlife.com/en-AE"
OUTPUT_DIR = Path("justlife_intel")
OUTPUT_DIR.mkdir(exist_ok=True)
SHOTS_DIR = OUTPUT_DIR / "shots"
SHOTS_DIR.mkdir(exist_ok=True)

SERVICE_URL_TARGET = 90
DEFAULT_HEADERS = {
    "accept-language": "en-AE,en;q=0.9",
}

CTA_TEXT_PATTERNS = [
    "book",
    "schedule",
    "continue",
    "checkout",
    "select date",
    "select time",
    "get started",
    "next",
]

BOOKING_WIDGET_PROBES = [
    "select",
    "[data-testid*='hour']",
    "[data-testid*='duration']",
    "[data-testid*='professional']",
    "[data-testid*='bed']",
    "[data-testid*='bath']",
    "[data-testid*='package']",
    "[data-testid*='quantity']",
    "[data-testid*='area']",
    "[data-testid*='slot']",
    "[role='radiogroup']",
    "[role='combobox']",
    "button[aria-pressed]",
    "input[type='number']",
    "input[type='radio']",
]

PRICE_SELECTORS = [
    "[data-testid*='total']",
    "[data-testid*='price']",
    "[data-testid*='grand']",
    "[class*='total']",
    "[class*='price']",
    "[class*='grand']",
    "[class*='summary'] [class*='amount']",
]
SUMMARY_SELECTORS = [
    "[data-testid*='summary']",
    "[class*='summary']",
    "[class*='fee']",
    "[class*='breakdown']",
]
CONTROL_LABEL_HINTS = {
    "HOURS_PROS": ["hour", "duration", "cleaner", "professional", "maid"],
    "ROOM_MATERIAL": ["bedroom", "bathroom", "kitchen", "room", "material"],
    "UNIT_QUANTITY": ["unit", "item", "piece", "appliance", "sofa"],
    "AREA_BASED": ["sq", "area", "meter", "m²", "sqft"],
    "FLAT_PACKAGE": ["package", "plan", "bundle", "basic", "premium"],
}
MODEL_SELECTORS = {
    "HOURS_PROS": [
        "[data-testid*='hour']",
        "[data-testid*='duration']",
        "[data-testid*='professional']",
        "[aria-label*='hour']",
        "[aria-label*='professional']",
    ],
    "ROOM_MATERIAL": [
        "[data-testid*='bed']",
        "[data-testid*='bath']",
        "[aria-label*='bed']",
        "[aria-label*='bath']",
    ],
    "UNIT_QUANTITY": [
        "[data-testid*='unit']",
        "[data-testid*='piece']",
        "[data-testid*='item']",
        "[aria-label*='unit']",
    ],
    "AREA_BASED": [
        "[data-testid*='area']",
        "[data-testid*='sq']",
        "[aria-label*='area']",
    ],
    "FLAT_PACKAGE": [
        "[data-testid*='package']",
        "[class*='package']",
        "[class*='plan']",
    ],
}
SERVICE_KEYWORDS = [
    "clean",
    "salon",
    "spa",
    "pest",
    "ac",
    "sofa",
    "mattress",
    "curtain",
    "carpet",
    "laundry",
    "maid",
    "deep",
    "move",
    "villa",
    "apartment",
    "wax",
    "brow",
    "lash",
    "nail",
    "physio",
    "nurse",
    "lab",
    "doctor",
    "pet",
    "groom",
]
EXCLUDED_PATTERNS = [
    r"/(dubai|abu-dhabi|sharjah|ajman)(?:$|/)",
    r"/my-account/",
    r"/checkout/flex",
    r"/faq",
    r"/privacy",
    r"/terms",
    r"/policy",
    r"/sitemap",
    r"/blog",
]
AVAILABILITY_TRIGGER_TEXT = ["schedule", "book", "time", "slot", "select date"]


@dataclass
class Config:
    base_url: str = os.environ.get("BASE", BASE_URL)
    market: str = os.environ.get("MARKET", "dubai")
    days: int = int(os.environ.get("DAYS", "14"))
    max_combos: int = int(os.environ.get("MAX_COMBOS", "60"))
    headless: bool = os.environ.get("HEADLESS", "1") == "1"
    throttle_seconds: float = 1.1
    jitter_seconds: float = 0.6
    retries: int = 2


@dataclass
class PageRecord:
    page_id: str
    parent_page_id: Optional[str]
    level: str
    url: str
    page_title_text: str
    detected_model_type: str
    has_addons: bool
    has_slots: bool
    notes: str


@dataclass
class OptionRecord:
    page_id: str
    control_group: str
    option_code: str
    option_label: str
    option_type: str
    is_default: bool
    sort_index: int


@dataclass
class PriceRecord:
    page_id: str
    service_name: str
    model_type: str
    hours: Optional[float] = None
    pros: Optional[int] = None
    bedrooms: Optional[int] = None
    bathrooms: Optional[int] = None
    kitchen_package: Optional[str] = None
    appliance_combo: Optional[str] = None
    package_name: Optional[str] = None
    duration: Optional[str] = None
    offer_state: str = "OFF"
    offer_name: Optional[str] = None
    discount_type: Optional[str] = None
    discount_value: Optional[float] = None
    base_price: Optional[float] = None
    fees_total: Optional[float] = None
    vat_amount: Optional[float] = None
    vat_percent: Optional[float] = None
    grand_total: Optional[float] = None
    currency: Optional[str] = None
    pricing_source_url: str = ""
    collected_at: str = ""
    qa_flags: str = ""
    screenshot_ref: Optional[str] = None


@dataclass
class AddonRecord:
    page_id: str
    addon_name: str
    addon_description: str
    addon_price: Optional[float]
    currency: Optional[str]
    is_recommended: bool
    is_required: bool
    seen_on_step: str
    pricing_source_url: str
    collected_at: str


@dataclass
class AvailabilityRecord:
    page_id: str
    date_str: str
    time_slot_label: str
    slot_status: str
    lead_time_days: int
    collected_at: str


@dataclass
class OfferRecord:
    page_id: str
    offer_name: str
    offer_badge_text: str
    description: str
    eligibility: str
    discount_type: Optional[str]
    discount_value: Optional[float]
    stacking_rules: str
    collected_at: str


@dataclass
class ElementRecord:
    page_id: str
    selector: str
    element_role: str
    raw_text: str
    normalized_text: str


class RunState:
    def __init__(self, config: Config) -> None:
        self.config = config
        self.pages: List[PageRecord] = []
        self.options: List[OptionRecord] = []
        self.prices: List[PriceRecord] = []
        self.addons: List[AddonRecord] = []
        self.availability: List[AvailabilityRecord] = []
        self.offers: List[OfferRecord] = []
        self.elements: List[ElementRecord] = []
        self.notes: List[str] = []
        self.errors: List[str] = []
        self.screenshots: Dict[str, str] = {}
        self.summary_counts = {
            "services": 0,
            "combos": 0,
            "addons": 0,
            "slots": 0,
            "offers": 0,
        }
        self.discovery_sources: Dict[str, int] = defaultdict(int)
        self.discovery_totals: Dict[str, int] = defaultdict(int)

    def log_note(self, message: str) -> None:
        print(message)
        self.notes.append(message)

    def log_error(self, message: str) -> None:
        print(f"[ERROR] {message}")
        self.errors.append(message)


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")


def slugify(text: str) -> str:
    base = re.sub(r"[^a-zA-Z0-9]+", "-", text.strip().lower()).strip("-")
    return base or "page"


def normalize_text(value: Optional[str]) -> str:
    if not value:
        return ""
    return re.sub(r"\s+", " ", value).strip()


def xpath_literal(text: str) -> str:
    if "'" not in text:
        return f"'{text}'"
    if '"' not in text:
        return f'"{text}"'
    parts = text.split("'")
    pieces: List[str] = []
    if parts[0]:
        pieces.append(f"'{parts[0]}'")
    for part in parts[1:]:
        pieces.append("\"'\"")
        if part:
            pieces.append(f"'{part}'")
    if not pieces:
        pieces.append("\"'\"")
    return "concat(" + ", ".join(pieces) + ")"


def text_locator(keyword: str) -> str:
    lowered = keyword.lower()
    return (
        "xpath=//*[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), "
        f"'{lowered}')]"
    )


def canonicalize_url(url: str) -> str:
    url = re.sub(r"(en-AE/)+", "en-AE/", url)
    url = url.split("#")[0]
    url = url.split("?")[0]
    if url.endswith("//"):
        url = url[:-1]
    if url.endswith("/"):
        url = url[:-1]
    return url


async def random_sleep(cfg: Config) -> None:
    base = cfg.throttle_seconds
    jitter = random.random() * cfg.jitter_seconds
    await asyncio.sleep(base + jitter)


async def take_screenshot(page, path: Path, highlight_rects: Optional[List[Dict[str, float]]] = None) -> None:
    raw = await page.screenshot(full_page=True)
    image = Image.open(io.BytesIO(raw)).convert("RGB")
    if highlight_rects:
        draw = ImageDraw.Draw(image)
        for rect in highlight_rects:
            x = rect.get("x", 0)
            y = rect.get("y", 0)
            width = rect.get("width", 0)
            height = rect.get("height", 0)
            draw.rectangle([x, y, x + width, y + height], outline="red", width=5)
    image.save(path)


async def gather_highlight_rects(page, selectors: Iterable[str]) -> List[Dict[str, float]]:
    rects: List[Dict[str, float]] = []
    for selector in selectors:
        try:
            elements = await page.query_selector_all(selector)
        except PlaywrightTimeout:
            continue
        for el in elements:
            with contextlib.suppress(Exception):
                box = await el.bounding_box()
                if box:
                    rects.append(box)
    return rects


async def extract_links(page) -> List[Tuple[str, str]]:
    raw_links = await page.evaluate(
        """
        () => {
            const anchors = Array.from(document.querySelectorAll('a[href*="/en-AE/"]'));
            return anchors.map(a => ({
                href: a.href,
                text: (a.textContent || '').trim(),
                inFooter: !!a.closest('footer'),
                inHeader: !!a.closest('header'),
            }));
        }
        """
    )
    results: List[Tuple[str, str]] = []
    for entry in raw_links:
        href = entry.get("href") if isinstance(entry, dict) else None
        if not isinstance(href, str):
            continue
        source = "body"
        if isinstance(entry, dict):
            if entry.get("inFooter"):
                source = "footer"
            elif entry.get("inHeader"):
                source = "header"
            elif isinstance(entry.get("text"), str) and "see all" in entry["text"].lower():
                source = "see_all"
        results.append((href, source))
    return results


def filter_service_links(links: Iterable[str]) -> List[str]:
    filtered: List[str] = []
    for href in links:
        if not href:
            continue
        url = canonicalize_url(href)
        low = url.lower()
        if not low.startswith("https://www.justlife.com/en-ae"):
            continue
        if any(re.search(pattern, low) for pattern in EXCLUDED_PATTERNS):
            continue
        slug = low.rsplit("/", 1)[-1]
        if not any(keyword in slug for keyword in SERVICE_KEYWORDS):
            continue
        filtered.append(url)
    return sorted(set(filtered))


async def ensure_market(page, cfg: Config, state: RunState) -> None:
    try:
        location_chip = await page.query_selector("[data-testid*='location']")
        if location_chip:
            text = (await location_chip.inner_text()).lower()
            if cfg.market.lower() not in text:
                await location_chip.click()
                await asyncio.sleep(1)
                input_box = await page.wait_for_selector("input[placeholder*='Search']", timeout=5000)
                await input_box.fill(cfg.market)
                await asyncio.sleep(0.5)
                await page.keyboard.press("Enter")
                await asyncio.sleep(2)
    except PlaywrightTimeout:
        state.log_note("Location chip not found; assuming default Dubai context")


async def has_booking_widget(page) -> bool:
    for selector in BOOKING_WIDGET_PROBES + PRICE_SELECTORS:
        try:
            handle = await page.query_selector(selector)
        except Exception:
            handle = None
        if handle:
            return True
    return False


async def reach_booking_widget(page, url: str, cfg: Config, state: RunState, page_id: str) -> bool:
    if await has_booking_widget(page):
        return True
    original_url = page.url
    anchors = page.locator("a[href*='/checkout/']")
    count = await anchors.count()
    for idx in range(min(count, 5)):
        href = await anchors.nth(idx).get_attribute("href")
        if not href:
            continue
        target = urljoin(original_url, href)
        try:
            await page.goto(target, wait_until="networkidle", timeout=60000)
            await random_sleep(cfg)
            await ensure_market(page, cfg, state)
        except Exception as exc:
            state.log_error(f"Checkout navigation failed for {page_id}: {exc}")
            await page.goto(original_url, wait_until="networkidle", timeout=60000)
            continue
        if await has_booking_widget(page):
            state.discovery_totals["checkout"] += 1
            state.log_note(f"Reached checkout via href for {page_id}: {target}")
            return True
        await page.goto(original_url, wait_until="networkidle", timeout=60000)
    LOWER = "translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')"
    for pattern in CTA_TEXT_PATTERNS:
        locator = page.locator(
            "xpath="
            "//a[contains(" + LOWER + ", " + xpath_literal(pattern) + ")] | "
            "//button[contains(" + LOWER + ", " + xpath_literal(pattern) + ")] | "
            "//*[@role='button'][contains(" + LOWER + ", " + xpath_literal(pattern) + ")]"
        )
        attempts = min(await locator.count(), 3)
        for idx in range(attempts):
            try:
                button = locator.nth(idx)
                await button.scroll_into_view_if_needed()
                with contextlib.suppress(Exception):
                    await button.click()
                with contextlib.suppress(Exception):
                    await page.wait_for_load_state("networkidle", timeout=60000)
                await random_sleep(cfg)
                if await has_booking_widget(page):
                    state.log_note(f"Reached widget via CTA '{pattern}' for {page_id}")
                    return True
            except Exception as exc:
                state.log_error(f"CTA click failed for {page_id}: {exc}")
            finally:
                if not await has_booking_widget(page):
                    with contextlib.suppress(Exception):
                        await page.goto(original_url, wait_until="networkidle", timeout=60000)
                        await random_sleep(cfg)
                        await ensure_market(page, cfg, state)
    return await has_booking_widget(page)


async def detect_model(page, state: RunState, page_id: str) -> Tuple[str, Dict[str, Any]]:
    model_scores: Dict[str, int] = defaultdict(int)
    evidence: Dict[str, List[Dict[str, str]]] = defaultdict(list)
    for model, selectors in MODEL_SELECTORS.items():
        for selector in selectors:
            try:
                handles = await page.query_selector_all(selector)
            except Exception:
                handles = []
            if not handles:
                continue
            for handle in handles[:3]:
                with contextlib.suppress(Exception):
                    text = normalize_text(await handle.inner_text())
                text = text or selector
                evidence[model].append({"selector": selector, "text": text})
                state.elements.append(
                    ElementRecord(page_id, selector, "control-candidate", text, text.lower())
                )
            model_scores[model] += len(handles)
        if evidence[model]:
            continue
        for hint in CONTROL_LABEL_HINTS.get(model, []):
            locator = page.locator(text_locator(hint))
            count = await locator.count()
            if not count:
                continue
            for idx in range(min(count, 2)):
                handle = locator.nth(idx)
                with contextlib.suppress(Exception):
                    text = normalize_text(await handle.inner_text())
                if not text:
                    continue
                selector_hint = f"text~{hint}"
                evidence[model].append({"selector": selector_hint, "text": text})
                state.elements.append(
                    ElementRecord(page_id, selector_hint, "control-label", text, text.lower())
                )
            if evidence[model]:
                model_scores[model] += len(evidence[model])
                break
    if not model_scores:
        state.elements.append(
            ElementRecord(page_id, "*", "model-detection", "no-hints", "no-hints")
        )
        return "OTHER", {}
    best_model = max(model_scores, key=model_scores.get)
    return best_model, {best_model: evidence[best_model]}


async def get_service_title(page) -> str:
    for selector in ["h1", "h2", "[data-testid*='title']"]:
        try:
            el = await page.query_selector(selector)
            if el:
                title = await el.inner_text()
                title = re.sub(r"\s+", " ", title).strip()
                if title:
                    return title
        except PlaywrightTimeout:
            continue
    return "Justlife Service"


async def wait_for_price_change(page, previous_text: str) -> str:
    checks = 0
    while checks < 16:
        for selector in PRICE_SELECTORS:
            try:
                el = await page.query_selector(selector)
            except Exception:
                el = None
            if el:
                try:
                    current = await el.inner_text()
                except Exception:
                    current = None
                if current and current != previous_text:
                    return current
        checks += 1
        await asyncio.sleep(0.5)
    return previous_text


def parse_price_text(text: str) -> Tuple[Optional[str], Optional[float]]:
    if not text:
        return None, None
    text = text.replace("\n", " ")
    currency_match = re.search(r"AED|د\.إ", text, re.I)
    number_match = re.search(r"([0-9]+(?:,[0-9]{3})*(?:\.[0-9]+)?)", text.replace(",", ""))
    currency = "AED" if currency_match else None
    amount = float(number_match.group(1)) if number_match else None
    return currency, amount


async def collect_summary_text(page) -> List[str]:
    snippets: List[str] = []
    for selector in SUMMARY_SELECTORS:
        try:
            nodes = await page.query_selector_all(selector)
        except Exception:
            nodes = []
        for node in nodes:
            try:
                text = normalize_text(await node.inner_text())
            except Exception:
                text = ""
            if text:
                snippets.append(text)
    return snippets


def parse_breakdown(snippets: List[str]) -> Dict[str, Optional[float]]:
    joined = " | ".join(snippets).lower()
    def find_amount(label: str) -> Optional[float]:
        match = re.search(rf"(?:{label})[^0-9]*([0-9]+(?:\.[0-9]+)?)", joined, re.I)
        return float(match.group(1)) if match else None

    def find_percent(label: str) -> Optional[float]:
        match = re.search(rf"(?:{label})[^0-9%]*([0-9]+(?:\.[0-9]+)?)%", joined, re.I)
        return float(match.group(1)) if match else None

    return {
        "base_price": find_amount("base|subtotal|service"),
        "fees_total": find_amount("fee"),
        "vat_amount": find_amount("vat"),
        "vat_percent": find_percent("vat"),
    }


async def enumerate_options(page, model_type: str, state: RunState, page_id: str) -> Tuple[List[Dict[str, Any]], Dict[str, List[Dict[str, Any]]]]:
    option_groups: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    hints = CONTROL_LABEL_HINTS.get(model_type, [])
    if not hints:
        flat_hints: List[str] = []
        for entries in CONTROL_LABEL_HINTS.values():
            flat_hints.extend(entries)
        hints = sorted(set(flat_hints))
    interactive_handles = await page.query_selector_all("[data-testid]")
    for handle in interactive_handles:
        try:
            data_testid = (await handle.get_attribute("data-testid") or "").lower()
        except Exception:
            data_testid = ""
        try:
            raw_text = normalize_text(await handle.inner_text())
        except Exception:
            raw_text = ""
        combined = f"{data_testid} {raw_text.lower()}" if raw_text else data_testid
        if not combined.strip():
            continue
        for hint in hints:
            if hint in combined:
                group_label = hint.replace("-", " ").replace("_", " ").title()
                try:
                    value = await handle.get_attribute("value") or await handle.get_attribute("data-value")
                except Exception:
                    value = None
                value = value or raw_text or data_testid
                try:
                    tag = (await handle.evaluate("el => el.tagName") or "").lower()
                except Exception:
                    tag = ""
                is_default = False
                for attr_name in ("aria-pressed", "aria-selected", "aria-checked"):
                    try:
                        attr_val = await handle.get_attribute(attr_name)
                    except Exception:
                        attr_val = None
                    if attr_val and attr_val.lower() == "true":
                        is_default = True
                option = {
                    "label": raw_text or value,
                    "value": value,
                    "selector": data_testid,
                    "option_type": "button" if tag == "button" else tag or "option",
                    "is_default": is_default,
                }
                if option["label"] and not any(existing["label"] == option["label"] for existing in option_groups[group_label]):
                    option_groups[group_label].append(option)
                    state.options.append(
                        OptionRecord(
                            page_id=page_id,
                            control_group=group_label,
                            option_code=str(option["value"]),
                            option_label=option["label"],
                            option_type=option["option_type"],
                            is_default=is_default,
                            sort_index=len(option_groups[group_label]) - 1,
                        )
                    )
                break
    select_handles = await page.query_selector_all("select")
    for select in select_handles:
        try:
            name_attr = (await select.get_attribute("name") or "").lower()
        except Exception:
            name_attr = ""
        try:
            text_block = normalize_text(await select.inner_text())
        except Exception:
            text_block = ""
        combined = f"{name_attr} {text_block.lower()}"
        for hint in hints:
            if hint in combined:
                group_label = hint.replace("-", " ").replace("_", " ").title()
                options = await select.query_selector_all("option")
                for idx, opt in enumerate(options):
                    try:
                        option_label_text = normalize_text(await opt.inner_text())
                    except Exception:
                        option_label_text = ""
                    try:
                        option_value_text = normalize_text(await opt.get_attribute("value"))
                    except Exception:
                        option_value_text = ""
                    label = option_label_text or option_value_text
                    if not label:
                        continue
                    try:
                        value = await opt.get_attribute("value")
                    except Exception:
                        value = None
                    value = value or label
                    try:
                        selected_attr = await opt.get_attribute("selected")
                    except Exception:
                        selected_attr = None
                    is_default = (selected_attr or "").lower() in ("true", "selected")
                    option = {
                        "label": label,
                        "value": value,
                        "selector": name_attr,
                        "option_type": "select",
                        "is_default": is_default,
                    }
                    if not any(existing["label"] == option["label"] for existing in option_groups[group_label]):
                        option_groups[group_label].append(option)
                        state.options.append(
                            OptionRecord(
                                page_id=page_id,
                                control_group=group_label,
                                option_code=str(value),
                                option_label=label,
                                option_type="select",
                                is_default=is_default,
                                sort_index=len(option_groups[group_label]) - 1,
                            )
                        )
                break
    if not option_groups:
        state.log_note(f"No option groups detected for {page_id}")
        state.elements.append(
            ElementRecord(page_id, "*", "option-discovery", "no-groups", "no-groups")
        )
    combos: List[Dict[str, Any]] = []
    keys = list(option_groups.keys())
    if not keys:
        return combos, option_groups

    def backtrack(index: int, current: Dict[str, Any]) -> None:
        if len(combos) >= state.config.max_combos:
            return
        if index >= len(keys):
            combos.append(current.copy())
            return
        group = keys[index]
        choices = option_groups[group]
        if not choices:
            backtrack(index + 1, current)
            return
        subset: List[Dict[str, Any]] = []
        subset.append(choices[0])
        if len(choices) > 1:
            subset.append(choices[-1])
        mid_index = len(choices) // 2
        if len(choices) > 2 and choices[mid_index] not in subset:
            subset.append(choices[mid_index])
        for idx in (1, 2):
            if idx < len(choices) and choices[idx] not in subset:
                subset.append(choices[idx])
        subset = subset[:5]
        for choice in subset:
            current[group] = choice
            backtrack(index + 1, current)
        current.pop(group, None)

    backtrack(0, {})
    return combos, option_groups


async def capture_price(page, combos: List[Dict[str, Any]], model_type: str, state: RunState, page_id: str, service_name: str, evidence: Dict[str, Any]) -> None:
    if not combos:
        text, selector = await get_price_text(page)
        currency, amount = parse_price_text(text)
        summary_snippets = await collect_summary_text(page)
        breakdown = parse_breakdown(summary_snippets)
        base_price = breakdown["base_price"] or amount
        fees_total = breakdown["fees_total"] or 0.0
        vat_amount = breakdown["vat_amount"] or 0.0
        vat_percent = breakdown["vat_percent"]
        qa_flags = ["NO_COMBOS"]
        if breakdown["fees_total"] is None and amount is not None:
            qa_flags.append("FEE_UNKNOWN")
        record = PriceRecord(
            page_id=page_id,
            service_name=service_name,
            model_type=model_type,
            base_price=base_price,
            fees_total=fees_total,
            vat_amount=vat_amount,
            vat_percent=vat_percent,
            grand_total=amount,
            currency=currency,
            pricing_source_url=page.url,
            collected_at=utc_now(),
            qa_flags=",".join(qa_flags),
            screenshot_ref=state.screenshots.get(page_id),
        )
        if selector and text:
            state.elements.append(
                ElementRecord(page_id, selector, "price-total", text, normalize_text(text))
            )
        state.prices.append(record)
        state.summary_counts["combos"] += 1
        return
    for combo in combos:
        previous_text, _ = await get_price_text(page)
        for group_label, choice in combo.items():
            await apply_option(page, group_label, choice)
            await asyncio.sleep(1)
        latest = await wait_for_price_change(page, previous_text)
        current_text, selector = await get_price_text(page)
        text_to_use = current_text or latest
        currency, amount = parse_price_text(text_to_use)
        summary_snippets = await collect_summary_text(page)
        breakdown = parse_breakdown(summary_snippets)
        base_price = breakdown["base_price"] or amount
        fees_total = breakdown["fees_total"] or 0.0
        vat_amount = breakdown["vat_amount"] or 0.0
        vat_percent = breakdown["vat_percent"]
        qa_flags = []
        if amount is None:
            qa_flags.append("NO_PRICE")
        if breakdown["fees_total"] is None and amount is not None:
            qa_flags.append("FEE_UNKNOWN")
        record = PriceRecord(
            page_id=page_id,
            service_name=service_name,
            model_type=model_type,
            base_price=base_price,
            fees_total=fees_total,
            vat_amount=vat_amount,
            vat_percent=vat_percent,
            grand_total=amount,
            currency=currency,
            pricing_source_url=page.url,
            collected_at=utc_now(),
            qa_flags=",".join(qa_flags),
            screenshot_ref=state.screenshots.get(page_id),
        )
        for group_label, choice in combo.items():
            normalized = group_label.lower()
            value = choice["label"]
            if "hour" in normalized:
                record.hours = float(re.findall(r"\d+", value)[0]) if re.findall(r"\d+", value) else None
            elif "professional" in normalized or "cleaner" in normalized:
                record.pros = int(re.findall(r"\d+", value)[0]) if re.findall(r"\d+", value) else None
            elif "bed" in normalized:
                record.bedrooms = int(re.findall(r"\d+", value)[0]) if re.findall(r"\d+", value) else None
            elif "bath" in normalized:
                record.bathrooms = int(re.findall(r"\d+", value)[0]) if re.findall(r"\d+", value) else None
            elif "package" in normalized:
                record.package_name = value
            elif "unit" in normalized or "item" in normalized or "piece" in normalized:
                record.appliance_combo = value
        if amount is not None and base_price is not None:
            computed = (base_price or 0.0) + (fees_total or 0.0) + (vat_amount or 0.0)
            if abs(computed - amount) > max(1.0, amount) * 0.02:
                record.qa_flags = ",".join(filter(None, [record.qa_flags, "MATH_MISMATCH"]))
        state.prices.append(record)
        state.summary_counts["combos"] += 1
        if selector and text_to_use:
            state.elements.append(
                ElementRecord(page_id, selector, "price-total", text_to_use, normalize_text(text_to_use))
            )
        await asyncio.sleep(0.5)


async def apply_option(page, group_label: str, choice: Dict[str, Any]) -> None:
    text = choice.get("label") or ""
    selector_hint = choice.get("selector") or ""
    value = choice.get("value")
    option_type = choice.get("option_type")
    if option_type == "select":
        potential_selectors = []
        if selector_hint:
            potential_selectors.append(page.locator(f"select[name='{selector_hint}']"))
            potential_selectors.append(page.locator(f"select[data-testid='{selector_hint}']"))
        potential_selectors.append(page.locator("select"))
        for locator in potential_selectors:
            try:
                if await locator.count():
                    await locator.first.select_option(str(value))
                    return
            except Exception:
                continue
    candidates = []
    if selector_hint:
        candidates.append(page.locator(f"[data-testid='{selector_hint}']"))
    if text:
        candidates.append(page.get_by_text(text, exact=True))
        candidates.append(page.locator(f"xpath=//*[normalize-space(text())={xpath_literal(text)}]"))
    if value:
        candidates.append(page.locator(f"[value='{value}']"))
    for locator in candidates:
        try:
            if await locator.count():
                await locator.first.click(timeout=4000)
                return
        except Exception:
            continue


async def get_price_text(page) -> Tuple[str, Optional[str]]:
    for selector in PRICE_SELECTORS:
        try:
            el = await page.query_selector(selector)
        except Exception:
            el = None
        if el:
            try:
                text = await el.inner_text()
            except Exception:
                text = None
            if text:
                return text, selector
    return "", None


async def capture_addons(page, page_id: str, state: RunState) -> None:
    addons_locator = page.locator("text=/add-ons|extras/i")
    if await addons_locator.count():
        with contextlib.suppress(Exception):
            await addons_locator.first.click()
            await asyncio.sleep(1)
    candidates = await page.query_selector_all("[data-testid*='addon'], [class*='addon'], [class*='upsell']")
    for candidate in candidates:
        try:
            text = await candidate.inner_text()
        except Exception:
            continue
        normalized = re.sub(r"\s+", " ", text).strip()
        if not normalized:
            continue
        currency, amount = parse_price_text(normalized)
        name = normalized.split("AED")[0].strip()
        record = AddonRecord(
            page_id=page_id,
            addon_name=name,
            addon_description=normalized,
            addon_price=amount,
            currency=currency,
            is_recommended="recommended" in normalized.lower(),
            is_required="required" in normalized.lower(),
            seen_on_step="pre-cart",
            pricing_source_url="",
            collected_at=utc_now(),
        )
        state.addons.append(record)
        state.summary_counts["addons"] += 1


async def capture_offers(page, page_id: str, state: RunState) -> None:
    offer_nodes = await page.query_selector_all("[class*='offer'], [data-testid*='offer'], [class*='discount']")
    for node in offer_nodes:
        text = await node.inner_text()
        normalized = re.sub(r"\s+", " ", text).strip()
        if not normalized:
            continue
        currency, amount = parse_price_text(normalized)
        record = OfferRecord(
            page_id=page_id,
            offer_name=normalized[:80],
            offer_badge_text=normalized[:40],
            description=normalized,
            eligibility="",
            discount_type="percent" if "%" in normalized else ("flat" if amount else None),
            discount_value=amount,
            stacking_rules="",
            collected_at=utc_now(),
        )
        state.offers.append(record)
        state.summary_counts["offers"] += 1


async def capture_availability(page, cfg: Config, page_id: str, state: RunState) -> None:
    trigger = None
    for text in AVAILABILITY_TRIGGER_TEXT:
        locator = page.locator(f"text=/{text}/i")
        if await locator.count():
            trigger = locator.first
            break
    if not trigger:
        return
    with contextlib.suppress(Exception):
        await trigger.click()
        await asyncio.sleep(1)
    for day_offset in range(cfg.days):
        target_date = datetime.now(timezone.utc).date() + timedelta(days=day_offset)
        date_str = target_date.strftime("%Y-%m-%d")
        locator = page.locator(f"text='{target_date.day}'")
        if await locator.count():
            with contextlib.suppress(Exception):
                await locator.nth(0).click()
                await asyncio.sleep(0.5)
        slots = await page.query_selector_all("[data-testid*='slot'], [class*='slot']")
        if not slots:
            continue
        for slot in slots:
            text = await slot.inner_text()
            normalized = re.sub(r"\s+", " ", text).strip()
            status = "available"
            classes = await slot.get_attribute("class") or ""
            if "disable" in classes.lower():
                status = "disabled"
            record = AvailabilityRecord(
                page_id=page_id,
                date_str=date_str,
                time_slot_label=normalized,
                slot_status=status,
                lead_time_days=day_offset,
                collected_at=utc_now(),
            )
            state.availability.append(record)
            state.summary_counts["slots"] += 1


async def process_service(browser_page, url: str, state: RunState) -> None:
    cfg = state.config
    await browser_page.goto(url, wait_until="domcontentloaded", timeout=60000)
    await random_sleep(cfg)
    await ensure_market(browser_page, cfg, state)
    await asyncio.sleep(1)
    page_stub = slugify(url.rsplit("/", 1)[-1])
    widget_present = await reach_booking_widget(browser_page, url, cfg, state, page_stub)
    service_name = await get_service_title(browser_page)
    composite_id = f"{service_name}-{page_stub}" if service_name else page_stub
    page_id = slugify(composite_id)
    screenshot_path = SHOTS_DIR / f"service_{page_id}.png"
    await take_screenshot(browser_page, screenshot_path)
    state.screenshots[page_id] = str(screenshot_path)
    if not widget_present:
        state.discovery_totals["skipped"] += 1
        state.log_note(f"Booking widget missing for {url}")
        state.elements.append(
            ElementRecord(page_id, url, "widget-check", "not-found", "not-found")
        )
        state.pages.append(
            PageRecord(
                page_id=page_id,
                parent_page_id=None,
                level="service",
                url=url,
                page_title_text=service_name or url,
                detected_model_type="OTHER",
                has_addons=False,
                has_slots=False,
                notes="NO_WIDGET",
            )
        )
        return
    state.discovery_totals["bookable"] += 1
    model_type, evidence = await detect_model(browser_page, state, page_id)
    combos, option_groups = await enumerate_options(browser_page, model_type, state, page_id)
    await capture_price(browser_page, combos, model_type, state, page_id, service_name or url, evidence)
    await capture_addons(browser_page, page_id, state)
    await capture_offers(browser_page, page_id, state)
    await capture_availability(browser_page, cfg, page_id, state)
    state.summary_counts["services"] += 1
    state.pages.append(
        PageRecord(
            page_id=page_id,
            parent_page_id=None,
            level="service",
            url=url,
            page_title_text=service_name or url,
            detected_model_type=model_type,
            has_addons=any(rec.page_id == page_id for rec in state.addons),
            has_slots=any(rec.page_id == page_id for rec in state.availability),
            notes=json.dumps(evidence)[:250] if evidence else "",
        )
    )


async def extract_sitemap_links(page) -> List[str]:
    sitemap_url = urljoin(BASE_URL + "/", "sitemap")
    try:
        await page.goto(sitemap_url, wait_until="domcontentloaded", timeout=60000)
    except Exception:
        return []
    entries = await extract_links(page)
    return [href for href, _ in entries]


async def discover_services(context_page, state: RunState) -> List[str]:
    await context_page.goto(state.config.base_url, wait_until="domcontentloaded", timeout=60000)
    await asyncio.sleep(3)
    rects = await gather_highlight_rects(
        context_page,
        [
            "section:has-text('Cleaning')",
            "section:has-text('Salon')",
            "section:has-text('Maintenance')",
            "footer",
        ],
    )
    home_shot = SHOTS_DIR / "home.png"
    await take_screenshot(context_page, home_shot, rects)
    state.screenshots["home"] = str(home_shot)
    link_entries = await extract_links(context_page)
    candidate_links = [href for href, _ in link_entries if href]
    for _, source in link_entries:
        state.discovery_sources[source] += 1
    filtered = filter_service_links(candidate_links)
    if len(filtered) < SERVICE_URL_TARGET:
        sitemap_links = await extract_sitemap_links(context_page)
        if sitemap_links:
            state.discovery_sources["sitemap"] += len(sitemap_links)
            candidate_links.extend(sitemap_links)
            filtered = filter_service_links(candidate_links)
        await context_page.goto(state.config.base_url, wait_until="domcontentloaded", timeout=60000)
    filtered = sorted(set(filtered))
    checkout_targets = [url for url in filtered if "/checkout/" in url]
    state.discovery_totals["discovered"] = len(filtered)
    state.discovery_totals["checkout"] = len(checkout_targets)
    breakdown = ", ".join(f"{src}:{count}" for src, count in sorted(state.discovery_sources.items()))
    state.log_note(f"Discovery breakdown -> {breakdown}")
    state.log_note(f"Discovered {len(filtered)} candidate service URLs")
    if len(filtered) < SERVICE_URL_TARGET:
        state.log_note("Warning: fewer than target service URLs detected; continuing anyway")
    state.log_note(f"Direct checkout targets identified: {len(checkout_targets)}")
    return filtered


def write_csv(path: Path, records: List[Dict[str, Any]]) -> None:
    if not records:
        df = pd.DataFrame()
    else:
        df = pd.DataFrame(records)
    df.to_csv(path, index=False)


def build_excel(state: RunState) -> Path:
    workbook_path = OUTPUT_DIR / "Justlife_CRM.xlsx"
    wb = Workbook()
    def get_sheet(name: str):
        if name in wb.sheetnames:
            return wb[name]
        else:
            return wb.create_sheet(title=name)
    # Summary sheet
    summary_ws = wb.active
    summary_ws.title = "Summary"
    summary_ws.append(["Metric", "Value"])
    summary_ws.append(["Services scraped", state.summary_counts["services"]])
    summary_ws.append(["Price combinations", state.summary_counts["combos"]])
    summary_ws.append(["Add-ons captured", state.summary_counts["addons"]])
    summary_ws.append(["Availability slots", state.summary_counts["slots"]])
    summary_ws.append(["Offers", state.summary_counts["offers"]])
    for cell in summary_ws[1]:
        cell.font = Font(bold=True)
    summary_ws.freeze_panes = "A2"
    # Additional sheets from pandas dataframes
    sheet_specs = {
        "Price Matrix": [PriceRecord, state.prices],
        "Add-ons": [AddonRecord, state.addons],
        "Availability": [AvailabilityRecord, state.availability],
        "Offers": [OfferRecord, state.offers],
        "Pages": [PageRecord, state.pages],
        "Options": [OptionRecord, state.options],
        "QA & Diagnostics": [ElementRecord, state.elements],
    }
    for sheet_name, (_, records) in sheet_specs.items():
        ws = get_sheet(sheet_name)
        df = pd.DataFrame([asdict(r) for r in records]) if records else pd.DataFrame()
        if df.empty:
            ws.append(["No data captured"])
            continue
        ws.append(list(df.columns))
        for row in df.itertuples(index=False):
            ws.append(list(row))
        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = "A2"
        apply_table_style(ws)
        if sheet_name == "Price Matrix" and df.shape[0] > 1:
            apply_price_formatting(ws, df)
    diagnostics_ws = get_sheet("QA & Diagnostics")
    diagnostics_ws.freeze_panes = "A2"
    # README sheet
    readme_ws = get_sheet("README")
    readme_ws.append(["Run timestamp", utc_now()])
    readme_ws.append(["Base URL", state.config.base_url])
    readme_ws.append(["Notes file", str(OUTPUT_DIR / "NOTES.txt")])
    readme_ws.append(["Screenshots", str(SHOTS_DIR)])
    readme_ws.append(["Errors", " | ".join(state.errors) if state.errors else "None"])
    wb.save(workbook_path)
    return workbook_path


def apply_table_style(ws) -> None:
    for column_cells in ws.columns:
        max_length = 0
        column = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            try:
                cell_value = str(cell.value)
            except Exception:
                cell_value = ""
            if cell.row == 1:
                cell.font = Font(bold=True)
                cell.fill = PatternFill("solid", fgColor="F3F4F6")
            max_length = max(max_length, len(cell_value))
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[column].width = min(60, max_length + 2)


def apply_price_formatting(ws, df: pd.DataFrame) -> None:
    if "grand_total" in df.columns:
        col_index = list(df.columns).index("grand_total") + 1
        color_rule = ColorScaleRule(
            start_type="percentile", start_value=10, start_color="F0F9FF",
            mid_type="percentile", mid_value=50, mid_color="80CBC4",
            end_type="percentile", end_value=90, end_color="004D40",
        )
        ws.conditional_formatting.add(f"{get_column_letter(col_index)}2:{get_column_letter(col_index)}{ws.max_row}", color_rule)
    if "qa_flags" in df.columns:
        qa_index = list(df.columns).index("qa_flags") + 1
        ws.conditional_formatting.add(
            f"{get_column_letter(qa_index)}2:{get_column_letter(qa_index)}{ws.max_row}",
            CellIsRule(operator="notEqual", formula=['""'], fill=PatternFill("solid", fgColor="FFCDD2"))
        )


def write_notes(state: RunState) -> Path:
    notes_path = OUTPUT_DIR / "NOTES.txt"
    content = ["Justlife competitive intelligence run notes", "=" * 60, ""]
    content.append(f"Timestamp: {utc_now()}")
    content.append(f"Base URL: {state.config.base_url}")
    content.append("")
    content.append("Model decisions:")
    for page in state.pages:
        content.append(f"- {page.page_title_text} ({page.url}) -> {page.detected_model_type}")
    if state.notes:
        content.append("")
        content.append("Additional notes:")
        content.extend(f"* {note}" for note in state.notes)
    if state.errors:
        content.append("")
        content.append("Errors:")
        content.extend(f"* {err}" for err in state.errors)
    notes_path.write_text("\n".join(content), encoding="utf-8")
    return notes_path


async def run_scraper() -> None:
    cfg = Config()
    ua = UserAgent()
    cfg_user_agent = cfg.__dict__.setdefault("user_agent", ua.chrome)
    state = RunState(cfg)
    playwright = await async_playwright().start()
    browser = await playwright.chromium.launch(headless=cfg.headless)
    context = await browser.new_context(locale="en-AE", user_agent=cfg_user_agent, extra_http_headers=DEFAULT_HEADERS)
    page = await context.new_page()
    try:
        service_urls = await discover_services(page, state)
        processed: set = set()
        for url in service_urls:
            if url in processed:
                continue
            processed.add(url)
            try:
                await process_service(page, url, state)
            except Exception as exc:
                state.log_error(f"Failed to process {url}: {exc}")
                state.discovery_totals["skipped"] += 1
            await random_sleep(cfg)
    finally:
        await context.close()
        await browser.close()
        await playwright.stop()
    # Write outputs
    pages_df = [asdict(record) for record in state.pages]
    options_df = [asdict(record) for record in state.options]
    prices_df = [asdict(record) for record in state.prices]
    addons_df = [asdict(record) for record in state.addons]
    availability_df = [asdict(record) for record in state.availability]
    offers_df = [asdict(record) for record in state.offers]
    elements_df = [asdict(record) for record in state.elements]
    write_csv(OUTPUT_DIR / "pages.csv", pages_df)
    write_csv(OUTPUT_DIR / "options.csv", options_df)
    write_csv(OUTPUT_DIR / "prices_raw.csv", prices_df)
    write_csv(OUTPUT_DIR / "addons.csv", addons_df)
    write_csv(OUTPUT_DIR / "availability.csv", availability_df)
    write_csv(OUTPUT_DIR / "offers.csv", offers_df)
    write_csv(OUTPUT_DIR / "elements.csv", elements_df)
    notes_path = write_notes(state)
    workbook_path = build_excel(state)
    print("Run summary")
    print("============")
    discovered = state.discovery_totals.get("discovered", 0)
    bookable = state.discovery_totals.get("bookable", 0)
    skipped = state.discovery_totals.get("skipped", 0)
    print(f"Discovered URLs: {discovered}   Bookable targets: {bookable}   Skipped (not bookable): {skipped}")
    print(f"Price combinations: {state.summary_counts['combos']}")
    print(f"Add-ons captured: {state.summary_counts['addons']}")
    print(f"Availability slots: {state.summary_counts['slots']}")
    print(f"Offers captured: {state.summary_counts['offers']}")
    print(f"Workbook: {workbook_path}")
    print(f"Notes: {notes_path}")
    print(f"Screenshots folder: {SHOTS_DIR}")


if __name__ == "__main__":
    try:
        asyncio.run(run_scraper())
    except RuntimeError:
        loop = asyncio.get_event_loop()
        loop.run_until_complete(run_scraper())
