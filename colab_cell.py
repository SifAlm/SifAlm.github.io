import importlib
import os
import random
import re
import subprocess
import sys
from collections import defaultdict
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple


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
from bs4 import BeautifulSoup
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

PRICE_SELECTORS = [
    "[data-testid*='total']",
    "[data-testid*='price']",
    "[class*='total']",
    "[class*='price']",
    "[class*='summary'] [class*='amount']",
]
CONTROL_LABEL_HINTS = {
    "HOURS_PROS": ["hour", "duration", "cleaner", "professional", "maid"],
    "ROOM_MATERIAL": ["bedroom", "bathroom", "kitchen", "room", "material"],
    "UNIT_QUANTITY": ["unit", "item", "piece", "appliance", "sofa"],
    "AREA_BASED": ["sq", "area", "meter", "m²", "sqft"],
    "FLAT_PACKAGE": ["package", "plan", "bundle", "basic", "premium"],
}
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

    def log_note(self, message: str) -> None:
        print(message)
        self.notes.append(message)

    def log_error(self, message: str) -> None:
        print(f"[ERROR] {message}")
        self.errors.append(message)


def utc_now() -> str:
    return datetime.utcnow().isoformat(timespec="seconds") + "Z"


def slugify(text: str) -> str:
    base = re.sub(r"[^a-zA-Z0-9]+", "-", text.strip().lower()).strip("-")
    return base or "page"


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


async def extract_links(page) -> List[str]:
    links = await page.eval_on_selector_all("a", "elements => elements.map(el => el.href)")
    return [href for href in links if isinstance(href, str) and href.startswith("https://www.justlife.com/en-AE")]


def filter_service_links(links: Iterable[str]) -> List[str]:
    filtered = []
    for href in links:
        url = canonicalize_url(href)
        if any(excl in url.lower() for excl in ["/blog", "faq", "privacy", "terms", "policy", "sitemap", "cart", "contact", "login"]):
            continue
        if len(url.split("/")) <= 4:
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


async def detect_model(page, state: RunState, page_id: str) -> Tuple[str, Dict[str, Any]]:
    html = await page.content()
    soup = BeautifulSoup(html, "lxml")
    model_scores: Dict[str, float] = defaultdict(float)
    group_evidence: Dict[str, List[str]] = defaultdict(list)
    for model, hints in CONTROL_LABEL_HINTS.items():
        for hint in hints:
            matches = soup.find_all(string=re.compile(hint, re.I))
            if matches:
                model_scores[model] += len(matches)
                for match in matches:
                    snippet = match.strip()
                    group_evidence[model].append(snippet[:80])
    if not model_scores:
        state.elements.append(
            ElementRecord(page_id, "*", "model-detection", "no-hints", "no-hints")
        )
        return "OTHER", {}
    best_model = max(model_scores, key=model_scores.get)
    evidence = {best_model: group_evidence[best_model]}
    for snippet in group_evidence[best_model][:3]:
        state.elements.append(
            ElementRecord(page_id, "text-match", "model-evidence", snippet, snippet.lower())
        )
    return best_model, evidence


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
    for _ in range(8):
        for selector in PRICE_SELECTORS:
            with contextlib.suppress(Exception):
                el = await page.query_selector(selector)
                if el:
                    current = await el.inner_text()
                    if current and current != previous_text:
                        return current
        await asyncio.sleep(1)
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


async def enumerate_options(page, model_type: str, state: RunState, page_id: str) -> Tuple[List[Dict[str, Any]], Dict[str, List[Dict[str, Any]]]]:
    option_groups: Dict[str, List[Dict[str, Any]]] = {}
    records: List[Dict[str, Any]] = []
    seen = set()
    labels = await page.eval_on_selector_all("label", "els => els.map(el => el.innerText)")
    for label in labels:
        normalized = label.strip().lower()
        if not normalized:
            continue
        if normalized in seen:
            continue
        seen.add(normalized)
        if any(hint in normalized for hint in CONTROL_LABEL_HINTS.get(model_type, [])):
            # attempt to find associated inputs
            option_groups[label] = []
    # fallback: look for button groups
    buttons: List[Dict[str, Any]] = []
    with contextlib.suppress(Exception):
        for el in await page.query_selector_all("button"):
            text = ""
            data_test_id = ""
            with contextlib.suppress(Exception):
                raw_text = await el.inner_text()
                text = re.sub(r"\s+", " ", raw_text or "").strip()
            with contextlib.suppress(Exception):
                attr = await el.get_attribute("data-testid")
                if attr:
                    data_test_id = attr
            if text:
                buttons.append({"text": text, "dataTestId": data_test_id})
    for button in buttons:
        text = (button.get("text") or "").strip()
        if len(text) > 40 or len(text) <= 1:
            continue
        lowered = text.lower()
        for model, hints in CONTROL_LABEL_HINTS.items():
            if any(hint in lowered for hint in hints):
                option_groups.setdefault(text, [])
    for group_label in list(option_groups.keys()):
        selector = f"label:has-text(\"{group_label}\")"
        try:
            input_container = await page.query_selector(selector)
        except Exception:
            input_container = None
        options: List[Dict[str, Any]] = []
        if input_container:
            descendants = await input_container.query_selector_all("input, button, option")
        else:
            descendants = await page.query_selector_all("button, [role='option']")
        for idx, desc in enumerate(descendants):
            with contextlib.suppress(Exception):
                text = await desc.inner_text()
                text = re.sub(r"\s+", " ", text).strip()
                if not text or len(text) > 50:
                    continue
                value = await desc.get_attribute("value") or await desc.get_attribute("data-value") or text
                options.append({
                    "label": text,
                    "value": value,
                    "index": idx,
                })
        if not options:
            continue
        option_groups[group_label] = options
        for idx, opt in enumerate(options):
            state.options.append(
                OptionRecord(
                    page_id=page_id,
                    control_group=group_label,
                    option_code=str(opt["value"]),
                    option_label=opt["label"],
                    option_type="button",
                    is_default=(idx == 0),
                    sort_index=idx,
                )
            )
    if not option_groups:
        state.log_note(f"No option groups detected for {page_id}")
    # Build a set of combinations limited by max combos
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
        subset = [choices[0]]
        if len(choices) > 1:
            subset.append(choices[-1])
        mid = len(choices) // 2
        if len(choices) > 2 and choices[mid] not in subset:
            subset.append(choices[mid])
        for choice in subset:
            current[group] = choice
            backtrack(index + 1, current)
        current.pop(group, None)
    backtrack(0, {})
    return combos, option_groups


async def capture_price(page, combos: List[Dict[str, Any]], model_type: str, state: RunState, page_id: str, service_name: str, evidence: Dict[str, Any]) -> None:
    if not combos:
        text = await get_price_text(page)
        currency, amount = parse_price_text(text)
        record = PriceRecord(
            page_id=page_id,
            service_name=service_name,
            model_type=model_type,
            base_price=amount,
            grand_total=amount,
            currency=currency,
            pricing_source_url=page.url,
            collected_at=utc_now(),
            qa_flags="NO_COMBOS",
        )
        state.prices.append(record)
        state.summary_counts["combos"] += 1
        return
    for combo in combos:
        previous = await get_price_text(page)
        for group_label, choice in combo.items():
            await apply_option(page, group_label, choice)
            await asyncio.sleep(1)
        latest = await wait_for_price_change(page, previous)
        currency, amount = parse_price_text(latest)
        qa_flags = []
        if amount is None:
            qa_flags.append("NO_PRICE")
        record = PriceRecord(
            page_id=page_id,
            service_name=service_name,
            model_type=model_type,
            base_price=amount,
            grand_total=amount,
            currency=currency,
            pricing_source_url=page.url,
            collected_at=utc_now(),
            qa_flags=",".join(qa_flags),
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
        state.prices.append(record)
        state.summary_counts["combos"] += 1
        await asyncio.sleep(0.5)


async def apply_option(page, group_label: str, choice: Dict[str, Any]) -> None:
    text = choice.get("label")
    if not text:
        return
    locator = page.locator(f"text={text}")
    try:
        await locator.first.click(timeout=4000)
    except PlaywrightTimeout:
        with contextlib.suppress(Exception):
            handle = await page.query_selector(f"[value='{choice.get('value')}']")
            if handle:
                await handle.click()


async def get_price_text(page) -> str:
    for selector in PRICE_SELECTORS:
        try:
            el = await page.query_selector(selector)
            if el:
                text = await el.inner_text()
                if text:
                    return text
        except PlaywrightTimeout:
            continue
    return ""


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
        target_date = datetime.utcnow().date() + timedelta(days=day_offset)
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
    await asyncio.sleep(2)
    service_name = await get_service_title(browser_page)
    page_id = slugify(service_name + "-" + url.split("/")[-1])
    model_type, evidence = await detect_model(browser_page, state, page_id)
    screenshot_path = SHOTS_DIR / f"service_{page_id}.png"
    await take_screenshot(browser_page, screenshot_path)
    state.screenshots[page_id] = str(screenshot_path)
    combos, option_groups = await enumerate_options(browser_page, model_type, state, page_id)
    await capture_price(browser_page, combos, model_type, state, page_id, service_name, evidence)
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
            page_title_text=service_name,
            detected_model_type=model_type,
            has_addons=any(rec.page_id == page_id for rec in state.addons),
            has_slots=any(rec.page_id == page_id for rec in state.availability),
            notes=json.dumps(evidence)[:250],
        )
    )


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
    links = await extract_links(context_page)
    filtered = filter_service_links(links)
    state.log_note(f"Discovered {len(filtered)} candidate service URLs")
    if len(filtered) < SERVICE_URL_TARGET:
        state.log_note("Warning: fewer than target service URLs detected; continuing anyway")
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
    print(f"Services processed: {state.summary_counts['services']}")
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
