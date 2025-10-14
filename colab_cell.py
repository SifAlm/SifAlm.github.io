"""Google Colab one-cell runner for the Justlife competitive intelligence scraper."""

COLAB_CELL = r"""
# Tamam Intelligence – Justlife Competitive Intelligence Scraper
# This cell installs its own dependencies, launches Playwright Chromium,
# crawls and enumerates service configurations, and produces a styled Excel workbook.
!pip install --quiet playwright requests beautifulsoup4 lxml pandas openpyxl tqdm fake_useragent nest_asyncio
!python -m playwright install --with-deps chromium

import asyncio
import json
import math
import os
import random
import re
import sys
import time
from collections import defaultdict, deque
from dataclasses import dataclass, field, asdict
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import nest_asyncio
import pandas as pd
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
from playwright.async_api import TimeoutError as PlaywrightTimeout, async_playwright
from tqdm import tqdm
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference, PieChart
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

nest_asyncio.apply()

OUTPUT_DIR = Path("justlife_intel")
OUTPUT_DIR.mkdir(exist_ok=True)

SERVICE_KEYWORDS = [
    "clean", "therapy", "nursing", "salon", "laundry", "maid", "pest",
    "tint", "wash", "plumb", "paint", "maintenance", "car", "home",
    "health", "beauty", "wellness", "relocation", "moving", "steril"
]
EXCLUDED_SEGMENTS = ["blog", "terms", "privacy", "faq", "policies", "career", "contact"]
PRICE_SELECTORS = [
    "[data-testid*='price']",
    "[data-test*='price']",
    "[class*='price']",
    "[class*='total']",
    "[data-testid*='summary'] [class*='amount']",
    "[class*='summary'] [class*='amount']",
]
ADDON_SELECTORS = [
    "[data-testid*='addon']",
    "[class*='addon']",
    "[class*='upsell']",
    "[data-test*='upsell']",
]
AVAILABILITY_TRIGGER_TEXT = ["schedule", "slot", "time", "date", "book", "choose"]
SLOT_SELECTORS = [
    "[data-testid*='slot']",
    "[class*='slot']",
    "button:has-text('AM')",
    "button:has-text('PM')",
]
OFFER_SELECTORS = [
    "[data-testid*='offer']",
    "[class*='offer']",
    "[class*='discount']",
]


@dataclass
class Config:
    base_url: str = os.environ.get("BASE", "https://www.justlife.com/en-AE")
    market: str = os.environ.get("MARKET", "dubai")
    days: int = int(os.environ.get("DAYS", "14"))
    max_combinations: int = int(os.environ.get("MAX_COMBOS", "60"))
    headless: bool = os.environ.get("HEADLESS", "1") == "1"
    throttle_seconds: float = 1.5
    jitter_seconds: float = 0.75
    accept_language: str = "en-AE,en;q=0.9"
    user_agent: Optional[str] = None
    discovery_limit: int = 80
    per_service_retry: int = 2


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
    dimensions: Dict[str, Any]
    offer_state: str
    offer_name: Optional[str]
    discount_type: Optional[str]
    discount_value: Optional[float]
    base_price: float
    fees_total: float
    vat_amount: float
    vat_percent: Optional[float]
    grand_total: float
    currency: str
    pricing_source_url: str
    collected_at: str
    qa_flags: List[str]
    selector_meta: Dict[str, Any]


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


@dataclass
class ControlOption:
    control_group: str
    option_label: str
    option_code: str
    selector: str
    option_type: str
    value: Optional[str]
    is_default: bool
    sort_index: int


@dataclass
class Control:
    group: str
    control_type: str
    options: List[ControlOption]


@dataclass
class ServiceArtifacts:
    page: PageRecord
    options: List[OptionRecord]
    prices: List[PriceRecord]
    addons: List[AddonRecord]
    availability: List[AvailabilityRecord]
    offers: List[OfferRecord]
    elements: List[ElementRecord]


@dataclass
class PipelineArtifacts:
    pages: List[PageRecord] = field(default_factory=list)
    options: List[OptionRecord] = field(default_factory=list)
    prices: List[PriceRecord] = field(default_factory=list)
    addons: List[AddonRecord] = field(default_factory=list)
    availability: List[AvailabilityRecord] = field(default_factory=list)
    offers: List[OfferRecord] = field(default_factory=list)
    elements: List[ElementRecord] = field(default_factory=list)
    notes: List[str] = field(default_factory=list)


# Utility helpers

def utc_now() -> str:
    return datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


def uuid5_url(url: str) -> str:
    import uuid
    return str(uuid.uuid5(uuid.NAMESPACE_URL, url))


def normalize_space(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def slugify(value: str) -> str:
    value = normalize_space(value)
    value = re.sub(r"[^0-9A-Za-z]+", "-", value)
    value = value.strip("-")
    return value.lower() or "value"


def level_for_depth(depth: int) -> str:
    if depth <= 0:
        return "home"
    if depth == 1:
        return "category"
    if depth == 2:
        return "subcategory"
    if depth == 3:
        return "service"
    return "flow"


def throttle_delay(config: Config) -> float:
    base = random.gauss(config.throttle_seconds, config.jitter_seconds / 2)
    return max(0.2, base)


def safe_float(value: Optional[str]) -> Optional[float]:
    if value is None:
        return None
    try:
        cleaned = value.replace(",", "").replace("AED", "").strip()
        return float(re.findall(r"[0-9]+(?:\.[0-9]+)?", cleaned)[0])
    except Exception:
        return None


def find_currency(text: str) -> Tuple[str, float]:
    match = re.search(r"(AED|د.إ)\s*([0-9.,]+)", text)
    if match:
        value = float(match.group(2).replace(",", ""))
        return match.group(1), value
    numbers = re.findall(r"[0-9]+(?:\.[0-9]+)?", text)
    if numbers:
        return "AED", float(numbers[0])
    return "AED", 0.0


async def read_first_text(page, selectors: List[str]) -> str:
    for selector in selectors:
        try:
            locator = page.locator(selector).first
            if await locator.count() == 0:
                continue
            text = normalize_space(await locator.inner_text())
            if text:
                return text
        except Exception:
            continue
    try:
        locator = page.locator("xpath=//*[contains(text(), 'AED')]").first
        if await locator.count():
            return normalize_space(await locator.inner_text())
    except Exception:
        pass
    return ""


async def wait_for_price_change(page, previous: str, timeout: float = 6.0) -> str:
    start = time.time()
    last = previous
    while time.time() - start < timeout:
        current = await read_first_text(page, PRICE_SELECTORS)
        if current and normalize_space(current) != normalize_space(last):
            return current
        await asyncio.sleep(0.4)
    return last


async def extract_breakdown(page) -> Dict[str, Any]:
    selectors = ["[data-testid*='summary']", "[class*='summary']", "[data-testid*='breakdown']", "[class*='breakdown']"]
    for selector in selectors:
        try:
            locator = page.locator(selector).first
            if await locator.count():
                text = normalize_space(await locator.inner_text())
                if text:
                    return {"text": text}
        except Exception:
            continue
    return {}


async def js_controls(page) -> List[Dict[str, Any]]:
    script = \"\"\"
    () => {
        const cssPath = (el) => {
            if (!el) return null;
            const path = [];
            while (el && el.nodeType === 1 && path.length < 8) {
                let selector = el.nodeName.toLowerCase();
                if (el.id) {
                    selector += '#' + CSS.escape(el.id);
                    path.unshift(selector);
                    break;
                } else {
                    let sibling = el;
                    let nth = 1;
                    while (sibling = sibling.previousElementSibling) {
                        if (sibling.nodeName === el.nodeName) nth += 1;
                    }
                    selector += `:nth-of-type(${nth})`;
                    path.unshift(selector);
                    el = el.parentElement;
                }
            }
            return path.join(' > ');
        };

        const controls = [];
        const pushControl = (group, type, el, extra = {}) => {
            if (!el) return;
            const label = (el.innerText || el.textContent || '').trim();
            if (!label) return;
            const selector = cssPath(el);
            if (!selector) return;
            controls.push({
                group,
                type,
                selector,
                optionLabel: label,
                value: el.value || el.getAttribute('data-value') || null,
                meta: extra
            });
        };

        const labelFor = (el) => {
            if (!el) return null;
            const labelledby = el.getAttribute('aria-labelledby');
            if (labelledby) {
                const labelEl = document.getElementById(labelledby);
                if (labelEl) return (labelEl.innerText || labelEl.textContent || '').trim();
            }
            const ariaLabel = el.getAttribute('aria-label');
            if (ariaLabel) return ariaLabel.trim();
            const id = el.id;
            if (id) {
                const label = document.querySelector(`label[for="${CSS.escape(id)}"]`);
                if (label) return (label.innerText || label.textContent || '').trim();
            }
            const prev = el.closest('div, section, article');
            if (prev) {
                const heading = prev.querySelector('h2, h3, h4, strong, p');
                if (heading) {
                    const text = (heading.innerText || heading.textContent || '').trim();
                    if (text.length <= 80) return text;
                }
            }
            return null;
        };

        const registerControl = (group, type, el, extra = {}) => {
            pushControl(group || 'Options', type, el, extra);
        };

        document.querySelectorAll('select').forEach((selectEl, idx) => {
            const group = labelFor(selectEl) || selectEl.name || `Select ${idx + 1}`;
            Array.from(selectEl.options).forEach((opt, optIndex) => {
                if (!opt.value && !opt.textContent.trim()) return;
                const option = {
                    group,
                    type: 'select',
                    selector: cssPath(selectEl),
                    optionLabel: (opt.innerText || opt.textContent || '').trim() || opt.value,
                    value: opt.value || opt.textContent.trim(),
                    meta: { optionIndex: optIndex }
                };
                controls.push(option);
            });
        });

        const clickableSelectors = [
            "button",
            "[role='button']",
            "[role='option']",
            "[data-testid*='option']",
            "[class*='pill']",
            "[class*='chip']",
            "[class*='option']",
            "[class*='selector']"
        ];
        clickableSelectors.forEach(sel => {
            document.querySelectorAll(sel).forEach((el, idx) => {
                if (!el || el.disabled || el.getAttribute('aria-hidden') === 'true') return;
                if (!el.offsetParent) return;
                const text = (el.innerText || el.textContent || '').trim();
                if (!text || text.length > 80) return;
                const groupEl = el.closest('[data-testid*="option"], [data-testid*="group"], [role="radiogroup"], [role="group"], [class*="options"], [class*="group"], section, article, div');
                let group = null;
                if (groupEl) {
                    group = groupEl.getAttribute('data-testid') || groupEl.getAttribute('aria-label');
                    if (!group) {
                        const header = groupEl.querySelector('h2, h3, h4, strong, label');
                        if (header) group = (header.innerText || header.textContent || '').trim();
                    }
                }
                if (!group) group = `Choice ${idx + 1}`;
                const selector = cssPath(el);
                if (!selector) return;
                controls.push({
                    group,
                    type: 'click',
                    selector,
                    optionLabel: text,
                    value: null,
                    meta: {}
                });
            });
        });

        return controls;
    }
    \"\"\"
    try:
        return await page.evaluate(script)
    except Exception:
        return []


class JustlifeScraper:
    def __init__(self, config: Config):
        self.config = config
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        self.artifacts = PipelineArtifacts()
        self.errors: List[str] = []

    async def __aenter__(self):
        ua = self.config.user_agent
        if not ua:
            try:
                ua = UserAgent().chrome
            except Exception:
                ua = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
        self.config.user_agent = ua
        self.playwright = await async_playwright().start()
        self.browser = await self.playwright.chromium.launch(headless=self.config.headless)
        self.context = await self.browser.new_context(
            locale="en-AE",
            user_agent=self.config.user_agent,
            extra_http_headers={"Accept-Language": self.config.accept_language},
            viewport={"width": 1280, "height": 720},
        )
        self.page = await self.context.new_page()
        self.page.set_default_timeout(45000)
        return self

    async def __aexit__(self, exc_type, exc, tb):
        if self.context:
            await self.context.close()
        if self.browser:
            await self.browser.close()
        if self.playwright:
            await self.playwright.stop()

    async def navigate(self, url: str) -> bool:
        for attempt in range(3):
            try:
                await self.page.goto(url, wait_until="domcontentloaded")
                await asyncio.sleep(throttle_delay(self.config))
                return True
            except PlaywrightTimeout:
                await asyncio.sleep(1.5)
            except Exception as err:
                self.errors.append(f"Navigation error {url}: {err}")
                await asyncio.sleep(1.5)
        return False

    async def discover(self):
        queue = deque([(self.config.base_url.rstrip('/'), None, 0)])
        seen = set()
        while queue and len(self.artifacts.pages) < self.config.discovery_limit:
            url, parent_id, depth = queue.popleft()
            if url in seen:
                continue
            seen.add(url)
            ok = await self.navigate(url)
            if not ok:
                self.artifacts.notes.append(f"Failed to navigate during discovery: {url}")
                continue
            content = await self.page.content()
            soup = BeautifulSoup(content, "lxml")
            title = normalize_space(soup.title.text if soup.title else await self.page.title())
            model_type, has_addons, has_slots, notes, element_logs = await self.detect_model(soup)
            page_id = uuid5_url(url)
            page_record = PageRecord(
                page_id=page_id,
                parent_page_id=parent_id,
                level=level_for_depth(depth),
                url=url,
                page_title_text=title or slugify(url.split("/")[-1]),
                detected_model_type=model_type,
                has_addons=has_addons,
                has_slots=has_slots,
                notes=notes,
            )
            self.artifacts.pages.append(page_record)
            for log in element_logs:
                log.page_id = page_id
                self.artifacts.elements.append(log)
            new_links = self.extract_links(url, soup)
            for link in new_links:
                if link not in seen:
                    queue.append((link, page_id, depth + 1))

    def extract_links(self, current_url: str, soup: BeautifulSoup) -> List[str]:
        base = self.config.base_url.split("//", 1)[-1]
        links: List[str] = []
        for anchor in soup.select("a[href]"):
            href = anchor.get("href")
            if not href:
                continue
            if href.startswith("javascript") or href.startswith("#"):
                continue
            if href.startswith("/"):
                href = f"https://{base}{href}"
            if not href.startswith(self.config.base_url):
                continue
            normalized = href.split("#")[0].rstrip("/")
            if any(token in normalized.lower() for token in EXCLUDED_SEGMENTS):
                continue
            if not any(keyword in normalized.lower() for keyword in SERVICE_KEYWORDS):
                continue
            links.append(normalized)
        return list(dict.fromkeys(links))

    async def detect_model(self, soup: BeautifulSoup) -> Tuple[str, bool, bool, str, List[ElementRecord]]:
        body_text = soup.get_text(" ")
        lowered = body_text.lower()
        detections: List[str] = []
        if any(word in lowered for word in ["hour", "hrs", "cleaner", "professional"]):
            model = "HOURS_PROS"
            detections.append("Detected HOURS_PROS via keywords")
        elif any(word in lowered for word in ["bedroom", "bathroom", "kitchen", "villa"]):
            model = "ROOM_MATERIAL"
            detections.append("Detected ROOM_MATERIAL via keywords")
        elif any(word in lowered for word in ["unit", "piece", "items", "pieces"]):
            model = "UNIT_QUANTITY"
            detections.append("Detected UNIT_QUANTITY via keywords")
        elif any(word in lowered for word in ["sqft", "square", "area"]):
            model = "AREA_BASED"
            detections.append("Detected AREA_BASED via keywords")
        elif any(word in lowered for word in ["package", "plan", "premium", "bundle"]):
            model = "FLAT_PACKAGE"
            detections.append("Detected FLAT_PACKAGE via keywords")
        else:
            model = "OTHER"
            detections.append("No model keyword match")
        has_addons = bool(soup.select_one(", ".join(ADDON_SELECTORS)))
        has_slots = "slot" in lowered or bool(soup.select_one(", ".join(SLOT_SELECTORS)))
        if has_addons:
            detections.append("Add-on markers found")
        if has_slots:
            detections.append("Potential availability slots")
        element_logs = [
            ElementRecord(
                page_id="",
                selector="body",
                element_role="detection",
                raw_text=body_text[:2000],
                normalized_text=", ".join(detections),
            )
        ]
        return model, has_addons, has_slots, "; ".join(detections), element_logs

    async def scrape_services(self):
        service_pages = [page for page in self.artifacts.pages if page.level in {"service", "flow"}]
        for page_record in tqdm(service_pages, desc="Scraping services"):
            success = False
            for attempt in range(self.config.per_service_retry):
                ok = await self.navigate(page_record.url)
                if not ok:
                    continue
                try:
                    artifacts = await self.capture_service(page_record)
                    self.artifacts.options.extend(artifacts.options)
                    self.artifacts.prices.extend(artifacts.prices)
                    self.artifacts.addons.extend(artifacts.addons)
                    self.artifacts.availability.extend(artifacts.availability)
                    self.artifacts.offers.extend(artifacts.offers)
                    self.artifacts.elements.extend(artifacts.elements)
                    success = True
                    break
                except Exception as err:
                    self.errors.append(f"Service capture failed {page_record.url}: {err}")
                    await asyncio.sleep(1.0)
            if not success:
                self.artifacts.notes.append(f"Failed to capture service after retries: {page_record.url}")

    async def capture_service(self, page_record: PageRecord) -> ServiceArtifacts:
        option_controls = await self.build_controls(page_record)
        price_records: List[PriceRecord] = []
        option_records: List[OptionRecord] = []
        element_records: List[ElementRecord] = []
        combinations: List[List[ControlOption]] = []
        if option_controls:
            pools = [control.options for control in option_controls]
            current: List[List[ControlOption]] = [[]]
            for options in pools:
                temp = []
                for prefix in current:
                    for option in options:
                        temp.append(prefix + [option])
                current = temp
            combinations = current[: self.config.max_combinations]
        else:
            combinations = [[]]
        visited_dimensions = set()
        for control in option_controls:
            for opt in control.options:
                option_records.append(
                    OptionRecord(
                        page_id=page_record.page_id,
                        control_group=control.group,
                        option_code=opt.option_code,
                        option_label=opt.option_label,
                        option_type=opt.option_type,
                        is_default=opt.is_default,
                        sort_index=opt.sort_index,
                    )
                )
        combo_progress = tqdm(combinations, desc=f"Combos {page_record.page_title_text}", leave=False)
        for combo in combo_progress:
            await self.navigate(page_record.url)
            await asyncio.sleep(0.5)
            dimensions: Dict[str, Any] = {}
            selector_meta = {"controls": []}
            for option in combo:
                dimensions[option.control_group] = option.option_label
                selector_meta["controls"].append({
                    "group": option.control_group,
                    "selector": option.selector,
                    "label": option.option_label,
                })
                await self.apply_option(option)
                await asyncio.sleep(0.5)
            key = tuple(sorted(dimensions.items()))
            if key in visited_dimensions:
                continue
            visited_dimensions.add(key)
            price_text_before = await read_first_text(self.page, PRICE_SELECTORS)
            price_text = await wait_for_price_change(self.page, price_text_before)
            qa_flags: List[str] = []
            if not price_text:
                qa_flags.append("NO_PRICE")
            currency, grand_total = find_currency(price_text)
            breakdown = await extract_breakdown(self.page)
            base_price = grand_total
            fees = 0.0
            vat = 0.0
            vat_percent = None
            if breakdown.get("text"):
                numbers = [float(n.replace(",", "")) for n in re.findall(r"[0-9]+(?:\.[0-9]+)?", breakdown["text"])]
                if numbers:
                    base_price = numbers[0]
                if len(numbers) >= 2:
                    fees = numbers[1]
                if len(numbers) >= 3:
                    vat = numbers[2]
                if vat and base_price:
                    vat_percent = round((vat / max(base_price, 1)) * 100, 2)
            else:
                breakdown_text = await self.page.content()
                if "vat" in breakdown_text.lower():
                    numbers = [float(n.replace(",", "")) for n in re.findall(r"[0-9]+(?:\.[0-9]+)?", breakdown_text)]
                    if len(numbers) >= 3:
                        base_price, fees, vat = numbers[:3]
                        vat_percent = round((vat / max(base_price, 1)) * 100, 2)
            if grand_total and abs((base_price + fees + vat) - grand_total) > max(1.0, grand_total) * 0.02:
                qa_flags.append("MATH_MISMATCH")
            element_records.append(
                ElementRecord(
                    page_id=page_record.page_id,
                    selector="price",
                    element_role="price",
                    raw_text=price_text,
                    normalized_text=json.dumps(breakdown),
                )
            )
            price_records.append(
                PriceRecord(
                    page_id=page_record.page_id,
                    service_name=page_record.page_title_text,
                    model_type=page_record.detected_model_type,
                    dimensions=dimensions,
                    offer_state="OFF",
                    offer_name=None,
                    discount_type=None,
                    discount_value=None,
                    base_price=base_price,
                    fees_total=fees,
                    vat_amount=vat,
                    vat_percent=vat_percent,
                    grand_total=grand_total,
                    currency=currency,
                    pricing_source_url=page_record.url,
                    collected_at=utc_now(),
                    qa_flags=qa_flags,
                    selector_meta=selector_meta,
                )
            )
        addons = await self.capture_addons(page_record)
        availability = await self.capture_availability(page_record)
        offers = await self.capture_offers(page_record)
        return ServiceArtifacts(
            page=page_record,
            options=option_records,
            prices=price_records,
            addons=addons,
            availability=availability,
            offers=offers,
            elements=element_records,
        )

    async def build_controls(self, page_record: PageRecord) -> List[Control]:
        controls_map: Dict[str, List[ControlOption]] = defaultdict(list)
        js_results = await js_controls(self.page)
        for index, entry in enumerate(js_results):
            group = normalize_space(entry.get("group") or "Options")
            option_label = normalize_space(entry.get("optionLabel", ""))
            if not option_label:
                continue
            option_type = entry.get("type") or "click"
            selector = entry.get("selector")
            value = entry.get("value")
            option = ControlOption(
                control_group=group,
                option_label=option_label,
                option_code=slugify(f"{group}-{option_label}"),
                selector=selector,
                option_type=option_type,
                value=value,
                is_default=index == 0,
                sort_index=len(controls_map[group]),
            )
            controls_map[group].append(option)
        controls: List[Control] = []
        for group, options in controls_map.items():
            if len(options) <= 1:
                continue
            controls.append(Control(group=group, control_type=options[0].option_type, options=options))
        if not controls:
            self.artifacts.notes.append(f"No controls detected on {page_record.url}")
        return controls

    async def apply_option(self, option: ControlOption) -> None:
        try:
            if option.option_type == "select":
                await self.page.select_option(option.selector, option.value)
            else:
                await self.page.locator(option.selector).click()
        except Exception as err:
            self.errors.append(f"Option interaction failed {option.option_label}: {err}")

    async def capture_addons(self, page_record: PageRecord) -> List[AddonRecord]:
        addons: List[AddonRecord] = []
        content = await self.page.content()
        soup = BeautifulSoup(content, "lxml")
        for selector in ADDON_SELECTORS:
            for block in soup.select(selector):
                name = normalize_space(block.get_text(" "))[:120]
                if not name:
                    continue
                price_match = re.search(r"(AED|د.إ)\s*([0-9.,]+)", name)
                currency = None
                price_value = None
                if price_match:
                    currency = price_match.group(1)
                    price_value = float(price_match.group(2).replace(",", ""))
                addons.append(
                    AddonRecord(
                        page_id=page_record.page_id,
                        addon_name=name[:80],
                        addon_description=name[:200],
                        addon_price=price_value,
                        currency=currency,
                        is_recommended="recommended" in name.lower(),
                        is_required="required" in name.lower(),
                        seen_on_step="service",
                        pricing_source_url=page_record.url,
                        collected_at=utc_now(),
                    )
                )
        return addons

    async def capture_availability(self, page_record: PageRecord) -> List[AvailabilityRecord]:
        records: List[AvailabilityRecord] = []
        try:
            trigger_found = False
            for text in AVAILABILITY_TRIGGER_TEXT:
                locator = self.page.get_by_text(re.compile(text, re.I))
                if await locator.count():
                    try:
                        await locator.first.click()
                        await asyncio.sleep(0.5)
                        trigger_found = True
                        break
                    except Exception:
                        continue
            if not trigger_found:
                return records
            for day_offset in range(min(self.config.days, 14)):
                date = datetime.utcnow().date() + timedelta(days=day_offset)
                date_str = date.isoformat()
                for selector in SLOT_SELECTORS:
                    slots = self.page.locator(selector)
                    count = await slots.count()
                    for idx in range(count):
                        slot = slots.nth(idx)
                        try:
                            label = normalize_space(await slot.inner_text())
                        except Exception:
                            continue
                        if not label:
                            continue
                        class_name = await slot.get_attribute("class") or ""
                        status = "disabled" if "disabled" in class_name.lower() else "available"
                        records.append(
                            AvailabilityRecord(
                                page_id=page_record.page_id,
                                date_str=date_str,
                                time_slot_label=label,
                                slot_status=status,
                                lead_time_days=day_offset,
                                collected_at=utc_now(),
                            )
                        )
                next_btn = self.page.get_by_role("button", name=re.compile("next", re.I))
                if await next_btn.count():
                    try:
                        await next_btn.click()
                        await asyncio.sleep(0.4)
                    except Exception:
                        break
                else:
                    break
        except Exception:
            self.artifacts.notes.append(f"Availability capture failed for {page_record.url}")
        return records

    async def capture_offers(self, page_record: PageRecord) -> List[OfferRecord]:
        offers: List[OfferRecord] = []
        try:
            content = await self.page.content()
            soup = BeautifulSoup(content, "lxml")
            for selector in OFFER_SELECTORS:
                for block in soup.select(selector):
                    text = normalize_space(block.get_text(" "))
                    if not text:
                        continue
                    badge = ""
                    desc = text
                    if ":" in text:
                        parts = text.split(":", 1)
                        badge = parts[0]
                        desc = parts[1]
                    offers.append(
                        OfferRecord(
                            page_id=page_record.page_id,
                            offer_name=badge or text[:60],
                            offer_badge_text=badge,
                            description=desc[:120],
                            eligibility="",
                            discount_type="percent" if "%" in text else None,
                            discount_value=safe_float(text),
                            stacking_rules="",
                            collected_at=utc_now(),
                        )
                    )
        except Exception:
            self.artifacts.notes.append(f"Offer capture failed for {page_record.url}")
        return offers

    async def run(self) -> PipelineArtifacts:
        await self.discover()
        await self.scrape_services()
        if self.errors:
            unique_errors = list(dict.fromkeys(self.errors))[:10]
            self.artifacts.notes.append("Top errors: " + " | ".join(unique_errors))
        return self.artifacts


class ExcelBuilder:
    def __init__(self, output_dir: Path):
        self.output_dir = output_dir
        self.path = output_dir / "Justlife_CRM.xlsx"

    def build(self, data: PipelineArtifacts, summary: Dict[str, Any]) -> Path:
        workbook = Workbook()
        self._build_readme(workbook, summary, data)
        self._build_pages_sheet(workbook, data)
        self._build_options_sheet(workbook, data)
        self._build_prices_sheet(workbook, data)
        self._build_addons_sheet(workbook, data)
        self._build_availability_sheet(workbook, data)
        self._build_offers_sheet(workbook, data)
        self._build_elements_sheet(workbook, data)
        workbook.save(self.path)
        return self.path

    def _style_table(self, ws, freeze: bool = True):
        header_font = Font(bold=True, color="FFFFFF")
        fill = PatternFill("solid", fgColor="1B4965")
        if ws.max_row:
            for cell in ws[1]:
                cell.font = header_font
                cell.fill = fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
        if freeze:
            ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for column in range(1, ws.max_column + 1):
            width = 14
            for row in range(1, ws.max_row + 1):
                value = ws.cell(row=row, column=column).value
                if value is None:
                    continue
                width = max(width, min(60, len(str(value)) + 2))
            ws.column_dimensions[get_column_letter(column)].width = width
        zebra = PatternFill("solid", fgColor="F5F5F5")
        for idx in range(2, ws.max_row + 1, 2):
            for col in range(1, ws.max_column + 1):
                ws.cell(row=idx, column=col).fill = zebra

    def _build_readme(self, wb: Workbook, summary: Dict[str, Any], data: PipelineArtifacts):
        ws = wb.active
        ws.title = "README"
        ws.append(["Tamam Intelligence – Justlife Competitive Intelligence Dashboard"])
        ws.append(["Generated:", summary["timestamp"]])
        ws.append(["Base URL:", summary["base_url"]])
        ws.append(["Market:", summary["market"]])
        ws.append(["Services scraped:", summary["services"]])
        ws.append(["Price combinations:", summary["price_rows"]])
        ws.append(["Add-ons captured:", summary["addons"]])
        ws.append(["Availability rows:", summary["availability"]])
        ws.append(["Offers:", summary["offers"]])
        ws.append(["Notes:"])
        notes_text = "\n".join(data.notes) if data.notes else "-"
        ws.append([notes_text])
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=2):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions["A"].width = 40
        ws.column_dimensions["B"].width = 60

        # Charts
        prices_by_model = defaultdict(int)
        for price in data.prices:
            prices_by_model[price.model_type] += 1
        if prices_by_model:
            start_row = ws.max_row + 2
            ws.append([])
            ws.append(["Model", "Price Rows"])
            for model, count in prices_by_model.items():
                ws.append([model, count])
            chart = BarChart()
            chart.title = "Price rows by model"
            chart.add_data(Reference(ws, min_col=2, min_row=start_row + 1, max_row=ws.max_row), titles_from_data=False)
            chart.set_categories(Reference(ws, min_col=1, min_row=start_row + 1, max_row=ws.max_row))
            chart.height = 8
            chart.width = 14
            ws.add_chart(chart, f"E2")
        addons_by_service = defaultdict(int)
        for addon in data.addons:
            addons_by_service[addon.page_id] += 1
        if addons_by_service:
            start_row = ws.max_row + 2
            ws.append([])
            ws.append(["Service", "Add-ons"])
            for service_id, count in addons_by_service.items():
                ws.append([service_id, count])
            pie = PieChart()
            pie.title = "Add-on volume by service"
            pie.add_data(Reference(ws, min_col=2, min_row=start_row + 1, max_row=ws.max_row), titles_from_data=False)
            pie.set_categories(Reference(ws, min_col=1, min_row=start_row + 1, max_row=ws.max_row))
            pie.height = 8
            pie.width = 8
            ws.add_chart(pie, f"M2")

    def _build_pages_sheet(self, wb: Workbook, data: PipelineArtifacts):
        ws = wb.create_sheet("Services Summary")
        rows = [
            [
                "page_id", "parent_page_id", "level", "url", "page_title_text",
                "detected_model_type", "has_addons", "has_slots", "notes"
            ]
        ]
        for page in data.pages:
            rows.append([
                page.page_id,
                page.parent_page_id,
                page.level,
                page.url,
                page.page_title_text,
                page.detected_model_type,
                page.has_addons,
                page.has_slots,
                page.notes,
            ])
        for row in rows:
            ws.append(row)
        self._style_table(ws)

    def _build_options_sheet(self, wb: Workbook, data: PipelineArtifacts):
        ws = wb.create_sheet("Options Matrix")
        ws.append(["page_id", "control_group", "option_code", "option_label", "option_type", "is_default", "sort_index"])
        for option in data.options:
            ws.append([
                option.page_id,
                option.control_group,
                option.option_code,
                option.option_label,
                option.option_type,
                option.is_default,
                option.sort_index,
            ])
        self._style_table(ws)

    def _build_prices_sheet(self, wb: Workbook, data: PipelineArtifacts):
        ws = wb.create_sheet("Price Matrix")
        base_headers = [
            "page_id", "service_name", "model_type", "offer_state", "offer_name",
            "discount_type", "discount_value", "base_price", "fees_total",
            "vat_amount", "vat_percent", "grand_total", "currency", "pricing_source_url",
            "collected_at", "qa_flags", "selector_meta"
        ]
        dimension_keys: List[str] = []
        for price in data.prices:
            for key in price.dimensions.keys():
                if key not in dimension_keys:
                    dimension_keys.append(key)
        headers = base_headers[:3] + dimension_keys + base_headers[3:]
        ws.append(headers)
        for price in data.prices:
            row = [price.page_id, price.service_name, price.model_type]
            for key in dimension_keys:
                row.append(price.dimensions.get(key))
            row.extend([
                price.offer_state,
                price.offer_name,
                price.discount_type,
                price.discount_value,
                price.base_price,
                price.fees_total,
                price.vat_amount,
                price.vat_percent,
                price.grand_total,
                price.currency,
                price.pricing_source_url,
                price.collected_at,
                ",".join(price.qa_flags),
                json.dumps(price.selector_meta, ensure_ascii=False),
            ])
            ws.append(row)
        self._style_table(ws)
        grand_total_col = headers.index("grand_total") + 1
        qa_col = headers.index("qa_flags") + 1
        offer_col = headers.index("offer_state") + 1
        if ws.max_row > 1:
            col_letter = get_column_letter(grand_total_col)
            ws.conditional_formatting.add(
                f"{col_letter}2:{col_letter}{ws.max_row}",
                ColorScaleRule(start_type="min", start_color="F2F6FF", mid_type="percentile", mid_value=50, mid_color="B7D3FF", end_type="max", end_color="124E96"),
            )
            qa_letter = get_column_letter(qa_col)
            ws.conditional_formatting.add(
                f"{qa_letter}2:{qa_letter}{ws.max_row}",
                FormulaRule(formula=[f"LEN(${qa_letter}2)>0"], fill=PatternFill("solid", fgColor="FFC7CE")),
            )
            offer_letter = get_column_letter(offer_col)
            ws.conditional_formatting.add(
                f"{offer_letter}2:{offer_letter}{ws.max_row}",
                FormulaRule(formula=[f"${offer_letter}2=\"ON\""], fill=PatternFill("solid", fgColor="E2F0CB")),
            )

    def _build_addons_sheet(self, wb: Workbook, data: PipelineArtifacts):
        ws = wb.create_sheet("Add-ons")
        ws.append(["page_id", "addon_name", "addon_description", "addon_price", "currency", "is_recommended", "is_required", "seen_on_step", "pricing_source_url", "collected_at"])
        for addon in data.addons:
            ws.append([
                addon.page_id,
                addon.addon_name,
                addon.addon_description,
                addon.addon_price,
                addon.currency,
                addon.is_recommended,
                addon.is_required,
                addon.seen_on_step,
                addon.pricing_source_url,
                addon.collected_at,
            ])
        self._style_table(ws)

    def _build_availability_sheet(self, wb: Workbook, data: PipelineArtifacts):
        ws = wb.create_sheet("Availability")
        ws.append(["page_id", "date_str", "time_slot_label", "slot_status", "lead_time_days", "collected_at"])
        for record in data.availability:
            ws.append([
                record.page_id,
                record.date_str,
                record.time_slot_label,
                record.slot_status,
                record.lead_time_days,
                record.collected_at,
            ])
        self._style_table(ws)

    def _build_offers_sheet(self, wb: Workbook, data: PipelineArtifacts):
        ws = wb.create_sheet("Offers")
        ws.append(["page_id", "offer_name", "offer_badge_text", "description", "eligibility", "discount_type", "discount_value", "stacking_rules", "collected_at"])
        for offer in data.offers:
            ws.append([
                offer.page_id,
                offer.offer_name,
                offer.offer_badge_text,
                offer.description,
                offer.eligibility,
                offer.discount_type,
                offer.discount_value,
                offer.stacking_rules,
                offer.collected_at,
            ])
        self._style_table(ws)

    def _build_elements_sheet(self, wb: Workbook, data: PipelineArtifacts):
        ws = wb.create_sheet("Elements")
        ws.append(["page_id", "selector", "element_role", "raw_text", "normalized_text"])
        for element in data.elements:
            ws.append([
                element.page_id,
                element.selector,
                element.element_role,
                element.raw_text,
                element.normalized_text,
            ])
        self._style_table(ws)


async def main():
    config = Config()
    print("Configuration:", json.dumps(asdict(config), indent=2))
    async with JustlifeScraper(config) as scraper:
        artifacts = await scraper.run()
    summary = {
        "timestamp": utc_now(),
        "base_url": config.base_url,
        "market": config.market,
        "services": len([p for p in artifacts.pages if p.level in {"service", "flow"}]),
        "price_rows": len(artifacts.prices),
        "addons": len(artifacts.addons),
        "availability": len(artifacts.availability),
        "offers": len(artifacts.offers),
    }
    csv_map = write_outputs(artifacts)
    excel_path = ExcelBuilder(OUTPUT_DIR).build(artifacts, summary)
    print("\nRun summary:")
    for key, value in summary.items():
        print(f"- {key}: {value}")
    print("\nGenerated files:")
    for name, path in csv_map.items():
        print(f"- {name}: {path}")
    print(f"- excel: {excel_path}")
    try:
        from google.colab import files
        files.download(str(excel_path))
    except Exception as err:
        print(f"Download skipped: {err}")


def write_outputs(artifacts: PipelineArtifacts) -> Dict[str, Path]:
    csv_map: Dict[str, Path] = {}
    csv_map["pages"] = OUTPUT_DIR / "pages.csv"
    csv_map["options"] = OUTPUT_DIR / "options.csv"
    csv_map["prices_raw"] = OUTPUT_DIR / "prices_raw.csv"
    csv_map["addons"] = OUTPUT_DIR / "addons.csv"
    csv_map["availability"] = OUTPUT_DIR / "availability.csv"
    csv_map["offers"] = OUTPUT_DIR / "offers.csv"
    csv_map["elements"] = OUTPUT_DIR / "elements.csv"
    csv_map["notes"] = OUTPUT_DIR / "NOTES.txt"

    def write_csv(path: Path, rows: List[Dict[str, Any]], headers: Optional[List[str]] = None):
        df = pd.DataFrame(rows, columns=headers)
        df.to_csv(path, index=False)

    write_csv(csv_map["pages"], [asdict(page) for page in artifacts.pages])
    write_csv(csv_map["options"], [asdict(option) for option in artifacts.options])
    price_rows = []
    for price in artifacts.prices:
        row = {
            "page_id": price.page_id,
            "service_name": price.service_name,
            "model_type": price.model_type,
            "offer_state": price.offer_state,
            "offer_name": price.offer_name,
            "discount_type": price.discount_type,
            "discount_value": price.discount_value,
            "base_price": price.base_price,
            "fees_total": price.fees_total,
            "vat_amount": price.vat_amount,
            "vat_percent": price.vat_percent,
            "grand_total": price.grand_total,
            "currency": price.currency,
            "pricing_source_url": price.pricing_source_url,
            "collected_at": price.collected_at,
            "qa_flags": ",".join(price.qa_flags),
            "selector_meta": json.dumps(price.selector_meta, ensure_ascii=False),
        }
        row.update(price.dimensions)
        price_rows.append(row)
    write_csv(csv_map["prices_raw"], price_rows)
    write_csv(csv_map["addons"], [asdict(addon) for addon in artifacts.addons])
    write_csv(csv_map["availability"], [asdict(record) for record in artifacts.availability])
    write_csv(csv_map["offers"], [asdict(offer) for offer in artifacts.offers])
    write_csv(csv_map["elements"], [asdict(element) for element in artifacts.elements])

    notes_path = csv_map["notes"]
    notes_lines = artifacts.notes or ["No additional notes"]
    notes_path.write_text("\n".join(notes_lines), encoding="utf-8")
    return csv_map


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except RuntimeError:
        loop = asyncio.get_event_loop()
        loop.run_until_complete(main())
"""
