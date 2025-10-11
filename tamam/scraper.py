"""Core scraping logic for the Tamam Justlife pipeline."""
from __future__ import annotations

import math
from collections import deque
from datetime import datetime
from pathlib import Path
import re
from typing import Any, Iterable

from bs4 import BeautifulSoup
from playwright.async_api import Browser, BrowserContext, Page, async_playwright

from .config import ScraperConfig
from .data_models import (
    AddonRecord,
    AvailabilityRecord,
    ElementRecord,
    OfferRecord,
    OptionRecord,
    PageRecord,
    PriceRecord,
)
from .notes import NotesLogger
from .storage import DataStore
from .utils import ensure_unique, extract_links, normalise_text, polite_wait, utc_now, uuid_from_url


PRICE_PATTERN = re.compile(r"AED\s*([0-9,]+(?:\.[0-9]+)?)", re.IGNORECASE)
VAT_PATTERN = re.compile(r"VAT\s*(?:\(|:)?\s*([0-9]{1,2})%", re.IGNORECASE)


class JustlifeScraper:
    """Encapsulates the three phases of the scraping workflow."""

    def __init__(self, config: ScraperConfig, datastore: DataStore, notes: NotesLogger) -> None:
        self.config = config
        self.datastore = datastore
        self.notes = notes
        self.service_pages: list[PageRecord] = []

    # ------------------------------------------------------------------
    async def run(self) -> None:
        self.config.ensure_directories()

        if not self.config.online:
            self.notes.add("ONLINE flag disabled – running in offline reconnaissance mode.")
            self.datastore.export_csvs()
            return

        async with async_playwright() as playwright:
            browser = await playwright.chromium.launch(headless=not self.config.debug)
            context = await browser.new_context(
                locale="en-US",
                timezone_id=self.config.timezone,
                java_script_enabled=True,
                device_scale_factor=1.0,
            )
            page = await context.new_page()
            await self._phase_a_recon(page)
            await self._phase_b_extract(context)
            await browser.close()

    # ------------------------------------------------------------------
    async def _phase_a_recon(self, page: Page) -> None:
        base_url = self.config.base_url.rstrip("/")
        queue = deque([base_url])
        visited: set[str] = set()
        parent_map: dict[str, str | None] = {base_url: None}

        while queue:
            url = queue.popleft()
            if url in visited:
                continue
            visited.add(url)
            try:
                await page.goto(url, wait_until="domcontentloaded", timeout=60000)
                polite_wait(self.config.polite_min_wait_ms, self.config.polite_max_wait_ms)
            except Exception as exc:  # pragma: no cover - network safety
                self.notes.add(f"Failed to load {url}: {exc}")
                await self._dump_html(page, url)
                continue

            html = await page.content()
            title_text = await page.title()
            soup = BeautifulSoup(html, "html.parser")

            model_type = self._detect_model_type(soup)
            has_addons = bool(soup.select("[data-testid*=addon], .addon"))
            has_slots = bool(re.search(r"slot|availability", html, re.IGNORECASE))
            level = self._classify_level(url, base_url)
            notes = self._summarise_model(model_type)
            page_id = uuid_from_url(url)
            parent_page_id = None
            parent = parent_map.get(url)
            if parent:
                parent_page_id = uuid_from_url(parent)

            page_record = PageRecord(
                page_id=page_id,
                parent_page_id=parent_page_id,
                level=level,
                url=url,
                page_title_text=title_text,
                detected_model_type=model_type,
                has_addons=has_addons,
                has_slots=has_slots,
                notes=notes,
            )
            self.datastore.add_page(page_record)
            if level in {"service", "flow", "checkout"}:
                self.service_pages.append(page_record)
                if level == "service":
                    self.notes.add(f"Service '{title_text}' classified as {model_type} via static heuristics.")

            selectors = self._discover_selectors(soup)
            self.datastore.export_selectors(page_id, selectors)

            links = extract_links(html, base_url)
            for link in links:
                if link not in visited and link.startswith(base_url):
                    parent_map.setdefault(link, url)
                    queue.append(link)

    # ------------------------------------------------------------------
    async def _phase_b_extract(self, context: BrowserContext) -> None:
        for page_record in self.service_pages:
            page = await context.new_page()
            try:
                await page.goto(page_record.url, wait_until="domcontentloaded", timeout=60000)
                polite_wait(self.config.polite_min_wait_ms, self.config.polite_max_wait_ms)
                await self._handle_location_gate(page)
                await self._collect_service_data(page_record, page)
            except Exception as exc:  # pragma: no cover - defensive
                self.notes.add(f"Extraction failed for {page_record.url}: {exc}")
                await self._dump_html(page, page_record.url)
            finally:
                await page.close()

    # ------------------------------------------------------------------
    def _detect_model_type(self, soup: BeautifulSoup) -> str:
        text = soup.get_text(" ", strip=True).lower()
        if any(keyword in text for keyword in ["hours", "hourly", "cleaners"]):
            return "HOURS_PROS"
        if any(keyword in text for keyword in ["bedroom", "bathroom", "room"]):
            return "ROOM_MATERIAL"
        if "appliance" in text:
            return "APPLIANCE_COMBO"
        if "sqft" in text or "square" in text:
            return "AREA_BASED"
        if any(keyword in text for keyword in ["package", "basic", "premium"]):
            return "FLAT_PACKAGE"
        if any(keyword in text for keyword in ["quantity", "units", "pieces"]):
            return "UNIT_QUANTITY"
        return "OTHER"

    def _summarise_model(self, model_type: str) -> str:
        mapping = {
            "HOURS_PROS": "Detected hours/professionals controls via keyword scan.",
            "ROOM_MATERIAL": "Detected bedroom/bathroom language indicating room count model.",
            "APPLIANCE_COMBO": "Detected appliance bundle language.",
            "AREA_BASED": "Detected sqft/area mentions.",
            "FLAT_PACKAGE": "Detected named packages (basic/premium/etc.).",
            "UNIT_QUANTITY": "Detected quantity or unit based pricing cues.",
            "OTHER": "Model undetermined from static content; requires manual follow-up.",
        }
        return mapping.get(model_type, "")

    def _classify_level(self, url: str, base_url: str) -> str:
        if url.rstrip("/") == base_url.rstrip("/"):
            return "home"
        relative = url.replace(base_url, "", 1).strip("/")
        depth = len([segment for segment in relative.split("/") if segment])
        if depth == 0:
            return "category"
        if depth == 1:
            return "subcategory"
        if depth == 2:
            return "service"
        if depth == 3:
            return "flow"
        return "checkout"

    def _discover_selectors(self, soup: BeautifulSoup) -> dict[str, list[str]]:
        selectors: dict[str, list[str]] = {
            "title": [],
            "price": [],
            "option": [],
            "addon": [],
            "slot": [],
            "fee": [],
        }
        for header in soup.select("h1, h2, [data-testid*=title]"):
            if header.get_text(strip=True):
                selector = header.name
                selectors["title"].append(selector)
        for price_el in soup.select("[class*=price], [data-testid*=price], span:contains('AED')"):
            selector = price_el.name
            selectors["price"].append(selector)
        for option_el in soup.select("button, [role=option], select"):
            selectors["option"].append(option_el.name)
        for addon_el in soup.select("[class*=addon], [data-testid*=addon]"):
            selectors["addon"].append(addon_el.name)
        for slot_el in soup.select("[class*=slot], [data-testid*=slot]"):
            selectors["slot"].append(slot_el.name)
        for fee_el in soup.select("[class*=fee], [data-testid*=fee], text:contains('VAT')"):
            selectors["fee"].append(fee_el.name if hasattr(fee_el, "name") else "text")
        return {key: ensure_unique(values) for key, values in selectors.items() if values}

    async def _handle_location_gate(self, page: Page) -> None:
        try:
            location_gate = page.locator("text=Select your city")
            if await location_gate.count() > 0:
                dubai_option = page.locator("role=button[name='Dubai']")
                if await dubai_option.count() == 0:
                    dubai_option = page.locator("text=Dubai")
                if await dubai_option.count() > 0:
                    await dubai_option.first.click()
                    polite_wait(self.config.polite_min_wait_ms, self.config.polite_max_wait_ms)
        except Exception:  # pragma: no cover - gating is best-effort
            return

    async def _collect_service_data(self, page_record: PageRecord, page: Page) -> None:
        controls = await self._discover_controls(page_record, page)
        for control in controls:
            self.datastore.add_option(control)

        combinations = self._enumerate_combinations(controls)
        if not combinations:
            combinations = [({}, [])]

        for dims, actions in combinations[: self.config.max_combos]:
            for action in actions:
                await action()
                polite_wait(self.config.polite_min_wait_ms, self.config.polite_max_wait_ms)
            price_info = await self._extract_price_breakdown(page_record, page)
            qa_flags: list[str] = []
            if price_info:
                base, fees, vat_amount, vat_percent, total, currency = price_info
                if base is None:
                    qa_flags.append("BASE_UNKNOWN")
                if all(value is not None for value in (base, fees, vat_amount, total)):
                    if not math.isclose((base or 0) + (fees or 0) + (vat_amount or 0), total or 0, rel_tol=0.01, abs_tol=1.0):
                        qa_flags.append("ROUNDING_WARNING")
            else:
                base = fees = vat_amount = vat_percent = total = None
                currency = "AED"
                qa_flags.append("MISSING_PRICE")

            price_record = PriceRecord(
                page_id=page_record.page_id,
                service_name=page_record.page_title_text,
                model_type=page_record.detected_model_type,
                dimensions=dims,
                offer_state="OFF",
                offer_name=None,
                discount_type=None,
                discount_value=None,
                base_price=base,
                fees_total=fees,
                vat_amount=vat_amount,
                vat_percent=vat_percent,
                grand_total=total,
                currency=currency,
                pricing_source_url=page_record.url,
                collected_at=utc_now(),
                qa_flags=qa_flags,
            )
            self.datastore.add_price(price_record)

        await self._collect_addons(page_record, page)
        await self._collect_availability(page_record, page)
        await self._collect_offers(page_record, page)

    async def _discover_controls(self, page_record: PageRecord, page: Page) -> list[OptionRecord]:
        controls: list[OptionRecord] = []
        groups = [
            ("Hours", "hours"),
            ("Professionals", "pros"),
            ("Bedrooms", "bedrooms"),
            ("Bathrooms", "bathrooms"),
            ("Packages", "package"),
        ]
        index_counter = 0
        for label_text, group_name in groups:
            locator = page.locator(f"text={label_text}")
            if await locator.count() == 0:
                continue
            group_root = locator.nth(0)
            container = group_root.locator("xpath=..")
            options = container.locator("button, [role=option], option")
            option_count = await options.count()
            for idx in range(option_count):
                option = options.nth(idx)
                option_label = normalise_text(await option.inner_text())
                option_value = await option.get_attribute("value") or option_label
                is_default = await option.get_attribute("aria-pressed") == "true"
                control = OptionRecord(
                    page_id=page_record.page_id,
                    control_group=label_text,
                    option_code=option_value,
                    option_label=option_label,
                    option_type="button",
                    is_default=is_default,
                    sort_index=index_counter,
                )
                controls.append(control)
                index_counter += 1
        return controls

    def _enumerate_combinations(self, controls: list[OptionRecord]):
        # Placeholder – combinations require page-specific logic.
        return []

    async def _extract_price_breakdown(self, page_record: PageRecord, page: Page):
        candidates = [
            "text=/AED/",
            "[data-testid*=price]",
            "[class*=price]",
        ]
        for selector in candidates:
            locator = page.locator(selector)
            if await locator.count() == 0:
                continue
            text = normalise_text(await locator.first.inner_text())
            self.datastore.add_element(
                ElementRecord(
                    page_id=page_record.page_id,
                    selector=selector,
                    element_role="price",
                    raw_text=text,
                    normalized_text=text,
                )
            )
            if match := PRICE_PATTERN.search(text):
                total = float(match.group(1).replace(",", ""))
                vat_match = VAT_PATTERN.search(text)
                vat_percent = float(vat_match.group(1)) if vat_match else None
                return (None, None, None, vat_percent, total, "AED")
        return None

    async def _collect_addons(self, page_record: PageRecord, page: Page) -> None:
        addon_locator = page.locator("[data-testid*=addon], .addon")
        count = await addon_locator.count()
        for idx in range(count):
            element = addon_locator.nth(idx)
            text = normalise_text(await element.inner_text())
            price_match = PRICE_PATTERN.search(text)
            price = float(price_match.group(1).replace(",", "")) if price_match else None
            addon = AddonRecord(
                page_id=page_record.page_id,
                addon_name=text.split("AED")[0].strip(),
                addon_description=text,
                addon_price=price,
                currency="AED" if price is not None else None,
                is_recommended="recommended" in text.lower(),
                is_required="required" in text.lower(),
                seen_on_step="pre-cart",
                pricing_source_url=page_record.url,
                collected_at=utc_now(),
            )
            self.datastore.add_addon(addon)

    async def _collect_availability(self, page_record: PageRecord, page: Page) -> None:
        slots_locator = page.locator("[data-testid*=slot], [class*=slot]")
        count = await slots_locator.count()
        today = datetime.utcnow().date()
        for idx in range(min(count, 50)):
            element = slots_locator.nth(idx)
            text = normalise_text(await element.inner_text())
            date_match = re.search(r"(\d{4}-\d{2}-\d{2})", text)
            slot_label = text
            if date_match:
                date_str = date_match.group(1)
            else:
                date_str = today.isoformat()
            record = AvailabilityRecord(
                page_id=page_record.page_id,
                date_str=date_str,
                time_slot_label=slot_label,
                slot_status="available",
                lead_time_days=(datetime.fromisoformat(date_str) - datetime.combine(today, datetime.min.time())).days,
                collected_at=utc_now(),
            )
            self.datastore.add_availability(record)

    async def _collect_offers(self, page_record: PageRecord, page: Page) -> None:
        offer_locator = page.locator("[class*=offer], [data-testid*=offer], text=/off/i")
        count = await offer_locator.count()
        for idx in range(count):
            element = offer_locator.nth(idx)
            text = normalise_text(await element.inner_text())
            discount_match = re.search(r"(\d+)%", text)
            offer = OfferRecord(
                page_id=page_record.page_id,
                offer_name=text.split(".")[0],
                offer_badge_text=text[:30],
                description=text,
                eligibility="",
                discount_type="percent" if discount_match else None,
                discount_value=float(discount_match.group(1)) if discount_match else None,
                stacking_rules="",
                collected_at=utc_now(),
            )
            self.datastore.add_offer(offer)

    async def _dump_html(self, page: Page, url: str) -> None:
        filename = uuid_from_url(url) + ".html"
        dump_dir = Path(self.config.html_dump_dir)
        dump_dir.mkdir(parents=True, exist_ok=True)
        path = dump_dir / filename
        path.parent.mkdir(parents=True, exist_ok=True)
        try:
            content = await page.content()
        except Exception:
            content = ""
        path.write_text(content, encoding="utf-8")

