"""In-memory storage for scraper output prior to serialisation."""
from __future__ import annotations

from dataclasses import asdict
from datetime import datetime
import json
import os
from pathlib import Path
from typing import Iterable

import pandas as pd

from .data_models import (
    AddonRecord,
    AvailabilityRecord,
    ElementRecord,
    OfferRecord,
    OptionRecord,
    PageRecord,
    PriceRecord,
    datetime_to_iso,
)


class DataStore:
    """Collects records and writes them to CSV/Excel."""

    def __init__(self, output_dir: str, selectors_dir: str | None = None) -> None:
        self.output_dir = Path(output_dir)
        self.selectors_dir = Path(selectors_dir) if selectors_dir else self.output_dir / "selectors"
        self.pages: list[PageRecord] = []
        self.options: list[OptionRecord] = []
        self.prices: list[PriceRecord] = []
        self.addons: list[AddonRecord] = []
        self.availability: list[AvailabilityRecord] = []
        self.offers: list[OfferRecord] = []
        self.elements: list[ElementRecord] = []

    # Record appenders -----------------------------------------------------
    def add_page(self, record: PageRecord) -> None:
        self.pages.append(record)

    def add_option(self, record: OptionRecord) -> None:
        self.options.append(record)

    def add_price(self, record: PriceRecord) -> None:
        self.prices.append(record)

    def add_addon(self, record: AddonRecord) -> None:
        self.addons.append(record)

    def add_availability(self, record: AvailabilityRecord) -> None:
        self.availability.append(record)

    def add_offer(self, record: OfferRecord) -> None:
        self.offers.append(record)

    def add_element(self, record: ElementRecord) -> None:
        self.elements.append(record)

    # Export helpers -------------------------------------------------------
    def _write_dataframe(self, df: pd.DataFrame, filename: str) -> Path:
        path = self.output_dir / filename
        df.to_csv(path, index=False)
        return path

    @staticmethod
    def _prices_to_dataframe(records: Iterable[PriceRecord]) -> pd.DataFrame:
        rows: list[dict[str, object]] = []
        for record in records:
            base = {
                "page_id": record.page_id,
                "service_name": record.service_name,
                "model_type": record.model_type,
                "offer_state": record.offer_state,
                "offer_name": record.offer_name,
                "discount_type": record.discount_type,
                "discount_value": record.discount_value,
                "base_price": record.base_price,
                "fees_total": record.fees_total,
                "vat_amount": record.vat_amount,
                "vat_percent": record.vat_percent,
                "grand_total": record.grand_total,
                "currency": record.currency,
                "pricing_source_url": record.pricing_source_url,
                "collected_at": datetime_to_iso(record.collected_at),
                "qa_flags": ",".join(record.qa_flags),
                "screenshot_ref": record.screenshot_ref,
            }
            rows.append({**base, **record.dimensions})
        return pd.DataFrame(rows)

    # DataFrame getters ---------------------------------------------------
    def pages_dataframe(self) -> pd.DataFrame:
        if self.pages:
            return pd.DataFrame([asdict(r) for r in self.pages])
        return pd.DataFrame(columns=[
            "page_id",
            "parent_page_id",
            "level",
            "url",
            "page_title_text",
            "detected_model_type",
            "has_addons",
            "has_slots",
            "notes",
        ])

    def options_dataframe(self) -> pd.DataFrame:
        if self.options:
            return pd.DataFrame([asdict(r) for r in self.options])
        return pd.DataFrame(columns=[
            "page_id",
            "control_group",
            "option_code",
            "option_label",
            "option_type",
            "is_default",
            "sort_index",
        ])

    def prices_dataframe(self) -> pd.DataFrame:
        df = self._prices_to_dataframe(self.prices)
        if df.empty:
            df = pd.DataFrame(columns=[
                "page_id",
                "service_name",
                "model_type",
                "offer_state",
                "offer_name",
                "discount_type",
                "discount_value",
                "base_price",
                "fees_total",
                "vat_amount",
                "vat_percent",
                "grand_total",
                "currency",
                "pricing_source_url",
                "collected_at",
                "qa_flags",
                "screenshot_ref",
            ])
        return df

    def addons_dataframe(self) -> pd.DataFrame:
        if self.addons:
            return pd.DataFrame([asdict(r) for r in self.addons])
        return pd.DataFrame(columns=[
            "page_id",
            "addon_name",
            "addon_description",
            "addon_price",
            "currency",
            "is_recommended",
            "is_required",
            "seen_on_step",
            "pricing_source_url",
            "collected_at",
        ])

    def availability_dataframe(self) -> pd.DataFrame:
        if self.availability:
            return pd.DataFrame([asdict(r) for r in self.availability])
        return pd.DataFrame(columns=[
            "page_id",
            "date_str",
            "time_slot_label",
            "slot_status",
            "lead_time_days",
            "collected_at",
        ])

    def offers_dataframe(self) -> pd.DataFrame:
        if self.offers:
            return pd.DataFrame([asdict(r) for r in self.offers])
        return pd.DataFrame(columns=[
            "page_id",
            "offer_name",
            "offer_badge_text",
            "description",
            "eligibility",
            "discount_type",
            "discount_value",
            "stacking_rules",
            "collected_at",
        ])

    def elements_dataframe(self) -> pd.DataFrame:
        if self.elements:
            return pd.DataFrame([asdict(r) for r in self.elements])
        return pd.DataFrame(columns=[
            "page_id",
            "selector",
            "element_role",
            "raw_text",
            "normalized_text",
        ])

    def export_csvs(self) -> dict[str, Path]:
        self.output_dir.mkdir(parents=True, exist_ok=True)
        paths: dict[str, Path] = {}
        paths["pages"] = self._write_dataframe(self.pages_dataframe(), "pages.csv")
        paths["options"] = self._write_dataframe(self.options_dataframe(), "options.csv")
        paths["prices_raw"] = self._write_dataframe(self.prices_dataframe(), "prices_raw.csv")
        paths["addons"] = self._write_dataframe(self.addons_dataframe(), "addons.csv")
        paths["availability"] = self._write_dataframe(self.availability_dataframe(), "availability.csv")
        paths["offers"] = self._write_dataframe(self.offers_dataframe(), "offers.csv")
        paths["elements"] = self._write_dataframe(self.elements_dataframe(), "elements.csv")

        return paths

    def export_selectors(self, page_id: str, selectors: dict[str, list[str]]) -> Path:
        directory = self.selectors_dir
        directory.mkdir(parents=True, exist_ok=True)
        path = directory / f"{page_id}.json"
        with path.open("w", encoding="utf-8") as handle:
            json.dump(selectors, handle, indent=2, ensure_ascii=False)
        return path

