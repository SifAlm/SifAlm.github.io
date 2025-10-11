"""Dataclasses representing the structured outputs of the scraper."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Optional


ISO_FORMAT = "%Y-%m-%dT%H:%M:%SZ"


@dataclass(slots=True)
class PageRecord:
    page_id: str
    parent_page_id: str | None
    level: str
    url: str
    page_title_text: str
    detected_model_type: str
    has_addons: bool
    has_slots: bool
    notes: str


@dataclass(slots=True)
class OptionRecord:
    page_id: str
    control_group: str
    option_code: str
    option_label: str
    option_type: str
    is_default: bool
    sort_index: int


@dataclass(slots=True)
class PriceRecord:
    page_id: str
    service_name: str
    model_type: str
    dimensions: dict[str, Any]
    offer_state: str
    offer_name: str | None
    discount_type: str | None
    discount_value: float | None
    base_price: float | None
    fees_total: float | None
    vat_amount: float | None
    vat_percent: float | None
    grand_total: float | None
    currency: str
    pricing_source_url: str
    collected_at: datetime
    qa_flags: list[str] = field(default_factory=list)
    screenshot_ref: str | None = None


@dataclass(slots=True)
class AddonRecord:
    page_id: str
    addon_name: str
    addon_description: str
    addon_price: float | None
    currency: str | None
    is_recommended: bool
    is_required: bool
    seen_on_step: str
    pricing_source_url: str
    collected_at: datetime


@dataclass(slots=True)
class AvailabilityRecord:
    page_id: str
    date_str: str
    time_slot_label: str
    slot_status: str
    lead_time_days: int
    collected_at: datetime


@dataclass(slots=True)
class OfferRecord:
    page_id: str
    offer_name: str
    offer_badge_text: str
    description: str
    eligibility: str
    discount_type: str | None
    discount_value: float | None
    stacking_rules: str
    collected_at: datetime


@dataclass(slots=True)
class ElementRecord:
    page_id: str
    selector: str
    element_role: str
    raw_text: str
    normalized_text: str


def datetime_to_iso(value: datetime | None) -> str | None:
    if value is None:
        return None
    return value.strftime(ISO_FORMAT)
