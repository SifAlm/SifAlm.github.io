"""Helper utilities for scraping, parsing and UUID generation."""
from __future__ import annotations

from datetime import datetime, timezone
import random
import time
import uuid
from typing import Iterable

from bs4 import BeautifulSoup


UUID_NAMESPACE = uuid.NAMESPACE_URL


def uuid_from_url(url: str) -> str:
    return str(uuid.uuid5(UUID_NAMESPACE, url))


def utc_now() -> datetime:
    return datetime.utcnow().replace(tzinfo=timezone.utc)


def polite_wait(min_ms: int, max_ms: int) -> None:
    delay = random.uniform(min_ms / 1000.0, max_ms / 1000.0)
    time.sleep(delay)


def extract_links(html: str, base_url: str) -> set[str]:
    soup = BeautifulSoup(html, "html.parser")
    links: set[str] = set()
    for tag in soup.find_all("a", href=True):
        href = tag["href"].strip()
        if href.startswith("http") and "justlife.com" not in href:
            continue
        if href.startswith("//"):
            href = "https:" + href
        if href.startswith("/"):
            href = base_url.rstrip("/") + href
        if href.startswith(base_url):
            links.add(href.split("#", 1)[0])
    return links


def normalise_text(value: str | None) -> str:
    if not value:
        return ""
    return " ".join(value.split())


def ensure_unique(seq: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    unique: list[str] = []
    for item in seq:
        if item not in seen:
            seen.add(item)
            unique.append(item)
    return unique
