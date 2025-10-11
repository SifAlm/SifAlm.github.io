"""Configuration objects for the Tamam Justlife scraper."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import timedelta
import os


@dataclass(slots=True)
class ScraperConfig:
    """High level configuration values for the scraping pipeline."""

    base_url: str = field(default_factory=lambda: os.environ.get("BASE", "https://www.justlife.com/en-AE"))
    output_dir: str = field(default_factory=lambda: os.environ.get("OUTPUT_DIR", "."))
    debug: bool = field(default_factory=lambda: os.environ.get("DEBUG") == "1")
    max_combos: int = field(default_factory=lambda: int(os.environ.get("MAX_COMBOS", "40")))
    availability_days: int = field(default_factory=lambda: int(os.environ.get("DAYS", "14")))
    polite_min_wait_ms: int = 300
    polite_max_wait_ms: int = 900
    timezone: str = "Asia/Dubai"
    online: bool = field(default_factory=lambda: os.environ.get("ONLINE", "true").lower() != "false")
    selectors_dir: str = "selectors"
    html_dump_dir: str = "html_dump"
    screenshots_dir: str = "screenshots"

    def ensure_directories(self) -> None:
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.selectors_dir, exist_ok=True)
        os.makedirs(self.html_dump_dir, exist_ok=True)
        os.makedirs(self.screenshots_dir, exist_ok=True)


DEFAULT_CONFIG = ScraperConfig()
