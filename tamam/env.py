"""Environment setup utilities for the Tamam Justlife scraper.

This module provides a helper that ensures the runtime has the Python
packages and Playwright browser binaries that the pipeline depends on.
It is intentionally conservative: if dependencies are already installed,
it avoids redundant work, but when missing it will install them in-place.
"""
from __future__ import annotations

import importlib
import os
import subprocess
import sys
from typing import Iterable


DEFAULT_PYTHON_PACKAGES: tuple[str, ...] = (
    "pandas",
    "openpyxl",
    "beautifulsoup4",
    "playwright",
    "pyyaml",
    "numpy",
)


def _run_subprocess(command: list[str], *, env: dict[str, str] | None = None) -> None:
    """Run a subprocess command and stream output live."""
    process_env = os.environ.copy()
    if env:
        process_env.update(env)

    print("[setup_env] Running:", " ".join(command))
    completed = subprocess.run(command, check=False, env=process_env)
    if completed.returncode != 0:
        raise RuntimeError(
            f"Command {' '.join(command)} failed with exit code {completed.returncode}"
        )


def _ensure_packages(packages: Iterable[str]) -> None:
    """Install any of the required packages that are missing."""
    missing: list[str] = []
    for package in packages:
        try:
            importlib.import_module(package)
        except ModuleNotFoundError:
            missing.append(package)
    if not missing:
        return

    command = [sys.executable, "-m", "pip", "install", "--upgrade", *missing]
    _run_subprocess(command)


def _ensure_playwright_browser() -> None:
    """Install the Chromium browser for Playwright if needed."""
    cache_flag = os.environ.get("PLAYWRIGHT_BROWSERS_PATH")
    command = [sys.executable, "-m", "playwright", "install", "chromium"]
    if cache_flag:
        env = {"PLAYWRIGHT_BROWSERS_PATH": cache_flag}
    else:
        env = {}
    _run_subprocess(command, env=env)


def setup_env() -> None:
    """Install Python packages and Playwright browsers.

    The function is idempotent and safe to call on every run. A DEBUG flag
    can be toggled via the environment to emit additional diagnostics.
    """
    debug = os.environ.get("DEBUG") == "1"
    if debug:
        print("[setup_env] DEBUG mode enabled")

    _ensure_packages(DEFAULT_PYTHON_PACKAGES)

    try:
        importlib.import_module("playwright.async_api")
    except ModuleNotFoundError as exc:  # pragma: no cover - defensive branch
        raise RuntimeError(
            "Playwright installation failed even after pip install"
        ) from exc

    try:
        _ensure_playwright_browser()
    except RuntimeError as exc:
        if debug:
            print("[setup_env] Playwright browser install failed:", exc)
        raise

    if debug:
        print("[setup_env] Environment ready")

