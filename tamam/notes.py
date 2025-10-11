"""Utility for accumulating NOTES.txt content."""
from __future__ import annotations

from pathlib import Path
from typing import Iterable


class NotesLogger:
    def __init__(self) -> None:
        self._lines: list[str] = []

    def add(self, line: str) -> None:
        self._lines.append(line.rstrip())

    def extend(self, lines: Iterable[str]) -> None:
        for line in lines:
            self.add(line)

    def write(self, path: Path) -> Path:
        text = "\n".join(self._lines) + "\n"
        path.write_text(text, encoding="utf-8")
        return path

    def __str__(self) -> str:  # pragma: no cover - convenience
        return "\n".join(self._lines)
