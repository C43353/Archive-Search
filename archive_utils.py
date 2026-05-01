"""General helpers and lightweight formatting utilities for Archive Search V11."""

from __future__ import annotations

import datetime as dt
import math
import sys
from pathlib import Path
from typing import List, Optional, Sequence


def get_app_folder() -> Path:
    """Return the folder that contains the running program."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def compact_whitespace(value) -> str:
    """Collapse repeated whitespace and convert the incoming value to a clean string."""
    if value is None:
        return ""
    return " ".join(str(value).split())


def normalize_text(value) -> str:
    """Convert incoming values into a lower-case search-friendly representation."""
    if value is None:
        return ""
    if isinstance(value, dt.datetime):
        return value.strftime("%d/%m/%Y").lower()
    if isinstance(value, dt.date):
        return value.strftime("%d/%m/%Y").lower()
    if isinstance(value, dt.time):
        return value.strftime("%H:%M").lower()
    return compact_whitespace(value).lower()


def display_text(value) -> str:
    """Convert a raw cell or document value into a display-safe string."""
    if value is None:
        return ""
    if isinstance(value, float):
        if math.isfinite(value) and value.is_integer():
            return str(int(value))
        return str(value)
    if isinstance(value, dt.datetime):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, dt.date):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, dt.time):
        return value.strftime("%H:%M")
    return compact_whitespace(value)


def build_highlight_terms(search_terms: Sequence[str]) -> List[str]:
    """Build a de-duplicated list of highlight terms from the current search input."""
    seen = set()
    items: List[str] = []
    for term in sorted(search_terms, key=len, reverse=True):
        whole = term.strip().lower()
        if whole and whole not in seen:
            items.append(whole)
            seen.add(whole)
        parts = [part.strip().lower() for part in whole.split() if part.strip()]
        for part in sorted(parts, key=len, reverse=True):
            if part and part not in seen:
                items.append(part)
                seen.add(part)
    return items


def pluralize(count: int, singular: str, plural: Optional[str] = None) -> str:
    """Return a singular or plural label that matches the supplied count."""
    return singular if count == 1 else (plural or singular + "s")


def should_ignore_filename(filename: str) -> bool:
    """Exclude dot-prefixed names and temporary Office lock files."""
    return filename.startswith(".") or filename.startswith("~$")


def truncate_text(text: str, limit: int) -> str:
    """Trim long preview text so the list and detail panes remain readable."""
    value = compact_whitespace(text)
    if len(value) <= limit:
        return value
    return value[: max(0, limit - 3)].rstrip() + "..."


def infer_type_label(file_type: str) -> str:
    """Convert an internal file-type key into the short label shown in the UI."""
    return {"excel": "Excel", "word": "Word", "pdf": "PDF", "text": "Text"}.get(file_type, file_type.title())


def format_utc_iso_for_display(value: str) -> str:
    """Convert a stored UTC ISO timestamp into a concise local display string."""
    if not value:
        return "Unknown"
    try:
        parsed = dt.datetime.fromisoformat(value)
    except Exception:
        return value
    try:
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=dt.timezone.utc)
        local_value = parsed.astimezone()
        return local_value.strftime("%d/%m/%Y %H:%M:%S")
    except Exception:
        return value


def utc_now_iso() -> str:
    """Return the current UTC time as an ISO 8601 string."""
    return dt.datetime.now(dt.timezone.utc).isoformat()
