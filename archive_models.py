"""Dataclasses shared across Archive Search V11."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Tuple


@dataclass(frozen=True)
class SearchRoot:
    """One search root configured in the UI."""

    label: str
    path: Path
    include_subfolders: bool


@dataclass(frozen=True)
class MatchSnippet:
    """A compact preview for one match inside a file."""

    location_label: str
    preview: str


@dataclass(frozen=True)
class SearchResult:
    """One grouped result rendered in the result list and details pane."""

    file_type: str
    document_name: str
    path: str
    root_label: str
    match_count: int
    snippets: Tuple[MatchSnippet, ...]
    root_path: Optional[str] = None
    open_sheet: Optional[str] = None
    open_row_number: Optional[int] = None
    first_location_label: Optional[str] = None
