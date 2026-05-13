"""Configuration constants for Archive Search V11.

Company/site-specific defaults should live in ``archive_search_local_config.py``.
That file is intentionally optional so the application can still start with safe,
generic values when a local configuration has not been deployed.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

APP_TITLE = "Archive Search"

DEFAULT_COMPANY_NAME = "Company Name"
DEFAULT_COMPANY_LOGO_PATH = None
DEFAULT_APP_ICON_PATH = None
DEFAULT_HEADER_LOGO_MAX_WIDTH = 360
DEFAULT_HEADER_LOGO_MAX_HEIGHT = 54
DEFAULT_ARCHIVE_SEARCH_PRIMARY_FOLDER = Path(r"C:\Primary Folder")
DEFAULT_ARCHIVE_SEARCH_SECONDARY_FOLDER = Path(r"C:\Secondary Folder")


def _load_local_setting(name: str, default: Any) -> Any:
    """Return a setting from the optional local config module, falling back safely."""
    try:
        import archive_search_local_config as local_config  # type: ignore
    except Exception:
        return default
    return getattr(local_config, name, default)


def _as_optional_path(value: Any) -> Path | None:
    """Convert a local config path-like value to an absolute Path, if supplied."""
    if value in (None, ""):
        return None
    path = Path(str(value)).expanduser()
    if not path.is_absolute():
        path = Path(__file__).resolve().parent / path
    return path.resolve()


def _as_positive_int(value: Any, default: int) -> int:
    """Return a positive integer local setting, falling back if it is invalid."""
    try:
        parsed = int(value)
    except (TypeError, ValueError):
        return default
    return parsed if parsed > 0 else default


COMPANY_NAME = str(_load_local_setting("LOCAL_COMPANY_NAME", DEFAULT_COMPANY_NAME)).strip() or DEFAULT_COMPANY_NAME
COMPANY_LOGO_PATH = _as_optional_path(_load_local_setting("LOCAL_COMPANY_LOGO_PATH", DEFAULT_COMPANY_LOGO_PATH))
APP_ICON_PATH = _as_optional_path(_load_local_setting("LOCAL_APP_ICON_PATH", DEFAULT_APP_ICON_PATH))
HEADER_LOGO_MAX_WIDTH = _as_positive_int(
    _load_local_setting("LOCAL_HEADER_LOGO_MAX_WIDTH", DEFAULT_HEADER_LOGO_MAX_WIDTH),
    DEFAULT_HEADER_LOGO_MAX_WIDTH,
)
HEADER_LOGO_MAX_HEIGHT = _as_positive_int(
    _load_local_setting("LOCAL_HEADER_LOGO_MAX_HEIGHT", DEFAULT_HEADER_LOGO_MAX_HEIGHT),
    DEFAULT_HEADER_LOGO_MAX_HEIGHT,
)
ARCHIVE_SEARCH_PRIMARY_FOLDER = Path(
    _load_local_setting("LOCAL_SEARCH_PRIMARY_FOLDER", DEFAULT_ARCHIVE_SEARCH_PRIMARY_FOLDER)
)
ARCHIVE_SEARCH_SECONDARY_FOLDER = Path(
    _load_local_setting("LOCAL_SEARCH_SECONDARY_FOLDER", DEFAULT_ARCHIVE_SEARCH_SECONDARY_FOLDER)
)

OPENPYXL_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
XLRD_EXTENSIONS = {".xls"}
XLSB_EXTENSIONS = {".xlsb"}
WORKBOOK_EXTENSIONS = OPENPYXL_EXTENSIONS | XLRD_EXTENSIONS | XLSB_EXTENSIONS

WORD_XML_EXTENSIONS = {".docx", ".docm", ".dotx", ".dotm"}
WORD_LEGACY_EXTENSIONS = {".doc"}
WORD_EXTENSIONS = WORD_XML_EXTENSIONS | WORD_LEGACY_EXTENSIONS

PDF_EXTENSIONS = {".pdf"}
PLAIN_TEXT_EXTENSIONS = {".txt"}
TEXT_DOCUMENT_EXTENSIONS = WORD_EXTENSIONS | PDF_EXTENSIONS | PLAIN_TEXT_EXTENSIONS
ALL_INDEXED_EXTENSIONS = WORKBOOK_EXTENSIONS | TEXT_DOCUMENT_EXTENSIONS

INDEX_DB_PREFIX = ".archive_search_"
INDEX_SCHEMA_VERSION = 5
INDEX_BATCH_SIZE = 25

INDEX_RETRY_ATTEMPTS = 4
INDEX_RETRY_DELAY_SECONDS = 0.75
INDEX_STABLE_PROBE_DELAY_SECONDS = 0.25
STATUS_UPDATE_INTERVAL_SECONDS = 0.35
RESULT_BATCH_SIZE = 20
RESULT_LIMIT = 250
DETAIL_SNIPPETS_PER_FILE = None  # Show every matching snippet in the details pane.
DETAIL_PREVIEW_LENGTH = 700

# Approved Primary Colour Palette
BRAND_BLUE = "#015eae"
BRAND_BLACK = "#000000"
BRAND_WHITE = "#ffffff"
BRAND_ORANGE = "#fdb813"
BRAND_GREY = "#cfd0d1"
KLAMP_GREY = "#697f90"
NCRT_BLUE = "#86a2d3"
VISE_BLUE = "#52c7de"
CHUCK_ORANGE = "#f39f63"

# UI tints derived from the approved brand palette. The main surfaces stay
# white/blue/black, with orange reserved for meaningful emphasis.
APP_BACKGROUND = "#f3f5f7"
SURFACE_BLUE = "#f7fbff"
BORDER_GREY = "#d8dcdf"
MUTED_TEXT = "#4f6271"
SOFT_BLUE = "#e8f2fc"
ROW_BLUE = "#eef7ff"
HOVER_BLUE = "#dceefd"
SELECTED_BLUE = "#b7d8f2"

HIGHLIGHT_TAG_NAME = "match_highlight"
HIGHLIGHT_BACKGROUND = BRAND_ORANGE
HIGHLIGHT_FOREGROUND = BRAND_BLACK
