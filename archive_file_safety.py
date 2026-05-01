"""Low-level safety checks and signature-based file-type detection."""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Optional

from archive_config import OPENPYXL_EXTENSIONS, PLAIN_TEXT_EXTENSIONS, WORD_LEGACY_EXTENSIONS, XLRD_EXTENSIONS

OLE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ZIP_SIGNATURES = (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")
PDF_SIGNATURE = b"%PDF-"
PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"
JPEG_SIGNATURE = b"\xff\xd8\xff"
GIF_SIGNATURES = (b"GIF87a", b"GIF89a")
BMP_SIGNATURE = b"BM"
TIFF_SIGNATURES = (b"II*\x00", b"MM\x00*")
SQLITE_SIGNATURE = b"SQLite format 3\x00"
SIGNATURE_BYTES = 4096
LEGACY_SCAN_BYTES = 262144


def can_open_binary_readonly(file_path: Path) -> bool:
    """Return True when the file can be opened read-only in binary mode."""
    try:
        with file_path.open("rb"):
            return True
    except OSError:
        return False


def ensure_binary_readonly_access(file_path: Path) -> None:
    """Raise an error when the target cannot be safely opened read-only."""
    if file_path.is_symlink():
        raise RuntimeError(f"Refusing to open symbolic link: {file_path}")
    if not file_path.is_file():
        raise RuntimeError(f"Not a regular file: {file_path}")
    try:
        with file_path.open("rb"):
            return
    except OSError as exc:
        raise RuntimeError(f"Could not open read-only: {file_path}: {exc}") from exc


def read_prefix(file_path: Path, size: int = SIGNATURE_BYTES) -> bytes:
    """Read a small read-only prefix used for signature sniffing."""
    with file_path.open("rb") as handle:
        return handle.read(size)


def _scan_legacy_office_markers(file_path: Path) -> bytes:
    """Read a larger prefix to look for common legacy Office stream names."""
    with file_path.open("rb") as handle:
        return handle.read(LEGACY_SCAN_BYTES)


def _looks_like_plain_text(blob: bytes) -> bool:
    """Heuristic used only for clearer troubleshooting messages."""
    if not blob:
        return True
    sample = blob[:SIGNATURE_BYTES]
    if b"\x00" in sample:
        return False
    control_bytes = sum(1 for byte in sample if byte < 32 and byte not in (9, 10, 12, 13))
    return (control_bytes / max(1, len(sample))) < 0.05


def _detect_ooxml_type(file_path: Path) -> Optional[str]:
    """Detect modern Office document types by inspecting ZIP members."""
    try:
        with zipfile.ZipFile(file_path, "r") as archive:
            names = archive.namelist()
    except (OSError, zipfile.BadZipFile, zipfile.LargeZipFile):
        return None

    if not names:
        return None
    if not any(name == "[Content_Types].xml" for name in names):
        return None
    if any(name.startswith("xl/") for name in names):
        return "excel"
    if any(name.startswith("word/") for name in names):
        return "word"
    return None


def _detect_legacy_ole_type(file_path: Path, *, allow_suffix_fallback: bool = True) -> Optional[str]:
    """Detect legacy Word/Excel documents from OLE signatures and markers."""
    try:
        blob = _scan_legacy_office_markers(file_path)
    except OSError:
        return None

    if b"W\x00o\x00r\x00d\x00D\x00o\x00c\x00u\x00m\x00e\x00n\x00t" in blob:
        return "word"
    if b"W\x00o\x00r\x00k\x00b\x00o\x00o\x00k" in blob or b"B\x00o\x00o\x00k" in blob:
        return "excel"

    if allow_suffix_fallback:
        suffix = file_path.suffix.lower()
        if suffix in WORD_LEGACY_EXTENSIONS:
            return "word"
        if suffix in XLRD_EXTENSIONS:
            return "excel"
    return None


def detect_supported_file_type(file_path: Path) -> Optional[str]:
    """Return the supported internal type after validating the file signature."""
    ensure_binary_readonly_access(file_path)
    prefix = read_prefix(file_path)

    if prefix.startswith(PDF_SIGNATURE):
        return "pdf"

    if any(prefix.startswith(signature) for signature in ZIP_SIGNATURES):
        return _detect_ooxml_type(file_path)

    if prefix.startswith(OLE_SIGNATURE):
        return _detect_legacy_ole_type(file_path)

    suffix = file_path.suffix.lower()
    if suffix in OPENPYXL_EXTENSIONS:
        return _detect_ooxml_type(file_path)
    if suffix in PLAIN_TEXT_EXTENSIONS:
        return "text"
    return None


def detect_file_content_kind(file_path: Path) -> Optional[str]:
    """Return the apparent content kind without trusting the filename extension."""
    ensure_binary_readonly_access(file_path)
    prefix = read_prefix(file_path)

    if not prefix:
        return "empty"
    if prefix.startswith(PDF_SIGNATURE):
        return "pdf"
    if prefix.startswith(SQLITE_SIGNATURE):
        return "sqlite"
    if prefix.startswith(PNG_SIGNATURE):
        return "png"
    if prefix.startswith(JPEG_SIGNATURE):
        return "jpeg"
    if any(prefix.startswith(signature) for signature in GIF_SIGNATURES):
        return "gif"
    if prefix.startswith(BMP_SIGNATURE):
        return "bmp"
    if any(prefix.startswith(signature) for signature in TIFF_SIGNATURES):
        return "tiff"
    if any(prefix.startswith(signature) for signature in ZIP_SIGNATURES):
        return _detect_ooxml_type(file_path) or "zip"
    if prefix.startswith(OLE_SIGNATURE):
        return _detect_legacy_ole_type(file_path, allow_suffix_fallback=False) or "ole"
    if _looks_like_plain_text(prefix):
        return "text"
    return None


def describe_file_content(file_path: Path) -> str:
    """Describe what a file appears to be based on content, not extension."""
    try:
        kind = detect_file_content_kind(file_path)
    except Exception as exc:
        return f"could not be checked ({exc})"

    return {
        "empty": "empty file",
        "pdf": "PDF document",
        "sqlite": "SQLite database",
        "png": "PNG image",
        "jpeg": "JPEG image",
        "gif": "GIF image",
        "bmp": "BMP image",
        "tiff": "TIFF image",
        "zip": "ZIP archive, not a supported Excel or Word document",
        "ole": "legacy Microsoft Office/OLE file, but not clearly a supported Word or Excel document",
        "excel": "Excel workbook",
        "word": "Word document",
        "text": "plain text",
    }.get(kind, "unknown binary file")
