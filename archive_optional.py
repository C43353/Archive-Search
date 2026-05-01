"""Optional third-party and platform dependencies for Archive Search V11."""

from __future__ import annotations

try:
    from docx import Document as DocxDocument
    HAS_DOCX = True
except ImportError:
    DocxDocument = None
    HAS_DOCX = False

try:
    from pypdf import PdfReader
    HAS_PDF = True
except ImportError:
    PdfReader = None
    HAS_PDF = False

try:
    import pythoncom
    import win32com.client
    HAS_WIN32_COM = True
except ImportError:
    pythoncom = None
    win32com = None
    HAS_WIN32_COM = False
