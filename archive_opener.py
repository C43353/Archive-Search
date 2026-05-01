"""Helpers for opening matched files as safely as the platform allows."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

from tkinter import messagebox

from archive_file_safety import can_open_binary_readonly
from archive_optional import HAS_WIN32_COM, pythoncom, win32com


class FileOpener:
    """Open result files as safely as the platform allows."""

    @staticmethod
    def open_result(file_type: str, file_path: str, sheet_name=None, row_number=None) -> None:
        """
        Open a selected result using the safest specialised opener available.

        If a read-only open cannot be guaranteed, the file is revealed in its
        containing folder instead of being opened directly.
        """
        target_path = Path(file_path)
        if not target_path.exists():
            messagebox.showerror("File not found", f"The file does not exist:\n{file_path}")
            return

        if file_type == "pdf":
            if FileOpener.open_pdf(file_path):
                return
            FileOpener.reveal_in_file_manager(file_path)
            return

        if not can_open_binary_readonly(target_path):
            FileOpener.reveal_in_file_manager(file_path)
            return

        if file_type == "excel" and HAS_WIN32_COM and sheet_name and row_number and sys.platform.startswith("win"):
            if FileOpener.open_excel_readonly(file_path=file_path, sheet_name=sheet_name, row_number=row_number):
                return
            FileOpener.reveal_in_file_manager(file_path)
            return

        if file_type == "word" and HAS_WIN32_COM and sys.platform.startswith("win"):
            if FileOpener.open_word_readonly(file_path=file_path):
                return
            FileOpener.reveal_in_file_manager(file_path)
            return

        FileOpener.reveal_in_file_manager(file_path)

    @staticmethod
    def reveal_in_file_manager(file_path: str) -> None:
        """Reveal the file in the operating system's file manager."""
        target_path = Path(file_path)
        if not target_path.exists():
            messagebox.showerror("File not found", f"The file does not exist:\n{file_path}")
            return

        resolved_path = target_path.resolve()
        try:
            if sys.platform.startswith("win"):
                subprocess.Popen(["explorer.exe", "/select,", str(resolved_path)])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-R", str(resolved_path)])
            else:
                subprocess.Popen(["xdg-open", str(resolved_path.parent)])
        except Exception as exc:
            messagebox.showerror("Could not reveal file", str(exc))

    @staticmethod
    def open_pdf(file_path: str) -> bool:
        """Open a PDF directly with the operating system default viewer."""
        target_path = Path(file_path)
        if not target_path.exists():
            messagebox.showerror("File not found", f"The PDF does not exist:\n{file_path}")
            return False

        resolved_path = target_path.resolve()
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(resolved_path))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(resolved_path)])
            else:
                subprocess.Popen(["xdg-open", str(resolved_path)])
            return True
        except Exception:
            return False

    @staticmethod
    def open_word_readonly(file_path: str) -> bool:
        """Open a Word document through COM in read-only mode when available."""
        target_path = Path(file_path)
        if not target_path.exists():
            messagebox.showerror("File not found", f"The document does not exist:\n{file_path}")
            return False

        try:
            pythoncom.CoInitialize()
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = True
            word.DisplayAlerts = 0
            try:
                word.AutomationSecurity = 3
            except Exception:
                pass
            document = word.Documents.Open(
                str(target_path.resolve()),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=True,
                Revert=False,
                OpenAndRepair=False,
                NoEncodingDialog=True,
            )
            document.Activate()
            return True
        except Exception:
            return False
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    @staticmethod
    def open_excel_readonly(file_path: str, sheet_name: str, row_number: int) -> bool:
        """Open an Excel workbook read-only and jump to the first matching row."""
        target_path = Path(file_path)
        if not target_path.exists():
            messagebox.showerror("File not found", f"The workbook does not exist:\n{file_path}")
            return False

        try:
            pythoncom.CoInitialize()
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = True
            try:
                excel.AutomationSecurity = 3
            except Exception:
                pass
            workbook = excel.Workbooks.Open(
                str(target_path.resolve()),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
                Notify=False,
            )
            worksheet = workbook.Worksheets(sheet_name)
            worksheet.Activate()
            target_cell = worksheet.Cells(row_number, 1)
            try:
                excel.Goto(target_cell, True)
            except Exception:
                target_cell.Select()
            try:
                excel.ActiveWindow.ScrollRow = max(1, row_number - 5)
                excel.ActiveWindow.ScrollColumn = 1
            except Exception:
                pass
            workbook.Activate()
            return True
        except Exception:
            return False
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
