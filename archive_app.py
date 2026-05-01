"""Tkinter user interface for Archive Search V11."""

from __future__ import annotations

import queue
import shlex
import threading
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional, Sequence, Tuple

from archive_config import (
    APP_TITLE,
    ARCHIVE_SEARCH_PRIMARY_FOLDER,
    ARCHIVE_SEARCH_SECONDARY_FOLDER,
    HIGHLIGHT_BACKGROUND,
    HIGHLIGHT_FOREGROUND,
    HIGHLIGHT_TAG_NAME,
    TEXT_DOCUMENT_EXTENSIONS,
    WORKBOOK_EXTENSIONS,
)

from archive_extractor import DocumentExtractor
from archive_index import SQLiteIndexManager
from archive_models import MatchSnippet, SearchResult, SearchRoot
from archive_opener import FileOpener
from archive_runner import SearchRunner
from archive_utils import (
    build_highlight_terms,
    get_app_folder,
    infer_type_label,
    pluralize,
)


class ArchiveSearchApp:
    """Main Tkinter application."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self._set_initial_window_geometry()

        self.app_folder = get_app_folder()
        # Worker threads never touch Tkinter directly; they post messages here
        # and the UI thread consumes them in _poll_queue().
        self.queue: queue.Queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.search_thread: Optional[threading.Thread] = None
        self.index_status_thread: Optional[threading.Thread] = None
        self.index_status_cancel_event = threading.Event()
        self.current_job_mode = "search"

        self.extractor = DocumentExtractor()
        self.index_manager = SQLiteIndexManager(self.extractor)

        self.current_highlight_terms: List[str] = []
        self.results_by_iid: Dict[str, SearchResult] = {}

        self.primary_folder_var = tk.StringVar(value=str(ARCHIVE_SEARCH_PRIMARY_FOLDER))
        self.secondary_folder_var = tk.StringVar(value=str(ARCHIVE_SEARCH_SECONDARY_FOLDER))
        self.use_primary_var = tk.BooleanVar(value=True)
        self.use_secondary_var = tk.BooleanVar(value=True)
        self.primary_subfolders_var = tk.BooleanVar(value=True)
        self.secondary_subfolders_var = tk.BooleanVar(value=True)

        self.search_workbooks_var = tk.BooleanVar(value=True)
        self.search_text_documents_var = tk.BooleanVar(value=True)

        self.search_var = tk.StringVar()
        self.match_mode_var = tk.StringVar(value="any")
        self.status_var = tk.StringVar(value="Ready.")
        self.index_info_var = tk.StringVar(value=self.index_manager.get_index_path())
        self.index_status_var = tk.StringVar(value="Index status: not checked yet. Click Check Index Status when needed.")

        # Folder and file-type changes affect the displayed index paths.
        for variable in (self.primary_folder_var, self.secondary_folder_var):
            variable.trace_add("write", lambda *_args: self._on_index_settings_changed())
        for variable in (
            self.use_primary_var,
            self.use_secondary_var,
            self.primary_subfolders_var,
            self.secondary_subfolders_var,
            self.search_workbooks_var,
            self.search_text_documents_var,
        ):
            variable.trace_add("write", lambda *_args: self._on_index_settings_changed())

        self._build_ui()
        self.root.bind("<Configure>", self._on_window_configure)
        self._update_status_wraplengths()
        self._update_folder_states()
        self._refresh_index_info()
        self._poll_queue()

    def _set_initial_window_geometry(self) -> None:
        """Open the app centred on the current screen and avoid oversized windows."""
        target_width = 1480
        target_height = 900
        try:
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            width = min(target_width, max(900, screen_width - 80))
            height = min(target_height, max(650, screen_height - 120))
            x = max(0, (screen_width - width) // 2)
            y = max(0, (screen_height - height) // 2)
            self.root.geometry(f"{width}x{height}+{x}+{y}")
        except Exception:
            self.root.geometry(f"{target_width}x{target_height}")

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=12)
        container.pack(fill="both", expand=True)

        top = ttk.Frame(container)
        top.pack(fill="x")

        style = ttk.Style(self.root)
        style.configure("FolderOn.TCheckbutton", font=("Segoe UI", 10, "bold"))
        style.configure("FolderOff.TCheckbutton", font=("Segoe UI", 10, "bold"), foreground="#808080")

        self.primary_check = ttk.Checkbutton(
            top,
            text="Primary Search Folder:",
            variable=self.use_primary_var,
            command=self._update_folder_states,
            style="FolderOn.TCheckbutton",
        )
        self.primary_check.grid(row=0, column=0, sticky="w", padx=(0, 4))

        self.primary_folder_entry = ttk.Entry(top, textvariable=self.primary_folder_var, width=90)
        self.primary_folder_entry.grid(row=0, column=1, sticky="we", padx=(4, 8))

        self.primary_browse_button = ttk.Button(
            top,
            text="Browse...",
            command=lambda: self._browse_folder(self.primary_folder_var, "Select primary search folder"),
        )
        self.primary_browse_button.grid(row=0, column=2, sticky="e")

        self.primary_subfolders_check = ttk.Checkbutton(top, text="Search subfolders", variable=self.primary_subfolders_var)
        self.primary_subfolders_check.grid(row=0, column=3, sticky="w", padx=(8, 0))

        self.secondary_check = ttk.Checkbutton(
            top,
            text="Secondary Search Folder:",
            variable=self.use_secondary_var,
            command=self._update_folder_states,
            style="FolderOn.TCheckbutton",
        )
        self.secondary_check.grid(row=1, column=0, sticky="w", padx=(0, 4), pady=(8, 0))

        self.secondary_folder_entry = ttk.Entry(top, textvariable=self.secondary_folder_var, width=90)
        self.secondary_folder_entry.grid(row=1, column=1, sticky="we", padx=(4, 8), pady=(8, 0))

        self.secondary_browse_button = ttk.Button(
            top,
            text="Browse...",
            command=lambda: self._browse_folder(self.secondary_folder_var, "Select secondary search folder"),
        )
        self.secondary_browse_button.grid(row=1, column=2, sticky="e", pady=(8, 0))

        self.secondary_subfolders_check = ttk.Checkbutton(top, text="Search subfolders", variable=self.secondary_subfolders_var)
        self.secondary_subfolders_check.grid(row=1, column=3, sticky="w", padx=(8, 0), pady=(8, 0))

        ttk.Label(top, text="Search Terms:", font=("Segoe UI", 10, "bold")).grid(row=2, column=0, sticky="w", pady=(12, 0))
        self.search_entry = ttk.Entry(top, textvariable=self.search_var, width=70)
        self.search_entry.grid(row=2, column=1, columnspan=3, sticky="we", padx=(4, 0), pady=(12, 0))
        self.search_entry.bind("<Return>", lambda event: self.start_search())

        options = ttk.Frame(container)
        options.pack(fill="x", pady=(10, 0))

        ttk.Label(options, text="Match mode:").pack(side="left")
        ttk.Radiobutton(options, text="Any word", variable=self.match_mode_var, value="any").pack(side="left", padx=(8, 0))
        ttk.Radiobutton(options, text="All words", variable=self.match_mode_var, value="all").pack(side="left", padx=(8, 0))

        ttk.Label(options, text="File types:").pack(side="left", padx=(20, 0))
        self.workbooks_check = ttk.Checkbutton(options, text="Workbooks", variable=self.search_workbooks_var)
        self.workbooks_check.pack(side="left", padx=(8, 0))
        self.text_documents_check = ttk.Checkbutton(options, text="Text documents", variable=self.search_text_documents_var)
        self.text_documents_check.pack(side="left", padx=(8, 0))

        self.search_button = ttk.Button(options, text="Search", command=self.start_search)
        self.search_button.pack(side="left", padx=(20, 0))
        self.update_button = ttk.Button(options, text="Update Index", command=self.update_index)
        self.update_button.pack(side="left", padx=(8, 0))
        self.check_index_button = ttk.Button(options, text="Check Index Status", command=self.check_index_status)
        self.check_index_button.pack(side="left", padx=(8, 0))
        self.cancel_button = ttk.Button(options, text="Cancel Process", command=self.cancel_process, state="disabled")
        self.cancel_button.pack(side="left", padx=(8, 0))
        self.clear_button = ttk.Button(options, text="Clear Output", command=self.clear_output)
        self.clear_button.pack(side="left", padx=(8, 0))

        help_text = (
            f"{APP_TITLE} keeps a saved search index so searches are faster. "
            "The index names are shown below; each selected archive folder stores its own index.\n"
            "Search uses the last completed index and does not change it. Update Index adds recent changes and will build the folder index first if one does not already exist.\n\n"
            "Use quotes for exact phrases, e.g. \"Sigma 5E\". Double-click a result or click 'Open Selected' to open the file read-only where supported."
        )
        tk.Label(container, text=help_text, justify="left", anchor="nw", height=4).pack(anchor="w", fill="x", pady=(10, 4))

        info_panel = ttk.Frame(container)
        info_panel.pack(fill="x", pady=(0, 4))

        self.index_info_label = tk.Label(
            info_panel,
            textvariable=self.index_info_var,
            anchor="nw",
            justify="left",
            fg="#444444",
            height=4,
        )
        self.index_info_label.pack(fill="x", anchor="w")

        self.index_status_label = tk.Label(
            info_panel,
            textvariable=self.index_status_var,
            anchor="nw",
            justify="left",
            fg="#8a5a00",
        )
        self.index_status_label.pack(fill="x", anchor="w", pady=(0, 0))

        self.status_label = tk.Label(
            info_panel,
            textvariable=self.status_var,
            anchor="nw",
            justify="left",
        )
        self.status_label.pack(fill="x", anchor="w", pady=(2, 0))

        self.panes = ttk.Panedwindow(container, orient="horizontal")
        panes = self.panes
        panes.pack(fill="both", expand=True)
        self.root.after_idle(self._set_initial_pane_sash)

        left_frame = ttk.Frame(panes)
        right_frame = ttk.Frame(panes)
        panes.add(left_frame, weight=5)
        panes.add(right_frame, weight=6)

        list_controls = ttk.Frame(left_frame)
        list_controls.pack(fill="x", pady=(0, 6))
        ttk.Button(list_controls, text="Open Selected", command=self.open_selected_result).pack(side="left")
        ttk.Button(list_controls, text="Copy Path", command=self.copy_selected_path).pack(side="left", padx=(8, 0))

        columns = ("document", "type", "matches", "location")
        self.results_tree = ttk.Treeview(left_frame, columns=columns, show="headings", selectmode="browse")
        self.results_tree.heading("document", text="Document")
        self.results_tree.heading("type", text="Type")
        self.results_tree.heading("matches", text="Matches")
        self.results_tree.heading("location", text="First location")

        self.results_tree.column("document", width=320, anchor="w")
        self.results_tree.column("type", width=80, anchor="center")
        self.results_tree.column("matches", width=90, anchor="center")
        self.results_tree.column("location", width=180, anchor="w")

        tree_scroll_y = ttk.Scrollbar(left_frame, orient="vertical", command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=tree_scroll_y.set)

        self.results_tree.pack(fill="both", expand=True, side="left")
        tree_scroll_y.pack(fill="y", side="right")

        self.results_tree.bind("<<TreeviewSelect>>", self._on_result_selected)
        self.results_tree.bind("<Double-1>", lambda event: self.open_selected_result())

        detail_header = ttk.Frame(right_frame)
        detail_header.pack(fill="x", pady=(0, 6))

        self.detail_title_var = tk.StringVar(value="Result details")
        self.detail_meta_var = tk.StringVar(value="Select a result to inspect content.")
        ttk.Label(detail_header, textvariable=self.detail_title_var, font=("Segoe UI", 12, "bold")).pack(anchor="w")
        ttk.Label(detail_header, textvariable=self.detail_meta_var).pack(anchor="w", pady=(2, 0))

        detail_text_frame = ttk.Frame(right_frame)
        detail_text_frame.pack(fill="both", expand=True)

        self.details_text = tk.Text(detail_text_frame, wrap="word", font=("Consolas", 10))
        detail_scroll = ttk.Scrollbar(detail_text_frame, orient="vertical", command=self.details_text.yview)
        self.details_text.configure(yscrollcommand=detail_scroll.set)

        self.details_text.pack(side="left", fill="both", expand=True)
        detail_scroll.pack(side="right", fill="y")

        self.details_text.tag_configure(HIGHLIGHT_TAG_NAME, background=HIGHLIGHT_BACKGROUND, foreground=HIGHLIGHT_FOREGROUND)
        self.details_text.tag_configure("heading", font=("Segoe UI", 10, "bold"))
        self.details_text.tag_configure("muted", foreground="#666666")
        self.details_text.tag_configure("snippet_label", font=("Segoe UI", 9, "bold"))
        self.details_text.tag_raise("sel")
        self.details_text.tag_lower(HIGHLIGHT_TAG_NAME, "sel")
        self.details_text.config(state="disabled")

        top.columnconfigure(1, weight=1)
        
        github_label = tk.Label(
        container,
        text="GitHub Project",
        fg="#666666",
        cursor="hand2",
        anchor="w",
        font=("Segoe UI", 9, "underline"),
        )
        github_label.pack(anchor="w", pady=(0, 6))
        github_label.bind("<Button-1>", lambda _event: webbrowser.open("https://github.com/C43353/Archive-Search"))


    def _set_initial_pane_sash(self) -> None:
        """Give the details pane slightly more starting width than the results pane."""
        panes = getattr(self, "panes", None)
        if panes is None:
            return
        try:
            width = panes.winfo_width()
            if width > 200:
                panes.sashpos(0, int(width * 0.46))
        except Exception:
            pass

    def _on_index_settings_changed(self) -> None:
        self._refresh_index_info()
        self.index_status_var.set("Index status: not checked yet. Click Check Index Status when needed.")

    def _collect_selected_search_roots(self) -> Tuple[SearchRoot, ...]:
        """Return only enabled root folders that currently exist on disk."""
        roots: List[SearchRoot] = []
        if self.use_primary_var.get():
            primary_text = self.primary_folder_var.get().strip()
            if primary_text:
                primary_path = Path(primary_text).expanduser()
                if primary_path.exists() and primary_path.is_dir():
                    roots.append(SearchRoot("Primary", primary_path, self.primary_subfolders_var.get()))
        if self.use_secondary_var.get():
            secondary_text = self.secondary_folder_var.get().strip()
            if secondary_text:
                secondary_path = Path(secondary_text).expanduser()
                if secondary_path.exists() and secondary_path.is_dir():
                    roots.append(SearchRoot("Secondary", secondary_path, self.secondary_subfolders_var.get()))
        return tuple(roots)

    def _collect_allowed_extensions(self) -> Tuple[str, ...]:
        """Translate the UI checkboxes into the suffixes the index should scan."""
        allowed_extensions = set()
        if self.search_workbooks_var.get():
            allowed_extensions |= WORKBOOK_EXTENSIONS
        if self.search_text_documents_var.get():
            allowed_extensions |= TEXT_DOCUMENT_EXTENSIONS
        return tuple(sorted(allowed_extensions))

    def check_index_status(self) -> None:
        if self.search_thread and self.search_thread.is_alive():
            self.index_status_var.set("Index status: wait until the current operation finishes, then click Check Index Status.")
            return
        if self.index_status_thread and self.index_status_thread.is_alive():
            self.index_status_var.set("Index status: already checking. Please wait.")
            return

        roots = self._collect_selected_search_roots()
        allowed_extensions = self._collect_allowed_extensions()
        if not roots or not allowed_extensions:
            self.index_status_var.set("Index status: select at least one valid folder and file type.")
            return

        self.index_status_cancel_event.set()
        self.index_status_cancel_event = threading.Event()
        cancel_event = self.index_status_cancel_event
        self.index_status_var.set("Index status: checking for new or changed files...")

        def worker() -> None:
            try:
                summary = self.index_manager.get_stale_index_summary(
                    roots=roots,
                    allowed_extensions=allowed_extensions,
                    cancel_event=cancel_event,
                )
                if not cancel_event.is_set():
                    self.queue.put(("index_status", summary))
            except Exception as exc:
                if not cancel_event.is_set():
                    self.queue.put(("index_status_error", str(exc)))

        self.index_status_thread = threading.Thread(target=worker, daemon=True)
        self.index_status_thread.start()

    def _format_index_status_summary(self, summary: Dict[str, object]) -> str:
        files_checked = int(summary.get("files_checked", 0))
        indexed_files = int(summary.get("indexed_files", 0))
        new_files = int(summary.get("new_files", 0))
        changed_files = int(summary.get("changed_files", 0))
        deleted_files = int(summary.get("deleted_files", 0))
        missing_databases = int(summary.get("missing_databases", 0))
        invalid_databases = int(summary.get("invalid_databases", 0))
        errors = int(summary.get("errors", 0))
        stale_total = new_files + changed_files + deleted_files

        if missing_databases or invalid_databases:
            parts = []
            if missing_databases:
                parts.append(f"{missing_databases} missing {pluralize(missing_databases, 'index', 'indexes')}")
            if invalid_databases:
                parts.append(f"{invalid_databases} invalid {pluralize(invalid_databases, 'index', 'indexes')}")
            return (
                "Index status: update recommended — "
                + ", ".join(parts)
                + f"; {files_checked} searchable {pluralize(files_checked, 'file')} found. "
                "Click Update Index to build or replace the affected index files."
            )
        if stale_total:
            details = []
            if new_files:
                details.append(f"{new_files} new")
            if changed_files:
                details.append(f"{changed_files} changed")
            if deleted_files:
                details.append(f"{deleted_files} deleted/missing")
            suffix = f"; {errors} scan {pluralize(errors, 'error')}" if errors else ""
            return (
                "Index status: update recommended — "
                + ", ".join(details)
                + f" since the last index update. Click Update Index to update the saved index{suffix}."
            )
        suffix = f"; {errors} scan {pluralize(errors, 'error')}" if errors else ""
        return (
            f"Index status: up to date for {indexed_files} indexed {pluralize(indexed_files, 'file')} "
            f"({files_checked} current {pluralize(files_checked, 'file')} checked{suffix})."
        )

    def _on_window_configure(self, event=None) -> None:
        self._update_status_wraplengths()

    def _update_status_wraplengths(self) -> None:
        wraplength = max(900, self.root.winfo_width() - 80)
        for widget in (
            getattr(self, "index_info_label", None),
            getattr(self, "index_status_label", None),
            getattr(self, "status_label", None),
        ):
            if widget is not None:
                widget.configure(wraplength=wraplength)

    def _browse_folder(self, variable: tk.StringVar, title: str) -> None:
        current = variable.get().strip()
        initial_dir = current if current and Path(current).exists() else str(self.app_folder)
        selected = filedialog.askdirectory(title=title, initialdir=initial_dir)
        if selected:
            variable.set(selected)

    def _update_folder_states(self) -> None:
        primary_enabled = self.use_primary_var.get()
        secondary_enabled = self.use_secondary_var.get()
        primary_state = "normal" if primary_enabled else "disabled"
        secondary_state = "normal" if secondary_enabled else "disabled"

        self.primary_folder_entry.config(state=primary_state)
        self.primary_browse_button.config(state=primary_state)
        self.primary_subfolders_check.config(state=primary_state)
        self.primary_check.config(style="FolderOn.TCheckbutton" if primary_enabled else "FolderOff.TCheckbutton")

        self.secondary_folder_entry.config(state=secondary_state)
        self.secondary_browse_button.config(state=secondary_state)
        self.secondary_subfolders_check.config(state=secondary_state)
        self.secondary_check.config(style="FolderOn.TCheckbutton" if secondary_enabled else "FolderOff.TCheckbutton")
        self._refresh_index_info()

    def _set_search_running_ui(self, running: bool) -> None:
        if running:
            self.search_button.config(state="disabled")
            self.update_button.config(state="disabled")
            self.check_index_button.config(state="disabled")
            self.cancel_button.config(state="normal")
            self.primary_check.config(state="disabled")
            self.secondary_check.config(state="disabled")
            self.primary_folder_entry.config(state="disabled")
            self.secondary_folder_entry.config(state="disabled")
            self.primary_browse_button.config(state="disabled")
            self.secondary_browse_button.config(state="disabled")
            self.primary_subfolders_check.config(state="disabled")
            self.secondary_subfolders_check.config(state="disabled")
            self.workbooks_check.config(state="disabled")
            self.text_documents_check.config(state="disabled")
            self.search_entry.config(state="disabled")
        else:
            self.search_button.config(state="normal")
            self.update_button.config(state="normal")
            self.check_index_button.config(state="normal")
            self.cancel_button.config(state="disabled")
            self.primary_check.config(state="normal")
            self.secondary_check.config(state="normal")
            self.workbooks_check.config(state="normal")
            self.text_documents_check.config(state="normal")
            self.search_entry.config(state="normal")
            self._update_folder_states()

    def _set_details_editable(self, editable: bool) -> None:
        self.details_text.config(state="normal" if editable else "disabled")

    def _clear_details(self) -> None:
        self.detail_title_var.set("Result details")
        self.detail_meta_var.set("Select a result to inspect content.")
        self._set_details_editable(True)
        try:
            self.details_text.delete("1.0", "end")
        finally:
            self._set_details_editable(False)

    def _highlight_details(self, start_index: str, end_index: str) -> None:
        if not self.current_highlight_terms:
            return
        for term in self.current_highlight_terms:
            if not term:
                continue
            search_start = start_index
            while True:
                match_index = self.details_text.search(term, search_start, stopindex=end_index, nocase=True)
                if not match_index:
                    break
                match_end = f"{match_index}+{len(term)}c"
                self.details_text.tag_add(HIGHLIGHT_TAG_NAME, match_index, match_end)
                search_start = match_end

    def _write_detail_text(self, text: str, tags: Tuple[str, ...] = ()) -> None:
        if not text:
            return
        start_index = self.details_text.index("end-1c")
        self.details_text.insert("end", text, tags)
        end_index = self.details_text.index("end-1c")
        if self.current_highlight_terms:
            self._highlight_details(start_index, end_index)

    def _write_result_folder_path(self, result: SearchResult) -> None:
        folder_path = str(Path(result.path).parent)
        self._write_detail_text(f"Folder: {folder_path}\n\n", ("muted",))

    def _show_result_details(self, result: SearchResult) -> None:
        type_label = infer_type_label(result.file_type)
        self.detail_title_var.set(result.document_name)
        meta = f"{type_label} | Root: {result.root_label}"
        if result.first_location_label:
            meta += f" | First match: {result.first_location_label}"
        self.detail_meta_var.set(meta)
        self._set_details_editable(True)
        try:
            self.details_text.delete("1.0", "end")
            self._write_result_folder_path(result)
            if result.file_type == "excel":
                for index, snippet in enumerate(result.snippets, start=1):
                    self._write_detail_text(f"{index}. ", ("muted",))
                    self._write_detail_text(f"{snippet.location_label}: ", ("snippet_label",))
                    self._write_detail_text(f"{snippet.preview}\n\n")
            else:
                full_text = self.index_manager.get_indexed_document_text(result)
                if full_text:
                    self._write_detail_text(full_text.rstrip() + "\n")
                else:
                    for snippet in result.snippets:
                        self._write_detail_text(f"{snippet.preview}\n\n")
            self.details_text.see("1.0")
        finally:
            self._set_details_editable(False)

    def _clear_results_tree(self) -> None:
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        self.results_by_iid.clear()

    def _append_result_to_tree(self, result: SearchResult) -> None:
        first_snippet = result.snippets[0] if result.snippets else MatchSnippet(location_label="", preview="")
        iid = self.results_tree.insert(
            "",
            "end",
            values=(
                result.document_name,
                infer_type_label(result.file_type),
                result.match_count,
                first_snippet.location_label,
            ),
        )
        self.results_by_iid[iid] = result

    def _on_result_selected(self, event=None) -> None:
        selected = self.results_tree.selection()
        if not selected:
            return
        result = self.results_by_iid.get(selected[0])
        if result is not None:
            self._show_result_details(result)

    def open_selected_result(self) -> None:
        selected = self.results_tree.selection()
        if not selected:
            messagebox.showinfo("No selection", "Please select a result first.")
            return
        result = self.results_by_iid.get(selected[0])
        if result is None:
            return
        FileOpener.open_result(
            file_type=result.file_type,
            file_path=result.path,
            sheet_name=result.open_sheet,
            row_number=result.open_row_number,
        )

    def copy_selected_path(self) -> None:
        selected = self.results_tree.selection()
        if not selected:
            messagebox.showinfo("No selection", "Please select a result first.")
            return
        result = self.results_by_iid.get(selected[0])
        if result is None:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(result.path)
        self.status_var.set("Copied selected result path to clipboard.")

    def clear_output(self) -> None:
        self._clear_results_tree()
        self._clear_details()
        self.status_var.set("Ready.")

    def _refresh_index_info(self) -> None:
        roots = []
        if self.use_primary_var.get():
            primary_text = self.primary_folder_var.get().strip()
            if primary_text:
                roots.append(SearchRoot("Primary", Path(primary_text).expanduser(), self.primary_subfolders_var.get()))
        if self.use_secondary_var.get():
            secondary_text = self.secondary_folder_var.get().strip()
            if secondary_text:
                roots.append(SearchRoot("Secondary", Path(secondary_text).expanduser(), self.secondary_subfolders_var.get()))
        allowed_extensions = self._collect_allowed_extensions()
        self.index_info_var.set(self.index_manager.describe_selected_index_paths(tuple(roots), allowed_extensions))

    def _get_selected_search_roots(self) -> Optional[Tuple[SearchRoot, ...]]:
        roots: List[SearchRoot] = []
        if not self.use_primary_var.get() and not self.use_secondary_var.get():
            messagebox.showwarning("No Search Folder Selected", "Please enable the primary folder, the secondary folder, or both.")
            return None
        if self.use_primary_var.get():
            primary_text = self.primary_folder_var.get().strip()
            if not primary_text:
                messagebox.showwarning("Missing Folder", "The primary folder is enabled, but no primary folder was entered.")
                return None
            primary_path = Path(primary_text).expanduser()
            if not primary_path.exists() or not primary_path.is_dir():
                messagebox.showerror("Invalid Primary Folder", f"The primary search folder does not exist or is not a folder:\n{primary_path}")
                return None
            roots.append(SearchRoot("Primary", primary_path, self.primary_subfolders_var.get()))
        if self.use_secondary_var.get():
            secondary_text = self.secondary_folder_var.get().strip()
            if not secondary_text:
                messagebox.showwarning("Missing Folder", "The secondary folder is enabled, but no secondary folder was entered.")
                return None
            secondary_path = Path(secondary_text).expanduser()
            if not secondary_path.exists() or not secondary_path.is_dir():
                messagebox.showerror("Invalid Secondary Folder", f"The secondary search folder does not exist or is not a folder:\n{secondary_path}")
                return None
            roots.append(SearchRoot("Secondary", secondary_path, self.secondary_subfolders_var.get()))
        return tuple(roots)

    def _get_allowed_extensions(self) -> Optional[Tuple[str, ...]]:
        allowed_extensions = set()
        if self.search_workbooks_var.get():
            allowed_extensions |= WORKBOOK_EXTENSIONS
        if self.search_text_documents_var.get():
            allowed_extensions |= TEXT_DOCUMENT_EXTENSIONS
        if not allowed_extensions:
            messagebox.showwarning("No File Types Selected", "Please tick Workbooks, Text documents, or both.")
            return None
        return tuple(sorted(allowed_extensions))

    @staticmethod
    def _format_blocking_file_type_issues(issues: Sequence[Dict[str, str]]) -> str:
        shown = []
        for index, issue in enumerate(issues[:10], start=1):
            shown.append(
                f"Folder:\n{issue.get('folder', '')}\n\n"
                f"File:\n{issue.get('file_name', '')}\n\n"
                f"Extension suggests: {issue.get('expected', '')}\n"
                f"Actual file content: {issue.get('actual', '')}"
            )
        extra = ""
        if len(issues) > 10:
            extra = f"\n\nOnly the first 10 issues are shown. There are {len(issues)} issues in total."
        return (
            "Update Index has not started because one or more files have an extension that does not match the file content.\n\n"
            + "\n\n".join(shown)
            + extra
            + "\n\nRename, remove, or replace the problem file(s), then click Update Index again."
        )

    def _stop_index_status_check(self) -> None:
        """Stop any in-progress manual index-status check before changing an index file."""
        self.index_status_cancel_event.set()
        if self.index_status_thread and self.index_status_thread.is_alive():
            self.index_status_thread.join(timeout=2.0)


    def _start_background_job(
        self,
        *,
        search_terms: Sequence[str],
        update_only: bool = False,
    ) -> None:
        """Start a search or incremental update in a daemon worker thread."""
        if self.search_thread and self.search_thread.is_alive():
            return

        search_roots = self._get_selected_search_roots()
        if not search_roots:
            return
        allowed_extensions = self._get_allowed_extensions()
        if not allowed_extensions:
            return

        if update_only:
            self._stop_index_status_check()
            self.status_var.set("Checking file types before updating the index...")
            self.root.update_idletasks()
            blocking_issues = self.index_manager.find_blocking_file_type_issues(search_roots, allowed_extensions)
            if blocking_issues:
                self.status_var.set("Update cancelled: fix the file type issue(s), then try again.")
                messagebox.showerror(
                    "File Type Issue Found",
                    self._format_blocking_file_type_issues(blocking_issues),
                )
                return

        if update_only:
            build_requirements = self.index_manager.get_index_build_requirements(search_roots)
            if build_requirements:
                missing_lines = []
                invalid_lines = []
                for item in build_requirements:
                    root = item["root"]
                    database_path = item["database_path"]
                    line = f"- {root.label}: {root.path}\n\n{database_path.name}"
                    if item.get("reason") == "invalid":
                        invalid_lines.append(line)
                    else:
                        missing_lines.append(line)

                message_parts = []
                if missing_lines:
                    message_parts.append(
                        "The following selected folder(s) do not already have a saved Archive Search index:\n\n"
                        + "\n".join(missing_lines)
                    )
                if invalid_lines:
                    message_parts.append(
                        "The following selected folder(s) already have a file at the expected index path, "
                        "but it is not a valid Archive Search index. It will be overwritten with a fresh build:\n\n"
                        + "\n".join(invalid_lines)
                    )

                confirmed = messagebox.askyesno(
                    "Index Will Be Built First",
                    "\n\n".join(message_parts)
                    + "\n\nUpdate Index will build the listed index file(s) from scratch, so it may take a while. Continue?",
                )
                if not confirmed:
                    self.status_var.set("Update cancelled.")
                    return

        self.cancel_event.clear()
        if not update_only:
            self.index_status_cancel_event.set()
        self.current_job_mode = "update" if update_only else "search"
        self.current_highlight_terms = build_highlight_terms(search_terms)
        self.clear_output()
        self.status_var.set("Working...")
        self._set_search_running_ui(True)

        runner = SearchRunner(
            index_manager=self.index_manager,
            queue_out=self.queue,
            cancel_event=self.cancel_event,
        )
        self.search_thread = threading.Thread(
            target=runner.run,
            args=(search_roots, allowed_extensions, tuple(search_terms), self.match_mode_var.get()),
            kwargs={
                "update_only": update_only,
            },
            daemon=True,
        )
        self.search_thread.start()

    def start_search(self) -> None:
        if self.search_thread and self.search_thread.is_alive():
            return
        search_input = self.search_var.get().strip()
        if not search_input:
            messagebox.showwarning("Missing Search Terms", "Please type one or more search terms.")
            return
        try:
            search_terms = tuple(term.lower() for term in shlex.split(search_input))
        except ValueError as exc:
            messagebox.showerror("Invalid Search Input", f"Could not parse search terms:\n{exc}")
            return
        self._start_background_job(search_terms=search_terms)

    def update_index(self) -> None:
        self._start_background_job(search_terms=(), update_only=True)

    def cancel_process(self) -> None:
        if self.search_thread and self.search_thread.is_alive():
            self.cancel_event.set()
            self.status_var.set("Cancelling...")
            self.cancel_button.config(state="disabled")

    def _poll_queue(self) -> None:
        pending_item_batches: List[List[Dict[str, object]]] = []
        latest_status: Optional[str] = None
        finished_payload: Optional[Dict[str, object]] = None
        fatal_payload: Optional[str] = None
        index_status_payload: Optional[Dict[str, object]] = None
        index_status_error: Optional[str] = None

        try:
            while True:
                item_type, payload = self.queue.get_nowait()
                if item_type == "items":
                    pending_item_batches.append(payload)
                elif item_type == "status":
                    latest_status = payload
                elif item_type == "finished":
                    finished_payload = payload
                elif item_type == "fatal":
                    fatal_payload = payload
                elif item_type == "index_status":
                    index_status_payload = payload
                elif item_type == "index_status_error":
                    index_status_error = payload
        except queue.Empty:
            pass

        if index_status_payload is not None:
            self._refresh_index_info()
            self.index_status_var.set(self._format_index_status_summary(index_status_payload))

        if index_status_error is not None:
            self.index_status_var.set(f"Index status: could not check index ({index_status_error}).")

        if pending_item_batches:
            for batch in pending_item_batches:
                for item in batch:
                    kind = item["kind"]
                    payload = item["payload"]
                    if kind == "result":
                        self._append_result_to_tree(payload)
            if not self.results_tree.selection() and self.results_tree.get_children():
                first = self.results_tree.get_children()[0]
                self.results_tree.selection_set(first)
                self.results_tree.focus(first)
                self.results_tree.see(first)
                self._on_result_selected()

        if latest_status is not None:
            self.status_var.set(latest_status)

        if finished_payload is not None:
            files_scanned = int(finished_payload.get("files_scanned", 0))
            indexed_files = int(finished_payload.get("indexed_files", 0))
            matched_files = int(finished_payload.get("matched_files", 0))
            total_matches = int(finished_payload.get("total_matches", 0))
            cancelled = bool(finished_payload.get("cancelled", False))
            elapsed_seconds = float(finished_payload.get("elapsed_seconds", 0.0))
            mode = str(finished_payload.get("mode", self.current_job_mode))
            files_reused = int(finished_payload.get("files_reused", 0))
            files_refreshed = int(finished_payload.get("files_refreshed", 0))
            files_deleted = int(finished_payload.get("files_deleted", 0))
            files_failed = int(finished_payload.get("files_failed", 0))
            first_error = str(finished_payload.get("first_error") or "")

            if cancelled:
                final_message = (
                    f"Cancelled. Scanned {files_scanned} {pluralize(files_scanned, 'file')}. "
                    f"Indexed set currently contains {indexed_files} {pluralize(indexed_files, 'file')}. "
                    f"Time: {elapsed_seconds:.2f} seconds."
                )
            elif mode == "update":
                warning = ""
                if files_failed:
                    warning = (
                        f" Warning: {files_failed} {pluralize(files_failed, 'file')} could not be indexed."
                    )
                    if first_error:
                        warning += f" First issue: {first_error}."
                final_message = (
                    f"Update finished. Checked {files_scanned} {pluralize(files_scanned, 'file')}. "
                    f"Kept {files_reused} unchanged, updated {files_refreshed}, removed {files_deleted}. "
                    f"The saved index now contains {indexed_files} {pluralize(indexed_files, 'file')}."
                    f"{warning} "
                    f"Time: {elapsed_seconds:.2f} seconds."
                )
            elif mode == "initial_index_required":
                missing_roots = str(finished_payload.get("missing_roots") or "the selected folder")
                final_message = (
                    f"Search not started. The saved index is missing or is not a valid Archive Search index for {missing_roots}. "
                    "Click Update Index first. After it finishes, click Search again."
                )
            else:
                final_message = (
                    f"Finished searching the saved index. "
                    f"Indexed {indexed_files} {pluralize(indexed_files, 'file')} total. "
                    f"Found {matched_files} matching {pluralize(matched_files, 'file')} with {total_matches} total {pluralize(total_matches, 'match')}. "
                    f"Time: {elapsed_seconds:.2f} seconds."
                )

            self.status_var.set(final_message)
            self._set_search_running_ui(False)
            if not cancelled and mode == "update":
                self._refresh_index_info()
                self.index_status_var.set("Index status: up to date.")

        if fatal_payload is not None:
            self.status_var.set("Operation failed.")
            self._set_search_running_ui(False)
            messagebox.showerror("Operation failed", fatal_payload)

        self.root.after(100, self._poll_queue)
