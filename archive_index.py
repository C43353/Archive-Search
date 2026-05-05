"""SQLite/FTS index management for Archive Search V11."""

from __future__ import annotations

import gc
import hashlib
import os
import re
import shutil
import socket
import sqlite3
import threading
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Callable, Dict, Iterator, List, Optional, Sequence, Tuple
from urllib.parse import quote, urlencode

from archive_config import (
    DETAIL_PREVIEW_LENGTH,
    DETAIL_SNIPPETS_PER_FILE,
    INDEX_BATCH_SIZE,
    INDEX_DB_PREFIX,
    INDEX_RETRY_ATTEMPTS,
    INDEX_RETRY_DELAY_SECONDS,
    INDEX_SCHEMA_VERSION,
    INDEX_STABLE_PROBE_DELAY_SECONDS,
    PDF_EXTENSIONS,
    PLAIN_TEXT_EXTENSIONS,
    RESULT_LIMIT,
    STATUS_UPDATE_INTERVAL_SECONDS,
    WORD_EXTENSIONS,
    WORD_LEGACY_EXTENSIONS,
    WORKBOOK_EXTENSIONS,
)
from archive_extractor import DocumentExtractor, ReadOnlyWordSession
from archive_file_safety import detect_file_content_kind, detect_supported_file_type, describe_file_content
from archive_models import MatchSnippet, SearchResult, SearchRoot
from archive_optional import HAS_WIN32_COM
from archive_utils import format_utc_iso_for_display, pluralize, should_ignore_filename, truncate_text, utc_now_iso


class ArchiveIndexLockedError(RuntimeError):
    """Raised when another user/computer is already updating an archive index."""


class ArchiveIndexReplaceError(RuntimeError):
    """Raised when a completed index cannot be swapped into place safely."""


class SQLiteIndexManager:
    """Maintain one SQLite/FTS index file inside each selected search root."""

    def __init__(self, extractor: DocumentExtractor) -> None:
        self.extractor = extractor
        self._schema_lock = threading.Lock()
        self.lock_stale_after_seconds = 2 * 60 * 60

    @staticmethod
    def _relative_path_in_scope(relative_path: str, include_subfolders: bool) -> bool:
        normalized = str(relative_path or "").replace("\\", "/").strip("/")
        return include_subfolders or "/" not in normalized

    def _filter_existing_rows_for_scope(
        self,
        root: SearchRoot,
        rows_by_path: Dict[str, sqlite3.Row],
    ) -> Dict[str, sqlite3.Row]:
        if root.include_subfolders:
            return dict(rows_by_path)
        return {
            path: row
            for path, row in rows_by_path.items()
            if self._relative_path_in_scope(str(row["relative_path"] or ""), root.include_subfolders)
        }

    @staticmethod
    def _scope_sql_condition(root: SearchRoot) -> tuple[str, tuple[object, ...]]:
        """Return extra SQL used when the user limits a root to top-level files."""
        if root.include_subfolders:
            return "", ()
        return " AND instr(files.relative_path, '/') = 0", ()

    @staticmethod
    def _quoted_match_term(term: str) -> str:
        return '"' + term.replace('"', '""') + '"'

    @staticmethod
    def _is_simple_prefix_token(term: str) -> bool:
        return bool(re.fullmatch(r"(?u)\w+", term))

    def _build_fts_term(self, term: str) -> str:
        cleaned = term.strip()
        if not cleaned:
            return ""
        if self._is_simple_prefix_token(cleaned):
            return f"{cleaned}*"
        return self._quoted_match_term(cleaned)

    def build_match_query(self, search_terms: Sequence[str], match_mode: str) -> str:
        """Build the SQLite FTS MATCH query for any-word or all-words mode."""
        operator = " AND " if match_mode == "all" else " OR "
        return operator.join(self._build_fts_term(term) for term in search_terms if term.strip())

    @staticmethod
    def _build_like_conditions(search_terms: Sequence[str], match_mode: str) -> tuple[str, list[str]]:
        cleaned_terms = [term.strip().lower() for term in search_terms if term.strip()]
        if not cleaned_terms:
            return "", []
        joiner = " AND " if match_mode == "all" else " OR "
        clause = joiner.join("LOWER(chunk_index.content) LIKE ?" for _ in cleaned_terms)
        params = [f"%{term}%" for term in cleaned_terms]
        return clause, params

    @staticmethod
    def _root_fingerprint(root_path: Path) -> str:
        """Return a stable non-reversible ID for a root folder path."""
        try:
            normalized = str(root_path.expanduser().resolve(strict=False))
        except Exception:
            normalized = str(root_path.expanduser().absolute())
        normalized = os.path.normcase(normalized)
        return hashlib.sha256(normalized.encode("utf-8", errors="surrogatepass")).hexdigest()[:16]

    def get_root_database_path(self, root: SearchRoot | Path) -> Path:
        """Return the SQLite database path for an archive folder.

        The database name is derived only from the canonical folder path, not from
        the UI slot (Primary/Secondary). This means the same folder uses the same
        index even if it is moved between Primary and Secondary, or accidentally
        entered in both places.
        """
        root_path = root.path if isinstance(root, SearchRoot) else Path(root)
        fingerprint = self._root_fingerprint(root_path)
        database_filename = f"{INDEX_DB_PREFIX}{fingerprint}.sqlite3"
        return root_path.expanduser() / database_filename

    def _lock_path_for_database(self, database_path: Path) -> Path:
        return database_path.with_suffix(database_path.suffix + ".lock")

    def _temp_path_for_database(self, database_path: Path) -> Path:
        return database_path.with_suffix(database_path.suffix + ".tmp")

    def _backup_path_for_database(self, database_path: Path) -> Path:
        return database_path.with_suffix(database_path.suffix + ".bak")

    def _remove_database_file_set(self, database_path: Path) -> None:
        """Remove a SQLite database file and any temporary sidecar files for it."""
        for candidate in (
            database_path,
            database_path.with_name(database_path.name + "-journal"),
            database_path.with_name(database_path.name + "-wal"),
            database_path.with_name(database_path.name + "-shm"),
        ):
            try:
                if candidate.exists():
                    candidate.unlink()
            except OSError as exc:
                raise ArchiveIndexReplaceError(
                    "A previous unfinished index file could not be removed. Please close other copies of Archive Search and try again."
                ) from exc

    def _write_lock_details(self, lock_file, root: SearchRoot, database_path: Path) -> None:
        user_name = os.environ.get("USERNAME") or os.environ.get("USER") or "Unknown user"
        computer_name = socket.gethostname() or "Unknown computer"
        lock_file.write("Archive Search index update in progress\n")
        lock_file.write(f"Folder: {root.path}\n")
        lock_file.write(f"Index: {database_path.name}\n")
        lock_file.write(f"User: {user_name}\n")
        lock_file.write(f"Computer: {computer_name}\n")
        lock_file.write(f"Started: {utc_now_iso()}\n")
        lock_file.flush()

    def _read_lock_details(self, lock_path: Path) -> str:
        try:
            return lock_path.read_text(encoding="utf-8", errors="replace").strip()
        except OSError:
            return "Another update is already running."

    def _remove_abandoned_lock_if_safe(
        self,
        *,
        lock_path: Path,
        database_path: Path,
        status_callback: Callable[[str], None],
        root: SearchRoot,
    ) -> bool:
        """Clear old lock marker files that were probably left behind after a crash."""
        if not lock_path.exists():
            return False

        try:
            lock_age = time.time() - lock_path.stat().st_mtime
        except OSError:
            lock_age = 0

        temp_path = self._temp_path_for_database(database_path)
        looks_like_abandoned_initial_build = (
            not database_path.exists()
            and not temp_path.exists()
            and lock_age > 60
        )
        looks_stale = lock_age > self.lock_stale_after_seconds

        if not (looks_stale or looks_like_abandoned_initial_build):
            return False

        try:
            lock_path.unlink()
            if looks_stale:
                status_callback(
                    f"A previous update marker for the {root.label.lower()} folder looked old, so it was cleared and the update will continue."
                )
            else:
                status_callback(
                    f"A previous setup marker for the {root.label.lower()} folder was left behind, so it was cleared and the index build will continue."
                )
            return True
        except OSError:
            return False

    @contextmanager
    def _root_update_lock(self, root: SearchRoot, database_path: Path, status_callback: Callable[[str], None]):
        """Take a simple cross-process lock before replacing an archive index."""
        lock_path = self._lock_path_for_database(database_path)
        database_path.parent.mkdir(parents=True, exist_ok=True)

        self._remove_abandoned_lock_if_safe(
            lock_path=lock_path,
            database_path=database_path,
            status_callback=status_callback,
            root=root,
        )

        try:
            lock_file = open(lock_path, "x", encoding="utf-8")
        except FileExistsError as exc:
            details = self._read_lock_details(lock_path)
            raise ArchiveIndexLockedError(
                "Another person or computer may already be updating this archive. "
                "You can still search the last completed index, but only one person can update it at a time.\n\n"
                "If you are sure nobody else is updating this folder, close Archive Search on other computers and delete the marker file shown below, then try again.\n\n"
                f"Folder: {root.path}\n"
                f"Marker file: {lock_path}\n\n"
                f"Details:\n{details}"
            ) from exc

        try:
            with lock_file:
                self._write_lock_details(lock_file, root, database_path)
                yield lock_path
        finally:
            try:
                lock_path.unlink()
            except FileNotFoundError:
                pass
            except OSError:
                pass
    @staticmethod
    def _root_identity(root_path: Path) -> str:
        """Return a canonical key used to identify duplicate selected folders."""
        try:
            normalized = str(root_path.expanduser().resolve(strict=False))
        except Exception:
            normalized = str(root_path.expanduser().absolute())
        return os.path.normcase(normalized)

    def _dedupe_roots(self, roots: Sequence[SearchRoot]) -> Tuple[SearchRoot, ...]:
        """Collapse duplicate selected roots while preserving user-facing labels.

        If the same folder is selected in both Primary and Secondary, searching it
        twice would duplicate results and can cause duplicated work. Duplicate paths are therefore merged into one root.

        If duplicate entries disagree on subfolder inclusion, the broader scope
        wins: include_subfolders=True.
        """
        merged: Dict[str, SearchRoot] = {}
        label_parts: Dict[str, List[str]] = {}
        for root in roots:
            key = self._root_identity(root.path)
            if key not in merged:
                merged[key] = root
                label_parts[key] = [root.label]
                continue

            existing = merged[key]
            if root.label not in label_parts[key]:
                label_parts[key].append(root.label)
            merged[key] = SearchRoot(
                label=" / ".join(label_parts[key]),
                path=existing.path,
                include_subfolders=existing.include_subfolders or root.include_subfolders,
            )

        return tuple(merged.values())

    def get_index_path(self) -> str:
        return "SQLite index files are stored inside each selected archive folder."

    @staticmethod
    def _metadata_key_for_allowed_extensions(prefix: str, allowed_extensions: Sequence[str]) -> str:
        suffix = "|".join(sorted({ext.lower() for ext in allowed_extensions})) or "all"
        return f"{prefix}:{suffix}"

    def _set_metadata_value(self, conn: sqlite3.Connection, key: str, value: str) -> None:
        conn.execute(
            "INSERT INTO metadata(key, value) VALUES(?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, value),
        )

    def _get_metadata_value(self, conn: sqlite3.Connection, key: str) -> Optional[str]:
        row = conn.execute("SELECT value FROM metadata WHERE key=?", (key,)).fetchone()
        return str(row[0]) if row and row[0] is not None else None

    def _get_root_last_run_display(
        self,
        root: SearchRoot,
        allowed_extensions: Sequence[str],
    ) -> str:
        database_path = self.get_root_database_path(root)
        if not database_path.exists():
            return "Never"
        allowed = tuple(sorted({ext.lower() for ext in allowed_extensions}))
        conn: Optional[sqlite3.Connection] = None
        try:
            conn = self._connect_readonly(database_path)
            value = None
            if allowed:
                value = self._get_metadata_value(
                    conn,
                    self._metadata_key_for_allowed_extensions("last_refresh_completed_utc", allowed),
                )
            value = value or self._get_metadata_value(conn, "last_refresh_completed_utc")
        except Exception:
            value = None
        finally:
            if conn is not None:
                self._close_connection_quietly(conn)
        return format_utc_iso_for_display(value) if value else "Unknown"

    def describe_selected_index_paths(
        self,
        roots: Sequence[SearchRoot],
        allowed_extensions: Sequence[str] = (),
    ) -> str:
        roots = self._dedupe_roots(roots)
        if not roots:
            return "SQLite index files are stored inside each selected archive folder."
        lines = ["SQLite index files are stored inside the selected archive folders:"]
        for root in roots:
            database_path = self.get_root_database_path(root)
            last_run = self._get_root_last_run_display(root, allowed_extensions)
            lines.append(f"{root.label}: Most recent index run: {last_run} | {database_path.name}")
        return "\n".join(lines)

    @staticmethod
    def _expected_content_type_for_suffix(suffix: str) -> Optional[str]:
        suffix = suffix.lower()
        if suffix in WORKBOOK_EXTENSIONS:
            return "excel"
        if suffix in WORD_EXTENSIONS:
            return "word"
        if suffix in PDF_EXTENSIONS:
            return "pdf"
        if suffix in PLAIN_TEXT_EXTENSIONS:
            return "text"
        return None

    @staticmethod
    def _describe_expected_content_type(expected_type: Optional[str], suffix: str) -> str:
        suffix_text = suffix or "no extension"
        labels = {
            "excel": "Excel workbook",
            "word": "Word document",
            "pdf": "PDF document",
            "text": "plain text document",
        }
        return f"{labels.get(expected_type, 'supported document')} ({suffix_text})"

    @staticmethod
    def _is_text_content_kind_acceptable(content_kind: Optional[str]) -> bool:
        return content_kind in {"text", "empty"}

    def _get_file_type_issue(self, file_path: Path, root: SearchRoot) -> Optional[Dict[str, str]]:
        """Return a blocking issue when a filename extension does not match the file content."""
        suffix = file_path.suffix.lower()
        expected_type = self._expected_content_type_for_suffix(suffix)
        if expected_type is None:
            return None

        try:
            content_kind = detect_file_content_kind(file_path)
            supported_type = detect_supported_file_type(file_path)
            actual_description = describe_file_content(file_path)
        except Exception as exc:
            return {
                "root_label": root.label,
                "file_name": file_path.name,
                "folder": str(file_path.parent),
                "path": str(file_path),
                "expected": self._describe_expected_content_type(expected_type, suffix),
                "actual": f"could not be checked ({exc})",
            }

        if expected_type == "text":
            mismatch = not self._is_text_content_kind_acceptable(content_kind)
        else:
            mismatch = supported_type != expected_type

        if not mismatch:
            return None

        return {
            "root_label": root.label,
            "file_name": file_path.name,
            "folder": str(file_path.parent),
            "path": str(file_path),
            "expected": self._describe_expected_content_type(expected_type, suffix),
            "actual": actual_description,
        }

    def find_blocking_file_type_issues(
        self,
        roots: Sequence[SearchRoot],
        allowed_extensions: Sequence[str],
        *,
        max_issues: int = 10,
    ) -> List[Dict[str, str]]:
        """Find files that should be fixed before an index update starts."""
        allowed = tuple(sorted({ext.lower() for ext in allowed_extensions}))
        issues: List[Dict[str, str]] = []
        for root in self._dedupe_roots(roots):
            for file_path in self._iter_root_files(root, allowed):
                issue = self._get_file_type_issue(file_path, root)
                if issue is not None:
                    issues.append(issue)
                    if len(issues) >= max_issues:
                        return issues
        return issues

    def _connect(self, database_path: Path) -> sqlite3.Connection:
        """Open a read/write connection. Only indexing/update code should use this."""
        database_path.parent.mkdir(parents=True, exist_ok=True)
        connection = sqlite3.connect(database_path, timeout=30.0)
        connection.row_factory = sqlite3.Row
        # DELETE journal mode avoids persistent WAL sidecar files and is simple to back up.
        connection.execute("PRAGMA journal_mode=DELETE")
        connection.execute("PRAGMA foreign_keys=ON")
        connection.execute("PRAGMA busy_timeout=30000")
        connection.execute("PRAGMA synchronous=NORMAL")
        return connection

    @staticmethod
    def _normalise_windows_path_for_sqlite_uri(path_text: str) -> str:
        r"""Return a Windows path string that SQLite can safely receive in a file URI.

        pathlib may resolve mapped drives such as S: to a UNC path. A normal
        pathlib URI for that UNC path would look like file://server/share/file,
        but the SQLite library bundled with Python rejects non-local URI
        authorities unless it was compiled with SQLITE_ALLOW_URI_AUTHORITY.

        For UNC paths, keep the server/share inside the URI path instead and
        use the allowed localhost authority, for example:
            \\server\share\db.sqlite3 -> file://localhost//server/share/db.sqlite3
        """
        if path_text.startswith("\\\\?\\UNC\\"):
            return r"\\" + path_text[8:]
        if path_text.startswith("\\\\?\\"):
            return path_text[4:]
        return path_text

    @classmethod
    def _sqlite_readonly_uri(cls, database_path: Path) -> str:
        """Build a SQLite read-only URI that also works for mapped/UNC drives."""
        resolved = database_path.expanduser().resolve(strict=False)
        path_text = cls._normalise_windows_path_for_sqlite_uri(str(resolved))

        # UNC paths start with two slashes/backslashes. Do not use Path.as_uri()
        # here because it creates file://server/share/... and SQLite rejects the
        # server name as an invalid URI authority.
        if path_text.startswith((r"\\", "//")):
            unc_path = path_text.replace("\\", "/")
            if not unc_path.startswith("//"):
                unc_path = "//" + unc_path.lstrip("/")
            encoded_path = quote(unc_path, safe="/:")
            return f"file://localhost{encoded_path}?{urlencode({'mode': 'ro'})}"

        try:
            uri = resolved.as_uri()
        except ValueError:
            uri_path = quote(resolved.as_posix(), safe="/:")
            uri = f"file:{uri_path}"
        separator = "&" if "?" in uri else "?"
        return f"{uri}{separator}{urlencode({'mode': 'ro'})}"

    def _connect_readonly(self, database_path: Path) -> sqlite3.Connection:
        """Open an existing database without creating or changing files.

        This is used for normal searches and status checks so many users can search
        at once without needing write permission or creating SQLite sidecar files.
        """
        readonly_uri = self._sqlite_readonly_uri(database_path)
        connection = sqlite3.connect(readonly_uri, uri=True, timeout=30.0)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA busy_timeout=30000")
        return connection

    @staticmethod
    def _has_required_index_tables(conn: sqlite3.Connection) -> bool:
        rows = conn.execute("SELECT name FROM sqlite_master WHERE type IN ('table', 'view')").fetchall()
        names = {str(row[0]) for row in rows}
        return {"metadata", "files", "chunk_index"}.issubset(names)

    def _has_usable_database(self, database_path: Path) -> bool:
        """Return True when an index file exists and has the expected tables."""
        if not database_path.exists():
            return False
        conn: Optional[sqlite3.Connection] = None
        try:
            conn = self._connect_readonly(database_path)
            return self._has_required_index_tables(conn)
        except sqlite3.Error:
            return False
        finally:
            if conn is not None:
                self._close_connection_quietly(conn)

    def get_index_build_requirements(self, roots: Sequence[SearchRoot]) -> Tuple[Dict[str, object], ...]:
        """Return selected roots where Update Index must build a new index first.

        The reason distinguishes a genuinely missing index from an existing file
        that is not a valid Archive Search SQLite index.
        """
        requirements: List[Dict[str, object]] = []
        for root in self._dedupe_roots(roots):
            database_path = self.get_root_database_path(root)
            if self._has_usable_database(database_path):
                continue
            requirements.append(
                {
                    "root": root,
                    "database_path": database_path,
                    "reason": "invalid" if database_path.exists() else "missing",
                }
            )
        return tuple(requirements)

    def get_roots_without_usable_index(self, roots: Sequence[SearchRoot]) -> Tuple[SearchRoot, ...]:
        """Return selected folders that do not yet have a usable saved index."""
        return tuple(item["root"] for item in self.get_index_build_requirements(roots))

    def _close_connection_quietly(self, conn: sqlite3.Connection) -> None:
        """Close a SQLite connection without hiding the real work error."""
        try:
            conn.close()
        except Exception:
            pass

    def _ensure_schema(self, database_path: Path) -> None:
        """Create or migrate a local per-root SQLite database if needed."""
        with self._schema_lock:
            conn = self._connect(database_path)
            try:
                conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS metadata (
                        key TEXT PRIMARY KEY,
                        value TEXT NOT NULL
                    )
                    """
                )
                current_version_row = conn.execute(
                    "SELECT value FROM metadata WHERE key='schema_version'"
                ).fetchone()
                current_version = int(current_version_row[0]) if current_version_row else None

                files_columns = (
                    "id",
                    "root_label",
                    "root_path",
                    "relative_path",
                    "path",
                    "file_type",
                    "suffix",
                    "size",
                    "mtime_ns",
                    "indexed_at_utc",
                    "last_error",
                )
                files_table_sql = """
                    CREATE TABLE IF NOT EXISTS files (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        root_label TEXT NOT NULL,
                        root_path TEXT NOT NULL,
                        relative_path TEXT NOT NULL,
                        path TEXT NOT NULL UNIQUE,
                        file_type TEXT NOT NULL,
                        suffix TEXT NOT NULL,
                        size INTEGER,
                        mtime_ns INTEGER,
                        indexed_at_utc TEXT NOT NULL,
                        last_error TEXT
                    )
                    """
                conn.execute(files_table_sql)

                existing_files_columns = tuple(row[1] for row in conn.execute("PRAGMA table_info(files)").fetchall())
                if existing_files_columns != files_columns:
                    # Rebuild the table when an older database shape is detected.
                    # Existing rows are preserved when all expected columns are available.
                    shared_columns = [name for name in files_columns if name in existing_files_columns]
                    conn.execute("DROP INDEX IF EXISTS idx_files_root_path")
                    conn.execute("DROP INDEX IF EXISTS idx_files_suffix")
                    conn.execute("ALTER TABLE files RENAME TO files_old")
                    conn.execute(files_table_sql)
                    if set(files_columns).issubset(set(shared_columns)):
                        column_list = ", ".join(files_columns)
                        conn.execute(
                            f"INSERT INTO files({column_list}) SELECT {column_list} FROM files_old"
                        )
                    conn.execute("DROP TABLE files_old")

                conn.execute("CREATE INDEX IF NOT EXISTS idx_files_root_path ON files(root_path)")
                conn.execute("CREATE INDEX IF NOT EXISTS idx_files_suffix ON files(suffix)")

                conn.execute(
                    """
                    CREATE VIRTUAL TABLE IF NOT EXISTS chunk_index USING fts5(
                        file_id UNINDEXED,
                        location_label UNINDEXED,
                        location_sort_key UNINDEXED,
                        sheet_name UNINDEXED,
                        row_number UNINDEXED,
                        line_number UNINDEXED,
                        page_number UNINDEXED,
                        content,
                        tokenize='unicode61 remove_diacritics 2',
                        prefix='2 3 4'
                    )
                    """
                )

                conn.execute(
                    "INSERT INTO metadata(key, value) VALUES('schema_version', ?) "
                    "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                    (str(INDEX_SCHEMA_VERSION),),
                )
                conn.commit()
            finally:
                self._close_connection_quietly(conn)

    def _iter_root_files(self, root: SearchRoot, allowed_extensions: Sequence[str]) -> Iterator[Path]:
        """Yield eligible files while skipping temporary/index artefacts."""
        allowed = {ext.lower() for ext in allowed_extensions}
        root_path = root.path
        db_path = self.get_root_database_path(root).resolve()
        if root.include_subfolders:
            for current_root, dirnames, filenames in os.walk(root_path):
                dirnames[:] = [name for name in dirnames if not should_ignore_filename(name)]
                for filename in sorted(filenames, key=str.lower):
                    if should_ignore_filename(filename):
                        continue
                    path = Path(current_root) / filename
                    try:
                        if path.resolve() == db_path:
                            continue
                    except Exception:
                        pass
                    if path.suffix.lower() in allowed:
                        yield path
        else:
            for entry in sorted(os.scandir(root_path), key=lambda item: item.name.lower()):
                if should_ignore_filename(entry.name):
                    continue
                try:
                    if entry.is_symlink() or not entry.is_file(follow_symlinks=False):
                        continue
                except OSError:
                    continue
                path = root_path / entry.name
                try:
                    if path.resolve() == db_path:
                        continue
                except Exception:
                    pass
                if path.suffix.lower() in allowed:
                    yield path

    def _fetch_existing_rows(self, conn: sqlite3.Connection, allowed_extensions: Sequence[str]) -> Dict[str, sqlite3.Row]:
        allowed = tuple(sorted({ext.lower() for ext in allowed_extensions}))
        if not allowed:
            return {}
        rows = conn.execute(
            f"SELECT * FROM files WHERE suffix IN ({','.join('?' for _ in allowed)})",
            (*allowed,),
        ).fetchall()
        return {row["path"]: row for row in rows}

    @staticmethod
    def _safe_stat(file_path: Path) -> Optional[os.stat_result]:
        try:
            return file_path.stat()
        except OSError:
            return None

    def _index_record_with_retries(
        self,
        *,
        file_path: Path,
        stat_result: os.stat_result,
        word_session: Optional[ReadOnlyWordSession] = None,
        retry_attempts: int = INDEX_RETRY_ATTEMPTS,
        retry_delay_seconds: float = INDEX_RETRY_DELAY_SECONDS,
    ) -> Tuple[Dict[str, object], os.stat_result]:
        """Retry transient read failures and verify a file is stable after extraction."""
        latest_stat = stat_result
        attempts = max(1, int(retry_attempts))
        last_record: Optional[Dict[str, object]] = None
        for attempt in range(1, attempts + 1):
            current_stat = self._safe_stat(file_path)
            if current_stat is None:
                return self.extractor.build_error_record(
                    file_path,
                    getattr(latest_stat, "st_size", None),
                    getattr(latest_stat, "st_mtime_ns", None),
                    f"File disappeared before it could be indexed: {file_path.name}",
                ), latest_stat
            latest_stat = current_stat
            record = self.extractor.build_index_record(
                file_path=file_path,
                size=current_stat.st_size,
                mtime_ns=current_stat.st_mtime_ns,
                word_session=word_session,
            )
            last_record = record
            if not record.get("error"):
                stable_before = (current_stat.st_size, current_stat.st_mtime_ns)
                time.sleep(INDEX_STABLE_PROBE_DELAY_SECONDS)
                after_stat = self._safe_stat(file_path)
                if after_stat is not None:
                    latest_stat = after_stat
                    stable_after = (after_stat.st_size, after_stat.st_mtime_ns)
                    if stable_after == stable_before:
                        return record, latest_stat
                if attempt < attempts:
                    time.sleep(retry_delay_seconds)
                    continue
                if after_stat is None:
                    return self.extractor.build_error_record(
                        file_path,
                        getattr(current_stat, "st_size", None),
                        getattr(current_stat, "st_mtime_ns", None),
                        f"File disappeared before it became stable: {file_path.name}",
                    ), current_stat
                return self.extractor.build_error_record(
                    file_path,
                    after_stat.st_size,
                    after_stat.st_mtime_ns,
                    f"File changed while being indexed and did not settle in time: {file_path.name}",
                ), after_stat
            if attempt < attempts:
                time.sleep(retry_delay_seconds)
                continue
            return record, latest_stat
        assert last_record is not None
        return last_record, latest_stat

    def _upsert_file_row(
        self,
        conn: sqlite3.Connection,
        *,
        root: SearchRoot,
        file_path: Path,
        stat_result,
        record: Dict[str, object],
    ) -> int:
        row = conn.execute(
            """
            INSERT INTO files(
                root_label, root_path, relative_path, path, file_type, suffix, size, mtime_ns,
                indexed_at_utc, last_error
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(path) DO UPDATE SET
                root_label=excluded.root_label,
                root_path=excluded.root_path,
                relative_path=excluded.relative_path,
                file_type=excluded.file_type,
                suffix=excluded.suffix,
                size=excluded.size,
                mtime_ns=excluded.mtime_ns,
                indexed_at_utc=excluded.indexed_at_utc,
                last_error=excluded.last_error
            RETURNING id
            """,
            (
                root.label,
                str(root.path),
                file_path.relative_to(root.path).as_posix(),
                str(file_path),
                record.get("file_type") or self.extractor.infer_file_type(file_path),
                file_path.suffix.lower(),
                getattr(stat_result, "st_size", None),
                getattr(stat_result, "st_mtime_ns", None),
                utc_now_iso(),
                record.get("error"),
            ),
        ).fetchone()
        return int(row[0])

    @staticmethod
    def _summarize_record_error(file_path: Path, record: Dict[str, object]) -> str:
        message = str(record.get("error") or "").strip()
        if not message:
            return ""
        if "full path:" in message.lower() or "folder:" in message.lower():
            return message
        return f"{file_path.name} in {file_path.parent}: {message}"

    @staticmethod
    def _record_indexing_error(stats: Dict[str, object], file_path: Path, record: Dict[str, object]) -> None:
        summary = SQLiteIndexManager._summarize_record_error(file_path, record)
        if not summary:
            return
        stats["files_failed"] = int(stats.get("files_failed", 0)) + 1
        if not str(stats.get("first_error") or ""):
            stats["first_error"] = summary

    def _replace_file_chunks(self, conn: sqlite3.Connection, file_id: int, chunks: Sequence[Dict[str, object]]) -> None:
        conn.execute("DELETE FROM chunk_index WHERE CAST(file_id AS INTEGER)=?", (file_id,))
        if not chunks:
            return
        conn.executemany(
            """
            INSERT INTO chunk_index(
                file_id, location_label, location_sort_key, sheet_name,
                row_number, line_number, page_number, content
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    str(file_id),
                    str(chunk.get("location_label") or ""),
                    str(chunk.get("location_sort_key") or 0),
                    chunk.get("sheet_name"),
                    None if chunk.get("row_number") is None else str(chunk.get("row_number")),
                    None if chunk.get("line_number") is None else str(chunk.get("line_number")),
                    None if chunk.get("page_number") is None else str(chunk.get("page_number")),
                    str(chunk.get("content") or ""),
                )
                for chunk in chunks
                if chunk.get("content")
            ],
        )

    def _validate_database(self, database_path: Path) -> None:
        """Check that a newly built database is safe to use before publishing it."""
        conn = self._connect_readonly(database_path)
        try:
            row = conn.execute("PRAGMA integrity_check").fetchone()
            result = str(row[0]) if row else ""
            if result.lower() != "ok":
                raise ArchiveIndexReplaceError(
                    "The new index did not pass its final safety check, so the previous index was kept."
                )
            required_tables = {"metadata", "files", "chunk_index"}
            rows = conn.execute("SELECT name FROM sqlite_master WHERE type IN ('table', 'view')").fetchall()
            names = {str(row[0]) for row in rows}
            missing = required_tables - names
            if missing:
                raise ArchiveIndexReplaceError(
                    "The new index was incomplete, so the previous index was kept."
                )
        finally:
            self._close_connection_quietly(conn)

    def _replace_live_database(self, *, temp_path: Path, database_path: Path, status_callback: Callable[[str], None]) -> None:
        """Replace the live index with a completed temp database, keeping a backup."""
        backup_path = self._backup_path_for_database(database_path)
        if database_path.exists():
            try:
                shutil.copy2(database_path, backup_path)
            except OSError:
                # A backup is useful, but the old live database is still in place if replace fails.
                pass

        gc.collect()
        last_error: Optional[BaseException] = None
        for attempt in range(1, 7):
            try:
                os.replace(temp_path, database_path)
                return
            except OSError as exc:
                last_error = exc
                if attempt == 1:
                    gc.collect()
                    status_callback(
                        "The new index is ready. Archive Search is retrying the final save in case the current index is briefly in use..."
                    )
                time.sleep(2)

        raise ArchiveIndexReplaceError(
            "The new index was built successfully, but the program could not put it in place because the current saved index file could not be replaced. "
            "The previous index has been kept. Make sure nobody is searching or updating this folder, then try again. If you are the only person using the program, close and reopen Archive Search before trying again."
        ) from last_error

    def _update_root_database(
        self,
        *,
        root: SearchRoot,
        database_path: Path,
        allowed: Sequence[str],
        cancel_event: threading.Event,
        status_callback: Callable[[str], None],
        stats: Dict[str, object],
    ) -> None:
        """Update only new, changed, and deleted files in the supplied database path."""
        self._ensure_schema(database_path)
        last_status_time = 0.0

        conn = self._connect(database_path)
        try:
            existing_rows = self._fetch_existing_rows(conn, allowed)
            scoped_existing_rows = self._filter_existing_rows_for_scope(root, existing_rows)
            seen_paths: set[str] = set()
            normal_jobs: List[Tuple[Path, os.stat_result]] = []
            legacy_jobs: List[Tuple[Path, os.stat_result]] = []

            status_callback(
                f"Checking the {root.label.lower()} folder for new, changed, or removed files..."
            )
            for file_path in self._iter_root_files(root, allowed):
                if cancel_event.is_set():
                    break
                try:
                    stat_result = file_path.stat()
                except OSError:
                    continue

                stats["files_seen"] += 1
                seen_paths.add(str(file_path))
                existing = scoped_existing_rows.get(str(file_path))
                suffix = file_path.suffix.lower()
                unchanged = (
                    existing is not None
                    and not existing["last_error"]
                    and int(existing["size"] or -1) == int(getattr(stat_result, "st_size", -2))
                    and int(existing["mtime_ns"] or -1) == int(getattr(stat_result, "st_mtime_ns", -2))
                )
                if unchanged:
                    stats["files_reused"] += 1
                else:
                    if suffix in WORD_LEGACY_EXTENSIONS and HAS_WIN32_COM:
                        legacy_jobs.append((file_path, stat_result))
                    else:
                        normal_jobs.append((file_path, stat_result))

                now = time.monotonic()
                if (now - last_status_time) >= STATUS_UPDATE_INTERVAL_SECONDS:
                    status_callback(
                        f"Checking the {root.label.lower()} folder... looked at {stats['files_seen']} {pluralize(stats['files_seen'], 'file')} so far"
                    )
                    last_status_time = now

            for batch_start in range(0, len(normal_jobs), INDEX_BATCH_SIZE):
                if cancel_event.is_set():
                    break
                batch = normal_jobs[batch_start: batch_start + INDEX_BATCH_SIZE]
                with conn:
                    for file_path, stat_result in batch:
                        record, latest_stat = self._index_record_with_retries(
                            file_path=file_path,
                            stat_result=stat_result,
                        )
                        file_id = self._upsert_file_row(
                            conn,
                            root=root,
                            file_path=file_path,
                            stat_result=latest_stat,
                            record=record,
                        )
                        self._replace_file_chunks(conn, file_id, record.get("chunks") or [])
                        self._record_indexing_error(stats, file_path, record)
                        stats["files_refreshed"] += 1
                status_callback(
                    f"Updating the saved index for the {root.label.lower()} folder... updated {stats['files_refreshed']} {pluralize(stats['files_refreshed'], 'file')}"
                )

            if legacy_jobs and not cancel_event.is_set():
                try:
                    with ReadOnlyWordSession() as word_session:
                        for batch_start in range(0, len(legacy_jobs), INDEX_BATCH_SIZE):
                            if cancel_event.is_set():
                                break
                            batch = legacy_jobs[batch_start: batch_start + INDEX_BATCH_SIZE]
                            with conn:
                                for file_path, stat_result in batch:
                                    record, latest_stat = self._index_record_with_retries(
                                        file_path=file_path,
                                        stat_result=stat_result,
                                        word_session=word_session,
                                    )
                                    file_id = self._upsert_file_row(
                                        conn,
                                        root=root,
                                        file_path=file_path,
                                        stat_result=latest_stat,
                                        record=record,
                                    )
                                    self._replace_file_chunks(conn, file_id, record.get("chunks") or [])
                                    self._record_indexing_error(stats, file_path, record)
                                    stats["files_refreshed"] += 1
                            status_callback(
                                f"Updating older Word files for the {root.label.lower()} folder... updated {stats['files_refreshed']} {pluralize(stats['files_refreshed'], 'file')}"
                            )
                except Exception as exc:
                    with conn:
                        for file_path, stat_result in legacy_jobs:
                            record = self.extractor.build_error_record(
                                file_path=file_path,
                                size=stat_result.st_size,
                                mtime_ns=stat_result.st_mtime_ns,
                                message=f"Could not start Microsoft Word safely to read this older file: {exc}",
                            )
                            file_id = self._upsert_file_row(
                                conn,
                                root=root,
                                file_path=file_path,
                                stat_result=stat_result,
                                record=record,
                            )
                            self._replace_file_chunks(conn, file_id, [])
                            self._record_indexing_error(stats, file_path, record)
                            stats["files_refreshed"] += 1

            missing_paths = sorted(set(scoped_existing_rows) - seen_paths)
            if missing_paths and not cancel_event.is_set():
                with conn:
                    for missing_path in missing_paths:
                        row = scoped_existing_rows.get(missing_path)
                        if row is None:
                            continue
                        conn.execute("DELETE FROM chunk_index WHERE CAST(file_id AS INTEGER)=?", (int(row["id"]),))
                        conn.execute("DELETE FROM files WHERE id=?", (int(row["id"]),))
                        stats["files_deleted"] += 1
                status_callback(
                    f"Removing deleted files from the {root.label.lower()} folder index... removed {stats['files_deleted']} {pluralize(stats['files_deleted'], 'file')}"
                )

            if not cancel_event.is_set():
                refresh_completed_utc = utc_now_iso()
                self._set_metadata_value(conn, "last_refresh_completed_utc", refresh_completed_utc)
                self._set_metadata_value(
                    conn,
                    self._metadata_key_for_allowed_extensions("last_refresh_completed_utc", allowed),
                    refresh_completed_utc,
                )

            if stats["files_refreshed"] > 0 or stats["files_deleted"] > 0:
                try:
                    conn.execute("INSERT INTO chunk_index(chunk_index) VALUES('optimize')")
                    conn.commit()
                except Exception:
                    pass

        finally:
            self._close_connection_quietly(conn)

    def _update_roots_safely(
        self,
        roots: Sequence[SearchRoot],
        allowed: Sequence[str],
        *,
        cancel_event: threading.Event,
        status_callback: Callable[[str], None],
    ) -> Dict[str, int]:
        """Incrementally update selected roots using lock files and temp databases."""
        stats: Dict[str, object] = {
            "files_seen": 0,
            "files_reused": 0,
            "files_refreshed": 0,
            "files_deleted": 0,
            "files_failed": 0,
            "first_error": "",
        }
        for root in roots:
            if cancel_event.is_set():
                break
            database_path = self.get_root_database_path(root)
            temp_path = self._temp_path_for_database(database_path)

            with self._root_update_lock(root, database_path, status_callback):
                self._remove_database_file_set(temp_path)

                status_callback(
                    f"Starting a safe update for the {root.label.lower()} folder. Searches can continue using the last completed index."
                )
                try:
                    if database_path.exists() and self._has_usable_database(database_path):
                        shutil.copy2(database_path, temp_path)
                    elif database_path.exists():
                        status_callback(
                            f"The existing index file for the {root.label.lower()} folder is not a valid Archive Search index. "
                            "It will be overwritten with a fresh index build."
                        )
                    self._update_root_database(
                        root=root,
                        database_path=temp_path,
                        allowed=allowed,
                        cancel_event=cancel_event,
                        status_callback=status_callback,
                        stats=stats,
                    )
                    if cancel_event.is_set():
                        try:
                            temp_path.unlink()
                        except OSError:
                            pass
                        break
                    status_callback(f"Checking the updated index for the {root.label.lower()} folder before using it...")
                    self._validate_database(temp_path)
                    status_callback(f"Putting the updated index in place for the {root.label.lower()} folder...")
                    self._replace_live_database(
                        temp_path=temp_path,
                        database_path=database_path,
                        status_callback=status_callback,
                    )
                    status_callback(f"The {root.label.lower()} folder index has been safely updated.")
                except Exception:
                    try:
                        if temp_path.exists():
                            temp_path.unlink()
                    except OSError:
                        pass
                    raise
        return stats

    def refresh_roots(
        self,
        roots: Sequence[SearchRoot],
        allowed_extensions: Sequence[str],
        *,
        cancel_event: threading.Event,
        status_callback: Callable[[str], None],
    ) -> Dict[str, int]:
        """Update selected SQLite/FTS indexes safely."""
        allowed = tuple(sorted({ext.lower() for ext in allowed_extensions}))
        if not allowed:
            return {
                "files_seen": 0,
                "files_reused": 0,
                "files_refreshed": 0,
                "files_deleted": 0,
                "files_failed": 0,
                "first_error": "",
            }

        roots = self._dedupe_roots(roots)
        return self._update_roots_safely(
            roots,
            allowed,
            cancel_event=cancel_event,
            status_callback=status_callback,
        )


    def get_stale_index_summary(
        self,
        roots: Sequence[SearchRoot],
        allowed_extensions: Sequence[str],
        *,
        cancel_event: Optional[threading.Event] = None,
    ) -> Dict[str, object]:
        """Return counts of files that appear newer than, missing from, or deleted from the index.

        This is intentionally read-only. It does not create or update the SQLite database.
        """
        allowed = tuple(sorted({ext.lower() for ext in allowed_extensions}))
        summary: Dict[str, object] = {
            "files_checked": 0,
            "indexed_files": 0,
            "new_files": 0,
            "changed_files": 0,
            "deleted_files": 0,
            "missing_databases": 0,
            "invalid_databases": 0,
            "roots_checked": 0,
            "errors": 0,
            "is_stale": False,
        }
        roots = self._dedupe_roots(roots)
        if not roots or not allowed:
            return summary

        for root in roots:
            if cancel_event is not None and cancel_event.is_set():
                break
            database_path = self.get_root_database_path(root)
            summary["roots_checked"] = int(summary["roots_checked"]) + 1

            existing_rows: Dict[str, sqlite3.Row] = {}
            if database_path.exists():
                conn: Optional[sqlite3.Connection] = None
                try:
                    conn = self._connect_readonly(database_path)
                    if not self._has_required_index_tables(conn):
                        summary["invalid_databases"] = int(summary["invalid_databases"]) + 1
                        existing_rows = {}
                    else:
                        existing_rows = self._fetch_existing_rows(conn, allowed)
                        existing_rows = self._filter_existing_rows_for_scope(root, existing_rows)
                        summary["indexed_files"] = int(summary["indexed_files"]) + len(existing_rows)
                except Exception:
                    summary["invalid_databases"] = int(summary["invalid_databases"]) + 1
                    existing_rows = {}
                finally:
                    if conn is not None:
                        self._close_connection_quietly(conn)
            else:
                summary["missing_databases"] = int(summary["missing_databases"]) + 1

            seen_paths: set[str] = set()
            for file_path in self._iter_root_files(root, allowed):
                if cancel_event is not None and cancel_event.is_set():
                    break
                try:
                    stat_result = file_path.stat()
                except OSError:
                    summary["errors"] = int(summary["errors"]) + 1
                    continue

                summary["files_checked"] = int(summary["files_checked"]) + 1
                path_key = str(file_path)
                seen_paths.add(path_key)
                existing = existing_rows.get(path_key)
                if existing is None:
                    summary["new_files"] = int(summary["new_files"]) + 1
                    continue
                try:
                    same_size = int(existing["size"] or -1) == int(getattr(stat_result, "st_size", -2))
                    same_mtime = int(existing["mtime_ns"] or -1) == int(getattr(stat_result, "st_mtime_ns", -2))
                except Exception:
                    same_size = False
                    same_mtime = False
                if not (same_size and same_mtime):
                    summary["changed_files"] = int(summary["changed_files"]) + 1

            deleted = len(set(existing_rows) - seen_paths)
            summary["deleted_files"] = int(summary["deleted_files"]) + deleted

        summary["is_stale"] = any(
            int(summary[key]) > 0
            for key in ("new_files", "changed_files", "deleted_files", "missing_databases", "invalid_databases")
        )
        return summary

    def count_indexed_files(self, roots: Sequence[SearchRoot], allowed_extensions: Sequence[str]) -> int:
        allowed = tuple(sorted({ext.lower() for ext in allowed_extensions}))
        roots = self._dedupe_roots(roots)
        if not roots or not allowed:
            return 0
        total = 0
        for root in roots:
            database_path = self.get_root_database_path(root)
            if not database_path.exists():
                continue
            try:
                conn = self._connect_readonly(database_path)
            except sqlite3.Error:
                continue
            try:
                scope_sql, scope_params = self._scope_sql_condition(root)
                row = conn.execute(
                    f"SELECT COUNT(*) FROM files WHERE suffix IN ({','.join('?' for _ in allowed)}){scope_sql}",
                    (*allowed, *scope_params),
                ).fetchone()
                total += int(row[0]) if row else 0
            finally:
                self._close_connection_quietly(conn)
        return total

    def get_indexed_document_text(self, result: SearchResult) -> str:
        """Return the indexed full text for a selected non-workbook result."""
        if result.file_type == "excel":
            return ""
        root_path_text = (result.root_path or "").strip()
        if not root_path_text:
            return ""
        database_path = self.get_root_database_path(Path(root_path_text))
        if not database_path.exists():
            return ""
        try:
            conn = self._connect_readonly(database_path)
        except sqlite3.Error:
            return ""
        try:
            row = conn.execute("SELECT id FROM files WHERE path=?", (result.path,)).fetchone()
            if row is None:
                return ""
            file_id = int(row["id"])
            rows = conn.execute(
                """
                SELECT content
                FROM chunk_index
                WHERE CAST(file_id AS INTEGER)=?
                ORDER BY CAST(COALESCE(location_sort_key, '0') AS INTEGER), location_label
                """,
                (file_id,),
            ).fetchall()
            parts = [
                str(item["content"] or "").strip()
                for item in rows
                if str(item["content"] or "").strip()
            ]
            return "\n\n".join(parts)
        except Exception:
            return ""
        finally:
            self._close_connection_quietly(conn)

    def _search_one_root(
        self,
        *,
        root: SearchRoot,
        allowed_extensions: Sequence[str],
        search_terms: Sequence[str],
        match_mode: str,
        match_query: str,
        cancel_event: threading.Event,
        status_callback: Callable[[str], None],
        snippets_per_file: Optional[int],
    ) -> Tuple[List[SearchResult], int]:
        allowed = tuple(sorted({ext.lower() for ext in allowed_extensions}))
        database_path = self.get_root_database_path(root)
        if not database_path.exists():
            return [], 0

        scope_sql, scope_params = self._scope_sql_condition(root)
        match_sql = f"""
            SELECT
                files.id AS file_id,
                files.path,
                files.root_path,
                files.file_type,
                chunk_index.location_label,
                chunk_index.sheet_name,
                chunk_index.row_number,
                chunk_index.line_number,
                chunk_index.page_number,
                chunk_index.location_sort_key,
                snippet(chunk_index, 7, '', '', ' ... ', 18) AS preview,
                bm25(chunk_index) AS score
            FROM chunk_index
            JOIN files ON files.id = CAST(chunk_index.file_id AS INTEGER)
            WHERE chunk_index MATCH ?
              AND files.suffix IN ({','.join('?' for _ in allowed)}){scope_sql}
            ORDER BY score, files.path, CAST(COALESCE(chunk_index.location_sort_key, '0') AS INTEGER), chunk_index.location_label
        """

        grouped: Dict[int, Dict[str, object]] = {}
        total_matches = 0

        try:
            conn = self._connect_readonly(database_path)
        except sqlite3.Error:
            status_callback(f"The {root.label.lower()} folder has no usable index yet. Please click Update Index when you are ready to create it.")
            return [], 0

        try:
            status_callback(f"Searching the saved index in the {root.label.lower()} folder...")
            rows = conn.execute(match_sql, (match_query, *allowed, *scope_params))
            for index, row in enumerate(rows, start=1):
                if cancel_event.is_set():
                    break
                file_id = int(row["file_id"])
                total_matches += 1
                info = grouped.get(file_id)
                if info is None:
                    info = {
                        "file_type": str(row["file_type"]),
                        "path": str(row["path"]),
                        "root_label": root.label,
                        "root_path": str(row["root_path"] or root.path),
                        "match_count": 0,
                        "best_score": float(row["score"]),
                        "snippets": [],
                        "open_sheet": None,
                        "open_row_number": None,
                        "first_location_label": None,
                    }
                    grouped[file_id] = info
                info["match_count"] = int(info["match_count"]) + 1
                if info["first_location_label"] is None:
                    info["first_location_label"] = str(row["location_label"] or "")
                if snippets_per_file is None or len(info["snippets"]) < snippets_per_file:
                    preview = truncate_text(str(row["preview"] or ""), DETAIL_PREVIEW_LENGTH)
                    info["snippets"].append(MatchSnippet(location_label=str(row["location_label"] or "Match"), preview=preview))
                if info["file_type"] == "excel" and info["open_sheet"] is None:
                    info["open_sheet"] = row["sheet_name"]
                    info["open_row_number"] = int(row["row_number"]) if row["row_number"] not in {None, ""} else None
                if index == 1 or (index % 100 == 0):
                    status_callback(
                        f"Searching the saved index in the {root.label.lower()} folder... checked {index} {pluralize(index, 'match')}"
                    )

            if not grouped:
                like_clause, like_params = self._build_like_conditions(search_terms=search_terms, match_mode=match_mode)
                if like_clause:
                    fallback_sql = f"""
                        SELECT
                            files.id AS file_id,
                            files.path,
                            files.root_path,
                            files.file_type,
                                        chunk_index.location_label,
                            chunk_index.sheet_name,
                            chunk_index.row_number,
                            chunk_index.line_number,
                            chunk_index.page_number,
                            chunk_index.location_sort_key,
                            chunk_index.content AS preview
                        FROM chunk_index
                        JOIN files ON files.id = CAST(chunk_index.file_id AS INTEGER)
                        WHERE files.suffix IN ({','.join('?' for _ in allowed)}){scope_sql}
                          AND ({like_clause})
                        ORDER BY files.path, CAST(COALESCE(chunk_index.location_sort_key, '0') AS INTEGER), chunk_index.location_label
                    """
                    rows = conn.execute(fallback_sql, (*allowed, *scope_params, *like_params))
                    for row in rows:
                        if cancel_event.is_set():
                            break
                        file_id = int(row["file_id"])
                        total_matches += 1
                        info = grouped.get(file_id)
                        if info is None:
                            info = {
                                "file_type": str(row["file_type"]),
                                "path": str(row["path"]),
                                "root_label": root.label,
                                "root_path": str(row["root_path"] or root.path),
                                "match_count": 0,
                                "best_score": 0.0,
                                "snippets": [],
                                "open_sheet": None,
                                "open_row_number": None,
                                "first_location_label": None,
                                    }
                            grouped[file_id] = info
                        info["match_count"] = int(info["match_count"]) + 1
                        if info["first_location_label"] is None:
                            info["first_location_label"] = str(row["location_label"] or "")
                        if snippets_per_file is None or len(info["snippets"]) < snippets_per_file:
                            preview = truncate_text(str(row["preview"] or ""), DETAIL_PREVIEW_LENGTH)
                            info["snippets"].append(MatchSnippet(location_label=str(row["location_label"] or "Match"), preview=preview))
                        if info["file_type"] == "excel" and info["open_sheet"] is None:
                            info["open_sheet"] = row["sheet_name"]
                            info["open_row_number"] = int(row["row_number"]) if row["row_number"] not in {None, ""} else None

        finally:
            self._close_connection_quietly(conn)

        file_results = [
            SearchResult(
                file_type=str(item["file_type"]),
                document_name=Path(str(item["path"])).name,
                path=str(item["path"]),
                root_label=str(item["root_label"]),
                match_count=int(item["match_count"]),
                snippets=tuple(item["snippets"]),
                root_path=str(item.get("root_path") or ""),
                open_sheet=item["open_sheet"],
                open_row_number=item["open_row_number"],
                first_location_label=item["first_location_label"],
            )
            for item in grouped.values()
        ]
        return file_results, total_matches

    def search(
        self,
        roots: Sequence[SearchRoot],
        allowed_extensions: Sequence[str],
        search_terms: Sequence[str],
        match_mode: str,
        *,
        cancel_event: threading.Event,
        status_callback: Callable[[str], None],
        result_limit: int = RESULT_LIMIT,
        snippets_per_file: Optional[int] = DETAIL_SNIPPETS_PER_FILE,
    ) -> Tuple[List[SearchResult], int]:
        roots = self._dedupe_roots(roots)
        if not search_terms or not roots:
            return [], 0

        allowed = tuple(sorted({ext.lower() for ext in allowed_extensions}))
        if not allowed:
            return [], 0

        match_query = self.build_match_query(search_terms, match_mode)
        all_results: List[SearchResult] = []
        total_matches = 0

        for root in roots:
            if cancel_event.is_set():
                break
            root_results, root_total_matches = self._search_one_root(
                root=root,
                allowed_extensions=allowed,
                search_terms=search_terms,
                match_mode=match_mode,
                match_query=match_query,
                cancel_event=cancel_event,
                status_callback=status_callback,
                snippets_per_file=snippets_per_file,
            )
            all_results.extend(root_results)
            total_matches += root_total_matches

        ordered = sorted(
            all_results,
            key=lambda item: (str(item.root_label).lower(), str(item.path).lower()),
        )[:result_limit]
        return ordered, total_matches
