"""Background worker that coordinates indexing and SQLite/FTS searching."""

from __future__ import annotations

import queue
import threading
import time
from typing import Dict, List, Sequence

from archive_config import RESULT_BATCH_SIZE
from archive_index import SQLiteIndexManager
from archive_models import SearchRoot


class SearchRunner:
    """Coordinate index refresh and FTS search work on a background thread."""

    def __init__(
        self,
        index_manager: SQLiteIndexManager,
        queue_out: queue.Queue,
        cancel_event: threading.Event,
    ) -> None:
        self.index_manager = index_manager
        self.queue = queue_out
        self.cancel_event = cancel_event

    def _status(self, message: str) -> None:
        self.queue.put(("status", message))

    def _emit_items(self, items: List[Dict[str, object]]) -> None:
        if items:
            self.queue.put(("items", items.copy()))
            items.clear()

    def _finish(
        self,
        started_at: float,
        *,
        files_scanned: int = 0,
        indexed_files: int = 0,
        matched_files: int = 0,
        total_matches: int = 0,
        cancelled: bool = False,
        mode: str = "search",
        **extra: object,
    ) -> None:
        """Publish a finished payload with the standard result fields."""
        payload: Dict[str, object] = {
            "files_scanned": files_scanned,
            "indexed_files": indexed_files,
            "matched_files": matched_files,
            "total_matches": total_matches,
            "cancelled": cancelled,
            "elapsed_seconds": time.perf_counter() - started_at,
            "mode": mode,
        }
        payload.update(extra)
        self.queue.put(("finished", payload))

    @staticmethod
    def _empty_refresh_stats() -> Dict[str, object]:
        """Return index-refresh counters used when a normal search does not update the index."""
        return {
            "files_seen": 0,
            "files_reused": 0,
            "files_refreshed": 0,
            "files_deleted": 0,
            "files_failed": 0,
            "first_error": "",
        }

    def run(
        self,
        roots: Sequence[SearchRoot],
        allowed_extensions: Sequence[str],
        search_terms: Sequence[str],
        match_mode: str,
        *,
        update_only: bool = False,
    ) -> None:
        started_at = time.perf_counter()
        try:
            if update_only:
                self._status("Preparing to update the saved index with recent changes...")
                refresh_stats = self.index_manager.refresh_roots(
                    roots=roots,
                    allowed_extensions=allowed_extensions,
                    cancel_event=self.cancel_event,
                    status_callback=self._status,
                )
            else:
                self._status(
                    "Searching the last completed index. New or changed files are included after you click Archive settings and run Update index."
                )
                refresh_stats = self._empty_refresh_stats()

            if not update_only and search_terms:
                build_requirements = self.index_manager.get_index_build_requirements(roots)
                if build_requirements:
                    root_names = ", ".join(str(item["root"].label) for item in build_requirements)
                    self._status(
                        "The saved index is missing or is not a valid Archive Search index. Click Archive settings, then click Update index first, then search again."
                    )
                    self._finish(
                        started_at,
                        mode="initial_index_required",
                        files_reused=0,
                        files_refreshed=0,
                        files_deleted=0,
                        missing_roots=root_names,
                    )
                    return

            indexed_files = self.index_manager.count_indexed_files(roots, allowed_extensions)
            if self.cancel_event.is_set():
                self._finish(
                    started_at,
                    files_scanned=int(refresh_stats["files_seen"]),
                    indexed_files=indexed_files,
                    cancelled=True,
                    mode="update" if update_only else "search",
                )
                return

            if update_only:
                self._finish(
                    started_at,
                    files_scanned=int(refresh_stats["files_seen"]),
                    indexed_files=indexed_files,
                    mode="update",
                    files_reused=refresh_stats["files_reused"],
                    files_refreshed=refresh_stats["files_refreshed"],
                    files_deleted=refresh_stats["files_deleted"],
                    files_failed=refresh_stats.get("files_failed", 0),
                    first_error=refresh_stats.get("first_error", ""),
                )
                return

            if not search_terms:
                self._finish(
                    started_at,
                    files_scanned=int(refresh_stats["files_seen"]),
                    indexed_files=indexed_files,
                    mode="search",
                    files_reused=refresh_stats["files_reused"],
                    files_refreshed=refresh_stats["files_refreshed"],
                    files_deleted=refresh_stats["files_deleted"],
                    files_failed=refresh_stats.get("files_failed", 0),
                    first_error=refresh_stats.get("first_error", ""),
                )
                return

            results, total_matches = self.index_manager.search(
                roots=roots,
                allowed_extensions=allowed_extensions,
                search_terms=search_terms,
                match_mode=match_mode,
                cancel_event=self.cancel_event,
                status_callback=self._status,
            )

            item_buffer: List[Dict[str, object]] = []
            for result in results:
                if self.cancel_event.is_set():
                    break
                item_buffer.append({"kind": "result", "payload": result})
                if len(item_buffer) >= RESULT_BATCH_SIZE:
                    self._emit_items(item_buffer)
            self._emit_items(item_buffer)

            self._finish(
                started_at,
                files_scanned=int(refresh_stats["files_seen"]),
                indexed_files=indexed_files,
                matched_files=len(results),
                total_matches=total_matches,
                cancelled=self.cancel_event.is_set(),
                mode="search",
                files_reused=refresh_stats["files_reused"],
                files_refreshed=refresh_stats["files_refreshed"],
                files_deleted=refresh_stats["files_deleted"],
            )
        except Exception as exc:
            if self.cancel_event.is_set():
                self._finish(
                    started_at,
                    cancelled=True,
                    mode="update" if update_only else "search",
                )
            else:
                self.queue.put(("fatal", str(exc)))
