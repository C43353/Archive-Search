# Archive Search

Archive Search is a desktop search tool for archived sales and technical documents.

The application builds a per-folder SQLite full-text index from supported archive files. Normal searches read the last completed index instead of opening every document again, so repeat searches are fast and multiple users can search a shared archive at the same time.

> **Important:** This project was developed with AI assistance. Review it, test it on dummy or copied data first, and do not rely on it as the only way to access important business records. (I didn't intend to use AI as much as I have, but it has done a better/faster job than I could have done in the timeframe)

## What it looks like

The screenshots below use sample files and the default `Company Name` branding. Your archive folders, company name, logo and icon can be configured locally.

## What the application does

Archive Search lets a user:

1. Choose a primary archive folder, a secondary archive folder, or both.
2. Decide whether each archive should include subfolders.
3. Choose whether to search workbooks, text documents, or both.
4. Click **Archive settings** > **Update index** to build or refresh the saved index.
5. Search from the main screen using any-word or all-words matching.
6. Select a result to review the matching evidence before opening the source file.
7. Open a selected result, show it in the file manager, or copy its full path.

Searches use the last completed SQLite index. New, edited, or deleted source files are included after **Update index** is run again. The main **Check index** button compares the selected folders against the saved index and reports whether the index appears out of date without changing anything.

## Supported file types

| Category | Extensions | Notes |
| --- | --- | --- |
| Excel workbooks | `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xls`, `.xlsb` | `.xlsx`/`.xlsm` are read through OpenPyXL; `.xls` uses xlrd; `.xlsb` uses pyxlsb. |
| Word documents | `.docx`, `.docm`, `.dotx`, `.dotm`, `.doc` | Modern Word files use python-docx when available. Legacy `.doc` requires Microsoft Word and pywin32. |
| PDFs | `.pdf` | Only embedded/selectable text is indexed. Scanned image-only PDFs need OCR outside this app first. |
| Plain text | `.txt` | Several common encodings are attempted. |

## Key features

- SQLite FTS5 indexing, with one index database stored in each selected archive folder.
- Any-word and all-words search modes.
- Quoted phrase support, for example `"Sigma 5E"`.
- Numeric/code substring fallback so a search such as `253` can still find a code such as `A9253`.
- Read-only search mode: normal searches open the existing index without writing to it.
- Incremental **Update index** workflow that adds new/changed files and removes deleted files from the index.
- **Check index** workflow that reports new, changed, deleted, missing-index, or invalid-index conditions without updating the index.
- File signature checks before indexing, so mismatched extensions such as a renamed file are flagged before an index update continues.
- Grouped result list with **File**, **Type**, and **Matches** columns.
- Details pane showing matching workbook rows or indexed document text with highlighted search terms.
- Workbook result details grouped by worksheet, with workbook headings shown where available.
- Double-click a workbook result to open it read-only at the first matching sheet/row when Microsoft Excel integration is available.
- Buttons and right-click actions for **Open selected**, **Show in folder**, and **Copy path**.
- Background worker thread for search and update jobs, keeping the Tkinter interface responsive.
- Optional local logo, app icon, and system-tray icon support.

## Requirements

- Python 3.10 or newer is recommended.
- Windows is recommended for the best Microsoft Office integration.
- Microsoft Excel is needed only for the optional “open workbook at matching row” integration.
- Microsoft Word plus pywin32 is needed for legacy `.doc` extraction and read-only Word opening.

Install dependencies from the project folder:

```bash
pip install -r requirements.txt
```

`pywin32` is listed with a Windows-only environment marker, so non-Windows systems should ignore it automatically. `Pillow` improves logo/icon scaling. `pystray` enables the optional system-tray icon when an app icon is configured.

## Running the program

From the project folder:

```bash
python main.py
```

The application starts with safe generic defaults:

- Primary folder: `C:\Primary Folder`
- Secondary folder: `C:\Secondary Folder`
- Company name: `Company Name`
- No logo or app icon

You can change the folders in the UI each time you run the program. For permanent local defaults, create a local config file as described below.

## Local configuration

Company-specific paths and branding should not be hard-coded into the shared application files. Instead, copy the template file:

```bash
copy archive_search_local_config.example.py archive_search_local_config.py
```

On PowerShell or macOS/Linux shells, use the normal copy command for that environment, for example:

```bash
cp archive_search_local_config.example.py archive_search_local_config.py
```

Then edit `archive_search_local_config.py` on that computer. The file is ignored by git through `.gitignore`.

Example:

```python
from pathlib import Path

LOCAL_COMPANY_NAME = "Company Name"

# Optional: show a logo in the header. Relative paths are resolved from this project folder.
LOCAL_COMPANY_LOGO_PATH = None
# LOCAL_COMPANY_LOGO_PATH = Path(r"C:\Path\To\logo.png")

LOCAL_HEADER_LOGO_MAX_WIDTH = 360
LOCAL_HEADER_LOGO_MAX_HEIGHT = 74

# Optional: title-bar/taskbar icon, and tray icon when pystray is installed.
LOCAL_APP_ICON_PATH = None
# LOCAL_APP_ICON_PATH = Path(r"C:\Path\To\icon.ico")

LOCAL_SEARCH_PRIMARY_FOLDER = Path(r"C:\Primary Folder")
LOCAL_SEARCH_SECONDARY_FOLDER = Path(r"C:\Secondary Folder")
```

Relative logo/icon paths are resolved from the application folder. Transparent PNG logos work well because the header has a white backing surface. `.ico` files are best for Windows title-bar/taskbar icons.

## Typical user workflow

### First-time setup

1. Start the app with `python main.py`.
2. Click **Archive settings**.
3. Enable the archive folder or folders you want to search.
4. Browse to the correct primary and/or secondary folder.
5. Tick **Subfolders** for each folder when nested archive folders should be included.
6. Choose **Workbooks**, **Text documents**, or both.
7. Click **Update index** and wait for it to finish.
8. Enter a search term and click **Search archive**.

The first index build may take a while on a large archive. Later updates are incremental and should usually be faster.

### Day-to-day searching

1. Type a table model, customer, order number, machine name, or phrase.
2. Use **Any word** when any term can match.
3. Use **All words** when every term must be present.
4. Use quotes for an exact phrase, for example `"Sigma 5E"`.
5. Select a result to review the matching evidence in the details pane.
6. Double-click the result, or click **Open selected**, to open the source document.

### Keeping the index current

- Click **Check index** when results look out of date.
- Click **Archive settings** > **Update index** after archive files are added, edited, moved, renamed, or deleted.
- Only one person or computer should run **Update index** for the same archive folder at the same time.
- Other users can continue to search the last completed index while an update is being prepared.

## Index location and privacy

Each selected archive folder stores its own SQLite database directly in that folder. For example:

```text
C:\Primary Folder\.archive_search_{path-code}.sqlite3
C:\Secondary Folder\.archive_search_{path-code}.sqlite3
```

`{path-code}` is a stable 16-character hash of the canonical folder path. The index is tied to the folder path, not to the Primary/Secondary UI slot. If the same folder is entered in both slots, Archive Search deduplicates it and searches it once.

The index contains extracted searchable text, file metadata, and status information. Treat `.archive_search_*.sqlite3` as sensitive if the source documents are sensitive. Users who only search need read access to the folder and index file. Users who run **Update index** need create/read/write/delete access in the archive folder so the app can create its `.lock`, `.tmp`, and `.bak` safety files.

## Multi-user and shared-folder behaviour

Normal searches open the saved index in read-only mode. This allows several people to search the same shared archive folder without changing the index and without creating SQLite sidecar files.

**Update index** uses a simple `.lock` marker beside the index so only one update can run for the same folder at once. The update is prepared in a temporary file first, the previous index is kept available during preparation, and a `.bak` backup is used when replacing the live index. If replacement fails, the previous index is kept.

When Windows resolves a mapped drive such as `S:` to a UNC path like `\\server\share`, the app builds a UNC-safe read-only SQLite URI so normal searches can still open the saved index.

## Safety notes

Archive Search is designed to avoid modifying source documents:

- Normal searches do not write to source documents or to the saved index.
- **Update index** writes only Archive Search index/safety files.
- Source files are opened through read-only extraction paths where supported.
- Legacy Word files are copied to a temporary read-only file before text extraction.
- File signatures are checked before indexing. If the extension and content type do not match, the update stops and reports the file name, folder, full path, expected type, and detected type.

No desktop search tool can guarantee every third-party library or Office integration will never affect a file, so test on copied data first and keep normal backups of important archives.

## Notes and limitations

- PDFs are searchable only when they contain embedded/selectable text.
- `.xlsb` support requires `pyxlsb`.
- `.doc` support requires Microsoft Word and `pywin32` on Windows.
- The app is primarily intended for Windows, although much of the indexing/search logic is cross-platform.
- When **Subfolders** is unticked, previously indexed subfolder entries remain in the database but are ignored for searching, status checks, and top-level refreshes until **Subfolders** is ticked again.
- The result limit is configured in `archive_config.py` as `RESULT_LIMIT`.

## Main project files

| File | Purpose |
| --- | --- |
| `main.py` | Launch entry point. |
| `archive_app.py` | Tkinter desktop UI and result rendering. |
| `archive_runner.py` | Background search/update worker that communicates with the UI queue. |
| `archive_index.py` | SQLite/FTS index creation, refresh, validation, status checking, and searching. |
| `archive_extractor.py` | Read-only text extraction for Excel, Word, PDF, and TXT files. |
| `archive_file_safety.py` | Low-level file access and file-signature checks. |
| `archive_opener.py` | Safer helpers for opening results and revealing files. |
| `archive_config.py` | Generic defaults, file-type sets, index constants, and UI colours. |
| `archive_search_local_config.example.py` | Template for local-only paths and branding. |
| `archive_optional.py` | Optional dependency imports for Word/PDF/Windows integration. |
| `archive_models.py` | Shared dataclasses. |
| `archive_utils.py` | Shared formatting and helper functions. |
| `requirements.txt` | Python package dependencies. |
