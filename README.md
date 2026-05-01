# Archive Search V11

Archive Search V11 is a Windows-friendly Tkinter desktop search tool for archived documents. It builds a SQLite full-text index so repeat searches are much faster after the first index build.

**Disclaimer:** This code was developed with AI assistance. Review it, test it on dummy data first, and do not use it on the only copy of important business records. (I didn't intend to use AI as much as I have, but it has done a better/faster job than I could have done in the timeframe)

## What the application does

The program searches one or two configured root folders for supported files. It extracts searchable text using read-only workflows where supported, stores the extracted text in a SQLite database, then searches that database.

Supported file types:

- Excel workbooks: `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xls`, `.xlsb`
- Word documents: `.docx`, `.docm`, `.dotx`, `.dotm`, `.doc`
- PDF documents: `.pdf`
- Plain text files: `.txt`

## Key features

- Search across one or two folder roots
- Optional recursive subfolder searching per root
- Search mode for **any word** or **all words**
- Quoted phrase support, for example `"Sigma 5E"`
- SQLite full-text search using one database per selected root folder
- Manual **Update Index** option that adds new/changed files and removes deleted files without starting over
- **Check Index Status** option to detect new, changed, deleted, or not-yet-indexed files
- Grouped result list with a details pane
- Result details show the folder path and highlighted search terms; text documents show their full indexed text
- Buttons for **Open Selected** and **Copy Path**
- Read-only extraction workflow where supported
- Background worker thread for search and update jobs so the UI stays responsive

## Main files

| File | Purpose |
| --- | --- |
| `main.py` | Launch entry point |
| `archive_app.py` | Tkinter desktop UI |
| `archive_runner.py` | Background indexing/search worker |
| `archive_index.py` | SQLite/FTS index management |
| `archive_extractor.py` | Read-only extraction for Excel, Word, PDF, and TXT files |
| `archive_opener.py` | Safer result-opening helpers |
| `archive_file_safety.py` | Low-level file access and signature checks |
| `archive_config.py` | File types, default folder paths, and UI constants |
| `archive_optional.py` | Optional third-party and Windows integration imports |
| `archive_models.py` | Shared dataclasses |
| `archive_utils.py` | Shared formatting and helper functions |
| `requirements.txt` | Python package dependencies |

## Requirements

- Python 3.10+ recommended
- Windows recommended for the best Microsoft Office integration

Install dependencies from the project folder:

```bash
pip install -r requirements.txt
```

Optional Windows note: `pywin32` is included in `requirements.txt` using a Windows-only environment marker, so non-Windows systems should ignore it automatically.

## Optional external software

Some features depend on software installed on the machine:

- **Microsoft Word + pywin32**: needed for legacy `.doc` extraction and read-only Word opening
- **Microsoft Excel + pywin32**: used to open Excel results directly to the matching sheet/row in read-only mode
- Without those integrations, the app falls back to safer generic opening behaviour where possible

## Running the program

Run the app from the project folder:

```bash
python main.py
```

## Default folders

The public code uses generic default folders:

- `C:\Primary Folder`
- `C:\Secondary Folder`

You can change the folders in the UI at runtime.

To make the app open with different default folders already filled in, edit these lines near the top of `archive_config.py`:

```python
ARCHIVE_SEARCH_PRIMARY_FOLDER = Path(r"C:\Primary Folder")
ARCHIVE_SEARCH_SECONDARY_FOLDER = Path(r"C:\Secondary Folder")
```

For example:

```python
ARCHIVE_SEARCH_PRIMARY_FOLDER = Path(r"C:\database1")
ARCHIVE_SEARCH_SECONDARY_FOLDER = Path(r"C:\database2")
```

## Index location and privacy

The index stores extracted searchable content, file metadata, and status information. Treat the SQLite index as sensitive if the source documents are sensitive.

Each selected archive folder stores its own SQLite database directly in that folder. For example:

- Primary folder `C:\Primary Folder` -> `C:\Primary Folder\.archive_search_{path code}.sqlite3`
- Secondary folder `C:\Secondary Folder` -> `C:\Secondary Folder\.archive_search_{path code}.sqlite3`

The `{path code}` is a stable 16-character hash of the canonical folder path. This keeps separate folders from accidentally sharing the same database while avoiding the need to store the full path in the filename.

Because the index contains extracted text from the source documents, make sure the archive folder permissions are appropriate for everyone who can see that folder.

### Multiple users and shared folders

Normal searches open the saved index in read-only mode. This means several people can search the same archive folder at the same time without changing the index file. If a folder has not been indexed yet, Search will not run; the app will ask the user to click **Update Index** first.

Only **Update Index** changes the saved index from the UI. It checks for new, changed, and deleted files and updates only those parts of the index. If a selected folder does not already have a saved index, **Update Index** warns the user and then builds that folder index first, which may take a while. If a file already exists at the expected index path but is not a valid Archive Search index, the app warns that the file will be overwritten with a fresh index build.

When an update starts, the program creates a small `.lock` marker file in the archive folder so only one person or computer can change that folder index at a time. If another person tries to update the same folder at the same time, they will see a message explaining that an update is already running. Other users can still search the last completed index. If the app finds an old marker left behind by a failed first-time setup, it clears it automatically and carries on.

Updates are prepared in a temporary file first. The old index remains available while the new or updated one is being prepared. The program checks the prepared index before using it, keeps a `.bak` backup of the previous index, and only swaps the prepared index into place after it has finished successfully. If anything goes wrong, the previous index is kept.

Users who only search need read permission for the archive folder and the `.archive_search_*.sqlite3` file. Users who run **Update Index** need create/read/write/delete permission in the archive folder so the program can create its `.lock`, `.tmp`, and `.bak` safety files.

The app does not automatically run index-status checks in the background. Use **Check Index Status** when you want to compare the saved index against the current folder contents before deciding whether to update.

### Why not use a disposable cache folder?

The SQLite files contain a performance index with extracted searchable text. They can be rebuilt, but rebuilding may take time, so they are treated as durable archive-folder data rather than temporary cache.

When **Search subfolders** is unticked, the program keeps previously indexed subfolder data in the SQLite database. Those entries are ignored for searching, status checks, and top-level refreshes until **Search subfolders** is turned back on.

## Safety notes

The program is designed to avoid modifying source documents:

- Source files are opened using read-only extraction paths **where supported**.
- Normal searches do not write to the index.
- Update Index writes only the program's own index and safety files.
- Legacy Word files are copied to a temporary read-only file before extraction.
- File signatures are checked before extraction. If a file extension and the actual file contents do not match, **Update Index** stops before generating the index and shows a popup naming the file, folder, full path, expected type, and detected content type so the issue can be fixed.

No Python program can promise an absolute guarantee against all third-party software behaviour, but this project deliberately avoids intentional write paths to indexed source documents.

## Notes and limitations

- PDFs are searchable only when they contain embedded text.
- `.xlsb` support requires `pyxlsb`.
- `.doc` support requires Microsoft Word and `pywin32`.
- The app is primarily intended for Windows, although much of the indexing/search logic is cross-platform.
- Indexes are keyed by the canonical folder path, not by the Primary/Secondary UI slot. If the same folder is entered as both Primary and Secondary, the app uses one shared index for that folder and searches it once to avoid duplicate results. If the two entries disagree about subfolder inclusion, the broader setting wins.
