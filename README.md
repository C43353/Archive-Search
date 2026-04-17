# Archive Search

A Windows-friendly Tkinter desktop search tool for archived documents, with two maintained program versions included in this repository.

This repository contains two related versions of the same application:

- **Archive_Search_V10.2.py** — classic text-output interface
- **Archive_Search_V10.3.py** — grouped-results interface with a details pane

**_Disclaimer: ChatGPT did a lot of heavy lifting in this code, it's improved a lot since the last code project I had. It has now exceeded my coding skill by a long way, I have checked and tried to understand as much as possible but I'm sure there will be gaps in my knowledge about these scripts._**

## What the application does

The program searches one or two configured root folders for supported files and caches extracted read-only content so repeat searches are much faster.

Supported file types:

- Excel workbooks: `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xls`
- Word documents: `.docx`, `.docm`, `.dotx`, `.dotm`, `.doc`
- PDF documents: `.pdf`

## Key features

- Search across one or two folder roots
- Optional recursive subfolder searching per root
- Search mode for **any word** or **all words**
- Quoted phrase support, for example `"Sigma 5E"`
- Read-only extraction workflow where supported
- Per-root JSON cache file for faster repeat searches
- Safer opening of matching files from the results UI
- Background worker thread so the UI stays responsive

## Version comparison

### V10.2
Best if you prefer a simple scrolling text log of matches.

Highlights:

- One large text output pane
- Each hit is rendered as a formatted result block
- Clickable result paths inside the output area
- Word/PDF results show first-match context
- Excel results are shown one matching row at a time

### V10.3
Best if you want a cleaner browsing workflow for large result sets.

Highlights:

- Grouped result list with one row per matching file
- Separate details pane for snippets
- Buttons for **Open Selected** and **Copy Path**
- Summary line shows both matching files and total matches separately
- Better for scanning many results quickly

## Requirements

- Python 3.10+ recommended
- Windows recommended for the best read-only integration with Microsoft Office

Install the core dependencies:

```bash
pip install openpyxl xlrd python-docx pypdf
```

Optional dependency for richer Office integration on Windows:

```bash
pip install pywin32
```

## Optional external software

Some features depend on software installed on the machine:

- **Microsoft Word + pywin32**: needed for legacy `.doc` extraction and read-only Word opening
- **Microsoft Excel + pywin32**: used to open Excel results directly to the matching sheet/row in read-only mode
- Without those integrations, the app falls back to the operating system default file opener where possible

## Running the program

Run either version directly:

```bash
python Archive_Search_V10.2_commented.py
```

or

```bash
python Archive_Search_V10.3_commented.py
```

## Default folders

Both scripts currently include these default search folders in code:

- `C:\Primary Folder`
- `C:\Secondary Folder`

You can change them in the UI at runtime, or edit the constants near the top of the script.

## Cache behaviour

Each indexed root folder gets its own cache file:

```text
.Archive_Search_cache.json
```

The cache stores extracted searchable content and a lightweight directory manifest to speed up later searches.

## Safety notes

The program is designed to avoid modifying source documents:

- source files are opened using read-only extraction paths where supported
- the program writes only its own cache file
- cache writes are atomic to reduce corruption risk

No Python program can promise an absolute guarantee against all third-party software behaviour, but this project deliberately avoids intentional write paths to indexed source documents.
