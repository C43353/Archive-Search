"""Example local configuration for Archive Search.

Copy this file to archive_search_local_config.py on each machine and edit the
values for that local installation. The real archive_search_local_config.py file
is ignored by git so company-specific paths and branding do not have to be
committed to the shared project.
"""

from pathlib import Path

# Shown in the white header when LOCAL_COMPANY_LOGO_PATH is not set or cannot be loaded.
LOCAL_COMPANY_NAME = "Company Name"

# Optional: show a local logo image instead of the company-name text in the app header.
# Transparent PNGs are recommended. Transparent areas inherit the white header surface.
# GIF also works with Tk's built-in loader; JPG works if Pillow is installed.
# Relative paths are resolved from this project folder.
LOCAL_COMPANY_LOGO_PATH = None
# Examples:
# LOCAL_COMPANY_LOGO_PATH = Path(r"C:\Path\To\logo.png")
# LOCAL_COMPANY_LOGO_PATH = Path("branding/company_logo.png")

# Maximum displayed logo size in the app header.
LOCAL_HEADER_LOGO_MAX_WIDTH = 360
LOCAL_HEADER_LOGO_MAX_HEIGHT = 74

# Optional: title-bar/taskbar icon, plus system tray icon when pystray is installed.
LOCAL_APP_ICON_PATH = None
# Examples:
# LOCAL_APP_ICON_PATH = Path(r"C:\Path\To\icon.ico")
# LOCAL_APP_ICON_PATH = Path("branding/app_icon.ico")

LOCAL_SEARCH_PRIMARY_FOLDER = Path(r"C:\Primary Folder")
LOCAL_SEARCH_SECONDARY_FOLDER = Path(r"C:\Secondary Folder")
