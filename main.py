"""Entry point for Archive Search."""

from __future__ import annotations

import tkinter as tk

from archive_app import ArchiveSearchApp


def main() -> None:
    root = tk.Tk()
    ArchiveSearchApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
