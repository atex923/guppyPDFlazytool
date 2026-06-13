# -*- coding: utf-8 -*-
"""No-console launcher for Guppy_PDFlazyTool.py.

Original source version: V0.2.20

Keep the implementation in Guppy_PDFlazyTool.py so the .py and .pyw entry
points do not drift apart.
"""

ORIGINAL_VERSION = "V0.2.20"


if __name__ == "__main__":
    try:
        from Guppy_PDFlazyTool import run_app

        run_app()
    except Exception:
        import traceback
        from pathlib import Path

        error_text = traceback.format_exc()

        try:
            Path(__file__).with_name("pdfname_error_log.txt").write_text(error_text, encoding="utf-8")
        except Exception:
            pass

        try:
            import tkinter as tk
            from tkinter import messagebox

            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Guppy PDF lazy tool startup error", error_text[-2000:])
            root.destroy()
        except Exception:
            pass
