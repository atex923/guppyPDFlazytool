# Nuitka Fast Build Notes

This project is structured so the GUI can start without importing large PDF,
image, OCR, and scientific packages at module import time.

For faster executable builds:

- Build from `Guppy_PDFlazyTool.py`, not from the full versioned `.pyw`.
- Keep OCR engines external in `ocr_packages`, `external_packages`, or `site-packages`.
- Do not let Nuitka follow OCR packages such as PaddleOCR, EasyOCR, RapidOCR,
  ONNXRuntime, OpenCV, Shapely, SciPy, Matplotlib, Torch, or TensorFlow.
- Keep `tkinterdnd2` optional. If it is not bundled, drag-and-drop is disabled
  but browse buttons still work.

Use `build_nuitka_fast.cmd` as the Windows reference command.

After building, copy required external packages beside the executable if you
want full PDF/OCR behavior from an external-package layout.
