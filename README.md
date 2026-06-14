# Guppy PDF手搓工具

`Guppy_PDFlazyTool` is a desktop PDF utility for Windows/macOS Python environments.

Current version: `V0.3.2`

## Latest main program

Use these root-level files for the latest version:

- `Guppy_PDFlazyTool.pyw`: latest no-console desktop launcher
- `Guppy_PDFlazyTool_V0.3.2.pyw`: versioned standalone no-console source
- `Guppy_PDFlazyTool.py`: latest Python source

The root `Guppy_PDFlazyTool.pyw` file is a lightweight no-console launcher
that imports and runs `Guppy_PDFlazyTool.py`. Keep feature edits in
`Guppy_PDFlazyTool.py`, then regenerate the versioned `.pyw` for standalone
delivery.

Older version-numbered full-source files are kept in `archive/`. Older Git
snapshots that used the `V0.3.x` number before this release are kept in
`archive/legacy_git/` so the current `V0.3.x` release line remains the canonical
version.

OCR engines are loaded externally at runtime. For Nuitka onefile builds, keep
OCR packages beside the app in `ocr_packages`, `external_packages`, or
`site-packages`.

## Included tools

- PDF rename and OCR-assisted field filling
- PDF file move workflow
- PDF watermark annotation
- PDF rotate, compress, merge, and page edit tools

## Files

- `Guppy_PDFlazyTool.py`: stable latest Python source
- `Guppy_PDFlazyTool.pyw`: stable latest no-console launcher
- `Guppy_PDFlazyTool_V0.3.2.pyw`: standalone no-console source for V0.3.2
- `archive/`: historical full-source version snapshots moved out of the root

## Suggested dependencies

```bash
pip install customtkinter PyMuPDF pillow numpy tkinterdnd2
```

Optional OCR engines:

```bash
pip install paddleocr paddlepaddle
pip install rapidocr-onnxruntime
pip install easyocr
```

## Git history

- `45c43d0` - Add Guppy PDF lazy tool v0.2.3
- `6f9d418` - Initial commit

See `CHANGELOG.md` for version history.
