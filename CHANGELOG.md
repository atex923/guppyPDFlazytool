# Changelog

## V0.3.1

- Explicitly imports `tkinter.font` so Nuitka folder builds include the module needed by externally loaded `customtkinter`.
- Stops frozen exe builds from trying to run `pip install` through the exe when startup packages are missing.
- Improves frozen-mode missing-package guidance for external `site-packages`.
- Restores readable OCR missing-file guidance for `ocr_packages`, `external_packages`, and `site-packages`.

## V0.3.0

- Promotes the current packaged-ready source line from `V0.2.20` to `V0.3.0`.
- Keeps the Nuitka folder-build fixes, external package DLL search paths, and external OCR loading behavior from `V0.2.20`.
- Reorganizes versioned source history so the root folder contains only the current versioned `.pyw`.
- Moves older Git snapshots that previously used `V0.3.x` numbering into `archive/legacy_git/`.
- Consolidates the in-program history block into major/minor version highlights when the first or second version number changes.

## V0.2.20

- Adds DLL search paths for external OCR packages such as ONNXRuntime, OpenCV, and Shapely.

## V0.2.19

- Adds external package directories to the Windows DLL search path.
- Refreshes executable-side package paths before lazy PDF/image imports.

## V0.2.18

- Adds executable-side external package paths before startup dependency checks.
- Fixes Nuitka folder builds loading `customtkinter` from the bundled `site-packages`.

## V0.2.17

- Builds the `customtkinter` module name at runtime to avoid Nuitka static following.
- Keeps `customtkinter` loadable from executable-side external package folders.

## V0.2.16

- Changed `customtkinter` startup import to dynamic loading.
- Allows Nuitka folder builds to load `customtkinter` from executable-side `site-packages`.

## V0.2.15

- Prioritized external package folders in `sys.path` for Nuitka folder builds.
- Fixed loading packages from the executable-side `site-packages` folder.

## V0.2.14

- Reformatted the main Python source with Ruff.
- Simplified redundant `try`/`except`/`pass` blocks with `contextlib.suppress`.
- Cleaned unused event argument names and the retained app instance reference.
- Removed a redundant return-time temporary variable and added an explicit fallback return.

## V0.2.13

- Removed dynamic responsive/compact layout adjustment from the main window.
- Removed dynamic font, row-height, padding, and preview-toolbar repositioning changes.
- Kept the fixed 12pt font and fixed control height rules from V0.2.12.

## V0.2.12

- Fixed all tab UI fonts at 12pt.
- Fixed vertical tab button height to six font units.
- Fixed input height and tightened vertical spacing.
- Reworked move tab labels to readable Chinese.
- Reworked watermark tab top controls into two compact rows.
- Removed watermark previous/next page buttons and operation tip text.
- Set watermark message list to 6pt font with a fixed lower message area.

## V0.2.11

- Added compact layout mode for low-resolution windows.
- Compact mode reduces font size, tree row height, button height, entry height, and panel padding.
- Keeps first/second tab bottom input areas visible by letting the middle tables shrink further.
- Changed first-tab preview toolbar buttons to stable ASCII labels: `-`, `+`, `Fit`, `<`, `>`.

## V0.2.10

- Reworked window sizing for lower-resolution displays.
- Reduced main column minimum widths and made preview toolbar wrap OCR controls on narrow windows.
- Moved first-tab folder action buttons to a second row so paths and buttons do not squeeze each other.
- Rewrote OCR engine loading to use external lazy imports cleanly.
- Prefer RapidOCR first to avoid PaddleOCR model host/network checks during box selection.
- OCR missing-package text now tells users to confirm OCR files are beside the app.

## V0.2.9

- Improved missing dependency messages for PyMuPDF / `fitz`.
- Shows the active Python executable in dependency install instructions.
- Installed core GUI/PDF packages into the Python 3.14 environment used by `.pyw` double-click launch.

## V0.2.8

- Changed OCR engines to external lazy imports for Nuitka builds.
- Added external OCR package search folders beside the app.
- Shows a clear message asking users to confirm OCR files are in the same folder when OCR cannot load.
- Confirmed PyMuPDF uses package name `PyMuPDF` and import name `fitz`.

## V0.2.7

- Fixed the first tab lower-left fields being pushed out by the file browser.
- Changed the left side to a fixed upper/lower grid layout with an elastic center table.

## V0.2.6

- Fixed maximize and resize layout behavior.
- Updated main column weights, preview area, toolbars, and tables to follow window resizing.

## V0.2.5

- Improved startup speed with lazy loading for PDF, Pillow, numpy, and OCR dependencies.
- Lets the main window open before feature-specific modules load.

## V0.2.4

- Fixed customtkinter font warning output for `.pyw` / executable startup.
- Avoids `font_shapes` warnings during launch.

## V0.2.3

- Changed third and fourth tab action buttons to rounded `CTkButton` styling.
- Kept button styling aligned with the first tab.
- Generated both `.py` and `.pyw` deliverables.

## V0.2.2

- Enlarged fourth-tab subtool tab labels.

## V0.2.1

- Removed the fourth-tab header slogan.
- Added undo-last-action buttons to all four fourth-tab subtools.
- Changed fourth-tab subtool labels to four peach-orange color levels.
- Unified feature button styling.

## V0.2.0

- Renamed the fourth main tab to `旋壓合切`.
- Unified the four main tab color system.
- Added program history in source comments.

## V0.1.3

- Integrated `PDF旋轉吧 V1.7.1` as the fourth main tab.

## V0.1.2

- Fixed drag-and-drop support in the watermark tab.
- Unified base colors across the first three tabs.

## V0.1.1

- Started `Guppy PDF手搓工具`.
- Integrated PDF rename/move and watermark workflows.
