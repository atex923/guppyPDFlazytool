# Guppy PDF手搓工具

`Guppy_PDFlazyTool` 是一套繁體中文桌面 PDF 手作工具，整合 PDF 更名、搬移、浮水印註記、頁面旋轉壓縮合併裁切，以及 Office/圖片文件轉 PDF 工作流。

Current version: `V1.4.0`

## 程式特色

- PDF 更名：支援欄位組合命名、讀取檔名、OCR 輔助填入收文資訊。
- 批次搬移：左側 PDF 清單支援 Ctrl/Shift 多選，重複檔會跳過，其餘檔案先搬移，並可整批回復上一批動作。
- 浮水印註記：提供 PDF 預覽、文字註記、簽名/字體處理與輸出後縮圖預覽。
- 旋壓合切：整合頁面編排、壓縮、合併、頁面刪取、插入空白頁、單頁右下角右轉、抽取/刪除與縮圖快取；頁面合併可切換清單/縮圖、依檔名排列與上下移動。
- 轉換氣體：支援 PowerPoint、Word、TXT、ODT、TIF/TIFF 批次拖曳轉 PDF，並提供左側批次統計與右側 PDF 預覽。
- PPT 轉 4 格：可將簡報轉為一般 PDF 後再排成 2x2 四格版面。
- 視窗操作：支援固定最上層，Windows 最大化會避開工作列可用區域。
- 系統匣按鈕：右側分頁欄最上方固定顯示 `.` 按鈕；Windows 可隱藏到系統匣，一般最小化維持縮到工具列。
- 置頂按鈕：固定最上層功能移到右側分頁欄最下方，方便和分頁操作集中。
- 封裝友善：OCR 與大型選用套件採延遲/外部載入，降低 Nuitka 編譯負擔。

## 支援格式

- PDF：更名、搬移、浮水印、頁面編排、壓縮、合併、頁面刪取。
- PowerPoint：`ppt`, `pptx`, `pptm`, `pps`, `ppsx`, `pot`, `potx`。
- Word / 文件：`doc`, `docx`, `txt`, `odt`。
- 圖片：`tif`, `tiff`，支援多頁 TIFF。

## Latest main program

Use these root-level files for the latest version:

- `Guppy_PDFlazyTool.pyw`: latest no-console desktop launcher
- `Guppy_PDFlazyTool.py`: latest Python source
- `Guppy_PDFlazyTool_V1.4.0.pyw`: versioned standalone no-console source
- `Guppy_PDFlazyTool_V1.4.0.py`: versioned Python source

Older version-numbered full-source files are kept in `archive/`. Older Git snapshots that used the `V0.3.x` number before this release are kept in `archive/legacy_git/`.

## 安裝

```bash
pip install -r requirements.txt
```

部分轉檔功能另外需要：

- Microsoft PowerPoint / Word for Windows COM conversion
- LibreOffice for ODT and Office fallback conversion

## 執行

```bash
python Guppy_PDFlazyTool.py
```

Windows 無 console 啟動可使用：

```bash
pythonw Guppy_PDFlazyTool.pyw
```

## Nuitka build

For faster exe builds, see `NUITKA_FAST_BUILD.md`. The recommended fast path keeps OCR engines external and prevents Nuitka from following large optional OCR/science packages during compilation.

## Files

- `Guppy_PDFlazyTool.py`: stable latest Python source
- `Guppy_PDFlazyTool.pyw`: stable latest no-console launcher
- `Guppy_PDFlazyTool_V1.4.0.py`: versioned Python source for V1.4.0
- `Guppy_PDFlazyTool_V1.4.0.pyw`: standalone no-console source for V1.4.0
- `requirements.txt`: Python dependencies
- `archive/`: historical full-source version snapshots grouped by version folder
- `NUITKA_FAST_BUILD.md`: notes for faster Nuitka builds
- `build_nuitka_fast.cmd`: Windows example build command

## Google Drive Release Folders

Google Drive handoff copies are organized under `12.Codex/Guppy_PDFlazyTool/versions/`. Each version keeps its own folder, for example `V1.4.0/`, with the matching `.py` and `.pyw` files. The `12.Codex/Guppy_PDFlazyTool/latest/` folder and `12.Codex` root may also keep the latest launcher and source for quick access.

See `CHANGELOG.md` for version history.
