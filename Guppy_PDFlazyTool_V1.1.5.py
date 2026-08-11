# -*- coding: utf-8 -*-
# =========================================================
# Guppy PDF手搓工具 V1.1.5
# =========================================================
# 程式歷史摘要：
# 說明：第一碼或第二碼進版時，本區整併為該碼號的改版重點；
#       細項逐版紀錄保留在 CHANGELOG.md 與 archive/。
# V0.1.x  建立 Guppy PDF手搓工具，整合 PDF 更名、搬移、浮水印與旋轉工具。
# V0.2.x  完成四分頁工具整合、版面調整、固定字體/欄位規格、延遲載入、
#         OCR 外部載入、PyMuPDF/Pillow 載入提示與 Nuitka 資料夾版相容修正。
# V0.3.0  正式整理版本歷史與發佈結構，保留 V0.2.20 的 Nuitka 資料夾版、
#         外部 OCR 與 DLL 搜尋路徑修正，根目錄只保留最新版程式入口。
# V0.3.1  補強 Nuitka exe 相容性，改善 frozen 缺套件提示與 OCR 外掛錯誤訊息。
# V0.3.2  修正第一分頁低解析度預覽工具列遮蔽，恢復第三分頁 PDF 拖曳開啟。
# V0.3.3  加速 Nuitka 編譯：拖曳套件改動態載入，補充 no-follow 編譯設定。
# V0.3.3.1 程式碼深度優化：移除冗餘全域設定、優化 O(N) 檔案清單排序、精簡模組檢查邏輯。
# V0.3.3.2 修正 ensure_required_modules() 進版後 required 未定義問題，補強缺少 customtkinter 時的檢查流程。
# V0.3.3.3 清除殘留空白 Warning Filter 區段；修正 OCR finally 不必要的 try/except del arr；
#          ensure_packages 去重只做一次；wrap_text_to_width 共用 draw 避免重建 dummy image；
#          load_tkinterdnd 簡化重複旗標檢查；移除 format_file_size 不可達 return。
# V0.3.3.4 修正 OCR recognize() finally 的 arr 清理方式，改用 suppress(Exception) 防護，避免未來流程變動造成二次例外。
# V0.3.3.5 第四分頁新增「頁面抽取」：可拖曳 PDF 載入，勾選所需頁面並於勾選框旁填寫事業碼、P+號碼，
#          依原檔名加上 _part_YYYYMMDD 另存選取頁面。
# V0.3.3.6 第四分頁「頁面抽取」移除事業碼與 P+號碼欄位，
#          並將「抽取另存」按鈕移至上方「輸出編輯 PDF」左側。
# V0.3.3.7 第四分頁「頁面抽取」新增另存資料夾選擇；刪除勾選改成只更新單頁縮圖標示，
#          避免每次刪除勾選都重新產生全部頁面縮圖而造成短暫停滯。
# V0.3.3.8 第四分頁移除「頁面抽取」功能區；「取消抽取」移至「回復上一動作」右側；
#          按下「抽取另存」時直接選擇儲放資料夾後輸出。
# V0.4.0   第四分頁改為單一編輯模式：任一頁勾選刪除時停用全部抽取，
#          任一頁勾選抽取時停用全部刪除；抽取另存完成後不再跳出確認視窗。
# V0.4.1   新增第五分頁「PPT轉換」：整合 PPToPDF 功能，可拖曳或選取 PowerPoint 檔，
#          以 PowerPoint COM 或 LibreOffice 轉成 PDF，輸出檔名為原檔名加 _rp_YYYYMMDD。
# V0.4.2   修正 PPT 轉 PDF：改用暫存英文路徑轉檔後搬回，降低中文檔名/路徑造成的
#          PowerPoint RPC 失敗與 LibreOffice 找不到輸出 PDF 問題；補強輸出檔搜尋與錯誤訊息。
# V0.4.3   第五分頁 PPT轉換：開啟 PDF 圓鈕改為 30x30；PowerPoint/LibreOffice
#          轉檔時改採隱藏或最小化背景執行，避免轉檔程式跳到前景。
# V0.4.4   新增第六分頁「PTT轉4格」：整合 PPToPDFoT4 功能，可拖曳或選取 PowerPoint 檔，
# V0.5     版本升級：承接 V0.4.4 六分頁功能整合成果，更新主程式版號與發佈檔名。
# V0.6     修正第二分頁「浮水印註記」中間區：左側作業預覽、右側儲存後 PDF 縮圖預覽。
# V0.6.1   第四分頁「頁面合併」中間區新增 8:2 分割畫面：左側合併清單，
# V0.7     正式版號升級，沿用 V0.6.1 功能整合成果。
# V0.8     OCR 結果統一轉為臺灣繁體中文；第一分頁新增「讀取檔名」。
# V0.8.1   OCR 狀態顯示區分「OCR載入中」與「OCR辨識中」，完成後顯示使用的 OCR 引擎。
# V0.8.2   新增第六分頁「DOC轉換」：DOC、DOCX、TXT 拖曳或選取後轉成 PDF；
#          右半邊顯示轉檔後 PDF 縮圖預覽，原「PTT轉4格」順延為第七分頁。
# V0.8.3   第四分頁「頁面編輯」新增「清空」工作區；合併第五、六分頁為
#          「轉換氣體」，自動依 PPT/PPTX/DOC/DOCX/TXT 副檔名轉成 PDF 並預覽。
# V0.9     正式版號升級：承接 V0.8.3 的頁面編輯清空、轉換氣體與轉檔預覽功能。
# V0.9.1   第四分頁「頁面編輯」新增插入空白頁；插入於目前點選頁面的下一頁。
# V0.9.2   優化第四分頁：點選頁面不再重繪全部縮圖，並加入縮圖快取以改善卡頓。
# V0.9.3   第四分頁新增頁碼範圍快速抽取；輸入「##-##」後自動另存 PDF。
# V0.9.4   第五分頁「轉換氣體」新增 TIF/TIFF 轉 PDF，支援多頁 TIFF 與右側預覽。
# V1.0.0   正式版號升級：承接 V0.9.4 全部功能，作為第一個穩定正式版本。
# V1.1.0   主視窗右下角新增「固定最上層」勾選，可即時切換視窗置頂狀態。
# V1.1.1   修正視窗最大化時下緣被 Windows 工作列遮住；改用目前螢幕可用工作區。
# V1.1.2   第五分頁「轉換氣體」新增 ODT 轉 PDF；使用 LibreOffice 並支援右側預覽。
# V1.1.3   第二分頁「搬移」左側 PDF 清單支援多選，可整批一次搬移並整批回復。
# V1.1.4   修正 V1.1.3 啟動失敗：第2分頁多選提示改用既有 FONT 與 MUTED_TEXT 樣式。
# V1.1.5   第2分頁批次搬移改為先搬可搬檔案；重複或失敗檔案最後才提示，正常時不跳確認視窗。
#          右側於輸出合併 PDF 後自動顯示輸出檔縮圖預覽，方便確認頁面排版。
#          先轉一般 PDF，再以 PyMuPDF 重新排成 2x2 四格 PDF；內頁黑框、間距 2pt，輸出 _rpT4_YYYYMMDD。
#
# 建議安裝：
# pip install customtkinter PyMuPDF pillow numpy tkinterdnd2
# pip install paddleocr paddlepaddle
# 備用：
# pip install rapidocr-onnxruntime
# pip install easyocr
# PPT轉換：pip install pywin32，或安裝 LibreOffice
# DOC轉換：DOC/DOCX 建議安裝 Microsoft Word；亦可使用 LibreOffice
# ODT轉換：需安裝 LibreOffice
# TIFF轉換：使用 Pillow + PyMuPDF，支援單頁及多頁 TIF/TIFF
# PTT轉4格：另需 PyMuPDF，已包含在主要 PDF 功能需求中
# =========================================================

from __future__ import annotations

import os
import re
import gc
import io
import sys
import warnings
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings("ignore", message=".*RequestsDependencyWarning.*")
warnings.filterwarnings("ignore", message=".*Preferred drawing method.*")

import traceback
import platform
import shutil
import subprocess
import importlib
import site
import threading
import time
from contextlib import suppress
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import tkinter as tk
import tkinter.font  # noqa: F401
from tkinter import ttk, filedialog, messagebox


# =========================================================
# PYW Startup / Dependency Guard
# =========================================================
def app_base_dir() -> Path:
    """Return a stable directory for logs whether running as .py/.pyw or frozen exe."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    try:
        return Path(__file__).resolve().parent
    except Exception:
        return Path.cwd()


BASE_DIR = app_base_dir()
ERROR_LOG = BASE_DIR / "pdfname_error_log.txt"
STARTUP_LOG = BASE_DIR / "pdfname_startup_log.txt"
EXTERNAL_PACKAGE_DIR_NAMES = ("ocr_packages", "external_packages", "site-packages")
EXTERNAL_DLL_SUBDIRS = (
    ("PIL",),
    ("pymupdf",),
    ("numpy.libs",),
    ("bin",),
    ("cv2",),
    ("onnxruntime", "capi"),
    ("shapely.libs",),
)
DLL_DIRECTORY_HANDLES: list[object] = []
DLL_DIRECTORY_KEYS: set[str] = set()


def add_dll_search_path(path: Path) -> None:
    if os.name != "nt" or not hasattr(os, "add_dll_directory"):
        return
    try:
        resolved = path.resolve()
    except Exception:
        return
    if not resolved.exists():
        return
    key = str(resolved).lower()
    if key in DLL_DIRECTORY_KEYS:
        return
    try:
        DLL_DIRECTORY_HANDLES.append(os.add_dll_directory(str(resolved)))
        DLL_DIRECTORY_KEYS.add(key)
    except Exception:
        pass


def add_external_package_paths() -> list[Path]:
    """Add package folders next to the script/exe before optional imports.

    This keeps large OCR engines outside Nuitka onefile builds. Put OCR
    packages beside the app in one of these folders:
    ocr_packages, external_packages, or site-packages.
    """
    roots: list[Path] = []
    with suppress(Exception):
        roots.append(Path(sys.executable).resolve().parent)
    with suppress(Exception):
        roots.append(Path(__file__).resolve().parent)
    roots.append(BASE_DIR)

    env_paths = []
    for env_name in ("GUPPY_OCR_PATH", "GUPPY_EXTERNAL_PACKAGES"):
        env_value = os.environ.get(env_name, "")
        env_paths.extend(part for part in env_value.split(os.pathsep) if part.strip())

    candidates: list[Path] = [Path(part).expanduser() for part in env_paths]
    for root in roots:
        candidates.append(root)
        candidates.extend(root / name for name in EXTERNAL_PACKAGE_DIR_NAMES)

    added: list[Path] = []
    seen = set()
    for path in candidates:
        try:
            resolved = path.resolve()
        except Exception:
            continue
        key = str(resolved).lower()
        if key in seen or not resolved.exists():
            continue
        seen.add(key)
        resolved_text = str(resolved)
        if resolved_text not in sys.path:
            sys.path.insert(0, resolved_text)
            site.addsitedir(resolved_text)
        add_dll_search_path(resolved)
        for child_parts in EXTERNAL_DLL_SUBDIRS:
            add_dll_search_path(resolved.joinpath(*child_parts))
        added.append(resolved)
    return added


def write_log(path: Path, text: str) -> None:
    with suppress(Exception):
        path.write_text(text, encoding="utf-8")


def append_startup_log(text: str) -> None:
    try:
        with STARTUP_LOG.open("a", encoding="utf-8") as f:
            f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {text}\n")
    except Exception:
        pass


def install_safe_stdio_for_pyw() -> None:
    """Give pythonw/.pyw a writable stdout/stderr so third-party libraries do not crash.

    In Windows .pyw mode sys.stdout and sys.stderr can be None. Some packages
    such as customtkinter write startup warnings directly to sys.stderr, which
    otherwise raises AttributeError before our GUI error dialog can appear.
    """
    global _PYW_STDIO_LOG_HANDLE
    try:
        if getattr(sys, "stdout", None) is None or getattr(sys, "stderr", None) is None:
            _PYW_STDIO_LOG_HANDLE = (BASE_DIR / "pdfname_stdio_log.txt").open(
                "a", encoding="utf-8", buffering=1
            )
            if getattr(sys, "stdout", None) is None:
                sys.stdout = _PYW_STDIO_LOG_HANDLE
            if getattr(sys, "stderr", None) is None:
                sys.stderr = _PYW_STDIO_LOG_HANDLE
    except Exception:

        class _NullWriter:
            def write(self, *_args, **_kwargs):
                return 0

            def flush(self):
                return None

        if getattr(sys, "stdout", None) is None:
            sys.stdout = _NullWriter()
        if getattr(sys, "stderr", None) is None:
            sys.stderr = _NullWriter()


install_safe_stdio_for_pyw()


class _FilteredCustomTkinterWarningWriter:
    """Filter noisy customtkinter font warning without hiding real errors."""

    _blocked_patterns = (
        "customtkinter.windows.widgets.font warning:",
        "Preferred drawing method 'font_shapes' can not be used because the font file could not be loaded",
        "Using 'circle_shapes' instead. The rendering quality will be bad!",
    )

    def __init__(self, target):
        self._target = target

    def write(self, text):
        try:
            if any(pattern in str(text) for pattern in self._blocked_patterns):
                return len(str(text))
            return self._target.write(text)
        except Exception:
            return 0

    def flush(self):
        try:
            return self._target.flush()
        except Exception:
            return None

    def __getattr__(self, name):
        return getattr(self._target, name)


def suppress_customtkinter_font_warning_output() -> None:
    """customtkinter sometimes prints this warning directly to stderr, not via warnings.warn.

    Keeping this wrapper active prevents the warning from showing in console/packaged builds,
    while traceback and normal errors still pass through.
    """
    try:
        if getattr(sys, "stderr", None) is not None and not isinstance(
            sys.stderr, _FilteredCustomTkinterWarningWriter
        ):
            sys.stderr = _FilteredCustomTkinterWarningWriter(sys.stderr)
    except Exception:
        pass


suppress_customtkinter_font_warning_output()


def show_startup_error(title: str, body: str) -> None:
    """Show visible errors in pythonw/.pyw mode and always write a log."""
    write_log(ERROR_LOG, body)
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(title, body[-4000:])
        root.destroy()
    except Exception:
        # In .pyw mode there may be no console, so the log is the reliable fallback.
        pass


def run_pip_install(packages: list[str]) -> tuple[bool, str]:
    creationflags = 0
    if os.name == "nt":
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

    commands = [
        [sys.executable, "-m", "pip", "install", "--upgrade", "pip"],
        [sys.executable, "-m", "pip", "install", *packages],
    ]
    output_parts: list[str] = []
    for cmd in commands:
        try:
            append_startup_log("執行：" + " ".join(cmd))
            completed = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=creationflags,
                timeout=600,
            )
            output_parts.append(completed.stdout or "")
            if completed.returncode != 0:
                return False, "\n".join(output_parts)
        except Exception:
            output_parts.append(traceback.format_exc())
            return False, "\n".join(output_parts)
    return True, "\n".join(output_parts)


def frozen_missing_package_message(packages: list[str]) -> str:
    package_text = ", ".join(dict.fromkeys(packages))
    folders = ", ".join(EXTERNAL_PACKAGE_DIR_NAMES)
    return (
        "程式缺少啟動必要套件。\n\n"
        f"缺少套件：{package_text}\n\n"
        "此程式目前是 exe 版本，無法用 exe 自己執行 pip install。\n"
        f"請確認 {folders} 是否與 exe 放在同一資料夾內，"
        "或重新執行 Nuitka 打包流程補齊外部套件。\n\n"
        f"詳細紀錄：\n{STARTUP_LOG}"
    )


def customtkinter_module_name() -> str:
    return os.environ.get("GUPPY_CUSTOMTKINTER_MODULE") or "".join(
        ("custom", "tkinter")
    )


def ensure_required_modules() -> None:
    """Only install the light GUI dependency during startup.

    Heavy modules such as PyMuPDF, Pillow, numpy, PaddleOCR, RapidOCR and
    EasyOCR are intentionally NOT imported here.  The main window can open
    first, and those modules are imported later when the user opens a PDF,
    creates thumbnails, runs OCR, or saves watermark output.
    """
    import importlib.util

    add_external_package_paths()
    ctk_module = customtkinter_module_name()
    missing_packages = [] if importlib.util.find_spec(ctk_module) else [ctk_module]

    if not missing_packages:
        return

    append_startup_log("啟動必要套件缺少：" + ", ".join(missing_packages))
    if getattr(sys, "frozen", False):
        raise RuntimeError(frozen_missing_package_message(missing_packages))

    ok, pip_output = run_pip_install(missing_packages)
    append_startup_log(pip_output[-4000:])

    # Re-check the same startup dependency after the install attempt.
    # V0.3.3.1 removed the previous `required` dictionary but still referenced it here;
    # keep this check single-source to avoid NameError when customtkinter is missing.
    still_missing = [] if importlib.util.find_spec(ctk_module) else [ctk_module]
    if not ok or still_missing:
        install_cmd = f'"{sys.executable}" -m pip install ' + " ".join(
            still_missing or missing_packages
        )
        raise RuntimeError(
            "程式缺少啟動必要套件，且自動安裝失敗。\n\n"
            f"請手動執行：\n{install_cmd}\n\n"
            f"詳細紀錄：\n{STARTUP_LOG}\n\n" + pip_output[-2500:]
        )


try:
    ensure_required_modules()
    ctk = importlib.import_module(customtkinter_module_name())
except Exception:
    show_startup_error("程式啟動失敗", traceback.format_exc())
    raise SystemExit(1)


class LazyImport:
    """Import a heavy module only when one of its attributes is first used."""

    def __init__(self, module_name: str, pip_name: str | None = None):
        self.module_name = module_name
        self.pip_name = pip_name or module_name
        self._module = None

    def _load(self):
        if self._module is None:
            add_external_package_paths()
            append_startup_log(f"延遲載入套件：{self.module_name}")
            try:
                self._module = importlib.import_module(self.module_name)
            except Exception as exc:
                package_note = f"套件名稱：{self.pip_name}"
                if self.module_name == "fitz":
                    package_note = "套件名稱：PyMuPDF；Python 匯入名稱：fitz"
                if getattr(sys, "frozen", False):
                    raise RuntimeError(
                        f"使用此功能需要外部套件。\n"
                        f"{package_note}\n\n"
                        "請確認 site-packages 是否與 exe 放在同一資料夾內，"
                        "或重新執行 Nuitka 打包流程補齊外部套件。"
                    ) from exc
                raise RuntimeError(
                    f"使用此功能需要套件。\n"
                    f"{package_note}\n\n"
                    f"目前執行的 Python：\n{sys.executable}\n\n"
                    f'請先執行：\n"{sys.executable}" -m pip install {self.pip_name}'
                ) from exc
        return self._module

    def __getattr__(self, name):
        return getattr(self._load(), name)

    def __call__(self, *args, **kwargs):
        return self._load()(*args, **kwargs)

    def available(self) -> bool:
        try:
            self._load()
            return True
        except Exception:
            return False


# Heavy modules are lazy proxies.  They are imported only when a PDF/preview/OCR
# function actually touches them, not while the application is opening.
fitz = LazyImport("fitz", "PyMuPDF")
np = LazyImport("numpy", "numpy")
Image = LazyImport("PIL.Image", "pillow")
ImageTk = LazyImport("PIL.ImageTk", "pillow")
ImageEnhance = LazyImport("PIL.ImageEnhance", "pillow")
ImageFilter = LazyImport("PIL.ImageFilter", "pillow")
ImageDraw = LazyImport("PIL.ImageDraw", "pillow")
ImageFont = LazyImport("PIL.ImageFont", "pillow")



# =========================================================
# Theme / Constants
# =========================================================
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

APP_VERSION = "1.1.5"
APP_TITLE = f"Guppy PDF手搓工具 V{APP_VERSION}"

BG = "#EEF2F7"
CARD = "#FFFFFF"
PRIMARY = "#2563EB"
PRIMARY_HOVER = "#1D4ED8"
PRIMARY_SOFT = "#EAF3FF"
PRIMARY_SOFT_HOVER = "#DCEBFF"
TAB_INACTIVE = "#CBD5E1"
MUTED_TEXT = "#4B5563"
FOOTER_TEXT = "#374151"
RED = "#DC2626"
RED_HOVER = "#B91C1C"
YELLOW = "#F59E0B"
YELLOW_HOVER = "#D97706"
PREVIEW_BLUE = PRIMARY_SOFT
PREVIEW_BORDER = "#CBD5E1"
FILENAME_BG = "#F8FAFC"
PREFIX_BG = PRIMARY_SOFT
OCR_BG = "#F8FAFC"
TURN_TAB_COLORS = ["#FFE5E0", "#FFD3C4", "#FFBEA6", "#FFA886"]
TURN_TAB_ACTIVE = "#FF8F66"
TURN_TAB_HOVER = "#FFB79B"
TEXT = "#111827"

UI_FONT_SIZE = 12
INPUT_HEIGHT = 28
TAB_BUTTON_HEIGHT = UI_FONT_SIZE * 6

FONT = ("Microsoft JhengHei UI", UI_FONT_SIZE)
TREE_FONT = ("Microsoft JhengHei UI", UI_FONT_SIZE)
BTN_FONT = ("Microsoft JhengHei UI", UI_FONT_SIZE)
LARGE_BTN_FONT = ("Microsoft JhengHei UI", UI_FONT_SIZE)
TURN_TAB_FONT = ("Microsoft JhengHei UI", UI_FONT_SIZE, "bold")
TITLE_FONT = ("Microsoft JhengHei UI", UI_FONT_SIZE, "bold")


def get_signature_font():
    system = platform.system()
    if system == "Darwin":
        return ("Baskerville", 13, "italic")
    if system == "Windows":
        return ("Old English Text MT", 13)
    return ("serif", 13, "italic")


SIGN_FONT = get_signature_font()


def rounded_button(parent, text, command, width=None, accent=False):
    """Create a shared rounded action button for embedded ttk-heavy pages."""
    width = width or max(72, int(len(text) * 18 + 30))
    return ctk.CTkButton(
        parent,
        text=text,
        command=command,
        width=width,
        height=30,
        corner_radius=10,
        fg_color=PRIMARY if accent else PREVIEW_BLUE,
        hover_color=PRIMARY_HOVER if accent else PRIMARY_SOFT_HOVER,
        text_color="white" if accent else "black",
        font=LARGE_BTN_FONT,
        border_width=0 if accent else 1,
        border_color=PREVIEW_BORDER,
    )


IMAGE_OFFSET = 20
MIN_ZOOM = 0.2
MAX_ZOOM = 5.0
MAX_PREVIEW_PIXELS = 8_000_000
PREVIEW_GC_INTERVAL = 8
DATE_FORMAT = "%Y-%m-%d %H:%M"
OCR_PLACEHOLDER = "開啟「框選辨識」後，在PDF預覽區拖曳框選文字，OCR結果會顯示在這裡。"


@dataclass
class PDFState:
    folder: str = ""
    selected_pdf: str = ""
    current_pdf_path: str = ""
    current_page: int = 0
    zoom: float = 1.0
    sort_column: str = "filename"
    sort_reverse: bool = False


# =========================================================
# Helper
# =========================================================
def normalize_receive_date(text: str) -> str:
    """收文日期：民國3碼年份不補0，月日補2碼。"""
    if not text:
        return ""

    s = str(text).strip()
    s = re.sub(r"[年月／－\-.]", "/", s)
    s = s.replace("中華民國", "").replace("民國", "").replace("日", "")
    s = re.sub(r"\s+", "", s)

    match = re.search(r"(\d{2,4})/(\d{1,2})/(\d{1,2})", s)
    if match:
        year, month, day = match.groups()
        return f"{year}{month.zfill(2)}{day.zfill(2)}"

    digits = re.sub(r"\D", "", s)

    if len(digits) in (7, 8):
        return digits

    match7 = re.search(r"\d{7}", digits)
    if match7:
        return match7.group(0)

    match8 = re.search(r"\d{8}", digits)
    if match8:
        return match8.group(0)

    return digits


def clean_one_line(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def format_timestamp(timestamp: float) -> str:
    try:
        return datetime.fromtimestamp(timestamp).strftime(DATE_FORMAT)
    except Exception:
        return ""


def toggle_sort(current_column: str, reverse: bool, column: str):
    return (column, not reverse) if current_column == column else (column, False)


def safe_pdf_filename(filename: str) -> str:
    filename = re.sub(r'[\\/:*?"<>|]', "", filename)
    return re.sub(r"\s+", " ", filename).strip()


def list_pdf_files(folder: str, sort_column: str, reverse: bool):
    result = []
    folder_path = Path(folder)

    if not folder_path.exists():
        return result

    for path in folder_path.iterdir():
        if path.is_file() and path.suffix.lower() == ".pdf":
            try:
                stat = path.stat()
                added = get_file_added_time(path, stat)
                result.append((path.name, added))
            except OSError:
                continue

    key = (
        (lambda item: item[0].lower())
        if sort_column == "filename"
        else (lambda item: item[1])
    )
    return sorted(result, key=key, reverse=reverse)


def format_file_size(size: int) -> str:
    try:
        size = int(size)
    except Exception:
        return ""
    units = ["B", "KB", "MB", "GB", "TB"]
    value = float(size)
    for unit in units:
        if value < 1024 or unit == units[-1]:
            return f"{value:.0f} {unit}" if unit == "B" else f"{value:.1f} {unit}"
        value /= 1024


def get_file_added_time(path: Path, stat_result=None) -> float:
    """回傳檔案加入/建立時間。
    macOS 優先使用 st_birthtime；Windows 的 st_ctime 是建立時間；
    Linux 若無 birth time，退回 st_ctime。
    """
    try:
        stat_result = stat_result or path.stat()
        return getattr(stat_result, "st_birthtime", stat_result.st_ctime)
    except Exception:
        return 0.0


def list_directory_items(folder: str, sort_column: str, reverse: bool):
    result = []
    folder_path = Path(folder)

    if not folder_path.exists():
        return result

    for path in folder_path.iterdir():
        try:
            stat = path.stat()
            is_dir = path.is_dir()
            result.append(
                {
                    "name": path.name,
                    "path": str(path),
                    "is_dir": is_dir,
                    "size": 0 if is_dir else stat.st_size,
                    "created": get_file_added_time(path, stat),
                    "modified": stat.st_mtime,
                }
            )
        except OSError:
            continue

    def item_key(item):
        if sort_column == "size":
            return (item["size"], item["name"].lower())
        if sort_column == "created":
            return (item["created"], item["name"].lower())
        if sort_column == "modified":
            return (item["modified"], item["name"].lower())
        return item["name"].lower()

    folders, files = [], []
    for item in result:
        if item["is_dir"]:
            folders.append(item)
        else:
            files.append(item)
            
    folders.sort(key=item_key, reverse=reverse)
    files.sort(key=item_key, reverse=reverse)
    return folders + files


# =========================================================
# OCR Engine
# =========================================================
class OCREngine:
    """
    延遲載入 OCR：
    程式啟動時不先載入 PaddleOCR，第一次框選辨識才載入。
    這樣可大幅加快開啟程式速度。
    """

    def __init__(self):
        self.engine_name = "尚未載入"
        self.ready = False
        self.paddleocr = None
        self.rapidocr = None
        self.easyocr_reader = None
        self.opencc_converter = None
        self.opencc_checked = False
        self.load_error = ""

    @staticmethod
    def _external_import(module_name: str):
        add_external_package_paths()
        return importlib.import_module(module_name)

    def load(self):
        if self.ready:
            return

        errors = []

        # OCR engines are optional external packages. Nuitka onefile builds
        # should leave them beside the exe instead of bundling them inside.
        try:
            RapidOCR = getattr(
                self._external_import("rapidocr_onnxruntime"), "RapidOCR"
            )
            self.rapidocr = RapidOCR()
            self.engine_name = "RapidOCR"
            self.ready = True
            return
        except Exception as exc:
            errors.append(f"RapidOCR import/init: {exc}")

        try:
            PaddleOCR = getattr(self._external_import("paddleocr"), "PaddleOCR")
            for kwargs in (
                {"use_textline_orientation": True, "lang": "chinese_cht"},
                {"lang": "chinese_cht"},
                {"use_textline_orientation": True, "lang": "ch"},
                {"lang": "ch"},
            ):
                try:
                    self.paddleocr = PaddleOCR(**kwargs)
                    self.engine_name = "PaddleOCR"
                    self.ready = True
                    return
                except Exception as exc:
                    errors.append(f"PaddleOCR init: {exc}")
        except Exception as exc:
            errors.append(f"PaddleOCR import: {exc}")

        try:
            easyocr = self._external_import("easyocr")
            self.easyocr_reader = easyocr.Reader(["ch_tra", "en"], gpu=False)
            self.engine_name = "EasyOCR"
            self.ready = True
            return
        except Exception as exc:
            errors.append(f"EasyOCR import/init: {exc}")

        self.engine_name = "未載入OCR"
        self.load_error = "\n".join(errors[-6:])
        self.ready = True

    def _load_opencc(self):
        """延遲載入 OpenCC，OCR 第一次使用時才建立繁體轉換器。"""
        if self.opencc_checked:
            return self.opencc_converter

        self.opencc_checked = True
        try:
            OpenCC = getattr(self._external_import("opencc"), "OpenCC")
            self.opencc_converter = OpenCC("s2twp")
            return self.opencc_converter
        except Exception as first_error:
            # 原始碼模式可嘗試自動安裝；封裝版則維持外部套件載入規則。
            if not getattr(sys, "frozen", False):
                try:
                    ok, _output = run_pip_install(["opencc-python-reimplemented"])
                    if ok:
                        importlib.invalidate_caches()
                        OpenCC = getattr(self._external_import("opencc"), "OpenCC")
                        self.opencc_converter = OpenCC("s2twp")
                        return self.opencc_converter
                except Exception:
                    pass
            self.load_error = (self.load_error + f"\nOpenCC：{first_error}").strip()
            return None

    @staticmethod
    def normalize_ocr_text(text: str) -> str:
        if not text:
            return ""
        text = str(text).replace("\u3000", " ").replace("\t", " ").replace("\r", "")
        text = re.sub(r"[ ]{2,}", " ", text)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip()

    def to_traditional_chinese(self, text: str) -> str:
        """將 OCR 結果統一轉為臺灣繁體中文。"""
        text = self.normalize_ocr_text(text)
        converter = self._load_opencc()
        if converter is not None:
            try:
                text = converter.convert(text)
            except Exception:
                pass

        # OCR 常見異體與誤辨字補正；OpenCC 未載入時也可做基本修正。
        replacements = {
            "鉄路": "鐵路",
            "臺铁": "臺鐵",
            "台铁": "臺鐵",
            "监造": "監造",
            "铁路": "鐵路",
            "桥": "橋",
        }
        for old, new in replacements.items():
            text = text.replace(old, new)
        return text

    @staticmethod
    def preprocess(img: Image.Image) -> Image.Image:
        if img.mode != "RGB":
            img = img.convert("RGB")

        w, h = img.size

        if max(w, h) < 1200:
            img = img.resize(
                (w * 2, h * 2), getattr(Image, "Resampling", Image).LANCZOS
            )

        img = ImageEnhance.Contrast(img).enhance(1.55)
        img = ImageEnhance.Sharpness(img).enhance(1.8)
        return img.filter(ImageFilter.SHARPEN)

    def recognize(self, img: Image.Image) -> str:
        if img is None:
            return ""

        self.load()

        if self.engine_name not in ("PaddleOCR", "RapidOCR", "EasyOCR"):
            return (
                "無法載入 OCR 套件。\n"
                "請確認 OCR 檔案是否在程式同一資料夾內。\n\n"
                "請將 OCR 套件放在 ocr_packages、external_packages "
                "或 site-packages 資料夾。\n"
                "可使用 rapidocr-onnxruntime、paddleocr+paddlepaddle 或 easyocr。\n\n"
                + self.load_error[-1200:]
            )

        processed_img = None
        arr = None
        try:
            processed_img = self.preprocess(img)
            arr = np.asarray(processed_img)

            if self.engine_name == "PaddleOCR":
                try:
                    result = self.paddleocr.ocr(arr)
                except TypeError:
                    result = self.paddleocr.ocr(arr, cls=True)
                return self.to_traditional_chinese(
                    self._parse_paddle_result(result)
                )

            if self.engine_name == "RapidOCR":
                result, _ = self.rapidocr(arr)
                text = "\n".join(
                    str(item[1]) for item in result or [] if len(item) >= 2
                ).strip()
                return self.to_traditional_chinese(text)

            if self.engine_name == "EasyOCR":
                result = self.easyocr_reader.readtext(arr, detail=0, paragraph=True)
                text = "\n".join(map(str, result)).strip()
                return self.to_traditional_chinese(text)

        except Exception as exc:
            return f"OCR失敗：{exc}"

        finally:
            with suppress(Exception):
                del arr
            if processed_img is not None and processed_img is not img:
                with suppress(Exception):
                    processed_img.close()
            gc.collect()

        return ""

    @staticmethod
    def _parse_paddle_result(result) -> str:
        texts = []

        if not result:
            return ""

        for page in result:
            if not page:
                continue

            for line in page:
                try:
                    if len(line) < 2:
                        continue
                    info = line[1]
                    if isinstance(info, (list, tuple)) and info:
                        texts.append(str(info[0]))
                    else:
                        texts.append(str(info))
                except Exception:
                    continue

        return "\n".join(texts).strip()


# =========================================================
# Watermark Tab (merged from PDF浮水印註記工具 V1.0.8)
# =========================================================
import tempfile
from io import BytesIO

# Drag-and-drop is optional and loaded dynamically.  Keeping tkinterdnd2 out of
# top-level imports prevents Nuitka from statically following it during builds.
DND_FILES = None
TkinterDnD = None
HAS_DND = False
_DND_LOAD_ATTEMPTED = False


def load_tkinterdnd() -> bool:
    global DND_FILES, TkinterDnD, HAS_DND, _DND_LOAD_ATTEMPTED
    if HAS_DND:
        return True
    if _DND_LOAD_ATTEMPTED:
        return False

    _DND_LOAD_ATTEMPTED = True
    try:
        module = importlib.import_module("tkinterdnd2")
        DND_FILES = getattr(module, "DND_FILES")
        TkinterDnD = getattr(module, "TkinterDnD")
        HAS_DND = True
        return True
    except Exception:
        DND_FILES = None
        TkinterDnD = None
        HAS_DND = False
        append_startup_log(
            "未載入 tkinterdnd2，拖曳功能停用；可手動安裝：pip install tkinterdnd2"
        )
        return False

WATERMARK_APP_VERSION = "V1.0.8"
WATERMARK_APP_TITLE = f"PDF浮水印註記工具 {WATERMARK_APP_VERSION}"
WM_DEFAULT_FONT_SIZE = 12
WM_MIN_FONT_SIZE = 6
WM_MAX_FONT_SIZE = 96
WM_DEFAULT_RENDER_SCALE = 1.5
WM_MIN_RENDER_SCALE = 0.5
WM_MAX_RENDER_SCALE = 4.0
WM_RENDER_SCALE_STEP = 0.25
WM_OUTPUT_IMAGE_SCALE = 3
WM_DEFAULT_BOX_MAX_WIDTH = 420
WM_DEFAULT_BOX_MIN_WIDTH = 180
WM_DEFAULT_BOX_VISIBLE_WIDTH_RATIO = 0.55
WM_DEFAULT_BOX_HEIGHT = 80
WM_RECENTER_DELAYS_MS = (50, 150, 350)
WM_MIN_USABLE_VIEW_SIZE = 80
WM_PAGE_MARGIN = 20
WM_WATERMARK_HEX_COLOR = "#d00000"
WM_WATERMARK_RGBA_COLOR = (208, 0, 0, 255)
WM_DEFAULT_NOTE_TEMPLATE = "附件「##」紙本資料1份置於施工室工程查核之用。"
WM_SETTINGS_FILE_NAME = "PDF浮水印註記工具設定.txt"


def ensure_packages(require_numpy: bool = False):
    """Load PDF/preview packages only when the active feature needs them."""
    missing = []

    for proxy, package in ((fitz, "PyMuPDF"), (Image, "pillow"), (ImageTk, "pillow")):
        try:
            proxy._load()
        except Exception:
            if package not in missing:
                missing.append(package)

    # Watermark output uses drawing/font modules; loading them here is still
    # delayed until the PDF/watermark page is actually used.
    for proxy, package in ((ImageDraw, "pillow"), (ImageFont, "pillow")):
        try:
            proxy._load()
        except Exception:
            if package not in missing:
                missing.append(package)

    if require_numpy:
        try:
            np._load()
        except Exception:
            missing.append("numpy")

    if missing:
        unique = list(dict.fromkeys(missing))
        pkgs = " ".join(unique)
        messagebox.showerror(
            "缺少套件",
            "使用此功能缺少必要套件：\n\n"
            + "\n".join(unique)
            + "\n\n請執行：\n"
            + f'"{sys.executable}" -m pip install {pkgs}',
        )
        return False
    return True


def parse_dropped_files(data, tk_root=None):
    if tk_root is not None:
        try:
            return list(tk_root.tk.splitlist(data))
        except Exception:
            pass

    files = []
    current = ""
    in_brace = False

    for ch in data:
        if ch == "{":
            in_brace = True
            current = ""
        elif ch == "}":
            in_brace = False
            if current:
                files.append(current)
                current = ""
        elif ch == " " and not in_brace:
            if current:
                files.append(current)
                current = ""
        else:
            current += ch

    if current:
        files.append(current)

    return files


def find_cjk_font_file():
    """
    尋找常見中文字型檔。
    Pillow 會直接把文字畫成透明圖片，所以中文字型會影響輸出品質。
    """
    candidates = []

    if sys.platform.startswith("win"):
        windir = os.environ.get("WINDIR", r"C:\Windows")
        fonts_dir = Path(windir) / "Fonts"
        candidates.extend(
            [
                fonts_dir / "msjh.ttc",  # 微軟正黑體
                fonts_dir / "msjhbd.ttc",  # 微軟正黑體粗體
                fonts_dir / "mingliu.ttc",  # 細明體
                fonts_dir / "kaiu.ttf",  # 標楷體
                fonts_dir / "simsun.ttc",
                fonts_dir / "simhei.ttf",
            ]
        )

    elif sys.platform == "darwin":
        candidates.extend(
            [
                Path("/System/Library/Fonts/PingFang.ttc"),
                Path("/System/Library/Fonts/STHeiti Light.ttc"),
                Path("/Library/Fonts/Arial Unicode.ttf"),
            ]
        )

    else:
        candidates.extend(
            [
                Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"),
                Path("/usr/share/fonts/opentype/noto/NotoSansCJKtc-Regular.otf"),
                Path("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"),
                Path("/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"),
                Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
            ]
        )

    for path in candidates:
        if path.exists():
            return str(path)

    return None


def load_font(font_size, font_file=None):
    if font_file and os.path.exists(font_file):
        try:
            return ImageFont.truetype(font_file, font_size)
        except Exception:
            pass

    # Windows 常見 fallback
    if sys.platform.startswith("win"):
        for name in ["msjh.ttc", "mingliu.ttc", "kaiu.ttf"]:
            try:
                return ImageFont.truetype(name, font_size)
            except Exception:
                continue

    try:
        return ImageFont.truetype("arial.ttf", font_size)
    except Exception:
        return ImageFont.load_default()


def measure_text(draw, text, font):
    if not text:
        return 0, 0
    bbox = draw.textbbox((0, 0), text, font=font)
    return bbox[2] - bbox[0], bbox[3] - bbox[1]


def wrap_text_to_width(text, font, max_width, draw=None):
    """
    依照像素寬度自動換行。
    支援中文沒有空白的情況，會逐字測量。
    draw 可由呼叫端傳入共用，省去重複建立 dummy image。
    """
    if draw is None:
        dummy_img = Image.new("RGBA", (10, 10), (255, 255, 255, 0))
        draw = ImageDraw.Draw(dummy_img)

    result_lines = []

    for paragraph in text.splitlines():
        if not paragraph:
            result_lines.append("")
            continue

        current = ""
        for ch in paragraph:
            test = current + ch
            w, _ = measure_text(draw, test, font)
            if w <= max_width or not current:
                current = test
            else:
                result_lines.append(current)
                current = ch

        if current:
            result_lines.append(current)

    return result_lines


def layout_text_lines(text, font, max_width, max_height=None, line_gap=2):
    dummy_img = Image.new("RGBA", (10, 10), (255, 255, 255, 0))
    draw = ImageDraw.Draw(dummy_img)
    lines = wrap_text_to_width(text, font, max_width, draw=draw)

    if max_height is None:
        return lines

    visible_lines = []
    used_height = 0
    for line in lines:
        _, line_h = measure_text(draw, line if line else "口", font)
        line_h = max(line_h, int(getattr(font, "size", WM_DEFAULT_FONT_SIZE) * 1.1))

        next_height = line_h if not visible_lines else used_height + line_gap + line_h
        if next_height > max_height:
            break

        visible_lines.append(line)
        used_height = next_height

    return visible_lines


def make_watermark_png_bytes(
    text,
    width_px,
    height_px,
    font_size_px,
    font_file=None,
    color=WM_WATERMARK_RGBA_COLOR,
):
    """
    將文字窗格轉成透明背景 PNG。
    回傳 PNG bytes。
    """
    width_px = max(20, int(width_px))
    height_px = max(20, int(height_px))
    font_size_px = max(WM_MIN_FONT_SIZE, int(font_size_px))

    img = Image.new("RGBA", (width_px, height_px), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    font = load_font(font_size_px, font_file)

    padding_x = max(2, int(font_size_px * 0.25))
    padding_y = max(2, int(font_size_px * 0.20))
    line_gap = max(2, int(font_size_px * 0.25))

    max_text_width = max(10, width_px - padding_x * 2)
    max_text_height = max(1, height_px - padding_y * 2)
    lines = layout_text_lines(text, font, max_text_width, max_text_height, line_gap)

    y = padding_y
    for line in lines:
        if y > height_px:
            break

        # 估算行高
        _, line_h = measure_text(draw, line if line else "口", font)
        line_h = max(line_h, int(font_size_px * 1.1))

        if y + line_h > height_px:
            break

        draw.text((padding_x, y), line, font=font, fill=color)
        y += line_h + line_gap

    buffer = BytesIO()
    img.save(buffer, format="PNG")
    return buffer.getvalue()


class SmoothWatermarkBox:
    HANDLE_SIZE = 10
    MIN_W = 90
    MIN_H = 45

    def __init__(self, app, x, y, width, height, text, font_size=WM_DEFAULT_FONT_SIZE):
        self.app = app
        self.canvas = app.canvas
        self.x = int(x)
        self.y = int(y)
        self.width = int(width)
        self.height = int(height)
        self.text = text
        self.font_size = int(font_size)
        self.active_mode = None
        self.start_mouse = None
        self.start_geom = None

        self.box_id = self.canvas.create_rectangle(
            self.x,
            self.y,
            self.x + self.width,
            self.y + self.height,
            outline=WM_WATERMARK_HEX_COLOR,
            width=2,
            fill="#ffffff",
            stipple="gray12",
            tags=("watermark_box",),
        )
        self.text_id = self.canvas.create_text(
            self.x + 6,
            self.y + 6,
            text=self.preview_text(),
            anchor="nw",
            fill=WM_WATERMARK_HEX_COLOR,
            font=("Microsoft JhengHei", self.font_size, "bold"),
            tags=("watermark_box",),
        )
        self.handles = {}
        self._create_handles()

        self._bind_drag_events()
        self.fit_inside_page()
        self.update_handles()

    def _create_handles(self):
        handle_defs = {
            "n": "sb_v_double_arrow",
            "s": "sb_v_double_arrow",
            "e": "sb_h_double_arrow",
            "w": "sb_h_double_arrow",
            "nw": "size_nw_se",
            "se": "size_nw_se",
            "ne": "size_ne_sw",
            "sw": "size_ne_sw",
        }

        for name, cursor in handle_defs.items():
            handle_id = self.canvas.create_rectangle(
                0,
                0,
                self.HANDLE_SIZE,
                self.HANDLE_SIZE,
                fill=WM_WATERMARK_HEX_COLOR,
                outline=WM_WATERMARK_HEX_COLOR,
                tags=("watermark_handle",),
            )
            self.canvas.tag_bind(
                handle_id, "<ButtonPress-1>", lambda e, n=name: self.start_resize(e, n)
            )
            self.canvas.tag_bind(handle_id, "<B1-Motion>", self.resize)
            self.canvas.tag_bind(handle_id, "<ButtonRelease-1>", self.stop_action)
            self.handles[name] = handle_id

    def _bind_drag_events(self):
        for item_id in (self.box_id, self.text_id):
            self.canvas.tag_bind(item_id, "<ButtonPress-1>", self.start_drag)
            self.canvas.tag_bind(item_id, "<B1-Motion>", self.drag)
            self.canvas.tag_bind(item_id, "<ButtonRelease-1>", self.stop_action)

    def get_text(self):
        return self.text.strip()

    def set_text(self, text):
        self.text = text
        self.refresh_text_layout()

    def set_font_size(self, size):
        self.font_size = int(size)
        self.canvas.itemconfigure(
            self.text_id, font=("Microsoft JhengHei", self.font_size, "bold")
        )
        self.refresh_text_layout()

    def preview_padding(self):
        padding_x = max(2, int(self.font_size * 0.25))
        padding_y = max(2, int(self.font_size * 0.20))
        return padding_x, padding_y

    def preview_text(self):
        font = load_font(self.font_size, self.app.cjk_font_file)
        padding_x, padding_y = self.preview_padding()
        line_gap = max(2, int(self.font_size * 0.25))
        max_text_width = max(10, self.width - padding_x * 2)
        max_text_height = max(1, self.height - padding_y * 2)
        lines = layout_text_lines(
            self.text, font, max_text_width, max_text_height, line_gap
        )
        return "\n".join(lines)

    def refresh_text_layout(self):
        self.canvas.itemconfigure(self.text_id, text=self.preview_text())

    def page_size(self):
        return max(1, int(self.app.page_image_width)), max(
            1, int(self.app.page_image_height)
        )

    def fit_inside_page(self):
        page_w, page_h = self.page_size()
        min_w = min(self.MIN_W, page_w)
        min_h = min(self.MIN_H, page_h)
        self.width = int(max(min_w, min(self.width, page_w)))
        self.height = int(max(min_h, min(self.height, page_h)))
        self.x, self.y = self.clamp_position(self.x, self.y)
        self._apply_geometry()

    def clamp_position(self, x, y, width=None, height=None):
        page_w, page_h = self.page_size()
        width = self.width if width is None else width
        height = self.height if height is None else height
        max_x = max(0, page_w - width)
        max_y = max(0, page_h - height)
        return int(min(max(0, x), max_x)), int(min(max(0, y), max_y))

    def visible_area(self):
        return self.app.get_visible_area()

    def center_on_visible_area(self):
        self.fit_inside_page()
        view_x, view_y, view_w, view_h = self.visible_area()
        target_x = view_x + (view_w - self.width) / 2
        target_y = view_y + (view_h - self.height) / 2
        self.x, self.y = self.clamp_position(target_x, target_y)
        self._apply_geometry()

    def center_on_page(self):
        self.center_on_visible_area()

    def start_drag(self, event):
        self.active_mode = "move"
        self.start_mouse = self._event_canvas_xy(event)
        self.start_geom = (self.x, self.y, self.width, self.height)
        return "break"

    def drag(self, event):
        if self.active_mode != "move" or not self.start_mouse:
            return "break"

        cx, cy = self._event_canvas_xy(event)
        sx, sy = self.start_mouse
        ox, oy, ow, oh = self.start_geom

        nx = ox + (cx - sx)
        ny = oy + (cy - sy)

        self.x, self.y = self.clamp_position(nx, ny)

        self._apply_geometry()
        return "break"

    def start_resize(self, event, mode):
        self.active_mode = mode
        self.start_mouse = self._event_canvas_xy(event)
        self.start_geom = (self.x, self.y, self.width, self.height)
        return "break"

    def resize(self, event):
        if not self.active_mode or not self.start_mouse:
            return "break"

        cx, cy = self._event_canvas_xy(event)
        sx, sy = self.start_mouse
        ox, oy, ow, oh = self.start_geom

        dx = cx - sx
        dy = cy - sy

        nx, ny, nw, nh = ox, oy, ow, oh
        mode = self.active_mode

        if "e" in mode:
            nw = ow + dx
        if "s" in mode:
            nh = oh + dy
        if "w" in mode:
            nx = ox + dx
            nw = ow - dx
        if "n" in mode:
            ny = oy + dy
            nh = oh - dy

        page_w, page_h = self.page_size()
        min_w = min(self.MIN_W, page_w)
        min_h = min(self.MIN_H, page_h)

        if nw < min_w:
            if "w" in mode:
                nx = ox + ow - min_w
            nw = min_w

        if nh < min_h:
            if "n" in mode:
                ny = oy + oh - min_h
            nh = min_h

        if nx < 0:
            if "w" in mode:
                nw += nx
            nx = 0

        if ny < 0:
            if "n" in mode:
                nh += ny
            ny = 0

        if nx + nw > page_w:
            nw = page_w - nx

        if ny + nh > page_h:
            nh = page_h - ny

        self.x = int(nx)
        self.y = int(ny)
        self.width = int(max(min_w, nw))
        self.height = int(max(min_h, nh))
        self.x, self.y = self.clamp_position(self.x, self.y)

        self._apply_geometry()
        return "break"

    def stop_action(self, _event=None):
        self.active_mode = None
        self.start_mouse = None
        self.start_geom = None
        return "break"

    def _event_canvas_xy(self, event):
        root_x = self.canvas.winfo_rootx()
        root_y = self.canvas.winfo_rooty()
        canvas_x = self.canvas.canvasx(event.x_root - root_x)
        canvas_y = self.canvas.canvasy(event.y_root - root_y)
        return canvas_x, canvas_y

    def _apply_geometry(self):
        padding_x, padding_y = self.preview_padding()
        self.canvas.coords(
            self.box_id, self.x, self.y, self.x + self.width, self.y + self.height
        )
        self.canvas.coords(self.text_id, self.x + padding_x, self.y + padding_y)
        self.refresh_text_layout()
        self.update_handles()
        self.canvas.configure(
            scrollregion=(0, 0, self.app.page_image_width, self.app.page_image_height)
        )

    def update_handles(self):
        hs = self.HANDLE_SIZE
        half = hs // 2

        positions = {
            "nw": (self.x - half, self.y - half),
            "n": (self.x + self.width / 2 - half, self.y - half),
            "ne": (self.x + self.width - half, self.y - half),
            "e": (self.x + self.width - half, self.y + self.height / 2 - half),
            "se": (self.x + self.width - half, self.y + self.height - half),
            "s": (self.x + self.width / 2 - half, self.y + self.height - half),
            "sw": (self.x - half, self.y + self.height - half),
            "w": (self.x - half, self.y + self.height / 2 - half),
        }

        self.canvas.tag_raise(self.box_id)
        self.canvas.tag_raise(self.text_id)
        for name, handle_id in self.handles.items():
            px, py = positions[name]
            self.canvas.coords(handle_id, px, py, px + hs, py + hs)
            self.canvas.tag_raise(handle_id)

    def destroy(self):
        try:
            self.canvas.delete(self.box_id)
            self.canvas.delete(self.text_id)
            for handle_id in self.handles.values():
                self.canvas.delete(handle_id)
        except Exception:
            pass

    def to_pdf_rect(self):
        scale = self.app.render_scale
        return fitz.Rect(
            self.x / scale,
            self.y / scale,
            (self.x + self.width) / scale,
            (self.y + self.height) / scale,
        )


class PDFWatermarkApp:
    def __init__(self, root, embedded=False):
        self.root = root
        self.embedded = embedded
        if not embedded:
            self.root.title(WATERMARK_APP_TITLE)
            self.root.geometry("1180x860")
            self.root.minsize(900, 650)
            self.root.resizable(True, True)

        self.pdf_path = None
        self.doc = None
        self.current_page_index = 0
        self.render_scale = WM_DEFAULT_RENDER_SCALE

        self.page_tk_image = None
        self.page_image_width = 1
        self.page_image_height = 1
        self.output_preview_images = []

        self.watermark_box = None
        self.last_output_path = None

        self.cjk_font_file = find_cjk_font_file()
        self.default_note_template = self.load_default_note_template()
        self.dnd_enabled = False

        self.build_ui()
        self.setup_drag_drop()
        self.root.after(250, self.setup_drag_drop)

        self.add_message(f"程式已啟動。預設字體大小 {WM_DEFAULT_FONT_SIZE}。")
        if self.dnd_enabled:
            self.add_message("拖曳開啟功能已啟用。")
        else:
            self.add_message(
                "拖曳開啟功能未啟用，若需要請安裝：pip install tkinterdnd2"
            )
        if self.cjk_font_file:
            self.add_message(f"偵測到中文字型：{self.cjk_font_file}")
        else:
            self.add_message(
                "未偵測到系統中文字型檔，會使用 Pillow 預設字型，中文可能無法正常顯示。"
            )
        self.add_message(f"已載入預設字串：{self.default_note_template}")

    def settings_path(self):
        try:
            base_dir = Path(__file__).resolve().parent
        except Exception:
            base_dir = Path.cwd()
        return base_dir / WM_SETTINGS_FILE_NAME

    def load_default_note_template(self):
        path = self.settings_path()
        try:
            if path.exists():
                text = path.read_text(encoding="utf-8").strip()
                if text:
                    return text
        except Exception:
            pass
        return WM_DEFAULT_NOTE_TEMPLATE

    def save_default_note_template(self):
        template = self.default_template_var.get().strip()
        if not template:
            messagebox.showwarning("預設字串為空", "請先輸入預設字串。")
            return

        try:
            self.settings_path().write_text(template, encoding="utf-8")
            self.default_note_template = template
            self.add_message(f"已儲存預設字串：{template}")
        except Exception as e:
            messagebox.showerror("儲存失敗", f"無法儲存預設字串：\n{e}")
            self.add_message(f"預設字串儲存失敗：{e}")

    def build_ui(self):
        self.main_frame = ttk.Frame(self.root, padding=6, style="App.TFrame")
        self.main_frame.pack(fill="both", expand=True)

        self.main_frame.rowconfigure(1, weight=1)
        self.main_frame.columnconfigure(0, weight=1)

        top = ttk.LabelFrame(
            self.main_frame, text="PDF浮水印註記", padding=6, style="App.TLabelframe"
        )
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)
        top.columnconfigure(3, weight=1)

        row1 = ttk.Frame(top, style="Card.TFrame")
        row1.grid(row=0, column=0, columnspan=8, sticky="ew", pady=1)
        row1.columnconfigure(1, weight=1)
        row1.columnconfigure(3, weight=1)

        ttk.Label(row1, text="輸入註記文字", style="Card.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 4), pady=1
        )
        self.note_var = tk.StringVar()
        self.note_entry = ttk.Entry(row1, textvariable=self.note_var, font=FONT)
        self.note_entry.grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=1)

        ttk.Label(row1, text="預設字串", style="Card.TLabel").grid(
            row=0, column=2, sticky="w", padx=(0, 4), pady=1
        )
        self.default_template_var = tk.StringVar(value=self.default_note_template)
        self.default_template_entry = ttk.Entry(
            row1, textvariable=self.default_template_var, font=FONT
        )
        self.default_template_entry.grid(
            row=0, column=3, sticky="ew", padx=(0, 8), pady=1
        )

        rounded_button(
            row1, "作為預設", self.save_default_note_template, width=88, accent=True
        ).grid(row=0, column=4, padx=2, pady=1)
        rounded_button(
            row1, "開啟PDF", self.open_pdf_dialog, width=88, accent=True
        ).grid(row=0, column=5, padx=2, pady=1)

        row2 = ttk.Frame(top, style="Card.TFrame")
        row2.grid(row=1, column=0, columnspan=8, sticky="ew", pady=1)

        rounded_button(
            row2, "確認插入", self.confirm_insert, width=88, accent=True
        ).pack(side="left", padx=2)
        ttk.Label(row2, text="PDF縮放", style="Card.TLabel").pack(
            side="left", padx=(10, 3)
        )
        rounded_button(row2, "-", self.zoom_out, width=38).pack(side="left", padx=1)
        rounded_button(row2, "+", self.zoom_in, width=38).pack(side="left", padx=1)
        rounded_button(row2, "適頁", self.fit_page_width, width=52).pack(
            side="left", padx=1
        )

        ttk.Label(row2, text="字體大小", style="Card.TLabel").pack(
            side="left", padx=(10, 3)
        )
        self.font_size_var = tk.IntVar(value=WM_DEFAULT_FONT_SIZE)
        self.font_size_spin = ttk.Spinbox(
            row2,
            from_=WM_MIN_FONT_SIZE,
            to=WM_MAX_FONT_SIZE,
            increment=1,
            width=5,
            textvariable=self.font_size_var,
            command=self.apply_font_size,
            font=FONT,
        )
        self.font_size_spin.pack(side="left", padx=2)
        self.font_size_spin.bind("<Return>", lambda _event: self.apply_font_size())
        self.font_size_spin.bind("<FocusOut>", lambda _event: self.apply_font_size())
        rounded_button(row2, "-", lambda: self.change_font_size(-1), width=38).pack(
            side="left", padx=1
        )
        rounded_button(row2, "+", lambda: self.change_font_size(1), width=38).pack(
            side="left", padx=1
        )

        rounded_button(row2, "儲存", self.save_pdf, width=70, accent=True).pack(
            side="left", padx=(10, 2)
        )
        rounded_button(row2, "列印", self.print_pdf, width=70).pack(side="left", padx=2)

        self.page_label = ttk.Label(row2, text="尚未開啟 PDF", style="Card.TLabel")
        self.page_label.pack(side="left", padx=(10, 4))
        self.status_label = ttk.Label(
            row2, text="請開啟 PDF 檔案", style="Muted.TLabel"
        )
        self.status_label.pack(side="left", padx=4)

        self.paned = ttk.PanedWindow(self.main_frame, orient="vertical")
        self.paned.grid(row=1, column=0, sticky="nsew", pady=(6, 0))

        # V0.6：浮水印分頁中間區固定切成左右兩欄，比例 8:2。
        # 左側保留原本可拖曳/縮放/放置浮水印的工作預覽；右側顯示儲存後 PDF 縮圖。
        preview_area = ttk.Frame(self.paned, style="App.TFrame")
        preview_area.rowconfigure(0, weight=1)
        preview_area.columnconfigure(0, weight=8, uniform="watermark_preview_split")
        preview_area.columnconfigure(1, weight=2, uniform="watermark_preview_split")

        preview = ttk.LabelFrame(
            preview_area, text="PDF預覽", padding=4, style="App.TLabelframe"
        )
        preview.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        output_preview = ttk.LabelFrame(
            preview_area, text="儲存PDF預覽", padding=4, style="App.TLabelframe"
        )
        output_preview.grid(row=0, column=1, sticky="nsew")

        message_frame = ttk.LabelFrame(
            self.paned, text="訊息列表", padding=3, style="App.TLabelframe"
        )
        message_frame.configure(height=72)
        message_frame.grid_propagate(False)

        self.paned.add(preview_area, weight=6)
        self.paned.add(message_frame, weight=0)

        preview.rowconfigure(0, weight=1)
        preview.columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(preview, bg=PREVIEW_BLUE, highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")

        vbar = ttk.Scrollbar(preview, orient="vertical", command=self.canvas.yview)
        vbar.grid(row=0, column=1, sticky="ns")

        hbar = ttk.Scrollbar(preview, orient="horizontal", command=self.canvas.xview)
        hbar.grid(row=1, column=0, sticky="ew")

        self.canvas.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)

        output_preview.rowconfigure(0, weight=1)
        output_preview.columnconfigure(0, weight=1)
        self.output_preview_canvas = tk.Canvas(
            output_preview, bg=FILENAME_BG, highlightthickness=0
        )
        self.output_preview_canvas.grid(row=0, column=0, sticky="nsew")
        output_vbar = ttk.Scrollbar(
            output_preview, orient="vertical", command=self.output_preview_canvas.yview
        )
        output_vbar.grid(row=0, column=1, sticky="ns")
        self.output_preview_canvas.configure(yscrollcommand=output_vbar.set)
        self.output_preview_inner = ttk.Frame(self.output_preview_canvas, style="Card.TFrame")
        self.output_preview_window = self.output_preview_canvas.create_window(
            (0, 0), window=self.output_preview_inner, anchor="nw"
        )
        self.output_preview_inner.bind(
            "<Configure>",
            lambda _event: self.output_preview_canvas.configure(
                scrollregion=self.output_preview_canvas.bbox("all")
            ),
        )
        self.output_preview_canvas.bind(
            "<Configure>",
            lambda event: self.output_preview_canvas.itemconfigure(
                self.output_preview_window, width=event.width
            ),
        )
        self.clear_output_preview("儲存後會在這裡顯示輸出 PDF 預覽。")

        message_frame.rowconfigure(0, weight=1)
        message_frame.columnconfigure(0, weight=1)

        self.message_list = tk.Listbox(
            message_frame,
            height=12,
            bg=CARD,
            fg=TEXT,
            font=("Microsoft JhengHei UI", 6),
            highlightthickness=1,
            highlightbackground=PREVIEW_BORDER,
            relief="flat",
        )
        self.message_list.grid(row=0, column=0, sticky="nsew")

        msg_vbar = ttk.Scrollbar(
            message_frame, orient="vertical", command=self.message_list.yview
        )
        msg_vbar.grid(row=0, column=1, sticky="ns")
        self.message_list.configure(yscrollcommand=msg_vbar.set)

    def clear_output_preview(self, message="尚未產生輸出 PDF 預覽。"):
        """Clear the right-side saved-PDF preview panel in the watermark tab."""
        if not hasattr(self, "output_preview_inner"):
            return
        for child in self.output_preview_inner.winfo_children():
            child.destroy()
        self.output_preview_images = []
        ttk.Label(
            self.output_preview_inner,
            text=message,
            style="Muted.TLabel",
            wraplength=180,
            justify="center",
        ).pack(fill="x", padx=8, pady=12)
        with suppress(Exception):
            self.output_preview_canvas.configure(
                scrollregion=self.output_preview_canvas.bbox("all")
            )

    def load_output_pdf_preview(self, pdf_path):
        """Render thumbnails of the saved PDF on the right side after watermark output."""
        if not pdf_path or not os.path.exists(pdf_path):
            self.clear_output_preview("找不到輸出的 PDF 檔案。")
            return
        if not ensure_packages():
            return

        for child in self.output_preview_inner.winfo_children():
            child.destroy()
        self.output_preview_images = []

        doc = None
        try:
            doc = fitz.open(pdf_path)
            with suppress(Exception):
                self.output_preview_canvas.update_idletasks()
            canvas_width = int(self.output_preview_canvas.winfo_width() or 180)
            thumb_width = max(120, min(240, canvas_width - 28))
            title = os.path.basename(pdf_path)
            ttk.Label(
                self.output_preview_inner,
                text=title,
                style="Card.TLabel",
                wraplength=max(120, thumb_width),
                justify="center",
            ).pack(fill="x", padx=6, pady=(6, 8))

            for index, page in enumerate(doc):
                zoom = max(thumb_width / max(page.rect.width, 1), 0.05)
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
                image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                photo = ImageTk.PhotoImage(image)
                self.output_preview_images.append(photo)

                item = ttk.Frame(self.output_preview_inner, style="Card.TFrame")
                item.pack(fill="x", padx=6, pady=(0, 10))
                tk.Label(item, image=photo, bg=CARD, bd=1, relief="solid").pack()
                ttk.Label(
                    item,
                    text=f"第 {index + 1} 頁",
                    style="Muted.TLabel",
                    anchor="center",
                ).pack(fill="x", pady=(2, 0))

            self.output_preview_canvas.configure(
                scrollregion=self.output_preview_canvas.bbox("all")
            )
            self.add_message(f"已在右側顯示輸出 PDF 預覽：{os.path.basename(pdf_path)}")
        except Exception as e:
            self.clear_output_preview("輸出 PDF 預覽失敗。")
            self.add_message(f"輸出 PDF 預覽失敗：{e}")
        finally:
            if doc is not None:
                doc.close()

    def add_message(self, text):
        self.message_list.insert("end", text)
        self.message_list.see("end")
        self.status_label.config(text=text)

    def get_visible_area(self):
        with suppress(Exception):
            self.root.update_idletasks()

        page_w = max(1, int(self.page_image_width))
        page_h = max(1, int(self.page_image_height))
        view_w = int(self.canvas.winfo_width())
        view_h = int(self.canvas.winfo_height())

        if view_w < WM_MIN_USABLE_VIEW_SIZE:
            view_w = int(
                self.main_frame.winfo_width() or self.root.winfo_width() or page_w
            )
        if view_h < WM_MIN_USABLE_VIEW_SIZE:
            try:
                top_h = int(self.main_frame.grid_bbox(0, 0)[3])
                root_h = int(self.root.winfo_height() or page_h)
                view_h = max(WM_MIN_USABLE_VIEW_SIZE, root_h - top_h - 80)
            except Exception:
                view_h = page_h

        view_w = max(1, min(view_w, page_w))
        view_h = max(1, min(view_h, page_h))
        view_x = self.canvas.canvasx(0)
        view_y = self.canvas.canvasy(0)
        view_x = min(max(0, view_x), max(0, page_w - view_w))
        view_y = min(max(0, view_y), max(0, page_h - view_h))
        return view_x, view_y, view_w, view_h

    def recenter_watermark_later(self):
        if not self.watermark_box:
            return

        self.watermark_box.center_on_visible_area()
        view_x, view_y, view_w, view_h = self.get_visible_area()
        self.add_message(
            f"窗格位置：x={self.watermark_box.x}, y={self.watermark_box.y}, "
            f"瀏覽區：{int(view_w)}x{int(view_h)}"
        )

        def recenter_once():
            if self.watermark_box:
                self.watermark_box.center_on_visible_area()

        for delay in WM_RECENTER_DELAYS_MS:
            self.root.after(delay, recenter_once)

    def setup_drag_drop(self):
        if not load_tkinterdnd():
            return

        widgets = [
            self.root,
            self.root.winfo_toplevel(),
            self.main_frame,
            self.canvas,
        ]
        widgets.extend(self._drop_widgets(self.main_frame))
        for widget in dict.fromkeys(widgets):
            try:
                if not hasattr(widget, "drop_target_register") or not hasattr(
                    widget, "dnd_bind"
                ):
                    continue
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self.on_drop_file)
                self.dnd_enabled = True
            except Exception:
                pass

    def _drop_widgets(self, widget):
        widgets = [widget]
        try:
            children = widget.winfo_children()
        except Exception:
            children = []
        for child in children:
            widgets.extend(self._drop_widgets(child))
        return widgets

    def on_drop_file(self, event):
        files = parse_dropped_files(event.data, self.root)
        if not files:
            return

        path = files[0]
        if not path.lower().endswith(".pdf"):
            messagebox.showwarning("格式不支援", "請拖曳 PDF 檔案。")
            return

        self.open_pdf_path(path)

    def open_pdf_dialog(self):
        if not ensure_packages():
            return

        path = filedialog.askopenfilename(
            title="選擇 PDF 檔案",
            filetypes=[("PDF 檔案", "*.pdf"), ("所有檔案", "*.*")],
        )
        if not path:
            return

        self.open_pdf_path(path)

    def open_pdf_path(self, path):
        if not ensure_packages():
            return

        try:
            if self.doc:
                self.doc.close()

            self.pdf_path = path
            self.doc = fitz.open(path)
            self.current_page_index = 0
            self.last_output_path = None
            self.clear_watermark()
            self.clear_output_preview("儲存後會在這裡顯示輸出 PDF 預覽。")
            self.render_page()
            self.add_message(f"已開啟：{os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("開啟失敗", f"無法開啟 PDF：\n{e}")
            self.add_message(f"開啟失敗：{e}")

    def render_page(self):
        if not self.doc:
            return

        self.canvas.delete("all")
        self.watermark_box = None

        page = self.doc[self.current_page_index]
        matrix = fitz.Matrix(self.render_scale, self.render_scale)
        pix = page.get_pixmap(matrix=matrix, alpha=False)

        image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.page_tk_image = ImageTk.PhotoImage(image)

        self.page_image_width = pix.width
        self.page_image_height = pix.height

        self.canvas.create_image(
            0, 0, anchor="nw", image=self.page_tk_image, tags=("pdf_page",)
        )
        self.canvas.configure(scrollregion=(0, 0, pix.width, pix.height))

        self.page_label.config(
            text=f"第 {self.current_page_index + 1} / {len(self.doc)} 頁"
        )

    def clear_watermark(self):
        if self.watermark_box:
            self.watermark_box.destroy()
            self.watermark_box = None

    def get_font_size(self):
        try:
            size = int(self.font_size_var.get())
        except Exception:
            size = WM_DEFAULT_FONT_SIZE

        size = max(WM_MIN_FONT_SIZE, min(WM_MAX_FONT_SIZE, size))
        self.font_size_var.set(size)
        return size

    def apply_font_size(self):
        size = self.get_font_size()
        if self.watermark_box:
            self.watermark_box.set_font_size(size)
            self.add_message(f"字體大小已調整為 {size}")

    def change_font_size(self, delta):
        size = self.get_font_size() + delta
        size = max(WM_MIN_FONT_SIZE, min(WM_MAX_FONT_SIZE, size))
        self.font_size_var.set(size)
        self.apply_font_size()

    def build_full_note_text(self):
        raw_text = self.note_var.get().strip()
        if not raw_text:
            return ""
        template = self.default_template_var.get().strip()
        if not template:
            template = WM_DEFAULT_NOTE_TEMPLATE
            self.default_template_var.set(template)

        if "##" in template:
            return template.replace("##", raw_text)

        return template

    def zoom_in(self):
        if not self.doc:
            return
        old_scale = self.render_scale
        self.render_scale = min(
            WM_MAX_RENDER_SCALE, self.render_scale + WM_RENDER_SCALE_STEP
        )
        self.refresh_page_after_zoom(old_scale)

    def zoom_out(self):
        if not self.doc:
            return
        old_scale = self.render_scale
        self.render_scale = max(
            WM_MIN_RENDER_SCALE, self.render_scale - WM_RENDER_SCALE_STEP
        )
        self.refresh_page_after_zoom(old_scale)

    def fit_page_width(self):
        if not self.doc:
            return

        try:
            page = self.doc[self.current_page_index]
            available_width = max(300, self.canvas.winfo_width() - 30)
            page_width = page.rect.width
            old_scale = self.render_scale
            self.render_scale = max(
                WM_MIN_RENDER_SCALE,
                min(WM_MAX_RENDER_SCALE, available_width / page_width),
            )
            self.refresh_page_after_zoom(old_scale)
        except Exception as e:
            self.add_message(f"適頁失敗：{e}")

    def refresh_page_after_zoom(self, old_scale=None):
        if not self.doc:
            return

        # 縮放 PDF 預覽時，保留目前浮水印框相對 PDF 的位置與大小。
        old_box = None
        if self.watermark_box and old_scale:
            old_box = {
                "x_pdf": self.watermark_box.x / old_scale,
                "y_pdf": self.watermark_box.y / old_scale,
                "w_pdf": self.watermark_box.width / old_scale,
                "h_pdf": self.watermark_box.height / old_scale,
                "text": self.watermark_box.get_text(),
                "font_size": self.get_font_size(),
            }

        self.render_page()

        if old_box:
            self.watermark_box = SmoothWatermarkBox(
                self,
                old_box["x_pdf"] * self.render_scale,
                old_box["y_pdf"] * self.render_scale,
                old_box["w_pdf"] * self.render_scale,
                old_box["h_pdf"] * self.render_scale,
                old_box["text"],
                font_size=old_box["font_size"],
            )

        self.add_message(f"PDF 預覽縮放：{self.render_scale:.2f}x")

    def confirm_insert(self):
        if not self.doc:
            messagebox.showwarning("尚未開啟 PDF", "請先開啟 PDF 檔案。")
            return

        text = self.build_full_note_text()
        if not text:
            messagebox.showwarning(
                "尚未輸入文字", "請先在「輸入註記文字」欄框輸入內容。"
            )
            return

        font_size = self.get_font_size()

        if self.watermark_box:
            self.watermark_box.set_text(text)
            self.watermark_box.set_font_size(font_size)
        else:
            page_w = max(1, int(self.page_image_width))
            page_h = max(1, int(self.page_image_height))
            view_x, view_y, view_w, view_h = self.get_visible_area()
            margin = (
                WM_PAGE_MARGIN
                if page_w > WM_PAGE_MARGIN * 2 and page_h > WM_PAGE_MARGIN * 2
                else 0
            )
            box_w = min(
                max(
                    WM_DEFAULT_BOX_MIN_WIDTH,
                    int(view_w * WM_DEFAULT_BOX_VISIBLE_WIDTH_RATIO),
                ),
                WM_DEFAULT_BOX_MAX_WIDTH,
                max(1, page_w - margin * 2),
            )
            box_h = min(
                max(WM_DEFAULT_BOX_HEIGHT, font_size * 5), max(1, page_h - margin * 2)
            )
            box_x = int(
                min(max(0, view_x + (view_w - box_w) / 2), max(0, page_w - box_w))
            )
            box_y = int(
                min(max(0, view_y + (view_h - box_h) / 2), max(0, page_h - box_h))
            )
            self.watermark_box = SmoothWatermarkBox(
                self, box_x, box_y, box_w, box_h, text, font_size=font_size
            )

        self.recenter_watermark_later()
        self.add_message(f"已插入註記字串：{text}")

    def save_pdf(self):
        if not self.doc or not self.pdf_path:
            messagebox.showwarning("尚未開啟 PDF", "請先開啟 PDF 檔案。")
            return None

        if not self.watermark_box:
            messagebox.showwarning("尚未插入註記", "請先按「確認插入」。")
            return None

        save_path = filedialog.asksaveasfilename(
            title="儲存 PDF 檔案",
            defaultextension=".pdf",
            initialfile=self.default_save_name(),
            filetypes=[("PDF 檔案", "*.pdf"), ("所有檔案", "*.*")],
        )
        if not save_path:
            self.add_message("已取消儲存。")
            return None

        result_path = self.write_watermark_to_pdf(save_path)
        if result_path:
            self.load_output_pdf_preview(result_path)
        return result_path

    def write_watermark_to_pdf(self, save_path):
        text = self.watermark_box.get_text() if self.watermark_box else ""

        if not text:
            messagebox.showwarning("文字為空", "註記文字框內沒有文字。")
            return None

        if Path(save_path).resolve() == Path(self.pdf_path).resolve():
            messagebox.showwarning(
                "儲存路徑不建議", "請另存為新 PDF，避免覆蓋目前開啟中的原始檔。"
            )
            self.add_message("已取消儲存：請選擇不同於原始 PDF 的儲存路徑。")
            return None

        out_doc = None
        try:
            out_doc = fitz.open(self.pdf_path)
            page = out_doc[self.current_page_index]
            rect = self.watermark_box.to_pdf_rect()

            # 高解析度合成：
            # 預覽用的是螢幕像素，若直接用同解析度 PNG 插入 PDF，放大列印時會模糊。
            # 這裡將文字圖提高倍率後再插入同一個 PDF rect，PDF 內的文字圖會更清楚。
            font_size_preview = self.get_font_size()
            output_scale = WM_OUTPUT_IMAGE_SCALE
            font_size_png = max(WM_MIN_FONT_SIZE, int(font_size_preview * output_scale))
            png_bytes = make_watermark_png_bytes(
                text=text,
                width_px=int(self.watermark_box.width * output_scale),
                height_px=int(self.watermark_box.height * output_scale),
                font_size_px=font_size_png,
                font_file=self.cjk_font_file,
                color=WM_WATERMARK_RGBA_COLOR,
            )

            page.insert_image(
                rect, stream=png_bytes, overlay=True, keep_proportion=False
            )

            out_doc.save(save_path, garbage=4, deflate=True)

            if os.path.exists(save_path):
                self.last_output_path = save_path
                self.add_message(f"已儲存並合成浮水印：{save_path}")
                return save_path

            self.add_message("儲存異常：程式執行完成，但沒有找到輸出檔案。")
            return None

        except Exception as e:
            messagebox.showerror("儲存失敗", f"無法儲存 PDF：\n{e}")
            self.add_message(f"儲存失敗：{e}")
            return None
        finally:
            if out_doc is not None:
                out_doc.close()

    def print_pdf(self):
        if not self.doc or not self.pdf_path:
            messagebox.showwarning("尚未開啟 PDF", "請先開啟 PDF 檔案。")
            return

        try:
            if self.watermark_box:
                temp_dir = tempfile.gettempdir()
                base = os.path.splitext(os.path.basename(self.pdf_path))[0]
                print_path = os.path.join(temp_dir, f"{base}_列印暫存_加入浮水印.pdf")
                result_path = self.write_watermark_to_pdf(print_path)
                if not result_path:
                    return
            else:
                print_path = self.pdf_path

            if sys.platform.startswith("win"):
                os.startfile(print_path, "print")
                self.add_message(f"已送出列印：{print_path}")
            elif sys.platform == "darwin":
                subprocess.Popen(["open", print_path])
                self.add_message("已開啟 PDF，請從預覽程式列印。")
            else:
                subprocess.Popen(["xdg-open", print_path])
                self.add_message("已開啟 PDF，請從 PDF 檢視器列印。")

        except Exception as e:
            messagebox.showerror("列印失敗", f"無法列印 PDF：\n{e}")
            self.add_message(f"列印失敗：{e}")

    def default_save_name(self):
        base = os.path.splitext(os.path.basename(self.pdf_path))[0]
        return f"{base}_加入浮水印.pdf"

    def prev_page(self):
        if not self.doc or self.current_page_index <= 0:
            return
        self.current_page_index -= 1
        self.render_page()
        self.add_message("已切換上一頁，請重新確認插入註記。")

    def next_page(self):
        if not self.doc or self.current_page_index >= len(self.doc) - 1:
            return
        self.current_page_index += 1
        self.render_page()
        self.add_message("已切換下一頁，請重新確認插入註記。")

    def on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


# =========================================================
# PDF Turn Tab (merged from PDF旋轉吧 V1.7.1)
# =========================================================
def file_size_text(path):
    if not path or not os.path.exists(path):
        return "-"
    size = os.path.getsize(path)
    units = ["B", "KB", "MB", "GB"]
    value = float(size)
    for unit in units:
        if value < 1024 or unit == units[-1]:
            return f"{value:.1f} {unit}"
        value /= 1024
    return "-"


def parse_drop_files(data):
    files = []
    current = []
    in_brace = False
    for char in data:
        if char == "{":
            in_brace = True
            current = []
        elif char == "}":
            in_brace = False
            files.append("".join(current))
            current = []
        elif char == " " and not in_brace:
            if current:
                files.append("".join(current))
                current = []
        else:
            current.append(char)
    if current:
        files.append("".join(current))
    return [path for path in files if path.lower().endswith(".pdf")]


def default_save_path(source_path, suffix):
    path = Path(source_path)
    return str(path.with_name(f"{path.stem}{suffix}.pdf"))


def make_thumbnail(pdf_path, page_index=0, width=150, rotation=0, mark=None):
    doc = fitz.open(pdf_path)
    try:
        page = doc[page_index]
        zoom = max(width / page.rect.width, 0.1)
        matrix = fitz.Matrix(zoom, zoom).prerotate(rotation)
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        image.thumbnail((width, int(width * 1.45)), Image.Resampling.LANCZOS)
        if mark:
            image = image.convert("RGBA")
            overlay = Image.new("RGBA", image.size, mark)
            image = Image.alpha_composite(image, overlay)
        return ImageTk.PhotoImage(image)
    finally:
        doc.close()


class ScrollArea(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, style="App.TFrame")
        self.canvas = tk.Canvas(self, bg=BG, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(
            self, orient="vertical", command=self.canvas.yview
        )
        self.content = ttk.Frame(self.canvas, style="App.TFrame")
        self.window_id = self.canvas.create_window(
            (0, 0), window=self.content, anchor="nw"
        )
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.content.bind("<Configure>", self._sync_scroll_region)
        self.canvas.bind("<Configure>", self._sync_width)
        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)
        self.content.bind("<Enter>", self._bind_mousewheel)
        self.content.bind("<Leave>", self._unbind_mousewheel)

    def _sync_scroll_region(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _sync_width(self, event):
        self.canvas.itemconfigure(self.window_id, width=event.width)

    def _bind_mousewheel(self, _event=None):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)

    def _unbind_mousewheel(self, _event=None):
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")

    def _on_mousewheel(self, event):
        if event.state & 0x0004:
            return
        if getattr(event, "num", None) == 4:
            delta = -1
        elif getattr(event, "num", None) == 5:
            delta = 1
        else:
            delta = -1 if event.delta > 0 else 1
        self.canvas.yview_scroll(delta, "units")


class BaseTab(ttk.Frame):
    def __init__(self, app, notebook):
        super().__init__(notebook, padding=12, style="App.TFrame")
        self.app = app
        self.thumb_size = tk.IntVar(value=150)
        self.thumbs = []
        self.drag_from = None
        self.drag_cards = []
        self.resize_after_id = None
        self.undo_stack = []

    def push_undo(self):
        snapshot = self.get_undo_snapshot()
        if snapshot is not None:
            self.undo_stack.append(snapshot)
            self.undo_stack = self.undo_stack[-30:]

    def undo_last_action(self):
        if not self.undo_stack:
            messagebox.showinfo("回復上一動作", "目前沒有可回復的動作。")
            return
        self.restore_undo_snapshot(self.undo_stack.pop())

    def get_undo_snapshot(self):
        return None

    def restore_undo_snapshot(self, _snapshot):
        return None

    def enable_drop(self, widget, callback):
        if not load_tkinterdnd() or DND_FILES is None:
            return
        try:
            if not hasattr(widget, "drop_target_register") or not hasattr(
                widget, "dnd_bind"
            ):
                return
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind(
                "<<Drop>>", lambda event: callback(parse_drop_files(event.data), event)
            )
        except Exception:
            return

    def clear_frame(self, frame):
        for child in frame.winfo_children():
            child.destroy()
        self.thumbs.clear()
        self.drag_cards = []

    def enable_responsive_layout(self, area):
        area.canvas.bind("<Configure>", self.schedule_layout, add="+")

    def schedule_layout(self, _event=None):
        if self.resize_after_id:
            self.after_cancel(self.resize_after_id)
        self.resize_after_id = self.after(80, self.relayout_cards)

    def columns_for_width(self, card_width):
        area_width = (
            self.area.canvas.winfo_width()
            if hasattr(self, "area")
            else self.winfo_width()
        )
        return max(1, area_width // card_width)

    def relayout_cards(self):
        self.resize_after_id = None
        if not self.drag_cards:
            return
        columns = self.get_columns()
        for pos, card in enumerate(self.drag_cards):
            card.grid_configure(row=pos // columns, column=pos % columns)

    def choose_pdf(self):
        initialdir = self.app.last_dir if self.app.last_dir else os.getcwd()
        path = filedialog.askopenfilename(
            title="選取 PDF",
            initialdir=initialdir,
            filetypes=[("PDF files", "*.pdf")],
        )
        if path:
            self.app.last_dir = os.path.dirname(path)
        return path

    def choose_pdfs(self):
        initialdir = self.app.last_dir if self.app.last_dir else os.getcwd()
        paths = filedialog.askopenfilenames(
            title="選取 PDF",
            initialdir=initialdir,
            filetypes=[("PDF files", "*.pdf")],
        )
        if paths:
            self.app.last_dir = os.path.dirname(paths[0])
        return list(paths)

    def bind_drag_sort(self, widget, pos):
        widget.bind("<ButtonPress-1>", lambda event, p=pos: self.start_drag(event, p))
        widget.bind("<ButtonRelease-1>", self.end_drag)

    def start_drag(self, _event, pos):
        self.drag_from = pos

    def find_drop_position(self, event):
        if not self.drag_cards:
            return None
        x_root = event.x_root
        y_root = event.y_root
        nearest = None
        nearest_distance = None
        for pos, card in enumerate(self.drag_cards):
            left = card.winfo_rootx()
            top = card.winfo_rooty()
            right = left + card.winfo_width()
            bottom = top + card.winfo_height()
            if left <= x_root <= right and top <= y_root <= bottom:
                return pos
            center_x = left + card.winfo_width() / 2
            center_y = top + card.winfo_height() / 2
            distance = (center_x - x_root) ** 2 + (center_y - y_root) ** 2
            if nearest_distance is None or distance < nearest_distance:
                nearest = pos
                nearest_distance = distance
        return nearest

    def move_item(self, items, from_pos, to_pos):
        if from_pos is None or to_pos is None or from_pos == to_pos:
            return False
        item = items.pop(from_pos)
        items.insert(to_pos, item)
        return True


class RotateTab(BaseTab):
    def __init__(self, app, notebook):
        super().__init__(app, notebook)
        self.pdf_path = None
        self.pages = []
        self.page_widgets = []
        self._build()

    def _build(self):
        toolbar = ttk.Frame(self, style="App.TFrame")
        toolbar.pack(fill="x", pady=(0, 10))
        rounded_button(toolbar, "開啟 PDF", self.open_pdf, accent=True).pack(
            side="left"
        )
        rounded_button(toolbar, "全部左轉", lambda: self.rotate_all(-90)).pack(
            side="left", padx=4
        )
        rounded_button(toolbar, "全部右轉", lambda: self.rotate_all(90)).pack(
            side="left"
        )
        rounded_button(toolbar, "全部重設", self.reset_all).pack(side="left", padx=4)
        rounded_button(toolbar, "回復上一動作", self.undo_last_action).pack(
            side="left", padx=4
        )
        ttk.Label(toolbar, text="縮圖", style="App.TLabel").pack(
            side="left", padx=(16, 4)
        )
        ttk.Scale(
            toolbar,
            from_=90,
            to=230,
            variable=self.thumb_size,
            command=lambda _v: self.render(),
        ).pack(side="left")
        rounded_button(toolbar, "輸出修改後 PDF", self.export_pdf, accent=True).pack(
            side="right"
        )
        self.title = ttk.Label(
            self, text="請開啟或拖曳 PDF 到此分頁", style="App.TLabel"
        )
        self.title.pack(anchor="w", pady=(0, 8))
        self.area = ScrollArea(self)
        self.area.pack(fill="both", expand=True)
        self.enable_responsive_layout(self.area)
        self.enable_drop(self, self.load_drop)
        self.bind_all("<Control-MouseWheel>", self.zoom)

    def zoom(self, event):
        value = self.thumb_size.get() + (10 if event.delta > 0 else -10)
        self.thumb_size.set(max(90, min(230, value)))
        self.render()

    def open_pdf(self):
        path = self.choose_pdf()
        if path:
            self.load_pdf(path)

    def load_drop(self, paths, _event=None):
        if paths:
            self.load_pdf(paths[0])

    def load_pdf(self, path):
        try:
            doc = fitz.open(path)
            count = doc.page_count
            doc.close()
        except Exception as exc:
            messagebox.showerror("開啟失敗", f"無法開啟 PDF：\n{exc}")
            return
        self.pdf_path = path
        self.pages = [{"index": i, "rotation": 0} for i in range(count)]
        self.undo_stack = []
        self.title.configure(text=f"{os.path.basename(path)} / {count} 頁")
        self.render()

    def get_undo_snapshot(self):
        if not self.pdf_path:
            return None
        return {
            "pdf_path": self.pdf_path,
            "pages": [dict(item) for item in self.pages],
        }

    def restore_undo_snapshot(self, snapshot):
        self.pdf_path = snapshot["pdf_path"]
        self.pages = [dict(item) for item in snapshot["pages"]]
        self.title.configure(
            text=f"{os.path.basename(self.pdf_path)} / {len(self.pages)} 頁"
        )
        self.render()

    def render(self):
        self.clear_frame(self.area.content)
        self.page_widgets = []
        if not self.pdf_path:
            return
        columns = self.get_columns()
        image_box_height = int(self.thumb_size.get() * 1.45)
        for pos, item in enumerate(self.pages):
            card = ttk.Frame(self.area.content, padding=8, style="Card.TFrame")
            card.grid(
                row=pos // columns, column=pos % columns, padx=6, pady=6, sticky="n"
            )
            self.drag_cards.append(card)
            thumb = make_thumbnail(
                self.pdf_path, item["index"], self.thumb_size.get(), item["rotation"]
            )
            self.thumbs.append(thumb)
            image_box = ttk.Frame(
                card,
                width=self.thumb_size.get(),
                height=image_box_height,
                style="Card.TFrame",
            )
            image_box.pack_propagate(False)
            image_box.pack()
            label = ttk.Label(image_box, image=thumb)
            label.place(relx=0.5, rely=0.5, anchor="center")
            self.bind_drag_sort(label, pos)
            page_label = ttk.Label(
                card,
                text=f"第 {pos + 1} 頁 / {item['rotation'] % 360}°",
                style="Card.TLabel",
            )
            page_label.pack(pady=(6, 4))
            buttons = ttk.Frame(card, style="Card.TFrame")
            buttons.pack()
            rounded_button(
                buttons, "左轉", lambda p=pos: self.rotate_one(p, -90), width=58
            ).pack(side="left")
            rounded_button(
                buttons, "右轉", lambda p=pos: self.rotate_one(p, 90), width=58
            ).pack(side="left", padx=2)
            rounded_button(
                buttons, "重設", lambda p=pos: self.reset_one(p), width=58
            ).pack(side="left")
            self.page_widgets.append({"image": label, "label": page_label})

    def get_columns(self):
        return self.columns_for_width(self.thumb_size.get() + 48)

    def update_page_view(self, pos):
        if not self.pdf_path or pos >= len(self.page_widgets):
            return
        item = self.pages[pos]
        thumb = make_thumbnail(
            self.pdf_path, item["index"], self.thumb_size.get(), item["rotation"]
        )
        self.thumbs[pos] = thumb
        widgets = self.page_widgets[pos]
        widgets["image"].configure(image=thumb)
        widgets["image"].image = thumb
        widgets["label"].configure(text=f"第 {pos + 1} 頁 / {item['rotation'] % 360}°")

    def end_drag(self, event):
        to_pos = self.find_drop_position(event)
        if (
            self.drag_from is not None
            and to_pos is not None
            and self.drag_from != to_pos
        ):
            self.push_undo()
        moved = self.move_item(self.pages, self.drag_from, to_pos)
        self.drag_from = None
        if moved:
            self.render()

    def rotate_one(self, pos, degrees):
        self.push_undo()
        self.pages[pos]["rotation"] = (self.pages[pos]["rotation"] + degrees) % 360
        self.update_page_view(pos)

    def reset_one(self, pos):
        self.push_undo()
        self.pages[pos]["rotation"] = 0
        self.update_page_view(pos)

    def rotate_all(self, degrees):
        self.push_undo()
        for pos, item in enumerate(self.pages):
            item["rotation"] = (item["rotation"] + degrees) % 360
            self.update_page_view(pos)

    def reset_all(self):
        self.push_undo()
        for pos, item in enumerate(self.pages):
            item["rotation"] = 0
            self.update_page_view(pos)

    def export_pdf(self):
        if not self.pdf_path:
            messagebox.showwarning("尚未開啟", "請先開啟 PDF。")
            return
        output = filedialog.asksaveasfilename(
            title="另存新檔",
            defaultextension=".pdf",
            initialfile=os.path.basename(default_save_path(self.pdf_path, "_r")),
            initialdir=os.path.dirname(self.pdf_path),
            filetypes=[("PDF files", "*.pdf")],
        )
        if not output:
            return
        source = fitz.open(self.pdf_path)
        result = fitz.open()
        try:
            for item in self.pages:
                result.insert_pdf(
                    source, from_page=item["index"], to_page=item["index"]
                )
                page = result[-1]
                page.set_rotation((page.rotation + item["rotation"]) % 360)
            result.save(output, garbage=4, deflate=True)
            messagebox.showinfo("完成", f"已輸出：\n{output}")
        except Exception as exc:
            messagebox.showerror("輸出失敗", str(exc))
        finally:
            result.close()
            source.close()


class CompressTab(BaseTab):
    def __init__(self, app, notebook):
        super().__init__(app, notebook)
        self.pdf_path = None
        self.quality = tk.IntVar(value=70)
        self.original_size = tk.StringVar(value="原有檔案大小\n-")
        self.estimated_size = tk.StringVar(value="預計壓縮後大小\n-")
        self.actual_size = tk.StringVar(value="壓縮後實際大小\n-")
        self._build()

    def _build(self):
        toolbar = ttk.Frame(self, style="App.TFrame")
        toolbar.pack(fill="x", pady=(0, 10))
        rounded_button(toolbar, "開啟 PDF", self.open_pdf, accent=True).pack(
            side="left"
        )
        rounded_button(toolbar, "回復上一動作", self.undo_last_action).pack(
            side="left", padx=4
        )
        ttk.Label(toolbar, text="壓縮比例", style="App.TLabel").pack(
            side="left", padx=(16, 6)
        )
        self.quality_scale = ttk.Scale(
            toolbar,
            from_=35,
            to=95,
            variable=self.quality,
            command=lambda _v: self.update_estimate(),
        )
        self.quality_scale.pack(side="left", fill="x", expand=True)
        self.quality_scale.bind("<ButtonPress-1>", lambda _event: self.push_undo())
        rounded_button(toolbar, "輸出壓縮PDF", self.export_pdf, accent=True).pack(
            side="right", padx=(10, 0)
        )
        self.info = ttk.Label(
            self, text="請開啟或拖曳 PDF 到此分頁", style="App.TLabel"
        )
        self.info.pack(anchor="w", pady=(0, 12))
        stats = ttk.Frame(self, style="App.TFrame")
        stats.pack(fill="x", pady=(0, 12))
        for variable in (self.original_size, self.estimated_size, self.actual_size):
            ttk.Label(
                stats,
                textvariable=variable,
                style="Stat.TLabel",
                anchor="center",
                justify="center",
                relief="raised",
                borderwidth=2,
            ).pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.enable_drop(self, self.load_drop)

    def open_pdf(self):
        path = self.choose_pdf()
        if path:
            self.load_pdf(path)

    def load_drop(self, paths, _event=None):
        if paths:
            self.load_pdf(paths[0])

    def load_pdf(self, path):
        self.push_undo()
        self.pdf_path = path
        self.actual_size.set("壓縮後實際大小\n-")
        self.update_estimate()

    def get_undo_snapshot(self):
        if not self.pdf_path:
            return None
        return {
            "pdf_path": self.pdf_path,
            "quality": self.quality.get(),
            "actual_size": self.actual_size.get(),
        }

    def restore_undo_snapshot(self, snapshot):
        self.pdf_path = snapshot["pdf_path"]
        self.quality.set(snapshot["quality"])
        self.actual_size.set(snapshot["actual_size"])
        self.update_estimate()

    def update_estimate(self):
        if not self.pdf_path:
            return
        original = os.path.getsize(self.pdf_path)
        quality = self.quality.get()
        factor = max(0.20, quality / 115)
        estimate = original * factor
        self.info.configure(text=f"{os.path.basename(self.pdf_path)} / 品質 {quality}")
        self.original_size.set(f"原有檔案大小\n{file_size_text(self.pdf_path)}")
        self.estimated_size.set(f"預計壓縮後大小\n{estimate / 1024 / 1024:.1f} MB")

    def export_pdf(self):
        if not self.pdf_path:
            messagebox.showwarning("尚未開啟", "請先開啟 PDF。")
            return
        output = filedialog.asksaveasfilename(
            title="另存新檔",
            defaultextension=".pdf",
            initialfile=os.path.basename(default_save_path(self.pdf_path, "_comp")),
            initialdir=os.path.dirname(self.pdf_path),
            filetypes=[("PDF files", "*.pdf")],
        )
        if not output:
            return
        quality = self.quality.get()
        zoom = 0.9 + (quality / 100)
        source = fitz.open(self.pdf_path)
        result = fitz.open()
        try:
            for page in source:
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                temp = Path(output).with_suffix(".jpg")
                img.save(temp, "JPEG", quality=quality, optimize=True)
                rect = fitz.Rect(0, 0, page.rect.width, page.rect.height)
                new_page = result.new_page(
                    width=page.rect.width, height=page.rect.height
                )
                new_page.insert_image(rect, filename=str(temp))
                temp.unlink(missing_ok=True)
            result.save(output, garbage=4, deflate=True)
            self.actual_size.set(f"壓縮後實際大小\n{file_size_text(output)}")
            messagebox.showinfo(
                "完成", f"已輸出：\n{output}\n\n實際大小：{file_size_text(output)}"
            )
        except Exception as exc:
            messagebox.showerror("輸出失敗", str(exc))
        finally:
            result.close()
            source.close()


class MergeTab(BaseTab):
    def __init__(self, app, notebook):
        super().__init__(app, notebook)
        self.files = []
        self.output_preview_images = []
        self._build()

    def _build(self):
        toolbar = ttk.Frame(self, style="App.TFrame")
        toolbar.pack(fill="x", pady=(0, 10))
        rounded_button(toolbar, "加入 PDF", self.add_files, accent=True).pack(
            side="left"
        )
        rounded_button(toolbar, "清空", self.clear).pack(side="left", padx=4)
        rounded_button(toolbar, "回復上一動作", self.undo_last_action).pack(side="left")
        rounded_button(toolbar, "輸出合併 PDF", self.export_pdf, accent=True).pack(
            side="right"
        )
        self.title = ttk.Label(
            self, text="請加入或拖曳 PDF 到此分頁", style="App.TLabel"
        )
        self.title.pack(anchor="w", pady=(0, 8))

        # V0.6.1：頁面合併中間區改成左右分割，左側作業清單、右側輸出後預覽，比例 8:2。
        split = ttk.Frame(self, style="App.TFrame")
        split.pack(fill="both", expand=True)
        split.rowconfigure(0, weight=1)
        split.columnconfigure(0, weight=8, uniform="merge_output_preview_split")
        split.columnconfigure(1, weight=2, uniform="merge_output_preview_split")

        list_frame = ttk.LabelFrame(
            split, text="合併清單", padding=4, style="App.TLabelframe"
        )
        list_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        output_preview = ttk.LabelFrame(
            split, text="合併輸出預覽", padding=4, style="App.TLabelframe"
        )
        output_preview.grid(row=0, column=1, sticky="nsew")

        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)
        self.area = ScrollArea(list_frame)
        self.area.pack(fill="both", expand=True)
        self.enable_responsive_layout(self.area)

        output_preview.rowconfigure(0, weight=1)
        output_preview.columnconfigure(0, weight=1)
        self.merge_output_preview_canvas = tk.Canvas(
            output_preview, bg=FILENAME_BG, highlightthickness=0
        )
        self.merge_output_preview_canvas.grid(row=0, column=0, sticky="nsew")
        output_vbar = ttk.Scrollbar(
            output_preview,
            orient="vertical",
            command=self.merge_output_preview_canvas.yview,
        )
        output_vbar.grid(row=0, column=1, sticky="ns")
        self.merge_output_preview_canvas.configure(yscrollcommand=output_vbar.set)
        self.merge_output_preview_inner = ttk.Frame(
            self.merge_output_preview_canvas, style="Card.TFrame"
        )
        self.merge_output_preview_window = self.merge_output_preview_canvas.create_window(
            (0, 0), window=self.merge_output_preview_inner, anchor="nw"
        )
        self.merge_output_preview_inner.bind(
            "<Configure>",
            lambda _event: self.merge_output_preview_canvas.configure(
                scrollregion=self.merge_output_preview_canvas.bbox("all")
            ),
        )
        self.merge_output_preview_canvas.bind(
            "<Configure>",
            lambda event: self.merge_output_preview_canvas.itemconfigure(
                self.merge_output_preview_window, width=event.width
            ),
        )
        self.clear_merge_output_preview("輸出合併 PDF 後會在這裡顯示預覽。")

        self.enable_drop(self, self.load_drop)
        self.enable_drop(list_frame, self.load_drop)
        self.enable_drop(self.area.canvas, self.load_drop)
        self.enable_drop(self.area.content, self.load_drop)
        self.enable_drop(output_preview, self.load_drop)

    def clear_merge_output_preview(self, message="尚未產生合併輸出預覽。"):
        if not hasattr(self, "merge_output_preview_inner"):
            return
        for child in self.merge_output_preview_inner.winfo_children():
            child.destroy()
        self.output_preview_images = []
        ttk.Label(
            self.merge_output_preview_inner,
            text=message,
            style="Muted.TLabel",
            wraplength=180,
            justify="center",
        ).pack(fill="x", padx=8, pady=12)
        with suppress(Exception):
            self.merge_output_preview_canvas.configure(
                scrollregion=self.merge_output_preview_canvas.bbox("all")
            )

    def load_merge_output_preview(self, pdf_path):
        if not pdf_path or not os.path.exists(pdf_path):
            self.clear_merge_output_preview("找不到合併輸出的 PDF 檔案。")
            return
        if not ensure_packages():
            return

        for child in self.merge_output_preview_inner.winfo_children():
            child.destroy()
        self.output_preview_images = []

        doc = None
        try:
            doc = fitz.open(pdf_path)
            with suppress(Exception):
                self.merge_output_preview_canvas.update_idletasks()
            canvas_width = int(self.merge_output_preview_canvas.winfo_width() or 180)
            thumb_width = max(120, min(240, canvas_width - 28))

            ttk.Label(
                self.merge_output_preview_inner,
                text=os.path.basename(pdf_path),
                style="Card.TLabel",
                wraplength=max(120, thumb_width),
                justify="center",
            ).pack(fill="x", padx=6, pady=(6, 8))

            for index, page in enumerate(doc):
                zoom = max(thumb_width / max(page.rect.width, 1), 0.05)
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
                image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                photo = ImageTk.PhotoImage(image)
                self.output_preview_images.append(photo)

                item = ttk.Frame(self.merge_output_preview_inner, style="Card.TFrame")
                item.pack(fill="x", padx=6, pady=(0, 10))
                tk.Label(item, image=photo, bg=CARD, bd=1, relief="solid").pack()
                ttk.Label(
                    item,
                    text=f"第 {index + 1} 頁",
                    style="Muted.TLabel",
                    anchor="center",
                ).pack(fill="x", pady=(2, 0))

            self.merge_output_preview_canvas.configure(
                scrollregion=self.merge_output_preview_canvas.bbox("all")
            )
            self.title.configure(
                text=f"合併清單 / {len(self.files)} 個檔案 / 已預覽：{os.path.basename(pdf_path)}"
            )
        except Exception as exc:
            self.clear_merge_output_preview("合併輸出 PDF 預覽失敗。")
            messagebox.showwarning("預覽失敗", f"合併 PDF 已輸出，但預覽失敗：\n{exc}")
        finally:
            if doc is not None:
                doc.close()

    def add_files(self):
        self.load_drop(self.choose_pdfs())

    def load_drop(self, paths, _event=None):
        if not paths:
            return
        self.push_undo()
        self.files.extend(paths)
        self.clear_merge_output_preview("合併清單已變更，輸出後會重新顯示預覽。")
        self.render()

    def clear(self):
        if self.files:
            self.push_undo()
        self.files = []
        self.clear_merge_output_preview("輸出合併 PDF 後會在這裡顯示預覽。")
        self.render()

    def get_undo_snapshot(self):
        return list(self.files)

    def restore_undo_snapshot(self, snapshot):
        self.files = list(snapshot)
        self.clear_merge_output_preview("合併清單已回復，輸出後會重新顯示預覽。")
        self.render()

    def render(self):
        self.clear_frame(self.area.content)
        self.title.configure(text=f"合併清單 / {len(self.files)} 個檔案")
        columns = self.get_columns()
        for pos, path in enumerate(self.files):
            card = ttk.Frame(self.area.content, padding=8, style="Card.TFrame")
            card.grid(
                row=pos // columns, column=pos % columns, padx=6, pady=6, sticky="n"
            )
            self.drag_cards.append(card)
            thumb = make_thumbnail(path, 0, 150)
            self.thumbs.append(thumb)
            label = ttk.Label(card, image=thumb)
            label.pack()
            self.bind_drag_sort(label, pos)
            ttk.Label(
                card,
                text=f"{pos + 1}. {os.path.basename(path)}",
                wraplength=160,
                style="Card.TLabel",
            ).pack(pady=(6, 4))
            rounded_button(card, "移除", lambda p=pos: self.remove(p), width=64).pack()

    def get_columns(self):
        return self.columns_for_width(198)

    def end_drag(self, event):
        to_pos = self.find_drop_position(event)
        if (
            self.drag_from is not None
            and to_pos is not None
            and self.drag_from != to_pos
        ):
            self.push_undo()
        moved = self.move_item(self.files, self.drag_from, to_pos)
        self.drag_from = None
        if moved:
            self.clear_merge_output_preview("合併順序已變更，輸出後會重新顯示預覽。")
            self.render()

    def remove(self, pos):
        self.push_undo()
        self.files.pop(pos)
        self.clear_merge_output_preview("合併清單已變更，輸出後會重新顯示預覽。")
        self.render()

    def export_pdf(self):
        if not self.files:
            messagebox.showwarning("沒有檔案", "請先加入 PDF。")
            return
        name = f"PDF_mer_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf"
        output = filedialog.asksaveasfilename(
            title="另存新檔",
            defaultextension=".pdf",
            initialfile=name,
            initialdir=os.path.dirname(self.files[0]),
            filetypes=[("PDF files", "*.pdf")],
        )
        if not output:
            return
        result = fitz.open()
        try:
            for path in self.files:
                doc = fitz.open(path)
                result.insert_pdf(doc)
                doc.close()
            result.save(output, garbage=4, deflate=True)
            self.load_merge_output_preview(output)
            messagebox.showinfo("完成", f"已輸出：\n{output}")
        except Exception as exc:
            messagebox.showerror("輸出失敗", str(exc))
        finally:
            result.close()


class EditTab(BaseTab):
    def __init__(self, app, notebook):
        super().__init__(app, notebook)
        self.pages = []
        self.base_path = None
        self.selected_page_pos = None
        self._edit_thumb_cache = {}
        self.quick_extract_var = tk.StringVar()
        self._build()

    def _build(self):
        edit_toolbar = ttk.Frame(self, style="App.TFrame")
        edit_toolbar.pack(fill="x", pady=(0, 6))
        rounded_button(edit_toolbar, "開啟 PDF", self.open_pdf, accent=True).pack(
            side="left"
        )
        rounded_button(edit_toolbar, "插入 PDF", self.insert_pdf).pack(side="left", padx=4)
        rounded_button(
            edit_toolbar,
            "插入空白頁",
            self.insert_blank_page,
        ).pack(side="left", padx=(0, 4))
        rounded_button(edit_toolbar, "清空", self.clear_workspace).pack(side="left", padx=(0, 4))
        rounded_button(edit_toolbar, "回復上一動作", self.undo_last_action).pack(side="left")
        rounded_button(edit_toolbar, "取消抽取", self.clear_extract_selection).pack(
            side="left", padx=4
        )
        rounded_button(edit_toolbar, "輸出編輯 PDF", self.export_pdf, accent=True).pack(
            side="right"
        )
        rounded_button(edit_toolbar, "抽取另存", self.export_extract_pdf, accent=True).pack(
            side="right", padx=(0, 4)
        )

        quick_extract_row = ttk.Frame(self, style="App.TFrame")
        quick_extract_row.pack(fill="x", pady=(0, 6))
        ttk.Label(
            quick_extract_row,
            text="快速抽取頁碼：",
            style="App.TLabel",
        ).pack(side="left")
        self.quick_extract_entry = ttk.Entry(
            quick_extract_row,
            textvariable=self.quick_extract_var,
            width=14,
        )
        self.quick_extract_entry.pack(side="left", padx=(0, 6))
        self.quick_extract_entry.bind(
            "<Return>",
            lambda _event: self.quick_extract_pages(),
        )
        rounded_button(
            quick_extract_row,
            "快速抽取",
            self.quick_extract_pages,
            accent=True,
        ).pack(side="left")
        ttk.Label(
            quick_extract_row,
            text="格式：01-08（包含起訖頁）",
            style="Muted.TLabel",
        ).pack(side="left", padx=(8, 0))

        self.title = ttk.Label(
            self,
            text="請開啟 PDF；拖曳其他 PDF 可插入頁面，勾選「抽取」後可另存部分頁面",
            style="App.TLabel",
        )
        self.title.pack(anchor="w", pady=(0, 8))
        self.area = ScrollArea(self)
        self.area.pack(fill="both", expand=True)
        self.enable_responsive_layout(self.area)
        self.enable_drop(self, self.load_drop)
        self.enable_drop(self.area.canvas, self.load_drop)
        self.enable_drop(self.area.content, self.load_drop)

    def clear_workspace(self):
        """清除頁面編輯工作區，並保留一次回復上一動作的能力。"""
        if not self.base_path and not self.pages:
            return
        self.push_undo()
        self.pages = []
        self.base_path = None
        self.selected_page_pos = None
        self.drag_from = None
        self._edit_thumb_cache.clear()
        self.quick_extract_var.set("")
        self.clear_frame(self.area.content)
        self.title.configure(
            text="請開啟 PDF；拖曳其他 PDF 可插入頁面，勾選「抽取」後可另存部分頁面"
        )
        with suppress(Exception):
            self.area.canvas.yview_moveto(0)
        with suppress(Exception):
            self.area.canvas.xview_moveto(0)

    def open_pdf(self):
        path = self.choose_pdf()
        if path:
            self.load_base(path)

    def insert_pdf(self):
        self.load_drop(self.choose_pdfs())

    def insert_blank_page(self):
        """在目前點選頁面的下一頁插入相同尺寸的空白頁。"""
        if not self.pages or not self.base_path:
            messagebox.showwarning("插入空白頁", "請先開啟 PDF。")
            return

        pos = self.selected_page_pos
        if pos is None or not (0 <= pos < len(self.pages)):
            messagebox.showwarning(
                "插入空白頁",
                "請先點選要插入位置前方的頁面縮圖。",
            )
            return

        source_item = self.pages[pos]
        temp_path = None
        try:
            source_doc = fitz.open(source_item["path"])
            try:
                source_page = source_doc[source_item["index"]]
                page_rect = source_page.rect
                width = float(page_rect.width)
                height = float(page_rect.height)
            finally:
                source_doc.close()

            fd, temp_path = tempfile.mkstemp(
                prefix="guppy_blank_page_",
                suffix=".pdf",
            )
            os.close(fd)

            blank_doc = fitz.open()
            try:
                blank_doc.new_page(width=width, height=height)
                blank_doc.save(temp_path, garbage=4, deflate=True)
            finally:
                blank_doc.close()

            self.push_undo()
            insert_at = pos + 1
            self.pages.insert(
                insert_at,
                self.make_page_item(
                    temp_path,
                    0,
                    inserted=True,
                    blank=True,
                ),
            )
            self.selected_page_pos = insert_at
            self.render()
        except Exception as exc:
            if temp_path:
                with suppress(Exception):
                    os.remove(temp_path)
            messagebox.showerror("插入空白頁失敗", str(exc))

    def load_drop(self, paths, event=None):
        if not paths:
            return
        if not self.base_path:
            self.load_base(paths[0])
            if len(paths) > 1:
                self.push_undo()
            for path in paths[1:]:
                self.add_insert(path)
        else:
            self.push_undo()
            insert_at = self.find_drop_position(event) if event else None
            if insert_at is not None:
                insert_at += 1
            for path in paths:
                insert_at = self.add_insert(path, insert_at)
        self.render()

    def make_page_item(
        self,
        path,
        index,
        inserted=False,
        delete=False,
        extract=False,
        blank=False,
    ):
        return {
            "path": path,
            "index": index,
            "inserted": inserted,
            "blank": blank,
            "delete": tk.BooleanVar(value=delete),
            "extract": tk.BooleanVar(value=extract),
        }

    def load_base(self, path):
        try:
            doc = fitz.open(path)
            count = doc.page_count
            doc.close()
        except Exception as exc:
            messagebox.showerror("開啟失敗", str(exc))
            return
        self.push_undo()
        self.base_path = path
        self.selected_page_pos = None
        self.quick_extract_var.set("")
        self._edit_thumb_cache.clear()
        self.pages = [self.make_page_item(path, i, inserted=False) for i in range(count)]
        self.render()

    def get_undo_snapshot(self):
        if not self.base_path:
            return None
        return {
            "base_path": self.base_path,
            "selected_page_pos": self.selected_page_pos,
            "pages": [
                {
                    "path": item["path"],
                    "index": item["index"],
                    "inserted": item["inserted"],
                    "blank": item.get("blank", False),
                    "delete": item["delete"].get(),
                    "extract": item.get("extract", tk.BooleanVar(value=False)).get(),
                }
                for item in self.pages
            ],
        }

    def restore_undo_snapshot(self, snapshot):
        self.base_path = snapshot["base_path"]
        self.selected_page_pos = snapshot.get("selected_page_pos")
        self.pages = [
            self.make_page_item(
                item["path"],
                item["index"],
                inserted=item["inserted"],
                delete=item.get("delete", False),
                extract=item.get("extract", False),
                blank=item.get("blank", False),
            )
            for item in snapshot["pages"]
        ]
        self.render()

    def add_insert(self, path, insert_at=None):
        try:
            doc = fitz.open(path)
            count = doc.page_count
            doc.close()
        except Exception as exc:
            messagebox.showerror("插入失敗", f"{path}\n{exc}")
            return insert_at
        new_pages = [
            self.make_page_item(path, i, inserted=True)
            for i in range(count)
        ]
        if insert_at is None:
            self.pages.extend(new_pages)
            return None
        self.pages[insert_at:insert_at] = new_pages
        return insert_at + len(new_pages)

    def page_display_text(self, pos, item):
        text = f"第 {pos + 1} 頁"
        if item.get("blank"):
            text += "（空白頁）"
        if pos == self.selected_page_pos:
            text = f"▶ {text}"
        return text

    @staticmethod
    def thumbnail_mark_key(mark):
        if mark is None:
            return None
        return tuple(mark)

    def get_edit_thumbnail(self, item, mark=None):
        """重用相同頁面縮圖，避免每次點選都重新開啟 PDF 轉圖。"""
        key = (
            os.path.abspath(item["path"]),
            int(item["index"]),
            150,
            self.thumbnail_mark_key(mark),
        )
        thumb = self._edit_thumb_cache.get(key)
        if thumb is None:
            thumb = make_thumbnail(
                item["path"],
                item["index"],
                150,
                mark=mark,
            )
            self._edit_thumb_cache[key] = thumb
        return thumb

    def update_selected_page_display(self, old_pos=None, new_pos=None):
        """只更新舊、新選取頁的文字標籤，不重畫全部縮圖。"""
        for pos in {old_pos, new_pos}:
            if pos is None or not (0 <= pos < len(self.pages)):
                continue
            item = self.pages[pos]
            widget = item.get("page_label_widget")
            if widget is not None and widget.winfo_exists():
                widget.configure(text=self.page_display_text(pos, item))
        self.update_extract_title()

    def render(self):
        self.clear_frame(self.area.content)
        if not self.base_path:
            return
        selected_count = sum(
            1 for item in self.pages if item.get("extract") and item["extract"].get()
        )
        self.title.configure(
            text=f"{os.path.basename(self.base_path)} / 目前 {len(self.pages)} 頁 / 頁面抽取已選 {selected_count} 頁"
        )
        columns = self.get_columns()
        for pos, item in enumerate(self.pages):
            if item.get("blank"):
                color = (255, 210, 80, 65)
            else:
                color = (80, 160, 255, 70) if item["inserted"] else None
            if item["delete"].get():
                color = (255, 80, 80, 90)
            card = ttk.Frame(self.area.content, padding=8, style="Card.TFrame")
            card.grid(
                row=pos // columns, column=pos % columns, padx=6, pady=6, sticky="n"
            )
            self.drag_cards.append(card)
            thumb = self.get_edit_thumbnail(item, mark=color)
            self.thumbs.append(thumb)
            label = ttk.Label(card, image=thumb)
            label.pack()
            self.bind_drag_sort(label, pos)
            page_label = ttk.Label(
                card,
                text=self.page_display_text(pos, item),
                wraplength=230,
                style="Card.TLabel",
            )
            page_label.pack(pady=(6, 2))
            item["page_label_widget"] = page_label

            delete_check = ttk.Checkbutton(
                card,
                text="刪除",
                variable=item["delete"],
                command=lambda i=item, l=label: self.update_delete_mark(i, l),
            )
            delete_check.pack(anchor="w")
            delete_check.bind(
                "<ButtonPress-1>", lambda _event: self.push_undo(), add="+"
            )
            item["delete_widget"] = delete_check

            extract_row = ttk.Frame(card, style="Card.TFrame")
            extract_row.pack(fill="x", pady=(4, 0))
            extract_check = ttk.Checkbutton(
                extract_row,
                text="抽取",
                variable=item["extract"],
                command=lambda i=item: self.update_extract_mark(i),
            )
            extract_check.grid(row=0, column=0, sticky="w", padx=(0, 4))
            extract_check.bind(
                "<ButtonPress-1>", lambda _event: self.push_undo(), add="+"
            )
            item["extract_widget"] = extract_check

        self.update_edit_mode_controls()

    def update_delete_mark(self, item, label):
        """Refresh only the clicked page thumbnail when the delete check changes.

        Rebuilding every thumbnail on each delete click makes large PDFs feel
        frozen for a moment.  This keeps the immediate red/blue page mark while
        avoiding a full render pass.
        V0.4.0 also makes delete/extract mutually exclusive across the whole tab.
        """
        if item["delete"].get() and self.has_extract_selection():
            item["delete"].set(False)
            self.update_edit_mode_controls()
            return

        if item["delete"].get() and item.get("extract") and item["extract"].get():
            item["extract"].set(False)

        if item.get("blank"):
            color = (255, 210, 80, 65)
        else:
            color = (80, 160, 255, 70) if item["inserted"] else None
        if item["delete"].get():
            color = (255, 80, 80, 90)
        thumb = self.get_edit_thumbnail(item, mark=color)
        label.configure(image=thumb)
        label.image = thumb
        item["thumb_ref"] = thumb
        self.update_edit_mode_controls()

    def get_columns(self):
        return self.columns_for_width(220)

    def has_delete_selection(self):
        return any(item["delete"].get() for item in self.pages)

    def has_extract_selection(self):
        return any(
            item.get("extract") and item["extract"].get() for item in self.pages
        )

    def edit_mode(self):
        if self.has_delete_selection():
            return "delete"
        if self.has_extract_selection():
            return "extract"
        return "none"

    def update_edit_mode_controls(self):
        mode = self.edit_mode()
        for item in self.pages:
            delete_widget = item.get("delete_widget")
            extract_widget = item.get("extract_widget")
            if delete_widget is not None:
                delete_widget.configure(state="disabled" if mode == "extract" else "normal")
            if extract_widget is not None:
                extract_widget.configure(state="disabled" if mode == "delete" else "normal")
        self.update_extract_title()

    def update_extract_mark(self, item):
        if item["extract"].get() and self.has_delete_selection():
            item["extract"].set(False)
            self.update_edit_mode_controls()
            return

        if item["extract"].get() and item["delete"].get():
            item["delete"].set(False)

        self.update_edit_mode_controls()

    def update_extract_title(self):
        if not self.base_path:
            return
        extract_count = sum(
            1 for item in self.pages if item.get("extract") and item["extract"].get()
        )
        delete_count = sum(1 for item in self.pages if item["delete"].get())
        if delete_count:
            mode_text = f"刪除已選 {delete_count} 頁 / 抽取停用"
        elif extract_count:
            mode_text = f"頁面抽取已選 {extract_count} 頁 / 刪除停用"
        else:
            mode_text = "尚未選擇刪除或抽取"
        selected_text = ""
        if (
            self.selected_page_pos is not None
            and 0 <= self.selected_page_pos < len(self.pages)
        ):
            selected_text = f" / 已點選第 {self.selected_page_pos + 1} 頁"
        self.title.configure(
            text=(
                f"{os.path.basename(self.base_path)} / 目前 {len(self.pages)} 頁"
                f" / {mode_text}{selected_text}"
            )
        )

    def select_all_extract(self):
        if not self.pages:
            messagebox.showinfo("頁面抽取", "請先開啟 PDF。")
            return
        if self.has_delete_selection():
            return
        for item in self.pages:
            item["extract"].set(True)
        self.update_edit_mode_controls()

    def clear_extract_selection(self):
        if not self.pages:
            return
        for item in self.pages:
            item["extract"].set(False)
        self.update_edit_mode_controls()

    @staticmethod
    def normalize_page_range_text(value):
        """將全形或長破折號統一為半形連字號。"""
        return (
            str(value or "")
            .strip()
            .replace("－", "-")
            .replace("–", "-")
            .replace("—", "-")
        )

    def parse_quick_extract_range(self):
        """解析 01-08 格式，回傳零起算起點、終點及檔名用範圍文字。"""
        raw = self.normalize_page_range_text(self.quick_extract_var.get())
        match = re.fullmatch(r"\s*(\d+)\s*-\s*(\d+)\s*", raw)
        if not match:
            raise ValueError("頁碼格式錯誤，請輸入例如：01-08")

        start_page = int(match.group(1))
        end_page = int(match.group(2))

        if start_page < 1 or end_page < 1:
            raise ValueError("頁碼必須從第 1 頁開始。")
        if start_page > end_page:
            raise ValueError("起始頁碼不可大於結束頁碼。")
        if end_page > len(self.pages):
            raise ValueError(
                f"結束頁碼超過目前工作區頁數（共 {len(self.pages)} 頁）。"
            )

        range_text = f"{start_page:02d}-{end_page:02d}"
        return start_page - 1, end_page - 1, range_text

    def quick_extract_output_path(self, range_text):
        """快速抽取固定儲存在原始 PDF 同一資料夾。"""
        source = Path(self.base_path)
        return str(source.with_name(f"{source.stem}-sp_{range_text}.pdf"))

    def quick_extract_pages(self):
        """依目前工作區頁序快速抽取指定的連續頁碼範圍。"""
        if not self.pages or not self.base_path:
            messagebox.showwarning("快速抽取", "請先開啟 PDF。")
            return

        try:
            start_pos, end_pos, range_text = self.parse_quick_extract_range()
        except ValueError as exc:
            messagebox.showwarning("快速抽取", str(exc))
            with suppress(Exception):
                self.quick_extract_entry.focus_set()
                self.quick_extract_entry.selection_range(0, "end")
            return

        selected_items = self.pages[start_pos : end_pos + 1]
        output = self.quick_extract_output_path(range_text)

        if os.path.exists(output):
            overwrite = messagebox.askyesno(
                "檔案已存在",
                f"檔案已存在，是否覆蓋？\n\n{output}",
            )
            if not overwrite:
                return
            with suppress(Exception):
                os.remove(output)

        result = fitz.open()
        open_docs = {}
        try:
            for item in selected_items:
                path = item["path"]
                if path not in open_docs:
                    open_docs[path] = fitz.open(path)
                result.insert_pdf(
                    open_docs[path],
                    from_page=item["index"],
                    to_page=item["index"],
                )

            if result.page_count == 0:
                messagebox.showwarning("快速抽取", "指定範圍沒有可輸出的頁面。")
                return

            result.save(output, garbage=4, deflate=True)
            self.title.configure(
                text=(
                    f"{os.path.basename(self.base_path)} / "
                    f"已快速抽取 {range_text}，共 {len(selected_items)} 頁 / "
                    f"輸出：{output}"
                )
            )
            messagebox.showinfo(
                "快速抽取完成",
                f"已抽取第 {range_text} 頁：\n{output}",
            )
        except Exception as exc:
            with suppress(Exception):
                if os.path.exists(output):
                    os.remove(output)
            messagebox.showerror("快速抽取失敗", str(exc))
        finally:
            result.close()
            for doc in open_docs.values():
                doc.close()

    def default_extract_output_path(self, output_folder=None):
        if not self.base_path:
            return ""
        source = Path(self.base_path)
        output_dir = Path(output_folder) if output_folder else source.parent
        date_text = datetime.now().strftime("%Y%m%d")
        return str(output_dir / f"{source.stem}_part_{date_text}.pdf")

    def export_extract_pdf(self):
        if not self.pages or not self.base_path:
            messagebox.showwarning("尚未開啟", "請先開啟 PDF。")
            return
        selected = [item for item in self.pages if item["extract"].get()]
        if not selected:
            messagebox.showwarning("尚未勾選", "請先勾選要抽取的頁面。")
            return

        initial_dir = os.path.dirname(self.base_path) if self.base_path else os.getcwd()
        output_folder = filedialog.askdirectory(
            title="選擇抽取另存資料夾",
            initialdir=initial_dir,
        )
        if not output_folder:
            return
        try:
            Path(output_folder).mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            messagebox.showerror("資料夾錯誤", f"無法建立或使用另存資料夾：\n{exc}")
            return

        output = self.default_extract_output_path(output_folder)
        if os.path.exists(output):
            overwrite = messagebox.askyesno(
                "檔案已存在",
                f"檔案已存在，是否覆蓋？\n\n{output}",
            )
            if not overwrite:
                return

        result = fitz.open()
        open_docs = {}
        try:
            for item in selected:
                path = item["path"]
                if path not in open_docs:
                    open_docs[path] = fitz.open(path)
                result.insert_pdf(
                    open_docs[path], from_page=item["index"], to_page=item["index"]
                )
            result.save(output, garbage=4, deflate=True)
            self.title.configure(
                text=f"{os.path.basename(self.base_path)} / 已抽取 {len(selected)} 頁另存至：{output}"
            )
        except Exception as exc:
            messagebox.showerror("頁面抽取失敗", str(exc))
        finally:
            result.close()
            for doc in open_docs.values():
                doc.close()

    def end_drag(self, event):
        from_pos = self.drag_from
        to_pos = self.find_drop_position(event)
        old_selected = self.selected_page_pos

        if (
            from_pos is not None
            and to_pos is not None
            and from_pos != to_pos
        ):
            self.push_undo()

        moved = self.move_item(self.pages, from_pos, to_pos)
        self.drag_from = None

        if moved and to_pos is not None:
            # 真正拖動換頁才重新排列；縮圖由快取直接取用。
            self.selected_page_pos = to_pos
            self.render()
            return

        if from_pos is not None and 0 <= from_pos < len(self.pages):
            # 單純點選只更新兩個文字標籤，不重新產生全部縮圖。
            self.selected_page_pos = from_pos
            self.update_selected_page_display(old_selected, from_pos)

    def export_pdf(self):
        if not self.pages:
            messagebox.showwarning("尚未開啟", "請先開啟 PDF。")
            return
        output = filedialog.asksaveasfilename(
            title="另存新檔",
            defaultextension=".pdf",
            initialfile=os.path.basename(default_save_path(self.base_path, "_edit")),
            initialdir=os.path.dirname(self.base_path),
            filetypes=[("PDF files", "*.pdf")],
        )
        if not output:
            return
        result = fitz.open()
        open_docs = {}
        try:
            for item in self.pages:
                if item["delete"].get():
                    continue
                path = item["path"]
                if path not in open_docs:
                    open_docs[path] = fitz.open(path)
                result.insert_pdf(
                    open_docs[path], from_page=item["index"], to_page=item["index"]
                )
            if result.page_count == 0:
                messagebox.showwarning("沒有頁面", "全部頁面都被刪除，無法輸出。")
                return
            result.save(output, garbage=4, deflate=True)
            messagebox.showinfo("完成", f"已輸出：\n{output}")
        except Exception as exc:
            messagebox.showerror("輸出失敗", str(exc))
        finally:
            result.close()
            for doc in open_docs.values():
                doc.close()


# =========================================================
# PPT Convert Tab (merged from PPToPDF V0.1)
# =========================================================
PPT_APP_NAME = "PPToPDF"
PPT_SUPPORTED_EXTENSIONS = {
    ".ppt",
    ".pptx",
    ".pptm",
    ".pps",
    ".ppsx",
    ".pot",
    ".potx",
}
PPT_PDF_FORMAT_POWERPOINT = 32  # PowerPoint ppSaveAsPDF


def ppt_output_pdf_path(source_path: Path) -> Path:
    today = datetime.now().strftime("%Y%m%d")
    return source_path.with_name(f"{source_path.stem}_rp_{today}.pdf")


def import_or_try_install(import_name: str, package_name: str):
    try:
        return importlib.import_module(import_name)
    except ImportError:
        if getattr(sys, "frozen", False):
            return None
        ok, output = run_pip_install([package_name])
        append_startup_log(output[-2500:])
        if ok:
            with suppress(Exception):
                return importlib.import_module(import_name)
        return None


def find_soffice() -> str | None:
    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return str(candidate)
    return None


def wait_for_pdf_file(path: Path, timeout_seconds: float = 12.0) -> bool:
    """Wait until a PDF exists and its size stops changing briefly."""
    deadline = time.time() + timeout_seconds
    last_size = -1
    stable_hits = 0
    while time.time() < deadline:
        try:
            if path.exists() and path.stat().st_size > 0:
                size = path.stat().st_size
                if size == last_size:
                    stable_hits += 1
                    if stable_hits >= 2:
                        return True
                else:
                    last_size = size
                    stable_hits = 0
        except Exception:
            pass
        time.sleep(0.25)
    with suppress(Exception):
        return path.exists() and path.stat().st_size > 0
    return False



def hidden_subprocess_options() -> dict:
    """Return subprocess options that keep helper converters out of foreground."""
    options = {}
    if os.name == "nt":
        options["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= getattr(subprocess, "STARTF_USESHOWWINDOW", 1)
            startupinfo.wShowWindow = getattr(subprocess, "SW_HIDE", 0)
            options["startupinfo"] = startupinfo
        except Exception:
            pass
    return options


def keep_powerpoint_in_background(powerpoint) -> None:
    """Best-effort PowerPoint COM settings to avoid foreground windows."""
    # Some PowerPoint versions do not allow Visible=False.  Keep every setting
    # best-effort so conversion does not fail merely because the UI cannot hide.
    with suppress(Exception):
        powerpoint.DisplayAlerts = 0
    with suppress(Exception):
        powerpoint.Visible = False
    with suppress(Exception):
        # 2 = ppWindowMinimized
        powerpoint.WindowState = 2


def prepare_ppt_output_path(output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.exists():
        try:
            output_path.unlink()
        except Exception as exc:
            raise RuntimeError(f"無法覆蓋既有 PDF，請先關閉檔案：{output_path}\n{exc}") from exc


def copy_ppt_to_safe_temp(source_path: Path, temp_dir: Path) -> Path:
    """Copy source PPT to an ASCII-only temporary filename for COM/LibreOffice."""
    safe_source = temp_dir / f"source{source_path.suffix.lower()}"
    shutil.copy2(source_path, safe_source)
    return safe_source


def move_temp_pdf_to_output(temp_pdf: Path, output_path: Path) -> Path:
    if not wait_for_pdf_file(temp_pdf):
        raise RuntimeError(f"暫存 PDF 未產生：{temp_pdf}")
    prepare_ppt_output_path(output_path)
    shutil.move(str(temp_pdf), str(output_path))
    if not wait_for_pdf_file(output_path):
        raise RuntimeError("PDF 已搬移但輸出檔不存在。")
    return output_path


def collect_pdf_candidates(folder: Path) -> list[Path]:
    try:
        return sorted(
            [p for p in folder.glob("*.pdf") if p.is_file()],
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
    except Exception:
        return []


def convert_ppt_with_powerpoint(source_path: Path, output_path: Path) -> Path:
    if platform.system().lower() != "windows":
        raise RuntimeError("PowerPoint COM 轉檔只支援 Windows。")

    win32com_client = import_or_try_install("win32com.client", "pywin32")
    pythoncom = import_or_try_install("pythoncom", "pywin32")
    if win32com_client is None or pythoncom is None:
        raise RuntimeError("找不到 pywin32，無法使用 Microsoft PowerPoint 轉檔。")

    prepare_ppt_output_path(output_path)
    powerpoint = None
    presentation = None
    pythoncom.CoInitialize()
    try:
        with tempfile.TemporaryDirectory(prefix="guppy_ppt_") as tmp:
            temp_dir = Path(tmp)
            temp_source = copy_ppt_to_safe_temp(source_path, temp_dir)
            temp_output = temp_dir / "source.pdf"

            # DispatchEx creates an isolated PowerPoint instance.  Keeping the
            # file in a plain temp path avoids many RPC/path issues with Chinese
            # filenames and cloud-synced folders.
            powerpoint = win32com_client.DispatchEx("PowerPoint.Application")
            keep_powerpoint_in_background(powerpoint)

            presentation = powerpoint.Presentations.Open(
                str(temp_source),
                WithWindow=False,
                ReadOnly=True,
                Untitled=False,
            )
            keep_powerpoint_in_background(powerpoint)

            export_errors = []
            try:
                # ppFixedFormatTypePDF = 2. ExportAsFixedFormat is more reliable
                # than SaveAs for some PowerPoint versions.
                presentation.ExportAsFixedFormat(str(temp_output), 2)
            except Exception as exc:
                export_errors.append(f"ExportAsFixedFormat：{exc}")
                try:
                    presentation.SaveAs(str(temp_output), PPT_PDF_FORMAT_POWERPOINT)
                except Exception as exc2:
                    export_errors.append(f"SaveAs：{exc2}")
                    raise RuntimeError("；".join(export_errors)) from exc2

            with suppress(Exception):
                presentation.Close()
            presentation = None
            with suppress(Exception):
                powerpoint.Quit()
            powerpoint = None

            if not wait_for_pdf_file(temp_output):
                raise RuntimeError("PowerPoint 已執行轉檔，但沒有產生 PDF。")
            return move_temp_pdf_to_output(temp_output, output_path)
    finally:
        with suppress(Exception):
            if presentation is not None:
                presentation.Close()
        with suppress(Exception):
            if powerpoint is not None:
                powerpoint.Quit()
        with suppress(Exception):
            pythoncom.CoUninitialize()


def convert_ppt_with_libreoffice(source_path: Path, output_path: Path) -> Path:
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("找不到 LibreOffice / soffice。")

    prepare_ppt_output_path(output_path)
    with tempfile.TemporaryDirectory(prefix="guppy_lo_") as tmp:
        temp_dir = Path(tmp)
        temp_source = copy_ppt_to_safe_temp(source_path, temp_dir)
        profile_dir = (temp_dir / "lo_profile").resolve()
        profile_dir.mkdir(parents=True, exist_ok=True)

        cmd = [
            soffice,
            f"-env:UserInstallation={profile_dir.as_uri()}",
            "--headless",
            "--invisible",
            "--nologo",
            "--nodefault",
            "--nofirststartwizard",
            "--nolockcheck",
            "--norestore",
            "--convert-to",
            "pdf:impress_pdf_Export",
            "--outdir",
            str(temp_dir),
            str(temp_source),
        ]
        completed = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
            cwd=str(temp_dir),
            timeout=180,
            **hidden_subprocess_options(),
        )

        stdout = (completed.stdout or "").strip()
        stderr = (completed.stderr or "").strip()
        if completed.returncode != 0:
            detail = stderr or stdout or "LibreOffice 轉檔失敗。"
            raise RuntimeError(detail)

        candidates = collect_pdf_candidates(temp_dir)
        preferred = temp_source.with_suffix(".pdf")
        if preferred.exists():
            return move_temp_pdf_to_output(preferred, output_path)
        if candidates:
            return move_temp_pdf_to_output(candidates[0], output_path)

        # Some LibreOffice builds ignore --outdir when already running.  Check the
        # original folder as a final fallback, then move the newest matching PDF.
        original_candidates = [
            source_path.with_suffix(".pdf"),
            *collect_pdf_candidates(source_path.parent)[:3],
        ]
        for candidate in original_candidates:
            if candidate.exists() and candidate.suffix.lower() == ".pdf":
                if candidate.resolve() == output_path.resolve():
                    return output_path
                try:
                    return move_temp_pdf_to_output(candidate, output_path)
                except Exception:
                    continue

        detail = "\n".join(part for part in (stdout, stderr) if part)
        if detail:
            raise RuntimeError("LibreOffice 已執行轉檔，但沒有找到產生的 PDF。\n" + detail[-1200:])
        raise RuntimeError("LibreOffice 已執行轉檔，但沒有找到產生的 PDF。")


def convert_powerpoint_to_pdf(source_file: str | Path) -> Path:
    source_path = Path(source_file).expanduser().resolve()
    if not source_path.exists():
        raise FileNotFoundError(f"找不到檔案：{source_path}")
    if source_path.suffix.lower() not in PPT_SUPPORTED_EXTENSIONS:
        raise ValueError(f"不支援的檔案格式：{source_path.suffix}")

    output_path = ppt_output_pdf_path(source_path)
    errors: list[str] = []

    try:
        return convert_ppt_with_powerpoint(source_path, output_path)
    except Exception as exc:
        errors.append(f"PowerPoint：{exc}")

    try:
        return convert_ppt_with_libreoffice(source_path, output_path)
    except Exception as exc:
        errors.append(f"LibreOffice：{exc}")

    raise RuntimeError("\n".join(errors))


class PPTConvertPanel:
    def __init__(self, root):
        self.root = root
        self.last_pdf_path: Path | None = None
        self.is_busy = False
        self.dnd_enabled = False
        self._build_ui()
        self.setup_drag_drop()
        self.root.after(250, self.setup_drag_drop)

    def _build_ui(self):
        self.main = tk.Frame(self.root, bg=BG)
        self.main.pack(fill="both", expand=True, padx=18, pady=18)
        self.main.grid_columnconfigure(0, weight=1)
        self.main.grid_rowconfigure(1, weight=1)

        header = self.card(self.main)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header.configure(padx=16, pady=14)

        tk.Label(
            header,
            text="PPT轉換",
            bg=CARD,
            fg=TEXT,
            font=("Microsoft JhengHei UI", 18, "bold"),
        ).pack(side="left")
        tk.Label(
            header,
            text="PowerPoint 拖進來，直接轉成 PDF",
            bg=CARD,
            fg=MUTED_TEXT,
            font=FONT,
        ).pack(side="left", padx=(12, 0))

        self.drop_card = self.card(self.main)
        self.drop_card.grid(row=1, column=0, sticky="nsew")
        self.drop_card.configure(padx=18, pady=18)
        self.drop_card.grid_columnconfigure(0, weight=1)
        self.drop_card.grid_rowconfigure(0, weight=1)

        self.drop_area = tk.Label(
            self.drop_card,
            text="把 PPT / PPTX 拖曳到這裡\n自動轉成 PDF",
            font=("Microsoft JhengHei UI", 24, "bold"),
            fg=PRIMARY,
            bg="#FFFFFF",
            relief="solid",
            bd=1,
            cursor="hand2",
        )
        self.drop_area.grid(row=0, column=0, sticky="nsew")
        self.drop_area.bind("<Button-1>", lambda _event: self.choose_files())

        controls = tk.Frame(self.main, bg=BG)
        controls.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        controls.grid_columnconfigure(1, weight=1)

        self.choose_btn = rounded_button(
            controls,
            "選擇 PPT 檔",
            self.choose_files,
            width=118,
            accent=True,
        )
        self.choose_btn.grid(row=0, column=0, sticky="w")

        self.status_var = tk.StringVar(
            value="拖曳 PPT/PPTX 檔案，或點選中間區域選取檔案。"
        )
        self.status_label = tk.Label(
            controls,
            textvariable=self.status_var,
            bg=BG,
            fg=MUTED_TEXT,
            font=FONT,
            anchor="w",
        )
        self.status_label.grid(row=0, column=1, sticky="ew", padx=12)

        # 30x30 圓形按鈕：開啟剛轉好的 PDF。
        self.open_pdf_dot = tk.Canvas(
            controls,
            width=30,
            height=30,
            bg=BG,
            highlightthickness=0,
            cursor="hand2",
        )
        self.open_pdf_dot.grid(row=0, column=2, sticky="e", padx=(8, 0))
        self.dot_item = self.open_pdf_dot.create_oval(
            2, 2, 28, 28, fill="#9AA0A6", outline="#9AA0A6"
        )
        self.open_pdf_dot.bind("<Button-1>", lambda _event: self.open_last_pdf())

    def card(self, parent):
        frame = tk.Frame(parent, bg=CARD, highlightthickness=1, highlightbackground=PREVIEW_BORDER)
        return frame

    def setup_drag_drop(self):
        if self.dnd_enabled:
            return
        if not load_tkinterdnd() or DND_FILES is None:
            self.set_status("拖曳套件 tkinterdnd2 未啟用；可點選中間區域選檔轉 PDF。")
            return

        for widget in dict.fromkeys((self.root, self.main, self.drop_card, self.drop_area)):
            try:
                if hasattr(widget, "drop_target_register") and hasattr(widget, "dnd_bind"):
                    widget.drop_target_register(DND_FILES)
                    widget.dnd_bind("<<Drop>>", self.on_drop)
                    self.dnd_enabled = True
            except Exception:
                pass
        if self.dnd_enabled:
            self.set_status("可拖曳 PPT/PPTX 檔案到此分頁轉成 PDF。")

    def set_status(self, text: str) -> None:
        self.status_var.set(text)

    def set_busy(self, busy: bool) -> None:
        self.is_busy = busy
        if busy:
            self.drop_area.config(text="轉檔中...", fg=PRIMARY)
            self.choose_btn.configure(state="disabled")
            self.open_pdf_dot.itemconfig(self.dot_item, fill="#DADCE0", outline="#DADCE0")
        else:
            self.drop_area.config(text="把 PPT / PPTX 拖曳到這裡\n自動轉成 PDF", fg=PRIMARY)
            self.choose_btn.configure(state="normal")
            color = PRIMARY if self.last_pdf_path else "#9AA0A6"
            self.open_pdf_dot.itemconfig(self.dot_item, fill=color, outline=color)

    def choose_files(self) -> None:
        if self.is_busy:
            return
        files = filedialog.askopenfilenames(
            title="選擇 PowerPoint 檔案",
            filetypes=[
                ("PowerPoint 檔案", "*.ppt *.pptx *.pptm *.pps *.ppsx *.pot *.potx"),
                ("全部檔案", "*.*"),
            ],
        )
        if files:
            self.start_conversion(list(files))

    def on_drop(self, event):
        if self.is_busy:
            return
        files = parse_dropped_files(event.data, self.root)
        self.start_conversion(files)

    def start_conversion(self, files: list[str]) -> None:
        valid_files: list[str] = []
        ignored: list[str] = []
        for file in files:
            path = Path(str(file).strip())
            if path.suffix.lower() in PPT_SUPPORTED_EXTENSIONS:
                valid_files.append(str(path))
            elif str(file).strip():
                ignored.append(path.name)

        if not valid_files:
            messagebox.showwarning("沒有可轉檔的 PPT", "請拖曳或選取 PowerPoint 檔案。")
            return

        if ignored:
            self.set_status(f"已略過非 PowerPoint 檔案：{', '.join(ignored[:3])}")
        else:
            self.set_status("開始轉檔...")

        self.set_busy(True)
        threading.Thread(
            target=self._conversion_worker,
            args=(valid_files,),
            name="GuppyPPTConvertWorker",
            daemon=True,
        ).start()

    def _conversion_worker(self, files: list[str]) -> None:
        success: list[Path] = []
        failed: list[str] = []
        for file in files:
            try:
                pdf_path = convert_powerpoint_to_pdf(file)
                success.append(pdf_path)
            except Exception as exc:
                failed.append(f"{Path(file).name}\n{exc}")
                self._write_error_log(file, exc)
        self.root.after(0, lambda: self._conversion_finished(success, failed))

    def _conversion_finished(self, success: list[Path], failed: list[str]) -> None:
        if success:
            self.last_pdf_path = success[-1]

        self.set_busy(False)

        if success and not failed:
            if len(success) == 1:
                self.set_status(f"完成：{success[0].name}")
            else:
                self.set_status(f"完成 {len(success)} 個 PDF，小圓點可開啟最後一個。")
        elif success and failed:
            self.set_status(f"完成 {len(success)} 個，失敗 {len(failed)} 個。")
            messagebox.showwarning("部分轉檔失敗", "\n\n".join(failed[:3]))
        else:
            self.set_status("轉檔失敗，請確認是否已安裝 PowerPoint 或 LibreOffice。")
            messagebox.showerror("轉檔失敗", "\n\n".join(failed[:3]))

    def _write_error_log(self, source_file: str, exc: Exception) -> None:
        try:
            log_path = Path(source_file).with_name(f"{PPT_APP_NAME}_error.log")
            with log_path.open("a", encoding="utf-8") as log_file:
                log_file.write("=" * 60 + "\n")
                log_file.write(datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
                log_file.write(f"Source: {source_file}\n")
                log_file.write(str(exc) + "\n")
                log_file.write(traceback.format_exc() + "\n")
        except Exception:
            pass

    def open_last_pdf(self) -> None:
        if not self.last_pdf_path or not self.last_pdf_path.exists():
            self.set_status("目前還沒有剛轉好的 PDF。")
            return
        try:
            if platform.system().lower() == "windows":
                os.startfile(str(self.last_pdf_path))  # type: ignore[attr-defined]
            elif platform.system().lower() == "darwin":
                subprocess.Popen(["open", str(self.last_pdf_path)])
            else:
                subprocess.Popen(["xdg-open", str(self.last_pdf_path)])
            self.set_status(f"已開啟：{self.last_pdf_path.name}")
        except Exception as exc:
            messagebox.showerror("無法開啟 PDF", str(exc))


# =========================================================
# DOC Convert Tab (V0.8.2)
# =========================================================
DOC_APP_NAME = "DOCToPDF"
DOC_SUPPORTED_EXTENSIONS = {".doc", ".docx", ".txt", ".odt"}
DOC_PDF_FORMAT_WORD = 17  # Word wdExportFormatPDF / wdFormatPDF
DOC_PREVIEW_PAGE_LIMIT = 30


def doc_output_pdf_path(source_path: Path) -> Path:
    today = datetime.now().strftime("%Y%m%d")
    return source_path.with_name(f"{source_path.stem}_rpDOC_{today}.pdf")


def copy_doc_to_safe_temp(source_path: Path, temp_dir: Path) -> Path:
    safe_source = temp_dir / f"source{source_path.suffix.lower()}"
    shutil.copy2(source_path, safe_source)
    return safe_source


def convert_doc_with_word(source_path: Path, output_path: Path) -> Path:
    if platform.system().lower() != "windows":
        raise RuntimeError("Microsoft Word COM 轉檔只支援 Windows。")

    win32com_client = import_or_try_install("win32com.client", "pywin32")
    pythoncom = import_or_try_install("pythoncom", "pywin32")
    if win32com_client is None or pythoncom is None:
        raise RuntimeError("找不到 pywin32，無法使用 Microsoft Word 轉檔。")

    prepare_ppt_output_path(output_path)
    word = None
    document = None
    pythoncom.CoInitialize()
    try:
        with tempfile.TemporaryDirectory(prefix="guppy_doc_") as tmp:
            temp_dir = Path(tmp)
            temp_source = copy_doc_to_safe_temp(source_path, temp_dir)
            temp_output = temp_dir / "source.pdf"

            word = win32com_client.DispatchEx("Word.Application")
            with suppress(Exception):
                word.Visible = False
            with suppress(Exception):
                word.DisplayAlerts = 0
            with suppress(Exception):
                word.ScreenUpdating = False

            document = word.Documents.Open(
                FileName=str(temp_source),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False,
            )

            export_errors = []
            try:
                document.ExportAsFixedFormat(
                    OutputFileName=str(temp_output),
                    ExportFormat=DOC_PDF_FORMAT_WORD,
                    OpenAfterExport=False,
                )
            except Exception as exc:
                export_errors.append(f"ExportAsFixedFormat：{exc}")
                try:
                    document.SaveAs2(
                        FileName=str(temp_output),
                        FileFormat=DOC_PDF_FORMAT_WORD,
                        AddToRecentFiles=False,
                    )
                except Exception as exc2:
                    export_errors.append(f"SaveAs2：{exc2}")
                    raise RuntimeError("；".join(export_errors)) from exc2

            with suppress(Exception):
                document.Close(False)
            document = None
            with suppress(Exception):
                word.Quit()
            word = None

            if not wait_for_pdf_file(temp_output):
                raise RuntimeError("Microsoft Word 已執行轉檔，但沒有產生 PDF。")
            return move_temp_pdf_to_output(temp_output, output_path)
    finally:
        with suppress(Exception):
            if document is not None:
                document.Close(False)
        with suppress(Exception):
            if word is not None:
                word.Quit()
        with suppress(Exception):
            pythoncom.CoUninitialize()


def convert_doc_with_libreoffice(source_path: Path, output_path: Path) -> Path:
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("找不到 LibreOffice / soffice。")

    prepare_ppt_output_path(output_path)
    with tempfile.TemporaryDirectory(prefix="guppy_doc_lo_") as tmp:
        temp_dir = Path(tmp)
        temp_source = copy_doc_to_safe_temp(source_path, temp_dir)
        profile_dir = (temp_dir / "lo_profile").resolve()
        profile_dir.mkdir(parents=True, exist_ok=True)

        cmd = [
            soffice,
            f"-env:UserInstallation={profile_dir.as_uri()}",
            "--headless",
            "--invisible",
            "--nologo",
            "--nodefault",
            "--nofirststartwizard",
            "--nolockcheck",
            "--norestore",
            "--convert-to",
            "pdf:writer_pdf_Export",
            "--outdir",
            str(temp_dir),
            str(temp_source),
        ]
        completed = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
            cwd=str(temp_dir),
            timeout=180,
            **hidden_subprocess_options(),
        )

        stdout = (completed.stdout or "").strip()
        stderr = (completed.stderr or "").strip()
        if completed.returncode != 0:
            raise RuntimeError(stderr or stdout or "LibreOffice 文件轉檔失敗。")

        preferred = temp_source.with_suffix(".pdf")
        candidates = collect_pdf_candidates(temp_dir)
        if preferred.exists():
            return move_temp_pdf_to_output(preferred, output_path)
        if candidates:
            return move_temp_pdf_to_output(candidates[0], output_path)

        original_candidates = [
            source_path.with_suffix(".pdf"),
            *collect_pdf_candidates(source_path.parent)[:3],
        ]
        for candidate in original_candidates:
            if candidate.exists() and candidate.suffix.lower() == ".pdf":
                if candidate.resolve() == output_path.resolve():
                    return output_path
                try:
                    return move_temp_pdf_to_output(candidate, output_path)
                except Exception:
                    continue

        raise RuntimeError(
            "LibreOffice 已執行轉檔，但沒有找到輸出 PDF。"
            + (f"\n{stdout}" if stdout else "")
            + (f"\n{stderr}" if stderr else "")
        )


def read_text_file(source_path: Path) -> str:
    raw = source_path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "cp950", "big5", "gb18030"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def wrap_text_for_pdf(text: str, max_units: int = 88) -> list[str]:
    """Wrap text using approximate East-Asian display width."""
    import unicodedata

    wrapped: list[str] = []
    for source_line in text.expandtabs(4).splitlines():
        if not source_line:
            wrapped.append("")
            continue
        current: list[str] = []
        units = 0
        for char in source_line:
            char_units = 2 if unicodedata.east_asian_width(char) in ("W", "F", "A") else 1
            if current and units + char_units > max_units:
                wrapped.append("".join(current))
                current = [char]
                units = char_units
            else:
                current.append(char)
                units += char_units
        wrapped.append("".join(current))
    return wrapped or [""]


def convert_txt_to_pdf(source_path: Path, output_path: Path) -> Path:
    prepare_ppt_output_path(output_path)
    text = read_text_file(source_path)
    lines = wrap_text_for_pdf(text)

    pdf = fitz.open()
    try:
        page_width, page_height = 595.0, 842.0  # A4 points
        margin_x, margin_y = 42.0, 42.0
        font_size = 11.0
        line_height = 16.0
        lines_per_page = max(1, int((page_height - margin_y * 2) // line_height))

        for start in range(0, len(lines), lines_per_page):
            page_lines = lines[start : start + lines_per_page]
            page = pdf.new_page(width=page_width, height=page_height)
            page.insert_textbox(
                fitz.Rect(
                    margin_x,
                    margin_y,
                    page_width - margin_x,
                    page_height - margin_y,
                ),
                "\n".join(page_lines),
                fontname="china-t",
                fontsize=font_size,
                lineheight=line_height / font_size,
                color=(0, 0, 0),
                align=0,
            )

        if pdf.page_count == 0:
            pdf.new_page(width=page_width, height=page_height)
        pdf.save(str(output_path), garbage=4, deflate=True)
    finally:
        pdf.close()

    if not wait_for_pdf_file(output_path):
        raise RuntimeError("TXT 已轉換，但沒有產生 PDF。")
    return output_path


def convert_document_to_pdf(source_file: str | Path) -> Path:
    source_path = Path(source_file).expanduser().resolve()
    if not source_path.exists():
        raise FileNotFoundError(f"找不到文件：{source_path}")

    suffix = source_path.suffix.lower()
    if suffix not in DOC_SUPPORTED_EXTENSIONS:
        raise ValueError(f"不支援的文件格式：{source_path.suffix}")

    output_path = doc_output_pdf_path(source_path)

    if suffix == ".txt":
        return convert_txt_to_pdf(source_path, output_path)

    # ODT 為 OpenDocument 格式，固定使用 LibreOffice，避免 Word 轉檔造成版面差異。
    if suffix == ".odt":
        try:
            return convert_doc_with_libreoffice(source_path, output_path)
        except Exception as exc:
            raise RuntimeError(f"ODT 轉檔需要 LibreOffice：\n{exc}") from exc

    errors: list[str] = []
    if platform.system().lower() == "windows":
        try:
            return convert_doc_with_word(source_path, output_path)
        except Exception as exc:
            errors.append(f"Microsoft Word：{exc}")

    try:
        return convert_doc_with_libreoffice(source_path, output_path)
    except Exception as exc:
        errors.append(f"LibreOffice：{exc}")

    raise RuntimeError("\n".join(errors))


# =========================================================
# TIFF Convert (V0.9.4)
# =========================================================
TIFF_SUPPORTED_EXTENSIONS = {".tif", ".tiff"}


def tiff_output_pdf_path(source_path: Path) -> Path:
    today = datetime.now().strftime("%Y%m%d")
    return source_path.with_name(f"{source_path.stem}_rpTIFF_{today}.pdf")


def normalize_image_dpi(value, default=300.0) -> float:
    """將 TIFF DPI 正規化，避免缺失或異常 DPI 產生極端 PDF 頁面尺寸。"""
    try:
        dpi = float(value)
    except (TypeError, ValueError, OverflowError):
        return default
    if dpi < 10.0 or dpi > 2400.0:
        return default
    return dpi


def tiff_frame_to_rgb(frame):
    """將 TIFF 頁面轉成白底 RGB，兼容透明、調色盤、灰階及 1-bit TIFF。"""
    if frame.mode in ("RGBA", "LA"):
        rgba = frame.convert("RGBA")
        background = Image.new("RGB", rgba.size, "white")
        background.paste(rgba, mask=rgba.getchannel("A"))
        return background

    if frame.mode == "P" and "transparency" in frame.info:
        rgba = frame.convert("RGBA")
        background = Image.new("RGB", rgba.size, "white")
        background.paste(rgba, mask=rgba.getchannel("A"))
        return background

    return frame.convert("RGB")


def convert_tiff_to_pdf(source_file: str | Path) -> Path:
    """將單頁或多頁 TIF/TIFF 轉為 PDF；每個 TIFF frame 對應一個 PDF 頁面。"""
    source_path = Path(source_file).expanduser().resolve()
    if not source_path.exists():
        raise FileNotFoundError(f"找不到 TIFF 檔案：{source_path}")
    if source_path.suffix.lower() not in TIFF_SUPPORTED_EXTENSIONS:
        raise ValueError(f"不支援的 TIFF 格式：{source_path.suffix}")

    output_path = tiff_output_pdf_path(source_path)
    prepare_ppt_output_path(output_path)

    source_image = None
    pdf = fitz.open()
    frame_count = 0

    try:
        source_image = Image.open(str(source_path))
        frame_count = max(1, int(getattr(source_image, "n_frames", 1)))

        for frame_index in range(frame_count):
            source_image.seek(frame_index)
            frame = source_image.copy()

            dpi_info = frame.info.get("dpi") or source_image.info.get("dpi") or (300, 300)
            if isinstance(dpi_info, (tuple, list)):
                dpi_x = normalize_image_dpi(dpi_info[0] if dpi_info else 300)
                dpi_y = normalize_image_dpi(
                    dpi_info[1] if len(dpi_info) > 1 else dpi_x
                )
            else:
                dpi_x = dpi_y = normalize_image_dpi(dpi_info)

            rgb = tiff_frame_to_rgb(frame)
            width_points = max(1.0, rgb.width * 72.0 / dpi_x)
            height_points = max(1.0, rgb.height * 72.0 / dpi_y)

            buffer = io.BytesIO()
            try:
                # PNG 為無損格式，適合掃描文字及線稿。
                rgb.save(buffer, format="PNG", optimize=False)
                page = pdf.new_page(width=width_points, height=height_points)
                page.insert_image(
                    page.rect,
                    stream=buffer.getvalue(),
                    keep_proportion=True,
                )
            finally:
                buffer.close()
                with suppress(Exception):
                    rgb.close()
                with suppress(Exception):
                    frame.close()

        if pdf.page_count == 0:
            raise RuntimeError("TIFF 中沒有可轉換的影像頁面。")

        pdf.save(str(output_path), garbage=4, deflate=True)
    except Exception:
        with suppress(Exception):
            if output_path.exists():
                output_path.unlink()
        raise
    finally:
        pdf.close()
        if source_image is not None:
            with suppress(Exception):
                source_image.close()

    if not wait_for_pdf_file(output_path):
        raise RuntimeError("TIFF 已轉換，但沒有產生 PDF。")

    # 確認輸出頁數，避免多頁 TIFF 靜默漏頁。
    verify = fitz.open(str(output_path))
    try:
        if verify.page_count != frame_count:
            raise RuntimeError(
                f"TIFF 原有 {frame_count} 頁，但輸出 PDF 為 {verify.page_count} 頁。"
            )
    finally:
        verify.close()

    return output_path


class UniversalConvertPanel:
    def __init__(self, root):
        self.root = root
        self.last_pdf_path: Path | None = None
        self.is_busy = False
        self.dnd_enabled = False
        self.preview_images: list[object] = []
        self._build_ui()
        self.setup_drag_drop()
        self.root.after(250, self.setup_drag_drop)

    def card(self, parent):
        return tk.Frame(
            parent,
            bg=CARD,
            highlightthickness=1,
            highlightbackground=PREVIEW_BORDER,
        )

    def _build_ui(self):
        self.main = tk.Frame(self.root, bg=BG)
        self.main.pack(fill="both", expand=True, padx=18, pady=18)
        self.main.grid_columnconfigure(0, weight=1)
        self.main.grid_rowconfigure(1, weight=1)

        header = self.card(self.main)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header.configure(padx=16, pady=14)
        tk.Label(
            header,
            text="轉換氣體",
            bg=CARD,
            fg=TEXT,
            font=("Microsoft JhengHei UI", 18, "bold"),
        ).pack(side="left")
        tk.Label(
            header,
            text="自動辨識 PPT、Word、TXT、ODT、TIFF 副檔名並轉成 PDF",
            bg=CARD,
            fg=MUTED_TEXT,
            font=FONT,
        ).pack(side="left", padx=(12, 0))

        split = tk.Frame(self.main, bg=BG)
        split.grid(row=1, column=0, sticky="nsew")
        split.grid_rowconfigure(0, weight=1)
        split.grid_columnconfigure(0, weight=1, uniform="docsplit")
        split.grid_columnconfigure(1, weight=1, uniform="docsplit")

        self.drop_card = self.card(split)
        self.drop_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self.drop_card.configure(padx=16, pady=16)
        self.drop_card.grid_columnconfigure(0, weight=1)
        self.drop_card.grid_rowconfigure(0, weight=1)

        self.drop_area = tk.Label(
            self.drop_card,
            text="把 PPT / PPTX / DOC / DOCX / TXT / ODT / TIFF 拖曳到這裡\n自動辨識格式並轉成 PDF",
            font=("Microsoft JhengHei UI", 22, "bold"),
            fg=PRIMARY,
            bg="#FFFFFF",
            relief="solid",
            bd=1,
            cursor="hand2",
        )
        self.drop_area.grid(row=0, column=0, sticky="nsew")
        self.drop_area.bind("<Button-1>", lambda _event: self.choose_files())

        self.preview_card = self.card(split)
        self.preview_card.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        self.preview_card.grid_rowconfigure(1, weight=1)
        self.preview_card.grid_columnconfigure(0, weight=1)

        self.preview_title_var = tk.StringVar(value="轉檔後 PDF 預覽")
        tk.Label(
            self.preview_card,
            textvariable=self.preview_title_var,
            bg=CARD,
            fg=TEXT,
            font=TITLE_FONT,
            anchor="w",
        ).grid(row=0, column=0, columnspan=2, sticky="ew", padx=12, pady=(10, 6))

        preview_shell = tk.Frame(self.preview_card, bg=CARD)
        preview_shell.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        preview_shell.grid_rowconfigure(0, weight=1)
        preview_shell.grid_columnconfigure(0, weight=1)

        self.preview_canvas = tk.Canvas(
            preview_shell,
            bg="#F5F7FA",
            highlightthickness=0,
        )
        self.preview_vbar = ttk.Scrollbar(
            preview_shell, orient="vertical", command=self.preview_canvas.yview
        )
        self.preview_hbar = ttk.Scrollbar(
            preview_shell, orient="horizontal", command=self.preview_canvas.xview
        )
        self.preview_canvas.configure(
            yscrollcommand=self.preview_vbar.set,
            xscrollcommand=self.preview_hbar.set,
        )
        self.preview_canvas.grid(row=0, column=0, sticky="nsew")
        self.preview_vbar.grid(row=0, column=1, sticky="ns")
        self.preview_hbar.grid(row=1, column=0, sticky="ew")

        self.preview_inner = tk.Frame(self.preview_canvas, bg="#F5F7FA")
        self.preview_window = self.preview_canvas.create_window(
            (0, 0), window=self.preview_inner, anchor="nw"
        )
        self.preview_inner.bind("<Configure>", self._update_preview_scrollregion)
        self.preview_canvas.bind("<Configure>", self._resize_preview_inner)
        self.preview_canvas.bind_all("<MouseWheel>", self._preview_mousewheel, add="+")
        self.clear_preview("轉檔完成後，右側會顯示 PDF 預覽。")

        controls = tk.Frame(self.main, bg=BG)
        controls.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        controls.grid_columnconfigure(1, weight=1)

        self.choose_btn = rounded_button(
            controls,
            "選擇文件",
            self.choose_files,
            width=120,
            accent=True,
        )
        self.choose_btn.grid(row=0, column=0, sticky="w")

        self.status_var = tk.StringVar(
            value="拖曳 PPT、PPTX、DOC、DOCX、TXT、ODT、TIF、TIFF，或點選左側區域選取文件。"
        )
        self.status_label = tk.Label(
            controls,
            textvariable=self.status_var,
            bg=BG,
            fg=MUTED_TEXT,
            font=FONT,
            anchor="w",
        )
        self.status_label.grid(row=0, column=1, sticky="ew", padx=12)

        self.open_pdf_dot = tk.Canvas(
            controls,
            width=30,
            height=30,
            bg=BG,
            highlightthickness=0,
            cursor="hand2",
        )
        self.open_pdf_dot.grid(row=0, column=2, sticky="e", padx=(8, 0))
        self.dot_item = self.open_pdf_dot.create_oval(
            2, 2, 28, 28, fill="#9AA0A6", outline="#9AA0A6"
        )
        self.open_pdf_dot.bind("<Button-1>", lambda _event: self.open_last_pdf())

    def _update_preview_scrollregion(self, _event=None):
        self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))

    def _resize_preview_inner(self, event):
        desired = max(event.width, self.preview_inner.winfo_reqwidth())
        self.preview_canvas.itemconfigure(self.preview_window, width=desired)

    def _preview_mousewheel(self, event):
        try:
            if self.root.winfo_containing(event.x_root, event.y_root) in (
                self.preview_canvas,
                self.preview_inner,
            ) or self._is_preview_child(self.root.winfo_containing(event.x_root, event.y_root)):
                self.preview_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass

    def _is_preview_child(self, widget):
        while widget is not None:
            if widget is self.preview_inner:
                return True
            widget = getattr(widget, "master", None)
        return False

    def clear_preview(self, message="尚未產生 PDF 預覽。"):
        for child in self.preview_inner.winfo_children():
            child.destroy()
        self.preview_images.clear()
        self.preview_title_var.set("轉檔後 PDF 預覽")
        tk.Label(
            self.preview_inner,
            text=message,
            bg="#F5F7FA",
            fg=MUTED_TEXT,
            font=FONT,
            justify="center",
        ).pack(fill="both", expand=True, padx=20, pady=40)
        self.preview_canvas.yview_moveto(0)
        self.preview_canvas.xview_moveto(0)

    def load_pdf_preview(self, pdf_path: Path):
        self.clear_preview("正在載入 PDF 預覽...")
        try:
            document = fitz.open(str(pdf_path))
            try:
                for child in self.preview_inner.winfo_children():
                    child.destroy()
                self.preview_images.clear()
                page_count = document.page_count
                show_count = min(page_count, DOC_PREVIEW_PAGE_LIMIT)
                self.preview_title_var.set(
                    f"轉檔後預覽：{pdf_path.name}（{page_count} 頁）"
                )

                available_width = max(320, self.preview_canvas.winfo_width() - 36)
                for page_index in range(show_count):
                    page = document.load_page(page_index)
                    pix = page.get_pixmap(matrix=fitz.Matrix(1.0, 1.0), alpha=False)
                    image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                    if image.width > available_width:
                        ratio = available_width / image.width
                        image = image.resize(
                            (available_width, max(1, int(image.height * ratio))),
                            Image.Resampling.LANCZOS,
                        )
                    photo = ImageTk.PhotoImage(image)
                    self.preview_images.append(photo)

                    card = tk.Frame(
                        self.preview_inner,
                        bg="#FFFFFF",
                        highlightthickness=1,
                        highlightbackground=PREVIEW_BORDER,
                    )
                    card.pack(padx=12, pady=8, anchor="n")
                    tk.Label(
                        card,
                        text=f"第 {page_index + 1} 頁",
                        bg="#FFFFFF",
                        fg=MUTED_TEXT,
                        font=FONT,
                    ).pack(anchor="w", padx=8, pady=(6, 2))
                    tk.Label(card, image=photo, bg="#FFFFFF").pack(padx=8, pady=(0, 8))
                    self.root.update_idletasks()

                if page_count > show_count:
                    tk.Label(
                        self.preview_inner,
                        text=f"為避免記憶體占用，僅預覽前 {show_count} 頁。",
                        bg="#F5F7FA",
                        fg=MUTED_TEXT,
                        font=FONT,
                    ).pack(pady=12)
            finally:
                document.close()
            self.preview_canvas.yview_moveto(0)
            self._update_preview_scrollregion()
        except Exception as exc:
            self.clear_preview("PDF 預覽失敗。")
            self.set_status(f"PDF 已轉出，但預覽失敗：{exc}")

    def setup_drag_drop(self):
        if self.dnd_enabled:
            return
        if not load_tkinterdnd() or DND_FILES is None:
            self.set_status("拖曳套件 tkinterdnd2 未啟用；可點選左側區域選檔轉 PDF。")
            return

        widgets = (
            self.root,
            self.main,
            self.drop_card,
            self.drop_area,
            self.preview_card,
            self.preview_canvas,
        )
        for widget in dict.fromkeys(widgets):
            try:
                if hasattr(widget, "drop_target_register") and hasattr(widget, "dnd_bind"):
                    widget.drop_target_register(DND_FILES)
                    widget.dnd_bind("<<Drop>>", self.on_drop)
                    self.dnd_enabled = True
            except Exception:
                pass
        if self.dnd_enabled:
            self.set_status("可拖曳 PPT、PPTX、DOC、DOCX、TXT、ODT、TIF、TIFF 到此分頁轉成 PDF。")

    def set_status(self, text: str):
        self.status_var.set(text)

    def set_busy(self, busy: bool):
        self.is_busy = busy
        if busy:
            self.drop_area.config(text="檔案轉 PDF 中...", fg=PRIMARY)
            self.choose_btn.configure(state="disabled")
            self.open_pdf_dot.itemconfig(
                self.dot_item, fill="#DADCE0", outline="#DADCE0"
            )
        else:
            self.drop_area.config(
                text="把 PPT / PPTX / DOC / DOCX / TXT / ODT / TIFF 拖曳到這裡\n自動辨識格式並轉成 PDF",
                fg=PRIMARY,
            )
            self.choose_btn.configure(state="normal")
            color = PRIMARY if self.last_pdf_path else "#9AA0A6"
            self.open_pdf_dot.itemconfig(self.dot_item, fill=color, outline=color)

    def choose_files(self):
        if self.is_busy:
            return
        files = filedialog.askopenfilenames(
            title="選擇要轉成 PDF 的文件",
            filetypes=[
                (
                    "可轉換文件",
                    "*.ppt *.pptx *.pptm *.pps *.ppsx *.pot *.potx *.doc *.docx *.txt *.odt *.tif *.tiff",
                ),
                ("PowerPoint", "*.ppt *.pptx *.pptm *.pps *.ppsx *.pot *.potx"),
                ("Word / 文字文件", "*.doc *.docx *.txt"),
                ("OpenDocument 文字文件", "*.odt"),
                ("TIFF 影像", "*.tif *.tiff"),
                ("全部檔案", "*.*"),
            ],
        )
        if files:
            self.start_conversion(list(files))

    def on_drop(self, event):
        if self.is_busy:
            return
        self.start_conversion(parse_dropped_files(event.data, self.root))

    def start_conversion(self, files: list[str]):
        supported = (
            PPT_SUPPORTED_EXTENSIONS
            | DOC_SUPPORTED_EXTENSIONS
            | TIFF_SUPPORTED_EXTENSIONS
        )
        valid_files: list[str] = []
        ignored: list[str] = []
        for file in files:
            path = Path(str(file).strip())
            if path.suffix.lower() in supported:
                valid_files.append(str(path))
            elif str(file).strip():
                ignored.append(path.name)

        if not valid_files:
            messagebox.showwarning(
                "沒有可轉檔的文件",
                "請拖曳或選取 PPT、PPTX、DOC、DOCX、TXT、ODT、TIF、TIFF 檔案。",
            )
            return

        self.clear_preview("文件轉檔中，完成後會顯示預覽。")
        if ignored:
            self.set_status(f"已略過不支援的檔案：{', '.join(ignored[:3])}")
        else:
            self.set_status("正在辨識副檔名並開始轉成 PDF...")
        self.set_busy(True)
        threading.Thread(
            target=self._conversion_worker,
            args=(valid_files,),
            name="GuppyUniversalConvertWorker",
            daemon=True,
        ).start()

    def _conversion_worker(self, files: list[str]):
        success: list[Path] = []
        failed: list[str] = []
        for file in files:
            try:
                source = Path(file)
                suffix = source.suffix.lower()
                if suffix in PPT_SUPPORTED_EXTENSIONS:
                    output = convert_powerpoint_to_pdf(source)
                elif suffix in DOC_SUPPORTED_EXTENSIONS:
                    output = convert_document_to_pdf(source)
                elif suffix in TIFF_SUPPORTED_EXTENSIONS:
                    output = convert_tiff_to_pdf(source)
                else:
                    raise ValueError(f"不支援的檔案格式：{suffix or '無副檔名'}")
                success.append(Path(output))
            except Exception as exc:
                failed.append(f"{Path(file).name}\n{exc}")
                self._write_error_log(file, exc)
        self.root.after(0, lambda: self._conversion_finished(success, failed))

    def _conversion_finished(self, success: list[Path], failed: list[str]):
        if success:
            self.last_pdf_path = success[-1]
        self.set_busy(False)

        if success:
            self.load_pdf_preview(success[-1])

        if success and not failed:
            if len(success) == 1:
                self.set_status(f"完成：{success[0].name}")
            else:
                self.set_status(f"完成 {len(success)} 個 PDF，右側預覽最後一個。")
        elif success and failed:
            self.set_status(f"完成 {len(success)} 個，失敗 {len(failed)} 個。")
            messagebox.showwarning("部分轉檔失敗", "\n\n".join(failed[:3]))
        else:
            self.clear_preview("轉檔失敗，沒有可預覽的 PDF。")
            self.set_status(
                "轉檔失敗；Office/ODT 文件請確認已安裝 PowerPoint、Word 或 LibreOffice；TIFF 請確認 Pillow 與 PyMuPDF 可用。"
            )
            messagebox.showerror("轉檔失敗", "\n\n".join(failed[:3]))

    def _write_error_log(self, source_file: str, exc: Exception):
        try:
            log_path = Path(source_file).with_name("ConvertGas_error.log")
            with log_path.open("a", encoding="utf-8") as log_file:
                log_file.write("=" * 60 + "\n")
                log_file.write(datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
                log_file.write(f"Source: {source_file}\n")
                log_file.write(str(exc) + "\n")
                log_file.write(traceback.format_exc() + "\n")
        except Exception:
            pass

    def open_last_pdf(self):
        if not self.last_pdf_path or not self.last_pdf_path.exists():
            self.set_status("目前還沒有剛轉好的 PDF。")
            return
        try:
            if platform.system().lower() == "windows":
                os.startfile(str(self.last_pdf_path))  # type: ignore[attr-defined]
            elif platform.system().lower() == "darwin":
                subprocess.Popen(["open", str(self.last_pdf_path)])
            else:
                subprocess.Popen(["xdg-open", str(self.last_pdf_path)])
            self.set_status(f"已開啟：{self.last_pdf_path.name}")
        except Exception as exc:
            messagebox.showerror("無法開啟 PDF", str(exc))


# =========================================================
# PPT 4-up Convert Tab (merged from PPToPDFoT4 V0.6)
# =========================================================
PPT4_APP_NAME = "PPToPDFoT4"


def ppt4_output_pdf_path(source_path: Path) -> Path:
    today = datetime.now().strftime("%Y%m%d")
    return source_path.with_name(f"{source_path.stem}_rpT4_{today}.pdf")


def unique_pdf_path(path: Path) -> Path:
    path = Path(path)
    if not path.exists():
        return path
    for i in range(1, 1000):
        candidate = path.with_name(f"{path.stem}_{i:02d}{path.suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"無法產生不重複檔名：{path}")


def convert_powerpoint_to_pdf_at(source_file: str | Path, output_path: str | Path) -> Path:
    source_path = Path(source_file).expanduser().resolve()
    output_path = Path(output_path).expanduser().resolve()
    if not source_path.exists():
        raise FileNotFoundError(f"找不到檔案：{source_path}")
    if source_path.suffix.lower() not in PPT_SUPPORTED_EXTENSIONS:
        raise ValueError(f"不支援的檔案格式：{source_path.suffix}")

    errors: list[str] = []
    try:
        return convert_ppt_with_powerpoint(source_path, output_path)
    except Exception as exc:
        errors.append(f"PowerPoint：{exc}")

    try:
        return convert_ppt_with_libreoffice(source_path, output_path)
    except Exception as exc:
        errors.append(f"LibreOffice：{exc}")

    raise RuntimeError("\n".join(errors))


def layout_pdf_as_2x2(source_pdf: str | Path, output_pdf: str | Path) -> Path:
    """Layout a normal PDF into 2x2 handout pages with 2pt gaps and black borders."""
    if not ensure_packages():
        raise RuntimeError("缺少 PyMuPDF / Pillow，無法進行 2x2 PDF 排版。")

    source_pdf = Path(source_pdf).resolve()
    output_pdf = Path(output_pdf).resolve()
    prepare_ppt_output_path(output_pdf)

    src_doc = None
    out_doc = None
    try:
        src_doc = fitz.open(str(source_pdf))
        if src_doc.page_count <= 0:
            raise RuntimeError("暫存 PDF 沒有頁面，無法排版。")

        out_doc = fitz.open()
        first_rect = src_doc[0].rect
        src_w = float(first_rect.width)
        src_h = float(first_rect.height)

        margin = 24.0
        gap = 2.0
        out_w = 842.0
        out_h = 595.0
        cell_w = (out_w - margin * 2 - gap) / 2
        cell_h = (out_h - margin * 2 - gap) / 2

        for base in range(0, src_doc.page_count, 4):
            page = out_doc.new_page(width=out_w, height=out_h)
            for offset in range(4):
                src_index = base + offset
                if src_index >= src_doc.page_count:
                    break
                row = offset // 2
                col = offset % 2
                x0 = margin + col * (cell_w + gap)
                y0 = margin + row * (cell_h + gap)
                x1 = x0 + cell_w
                y1 = y0 + cell_h

                scale = min(cell_w / src_w, cell_h / src_h)
                draw_w = src_w * scale
                draw_h = src_h * scale
                dx = (cell_w - draw_w) / 2
                dy = (cell_h - draw_h) / 2
                rect = fitz.Rect(x0 + dx, y0 + dy, x1 - dx, y1 - dy)

                page.show_pdf_page(rect, src_doc, src_index)
                with suppress(Exception):
                    page.draw_rect(rect, color=(0, 0, 0), width=0.9, fill=None, overlay=True)

        out_doc.save(str(output_pdf), garbage=4, deflate=True, clean=True)
        out_doc.close()
        out_doc = None
        if not wait_for_pdf_file(output_pdf, timeout_seconds=12.0):
            raise RuntimeError("PDF 2x2 排版失敗，沒有產生有效輸出檔。")
        return output_pdf
    finally:
        with suppress(Exception):
            if out_doc is not None:
                out_doc.close()
        with suppress(Exception):
            if src_doc is not None:
                src_doc.close()


def convert_powerpoint_to_4up_pdf(source_file: str | Path) -> Path:
    source_path = Path(source_file).expanduser().resolve()
    if not source_path.exists():
        raise FileNotFoundError(f"找不到檔案：{source_path}")
    if source_path.suffix.lower() not in PPT_SUPPORTED_EXTENSIONS:
        raise ValueError(f"不支援的檔案格式：{source_path.suffix}")

    output_path = unique_pdf_path(ppt4_output_pdf_path(source_path)).resolve()
    with tempfile.TemporaryDirectory(prefix="guppy_ppt4_") as tmp:
        temp_dir = Path(tmp)
        plain_pdf = temp_dir / "plain.pdf"
        convert_powerpoint_to_pdf_at(source_path, plain_pdf)
        return layout_pdf_as_2x2(plain_pdf, output_path)


class PPT4ConvertPanel(PPTConvertPanel):
    def _build_ui(self):
        self.main = tk.Frame(self.root, bg=BG)
        self.main.pack(fill="both", expand=True, padx=18, pady=18)
        self.main.grid_columnconfigure(0, weight=1)
        self.main.grid_rowconfigure(1, weight=1)

        header = self.card(self.main)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header.configure(padx=16, pady=14)

        tk.Label(
            header,
            text="PTT轉4格",
            bg=CARD,
            fg=TEXT,
            font=("Microsoft JhengHei UI", 18, "bold"),
        ).pack(side="left")
        tk.Label(
            header,
            text="PowerPoint 先轉 PDF，再排成 2×2 四格講義",
            bg=CARD,
            fg=MUTED_TEXT,
            font=FONT,
        ).pack(side="left", padx=(12, 0))

        self.drop_card = self.card(self.main)
        self.drop_card.grid(row=1, column=0, sticky="nsew")
        self.drop_card.configure(padx=18, pady=18)
        self.drop_card.grid_columnconfigure(0, weight=1)
        self.drop_card.grid_rowconfigure(0, weight=1)

        self.drop_area = tk.Label(
            self.drop_card,
            text="把 PPT / PPTX 拖曳到這裡\n自動轉成 2×2 四格 PDF",
            font=("Microsoft JhengHei UI", 24, "bold"),
            fg=PRIMARY,
            bg="#FFFFFF",
            relief="solid",
            bd=1,
            cursor="hand2",
        )
        self.drop_area.grid(row=0, column=0, sticky="nsew")
        self.drop_area.bind("<Button-1>", lambda _event: self.choose_files())

        controls = tk.Frame(self.main, bg=BG)
        controls.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        controls.grid_columnconfigure(1, weight=1)

        self.choose_btn = rounded_button(
            controls,
            "選擇 PPT 檔",
            self.choose_files,
            width=118,
            accent=True,
        )
        self.choose_btn.grid(row=0, column=0, sticky="w")

        self.status_var = tk.StringVar(
            value="拖曳 PPT/PPTX 檔案，或點選中間區域選取檔案。"
        )
        self.status_label = tk.Label(
            controls,
            textvariable=self.status_var,
            bg=BG,
            fg=MUTED_TEXT,
            font=FONT,
            anchor="w",
        )
        self.status_label.grid(row=0, column=1, sticky="ew", padx=12)

        self.open_pdf_dot = tk.Canvas(
            controls,
            width=30,
            height=30,
            bg=BG,
            highlightthickness=0,
            cursor="hand2",
        )
        self.open_pdf_dot.grid(row=0, column=2, sticky="e", padx=(8, 0))
        self.dot_item = self.open_pdf_dot.create_oval(
            2, 2, 28, 28, fill="#9AA0A6", outline="#9AA0A6"
        )
        self.open_pdf_dot.bind("<Button-1>", lambda _event: self.open_last_pdf())

    def setup_drag_drop(self):
        super().setup_drag_drop()
        if self.dnd_enabled:
            self.set_status("可拖曳 PPT/PPTX 檔案到此分頁轉成 2×2 四格 PDF。")
        else:
            self.set_status("拖曳套件 tkinterdnd2 未啟用；可點選中間區域選檔轉 2×2 四格 PDF。")

    def set_busy(self, busy: bool) -> None:
        self.is_busy = busy
        if busy:
            self.drop_area.config(text="轉 2×2 四格 PDF 中...", fg=PRIMARY)
            self.choose_btn.configure(state="disabled")
            self.open_pdf_dot.itemconfig(self.dot_item, fill="#DADCE0", outline="#DADCE0")
        else:
            self.drop_area.config(text="把 PPT / PPTX 拖曳到這裡\n自動轉成 2×2 四格 PDF", fg=PRIMARY)
            self.choose_btn.configure(state="normal")
            color = PRIMARY if self.last_pdf_path else "#9AA0A6"
            self.open_pdf_dot.itemconfig(self.dot_item, fill=color, outline=color)

    def _conversion_worker(self, files: list[str]) -> None:
        success: list[Path] = []
        failed: list[str] = []
        for file in files:
            try:
                pdf_path = convert_powerpoint_to_4up_pdf(file)
                success.append(pdf_path)
            except Exception as exc:
                failed.append(f"{Path(file).name}\n{exc}")
                self._write_error_log(file, exc)
        self.root.after(0, lambda: self._conversion_finished(success, failed))

    def _conversion_finished(self, success: list[Path], failed: list[str]) -> None:
        if success:
            self.last_pdf_path = success[-1]

        self.set_busy(False)

        if success and not failed:
            if len(success) == 1:
                self.set_status(f"完成四格 PDF：{success[0].name}")
            else:
                self.set_status(f"完成 {len(success)} 個四格 PDF，小圓點可開啟最後一個。")
        elif success and failed:
            self.set_status(f"完成 {len(success)} 個，失敗 {len(failed)} 個。")
            messagebox.showwarning("部分轉檔失敗", "\n\n".join(failed[:3]))
        else:
            self.set_status("轉檔失敗，請確認是否已安裝 PowerPoint、LibreOffice 或 PyMuPDF。")
            messagebox.showerror("轉檔失敗", "\n\n".join(failed[:3]))

    def _write_error_log(self, source_file: str, exc: Exception) -> None:
        try:
            log_path = Path(source_file).with_name(f"{PPT4_APP_NAME}_error.log")
            with log_path.open("a", encoding="utf-8") as log_file:
                log_file.write("=" * 60 + "\n")
                log_file.write(datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
                log_file.write(f"Source: {source_file}\n")
                log_file.write(str(exc) + "\n")
                log_file.write(traceback.format_exc() + "\n")
        except Exception:
            pass


class PDFTurnPanel:
    def __init__(self, root):
        self.root = root
        self.last_dir = None
        self._style()
        self._build()

    def _style(self):
        style = ttk.Style(self.root)
        style.configure(".", background=BG, foreground=TEXT, font=FONT)
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, foreground=TEXT, font=FONT)
        style.configure("TButton", padding=(10, 5), font=LARGE_BTN_FONT)
        style.configure(
            "Stat.TLabel",
            background=CARD,
            foreground=TEXT,
            font=TITLE_FONT,
            padding=(18, 18),
        )
        style.configure("TurnTabHost.TFrame", background=BG)

    def _build(self):
        self.tab_bar = ttk.Frame(
            self.root, padding=(12, 12, 12, 4), style="TurnTabHost.TFrame"
        )
        self.tab_bar.pack(fill="x")
        self.content = ttk.Frame(self.root, style="App.TFrame")
        self.content.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.tabs = [
            ("頁面移轉", RotateTab(self, self.content)),
            ("壓縮檔案", CompressTab(self, self.content)),
            ("頁面合併", MergeTab(self, self.content)),
            ("頁面編輯", EditTab(self, self.content)),
        ]
        self.tab_buttons = []
        for index, (title, _tab) in enumerate(self.tabs):
            button = ctk.CTkButton(
                self.tab_bar,
                text=title,
                command=lambda i=index: self.show_tab(i),
                width=142,
                height=44,
                corner_radius=12,
                fg_color=TURN_TAB_COLORS[index],
                hover_color=TURN_TAB_HOVER,
                text_color=TEXT,
                font=TURN_TAB_FONT,
            )
            button.pack(side="left", padx=(0, 8))
            self.tab_buttons.append(button)
        self.show_tab(0)

    def show_tab(self, active_index):
        for index, (_title, tab) in enumerate(self.tabs):
            tab.pack_forget()
            self.tab_buttons[index].configure(
                fg_color=TURN_TAB_ACTIVE
                if index == active_index
                else TURN_TAB_COLORS[index],
                text_color="white" if index == active_index else TEXT,
            )
        self.tabs[active_index][1].pack(fill="both", expand=True)


# =========================================================
# Main
# =========================================================
class PDFRenameTool:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self._work_area_refresh_job = None
        self._last_work_area = None
        self.configure_root_window()
        self.root.resizable(True, True)

        self.state = PDFState()
        self.deleted_files = []
        self.move_history = []
        self.move_folder = ""
        self.move_recent_folders = []
        self.move_sort_column = "name"
        self.move_sort_reverse = False
        self.current_mode = "rename"
        self.pdf_doc = None

        self.preview_img = None
        self.preview_pil_img = None
        self.preview_gc_counter = 0

        self.ocr_engine = OCREngine()
        self.ocr_select_mode = tk.BooleanVar(value=False)
        self.topmost_var = tk.BooleanVar(value=False)
        self.ocr_start = None
        self.ocr_rect_id = None

        self.company_options = ["中工段", "中興監造", "聖穎", "建業"]

        self.vars = {}
        self.entry_widgets = {}

        self._auto_fit_preview = True

        self.create_style()
        self.create_ui()

        # 先讓主視窗顯示出來，再用背景執行緒逐步預熱 PDF 相關模組。
        # OCR 套件仍維持真正使用時才載入，避免啟動後背景也吃太多資源。
        self._deferred_preload_started = False
        self.root.after(700, self.start_deferred_preload)

    def get_current_monitor_work_area(self):
        """取得目前視窗所在螢幕扣除工作列後的可用範圍。"""
        try:
            self.root.update_idletasks()
        except Exception:
            pass

        if platform.system().lower() == "windows":
            try:
                import ctypes
                from ctypes import wintypes

                class MONITORINFO(ctypes.Structure):
                    _fields_ = [
                        ("cbSize", wintypes.DWORD),
                        ("rcMonitor", wintypes.RECT),
                        ("rcWork", wintypes.RECT),
                        ("dwFlags", wintypes.DWORD),
                    ]

                user32 = ctypes.windll.user32
                hwnd = int(self.root.winfo_id())

                # Tk 在 Windows 上可能回傳內層 HWND；優先取得其頂層父視窗。
                parent_hwnd = user32.GetParent(hwnd)
                if parent_hwnd:
                    hwnd = parent_hwnd

                MONITOR_DEFAULTTONEAREST = 2
                monitor = user32.MonitorFromWindow(
                    hwnd,
                    MONITOR_DEFAULTTONEAREST,
                )

                info = MONITORINFO()
                info.cbSize = ctypes.sizeof(MONITORINFO)

                if monitor and user32.GetMonitorInfoW(monitor, ctypes.byref(info)):
                    work = info.rcWork
                    return (
                        int(work.left),
                        int(work.top),
                        int(work.right),
                        int(work.bottom),
                    )
            except Exception:
                pass

            # Windows 備援：取得主要螢幕的系統工作區。
            try:
                import ctypes
                from ctypes import wintypes

                rect = wintypes.RECT()
                SPI_GETWORKAREA = 0x0030
                if ctypes.windll.user32.SystemParametersInfoW(
                    SPI_GETWORKAREA,
                    0,
                    ctypes.byref(rect),
                    0,
                ):
                    return (
                        int(rect.left),
                        int(rect.top),
                        int(rect.right),
                        int(rect.bottom),
                    )
            except Exception:
                pass

        try:
            screen_w = max(900, int(self.root.winfo_screenwidth()))
            screen_h = max(650, int(self.root.winfo_screenheight()))
        except Exception:
            screen_w, screen_h = 1500, 900

        return 0, 0, screen_w, screen_h

    def apply_work_area_limits(self, center_window=False):
        """限制主視窗在工作區內，避免最大化後被工作列遮住。"""
        left, top, right, bottom = self.get_current_monitor_work_area()
        work_w = max(900, right - left)
        work_h = max(650, bottom - top)

        # 額外保留少量安全距離，兼容 Windows 顯示比例及自動隱藏工作列。
        side_margin = 8
        bottom_margin = 12
        max_width = max(760, work_w - side_margin)
        max_height = max(560, work_h - bottom_margin)

        self._last_work_area = (left, top, right, bottom)

        try:
            self.root.maxsize(max_width, max_height)
        except Exception:
            pass

        if center_window:
            width = min(1500, max(900, int(max_width * 0.92)))
            height = min(900, max(640, int(max_height * 0.88)))
            min_width = min(900, max(760, int(max_width * 0.74)))
            min_height = min(650, max(560, int(max_height * 0.68)))

            x = left + max(0, (work_w - width) // 2)
            y = top + max(0, (work_h - height) // 2)

            self.root.geometry(f"{width}x{height}+{x}+{y}")
            self.root.minsize(min_width, min_height)

    def schedule_work_area_refresh(self, event=None):
        """視窗移到其他螢幕或最大化時，重新套用該螢幕工作區。"""
        if event is not None and event.widget is not self.root:
            return

        if self._work_area_refresh_job is not None:
            with suppress(Exception):
                self.root.after_cancel(self._work_area_refresh_job)

        self._work_area_refresh_job = self.root.after(
            180,
            self.refresh_work_area_limits,
        )

    def refresh_work_area_limits(self):
        self._work_area_refresh_job = None
        current = self.get_current_monitor_work_area()

        if current != self._last_work_area:
            self.apply_work_area_limits(center_window=False)

    def configure_root_window(self):
        self.apply_work_area_limits(center_window=True)
        self.root.bind(
            "<Configure>",
            self.schedule_work_area_refresh,
            add="+",
        )

    def start_deferred_preload(self):
        if self._deferred_preload_started:
            return
        self._deferred_preload_started = True

        def worker():
            preload_list = (
                (fitz, "PyMuPDF / PDF核心"),
                (Image, "Pillow / 影像核心"),
                (ImageDraw, "Pillow / 文字繪製"),
                (ImageFont, "Pillow / 字型處理"),
            )
            for proxy, label in preload_list:
                try:
                    proxy._load()
                    append_startup_log(f"背景預熱完成：{label}")
                except Exception as exc:
                    append_startup_log(f"背景預熱略過：{label}：{exc}")
                time.sleep(0.25)

        threading.Thread(
            target=worker, name="GuppyDeferredPreload", daemon=True
        ).start()

    # =====================================================
    # UI Factory
    # =====================================================
    def create_style(self):
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Treeview",
            rowheight=34,
            font=TREE_FONT,
            background="white",
            fieldbackground="white",
        )
        style.configure("Treeview.Heading", font=TITLE_FONT)
        style.configure(".", background=BG, foreground=TEXT, font=FONT)
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, foreground=TEXT, font=FONT)
        style.configure("TButton", padding=(10, 5), font=LARGE_BTN_FONT)
        style.configure("TCheckbutton", background=CARD, foreground=TEXT, font=FONT)
        style.configure("App.TFrame", background=BG)
        style.configure("Card.TFrame", background=CARD)
        style.configure(
            "App.TLabelframe",
            background=CARD,
            bordercolor=PREVIEW_BORDER,
            relief="solid",
        )
        style.configure(
            "App.TLabelframe.Label", background=CARD, foreground=TEXT, font=TITLE_FONT
        )
        style.configure("App.TLabel", background=BG, foreground=TEXT, font=FONT)
        style.configure("Card.TLabel", background=CARD, foreground=TEXT, font=FONT)
        style.configure(
            "Muted.TLabel", background=CARD, foreground=MUTED_TEXT, font=FONT
        )
        style.configure(
            "Accent.TButton",
            font=LARGE_BTN_FONT,
            foreground="white",
            background=PRIMARY,
            borderwidth=1,
        )
        style.map(
            "Accent.TButton",
            background=[("active", PRIMARY_HOVER), ("pressed", PRIMARY_HOVER)],
        )
        style.configure(
            "Soft.TButton",
            font=LARGE_BTN_FONT,
            foreground="black",
            background=PREVIEW_BLUE,
            borderwidth=1,
        )
        style.map(
            "Soft.TButton",
            background=[
                ("active", PRIMARY_SOFT_HOVER),
                ("pressed", PRIMARY_SOFT_HOVER),
            ],
        )

    def button(
        self,
        parent,
        text,
        command,
        width=90,
        color=PRIMARY,
        hover=PRIMARY_HOVER,
        text_color="white",
        border=False,
    ):
        width = max(width, int(len(text) * 18 + 28))
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            width=width,
            height=30,
            corner_radius=10,
            fg_color=color,
            hover_color=hover,
            text_color=text_color,
            border_width=1 if border else 0,
            border_color=PREVIEW_BORDER,
            font=LARGE_BTN_FONT,
        )

    def entry(self, parent, var, color="white"):
        return ctk.CTkEntry(
            parent,
            textvariable=var,
            height=INPUT_HEIGHT,
            font=FONT,
            corner_radius=10,
            fg_color=color,
        )

    def combo(self, parent, var, values):
        return ctk.CTkComboBox(
            parent,
            variable=var,
            values=values,
            height=INPUT_HEIGHT,
            font=FONT,
            corner_radius=10,
            state="normal",
        )

    # =====================================================
    # Layout
    # =====================================================
    def create_ui(self):
        self.root.configure(bg=BG)

        self.main = tk.Frame(self.root, bg=BG)
        self.main.pack(side="top", fill="both", expand=True)
        self.main.grid_rowconfigure(0, weight=1)
        # 兩個主要瀏覽區共同分配空間；中間搬移欄與右側分頁欄維持窄欄，避免最大化後被固定寬度推擠。
        self.main.grid_columnconfigure(0, weight=1, minsize=300, uniform="maincols")
        self.main.grid_columnconfigure(1, weight=0, minsize=54)
        self.main.grid_columnconfigure(2, weight=1, minsize=320, uniform="maincols")
        self.main.grid_columnconfigure(3, weight=0, minsize=58)

        # 左半部：原本更名檔案瀏覽區保留，但可隨視窗縮放
        self.left_shell = tk.Frame(self.main, bg=BG)
        self.left_shell.grid(row=0, column=0, sticky="nsew", padx=(12, 6), pady=12)
        # 第一分頁左側改用 grid：資料夾列固定、檔案瀏覽區伸縮、左下欄位固定顯示。
        # 避免視窗最大化或縮放時 Treeview 把下方欄位擠出畫面。
        self.left_shell.grid_rowconfigure(0, weight=0)
        self.left_shell.grid_rowconfigure(1, weight=1, minsize=180)
        self.left_shell.grid_rowconfigure(2, weight=0)
        self.left_shell.grid_columnconfigure(0, weight=1)

        # 中間：搬移箭頭固定放在左右兩半正中間；更名模式時只隱藏按鈕，不改變版面結構
        self.center_move_bar = tk.Frame(self.main, bg=BG, width=64)
        self.center_move_bar.grid(row=0, column=1, sticky="ns", padx=(0, 0), pady=12)
        self.center_move_bar.grid_propagate(False)

        # 最右側縱向分頁標籤：固定窄欄，主內容縮放時不被擠出視窗
        self.tab_bar = tk.Frame(self.main, bg=BG, width=68)
        self.tab_bar.grid(row=0, column=3, sticky="ns", padx=(0, 12), pady=12)
        self.tab_bar.grid_propagate(False)

        # 右半部：更名預覽 / 搬移瀏覽區堆疊切換
        self.right_shell = tk.Frame(self.main, bg=BG)
        self.right_shell.grid(row=0, column=2, sticky="nsew", padx=(6, 6), pady=12)
        self.right_stack = tk.Frame(self.right_shell, bg=BG)
        self.right_stack.pack(side="left", fill="both", expand=True)

        self.rename_page = tk.Frame(self.right_stack, bg=BG)
        self.move_page = tk.Frame(self.right_stack, bg=BG)
        self.watermark_page = tk.Frame(self.main, bg=BG)
        self.turn_page = tk.Frame(self.main, bg=BG)
        self.convert_page = tk.Frame(self.main, bg=BG)
        self.ppt4_page = tk.Frame(self.main, bg=BG)

        self.create_folder_area(self.left_shell)
        self.create_treeview(self.left_shell)
        self.create_form(self.left_shell)

        self.create_preview_toolbar(self.rename_page)
        self.create_preview_area(self.rename_page)

        self.create_move_area(self.move_page)
        self.watermark_app = PDFWatermarkApp(self.watermark_page, embedded=True)
        self.turn_app = PDFTurnPanel(self.turn_page)
        self.convert_app = UniversalConvertPanel(self.convert_page)
        self.ppt4_app = PPT4ConvertPanel(self.ppt4_page)
        self.create_center_move_button()
        self.create_vertical_tabs()
        self.switch_mode("rename")
        self.create_signature_footer()

    def create_signature_footer(self):
        footer = tk.Frame(self.root, bg=BG, height=38)
        footer.pack(side="bottom", fill="x")
        footer.pack_propagate(False)

        # 左右兩側採相同權重，讓簽名維持視窗正中央；
        # 固定最上層勾選框則貼齊右下角。
        footer.grid_columnconfigure(0, weight=1)
        footer.grid_columnconfigure(1, weight=0)
        footer.grid_columnconfigure(2, weight=1)

        tk.Label(
            footer,
            text="Inspired by Atex's high thoughts.",
            bg=BG,
            fg=FOOTER_TEXT,
            font=SIGN_FONT,
        ).grid(row=0, column=1, sticky="s", pady=(5, 5))

        self.topmost_checkbox = ctk.CTkCheckBox(
            footer,
            text="固定最上層",
            variable=self.topmost_var,
            command=self.toggle_topmost,
            width=118,
            height=24,
            checkbox_width=18,
            checkbox_height=18,
            corner_radius=4,
            border_width=2,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color=TEXT,
            font=BTN_FONT,
        )
        self.topmost_checkbox.grid(
            row=0,
            column=2,
            sticky="se",
            padx=(8, 14),
            pady=(3, 7),
        )

    def toggle_topmost(self):
        """依右下角勾選狀態即時切換主視窗是否固定在最上層。"""
        enabled = bool(self.topmost_var.get())
        try:
            self.root.attributes("-topmost", enabled)
            if enabled:
                # 啟用時立即將視窗提到前方；後續由作業系統維持置頂。
                self.root.lift()
                self.root.after_idle(self.root.lift)

            # 置頂狀態切換後重新套用工作區，避免最大化時壓到工作列。
            self.root.after(
                80,
                lambda: self.apply_work_area_limits(center_window=False),
            )
        except Exception as exc:
            self.topmost_var.set(False)
            messagebox.showerror(
                "固定最上層",
                f"無法切換視窗最上層狀態：\n{exc}",
            )

    def create_center_move_button(self):
        """建立左右瀏覽區中央的向右搬移按鈕。
        使用 place 固定在中間偏上位置，避免切換分頁後被 pack 版面擠到下方。
        """
        self.center_move_btn = ctk.CTkButton(
            self.center_move_bar,
            text="→",
            command=self.move_selected_files_to_right,
            width=54,
            height=78,
            corner_radius=20,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="white",
            font=TITLE_FONT,
        )
        self.center_move_btn.place(relx=0.5, rely=0.38, anchor="center")
        self.center_move_btn.place_forget()

    def create_vertical_tabs(self):
        self.rename_tab_btn = ctk.CTkButton(
            self.tab_bar,
            text="更\n名",
            command=lambda: self.switch_mode("rename"),
            width=50,
            height=TAB_BUTTON_HEIGHT,
            corner_radius=14,
            font=TITLE_FONT,
        )
        self.rename_tab_btn.pack(pady=(10, 10), padx=8)

        self.move_tab_btn = ctk.CTkButton(
            self.tab_bar,
            text="搬\n移",
            command=lambda: self.switch_mode("move"),
            width=50,
            height=TAB_BUTTON_HEIGHT,
            corner_radius=14,
            font=TITLE_FONT,
        )
        self.move_tab_btn.pack(pady=(0, 10), padx=8)

        self.watermark_tab_btn = ctk.CTkButton(
            self.tab_bar,
            text="浮\n水\n印",
            command=lambda: self.switch_mode("watermark"),
            width=50,
            height=TAB_BUTTON_HEIGHT,
            corner_radius=14,
            font=TITLE_FONT,
        )
        self.watermark_tab_btn.pack(pady=(0, 10), padx=8)

        self.turn_tab_btn = ctk.CTkButton(
            self.tab_bar,
            text="旋\n壓\n合\n切",
            command=lambda: self.switch_mode("turn"),
            width=50,
            height=TAB_BUTTON_HEIGHT,
            corner_radius=14,
            font=TITLE_FONT,
        )
        self.turn_tab_btn.pack(pady=(0, 10), padx=8)

        self.convert_tab_btn = ctk.CTkButton(
            self.tab_bar,
            text="轉\n換\n氣\n體",
            command=lambda: self.switch_mode("convert"),
            width=50,
            height=TAB_BUTTON_HEIGHT,
            corner_radius=14,
            font=TITLE_FONT,
        )
        self.convert_tab_btn.pack(pady=(0, 10), padx=8)

        self.ppt4_tab_btn = ctk.CTkButton(
            self.tab_bar,
            text="PTT\n轉\n4格",
            command=lambda: self.switch_mode("ppt4"),
            width=50,
            height=TAB_BUTTON_HEIGHT,
            corner_radius=14,
            font=TITLE_FONT,
        )
        self.ppt4_tab_btn.pack(pady=(0, 10), padx=8)

    def switch_mode(self, mode):
        self.current_mode = mode
        self.rename_page.pack_forget()
        self.move_page.pack_forget()
        self.watermark_page.grid_forget()
        self.turn_page.grid_forget()
        self.convert_page.grid_forget()
        self.ppt4_page.grid_forget()

        inactive_color = TAB_INACTIVE
        self.rename_tab_btn.configure(fg_color=inactive_color, text_color="black")
        self.move_tab_btn.configure(fg_color=inactive_color, text_color="black")
        self.watermark_tab_btn.configure(fg_color=inactive_color, text_color="black")
        self.turn_tab_btn.configure(fg_color=inactive_color, text_color="black")
        self.convert_tab_btn.configure(fg_color=inactive_color, text_color="black")
        self.ppt4_tab_btn.configure(fg_color=inactive_color, text_color="black")

        if mode in ("watermark", "turn", "convert", "ppt4"):
            self.left_shell.grid_forget()
            self.center_move_bar.grid_forget()
            self.right_shell.grid_forget()
            self.center_move_btn.place_forget()
            if mode == "watermark":
                self.watermark_page.grid(
                    row=0, column=0, columnspan=3, sticky="nsew", padx=(12, 6), pady=12
                )
                self.watermark_page.grid_rowconfigure(0, weight=1)
                self.watermark_page.grid_columnconfigure(0, weight=1)
                self.watermark_tab_btn.configure(fg_color=PRIMARY, text_color="white")
                self.root.after(100, self.watermark_app.setup_drag_drop)
            elif mode == "turn":
                self.turn_page.grid(
                    row=0, column=0, columnspan=3, sticky="nsew", padx=(12, 6), pady=12
                )
                self.turn_page.grid_rowconfigure(0, weight=1)
                self.turn_page.grid_columnconfigure(0, weight=1)
                self.turn_tab_btn.configure(fg_color=PRIMARY, text_color="white")
            elif mode == "convert":
                self.convert_page.grid(
                    row=0, column=0, columnspan=3, sticky="nsew", padx=(12, 6), pady=12
                )
                self.convert_page.grid_rowconfigure(0, weight=1)
                self.convert_page.grid_columnconfigure(0, weight=1)
                self.convert_tab_btn.configure(fg_color=PRIMARY, text_color="white")
                self.root.after(100, self.convert_app.setup_drag_drop)
            else:
                self.ppt4_page.grid(
                    row=0, column=0, columnspan=3, sticky="nsew", padx=(12, 6), pady=12
                )
                self.ppt4_page.grid_rowconfigure(0, weight=1)
                self.ppt4_page.grid_columnconfigure(0, weight=1)
                self.ppt4_tab_btn.configure(fg_color=PRIMARY, text_color="white")
                self.root.after(100, self.ppt4_app.setup_drag_drop)
            return

        self.left_shell.grid(row=0, column=0, sticky="nsew", padx=(12, 6), pady=12)
        self.center_move_bar.grid(row=0, column=1, sticky="ns", padx=(0, 0), pady=12)
        self.right_shell.grid(row=0, column=2, sticky="nsew", padx=(6, 6), pady=12)

        if mode == "move":
            self.move_page.pack(fill="both", expand=True)
            self.center_move_btn.place(relx=0.5, rely=0.38, anchor="center")
            self.move_tab_btn.configure(fg_color=PRIMARY, text_color="white")
            if not self.move_folder and self.state.folder:
                self.set_move_folder(self.state.folder)
        else:
            self.rename_page.pack(fill="both", expand=True)
            self.center_move_btn.place_forget()
            self.rename_tab_btn.configure(fg_color=PRIMARY, text_color="white")

    def create_folder_area(self, parent):
        frame = self.card(parent)
        self.folder_area_frame = frame
        frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        frame.configure(padx=12, pady=12)

        tk.Label(frame, text="選擇資料夾", bg=CARD, fg=TEXT, font=TITLE_FONT).pack(
            anchor="w", pady=(0, 8)
        )

        row = tk.Frame(frame, bg=CARD)
        row.pack(fill="x")

        self.folder_var = tk.StringVar()
        self.entry(row, self.folder_var).pack(side="left", fill="x", expand=True)

        button_row = tk.Frame(frame, bg=CARD)
        self.folder_button_row = button_row
        button_row.pack(fill="x", pady=(8, 0))

        buttons = [
            ("瀏覽資料夾", self.browse_folder, 110, PRIMARY, PRIMARY_HOVER, "white"),
            ("讀取檔名", self.read_selected_filename, 100, PREVIEW_BLUE, PRIMARY_SOFT_HOVER, "black"),
            ("刪除檔案", self.delete_pdf, 100, RED, RED_HOVER, "white"),
            ("回復刪除", self.restore_pdf, 100, YELLOW, YELLOW_HOVER, "black"),
        ]

        for text, cmd, width, color, hover, text_color in buttons:
            self.button(button_row, text, cmd, width, color, hover, text_color).pack(
                side="left", padx=2
            )

    def create_treeview(self, parent):
        frame = self.card(parent)
        self.tree_frame = frame
        frame.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        columns = ("no", "filename", "date")
        self.tree = ttk.Treeview(
            frame,
            columns=columns,
            show="headings",
            height=8,
            selectmode="extended",
        )

        headers = {
            "no": ("項次", None),
            "filename": ("檔名", lambda: self.sort_tree("filename")),
            "date": ("加入時間", lambda: self.sort_tree("date")),
        }

        for col, (text, cmd) in headers.items():
            if cmd:
                self.tree.heading(col, text=text, command=cmd)
            else:
                self.tree.heading(col, text=text)

        self.tree.column("no", width=70, minwidth=70, anchor="center", stretch=False)
        self.tree.column("filename", width=430, minwidth=220, stretch=True)
        self.tree.column(
            "date", width=180, minwidth=160, anchor="center", stretch=False
        )

        scroll = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll.set, xscrollcommand=x_scroll.set)

        scroll.pack(side="right", fill="y")
        x_scroll.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.select_pdf)

    def create_form(self, parent):
        outer = self.card(parent)
        self.form_outer = outer
        outer.grid(row=2, column=0, sticky="ew")
        outer.configure(padx=12, pady=12)

        form = tk.Frame(outer, bg=CARD)
        self.rename_form = form
        form.pack(fill="x")

        fields = [
            ("發文單位", "combo"),
            ("文號", "entry"),
            ("主旨", "entry"),
            ("收文號碼", "entry"),
            ("收文日期", "entry"),
            ("增加前名", "combo"),
        ]

        for row, (title, kind) in enumerate(fields):
            tk.Label(form, text=title, bg=CARD, fg=TEXT, font=FONT).grid(
                row=row, column=0, sticky="w", pady=1
            )

            var = tk.StringVar()
            self.vars[title] = var

            if kind == "combo":
                widget = self.combo(form, var, self.company_options)
                widget.set("")
            else:
                widget = self.entry(form, var)
                widget.bind(
                    "<Button-1>",
                    lambda _event, field=title: self.fill_field_from_ocr(field),
                )

            self.entry_widgets[title] = widget
            widget.grid(row=row, column=1, sticky="ew", padx=6, pady=1)
            var.trace_add("write", self.update_preview)

        form.columnconfigure(1, weight=1)
        self.create_bottom_rename_area(outer)

    def create_bottom_rename_area(self, parent):
        bottom = tk.Frame(
            parent,
            bg=PRIMARY_SOFT,
            highlightbackground=PREVIEW_BORDER,
            highlightthickness=1,
        )
        self.bottom_rename_area = bottom
        bottom.pack(fill="x", pady=(6, 0))
        bottom.configure(padx=8, pady=8)

        self.preview_var = tk.StringVar()
        self.prefix_var = tk.StringVar()

        rows = [
            ("變更檔名", self.preview_var, FILENAME_BG, "更名確認", self.rename_pdf, 0),
            ("前名調整", self.prefix_var, PREFIX_BG, "增名確認", self.rename_prefix, 1),
        ]

        for label, var, color, btn_text, cmd, row in rows:
            pady = (2, 0) if row else (0, 1)

            tk.Label(bottom, text=label, bg=PRIMARY_SOFT, fg=TEXT, font=FONT).grid(
                row=row, column=0, sticky="w", pady=pady
            )

            self.entry(bottom, var, color).grid(
                row=row, column=1, sticky="ew", padx=6, pady=(2, 0) if row else 0
            )

            self.button(bottom, btn_text, cmd, width=110 if row else 100).grid(
                row=row, column=2, padx=3, pady=(2, 0) if row else 0
            )

        bottom.columnconfigure(1, weight=1)

    def create_preview_toolbar(self, parent):
        # 低解析度時分成兩列，先保住翻頁與縮放控制，避免被 OCR 狀態列擠出可視範圍。
        bar = tk.Frame(parent, bg=CARD)
        bar.pack(fill="x", pady=(0, 10))
        bar.grid_columnconfigure(0, weight=1)

        left = tk.Frame(bar, bg=CARD)
        left.grid(row=0, column=0, sticky="w", padx=10, pady=(8, 3))

        for text, cmd, width in (
            ("-", self.zoom_out, 44),
            ("+", self.zoom_in, 44),
            ("Fit", self.fit_page, 58),
        ):
            self.preview_button(left, text, cmd, width).pack(side="left", padx=2)

        self.page_var = tk.StringVar(value="0 / 0")
        tk.Label(left, textvariable=self.page_var, bg=CARD, fg=TEXT, font=FONT).pack(
            side="left", padx=(10, 6)
        )

        for text, cmd in (("<", self.prev_page), (">", self.next_page)):
            self.preview_button(left, text, cmd, 38).pack(side="left", padx=2)

        right = tk.Frame(bar, bg=CARD)
        right.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
        right.grid_columnconfigure(0, weight=1)

        self.ocr_status_var = tk.StringVar(value="OCR：尚未載入")
        tk.Label(
            right, textvariable=self.ocr_status_var, bg=CARD, fg=TEXT, font=FONT
        ).grid(row=0, column=0, sticky="w", padx=(0, 10))

        self.ocr_check = ctk.CTkCheckBox(
            right,
            text="框選辨識",
            variable=self.ocr_select_mode,
            command=self.toggle_ocr_mode,
            font=BTN_FONT,
        )
        self.ocr_check.grid(row=0, column=1, sticky="e", padx=2)

    def create_preview_area(self, parent):
        frame = self.card(parent)
        frame.pack(fill="both", expand=True)

        canvas_frame = tk.Frame(frame, bg=CARD)
        canvas_frame.pack(fill="both", expand=True)

        x_scroll = tk.Scrollbar(canvas_frame, orient="horizontal")
        y_scroll = tk.Scrollbar(canvas_frame, orient="vertical")

        self.canvas = tk.Canvas(
            canvas_frame,
            bg="white",
            highlightthickness=0,
            xscrollcommand=x_scroll.set,
            yscrollcommand=y_scroll.set,
        )

        x_scroll.config(command=self.canvas.xview)
        y_scroll.config(command=self.canvas.yview)

        x_scroll.pack(side="bottom", fill="x")
        y_scroll.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        for event, handler in {
            "<MouseWheel>": self.mouse_zoom,
            "<ButtonPress-1>": self.canvas_mouse_down,
            "<B1-Motion>": self.canvas_mouse_move,
            "<ButtonRelease-1>": self.canvas_mouse_up,
        }.items():
            self.canvas.bind(event, handler)

        self.create_ocr_text_area(frame)

    def create_ocr_text_area(self, parent):
        frame = tk.Frame(parent, bg=CARD, height=125)
        frame.pack(fill="x", pady=(10, 0))
        frame.pack_propagate(False)

        title_row = tk.Frame(frame, bg=CARD)
        title_row.pack(fill="x", padx=10, pady=(8, 4))

        tk.Label(
            title_row, text="測試辨識字串", bg=CARD, fg=TEXT, font=TITLE_FONT
        ).pack(side="left")
        self.preview_button(title_row, "清除", self.clear_ocr_text, 70).pack(
            side="right"
        )

        self.ocr_text = tk.Text(
            frame,
            height=3,
            font=FONT,
            bg=OCR_BG,
            fg=TEXT,
            wrap="word",
            relief="solid",
            bd=1,
        )
        self.ocr_text.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        self.set_text(self.ocr_text, OCR_PLACEHOLDER)

    def card(self, parent):
        return tk.Frame(parent, bg=CARD)

    def preview_button(self, parent, text, command, width):
        return self.button(
            parent,
            text,
            command,
            width=width,
            color=PREVIEW_BLUE,
            hover=PRIMARY_SOFT_HOVER,
            text_color="black",
            border=True,
        )

    @staticmethod
    def set_text(widget, text):
        widget.delete("1.0", "end")
        widget.insert("1.0", text)

    def refresh_file_views(self, refresh_move=True):
        self.load_pdfs()
        if refresh_move and self.move_folder:
            self.load_move_tree()

    def clear_current_pdf(self, clear_canvas=False):
        self.close_pdf()
        self.state.selected_pdf = ""
        self.state.current_pdf_path = ""
        self.state.current_page = 0
        self.release_preview_image(clear_canvas=clear_canvas, force_collect=True)
        self.page_var.set("0 / 0")

    def select_pdf_in_tree(self, filename):
        for item_id in self.tree.get_children():
            values = self.tree.item(item_id, "values")
            if values and values[1] == filename:
                self.tree.selection_set(item_id)
                self.tree.see(item_id)
                return True
        return False

    def collect_preview_memory(self, force=False):
        self.preview_gc_counter += 1
        if force or self.preview_gc_counter >= PREVIEW_GC_INTERVAL:
            self.preview_gc_counter = 0
            gc.collect()

    def release_preview_image(self, clear_canvas=False, force_collect=False):
        if clear_canvas:
            self.canvas.delete("all")

        if self.preview_pil_img is not None:
            with suppress(Exception):
                self.preview_pil_img.close()

        self.preview_pil_img = None
        self.preview_img = None
        self.collect_preview_memory(force=force_collect)

    @staticmethod
    def limited_preview_zoom(page, requested_zoom):
        page_pixels = max(float(page.rect.width * page.rect.height), 1.0)
        pixel_limited_zoom = (MAX_PREVIEW_PIXELS / page_pixels) ** 0.5
        return max(MIN_ZOOM, min(requested_zoom, MAX_ZOOM, pixel_limited_zoom))

    # =====================================================
    # Move Browser
    # =====================================================
    def create_move_area(self, parent):
        top = self.card(parent)
        self.move_top_area = top
        top.pack(fill="x", pady=(0, 8))
        top.configure(padx=8, pady=8)

        tk.Label(top, text="搬移目的資料夾", bg=CARD, fg=TEXT, font=TITLE_FONT).pack(
            anchor="w", pady=(0, 4)
        )

        row = tk.Frame(top, bg=CARD)
        row.pack(fill="x")

        self.move_folder_var = tk.StringVar()
        self.move_folder_combo = ctk.CTkComboBox(
            row,
            variable=self.move_folder_var,
            values=[],
            height=INPUT_HEIGHT,
            font=FONT,
            corner_radius=10,
            state="normal",
            command=self.on_move_folder_combo,
        )
        self.move_folder_combo.pack(side="left", fill="x", expand=True)
        self.move_folder_combo.bind(
            "<Return>", lambda _event: self.set_move_folder(self.move_folder_var.get())
        )

        move_button_row = tk.Frame(top, bg=CARD)
        self.move_button_row = move_button_row
        move_button_row.pack(fill="x", pady=(4, 0))

        self.button(move_button_row, "瀏覽", self.browse_move_folder, 80).pack(
            side="left", padx=1
        )
        self.button(
            move_button_row,
            "上一層",
            self.move_parent_folder,
            85,
            color=PREVIEW_BLUE,
            hover=PRIMARY_SOFT_HOVER,
            text_color="black",
            border=True,
        ).pack(side="left", padx=1)
        self.button(
            move_button_row,
            "回復移動",
            self.undo_last_move,
            100,
            color=YELLOW,
            hover=YELLOW_HOVER,
            text_color="black",
        ).pack(side="left", padx=1)

        body = self.card(parent)
        self.move_body_area = body
        body.pack(fill="both", expand=True)

        tree_area = tk.Frame(body, bg=CARD)
        tree_area.pack(fill="both", expand=True)

        columns = ("name", "size", "created", "modified")
        self.move_tree = ttk.Treeview(tree_area, columns=columns, show="headings")
        headers = {
            "name": "名稱",
            "size": "大小",
            "created": "加入時間",
            "modified": "修改時間",
        }
        for col, title in headers.items():
            self.move_tree.heading(
                col, text=title, command=lambda c=col: self.sort_move_tree(c)
            )

        self.move_tree.column("name", width=360, minwidth=180, stretch=True)
        self.move_tree.column("size", width=90, minwidth=70, anchor="e", stretch=False)
        self.move_tree.column(
            "created", width=140, minwidth=120, anchor="center", stretch=False
        )
        self.move_tree.column(
            "modified", width=140, minwidth=120, anchor="center", stretch=False
        )

        y_scroll = ttk.Scrollbar(
            tree_area, orient="vertical", command=self.move_tree.yview
        )
        self.move_tree.configure(yscrollcommand=y_scroll.set)
        self.move_tree.pack(side="left", fill="both", expand=True)
        y_scroll.pack(side="right", fill="y")
        self.move_tree.bind("<Double-1>", self.enter_selected_move_folder)

        x_scroll = ttk.Scrollbar(
            body, orient="horizontal", command=self.move_tree.xview
        )
        self.move_tree.configure(xscrollcommand=x_scroll.set)
        x_scroll.pack(fill="x", padx=0, pady=(1, 0))

        bottom = self.card(parent)
        self.move_bottom_area = bottom
        bottom.pack(fill="x", pady=(6, 0))
        bottom.configure(padx=8, pady=8)

        self.move_target_var = tk.StringVar(value="目的地：尚未選擇")
        tk.Label(
            bottom,
            textvariable=self.move_target_var,
            bg=CARD,
            fg=TEXT,
            font=FONT,
        ).pack(side="left", fill="x", expand=True)

        tk.Label(
            bottom,
            text="左側可 Ctrl 逐筆多選 / Shift 連續多選",
            bg=CARD,
            fg=MUTED_TEXT,
            font=FONT,
        ).pack(side="right", padx=(10, 0))

    def browse_move_folder(self):
        folder = filedialog.askdirectory(
            initialdir=self.move_folder or self.state.folder or None
        )
        if folder:
            self.set_move_folder(folder)

    def on_move_folder_combo(self, value):
        self.set_move_folder(value)

    def set_move_folder(self, folder):
        if not folder:
            return
        path = Path(folder).expanduser()
        if not path.exists() or not path.is_dir():
            messagebox.showerror("錯誤", f"資料夾不存在：\n{folder}")
            return

        self.move_folder = str(path)
        self.move_folder_var.set(self.move_folder)
        self.move_target_var.set(f"目的地：{self.move_folder}")

        if self.move_folder not in self.move_recent_folders:
            self.move_recent_folders.insert(0, self.move_folder)
            self.move_recent_folders = self.move_recent_folders[:12]
            self.move_folder_combo.configure(values=self.move_recent_folders)

        self.load_move_tree()

    def load_move_tree(self):
        self.move_tree.delete(*self.move_tree.get_children())
        for item in list_directory_items(
            self.move_folder, self.move_sort_column, self.move_sort_reverse
        ):
            icon_name = f"📁 {item['name']}" if item["is_dir"] else f"📄 {item['name']}"
            size_text = "<資料夾>" if item["is_dir"] else format_file_size(item["size"])
            created = format_timestamp(item["created"])
            modified = format_timestamp(item["modified"])
            self.move_tree.insert(
                "",
                "end",
                values=(icon_name, size_text, created, modified),
                tags=(item["path"], "dir" if item["is_dir"] else "file"),
            )

    def sort_move_tree(self, column):
        self.move_sort_column, self.move_sort_reverse = toggle_sort(
            self.move_sort_column, self.move_sort_reverse, column
        )
        self.load_move_tree()

    def move_parent_folder(self):
        if not self.move_folder:
            return
        parent = Path(self.move_folder).parent
        if parent and str(parent) != self.move_folder:
            self.set_move_folder(str(parent))

    def get_selected_move_path(self):
        selected = self.move_tree.selection()
        if not selected:
            return None
        tags = self.move_tree.item(selected[0], "tags")
        return Path(tags[0]) if tags else None

    def enter_selected_move_folder(self, _event=None):
        path = self.get_selected_move_path()
        if path and path.is_dir():
            self.set_move_folder(str(path))

    def get_selected_left_pdf_paths(self):
        """取得左側目前所有選取的 PDF 路徑，維持畫面選取順序。"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning(
                "提醒",
                "請先在左側檔案瀏覽區選擇要搬移的 PDF。\n"
                "可使用 Ctrl 逐筆多選，或 Shift 連續多選。",
            )
            return []

        paths = []
        missing = []
        for item_id in selected:
            values = self.tree.item(item_id, "values")
            if not values or len(values) < 2:
                continue

            src = Path(self.state.folder) / str(values[1])
            if src.exists():
                paths.append(src)
            else:
                missing.append(src)

        if missing:
            messagebox.showerror(
                "錯誤",
                "下列檔案不存在，請重新整理後再試：\n\n"
                + "\n".join(str(path) for path in missing[:10]),
            )
            return []

        return paths

    def get_selected_left_pdf_path(self):
        """相容舊流程：需要單一路徑時回傳目前焦點或第一個選取檔案。"""
        paths = self.get_selected_left_pdf_paths()
        if not paths:
            return None

        focus_id = self.tree.focus()
        if focus_id and focus_id in self.tree.selection():
            values = self.tree.item(focus_id, "values")
            if values and len(values) >= 2:
                focus_path = Path(self.state.folder) / str(values[1])
                if focus_path.exists():
                    return focus_path

        return paths[0]

    def get_move_destination_folder(self):
        selected_path = self.get_selected_move_path()
        if selected_path and selected_path.is_dir():
            return selected_path
        if self.move_folder:
            return Path(self.move_folder)
        return None

    def move_selected_files_to_right(self):
        sources = self.get_selected_left_pdf_paths()
        dst_folder = self.get_move_destination_folder()

        if not sources:
            return
        if not dst_folder:
            messagebox.showwarning("提醒", "請先選擇右側目的資料夾。")
            return

        moved_batch = []
        problems = []
        seen_names = set()

        # 若目前預覽的 PDF 也在本次選取清單內，先關閉檔案控制代碼。
        if self.state.current_pdf_path:
            current = Path(self.state.current_pdf_path)
            if any(current == src for src in sources):
                self.clear_current_pdf(clear_canvas=True)

        for src in sources:
            dst = dst_folder / src.name
            name_key = src.name.lower()

            # 同一批選取中若出現相同檔名，只處理第一個。
            if name_key in seen_names:
                problems.append(f"{src.name}（選取清單中檔名重複）")
                continue
            seen_names.add(name_key)

            # 目的地已存在同名檔：跳過，但不影響其他檔案。
            if dst.exists():
                problems.append(f"{src.name}（目的地已有同名檔案）")
                continue

            try:
                src.rename(dst)
                moved_batch.append((dst, src))
            except Exception as exc:
                problems.append(f"{src.name}（搬移失敗：{exc}）")

        # 有成功搬移的檔案才寫入回復歷史。
        if moved_batch:
            self.move_history.append(moved_batch)

        self.refresh_file_views()

        moved_count = len(moved_batch)
        if moved_count:
            self.move_target_var.set(
                f"已搬移 {moved_count} 個檔案到：{dst_folder}"
            )
        else:
            self.move_target_var.set(f"未搬移檔案：{dst_folder}")

        # 正常完成時完全不跳視窗；只有遇到問題才提示。
        if problems:
            messagebox.showwarning(
                "部分檔案未搬移",
                (
                    f"已成功搬移 {moved_count} 個檔案。\n"
                    f"以下 {len(problems)} 個檔案未搬移：\n\n"
                    + "\n".join(problems[:30])
                    + (
                        f"\n\n其餘 {len(problems) - 30} 個問題檔案未列出。"
                        if len(problems) > 30
                        else ""
                    )
                ),
            )

    def undo_last_move(self):
        if not self.move_history:
            messagebox.showinfo("提醒", "目前沒有可回復的搬移步驟。")
            return

        history_item = self.move_history[-1]

        # V1.1.3 起每個歷史項目是一批 [(搬後路徑, 原路徑), ...]；
        # 同時相容舊版執行中的單筆 tuple 格式。
        if (
            isinstance(history_item, tuple)
            and len(history_item) == 2
            and isinstance(history_item[0], Path)
        ):
            batch = [history_item]
        else:
            batch = list(history_item)

        problems = []
        for moved_path, original_path in batch:
            if not moved_path.exists():
                problems.append(f"找不到已搬移檔案：{moved_path}")
            elif original_path.exists():
                problems.append(f"原位置已有同名檔案：{original_path.name}")

        if problems:
            messagebox.showerror(
                "無法整批回復",
                "為避免只回復部分檔案，本次沒有回復任何檔案。\n\n"
                + "\n".join(problems[:15]),
            )
            return

        restored = []
        try:
            for moved_path, original_path in reversed(batch):
                moved_path.rename(original_path)
                restored.append((original_path, moved_path))

            self.move_history.pop()
            self.refresh_file_views()

            if len(batch) == 1:
                self.move_target_var.set(f"已回復：{batch[0][1]}")
            else:
                self.move_target_var.set(
                    f"已回復上一批搬移：{len(batch)} 個檔案"
                )
        except Exception as exc:
            # 回復途中出錯時，盡可能恢復成搬移後的狀態。
            rollback_errors = []
            for original_path, moved_path in reversed(restored):
                try:
                    if original_path.exists() and not moved_path.exists():
                        original_path.rename(moved_path)
                except Exception as rollback_exc:
                    rollback_errors.append(str(rollback_exc))

            self.refresh_file_views()

            detail = str(exc)
            if rollback_errors:
                detail += (
                    "\n\n部分檔案重新搬回目的地失敗：\n"
                    + "\n".join(rollback_errors[:10])
                )
            messagebox.showerror("回復失敗", detail)

    # =====================================================
    # OCR Fill
    # =====================================================
    def get_ocr_text(self):
        return self.ocr_text.get("1.0", "end").strip()

    def clear_ocr_text(self):
        self.set_text(self.ocr_text, "")

    def fill_field_from_ocr(self, field):
        if not self.ocr_select_mode.get():
            return

        text = self.get_ocr_text()

        if not text or OCR_PLACEHOLDER in text:
            return

        value = (
            normalize_receive_date(text)
            if field == "收文日期"
            else clean_one_line(text)
        )
        self.vars[field].set(value)

    # =====================================================
    # Folder / Tree
    # =====================================================
    def browse_folder(self):
        folder = filedialog.askdirectory()
        if not folder:
            return

        self.state.folder = folder
        self.folder_var.set(folder)
        self.load_pdfs()
        if not self.move_folder:
            self.set_move_folder(folder)

    def read_selected_filename(self):
        """將左側目前選取的 PDF 檔名帶入右下方「變更檔名」欄位。"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("提醒", "請先在左側檔案瀏覽區選擇 PDF。")
            return

        values = self.tree.item(selected[0], "values")
        if not values or len(values) < 2:
            messagebox.showwarning("提醒", "無法讀取目前選取的檔名。")
            return

        filename = safe_pdf_filename(str(values[1]))
        if not filename.lower().endswith(".pdf"):
            filename += ".pdf"
        self.preview_var.set(filename)

    def load_pdfs(self):
        self.tree.delete(*self.tree.get_children())

        pdfs = list_pdf_files(
            self.state.folder, self.state.sort_column, self.state.sort_reverse
        )

        for index, (filename, added) in enumerate(pdfs, start=1):
            date_text = format_timestamp(added)
            self.tree.insert("", "end", values=(index, filename, date_text))

    def sort_tree(self, column):
        self.state.sort_column, self.state.sort_reverse = toggle_sort(
            self.state.sort_column, self.state.sort_reverse, column
        )
        self.load_pdfs()

    def select_pdf(self, _event=None):
        selected = self.tree.selection()
        if not selected:
            return

        # 多選時以目前有焦點的那一列做預覽，不影響其他已選取列。
        item_id = self.tree.focus()
        if not item_id or item_id not in selected:
            item_id = selected[-1]

        values = self.tree.item(item_id, "values")
        if not values:
            return

        self.state.selected_pdf = values[1]
        self.state.current_pdf_path = str(
            Path(self.state.folder) / self.state.selected_pdf
        )
        self.state.current_page = 0

        self.open_pdf(self.state.current_pdf_path)
        self.fit_page()

    # =====================================================
    # PDF Preview
    # =====================================================
    def open_pdf(self, path):
        try:
            self.close_pdf()
            self.pdf_doc = fitz.open(path)
        except Exception as exc:
            messagebox.showerror("錯誤", str(exc))

    def close_pdf(self):
        if self.pdf_doc:
            self.pdf_doc.close()
            self.pdf_doc = None
            self.collect_preview_memory(force=True)

    def show_preview(self):
        if not self.pdf_doc:
            return

        try:
            page = self.pdf_doc.load_page(self.state.current_page)
            render_zoom = self.limited_preview_zoom(page, self.state.zoom)
            self.state.zoom = render_zoom
            matrix = fitz.Matrix(render_zoom, render_zoom)

            self.release_preview_image(clear_canvas=True)
            pix = page.get_pixmap(matrix=matrix, alpha=False)

            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            self.preview_pil_img = img
            self.preview_img = ImageTk.PhotoImage(img)

            self.canvas.create_image(
                IMAGE_OFFSET,
                IMAGE_OFFSET,
                anchor="nw",
                image=self.preview_img,
                tags=("pdf_image",),
            )
            self.canvas.config(scrollregion=self.canvas.bbox("all"))

            self.page_var.set(f"{self.state.current_page + 1} / {len(self.pdf_doc)}")
            del pix
            self.collect_preview_memory()

        except Exception as exc:
            messagebox.showerror("錯誤", f"PDF預覽失敗：{exc}")

    def fit_page(self):
        if not self.pdf_doc:
            return

        page = self.pdf_doc.load_page(self.state.current_page)
        canvas_width = max(self.canvas.winfo_width(), 220)
        canvas_height = max(self.canvas.winfo_height(), 220)
        zoom_w = max(MIN_ZOOM, (canvas_width - IMAGE_OFFSET * 2 - 24) / page.rect.width)
        zoom_h = max(
            MIN_ZOOM, (canvas_height - IMAGE_OFFSET * 2 - 24) / page.rect.height
        )
        self.state.zoom = min(zoom_w, zoom_h, MAX_ZOOM)
        self._auto_fit_preview = True
        self.show_preview()

    def zoom_in(self):
        self._auto_fit_preview = False
        self.state.zoom = min(MAX_ZOOM, self.state.zoom + 0.1)
        self.show_preview()

    def zoom_out(self):
        self._auto_fit_preview = False
        self.state.zoom = max(0.3, self.state.zoom - 0.1)
        self.show_preview()

    def mouse_zoom(self, event):
        if self.ocr_select_mode.get():
            return
        self.zoom_in() if event.delta > 0 else self.zoom_out()

    def next_page(self):
        if self.pdf_doc and self.state.current_page < len(self.pdf_doc) - 1:
            self.state.current_page += 1
            self.show_preview()

    def prev_page(self):
        if self.pdf_doc and self.state.current_page > 0:
            self.state.current_page -= 1
            self.show_preview()

    # =====================================================
    # Pan / OCR Selection
    # =====================================================
    def toggle_ocr_mode(self):
        enabled = self.ocr_select_mode.get()
        self.canvas.config(cursor="crosshair" if enabled else "")

        if not enabled and self.ocr_rect_id:
            self.canvas.delete(self.ocr_rect_id)
            self.ocr_rect_id = None

    def canvas_mouse_down(self, event):
        if self.ocr_select_mode.get():
            self.start_ocr_select(event)
        else:
            self.canvas.scan_mark(event.x, event.y)

    def canvas_mouse_move(self, event):
        if self.ocr_select_mode.get():
            self.move_ocr_select(event)
        else:
            self.canvas.scan_dragto(event.x, event.y, gain=1)

    def canvas_mouse_up(self, event):
        if self.ocr_select_mode.get():
            self.end_ocr_select(event)

    def start_ocr_select(self, event):
        self.ocr_start = (self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))

        if self.ocr_rect_id:
            self.canvas.delete(self.ocr_rect_id)

        x, y = self.ocr_start
        self.ocr_rect_id = self.canvas.create_rectangle(
            x, y, x, y, outline=PRIMARY, width=2, dash=(4, 2)
        )

    def move_ocr_select(self, event):
        if not self.ocr_start or not self.ocr_rect_id:
            return

        x0, y0 = self.ocr_start
        x1 = self.canvas.canvasx(event.x)
        y1 = self.canvas.canvasy(event.y)
        self.canvas.coords(self.ocr_rect_id, x0, y0, x1, y1)

    def end_ocr_select(self, event):
        if not self.ocr_start:
            return

        x0, y0 = self.ocr_start
        x1 = self.canvas.canvasx(event.x)
        y1 = self.canvas.canvasy(event.y)
        self.ocr_start = None

        if self.ocr_rect_id:
            self.canvas.delete(self.ocr_rect_id)
            self.ocr_rect_id = None

        left, right = sorted((x0, x1))
        top, bottom = sorted((y0, y1))

        if right - left < 8 or bottom - top < 8:
            return

        self.run_ocr_for_canvas_rect(left, top, right, bottom)

    def run_ocr_for_canvas_rect(self, left, top, right, bottom):
        if self.preview_pil_img is None:
            return

        img_left = IMAGE_OFFSET
        img_top = IMAGE_OFFSET
        img_right = img_left + self.preview_pil_img.width
        img_bottom = img_top + self.preview_pil_img.height

        crop_left = max(left, img_left)
        crop_top = max(top, img_top)
        crop_right = min(right, img_right)
        crop_bottom = min(bottom, img_bottom)

        if crop_right <= crop_left or crop_bottom <= crop_top:
            return

        crop_box = (
            int(crop_left - img_left),
            int(crop_top - img_top),
            int(crop_right - img_left),
            int(crop_bottom - img_top),
        )

        crop_img = None
        try:
            crop_img = self.preview_pil_img.crop(crop_box)

            # 第一次辨識前，OCR 引擎與模型尚未載入；先顯示載入狀態。
            if not self.ocr_engine.ready:
                self.ocr_status_var.set("OCR載入中")
                self.set_text(self.ocr_text, "OCR載入中...")
                self.root.update_idletasks()
                self.ocr_engine.load()

            if self.ocr_engine.engine_name not in (
                "PaddleOCR",
                "RapidOCR",
                "EasyOCR",
            ):
                self.ocr_status_var.set("OCR載入失敗")
            else:
                self.ocr_status_var.set("OCR辨識中")
                self.set_text(self.ocr_text, "OCR辨識中...")
                self.root.update_idletasks()

            result = self.ocr_engine.recognize(crop_img)

            if self.ocr_engine.engine_name in (
                "PaddleOCR",
                "RapidOCR",
                "EasyOCR",
            ):
                self.ocr_status_var.set(
                    f"OCR辨識完成：{self.ocr_engine.engine_name}"
                )
            else:
                self.ocr_status_var.set("OCR無法使用")

            self.set_text(self.ocr_text, result)

        except Exception as exc:
            self.ocr_status_var.set("OCR辨識失敗")
            self.set_text(self.ocr_text, f"OCR錯誤：{exc}")

        finally:
            if crop_img is not None:
                with suppress(Exception):
                    crop_img.close()
            gc.collect()

    # =====================================================
    # Rename
    # =====================================================
    def update_preview(self, *_args):
        base_parts = [
            self.vars["發文單位"].get(),
            self.vars["文號"].get(),
            self.vars["主旨"].get(),
        ]
        base = "_".join(part for part in base_parts if part)

        extra = [self.vars["收文號碼"].get(), self.vars["收文日期"].get()]
        extra = [part for part in extra if part]

        filename = f"{base}({'_'.join(extra)}).pdf" if extra else f"{base}.pdf"
        self.preview_var.set(safe_pdf_filename(filename))

        prefix = self.vars["增加前名"].get()
        self.prefix_var.set(
            f"{prefix}_{self.state.selected_pdf}"
            if prefix and self.state.selected_pdf
            else ""
        )

    def rename_file(self, new_name):
        new_name = safe_pdf_filename(new_name)

        if not self.state.selected_pdf or not new_name:
            return

        old_path = Path(self.state.folder) / self.state.selected_pdf
        new_path = Path(self.state.folder) / new_name

        if old_path == new_path:
            return

        if new_path.exists():
            messagebox.showerror("錯誤", f"檔案已存在：\n{new_path.name}")
            return

        try:
            self.close_pdf()
            old_path.rename(new_path)

            self.state.selected_pdf = new_path.name
            self.state.current_pdf_path = str(new_path)

            self.open_pdf(str(new_path))
            self.load_pdfs()
            self.select_pdf_in_tree(new_path.name)
            self.show_preview()

        except Exception as exc:
            messagebox.showerror("錯誤", str(exc))

    def rename_pdf(self):
        self.rename_file(self.preview_var.get())

    def rename_prefix(self):
        self.rename_file(self.prefix_var.get())

    # =====================================================
    # Delete / Restore
    # =====================================================
    def delete_pdf(self):
        if not self.state.selected_pdf:
            return

        recycle = Path(self.state.folder) / "_deleted_temp"
        recycle.mkdir(exist_ok=True)

        old_path = Path(self.state.folder) / self.state.selected_pdf
        deleted_path = recycle / self.state.selected_pdf

        if deleted_path.exists():
            messagebox.showerror(
                "錯誤", f"暫存刪除區已有同名檔案：\n{deleted_path.name}"
            )
            return

        try:
            self.clear_current_pdf(clear_canvas=True)
            old_path.rename(deleted_path)
            self.deleted_files.append((deleted_path, old_path))
            self.load_pdfs()

        except Exception as exc:
            messagebox.showerror("錯誤", str(exc))

    def restore_pdf(self):
        if not self.deleted_files:
            return

        deleted_path, original_path = self.deleted_files[-1]

        try:
            if original_path.exists():
                messagebox.showerror(
                    "錯誤", f"原位置已有同名檔案：\n{original_path.name}"
                )
                return

            deleted_path.rename(original_path)
            self.deleted_files.pop()
            self.load_pdfs()

        except Exception as exc:
            messagebox.showerror("錯誤", str(exc))


# =========================================================
# Run
# =========================================================
def run_app():
    if load_tkinterdnd():
        try:
            root = TkinterDnD.Tk()
        except Exception:
            root = None
    else:
        root = None

    try:
        if root is None:
            root = ctk.CTk()
    except Exception:
        root = tk.Tk()

    _app = PDFRenameTool(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        run_app()
    except Exception:
        error_text = traceback.format_exc()

        try:
            log_path = Path(__file__).with_name("pdfname_error_log.txt")
            log_path.write_text(error_text, encoding="utf-8")
        except Exception:
            pass

        try:
            temp_root = tk.Tk()
            temp_root.withdraw()
            messagebox.showerror(
                "程式啟動失敗",
                "程式發生錯誤，已產生 pdfname_error_log.txt。\n\n" + error_text[-2000:],
            )
            temp_root.destroy()
        except Exception:
            print(error_text)
