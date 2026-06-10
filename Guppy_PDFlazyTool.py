# -*- coding: utf-8 -*-
# =========================================================
# Guppy PDF手搓工具 V0.2.0
# =========================================================
# 程式歷史：
# V0.1.1  建立 Guppy PDF手搓工具，整合更名、搬移、浮水印第三分頁。
# V0.1.2  修正浮水印分頁拖曳 PDF，統一前三分頁基礎配色。
# V0.1.3  整合 PDF旋轉吧 V1.7.1 為第四分頁。
# V0.2.0  第四分頁改名「旋壓合切」，重新統一四個分頁配色並整理整合程式碼。
# V0.2.1  移除第四分頁標語，加入四個小功能回復上一動作，統一功能按鈕風格。
# V0.2.2  第四分頁小工具分頁標籤文字放大。
# V0.2.3  第三、第四分頁功能按鈕改為與第一分頁一致的圓角按鈕。
#
# 建議安裝：
# pip install customtkinter PyMuPDF pillow numpy tkinterdnd2
# pip install paddleocr paddlepaddle
# 備用：
# pip install rapidocr-onnxruntime
# pip install easyocr
# =========================================================

import os
import re
import gc
import sys
import warnings
import traceback
import platform
import subprocess
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import tkinter as tk
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


def write_log(path: Path, text: str) -> None:
    try:
        path.write_text(text, encoding="utf-8")
    except Exception:
        pass


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
            _PYW_STDIO_LOG_HANDLE = (BASE_DIR / "pdfname_stdio_log.txt").open("a", encoding="utf-8", buffering=1)
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


def ensure_required_modules() -> None:
    """Install core GUI/PDF dependencies before importing them.

    This prevents the common .pyw symptom: double-clicking and seeing nothing
    because a top-level import failed before the app's error dialog existed.
    """
    import importlib.util

    required = {
        "customtkinter": "customtkinter",
        "fitz": "PyMuPDF",
        "PIL": "pillow",
        "numpy": "numpy",
    }
    missing_packages = [pkg for module, pkg in required.items() if importlib.util.find_spec(module) is None]
    if not missing_packages:
        return

    append_startup_log("缺少套件：" + ", ".join(missing_packages))
    ok, pip_output = run_pip_install(missing_packages)
    append_startup_log(pip_output[-4000:])

    still_missing = [pkg for module, pkg in required.items() if importlib.util.find_spec(module) is None]
    if not ok or still_missing:
        install_cmd = f'"{sys.executable}" -m pip install ' + " ".join(still_missing or missing_packages)
        raise RuntimeError(
            "程式缺少必要套件，且自動安裝失敗。\n\n"
            f"請手動執行：\n{install_cmd}\n\n"
            f"詳細紀錄：\n{STARTUP_LOG}\n\n"
            + pip_output[-2500:]
        )


warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings("ignore", message=".*RequestsDependencyWarning.*")
warnings.filterwarnings("ignore", message=".*Preferred drawing method.*")

try:
    ensure_required_modules()
    import customtkinter as ctk
    import fitz
    import numpy as np
    from PIL import Image, ImageTk, ImageEnhance, ImageFilter
except Exception:
    show_startup_error("程式啟動失敗", traceback.format_exc())
    raise SystemExit(1)


# =========================================================
# Warning Filter
# =========================================================
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings("ignore", message=".*RequestsDependencyWarning.*")
warnings.filterwarnings("ignore", message=".*Preferred drawing method.*")


# =========================================================
# Theme / Constants
# =========================================================
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

APP_VERSION = "0.2.3"
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

FONT = ("Microsoft JhengHei UI", 11)
TREE_FONT = ("Microsoft JhengHei UI", 16)
BTN_FONT = ("Microsoft JhengHei UI", 11)
LARGE_BTN_FONT = ("Microsoft JhengHei UI", 14)
TURN_TAB_FONT = ("Microsoft JhengHei UI", 20)
TITLE_FONT = ("Microsoft JhengHei UI", 13, "bold")


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
        height=34,
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
    replacements = {
        "中華民國": "",
        "民國": "",
        "年": "/",
        "月": "/",
        "日": "",
        "／": "/",
        "-": "/",
        "－": "/",
        ".": "/",
    }

    for old, new in replacements.items():
        s = s.replace(old, new)

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
    filename = re.sub(r"\s+", " ", filename).strip()
    return filename


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

    key = (lambda item: item[0].lower()) if sort_column == "filename" else (lambda item: item[1])
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
    return str(size)


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
            result.append({
                "name": path.name,
                "path": str(path),
                "is_dir": is_dir,
                "size": 0 if is_dir else stat.st_size,
                "created": get_file_added_time(path, stat),
                "modified": stat.st_mtime,
            })
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

    folders = sorted((item for item in result if item["is_dir"]), key=item_key, reverse=reverse)
    files = sorted((item for item in result if not item["is_dir"]), key=item_key, reverse=reverse)
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

    def load(self):
        if self.ready:
            return

        # 1. PaddleOCR：繁中、數字、英文較準
        try:
            from paddleocr import PaddleOCR

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
                except Exception:
                    continue

        except Exception:
            pass

        # 2. RapidOCR：速度快、易打包
        try:
            from rapidocr_onnxruntime import RapidOCR
            self.rapidocr = RapidOCR()
            self.engine_name = "RapidOCR"
            self.ready = True
            return
        except Exception:
            pass

        # 3. EasyOCR：備用
        try:
            import easyocr
            self.easyocr_reader = easyocr.Reader(["ch_tra", "en"], gpu=False)
            self.engine_name = "EasyOCR"
            self.ready = True
            return
        except Exception:
            pass

        self.engine_name = "未安裝OCR"
        self.ready = True

    @staticmethod
    def preprocess(img: Image.Image) -> Image.Image:
        if img.mode != "RGB":
            img = img.convert("RGB")

        w, h = img.size

        if max(w, h) < 1200:
            img = img.resize((w * 2, h * 2), getattr(Image, 'Resampling', Image).LANCZOS)

        img = ImageEnhance.Contrast(img).enhance(1.55)
        img = ImageEnhance.Sharpness(img).enhance(1.8)
        return img.filter(ImageFilter.SHARPEN)

    def recognize(self, img: Image.Image) -> str:
        if img is None:
            return ""

        self.load()

        if self.engine_name == "未安裝OCR":
            return (
                "尚未安裝 OCR 套件。\n"
                "建議：pip install paddleocr paddlepaddle\n"
                "備用：pip install rapidocr-onnxruntime numpy\n"
                "再備用：pip install easyocr numpy"
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
                return self._parse_paddle_result(result)

            if self.engine_name == "RapidOCR":
                result, _ = self.rapidocr(arr)
                return "\n".join(str(item[1]) for item in result or [] if len(item) >= 2).strip()

            if self.engine_name == "EasyOCR":
                result = self.easyocr_reader.readtext(arr, detail=0, paragraph=True)
                return "\n".join(map(str, result)).strip()

        except Exception as exc:
            return f"OCR失敗：{exc}"

        finally:
            del arr
            if processed_img is not None and processed_img is not img:
                try:
                    processed_img.close()
                except Exception:
                    pass
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
from PIL import ImageDraw, ImageFont

try:
    import importlib.util
    if importlib.util.find_spec("tkinterdnd2") is None:
        append_startup_log("缺少拖曳套件 tkinterdnd2，嘗試自動安裝。")
        ok, pip_output = run_pip_install(["tkinterdnd2"])
        append_startup_log(pip_output[-4000:])
        if not ok:
            append_startup_log("tkinterdnd2 自動安裝失敗，拖曳功能將停用。")
except Exception:
    append_startup_log("檢查 tkinterdnd2 時發生錯誤：\n" + traceback.format_exc())

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    HAS_DND = False

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


def ensure_packages():
    missing = []
    if fitz is None:
        missing.append("pymupdf")
    if Image is None or ImageTk is None or ImageDraw is None or ImageFont is None:
        missing.append("pillow")

    if missing:
        messagebox.showerror(
            "缺少套件",
            "缺少必要套件：\n\n"
            + "\n".join(missing)
            + "\n\n請執行：\n"
            + "pip install " + " ".join(missing)
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
        candidates.extend([
            fonts_dir / "msjh.ttc",        # 微軟正黑體
            fonts_dir / "msjhbd.ttc",      # 微軟正黑體粗體
            fonts_dir / "mingliu.ttc",     # 細明體
            fonts_dir / "kaiu.ttf",        # 標楷體
            fonts_dir / "simsun.ttc",
            fonts_dir / "simhei.ttf",
        ])

    elif sys.platform == "darwin":
        candidates.extend([
            Path("/System/Library/Fonts/PingFang.ttc"),
            Path("/System/Library/Fonts/STHeiti Light.ttc"),
            Path("/Library/Fonts/Arial Unicode.ttf"),
        ])

    else:
        candidates.extend([
            Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"),
            Path("/usr/share/fonts/opentype/noto/NotoSansCJKtc-Regular.otf"),
            Path("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"),
            Path("/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"),
            Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        ])

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


def wrap_text_to_width(text, font, max_width):
    """
    依照像素寬度自動換行。
    支援中文沒有空白的情況，會逐字測量。
    """
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
    lines = wrap_text_to_width(text, font, max_width)

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


def make_watermark_png_bytes(text, width_px, height_px, font_size_px, font_file=None, color=WM_WATERMARK_RGBA_COLOR):
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
            tags=("watermark_box",)
        )
        self.text_id = self.canvas.create_text(
            self.x + 6,
            self.y + 6,
            text=self.preview_text(),
            anchor="nw",
            fill=WM_WATERMARK_HEX_COLOR,
            font=("Microsoft JhengHei", self.font_size, "bold"),
            tags=("watermark_box",)
        )
        self.handles = {}
        self._create_handles()

        self._bind_drag_events()
        self.fit_inside_page()
        self.update_handles()

    def _create_handles(self):
        handle_defs = {
            "n":  "sb_v_double_arrow",
            "s":  "sb_v_double_arrow",
            "e":  "sb_h_double_arrow",
            "w":  "sb_h_double_arrow",
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
                tags=("watermark_handle",)
            )
            self.canvas.tag_bind(handle_id, "<ButtonPress-1>", lambda e, n=name: self.start_resize(e, n))
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
        self.canvas.itemconfigure(self.text_id, font=("Microsoft JhengHei", self.font_size, "bold"))
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
        lines = layout_text_lines(self.text, font, max_text_width, max_text_height, line_gap)
        return "\n".join(lines)

    def refresh_text_layout(self):
        self.canvas.itemconfigure(self.text_id, text=self.preview_text())

    def page_size(self):
        return max(1, int(self.app.page_image_width)), max(1, int(self.app.page_image_height))

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

    def stop_action(self, event=None):
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
        self.canvas.coords(self.box_id, self.x, self.y, self.x + self.width, self.y + self.height)
        self.canvas.coords(self.text_id, self.x + padding_x, self.y + padding_y)
        self.refresh_text_layout()
        self.update_handles()
        self.canvas.configure(scrollregion=(0, 0, self.app.page_image_width, self.app.page_image_height))

    def update_handles(self):
        hs = self.HANDLE_SIZE
        half = hs // 2

        positions = {
            "nw": (self.x - half, self.y - half),
            "n":  (self.x + self.width / 2 - half, self.y - half),
            "ne": (self.x + self.width - half, self.y - half),
            "e":  (self.x + self.width - half, self.y + self.height / 2 - half),
            "se": (self.x + self.width - half, self.y + self.height - half),
            "s":  (self.x + self.width / 2 - half, self.y + self.height - half),
            "sw": (self.x - half, self.y + self.height - half),
            "w":  (self.x - half, self.y + self.height / 2 - half),
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
            (self.y + self.height) / scale
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

        self.watermark_box = None
        self.last_output_path = None

        self.cjk_font_file = find_cjk_font_file()
        self.default_note_template = self.load_default_note_template()
        self.dnd_enabled = False

        self.build_ui()
        self.setup_drag_drop()

        self.add_message(f"程式已啟動。預設字體大小 {WM_DEFAULT_FONT_SIZE}。")
        if self.dnd_enabled:
            self.add_message("拖曳開啟功能已啟用。")
        else:
            self.add_message("拖曳開啟功能未啟用，若需要請安裝：pip install tkinterdnd2")
        if self.cjk_font_file:
            self.add_message(f"偵測到中文字型：{self.cjk_font_file}")
        else:
            self.add_message("未偵測到系統中文字型檔，會使用 Pillow 預設字型，中文可能無法正常顯示。")
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
        self.main_frame = ttk.Frame(self.root, padding=8, style="App.TFrame")
        self.main_frame.pack(fill="both", expand=True)

        self.main_frame.rowconfigure(1, weight=1)
        self.main_frame.columnconfigure(0, weight=1)

        top = ttk.LabelFrame(self.main_frame, text="浮水印設定", padding=8, style="App.TLabelframe")
        top.grid(row=0, column=0, sticky="ew")

        row1 = ttk.Frame(top, style="Card.TFrame")
        row1.pack(fill="x", pady=3)

        ttk.Label(row1, text="輸入註記文字", width=14, style="Card.TLabel").pack(side="left")
        self.note_var = tk.StringVar()
        self.note_entry = ttk.Entry(row1, textvariable=self.note_var)
        self.note_entry.pack(side="left", fill="x", expand=True, padx=6)

        rounded_button(row1, "開啟PDF檔案", self.open_pdf_dialog, accent=True).pack(side="left", padx=3)

        row_template = ttk.Frame(top, style="Card.TFrame")
        row_template.pack(fill="x", pady=3)

        ttk.Label(row_template, text="預設字串", width=14, style="Card.TLabel").pack(side="left")
        self.default_template_var = tk.StringVar(value=self.default_note_template)
        self.default_template_entry = ttk.Entry(row_template, textvariable=self.default_template_var)
        self.default_template_entry.pack(side="left", fill="x", expand=True, padx=6)

        rounded_button(row_template, "作為預設字串", self.save_default_note_template, accent=True).pack(side="left", padx=3)

        row2 = ttk.Frame(top, style="Card.TFrame")
        row2.pack(fill="x", pady=3)

        rounded_button(row2, "確認插入", self.confirm_insert, accent=True).pack(side="left", padx=3)
        rounded_button(row2, "上一頁", self.prev_page).pack(side="left", padx=3)
        rounded_button(row2, "下一頁", self.next_page).pack(side="left", padx=3)

        ttk.Label(row2, text="PDF縮放", style="Card.TLabel").pack(side="left", padx=(16, 3))
        rounded_button(row2, "－", self.zoom_out, width=44).pack(side="left", padx=2)
        rounded_button(row2, "＋", self.zoom_in, width=44).pack(side="left", padx=2)
        rounded_button(row2, "適頁", self.fit_page_width, width=58).pack(side="left", padx=2)

        ttk.Label(row2, text="字體大小", style="Card.TLabel").pack(side="left", padx=(16, 3))

        self.font_size_var = tk.IntVar(value=WM_DEFAULT_FONT_SIZE)
        self.font_size_spin = ttk.Spinbox(
            row2,
            from_=WM_MIN_FONT_SIZE,
            to=WM_MAX_FONT_SIZE,
            increment=1,
            width=6,
            textvariable=self.font_size_var,
            command=self.apply_font_size
        )
        self.font_size_spin.pack(side="left", padx=3)
        self.font_size_spin.bind("<Return>", lambda e: self.apply_font_size())
        self.font_size_spin.bind("<FocusOut>", lambda e: self.apply_font_size())

        rounded_button(row2, "－", lambda: self.change_font_size(-1), width=44).pack(side="left", padx=2)
        rounded_button(row2, "＋", lambda: self.change_font_size(1), width=44).pack(side="left", padx=2)

        self.page_label = ttk.Label(row2, text="尚未開啟 PDF", style="Card.TLabel")
        self.page_label.pack(side="left", padx=12)

        self.tip_label = ttk.Label(
            row2,
            text="拖曳文字可移動；拖曳紅點可拉伸；PDF 可縮放瀏覽",
            style="Muted.TLabel"
        )
        self.tip_label.pack(side="left", padx=12)

        row3 = ttk.Frame(top, style="Card.TFrame")
        row3.pack(fill="x", pady=3)

        rounded_button(row3, "儲存檔案", self.save_pdf, accent=True).pack(side="left", padx=3)
        rounded_button(row3, "列印檔案", self.print_pdf).pack(side="left", padx=3)

        self.status_label = ttk.Label(row3, text="請先開啟 PDF 檔案", style="Muted.TLabel")
        self.status_label.pack(side="left", padx=12)

        self.paned = ttk.PanedWindow(self.main_frame, orient="vertical")
        self.paned.grid(row=1, column=0, sticky="nsew", pady=(8, 0))

        preview = ttk.LabelFrame(self.paned, text="PDF預覽", padding=4, style="App.TLabelframe")
        message_frame = ttk.LabelFrame(self.paned, text="訊息列表", padding=4, style="App.TLabelframe")

        self.paned.add(preview, weight=5)
        self.paned.add(message_frame, weight=1)

        preview.rowconfigure(0, weight=1)
        preview.columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(preview, bg=PREVIEW_BLUE, highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")

        vbar = ttk.Scrollbar(preview, orient="vertical", command=self.canvas.yview)
        vbar.grid(row=0, column=1, sticky="ns")

        hbar = ttk.Scrollbar(preview, orient="horizontal", command=self.canvas.xview)
        hbar.grid(row=1, column=0, sticky="ew")

        self.canvas.configure(
            yscrollcommand=vbar.set,
            xscrollcommand=hbar.set
        )

        self.canvas.bind("<MouseWheel>", self.on_mousewheel)

        message_frame.rowconfigure(0, weight=1)
        message_frame.columnconfigure(0, weight=1)

        self.message_list = tk.Listbox(
            message_frame,
            height=6,
            bg=CARD,
            fg=TEXT,
            font=FONT,
            highlightthickness=1,
            highlightbackground=PREVIEW_BORDER,
            relief="flat",
        )
        self.message_list.grid(row=0, column=0, sticky="nsew")

        msg_vbar = ttk.Scrollbar(message_frame, orient="vertical", command=self.message_list.yview)
        msg_vbar.grid(row=0, column=1, sticky="ns")
        self.message_list.configure(yscrollcommand=msg_vbar.set)

    def add_message(self, text):
        self.message_list.insert("end", text)
        self.message_list.see("end")
        self.status_label.config(text=text)

    def get_visible_area(self):
        try:
            self.root.update_idletasks()
        except Exception:
            pass

        page_w = max(1, int(self.page_image_width))
        page_h = max(1, int(self.page_image_height))
        view_w = int(self.canvas.winfo_width())
        view_h = int(self.canvas.winfo_height())

        if view_w < WM_MIN_USABLE_VIEW_SIZE:
            view_w = int(self.main_frame.winfo_width() or self.root.winfo_width() or page_w)
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
        if not HAS_DND:
            return

        widgets = [self.root, self.root.winfo_toplevel(), self.main_frame, self.canvas]
        for widget in dict.fromkeys(widgets):
            try:
                if not hasattr(widget, "drop_target_register") or not hasattr(widget, "dnd_bind"):
                    continue
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self.on_drop_file)
                self.dnd_enabled = True
            except Exception:
                pass

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
            filetypes=[("PDF 檔案", "*.pdf"), ("所有檔案", "*.*")]
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

        self.canvas.create_image(0, 0, anchor="nw", image=self.page_tk_image, tags=("pdf_page",))
        self.canvas.configure(scrollregion=(0, 0, pix.width, pix.height))

        self.page_label.config(text=f"第 {self.current_page_index + 1} / {len(self.doc)} 頁")

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
        self.render_scale = min(WM_MAX_RENDER_SCALE, self.render_scale + WM_RENDER_SCALE_STEP)
        self.refresh_page_after_zoom(old_scale)

    def zoom_out(self):
        if not self.doc:
            return
        old_scale = self.render_scale
        self.render_scale = max(WM_MIN_RENDER_SCALE, self.render_scale - WM_RENDER_SCALE_STEP)
        self.refresh_page_after_zoom(old_scale)

    def fit_page_width(self):
        if not self.doc:
            return

        try:
            page = self.doc[self.current_page_index]
            available_width = max(300, self.canvas.winfo_width() - 30)
            page_width = page.rect.width
            old_scale = self.render_scale
            self.render_scale = max(WM_MIN_RENDER_SCALE, min(WM_MAX_RENDER_SCALE, available_width / page_width))
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
                font_size=old_box["font_size"]
            )

        self.add_message(f"PDF 預覽縮放：{self.render_scale:.2f}x")

    def confirm_insert(self):
        if not self.doc:
            messagebox.showwarning("尚未開啟 PDF", "請先開啟 PDF 檔案。")
            return

        text = self.build_full_note_text()
        if not text:
            messagebox.showwarning("尚未輸入文字", "請先在「輸入註記文字」欄框輸入內容。")
            return

        font_size = self.get_font_size()

        if self.watermark_box:
            self.watermark_box.set_text(text)
            self.watermark_box.set_font_size(font_size)
        else:
            page_w = max(1, int(self.page_image_width))
            page_h = max(1, int(self.page_image_height))
            view_x, view_y, view_w, view_h = self.get_visible_area()
            margin = WM_PAGE_MARGIN if page_w > WM_PAGE_MARGIN * 2 and page_h > WM_PAGE_MARGIN * 2 else 0
            box_w = min(
                max(WM_DEFAULT_BOX_MIN_WIDTH, int(view_w * WM_DEFAULT_BOX_VISIBLE_WIDTH_RATIO)),
                WM_DEFAULT_BOX_MAX_WIDTH,
                max(1, page_w - margin * 2)
            )
            box_h = min(max(WM_DEFAULT_BOX_HEIGHT, font_size * 5), max(1, page_h - margin * 2))
            box_x = int(min(max(0, view_x + (view_w - box_w) / 2), max(0, page_w - box_w)))
            box_y = int(min(max(0, view_y + (view_h - box_h) / 2), max(0, page_h - box_h)))
            self.watermark_box = SmoothWatermarkBox(
                self,
                box_x,
                box_y,
                box_w,
                box_h,
                text,
                font_size=font_size
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
            filetypes=[("PDF 檔案", "*.pdf"), ("所有檔案", "*.*")]
        )
        if not save_path:
            self.add_message("已取消儲存。")
            return None

        return self.write_watermark_to_pdf(save_path)

    def write_watermark_to_pdf(self, save_path):
        text = self.watermark_box.get_text() if self.watermark_box else ""

        if not text:
            messagebox.showwarning("文字為空", "註記文字框內沒有文字。")
            return None

        if Path(save_path).resolve() == Path(self.pdf_path).resolve():
            messagebox.showwarning("儲存路徑不建議", "請另存為新 PDF，避免覆蓋目前開啟中的原始檔。")
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
                color=WM_WATERMARK_RGBA_COLOR
            )

            page.insert_image(
                rect,
                stream=png_bytes,
                overlay=True,
                keep_proportion=False
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
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.content = ttk.Frame(self.canvas, style="App.TFrame")
        self.window_id = self.canvas.create_window((0, 0), window=self.content, anchor="nw")
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

    def restore_undo_snapshot(self, snapshot):
        return None

    def enable_drop(self, widget, callback):
        if DND_FILES is None:
            return
        try:
            if not hasattr(widget, "drop_target_register") or not hasattr(widget, "dnd_bind"):
                return
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", lambda event: callback(parse_drop_files(event.data), event))
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
        area_width = self.area.canvas.winfo_width() if hasattr(self, "area") else self.winfo_width()
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
        rounded_button(toolbar, "開啟 PDF", self.open_pdf, accent=True).pack(side="left")
        rounded_button(toolbar, "全部左轉", lambda: self.rotate_all(-90)).pack(side="left", padx=4)
        rounded_button(toolbar, "全部右轉", lambda: self.rotate_all(90)).pack(side="left")
        rounded_button(toolbar, "全部重設", self.reset_all).pack(side="left", padx=4)
        rounded_button(toolbar, "回復上一動作", self.undo_last_action).pack(side="left", padx=4)
        ttk.Label(toolbar, text="縮圖", style="App.TLabel").pack(side="left", padx=(16, 4))
        ttk.Scale(toolbar, from_=90, to=230, variable=self.thumb_size, command=lambda _v: self.render()).pack(side="left")
        rounded_button(toolbar, "輸出修改後 PDF", self.export_pdf, accent=True).pack(side="right")
        self.title = ttk.Label(self, text="請開啟或拖曳 PDF 到此分頁", style="App.TLabel")
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
        self.title.configure(text=f"{os.path.basename(self.pdf_path)} / {len(self.pages)} 頁")
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
            card.grid(row=pos // columns, column=pos % columns, padx=6, pady=6, sticky="n")
            self.drag_cards.append(card)
            thumb = make_thumbnail(self.pdf_path, item["index"], self.thumb_size.get(), item["rotation"])
            self.thumbs.append(thumb)
            image_box = ttk.Frame(card, width=self.thumb_size.get(), height=image_box_height, style="Card.TFrame")
            image_box.pack_propagate(False)
            image_box.pack()
            label = ttk.Label(image_box, image=thumb)
            label.place(relx=0.5, rely=0.5, anchor="center")
            self.bind_drag_sort(label, pos)
            page_label = ttk.Label(card, text=f"第 {pos + 1} 頁 / {item['rotation'] % 360}°", style="Card.TLabel")
            page_label.pack(pady=(6, 4))
            buttons = ttk.Frame(card, style="Card.TFrame")
            buttons.pack()
            rounded_button(buttons, "左轉", lambda p=pos: self.rotate_one(p, -90), width=58).pack(side="left")
            rounded_button(buttons, "右轉", lambda p=pos: self.rotate_one(p, 90), width=58).pack(side="left", padx=2)
            rounded_button(buttons, "重設", lambda p=pos: self.reset_one(p), width=58).pack(side="left")
            self.page_widgets.append({"image": label, "label": page_label})

    def get_columns(self):
        return self.columns_for_width(self.thumb_size.get() + 48)

    def update_page_view(self, pos):
        if not self.pdf_path or pos >= len(self.page_widgets):
            return
        item = self.pages[pos]
        thumb = make_thumbnail(self.pdf_path, item["index"], self.thumb_size.get(), item["rotation"])
        self.thumbs[pos] = thumb
        widgets = self.page_widgets[pos]
        widgets["image"].configure(image=thumb)
        widgets["image"].image = thumb
        widgets["label"].configure(text=f"第 {pos + 1} 頁 / {item['rotation'] % 360}°")

    def end_drag(self, event):
        to_pos = self.find_drop_position(event)
        if self.drag_from is not None and to_pos is not None and self.drag_from != to_pos:
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
                result.insert_pdf(source, from_page=item["index"], to_page=item["index"])
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
        rounded_button(toolbar, "開啟 PDF", self.open_pdf, accent=True).pack(side="left")
        rounded_button(toolbar, "回復上一動作", self.undo_last_action).pack(side="left", padx=4)
        ttk.Label(toolbar, text="壓縮比例", style="App.TLabel").pack(side="left", padx=(16, 6))
        self.quality_scale = ttk.Scale(toolbar, from_=35, to=95, variable=self.quality, command=lambda _v: self.update_estimate())
        self.quality_scale.pack(side="left", fill="x", expand=True)
        self.quality_scale.bind("<ButtonPress-1>", lambda _event: self.push_undo())
        rounded_button(toolbar, "輸出壓縮PDF", self.export_pdf, accent=True).pack(side="right", padx=(10, 0))
        self.info = ttk.Label(self, text="請開啟或拖曳 PDF 到此分頁", style="App.TLabel")
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
                new_page = result.new_page(width=page.rect.width, height=page.rect.height)
                new_page.insert_image(rect, filename=str(temp))
                temp.unlink(missing_ok=True)
            result.save(output, garbage=4, deflate=True)
            self.actual_size.set(f"壓縮後實際大小\n{file_size_text(output)}")
            messagebox.showinfo("完成", f"已輸出：\n{output}\n\n實際大小：{file_size_text(output)}")
        except Exception as exc:
            messagebox.showerror("輸出失敗", str(exc))
        finally:
            result.close()
            source.close()


class MergeTab(BaseTab):
    def __init__(self, app, notebook):
        super().__init__(app, notebook)
        self.files = []
        self._build()

    def _build(self):
        toolbar = ttk.Frame(self, style="App.TFrame")
        toolbar.pack(fill="x", pady=(0, 10))
        rounded_button(toolbar, "加入 PDF", self.add_files, accent=True).pack(side="left")
        rounded_button(toolbar, "清空", self.clear).pack(side="left", padx=4)
        rounded_button(toolbar, "回復上一動作", self.undo_last_action).pack(side="left")
        rounded_button(toolbar, "輸出合併 PDF", self.export_pdf, accent=True).pack(side="right")
        self.title = ttk.Label(self, text="請加入或拖曳 PDF 到此分頁", style="App.TLabel")
        self.title.pack(anchor="w", pady=(0, 8))
        self.area = ScrollArea(self)
        self.area.pack(fill="both", expand=True)
        self.enable_responsive_layout(self.area)
        self.enable_drop(self, self.load_drop)

    def add_files(self):
        self.load_drop(self.choose_pdfs())

    def load_drop(self, paths, _event=None):
        if not paths:
            return
        self.push_undo()
        self.files.extend(paths)
        self.render()

    def clear(self):
        if self.files:
            self.push_undo()
        self.files = []
        self.render()

    def get_undo_snapshot(self):
        return list(self.files)

    def restore_undo_snapshot(self, snapshot):
        self.files = list(snapshot)
        self.render()

    def render(self):
        self.clear_frame(self.area.content)
        self.title.configure(text=f"合併清單 / {len(self.files)} 個檔案")
        columns = self.get_columns()
        for pos, path in enumerate(self.files):
            card = ttk.Frame(self.area.content, padding=8, style="Card.TFrame")
            card.grid(row=pos // columns, column=pos % columns, padx=6, pady=6, sticky="n")
            self.drag_cards.append(card)
            thumb = make_thumbnail(path, 0, 150)
            self.thumbs.append(thumb)
            label = ttk.Label(card, image=thumb)
            label.pack()
            self.bind_drag_sort(label, pos)
            ttk.Label(card, text=f"{pos + 1}. {os.path.basename(path)}", wraplength=160, style="Card.TLabel").pack(pady=(6, 4))
            rounded_button(card, "移除", lambda p=pos: self.remove(p), width=64).pack()

    def get_columns(self):
        return self.columns_for_width(198)

    def end_drag(self, event):
        to_pos = self.find_drop_position(event)
        if self.drag_from is not None and to_pos is not None and self.drag_from != to_pos:
            self.push_undo()
        moved = self.move_item(self.files, self.drag_from, to_pos)
        self.drag_from = None
        if moved:
            self.render()

    def remove(self, pos):
        self.push_undo()
        self.files.pop(pos)
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
        self._build()

    def _build(self):
        toolbar = ttk.Frame(self, style="App.TFrame")
        toolbar.pack(fill="x", pady=(0, 10))
        rounded_button(toolbar, "開啟 PDF", self.open_pdf, accent=True).pack(side="left")
        rounded_button(toolbar, "插入 PDF", self.insert_pdf).pack(side="left", padx=4)
        rounded_button(toolbar, "回復上一動作", self.undo_last_action).pack(side="left")
        rounded_button(toolbar, "輸出編輯 PDF", self.export_pdf, accent=True).pack(side="right")
        self.title = ttk.Label(self, text="請開啟 PDF；拖曳其他 PDF 可插入頁面", style="App.TLabel")
        self.title.pack(anchor="w", pady=(0, 8))
        self.area = ScrollArea(self)
        self.area.pack(fill="both", expand=True)
        self.enable_responsive_layout(self.area)
        self.enable_drop(self, self.load_drop)

    def open_pdf(self):
        path = self.choose_pdf()
        if path:
            self.load_base(path)

    def insert_pdf(self):
        self.load_drop(self.choose_pdfs())

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
        self.pages = [{"path": path, "index": i, "inserted": False, "delete": tk.BooleanVar(value=False)} for i in range(count)]
        self.render()

    def get_undo_snapshot(self):
        if not self.base_path:
            return None
        return {
            "base_path": self.base_path,
            "pages": [
                {
                    "path": item["path"],
                    "index": item["index"],
                    "inserted": item["inserted"],
                    "delete": item["delete"].get(),
                }
                for item in self.pages
            ],
        }

    def restore_undo_snapshot(self, snapshot):
        self.base_path = snapshot["base_path"]
        self.pages = [
            {
                "path": item["path"],
                "index": item["index"],
                "inserted": item["inserted"],
                "delete": tk.BooleanVar(value=item["delete"]),
            }
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
            {"path": path, "index": i, "inserted": True, "delete": tk.BooleanVar(value=False)}
            for i in range(count)
        ]
        if insert_at is None:
            self.pages.extend(new_pages)
            return None
        self.pages[insert_at:insert_at] = new_pages
        return insert_at + len(new_pages)

    def render(self):
        self.clear_frame(self.area.content)
        if not self.base_path:
            return
        self.title.configure(text=f"{os.path.basename(self.base_path)} / 目前 {len(self.pages)} 頁")
        columns = self.get_columns()
        for pos, item in enumerate(self.pages):
            color = (80, 160, 255, 70) if item["inserted"] else None
            if item["delete"].get():
                color = (255, 80, 80, 90)
            card = ttk.Frame(self.area.content, padding=8, style="Card.TFrame")
            card.grid(row=pos // columns, column=pos % columns, padx=6, pady=6, sticky="n")
            self.drag_cards.append(card)
            thumb = make_thumbnail(item["path"], item["index"], 150, mark=color)
            self.thumbs.append(thumb)
            label = ttk.Label(card, image=thumb)
            label.pack()
            self.bind_drag_sort(label, pos)
            ttk.Label(card, text=f"第 {pos + 1} 頁", wraplength=150, style="Card.TLabel").pack(pady=(6, 2))
            delete_check = ttk.Checkbutton(card, text="刪除", variable=item["delete"], command=self.render)
            delete_check.pack()
            delete_check.bind("<ButtonPress-1>", lambda _event: self.push_undo(), add="+")

    def get_columns(self):
        return self.columns_for_width(198)

    def end_drag(self, event):
        to_pos = self.find_drop_position(event)
        if self.drag_from is not None and to_pos is not None and self.drag_from != to_pos:
            self.push_undo()
        moved = self.move_item(self.pages, self.drag_from, to_pos)
        self.drag_from = None
        if moved:
            self.render()

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
                result.insert_pdf(open_docs[path], from_page=item["index"], to_page=item["index"])
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
        style.configure("Stat.TLabel", background=CARD, foreground=TEXT, font=("Microsoft JhengHei UI", 16, "bold"), padding=(18, 18))
        style.configure("TurnTabHost.TFrame", background=BG)

    def _build(self):
        self.tab_bar = ttk.Frame(self.root, padding=(12, 12, 12, 4), style="TurnTabHost.TFrame")
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
                fg_color=TURN_TAB_ACTIVE if index == active_index else TURN_TAB_COLORS[index],
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
        self.root.geometry("1600x920")
        self.root.minsize(1180, 720)
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
        self.ocr_start = None
        self.ocr_rect_id = None

        self.company_options = ["中工段", "中興監造", "聖穎", "建業"]

        self.vars = {}
        self.entry_widgets = {}

        self.create_style()
        self.create_ui()

    # =====================================================
    # UI Factory
    # =====================================================
    def create_style(self):
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", rowheight=34, font=TREE_FONT, background="white", fieldbackground="white")
        style.configure("Treeview.Heading", font=("Microsoft JhengHei UI", 13, "bold"))
        style.configure(".", background=BG, foreground=TEXT, font=FONT)
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, foreground=TEXT, font=FONT)
        style.configure("TButton", padding=(10, 5), font=LARGE_BTN_FONT)
        style.configure("TCheckbutton", background=CARD, foreground=TEXT, font=FONT)
        style.configure("App.TFrame", background=BG)
        style.configure("Card.TFrame", background=CARD)
        style.configure("App.TLabelframe", background=CARD, bordercolor=PREVIEW_BORDER, relief="solid")
        style.configure("App.TLabelframe.Label", background=CARD, foreground=TEXT, font=TITLE_FONT)
        style.configure("App.TLabel", background=BG, foreground=TEXT, font=FONT)
        style.configure("Card.TLabel", background=CARD, foreground=TEXT, font=FONT)
        style.configure("Muted.TLabel", background=CARD, foreground=MUTED_TEXT, font=FONT)
        style.configure("Accent.TButton", font=LARGE_BTN_FONT, foreground="white", background=PRIMARY, borderwidth=1)
        style.map("Accent.TButton", background=[("active", PRIMARY_HOVER), ("pressed", PRIMARY_HOVER)])
        style.configure("Soft.TButton", font=LARGE_BTN_FONT, foreground="black", background=PREVIEW_BLUE, borderwidth=1)
        style.map("Soft.TButton", background=[("active", PRIMARY_SOFT_HOVER), ("pressed", PRIMARY_SOFT_HOVER)])

    def button(self, parent, text, command, width=90, color=PRIMARY, hover=PRIMARY_HOVER,
               text_color="white", border=False):
        width = max(width, int(len(text) * 18 + 28))
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            width=width,
            height=34,
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
            height=24,
            font=FONT,
            corner_radius=10,
            fg_color=color,
        )

    def combo(self, parent, var, values):
        return ctk.CTkComboBox(
            parent,
            variable=var,
            values=values,
            height=24,
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
        self.main.grid_columnconfigure(0, weight=1, minsize=500)
        self.main.grid_columnconfigure(1, weight=0, minsize=84)
        self.main.grid_columnconfigure(2, weight=2, minsize=480)
        self.main.grid_columnconfigure(3, weight=0, minsize=72)

        # 左半部：原本更名檔案瀏覽區保留，但可隨視窗縮放
        self.left_shell = tk.Frame(self.main, bg=BG, width=720)
        self.left_shell.grid(row=0, column=0, sticky="nsew", padx=(12, 6), pady=12)

        # 中間：搬移箭頭固定放在左右兩半正中間；更名模式時只隱藏按鈕，不改變版面結構
        self.center_move_bar = tk.Frame(self.main, bg=BG, width=84)
        self.center_move_bar.grid(row=0, column=1, sticky="ns", padx=(0, 0), pady=12)
        self.center_move_bar.grid_propagate(False)

        # 最右側縱向分頁標籤：先固定在右側，避免長檔名把切換標籤擠出視窗
        self.tab_bar = tk.Frame(self.main, bg=BG, width=72)
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

        self.create_folder_area(self.left_shell)
        self.create_treeview(self.left_shell)
        self.create_form(self.left_shell)

        self.create_preview_toolbar(self.rename_page)
        self.create_preview_area(self.rename_page)

        self.create_move_area(self.move_page)
        self.watermark_app = PDFWatermarkApp(self.watermark_page, embedded=True)
        self.turn_app = PDFTurnPanel(self.turn_page)
        self.create_center_move_button()
        self.create_vertical_tabs()
        self.switch_mode("rename")
        self.create_signature_footer()

    def create_signature_footer(self):
        footer = tk.Frame(self.root, bg=BG, height=28)
        footer.pack(side="bottom", fill="x")
        footer.pack_propagate(False)

        tk.Label(
            footer,
            text="Inspired by Atex's high thoughts.",
            bg=BG,
            fg=FOOTER_TEXT,
            font=SIGN_FONT,
        ).pack(side="bottom", pady=(0, 6))

    def create_center_move_button(self):
        """建立左右瀏覽區中央的向右搬移按鈕。
        使用 place 固定在中間偏上位置，避免切換分頁後被 pack 版面擠到下方。
        """
        self.center_move_btn = ctk.CTkButton(
            self.center_move_bar,
            text="→",
            command=self.move_selected_files_to_right,
            width=64,
            height=86,
            corner_radius=20,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="white",
            font=("Microsoft JhengHei UI", 34, "bold"),
        )
        self.center_move_btn.place(relx=0.5, rely=0.38, anchor="center")
        self.center_move_btn.place_forget()

    def create_vertical_tabs(self):
        self.rename_tab_btn = ctk.CTkButton(
            self.tab_bar,
            text="更\n名",
            command=lambda: self.switch_mode("rename"),
            width=52,
            height=135,
            corner_radius=14,
            font=("Microsoft JhengHei UI", 18, "bold"),
        )
        self.rename_tab_btn.pack(pady=(10, 10), padx=8)

        self.move_tab_btn = ctk.CTkButton(
            self.tab_bar,
            text="搬\n移",
            command=lambda: self.switch_mode("move"),
            width=52,
            height=135,
            corner_radius=14,
            font=("Microsoft JhengHei UI", 18, "bold"),
        )
        self.move_tab_btn.pack(pady=(0, 10), padx=8)

        self.watermark_tab_btn = ctk.CTkButton(
            self.tab_bar,
            text="浮\n水\n印",
            command=lambda: self.switch_mode("watermark"),
            width=52,
            height=150,
            corner_radius=14,
            font=("Microsoft JhengHei UI", 17, "bold"),
        )
        self.watermark_tab_btn.pack(pady=(0, 10), padx=8)

        self.turn_tab_btn = ctk.CTkButton(
            self.tab_bar,
            text="旋\n壓\n合\n切",
            command=lambda: self.switch_mode("turn"),
            width=52,
            height=170,
            corner_radius=14,
            font=("Microsoft JhengHei UI", 16, "bold"),
        )
        self.turn_tab_btn.pack(pady=(0, 10), padx=8)

    def switch_mode(self, mode):
        self.current_mode = mode
        self.rename_page.pack_forget()
        self.move_page.pack_forget()
        self.watermark_page.grid_forget()
        self.turn_page.grid_forget()

        inactive_color = TAB_INACTIVE
        self.rename_tab_btn.configure(fg_color=inactive_color, text_color="black")
        self.move_tab_btn.configure(fg_color=inactive_color, text_color="black")
        self.watermark_tab_btn.configure(fg_color=inactive_color, text_color="black")
        self.turn_tab_btn.configure(fg_color=inactive_color, text_color="black")

        if mode in ("watermark", "turn"):
            self.left_shell.grid_forget()
            self.center_move_bar.grid_forget()
            self.right_shell.grid_forget()
            self.center_move_btn.place_forget()
            if mode == "watermark":
                self.watermark_page.grid(row=0, column=0, columnspan=3, sticky="nsew", padx=(12, 6), pady=12)
                self.watermark_tab_btn.configure(fg_color=PRIMARY, text_color="white")
            else:
                self.turn_page.grid(row=0, column=0, columnspan=3, sticky="nsew", padx=(12, 6), pady=12)
                self.turn_tab_btn.configure(fg_color=PRIMARY, text_color="white")
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
        frame.pack(fill="x", pady=(0, 10))
        frame.configure(padx=12, pady=12)

        tk.Label(frame, text="選擇資料夾", bg=CARD, fg=TEXT, font=TITLE_FONT).pack(anchor="w", pady=(0, 8))

        row = tk.Frame(frame, bg=CARD)
        row.pack(fill="x")

        self.folder_var = tk.StringVar()
        self.entry(row, self.folder_var).pack(side="left", fill="x", expand=True, padx=(0, 8))

        buttons = [
            ("瀏覽資料夾", self.browse_folder, 110, PRIMARY, PRIMARY_HOVER, "white"),
            ("刪除檔案", self.delete_pdf, 100, RED, RED_HOVER, "white"),
            ("回復刪除", self.restore_pdf, 100, YELLOW, YELLOW_HOVER, "black"),
        ]

        for text, cmd, width, color, hover, text_color in buttons:
            self.button(row, text, cmd, width, color, hover, text_color).pack(side="left", padx=2)

    def create_treeview(self, parent):
        frame = self.card(parent)
        frame.pack(fill="both", expand=True, pady=(0, 10))

        columns = ("no", "filename", "date")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings")

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
        self.tree.column("date", width=180, minwidth=160, anchor="center", stretch=False)

        scroll = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll.set, xscrollcommand=x_scroll.set)

        scroll.pack(side="right", fill="y")
        x_scroll.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.select_pdf)

    def create_form(self, parent):
        outer = self.card(parent)
        outer.pack(fill="x")
        outer.configure(padx=12, pady=12)

        form = tk.Frame(outer, bg=CARD)
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
            tk.Label(form, text=title, bg=CARD, fg=TEXT, font=FONT).grid(row=row, column=0, sticky="w", pady=4)

            var = tk.StringVar()
            self.vars[title] = var

            if kind == "combo":
                widget = self.combo(form, var, self.company_options)
                widget.set("")
            else:
                widget = self.entry(form, var)
                widget.bind("<Button-1>", lambda _event, field=title: self.fill_field_from_ocr(field))

            self.entry_widgets[title] = widget
            widget.grid(row=row, column=1, sticky="ew", padx=8)
            var.trace_add("write", self.update_preview)

        form.columnconfigure(1, weight=1)
        self.create_bottom_rename_area(outer)

    def create_bottom_rename_area(self, parent):
        bottom = tk.Frame(parent, bg=PRIMARY_SOFT, highlightbackground=PREVIEW_BORDER, highlightthickness=1)
        bottom.pack(fill="x", pady=(14, 0))
        bottom.configure(padx=10, pady=10)

        self.preview_var = tk.StringVar()
        self.prefix_var = tk.StringVar()

        rows = [
            ("變更檔名", self.preview_var, FILENAME_BG, "更名確認", self.rename_pdf, 0),
            ("前名調整", self.prefix_var, PREFIX_BG, "增名確認", self.rename_prefix, 1),
        ]

        for label, var, color, btn_text, cmd, row in rows:
            pady = (10, 0) if row else (0, 6)

            tk.Label(bottom, text=label, bg=PRIMARY_SOFT, fg=TEXT, font=FONT).grid(
                row=row, column=0, sticky="w", pady=pady
            )

            self.entry(bottom, var, color).grid(row=row, column=1, sticky="ew", padx=8, pady=(10, 0) if row else 0)

            self.button(bottom, btn_text, cmd, width=110 if row else 100).grid(
                row=row, column=2, padx=4, pady=(10, 0) if row else 0
            )

        bottom.columnconfigure(1, weight=1)

    def create_preview_toolbar(self, parent):
        bar = tk.Frame(parent, bg=CARD, height=56)
        bar.pack(fill="x", pady=(0, 10))
        bar.pack_propagate(False)

        left = tk.Frame(bar, bg=CARD)
        left.pack(side="left", padx=10)

        for text, cmd, width in (("－", self.zoom_out, 45), ("＋", self.zoom_in, 45), ("整頁", self.fit_page, 70)):
            self.preview_button(left, text, cmd, width).pack(side="left", padx=2, pady=10)

        self.page_var = tk.StringVar(value="0 / 0")
        tk.Label(left, textvariable=self.page_var, bg=CARD, fg=TEXT, font=FONT).pack(side="left", padx=(18, 8))

        for text, cmd in (("上一頁", self.prev_page), ("下一頁", self.next_page)):
            self.preview_button(left, text, cmd, 90).pack(side="left", padx=2)

        right = tk.Frame(bar, bg=CARD)
        right.pack(side="right", padx=10)

        self.ocr_status_var = tk.StringVar(value="OCR：尚未載入")
        tk.Label(right, textvariable=self.ocr_status_var, bg=CARD, fg=TEXT, font=FONT).pack(side="left", padx=(0, 10))

        self.ocr_check = ctk.CTkCheckBox(
            right,
            text="框選辨識",
            variable=self.ocr_select_mode,
            command=self.toggle_ocr_mode,
            font=BTN_FONT,
        )
        self.ocr_check.pack(side="left", padx=2)

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

        tk.Label(title_row, text="測試辨識字串", bg=CARD, fg=TEXT, font=TITLE_FONT).pack(side="left")
        self.preview_button(title_row, "清除", self.clear_ocr_text, 70).pack(side="right")

        self.ocr_text = tk.Text(frame, height=3, font=FONT, bg=OCR_BG, fg=TEXT, wrap="word", relief="solid", bd=1)
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
            try:
                self.preview_pil_img.close()
            except Exception:
                pass

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
        top.pack(fill="x", pady=(0, 10))
        top.configure(padx=12, pady=12)

        tk.Label(top, text="搬移目的資料夾", bg=CARD, fg=TEXT, font=TITLE_FONT).pack(anchor="w", pady=(0, 8))

        row = tk.Frame(top, bg=CARD)
        row.pack(fill="x")

        self.move_folder_var = tk.StringVar()
        self.move_folder_combo = ctk.CTkComboBox(
            row,
            variable=self.move_folder_var,
            values=[],
            height=28,
            font=FONT,
            corner_radius=10,
            state="normal",
            command=self.on_move_folder_combo,
        )
        self.move_folder_combo.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.move_folder_combo.bind("<Return>", lambda _event: self.set_move_folder(self.move_folder_var.get()))

        self.button(row, "瀏覽", self.browse_move_folder, 80).pack(side="left", padx=2)
        self.button(row, "上一層", self.move_parent_folder, 85, color=PREVIEW_BLUE, hover=PRIMARY_SOFT_HOVER, text_color="black", border=True).pack(side="left", padx=2)
        self.button(row, "回復移動", self.undo_last_move, 100, color=YELLOW, hover=YELLOW_HOVER, text_color="black").pack(side="left", padx=2)

        body = self.card(parent)
        body.pack(fill="both", expand=True)

        # 瀏覽區本體
        tree_area = tk.Frame(body, bg=CARD)
        tree_area.pack(fill="both", expand=True)

        columns = ("name", "size", "created", "modified")
        self.move_tree = ttk.Treeview(tree_area, columns=columns, show="headings")
        headers = {
            "name": "檔名",
            "size": "檔案大小",
            "created": "加入時間",
            "modified": "修改時間",
        }
        for col, title in headers.items():
            self.move_tree.heading(col, text=title, command=lambda c=col: self.sort_move_tree(c))

        self.move_tree.column("name", width=360, minwidth=220, stretch=True)
        self.move_tree.column("size", width=110, minwidth=90, anchor="e", stretch=False)
        self.move_tree.column("created", width=170, minwidth=150, anchor="center", stretch=False)
        self.move_tree.column("modified", width=170, minwidth=150, anchor="center", stretch=False)

        y_scroll = ttk.Scrollbar(tree_area, orient="vertical", command=self.move_tree.yview)
        self.move_tree.configure(yscrollcommand=y_scroll.set)
        self.move_tree.pack(side="left", fill="both", expand=True)
        y_scroll.pack(side="right", fill="y")
        self.move_tree.bind("<Double-1>", self.enter_selected_move_folder)

        # 左右移動欄固定放在搬移瀏覽區正下方
        x_scroll = ttk.Scrollbar(body, orient="horizontal", command=self.move_tree.xview)
        self.move_tree.configure(xscrollcommand=x_scroll.set)
        x_scroll.pack(fill="x", padx=0, pady=(2, 0))

        bottom = self.card(parent)
        bottom.pack(fill="x", pady=(10, 0))
        bottom.configure(padx=12, pady=12)

        self.move_target_var = tk.StringVar(value="目的地：尚未選擇")
        tk.Label(bottom, textvariable=self.move_target_var, bg=CARD, fg=TEXT, font=FONT).pack(side="left", fill="x", expand=True)

    def browse_move_folder(self):
        folder = filedialog.askdirectory(initialdir=self.move_folder or self.state.folder or None)
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
        for item in list_directory_items(self.move_folder, self.move_sort_column, self.move_sort_reverse):
            icon_name = f"📁 {item['name']}" if item["is_dir"] else f"📄 {item['name']}"
            size_text = "<資料夾>" if item["is_dir"] else format_file_size(item["size"])
            created = format_timestamp(item["created"])
            modified = format_timestamp(item["modified"])
            self.move_tree.insert("", "end", values=(icon_name, size_text, created, modified), tags=(item["path"], "dir" if item["is_dir"] else "file"))

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

    def get_selected_left_pdf_path(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("提醒", "請先在左側檔案瀏覽區選擇要搬移的 PDF。")
            return None
        values = self.tree.item(selected[0], "values")
        if not values:
            return None
        src = Path(self.state.folder) / values[1]
        if not src.exists():
            messagebox.showerror("錯誤", f"找不到檔案：\n{src}")
            return None
        return src

    def get_move_destination_folder(self):
        selected_path = self.get_selected_move_path()
        if selected_path and selected_path.is_dir():
            return selected_path
        if self.move_folder:
            return Path(self.move_folder)
        return None

    def move_selected_files_to_right(self):
        src = self.get_selected_left_pdf_path()
        dst_folder = self.get_move_destination_folder()
        if not src or not dst_folder:
            messagebox.showwarning("提醒", "請先選擇右側目的資料夾。")
            return

        dst = dst_folder / src.name
        if dst.exists():
            messagebox.showerror("錯誤", f"目的資料夾已有同名檔案：\n{dst.name}")
            return

        try:
            if self.state.current_pdf_path and Path(self.state.current_pdf_path) == src:
                self.clear_current_pdf(clear_canvas=True)

            src.rename(dst)
            self.move_history.append((dst, src))
            self.refresh_file_views()
            self.move_target_var.set(f"已搬移到：{dst_folder}")
        except Exception as exc:
            messagebox.showerror("搬移失敗", str(exc))

    def undo_last_move(self):
        if not self.move_history:
            messagebox.showinfo("提醒", "目前沒有可回復的搬移步驟。")
            return

        moved_path, original_path = self.move_history[-1]
        if not moved_path.exists():
            messagebox.showerror("錯誤", f"找不到已搬移的檔案：\n{moved_path}")
            return
        if original_path.exists():
            messagebox.showerror("錯誤", f"原位置已有同名檔案：\n{original_path.name}")
            return

        try:
            moved_path.rename(original_path)
            self.move_history.pop()
            self.refresh_file_views()
            self.move_target_var.set(f"已回復：{original_path}")
        except Exception as exc:
            messagebox.showerror("回復失敗", str(exc))

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

        value = normalize_receive_date(text) if field == "收文日期" else clean_one_line(text)
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

    def load_pdfs(self):
        self.tree.delete(*self.tree.get_children())

        pdfs = list_pdf_files(self.state.folder, self.state.sort_column, self.state.sort_reverse)

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

        values = self.tree.item(selected[0], "values")
        if not values:
            return

        self.state.selected_pdf = values[1]
        self.state.current_pdf_path = str(Path(self.state.folder) / self.state.selected_pdf)
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

            self.canvas.create_image(IMAGE_OFFSET, IMAGE_OFFSET, anchor="nw", image=self.preview_img, tags=("pdf_image",))
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
        canvas_width = max(self.canvas.winfo_width(), 820)
        self.state.zoom = max(MIN_ZOOM, (canvas_width - 80) / page.rect.width)
        self.show_preview()

    def zoom_in(self):
        self.state.zoom = min(MAX_ZOOM, self.state.zoom + 0.1)
        self.show_preview()

    def zoom_out(self):
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
        self.ocr_rect_id = self.canvas.create_rectangle(x, y, x, y, outline=PRIMARY, width=2, dash=(4, 2))

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

            self.set_text(self.ocr_text, "辨識中...")
            self.root.update_idletasks()

            result = self.ocr_engine.recognize(crop_img)
            self.ocr_status_var.set(f"OCR：{self.ocr_engine.engine_name}")

            self.set_text(self.ocr_text, result)

        except Exception as exc:
            self.set_text(self.ocr_text, f"OCR錯誤：{exc}")

        finally:
            if crop_img is not None:
                try:
                    crop_img.close()
                except Exception:
                    pass
            gc.collect()

    # =====================================================
    # Rename
    # =====================================================
    def update_preview(self, *_args):
        base_parts = [self.vars["發文單位"].get(), self.vars["文號"].get(), self.vars["主旨"].get()]
        base = "_".join(part for part in base_parts if part)

        extra = [self.vars["收文號碼"].get(), self.vars["收文日期"].get()]
        extra = [part for part in extra if part]

        filename = f"{base}({'_'.join(extra)}).pdf" if extra else f"{base}.pdf"
        self.preview_var.set(safe_pdf_filename(filename))

        prefix = self.vars["增加前名"].get()
        self.prefix_var.set(f"{prefix}_{self.state.selected_pdf}" if prefix and self.state.selected_pdf else "")

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
            messagebox.showerror("錯誤", f"暫存刪除區已有同名檔案：\n{deleted_path.name}")
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
                messagebox.showerror("錯誤", f"原位置已有同名檔案：\n{original_path.name}")
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
    if HAS_DND:
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

    app = PDFRenameTool(root)
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
                "程式發生錯誤，已產生 pdfname_error_log.txt。\n\n"
                + error_text[-2000:]
            )
            temp_root.destroy()
        except Exception:
            print(error_text)
