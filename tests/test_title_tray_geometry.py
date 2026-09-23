"""Exercise the Windows geometry without importing GUI/optional dependencies."""

import ast
import ctypes
from ctypes import wintypes
from pathlib import Path
from types import SimpleNamespace
import unittest


source = Path(__file__).resolve().parents[1] / "Guppy_PDFlazyTool.py"
tree = ast.parse(source.read_text(encoding="utf-8-sig"))
node = next(n for n in tree.body if isinstance(n, ast.ClassDef)
            and n.name == "NativeTitleTrayButton")
namespace = {}
exec(compile(ast.Module(body=[node], type_ignores=[]), str(source), "exec"), namespace)
Control = namespace["NativeTitleTrayButton"]


class CaptionGeometryTests(unittest.TestCase):
    def test_adjacent_at_multiple_scales_and_monitor_origins(self):
        for scale in (1, 1.25, 1.5, 2):
            for origin in ((100, 80), (-1920, -1080), (0, 0)):
                with self.subTest(scale=scale, origin=origin):
                    width, height = int(46 * scale), int(30 * scale)
                    x, y = origin
                    window = (x, y, x + 1600, y + 1000)
                    buttons = (1600 - 3 * width, 8, 1600, 8 + height)
                    actual = Control.caption_geometry(window, buttons)
                    self.assertEqual(actual, (x + buttons[0] - width, y + 8, width, height))
                    self.assertEqual(actual[0] + actual[2], x + buttons[0])

    def test_invalid_bounds(self):
        for bounds in ((0, 0, 0, 0), (200, 30, 100, 50),
                       (862, 10, 1000, 10), (862, -1, 1000, 30),
                       (900, 8, 1100, 38), (10, 8, 148, 38),
                       (862, 700, 1000, 900)):
            with self.subTest(bounds=bounds):
                self.assertIsNone(Control.caption_geometry((0, 0, 1000, 800), bounds))

    def make_control(self, bounds=(862, 8, 1000, 38), result=0, window_ok=True):
        control = Control.__new__(Control)
        control.ctypes, control.wt, control.hwnd = ctypes, wintypes, 1
        def write_rect(pointer, values):
            rect = ctypes.cast(pointer, ctypes.POINTER(wintypes.RECT)).contents
            rect.left, rect.top, rect.right, rect.bottom = values
        def dwm(hwnd, attribute, pointer, size):
            self.assertEqual(attribute, 5)
            self.assertEqual(size, ctypes.sizeof(wintypes.RECT))
            write_rect(pointer, bounds)
            return result
        def window_rect(hwnd, pointer):
            write_rect(pointer, (-1000, 50, 0, 850))
            return window_ok
        control.get_dwm_attribute = dwm
        control.u = SimpleNamespace(GetWindowRect=window_rect)
        control.get_minimize_geometry = lambda: (10, 20, 46, 30)
        return control

    def test_dwm_preferred_and_screen_coordinates(self):
        self.assertEqual(self.make_control().get_caption_geometry(), (-184, 58, 46, 30))

    def test_fallback_for_failed_or_invalid_dwm(self):
        for control in (self.make_control(result=-1),
                        self.make_control(bounds=(0, 0, 0, 0)),
                        self.make_control(window_ok=False)):
            self.assertEqual(control.get_caption_geometry(), (10, 20, 46, 30))
        control = self.make_control()
        control.get_dwm_attribute = None
        self.assertEqual(control.get_caption_geometry(), (10, 20, 46, 30))
        control.get_minimize_geometry = lambda: None
        self.assertIsNone(control.get_caption_geometry())


if __name__ == "__main__":
    unittest.main()
