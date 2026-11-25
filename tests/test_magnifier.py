"""Exercise the real GDI magnifier with a deterministic pixel source."""
import ctypes
import os
from pathlib import Path
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
os.environ['APPDATA'] = str(ROOT / '.test-runtime')
from color_tools import eyedropper as eye


class MagnifierTests(unittest.TestCase):
    def test_physical_pixel_context_restores_after_error(self):
        previous = 0x123456789
        with patch.object(eye, '_pip_set_thread_dpi', return_value=previous) as setter:
            with self.assertRaisesRegex(RuntimeError, 'drawing failed'):
                with eye._pip_physical_pixels():
                    raise RuntimeError('drawing failed')
        self.assertEqual(setter.call_args_list[0].args, (-3,))
        self.assertEqual(setter.call_args_list[-1].args, (previous,))

    def test_physical_pixel_context_without_thread_dpi_api(self):
        with patch.object(eye, '_pip_set_thread_dpi', None):
            with eye._pip_physical_pixels():
                pass

    @unittest.skipUnless(hasattr(ctypes.windll.user32, 'GetWindowDpiAwarenessContext'),
                         'Per-window DPI awareness requires Windows 10')
    def test_overlay_is_per_monitor_aware_and_restores_calling_context(self):
        native = ctypes.WinDLL('user32')
        native.GetThreadDpiAwarenessContext.restype = ctypes.c_void_p
        native.GetWindowDpiAwarenessContext.argtypes = [ctypes.c_void_p]
        native.GetWindowDpiAwarenessContext.restype = ctypes.c_void_p
        native.GetAwarenessFromDpiAwarenessContext.argtypes = [ctypes.c_void_p]
        native.GetAwarenessFromDpiAwarenessContext.restype = ctypes.c_int
        native.AreDpiAwarenessContextsEqual.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
        native.AreDpiAwarenessContextsEqual.restype = ctypes.c_bool
        previous = native.GetThreadDpiAwarenessContext()
        instance = None
        with patch.object(eye, '_pip_hwnd', [None]), patch.object(
                eye, '_pip_cls', [None]), patch.object(
                eye.app, '_pip_wndproc_cb', None, create=True):
            try:
                instance = eye.init_pip_overlay()
                self.assertTrue(eye._pip_hwnd[0])
                context = native.GetWindowDpiAwarenessContext(eye._pip_hwnd[0])
                self.assertEqual(native.GetAwarenessFromDpiAwarenessContext(context), 2)
                self.assertTrue(native.AreDpiAwarenessContextsEqual(
                    native.GetThreadDpiAwarenessContext(), previous))
            finally:
                if eye._pip_hwnd[0]:
                    eye._pip_u32.DestroyWindow(eye._pip_hwnd[0])
                if instance and eye._pip_cls[0]:
                    eye._pip_u32.UnregisterClassW(eye._pip_cls[0], instance)

    def test_cursor_restored_after_repeated_hide(self):
        with patch.object(eye, '_pip_previous_cursor', [None]), patch.object(
                eye._pip_u32, 'SetCursor', side_effect=[123, None, None]) as cursor:
            eye._pip_hide_cursor()
            eye._pip_hide_cursor()
            eye._pip_restore_cursor()
            self.assertEqual(cursor.call_args_list[-1].args, (123,))
            self.assertIsNone(eye._pip_previous_cursor[0])

    def test_overlay_target_stays_on_pointer_at_screen_edges(self):
        position = [0, 0]
        contexts = []
        native = ctypes.WinDLL('user32')
        native.GetThreadDpiAwarenessContext.restype = ctypes.c_void_p
        native.GetAwarenessFromDpiAwarenessContext.argtypes = [ctypes.c_void_p]
        native.GetAwarenessFromDpiAwarenessContext.restype = ctypes.c_int
        def physical_cursor(pointer):
            point = ctypes.cast(pointer, ctypes.POINTER(eye._wt.POINT)).contents
            point.x, point.y = position
            return True
        def record_context(*args):
            contexts.append(native.GetAwarenessFromDpiAwarenessContext(
                native.GetThreadDpiAwarenessContext()))
        with patch.object(eye, '_pip_hwnd', [123]), patch.object(
                eye.app, 'pipette_active', True), patch.object(
                eye._pip_u32, 'PeekMessageW', return_value=0), patch.object(
                eye._pip_u32, 'GetCursorPos', side_effect=AssertionError('Logical cursor used')), patch.object(
                eye._pip_u32, 'GetPhysicalCursorPos', side_effect=physical_cursor), patch.object(
                eye._pip_u32, 'ShowWindow'), patch.object(
                eye._pip_u32, 'SetWindowPos'), patch.object(
                eye, '_pip_hide_cursor'), patch.object(
                eye, '_pip_was_visible', [False]), patch.object(
                eye, '_pip_draw', side_effect=record_context) as draw:
            for x, y in ((0, 0), (-1920, 0), (1919, 1079), (1400, 1800), (3200, -700)):
                position[:] = x, y
                # A stale/logical point must be replaced by the physical point.
                with patch.object(eye, '_pip_pt', eye._wt.POINT(1, 2)):
                    eye._pip_overlay_update()
                ox, oy = draw.call_args.args[:2]
                self.assertEqual(ox + eye._PIP_LENS_X + eye._PIP_LENS_D // 2, x)
                self.assertEqual(oy + eye._PIP_LENS_Y + eye._PIP_LENS_D // 2, y)
            self.assertEqual(contexts, [2] * 5)

    def test_pixels_target_transparency_and_desktop_edge(self):
        eye._pip_init_gdi()
        gdi = eye._pip_gdi
        source = eye._pip_g32.CreateCompatibleDC(gdi['hdc_screen'])
        info = eye._BITMAPINFO()
        info.bmiHeader.biSize = ctypes.sizeof(eye._BITMAPINFOHEADER)
        info.bmiHeader.biWidth = 9
        info.bmiHeader.biHeight = -9
        info.bmiHeader.biPlanes = 1
        info.bmiHeader.biBitCount = 32
        bits = ctypes.c_void_p()
        bitmap = eye._pip_g32.CreateDIBSection(source, ctypes.byref(info), 0,
                                              ctypes.byref(bits), None, 0)
        old = eye._pip_g32.SelectObject(source, bitmap)
        try:
            original = (ctypes.c_uint32 * 81).from_address(bits.value)
            for y in range(9):
                for x in range(9):
                    original[y * 9 + x] = (x * 25 << 16) | (y * 25 << 8) | 60
            output = (ctypes.c_uint32 * (eye._PIP_OW * eye._PIP_OH)).from_address(
                gdi['bits_ptr'].value)
            def pixel(x, y):
                return output[(eye._PIP_LENS_Y + y) * eye._PIP_OW + eye._PIP_LENS_X + x]
            with patch.object(eye._pip_u32, 'GetSystemMetrics',
                              side_effect=lambda key: {76: 0, 77: 0, 78: 9, 79: 9}[key]):
                for cx, cy in ((4, 4), (0, 0), (8, 8)):
                    ctypes.memmove(gdi['bits_ptr'], eye._pip_pixel_template,
                                   ctypes.sizeof(eye._pip_pixel_template))
                    eye._pip_draw_lens(source, gdi['hdc_mem'], gdi['bits_ptr'], cx, cy)
                    center = 4 * eye._PIP_CELL + eye._PIP_CELL // 2
                    self.assertEqual(pixel(center, center), 0xFF000000 | original[cy * 9 + cx])
                    self.assertEqual(pixel(0, 0), 0)
                    self.assertEqual(pixel(4 * eye._PIP_CELL, center), 0xFFFFFFFF)
                    self.assertEqual(pixel(4 * eye._PIP_CELL - 1, center), 0xFF000000)
                    if (cx, cy) == (4, 4):
                        from PIL import Image
                        Image.frombuffer('RGBA', (eye._PIP_OW, eye._PIP_OH),
                                         bytes(output), 'raw', 'BGRA', 0, 1).save(
                                             ROOT / '.test-runtime' / 'magnifier.png')
        finally:
            eye._pip_g32.SelectObject(source, old)
            eye._pip_g32.DeleteObject(bitmap)
            eye._pip_g32.DeleteDC(source)
            eye._pip_free_gdi()
