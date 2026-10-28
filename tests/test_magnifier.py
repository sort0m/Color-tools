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
    def test_cursor_restored_after_repeated_hide(self):
        with patch.object(eye, '_pip_previous_cursor', [None]), patch.object(
                eye._pip_u32, 'SetCursor', side_effect=[123, None, None]) as cursor:
            eye._pip_hide_cursor()
            eye._pip_hide_cursor()
            eye._pip_restore_cursor()
            self.assertEqual(cursor.call_args_list[-1].args, (123,))
            self.assertIsNone(eye._pip_previous_cursor[0])

    def test_overlay_target_stays_on_pointer_at_screen_edges(self):
        with patch.object(eye, '_pip_hwnd', [123]), patch.object(
                eye.app, 'pipette_active', True), patch.object(
                eye._pip_u32, 'PeekMessageW', return_value=0), patch.object(
                eye._pip_u32, 'GetCursorPos'), patch.object(
                eye._pip_u32, 'ShowWindow'), patch.object(
                eye._pip_u32, 'SetWindowPos'), patch.object(
                eye, '_pip_hide_cursor'), patch.object(
                eye, '_pip_was_visible', [False]), patch.object(eye, '_pip_draw') as draw:
            for x, y in ((0, 0), (-1920, 0), (1919, 1079)):
                with patch.object(eye, '_pip_pt', eye._wt.POINT(x, y)):
                    eye._pip_overlay_update()
                ox, oy = draw.call_args.args[:2]
                self.assertEqual(ox + eye._PIP_LENS_X + eye._PIP_LENS_D // 2, x)
                self.assertEqual(oy + eye._PIP_LENS_Y + eye._PIP_LENS_D // 2, y)

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
