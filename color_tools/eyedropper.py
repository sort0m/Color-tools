"""
eyedropper.py — Pipetti (eyedropper), Win32-overlay-ikkuna, GDI-piirto.
"""

import ctypes
import ctypes.wintypes as _wt
import threading
import time as _time


from .bootstrap import _sc, _DPI_SCALE, _write_log
from .state import app, _pip_lock

# ── Message-pumping sleep ─────────────────────────────────────────────────────
_mwfmo_user32 = ctypes.windll.user32
_mwfmo_user32.PeekMessageW.restype  = ctypes.c_bool
_mwfmo_user32.PeekMessageW.argtypes = [
    ctypes.c_void_p, ctypes.c_void_p,
    ctypes.c_uint, ctypes.c_uint, ctypes.c_uint]
_mwfmo_user32.TranslateMessage.restype  = ctypes.c_bool
_mwfmo_user32.TranslateMessage.argtypes = [ctypes.c_void_p]
_mwfmo_user32.DispatchMessageW.restype  = ctypes.c_ssize_t
_mwfmo_user32.DispatchMessageW.argtypes = [ctypes.c_void_p]
_mwfmo_user32.MsgWaitForMultipleObjects.restype  = ctypes.c_ulong
_mwfmo_user32.MsgWaitForMultipleObjects.argtypes = [
    ctypes.c_ulong, ctypes.c_void_p, ctypes.c_bool,
    ctypes.c_ulong, ctypes.c_ulong]

_PM_REMOVE   = 0x0001
_QS_ALLINPUT = 0x04FF
_MSG_BUF     = _wt.MSG()

def _pumping_sleep(seconds):
    """Sleep `seconds` samalla pumppaamalla Win32-viestijono."""
    deadline = _time.perf_counter() + seconds
    while True:
        remaining_ms = int((deadline - _time.perf_counter()) * 1000)
        if remaining_ms <= 0:
            break
        ret = _mwfmo_user32.MsgWaitForMultipleObjects(
            0, None, False,
            ctypes.c_ulong(remaining_ms),
            ctypes.c_ulong(_QS_ALLINPUT))
        if ret == 0:
            while _mwfmo_user32.PeekMessageW(
                    ctypes.byref(_MSG_BUF), None, 0, 0, _PM_REMOVE):
                _mwfmo_user32.TranslateMessage(ctypes.byref(_MSG_BUF))
                _mwfmo_user32.DispatchMessageW(ctypes.byref(_MSG_BUF))
        else:
            break

# ── Pipetti-säie ──────────────────────────────────────────────────────────────
def _pipette_run():
    app.pipette_active = True

    WH_MOUSE_LL    = 14
    WM_MOUSEMOVE   = 0x0200
    WM_LBUTTONDOWN = 0x0201
    WM_LBUTTONUP   = 0x0202
    WM_RBUTTONDOWN = 0x0204
    WM_RBUTTONUP   = 0x0205
    WM_QUIT        = 0x0012
    VK_ESCAPE      = 0x1B

    class _MSLLHOOKSTRUCT(ctypes.Structure):
        _fields_ = [
            ('x',           ctypes.c_long),
            ('y',           ctypes.c_long),
            ('mouseData',   ctypes.c_ulong),
            ('flags',       ctypes.c_ulong),
            ('time',        ctypes.c_ulong),
            ('dwExtraInfo', ctypes.c_size_t),
        ]

    _user32  = ctypes.windll.user32
    _kernel32 = ctypes.windll.kernel32
    _tid     = _kernel32.GetCurrentThreadId()
    app._pipette_thread_tid = _tid
    _hook_id = [None]

    _HOOKPROC = ctypes.WINFUNCTYPE(
        ctypes.c_longlong,
        ctypes.c_int,
        ctypes.c_size_t,
        ctypes.c_ssize_t,
    )

    _pt0 = _wt.POINT()
    _user32.GetCursorPos(ctypes.byref(_pt0))
    _pip_pos   = [_pt0.x, _pt0.y]          # aloita kursorin kohdasta, ei (0, 0):sta
    _pip_event = threading.Event()
    _selected_pos = [None]

    def _post_quit():
        _user32.PostThreadMessageW(_tid, WM_QUIT, 0, 0)

    def _cancel():
        app.pipette_active = False
        _pip_event.set()
        _post_quit()

    def _capture_loop():
        _gdi32  = ctypes.windll.gdi32
        _u32cap = ctypes.windll.user32
        _u32cap.GetDC.restype                = ctypes.c_void_p
        _u32cap.GetDC.argtypes               = [ctypes.c_void_p]
        _u32cap.ReleaseDC.restype            = ctypes.c_int
        _u32cap.ReleaseDC.argtypes           = [ctypes.c_void_p, ctypes.c_void_p]
        _gdi32.CreateCompatibleDC.restype    = ctypes.c_void_p
        _gdi32.CreateCompatibleDC.argtypes   = [ctypes.c_void_p]
        _gdi32.CreateCompatibleBitmap.restype  = ctypes.c_void_p
        _gdi32.CreateCompatibleBitmap.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int]
        _gdi32.SelectObject.restype          = ctypes.c_void_p
        _gdi32.SelectObject.argtypes         = [ctypes.c_void_p, ctypes.c_void_p]
        _gdi32.BitBlt.restype                = ctypes.c_bool
        _gdi32.BitBlt.argtypes               = [
            ctypes.c_void_p, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
            ctypes.c_void_p, ctypes.c_int, ctypes.c_int, ctypes.c_ulong,
        ]
        _gdi32.GetPixel.restype              = ctypes.c_ulong
        _gdi32.GetPixel.argtypes             = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int]
        _gdi32.DeleteObject.restype          = ctypes.c_bool
        _gdi32.DeleteObject.argtypes         = [ctypes.c_void_p]
        _gdi32.DeleteDC.restype              = ctypes.c_bool
        _gdi32.DeleteDC.argtypes             = [ctypes.c_void_p]
        hdc_screen = _u32cap.GetDC(0)
        hdc_mem    = _gdi32.CreateCompatibleDC(hdc_screen)
        hbmp_1px = _gdi32.CreateCompatibleBitmap(hdc_screen, 1, 1)
        _old_bmp = _gdi32.SelectObject(hdc_mem, hbmp_1px)
        try:
            if not all((hdc_screen, hdc_mem, hbmp_1px, _old_bmp)):
                _write_log('Eyedropper: could not allocate capture resources')
                _cancel()
                return
            while True:
                _pip_event.wait(timeout=0.05)
                _pip_event.clear()
                selected = _selected_pos[0]
                if selected is None and not app.pipette_active:
                    break
                if selected is None and _u32cap.GetAsyncKeyState(VK_ESCAPE) & 0x8000:
                    _cancel()
                    break
                x, y = selected if selected is not None else tuple(_pip_pos)
                try:
                    if not _gdi32.BitBlt(hdc_mem, 0, 0, 1, 1, hdc_screen, x, y, 0xCC0020):
                        raise OSError('BitBlt failed')
                    pixel = _gdi32.GetPixel(hdc_mem, 0, 0)
                    if pixel != 0xFFFFFFFF:
                        r2 = pixel & 0xFF
                        g2 = (pixel >> 8) & 0xFF
                        b2 = (pixel >> 16) & 0xFF
                        with _pip_lock:
                            app.pip_color = [r2, g2, b2, 255]
                        if selected is not None:
                            app._pending_pipette_color = [r2 / 255., g2 / 255., b2 / 255., 1.]
                    elif selected is not None:
                        _write_log('Eyedropper: GetPixel failed for selected point')
                except Exception as exc:
                    if selected is not None:
                        _write_log(f'Eyedropper final sample failed: {exc}')
                if selected is not None:
                    break
        finally:
            _gdi32.SelectObject(hdc_mem, _old_bmp)
            _gdi32.DeleteObject(hbmp_1px)
            _gdi32.DeleteDC(hdc_mem)
            _u32cap.ReleaseDC(0, hdc_screen)

    capture_thread = threading.Thread(target=_capture_loop, daemon=True)
    capture_thread.start()

    _await_up     = [0]      # odotettava vapautusviesti (nielaistaan: ei vuoda alla olevalle ikkunalle)
    _finish_timer = [None]

    def _hook_proc(nCode, wParam, lParam):
        if nCode >= 0:
            if wParam == WM_MOUSEMOVE:
                ms = ctypes.cast(lParam, ctypes.POINTER(_MSLLHOOKSTRUCT)).contents
                _pip_pos[0] = ms.x
                _pip_pos[1] = ms.y
                _pip_event.set()
            elif wParam in (WM_LBUTTONDOWN, WM_RBUTTONDOWN):
                if app.pipette_active:
                    if wParam == WM_LBUTTONDOWN:
                        ms = ctypes.cast(lParam, ctypes.POINTER(_MSLLHOOKSTRUCT)).contents
                        _selected_pos[0] = (ms.x, ms.y)
                    app.pipette_active = False
                    _pip_event.set()
                    _await_up[0] = wParam + 1
                    # Varalukko: jos vapautusviesti ei tule, hook poistuu silti
                    _finish_timer[0] = threading.Timer(1.0, _post_quit)
                    _finish_timer[0].daemon = True
                    _finish_timer[0].start()
                return 1
            elif wParam in (WM_LBUTTONUP, WM_RBUTTONUP):
                if _await_up[0]:
                    if wParam == _await_up[0]:
                        _post_quit()
                    return 1
        return _user32.CallNextHookEx(_hook_id[0], nCode, wParam, lParam)

    _cb = _HOOKPROC(_hook_proc)
    app._pipette_hook_cb = _cb
    _user32.SetWindowsHookExW.argtypes = [
        ctypes.c_int, _HOOKPROC, ctypes.c_void_p, ctypes.c_uint
    ]
    _user32.SetWindowsHookExW.restype = ctypes.c_void_p
    _user32.CallNextHookEx.restype  = ctypes.c_longlong
    _user32.CallNextHookEx.argtypes = [
        ctypes.c_void_p, ctypes.c_int, ctypes.c_size_t, ctypes.c_ssize_t,
    ]

    _user32.UnhookWindowsHookEx.argtypes = [ctypes.c_void_p]
    _user32.UnhookWindowsHookEx.restype = ctypes.c_bool
    _msg = _wt.MSG()
    try:
        _hook_id[0] = _user32.SetWindowsHookExW(WH_MOUSE_LL, _cb, None, 0)
        if not _hook_id[0]:
            _write_log(f"Eyedropper: SetWindowsHookExW failed (error {_kernel32.GetLastError()})")
            return
        while _user32.GetMessageW(ctypes.byref(_msg), None, 0, 0) > 0:
            _user32.TranslateMessage(ctypes.byref(_msg))
            _user32.DispatchMessageW(ctypes.byref(_msg))
    finally:
        app.pipette_active = False
        _pip_event.set()
        if _finish_timer[0] is not None:
            _finish_timer[0].cancel()
        capture_thread.join()
        if _hook_id[0]:
            _user32.UnhookWindowsHookEx(_hook_id[0])
            _hook_id[0] = None
        app._pipette_thread_tid = 0

def pipette_thread():
    """Pipettisäikeen kääre: lippu nollautuu aina, vaikka toteutus kaatuisi."""
    try:
        _pipette_run()
    except Exception:
        import traceback
        _write_log("pipette_thread error:\n" + traceback.format_exc())
    finally:
        app.pipette_active = False
        app._pipette_thread_tid = 0

_pipette_lock = threading.Lock()

def start_pipette():
    """Käynnistää pipetti-säikeen. Suojattu useita samanaikaisia instansseja vastaan."""
    if app.pipette_active:
        return
    if not _pipette_lock.acquire(blocking=False):
        return
    def _run():
        try:
            pipette_thread()
        finally:
            _pipette_lock.release()
    threading.Thread(target=_run, daemon=True).start()

# ── Win32 overlay-ikkuna ──────────────────────────────────────────────────────
_pip_u32 = ctypes.windll.user32
_pip_g32 = ctypes.windll.gdi32
_PIP_OW       = max(130, int(130 * _DPI_SCALE))
_PIP_OH       = max(28,  int(28  * _DPI_SCALE))
_PIP_SW       = max(28,  int(28  * _DPI_SCALE))
_pip_hwnd = [None]
_pip_cls  = [None]

_PIP_BG_A   = 185
_PIP_BG_VAL = (
    (20 * _PIP_BG_A // 255)
    | ((20 * _PIP_BG_A // 255) << 8)
    | ((20 * _PIP_BG_A // 255) << 16)
    | (_PIP_BG_A << 24)
)
_pip_pixel_template = (ctypes.c_uint32 * (_PIP_OW * _PIP_OH))()
for _i in range(_PIP_OW * _PIP_OH):
    _pip_pixel_template[_i] = _PIP_BG_VAL
for _ty in range(_PIP_OH):
    for _tx in range(_PIP_SW, _PIP_OW):
        _pip_pixel_template[_ty * _PIP_OW + _tx] = 0xFF000000

def _pip_wndproc(hwnd, msg, wp, lp):
    if msg == 0x0002:  # WM_DESTROY
        return 0
    return _pip_u32.DefWindowProcW(hwnd, msg, wp, lp)

_PIP_WNDPROC_T = ctypes.WINFUNCTYPE(
    ctypes.c_ssize_t, ctypes.c_void_p,
    ctypes.c_uint, ctypes.c_ssize_t, ctypes.c_ssize_t,
)
_pip_u32.DefWindowProcW.restype  = ctypes.c_ssize_t
_pip_u32.DefWindowProcW.argtypes = [
    ctypes.c_void_p, ctypes.c_uint, ctypes.c_ssize_t, ctypes.c_ssize_t,
]

def init_pip_overlay():
    """Luo overlay-ikkunan. Kutsuttava dpg.create_context():n jälkeen."""
    import os as _os
    app._pip_wndproc_cb = _PIP_WNDPROC_T(_pip_wndproc)
    _pip_u32.RegisterClassW.restype   = ctypes.c_uint16
    _pip_u32.RegisterClassW.argtypes  = [ctypes.c_void_p]
    _pip_u32.CreateWindowExW.restype  = ctypes.c_void_p
    _pip_u32.CreateWindowExW.argtypes = [
        ctypes.c_ulong, ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_ulong,
        ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
        ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p,
    ]

    class _PipWNDCLASS(ctypes.Structure):
        _fields_ = [
            ('style',         ctypes.c_uint),
            ('lpfnWndProc',   ctypes.c_void_p),
            ('cbClsExtra',    ctypes.c_int),
            ('cbWndExtra',    ctypes.c_int),
            ('hInstance',     ctypes.c_void_p),
            ('hIcon',         ctypes.c_void_p),
            ('hCursor',       ctypes.c_void_p),
            ('hbrBackground', ctypes.c_void_p),
            ('lpszMenuName',  ctypes.c_wchar_p),
            ('lpszClassName', ctypes.c_wchar_p),
        ]

    ctypes.windll.kernel32.GetModuleHandleW.restype = ctypes.c_void_p
    _pip_hInst = ctypes.windll.kernel32.GetModuleHandleW(None)
    _pip_cls[0] = f"PipOvl_{_os.getpid()}"
    _pip_wc = _PipWNDCLASS()
    _pip_wc.lpfnWndProc   = ctypes.cast(app._pip_wndproc_cb, ctypes.c_void_p).value
    _pip_wc.hInstance     = _pip_hInst
    _pip_wc.lpszClassName = _pip_cls[0]
    if _pip_u32.RegisterClassW(ctypes.byref(_pip_wc)):
        _pip_hwnd[0] = _pip_u32.CreateWindowExW(
            0x00000008 | 0x00000080 | 0x08000000 | 0x00080000,
            _pip_cls[0], "", 0x80000000,
            0, 0, _PIP_OW, _PIP_OH,
            None, None, _pip_hInst, None,
        )
    return _pip_hInst

# ── GDI-resurssit ─────────────────────────────────────────────────────────────
class _BLENDFUNCTION(ctypes.Structure):
    _fields_ = [('BlendOp',             ctypes.c_byte),
                ('BlendFlags',           ctypes.c_byte),
                ('SourceConstantAlpha',  ctypes.c_byte),
                ('AlphaFormat',          ctypes.c_byte)]

class _BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [('biSize',          ctypes.c_uint32),
                ('biWidth',         ctypes.c_int32),
                ('biHeight',        ctypes.c_int32),
                ('biPlanes',        ctypes.c_uint16),
                ('biBitCount',      ctypes.c_uint16),
                ('biCompression',   ctypes.c_uint32),
                ('biSizeImage',     ctypes.c_uint32),
                ('biXPelsPerMeter', ctypes.c_int32),
                ('biYPelsPerMeter', ctypes.c_int32),
                ('biClrUsed',       ctypes.c_uint32),
                ('biClrImportant',  ctypes.c_uint32)]

class _BITMAPINFO(ctypes.Structure):
    _fields_ = [('bmiHeader', _BITMAPINFOHEADER),
                ('bmiColors', ctypes.c_uint32 * 3)]

_pip_gdi = {
    'hdc_screen': None, 'hdc_mem': None, 'hbmp': None,
    'old_hbmp': None, 'bits_ptr': None,
}

_pip_u32.GetDC.restype                  = ctypes.c_void_p
_pip_u32.GetDC.argtypes                 = [ctypes.c_void_p]
_pip_u32.ReleaseDC.restype              = ctypes.c_int
_pip_u32.ReleaseDC.argtypes             = [ctypes.c_void_p, ctypes.c_void_p]
_pip_g32.CreateCompatibleDC.restype     = ctypes.c_void_p
_pip_g32.CreateCompatibleDC.argtypes    = [ctypes.c_void_p]
_pip_g32.CreateDIBSection.restype       = ctypes.c_void_p
_pip_g32.CreateDIBSection.argtypes      = [
    ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint,
    ctypes.POINTER(ctypes.c_void_p), ctypes.c_void_p, ctypes.c_ulong,
]
_pip_g32.SelectObject.restype           = ctypes.c_void_p
_pip_g32.SelectObject.argtypes          = [ctypes.c_void_p, ctypes.c_void_p]
_pip_g32.DeleteObject.restype           = ctypes.c_bool
_pip_g32.DeleteObject.argtypes          = [ctypes.c_void_p]
_pip_g32.DeleteDC.restype               = ctypes.c_bool
_pip_g32.DeleteDC.argtypes              = [ctypes.c_void_p]
_pip_g32.SetBkMode.restype              = ctypes.c_int
_pip_g32.SetBkMode.argtypes             = [ctypes.c_void_p, ctypes.c_int]
_pip_g32.SetBkColor.restype             = ctypes.c_ulong
_pip_g32.SetBkColor.argtypes            = [ctypes.c_void_p, ctypes.c_ulong]
_pip_g32.SetTextColor.restype           = ctypes.c_ulong
_pip_g32.SetTextColor.argtypes          = [ctypes.c_void_p, ctypes.c_ulong]
_pip_u32.DrawTextW.restype              = ctypes.c_int
_pip_u32.DrawTextW.argtypes             = [
    ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_int,
    ctypes.c_void_p, ctypes.c_uint,
]

def _pip_init_gdi():
    if _pip_gdi['hdc_screen'] is not None:
        return
    hdc_screen = _pip_u32.GetDC(None)
    hdc_mem    = _pip_g32.CreateCompatibleDC(hdc_screen)
    bmi = _BITMAPINFO()
    bmi.bmiHeader.biSize        = ctypes.sizeof(_BITMAPINFOHEADER)
    bmi.bmiHeader.biWidth       = _PIP_OW
    bmi.bmiHeader.biHeight      = -_PIP_OH
    bmi.bmiHeader.biPlanes      = 1
    bmi.bmiHeader.biBitCount    = 32
    bmi.bmiHeader.biCompression = 0
    bits_ptr = ctypes.c_void_p()
    hbmp = _pip_g32.CreateDIBSection(
        hdc_mem, ctypes.byref(bmi), 0, ctypes.byref(bits_ptr), None, 0)
    if hbmp:
        _old_hbmp = _pip_g32.SelectObject(hdc_mem, hbmp)
        _pip_gdi['hdc_screen'] = hdc_screen
        _pip_gdi['hdc_mem']    = hdc_mem
        _pip_gdi['hbmp']       = hbmp
        _pip_gdi['old_hbmp']   = _old_hbmp
        _pip_gdi['bits_ptr']   = bits_ptr
    else:
        _pip_g32.DeleteDC(hdc_mem)
        _pip_u32.ReleaseDC(None, hdc_screen)

def _pip_free_gdi():
    if _pip_gdi['hdc_mem'] and _pip_gdi.get('old_hbmp'):
        _pip_g32.SelectObject(_pip_gdi['hdc_mem'], _pip_gdi['old_hbmp'])
    if _pip_gdi['hbmp']:
        _pip_g32.DeleteObject(_pip_gdi['hbmp'])
    if _pip_gdi['hdc_mem']:
        _pip_g32.DeleteDC(_pip_gdi['hdc_mem'])
    if _pip_gdi['hdc_screen']:
        _pip_u32.ReleaseDC(None, _pip_gdi['hdc_screen'])
    for k in _pip_gdi:
        _pip_gdi[k] = None

_PIP_SW_INNER = _PIP_SW - 8
_pip_sw_row = (ctypes.c_uint32 * _PIP_SW_INNER)()
_pip_sw_row_bytes = ctypes.sizeof(_pip_sw_row)
_PIP_TEXT_W = _PIP_OW - _PIP_SW
_pip_txt_row = (ctypes.c_uint32 * _PIP_TEXT_W)()
_pip_txt_row_bytes = ctypes.sizeof(_pip_txt_row)

def _pip_draw(ox, oy, pr, pg, pb):
    """Piirtää eyedropper-overlayn UpdateLayeredWindow:lla."""
    if not _pip_hwnd[0]:
        return
    _pip_init_gdi()
    hdc_screen = _pip_gdi['hdc_screen']
    hdc_mem    = _pip_gdi['hdc_mem']
    bits_ptr   = _pip_gdi['bits_ptr']
    if hdc_screen is None:
        return

    ctypes.memmove(bits_ptr, _pip_pixel_template,
                    _PIP_OW * _PIP_OH * ctypes.sizeof(ctypes.c_uint32))
    sw_val = pb | (pg << 8) | (pr << 16) | (255 << 24)
    for _k in range(_PIP_SW_INNER):
        _pip_sw_row[_k] = sw_val
    _row_stride = _PIP_OW * 4
    _sw_x_off   = 4 * 4
    for sy in range(4, _PIP_OH - 4):
        dst = bits_ptr.value + sy * _row_stride + _sw_x_off
        ctypes.memmove(dst, _pip_sw_row, _pip_sw_row_bytes)

    hex_str = f"#{pr:02X}{pg:02X}{pb:02X}"
    _pip_g32.SetBkMode(hdc_mem, 2)
    _pip_g32.SetBkColor(hdc_mem, 0x000000)
    _pip_g32.SetTextColor(hdc_mem, 0xFFFFFF)
    rc3 = _wt.RECT(_PIP_SW, 0, _PIP_OW - 2, _PIP_OH)
    _pip_u32.DrawTextW(hdc_mem, hex_str, -1, ctypes.byref(rc3), 0x0025)

    _bg  = _PIP_BG_VAL
    _txt_x_off = _PIP_SW * 4
    for _ty in range(_PIP_OH):
        src = bits_ptr.value + _ty * _row_stride + _txt_x_off
        ctypes.memmove(_pip_txt_row, src, _pip_txt_row_bytes)
        for _k in range(_PIP_TEXT_W):
            if (_pip_txt_row[_k] & 0xFF) > 40:
                _pip_txt_row[_k] = 0xFFFFFFFF
            else:
                _pip_txt_row[_k] = _bg
        ctypes.memmove(src, _pip_txt_row, _pip_txt_row_bytes)

    pt_dst = _wt.POINT(ox, oy)
    pt_src = _wt.POINT(0, 0)
    sz     = _wt.SIZE(_PIP_OW, _PIP_OH)
    blend  = _BLENDFUNCTION(0, 0, 255, 1)
    _pip_u32.UpdateLayeredWindow(
        _pip_hwnd[0], hdc_screen, ctypes.byref(pt_dst),
        ctypes.byref(sz), hdc_mem, ctypes.byref(pt_src),
        0, ctypes.byref(blend), 2,
    )

# Argtyypit overlayta varten
_pip_u32.SetWindowPos.restype  = ctypes.c_bool
_pip_u32.SetWindowPos.argtypes = [
    ctypes.c_void_p, ctypes.c_void_p,
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_uint,
]
_pip_u32.UpdateLayeredWindow.restype  = ctypes.c_bool
_pip_u32.UpdateLayeredWindow.argtypes = [
    ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p,
    ctypes.c_void_p, ctypes.c_void_p, ctypes.c_ulong, ctypes.c_void_p, ctypes.c_ulong,
]
_pip_u32.ShowWindow.restype  = ctypes.c_bool
_pip_u32.ShowWindow.argtypes = [ctypes.c_void_p, ctypes.c_int]
_pip_u32.GetCursorPos.restype  = ctypes.c_bool
_pip_u32.GetCursorPos.argtypes = [ctypes.c_void_p]
_pip_u32.GetSystemMetrics.restype  = ctypes.c_int
_pip_u32.GetSystemMetrics.argtypes = [ctypes.c_int]
_pip_u32.PeekMessageW.restype  = ctypes.c_bool
_pip_u32.PeekMessageW.argtypes = [
    ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint, ctypes.c_uint, ctypes.c_uint,
]
_pip_u32.DestroyWindow.restype   = ctypes.c_bool
_pip_u32.DestroyWindow.argtypes  = [ctypes.c_void_p]
_pip_u32.UnregisterClassW.restype  = ctypes.c_bool
_pip_u32.UnregisterClassW.argtypes = [ctypes.c_wchar_p, ctypes.c_void_p]
_pip_u32.PostQuitMessage.restype   = None
_pip_u32.PostQuitMessage.argtypes  = [ctypes.c_int]

_pip_was_visible  = [False]
_PIP_HWND_TOPMOST = ctypes.c_void_p(-1)
_PIP_SWP_FLAGS    = ctypes.c_uint(0x0001 | 0x0002 | 0x0010)
_pip_pt = _wt.POINT()

def _pip_overlay_update():
    """Päivittää overlay-ikkunan sijainnin ja piirron. Kutsutaan update()-silmukassa."""
    if not _pip_hwnd[0]:
        return
    _msg2 = _wt.MSG()
    while _pip_u32.PeekMessageW(ctypes.byref(_msg2), _pip_hwnd[0], 0, 0, 1) > 0:
        _pip_u32.TranslateMessage(ctypes.byref(_msg2))
        _pip_u32.DispatchMessageW(ctypes.byref(_msg2))
    if app.pipette_active:
        _pip_u32.GetCursorPos(ctypes.byref(_pip_pt))
        _vx0 = _pip_u32.GetSystemMetrics(76)
        _vy0 = _pip_u32.GetSystemMetrics(77)
        _vsw = _pip_u32.GetSystemMetrics(78)
        _vsh = _pip_u32.GetSystemMetrics(79)
        ox = _pip_pt.x + _sc(18)
        oy = _pip_pt.y + _sc(18)
        if ox + _PIP_OW > _vx0 + _vsw: ox = _pip_pt.x - _PIP_OW - _sc(4)
        if oy + _PIP_OH > _vy0 + _vsh: oy = _pip_pt.y - _PIP_OH - _sc(4)
        with _pip_lock:
            _pr, _pg, _pb = app.pip_color[0], app.pip_color[1], app.pip_color[2]
        _pip_draw(ox, oy, _pr, _pg, _pb)
        if not _pip_was_visible[0]:
            _pip_u32.SetWindowPos(_pip_hwnd[0], _PIP_HWND_TOPMOST, 0, 0, 0, 0, _PIP_SWP_FLAGS)
            _pip_was_visible[0] = True
        _pip_u32.ShowWindow(_pip_hwnd[0], 4)
    else:
        if _pip_was_visible[0]:
            _pip_was_visible[0] = False
        _pip_u32.ShowWindow(_pip_hwnd[0], 0)


# ── Globaali pikanäppäin Alt+P (RegisterHotKey) ──────────────────────────────
_hotkey_tid = [0]
_hotkey_ok  = [False]

def _hotkey_thread(ready, result):
    u32 = ctypes.windll.user32
    MOD_ALT, MOD_NOREPEAT, VK_P, WM_HOTKEY = 0x0001, 0x4000, 0x50, 0x0312
    _hotkey_tid[0] = ctypes.windll.kernel32.GetCurrentThreadId()
    ok = bool(u32.RegisterHotKey(None, 1, MOD_ALT | MOD_NOREPEAT, VK_P))
    _hotkey_ok[0] = ok
    result[0] = ok
    ready.set()
    if not ok:
        _write_log("RegisterHotKey(Alt+P) failed - falling back to key polling.")
        _hotkey_tid[0] = 0
        return
    msg = _wt.MSG()
    try:
        while u32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            if msg.message == WM_HOTKEY:
                start_pipette()
    finally:
        u32.UnregisterHotKey(None, 1)
        _hotkey_ok[0] = False
        _hotkey_tid[0] = 0

def start_hotkey():
    """Rekisteröi Alt+P. Palauttaa True jos onnistui (muuten käytetään pollausta)."""
    ready  = threading.Event()
    result = [False]
    threading.Thread(target=_hotkey_thread, args=(ready, result), daemon=True).start()
    ready.wait(1.0)
    return result[0]

def stop_hotkey():
    tid = _hotkey_tid[0]
    if tid:
        ctypes.windll.user32.PostThreadMessageW(tid, 0x0012, 0, 0)

def hotkey_registered():
    return _hotkey_ok[0]
