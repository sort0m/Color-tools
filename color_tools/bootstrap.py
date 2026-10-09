"""
bootstrap.py — Alustus, polut, Win32-tiedostodialogi, virheenkäsittely.

Täytyy importoida ensimmäisenä ennen muita moduuleja.
"""

import os
import sys
import traceback
import ctypes
import time
import threading
import faulthandler as _fh

# Sovellus käyttää Win32-API:a (ctypes.windll) -> vain Windows. Selkeä viesti kaatumisen sijaan.
if os.name != 'nt':
    sys.stderr.write("Color Tools requires Windows (it uses the Win32 API).\n")
    sys.exit(1)

# ── Crash log ────────────────────────────────────────────────────────────────
try:
    _fh_dir = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), 'Color Tools')
    os.makedirs(_fh_dir, exist_ok=True)
    _fh_path = os.path.join(_fh_dir, 'crash.log')
    # Append so the previous run's crash info survives; start fresh only if the log grows large.
    _fh_mode = 'w' if (os.path.exists(_fh_path) and os.path.getsize(_fh_path) > 200000) else 'a'
    _fh_file = open(_fh_path, _fh_mode, encoding='utf-8')
    _fh.enable(file=_fh_file)
    import atexit as _atexit
    _atexit.register(_fh_file.close)
except Exception:
    pass

# ── DPI awareness ─────────────────────────────────────────────────────────────
# System-aware layout: Windows scales the complete window when it moves to a
# monitor with a different DPI. Do not advertise per-monitor awareness while
# the layout and fonts are calculated only once.
_DPI_SCALE = 1.0
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    ctypes.windll.user32.GetDC.restype = ctypes.c_void_p
    ctypes.windll.user32.GetDC.argtypes = [ctypes.c_void_p]
    ctypes.windll.user32.ReleaseDC.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
    ctypes.windll.gdi32.GetDeviceCaps.argtypes = [ctypes.c_void_p, ctypes.c_int]
    _hdc = ctypes.windll.user32.GetDC(0)
    if _hdc:
        try:
            _dpi = ctypes.windll.gdi32.GetDeviceCaps(_hdc, 88)
            _DPI_SCALE = max(96, _dpi) / 96.0
        finally:
            ctypes.windll.user32.ReleaseDC(0, _hdc)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

# ── Config-hakemisto ──────────────────────────────────────────────────────────
_APPDATA    = os.environ.get('APPDATA', os.path.expanduser('~'))
CFG_DIR     = os.path.join(_APPDATA, 'Color Tools')
CFG_FILE    = os.path.join(CFG_DIR, 'config.json')
_LOG        = os.path.join(CFG_DIR, 'error.log')
if getattr(sys, 'frozen', False):
    _SCRIPT_DIR = sys._MEIPASS
else:
    _SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
    # __file__ on nyt color_tools/bootstrap.py — pakettihakemisto on sen parent
    _SCRIPT_DIR = os.path.dirname(_SCRIPT_DIR)

# ── Win32 native file dialogs ─────────────────────────────────────────────────
class _OPENFILENAMEW(ctypes.Structure):
    _fields_ = [
        ('lStructSize',       ctypes.c_uint32),
        ('hwndOwner',         ctypes.c_void_p),
        ('hInstance',         ctypes.c_void_p),
        ('lpstrFilter',       ctypes.c_wchar_p),
        ('lpstrCustomFilter', ctypes.c_wchar_p),
        ('nMaxCustFilter',    ctypes.c_uint32),
        ('nFilterIndex',      ctypes.c_uint32),
        ('lpstrFile',         ctypes.c_wchar_p),
        ('nMaxFile',          ctypes.c_uint32),
        ('lpstrFileTitle',    ctypes.c_wchar_p),
        ('nMaxFileTitle',     ctypes.c_uint32),
        ('lpstrInitialDir',   ctypes.c_wchar_p),
        ('lpstrTitle',        ctypes.c_wchar_p),
        ('Flags',             ctypes.c_uint32),
        ('nFileOffset',       ctypes.c_uint16),
        ('nFileExtension',    ctypes.c_uint16),
        ('lpstrDefExt',       ctypes.c_wchar_p),
        ('lCustData',         ctypes.c_ssize_t),
        ('lpfnHook',          ctypes.c_void_p),
        ('lpTemplateName',    ctypes.c_wchar_p),
        ('pvReserved',        ctypes.c_void_p),
        ('dwReserved',        ctypes.c_uint32),
        ('FlagsEx',           ctypes.c_uint32),
    ]

_comdlg32 = ctypes.windll.comdlg32
for _dialog_name in ('GetOpenFileNameW', 'GetSaveFileNameW'):
    _dialog = getattr(_comdlg32, _dialog_name)
    _dialog.argtypes = [ctypes.POINTER(_OPENFILENAMEW)]
    _dialog.restype = ctypes.c_int

# OPENFILENAME-liput
_OFN_OVERWRITEPROMPT = 0x00000002
_OFN_NOCHANGEDIR     = 0x00000008  # älä vaihda prosessin työhakemistoa
_OFN_PATHMUSTEXIST   = 0x00000800
_OFN_FILEMUSTEXIST   = 0x00001000

def _win32_open_file(title, filter_pairs, owner_hwnd=None):
    """Avaa Win32 GetOpenFileNameW-dialogi. Palauttaa valitun polun tai None."""
    filt = chr(0).join(f"{d}{chr(0)}{p}" for d, p in filter_pairs) + chr(0)*2
    buf  = ctypes.create_unicode_buffer(32768)
    ofn  = _OPENFILENAMEW()
    ofn.lStructSize  = ctypes.sizeof(_OPENFILENAMEW)
    ofn.hwndOwner    = owner_hwnd
    ofn.lpstrFilter  = filt
    ofn.nFilterIndex = 1
    ofn.lpstrFile    = ctypes.cast(buf, ctypes.c_wchar_p)
    ofn.nMaxFile     = len(buf)
    ofn.lpstrTitle   = title
    ofn.Flags        = _OFN_FILEMUSTEXIST | _OFN_PATHMUSTEXIST | _OFN_NOCHANGEDIR
    if _comdlg32.GetOpenFileNameW(ctypes.byref(ofn)):
        return buf.value
    return None

def _win32_save_file(title, filter_pairs, default_ext, initial_name=None, owner_hwnd=None):
    """Avaa Win32 GetSaveFileNameW-dialogi. Palauttaa valitun polun tai None."""
    filt = chr(0).join(f"{d}{chr(0)}{p}" for d, p in filter_pairs) + chr(0)*2
    buf  = ctypes.create_unicode_buffer(initial_name or "", 32768)
    ofn  = _OPENFILENAMEW()
    ofn.lStructSize  = ctypes.sizeof(_OPENFILENAMEW)
    ofn.hwndOwner    = owner_hwnd
    ofn.lpstrFilter  = filt
    ofn.nFilterIndex = 1
    ofn.lpstrFile    = ctypes.cast(buf, ctypes.c_wchar_p)
    ofn.nMaxFile     = len(buf)
    ofn.lpstrTitle   = title
    ofn.lpstrDefExt  = default_ext.lstrip('.')
    ofn.Flags        = _OFN_OVERWRITEPROMPT | _OFN_NOCHANGEDIR
    if _comdlg32.GetSaveFileNameW(ctypes.byref(ofn)):
        return buf.value
    return None

# ── Lokikirjoitus ja fatal-error ──────────────────────────────────────────────
def _write_log(msg):
    try:
        os.makedirs(CFG_DIR, exist_ok=True)
        # Älä anna lokin kasvaa rajatta: yli 500 kt -> vanha talteen (.old)
        if os.path.exists(_LOG) and os.path.getsize(_LOG) > 500000:
            os.replace(_LOG, _LOG + '.old')
        with open(_LOG, 'a', encoding='utf-8') as f:
            f.write(f"\n=== {time.strftime('%Y-%m-%d %H:%M:%S')} ===\n{msg}\n")
    except Exception:
        pass

def _fatal(msg):
    _write_log(msg)
    try:
        ctypes.windll.user32.MessageBoxW(0, msg[:1200], "Color Tools - Error", 0x10)
    except Exception:
        pass
    sys.exit(1)

# ── Riippuvuudet ──────────────────────────────────────────────────────────────
# ── Leikepöytä (Win32; ei ulkoista riippuvuutta) ─────────────────────────────
_CF_UNICODETEXT = 13
_GMEM_MOVEABLE  = 0x0002
_k32 = ctypes.windll.kernel32
_u32 = ctypes.windll.user32
_k32.GlobalAlloc.restype   = ctypes.c_void_p
_k32.GlobalAlloc.argtypes  = [ctypes.c_uint, ctypes.c_size_t]
_k32.GlobalLock.restype    = ctypes.c_void_p
_k32.GlobalLock.argtypes   = [ctypes.c_void_p]
_k32.GlobalUnlock.restype  = ctypes.c_int
_k32.GlobalUnlock.argtypes = [ctypes.c_void_p]
_k32.GlobalFree.restype    = ctypes.c_void_p
_k32.GlobalFree.argtypes   = [ctypes.c_void_p]
_u32.OpenClipboard.restype      = ctypes.c_int
_u32.OpenClipboard.argtypes     = [ctypes.c_void_p]
_u32.CloseClipboard.restype     = ctypes.c_int
_u32.EmptyClipboard.restype     = ctypes.c_int
_u32.SetClipboardData.restype   = ctypes.c_void_p
_u32.SetClipboardData.argtypes  = [ctypes.c_uint, ctypes.c_void_p]

def _clipboard_set(text):
    buf  = ctypes.create_unicode_buffer(text)
    size = ctypes.sizeof(buf)                       # sis. päättävän nollan
    for _ in range(20):                             # toinen sovellus voi pitää leikepöytää hetken lukittuna
        if _u32.OpenClipboard(None):
            break
        time.sleep(0.01)
    else:
        raise OSError("OpenClipboard failed")
    try:
        _u32.EmptyClipboard()
        h = _k32.GlobalAlloc(_GMEM_MOVEABLE, size)
        if not h:
            raise MemoryError("GlobalAlloc failed")
        p = _k32.GlobalLock(h)
        if not p:
            _k32.GlobalFree(h)
            raise MemoryError("GlobalLock failed")
        ctypes.memmove(p, buf, size)
        _k32.GlobalUnlock(h)
        if not _u32.SetClipboardData(_CF_UNICODETEXT, h):
            _k32.GlobalFree(h)
            raise OSError("SetClipboardData failed")
        # onnistui: muistin omistus siirtyi järjestelmälle
    finally:
        _u32.CloseClipboard()

def safe_copy(text):
    """Kopioi tekstin leikepöydälle; ohittaa virheet (esim. leikepöytä lukittuna)."""
    try:
        _clipboard_set(text)
    except Exception as e:
        _write_log(f"Clipboard copy failed: {e}")

try:
    import dearpygui.dearpygui  # noqa: F401
except Exception:
    _fatal("Could not import dearpygui:\n\n" + traceback.format_exc())

# PIL on valinnainen — käytetään vain kuvan paletti-tuontiin.
try:
    import PIL
    import PIL.Image  # noqa: F401
    _PIL_AVAILABLE = tuple(int(v) for v in PIL.__version__.split('.')[:3]) >= (12, 3, 0)
    if not _PIL_AVAILABLE:
        _write_log("Pillow >= 12.3.0 required for safe image import; install requirements.txt.")
except Exception:
    _PIL_AVAILABLE = False
    _write_log("Pillow unavailable — image palette import disabled.\nInstall with: pip install -r requirements.txt")

# ── HWND-haku prosessi-ID:n perusteella ──────────────────────────────────────
_OWN_PID = ctypes.windll.kernel32.GetCurrentProcessId()
_EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
_enum_windows_lock = threading.Lock()
_u32.GetWindowThreadProcessId.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_ulong)]
_u32.GetWindowThreadProcessId.restype = ctypes.c_ulong
_u32.GetWindowTextW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_int]
_u32.GetWindowTextW.restype = ctypes.c_int
_u32.EnumWindows.argtypes = [_EnumWindowsProc, ctypes.c_void_p]
_u32.EnumWindows.restype = ctypes.c_int


def acquire_instance_lock():
    """Keep one settings writer per user/session; the OS releases on exit."""
    import hashlib
    key = hashlib.sha256(os.path.normcase(CFG_FILE).encode('utf-8')).hexdigest()
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.CreateMutexW.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_wchar_p]
    kernel.CreateMutexW.restype = ctypes.c_void_p
    kernel.CloseHandle.argtypes = [ctypes.c_void_p]
    kernel.CloseHandle.restype = ctypes.c_int
    ctypes.set_last_error(0)
    handle = kernel.CreateMutexW(None, False, 'Local\\ColorTools-' + key)
    error = ctypes.get_last_error()
    if not handle:
        raise ctypes.WinError(error)
    if error == 183:  # ERROR_ALREADY_EXISTS
        kernel.CloseHandle(handle)
        raise RuntimeError('Color Tools is already running. / Color Tools on jo käynnissä.')
    import atexit
    atexit.register(kernel.CloseHandle, handle)
    return handle

def _get_own_hwnd():
    result = [None]
    def _cb(hwnd, _):
        w_pid = ctypes.c_ulong(0)
        ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(w_pid))
        if w_pid.value == _OWN_PID:
            buf = ctypes.create_unicode_buffer(256)
            ctypes.windll.user32.GetWindowTextW(hwnd, buf, 256)
            if buf.value == "Color Tools":
                result[0] = hwnd
                return False
        return True
    with _enum_windows_lock:
        _cb_ref = _EnumWindowsProc(_cb)
        ctypes.windll.user32.EnumWindows(_cb_ref, 0)
    return result[0]

# ── DPI-skaalaus-apufunktio ───────────────────────────────────────────────────
def _sc(v):
    """Skaalaa pikseliarvo DPI:llä."""
    return max(1, int(round(v * _DPI_SCALE)))
