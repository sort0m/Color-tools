import os
import sys
import traceback
import ctypes
import faulthandler as _fh
# Also write to file so crashes are visible without a console
try:
    _fh_dir = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), 'Color Tools')
    os.makedirs(_fh_dir, exist_ok=True)
    _fh_file = open(os.path.join(_fh_dir, 'crash.log'), 'w', encoding='utf-8')
    _fh.enable(file=_fh_file)
    import atexit as _atexit
    _atexit.register(_fh_file.close)
except Exception:
    pass

import colorsys
import threading
import time
import json
import struct
from functools import lru_cache
from collections import deque
import ctypes.wintypes as _wt

# ── DPI awareness (must run before everything else) ─────────────────────────
# Level 1 (system-aware): the process declares DPI awareness so Windows does
# not automatically scale the window. DPI is read at startup and all UI
# measurements are scaled accordingly -- the window stays the correct size on
# both monitors and content does not shrink when moving between screens.
_DPI_SCALE = 1.0
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
    ctypes.windll.user32.GetDC.restype = ctypes.c_void_p  # HDC is a pointer on 64-bit
    _hdc = ctypes.windll.user32.GetDC(0)
    _dpi = ctypes.windll.gdi32.GetDeviceCaps(_hdc, 88)  # LOGPIXELSX
    ctypes.windll.user32.ReleaseDC(0, _hdc)
    _DPI_SCALE = _dpi / 96.0
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

# ── Config directory ─────────────────────────────────────────────────────────
_APPDATA    = os.environ.get('APPDATA', os.path.expanduser('~'))
CFG_DIR     = os.path.join(_APPDATA, 'Color Tools')
CFG_FILE    = os.path.join(CFG_DIR, 'config.json')
_LOG        = os.path.join(CFG_DIR, 'error.log')
if getattr(sys, 'frozen', False):
    _SCRIPT_DIR = sys._MEIPASS
else:
    _SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))



# ── Win32 native file dialogs ────────────────────────────────────────────────
# GetOpenFileNameW / GetSaveFileNameW are direct, COM-free Win32 calls.
# They work safely from a background thread (no Tcl/Tk thread restriction)
# and do not leave processes hanging after close.
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

def _win32_open_file(title, filter_pairs, owner_hwnd=None):
    """Open the Win32 GetOpenFileNameW dialog. Returns the selected path or None."""
    filt = chr(0).join(f"{d}{chr(0)}{p}" for d, p in filter_pairs) + chr(0)*2
    buf  = ctypes.create_unicode_buffer(32768)
    ofn  = _OPENFILENAMEW()
    ofn.lStructSize  = ctypes.sizeof(_OPENFILENAMEW)
    ofn.hwndOwner    = owner_hwnd
    ofn.lpstrFilter  = filt
    ofn.nFilterIndex = 1
    # FIX: create_unicode_buffer returns c_wchar_Array_N which is not directly
    # assignable to c_wchar_p struct fields. Cast to a pointer first.
    ofn.lpstrFile    = ctypes.cast(buf, ctypes.c_wchar_p)
    ofn.nMaxFile     = len(buf)
    ofn.lpstrTitle   = title
    ofn.Flags        = 0x00001000 | 0x00000800 | 0x00000008  # OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR
    if _comdlg32.GetOpenFileNameW(ctypes.byref(ofn)):
        return buf.value
    return None

def _win32_save_file(title, filter_pairs, default_ext, initial_name=None, owner_hwnd=None):
    """Open the Win32 GetSaveFileNameW dialog. Returns the selected path or None."""
    filt = chr(0).join(f"{d}{chr(0)}{p}" for d, p in filter_pairs) + chr(0)*2
    buf  = ctypes.create_unicode_buffer(initial_name or "", 32768)
    ofn  = _OPENFILENAMEW()
    ofn.lStructSize  = ctypes.sizeof(_OPENFILENAMEW)
    ofn.hwndOwner    = owner_hwnd
    ofn.lpstrFilter  = filt
    ofn.nFilterIndex = 1
    # FIX: cast c_wchar_Array_N to c_wchar_p before assigning to struct field
    ofn.lpstrFile    = ctypes.cast(buf, ctypes.c_wchar_p)
    ofn.nMaxFile     = len(buf)
    ofn.lpstrTitle   = title
    ofn.lpstrDefExt  = default_ext.lstrip('.')
    ofn.Flags        = 0x00000002  # OFN_OVERWRITEPROMPT
    if _comdlg32.GetSaveFileNameW(ctypes.byref(ofn)):
        return buf.value
    return None

def _write_log(msg):
    try:
        os.makedirs(CFG_DIR, exist_ok=True)
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


# Dependency imports come after _fatal is defined so errors are reported correctly.
try:
    import pyperclip
except Exception:
    _fatal(
        "Missing dependency. Install with:\n"
        "  pip install pyperclip\n\n"
        + traceback.format_exc()
    )

try:
    import dearpygui.dearpygui as dpg
except Exception:
    _fatal("Could not import dearpygui:\n\n" + traceback.format_exc())

# PIL is optional — used only for image palette import.
# If missing, the button still works but shows an error message.
try:
    from PIL import Image as _PIL_Image
    _PIL_AVAILABLE = True
except Exception:
    _PIL_AVAILABLE = False
    _write_log("PIL (Pillow) not found — image palette import disabled.\nInstall with: pip install Pillow")

_write_log("Imports OK, starting up...")

try:
    # -- HWND helper: locates the window by process ID, not title.
    # FindWindowW(None, title) could return the wrong window if another
    # process happens to use the same title. Process ID is unambiguous.
    _OWN_PID = ctypes.windll.kernel32.GetCurrentProcessId()
    _EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    _enum_windows_lock = threading.Lock()  # Serializes concurrent EnumWindows calls

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
                    return False  # stop enumeration
            return True
        with _enum_windows_lock:
            # Keep a local reference to prevent premature GC of the callback.
            # EnumWindows is synchronous so the risk is low, but this is the
            # correct defensive pattern.
            _cb_ref = _EnumWindowsProc(_cb)
            ctypes.windll.user32.EnumWindows(_cb_ref, 0)
        return result[0]

    def load_config():
        try:
            with open(CFG_FILE, encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}

    def save_config():
        try:
            os.makedirs(CFG_DIR, exist_ok=True)
            with _history_lock:
                history_snapshot = list(app.history)
            with _palettes_lock:
                palettes_snapshot  = {k: list(v) for k, v in app.palettes.items()}
                pal_order_snapshot = list(app._pal_order)
            # Atomic write: write to a temp file then rename over the original.
            # If the process crashes mid-write, config.json stays intact.
            _tmp = CFG_FILE + '.tmp'
            with open(_tmp, 'w', encoding='utf-8') as f:
                json.dump({
                    'history':   history_snapshot,
                    'palettes':  palettes_snapshot,
                    'pal_order': pal_order_snapshot,
                    'win_pos':    list(app._last_viewport_pos),
                    'theme_name': app.theme_name,
                    'always_on_top': app.always_on_top,
                }, f)
            os.replace(_tmp, CFG_FILE)
        except Exception:
            _write_log("save_config failed:\n" + traceback.format_exc())

    # -- CSS color library ------------------------------------------------
    CSS_COLORS = {
        "Black": (0, 0, 0),       "White": (255, 255, 255),  "Red": (255, 0, 0),
        "Lime": (0, 255, 0),      "Blue": (0, 0, 255),       "Yellow": (255, 255, 0),
        "Cyan": (0, 255, 255),    "Magenta": (255, 0, 255),
        "Silver": (192, 192, 192),"Gray": (128, 128, 128),   "Maroon": (128, 0, 0),
        "Olive": (128, 128, 0),   "Green": (0, 128, 0),      "Purple": (128, 0, 128),
        "Teal": (0, 128, 128),    "Navy": (0, 0, 128),       "Orange": (255, 165, 0),
        "Aquamarine": (127, 255, 212),  "BlueViolet": (138, 43, 226),
        "Brown": (165, 42, 42),         "CadetBlue": (95, 158, 160),
        "Chocolate": (210, 105, 30),    "Coral": (255, 127, 80),
        "Crimson": (220, 20, 60),       "DarkBlue": (0, 0, 139),
        "DarkCyan": (0, 139, 139),      "DarkGoldenrod": (184, 134, 11),
        "DarkGray": (169, 169, 169),    "DarkGreen": (0, 100, 0),
        "DarkMagenta": (139, 0, 139),   "DarkOrange": (255, 140, 0),
        "DarkRed": (139, 0, 0),         "DarkViolet": (148, 0, 211),
        "DeepPink": (255, 20, 147),     "DeepSkyBlue": (0, 191, 255),
        "DodgerBlue": (30, 144, 255),   "FireBrick": (178, 34, 34),
        "ForestGreen": (34, 139, 34),   "Gold": (255, 215, 0),
        "HotPink": (255, 105, 180),     "IndianRed": (205, 92, 92),
        "Indigo": (75, 0, 130),         "Khaki": (240, 230, 140),
        "Lavender": (230, 230, 250),    "LawnGreen": (124, 252, 0),
        "LightBlue": (173, 216, 230),   "LightCoral": (240, 128, 128),
        "LightGray": (211, 211, 211),   "LightGreen": (144, 238, 144),
        "LightPink": (255, 182, 193),   "LightSkyBlue": (135, 206, 250),
        "LimeGreen": (50, 205, 50),     "MediumOrchid": (186, 85, 211),
        "MediumPurple": (147, 112, 219),"MediumTurquoise": (72, 209, 204),
        "MidnightBlue": (25, 25, 112),  "OrangeRed": (255, 69, 0),
        "Orchid": (218, 112, 214),      "PaleGreen": (152, 251, 152),
        "PaleTurquoise": (175, 238, 238),"Pink": (255, 192, 203),
        "Plum": (221, 160, 221),        "RoyalBlue": (65, 105, 225),
        "Salmon": (250, 128, 114),      "SeaGreen": (46, 139, 87),
        "SkyBlue": (135, 206, 235),     "SlateBlue": (106, 90, 205),
        "SpringGreen": (0, 255, 127),   "SteelBlue": (70, 130, 180),
        "Tomato": (255, 99, 71),        "Turquoise": (64, 224, 208),
        "Violet": (238, 130, 238),      "YellowGreen": (154, 205, 50),
    }

    @lru_cache(maxsize=32768)
    def nearest_css(r, g, b):
        best, bd = "Unknown", float('inf')
        for n, (cr, cg, cb) in CSS_COLORS.items():
            d = (r - cr) ** 2 + (g - cg) ** 2 + (b - cb) ** 2
            if d < bd:
                bd = d
                best = n
        return best

    # -- Color conversions ------------------------------------------------
    # Threshold value 0.04045 per the IEC 61966-2-1 standard.
    # _lin_rgb takes a float [0,1]; _lin normalises an integer [0,255] before calling it.
    _LIN_RGB_LUT = [
        (i/255.)/12.92 if (i/255.) <= 0.04045 else (((i/255.)+0.055)/1.055)**2.4
        for i in range(256)
    ]
    def _lin(c):
        return _LIN_RGB_LUT[max(0, min(255, int(c)))]

    def luminance(r, g, b):
        return 0.2126 * _lin(r) + 0.7152 * _lin(g) + 0.0722 * _lin(b)

    def contrast_ratio(r1, g1, b1, r2, g2, b2):
        l1 = luminance(r1, g1, b1) + 0.05
        l2 = luminance(r2, g2, b2) + 0.05
        return max(l1, l2) / min(l1, l2)

    def wcag_level(r):
        return "AAA" if r >= 7 else "AA" if r >= 4.5 else "AA*" if r >= 3 else "Fail"

    def wcag_color(r):
        """Return an [R,G,B] colour matching the WCAG level."""
        if r >= 7:   return [80,  210,  80]   # AAA  -- green
        if r >= 4.5: return [170, 220,  60]   # AA   -- yellow-green
        if r >= 3:   return [230, 160,  40]   # AA*  -- orange
        return           [200,  70,  70]       # Fail -- red

    def rgb_to_cmyk(r, g, b):
        rf = r / 255.; gf = g / 255.; bf = b / 255.
        k  = 1. - max(rf, gf, bf)
        if k < 1.:
            c = (1 - rf - k) / (1 - k)
            m = (1 - gf - k) / (1 - k)
            y = (1 - bf - k) / (1 - k)
        else:
            c = m = y = 0.
        return c * 100, m * 100, y * 100, k * 100

    def cmyk_to_rgb(c, m, y, k):
        c /= 100; m /= 100; y /= 100; k /= 100
        return (
            max(0, min(255, int(round(255 * (1 - c) * (1 - k))))),
            max(0, min(255, int(round(255 * (1 - m) * (1 - k))))),
            max(0, min(255, int(round(255 * (1 - y) * (1 - k))))),
        )

    # ── LUT acceleration for LAB helpers ─────────────────────────────────────
    # Pre-compute 4096-entry LUTs over [0, 1]. lab_to_rgb is called ~48x per
    # slider per redraw in LAB mode — the LUTs avoid hundreds of pow() calls.
    _LAB_LUT_SIZE = 4096
    _LAB_LUT_MAX  = _LAB_LUT_SIZE - 1

    # _xyz_f LUT: input t in [0, 1]
    _XYZ_F_LUT = [
        (i / _LAB_LUT_MAX) ** (1 / 3)
        if (i / _LAB_LUT_MAX) > 0.008856
        else 7.787 * (i / _LAB_LUT_MAX) + 16 / 116
        for i in range(_LAB_LUT_SIZE)
    ]

    # _delin LUT: input c in [0, 1]
    _DELIN_LUT = [
        (
            12.92 * (i / _LAB_LUT_MAX)
            if (i / _LAB_LUT_MAX) <= 0.0031308
            else 1.055 * (i / _LAB_LUT_MAX) ** (1 / 2.4) - 0.055
        )
        for i in range(_LAB_LUT_SIZE)
    ]

    def _xyz_f_fast(t):
        """LUT-accelerated _xyz_f; used in rgb_to_lab."""
        if t <= 0.0:
            # Correct CIE linear value, not 0.0
            return 7.787 * t + 16 / 116
        if t >= 1.0:
            return _XYZ_F_LUT[_LAB_LUT_MAX]
        return _XYZ_F_LUT[int(t * _LAB_LUT_MAX + 0.5)]

    def _delin_fast(c):
        """LUT-accelerated _delin; used in lab_to_rgb."""
        if c <= 0.0:
            return 0.0
        if c >= 1.0:
            return _DELIN_LUT[_LAB_LUT_MAX]
        return _DELIN_LUT[int(c * _LAB_LUT_MAX + 0.5)]

    def rgb_to_lab(r, g, b):
        # r, g, b are already int 0-255; _lin() handles normalisation internally.
        rl = _lin(r); gl = _lin(g); bl = _lin(b)
        X = rl * .4124 + gl * .3576 + bl * .1805
        Y = rl * .2126 + gl * .7152 + bl * .0722
        Z = rl * .0193 + gl * .1192 + bl * .9505
        X /= .95047; Z /= 1.08883
        return (
            116 * _xyz_f_fast(Y) - 16,
            500 * (_xyz_f_fast(X) - _xyz_f_fast(Y)),
            200 * (_xyz_f_fast(Y) - _xyz_f_fast(Z)),
        )

    # _fi_lab LUT: input t in [0, 1], maps to t**3
    _FI_LAB_LUT = [(i / _LAB_LUT_MAX) ** 3 for i in range(_LAB_LUT_SIZE)]

    def _fi_lab(t):
        """LUT-accelerated inverse of the CIE f function; used in lab_to_rgb.
        Defined at module scope to avoid creating a new closure on every call
        (lab_to_rgb is called ~48x per slider redraw in LAB mode)."""
        if t <= 0.0:
            t3 = 0.0
        elif t >= 1.0:
            t3 = _FI_LAB_LUT[_LAB_LUT_MAX]
        else:
            t3 = _FI_LAB_LUT[int(t * _LAB_LUT_MAX + 0.5)]
        return t3 if t3 > 0.008856 else (t - 16 / 116) / 7.787

    def lab_to_rgb(L, a, b_):
        fy = (L + 16) / 116
        fx = a / 500 + fy
        fz = fy - b_ / 200

        X  = _fi_lab(fx) * .95047; Y = _fi_lab(fy); Z = _fi_lab(fz) * 1.08883
        rl =  X * 3.2406 - Y * 1.5372 - Z * .4986
        gl = -X * .9689  + Y * 1.8758  + Z * .0415
        bl =  X * .0557  - Y * .2040   + Z * 1.0570
        return (
            max(0, min(255, int(round(_delin_fast(rl) * 255)))),
            max(0, min(255, int(round(_delin_fast(gl) * 255)))),
            max(0, min(255, int(round(_delin_fast(bl) * 255)))),
        )

    def rgb_to_gray(r, g, b):
        return int(round(.2126 * r + .7152 * g + .0722 * b))

    # -- Formats ----------------------------------------------------------
    FORMAT_OPTIONS = ["HEX", "RGB", "HSL", "HSV", "CSS Name", "Contrast", "CMYK"]

    def format_value(tag, fmt):
        rgb = app.harmony_rgb.get(tag)
        if rgb is None:
            return ""
        r, g, b = rgb
        if fmt == "HEX":
            return f"#{r:02x}{g:02x}{b:02x}".upper()
        if fmt == "RGB":
            return f"{r}, {g}, {b}"
        if fmt == "HSL":
            # colorsys.rgb_to_hls returns (h, l, s) -- not (h, s, l).
            h, l, s = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
            return f"{h*360:.0f}  {s*100:.0f}%  {l*100:.0f}%"
        if fmt == "HSV":
            h, s, v = colorsys.rgb_to_hsv(r / 255, g / 255, b / 255)
            return f"{h*360:.0f}  {s*100:.0f}%  {v*100:.0f}%"
        if fmt == "CSS Name":
            return nearest_css(r, g, b)
        if fmt == "Contrast":
            cw = contrast_ratio(r, g, b, 255, 255, 255)
            cb = contrast_ratio(r, g, b, 0,   0,   0  )
            # Plain text used only for the Copy button
            return f"W {cw:.1f}:1 {wcag_level(cw)}  B {cb:.1f}:1 {wcag_level(cb)}"
        if fmt == "CMYK":
            c, m, y, k = rgb_to_cmyk(r, g, b)
            return f"C{c:.0f}% M{m:.0f}% Y{y:.0f}% K{k:.0f}%"
        return ""

    HARMONY_TAGS = ["main", "h1", "h2", "h3", "h4", "h5", "h6", "h7"]
    HIST_TAGS    = [f"hist_{i}" for i in range(60)]
    HARMONY_KEYS = {tag: {
        "group":       f"group_{tag}",
        "rect":        f"rect_{tag}",
        "val":         f"val_{tag}",
        "ctr_grp":     f"ctr_grp_{tag}",
        "ctr_bars":    f"ctr_bars_{tag}",
        "copy_btn":    f"copy_btn_{tag}",
        "ctr_w":       f"ctr_w_{tag}",
        "ctr_b":       f"ctr_b_{tag}",
        "ctr_bar_w":   f"ctr_bar_w_txt_{tag}",
        "ctr_bar_b":   f"ctr_bar_b_txt_{tag}",
    } for tag in HARMONY_TAGS}

    def refresh_harmony_values():
        fmt         = app.fmt_mode
        mode        = app.harmony_mode
        is_contrast = (fmt == "Contrast")
        for tag in HARMONY_TAGS:
            keys = HARMONY_KEYS[tag]

            if not dpg.does_item_exist(keys["val"]):
                continue
            if not dpg.is_item_shown(keys["group"]):
                continue
            dpg.configure_item(keys["val"],      show=not is_contrast)
            dpg.configure_item(keys["ctr_grp"],  show=is_contrast)
            dpg.configure_item(keys["ctr_bars"], show=is_contrast)
            dpg.configure_item(keys["copy_btn"], show=not is_contrast and mode not in ("Tints", "Shades", "Tones"))
            if is_contrast:
                rgb = app.harmony_rgb.get(tag)
                if rgb:
                    r, g, b = rgb
                    cw = contrast_ratio(r, g, b, 255, 255, 255)
                    cb = contrast_ratio(r, g, b, 0,   0,   0  )
                    dpg.set_value(f"ctr_w_{tag}",
                                  f"W  {cw:.2f}:1  {wcag_level(cw)}")
                    dpg.configure_item(f"ctr_w_{tag}", color=wcag_color(cw))
                    dpg.set_value(f"ctr_b_{tag}",
                                  f"B  {cb:.2f}:1  {wcag_level(cb)}")
                    dpg.configure_item(f"ctr_b_{tag}", color=wcag_color(cb))
                    # Update preview bar text colors
                    if dpg.does_item_exist(f"ctr_bar_w_txt_{tag}"):
                        dpg.configure_item(f"ctr_bar_w_txt_{tag}", color=[r, g, b, 255])
                    if dpg.does_item_exist(f"ctr_bar_b_txt_{tag}"):
                        dpg.configure_item(f"ctr_bar_b_txt_{tag}", color=[r, g, b, 255])
            else:
                dpg.set_value(f"val_{tag}", format_value(tag, fmt))

    # -- Layout constants ─────────────────────────────────────────────────
    def _sc(v):
        return max(1, int(round(v * _DPI_SCALE)))

    W            = _sc(268)
    BW           = _sc(130)
    VALW         = W - _sc(56)
    HSQ          = _sc(26)
    PAL_SW       = _sc(24)
    MAX_PAL_ROWS = 100

    GRAD_W  = W - _sc(30)
    GRAD_H  = _sc(15)
    THUMB_H = GRAD_H + _sc(4)
    SEGS    = 48

    _ESW   = _sc(28)   # palette-editor swatch size
    _ECOLS = 8          # swatches per row in edit mode

    # Pre-computed swatch tags — avoids f-string allocation every frame.
    _PEDIT_SW_TAGS = {i: f"pedit_sw_{i}" for i in range(MAX_PAL_ROWS * _ECOLS + 2)}

    _history_lock  = threading.Lock()
    _palettes_lock = threading.Lock()
    _pip_lock      = threading.Lock()

    # -- Application state ────────────────────────────────────────────────
    class AppState:
        def __init__(self):
            self.pipette_active       = False
            self.pip_color            = [0, 0, 0, 255]
            self._pipette_thread_tid  = 0
            self.current_color   = [0.2, 0.5, 1.0, 1.0]
            self.history         = deque(maxlen=60)
            self.change_time     = 0.
            self.use_wheel       = True
            self.harmony_rgb     = {}
            self.palettes        = {}
            self.sl_vals         = {
                "R": 51,  "G": 127, "B": 255,
                "H": 210, "S": 100, "L": 60,
                "C": 0,   "M": 0,   "Yv": 0,   "K": 0,
                "LA": 50, "La": 0,  "Lb": 0,   "Gray": 128,
            }
            self._dragging          = None
            self._last_sig          = None
            self._last_mode         = None
            self._last_fmt          = None
            self._last_hist         = None
            self._hist_dirty        = True   # set True whenever history changes
            self._last_pal          = None
            self._last_sl_vals      = {}
            self._pal_order         = []
            self._editing_pal       = None
            self.theme_name         = "Dark"
            self.always_on_top      = False
            self.palette_select_mode          = False
            self.selected_history_indices     = set()
            self.selected_history_colors      = {}  # {index: [r,g,b,255]} captured at click time
            self._import_status       = ("", 0.0)  # (error_msg, show_until) — single attribute write is GIL-atomic
            self._pal_drag_idx        = None   # index of the swatch being dragged
            self._pal_drag_active     = False  # threshold exceeded -> real drag
            self._pal_drag_start      = (0, 0) # mouse position at press
            self._pal_drag_insert     = None   # insertion index (0..n), None = no drag
            self._pending_pal_deselect = False
            # -- Custom titlebar drag state --
            self._tb_dragging         = False
            self._tb_drag_start_vp    = (0, 0)
            self._tb_hov              = {"dl_tb_min": None, "dl_tb_close": None}

            self._pal_selected_idx    = None
            self._pal_pre_edit_snapshot = None # Palette snapshot taken at swatch selection, not yet on stack
            self._pal_edit_dirty      = False  # True once color actually changed since last swatch click
            self._pal_undo_freeze     = 0       # Frames left to block _update_selected_pal_color after undo
            self._last_picker_force_time = 0.0
            self._pending_wheel_sync     = None # Frame-delayed sync for DPG wheel bug
            self._pending_picker_color   = None # Deferred delete+recreate, consumed in update()
            self._pal_undo_stack         = deque(maxlen=30)  # Snapshots for undo
            self._mouse_wheel_delta      = 0    # Mouse wheel scroll delta for sliders
            self._pending_pipette_color  = None # Set by pipette thread, consumed by update()
            self._picker_processing      = False # Guard against recursive picker callbacks
            self._pending_save           = False # Set by background threads, flushed in update()
            self._pending_image_import   = None  # PIL Image set by bg thread; consumed by update()
            self._image_modal_open       = False # True while DPG color-count modal is showing
            self._image_modal_pil        = None  # PIL Image kept alive while modal is open
            self._image_modal_colors     = []    # Current quantized colours shown in modal
            self._last_edit_color_sig    = None  # Tracks color changes for palette edit panel refresh
            self._last_preview_sig       = None  # Cache: avoid redundant preview_hex set_value calls
            self._last_viewport_pos      = [100, 100]  # Cached; used by save_config at shutdown
            self._cached_vp_w           = 0     # Cached viewport width
            self._cached_vp_h           = 0     # Cached viewport height
            self._cached_vp_pos         = [0, 0]  # Cached viewport position
            self._pal_dirty              = True   # Set True whenever palette list changes
            self._harmony_text_dirty     = True   # Throttle: text/WCAG updates deferred during drag
            self._vp_pos_frame_ctr       = 0      # Frame counter for viewport position poll rate
            self._last_drawn_preview     = None   # Cache: avoids redundant preview set_value calls
            # PERF: Cached combo values — read from memory instead of DPG every frame
            self.harmony_mode            = "Triadic"  # mirrors "harmony_combo" widget value
            self.fmt_mode                = "HEX"      # mirrors "fmt_combo" widget value
            self.slider_mode             = "RGB"       # mirrors "slider_mode_combo" widget value


    app  = AppState()
    _cfg = load_config()
    if 'history'  in _cfg and isinstance(_cfg['history'],  list):
        app.history = deque(_cfg['history'], maxlen=60)
    if 'palettes' in _cfg and isinstance(_cfg['palettes'], dict):
        app.palettes   = _cfg['palettes']
        app._pal_order = [k for k in _cfg.get('pal_order', []) if k in app.palettes]
        for k in app.palettes:
            if k not in app._pal_order:
                app._pal_order.append(k)
    _start_pos = _cfg.get('win_pos', [100, 100])
    app.theme_name = _cfg.get('theme_name', 'Dark')
    app.always_on_top = bool(_cfg.get('always_on_top', False))

    # -- Helpers ──────────────────────────────────────────────────────────
    def clamp01(v):
        return max(0., min(1., v))

    # round() preserves colour accuracy.
    def to_int(v):
        if v > 1.5:
            return max(0, min(255, int(round(v))))
        return max(0, min(255, int(round(v * 255))))

    def color_sig(c):
        return tuple(to_int(v) for v in c[:3])

    def add_to_history(color):
        r, g, b = color_sig(color)
        e = [r, g, b, 255]
        with _history_lock:
            if app.history and app.history[0] == e:
                return
            app.history.appendleft(e)
            app._hist_dirty = True

    # -- Gradient calculations for sliders --------------------------------
    _grad_ctx_cache = [None, None]  # [last vsig, last result]

    def _grad_ctx():
        """Compute base colour values ONCE per slider redraw — passed to _grad_color.
        Avoids repeating rgb_to_hls / rgb_to_cmyk / rgb_to_lab for every segment.
        Results are cached: if the colour has not changed since the last call,
        the previous result is returned directly without recalculation."""
        vsig = color_sig(app.current_color)
        if _grad_ctx_cache[0] == vsig:
            return _grad_ctx_cache[1]
        r, g, b    = vsig
        rf, gf, bf = r / 255., g / 255., b / 255.
        # colorsys.rgb_to_hls returns (h, l, s) -- not (h, s, l).
        h_, l_, s_ = colorsys.rgb_to_hls(rf, gf, bf)
        c0, m0, y0, k0 = rgb_to_cmyk(r, g, b)
        L0, a0, b0_    = rgb_to_lab(r, g, b)
        result = rf, gf, bf, h_, l_, s_, c0, m0, y0, k0, L0, a0, b0_
        _grad_ctx_cache[0] = vsig
        _grad_ctx_cache[1] = result
        return result

    def _grad_color(name, t, ctx):
        """Return the RGB tuple for gradient segment t given precomputed ctx."""
        rf, gf, bf, h_, l_, s_, c0, m0, y0, k0, L0, a0, b0_ = ctx

        def f(px):
            return (
                max(0, min(255, int(round(px[0] * 255)))),
                max(0, min(255, int(round(px[1] * 255)))),
                max(0, min(255, int(round(px[2] * 255)))),
            )

        if   name == "R":    return f((t, gf, bf))
        elif name == "G":    return f((rf, t, bf))
        elif name == "B":    return f((rf, gf, t))
        elif name == "H":
            hr, hg, hb = colorsys.hls_to_rgb(
                t,
                l_ if .05 < l_ < .95 else .5,
                s_ if s_ > .05 else 1.,
            )
            return f((hr, hg, hb))
        elif name == "S":    return f(colorsys.hls_to_rgb(h_, l_, t))
        elif name == "L":    return f(colorsys.hls_to_rgb(h_, t, s_))
        elif name == "C":    return cmyk_to_rgb(t * 100, m0, y0, k0)
        elif name == "M":    return cmyk_to_rgb(c0, t * 100, y0, k0)
        elif name == "Yv":   return cmyk_to_rgb(c0, m0, t * 100, k0)
        elif name == "K":    return cmyk_to_rgb(c0, m0, y0, t * 100)
        elif name == "LA":   return lab_to_rgb(t * 100, a0, b0_)
        elif name == "La":   return lab_to_rgb(L0, -128 + t * 255, b0_)
        elif name == "Lb":   return lab_to_rgb(L0, a0, -128 + t * 255)
        elif name == "Gray":
            v = int(t * 255); return (v, v, v)
        return (128, 128, 128)

    SLIDER_MODES = ["RGB", "HSL", "CMYK", "LAB", "Grayscale"]
    SLIDERS_BY_MODE = {
        "RGB":       [("R",  "R",    [255, 110, 110],   0, 255),
                      ("G",  "G",    [110, 220, 110],   0, 255),
                      ("B",  "B",    [110, 160, 255],   0, 255)],
        "HSL":       [("H",  "H",    [220, 200,  80],   0, 359),
                      ("S",  "S",    [180, 220, 180],   0, 100),
                      ("L",  "L",    [210, 210, 210],   0, 100)],
        "CMYK":      [("C",  "C",    [ 80, 200, 200],   0, 100),
                      ("M",  "M",    [220,  80, 180],   0, 100),
                      ("Y",  "Yv",   [220, 200,  50],   0, 100),
                      ("K",  "K",    [160, 160, 160],   0, 100)],
        "LAB":       [("L",  "LA",   [220, 220, 220],   0, 100),
                      ("a",  "La",   [220,  80,  80], -128, 127),
                      ("b",  "Lb",   [220, 200,  50], -128, 127)],
        "Grayscale": [("V",  "Gray", [180, 180, 180],   0, 255)],
    }
    # Pre-computed DPG tag names for slider drawlists.
    _DL_TAGS = {gname: f"dl_{gname}"
                for sliders in SLIDERS_BY_MODE.values()
                for _, gname, *_ in sliders}
    # Pre-computed segment X positions for gradient slider bars.
    # Computed once at startup — do not change after DPI is set.
    _sw_unit = GRAD_W / SEGS
    _GRAD_SEG_X = [
        (int(i * _sw_unit), max(int(i * _sw_unit) + 1, int((i + 1) * _sw_unit)))
        for i in range(SEGS)
    ]
    del _sw_unit

    # Cache: {gname: {'seg_ids': [...], 'indicator_id': int, 'y0': int, 'y1': int}}
    # Items are created once and updated via configure_item() — no delete/recreate.
    _GRAD_ITEM_CACHE: dict = {}

    def _apply_from_mode(mode):
        sv = app.sl_vals
        if mode == "RGB":
            r, g, b = sv["R"], sv["G"], sv["B"]
        elif mode == "HSL":
            # hls_to_rgb takes (h, l, s)
            rr, gg, bb = colorsys.hls_to_rgb(sv["H"] / 360., sv["L"] / 100., sv["S"] / 100.)
            r, g, b = int(round(rr * 255)), int(round(gg * 255)), int(round(bb * 255))
        elif mode == "CMYK":
            r, g, b = cmyk_to_rgb(sv["C"], sv["M"], sv["Yv"], sv["K"])
        elif mode == "LAB":
            r, g, b = lab_to_rgb(sv["LA"], sv["La"], sv["Lb"])
            # lab_to_rgb already clamps internally; no extra check needed.
        elif mode == "Grayscale":
            r = g = b = sv["Gray"]
        else:
            return

        col = [r / 255., g / 255., b / 255., 1.]
        app.current_color = col
        app.change_time   = time.perf_counter()
        # Mark picker as processing to prevent recursive callbacks from dpg.set_value()
        app._picker_processing = True
        try:
            dpg.set_value("picker_wheel", [r, g, b, 255])
        finally:
            app._picker_processing = False

        rf = r / 255.; gf = g / 255.; bf = b / 255.

        if mode != "RGB":
            sv["R"] = r; sv["G"] = g; sv["B"] = b
        if mode != "HSL":
            # colorsys.rgb_to_hls returns (h, l, s).
            h, l, s = colorsys.rgb_to_hls(rf, gf, bf)
            sv["H"] = int(h * 360) % 360; sv["S"] = int(s * 100); sv["L"] = int(l * 100)
        if mode != "CMYK":
            c, m, y, k = rgb_to_cmyk(r, g, b)
            sv["C"] = int(round(c)); sv["M"] = int(round(m))
            sv["Yv"] = int(round(y)); sv["K"] = int(round(k))
        if mode != "LAB":
            L, a, b_ = rgb_to_lab(r, g, b)
            sv["LA"] = int(round(L)); sv["La"] = int(round(a)); sv["Lb"] = int(round(b_))
        if mode != "Grayscale":
            sv["Gray"] = rgb_to_gray(r, g, b)

    def _sync_all_modes():
        r, g, b    = color_sig(app.current_color)
        rf, gf, bf = r / 255., g / 255., b / 255.
        sv = app.sl_vals
        sv["R"] = r; sv["G"] = g; sv["B"] = b
        # colorsys.rgb_to_hls returns (h, l, s).
        h, l, s = colorsys.rgb_to_hls(rf, gf, bf)
        sv["H"] = int(h * 360) % 360; sv["S"] = int(s * 100); sv["L"] = int(l * 100)
        c, m, y, k = rgb_to_cmyk(r, g, b)
        sv["C"] = int(round(c)); sv["M"] = int(round(m))
        sv["Yv"] = int(round(y)); sv["K"] = int(round(k))
        L, a, b_ = rgb_to_lab(r, g, b)
        sv["LA"] = int(round(L)); sv["La"] = int(round(a)); sv["Lb"] = int(round(b_))
        sv["Gray"] = rgb_to_gray(r, g, b)

    def _patch_safe_joystick_apis():
        """Stub out joystick/gamepad APIs to prevent DearPyGui from polling devices."""
        k32 = ctypes.windll.kernel32
        k32.GetProcAddress.restype   = ctypes.c_void_p
        k32.GetProcAddress.argtypes  = [ctypes.c_void_p, ctypes.c_char_p]
        k32.LoadLibraryW.restype     = ctypes.c_void_p
        k32.LoadLibraryW.argtypes    = [ctypes.c_wchar_p]
        k32.VirtualProtect.restype   = ctypes.c_bool
        k32.VirtualProtect.argtypes  = [
            ctypes.c_void_p, ctypes.c_size_t,
            ctypes.c_ulong, ctypes.POINTER(ctypes.c_ulong)]
        k32.FlushInstructionCache.restype  = ctypes.c_bool
        k32.FlushInstructionCache.argtypes = [
            ctypes.c_void_p, ctypes.c_void_p, ctypes.c_size_t]
        k32.GetCurrentProcess.restype  = ctypes.c_void_p
        k32.GetCurrentProcess.argtypes = []
        stub  = (ctypes.c_ubyte * 3)(0x31, 0xC0, 0xC3)
        hproc = k32.GetCurrentProcess()
        def _p(dll, fn):
            try:
                hmod = k32.LoadLibraryW(dll)
                if not hmod: return
                addr = k32.GetProcAddress(hmod, fn.encode('ascii'))
                if not addr: return
                old = ctypes.c_ulong(0)
                if k32.VirtualProtect(ctypes.c_void_p(addr), ctypes.c_size_t(3),
                                      ctypes.c_ulong(0x40), ctypes.byref(old)):
                    ctypes.memmove(addr, stub, 3)
                    k32.VirtualProtect(ctypes.c_void_p(addr), ctypes.c_size_t(3),
                                       old, ctypes.byref(ctypes.c_ulong(0)))
                    k32.FlushInstructionCache(
                        hproc, ctypes.c_void_p(addr), ctypes.c_size_t(3))
            except Exception: pass
        _p('winmm',   'joyGetNumDevs')
        for xi in ('xinput1_4', 'xinput1_3', 'xinput9_1_0'):
            _p(xi, 'XInputGetCapabilities')
            _p(xi, 'XInputGetState')
        _p('dinput8', 'DirectInput8Create')
        _p('setupapi','SetupDiEnumDeviceInterfaces')
        _p('setupapi','SetupDiEnumDeviceInfo')
        _p('hid',     'HidD_GetHidGuid')
        _p('winmm',   'waveOutGetNumDevs')
        _p('winmm',   'midiOutGetNumDevs')

    _patch_safe_joystick_apis()

    # ── Message-pumping sleep ─────────────────────────────────────────────
    # time.sleep() blocks the entire Python thread and the Win32 message queue
    # stalls. Windows marks the window "not responding" after ~500 ms.
    # MsgWaitForMultipleObjects wakes immediately on a message, pumps the queue,
    # then sleeps again — the window stays responsive.


    _mwfmo_user32 = ctypes.windll.user32
    _mwfmo_k32    = ctypes.windll.kernel32

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
    _MSG_BUF     = (ctypes.c_ubyte * 48)()   # sizeof(MSG) = 48 bytes x64

    def _pumping_sleep(seconds):
        """Sleep up to `seconds` while continuously pumping the Win32 message queue."""
        deadline = time.perf_counter() + seconds
        while True:
            remaining_ms = int((deadline - time.perf_counter()) * 1000)
            if remaining_ms <= 0:
                break
            # Wait up to remaining_ms or until a message arrives
            ret = _mwfmo_user32.MsgWaitForMultipleObjects(
                0, None, False,
                ctypes.c_ulong(remaining_ms),
                ctypes.c_ulong(_QS_ALLINPUT))
            # ret == 0 (WAIT_OBJECT_0): message arrived — pump it
            if ret == 0:
                while _mwfmo_user32.PeekMessageW(
                        _MSG_BUF, None, 0, 0, _PM_REMOVE):
                    _mwfmo_user32.TranslateMessage(_MSG_BUF)
                    _mwfmo_user32.DispatchMessageW(_MSG_BUF)
            # ret == 258 (WAIT_TIMEOUT): time elapsed — done
            else:
                break

    def _update_selected_pal_color():
        # Block writes for a few frames after undo so that stale DPG picker
        # callbacks cannot corrupt the restored state.
        if app._pal_undo_freeze > 0:
            return
        if app._editing_pal and app._pal_selected_idx is not None:
            with _palettes_lock:
                if app._editing_pal in app.palettes:
                    colors = app.palettes[app._editing_pal]
                    if app._pal_selected_idx < len(colors):
                        r, g, b = color_sig(app.current_color)
                        # Lazy undo push: only push a snapshot the first time the colour
                        # actually changes after a swatch click, so merely browsing
                        # swatches does not fill the undo stack with duplicates.
                        if not app._pal_edit_dirty:
                            snap = app._pal_pre_edit_snapshot
                            if snap is not None:
                                nm = app._editing_pal
                                app._pal_undo_stack.append((nm, snap))
                                if dpg.does_item_exist("pal_undo_btn"):
                                    dpg.configure_item("pal_undo_btn", enabled=True)
                            app._pal_edit_dirty = True
                            app._pal_pre_edit_snapshot = None
                        colors[app._pal_selected_idx] = [r, g, b, 255]
            # Panel refresh is handled by the main update() loop, not here,
            # to avoid constant rebuilds while dragging the colour wheel.

    def _force_update_color_picker(col_norm):
        # Called from DPG callbacks (hex input, history, undo, swatch click).
        # Calling delete_item or add_color_picker inside a callback causes a
        # C-level segfault. Only set flags here; the update() loop performs the
        # delete+recreate safely between render frames.
        app.current_color = list(col_norm)
        app._last_picker_force_time = time.perf_counter()
        app._pending_picker_color   = list(col_norm)


    def picker_cb(s, v):
        # ── CRITICAL GUARDS ──
        # (1) Ignore for 500ms after a forced update to swallow DPG's internal sync noise.
        if time.perf_counter() - app._last_picker_force_time < 0.5:
            return
        # (2) Prevent recursive callbacks from _apply_from_mode calling dpg.set_value()
        if app._picker_processing:
            return
            
        app._picker_processing = True
        try:
            # Normalize if we get 0-255 ints from uint8 mode
            if any(x > 1.01 for x in v):
                v = [x / 255.0 for x in v]
            app.current_color = list(v)
            app.change_time   = time.perf_counter()
            _sync_all_modes()
            _update_selected_pal_color()
        finally:
            app._picker_processing = False

    def copy_preview(s, v):
        r, g, b = color_sig(app.current_color)
        pyperclip.copy(f"#{r:02x}{g:02x}{b:02x}".upper())

    def copy_harmony_val(s, a, u):
        pyperclip.copy(dpg.get_value(f"val_{u}"))

    def hex_input_cb(s, v):
        """Called when the user edits the hex input field and presses Enter."""
        raw = v.strip().lstrip('#')
        if len(raw) == 3:
            raw = raw[0]*2 + raw[1]*2 + raw[2]*2
        if len(raw) != 6:
            return
        try:
            r = int(raw[0:2], 16)
            g = int(raw[2:4], 16)
            b = int(raw[4:6], 16)
        except ValueError:
            return
        col_norm = [r / 255., g / 255., b / 255., 1.]
        _force_update_color_picker(col_norm)
        _sync_all_modes()
        add_to_history(col_norm)

    def toggle_on_top(s, v):
        app.always_on_top = bool(v)
        dpg.set_viewport_always_top(v)
        save_config()

    # -- Palette helpers ──────────────────────────────────────────────────
    def _auto_pal_name():
        i = 1
        while f"Palette {i}" in app.palettes:
            i += 1
        return f"Palette {i}"

    def _register_palette(name):
        # Must be called inside _palettes_lock — _pal_order update must be
        # atomic with the dict write.
        if name not in app._pal_order:
            app._pal_order.append(name)
            app._pal_dirty = True

    def palette_names():
        order = [k for k in reversed(app._pal_order) if k in app.palettes]
        # Use a set for O(1) lookup — list search is O(n) per iteration.
        order_set = set(order)
        for k in sorted(app.palettes.keys()):
            if k not in order_set:
                order.append(k)
        return order

    _PCOLS = 9

    def _rebuild_pal_rows():
        """Rebuild the palette list UI. Must be called from the main thread only."""
        editing = bool(app._editing_pal)
        if dpg.does_item_exist("pal_scroll"):
            dpg.configure_item("pal_scroll",   show=not editing)
        if dpg.does_item_exist("pal_edit_win"):
            dpg.configure_item("pal_edit_win", show=editing)
        if editing:
            return
        with _palettes_lock:
            names = palette_names()
            colors_snapshot = {n: list(app.palettes.get(n, [])) for n in names}
        for pi in range(MAX_PAL_ROWS):
            row_tag = f"pal_row_{pi}"
            if not dpg.does_item_exist(row_tag):
                continue
            if pi >= len(names):
                dpg.configure_item(row_tag, show=False)
                continue
            name   = names[pi]
            colors = colors_snapshot.get(name, [])
            dpg.delete_item(row_tag, children_only=True)
            dpg.configure_item(row_tag, show=True)
            dpg.bind_item_theme(row_tag, _nospc_theme)
            if colors:
                n_rows = -(-len(colors) // _PCOLS)  # ceiling division
                for row_idx, row_start in enumerate(range(0, len(colors), _PCOLS)):
                    is_last = (row_idx == n_rows - 1)
                    grp = dpg.add_group(parent=row_tag, horizontal=True)
                    dpg.bind_item_theme(grp, _nospc_theme)
                    for si in range(row_start, min(row_start + _PCOLS, len(colors))):
                        c    = colors[si]
                        norm = [c[j] / 255. for j in range(3)] + [1.]

                        def _swatch_cb(s, a, u):
                            _force_update_color_picker(u)
                            _sync_all_modes()

                        dpg.add_color_button(
                            parent=grp, default_value=c,
                            width=PAL_SW, height=PAL_SW,
                            callback=_swatch_cb, user_data=norm,
                            no_alpha=True,
                        )
                    if is_last:
                        dpg.add_button(
                            parent=grp, label=">>", width=24, height=PAL_SW,
                            callback=lambda s, a, u: open_palette_editor(u),
                            user_data=name,
                        )
            else:
                # Empty palette: show only action buttons with a placeholder label
                grp = dpg.add_group(parent=row_tag, horizontal=True)
                dpg.bind_item_theme(grp, _nospc_theme)
                dpg.add_text("(empty)", parent=grp, color=[100, 100, 100])
                dpg.add_button(
                    parent=grp, label=">>", width=24, height=PAL_SW,
                    callback=lambda s, a, u: open_palette_editor(u),
                    user_data=name,
                )
            dpg.add_spacer(height=10, parent=row_tag)

    # -- Palette editor ---------------------------------------------------
    def open_palette_editor(name):
        if not name or name not in app.palettes:
            return
        app._editing_pal = name
        if dpg.does_item_exist("pal_scroll"):
            dpg.configure_item("pal_scroll",   show=False)
        if dpg.does_item_exist("pal_edit_win"):
            dpg.configure_item("pal_edit_win", show=True)
        _refresh_pal_edit_panel()

    def _push_pal_undo(name):
        with _palettes_lock:
            if not name or name not in app.palettes:
                return
            # Deep copy each colour list so later in-place mutations cannot
            # corrupt the snapshot.
            colors_copy = [list(c) for c in app.palettes[name]]
        app._pal_undo_stack.append((name, colors_copy))
        if dpg.does_item_exist("pal_undo_btn"):
            dpg.configure_item("pal_undo_btn", enabled=True)

    def _pal_undo(s, a, u):
        if not app._pal_undo_stack:
            return
        name, colors = app._pal_undo_stack.pop()
        idx = app._pal_selected_idx
        app._pal_selected_idx = None
        app._pal_edit_dirty = False
        # Freeze writes for a few frames so stale DPG callbacks cannot
        # overwrite the restored state.
        app._pal_undo_freeze = 4
        with _palettes_lock:
            app.palettes[name] = colors
            # Restore palette order if missing (e.g. after an undo of a deletion)
            if name not in app._pal_order:
                app._pal_order.append(name)
        # Reset pre-edit snapshot to the restored state immediately so the
        # next change can be lazy-pushed onto the undo stack.
        app._pal_pre_edit_snapshot = [list(c) for c in colors]
        # Update the picker to the restored colour.
        if (name == app._editing_pal
                and idx is not None
                and idx < len(colors)):
            c = colors[idx]
            _force_update_color_picker([c[0] / 255., c[1] / 255., c[2] / 255., 1.0])
            _sync_all_modes()
        # Restore selection last, after all other state is consistent.
        app._pal_selected_idx = idx if (idx is not None and idx < len(colors)) else None
        save_config()
        _rebuild_pal_rows()
        _refresh_pal_edit_panel()
        if dpg.does_item_exist("pal_undo_btn"):
            dpg.configure_item("pal_undo_btn", enabled=len(app._pal_undo_stack) > 0)

    def _refresh_pal_edit_panel(drag_src=None, drag_insert=None):
        """Rebuild the palette editor panel.

        Normal mode: draw all colours.
        Drag preview: omit the dragged colour and insert an empty gap slot
        (yellow border) at drag_insert.
        """
        if not dpg.does_item_exist("pal_edit_panel"):
            return
        dpg.delete_item("pal_edit_panel", children_only=True)
        name = app._editing_pal
        if not name or name not in app.palettes:
            if dpg.does_item_exist("pal_edit_win"):
                dpg.configure_item("pal_edit_win", show=False)
            if dpg.does_item_exist("pal_scroll"):
                dpg.configure_item("pal_scroll",   show=True)
            return
        dpg.configure_item("pal_edit_panel", show=True)
        # Read palette inside lock — background threads may modify app.palettes concurrently.
        with _palettes_lock:
            if name not in app.palettes:
                return
            colors = list(app.palettes[name])

        _BW2 = (W - _sc(8)) // 2 - _sc(8)

        # ── Callback definitions ─────────────────────────────────────────
        def _close_edit(s, a, u):
            app._editing_pal = None
            app._pal_selected_idx = None
            app._pending_pal_deselect = False
            _rebuild_pal_rows()
            _refresh_pal_edit_panel()

        def _add_current(s, a, u):
            nm = app._editing_pal
            if not nm or nm not in app.palettes:
                return
            r, g, b = color_sig(app.current_color)
            _push_pal_undo(nm)
            with _palettes_lock:
                app.palettes[nm].append([r, g, b, 255])
            # Deselect so the picker is free to choose the next colour to add.
            app._pal_selected_idx = None
            app._pal_dirty = True
            save_config()
            _rebuild_pal_rows()
            _refresh_pal_edit_panel()

        def _exp_pal(s, a, u):
            export_palette_by_name(app._editing_pal)

        def _exp_pal_ase(s, a, u):
            export_palette_ase_by_name(app._editing_pal)

        def _del_pal(s, a, u):
            nm = app._editing_pal
            with _palettes_lock:
                if nm and nm in app.palettes:
                    del app.palettes[nm]
                    if nm in app._pal_order:
                        app._pal_order.remove(nm)
            app._editing_pal = None
            app._pal_selected_idx = None
            app._pending_pal_deselect = False
            app._pal_dirty = True
            save_config()
            _rebuild_pal_rows()
            _refresh_pal_edit_panel()

        # ── Colour swatches (top) ─────────────────────────────────────────
        dpg.add_text("Drag to reorder", parent="pal_edit_panel", color=[90, 90, 90])
        dpg.add_spacer(height=2, parent="pal_edit_panel")

        dragging = (drag_src is not None and drag_insert is not None
                    and 0 <= drag_src < len(colors) and len(colors) > 1)

        if dragging:
            stripped = [c for i, c in enumerate(colors) if i != drag_src]
            # Defensive guard: drag_insert should never be None when dragging=True.
            safe_ins = drag_insert if drag_insert is not None else len(stripped)
            ins = max(0, min(len(stripped), safe_ins))
            preview = stripped[:ins] + [None] + stripped[ins:]
        else:
            preview = list(colors)

        if preview:
            for row_start in range(0, len(preview), _ECOLS):
                grp = dpg.add_group(parent="pal_edit_panel", horizontal=True)
                dpg.bind_item_theme(grp, _nospc_theme)
                for slot in range(row_start, min(row_start + _ECOLS, len(preview))):
                    c  = preview[slot]
                    dl = dpg.add_drawlist(parent=grp, tag=f"pedit_sw_{slot}",
                                         width=_ESW, height=_ESW)
                    if c is None:
                        dc = colors[drag_src]
                        dpg.draw_rectangle(parent=dl,
                                           pmin=(1, 1), pmax=(_ESW - 2, _ESW - 2),
                                           fill=[dc[0], dc[1], dc[2], 255],
                                           color=[220, 200, 50, 255], thickness=2)
                    else:
                        is_sel = (app._pal_selected_idx == slot and not dragging)
                        is_light = _THEME_MAP.get(app.theme_name, (None, None, False))[2]
                        sel_border = [40, 40, 40, 255] if is_light else [255, 255, 255, 255]
                        border = sel_border if is_sel else [40, 40, 40, 180]
                        thick = 2 if is_sel else 1
                        dpg.draw_rectangle(parent=dl,
                                           pmin=(1, 1), pmax=(_ESW - 2, _ESW - 2),
                                           fill=[c[0], c[1], c[2], 255],
                                           color=border, thickness=thick)
        else:
            dpg.add_text("(empty palette)", parent="pal_edit_panel", color=[120, 120, 120])

        dpg.add_spacer(height=2, parent="pal_edit_panel")
        dpg.add_text("Right-click to remove", parent="pal_edit_panel", color=[90, 90, 90])

        # ── Action buttons (bottom) ───────────────────────────────────────
        dpg.add_separator(parent="pal_edit_panel")
        dpg.add_spacer(height=2, parent="pal_edit_panel")

        _grp2 = dpg.add_group(parent="pal_edit_panel", horizontal=True, tag="pal_action_btns")
        dpg.add_button(parent=_grp2, label="+ Add color", width=_BW2, height=THUMB_H, callback=_add_current)
        dpg.add_button(parent=_grp2, tag="pal_undo_btn", label="Undo", width=_BW2, height=THUMB_H,
                       callback=_pal_undo, enabled=(len(app._pal_undo_stack) > 0))

        dpg.add_spacer(height=2, parent="pal_edit_panel")

        _grp1 = dpg.add_group(parent="pal_edit_panel", horizontal=True, tag="pal_export_btns")
        dpg.add_button(parent=_grp1, label="Export HTML", width=_BW2, height=THUMB_H, callback=_exp_pal)
        dpg.add_button(parent=_grp1, label="Export ASE",  width=_BW2, height=THUMB_H, callback=_exp_pal_ase)

        dpg.add_spacer(height=2, parent="pal_edit_panel")

        _grp3 = dpg.add_group(parent="pal_edit_panel", horizontal=True, tag="pal_close_btns")
        dpg.add_button(parent=_grp3, label="Delete palette", width=_BW2, height=THUMB_H, callback=_del_pal)
        dpg.add_button(parent=_grp3, label="Close editor",   width=_BW2, height=THUMB_H, callback=_close_edit)

        dpg.add_spacer(height=2, parent="pal_edit_panel")

    def export_palette_by_name(name):
        with _palettes_lock:
            if not name or name not in app.palettes:
                return
            colors = list(app.palettes[name])
        h = (
            f"<html><head><meta charset='utf-8'>HTMLTITLE_PLACEHOLDER{_HCSS}</head><body>"
            f"<h2>HTMLH2_PLACEHOLDER</h2>"
            f"<table>{_html_table_header()}"
        )
        for v in colors:
            h += _html_color_row(v[0], v[1], v[2])
        h += "</table></body></html>"
        _save_html_dialog(h, name)

    def _pal_swatch_right_click(rmb_clicked):
        """Handle right-click colour removal in the palette editor. Called every frame."""
        if not rmb_clicked:
            return
        if app._editing_pal and app._editing_pal in app.palettes:
            # Read ncolors inside the lock — a background thread could shrink
            # the list between the release and the hover loop.
            with _palettes_lock:
                ncolors = len(app.palettes.get(app._editing_pal, []))
            for idx in range(ncolors):
                stag = _PEDIT_SW_TAGS[idx]
                if dpg.does_item_exist(stag) and dpg.is_item_hovered(stag):
                    _push_pal_undo(app._editing_pal)
                    with _palettes_lock:
                        nm = app._editing_pal
                        if nm and nm in app.palettes and idx < len(app.palettes[nm]):
                            app.palettes[nm].pop(idx)
                            if not app.palettes[nm]:
                                del app.palettes[nm]
                                if nm in app._pal_order:
                                    app._pal_order.remove(nm)
                                app._editing_pal = None
                                app._pal_selected_idx = None
                                app._pal_pre_edit_snapshot = None
                                app._pal_edit_dirty = False
                            else:
                                # Update _pal_selected_idx after removal so the picker
                                # does not write to the wrong slot.
                                if app._pal_selected_idx is not None:
                                    if app._pal_selected_idx == idx:
                                        # The selected color was deleted — deselect
                                        app._pal_selected_idx = None
                                        app._pal_pre_edit_snapshot = None
                                        app._pal_edit_dirty = False
                                    elif app._pal_selected_idx > idx:
                                        # Selected color shifted one position left
                                        app._pal_selected_idx -= 1
                    app._pal_dirty = True
                    save_config()
                    _rebuild_pal_rows()
                    _refresh_pal_edit_panel()
                    return

    def _pal_swatch_drag(lmb_dn, lmb_clk, lmb_rel):
        """Palette drag-and-drop reorder. Called every frame.

        Target slot determined by DPG hover-testing rather than distance,
        avoiding coordinate-system confusion on DPI-scaled displays.
        """
        if not app._editing_pal or app._editing_pal not in app.palettes:
            return

        # Take a snapshot inside the lock — background threads may replace or
        # mutate app.palettes[name] after the lock is released.
        with _palettes_lock:
            if app._editing_pal not in app.palettes:
                return
            colors = list(app.palettes[app._editing_pal])  # snapshot copy
            n      = len(colors)

        _mp = dpg.get_mouse_pos(local=False)
        if not _mp or len(_mp) < 2: return
        mx, my = _mp

        def _hovered_slot():
            """Return the index of the slot currently under the mouse pointer."""
            for slot in range(n + 1):   # up to n slots in the preview
                t = _PEDIT_SW_TAGS[slot]
                if dpg.does_item_exist(t) and dpg.is_item_hovered(t):
                    return slot
            return None

        # ── Start drag ──────────────────────────────────────────────────────────────────────
        if lmb_clk:
            for i in range(n):
                tag = _PEDIT_SW_TAGS[i]
                if dpg.does_item_exist(tag) and dpg.is_item_hovered(tag):
                    app._pal_drag_idx    = i
                    app._pal_drag_insert = i
                    app._pal_drag_active = False
                    app._pal_drag_start  = (mx, my)

                    # Deselect if the already-selected swatch is clicked again
                    if app._pal_selected_idx == i:
                        app._pal_selected_idx = None
                        app._pal_pre_edit_snapshot = None
                        app._pal_edit_dirty = False
                        _refresh_pal_edit_panel()
                        break
                    # Take a snapshot now (before any colour change),
                    # but do NOT push yet — lazy push happens in _update_selected_pal_color
                    # only when the colour actually changes. This prevents mere swatch
                    # browsing from filling the undo stack with identical snapshots.
                    if app._editing_pal in app.palettes:
                        app._pal_pre_edit_snapshot = [list(c) for c in app.palettes[app._editing_pal]]
                    app._pal_edit_dirty = False
                    app._pal_selected_idx = i
                    app._last_edit_color_sig = None  # Force immediate edit panel refresh
                    c = colors[i]
                    v = [c[0] / 255., c[1] / 255., c[2] / 255., 1.0]
                    _force_update_color_picker(v)
                    _sync_all_modes()
                    _refresh_pal_edit_panel()
                    break



        # ── Mouse moves with LMB held ─────────────────────────────────────────────────────
        if lmb_dn and app._pal_drag_idx is not None:
            dx = mx - app._pal_drag_start[0]
            dy = my - app._pal_drag_start[1]
            if not app._pal_drag_active and (dx * dx + dy * dy) > 16:
                app._pal_drag_active = True
                # Drag is a reorder — always push the current state onto the undo stack.
                # Use pre-edit snapshot if the colour hasn't been changed yet,
                # otherwise save the current state.
                if not app._pal_edit_dirty and app._pal_pre_edit_snapshot is not None:
                    nm = app._editing_pal
                    if nm and nm in app.palettes:
                        app._pal_undo_stack.append((nm, app._pal_pre_edit_snapshot))
                        if dpg.does_item_exist("pal_undo_btn"):
                            dpg.configure_item("pal_undo_btn", enabled=True)
                else:
                    _push_pal_undo(app._editing_pal)
                app._pal_pre_edit_snapshot = None
                app._pal_edit_dirty = True
                _refresh_pal_edit_panel(
                    drag_src    = app._pal_drag_idx,
                    drag_insert = app._pal_drag_insert,
                )

            if app._pal_drag_active:
                hit = _hovered_slot()
                if hit is not None and hit != app._pal_drag_insert:
                    app._pal_drag_insert = hit
                    _refresh_pal_edit_panel(
                        drag_src    = app._pal_drag_idx,
                        drag_insert = app._pal_drag_insert,
                    )

        # ── LMB released ──────────────────────────────────────────────────────────────────────
        if lmb_rel and app._pal_drag_idx is not None:
            src    = app._pal_drag_idx
            ins    = app._pal_drag_insert
            active = app._pal_drag_active

            app._pal_drag_idx    = None
            app._pal_drag_insert = None
            app._pal_drag_active = False

            if not active:
                _refresh_pal_edit_panel()

            else:
                if ins is not None and 0 <= src < n:
                    stripped = [c for i, c in enumerate(colors) if i != src]
                    real_ins = max(0, min(len(stripped), ins))
                    stripped.insert(real_ins, colors[src])
                    with _palettes_lock:
                        app.palettes[app._editing_pal][:] = stripped
                    # Keep the selection tracking the dragged swatch at its new position,
                    # so the color picker stays in sync with the highlighted slot.
                    if app._pal_selected_idx == src:
                        app._pal_selected_idx = real_ins
                    app._pal_dirty = True
                    save_config()
                _rebuild_pal_rows()
                _refresh_pal_edit_panel()

    # -- History selection -----------------------------------------------
    _hist_selected_theme = None  # Placeholder; real theme created after dpg.create_context()

    _hist_prev_selection: set = set()

    def _update_history_selection_style():
        # Only update buttons whose state changed since the last call,
        # eliminating redundant bind_item_theme calls on every click.
        prev = _hist_prev_selection
        curr = app.selected_history_indices
        changed = prev.symmetric_difference(curr)
        for i in changed:
            if i < len(HIST_TAGS):
                tag = HIST_TAGS[i]
                if dpg.does_item_exist(tag):
                    dpg.bind_item_theme(tag, _hist_selected_theme if i in curr else None)
        prev.clear(); prev.update(curr)

    def save_selected_as_palette():
        if not app.selected_history_indices:
            return
        # Use colours captured at click time, not current history indices.
        # The history deque may shift (appendleft) between selection and save.
        colors = [
            app.selected_history_colors[i]
            for i in sorted(app.selected_history_indices)
            if i in app.selected_history_colors
        ]
        if not colors:
            return
        with _palettes_lock:
            name = _auto_pal_name()
            app.palettes[name] = colors
            _register_palette(name)
        save_config()
        _rebuild_pal_rows()
        _exit_palette_select_mode()

    def _exit_palette_select_mode():
        """Exit palette creation mode without saving."""
        app.palette_select_mode = False
        app.selected_history_indices.clear()
        app.selected_history_colors.clear()
        dpg.set_item_label("btn_new_palette", "New Palette")
        dpg.bind_item_theme("btn_new_palette", None)
        if dpg.does_item_exist("btn_cancel_palette"):
            dpg.configure_item("btn_cancel_palette", show=False)
        _update_history_selection_style()

    def toggle_palette_select_mode():
        if app.palette_select_mode:
            if app.selected_history_indices:
                save_selected_as_palette()
                if dpg.does_item_exist("btn_cancel_palette"):
                    dpg.configure_item("btn_cancel_palette", show=False)
            else:
                _exit_palette_select_mode()
        else:
            app.palette_select_mode = True
            app.selected_history_indices.clear()
            app.selected_history_colors.clear()
            dpg.set_item_label("btn_new_palette", "Save")
            dpg.bind_item_theme("btn_new_palette", _new_palette_select_theme)
            if dpg.does_item_exist("btn_cancel_palette"):
                dpg.configure_item("btn_cancel_palette", show=True)
            _update_history_selection_style()

    def history_click_cb(sender, app_data, user_data):
        i = user_data
        _do_style_update = False
        with _history_lock:
            if i >= len(app.history):
                return
            if app.palette_select_mode:
                if i in app.selected_history_indices:
                    app.selected_history_indices.remove(i)
                    app.selected_history_colors.pop(i, None)
                else:
                    app.selected_history_indices.add(i)
                    # Capture colour at click time — history may shift before save.
                    app.selected_history_colors[i] = list(app.history[i])
                _do_style_update = True
            else:
                col      = app.history[i]
                col_norm = [col[0] / 255., col[1] / 255., col[2] / 255., 1.]
                _force_update_color_picker(col_norm)
                _sync_all_modes()
        # DPG calls made outside the lock to avoid holding it during UI updates.
        if _do_style_update:
            _update_history_selection_style()

    # -- Harmony -> palette -----------------------------------------------
    def save_harmony_as_palette():
        """Save the current harmony colours as a new palette."""
        colors = []
        for tag in HARMONY_TAGS:
            if app.harmony_rgb.get(tag) is not None:
                r, g, b = app.harmony_rgb[tag]
                colors.append([r, g, b, 255])
        if not colors:
            return
        with _palettes_lock:
            name = _auto_pal_name()
            app.palettes[name] = colors
            _register_palette(name)
        save_config()
        _rebuild_pal_rows()

    # -- Image import -----------------------------------------------------
    def _commit_image_palette(colors):
        """Commit imported colours to history and palettes. Thread-safe; no DPG calls."""
        with _history_lock:
            for c in reversed(colors):
                if c not in app.history:
                    app.history.appendleft(c)
        with _palettes_lock:
            # Name generation, dict write and order registration in one lock.
            name = _auto_pal_name()
            app.palettes[name] = list(colors)
            _register_palette(name)
        # save_config() is not thread-safe; signal the main thread instead.
        app._pending_save = True
        app._pal_dirty = True

    def import_image_palette():
        """Open a Win32 file dialog and load the image in a background thread.
        The colour-count panel is shown by the main thread (update loop)."""
        if not _PIL_AVAILABLE:
            app._import_status = ("Pillow not installed. Run: pip install Pillow", time.perf_counter() + 6.0)
            return

        def _run():
            try:
                # Step 1: Win32 file dialog (thread-safe)
                path = _win32_open_file(
                    "Open Image",
                    [
                        ("Image files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp;*.tiff"),
                        ("All files",   "*.*"),
                    ],
                )
                if not path:
                    return

                # Step 2: Load and shrink the image
                try:
                    pil_img = _PIL_Image.open(path).convert("RGB")
                    pil_img.thumbnail((200, 200))
                except Exception:
                    _write_log("import_image: PIL open failed:\n" + traceback.format_exc())
                    app._import_status = ("Could not open image — see error.log", time.perf_counter() + 5.0)
                    return

                # Step 3: Hand the image to the main thread; it will open a DPG modal.
                # Hand image to the main thread for the colour-count panel.
                app._pending_image_import = pil_img

            except Exception:
                _write_log("import_image error:\n" + traceback.format_exc())
                app._import_status = ("Image import failed — see error.log", time.perf_counter() + 5.0)
                app._pal_dirty = True

        threading.Thread(target=_run, daemon=True).start()

    # -- Image import: quantisation + DPG modal (all on main thread) ----------

    def _quantize_image(pil_img, n):
        """Return a sorted list of [R,G,B,255] colours from PIL quantisation."""
        try:
            method = _PIL_Image.Quantize.FASTOCTREE
        except AttributeError:
            method = getattr(_PIL_Image, 'FASTOCTREE', 2)
        quantized    = pil_img.quantize(colors=n, method=method)
        palette_data = quantized.getpalette()[:n * 3]
        colors = [
            [palette_data[i*3], palette_data[i*3+1], palette_data[i*3+2], 255]
            for i in range(n)
        ]
        colors.sort(key=lambda c: 0.2126*c[0] + 0.7152*c[1] + 0.0722*c[2])
        return colors

    # Colours per row in import preview (max 20 = max 2 rows)
    _IMG_PREVIEW_COLS = 10

    def _rebuild_image_import_preview():
        """Redraw colour swatches in the import panel (10 per row, max 2 rows)."""
        if not dpg.does_item_exist("img_import_preview"):
            return
        dpg.delete_item("img_import_preview", children_only=True)
        colors = app._image_modal_colors
        for row_start in range(0, len(colors), _IMG_PREVIEW_COLS):
            grp = dpg.add_group(parent="img_import_preview", horizontal=True)
            dpg.bind_item_theme(grp, _nospc_theme)
            for c in colors[row_start : row_start + _IMG_PREVIEW_COLS]:
                dpg.add_color_button(
                    parent=grp,
                    default_value=c, width=PAL_SW, height=PAL_SW,
                    no_alpha=True, no_border=True,
                )

    def _image_import_slider_cb(s, v):
        app._image_modal_colors = _quantize_image(app._image_modal_pil, int(v))
        _rebuild_image_import_preview()

    def _image_import_ok(s, a, u):
        colors = list(app._image_modal_colors)
        _image_import_close()
        if colors:
            _commit_image_palette(colors)

    def _image_import_cancel(s, a, u):
        _image_import_close()

    def _image_import_close():
        """Close the import panel and restore the normal palette list."""
        app._image_modal_open   = False
        app._image_modal_pil    = None
        app._image_modal_colors = []
        if dpg.does_item_exist("img_import_win"):
            dpg.configure_item("img_import_win", show=False)
        if dpg.does_item_exist("pal_scroll"):
            dpg.configure_item("pal_scroll", show=not bool(app._editing_pal))
        if dpg.does_item_exist("pal_edit_win"):
            dpg.configure_item("pal_edit_win", show=bool(app._editing_pal))

    def _open_image_import_panel(pil_img):
        """Show the colour-count selection panel in place of the palette list.
        Must be called from the main thread."""
        if app._image_modal_open:
            return
        app._image_modal_open   = True
        app._image_modal_pil    = pil_img
        app._image_modal_colors = _quantize_image(pil_img, 8)

        # Hide the normal palette view; show the import panel instead.
        if dpg.does_item_exist("pal_scroll"):
            dpg.configure_item("pal_scroll",   show=False)
        if dpg.does_item_exist("pal_edit_win"):
            dpg.configure_item("pal_edit_win", show=False)
        if dpg.does_item_exist("img_import_win"):
            dpg.configure_item("img_import_win", show=True)

        # Reset slider to 8 and refresh the preview.
        if dpg.does_item_exist("img_import_slider"):
            dpg.set_value("img_import_slider", 8)
        _rebuild_image_import_preview()



    def import_ase_palette():
        """Import colours from an ASE (Adobe Swatch Exchange) file."""
        def _run():
            try:
                # Win32 save dialog — no Tkinter, thread-safe
                path = _win32_open_file(
                    "Import ASE Palette",
                    [("ASE files", "*.ase"), ("All files", "*.*")],
                )
                if not path:
                    return
                with open(path, 'rb') as f:
                    data = f.read()
                # Validate header
                if data[:4] != b'ASEF':
                    app._import_status = ("Not a valid ASE file.", time.perf_counter() + 5.0)
                    return
                num_blocks = struct.unpack_from('>I', data, 8)[0]
                pos = 12
                colors = []
                for _ in range(num_blocks):
                    if pos + 6 > len(data):
                        break
                    block_type, block_len = struct.unpack_from('>HI', data, pos)
                    pos += 6
                    block_end = pos + block_len
                    if block_type == 0x0001:  # color entry
                        # name length (uint16) + name (utf-16-be)
                        if pos + 2 > len(data):
                            pos = block_end
                            continue
                        name_len = struct.unpack_from('>H', data, pos)[0]
                        pos += 2 + name_len * 2
                        # color model (4 chars)
                        if pos + 4 > len(data):
                            pos = block_end
                            continue
                        model = data[pos:pos+4]
                        pos += 4
                        if model == b'RGB ':
                            if pos + 12 > len(data):
                                pos = block_end
                                continue
                            r, g, b = struct.unpack_from('>fff', data, pos)
                            r = max(0, min(255, int(round(r * 255))))
                            g = max(0, min(255, int(round(g * 255))))
                            b = max(0, min(255, int(round(b * 255))))
                            colors.append([r, g, b, 255])
                    pos = block_end
                if not colors:
                    app._import_status = ("No RGB colours found in ASE file.", time.perf_counter() + 5.0)
                    return
                _commit_image_palette(colors)
            except Exception:
                _write_log("import_ase error:\n" + traceback.format_exc())
                app._import_status = ("ASE import failed — see error.log", time.perf_counter() + 5.0)
                app._pal_dirty = True
        threading.Thread(target=_run, daemon=True).start()

    # -- Eyedropper ───────────────────────────────────────────────────────
    def pipette_thread():
        app.pipette_active = True


        WH_MOUSE_LL    = 14
        WM_MOUSEMOVE   = 0x0200
        WM_LBUTTONDOWN = 0x0201
        WM_QUIT        = 0x0012

        class _MSLLHOOKSTRUCT(ctypes.Structure):
            _fields_ = [
                ('x',           ctypes.c_long),
                ('y',           ctypes.c_long),
                ('mouseData',   ctypes.c_ulong),
                ('flags',       ctypes.c_ulong),
                ('time',        ctypes.c_ulong),
                # dwExtraInfo is ULONG_PTR (pointer-sized) — c_ulong is only
                # 32-bit on Windows even in 64-bit mode, causing struct misalignment.
                ('dwExtraInfo', ctypes.c_size_t),
            ]

        _user32  = ctypes.windll.user32
        _kernel32 = ctypes.windll.kernel32
        _tid     = _kernel32.GetCurrentThreadId()
        app._pipette_thread_tid = _tid  # Store for shutdown
        _hook_id = [None]

        _HOOKPROC = ctypes.WINFUNCTYPE(
            ctypes.c_longlong,
            ctypes.c_int,
            ctypes.c_size_t,
            ctypes.c_size_t,
        )

        # Latest mouse position used by the capture thread.
        # The hook stores coordinates and returns immediately.
        _pip_pos   = [0, 0]
        _pip_event = threading.Event()

        def _capture_loop():
            """Pixel capture via BitBlt — faster than GetPixel, which forces
            DWM to synchronise the entire framebuffer."""
            _gdi32  = ctypes.windll.gdi32
            _u32cap = ctypes.windll.user32
            # All GDI/User32 functions receiving or returning handles need explicit
            # restype + argtypes to avoid 64-bit integer truncation.
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
            # 1x1 bitmap for the memory DC
            hbmp_1px = _gdi32.CreateCompatibleBitmap(hdc_screen, 1, 1)
            _gdi32.SelectObject(hdc_mem, hbmp_1px)
            try:
                while app.pipette_active:
                    _pip_event.wait(timeout=0.05)
                    _pip_event.clear()
                    if not app.pipette_active:
                        break
                    x, y = _pip_pos
                    try:
                        # SRCCOPY=0xCC0020: copy 1x1 pixel from screen to memory DC
                        _gdi32.BitBlt(hdc_mem, 0, 0, 1, 1, hdc_screen, x, y, 0xCC0020)
                        pixel = _gdi32.GetPixel(hdc_mem, 0, 0)
                        if pixel != -1:
                            r2 = pixel & 0xFF
                            g2 = (pixel >> 8) & 0xFF
                            b2 = (pixel >> 16) & 0xFF
                            with _pip_lock:
                                app.pip_color = [r2, g2, b2, 255]
                    except Exception:
                        pass
            finally:
                _gdi32.DeleteObject(hbmp_1px)
                _gdi32.DeleteDC(hdc_mem)
                _u32cap.ReleaseDC(0, hdc_screen)

        threading.Thread(target=_capture_loop, daemon=True).start()

        def _hook_proc(nCode, wParam, lParam):
            if nCode >= 0:
                ms = ctypes.cast(lParam, ctypes.POINTER(_MSLLHOOKSTRUCT)).contents
                if wParam == WM_MOUSEMOVE:
                    # Store coordinates and signal the capture thread.
                    # Minimal work here so the WH_MOUSE_LL hook returns immediately.
                    _pip_pos[0] = ms.x
                    _pip_pos[1] = ms.y
                    _pip_event.set()
                elif wParam == WM_LBUTTONDOWN:
                    app.pipette_active = False
                    _pip_event.set()  # wake capture thread for shutdown
                    with _pip_lock:
                        c = list(app.pip_color)
                    app._pending_pipette_color = [c[0]/255., c[1]/255., c[2]/255., 1.]
                    # Suppress click and exit the message loop
                    _user32.PostThreadMessageW(_tid, WM_QUIT, 0, 0)
                    return 1  # suppress — not forwarded to other windows

            return _user32.CallNextHookEx(_hook_id[0], nCode, wParam, lParam)

        _cb = _HOOKPROC(_hook_proc)
        # Keep a reference on the app object to prevent premature GC.
        # Consistent with app._pip_wndproc_cb.
        app._pipette_hook_cb = _cb
        # Set argtypes explicitly for this call — avoids TypeError from
        # stale cached argtypes.
        _user32.SetWindowsHookExW.argtypes = [
            ctypes.c_int, _HOOKPROC, ctypes.c_void_p, ctypes.c_uint
        ]
        _user32.SetWindowsHookExW.restype = ctypes.c_void_p
        # CallNextHookEx passes wParam/lParam as pointer-sized values.
        # Without argtypes ctypes defaults to c_int and overflows on 64-bit Windows.
        _user32.CallNextHookEx.restype  = ctypes.c_longlong
        _user32.CallNextHookEx.argtypes = [
            ctypes.c_void_p,   # hhk  (HHOOK, pointer-sized)
            ctypes.c_int,      # nCode
            ctypes.c_size_t,   # wParam (WPARAM = UINT_PTR)
            ctypes.c_ssize_t,  # lParam (LPARAM = LONG_PTR)
        ]
        _hook_id[0] = _user32.SetWindowsHookExW(WH_MOUSE_LL, _cb, None, 0)

        # Message loop — required for WH_MOUSE_LL; hook does not fire without it
        _msg = _wt.MSG()
        while _user32.GetMessageW(ctypes.byref(_msg), None, 0, 0) > 0:
            _user32.TranslateMessage(ctypes.byref(_msg))
            _user32.DispatchMessageW(ctypes.byref(_msg))

        if _hook_id[0]:
            _user32.UnhookWindowsHookEx(_hook_id[0])
            _hook_id[0] = None
        app._pipette_thread_tid = 0

    _pipette_lock = threading.Lock()

    def start_pipette():
        """Start the eyedropper. Guards against multiple concurrent instances."""
        # Multiple simultaneous threads each creating their own overlay window
        # causes a Win32-level crash.
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

    # -- HTML export ------------------------------------------------------
    _HCSS = (
        "<script>"
        "function cc(t,b){"
        "navigator.clipboard.writeText(t).then(()=>{"
        "var o=b.innerText;b.innerText='Copied!';b.style.background='#4CAF50';"
        "setTimeout(()=>{b.innerText=o;b.style.background=''},1200)})}"
        "</script>"
        "<style>"
        "body{font-family:'Segoe UI',sans-serif;background:#111;color:#ddd;padding:30px;margin:0}"
        "h2{margin-bottom:18px;font-weight:600;letter-spacing:.03em}"
        "table{border-collapse:collapse;background:#1a1a1a;white-space:nowrap}"
        "td,th{border:1px solid #2a2a2a;padding:7px 10px;font-size:13px;vertical-align:middle}"
        "th{background:#222;color:#aaa;font-weight:600;text-transform:uppercase;"
        "font-size:11px;letter-spacing:.06em;white-space:nowrap}"
        "tr:hover td{background:#222}"
        ".sw{width:52px;height:30px;border-radius:5px;border:1px solid #333;display:block}"
        ".cb{cursor:pointer;padding:3px 10px;background:#2e2e2e;color:#bbb;"
        "border:1px solid #444;border-radius:4px;font-size:11px;display:block;width:100%;"
        "box-sizing:border-box;margin-top:2px}"
        ".cb:first-child{margin-top:0}"
        ".cb:hover{background:#3d3d3d;color:#fff;border-color:#666}"
        ".mono{font-family:'Consolas','Courier New',monospace;font-size:12px}"
        ".wcag-aaa{color:#4caf50;font-weight:700}"
        ".wcag-aa{color:#8bc34a;font-weight:700}"
        ".wcag-aa-lg{color:#cddc39}"
        ".wcag-fail{color:#e05555}"
        ".ctr-row{display:flex;align-items:center;gap:6px;margin-bottom:3px}"
        ".ctr-bar{display:inline-block;padding:1px 8px;border-radius:3px;"
        "font-size:13px;font-weight:600;white-space:nowrap}"
        ".ctr-bar-w{background:#ffffff;border:1px solid #555}"
        ".ctr-bar-b{background:#000000;border:1px solid #555}"
        "</style>"
    )

    def _html_color_row(r, g, b):
        """Return an HTML <tr> with colour data columns and a Copy button."""
        hx  = f"#{r:02x}{g:02x}{b:02x}".upper()
        rgb = f"rgb({r}, {g}, {b})"

        h, l, s   = colorsys.rgb_to_hls(r/255, g/255, b/255)
        hsl = f"hsl({h*360:.0f}, {s*100:.0f}%, {l*100:.0f}%)"

        h2, s2, v2 = colorsys.rgb_to_hsv(r/255, g/255, b/255)
        hsv = f"hsv({h2*360:.0f}, {s2*100:.0f}%, {v2*100:.0f}%)"

        c_, m_, y_, k_ = rgb_to_cmyk(r, g, b)
        cmyk = f"C{c_:.0f}% M{m_:.0f}% Y{y_:.0f}% K{k_:.0f}%"

        css = nearest_css(r, g, b)

        cw = contrast_ratio(r, g, b, 255, 255, 255)
        cb = contrast_ratio(r, g, b, 0,   0,   0  )

        def _wcag_cls(ratio):
            if ratio >= 7:   return "wcag-aaa", "AAA"
            if ratio >= 4.5: return "wcag-aa",  "AA"
            if ratio >= 3:   return "wcag-aa-lg","AA*"
            return "wcag-fail", "Fail"

        cw_cls, cw_lbl = _wcag_cls(cw)
        cb_cls, cb_lbl = _wcag_cls(cb)

        def _btn(val):
            return f"<button class=\"cb\" onclick=\"cc(\'{val}\',this)\">Copy</button>"

        return (
            f"<tr>"
            f"<td><span class=\"sw\" style=\"background:{hx}\"></span></td>"
            f"<td class=\"mono\">{hx} {_btn(hx)}</td>"
            f"<td class=\"mono\">{rgb}</td>"
            f"<td class=\"mono\">{hsl}</td>"
            f"<td class=\"mono\">{hsv}</td>"
            f"<td class=\"mono\">{cmyk}</td>"
            f"<td>{css}</td>"
            f"<td>"
            f"<div class=\"ctr-row\"><span class=\"{cw_cls}\">W &nbsp;{cw:.2f}:1 &nbsp;{cw_lbl}</span>"
            f"<span class=\"ctr-bar ctr-bar-w\" style=\"color:{hx}\">Text</span></div>"
            f"<div class=\"ctr-row\"><span class=\"{cb_cls}\">B &nbsp;{cb:.2f}:1 &nbsp;{cb_lbl}</span>"
            f"<span class=\"ctr-bar ctr-bar-b\" style=\"color:{hx}\">Text</span></div>"
            f"</td>"
            f"</tr>"
        )

    def _html_table_header():
        return (
            "<tr>"
            "<th>Preview</th><th>HEX</th><th>RGB</th>"
            "<th>HSL</th><th>HSV</th><th>CMYK</th>"
            "<th>CSS Name</th><th>Contrast (WCAG W/B)</th>"
            "</tr>"
        )

    def _save_html_dialog(content, default_name):
        def _run():
            try:
                # Win32 save dialog — no Tkinter, thread-safe
                path = _win32_save_file(
                    "Save as",
                    [("HTML file", "*.html"), ("All files", "*.*")],
                    default_ext=".html",
                    initial_name=default_name,
                )
                if path:
                    title = os.path.splitext(os.path.basename(path))[0]
                    final = content.replace(
                        'HTMLTITLE_PLACEHOLDER',
                        f'<title>{title}</title>'
                    ).replace('HTMLH2_PLACEHOLDER', title)
                    with open(path, 'w', encoding='utf-8') as f:
                        f.write(final)
                    os.startfile(path)
            except Exception as e:
                _write_log(f"Save error: {e}\n" + traceback.format_exc())

        threading.Thread(target=_run, daemon=True).start()

    def _save_ase_dialog(colors, default_name):
        """Save colours as an ASE file in a background thread."""
        def _build_ase(palette_name, color_list):
            # Standalone color entries without group wrapper.
            # Group blocks (0xC002/0xC003) cause import failures in several apps.
            # Color type 2 = normal (0=global, 1=spot); normal has widest compatibility.
            def _enc_name(n):
                return struct.pack('>H', len(n) + 1) + (n + '\x00').encode('utf-16-be')

            entries = []
            for c in color_list:
                rf = c[0] / 255.0
                gf = c[1] / 255.0
                bf = c[2] / 255.0
                label = f"#{c[0]:02X}{c[1]:02X}{c[2]:02X}"
                block_data = (
                    _enc_name(label)
                    + b'RGB '
                    + struct.pack('>fff', rf, gf, bf)
                    + struct.pack('>H', 2)   # 2 = normal
                )
                entries.append(struct.pack('>HI', 0x0001, len(block_data)) + block_data)

            header = b'ASEF' + struct.pack('>HH', 1, 0) + struct.pack('>I', len(entries))
            return header + b''.join(entries)

        def _run():
            try:
                # Win32 save dialog — no Tkinter, thread-safe
                path = _win32_save_file(
                    "Export ASE",
                    [("Adobe Swatch Exchange", "*.ase"), ("All files", "*.*")],
                    default_ext=".ase",
                    initial_name=default_name,
                )
                if path:
                    pal_name = os.path.splitext(os.path.basename(path))[0]
                    data = _build_ase(pal_name, colors)
                    with open(path, 'wb') as f:
                        f.write(data)
            except Exception as e:
                _write_log(f"ASE export error: {e}\n" + traceback.format_exc())

        threading.Thread(target=_run, daemon=True).start()

    def export_palette_ase_by_name(name):
        with _palettes_lock:
            if not name or name not in app.palettes:
                return
            colors = list(app.palettes[name])
        _save_ase_dialog(colors, name)

    def export_palette():
        h = (
            f"<html><head><meta charset='utf-8'>HTMLTITLE_PLACEHOLDER{_HCSS}</head><body>"
            f"<h2>HTMLH2_PLACEHOLDER</h2>"
            f"<table>{_html_table_header()}"
        )
        for tag in HARMONY_TAGS:
            if not dpg.is_item_shown(f"group_{tag}"):
                continue
            rgb = app.harmony_rgb.get(tag)
            if not rgb:
                continue
            r, g, b = rgb
            h += _html_color_row(r, g, b)
        h += "</table></body></html>"
        _save_html_dialog(h, "harmony_palette")

    def export_history():
        with _history_lock:
            hist_snapshot = list(app.history)
        if not hist_snapshot:
            return
        h = (
            f"<html><head><meta charset='utf-8'>HTMLTITLE_PLACEHOLDER{_HCSS}</head><body>"
            f"<h2>HTMLH2_PLACEHOLDER</h2>"
            f"<table>{_html_table_header()}"
        )
        for v in hist_snapshot:
            h += _html_color_row(v[0], v[1], v[2])
        h += "</table></body></html>"
        _save_html_dialog(h, "color_history")

    def _minimize_win():
        """Minimize the window via Win32 (dpg.minimize_viewport is not always available)."""
        try:
            hwnd = _get_own_hwnd()
            if hwnd:
                ctypes.windll.user32.ShowWindow(hwnd, 6)  # SW_MINIMIZE
        except Exception:
            _write_log("_minimize_win error:\n" + traceback.format_exc())

    def set_theme(s, v):
        name = v if v in _THEME_MAP else "Dark"
        app.theme_name = name
        app_th, tb_th, is_light = _THEME_MAP[name]
        dpg.bind_theme(app_th)
        if dpg.does_item_exist("tb_logo"):
            new_tex = "logo_light" if is_light else "logo_dark"
            new_w   = _logo_light_w if is_light else _logo_dark_w
            new_h   = _logo_light_h if is_light else _logo_dark_h
            if new_w > 0:
                dpg.configure_item("tb_logo", texture_tag=new_tex, width=new_w, height=new_h)
        if dpg.does_item_exist("titlebar_cw"):
            dpg.bind_item_theme("titlebar_cw", tb_th)
        # Update history selection border colour for the new theme.
        border_col = [60, 60, 60, 255] if is_light else [220, 220, 220, 255]
        try:
            dpg.delete_item(_hist_selected_theme, children_only=True)
            with dpg.theme_component(dpg.mvColorButton, parent=_hist_selected_theme):
                dpg.add_theme_color(dpg.mvThemeCol_Border, border_col,
                                    category=dpg.mvThemeCat_Core)
                dpg.add_theme_style(dpg.mvStyleVar_FrameBorderSize, 2)
        except Exception:
            pass

    def slider_mode_cb(s, v):
        app.slider_mode = v
        for m in SLIDER_MODES:
            dpg.configure_item(f"grp_sl_{m}", show=(m == v))
        app._last_sl_vals = {}

    def harmony_combo_cb(s, v):
        app.harmony_mode = v

    def fmt_combo_cb(s, v):
        app.fmt_mode = v
        refresh_harmony_values()

    def toggle_mode(s, v):
        app.use_wheel = not app.use_wheel
        if app.use_wheel:
            dpg.configure_item("grp_wheel",       show=True)
            dpg.configure_item("grp_sliders_all", show=False)
            dpg.set_item_label("btn_mode", "Sliders")
        else:
            dpg.configure_item("grp_wheel",       show=False)
            dpg.configure_item("grp_sliders_all", show=True)
            dpg.set_item_label("btn_mode", "Color Wheel")
            _sync_all_modes()
            app._last_sl_vals = {}

    # ─────────────────────────────────────────────────────────────────────
    #  UI BUILD
    # ─────────────────────────────────────────────────────────────────────
    dpg.create_context()

    # Mouse wheel handler via DPG — works only while the viewport is active
    with dpg.handler_registry():
        dpg.add_mouse_wheel_handler(
            callback=lambda s, a, u: setattr(app, '_mouse_wheel_delta', a)
        )

    # ── Logo textures ──────────────────────────────────────────────────────────────
    _TB_H        = _sc(34)   # total titlebar height — compact, like a native bar
    _logo_dark_w  = 0
    _logo_dark_h  = 0
    _logo_light_w = 0
    _logo_light_h = 0

    with dpg.texture_registry():
        for _tex_tag, _fname, _wattr in [
            ("logo_dark",  "colortoolsd.png", "_logo_dark_w"),
            ("logo_light", "colortoolsw.png", "_logo_light_w"),
        ]:
            _path = os.path.join(_SCRIPT_DIR, _fname)
            try:
                _lw, _lh, _, _ldata = dpg.load_image(_path)
                dpg.add_static_texture(_lw, _lh, _ldata, tag=_tex_tag)
                # Render at native size but cap height to titlebar
                _scale  = min(1.0, (_TB_H - _sc(4)) / max(1, _lh))
                _rw     = max(1, int(_lw * _scale))
                _rh     = max(1, int(_lh * _scale))
                if _wattr == "_logo_dark_w":
                    _logo_dark_w  = _rw;  _logo_dark_h  = _rh
                else:
                    _logo_light_w = _rw;  _logo_light_h = _rh
            except Exception:
                dpg.add_static_texture(1, 1, [0, 0, 0, 0], tag=_tex_tag)

    if _logo_light_w == 0 and _logo_dark_w > 0:
        _logo_light_w = _logo_dark_w;  _logo_light_h = _logo_dark_h

    # Icon button dimensions — kept small
    _ICON_W      = _sc(28)
    _IC          = _sc(6)    # half-length of drawn lines
    _THICK_ICON  = _sc(2)    # line thickness

    _VP_W        = int(283 * _DPI_SCALE)

    with dpg.theme() as _nospc_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing,  0, 0)
            dpg.add_theme_style(dpg.mvStyleVar_FramePadding, 0, 0)

    with dpg.theme() as _dark_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_WindowBg,       [26,  26,  26,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ChildBg,        [26,  26,  26,  255])
            dpg.add_theme_color(dpg.mvThemeCol_FrameBg,        [45,  45,  45,  255])
            dpg.add_theme_color(dpg.mvThemeCol_FrameBgHovered, [55,  55,  55,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Text,           [220, 220, 220, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Button,         [60,  60,  60,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered,  [80,  80,  80,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Header,         [60,  60,  60,  255])
            dpg.add_theme_color(dpg.mvThemeCol_PopupBg,        [30,  30,  30,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ScrollbarBg,    [20,  20,  20,  255])
            dpg.add_theme_style(dpg.mvStyleVar_TabRounding,    2)
            dpg.add_theme_style(dpg.mvStyleVar_FramePadding,   4, 3)

    with dpg.theme() as _light_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_WindowBg,           [235, 235, 235, 255])
            dpg.add_theme_color(dpg.mvThemeCol_ChildBg,            [235, 235, 235, 255])
            dpg.add_theme_color(dpg.mvThemeCol_FrameBg,            [210, 210, 210, 255])
            dpg.add_theme_color(dpg.mvThemeCol_FrameBgHovered,     [195, 195, 195, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Text,               [30,  30,  30,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Button,             [200, 200, 200, 255])
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered,      [180, 180, 180, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Header,             [190, 190, 190, 255])
            dpg.add_theme_color(dpg.mvThemeCol_PopupBg,            [230, 230, 230, 255])
            dpg.add_theme_color(dpg.mvThemeCol_ScrollbarBg,        [220, 220, 220, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Separator,          [160, 160, 160, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Tab,                [200, 200, 200, 255])
            dpg.add_theme_color(dpg.mvThemeCol_TabHovered,         [66,  150, 250, 200])
            dpg.add_theme_color(dpg.mvThemeCol_TabActive,          [66,  150, 250, 255])
            dpg.add_theme_color(dpg.mvThemeCol_TabUnfocused,       [200, 200, 200, 255])
            dpg.add_theme_color(dpg.mvThemeCol_TabUnfocusedActive, [66,  150, 250, 255])
            dpg.add_theme_style(dpg.mvStyleVar_TabRounding,  2)
            dpg.add_theme_style(dpg.mvStyleVar_FramePadding, 4, 3)

    with dpg.theme() as _midnight_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_WindowBg,       [13,  17,  23,  255])  # #0d1117
            dpg.add_theme_color(dpg.mvThemeCol_ChildBg,        [13,  17,  23,  255])
            dpg.add_theme_color(dpg.mvThemeCol_FrameBg,        [22,  30,  40,  255])
            dpg.add_theme_color(dpg.mvThemeCol_FrameBgHovered, [30,  42,  56,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Text,           [201, 209, 217, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Button,         [31,  54,  84,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered,  [48,  84,  130, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Header,         [31,  54,  84,  255])
            dpg.add_theme_color(dpg.mvThemeCol_PopupBg,        [16,  22,  30,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ScrollbarBg,    [10,  13,  18,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Tab,            [22,  30,  40,  255])
            dpg.add_theme_color(dpg.mvThemeCol_TabHovered,     [48,  84,  130, 200])
            dpg.add_theme_color(dpg.mvThemeCol_TabActive,      [31,  54,  84,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Separator,      [30,  42,  56,  255])
            dpg.add_theme_style(dpg.mvStyleVar_TabRounding,    2)
            dpg.add_theme_style(dpg.mvStyleVar_FramePadding,   4, 3)

    with dpg.theme() as _mocha_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_WindowBg,       [30,  26,  23,  255])  # warm dark brown
            dpg.add_theme_color(dpg.mvThemeCol_ChildBg,        [30,  26,  23,  255])
            dpg.add_theme_color(dpg.mvThemeCol_FrameBg,        [50,  42,  36,  255])
            dpg.add_theme_color(dpg.mvThemeCol_FrameBgHovered, [65,  54,  46,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Text,           [225, 210, 195, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Button,         [90,  65,  45,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered,  [120, 88,  60,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Header,         [90,  65,  45,  255])
            dpg.add_theme_color(dpg.mvThemeCol_PopupBg,        [38,  32,  28,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ScrollbarBg,    [22,  18,  15,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Tab,            [50,  42,  36,  255])
            dpg.add_theme_color(dpg.mvThemeCol_TabHovered,     [120, 88,  60,  200])
            dpg.add_theme_color(dpg.mvThemeCol_TabActive,      [90,  65,  45,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Separator,      [65,  54,  46,  255])
            dpg.add_theme_style(dpg.mvStyleVar_TabRounding,    2)
            dpg.add_theme_style(dpg.mvStyleVar_FramePadding,   4, 3)

    with dpg.theme() as _nord_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_WindowBg,       [46,  52,  64,  255])  # Nord Polar Night
            dpg.add_theme_color(dpg.mvThemeCol_ChildBg,        [46,  52,  64,  255])
            dpg.add_theme_color(dpg.mvThemeCol_FrameBg,        [59,  66,  82,  255])
            dpg.add_theme_color(dpg.mvThemeCol_FrameBgHovered, [67,  76,  94,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Text,           [216, 222, 233, 255])  # Nord Snow Storm
            dpg.add_theme_color(dpg.mvThemeCol_Button,         [67,  76,  94,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered,  [76,  86,  106, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Header,         [94,  129, 172, 255])  # Nord Frost blue
            dpg.add_theme_color(dpg.mvThemeCol_PopupBg,        [59,  66,  82,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ScrollbarBg,    [36,  41,  51,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Tab,            [59,  66,  82,  255])
            dpg.add_theme_color(dpg.mvThemeCol_TabHovered,     [94,  129, 172, 200])
            dpg.add_theme_color(dpg.mvThemeCol_TabActive,      [58,  90,  130, 255])  # darkened Nord Frost
            dpg.add_theme_color(dpg.mvThemeCol_Separator,      [67,  76,  94,  255])
            dpg.add_theme_style(dpg.mvStyleVar_TabRounding,    2)
            dpg.add_theme_style(dpg.mvStyleVar_FramePadding,   4, 3)

    with dpg.theme() as _solarized_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_WindowBg,       [0,   43,  54,  255])  # Solarized base03
            dpg.add_theme_color(dpg.mvThemeCol_ChildBg,        [0,   43,  54,  255])
            dpg.add_theme_color(dpg.mvThemeCol_FrameBg,        [7,   54,  66,  255])  # base02
            dpg.add_theme_color(dpg.mvThemeCol_FrameBgHovered, [0,   72,  88,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Text,           [253, 246, 227, 255])  # base3 — sufficient contrast on the active tab
            dpg.add_theme_color(dpg.mvThemeCol_Button,         [0,   72,  88,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered,  [0,   95,  115, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Header,         [42,  161, 152, 255])  # cyan accent
            dpg.add_theme_color(dpg.mvThemeCol_PopupBg,        [7,   54,  66,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ScrollbarBg,    [0,   33,  42,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Tab,            [7,   54,  66,  255])
            dpg.add_theme_color(dpg.mvThemeCol_TabHovered,     [42,  161, 152, 200])
            dpg.add_theme_color(dpg.mvThemeCol_TabActive,      [0,   95,  89,  255])  # darkened cyan — distinct from inactive tabs, contrasts with base3 text
            dpg.add_theme_color(dpg.mvThemeCol_Separator,      [0,   72,  88,  255])
            dpg.add_theme_style(dpg.mvStyleVar_TabRounding,    2)
            dpg.add_theme_style(dpg.mvStyleVar_FramePadding,   4, 3)

    with dpg.theme() as _new_palette_select_theme:
        with dpg.theme_component(dpg.mvButton):
            dpg.add_theme_color(dpg.mvThemeCol_Button,        [0, 150, 0, 255])
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, [0, 180, 0, 255])
            dpg.add_theme_color(dpg.mvThemeCol_ButtonActive,  [0, 200, 0, 255])


    # History selection theme created once globally.
    with dpg.theme() as _hist_selected_theme:
        with dpg.theme_component(dpg.mvColorButton):
            dpg.add_theme_color(
                dpg.mvThemeCol_Border, [220, 220, 220, 255],
                category=dpg.mvThemeCat_Core,
            )
            dpg.add_theme_style(dpg.mvStyleVar_FrameBorderSize, 2)

    # ── Titlebar themes ────────────────────────────────────────────────────────────
    # Background matches the app WindowBg so there is no visible seam.
    with dpg.theme() as _titlebar_dark_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_ChildBg,        [26, 26, 26, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Border,         [26, 26, 26, 255])
            dpg.add_theme_style(dpg.mvStyleVar_WindowPadding,  0, _sc(5))
            dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing,    _sc(2), 0)

    with dpg.theme() as _titlebar_light_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_ChildBg,        [235, 235, 235, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Border,         [235, 235, 235, 255])
            dpg.add_theme_style(dpg.mvStyleVar_WindowPadding,  0, _sc(5))
            dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing,    _sc(2), 0)

    with dpg.theme() as _titlebar_midnight_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_ChildBg,        [13, 17, 23, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Border,         [13, 17, 23, 255])
            dpg.add_theme_style(dpg.mvStyleVar_WindowPadding,  0, _sc(5))
            dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing,    _sc(2), 0)

    with dpg.theme() as _titlebar_mocha_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_ChildBg,        [30, 26, 23, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Border,         [30, 26, 23, 255])
            dpg.add_theme_style(dpg.mvStyleVar_WindowPadding,  0, _sc(5))
            dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing,    _sc(2), 0)

    with dpg.theme() as _titlebar_nord_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_ChildBg,        [46, 52, 64, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Border,         [46, 52, 64, 255])
            dpg.add_theme_style(dpg.mvStyleVar_WindowPadding,  0, _sc(5))
            dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing,    _sc(2), 0)

    with dpg.theme() as _titlebar_solarized_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_ChildBg,        [0, 43, 54, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Border,         [0, 43, 54, 255])
            dpg.add_theme_style(dpg.mvStyleVar_WindowPadding,  0, _sc(5))
            dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing,    _sc(2), 0)

    _THEME_MAP = {
        "Dark":      (_dark_theme,      _titlebar_dark_theme,      False),
        "Light":     (_light_theme,     _titlebar_light_theme,     True),
        "Midnight":  (_midnight_theme,  _titlebar_midnight_theme,  False),
        "Mocha":     (_mocha_theme,     _titlebar_mocha_theme,     False),
        "Nord":      (_nord_theme,      _titlebar_nord_theme,      False),
        "Solarized": (_solarized_theme, _titlebar_solarized_theme, False),
    }
    THEME_NAMES = list(_THEME_MAP.keys())

    # Dark background for the Contrast display — overrides light theme so WCAG
    # labels are always legible on both dark and light global themes.
    with dpg.theme() as _contrast_bg_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_ChildBg, [26, 26, 26, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Text,    [210, 210, 210, 255])
            dpg.add_theme_style(dpg.mvStyleVar_WindowPadding, 8, 3)
            dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing,   6, 3)


    def _make_grad_slider(label, label_color, gname):
        with dpg.group(horizontal=True):
            dpg.add_text(label, color=label_color)
            dpg.add_drawlist(width=GRAD_W, height=THUMB_H, tag=f"dl_{gname}")
            dpg.add_text("   ", tag=f"slval_{gname}")

    # ── Win32 eyedropper overlay ──────────────────────────────────────────
    # Created once on the main thread. The update() loop moves and repaints it.

    _pip_u32 = ctypes.windll.user32
    _pip_g32 = ctypes.windll.gdi32
    # Overlay dimensions scaled by _DPI_SCALE for high-DPI displays.
    _PIP_OW       = max(130, int(130 * _DPI_SCALE))
    _PIP_OH       = max(28,  int(28  * _DPI_SCALE))
    _PIP_SW       = max(28,  int(28  * _DPI_SCALE))  # swatch area width
    _pip_hwnd = [None]
    _pip_cls  = [None]

    # ── Pre-compute static pip overlay pixel template ────────────────────────
    # Background: dark semi-transparent (pre-multiplied alpha).
    # Text area (tx >= 28): black with full alpha so GDI DrawText is visible.
    # Only the swatch area changes per frame — template is copied each call.
    _PIP_BG_A   = 185
    _PIP_BG_VAL = (
        (20 * _PIP_BG_A // 255)
        | ((20 * _PIP_BG_A // 255) << 8)
        | ((20 * _PIP_BG_A // 255) << 16)
        | (_PIP_BG_A << 24)
    )
    # Pre-computed per-pixel byte patterns reused by _pip_draw every frame.
    _PIP_BG_BYTES = _PIP_BG_VAL.to_bytes(4, "little")
    _PIP_WH_BYTES = b"\xff\xff\xff\xff"
    _pip_pixel_template = (ctypes.c_uint32 * (_PIP_OW * _PIP_OH))()
    for _i in range(_PIP_OW * _PIP_OH):
        _pip_pixel_template[_i] = _PIP_BG_VAL
    for _ty in range(_PIP_OH):
        for _tx in range(_PIP_SW, _PIP_OW):
            _pip_pixel_template[_ty * _PIP_OW + _tx] = 0xFF000000

    def _pip_wndproc(hwnd, msg, wp, lp):
        # UpdateLayeredWindow windows do not receive WM_PAINT; all drawing
        # is done manually from the update() loop via _pip_draw().
        if msg == 0x0002:  # WM_DESTROY
            return 0
        return _pip_u32.DefWindowProcW(hwnd, msg, wp, lp)

    _PIP_WNDPROC_T = ctypes.WINFUNCTYPE(
        ctypes.c_ssize_t, ctypes.c_void_p,
        ctypes.c_uint, ctypes.c_ssize_t, ctypes.c_ssize_t,
    )
    # Set argtypes for DefWindowProcW — without them 64-bit LPARAM overflows
    _pip_u32.DefWindowProcW.restype  = ctypes.c_ssize_t
    _pip_u32.DefWindowProcW.argtypes = [
        ctypes.c_void_p, ctypes.c_uint,
        ctypes.c_ssize_t, ctypes.c_ssize_t,
    ]
    # Keep a reference on the app object to prevent premature GC.
    app._pip_wndproc_cb = _PIP_WNDPROC_T(_pip_wndproc)

    # Set restype before calling — HWND is pointer-sized on 64-bit Windows.
    # The default c_int truncates the upper 32 bits, producing a corrupt handle.
    _pip_u32.RegisterClassW.restype   = ctypes.c_uint16   # ATOM (WORD)
    _pip_u32.RegisterClassW.argtypes  = [ctypes.c_void_p]
    _pip_u32.CreateWindowExW.restype  = ctypes.c_void_p   # HWND (pointer-sized)
    _pip_u32.CreateWindowExW.argtypes = [
        ctypes.c_ulong,    # dwExStyle
        ctypes.c_wchar_p,  # lpClassName
        ctypes.c_wchar_p,  # lpWindowName
        ctypes.c_ulong,    # dwStyle
        ctypes.c_int,      # X
        ctypes.c_int,      # Y
        ctypes.c_int,      # nWidth
        ctypes.c_int,      # nHeight
        ctypes.c_void_p,   # hWndParent
        ctypes.c_void_p,   # hMenu
        ctypes.c_void_p,   # hInstance
        ctypes.c_void_p,   # lpParam
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
    # GetModuleHandleW returns an HMODULE (pointer-sized). Without restype
    # ctypes defaults to c_int and truncates the upper 32 bits.
    ctypes.windll.kernel32.GetModuleHandleW.restype = ctypes.c_void_p
    _pip_hInst = ctypes.windll.kernel32.GetModuleHandleW(None)
    _pip_cls[0] = f"PipOvl_{os.getpid()}"
    _pip_wc = _PipWNDCLASS()
    _pip_wc.lpfnWndProc   = ctypes.cast(app._pip_wndproc_cb, ctypes.c_void_p).value
    _pip_wc.hInstance     = _pip_hInst
    _pip_wc.lpszClassName = _pip_cls[0]
    if _pip_u32.RegisterClassW(ctypes.byref(_pip_wc)):
        _pip_hwnd[0] = _pip_u32.CreateWindowExW(
            0x00000008 | 0x00000080 | 0x08000000 | 0x00080000,  # TOPMOST|TOOLWINDOW|NOACTIVATE|LAYERED
            _pip_cls[0], "", 0x80000000,           # WS_POPUP
            0, 0, _PIP_OW, _PIP_OH,
            None, None, _pip_hInst, None,
        )
        if _pip_hwnd[0]:
            pass  # Content drawn via UpdateLayeredWindow in _pip_overlay_update()

    _pip_pt = _wt.POINT()

    # ── UpdateLayeredWindow structures ──────────────────────────────
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

    # ── Eyedropper overlay: GDI resources initialised ONCE here ──────────────────
    # CreateCompatibleDC / CreateDIBSection per frame is expensive at 30 fps.
    # Initialised once at startup and freed at shutdown.
    _pip_gdi = {
        'hdc_screen': None,
        'hdc_mem':    None,
        'hbmp':       None,
        'bits_ptr':   None,
        'pixels':     None,
        'buf':        None,  # PERF: pre-allocated; reused every frame in _pip_draw
        'ba':         None,  # PERF: pre-allocated bytearray; reused every frame
        'ba_ctype':   None,  # PERF: ctypes view of ba — avoids from_buffer() every frame
    }

    # GDI/User32 argtypes — set once here so _pip_init_gdi() has no overhead.
    _pip_u32.GetDC.restype                  = ctypes.c_void_p
    _pip_u32.GetDC.argtypes                 = [ctypes.c_void_p]
    _pip_u32.ReleaseDC.restype              = ctypes.c_int
    _pip_u32.ReleaseDC.argtypes             = [ctypes.c_void_p, ctypes.c_void_p]
    _pip_g32.CreateCompatibleDC.restype     = ctypes.c_void_p
    _pip_g32.CreateCompatibleDC.argtypes    = [ctypes.c_void_p]
    _pip_g32.CreateCompatibleBitmap.restype  = ctypes.c_void_p
    _pip_g32.CreateCompatibleBitmap.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int]
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
        """Initialise eyedropper GDI resources once before the first _pip_draw call."""
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
            hdc_mem, ctypes.byref(bmi), 0,
            ctypes.byref(bits_ptr), None, 0)
        if hbmp:
            _pip_g32.SelectObject(hdc_mem, hbmp)
            _pip_gdi['hdc_screen'] = hdc_screen
            _pip_gdi['hdc_mem']    = hdc_mem
            _pip_gdi['hbmp']       = hbmp
            _pip_gdi['bits_ptr']   = bits_ptr
            _pip_gdi['pixels']     = (ctypes.c_uint32 * (_PIP_OW * _PIP_OH))()
            # Allocated once; reused by _pip_draw every frame (~420 kB/s GC pressure avoided).
            _pip_gdi['buf']        = (ctypes.c_char * (4 * _PIP_OW * _PIP_OH))()
            _pip_gdi['ba']         = bytearray(4 * _PIP_OW * _PIP_OH)
            # Fixed ctypes view of ba — avoids from_buffer() allocation every frame.
            _pip_gdi['ba_ctype']   = (ctypes.c_char * (4 * _PIP_OW * _PIP_OH)).from_buffer(
                _pip_gdi['ba']
            )
        else:
            _pip_g32.DeleteDC(hdc_mem)
            _pip_u32.ReleaseDC(None, hdc_screen)

    def _pip_free_gdi():
        """Free GDI resources at shutdown."""
        if _pip_gdi['hbmp']:
            _pip_g32.DeleteObject(_pip_gdi['hbmp'])
        if _pip_gdi['hdc_mem']:
            _pip_g32.DeleteDC(_pip_gdi['hdc_mem'])
        if _pip_gdi['hdc_screen']:
            _pip_u32.ReleaseDC(None, _pip_gdi['hdc_screen'])
        for k in _pip_gdi:
            _pip_gdi[k] = None

    def _pip_draw(ox, oy, pr, pg, pb):
        """Draw the eyedropper overlay via UpdateLayeredWindow (per-pixel alpha).
        Background: BGRA (20,20,20,185) — dark semi-transparent.
        Swatch:     BGRA (pb,pg,pr,255) — fully opaque.
        Text:       BGRA (255,255,255,255) — white.
        GDI resources are allocated once in _pip_init_gdi(); no allocations here.
        """
        if not _pip_hwnd[0]:
            return
        _pip_init_gdi()
        hdc_screen = _pip_gdi['hdc_screen']
        hdc_mem    = _pip_gdi['hdc_mem']
        bits_ptr   = _pip_gdi['bits_ptr']
        pixels     = _pip_gdi['pixels']
        if hdc_screen is None:
            return

        # ── 1. Copy the pre-built template: background + black text area ────
        ctypes.memmove(pixels, _pip_pixel_template, ctypes.sizeof(pixels))

        # ── 2. Swatch: fully opaque (alpha=255, pre-mult = value as-is)
        sw_val = pb | (pg << 8) | (pr << 16) | (255 << 24)
        for sy in range(4, _PIP_OH - 4):
            for sx in range(4, _PIP_SW - 4):
                pixels[sy * _PIP_OW + sx] = sw_val

        # ── 3. Copy pixels to DIB memory and draw text via GDI ────────────
        ctypes.memmove(bits_ptr, pixels, ctypes.sizeof(pixels))

        hex_str = f"#{pr:02X}{pg:02X}{pb:02X}"
        _pip_g32.SetBkMode(hdc_mem, 2)       # OPAQUE
        _pip_g32.SetBkColor(hdc_mem, 0x000000)
        _pip_g32.SetTextColor(hdc_mem, 0xFFFFFF)
        rc3 = _wt.RECT(_PIP_SW, 0, _PIP_OW - 2, _PIP_OH)
        _pip_u32.DrawTextW(hdc_mem, hex_str, -1, ctypes.byref(rc3), 0x0025)

        # ── 4. Fix alpha in the text area (optimised: single scan with bitmask)
        # Text pixels (drawn by GDI in white) show up as bright (B-channel > 40).
        # All other pixels in the text area are restored to the background colour.
        # GDI DrawText writes white text on a black background (BGRA format).
        # In the text area (tx >= _PIP_SW): B-channel > 40 => text pixel => 0xFFFFFFFF,
        # otherwise restore background.
        # Note: DIB format is BGRA (little-endian), so byte order is B,G,R,A.
        # PERF: uses pre-allocated buffers — no allocation per frame
        buf = _pip_gdi['buf']
        ba  = _pip_gdi['ba']
        ctypes.memmove(buf, bits_ptr, ctypes.sizeof(buf))
        ba[:] = buf
        _bg_bytes = _PIP_BG_BYTES
        _wh_bytes = _PIP_WH_BYTES
        row_stride = _PIP_OW * 4
        text_col_start = _PIP_SW * 4
        for ty in range(_PIP_OH):
            row_off = ty * row_stride
            for tx4 in range(text_col_start, row_stride, 4):
                idx = row_off + tx4
                # B-channel is at byte idx (BGRA). GDI writes only
                # bright pixels as white — all channels > 40 when text.
                # Checking only the B-channel is sufficient to distinguish text pixels.
                if ba[idx] > 40:
                    ba[idx:idx+4] = _wh_bytes
                else:
                    ba[idx:idx+4] = _bg_bytes
        # Use the pre-created ctypes view — avoids from_buffer() per frame.
        ctypes.memmove(bits_ptr, _pip_gdi['ba_ctype'], len(ba))

        # UpdateLayeredWindow
        pt_dst = _wt.POINT(ox, oy)
        pt_src = _wt.POINT(0, 0)
        sz     = _wt.SIZE(_PIP_OW, _PIP_OH)
        blend  = _BLENDFUNCTION(0, 0, 255, 1)  # AC_SRC_OVER, AC_SRC_ALPHA
        _pip_u32.UpdateLayeredWindow(
            _pip_hwnd[0], hdc_screen, ctypes.byref(pt_dst),
            ctypes.byref(sz), hdc_mem, ctypes.byref(pt_src),
            0, ctypes.byref(blend), 2,  # ULW_ALPHA=2
        )

    # SetWindowPos argtypes set once — prevents ctypes marshalling errors.
    _pip_u32.SetWindowPos.restype  = ctypes.c_bool
    _pip_u32.SetWindowPos.argtypes = [
        ctypes.c_void_p, ctypes.c_void_p,
        ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
        ctypes.c_uint,
    ]
    # UpdateLayeredWindow takes 64-bit HDC handles; argtypes required to avoid overflow.
    _pip_u32.UpdateLayeredWindow.restype  = ctypes.c_bool
    _pip_u32.UpdateLayeredWindow.argtypes = [
        ctypes.c_void_p,   # hWnd
        ctypes.c_void_p,   # hdcDst  (screen DC)
        ctypes.c_void_p,   # pptDst  (POINT*)
        ctypes.c_void_p,   # psize   (SIZE*)
        ctypes.c_void_p,   # hdcSrc  (memory DC)
        ctypes.c_void_p,   # pptSrc  (POINT*)
        ctypes.c_ulong,    # crKey   (COLORREF)
        ctypes.c_void_p,   # pblend  (BLENDFUNCTION*)
        ctypes.c_ulong,    # dwFlags
    ]
    _pip_u32.ShowWindow.restype  = ctypes.c_bool
    _pip_u32.ShowWindow.argtypes = [ctypes.c_void_p, ctypes.c_int]
    _pip_u32.GetCursorPos.restype  = ctypes.c_bool
    _pip_u32.GetCursorPos.argtypes = [ctypes.c_void_p]
    _pip_u32.GetSystemMetrics.restype  = ctypes.c_int
    _pip_u32.GetSystemMetrics.argtypes = [ctypes.c_int]
    _pip_u32.PeekMessageW.restype  = ctypes.c_bool
    _pip_u32.PeekMessageW.argtypes = [
        ctypes.c_void_p, ctypes.c_void_p,
        ctypes.c_uint, ctypes.c_uint, ctypes.c_uint,
    ]
    _pip_u32.DestroyWindow.restype   = ctypes.c_bool
    _pip_u32.DestroyWindow.argtypes  = [ctypes.c_void_p]
    _pip_u32.UnregisterClassW.restype  = ctypes.c_bool
    _pip_u32.UnregisterClassW.argtypes = [ctypes.c_wchar_p, ctypes.c_void_p]
    _pip_u32.PostQuitMessage.restype   = None
    _pip_u32.PostQuitMessage.argtypes  = [ctypes.c_int]
    _pip_was_visible = [False]  # track overlay visibility state
    _PIP_HWND_TOPMOST = ctypes.c_void_p(-1)                   # HWND_TOPMOST
    _PIP_SWP_FLAGS    = ctypes.c_uint(0x0001 | 0x0002 | 0x0010)  # SWP_NOSIZE|SWP_NOMOVE|SWP_NOACTIVATE

    def _pip_overlay_update():
        """Move and repaint the Win32 overlay. Called from the update() loop."""
        if not _pip_hwnd[0]:
            return
        # Pump the overlay window's own messages
        _msg = _wt.MSG()
        while _pip_u32.PeekMessageW(ctypes.byref(_msg), _pip_hwnd[0], 0, 0, 1) > 0:
            _pip_u32.TranslateMessage(ctypes.byref(_msg))
            _pip_u32.DispatchMessageW(ctypes.byref(_msg))
        if app.pipette_active:
            _pip_u32.GetCursorPos(ctypes.byref(_pip_pt))
            sw = _pip_u32.GetSystemMetrics(0)
            sh = _pip_u32.GetSystemMetrics(1)
            ox = _pip_pt.x + 18
            oy = _pip_pt.y + 18
            if ox + _PIP_OW > sw: ox = _pip_pt.x - _PIP_OW - 4
            if oy + _PIP_OH > sh: oy = _pip_pt.y - _PIP_OH - 4
            with _pip_lock:
                _pr, _pg, _pb = app.pip_color[0], app.pip_color[1], app.pip_color[2]
            _pip_draw(ox, oy, _pr, _pg, _pb)
            # On first show, bring to front of TOPMOST z-order so it stays above
            # the always-on-top main window.
            if not _pip_was_visible[0]:
                _pip_u32.SetWindowPos(_pip_hwnd[0], _PIP_HWND_TOPMOST, 0, 0, 0, 0, _PIP_SWP_FLAGS)
                _pip_was_visible[0] = True
            _pip_u32.ShowWindow(_pip_hwnd[0], 4)  # SW_SHOWNOACTIVATE
        else:
            if _pip_was_visible[0]:
                _pip_was_visible[0] = False
            _pip_u32.ShowWindow(_pip_hwnd[0], 0)  # SW_HIDE

    with dpg.window(tag="PrimaryWindow"):

        # ── Custom titlebar (replaces the native Windows title bar) ──────────────
        with dpg.child_window(
            tag="titlebar_cw", width=_VP_W, height=_TB_H,
            border=False, no_scrollbar=True,
        ):
            dpg.bind_item_theme("titlebar_cw", _titlebar_dark_theme)
            with dpg.group(tag="titlebar", horizontal=True):
                # Logo at native/near-native size for sharpness
                if _logo_dark_w > 0:
                    _tb_pad_y = max(0, (_TB_H - _logo_dark_h) // 2)
                    if _tb_pad_y > 0:
                        dpg.add_spacer(height=_tb_pad_y)
                    dpg.add_image(
                        "logo_dark", tag="tb_logo",
                        width=_logo_dark_w, height=_logo_dark_h,
                    )
                # Spacer pushes control buttons to the right
                dpg.add_spacer(width=2, tag="tb_spacer")  # resized after first render

                # ── Control buttons as drawlists (hover + click handled in update) ──
                dpg.add_drawlist(width=_ICON_W, height=_TB_H, tag="dl_tb_min")
                dpg.add_drawlist(width=_ICON_W, height=_TB_H, tag="dl_tb_close")
        dpg.add_spacer(height=2)

        with dpg.group(horizontal=True):
            dpg.add_button(
                tag="btn_mode", label="Sliders",
                callback=toggle_mode, width=BW, height=THUMB_H,
            )
            dpg.add_button(label="Color picker", callback=start_pipette, width=BW, height=THUMB_H)
        dpg.add_spacer(height=3)

        _PREV_SW = BW
        _PREV_HW = W - BW - _sc(8) - _sc(8) - 20
        with dpg.group(horizontal=True):
            dpg.add_color_button(
                tag="preview", width=_PREV_SW, height=THUMB_H,
                no_border=True, callback=copy_preview,
            )
            dpg.add_input_text(
                tag="preview_hex", width=_PREV_HW, height=THUMB_H,
                default_value="#000000",
                hint="#RRGGBB",
                on_enter=True,
                callback=hex_input_cb,
            )
            dpg.add_button(
                label="C", width=20, height=THUMB_H,
                callback=lambda s, a, u: pyperclip.copy(dpg.get_value("preview_hex")),
            )
        dpg.add_spacer(height=2)

        with dpg.group(tag="grp_wheel", show=True):
            _ir, _ig, _ib = [int(round(c * 255)) for c in app.current_color[:3]]
            dpg.add_color_picker(
                tag="picker_wheel", width=W,
                no_inputs=True, no_label=True,
                no_side_preview=True, no_small_preview=True,
                picker_mode=dpg.mvColorPicker_wheel,
                display_type=dpg.mvColorEdit_uint8,
                display_hsv=True,
                default_value=[_ir, _ig, _ib, 255],
                callback=picker_cb,
            )

        with dpg.group(tag="grp_sliders_all", show=False):
            dpg.add_combo(
                SLIDER_MODES, tag="slider_mode_combo",
                default_value="RGB", width=W, callback=slider_mode_cb,
            )
            dpg.add_spacer(height=3)
            _SLIDER_AREA_H = _sc(215)
            with dpg.child_window(
                width=W, height=_SLIDER_AREA_H, border=False, no_scrollbar=True,
            ):
                for mode, sliders in SLIDERS_BY_MODE.items():
                    with dpg.group(tag=f"grp_sl_{mode}", show=(mode == "RGB")):
                        for (lbl, gname, lcol, smin, smax) in sliders:
                            _make_grad_slider(lbl, lcol, gname)

        with dpg.child_window(
            width=W, height=_sc(232), border=False, no_scrollbar=True,
            no_scroll_with_mouse=True, tag="tab_area",
        ):
            with dpg.tab_bar():

                # -- Harmony ------------------------------------------
                with dpg.tab(label="  Harmony  "):
                    dpg.add_spacer(height=1)
                    with dpg.group(horizontal=True):
                        dpg.add_combo(
                            ["Complementary", "Split Complementary", "Analogous",
                             "Triadic", "Tetradic", "Rectangle", "Tints", "Shades", "Tones"],
                            tag="harmony_combo", default_value="Triadic", width=170,
                            callback=harmony_combo_cb,
                        )
                        dpg.add_combo(
                            FORMAT_OPTIONS, tag="fmt_combo",
                            default_value="HEX", width=89,
                            callback=fmt_combo_cb,
                        )
                    dpg.add_spacer(height=1)
                    # Scrollable area for the 8 colour rows.
                    with dpg.child_window(
                        width=W, height=_sc(147),
                        border=False, no_scrollbar=False,
                        tag="harmony_rows_scroll",
                    ):
                        for tag in ["main", "h1", "h2", "h3", "h4", "h5", "h6", "h7"]:
                            with dpg.group(tag=f"group_{tag}"):
                                with dpg.group(horizontal=True):
                                    dpg.add_color_button(tag=f"rect_{tag}", width=THUMB_H, height=THUMB_H)
                                    dpg.add_input_text(tag=f"val_{tag}", width=VALW, readonly=True)
                                    dpg.add_button(
                                        label="C", width=20, height=THUMB_H,
                                        tag=f"copy_btn_{tag}",
                                        callback=copy_harmony_val, user_data=tag,
                                    )
                                    # Contrast display: dark box with W/B labels on the left,
                                    # white/black preview bars on the right.
                                    _TXT_SZ = _sc(14)
                                    _BAR_W = _sc(75)
                                    _BAR_H = _sc(21)
                                    _CTR_TXT_W = VALW
                                    with dpg.child_window(
                                        tag=f"ctr_grp_{tag}", show=False,
                                        width=_CTR_TXT_W, height=_sc(48),
                                        border=False, no_scrollbar=True,
                                    ):
                                        # W/B ratio labels — normal left-to-right flow
                                        with dpg.group():
                                            with dpg.group(horizontal=True):
                                                dpg.add_spacer(width=_sc(5))
                                                dpg.add_text("W --", tag=f"ctr_w_{tag}", color=[120,120,120])
                                            with dpg.group(horizontal=True):
                                                dpg.add_spacer(width=_sc(5))
                                                dpg.add_text("B --", tag=f"ctr_b_{tag}", color=[120,120,120])
                                        # White and black Text bars — absolutely positioned to the right edge
                                        with dpg.group(tag=f"ctr_bars_{tag}", show=False):
                                            dpg.add_drawlist(tag=f"ctr_bar_w_{tag}", width=_BAR_W, height=_BAR_H)
                                            dpg.draw_rectangle(parent=f"ctr_bar_w_{tag}",
                                                pmin=(0,0), pmax=(_BAR_W, _BAR_H),
                                                fill=[255,255,255,255], color=[0,0,0,0])
                                            dpg.draw_text(parent=f"ctr_bar_w_{tag}",
                                                pos=(_sc(8), (_BAR_H - _TXT_SZ) // 2),
                                                text="Text", tag=f"ctr_bar_w_txt_{tag}",
                                                color=[0,0,0,255], size=_TXT_SZ)
                                            dpg.add_drawlist(tag=f"ctr_bar_b_{tag}", width=_BAR_W, height=_BAR_H)
                                            dpg.draw_rectangle(parent=f"ctr_bar_b_{tag}",
                                                pmin=(0,0), pmax=(_BAR_W, _BAR_H),
                                                fill=[0,0,0,255], color=[0,0,0,0])
                                            dpg.draw_text(parent=f"ctr_bar_b_{tag}",
                                                pos=(_sc(8), (_BAR_H - _TXT_SZ) // 2),
                                                text="Text", tag=f"ctr_bar_b_txt_{tag}",
                                                color=[0,0,0,255], size=_TXT_SZ)
                                        dpg.set_item_pos(f"ctr_bars_{tag}", [_CTR_TXT_W - _BAR_W, 0])
                                    dpg.bind_item_theme(f"ctr_grp_{tag}", _contrast_bg_theme)
                    dpg.add_spacer(height=2)
                    with dpg.group(horizontal=True):
                        dpg.add_button(
                            label="Export HTML",  callback=export_palette,
                            width=(W - _sc(6)) // 2, height=THUMB_H,
                        )
                        dpg.add_button(
                            label="Save Palette", callback=save_harmony_as_palette,
                            width=(W - _sc(6)) // 2, height=THUMB_H,
                        )

                # -- History ------------------------------------------
                with dpg.tab(label="  History  "):
                    with dpg.group(horizontal=True):
                        dpg.add_button(
                            tag="btn_new_palette", label="New Palette",
                            callback=toggle_palette_select_mode,
                            width=90, height=THUMB_H,
                        )
                        dpg.add_button(
                            tag="btn_cancel_palette", label="Cancel",
                            callback=_exit_palette_select_mode,
                            width=60, height=THUMB_H,
                            show=False,
                        )
                    dpg.bind_item_theme("btn_new_palette", None)
                    dpg.add_spacer(height=1)
                    for row in range(6):
                        with dpg.group(horizontal=True) as row_group:
                            dpg.bind_item_theme(row_group, _nospc_theme)
                            for i in range(row * 10, row * 10 + 10):
                                dpg.add_color_button(
                                    tag=f"hist_{i}", width=HSQ, height=HSQ,
                                    show=False,
                                    callback=history_click_cb, user_data=i,
                                )
                                dpg.bind_item_theme(f"hist_{i}", _nospc_theme)
                    dpg.add_spacer(height=2)
                    dpg.add_button(
                        label="Export History (HTML)",
                        callback=export_history, width=W, height=THUMB_H,
                    )

                # -- Palettes -----------------------------------------
                with dpg.tab(label=" Palettes  "):
                    with dpg.group(horizontal=True):
                        dpg.add_button(
                            label="Import from Image",
                            callback=import_image_palette,
                            width=(W - _sc(4)) // 2, height=THUMB_H,
                        )
                        dpg.add_button(
                            label="Import ASE",
                            callback=import_ase_palette,
                            width=(W - _sc(4)) // 2, height=THUMB_H,
                        )
                    dpg.add_text("", tag="import_error_text", color=[220, 80, 80], show=False)
                    dpg.add_spacer(height=1)
                    PAL_SCROLL_H = _sc(181)
                    with dpg.child_window(
                        width=W, height=PAL_SCROLL_H,
                        border=True, no_scrollbar=False,
                        tag="pal_scroll",
                    ):
                        for pi in range(MAX_PAL_ROWS):
                            with dpg.group(tag=f"pal_row_{pi}", show=False):
                                pass

                    with dpg.child_window(
                        width=W, height=PAL_SCROLL_H,
                        border=True, no_scrollbar=False,
                        tag="pal_edit_win", show=False,
                    ):
                        with dpg.group(tag="pal_edit_panel"):
                            pass

                    # Import panel: shown in place of the palette list during image import.
                    with dpg.child_window(
                        width=W, height=PAL_SCROLL_H,
                        border=True, no_scrollbar=True,
                        tag="img_import_win", show=False,
                    ):
                        dpg.add_text("How many colors?  (2 – 20)")
                        dpg.add_spacer(height=3)
                        dpg.add_slider_int(
                            tag="img_import_slider",
                            default_value=8, min_value=2, max_value=20,
                            callback=_image_import_slider_cb,
                            width=W - _sc(20),
                        )
                        dpg.add_spacer(height=5)
                        dpg.add_text("Preview:")
                        with dpg.group(tag="img_import_preview"):
                            pass  # filled by _rebuild_image_import_preview()
                        dpg.add_spacer(height=6)
                        _btn_w2 = (W - _sc(8)) // 2
                        with dpg.group(horizontal=True):
                            dpg.add_button(label="OK",     width=_btn_w2, height=THUMB_H,
                                           callback=_image_import_ok)
                            dpg.add_button(label="Cancel", width=_btn_w2, height=THUMB_H,
                                           callback=_image_import_cancel)

        dpg.add_separator()
        with dpg.group(horizontal=True):
            dpg.add_checkbox(label="Always on Top", callback=toggle_on_top,
                             default_value=app.always_on_top)
            dpg.add_text("  Theme:")
            dpg.add_combo(
                THEME_NAMES, tag="theme_combo",
                default_value=app.theme_name,
                width=_sc(82), callback=set_theme,
            )

    # ─────────────────────────────────────────────────────────────────────
    #  UPDATE LOOP
    # ─────────────────────────────────────────────────────────────────────
    def _draw_grad_slider(gname, smin, smax, ctx=None):
        tag_dl = _DL_TAGS[gname]
        if not dpg.does_item_exist(tag_dl):
            return
        if ctx is None:
            ctx = _grad_ctx()

        val = app.sl_vals.get(gname, smin)
        y0  = (THUMB_H - GRAD_H) // 2
        y1  = y0 + GRAD_H

        cache = _GRAD_ITEM_CACHE.get(gname)

        if cache is None:
            # ── First draw: create all items once ──────────────────────────
            # Segment rectangles (colours set immediately below)
            seg_ids = []
            for x0s, x1s in _GRAD_SEG_X:
                rid = dpg.draw_rectangle(
                    parent=tag_dl,
                    pmin=(x0s, y0), pmax=(x1s, y1),
                    fill=[128, 128, 128, 255], color=[0, 0, 0, 0],
                )
                seg_ids.append(rid)
            # Border line (never changes)
            dpg.draw_rectangle(
                parent=tag_dl,
                pmin=(0, y0), pmax=(GRAD_W, y1),
                fill=[0, 0, 0, 0], color=[70, 70, 70, 180],
            )
            # Value indicator line
            ind_col = [30, 30, 30, 240] if app.theme_name == "Light" else [255, 255, 255, 240]
            indicator_id = dpg.draw_line(
                parent=tag_dl,
                p1=(0, y0 - 2), p2=(0, y1 + 2),
                color=ind_col, thickness=2,
            )
            cache = {'seg_ids': seg_ids, 'indicator_id': indicator_id, 'y0': y0, 'y1': y1}
            _GRAD_ITEM_CACHE[gname] = cache

        # ── Update segment colours and indicator (no delete/recreate) ────
        seg_ids      = cache['seg_ids']
        indicator_id = cache['indicator_id']
        cy0          = cache['y0']
        cy1          = cache['y1']

        for i, rid in enumerate(seg_ids):
            t = (i + 0.5) / SEGS
            r, g, b = _grad_color(gname, t, ctx)
            dpg.configure_item(rid, fill=[r, g, b, 255])

        t_val   = max(0., min(1., (val - smin) / (smax - smin))) if smax > smin else 0.  # FIX: clamp
        tx      = int(t_val * (GRAD_W - 1))
        ind_col = [30, 30, 30, 240] if app.theme_name == "Light" else [255, 255, 255, 240]
        dpg.configure_item(indicator_id,
                           p1=(tx, cy0 - 2), p2=(tx, cy1 + 2), color=ind_col)

        dpg.set_value(f"slval_{gname}", f"{val:>4}")

    # _POINT defined once here, not inside update() to avoid recreation every frame.
    class _POINT(ctypes.Structure):
        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
    _pt = _POINT()   # Single global instance — no allocation per frame

    def update():
        # -- Frame-level mouse state ──────────────────────────────────────
        _lmb_down    = dpg.is_mouse_button_down(0)
        _lmb_clicked = dpg.is_mouse_button_clicked(0)
        _lmb_rel     = dpg.is_mouse_button_released(0)
        _rmb_clicked = dpg.is_mouse_button_clicked(1)

        # -- Viewport size guard ──────────────────────────────────────────
        # DPG/GLFW auto-scales the viewport on DPI change. Because the window
        # is resizable=False and content is fixed-size, force dimensions back.
        _vw = dpg.get_viewport_width()
        _vh = dpg.get_viewport_height()
        if _vw != app._cached_vp_w or _vh != app._cached_vp_h:
            app._cached_vp_w = _vw
            app._cached_vp_h = _vh
            if _vw != _FIXED_VP_W or _vh != _FIXED_VP_H:
                dpg.set_viewport_width(_FIXED_VP_W)
                dpg.set_viewport_height(_FIXED_VP_H)

        # -- Frame-delayed color picker sync ──────────────────────────
        # Count down the undo write-freeze counter
        if app._pal_undo_freeze > 0:
            app._pal_undo_freeze -= 1


        # Picker delete+recreate: done ONLY here, between render frames.
        # Never inside a DPG callback — that causes a C-level segfault.
        if app._pending_picker_color is not None:
            # Do not delete the picker while the user is dragging — is_item_active=True
            # means DPG is currently processing it.
            _wheel_active = (dpg.does_item_exist("picker_wheel")
                             and dpg.is_item_active("picker_wheel"))
            if not _wheel_active:
                _pc = app._pending_picker_color
                app._pending_picker_color = None
                ir, ig, ib = [int(round(c * 255)) for c in _pc[:3]]
                col_255 = [ir, ig, ib, 255]
                if dpg.does_item_exist("picker_wheel"):
                    dpg.delete_item("picker_wheel")
                dpg.add_color_picker(
                    parent="grp_wheel",
                    tag="picker_wheel", width=W,
                    no_inputs=True, no_label=True,
                    no_side_preview=True, no_small_preview=True,
                    picker_mode=dpg.mvColorPicker_wheel,
                    display_type=dpg.mvColorEdit_uint8,
                    display_hsv=True,
                    default_value=col_255,
                    callback=picker_cb,
                )
                # Second set_value on the next frame syncs the triangle
                app._pending_wheel_sync = col_255
        if app._pending_wheel_sync is not None:
            if dpg.does_item_exist("picker_wheel"):
                app._picker_processing = True
                try:
                    dpg.set_value("picker_wheel", app._pending_wheel_sync)
                finally:
                    app._picker_processing = False
            app._pending_wheel_sync = None

        # -- Eyedropper: apply colour set by background thread ────────────
        if app._pending_pipette_color is not None:
            v = app._pending_pipette_color
            app._pending_pipette_color = None
            _force_update_color_picker(v)
            _sync_all_modes()
            add_to_history(v)

        # -- Flush deferred save requested by background threads ─────────
        if app._pending_save:
            app._pending_save = False
            save_config()

        # -- Open image-import panel when background thread has loaded an image ─
        if app._pending_image_import is not None and not app._image_modal_open:
            pil = app._pending_image_import
            app._pending_image_import = None
            _open_image_import_panel(pil)

        vsig = color_sig(app.current_color)
        # Cache viewport position: fetch only while dragging or every 30 frames.
        # save_config() reads _last_viewport_pos to avoid calling
        # get_viewport_pos() after the DPG context is destroyed.
        if app._tb_dragging or _lmb_down:
            try:
                app._last_viewport_pos = list(dpg.get_viewport_pos())
                app._cached_vp_pos     = app._last_viewport_pos
            except Exception:
                pass
        else:
            app._vp_pos_frame_ctr += 1
            if app._vp_pos_frame_ctr >= 30:
                app._vp_pos_frame_ctr = 0
                try:
                    app._last_viewport_pos = list(dpg.get_viewport_pos())
                except Exception:
                    pass
        mode = app.harmony_mode
        fmt  = app.fmt_mode
        pip  = app.pipette_active
        if pip:
            with _pip_lock:
                pip_color = list(app.pip_color)
        else:
            pip_color = None
        _new_preview = pip_color if pip else list(vsig) + [255]
        if _new_preview != app._last_drawn_preview:
            dpg.set_value("preview", _new_preview)
            app._last_drawn_preview = _new_preview
        _pip_overlay_update()

        # Only update preview_hex when colour changed and user is not typing
        if not pip and vsig != app._last_preview_sig:
            if not dpg.is_item_active("preview_hex"):
                r, g, b = vsig
                dpg.set_value("preview_hex", f"#{r:02x}{g:02x}{b:02x}".upper())
                app._last_preview_sig = vsig

        # -- Titlebar icon buttons (minimize / close) ─────────────────────
        for _dl_tag, _is_close in (("dl_tb_min", False), ("dl_tb_close", True)):
            if not dpg.does_item_exist(_dl_tag):
                continue
            _hov = dpg.is_item_hovered(_dl_tag)
            _clk = _lmb_clicked and _hov
            if _clk:
                if _is_close:
                    dpg.stop_dearpygui()
                else:
                    _minimize_win()
            if _hov == app._tb_hov[_dl_tag]:
                continue
            app._tb_hov[_dl_tag] = _hov
            dpg.delete_item(_dl_tag, children_only=True)
            if _hov:
                _bg = [200, 40, 40, 200] if _is_close else [80, 80, 80, 120]
                dpg.draw_rectangle(
                    parent=_dl_tag,
                    pmin=(0, 0), pmax=(_ICON_W, _TB_H),
                    fill=_bg, color=[0, 0, 0, 0],
                )
            _col = [255, 255, 255, 255] if _hov else [190, 190, 190, 255]
            _cx3 = _ICON_W // 2
            _cy3 = _TB_H  // 2
            if _is_close:
                dpg.draw_line(parent=_dl_tag,
                    p1=(_cx3 - _IC, _cy3 - _IC), p2=(_cx3 + _IC, _cy3 + _IC),
                    color=_col, thickness=_THICK_ICON)
                dpg.draw_line(parent=_dl_tag,
                    p1=(_cx3 + _IC, _cy3 - _IC), p2=(_cx3 - _IC, _cy3 + _IC),
                    color=_col, thickness=_THICK_ICON)
            else:
                dpg.draw_line(parent=_dl_tag,
                    p1=(_cx3 - _IC, _cy3 + _sc(3)), p2=(_cx3 + _IC, _cy3 + _sc(3)),
                    color=_col, thickness=_THICK_ICON)

        # -- Titlebar drag ────────────────────────────────────────────────
        # GetCursorPos only when LMB is down to avoid a syscall on every idle frame.
        lmb_dn_tb  = _lmb_down
        lmb_clk_tb = _lmb_clicked
        mx_g = my_g = 0
        if lmb_dn_tb or lmb_clk_tb:
            ctypes.windll.user32.GetCursorPos(ctypes.byref(_pt))
            mx_g, my_g = _pt.x, _pt.y
        if lmb_clk_tb and dpg.does_item_exist("titlebar_cw"):
            vx, vy = dpg.get_viewport_pos()
            _rel_y = my_g - vy
            _rel_x = mx_g - vx
            _vw    = app._cached_vp_w or dpg.get_viewport_width()
            # Use coordinate check instead of is_item_hovered — ImGui WindowPadding
            # (~8 px) leaves a border area where hover does not register.
            _in_tb_zone = (0 <= _rel_y < _TB_H + _sc(10)
                           and 0 <= _rel_x < _vw - _ICON_W * 2)
            if (_in_tb_zone
                    and not dpg.is_item_hovered("dl_tb_min")
                    and not dpg.is_item_hovered("dl_tb_close")):
                app._tb_dragging = True
                app._tb_drag_start_vp = (mx_g - vx, my_g - vy)
        if not lmb_dn_tb:
            app._tb_dragging = False
        if app._tb_dragging and lmb_dn_tb:
            ox, oy = app._tb_drag_start_vp
            dpg.set_viewport_pos([mx_g - ox, my_g - oy])

        # -- Sliders ──────────────────────────────────────────────────────
        # Always reset wheel delta regardless of mode. Without this, delta
        # accumulates in wheel mode and fires a jump when switching to sliders.
        wheel_delta            = app._mouse_wheel_delta
        app._mouse_wheel_delta = 0
        if not app.use_wheel:
            smode    = app.slider_mode
            current  = SLIDERS_BY_MODE.get(smode, [])
            # Compute _grad_ctx() once before the loop, not once per slider.
            _frame_grad_ctx = _grad_ctx()
            _mp2 = dpg.get_mouse_pos()
            mx, my = (_mp2[0], _mp2[1]) if _mp2 and len(_mp2) >= 2 else (0, 0)

            for (lbl, gname, _lcol, smin, smax) in current:
                tag_dl = _DL_TAGS[gname]
                if not dpg.does_item_exist(tag_dl):
                    continue
                hovered = dpg.is_item_hovered(tag_dl)
                if _lmb_clicked and hovered:
                    app._dragging = gname
                if not _lmb_down and app._dragging == gname:
                    app._dragging = None
                if app._dragging == gname and _lmb_down:
                    rect_min = dpg.get_item_rect_min(tag_dl)
                    lx = mx - rect_min[0]
                    t  = max(0., min(1., lx / (GRAD_W - 1)))
                    app.sl_vals[gname] = int(round(smin + (smax - smin) * t))
                    _apply_from_mode(smode)
                    _update_selected_pal_color()
                # Mouse wheel: delta > 0 = up (increase), < 0 = down (decrease)
                if hovered and wheel_delta != 0 and not _lmb_down:
                    step = max(1, (smax - smin) // 50)
                    new_val = app.sl_vals.get(gname, smin) + (step if wheel_delta > 0 else -step)
                    app.sl_vals[gname] = max(smin, min(smax, new_val))
                    _apply_from_mode(smode)
                    _update_selected_pal_color()
                if (app._last_sl_vals.get(gname) != app.sl_vals.get(gname)
                        or app._last_sl_vals.get('_sig') != vsig):
                    _draw_grad_slider(gname, smin, smax, _frame_grad_ctx)
            app._last_sl_vals = dict(app.sl_vals)
            app._last_sl_vals['_sig'] = vsig

        # -- Harmony colours ──────────────────────────────────────────────
        if vsig != app._last_sig or mode != app._last_mode or fmt != app._last_fmt:
            rf, gf, bf = [clamp01(c) for c in app.current_color[:3]]
            # colorsys.rgb_to_hls returns (h, l, s).
            h_, l_, s_ = colorsys.rgb_to_hls(rf, gf, bf)

            updates = [("main", rf, gf, bf)]

            angles = {
                "Complementary":       [180],
                "Split Complementary": [150, 210],
                "Analogous":           [30, -30],
                "Triadic":             [120, 240],
                "Tetradic":            [90, 180, 270],
                "Rectangle":           [60, 180, 240],
            }.get(mode, [])

            if mode in ["Tints", "Shades", "Tones"]:
                if mode == "Tints":
                    for i in range(1, 7):
                        t = i / 7.0
                        nr, ng, nb = colorsys.hls_to_rgb(h_, l_ + (1 - l_) * t, s_)
                        updates.append((f"h{i}", nr, ng, nb))
                elif mode == "Shades":
                    for i in range(1, 7):
                        t = i / 7.0
                        nr, ng, nb = colorsys.hls_to_rgb(h_, l_ * (1 - t), s_)
                        updates.append((f"h{i}", nr, ng, nb))
                elif mode == "Tones":
                    for i in range(1, 7):
                        t = i / 7.0
                        nr, ng, nb = colorsys.hls_to_rgb(h_, l_, s_ * (1 - t))
                        updates.append((f"h{i}", nr, ng, nb))
            elif angles:
                for i, ang in enumerate(angles):
                    nh = (h_ + ang / 360.) % 1.
                    nr, ng, nb = colorsys.hls_to_rgb(nh, l_, s_)
                    updates.append((f"h{i+1}", nr, ng, nb))

            # Check whether the user is dragging (picker or slider active).
            # During drag: update only colour swatches (fast, no configure_item).
            # Text, WCAG values and layout are deferred until drag ends.
            _picker_dragging = (dpg.does_item_exist("picker_wheel")
                                and dpg.is_item_active("picker_wheel"))
            _slider_dragging = (app._dragging is not None)
            _is_dragging     = _picker_dragging or _slider_dragging

            # Colour-only change during drag: use lightweight update path.
            _mode_or_fmt_changed = (mode != app._last_mode or fmt != app._last_fmt)
            _do_full_update = not _is_dragging or _mode_or_fmt_changed
            if _is_dragging and not _mode_or_fmt_changed:
                app._harmony_text_dirty = True  # defer text/WCAG update

            app.harmony_rgb = {}
            is_contrast = (fmt == "Contrast")
            show_copy   = not is_contrast and mode not in ("Tints", "Shades", "Tones")

            for i, tag in enumerate(HARMONY_TAGS):
                keys = HARMONY_KEYS[tag]
                if i < len(updates):
                    rgb2 = tuple(to_int(c) for c in updates[i][1:])
                    app.harmony_rgb[tag] = rgb2
                    # Fast update: always — colour swatch reflects colour in real time
                    dpg.set_value(keys["rect"], list(rgb2) + [255])
                    if _do_full_update:
                        # Slow update: visibility, text, WCAG, layout
                        dpg.configure_item(keys["group"],    show=True)
                        dpg.configure_item(keys["val"],      show=not is_contrast)
                        dpg.configure_item(keys["ctr_grp"],  show=is_contrast)
                        dpg.configure_item(keys["ctr_bars"], show=is_contrast)
                        dpg.configure_item(keys["copy_btn"], show=show_copy)
                        if is_contrast:
                            r2, g2, b2 = rgb2
                            cw = contrast_ratio(r2, g2, b2, 255, 255, 255)
                            cb = contrast_ratio(r2, g2, b2, 0,   0,   0  )
                            dpg.set_value(keys["ctr_w"],
                                          f"W  {cw:.2f}:1  {wcag_level(cw)}")
                            dpg.configure_item(keys["ctr_w"], color=wcag_color(cw))
                            dpg.set_value(keys["ctr_b"],
                                          f"B  {cb:.2f}:1  {wcag_level(cb)}")
                            dpg.configure_item(keys["ctr_b"], color=wcag_color(cb))
                            if dpg.does_item_exist(keys["ctr_bar_w"]):
                                dpg.configure_item(keys["ctr_bar_w"], color=[r2, g2, b2, 255])
                            if dpg.does_item_exist(keys["ctr_bar_b"]):
                                dpg.configure_item(keys["ctr_bar_b"], color=[r2, g2, b2, 255])
                        else:
                            dpg.set_value(keys["val"], format_value(tag, fmt))
                else:
                    if _do_full_update:
                        dpg.configure_item(keys["group"], show=False)

            app._last_sig  = vsig
            app._last_mode = mode
            app._last_fmt  = fmt
            if _do_full_update:
                app._harmony_text_dirty = False

        # When drag ends, flush the deferred text/WCAG update
        elif app._harmony_text_dirty and not (
            (dpg.does_item_exist("picker_wheel") and dpg.is_item_active("picker_wheel"))
            or app._dragging is not None
        ):
            app._harmony_text_dirty = False
            refresh_harmony_values()

        # -- History ──────────────────────────────────────────────────────
        if app._hist_dirty:
            app._hist_dirty = False
            with _history_lock:
                hist_snap = list(app.history)
            n = len(hist_snap)
            for i, t in enumerate(HIST_TAGS):
                if i < n:
                    dpg.configure_item(t, show=True)
                    dpg.set_value(t, hist_snap[i])
                else:
                    dpg.configure_item(t, show=False)
            _update_history_selection_style()

        # -- Image import error message (auto-hides after 5 s) ────────────
        if dpg.does_item_exist("import_error_text"):
            _imp_msg, _imp_until = app._import_status
            if _imp_msg and time.perf_counter() < _imp_until:
                dpg.set_value("import_error_text", _imp_msg)
                dpg.configure_item("import_error_text", show=True)
            else:
                if _imp_msg:
                    app._import_status = ("", 0.0)
                dpg.configure_item("import_error_text", show=False)

        # History right-click removal
        if _rmb_clicked:
            removed = False
            hovered_idx = None
            with _history_lock:
                hist_len = len(app.history)
            for i in range(hist_len):
                if i < len(HIST_TAGS) and dpg.is_item_hovered(HIST_TAGS[i]):
                    hovered_idx = i
                    break
            if hovered_idx is not None:
                with _history_lock:
                    if hovered_idx < len(app.history):
                        del app.history[hovered_idx]
                        app._hist_dirty = True
                        app.selected_history_indices.discard(hovered_idx)
                        removed = True
            if removed:
                # Reset selection so indices do not point to wrong colours.
                if app.palette_select_mode:
                    _exit_palette_select_mode()
                save_config()

        if app.change_time > 0 and (time.perf_counter() - app.change_time) > 0.5:
            if not dpg.is_item_active("picker_wheel"):
                add_to_history(app.current_color)
                app.change_time = 0.
                # Save palette colour changes when the user stops dragging.
                if app._editing_pal and app._pal_edit_dirty:
                    save_config()

        _pal_swatch_right_click(_rmb_clicked)
        _pal_swatch_drag(_lmb_down, _lmb_clicked, _lmb_rel)

        # -- Palette edit panel: refresh when active slot colour changes ──────
        # Done once per frame to avoid hundreds of rebuilds per second while
        # dragging. Skipped during drag reorder — _pal_swatch_drag() controls
        # the panel state.
        if app._editing_pal and app._pal_selected_idx is not None:
            if vsig != app._last_edit_color_sig:
                app._last_edit_color_sig = vsig
                if not app._pal_drag_active:
                    _refresh_pal_edit_panel()

        # Palette rebuild: _pal_dirty is set by the main thread and background threads.
        if app._pal_dirty:
            app._pal_dirty = False
            _rebuild_pal_rows()

    # ─────────────────────────────────────────────────────────────────────
    #  STARTUP
    # ─────────────────────────────────────────────────────────────────────
    dpg.create_viewport(
        title='Color Tools',
        width=int(283*_DPI_SCALE),
        height=int(623*_DPI_SCALE),   # same total height as original; custom bar replaces native
        x_pos=_start_pos[0], y_pos=_start_pos[1],
        resizable=False,
        decorated=False,   # hide native Windows title bar; custom bar is used instead
        vsync=False,       # manual FPS limiter handles frame pacing
    )
    dpg.setup_dearpygui()
    dpg.bind_theme(_dark_theme)

    # ── Rainbow gradient bar ─────────────────────────────────────────────────
    # viewport_drawlist draws directly on the viewport surface, outside any
    # window, so it is always flush against the top edge with no padding.
    _RB_H    = _sc(3)
    _RB_SEGS = 64
    with dpg.viewport_drawlist(front=True, tag="dl_rainbow"):
        _sw = _VP_W / _RB_SEGS
        for _i in range(_RB_SEGS):
            _hue = _i / _RB_SEGS
            _rr, _gg, _bb = colorsys.hsv_to_rgb(_hue, 1.0, 1.0)
            _x0 = int(_i * _sw)
            _x1 = max(_x0 + 1, int((_i + 1) * _sw))
            dpg.draw_rectangle(
                parent="dl_rainbow",
                pmin=(_x0, 0), pmax=(_x1, _RB_H),
                fill=[int(_rr*255), int(_gg*255), int(_bb*255), 255],
                color=[0, 0, 0, 0],
            )

    dpg.show_viewport()
    dpg.set_primary_window("PrimaryWindow", True)
    # Restore always-on-top state from config before the first frame.
    if app.always_on_top:
        dpg.set_viewport_always_top(True)
    dpg.render_dearpygui_frame()

    # Record the desired viewport size after the first frame. If DPG or Windows
    # resizes it later (e.g. DPI change), restore this size every frame.
    _FIXED_VP_W = dpg.get_viewport_width()
    _FIXED_VP_H = dpg.get_viewport_height()

    # Set titlebar spacer once — resizable=False so viewport width never changes.
    if dpg.does_item_exist("tb_spacer"):
        _vw     = dpg.get_viewport_width()
        _logo_w = _logo_dark_w or 0
        _sp_w   = max(2, _vw - _logo_w - 2 * _ICON_W - 3 * _sc(2) - 17)
        dpg.configure_item("tb_spacer", width=_sp_w)


    set_theme(None, app.theme_name)

    _sync_all_modes()
    _rebuild_pal_rows()

    # Apply Win32 DWM dark titlebar style. Wait until the HWND is available.
    for _ in range(20):
        _hwnd = _get_own_hwnd()
        if _hwnd:
            try:
                _dark = ctypes.c_int(1)
                ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    _hwnd, 20, ctypes.byref(_dark), ctypes.sizeof(_dark)
                )
            except Exception:
                pass
            break
        time.sleep(0.05)

    # Set Windows timer resolution to 1 ms (default ~15.6 ms) for smooth
    # FPS throttling and accurate sleep.
    try:
        ctypes.windll.winmm.timeBeginPeriod(1)
        _timer_period_set = True
    except Exception:
        _timer_period_set = False

    _TARGET_FPS   = 30
    _FRAME_BUDGET = 1.0 / _TARGET_FPS

    while dpg.is_dearpygui_running():
        _frame_start = time.perf_counter()
        update()
        dpg.render_dearpygui_frame()
        _elapsed = time.perf_counter() - _frame_start
        _sleep   = _FRAME_BUDGET - _elapsed
        if _sleep > 0.001:
            _pumping_sleep(_sleep)

    save_config()

    # ── Clean shutdown ────────────────────────────────────────────
    # 1. Clear the pipette flag so any active pipette thread exits its loop.
    app.pipette_active = False
    # If the pipette thread is blocked in GetMessageW, send it WM_QUIT
    _pip_tid_shutdown = getattr(app, '_pipette_thread_tid', 0)
    if _pip_tid_shutdown:
        try:
            ctypes.windll.user32.PostThreadMessageW(_pip_tid_shutdown, 0x0012, 0, 0)
        except Exception:
            pass

    # Close the Win32 eyedropper overlay
    if _pip_hwnd[0]:
        try:
            _pip_u32.DestroyWindow(_pip_hwnd[0])
            _pip_u32.UnregisterClassW(_pip_cls[0], _pip_hInst)
        except Exception:
            pass
        _pip_hwnd[0] = None
    # Free eyedropper GDI resources
    try:
        _pip_free_gdi()
    except Exception:
        pass
    # Keep app._pip_wndproc_cb alive — the OS may deliver messages after
    # DestroyWindow. Setting it to None would cause a segfault.

    # 2. Clear pending picker state — no more DPG calls beyond this point.

    # 3. Restore Windows timer resolution.
    if _timer_period_set:
        try:
            ctypes.windll.winmm.timeEndPeriod(1)
        except Exception:
            pass

    dpg.destroy_context()

except Exception:
    _fatal(traceback.format_exc())