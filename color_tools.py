import os
import sys
import traceback
import ctypes
import colorsys
import threading
import time
import json
from functools import lru_cache
from collections import deque

# ── DPI awareness (must run before everything else) ──────────────────────────────────
# Level 1 (system-aware): the process declares DPI awareness so Windows does
# not automatically scale the window. DPI is read at startup and all UI
# measurements are scaled accordingly -- the window stays the correct size on
# both monitors and content does not shrink when moving between screens.
_DPI_SCALE = 1.0
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
    _hdc = ctypes.windll.user32.GetDC(0)
    _dpi = ctypes.windll.gdi32.GetDeviceCaps(_hdc, 88)  # LOGPIXELSX
    ctypes.windll.user32.ReleaseDC(0, _hdc)
    _DPI_SCALE = _dpi / 96.0
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

# ── Config directory (must be defined before anything else) ────────────────────────────────────
_APPDATA    = os.environ.get('APPDATA', os.path.expanduser('~'))
CFG_DIR     = os.path.join(_APPDATA, 'Color Tools')
CFG_FILE    = os.path.join(CFG_DIR, 'config.json')
_LOG        = os.path.join(CFG_DIR, 'error.log')
if getattr(sys, 'frozen', False):
    _SCRIPT_DIR = sys._MEIPASS
else:
    _SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


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
    import mss as _mss_mod
    from pynput import mouse as _mouse_mod
    import pyperclip
except Exception:
    _fatal(
        "Missing dependency. Install with:\n"
        "  pip install mss pynput pyperclip\n\n"
        + traceback.format_exc()
    )

try:
    import dearpygui.dearpygui as dpg
except Exception:
    _fatal("Could not import dearpygui:\n\n" + traceback.format_exc())

# PIL is optional — used only for image palette import.
# If missing the button still works but shows a clear error message.
try:
    from PIL import Image as _PIL_Image
    _PIL_AVAILABLE = True
except Exception:
    _PIL_AVAILABLE = False
    _write_log("PIL (Pillow) not found — image palette import will be disabled.\n"
               "Install with:  pip install Pillow")

_write_log("Import OK, starting up...")

try:
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
            with open(CFG_FILE, 'w', encoding='utf-8') as f:
                json.dump({
                    'history':   history_snapshot,
                    'palettes':  palettes_snapshot,
                    'pal_order': pal_order_snapshot,
                    'win_pos':    list(dpg.get_viewport_pos()),
                    'theme_name': app.theme_name,
                }, f)
        except Exception:
            pass

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

    @lru_cache(maxsize=4096)
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
    @lru_cache(maxsize=256)
    def _lin_rgb(c):
        c = max(0., min(1., c))
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4

    def _lin(c):
        return _lin_rgb(c / 255.)

    @lru_cache(maxsize=1024)
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

    def _xyz_f(t):
        # Guard against negative values (would produce complex numbers in Python).
        if t < 0:
            return 0.0
        return t ** (1 / 3) if t > 0.008856 else 7.787 * t + 16 / 116

    def _delin(c):
        c = max(0, min(1, c))
        return 12.92 * c if c <= 0.0031308 else 1.055 * c ** (1 / 2.4) - 0.055

    def rgb_to_lab(r, g, b):
        rl = _lin_rgb(r / 255); gl = _lin_rgb(g / 255); bl = _lin_rgb(b / 255)
        X = rl * .4124 + gl * .3576 + bl * .1805
        Y = rl * .2126 + gl * .7152 + bl * .0722
        Z = rl * .0193 + gl * .1192 + bl * .9505
        X /= .95047; Z /= 1.08883
        return (
            116 * _xyz_f(Y) - 16,
            500 * (_xyz_f(X) - _xyz_f(Y)),
            200 * (_xyz_f(Y) - _xyz_f(Z)),
        )

    def lab_to_rgb(L, a, b_):
        fy = (L + 16) / 116
        fx = a / 500 + fy
        fz = fy - b_ / 200

        def fi(t):
            # Compute t**3 only once.
            t3 = t ** 3
            return t3 if t3 > 0.008856 else (t - 16 / 116) / 7.787

        X  = fi(fx) * .95047; Y = fi(fy); Z = fi(fz) * 1.08883
        rl =  X * 3.2406 - Y * 1.5372 - Z * .4986
        gl = -X * .9689  + Y * 1.8758  + Z * .0415
        bl =  X * .0557  - Y * .2040   + Z * 1.0570
        return (
            max(0, min(255, int(round(_delin(rl) * 255)))),
            max(0, min(255, int(round(_delin(gl) * 255)))),
            max(0, min(255, int(round(_delin(bl) * 255)))),
        )

    def rgb_to_gray(r, g, b):
        return int(round(.2126 * r + .7152 * g + .0722 * b))

    # -- Formaatit --------------------------------------------------------
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
        fmt         = dpg.get_value("fmt_combo")
        mode        = dpg.get_value("harmony_combo")
        is_contrast = (fmt == "Contrast")
        for tag in HARMONY_TAGS:
            if not dpg.does_item_exist(f"val_{tag}"):
                continue
            if not dpg.is_item_shown(f"group_{tag}"):
                continue
            dpg.configure_item(f"val_{tag}",      show=not is_contrast)
            dpg.configure_item(f"ctr_grp_{tag}",  show=is_contrast)
            dpg.configure_item(f"ctr_bars_{tag}", show=is_contrast)
            dpg.configure_item(f"copy_btn_{tag}", show=not is_contrast and mode not in ("Tints", "Shades", "Tones"))
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

    # -- Layout-vakiot ----------------------------------------------------
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

    _ESW   = _sc(28)   # edit swatch size
    _ECOLS = 8          # swatches per row in edit mode

    _history_lock  = threading.Lock()
    _palettes_lock = threading.Lock()
    _pip_lock      = threading.Lock()
    _wheel_lock    = threading.Lock()

    # -- Sovellustila -----------------------------------------------------
    class AppState:
        def __init__(self):
            self.pipette_active  = False
            self.pip_color       = [0, 0, 0, 255]
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
            self.palette_select_mode          = False
            self.selected_history_indices     = set()
            self._pending_pal_rebuild = False
            self._import_error        = ""
            self._import_error_until  = 0.0
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
            self._pal_drag_undo_done  = False  # True when undo was already pushed on this click (drag)
            self._pal_pre_edit_snapshot = None # Palette snapshot taken at swatch selection, not yet on stack
            self._pal_edit_dirty      = False  # True once color actually changed since last swatch click
            self._pal_undo_freeze     = 0       # Frames left to block _update_selected_pal_color after undo
            self._last_picker_force_time = 0.0
            self._pending_wheel_sync     = None # Frame-delayed sync for DPG wheel bug
            self._pal_undo_stack         = deque(maxlen=30)  # Snapshots for undo
            self._mouse_wheel_delta      = 0    # Mouse wheel scroll delta for sliders
            self._pending_pipette_color  = None # Set by pipette thread, consumed by update()
            self._picker_processing      = False # Guard against recursive picker callbacks
            self._pending_save           = False # Set by background threads, flushed in update()
            self._last_edit_color_sig    = None  # Tracks color changes for palette edit panel refresh
            self._last_preview_sig       = None  # Cache: avoid redundant preview_hex set_value calls


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

    # -- Apufunktiot ------------------------------------------------------
    def clamp01(v):
        return max(0., min(1., v))

    # round() preserves colour accuracy.
    def to_int(v):
        if v > 1.0001: 
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
    def _grad_ctx():
        """Compute base colour values ONCE per slider redraw — passed to _grad_color.
        Avoids repeating rgb_to_hls / rgb_to_cmyk / rgb_to_lab for every segment."""
        r, g, b    = color_sig(app.current_color)
        rf, gf, bf = r / 255., g / 255., b / 255.
        # colorsys.rgb_to_hls returns (h, l, s) -- not (h, s, l).
        h_, l_, s_ = colorsys.rgb_to_hls(rf, gf, bf)
        c0, m0, y0, k0 = rgb_to_cmyk(r, g, b)
        L0, a0, b0_    = rgb_to_lab(r, g, b)
        return rf, gf, bf, h_, l_, s_, c0, m0, y0, k0, L0, a0, b0_

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
    # Pre-computed DPG tag names for slider drawlists — avoids f-string creation every frame.
    _DL_TAGS = {gname: f"dl_{gname}"
                for sliders in SLIDERS_BY_MODE.values()
                for _, gname, *_ in sliders}

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
        app.change_time   = time.time()
        
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

    def set_window_style():
        """Retries until FindWindowW succeeds (window may not be registered yet)."""
        for _ in range(10):
            try:
                hwnd = ctypes.windll.user32.FindWindowW(None, "Color Tools")
                if hwnd:
                    dark = ctypes.c_int(1)
                    ctypes.windll.dwmapi.DwmSetWindowAttribute(
                        hwnd, 20, ctypes.byref(dark), ctypes.sizeof(dark)
                    )
                    return
            except Exception:
                pass
            time.sleep(0.1)

    def _update_selected_pal_color():
        # Estetään kirjoitus muutaman framen ajan undoamisen jälkeen, jottei
        # DPG:n jonossa olevat vanhat picker-callbackit turmele palautettua tilaa.
        if getattr(app, '_pal_undo_freeze', 0) > 0:
            return
        if app._editing_pal and getattr(app, '_pal_selected_idx', None) is not None:
            with _palettes_lock:
                if app._editing_pal in app.palettes:
                    colors = app.palettes[app._editing_pal]
                    if app._pal_selected_idx < len(colors):
                        r, g, b = color_sig(app.current_color)
                        # Lazy undo push: pushataan snapshot pinoon vasta kun väri
                        # OIKEASTI muuttuu ensimmäistä kertaa tämän swatch-klikkauksen jälkeen.
                        # Näin pelkkä swatchin valinta ei täytä pinoa identtisillä snapshotilla.
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
            # Don't refresh the panel on every picker movement — the palette signature
            # change will be detected in the main update() loop, avoiding constant redraws.
            # This prevents freezing when dragging the color wheel while editing a palette.

    def _force_update_color_picker(col_norm):
        app._last_picker_force_time = time.time()
        # Convert to 0-255 as uint8 is often more robust for wheel sync in DPG
        ir, ig, ib = [int(round(c * 255)) for c in col_norm[:3]]
        col_255 = [ir, ig, ib, 255]
        
        app.current_color = list(col_norm)
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
        # Schedule a second set_value in the next frame to force triangle sync
        app._pending_wheel_sync = col_255


    def picker_cb(s, v):
        # ── CRITICAL GUARDS ──
        # (1) Ignore for 500ms after a forced update to swallow DPG's internal sync noise.
        if time.time() - getattr(app, '_last_picker_force_time', 0.0) < 0.5:
            return
        # (2) Prevent recursive callbacks from _apply_from_mode calling dpg.set_value()
        if getattr(app, '_picker_processing', False):
            return
            
        # Mark that we're processing to block recursive calls
        app._picker_processing = True
        try:
            # Normalize if we get 0-255 ints from uint8 mode
            if any(x > 1.01 for x in v):
                v = [x / 255.0 for x in v]
            app.current_color = list(v)
            app.change_time   = time.time()
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
        # Accept 3-char shorthand (e.g. "F0A" -> "FF00AA")
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
        dpg.set_viewport_always_top(v)

    # -- Paletti-apufunktiot ----------------------------------------------
    def _auto_pal_name():
        i = 1
        while f"Palette {i}" in app.palettes:
            i += 1
        return f"Palette {i}"

    def _register_palette(name):
        if name not in app._pal_order:
            app._pal_order.append(name)

    def palette_names():
        order = [k for k in reversed(app._pal_order) if k in app.palettes]
        for k in sorted(app.palettes.keys()):
            if k not in order:
                order.append(k)
        return order

    _PCOLS = 9

    def _rebuild_pal_rows():
        """Must ALWAYS be called from the main thread -- never from a background thread."""
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
                            dpg.set_value("picker_wheel", u)

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
        if not name or name not in app.palettes:
            return
        # Syvä kopio: jokainen värilista kopioidaan erikseen jotta myöhemmät
        # paikallaan-mutaatiot (pop, suora alkion muokkaus) eivät korruptoi snapshotia.
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
        app._pal_drag_undo_done = False
        # Kirjoituslukko: estää vanhat DPG-callbackit ylikirjoittamasta palautettua tilaa.
        app._pal_undo_freeze = 4
        with _palettes_lock:
            app.palettes[name] = colors
        # KRIITTINEN: aseta pre-edit snapshot HETI palautettuun tilaan.
        # Jos snapshot jätetään None:ksi, lazy push ei voi tallentaa seuraavaa
        # muutosta pinoon — käyttäjä menettää undo-toiminnon palautuksen jälkeen.
        app._pal_pre_edit_snapshot = [list(c) for c in colors]
        # Päivitä pikeri palautettuun väriin.
        if (name == app._editing_pal
                and idx is not None
                and idx < len(colors)):
            c = colors[idx]
            _force_update_color_picker([c[0] / 255., c[1] / 255., c[2] / 255., 1.0])
            _sync_all_modes()
        # Aseta valinta takaisin vasta kaiken muun jälkeen.
        app._pal_selected_idx = idx if (idx is not None and idx < len(colors)) else None
        save_config()
        _rebuild_pal_rows()
        _refresh_pal_edit_panel()
        if dpg.does_item_exist("pal_undo_btn"):
            dpg.configure_item("pal_undo_btn", enabled=len(app._pal_undo_stack) > 0)

    def _refresh_pal_edit_panel(drag_src=None, drag_insert=None):
        """Rakentaa paletin edit-panelin uudelleen.

        Normal mode: draw all colours.
        Drag preview: remove the dragged colour from the list and insert an
        empty gap slot (yellow border) at drag_insert.
        Slot positions are read live every frame in _pal_swatch_drag() via
        dpg.get_item_rect_min(), so no separate rect cache is maintained.
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
        colors = app.palettes[name]

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
            # Deselect so the picker is no longer locked to any slot —
            # the user can now freely pick a new color to add next.
            app._pal_selected_idx = None
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
            ins = max(0, min(len(stripped), drag_insert))
            preview = stripped[:ins] + [None] + stripped[ins:]
        else:
            preview = list(colors)

        if preview:
            for row_start in range(0, len(preview), _ECOLS):
                grp = dpg.add_group(parent="pal_edit_panel", horizontal=True, tag=dpg.generate_uuid())
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
                        is_sel = (getattr(app, '_pal_selected_idx', None) == slot and not dragging)
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
        if not name or name not in app.palettes:
            return
        colors = app.palettes[name]
        h = (
            f"<html><head><meta charset='utf-8'>HTMLTITLE_PLACEHOLDER{_HCSS}</head><body>"
            f"<h2>HTMLH2_PLACEHOLDER</h2>"
            f"<table>{_html_table_header()}"
        )
        for v in colors:
            h += _html_color_row(v[0], v[1], v[2])
        h += "</table></body></html>"
        _save_html_dialog(h, name)

    def _pal_swatch_right_click():
        """Called every frame from the update() loop."""
        if not dpg.is_mouse_button_clicked(1):
            return
        if app._editing_pal and app._editing_pal in app.palettes:
            with _palettes_lock:
                colors = app.palettes.get(app._editing_pal, [])
                ncolors = len(colors)
            for idx in range(ncolors):
                stag = f"pedit_sw_{idx}"
                if dpg.does_item_exist(stag) and dpg.is_item_hovered(stag):
                    _push_pal_undo(app._editing_pal)
                    with _palettes_lock:
                        colors.pop(idx)
                        if not colors:
                            del app.palettes[app._editing_pal]
                            if app._editing_pal in app._pal_order:
                                app._pal_order.remove(app._editing_pal)
                            app._editing_pal = None
                    save_config()
                    _rebuild_pal_rows()
                    _refresh_pal_edit_panel()
                    return

    def _pal_swatch_drag():
        """Drag-and-drop -- kutsutaan joka framessa.

        Rects are read at the start of every frame (after the previous render pass),
        so they are always up to date regardless of when the panel was last rebuilt.
        """
        if not app._editing_pal or app._editing_pal not in app.palettes:
            return

        colors = app.palettes[app._editing_pal]
        n      = len(colors)

        lmb_dn  = dpg.is_mouse_button_down(0)
        lmb_clk = dpg.is_mouse_button_clicked(0)
        lmb_rel = dpg.is_mouse_button_released(0)
        mx, my  = dpg.get_mouse_pos(local=False)

        # ── Read slot positions every frame ─────────────────────────────────────────
        # In preview mode there are n slots (n-1 colours + 1 gap).
        # In normal mode there are n slots. Always read all existing slots.
        if app._pal_drag_active:
            live_rects = {}
            for slot in range(n + 1):   # max n slots in preview
                t = f"pedit_sw_{slot}"
                if dpg.does_item_exist(t):
                    try:
                        rm = dpg.get_item_rect_min(t)
                        if rm is not None:
                            live_rects[slot] = (rm[0], rm[1])
                    except Exception:
                        pass
        else:
            live_rects = {}

        def _nearest_slot(px, py):
            if not live_rects:
                return None
            best_i, best_d = None, float('inf')
            for i, (rx, ry) in live_rects.items():
                cx = rx + _ESW / 2
                cy = ry + _ESW / 2
                d  = (px - cx) ** 2 + (py - cy) ** 2
                if d < best_d:
                    best_d = d
                    best_i = i
            return best_i

        # ── Start drag ──────────────────────────────────────────────────────────────────────
        if lmb_clk:
            clicked_swatch = False
            for i in range(n):
                tag = f"pedit_sw_{i}"
                if dpg.does_item_exist(tag) and dpg.is_item_hovered(tag):
                    clicked_swatch = True
                    app._pal_drag_idx    = i
                    app._pal_drag_insert = i
                    app._pal_drag_active = False
                    app._pal_drag_start  = (mx, my)

                    # Deselektoi jos klikataan jo valittua swatchia
                    if app._pal_selected_idx == i:
                        app._pal_selected_idx = None
                        app._pal_pre_edit_snapshot = None
                        app._pal_edit_dirty = False
                        app._pal_drag_undo_done = False
                        _refresh_pal_edit_panel()
                        break
                    # Ota snapshot tällä hetkellä (ennen mahdollista värinmuutosta),
                    # mutta ÄLÄ pushä vielä — lazy push tapahtuu _update_selected_pal_color:ssa
                    # vasta kun väri oikeasti muuttuu. Näin pelkkä swatchien selaus ei
                    # täytä undo-pinoa identtisillä snapshotilla.
                    if app._editing_pal in app.palettes:
                        app._pal_pre_edit_snapshot = [list(c) for c in app.palettes[app._editing_pal]]
                    app._pal_edit_dirty = False
                    app._pal_drag_undo_done = False
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
                # Drag on järjestyksenvaihto — pushätaan aina nykyinen tila pinoon.
                # Käytetään pre-edit snapshotia jos väriä ei ole vielä muutettu,
                # muuten tallennetaan sen hetkinen tila.
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
                app._pal_drag_undo_done = False
                _refresh_pal_edit_panel(
                    drag_src    = app._pal_drag_idx,
                    drag_insert = app._pal_drag_insert,
                )

            if app._pal_drag_active:
                hit = _nearest_slot(mx, my)
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
                # Redundant picker set removed (was causing black jump due to float mismatch)
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
                    save_config()
                _rebuild_pal_rows()
                _refresh_pal_edit_panel()

    # -- Historia-valinta -------------------------------------------------
    # Placeholder -- varsinainen teema luodaan dpg.create_context():n jalkeen
    _hist_selected_theme = None

    def _update_history_selection_style():
        for i, tag in enumerate(HIST_TAGS):
            if not dpg.does_item_exist(tag):
                continue
            # Use the global theme -- do not create a new object every update.
            dpg.bind_item_theme(tag, _hist_selected_theme if i in app.selected_history_indices else None)

    def save_selected_as_palette():
        if not app.selected_history_indices:
            return
        with _history_lock:
            colors = [
                app.history[i]
                for i in sorted(app.selected_history_indices)
                if i < len(app.history)
            ]
        if not colors:
            return
        name = _auto_pal_name()
        with _palettes_lock:
            app.palettes[name] = colors
        _register_palette(name)
        save_config()
        _rebuild_pal_rows()
        _exit_palette_select_mode()

    def _exit_palette_select_mode():
        """Poistu paletin luontitilasta ilman tallennusta."""
        app.palette_select_mode = False
        app.selected_history_indices.clear()
        dpg.set_item_label("btn_new_palette", "New Palette")
        dpg.bind_item_theme("btn_new_palette", None)
        if dpg.does_item_exist("btn_cancel_palette"):
            dpg.configure_item("btn_cancel_palette", show=False)
        _update_history_selection_style()

    def cancel_palette_select_mode():
        _exit_palette_select_mode()

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
            dpg.set_item_label("btn_new_palette", "Save")
            dpg.bind_item_theme("btn_new_palette", _new_palette_select_theme)
            if dpg.does_item_exist("btn_cancel_palette"):
                dpg.configure_item("btn_cancel_palette", show=True)
            _update_history_selection_style()

    def history_click_cb(sender, app_data, user_data):
        i = user_data
        with _history_lock:
            if i >= len(app.history):
                return
            if app.palette_select_mode:
                if i in app.selected_history_indices:
                    app.selected_history_indices.remove(i)
                else:
                    app.selected_history_indices.add(i)
                _update_history_selection_style()
            else:
                col      = app.history[i]
                col_norm = [col[0] / 255., col[1] / 255., col[2] / 255., 1.]
                _force_update_color_picker(col_norm)
                # manual redundant picker_cb call removed.

    # -- Harmonia -> paletti ----------------------------------------------
    def save_harmony_as_palette():
        colors = []
        for tag in ["main", "h1", "h2", "h3", "h4", "h5", "h6", "h7"]:
            if app.harmony_rgb.get(tag) is not None:
                r, g, b = app.harmony_rgb[tag]
                colors.append([r, g, b, 255])
        if not colors:
            return
        name = _auto_pal_name()
        with _palettes_lock:
            app.palettes[name] = colors
        _register_palette(name)
        save_config()
        _rebuild_pal_rows()

    # -- Kuvan tuonti -----------------------------------------------------
    def _commit_image_palette(colors):
        """Thread-safe -- does not touch any DearPyGui elements."""
        with _history_lock:
            for c in reversed(colors):
                if c not in app.history:
                    app.history.appendleft(c)
        with _palettes_lock:
            # Generate name inside the lock to avoid a race condition where two
            # background threads could both get the same name before either writes.
            name = _auto_pal_name()
            app.palettes[name] = list(colors)
        _register_palette(name)
        # save_config() calls dpg.get_viewport_pos() which is not thread-safe.
        # Signal the main thread to save on the next frame instead.
        app._pending_save = True
        app._pending_pal_rebuild = True

    def import_image_palette():
        if not _PIL_AVAILABLE:
            app._import_error       = "Pillow not installed — run: pip install Pillow"
            app._import_error_until = time.time() + 6.0
            return

        def _run():
            try:
                import tkinter as tk
                from tkinter import filedialog, simpledialog
                root = tk.Tk()
                root.withdraw()
                root.attributes('-topmost', True)
                path = filedialog.askopenfilename(
                    title="Open Image",
                    filetypes=[
                        ("Image files", "*.png *.jpg *.jpeg *.bmp *.gif *.webp *.tiff"),
                        ("All files", "*.*"),
                    ],
                )
                if not path:
                    root.destroy()
                    return
                try:
                    pil_img = _PIL_Image.open(path).convert("RGB")
                    pil_img.thumbnail((200, 200))
                except Exception:
                    root.destroy()
                    return

                dlg = tk.Toplevel(root)
                dlg.title("Select Colors")
                dlg.geometry("350x230")
                dlg.attributes('-topmost', True)
                dlg.grab_set()

                # Theme matching for Tkinter
                themes = {
                    "Dark":      {"bg": "#1a1a1a", "fg": "#dcdcdc", "btn": "#3c3c3c", "btnhov": "#505050", "hl": "#4296fa"},
                    "Light":     {"bg": "#ebebeb", "fg": "#1e1e1e", "btn": "#c8c8c8", "btnhov": "#b4b4b4", "hl": "#4296fa"},
                    "Midnight":  {"bg": "#0D1117", "fg": "#c9d1d9", "btn": "#1f3654", "btnhov": "#305482", "hl": "#58a6ff"},
                    "Mocha":     {"bg": "#1e1a17", "fg": "#e1d2c3", "btn": "#5a412d", "btnhov": "#78583c", "hl": "#a67c52"},
                    "Nord":      {"bg": "#2e3440", "fg": "#d8dee9", "btn": "#434c5e", "btnhov": "#4c566a", "hl": "#88c0d0"},
                    "Solarized": {"bg": "#002b36", "fg": "#fdf6e3", "btn": "#004858", "btnhov": "#005f73", "hl": "#2aa198"},
                }
                th = themes.get(app.theme_name, themes["Dark"])
                dlg.configure(bg=th["bg"])
                
                # Attempt to set dark titlebar on Windows 10/11 for dark themes
                try:
                    hwnd = ctypes.windll.user32.GetParent(dlg.winfo_id())
                    if app.theme_name != "Light":
                        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(ctypes.c_int(1)), 4)
                except Exception:
                    pass

                dlg_accepted = False
                selected_colors = []

                def update_preview(*args):
                    n_val = int(scale.get())
                    quantized = pil_img.quantize(colors=n_val, method=_PIL_Image.Quantize.FASTOCTREE)
                    palette_data = quantized.getpalette()[:n_val * 3]
                    colors = [
                        [palette_data[i * 3], palette_data[i * 3 + 1], palette_data[i * 3 + 2], 255]
                        for i in range(n_val)
                    ]
                    colors.sort(key=lambda c: 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2])
                    
                    for w in preview_frame.winfo_children():
                        w.destroy()
                    
                    for c in colors:
                        hx = f"#{c[0]:02x}{c[1]:02x}{c[2]:02x}"
                        lbl = tk.Label(preview_frame, bg=hx)
                        lbl.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

                    nonlocal selected_colors
                    selected_colors = colors

                lbl = tk.Label(dlg, text="How many colors? (2-20)", font=("Segoe UI", 10), bg=th["bg"], fg=th["fg"])
                lbl.pack(pady=10)
                
                scale = tk.Scale(
                    dlg, from_=2, to=20, orient=tk.HORIZONTAL, command=update_preview,
                    bg=th["bg"], fg=th["fg"], highlightthickness=0,
                    activebackground=th["hl"], troughcolor=th["btn"]
                )
                scale.set(8)
                scale.pack(fill=tk.X, padx=20)
                
                preview_frame = tk.Frame(dlg, height=50, bg=th["bg"])
                preview_frame.pack(fill=tk.X, padx=20, pady=15)
                preview_frame.pack_propagate(False)
                
                def on_ok():
                    nonlocal dlg_accepted
                    dlg_accepted = True
                    dlg.destroy()

                btn_frame = tk.Frame(dlg, bg=th["bg"])
                btn_frame.pack(pady=5)
                
                for txt, cmd in [("OK", on_ok), ("Cancel", dlg.destroy)]:
                    tk.Button(
                        btn_frame, text=txt, width=10, command=cmd, cursor="hand2",
                        bg=th["btn"], fg=th["fg"], activebackground=th["btnhov"],
                        activeforeground=th["fg"], relief=tk.FLAT, bd=0, padx=10, pady=4
                    ).pack(side=tk.LEFT, padx=5)
                
                dlg.update_idletasks()
                x = root.winfo_screenwidth() // 2 - 175
                y = root.winfo_screenheight() // 2 - 115
                dlg.geometry(f"+{x}+{y}")
                
                root.wait_window(dlg)
                root.destroy()
                
                if dlg_accepted and selected_colors:
                    _commit_image_palette(selected_colors)
            except Exception:
                _write_log("import_image error:\n" + traceback.format_exc())
                app._import_error       = "Image import failed — check error.log"
                app._import_error_until = time.time() + 5.0
                app._pending_pal_rebuild = True

        threading.Thread(target=_run, daemon=True).start()

    def import_ase_palette():
        """Import colors from an ASE (Adobe Swatch Exchange) file."""
        def _run():
            try:
                import struct
                import tkinter as tk
                from tkinter import filedialog
                root = tk.Tk()
                root.withdraw()
                root.attributes('-topmost', True)
                path = filedialog.askopenfilename(
                    title="Import ASE Palette",
                    filetypes=[
                        ("ASE files", "*.ase"),
                        ("All files", "*.*"),
                    ],
                )
                root.destroy()
                if not path:
                    return
                with open(path, 'rb') as f:
                    data = f.read()
                # Validate header
                if data[:4] != b'ASEF':
                    app._import_error       = "Not a valid ASE file"
                    app._import_error_until = time.time() + 5.0
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
                    app._import_error       = "No RGB colors found in ASE file"
                    app._import_error_until = time.time() + 5.0
                    return
                _commit_image_palette(colors)
            except Exception:
                _write_log("import_ase error:\n" + traceback.format_exc())
                app._import_error       = "ASE import failed — check error.log"
                app._import_error_until = time.time() + 5.0
                app._pending_pal_rebuild = True
        threading.Thread(target=_run, daemon=True).start()

    # -- Pipette ----------------------------------------------------------
    def _pipette_overlay():
        """Win32-pohjainen overlay — ei Tkinteriä, toimii taustasäikeessä."""
        import ctypes.wintypes as _wt

        _u32 = ctypes.windll.user32
        _g32 = ctypes.windll.gdi32

        WS_POPUP       = 0x80000000
        WS_EX_TOPMOST  = 0x00000008
        WS_EX_TOOLWINDOW = 0x00000080
        WS_EX_LAYERED  = 0x00080000
        WS_EX_NOACTIVATE = 0x08000000
        WM_PAINT       = 0x000F
        WM_DESTROY     = 0x0002

        CLS_NAME = "PipetteOverlay"
        OW, OH   = 120, 28   # overlay koko pikseleinä

        def _wndproc(hwnd, msg, wp, lp):
            if msg == WM_PAINT:
                ps = ctypes.create_string_buffer(64)
                hdc = _u32.BeginPaint(hwnd, ps)
                with _pip_lock:
                    r2, g2, b2 = app.pip_color[0], app.pip_color[1], app.pip_color[2]
                # Tausta
                hbr = _g32.CreateSolidBrush(_rgb(20, 20, 20))
                rc  = _wt.RECT(0, 0, OW, OH)
                _u32.FillRect(hdc, ctypes.byref(rc), hbr)
                _g32.DeleteObject(hbr)
                # Värilaatikko
                hbr2 = _g32.CreateSolidBrush(_rgb(r2, g2, b2))
                rc2  = _wt.RECT(4, 4, 24, OH - 4)
                _u32.FillRect(hdc, ctypes.byref(rc2), hbr2)
                _g32.DeleteObject(hbr2)
                # Hex-teksti valkoisella
                hex_str = f"#{r2:02X}{g2:02X}{b2:02X}"
                _g32.SetBkMode(hdc, 1)   # TRANSPARENT
                _g32.SetTextColor(hdc, _rgb(255, 255, 255))
                rc3 = _wt.RECT(28, 0, OW - 2, OH)
                _u32.DrawTextW(hdc, hex_str, -1, ctypes.byref(rc3), 0x0025)  # DT_SINGLELINE|DT_VCENTER|DT_LEFT
                _u32.EndPaint(hwnd, ps)
                return 0
            if msg == WM_DESTROY:
                _u32.PostQuitMessage(0)
                return 0
            return _u32.DefWindowProcW(hwnd, msg, wp, lp)

        def _rgb(r, g, b):
            return r | (g << 8) | (b << 16)

        # WNDPROC: LRESULT(HWND, UINT, WPARAM, LPARAM)
        # LPARAM on allekirjoitettu 64-bit arvo — c_ssize_t estää ylivuodon
        WNDPROC_TYPE = ctypes.WINFUNCTYPE(
            ctypes.c_ssize_t,
            ctypes.c_void_p,
            ctypes.c_uint,
            ctypes.c_ssize_t,
            ctypes.c_ssize_t,
        )
        _u32.DefWindowProcW.restype  = ctypes.c_ssize_t
        _u32.DefWindowProcW.argtypes = [
            ctypes.c_void_p, ctypes.c_uint,
            ctypes.c_ssize_t, ctypes.c_ssize_t,
        ]
        _wndproc_cb = WNDPROC_TYPE(_wndproc)

        hInst = ctypes.windll.kernel32.GetModuleHandleW(None)

        class WNDCLASSW(ctypes.Structure):
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

        wc2 = WNDCLASSW()
        wc2.lpfnWndProc   = ctypes.cast(_wndproc_cb, ctypes.c_void_p).value
        wc2.hInstance     = hInst
        wc2.lpszClassName = CLS_NAME
        _u32.RegisterClassW(ctypes.byref(wc2))

        hwnd = _u32.CreateWindowExW(
            WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_LAYERED | WS_EX_NOACTIVATE,
            CLS_NAME, "",
            WS_POPUP,
            0, 0, OW, OH,
            None, None, hInst, None
        )

        if not hwnd:
            return

        # Pyöristetty läpinäkyvyys
        ctypes.windll.user32.SetLayeredWindowAttributes(hwnd, 0, 230, 2)  # LWA_ALPHA
        _u32.ShowWindow(hwnd, 1)

        msg = _wt.MSG()
        while app.pipette_active:
            # Siirrä ikkuna kursorin viereen
            pt = _wt.POINT()
            _u32.GetCursorPos(ctypes.byref(pt))
            sw = _u32.GetSystemMetrics(0)
            sh = _u32.GetSystemMetrics(1)
            ox = pt.x + 20 if pt.x + 20 + OW < sw else pt.x - OW - 4
            oy = pt.y + 20 if pt.y + 20 + OH < sh else pt.y - OH - 4
            _u32.SetWindowPos(hwnd, None, ox, oy, 0, 0, 0x0001 | 0x0004)  # SWP_NOSIZE|SWP_NOZORDER
            _u32.InvalidateRect(hwnd, None, False)
            # Käsittele viestit
            while _u32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 1) > 0:
                _u32.TranslateMessage(ctypes.byref(msg))
                _u32.DispatchMessageW(ctypes.byref(msg))
            time.sleep(0.03)

        _u32.DestroyWindow(hwnd)
        _u32.UnregisterClassW(CLS_NAME, hInst)

    def pipette_thread():
        app.pipette_active = True

        # Käynnistä overlay omassa säikeessä
        threading.Thread(target=_pipette_overlay, daemon=True).start()

        import ctypes.wintypes as _wt

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
                ('dwExtraInfo', ctypes.c_ulong),
            ]

        _user32  = ctypes.windll.user32
        _kernel32 = ctypes.windll.kernel32
        _tid     = _kernel32.GetCurrentThreadId()
        _hook_id = [None]

        _HOOKPROC = ctypes.WINFUNCTYPE(
            ctypes.c_longlong,
            ctypes.c_int,
            ctypes.c_size_t,
            ctypes.c_size_t,
        )

        # Viimeisin hiiren sijainti jota capture-säie käyttää.
        # Hook tallentaa koordinaatit ja palataa HETI — ei raskaita operaatioita hookissa.
        _pip_pos   = [0, 0]
        _pip_event = threading.Event()

        def _capture_loop():
            """Erillinen säie joka tekee mss-kaappauksen kun hook signaloi."""
            try:
                sct = _mss_mod.mss()  # Luodaan KERRAN, käytetään koko session ajan
                while app.pipette_active:
                    _pip_event.wait(timeout=0.05)
                    _pip_event.clear()
                    if not app.pipette_active:
                        break
                    x, y = _pip_pos
                    try:
                        img = sct.grab({'top': y, 'left': x, 'width': 1, 'height': 1})
                        b2, g2, r2 = img.raw[0], img.raw[1], img.raw[2]
                        with _pip_lock:
                            app.pip_color = [r2, g2, b2, 255]
                    except Exception:
                        pass
                sct.close()
            except Exception:
                pass

        threading.Thread(target=_capture_loop, daemon=True).start()

        def _hook_proc(nCode, wParam, lParam):
            if nCode >= 0:
                ms = ctypes.cast(lParam, ctypes.POINTER(_MSLLHOOKSTRUCT)).contents
                if wParam == WM_MOUSEMOVE:
                    # Tallenna koordinaatit ja signaloi capture-säikeelle — EI raskaita
                    # operaatioita tässä, jotta WH_MOUSE_LL hook palaa välittömästi.
                    _pip_pos[0] = ms.x
                    _pip_pos[1] = ms.y
                    _pip_event.set()
                elif wParam == WM_LBUTTONDOWN:
                    app.pipette_active = False
                    _pip_event.set()  # Herätä capture-säie lopetusta varten
                    with _pip_lock:
                        c = list(app.pip_color)
                    app._pending_pipette_color = [c[0]/255., c[1]/255., c[2]/255., 1.]
                    # Tukahduta klikkaus ja poistu viestisilmukasta
                    _user32.PostThreadMessageW(_tid, WM_QUIT, 0, 0)
                    return 1  # suppress — ei välitetä muille ikkunoille

            return _user32.CallNextHookEx(_hook_id[0], nCode, wParam, lParam)

        _cb = _HOOKPROC(_hook_proc)
        # Aseta argtypes eksplisiittisesti juuri tälle kutsulle — välttää
        # välimuistissa olevan vanhan argtypes-asetuksen aiheuttaman TypeError
        _user32.SetWindowsHookExW.argtypes = [
            ctypes.c_int, _HOOKPROC, ctypes.c_void_p, ctypes.c_uint
        ]
        _user32.SetWindowsHookExW.restype = ctypes.c_void_p
        _hook_id[0] = _user32.SetWindowsHookExW(WH_MOUSE_LL, _cb, None, 0)

        # Viestisilmukka — PAKOLLINEN WH_MOUSE_LL:lle, ilman tätä hook ei toimi
        _msg = _wt.MSG()
        while _user32.GetMessageW(ctypes.byref(_msg), None, 0, 0) > 0:
            _user32.TranslateMessage(ctypes.byref(_msg))
            _user32.DispatchMessageW(ctypes.byref(_msg))

        if _hook_id[0]:
            _user32.UnhookWindowsHookEx(_hook_id[0])
            _hook_id[0] = None

    def start_pipette():
        threading.Thread(target=pipette_thread, daemon=True).start()

    _wheel_listener = [None]   # viittaus pysäytystä varten

    def mouse_wheel_thread():
        """Background thread to capture mouse wheel events for sliders."""
        def on_scroll(x, y, dx, dy):
            with _wheel_lock:
                app._mouse_wheel_delta = dy
        try:
            with _mouse_mod.Listener(on_scroll=on_scroll) as listener:
                _wheel_listener[0] = listener
                listener.join()
        except Exception:
            pass
        finally:
            _wheel_listener[0] = None

    # Start mouse wheel thread
    threading.Thread(target=mouse_wheel_thread, daemon=True).start()

    # -- HTML-vienti ------------------------------------------------------
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
        """Return a rich <tr>: data columns on the left, one Copy column on the right."""
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
                import tkinter as tk
                from tkinter import filedialog
                root = tk.Tk()
                root.withdraw()
                root.attributes('-topmost', True)
                path = filedialog.asksaveasfilename(
                    defaultextension=".html",
                    filetypes=[("HTML file", "*.html"), ("All files", "*.*")],
                    initialfile=default_name,
                    title="Save as",
                )
                root.destroy()
                if path:
                    title = os.path.splitext(os.path.basename(path))[0]
                    final = content.replace(
                        'HTMLTITLE_PLACEHOLDER',
                        f'<title>{title}</title>'
                    ).replace('HTMLH2_PLACEHOLDER', title)
                    with open(path, 'w', encoding='utf-8') as f:
                        f.write(final)
                    os.startfile(path)
            # _write_log() keeps error logging consistent.
            except Exception as e:
                _write_log(f"Save error: {e}\n" + traceback.format_exc())

        threading.Thread(target=_run, daemon=True).start()

    def _save_ase_dialog(colors, default_name):
        """Save colors as an ASE (Adobe Swatch Exchange) file in a background thread."""
        def _build_ase(palette_name, color_list):
            import struct
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
                import tkinter as tk
                from tkinter import filedialog
                root = tk.Tk()
                root.withdraw()
                root.attributes('-topmost', True)
                path = filedialog.asksaveasfilename(
                    defaultextension=".ase",
                    filetypes=[("Adobe Swatch Exchange", "*.ase"), ("All files", "*.*")],
                    initialfile=default_name,
                    title="Export ASE",
                )
                root.destroy()
                if path:
                    pal_name = os.path.splitext(os.path.basename(path))[0]
                    data = _build_ase(pal_name, colors)
                    with open(path, 'wb') as f:
                        f.write(data)
            except Exception as e:
                _write_log(f"ASE export error: {e}\n" + traceback.format_exc())

        threading.Thread(target=_run, daemon=True).start()

    def export_palette_ase_by_name(name):
        if not name or name not in app.palettes:
            return
        _save_ase_dialog(list(app.palettes[name]), name)

    def export_palette():
        mode = dpg.get_value("harmony_combo")
        h = (
            f"<html><head><meta charset='utf-8'>HTMLTITLE_PLACEHOLDER{_HCSS}</head><body>"
            f"<h2>HTMLH2_PLACEHOLDER</h2>"
            f"<table>{_html_table_header()}"
        )
        for tag in ["main", "h1", "h2", "h3", "h4", "h5", "h6", "h7"]:
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
        if not app.history:
            return
        h = (
            f"<html><head><meta charset='utf-8'>HTMLTITLE_PLACEHOLDER{_HCSS}</head><body>"
            f"<h2>HTMLH2_PLACEHOLDER</h2>"
            f"<table>{_html_table_header()}"
        )
        for v in app.history:
            h += _html_color_row(v[0], v[1], v[2])
        h += "</table></body></html>"
        _save_html_dialog(h, "color_history")

    def _minimize_win():
        """Minimize via Win32 since dpg.minimize_viewport() may not exist in all builds."""
        try:
            hwnd = ctypes.windll.user32.FindWindowW(None, "Color Tools")
            if hwnd:
                ctypes.windll.user32.ShowWindow(hwnd, 6)  # SW_MINIMIZE
        except Exception:
            pass

    # Map theme name → (app_theme, titlebar_theme, is_light)
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
        # Päivitä historian valintareunus teeman mukaan
        # Light-teemassa tumma reunus, muissa vaalea
        border_col = [60, 60, 60, 255] if is_light else [220, 220, 220, 255]
        if dpg.does_item_exist(_hist_selected_theme):
            dpg.delete_item(_hist_selected_theme, children_only=True)
            with dpg.theme_component(dpg.mvColorButton, parent=_hist_selected_theme):
                dpg.add_theme_color(dpg.mvThemeCol_Border, border_col,
                                    category=dpg.mvThemeCat_Core)
                dpg.add_theme_style(dpg.mvStyleVar_FrameBorderSize, 2)

    def slider_mode_cb(s, v):
        for m in SLIDER_MODES:
            dpg.configure_item(f"grp_sl_{m}", show=(m == v))
        app._last_sl_vals = {}

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

    # ================================================================
    #  UI BUILD
    # ================================================================
    dpg.create_context()

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
            dpg.add_theme_color(dpg.mvThemeCol_TabActive,      [58,  90,  130, 255])  # tummennettu Nord Frost
            dpg.add_theme_color(dpg.mvThemeCol_Separator,      [67,  76,  94,  255])
            dpg.add_theme_style(dpg.mvStyleVar_TabRounding,    2)
            dpg.add_theme_style(dpg.mvStyleVar_FramePadding,   4, 3)

    with dpg.theme() as _solarized_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_WindowBg,       [0,   43,  54,  255])  # Solarized base03
            dpg.add_theme_color(dpg.mvThemeCol_ChildBg,        [0,   43,  54,  255])
            dpg.add_theme_color(dpg.mvThemeCol_FrameBg,        [7,   54,  66,  255])  # base02
            dpg.add_theme_color(dpg.mvThemeCol_FrameBgHovered, [0,   72,  88,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Text,           [253, 246, 227, 255])  # base3 – riittävä kontrasti aktiivisella välilehdellä
            dpg.add_theme_color(dpg.mvThemeCol_Button,         [0,   72,  88,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered,  [0,   95,  115, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Header,         [42,  161, 152, 255])  # cyan accent
            dpg.add_theme_color(dpg.mvThemeCol_PopupBg,        [7,   54,  66,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ScrollbarBg,    [0,   33,  42,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Tab,            [7,   54,  66,  255])
            dpg.add_theme_color(dpg.mvThemeCol_TabHovered,     [42,  161, 152, 200])
            dpg.add_theme_color(dpg.mvThemeCol_TabActive,      [0,   95,  89,  255])  # tummennettu syaani – eroaa inaktiivisesta ja kontrastoi base3-tekstin kanssa
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

    # NOTE: _make_grad_slider -- smin/smax are not needed at build time;
    # gradient scaling is performed inside _draw_grad_slider.
    def _make_grad_slider(label, label_color, gname):
        with dpg.group(horizontal=True):
            dpg.add_text(label, color=label_color)
            dpg.add_drawlist(width=GRAD_W, height=THUMB_H, tag=f"dl_{gname}")
            dpg.add_text("   ", tag=f"slval_{gname}")

    with dpg.window(tag="PrimaryWindow"):

        # ── Custom titlebar (replaces the native Windows title bar) ──────────────
        with dpg.child_window(
            tag="titlebar_cw", width=_VP_W, height=_TB_H,
            border=False, no_scrollbar=True,
        ):
            dpg.bind_item_theme("titlebar_cw", _titlebar_dark_theme)
            with dpg.group(tag="titlebar", horizontal=True):
                # Logo image at native/near-native size for sharpness
                if _logo_dark_w > 0:
                    _tb_pad_y = max(0, (_TB_H - _logo_dark_h) // 2)
                    if _tb_pad_y > 0:
                        dpg.add_spacer(height=_tb_pad_y)
                    dpg.add_image(
                        "logo_dark", tag="tb_logo",
                        width=_logo_dark_w, height=_logo_dark_h,
                    )
                # Flexible spacer pushes control buttons to the right
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
                        )
                        dpg.add_combo(
                            FORMAT_OPTIONS, tag="fmt_combo",
                            default_value="HEX", width=89,
                            callback=refresh_harmony_values,
                        )
                    dpg.add_spacer(height=1)
                    # Scrollable area for the 8 colour rows.
                    # Scroll events are consumed here and do NOT propagate to
                    # the tab bar or the dropdowns above.
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
                            callback=cancel_palette_select_mode,
                            width=60, height=THUMB_H,
                            show=False,
                        )
                    # None resets the theme.
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

        dpg.add_separator()
        with dpg.group(horizontal=True):
            dpg.add_checkbox(label="Always on Top", callback=toggle_on_top)
            dpg.add_text("  Theme:")
            dpg.add_combo(
                THEME_NAMES, tag="theme_combo",
                default_value=app.theme_name,
                width=_sc(82), callback=set_theme,
            )

    # ================================================================
    #  UPDATE LOOP
    # ================================================================
    def _draw_grad_slider(gname, smin, smax):
        tag_dl = f"dl_{gname}"
        if not dpg.does_item_exist(tag_dl):
            return
        dpg.delete_item(tag_dl, children_only=True)
        val = app.sl_vals.get(gname, smin)
        y0  = (THUMB_H - GRAD_H) // 2
        y1  = y0 + GRAD_H
        sw  = GRAD_W / SEGS
        ctx = _grad_ctx()  # Compute base colour values once for all 48 segments
        for i in range(SEGS):
            t = (i + 0.5) / SEGS
            r, g, b = _grad_color(gname, t, ctx)
            x0s = int(i * sw)
            x1s = max(x0s + 1, int((i + 1) * sw))
            dpg.draw_rectangle(
                parent=tag_dl,
                pmin=(x0s, y0), pmax=(x1s, y1),
                fill=[r, g, b, 255], color=[0, 0, 0, 0],
            )
        dpg.draw_rectangle(
            parent=tag_dl,
            pmin=(0, y0), pmax=(GRAD_W, y1),
            fill=[0, 0, 0, 0], color=[70, 70, 70, 180],
        )
        t_val   = (val - smin) / (smax - smin) if smax > smin else 0.
        tx      = int(t_val * (GRAD_W - 1))
        ind_col = [30, 30, 30, 240] if app.theme_name == "Light" else [255, 255, 255, 240]
        dpg.draw_line(
            parent=tag_dl,
            p1=(tx, y0 - 2), p2=(tx, y1 + 2),
            color=ind_col, thickness=2,
        )
        dpg.set_value(f"slval_{gname}", f"{val:>4}")

    # _POINT defined once here, not inside update() to avoid recreation every frame.
    class _POINT(ctypes.Structure):
        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

    def update():
        # -- Frame-delayed Color Picker Sync -------------------------
        # Laske undo-kirjoituslukko alas framea kohti
        if getattr(app, '_pal_undo_freeze', 0) > 0:
            app._pal_undo_freeze -= 1
        if getattr(app, '_pending_wheel_sync', None) is not None:
            if dpg.does_item_exist("picker_wheel"):
                dpg.set_value("picker_wheel", app._pending_wheel_sync)
            app._pending_wheel_sync = None

        # -- Pipette: apply color set by background thread (DPG is not thread-safe) --
        if app._pending_pipette_color is not None:
            v = app._pending_pipette_color
            app._pending_pipette_color = None
            _force_update_color_picker(v)
            _sync_all_modes()
            add_to_history(v)

        # -- Flush deferred save requested by background threads --------------
        if app._pending_save:
            app._pending_save = False
            save_config()

        vsig = color_sig(app.current_color)
        mode = dpg.get_value("harmony_combo")
        fmt  = dpg.get_value("fmt_combo")
        pip  = app.pipette_active
        # Only acquire pip_lock when pipette is actually running
        if pip:
            with _pip_lock:
                pip_color = list(app.pip_color)
        else:
            pip_color = None
        dpg.set_value("preview", pip_color if pip else list(vsig) + [255])
        # Only update preview_hex when colour changed and user is not typing
        if not pip and vsig != app._last_preview_sig:
            if not dpg.is_item_active("preview_hex"):
                r, g, b = vsig
                dpg.set_value("preview_hex", f"#{r:02x}{g:02x}{b:02x}".upper())
                app._last_preview_sig = vsig

        # -- Titlebar icon buttons (minimize / close) -------------------------
        # Redrawn only when hover state changes.
        for _dl_tag, _is_close in (("dl_tb_min", False), ("dl_tb_close", True)):
            if not dpg.does_item_exist(_dl_tag):
                continue
            _hov = dpg.is_item_hovered(_dl_tag)
            _clk = dpg.is_item_clicked(_dl_tag)
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

        # -- Titlebar drag (moves the borderless window) ----------------------
        # GetCursorPos only when LMB is down to avoid a syscall on every idle frame.
        lmb_dn_tb  = dpg.is_mouse_button_down(0)
        lmb_clk_tb = dpg.is_mouse_button_clicked(0)
        mx_g = my_g = 0
        if lmb_dn_tb or lmb_clk_tb:
            _pt = _POINT()
            ctypes.windll.user32.GetCursorPos(ctypes.byref(_pt))
            mx_g, my_g = _pt.x, _pt.y
        if lmb_clk_tb and dpg.does_item_exist("titlebar_cw"):
            if (dpg.is_item_hovered("titlebar_cw")
                    and not dpg.is_item_hovered("dl_tb_min")
                    and not dpg.is_item_hovered("dl_tb_close")):
                app._tb_dragging = True
                vx, vy = dpg.get_viewport_pos()
                app._tb_drag_start_vp = (mx_g - vx, my_g - vy)
        if not lmb_dn_tb:
            app._tb_dragging = False
        if app._tb_dragging and lmb_dn_tb:
            ox, oy = app._tb_drag_start_vp
            dpg.set_viewport_pos([mx_g - ox, my_g - oy])

        # -- Sliders --------------------------------------------------
        if not app.use_wheel:
            smode    = dpg.get_value("slider_mode_combo")
            current  = SLIDERS_BY_MODE.get(smode, [])
            mx, my   = dpg.get_mouse_pos()
            lmb_down    = dpg.is_mouse_button_down(0)
            lmb_clicked = dpg.is_mouse_button_clicked(0)
            with _wheel_lock:
                wheel_delta = app._mouse_wheel_delta
                app._mouse_wheel_delta = 0  # Reset for next frame
            
            for (lbl, gname, lcol, smin, smax) in current:
                tag_dl = _DL_TAGS[gname]
                if not dpg.does_item_exist(tag_dl):
                    continue
                hovered = dpg.is_item_hovered(tag_dl)
                if lmb_clicked and hovered:
                    app._dragging = gname
                if not lmb_down and app._dragging == gname:
                    app._dragging = None
                if app._dragging == gname and lmb_down:
                    rect_min = dpg.get_item_rect_min(tag_dl)
                    lx = mx - rect_min[0]
                    t  = max(0., min(1., lx / (GRAD_W - 1)))
                    app.sl_vals[gname] = int(round(smin + (smax - smin) * t))
                    _apply_from_mode(smode)
                    _update_selected_pal_color()
                # Mouse wheel scroll: wheel_delta > 0 = up (increase), < 0 = down (decrease)
                if hovered and wheel_delta != 0 and not lmb_down:
                    step = max(1, (smax - smin) // 50)
                    new_val = app.sl_vals.get(gname, smin) + int(wheel_delta * step / abs(wheel_delta))
                    app.sl_vals[gname] = max(smin, min(smax, new_val))
                    _apply_from_mode(smode)
                    _update_selected_pal_color()
                if (app._last_sl_vals.get(gname) != app.sl_vals.get(gname)
                        or app._last_sl_vals.get('_sig') != vsig):
                    _draw_grad_slider(gname, smin, smax)
            app._last_sl_vals = dict(app.sl_vals)
            app._last_sl_vals['_sig'] = vsig

        # -- Harmony colours ------------------------------------------
        if vsig != app._last_sig or mode != app._last_mode or fmt != app._last_fmt:
            rf, gf, bf = [clamp01(c) for c in app.current_color[:3]]
            # colorsys.rgb_to_hls returns (h, l, s).
            h_, l_, s_ = colorsys.rgb_to_hls(rf, gf, bf)
            
            updates = [("main", rf, gf, bf)]
            
            # Handle angle-based harmonies
            angles = {
                "Complementary":       [180],
                "Split Complementary": [150, 210],
                "Analogous":           [30, -30],
                "Triadic":             [120, 240],
                "Tetradic":            [90, 180, 270],
                "Rectangle":           [60, 180, 240],
            }.get(mode, [])
            
            if mode in ["Tints", "Shades", "Tones"]:
                # Generate tints, shades, or tones (7 colors total)
                if mode == "Tints":
                    # Tint = mix with white (increase lightness)
                    for i in range(1, 7):
                        t = i / 7.0
                        new_l = l_ + (1 - l_) * t
                        nr, ng, nb = colorsys.hls_to_rgb(h_, new_l, s_)
                        updates.append((f"h{i}", nr, ng, nb))
                elif mode == "Shades":
                    # Shade = mix with black (decrease lightness)
                    for i in range(1, 7):
                        t = i / 7.0
                        new_l = l_ * (1 - t)
                        nr, ng, nb = colorsys.hls_to_rgb(h_, new_l, s_)
                        updates.append((f"h{i}", nr, ng, nb))
                elif mode == "Tones":
                    # Tone = mix with gray (decrease saturation)
                    for i in range(1, 7):
                        t = i / 7.0
                        new_s = s_ * (1 - t)
                        nr, ng, nb = colorsys.hls_to_rgb(h_, l_, new_s)
                        updates.append((f"h{i}", nr, ng, nb))
            elif angles:
                # Angle-based harmonies
                for i, ang in enumerate(angles):
                    nh = (h_ + ang / 360.) % 1.
                    # colorsys.hls_to_rgb takes (h, l, s).
                    nr, ng, nb = colorsys.hls_to_rgb(nh, l_, s_)
                    updates.append((f"h{i+1}", nr, ng, nb))
            
            app.harmony_rgb = {}
            is_contrast = (fmt == "Contrast")
            show_copy = not is_contrast and mode not in ("Tints", "Shades", "Tones")
            for i, tag in enumerate(HARMONY_TAGS):
                keys = HARMONY_KEYS[tag]
                if i < len(updates):
                    dpg.configure_item(keys["group"], show=True)
                    rgb2 = tuple(to_int(c) for c in updates[i][1:])
                    app.harmony_rgb[tag] = rgb2
                    dpg.set_value(keys["rect"], list(rgb2) + [255])
                    # Show either plain text field or coloured contrast display
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
                        # Update preview bar text colors
                        if dpg.does_item_exist(keys["ctr_bar_w"]):
                            dpg.configure_item(keys["ctr_bar_w"], color=[r2, g2, b2, 255])
                        if dpg.does_item_exist(keys["ctr_bar_b"]):
                            dpg.configure_item(keys["ctr_bar_b"], color=[r2, g2, b2, 255])
                    else:
                        dpg.set_value(keys["val"], format_value(tag, fmt))
                else:
                    dpg.configure_item(keys["group"], show=False)
            app._last_sig  = vsig
            app._last_mode = mode
            app._last_fmt  = fmt

        # -- History --------------------------------------------------
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

        # -- Image import error message (auto-hides after 5 s) ---------------
        if dpg.does_item_exist("import_error_text"):
            if app._import_error and time.time() < app._import_error_until:
                dpg.set_value("import_error_text", app._import_error)
                dpg.configure_item("import_error_text", show=True)
            else:
                if app._import_error:
                    app._import_error = ""
                dpg.configure_item("import_error_text", show=False)

        # Historia right-click poisto + save
        if dpg.is_mouse_button_clicked(1):
            removed = False
            with _history_lock:
                for i in range(len(app.history)):
                    if dpg.is_item_hovered(HIST_TAGS[i]):
                        tmp = list(app.history)
                        del tmp[i]
                        app.history = deque(tmp, maxlen=60)
                        app._hist_dirty = True
                        app.selected_history_indices.discard(i)
                        removed = True
                        break
            if removed:
                # Reset selection state on removal so indices do not
                # point to wrong colours when saving a palette.
                if app.palette_select_mode:
                    _exit_palette_select_mode()
                save_config()

        if app.change_time > 0 and (time.time() - app.change_time) > 0.5:
            if not dpg.is_item_active("picker_wheel"):
                add_to_history(app.current_color)
                app.change_time = 0.

        _pal_swatch_right_click()
        _pal_swatch_drag()

        # -- Palette edit panel: refresh when active slot color changes -------
        # Done here (once per frame) instead of inside picker_cb to avoid
        # rebuilding the panel hundreds of times per second while dragging.
        if app._editing_pal and getattr(app, '_pal_selected_idx', None) is not None:
            if vsig != app._last_edit_color_sig:
                app._last_edit_color_sig = vsig
                _refresh_pal_edit_panel()

        # Palette rebuild: requested by background thread, executed by the main thread.
        if app._pending_pal_rebuild:
            app._pending_pal_rebuild = False
            _rebuild_pal_rows()

        _pal_order_set = set(app._pal_order)
        pal_sig = (
            tuple(app._pal_order)
            + tuple(k for k in sorted(app.palettes) if k not in _pal_order_set)
        )
        if pal_sig != app._last_pal:
            _rebuild_pal_rows()
            app._last_pal = pal_sig

    # ================================================================
    #  STARTUP
    # ================================================================
    dpg.create_viewport(
        title='Color Tools',
        width=int(283*_DPI_SCALE),
        height=int(623*_DPI_SCALE),   # same total height as original; custom bar replaces native
        x_pos=_start_pos[0], y_pos=_start_pos[1],
        resizable=False,
        decorated=False,   # hide native Windows title bar; custom bar is used instead
    )
    dpg.setup_dearpygui()
    dpg.bind_theme(_dark_theme)  # overridden below after context is ready

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
    dpg.render_dearpygui_frame()

    # Set titlebar spacer once — resizable=False so viewport width never changes.
    if dpg.does_item_exist("tb_spacer"):
        _vw     = dpg.get_viewport_width()
        _logo_w = _logo_dark_w or 0
        _sp_w   = max(2, _vw - _logo_w - 2 * _ICON_W - 3 * _sc(2) - 17)
        dpg.configure_item("tb_spacer", width=_sp_w)

    # Apply saved theme (titlebar + logo + global)
    set_theme(None, app.theme_name)

    _sync_all_modes()
    _rebuild_pal_rows()

    # Win32 DWM calls run in a background thread (does not touch DPG elements).
    threading.Thread(target=set_window_style, daemon=True).start()

    while dpg.is_dearpygui_running():
        update()
        dpg.render_dearpygui_frame()

    save_config()
    # Pysäytä wheel-listener siististi jotta hiiri ei jumita sulkeutuessa
    if _wheel_listener[0] is not None:
        try:
            _wheel_listener[0].stop()
        except Exception:
            pass
    dpg.destroy_context()

except Exception:
    _fatal(traceback.format_exc())
