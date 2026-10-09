"""
state.py — Sovelluksen tila, lukot, konfiguraatio.

AppState-luokka sisältää kaiken ajoaikaisen tilan.
load_config / save_config lukevat ja kirjoittavat config.json:n.
"""

import ctypes
import ctypes.wintypes as _wt
import json
import math
import os
import shutil
import threading
import traceback
import uuid
from collections import deque

from .storage import atomic_write

from .bootstrap import (
    CFG_DIR, CFG_FILE, _write_log,
)

# ── Lukot ────────────────────────────────────────────────────────────────────
_history_lock  = threading.Lock()
_palettes_lock = threading.Lock()
_pip_lock      = threading.Lock()
_save_lock     = threading.Lock()
_config_write_blocked = False
_MAX_CONFIG_BYTES = 20 * 1024 * 1024

# ── Sovelluksen tila ──────────────────────────────────────────────────────────
class AppState:
    def __init__(self):
        self.pipette_active       = False
        self.pip_color            = [0, 0, 0, 255]
        self._pipette_thread_tid  = 0
        self._pipette_hotkey_held = False
        self.current_color   = [0.2, 0.5, 1.0, 1.0]
        self.contrast_bg     = [255, 255, 255, 255]
        self.history         = deque(maxlen=60)
        self.change_time     = 0.
        self.use_wheel       = True
        self.top_tab         = "wheel"  # "wheel" | "sliders" | "contrast"
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
        self._hist_dirty        = True
        self._last_sl_vals      = {}
        self._pal_order         = []
        self._editing_pal       = None
        self.theme_name         = "Dark"
        self.always_on_top      = False
        self.tab_area_collapsed = False
        self.utility_collapsed  = False
        self.top_view_collapsed = False
        self.palette_select_mode          = False
        self.selected_history_indices     = set()
        self.selected_history_colors      = {}
        self._import_status       = ("", 0.0)
        self._pal_drag_idx        = None
        self._pal_drag_active     = False
        self._pal_drag_start      = (0, 0)
        self._pal_drag_insert     = None
        # -- Custom titlebar drag state --
        self._tb_dragging         = False
        self._tb_drag_start_vp    = (0, 0)
        self._tb_hwnd_cache       = None
        self._tb_lmb_prev         = False
        self._tb_hov              = {"dl_tb_min": None, "dl_tb_close": None}
        self._vp_size_guard_skip  = 0

        self._pal_selected_idx    = None
        self._pal_page = 0
        self._pal_edit_page = 0
        self._pal_pre_edit_snapshot = None
        self._pal_edit_dirty      = False
        self._pal_undo_freeze     = 0
        self._last_picker_force_time = 0.0
        self._pending_wheel_sync     = None
        self._pending_picker_color   = None
        self._pal_undo_stack         = deque(maxlen=30)
        self._mouse_wheel_delta      = 0
        self._pending_pipette_color  = None
        self._picker_processing      = False
        self._pending_save           = False
        self._pending_close_edit     = False
        self._palette_select_cancel_pending = False
        self._pending_drop_path      = None
        self._drop_preview_colors_pending = []
        self._image_modal_open       = False
        self._image_modal_pil        = None
        self._image_modal_colors     = []
        self._drop_preview_open      = False
        self._config_error           = ""
        self._last_edit_color_sig    = None
        self._last_preview_sig       = None
        self._last_viewport_pos      = [100, 100]
        self._cached_vp_w           = 0
        self._cached_vp_h           = 0
        self._pal_dirty              = True
        self._harmony_text_dirty     = True
        self._vp_pos_frame_ctr       = 0
        self._last_drawn_preview     = None
        # Peilataan combo-widgettien arvoja muistissa — luetaan muistista joka ruutu
        self.harmony_mode            = "Triadic"
        self.fmt_mode                = "HEX"
        self.slider_mode             = "RGB"
        self.language                = "en"
        self._lang_dirty             = False
        self._last_theory_key        = None
        self._tabs_calibrated        = False
        self._history_save_due       = 0.0   # viivästetty historian tallennus (perf_counter)
        # Asetetaan myöhemmin update_loop / ui_main / eyedropper -moduuleissa
        self._fixed_vp_w             = 0
        self._fixed_vp_h             = 0
        self._pipette_hook_cb        = None
        self._pip_wndproc_cb         = None


# ── Globaali app-olio (luodaan kerran) ───────────────────────────────────────
app = AppState()

# ── Konfiguraatio ────────────────────────────────────────────────────────────
def _backup_broken_config():
    """Kopioi lukukelvottoman config.json:n talteen, ennen kuin save_config
    ehtii ylikirjoittaa sen oletusarvoilla."""
    try:
        if not os.path.exists(CFG_FILE):
            return None
        bak = CFG_FILE + '.bak'
        if os.path.exists(bak):
            # Älä ylikirjoita aiempaa varmuuskopiota
            bak = f"{CFG_FILE}.{uuid.uuid4().hex}.bak"
        shutil.copy2(CFG_FILE, bak)
        _write_log(f"config.json could not be read. Backup saved to: {bak}")
        return bak
    except Exception:
        _write_log("Backing up broken config failed:\n" + traceback.format_exc())
        return None

def load_config():
    global _config_write_blocked
    _config_write_blocked = False
    try:
        with open(CFG_FILE, 'rb') as f:
            data = f.read(_MAX_CONFIG_BYTES + 1)
        if len(data) > _MAX_CONFIG_BYTES:
            # Do not replace a possibly valid large file with defaults.
            _config_write_blocked = True
            raise ValueError("config exceeds the size limit")
        cfg = json.loads(data)
        if not isinstance(cfg, dict):
            raise ValueError("config root is not a JSON object")
        return cfg
    except FileNotFoundError:
        return {}
    except Exception:
        _write_log("load_config failed:\n" + traceback.format_exc())
        if not _backup_broken_config():
            _config_write_blocked = True
        app._config_error = 'config_blocked' if _config_write_blocked else 'config_recovered'
        return {}

def save_config():
    """Serialize snapshots and atomic writes, including callback callers."""
    with _save_lock:
        return _save_config_locked()


def _save_config_locked():
    if _config_write_blocked:
        app._config_error = 'config_blocked'
        return False
    try:
        os.makedirs(CFG_DIR, exist_ok=True)
        with _history_lock:
            history_snapshot = [list(c) for c in app.history]
        with _palettes_lock:
            palettes_snapshot  = {k: [list(c) for c in v] for k, v in app.palettes.items()}
            pal_order_snapshot = list(app._pal_order)
        data = json.dumps({
                'history':   history_snapshot,
                'palettes':  palettes_snapshot,
                'pal_order': pal_order_snapshot,
                'win_pos':    list(app._last_viewport_pos),
                'theme_name': app.theme_name,
                'always_on_top': app.always_on_top,
                'language':   app.language,
                'tab_area_collapsed': app.tab_area_collapsed,
                'utility_collapsed': app.utility_collapsed,
                'top_view_collapsed': app.top_view_collapsed,
            }, ensure_ascii=False, allow_nan=False)
        if len(data.encode('utf-8')) > _MAX_CONFIG_BYTES:
            raise ValueError('settings exceed the reloadable size limit')
        atomic_write(CFG_FILE, data)
        if app._config_error == 'config_save_failed':
            app._config_error = ''
        return True
    except Exception:
        _write_log("save_config failed:\n" + traceback.format_exc())
        app._config_error = 'config_save_failed'
        return False


THEME_CHOICES    = ("Dark", "Light", "Midnight", "Mocha", "Nord", "Solarized", "Black")
LANGUAGE_CHOICES = ("en", "fi")

def _is_num(v):
    return (isinstance(v, int) and not isinstance(v, bool)
            or isinstance(v, float) and math.isfinite(v))

def _clean_color(c):
    """Palauttaa [r, g, b, 255] (kokonaisluvut 0-255) tai None jos c ei kelpaa."""
    if (isinstance(c, (list, tuple)) and len(c) >= 3
            and all(_is_num(v) for v in c[:3])):
        return [max(0, min(255, int(round(v)))) for v in c[:3]] + [255]
    return None

def _clean_colors(seq):
    out = []
    for c in seq:
        cc = _clean_color(c)
        if cc is not None:
            out.append(cc)
    return out

def apply_config(cfg):
    """Soveltaa ladatun konfiguraation app-olioon.

    Kaikki arvot validoidaan: rikkinäinen mutta kelvollinen JSON (esim.
    "pal_order": null tai väärän muotoinen win_pos) ei saa kaataa käynnistystä."""
    if not isinstance(cfg, dict):
        cfg = {}

    hist = cfg.get('history')
    if isinstance(hist, list):
        app.history = deque(_clean_colors(hist)[:60], maxlen=60)

    pals = cfg.get('palettes')
    if isinstance(pals, dict):
        app.palettes = {str(k): _clean_colors(v)
                        for k, v in pals.items() if isinstance(v, list)}
        order = cfg.get('pal_order')
        if not isinstance(order, list):
            order = []
        app._pal_order = []
        for k in order:
            if isinstance(k, str) and k in app.palettes and k not in app._pal_order:
                app._pal_order.append(k)
        for k in app.palettes:
            if k not in app._pal_order:
                app._pal_order.append(k)

    pos = cfg.get('win_pos')
    if (isinstance(pos, (list, tuple)) and len(pos) == 2
            and all(_is_num(v) and -2**31 <= v < 2**31 for v in pos)):
        start_pos = [int(pos[0]), int(pos[1])]
    else:
        start_pos = [100, 100]
    # Tarkista että ikkuna on näkyvällä monitorilla
    try:
        _MONITOR_DEFAULTTONULL = 0
        _pt_check = _wt.POINT(start_pos[0], start_pos[1])
        if not ctypes.windll.user32.MonitorFromPoint(_pt_check, _MONITOR_DEFAULTTONULL):
            start_pos = [100, 100]
    except Exception:
        pass
    app._last_viewport_pos = list(start_pos)

    theme = cfg.get('theme_name')
    app.theme_name    = theme if theme in THEME_CHOICES else 'Dark'
    app.always_on_top = bool(cfg.get('always_on_top', False))
    lang = cfg.get('language')
    app.language      = lang if lang in LANGUAGE_CHOICES else 'en'
    app.tab_area_collapsed = bool(cfg.get('tab_area_collapsed', False))
    app.utility_collapsed  = bool(cfg.get('utility_collapsed', False))
    app.top_view_collapsed = bool(cfg.get('top_view_collapsed', False))
    return start_pos
