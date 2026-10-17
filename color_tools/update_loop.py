"""
update_loop.py — update()-silmukka, käynnistys ja sammutus.
"""

import ctypes
import ctypes.wintypes as _wt
import time
import traceback
import colorsys

import dearpygui.dearpygui as dpg

from .bootstrap import _sc, _get_own_hwnd, _write_log, acquire_instance_lock
from .state import app, _pip_lock, _history_lock, save_config, apply_config, load_config
from .color_math import color_sig, clamp01, to_int
from .sliders import (
    SLIDERS_BY_MODE, _DL_TAGS, _grad_ctx, _draw_grad_slider,
    _apply_from_mode, _sync_all_modes,
)
from .harmony import (
    HARMONY_TAGS, HIST_TAGS, HARMONY_KEYS, _HARMONY_ANGLES,
    refresh_harmony_values, format_value, color_theory_info,
    add_to_history, _force_update_color_picker, _update_wheel_contrast,
    _update_selected_pal_color, picker_cb, _wc_last_sig,
    _apply_external_color_to_pal,
)
from .palette import (
    _update_history_selection_style, _exit_palette_select_mode,
    _rebuild_pal_rows, _refresh_pal_edit_panel,
    _pal_swatch_right_click, _pal_swatch_drag,
    _update_active_pal_swatch,
)
from .import_export import (
    _import_from_path, process_import_results, shutdown_io,
)
from .eyedropper import (
    _pip_overlay_update, _pip_free_gdi, _pip_hwnd, _pip_cls, _pip_u32,
    _pumping_sleep, init_pip_overlay,
    start_hotkey, stop_hotkey,
)
from . import sliders as _sliders_mod
from . import themes as _themes_mod
from .themes import set_theme
from .ui_main import (
    build_ui, W, _TB_H, _ICON_W, _IC, _THICK_ICON, _minimize_win,
    _VP_W, _VP_H, tab_label, calibrate_tabs, pick_color_label,
)

# ── Update-silmukka ────────────────────────────────────────────────────────────
_pt = _wt.POINT()

_ut_u32 = ctypes.windll.user32
_ut_u32.WindowFromPoint.restype  = ctypes.c_void_p
_ut_u32.WindowFromPoint.argtypes = [_wt.POINT]
_ut_u32.GetAncestor.restype      = ctypes.c_void_p
_ut_u32.GetAncestor.argtypes     = [ctypes.c_void_p, ctypes.c_uint]

_ut_u32.GetWindowRect.argtypes = [ctypes.c_void_p, ctypes.POINTER(_wt.RECT)]
_ut_u32.GetWindowRect.restype = ctypes.c_bool
_ut_u32.IsWindow.argtypes = [ctypes.c_void_p]
_ut_u32.IsWindow.restype = ctypes.c_bool
_ut_u32.IsIconic.argtypes = [ctypes.c_void_p]
_ut_u32.IsIconic.restype = ctypes.c_bool
_ut_u32.SetWindowPos.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_int,
                               ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_uint]
_ut_u32.SetWindowPos.restype = ctypes.c_bool


def _native_window_rect():
    hwnd = app._tb_hwnd_cache
    if not hwnd or not _ut_u32.IsWindow(hwnd):
        hwnd = _get_own_hwnd()
        app._tb_hwnd_cache = hwnd
    rect = _wt.RECT()
    if not hwnd or _ut_u32.IsIconic(hwnd) or not _ut_u32.GetWindowRect(hwnd, ctypes.byref(rect)):
        return None
    return hwnd, rect


def _cursor_over_window(hwnd, x, y):
    """True jos annetussa pisteessä on (päällimmäisenä) tämän sovelluksen ikkuna."""
    try:
        h = _ut_u32.WindowFromPoint(_wt.POINT(int(x), int(y)))
        if not h:
            return False
        return _ut_u32.GetAncestor(h, 2) == hwnd  # GA_ROOT
    except Exception:
        return True

def _guard_viewport_size():
    """Pitää viewportin koon kiinteänä (ei käyttäjän venytystä)."""
    # Viewport-kokosuojaus
    if app._vp_size_guard_skip > 0:
        app._vp_size_guard_skip -= 1
    elif not app._tb_dragging:
        _vw = dpg.get_viewport_width()
        _vh = dpg.get_viewport_height()
        if _vw != app._cached_vp_w or _vh != app._cached_vp_h:
            app._cached_vp_w = _vw
            app._cached_vp_h = _vh
            if _vw != app._fixed_vp_w or _vh != app._fixed_vp_h:
                dpg.set_viewport_width(app._fixed_vp_w)
                dpg.set_viewport_height(app._fixed_vp_h)
                app._vp_size_guard_skip = 5


def _apply_pending_picker():
    """Soveltaa odottavan väripyörä-/pipettivärin ja pakottaa pickerin päivityksen."""
    # Väripyörän päivitys (delete+recreate välillä ruutujen)
    if app._pending_picker_color is not None:
        _wheel_active = (dpg.does_item_exist("picker_wheel")
                         and dpg.is_item_active("picker_wheel"))
        if not _wheel_active:
            # A later slider/callback edit may have superseded the queued color.
            _pc = list(app.current_color)
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
            app._pending_wheel_sync = col_255
    if app._pending_wheel_sync is not None:
        if dpg.does_item_exist("picker_wheel"):
            app._picker_processing = True
            try:
                dpg.set_value("picker_wheel", app._pending_wheel_sync)
            finally:
                app._picker_processing = False
        app._pending_wheel_sync = None


def _apply_pending_actions():
    """Käsittelee muiden säikeiden/callbackien jättämät odottavat toiminnot."""
    process_import_results()
    # Pipetti: sovella taustasilmukan asettama väri
    if app._pending_pipette_color is not None:
        v = app._pending_pipette_color
        app._pending_pipette_color = None
        _force_update_color_picker(v)
        _sync_all_modes()
        _update_wheel_contrast()
        add_to_history(v)
        _apply_external_color_to_pal()

    if app._palette_save_due and time.perf_counter() >= app._palette_save_due:
        save_config()

    if app._pending_save:
        app._pending_save = False
        save_config()

    # Historian viivästetty tallennus (kaatuminen/pakotettu sulkeminen ei vie historiaa)
    if app._history_save_due and time.perf_counter() >= app._history_save_due:
        app._history_save_due = 0.0
        save_config()

    if app._pending_close_edit:
        app._pending_close_edit = False
        _rebuild_pal_rows()
        _refresh_pal_edit_panel()

    if app._pending_drop_path is not None:
        _path = app._pending_drop_path
        app._pending_drop_path = None
        _import_from_path(_path)


def _track_viewport_pos(_lmb_down):
    """Persist signed native coordinates, never DPG's cached viewport position."""
    app._vp_pos_frame_ctr += 1
    if app._tb_dragging or _lmb_down or app._vp_pos_frame_ctr >= 30:
        app._vp_pos_frame_ctr = 0
        native = _native_window_rect()
        if native:
            _, rect = native
            app._last_viewport_pos = [rect.left, rect.top]


def _update_preview(vsig):
    """Esikatselu, pipetin väri ja hex-kenttä."""
    pip  = app.pipette_active
    if pip:
        with _pip_lock:
            pip_color = list(app.pip_color)
    else:
        pip_color = None
    _new_preview = pip_color if pip else [*list(vsig), 255]
    if _new_preview != app._last_drawn_preview:
        dpg.set_value("preview", _new_preview)
        app._last_drawn_preview = _new_preview
    _pip_overlay_update()

    if not pip and vsig != app._last_preview_sig:
        if not dpg.is_item_active("preview_hex"):
            r, g, b = vsig
            dpg.set_value("preview_hex", f"#{r:02x}{g:02x}{b:02x}".upper())
            app._last_preview_sig = vsig


def _update_theory_tab(vsig):
    """Teoria-välilehden teksti (vain kun väri tai kieli muuttuu)."""
    if dpg.does_item_exist("theory_text"):
        # Päivitetään vain kun väri tai kieli muuttuu (ei joka ruutu)
        _theory_key = (vsig, app.language)
        if _theory_key != app._last_theory_key:
            app._last_theory_key = _theory_key
            info = color_theory_info(vsig)
            dpg.set_value("theory_title", info["name"])
            dpg.set_value("theory_hex", f'#{info["hex"]}')
            dpg.set_value("theory_tags", info["tags_text"])
            dpg.set_value("theory_association_lbl", info["association_heading"])
            dpg.set_value("theory_temp_lbl", info["temp_heading"])
            dpg.set_value("theory_text", info["body"])
            if dpg.does_item_exist("theory_temp_dl"):
                _w = dpg.get_item_width("theory_temp_dl") or 1
                _h = dpg.get_item_height("theory_temp_dl") or 1
                _x = max(3, min(_w - 3, int(round(info["temp_value"] * (_w - 1)))))
                for _mark in ("theory_temp_mark_bg", "theory_temp_mark"):
                    dpg.configure_item(_mark, p1=(_x, 0), p2=(_x, _h))


def _update_titlebar_icons(_lmb_clicked):
    """Mukautetun otsikkopalkin ikonit (piirto, hover, klikkaus)."""
    # Titlebar-ikonit
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


def _drag_titlebar():
    """Hit-test in the UI; move using one native screen-coordinate space."""
    primary = 0x02 if _ut_u32.GetSystemMetrics(23) else 0x01
    down = bool(_ut_u32.GetAsyncKeyState(primary) & 0x8000)
    clicked = down and not app._tb_lmb_prev
    app._tb_lmb_prev = down
    if not down:
        if app._tb_dragging:
            app._vp_size_guard_skip = 5
        app._tb_dragging = False
        return
    if not clicked and not app._tb_dragging:
        return
    native = _native_window_rect()
    if not native or not _ut_u32.GetCursorPos(ctypes.byref(_pt)):
        app._tb_dragging = False
        return
    hwnd, rect = native
    if clicked:
        # DPG's hit-test is already relative to the rendered title bar and its
        # DPI. Subtracting its cached viewport origin from Win32 coordinates
        # can reject every click after moving to another monitor.
        app._tb_dragging = (
            dpg.does_item_exist('titlebar_cw')
            and dpg.is_item_hovered('titlebar_cw')
            and not dpg.is_item_hovered('dl_tb_min')
            and not dpg.is_item_hovered('dl_tb_close')
            and _cursor_over_window(hwnd, _pt.x, _pt.y)
        )
        if app._tb_dragging:
            app._tb_drag_start_vp = (_pt.x - rect.left, _pt.y - rect.top)
    if app._tb_dragging:
        ox, oy = app._tb_drag_start_vp
        # Preserve size, z-order and activation while moving (including x/y < 0).
        if not _ut_u32.SetWindowPos(hwnd, None, _pt.x - ox, _pt.y - oy,
                                   0, 0, 0x0001 | 0x0004 | 0x0010):
            app._tb_dragging = False
            _write_log('Titlebar: SetWindowPos failed')
        else:
            actual = _native_window_rect()
            if actual:
                app._last_viewport_pos = [actual[1].left, actual[1].top]


def _update_sliders(_lmb_down, _lmb_clicked, vsig):
    """Gradienttiliukusäätimet: veto, rullalla säätö, piirto."""
    # Liukusäätimet
    wheel_delta            = app._mouse_wheel_delta
    app._mouse_wheel_delta = 0
    if not app.use_wheel and app.top_tab == "sliders":
        smode    = app.slider_mode
        current  = SLIDERS_BY_MODE.get(smode, [])
        _mp2 = dpg.get_mouse_pos()
        mx = _mp2[0] if _mp2 and len(_mp2) >= 2 else 0
        for (_, gname, _lcol, smin, smax) in current:
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
                t  = max(0., min(1., lx / (_sliders_mod.gradient_width(gname) - 1)))
                app.sl_vals[gname] = _sliders_mod._display_value(gname, smin + (smax - smin) * t)
                _apply_from_mode(smode)
                _update_selected_pal_color()
                _update_wheel_contrast()
            if hovered and wheel_delta != 0 and not _lmb_down:
                step = _sliders_mod.slider_step(gname)
                new_val = app.sl_vals.get(gname, smin) + (step if wheel_delta > 0 else -step)
                app.sl_vals[gname] = _sliders_mod._display_value(gname, max(smin, min(smax, new_val)))
                _apply_from_mode(smode)
                _update_selected_pal_color()
                _update_wheel_contrast()
        # All input changes must precede drawing: even unchanged RGB can have
        # a different requested hue/chroma or out-of-gamut component.
        _frame_grad_ctx = _grad_ctx()
        for (_, gname, _lcol, smin, smax) in current:
            if not dpg.does_item_exist(_DL_TAGS[gname]):
                continue
            if (app._last_sl_vals.get(gname) != app.sl_vals.get(gname)
                    or app._last_sl_vals.get('_ctx') != _frame_grad_ctx):
                _draw_grad_slider(gname, smin, smax, _frame_grad_ctx)
        app._last_sl_vals = dict(app.sl_vals)
        app._last_sl_vals['_ctx'] = _frame_grad_ctx


def _update_harmony(vsig):
    """Harmoniaswatchit ja niiden arvot."""
    from .ui_main import align_harmony_values, align_history_export
    align_harmony_values()
    align_history_export()
    mode = app.harmony_mode
    fmt  = app.fmt_mode
    # Harmoniasävyt
    if vsig != app._last_sig or mode != app._last_mode or fmt != app._last_fmt:
        rf, gf, bf = [clamp01(c) for c in app.current_color[:3]]
        h_, l_, s_ = colorsys.rgb_to_hls(rf, gf, bf)
        updates = [("main", rf, gf, bf)]
        angles  = _HARMONY_ANGLES.get(mode, [])
        if mode in ["Tints", "Shades", "Tones"]:
            for i in range(1, 7):
                t = i / 7.0
                if mode == "Tints":
                    nr, ng, nb = colorsys.hls_to_rgb(h_, l_ + (1 - l_) * t, s_)
                elif mode == "Shades":
                    nr, ng, nb = colorsys.hls_to_rgb(h_, l_ * (1 - t), s_)
                else:
                    nr, ng, nb = colorsys.hls_to_rgb(h_, l_, s_ * (1 - t))
                updates.append((f"h{i}", nr, ng, nb))
        elif angles:
            for i, ang in enumerate(angles):
                nh = (h_ + ang / 360.) % 1.
                nr, ng, nb = colorsys.hls_to_rgb(nh, l_, s_)
                updates.append((f"h{i+1}", nr, ng, nb))

        _picker_dragging = (dpg.does_item_exist("picker_wheel")
                            and dpg.is_item_active("picker_wheel"))
        _slider_dragging = (app._dragging is not None)
        _is_dragging     = _picker_dragging or _slider_dragging
        _mode_or_fmt_changed = (mode != app._last_mode or fmt != app._last_fmt)
        _do_full_update = not _is_dragging or _mode_or_fmt_changed
        if _is_dragging and not _mode_or_fmt_changed:
            app._harmony_text_dirty = True

        app.harmony_rgb = {}
        for i, tag in enumerate(HARMONY_TAGS):
            keys = HARMONY_KEYS[tag]
            if i < len(updates):
                rgb2 = tuple(to_int(c) for c in updates[i][1:])
                app.harmony_rgb[tag] = rgb2
                dpg.set_value(keys["rect"], [*list(rgb2), 255])
                if _do_full_update:
                    dpg.configure_item(keys["group"],    show=True)
                    dpg.configure_item(keys["val"],      show=True)
                    dpg.configure_item(keys["copy_btn"], show=True)
                    dpg.set_value(keys["val"], format_value(tag, fmt))
                    from .fonts import bind_value_font
                    bind_value_font(keys["val"])
            else:
                if _do_full_update:
                    dpg.configure_item(keys["group"], show=False)

        app._last_sig  = vsig
        app._last_mode = mode
        app._last_fmt  = fmt
        if _do_full_update:
            app._harmony_text_dirty = False

    elif app._harmony_text_dirty and not (
        (dpg.does_item_exist("picker_wheel") and dpg.is_item_active("picker_wheel"))
        or app._dragging is not None
    ):
        app._harmony_text_dirty = False
        refresh_harmony_values()


_last_notice = None


def _show_operation_errors():
    global _last_notice
    from .i18n import t
    _, until = app._import_status
    code = app._import_status_code
    export_error = t(code) if code in ('export_failed', 'export_open_failed', 'export_aco_limit') and time.perf_counter() < until else ''
    notice = (app._config_error, export_error)
    if notice != _last_notice:
        _last_notice = notice
        text = '\n\n'.join(filter(None, (t(app._config_error) if app._config_error else '', export_error)))
        if text and dpg.does_item_exist('operation_status_window'):
            dpg.set_value('operation_status_text', text)
            dpg.configure_item('operation_status_window', show=True)


def _update_history_and_status(_rmb_clicked):
    """Historia, tuontiviesti, historiarivin poisto ja värimuutoksen kirjaus."""
    _show_operation_errors()
    if app._palette_select_cancel_pending:
        app._palette_select_cancel_pending = False
        if app.palette_select_mode:
            _exit_palette_select_mode()

    # Historia
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

    # Tuontivirhe-viesti
    if dpg.does_item_exist("import_error_text"):
        _imp_msg, _imp_until = app._import_status
        if _imp_msg and time.perf_counter() < _imp_until:
            dpg.set_value("import_error_text", _imp_msg)
            dpg.configure_item("import_error_text", show=True)
        else:
            if _imp_msg:
                app._import_status = ("", 0.0)
                app._import_status_code = None
            dpg.configure_item("import_error_text", show=False)

    # Historia-hiiren oikea nappi
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
            if app.palette_select_mode:
                _exit_palette_select_mode()
            save_config()

    if app.change_time > 0 and (time.perf_counter() - app.change_time) > 0.5:
        if not dpg.is_item_active("picker_wheel"):
            add_to_history(app.current_color)
            app.change_time = 0.
            if app._editing_pal and app._pal_edit_dirty:
                save_config()


def _update_palette_editor(_lmb_down, _lmb_clicked, _lmb_rel, _rmb_clicked, vsig):
    """Paletin muokkaus: oikea nappi, veto, aktiivinen swatch, uudelleenrakennus."""
    _pal_swatch_right_click(_rmb_clicked)
    _pal_swatch_drag(_lmb_down, _lmb_clicked, _lmb_rel)

    if app._editing_pal and app._pal_selected_idx is not None:
        if vsig != app._last_edit_color_sig:
            app._last_edit_color_sig = vsig
            if not app._pal_drag_active:
                _r, _g, _b = vsig
                if not _update_active_pal_swatch(app._pal_selected_idx, _r, _g, _b):
                    _refresh_pal_edit_panel()

    if app._pal_dirty:
        app._pal_dirty = False
        _rebuild_pal_rows()


def update():
    """Yksi UI-päivitys per ruutu. Osat on eroteltu apufunktioiksi (järjestys on tärkeä)."""
    _lmb_down    = dpg.is_mouse_button_down(0)
    _lmb_clicked = dpg.is_mouse_button_clicked(0)
    _lmb_rel     = dpg.is_mouse_button_released(0)
    _rmb_clicked = dpg.is_mouse_button_clicked(1)

    # Tabien leveyden tasaus (ajetaan muutama ruutu käynnistyksen jälkeen)
    if not app._tabs_calibrated:
        app._tabs_calibrated = calibrate_tabs()

    _guard_viewport_size()
    _apply_pending_picker()
    _apply_pending_actions()

    vsig = color_sig(app.current_color)

    _track_viewport_pos(_lmb_down)
    _update_preview(vsig)
    _update_theory_tab(vsig)
    _update_titlebar_icons(_lmb_clicked)
    _drag_titlebar()
    _update_sliders(_lmb_down, _lmb_clicked, vsig)
    _update_harmony(vsig)
    _update_history_and_status(_rmb_clicked)
    _update_palette_editor(_lmb_down, _lmb_clicked, _lmb_rel, _rmb_clicked, vsig)

    # Kielipäivitys
    if app._lang_dirty:
        app._lang_dirty = False
        _refresh_ui_language()


def _refresh_ui_language():
    """Päivittää kaikki käännetyt UI-tekstit uudelle kielelle."""
    from .i18n import t
    app._tabs_calibrated = False
    _items = {
        "always_on_top_chk": ("label", t("always_on_top")),
        "lbl_appearance":    ("value", t("info_appearance")),
        "lbl_theme":         ("value", t("theme_label")),
        "lbl_lang":          ("value", t("language_label")),
        "lbl_pal_edit_hdr":  ("value", t("info_palette_editing")),
        "lbl_pal_edit_help": ("value", t("info_palette_help")),
        "lbl_kb_hdr":        ("value", t("info_keyboard")),
        "lbl_kb_help":       ("value", t("info_keyboard_help")),
        "lbl_drop_hdr":      ("value", t("info_drop")),
        "lbl_drop_help":     ("value", t("info_drop_help")),
        "lbl_version":       ("value", t("info_version")),
        "lbl_author":        ("value", t("info_author")),
        "btn_github":        ("label", t("info_github")),
        "lbl_wcag_title":    ("value", t("wcag_title")),
        "lbl_wcag_fail":     ("value", t("wcag_fail")),
        "lbl_wcag_aalg":     ("value", t("wcag_aa_lg")),
        "lbl_wcag_aa":       ("value", t("wcag_aa")),
        "lbl_wcag_aaa":      ("value", t("wcag_aaa")),
        "lbl_wcag_w":        ("value", t("wcag_w")),
        "lbl_wcag_b":        ("value", t("wcag_b")),
        "lbl_wcag_c":        ("value", t("wcag_c")),
        "wheel_preview_w_txt": ("label", t("sample_text")),
        "tab_wheel":         ("label", tab_label(t("tab_wheel"))),
        "tab_sliders":       ("label", tab_label(t("tab_sliders"))),
        "tab_info":          ("label", tab_label(t("tab_info"))),
        "tab_theory":        ("label", tab_label(t("tab_theory"))),
        "tab_contrast":      ("label", tab_label(t("tab_contrast"))),
        "tab_simulation":    ("label", tab_label(t("tab_simulation"))),
        "btab_harmony":      ("label", tab_label(t("tab_harmony"))),
        "btab_history":      ("label", tab_label(t("tab_history"))),
        "tab_palettes":      ("label", tab_label(t("tab_palettes"))),
        "btn_pick_color":    ("label", pick_color_label()),
        "btn_export_html":   ("label", t("export_html")),
        "btn_save_palette":  ("label", t("save_palette")),
        "btn_new_palette":   ("label", t("save") if app.palette_select_mode else t("new_palette")),
        "btn_cancel_palette":("label", t("cancel")),
        "btn_export_history":("label", t("export_history")),
        "btn_import":        ("label", t("import")),
        "lbl_img_how_many":  ("value", t("img_how_many")),
        "lbl_img_preview":   ("value", t("img_preview")),
        "btn_img_ok":        ("label", t("img_ok")),
        "btn_img_cancel":    ("label", t("cancel")),
        "lbl_drop_import":   ("value", t("drop_import_label")),
        "btn_drop_import":   ("label", t("import")),
        "btn_drop_cancel":   ("label", t("cancel")),
        "btn_add_to_bg":     ("label", t("add_to_bg")),
        "btn_sim_add_to_bg": ("label", t("add_to_bg")),
        "lbl_sim_preview":   ("value", t("sim_preview")),
        "btn_sim_prot":      ("label", t("sim_prot")),
        "btn_sim_deut":      ("label", t("sim_deut")),
        "btn_sim_trit":      ("label", t("sim_trit")),
        "btn_sim_norm":      ("label", t("sim_norm")),
    }
    for tag, (method, val) in _items.items():
        if not dpg.does_item_exist(tag):
            continue
        if method == "label":
            dpg.configure_item(tag, label=val)
        else:
            dpg.set_value(tag, val)
    from .ui_main import refresh_hotkey_settings
    refresh_hotkey_settings()
    from .import_export import refresh_import_warning
    refresh_import_warning()
    from .palette import _refresh_pal_edit_panel
    _refresh_pal_edit_panel()
    app._pal_dirty = True      # palettilistan "(tyhjä)"-tekstit uudelle kielelle

    if dpg.does_item_exist("harmony_combo"):
        from .i18n import get_translated_harmony_modes, to_translated_harmony
        modes = list(get_translated_harmony_modes().keys())
        dpg.configure_item("harmony_combo", items=modes)
        dpg.set_value("harmony_combo", to_translated_harmony(app.harmony_mode))
    app._harmony_text_dirty = True
    
    # Pakota kontrasti-päivitys (W/B/C-napit, APCA ja ehdotus uudelle kielelle)
    _wc_last_sig[0] = None
    _update_wheel_contrast()


# ── Käynnistys ────────────────────────────────────────────────────────────────
def main():
    """Sovelluksen pääfunktio."""
    if acquire_instance_lock() is None:
        ctypes.windll.user32.MessageBoxW(
            0, 'Color Tools is already running. / Color Tools on jo käynnissä.',
            'Color Tools', 0x40)
        return
    cfg = load_config()
    start_pos = apply_config(cfg)

    dpg.create_context()
    dpg.configure_app(manual_callback_management=True)
    ui_constants = build_ui()

    # Eyedropper-overlay
    _pip_hInst = init_pip_overlay()
    if not start_hotkey():
        app._hotkey_notice = "hotkey_unavailable"
    from .ui_main import refresh_hotkey_settings
    refresh_hotkey_settings()

    # Käynnistys
    dpg.create_viewport(
        title='Color Tools',
        width=_VP_W,
        height=_VP_H,
        x_pos=start_pos[0], y_pos=start_pos[1],
        resizable=False,
        decorated=False,
        vsync=False,
    )
    dpg.setup_dearpygui()

    _TM = _themes_mod._THEME_MAP
    if _TM:
        first_theme = _TM.get(app.theme_name, list(_TM.values())[0])
        dpg.bind_theme(first_theme[0])

    # Sateenkaarigradientin palkki
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

    from .ui_main import activate_primary_window
    activate_primary_window()
    if app.always_on_top:
        dpg.set_viewport_always_top(True)
    app._fixed_vp_w = dpg.get_viewport_width()
    # app._fixed_vp_h is already correctly calculated by build_ui()
    dpg.set_viewport_width(app._fixed_vp_w)
    dpg.set_viewport_height(app._fixed_vp_h)
    dpg.show_viewport()

    if dpg.does_item_exist("tb_spacer"):
        _logo_dark_w = ui_constants["logo_dark_w"]
        _vw     = dpg.get_viewport_width()
        _sp_w   = max(2, _vw - _logo_dark_w - 2 * ui_constants["ICON_W"] - 3 * _sc(2) - 17)
        dpg.configure_item("tb_spacer", width=_sp_w)

    set_theme(None, app.theme_name)
    _sync_all_modes()
    _update_wheel_contrast()
    _rebuild_pal_rows()

    # DWM dark titlebar
    for _ in range(20):
        _hwnd = _get_own_hwnd()
        if _hwnd:
            try:
                ctypes.windll.dwmapi.DwmSetWindowAttribute.restype  = ctypes.c_long
                ctypes.windll.dwmapi.DwmSetWindowAttribute.argtypes = [
                    ctypes.c_void_p, ctypes.c_ulong, ctypes.c_void_p, ctypes.c_ulong,
                ]
                _dark = ctypes.c_int(1)
                ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    _hwnd, 20, ctypes.byref(_dark), ctypes.sizeof(_dark))
            except Exception:
                pass
            break
        time.sleep(0.05)

    # Drag-and-drop WndProc
    _WM_DROPFILES = 0x0233
    _GWLP_WNDPROC = -4
    _WNDPROC_TYPE = ctypes.WINFUNCTYPE(
        ctypes.c_ssize_t, ctypes.c_void_p, ctypes.c_uint,
        ctypes.c_size_t, ctypes.c_ssize_t,
    )
    _wndproc_refs = []
    _drop_installed = False

    if _hwnd:
        try:
            _user32  = ctypes.windll.user32
            _shell32 = ctypes.windll.shell32
            try:
                _user32.ChangeWindowMessageFilterEx(_hwnd, _WM_DROPFILES, 1, None)
                _user32.ChangeWindowMessageFilterEx(_hwnd, 0x0049, 1, None)
            except Exception:
                pass
            _shell32.DragAcceptFiles.argtypes = [ctypes.c_void_p, ctypes.c_int]
            _shell32.DragAcceptFiles.restype = None
            _shell32.DragAcceptFiles(_hwnd, True)
            _shell32.DragQueryFileW.restype  = ctypes.c_uint
            _shell32.DragQueryFileW.argtypes = [
                ctypes.c_void_p, ctypes.c_uint, ctypes.c_wchar_p, ctypes.c_uint,
            ]
            _shell32.DragFinish.restype  = None
            _shell32.DragFinish.argtypes = [ctypes.c_void_p]
            if hasattr(_user32, 'GetWindowLongPtrW'):
                _GetWindowLongPtr = _user32.GetWindowLongPtrW
                _SetWindowLongPtr = _user32.SetWindowLongPtrW
            else:
                _GetWindowLongPtr = _user32.GetWindowLongW
                _SetWindowLongPtr = _user32.SetWindowLongW
            _GetWindowLongPtr.restype  = ctypes.c_ssize_t
            _GetWindowLongPtr.argtypes = [ctypes.c_void_p, ctypes.c_int]
            _SetWindowLongPtr.restype  = ctypes.c_ssize_t
            _SetWindowLongPtr.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_ssize_t]
            _user32.CallWindowProcW.restype    = ctypes.c_ssize_t
            _user32.CallWindowProcW.argtypes   = [
                ctypes.c_ssize_t, ctypes.c_void_p, ctypes.c_uint,
                ctypes.c_size_t, ctypes.c_ssize_t,
            ]
            _old_wndproc_ptr = _GetWindowLongPtr(_hwnd, _GWLP_WNDPROC)
            def _drop_wndproc(hwnd_w, msg, wp, lp):
                if msg == _WM_DROPFILES:
                    try:
                        _n   = _shell32.DragQueryFileW(wp, 0, None, 0)   # pituus (ei 260 merkin rajaa)
                        _buf = ctypes.create_unicode_buffer(_n + 1)
                        _shell32.DragQueryFileW(wp, 0, _buf, _n + 1)
                        _dpath = _buf.value
                        if _dpath:
                            app._pending_drop_path = _dpath
                    except Exception:
                        _write_log("File drop failed:\n" + traceback.format_exc())
                    finally:
                        _shell32.DragFinish(wp)
                    return 0
                return _user32.CallWindowProcW(_old_wndproc_ptr, hwnd_w, msg, wp, lp)
            _new_proc = _WNDPROC_TYPE(_drop_wndproc)
            _wndproc_refs.append(_new_proc)
            _new_proc_addr = ctypes.cast(_new_proc, ctypes.c_void_p).value
            _SetWindowLongPtr(_hwnd, _GWLP_WNDPROC, _new_proc_addr)
            _drop_installed = True
        except Exception:
            _write_log("drop wndproc setup failed:\n" + traceback.format_exc())

    # Windows timer-tarkkuus
    try:
        ctypes.windll.winmm.timeBeginPeriod(1)
        _timer_period_set = True
    except Exception:
        _timer_period_set = False

    _TARGET_FPS   = 30
    _FRAME_BUDGET = 1.0 / _TARGET_FPS

    _err_count = 0
    try:
        while dpg.is_dearpygui_running():
            _frame_start = time.perf_counter()
            try:
                for job in dpg.get_callback_queue() or ():
                    try:
                        dpg.run_callbacks([job])
                    except Exception:
                        _write_log("UI callback failed:\n" + traceback.format_exc())
                update()
                _err_count = 0
            except Exception:
                # Yksittäinen virhe ei kaada sovellusta; toistuva (60 ruutua) kaataa siististi
                _err_count += 1
                if _err_count == 1:
                    _write_log("update() error:\n" + traceback.format_exc())
                if _err_count >= 60:
                    raise
            dpg.render_dearpygui_frame()
            _elapsed = time.perf_counter() - _frame_start
            _sleep   = _FRAME_BUDGET - _elapsed
            if _sleep > 0.001:
                _pumping_sleep(_sleep)
    finally:
        # Sammutus ajetaan aina: tallennus ja resurssien vapautus myös virhetilanteessa
        shutdown_io()
        try:
            save_config()
        except Exception:
            pass
        try:
            stop_hotkey()
        except Exception:
            pass

        app.pipette_active = False
        _pip_tid_shutdown = getattr(app, '_pipette_thread_tid', 0)
        if _pip_tid_shutdown:
            try:
                ctypes.windll.user32.PostThreadMessageW(_pip_tid_shutdown, 0x0012, 0, 0)
            except Exception:
                pass

        if _pip_hwnd[0]:
            try:
                _pip_u32.DestroyWindow(_pip_hwnd[0])
                _pip_u32.UnregisterClassW(_pip_cls[0], _pip_hInst)
            except Exception:
                pass
            _pip_hwnd[0] = None
        try:
            _pip_free_gdi()
        except Exception:
            pass

        if _timer_period_set:
            try:
                ctypes.windll.winmm.timeEndPeriod(1)
            except Exception:
                pass

        try:
            if _drop_installed and _GetWindowLongPtr(_hwnd, _GWLP_WNDPROC) == _new_proc_addr:
                _shell32.DragAcceptFiles(_hwnd, False)
                _SetWindowLongPtr(_hwnd, _GWLP_WNDPROC, _old_wndproc_ptr)
        finally:
            dpg.destroy_context()
