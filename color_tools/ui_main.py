"""
ui_main.py — Koko DPG-käyttöliittymän rakennus.

build_ui() kutsutaan dpg.create_context():n ja create_themes():n jälkeen.
"""

import traceback
import ctypes
import os

import dearpygui.dearpygui as dpg
from . import fonts

from .bootstrap import _sc, _DPI_SCALE, _SCRIPT_DIR, safe_copy
from .state import app, save_config
from .i18n import t
from .color_math import (
    FORMAT_OPTIONS,
)
from .sliders import (
    SLIDER_MODES, SLIDERS_BY_MODE, init_layout_constants, _make_grad_slider,
)
from .harmony import (
    HARMONY_TAGS, copy_preview,
    copy_harmony_val, hex_input_cb,
    toggle_on_top, picker_cb, set_current_as_sample_background,
)
from .palette import (
    MAX_PAL_ROWS, save_harmony_as_palette,
    toggle_palette_select_mode, _exit_palette_select_mode,
    history_click_cb, init_palette_constants, change_palette_page, _PALETTES_PER_PAGE,
)
from .import_export import (
    _import_open_dialog, _close_drop_preview,
    _image_import_slider_cb, _image_import_ok,
    _image_import_cancel, export_palette, export_history,
)
from . import themes as _themes_mod
from .themes import (
    create_themes,
    set_theme, slider_mode_cb, harmony_combo_cb,
    _redraw_harmony_combo_icon, fmt_combo_cb, _on_top_tab,
)
from .eyedropper import start_pipette, change_hotkey
from .hotkey_config import MODIFIERS, KEYS, DEFAULT_HOTKEY, hotkey_label

def pick_color_label():
    label = t('pick_color')
    if app.pipette_hotkey['enabled']:
        label += f" ({hotkey_label(app.pipette_hotkey)})"
    return label


def refresh_hotkey_settings():
    for tag, key in (('hotkey_enabled', 'hotkey_enabled'), ('hotkey_apply', 'hotkey_apply'),
                     ('hotkey_reset', 'hotkey_reset')):
        if dpg.does_item_exist(tag):
            dpg.configure_item(tag, label=t(key))
    if dpg.does_item_exist('hotkey_status'):
        dpg.set_value('hotkey_status', t(app._hotkey_notice) if app._hotkey_notice else '')
    if dpg.does_item_exist('btn_pick_color'):
        dpg.configure_item('btn_pick_color', label=pick_color_label())


def apply_hotkey_settings(sender=None, value=None, user_data=None):
    settings = dict(DEFAULT_HOTKEY) if user_data == 'reset' else {
        'enabled': dpg.get_value('hotkey_enabled'),
        'modifiers': dpg.get_value('hotkey_modifiers'),
        'key': dpg.get_value('hotkey_key'),
    }
    if change_hotkey(settings):
        save_config()
    for tag, key in (('hotkey_enabled', 'enabled'), ('hotkey_modifiers', 'modifiers'), ('hotkey_key', 'key')):
        dpg.set_value(tag, app.pipette_hotkey[key])
    refresh_hotkey_settings()


# ── Layout-vakiot (lasketaan DPI:stä) ─────────────────────────────────────────
W = _sc(268)
BW = _sc(130)
HSQ = _sc(26)
PAL_SW = _sc(24)
GRAD_W = W - _sc(42) - _sc(14) - 16
THUMB_H = max(_sc(19), _sc(fonts.UI_FONT_SIZE) + 6)
GRAD_H = THUMB_H  # Match the visible gradient to its numeric field.
# Swatch, copy button and two ImGui 8px gaps; scrollbar space is dynamic.
VALW = W - THUMB_H - _sc(20) - 16
_ESW = _sc(28)
_ECOLS = 8
_BW2 = (W - _sc(8)) // 2 - _sc(8)
_VP_W = int(283 * _DPI_SCALE)
# Two taller action rows and the final checkbox need room with the larger UI font.
_VP_H = _sc(898 + 3 * max(0, fonts.UI_FONT_SIZE - 13))
_TB_H = _sc(34)
_ICON_W = _sc(28)
_IC = _sc(6)
_THICK_ICON = _sc(2)

# Osioiden korkeudet (laajennettu / tiivistetty) ja viewportin lisätila
_COLLAPSED_H = _sc(28)
_TAB_AREA_H  = _sc(240)
_UTILITY_H   = _sc(273)
_TOP_VIEW_H  = _sc(280)
_VP_H_PAD    = 47

def _temp_gradient_color(x):
    """Lämpötilamittarin väri kohdassa x (0 = viileä sininen, 0.5 = neutraali harmaa, 1 = lämmin oranssi)."""
    cool, mid, warm = (70, 130, 235), (160, 160, 160), (240, 130, 50)
    a, b, f = (cool, mid, x * 2) if x < 0.5 else (mid, warm, (x - 0.5) * 2)
    return [int(round(a[i] + (b[i] - a[i]) * f)) for i in range(3)] + [255]


# Tiivistys-/laajennusnapin nuolet. Jos mitään symbolifonttia ei löydy
# (Tahoma / Segoe UI / Arial / Calibri puuttuu), build_ui vaihtaa ne
# ASCII-merkkeihin, koska DPG:n oletusfontissa ei ole ▲/▼-merkkejä.
_ARROW_UP, _ARROW_DN = "▲", "▼"

# Stock DearPyGui tabs use measured whitespace padding after the first frame.
_TRI_FONT = None
_TAB_BARS = {
    "main_tab_bar": ("tab_wheel", "tab_sliders", "tab_info"),
    "bottom_tab_bar": ("btab_harmony", "btab_history", "tab_palettes"),
    "utility_tab_bar": ("tab_theory", "tab_contrast", "tab_simulation"),
}


def tab_label(text):
    """Start unpadded; font metrics become available after rendering."""
    return text.strip()


def calibrate_tabs():
    """Approximate equal widths using spaces, reserving the collapse button."""
    space = dpg.get_text_size(' ', font=fonts.hex_font or 0)
    if not space or space[0] <= 0:
        return False
    for tabs in _TAB_BARS.values():
        if any(not dpg.does_item_exist(tag) or
               dpg.get_item_state(tag).get('rect_size', (0, 0))[0] <= 0 for tag in tabs):
            return False
    # Every theme uses FramePadding.x=4; ImGui's tab gap is 4 pixels.
    # Stock tab buttons do not expose rectangle measurements via DPG.
    arrow_width = max(dpg.get_text_size(arrow, font=_TRI_FONT or 0)[0]
                      for arrow in (_ARROW_UP, _ARROW_DN)) + 8
    target = (W - arrow_width - 3 * 4) / 3
    for tabs in _TAB_BARS.values():
        for tag in tabs:
            label = dpg.get_item_label(tag)
            text = label.strip()
            text_width = dpg.get_text_size(text, font=fonts.hex_font or 0)[0]
            count = max(0, int((target - 8 - text_width) / space[0]))
            padded = ' ' * (count // 2) + text + ' ' * (count - count // 2)
            dpg.set_item_label(tag, padded)
    return True


def align_harmony_values():
    """Reserve scrollbar space only when the harmony list actually scrolls."""
    if not dpg.does_item_exist("harmony_rows_scroll"):
        return
    width = VALW - (_sc(16) if dpg.get_y_scroll_max("harmony_rows_scroll") > 0 else 0)
    for tag in HARMONY_TAGS:
        field = f"val_{tag}"
        if dpg.does_item_exist(field) and dpg.get_item_width(field) != width:
            dpg.configure_item(field, width=width)


def align_history_export():
    """Keep both tabs' actions at the harmony row's rendered position."""
    if not dpg.does_item_exist("harmony_actions"):
        return
    if dpg.get_item_state("harmony_actions").get("rect_size", (0, 0))[1] <= 0:
        return
    # Both tabs use the same child window coordinate system. The harmony
    # position remains available when history is shown, even with fewer rows.
    position = dpg.get_item_pos("harmony_actions")
    if list(dpg.get_item_pos("btn_export_history")) != list(position):
        dpg.set_item_pos("btn_export_history", position)


# Logo-koot (asetetaan _load_textures:n jälkeen)
_logo_dark_w = 0
_logo_dark_h = 0
_logo_light_w = 0
_logo_light_h = 0
_font_large = None

# ── Win32 ikkunan minimointi ──────────────────────────────────────────────────


def _minimize_win():
    from .bootstrap import _get_own_hwnd
    try:
        hwnd = _get_own_hwnd()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 6)
    except Exception:
        from .bootstrap import _write_log
        _write_log("_minimize_win error:\n" + traceback.format_exc())


ctypes.windll.user32.MoveWindow.restype = ctypes.c_bool
ctypes.windll.user32.MoveWindow.argtypes = [
    ctypes.c_void_p, ctypes.c_int, ctypes.c_int,
    ctypes.c_int, ctypes.c_int, ctypes.c_bool,
]

# ── Ikkunan koko / Välilehtien piilotus ─────────────────────────────────────────


def apply_collapse_state():
    h1 = _COLLAPSED_H if app.tab_area_collapsed else _TAB_AREA_H
    h2 = _COLLAPSED_H if app.utility_collapsed else _UTILITY_H
    h3 = _COLLAPSED_H if app.top_view_collapsed else _TOP_VIEW_H
    if dpg.does_item_exist("tab_area"):
        dpg.configure_item("tab_area", height=h1)
    if dpg.does_item_exist("utility_section"):
        dpg.configure_item("utility_section", height=h2)
    if dpg.does_item_exist("top_view_cw"):
        dpg.configure_item("top_view_cw", height=h3)
    if dpg.does_item_exist("btn_collapse_tab_area"):
        dpg.configure_item("btn_collapse_tab_area",
                           label=_ARROW_DN if app.tab_area_collapsed else _ARROW_UP)
    if dpg.does_item_exist("btn_collapse_utility"):
        dpg.configure_item("btn_collapse_utility",
                           label=_ARROW_DN if app.utility_collapsed else _ARROW_UP)
    if dpg.does_item_exist("btn_collapse_top_view"):
        dpg.configure_item("btn_collapse_top_view",
                           label=_ARROW_DN if app.top_view_collapsed else _ARROW_UP)

    base_h = _VP_H
    if app.tab_area_collapsed:
        base_h -= (_TAB_AREA_H - _COLLAPSED_H)
    if app.utility_collapsed:
        base_h -= (_UTILITY_H - _COLLAPSED_H)
    if app.top_view_collapsed:
        base_h -= (_TOP_VIEW_H - _COLLAPSED_H)

    app._fixed_vp_h = base_h + _VP_H_PAD
    try:
        dpg.set_viewport_height(app._fixed_vp_h)
    except Exception:
        pass


def toggle_tab_area(s, a, u):
    app.tab_area_collapsed = not app.tab_area_collapsed
    apply_collapse_state()
    save_config()


def toggle_utility(s, a, u):
    app.utility_collapsed = not app.utility_collapsed
    apply_collapse_state()
    save_config()

def toggle_top_view(s, a, u):
    app.top_view_collapsed = not app.top_view_collapsed
    apply_collapse_state()
    save_config()

# ── Päärakennusfunktio ────────────────────────────────────────────────────────


def build_ui():
    """Rakentaa koko DPG-käyttöliittymän. Kutsuttava dpg.create_context():n jälkeen."""
    global _logo_dark_w, _logo_dark_h, _logo_light_w, _logo_light_h, _font_large

    # -- Alustetaan layout-vakiot liukusäätimille --
    init_layout_constants(GRAD_W, GRAD_H, THUMB_H)

    # Hiiripyörä-handler
    with dpg.handler_registry():
        dpg.add_mouse_wheel_handler(
            callback=lambda s, a, u: setattr(app, '_mouse_wheel_delta',
                                              app._mouse_wheel_delta + a)
        )

    # -- Logotekstuurit --
    with dpg.texture_registry():
        for _tex_tag, _fname, _wattr in [
            ("logo_dark",  "colortoolsd.png", "dark"),
            ("logo_light", "colortoolsw.png", "light"),
        ]:
            _path = os.path.join(_SCRIPT_DIR, _fname)
            try:
                _lw, _lh, _, _ldata = dpg.load_image(_path)
                dpg.add_static_texture(_lw, _lh, _ldata, tag=_tex_tag)
                _scale = min(1.0, (_TB_H - _sc(4)) / max(1, _lh))
                _rw = max(1, int(_lw * _scale))
                _rh = max(1, int(_lh * _scale))
                if _wattr == "dark":
                    _logo_dark_w = _rw;  _logo_dark_h = _rh
                else:
                    _logo_light_w = _rw;  _logo_light_h = _rh
            except Exception:
                dpg.add_static_texture(1, 1, [0, 0, 0, 0], tag=_tex_tag)

    if _logo_light_w == 0 and _logo_dark_w > 0:
        _logo_light_w = _logo_dark_w
        _logo_light_h = _logo_dark_h

    # -- Fontit --
    # Fontin on sisällettävä ▲ ja ▼ (U+25B2, U+25BC). Ensimmäinen löytyvä käytetään.
    _FONTS_DIR = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    _FONT_PATHS = [os.path.join(_FONTS_DIR, _fn) for _fn in ("segoeui.ttf", "tahoma.ttf", "arial.ttf", "calibri.ttf")]
    # Rasteroi suoraan lopulliseen pikselikokoon: älä venytä fonttitekstuuria.
    dpg.set_global_font_scale(1.0)
    fonts.hex_font = fonts.value_font = None
    with dpg.font_registry():
        _font_triangle = None
        for _fp in _FONT_PATHS:
            if os.path.exists(_fp):
                with dpg.font(_fp, _sc(fonts.UI_FONT_SIZE), pixel_snapH=True) as _font_main:
                    dpg.add_font_range_hint(dpg.mvFontRangeHint_Default)
                    # Dashes and ellipsis used by translated UI strings.
                    dpg.add_font_chars([0x2013, 0x2014, 0x2026, 0x25B2, 0x25BC])
                dpg.bind_font(_font_main)
                fonts.hex_font = _font_main
                _font_large = dpg.add_font(_fp, _sc(22), pixel_snapH=True)
                # Tab layout measures symbols with the main font too.
                _font_triangle = _font_main
                break
        _value_path = os.path.join(_FONTS_DIR, "consola.ttf")
        if os.path.exists(_value_path):
            with dpg.font(_value_path, _sc(fonts.VALUE_FONT_SIZE), pixel_snapH=True) as fonts.value_font:
                dpg.add_font_range_hint(dpg.mvFontRangeHint_Default)
                dpg.add_font_chars([0x00B0, 0x2212])

    # Ei symbolifonttia -> ASCII-nuolet (muuten ▲ näkyisi '?'-merkkinä)
    global _ARROW_UP, _ARROW_DN, _TRI_FONT
    _TRI_FONT = _font_triangle
    if _font_triangle is None:
        _ARROW_UP, _ARROW_DN = "^", "v"

    # -- Teemat --
    create_themes()
    fonts.create_value_theme(THUMB_H, _sc(fonts.VALUE_FONT_SIZE if fonts.value_font else fonts.UI_FONT_SIZE))
    # Päivitä palette-moduulin viitteet teemoihin (nyt luotu)
    from . import themes as _t
    init_palette_constants(
        PAL_SW, MAX_PAL_ROWS, _ECOLS, _ESW, _BW2, THUMB_H,
        _t._nospc_theme, _t._close_btn_theme,
    )

    # ── Pääikkuna ─────────────────────────────────────────────────────────────
    # Fixed layout: wheel events must never scroll the entire application.
    # Scrollable child panels retain their own wheel handling.
    with dpg.window(tag="PrimaryWindow", no_scrollbar=True, no_scroll_with_mouse=True):

        # Titlebar
        with dpg.child_window(
            tag="titlebar_cw", width=_VP_W, height=_TB_H,
            border=False, no_scrollbar=True, no_scroll_with_mouse=True,
        ):
            dpg.bind_item_theme("titlebar_cw", _t._titlebar_dark_theme)
            with dpg.group(tag="titlebar", horizontal=True):
                if _logo_dark_w > 0:
                    _tb_pad_y = max(0, (_TB_H - _logo_dark_h) // 2)
                    if _tb_pad_y > 0:
                        dpg.add_spacer(height=_tb_pad_y)
                    dpg.add_image(
                        "logo_dark", tag="tb_logo",
                        width=_logo_dark_w, height=_logo_dark_h,
                    )
                dpg.add_spacer(width=2, tag="tb_spacer")
                dpg.add_drawlist(width=_ICON_W, height=_TB_H, tag="dl_tb_min")
                dpg.add_drawlist(width=_ICON_W, height=_TB_H,
                                 tag="dl_tb_close")
        dpg.add_spacer(height=2)

        # ── Yläosan välilehdet: Pyörä / Liukusäätimet / Info ──────────────────

        def _main_tab_cb(s, active_tab, u):
            tag = active_tab
            if isinstance(tag, int):
                tag = dpg.get_item_alias(tag) or tag
            
            if tag == "tab_wheel":
                _on_top_tab("wheel")
            elif tag == "tab_sliders":
                _on_top_tab("sliders")
            elif tag == "tab_info":
                _on_top_tab("info")
        with dpg.child_window(
            width=0, height=_TOP_VIEW_H, border=False, no_scrollbar=True,
            no_scroll_with_mouse=True, tag="top_view_cw",
        ):
            dpg.bind_item_theme("top_view_cw", _t._nopad_cw_theme)
            with dpg.tab_bar(tag="main_tab_bar", callback=_main_tab_cb):
                with dpg.tab(label=tab_label(t("tab_wheel")), tag="tab_wheel"):
                    with dpg.group(tag="grp_wheel"):
                        dpg.add_spacer(height=_sc(8))
                        _ir, _ig, _ib = [int(round(c * 255))
                                             for c in app.current_color[:3]]
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

                with dpg.tab(label=tab_label(t("tab_sliders")), tag="tab_sliders"):
                    with dpg.group(tag="grp_sliders_all"):
                        dpg.add_spacer(height=_sc(8))
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
                                    for (_lbl, gname, lcol, _smin, _smax) in sliders:
                                        _make_grad_slider(_lbl, lcol, gname)

            # ── Info-välilehti ─────────────────────────────────────────────
                with dpg.tab(label=tab_label(t("tab_info")), tag="tab_info"):
                    with dpg.child_window(
                        width=W, height=_sc(252), border=False, no_scrollbar=False,
                        tag="info_scroll",
                    ):
                        def _change_language(s, v, u):
                            from .i18n import set_language
                            set_language("fi" if v == "Suomi" else "en")
                            save_config()
                        _themes_mod.add_text(t("info_appearance"), tag="lbl_appearance", color=[160, 200, 255])
                        dpg.add_separator()
                        dpg.add_spacer(height=_sc(3))
                        with dpg.group(horizontal=True):
                            _themes_mod.add_text(t("theme_label"), tag="lbl_theme", color=[180, 180, 180])
                            dpg.add_combo(
                                _themes_mod.THEME_NAMES, tag="theme_combo",
                                default_value=app.theme_name,
                                width=_sc(120), callback=set_theme,
                            )
                        dpg.add_spacer(height=_sc(4))
                        with dpg.group(horizontal=True):
                            _themes_mod.add_text(t("language_label"), tag="lbl_lang", color=[180, 180, 180])
                            dpg.add_combo(
                                ["English", "Suomi"], tag="language_combo",
                                default_value="Suomi" if app.language == "fi" else "English",
                                width=_sc(120), callback=_change_language,
                            )
                        dpg.add_spacer(height=_sc(8))
                        _themes_mod.add_text(t("info_keyboard"), tag="lbl_kb_hdr", color=[160, 200, 255])
                        dpg.add_separator()
                        dpg.add_spacer(height=_sc(3))
                        _themes_mod.add_text(t("info_keyboard_help"), tag="lbl_kb_help", color=[180, 180, 180], wrap=W - _sc(16))
                        dpg.add_checkbox(label=t('hotkey_enabled'), tag='hotkey_enabled', default_value=app.pipette_hotkey['enabled'])
                        with dpg.group(horizontal=True):
                            dpg.add_combo(list(MODIFIERS), tag='hotkey_modifiers', width=_sc(132), default_value=app.pipette_hotkey['modifiers'])
                            dpg.add_combo(list(KEYS), tag='hotkey_key', width=_sc(72), default_value=app.pipette_hotkey['key'])
                        with dpg.group(horizontal=True):
                            dpg.add_button(label=t('hotkey_apply'), tag='hotkey_apply', callback=apply_hotkey_settings)
                            dpg.add_button(label=t('hotkey_reset'), tag='hotkey_reset', callback=apply_hotkey_settings, user_data='reset')
                        _themes_mod.add_text('', tag='hotkey_status', wrap=W - _sc(16))
                        dpg.add_spacer(height=_sc(8))
                        _themes_mod.add_text(t("info_palette_editing"), tag="lbl_pal_edit_hdr", color=[160, 200, 255])
                        dpg.add_separator()
                        dpg.add_spacer(height=_sc(3))
                        _themes_mod.add_text(t("info_palette_help"), tag="lbl_pal_edit_help",
                                     color=[180, 180, 180], wrap=W - _sc(16))
                        dpg.add_spacer(height=_sc(8))
                        _themes_mod.add_text(t("info_drop"), tag="lbl_drop_hdr", color=[160, 200, 255])
                        dpg.add_separator()
                        dpg.add_spacer(height=_sc(3))
                        _themes_mod.add_text(t("info_drop_help"), tag="lbl_drop_help",
                                     color=[180, 180, 180], wrap=W - _sc(16))
                        dpg.add_spacer(height=_sc(12))
                        _themes_mod.add_text(t("info_version"), tag="lbl_version", color=[220, 220, 255])
                        _themes_mod.add_text(t("info_author"), tag="lbl_author", color=[160, 160, 160])
                        dpg.add_spacer(height=_sc(4))
                        def _open_github(s, a, u):
                            try:
                                os.startfile("https://github.com/sort0m/Color-tools")
                            except Exception:
                                pass
                        dpg.add_button(
                            label=t("info_github"), tag="btn_github",
                            callback=_open_github,
                            width=W - _sc(16), height=THUMB_H,
                        )
                dpg.add_tab_button(label=_ARROW_UP, trailing=True, tag="btn_collapse_top_view", callback=toggle_top_view, no_tooltip=True)

        dpg.add_spacer(height=3)

        # Värien esikatselupalkki
        _PREV_SW = BW
        _PREV_HW = W - BW - 16 - _sc(20)
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
            label="C", tag="preview_copy", width=_sc(20), height=THUMB_H,
            callback=lambda s, a, u: safe_copy(dpg.get_value("preview_hex")),
            )
        dpg.add_button(
            label=pick_color_label(), tag="btn_pick_color",
            width=W, height=THUMB_H,
            callback=lambda s, a, u: start_pipette(),
        )
        dpg.add_spacer(height=2)

        with dpg.child_window(
            width=0, height=_TAB_AREA_H, border=False, no_scrollbar=True,
            no_scroll_with_mouse=True, tag="tab_area",
        ):
            with dpg.tab_bar(tag="bottom_tab_bar"):
                with dpg.tab(label=tab_label(t("tab_harmony")), tag="btab_harmony"):
                    dpg.add_spacer(height=1)
                    with dpg.group(horizontal=True):
                        with dpg.group():
                            from .i18n import get_translated_harmony_modes, to_translated_harmony
                            dpg.add_combo(
                                list(get_translated_harmony_modes().keys()),
                                tag="harmony_combo", default_value=to_translated_harmony(app.harmony_mode),
                                width=_sc(212), callback=harmony_combo_cb,
                            )
                            dpg.add_combo(
                                FORMAT_OPTIONS, tag="fmt_combo",
                                default_value="HEX", width=_sc(212),
                                callback=fmt_combo_cb,
                            )
                        dpg.add_drawlist(tag="hm_icon_dl", width=W - _sc(212) - 8,
                                         height=2 * THUMB_H + 4)
                    _redraw_harmony_combo_icon("Triadic")
                    dpg.add_spacer(height=1)
                    with dpg.child_window(
                        width=W, height=_sc(118),
                        border=False, no_scrollbar=False,
                        tag="harmony_rows_scroll",
                    ):
                        for tag in HARMONY_TAGS:
                            with dpg.group(tag=f"group_{tag}"):
                                with dpg.group(horizontal=True):
                                    dpg.add_color_button(tag=f"rect_{tag}", width=THUMB_H, height=THUMB_H)
                                    dpg.add_input_text(tag=f"val_{tag}", width=VALW, readonly=True)
                                    dpg.add_button(
                                        label="C", width=_sc(20), height=THUMB_H,
                                        tag=f"copy_btn_{tag}",
                                        callback=copy_harmony_val, user_data=tag,
                                    )
                    dpg.add_spacer(height=0)
                    with dpg.group(horizontal=True, tag="harmony_actions"):
                        dpg.add_button(
                            label=t("export_html"),  callback=export_palette, tag="btn_export_html",
                            width=(W - 8) // 2 - 1, height=THUMB_H,
                        )
                        dpg.add_button(
                            label=t("save_palette"), tag="btn_save_palette",
                            callback=save_harmony_as_palette,
                            width=(W - 8) // 2 + 1, height=THUMB_H,
                        )

        # -- Historia --------------------------------------------------
                with dpg.tab(label=tab_label(t("tab_history")), tag="btab_history"):
                    with dpg.group(horizontal=True):
                        dpg.add_button(
                            tag="btn_new_palette", label=t("new_palette"),
                            callback=toggle_palette_select_mode,
                            width=_sc(90), height=THUMB_H,
                        )
                        dpg.add_button(
                            tag="btn_cancel_palette", label=t("cancel"),
                            callback=_exit_palette_select_mode,
                            width=_sc(60), height=THUMB_H,
                            show=False,
                        )
                    dpg.bind_item_theme("btn_new_palette", None)
                    dpg.add_spacer(height=1)
                    with dpg.group() as _history_grid:
                        dpg.bind_item_theme(_history_grid, _t._nospc_theme)
                        for row in range(6):
                            with dpg.group(horizontal=True) as row_group:
                                dpg.bind_item_theme(row_group, _t._nospc_theme)
                                for i in range(row * 10, row * 10 + 10):
                                    # Fill the content width without accumulating rounding gaps.
                                    col = i % 10
                                    swatch_width = (col + 1) * W // 10 - col * W // 10
                                    dpg.add_color_button(
                                        tag=f"hist_{i}", width=swatch_width, height=HSQ,
                                        show=False,
                                        callback=history_click_cb, user_data=i,
                                    )
                                    dpg.bind_item_theme(f"hist_{i}", _t._nospc_theme)
                    dpg.add_spacer(height=4)
                    dpg.add_button(
                        label=t("export_history"), tag="btn_export_history",
                        callback=export_history, width=W, height=THUMB_H,
                    )

        # -- Paletit ---------------------------------------------------
                with dpg.tab(label=tab_label(t("tab_palettes")), tag="tab_palettes"):
                    _themes_mod.add_text("", tag="import_error_text", color=[220, 80, 80], show=False)
                    PAL_SCROLL_H = -_sc(6)  # Fit remaining tab height, keeping the bottom border visible.
                    with dpg.child_window(
                        width=W, height=PAL_SCROLL_H,
                        border=True, no_scrollbar=False,
                        tag="pal_scroll",
                    ):
                        with dpg.table(header_row=False, no_pad_outerX=True,
                                       no_pad_innerX=True, policy=dpg.mvTable_SizingStretchProp):
                            dpg.add_table_column(width_stretch=True)
                            dpg.add_table_column(width_fixed=True, init_width_or_weight=_sc(70))
                            with dpg.table_row():
                                with dpg.group(horizontal=True, tag='pal_page_controls', show=False):
                                    dpg.add_button(label='<', tag='pal_page_prev', height=THUMB_H,
                                                   callback=change_palette_page, user_data=-1)
                                    _themes_mod.add_text('1 / 1', tag='pal_page_label')
                                    dpg.add_button(label='>', tag='pal_page_next', height=THUMB_H,
                                                   callback=change_palette_page, user_data=1)
                                with dpg.group(tag="pal_import_buttons"):
                                    dpg.add_button(
                                        label=t("import"), tag="btn_import",
                                        callback=_import_open_dialog,
                                        width=-1, height=THUMB_H,
                                    )
                        for pi in range(_PALETTES_PER_PAGE):
                            with dpg.group(tag=f"pal_row_{pi}", show=False):
                                pass

                    with dpg.child_window(
                        width=W, height=PAL_SCROLL_H,
                        border=True, no_scrollbar=False,
                        tag="pal_edit_win", show=False,
                    ):
                        with dpg.group(tag="pal_edit_panel"):
                            pass

                    with dpg.child_window(
                        width=W, height=PAL_SCROLL_H,
                        border=True, no_scrollbar=True,
                        tag="img_import_win", show=False,
                    ):
                        _themes_mod.add_text(t("img_how_many"), tag="lbl_img_how_many")
                        dpg.add_spacer(height=3)
                        dpg.add_slider_int(
                            tag="img_import_slider",
                            default_value=8, min_value=2, max_value=30,
                            callback=_image_import_slider_cb,
                            width=W - _sc(20),
                        )
                        dpg.add_spacer(height=5)
                        _themes_mod.add_text(t("img_preview"), tag="lbl_img_preview")
                        with dpg.group(tag="img_import_preview"):
                            pass
                        dpg.add_spacer(height=6)
                        _btn_w2 = (W - _sc(8)) // 2
                        with dpg.group(horizontal=True):
                            dpg.add_button(label=t("img_ok"), width=_btn_w2, height=THUMB_H, tag="btn_img_ok",
                                           callback=_image_import_ok)
                            dpg.add_button(label=t("cancel"), width=_btn_w2, height=THUMB_H, tag="btn_img_cancel",
                                           callback=_image_import_cancel)

                    with dpg.child_window(
                        width=W, height=PAL_SCROLL_H,
                        border=True, no_scrollbar=True,
                        tag="drop_preview_win", show=False,
                    ):
                        _themes_mod.add_text("", tag="drop_preview_label",
                                     color=[180, 210, 255], wrap=W - _sc(16))
                        dpg.add_spacer(height=_sc(4))
                        _themes_mod.add_text("", tag="drop_preview_warning",
                                             color=[255, 200, 100], wrap=W - _sc(16), show=False)
                        with dpg.group(tag="drop_preview_colors"):
                            pass
                        dpg.add_spacer(height=_sc(6))
                        _themes_mod.add_text(t("drop_import_label"), tag="lbl_drop_import", color=[200, 200, 200])
                        dpg.add_spacer(height=_sc(4))
                        _btn_dp = (W - _sc(8)) // 2
                        with dpg.group(horizontal=True):
                            def _drop_preview_ok(s, a, u):
                                colors_to_commit = getattr(app, '_drop_preview_colors_pending', [])
                                _close_drop_preview()
                                app._drop_preview_colors_pending = []
                                if colors_to_commit:
                                    from .import_export import _commit_image_palette
                                    _commit_image_palette(colors_to_commit)
                            def _drop_preview_cancel(s, a, u):
                                app._drop_preview_colors_pending = []
                                _close_drop_preview()
                            dpg.add_button(label=t("import"), width=_btn_dp, height=THUMB_H, tag="btn_drop_import",
                                           callback=_drop_preview_ok)
                            dpg.add_button(label=t("cancel"), width=_btn_dp, height=THUMB_H, tag="btn_drop_cancel",
                                           callback=_drop_preview_cancel)
                dpg.add_tab_button(label=_ARROW_UP, trailing=True, tag="btn_collapse_tab_area", callback=toggle_tab_area, no_tooltip=True)

        # -- Teoria + Kontrasti -------------------------------------------------
        with dpg.child_window(
            width=0, height=_UTILITY_H, border=False, no_scrollbar=True,
            no_scroll_with_mouse=True, tag="utility_section",
        ):
            dpg.bind_item_theme("utility_section", _t._nopad_cw_theme)
            with dpg.tab_bar(tag="utility_tab_bar"):
                with dpg.tab(label=tab_label(t("tab_theory")), tag="tab_theory"):
                    dpg.add_spacer(height=2)
                    with dpg.child_window(
                        width=W, height=_sc(238), border=False, no_scrollbar=False,
                        tag="theory_scroll"
                    ):
                        with dpg.group(horizontal=True):
                            _themes_mod.add_text("", tag="theory_title", color=[210, 220, 255])
                            _themes_mod.add_text("", tag="theory_hex", color=[210, 220, 255])
                        _themes_mod.add_text("", tag="theory_temp_lbl", color=[200, 200, 200])
                        # Lämpötilamittari: viileä (sininen) -> lämmin (oranssi); merkki päivittyy update_loopissa
                        _tw, _th, _tn = W - _sc(30), _sc(12), 40
                        dpg.add_drawlist(width=_tw, height=_th, tag="theory_temp_dl")
                        for _i in range(_tn):
                            dpg.draw_rectangle(
                                parent="theory_temp_dl",
                                pmin=(int(_i * _tw / _tn), 0), pmax=(int((_i + 1) * _tw / _tn) + 1, _th),
                                fill=_temp_gradient_color(_i / (_tn - 1)), color=[0, 0, 0, 0],
                            )
                        dpg.draw_line(parent="theory_temp_dl", p1=(3, 0), p2=(3, _th),
                                      color=[0, 0, 0, 255], thickness=5, tag="theory_temp_mark_bg")
                        dpg.draw_line(parent="theory_temp_dl", p1=(3, 0), p2=(3, _th),
                                      color=[255, 255, 255, 255], thickness=2, tag="theory_temp_mark")
                        dpg.add_spacer(height=4)
                        _themes_mod.add_text("", tag="theory_association_lbl", color=[200, 200, 200])
                        _themes_mod.add_text("", tag="theory_tags", wrap=W - _sc(30), color=[130, 210, 200])
                        dpg.add_spacer(height=4)
                        _themes_mod.add_text("", tag="theory_text", wrap=W - _sc(30), color=[200, 200, 200])

                with dpg.tab(label=tab_label(t("tab_contrast")), tag="tab_contrast"):
                    with dpg.child_window(
                        width=W, height=_sc(238), border=False, no_scrollbar=False,
                        tag="contrast_scroll"
                    ):
                        _BOX_H = _sc(28)
                        _BOX_W = (W - _sc(34)) // 3
                        dpg.add_spacer(height=_sc(6))
                        with dpg.group(horizontal=True):
                            with dpg.group(width=_BOX_W):
                                dpg.add_button(label="W --", tag="wheel_ctr_w", width=_BOX_W, height=_BOX_H)
                                _themes_mod.add_text("--", tag="wheel_ctr_w_res")
                            with dpg.group(width=_BOX_W):
                                dpg.add_button(label="B --", tag="wheel_ctr_b", width=_BOX_W, height=_BOX_H)
                                _themes_mod.add_text("--", tag="wheel_ctr_b_res")
                            with dpg.group(width=_BOX_W):
                                dpg.add_button(label="C --", tag="wheel_ctr_custom", width=_BOX_W, height=_BOX_H)
                                _themes_mod.add_text("--", tag="wheel_ctr_custom_res")
                        dpg.add_spacer(height=_sc(4))
                        with dpg.group(horizontal=True):
                            dpg.add_button(
                                label=t("add_to_bg"), tag="btn_add_to_bg",
                                width=_BOX_W, height=_BOX_H,
                                callback=lambda s, a, u: set_current_as_sample_background(),
                            )
                            dpg.add_button(label=t("sample_text"), tag="wheel_preview_w_txt",
                                           width=_BOX_W * 2 + _sc(8), height=_BOX_H)
                        if _font_large is not None:
                            dpg.bind_item_font("wheel_preview_w_txt", _font_large)
                        
                        dpg.add_spacer(height=_sc(4))
                        with dpg.group(tag="ctr_legend_grp") as _ctr_legend_grp:
                            _themes_mod.add_text(t("wcag_title"),  tag="lbl_wcag_title", color=[150, 150, 150, 255])
                            _themes_mod.add_text(t("wcag_fail"),   tag="lbl_wcag_fail",  color=[200, 70,  70,  255])
                            _themes_mod.add_text(t("wcag_aa_lg"),  tag="lbl_wcag_aalg",  color=[230, 160, 40,  255])
                            _themes_mod.add_text(t("wcag_aa"),     tag="lbl_wcag_aa",    color=[170, 220, 60,  255])
                            _themes_mod.add_text(t("wcag_aaa"),    tag="lbl_wcag_aaa",   color=[80,  210, 80,  255])
                            dpg.add_spacer(height=_sc(2))
                            _themes_mod.add_text(t("wcag_w"),      tag="lbl_wcag_w",     color=[120, 120, 120, 255])
                            _themes_mod.add_text(t("wcag_b"),      tag="lbl_wcag_b",     color=[120, 120, 120, 255])
                            _themes_mod.add_text(t("wcag_c"),      tag="lbl_wcag_c",     color=[120, 120, 120, 255])
                        dpg.bind_item_theme(_ctr_legend_grp, _t._tight_text_theme)

                with dpg.tab(label=tab_label(t("tab_simulation")), tag="tab_simulation"):
                    with dpg.child_window(
                        width=W, height=_sc(238), border=False, no_scrollbar=False,
                        tag="simulation_scroll"
                    ):
                        dpg.add_spacer(height=_sc(6))
                        _themes_mod.add_text(t("sim_preview"), tag="lbl_sim_preview", color=[160, 200, 255])
                        dpg.add_spacer(height=_sc(4))
                        
                        _BOX_H = _sc(28)
                        dpg.add_button(
                            label=t("add_to_bg"), tag="btn_sim_add_to_bg",
                            width=W - _sc(20), height=_BOX_H,
                            callback=lambda s, a, u: set_current_as_sample_background(),
                        )
                        dpg.add_spacer(height=_sc(6))
                        
                        dpg.add_button(label=t("sim_prot"), tag="btn_sim_prot", width=W - _sc(20), height=_BOX_H)
                        dpg.add_button(label=t("sim_deut"), tag="btn_sim_deut", width=W - _sc(20), height=_BOX_H)
                        dpg.add_button(label=t("sim_trit"), tag="btn_sim_trit", width=W - _sc(20), height=_BOX_H)
                        dpg.add_spacer(height=_sc(4))
                        dpg.add_button(label=t("sim_norm"), tag="btn_sim_norm", width=W - _sc(20), height=_BOX_H)
                        
                        if _font_large is not None:
                            dpg.bind_item_font("btn_sim_prot", _font_large)
                            dpg.bind_item_font("btn_sim_deut", _font_large)
                            dpg.bind_item_font("btn_sim_trit", _font_large)
                            dpg.bind_item_font("btn_sim_norm", _font_large)
                dpg.add_tab_button(label=_ARROW_UP, trailing=True, tag="btn_collapse_utility", callback=toggle_utility, no_tooltip=True)

        dpg.add_separator()
        dpg.add_checkbox(
            label=t("always_on_top"),
            tag="always_on_top_chk",
            callback=toggle_on_top,
            default_value=app.always_on_top,
        )

        apply_collapse_state()

        if _font_triangle is not None:
            if dpg.does_item_exist("btn_collapse_tab_area"):
                dpg.bind_item_font("btn_collapse_tab_area", _font_triangle)
            if dpg.does_item_exist("btn_collapse_utility"):
                dpg.bind_item_font("btn_collapse_utility", _font_triangle)
            if dpg.does_item_exist("btn_collapse_top_view"):
                dpg.bind_item_font("btn_collapse_top_view", _font_triangle)

    with dpg.window(label='Color Tools', tag='operation_status_window',
                    show=False, autosize=True, no_saved_settings=True, pos=(_sc(8), _sc(45))):
        _themes_mod.add_text('', tag='operation_status_text', wrap=_sc(230))
        dpg.add_button(label='OK', callback=lambda: dpg.configure_item('operation_status_window', show=False))

    dpg.bind_item_theme("tab_area", _t._nopad_cw_theme)
    dpg.bind_item_theme("harmony_rows_scroll", _t._nopad_cw_theme)
    with dpg.theme() as _palette_inset_theme:
        with dpg.theme_component(dpg.mvChildWindow):
            dpg.add_theme_style(dpg.mvStyleVar_WindowPadding, _sc(6), _sc(6))
    dpg.bind_item_theme("pal_scroll", _palette_inset_theme)
    dpg.bind_item_theme("pal_edit_win", _palette_inset_theme)
    fonts.bind_value_font("preview_hex")
    # The inline HEX shares the name's UI font and baseline in the Theory panel.
    for tag in HARMONY_TAGS:
        fonts.bind_value_font(f"val_{tag}")

    return {
        "W": W, "BW": BW, "VP_W": _VP_W, "TB_H": _TB_H,
        "ICON_W": _ICON_W, "IC": _IC, "THICK_ICON": _THICK_ICON,
        "logo_dark_w": _logo_dark_w, "logo_dark_h": _logo_dark_h,
        "logo_light_w": _logo_light_w, "logo_light_h": _logo_light_h,
        "font_large": _font_large,
    }


def activate_primary_window():
    """Apply fixed-layout flags after Dear PyGui resets primary-window flags."""
    dpg.set_primary_window("PrimaryWindow", True)
    dpg.configure_item("PrimaryWindow", no_scrollbar=True, no_scroll_with_mouse=True)
    dpg.set_x_scroll("PrimaryWindow", 0)
    dpg.set_y_scroll("PrimaryWindow", 0)
