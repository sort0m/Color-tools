"""
themes.py — Kaikki DPG-teemat ja set_theme-funktio.

Teemat luodaan funktiossa create_themes() jota kutsutaan
dpg.create_context():n jälkeen.
"""

import dearpygui.dearpygui as dpg

from .bootstrap import _sc
from .accessibility import contrast_color

# Teema-muuttujat — asetetaan create_themes()-funktion jälkeen
_nospc_theme          = None
_nopad_cw_theme       = None
_tight_text_theme     = None
_close_btn_theme      = None
_new_palette_select_theme = None
_hist_selected_theme  = None

_dark_theme           = None
_light_theme          = None
_midnight_theme       = None
_mocha_theme          = None
_nord_theme           = None
_solarized_theme      = None
_black_theme          = None

_titlebar_dark_theme      = None
_titlebar_light_theme     = None
_titlebar_midnight_theme  = None
_titlebar_mocha_theme     = None
_titlebar_nord_theme      = None
_titlebar_solarized_theme = None
_titlebar_black_theme     = None

_THEME_MAP  = {}
THEME_NAMES = []
TEXT_SURFACES = (
    'WindowBg', 'ChildBg', 'PopupBg', 'FrameBg', 'FrameBgHovered', 'FrameBgActive',
    'Button', 'ButtonHovered', 'ButtonActive', 'Header', 'HeaderHovered', 'HeaderActive',
    'Tab', 'TabHovered', 'TabActive', 'TabUnfocused', 'TabUnfocusedActive',
    'TitleBg', 'TitleBgActive', 'TitleBgCollapsed', 'MenuBarBg', 'TextSelectedBg',
    'TableHeaderBg', 'TableRowBg', 'TableRowBgAlt',
)
_TEXT_ITEMS = {}
_ACCESSIBLE_PALETTES = {}


def text_color(preferred):
    from .state import app
    palette = _ACCESSIBLE_PALETTES.get(app.theme_name, _ACCESSIBLE_PALETTES.get('Dark'))
    backgrounds = [palette[key] for key in ('WindowBg', 'ChildBg', 'PopupBg')] if palette else [[26, 26, 26]]
    return contrast_color(preferred, backgrounds)


def _prune_text_items():
    """Discard deleted widgets before registering another dynamic label."""
    for item in list(_TEXT_ITEMS):
        if not dpg.does_item_exist(item):
            del _TEXT_ITEMS[item]


def register_text(item, preferred):
    _prune_text_items()
    _TEXT_ITEMS[item] = list(preferred)
    dpg.configure_item(item, color=text_color(preferred))
    return item


def add_text(*args, **kwargs):
    """Register explicit text colors so newly opened dialogs follow the theme."""
    preferred = kwargs.get('color')
    if preferred is not None:
        _prune_text_items()
        kwargs['color'] = text_color(preferred)
    item = dpg.add_text(*args, **kwargs)
    if preferred is not None:
        _TEXT_ITEMS[item] = list(preferred)
    return item


def _make_accessible(theme, is_light=False):
    component = dpg.get_item_children(theme, 1)[0]
    nodes = {dpg.get_item_configuration(item)['target']: item
             for item in dpg.get_item_children(component, 1)
             if dpg.get_item_type(item) == 'mvAppItemType::mvThemeColor'}
    def original(name, fallback):
        item = nodes.get(getattr(dpg, 'mvThemeCol_' + name))
        return dpg.get_value(item) if item else fallback
    foreground = [20, 20, 20, 255] if is_light else [245, 245, 245, 255]
    window = original('WindowBg', [26, 26, 26, 255])
    palette = {'Text': foreground, 'TextDisabled': foreground}
    for name in TEXT_SURFACES:
        fallback = original('ButtonHovered', window) if name.endswith(('Active', 'Hovered')) else window
        if name == 'TextSelectedBg':
            fallback = original('Header', original('ButtonHovered', window))
        palette[name] = contrast_color(original(name, fallback), [foreground])
    # Keep the same hover relationship in every theme. Light text already
    # gains contrast when darkened; dark text may need a lighter active base.
    active = palette['TabActive']
    for step in range(256):
        base = [round(c + (255 - c) * step / 255) for c in active[:3]]
        hover = [round(c * 0.88) for c in base] + [255]
        if contrast_color(hover, [foreground]) == hover:
            palette['TabActive'] = base + [255]
            palette['TabUnfocusedActive'] = base + [255]
            palette['TabHovered'] = hover
            break
    for name, color in palette.items():
        target = getattr(dpg, 'mvThemeCol_' + name)
        if target in nodes:
            dpg.set_value(nodes[target], color)
        else:
            dpg.add_theme_color(target, color, parent=component)
    dpg.add_theme_style(dpg.mvStyleVar_DisabledAlpha, 1., parent=component)
    with dpg.theme_component(dpg.mvAll, parent=theme, enabled_state=False):
        # Flatten disabled controls into the window, retaining readable text.
        for name, color in palette.items():
            if name in ('Button', 'ButtonHovered', 'ButtonActive', 'FrameBg'):
                color = palette['WindowBg']
            dpg.add_theme_color(getattr(dpg, 'mvThemeCol_' + name), color)
        dpg.add_theme_color(dpg.mvThemeCol_Border, contrast_color([110, 110, 110], [foreground]))
        dpg.add_theme_style(dpg.mvStyleVar_FrameBorderSize, 1.)
        dpg.add_theme_style(dpg.mvStyleVar_DisabledAlpha, 1.)
    return palette


def _accessible_button(theme, colors):
    component = dpg.get_item_children(theme, 1)[0]
    # Explicit opaque colors for every interaction state, independent of theme.
    dpg.delete_item(component, children_only=True)
    for name, color in zip(('Button', 'ButtonHovered', 'ButtonActive'), colors):
        dpg.add_theme_color(getattr(dpg, 'mvThemeCol_' + name), contrast_color(color, [[255, 255, 255]]), parent=component)
    dpg.add_theme_color(dpg.mvThemeCol_Text, [255, 255, 255, 255], parent=component)

def create_themes():
    """Luo kaikki DPG-teemat. Kutsuttava dpg.create_context():n jälkeen."""
    global _nospc_theme, _nopad_cw_theme, _tight_text_theme
    global _close_btn_theme
    global _new_palette_select_theme, _hist_selected_theme
    global _dark_theme, _light_theme, _midnight_theme
    global _mocha_theme, _nord_theme, _solarized_theme, _black_theme
    global _titlebar_dark_theme, _titlebar_light_theme, _titlebar_midnight_theme
    global _titlebar_mocha_theme, _titlebar_nord_theme, _titlebar_solarized_theme
    global _titlebar_black_theme, _THEME_MAP, THEME_NAMES
    _TEXT_ITEMS.clear()
    _ACCESSIBLE_PALETTES.clear()

    with dpg.theme() as _nospc_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing,  0, 0)
            dpg.add_theme_style(dpg.mvStyleVar_FramePadding, 0, 0)

    with dpg.theme() as _nopad_cw_theme:
        with dpg.theme_component(dpg.mvChildWindow):
            dpg.add_theme_style(dpg.mvStyleVar_WindowPadding, 0, 0)

    with dpg.theme() as _tight_text_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing, 4, 1)

    with dpg.theme() as _close_btn_theme:
        with dpg.theme_component(dpg.mvButton):
            dpg.add_theme_color(dpg.mvThemeCol_Button,        [180, 50, 50, 255])
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, [220, 80, 80, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Text,          [255, 255, 255, 255])

    with dpg.theme() as _new_palette_select_theme:
        with dpg.theme_component(dpg.mvButton):
            dpg.add_theme_color(dpg.mvThemeCol_Button,        [0, 150, 0, 255])
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, [0, 180, 0, 255])
            dpg.add_theme_color(dpg.mvThemeCol_ButtonActive,  [0, 200, 0, 255])

    with dpg.theme() as _hist_selected_theme:
        with dpg.theme_component(dpg.mvColorButton):
            dpg.add_theme_color(
                dpg.mvThemeCol_Border, [220, 220, 220, 255],
                category=dpg.mvThemeCat_Core,
            )
            dpg.add_theme_style(dpg.mvStyleVar_FrameBorderSize, 2)

    # ── Pääväriteemat ──────────────────────────────────────────────────────────
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
            dpg.add_theme_color(dpg.mvThemeCol_Tab,            [45,  45,  45,  255])
            dpg.add_theme_color(dpg.mvThemeCol_TabActive,      [25,  85, 140,  255])
            dpg.add_theme_color(dpg.mvThemeCol_TabUnfocused,   [45,  45,  45,  255])
            dpg.add_theme_color(dpg.mvThemeCol_TabUnfocusedActive, [25, 85, 140, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Separator,      [55,  55,  55,  255])
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
            dpg.add_theme_color(dpg.mvThemeCol_TabActive,          [66,  150, 250, 255])
            dpg.add_theme_color(dpg.mvThemeCol_TabUnfocused,       [200, 200, 200, 255])
            dpg.add_theme_color(dpg.mvThemeCol_TabUnfocusedActive, [66,  150, 250, 255])
            dpg.add_theme_style(dpg.mvStyleVar_TabRounding,  2)
            dpg.add_theme_style(dpg.mvStyleVar_FramePadding, 4, 3)

    with dpg.theme() as _midnight_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_WindowBg,       [13,  17,  23,  255])
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
            dpg.add_theme_color(dpg.mvThemeCol_TabActive,      [31,  54,  84,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Separator,      [30,  42,  56,  255])
            dpg.add_theme_style(dpg.mvStyleVar_TabRounding,    2)
            dpg.add_theme_style(dpg.mvStyleVar_FramePadding,   4, 3)

    with dpg.theme() as _mocha_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_WindowBg,       [30,  26,  23,  255])
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
            dpg.add_theme_color(dpg.mvThemeCol_TabActive,      [90,  65,  45,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Separator,      [65,  54,  46,  255])
            dpg.add_theme_style(dpg.mvStyleVar_TabRounding,    2)
            dpg.add_theme_style(dpg.mvStyleVar_FramePadding,   4, 3)

    with dpg.theme() as _nord_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_WindowBg,       [46,  52,  64,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ChildBg,        [46,  52,  64,  255])
            dpg.add_theme_color(dpg.mvThemeCol_FrameBg,        [59,  66,  82,  255])
            dpg.add_theme_color(dpg.mvThemeCol_FrameBgHovered, [67,  76,  94,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Text,           [216, 222, 233, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Button,         [67,  76,  94,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered,  [76,  86,  106, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Header,         [94,  129, 172, 255])
            dpg.add_theme_color(dpg.mvThemeCol_PopupBg,        [59,  66,  82,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ScrollbarBg,    [36,  41,  51,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Tab,            [59,  66,  82,  255])
            dpg.add_theme_color(dpg.mvThemeCol_TabActive,      [58,  90,  130, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Separator,      [67,  76,  94,  255])
            dpg.add_theme_style(dpg.mvStyleVar_TabRounding,    2)
            dpg.add_theme_style(dpg.mvStyleVar_FramePadding,   4, 3)

    with dpg.theme() as _solarized_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_WindowBg,       [0,   43,  54,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ChildBg,        [0,   43,  54,  255])
            dpg.add_theme_color(dpg.mvThemeCol_FrameBg,        [7,   54,  66,  255])
            dpg.add_theme_color(dpg.mvThemeCol_FrameBgHovered, [0,   72,  88,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Text,           [253, 246, 227, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Button,         [0,   72,  88,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered,  [0,   95,  115, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Header,         [42,  161, 152, 255])
            dpg.add_theme_color(dpg.mvThemeCol_PopupBg,        [7,   54,  66,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ScrollbarBg,    [0,   33,  42,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Tab,            [7,   54,  66,  255])
            dpg.add_theme_color(dpg.mvThemeCol_TabActive,      [0,   95,  89,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Separator,      [0,   72,  88,  255])
            dpg.add_theme_style(dpg.mvStyleVar_TabRounding,    2)
            dpg.add_theme_style(dpg.mvStyleVar_FramePadding,   4, 3)

    with dpg.theme() as _black_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_WindowBg,       [0,   0,   0,   255])
            dpg.add_theme_color(dpg.mvThemeCol_ChildBg,        [0,   0,   0,   255])
            dpg.add_theme_color(dpg.mvThemeCol_FrameBg,        [50,  50,  50,  255])
            dpg.add_theme_color(dpg.mvThemeCol_FrameBgHovered, [70,  70,  70,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Text,           [230, 230, 230, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Button,         [45,  45,  45,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered,  [70,  70,  70,  255])
            dpg.add_theme_color(dpg.mvThemeCol_Header,         [60,  60,  60,  255])
            dpg.add_theme_color(dpg.mvThemeCol_PopupBg,        [15,  15,  15,  255])
            dpg.add_theme_color(dpg.mvThemeCol_ScrollbarBg,    [0,   0,   0,   255])
            dpg.add_theme_color(dpg.mvThemeCol_Tab,            [45,  45,  45,  255])
            dpg.add_theme_color(dpg.mvThemeCol_TabActive,      [100, 100, 100, 255])
            dpg.add_theme_color(dpg.mvThemeCol_Separator,      [55,  55,  55,  255])
            dpg.add_theme_style(dpg.mvStyleVar_TabRounding,    2)
            dpg.add_theme_style(dpg.mvStyleVar_FramePadding,   4, 3)

    # ── Titlebar-teemat ────────────────────────────────────────────────────────
    def _make_titlebar_theme(bg):
        with dpg.theme() as th:
            with dpg.theme_component(dpg.mvAll):
                dpg.add_theme_color(dpg.mvThemeCol_ChildBg, bg)
                dpg.add_theme_color(dpg.mvThemeCol_Border,  bg)
                dpg.add_theme_style(dpg.mvStyleVar_WindowPadding, 0, _sc(5))
                dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing,   _sc(2), 0)
        return th

    _titlebar_dark_theme      = _make_titlebar_theme([26,  26,  26,  255])
    _titlebar_light_theme     = _make_titlebar_theme([235, 235, 235, 255])
    _titlebar_midnight_theme  = _make_titlebar_theme([13,  17,  23,  255])
    _titlebar_mocha_theme     = _make_titlebar_theme([30,  26,  23,  255])
    _titlebar_nord_theme      = _make_titlebar_theme([46,  52,  64,  255])
    _titlebar_solarized_theme = _make_titlebar_theme([0,   43,  54,  255])
    _titlebar_black_theme     = _make_titlebar_theme([0,   0,   0,   255])

    _THEME_MAP = {
        "Dark":      (_dark_theme,      _titlebar_dark_theme,      False),
        "Light":     (_light_theme,     _titlebar_light_theme,     True),
        "Midnight":  (_midnight_theme,  _titlebar_midnight_theme,  False),
        "Mocha":     (_mocha_theme,     _titlebar_mocha_theme,     False),
        "Nord":      (_nord_theme,      _titlebar_nord_theme,      False),
        "Solarized": (_solarized_theme, _titlebar_solarized_theme, False),
        "Black":     (_black_theme,     _titlebar_black_theme,     False),
    }
    THEME_NAMES = list(_THEME_MAP.keys())
    for name, (theme, _, is_light) in _THEME_MAP.items():
        _ACCESSIBLE_PALETTES[name] = _make_accessible(theme, is_light)
    _accessible_button(_close_btn_theme, ([180, 50, 50], [220, 80, 80], [160, 35, 35]))
    _accessible_button(_new_palette_select_theme, ([0, 150, 0], [0, 180, 0], [0, 200, 0]))

# ── set_theme ─────────────────────────────────────────────────────────────────
def set_theme(sender, name, _u=None):
    from .state import app, save_config
    theme_name = str(name)
    if theme_name not in _THEME_MAP:
        return
    main_theme, titlebar_theme, is_light = _THEME_MAP[theme_name]
    dpg.bind_theme(main_theme)
    if dpg.does_item_exist("titlebar_cw"):
        dpg.bind_item_theme("titlebar_cw", titlebar_theme)
    app.theme_name = theme_name
    app._last_sl_vals = {}  # pakota sliderien uudelleenpiirto (indikaattorin väri)
    if dpg.does_item_exist("theme_combo"):
        dpg.set_value("theme_combo", theme_name)

    for item, preferred in list(_TEXT_ITEMS.items()):
        if dpg.does_item_exist(item):
            dpg.configure_item(item, color=text_color(preferred))
        else:
            del _TEXT_ITEMS[item]

    # Päivitä logo
    if is_light:
        logo_tag = "logo_light"
    else:
        logo_tag = "logo_dark"
    if dpg.does_item_exist("tb_logo"):
        dpg.configure_item("tb_logo", texture_tag=logo_tag)
    save_config()


def slider_mode_cb(s, v):
    from .state import app
    from .sliders import SLIDERS_BY_MODE
    app.slider_mode = v
    for mode in SLIDERS_BY_MODE.keys():
        tag = f"grp_sl_{mode}"
        if dpg.does_item_exist(tag):
            dpg.configure_item(tag, show=(mode == v))
    app._last_sl_vals = {}

def harmony_combo_cb(s, v):
    from .state import app
    from .i18n import to_internal_harmony
    internal = to_internal_harmony(v)
    app.harmony_mode = internal
    _redraw_harmony_combo_icon(internal)

def _redraw_harmony_combo_icon(mode):
    import math
    from .bootstrap import _sc
    from .harmony import _HARMONY_ICON_DATA
    if not dpg.does_item_exist("hm_icon_dl"):
        return
    dpg.delete_item("hm_icon_dl", children_only=True)
    angles, tint_shd = _HARMONY_ICON_DATA.get(mode, ([], None))
    width = dpg.get_item_width("hm_icon_dl")
    height = dpg.get_item_height("hm_icon_dl")
    R  = _sc(17)
    CX = width / 2
    CY = height / 2
    DR = _sc(4)
    COL_RING = [100, 100, 100, 180]
    COL_MAIN = [255, 200,  50, 255]
    COL_HRM  = [ 80, 160, 255, 255]
    COL_LINE = [ 80,  80,  80, 140]
    BASE_A   = -math.pi / 2
    dpg.draw_circle(parent="hm_icon_dl", center=(CX, CY), radius=R,
                    color=COL_RING, thickness=1)
    if tint_shd:
        dot_radius = _sc(4)
        step = (width - 2 * (dot_radius + _sc(1))) / 4
        for i in range(5):
            tx = CX + (i - 2) * step
            t  = i / 4.0
            if tint_shd == "tint":
                c = [int(100+155*t), int(60+100*t), int(20+40*t), 255]
            elif tint_shd == "shade":
                c = [int(200-180*t), int(80-70*t), int(20-15*t), 255]
            else:
                c = [int(120+40*t), int(80+40*t), int(20+20*t), 255]
            dpg.draw_circle(parent="hm_icon_dl", center=(tx, CY), radius=dot_radius,
                            color=[0,0,0,0], fill=c)
    else:
        bx = int(CX + R * math.cos(BASE_A))
        by = int(CY + R * math.sin(BASE_A))
        for deg in angles:
            a  = BASE_A + math.radians(deg)
            hx = int(CX + R * math.cos(a))
            hy = int(CY + R * math.sin(a))
            dpg.draw_line(parent="hm_icon_dl", p1=(CX, CY), p2=(hx, hy),
                          color=COL_LINE, thickness=1)
            dpg.draw_circle(parent="hm_icon_dl", center=(hx, hy), radius=DR,
                            color=[0,0,0,0], fill=COL_HRM)
        dpg.draw_line(parent="hm_icon_dl", p1=(CX, CY), p2=(bx, by),
                      color=COL_LINE, thickness=1)
        dpg.draw_circle(parent="hm_icon_dl", center=(bx, by), radius=DR,
                        color=[0,0,0,0], fill=COL_MAIN)

def fmt_combo_cb(s, v):
    from .state import app
    app.fmt_mode = v
    from .harmony import refresh_harmony_values
    refresh_harmony_values()

def _on_top_tab(tab):
    from .state import app
    app.top_tab  = tab
    app.use_wheel = (tab == "wheel")
