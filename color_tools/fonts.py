"""Segoe UI labels and Consolas color values and numeric fields."""

import dearpygui.dearpygui as dpg

UI_FONT_SIZE = 17
VALUE_FONT_SIZE = 13
HEX_FONT_SIZE = UI_FONT_SIZE
# HTML uses Segoe UI throughout, independently of native value font sizing.
HTML_UI_FONT_SIZE = 14
UI_FONT_FAMILY = "Segoe UI"
hex_font = None  # Main UI font, also used for tab measurements.
value_font = None
_value_theme = None


def create_value_theme(row_height, font_height):
    """Center smaller value text without shrinking input fields."""
    global _value_theme
    with dpg.theme() as _value_theme:
        for enabled in (True, False):
            with dpg.theme_component(dpg.mvAll, enabled_state=enabled):
                dpg.add_theme_style(dpg.mvStyleVar_FramePadding, 4, (row_height - font_height) / 2)


def bind_value_font(item, enabled=True):
    """Use Consolas for values; fall back to the UI font if unavailable."""
    dpg.bind_item_font(item, (value_font or hex_font or 0) if enabled else 0)
    if _value_theme:
        dpg.bind_item_theme(item, _value_theme if enabled else 0)
