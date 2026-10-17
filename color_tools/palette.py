"""
palette.py — Palettien hallinta, muokkaus, undo, drag-and-drop, history-valinta.
"""


import time

import dearpygui.dearpygui as dpg
from . import themes as _themes_mod

from .bootstrap import _sc
from .color_math import color_sig
from .i18n import t
from .state import app, _palettes_lock, _history_lock, request_palette_save
from .palette_formats import MAX_COLORS
from .palette_history import pack_colors, unpack_colors, record_snapshot
from .harmony import (
    _force_update_color_picker, _update_wheel_contrast, HARMONY_TAGS,
    HIST_TAGS, _apply_external_color_to_pal,
)

# Palettien layout-vakio (asetetaan ui_main:ssa)
PAL_SW      = 0
MAX_PAL_ROWS = 100
_PCOLS      = 9
_ECOLS      = 8
_ESW        = 0
_BW2        = 0
THUMB_H     = 0

_nospc_theme    = None  # asetetaan ui_main:ssa
_close_btn_theme = None  # asetetaan ui_main:ssa

def init_palette_constants(pal_sw, max_pal_rows, ecols, esw, bw2, thumb_h, nospc, close_btn):
    global PAL_SW, MAX_PAL_ROWS, _ECOLS, _ESW, _BW2, THUMB_H, _nospc_theme, _close_btn_theme
    PAL_SW      = pal_sw
    MAX_PAL_ROWS = max_pal_rows
    _ECOLS      = ecols
    _ESW        = esw
    _BW2        = bw2
    THUMB_H     = thumb_h
    _nospc_theme    = nospc
    _close_btn_theme = close_btn

# ── Paletti-apufunktiot ───────────────────────────────────────────────────────
def _auto_pal_name():
    i = 1
    while f"Palette {i}" in app.palettes:
        i += 1
    return f"Palette {i}"

def can_add_palette():
    """Limit new palettes; larger existing collections remain pageable."""
    if len(app.palettes) >= MAX_PAL_ROWS:
        from .import_export import _navigate_to_palettes
        _navigate_to_palettes()
        app._import_status_code = None
        app._import_status = (t("pal_limit").format(n=MAX_PAL_ROWS),
                              time.perf_counter() + 5.0)
        return False
    return True

def _register_palette(name):
    if name not in app._pal_order:
        app._pal_order.append(name)
        app._pal_page = 0
        app._pal_dirty = True

def palette_names():
    order = [k for k in reversed(app._pal_order) if k in app.palettes]
    order_set = set(order)
    for k in sorted(app.palettes.keys()):
        if k not in order_set:
            order.append(k)
    return order

# ── Palettilistan piirtäminen ─────────────────────────────────────────────────
_PALETTES_PER_PAGE = 20
_pal_row_cache = {}


def sync_palette_panels():
    importing = app._image_modal_open or app._drop_preview_open
    for tag, visible in (
        ('pal_scroll', not importing and not app._editing_pal),
        ('pal_edit_win', not importing and bool(app._editing_pal)),
        ('pal_import_buttons', not importing),
        ('img_import_win', app._image_modal_open),
        ('drop_preview_win', app._drop_preview_open),
    ):
        if dpg.does_item_exist(tag):
            dpg.configure_item(tag, show=bool(visible))


def change_palette_page(sender, value, delta):
    app._pal_page = max(0, app._pal_page + delta)
    _rebuild_pal_rows()
    dpg.set_y_scroll('pal_scroll', 0)


def _rebuild_pal_rows():
    """Refresh only visible palette summaries, on the UI thread."""
    sync_palette_panels()
    if app._editing_pal:
        return
    with _palettes_lock:
        names = palette_names()
        pages = max(1, (len(names) + _PALETTES_PER_PAGE - 1) // _PALETTES_PER_PAGE)
        app._pal_page = min(app._pal_page, pages - 1)
        start = app._pal_page * _PALETTES_PER_PAGE
        rows = [(name, len(app.palettes[name]),
                 [list(c) for c in app.palettes[name][:_PCOLS]])
                for name in names[start:start + _PALETTES_PER_PAGE]]
    if dpg.does_item_exist('pal_page_controls'):
        dpg.configure_item('pal_page_controls', show=pages > 1)
        dpg.set_value('pal_page_label', f'{app._pal_page + 1} / {pages}')
        dpg.configure_item('pal_page_prev', enabled=app._pal_page > 0)
        dpg.configure_item('pal_page_next', enabled=app._pal_page < pages - 1)

    def swatch_cb(sender, value, color):
        _force_update_color_picker(color)
        from .sliders import _sync_all_modes
        _sync_all_modes()
        _update_wheel_contrast()

    for index in range(_PALETTES_PER_PAGE):
        tag = f'pal_row_{index}'
        if not dpg.does_item_exist(tag):
            continue
        dpg.configure_item(tag, show=index < len(rows))
        if index >= len(rows):
            continue
        name, count, colors = rows[index]
        signature = (name, count, tuple(map(tuple, colors)), app.language)
        children = dpg.get_item_children(tag, 1)
        if _pal_row_cache.get(index) == signature and children:
            continue
        dpg.delete_item(tag, children_only=True)
        dpg.bind_item_theme(tag, _nospc_theme)
        group = dpg.add_group(parent=tag, horizontal=True)
        dpg.bind_item_theme(group, _nospc_theme)
        for color in colors:
            dpg.add_color_button(parent=group, default_value=color,
                width=PAL_SW, height=PAL_SW, no_alpha=True,
                callback=swatch_cb, user_data=[c / 255. for c in color[:3]] + [1.])
        if not colors:
            _themes_mod.add_text(t('pal_empty'), parent=group)
        dpg.add_button(parent=group, label='>>', width=_sc(24), height=PAL_SW,
                       callback=lambda s, a, u: open_palette_editor(u), user_data=name)
        dpg.add_spacer(height=10, parent=tag)
        _pal_row_cache[index] = signature


_EDIT_COLORS_PER_PAGE = 96


def _pal_edit_range(count):
    pages = max(1, (count + _EDIT_COLORS_PER_PAGE - 1) // _EDIT_COLORS_PER_PAGE)
    app._pal_edit_page = max(0, min(app._pal_edit_page, pages - 1))
    start = app._pal_edit_page * _EDIT_COLORS_PER_PAGE
    return start, min(count, start + _EDIT_COLORS_PER_PAGE), pages


def _cancel_pal_drag():
    app._pal_drag_idx = None
    app._pal_drag_insert = None
    app._pal_drag_active = False


def change_editor_page(sender, value, delta):
    if app._pal_edit_dirty:
        request_palette_save()
    _cancel_pal_drag()
    # A hidden selected color must not change while browsing another page.
    app._pal_selected_idx = None
    app._pal_pre_edit_snapshot = None
    app._pal_edit_dirty = False
    app._pal_edit_page += delta
    _refresh_pal_edit_panel()
    dpg.set_y_scroll('pal_edit_win', 0)


def _move_palette_color(source, destination):
    """Move to a zero-based final position; preserve every color and undo."""
    name = app._editing_pal
    with _palettes_lock:
        colors = app.palettes.get(name, [])
        if not (0 <= source < len(colors) and 0 <= destination < len(colors)):
            return False
        if source == destination:
            return False
    _push_pal_undo(name)
    with _palettes_lock:
        colors.insert(destination, colors.pop(source))
        selected = app._pal_selected_idx
        if selected == source:
            app._pal_selected_idx = destination
        elif selected is not None:
            if source < selected <= destination:
                app._pal_selected_idx -= 1
            elif destination <= selected < source:
                app._pal_selected_idx += 1
        app._pal_pre_edit_snapshot = pack_colors(colors)
    app._pal_edit_dirty = False
    app._pal_dirty = True
    request_palette_save()
    return True


def open_palette_editor(name):
    if not name or name not in app.palettes:
        return
    if app._pal_edit_dirty:
        request_palette_save()
    _cancel_pal_drag()
    app._pal_edit_page = 0
    app._pal_selected_idx = None
    app._pal_pre_edit_snapshot = None
    app._pal_edit_dirty = False
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
        colors_copy = pack_colors(app.palettes[name])
    record_snapshot(app._pal_undo_stack, name, colors_copy)
    if dpg.does_item_exist("pal_undo_btn"):
        dpg.configure_item("pal_undo_btn", enabled=_pal_undo_available())

def _pal_undo_available():
    """True jos muokattavalle palettille on olemassa undo-merkintä."""
    cur = app._editing_pal
    return any(e[0] == cur for e in app._pal_undo_stack)

def _purge_pal_undo(name):
    """Poistaa paletin undo-merkinnät (esim. paletin poiston jälkeen)."""
    if not name:
        return
    kept = [e for e in app._pal_undo_stack if e[0] != name]
    app._pal_undo_stack.clear()
    app._pal_undo_stack.extend(kept)

def _pal_undo(s, a, u):
    # Undo koskee vain parhaillaan muokattavaa palettia.
    _cancel_pal_drag()
    cur = app._editing_pal
    pos = None
    for k in range(len(app._pal_undo_stack) - 1, -1, -1):
        if app._pal_undo_stack[k][0] == cur:
            pos = k
            break
    if pos is None:
        return
    name, snapshot = app._pal_undo_stack[pos]
    colors = unpack_colors(snapshot)
    del app._pal_undo_stack[pos]
    idx = app._pal_selected_idx
    # Suppress palette writes only while restoring the selected color, not
    # for subsequent user actions. Programmatic picker updates are deferred.
    app._pal_selected_idx = None
    app._pal_edit_dirty = False
    with _palettes_lock:
        app.palettes[name] = colors
        if name not in app._pal_order:
            app._pal_order.append(name)
    app._pal_pre_edit_snapshot = pack_colors(colors)
    if (name == app._editing_pal
            and idx is not None
            and idx < len(colors)):
        c = colors[idx]
        _force_update_color_picker([c[0] / 255., c[1] / 255., c[2] / 255., 1.0])
        from .sliders import _sync_all_modes
        _sync_all_modes()
        _update_wheel_contrast()
    app._pal_selected_idx = idx if (idx is not None and idx < len(colors)) else None
    request_palette_save()
    _rebuild_pal_rows()
    _refresh_pal_edit_panel()
    if dpg.does_item_exist("pal_undo_btn"):
        dpg.configure_item("pal_undo_btn", enabled=_pal_undo_available())

def _refresh_pal_edit_panel(drag_src=None, drag_insert=None):
    """Rakentaa palettieditorin uudelleen."""
    if not dpg.does_item_exist("pal_edit_panel"):
        return
    dpg.delete_item("pal_edit_panel", children_only=True)
    name = app._editing_pal
    if not name or name not in app.palettes:
        sync_palette_panels()
        return
    dpg.configure_item("pal_edit_panel", show=True)
    with _palettes_lock:
        if name not in app.palettes:
            return
        count = len(app.palettes[name])
        start, end, pages = _pal_edit_range(count)
        colors = list(app.palettes[name][start:end])

    # ── Callbackit ──────────────────────────────────────────────────────
    def _close_edit(s, a, u):
        # Schedule the last edit even when the editor closes before the history delay.
        if app._pal_edit_dirty:
            request_palette_save()
        app._editing_pal = None
        app._pal_selected_idx = None
        _cancel_pal_drag()
        app._pending_close_edit = True

    def _add_current(s, a, u):
        nm = app._editing_pal
        if not nm or nm not in app.palettes:
            return
        if len(app.palettes[nm]) >= MAX_COLORS:
            app._import_status_code = None
            app._import_status = (t('pal_color_limit'), time.perf_counter() + 5)
            return
        r, g, b = color_sig(app.current_color)
        _push_pal_undo(nm)
        with _palettes_lock:
            app.palettes[nm].append([r, g, b, 255])
            app._pal_edit_page = (len(app.palettes[nm]) - 1) // _EDIT_COLORS_PER_PAGE
        app._pal_selected_idx = None
        app._pal_pre_edit_snapshot = None
        app._pal_edit_dirty = False
        app._pal_dirty = True
        request_palette_save()
        _rebuild_pal_rows()
        _refresh_pal_edit_panel()

    def _del_pal(s, a, u):
        if dpg.does_item_exist("pal_normal_btns"):
            dpg.configure_item("pal_normal_btns", show=False)
        if dpg.does_item_exist("pal_del_confirm_grp"):
            dpg.configure_item("pal_del_confirm_grp", show=True)
        if dpg.does_item_exist("pal_close_btn"):
            dpg.configure_item("pal_close_btn", show=False)

    def _del_pal_confirmed(s, a, u):
        if dpg.does_item_exist("pal_del_confirm_grp"):
            dpg.configure_item("pal_del_confirm_grp", show=False)
        if dpg.does_item_exist("pal_normal_btns"):
            dpg.configure_item("pal_normal_btns", show=True)
        nm = app._editing_pal
        with _palettes_lock:
            if nm and nm in app.palettes:
                del app.palettes[nm]
                if nm in app._pal_order:
                    app._pal_order.remove(nm)
        _purge_pal_undo(nm)
        app._editing_pal = None
        app._pal_selected_idx = None
        app._pal_dirty = True
        request_palette_save()
        _rebuild_pal_rows()
        _refresh_pal_edit_panel()

    def _del_pal_cancel(s, a, u):
        if dpg.does_item_exist("pal_del_confirm_grp"):
            dpg.configure_item("pal_del_confirm_grp", show=False)
        if dpg.does_item_exist("pal_normal_btns"):
            dpg.configure_item("pal_normal_btns", show=True)
        if dpg.does_item_exist("pal_close_btn"):
            dpg.configure_item("pal_close_btn", show=True)

    # ── Väriswatch:t ─────────────────────────────────────────────────────
    from .themes import _THEME_MAP
    dragging = (drag_src is not None and drag_insert is not None
                and start <= drag_src < end and len(colors) > 1)
    if dragging:
        stripped = [c for i, c in enumerate(colors) if i + start != drag_src]
        safe_ins = drag_insert - start
        ins = max(0, min(len(stripped), safe_ins))
        preview = stripped[:ins] + [None] + stripped[ins:]
    else:
        preview = list(colors)

    with dpg.group(parent="pal_edit_panel", tag="pal_edit_main"):
        with dpg.group(horizontal=True, show=pages > 1):
            dpg.add_button(label='<', tag='pedit_prev', callback=change_editor_page,
                           user_data=-1, enabled=app._pal_edit_page > 0)
            _themes_mod.add_text(t('pal_editor_page').format(page=app._pal_edit_page + 1, pages=pages), tag='pedit_page_label')
            dpg.add_button(label='>', tag='pedit_next', callback=change_editor_page,
                           user_data=1, enabled=app._pal_edit_page < pages - 1)
        if preview:
            for row_start in range(0, len(preview), _ECOLS):
                grp = dpg.add_group(horizontal=True)
                dpg.bind_item_theme(grp, _nospc_theme)
                for slot in range(row_start, min(row_start + _ECOLS, len(preview))):
                    c  = preview[slot]
                    dl = dpg.add_drawlist(parent=grp, tag=f"pedit_sw_{start + slot}",
                                         width=_ESW, height=_ESW)
                    if c is None:
                        dc = colors[drag_src - start]
                        dpg.draw_rectangle(parent=dl,
                                           pmin=(1, 1), pmax=(_ESW - 2, _ESW - 2),
                                           fill=[dc[0], dc[1], dc[2], 255],
                                           color=[220, 200, 50, 255], thickness=2)
                    else:
                        is_sel = (app._pal_selected_idx == start + slot and not dragging)
                        is_light = _THEME_MAP.get(app.theme_name, (None, None, False))[2]
                        sel_border = [40, 40, 40, 255] if is_light else [255, 255, 255, 255]
                        border = sel_border if is_sel else [40, 40, 40, 180]
                        thick = 2 if is_sel else 1
                        dpg.draw_rectangle(parent=dl,
                                           pmin=(1, 1), pmax=(_ESW - 2, _ESW - 2),
                                           fill=[c[0], c[1], c[2], 255],
                                           color=border, thickness=thick)
        else:
            _themes_mod.add_text(t("pal_empty_edit"), color=[120, 120, 120])

        dpg.add_spacer(height=2)
        dpg.add_separator()
        dpg.add_spacer(height=2)

        with dpg.group(tag="pal_normal_btns"):
            _grp2 = dpg.add_group(horizontal=True, tag="pal_action_btns")
            dpg.add_button(parent=_grp2, label=t("add_color"), width=_BW2, height=THUMB_H, callback=_add_current)
            dpg.add_button(parent=_grp2, tag="pal_undo_btn", label=t("undo"), width=_BW2, height=THUMB_H,
                           callback=_pal_undo, enabled=_pal_undo_available())
            dpg.add_spacer(height=2)
            _grp1 = dpg.add_group(horizontal=True, tag="pal_export_btns")
            from .import_export import _show_export_format
            dpg.add_button(parent=_grp1, label=t("export"), width=_BW2, height=THUMB_H, callback=_show_export_format)
            dpg.add_button(parent=_grp1, label=t("delete_palette"), width=_BW2, height=THUMB_H, callback=_del_pal, tag="pal_delete_btn")

        with dpg.group(tag="pal_del_confirm_grp", show=False):
            dpg.add_spacer(height=_sc(6))
            _themes_mod.add_text(t("delete_confirm"), color=[220, 80, 80])
            dpg.add_spacer(height=_sc(6))
            _grp_del = dpg.add_group(horizontal=True)
            dpg.add_button(parent=_grp_del, label=t("delete_btn"), width=_BW2, height=THUMB_H, callback=_del_pal_confirmed)
            dpg.add_button(parent=_grp_del, label=t("cancel"), width=_BW2, height=THUMB_H, callback=_del_pal_cancel)

        dpg.add_spacer(height=2)
        dpg.add_button(
            label=t("close"), width=_BW2, height=THUMB_H,
            callback=_close_edit, tag="pal_close_btn",
        )

    # Vientimuoto-paneeli
    from .import_export import _hide_export_format, _execute_palette_export
    with dpg.group(parent="pal_edit_panel", horizontal=False, tag="pal_export_format_panel", show=False):
        _themes_mod.add_text(t("export_format"), tag="lbl_export_format", color=[200, 200, 200])
        dpg.add_spacer(height=_sc(8))
        dpg.add_radio_button(
            items=["HTML (web)", "ASE (Adobe)", "ACO (Adobe)", "GPL (GIMP)",
                   "Procreate (.swatches)", "Affinity (.afpalette)"],
            default_value="HTML (web)",
            tag="pal_export_format_radio",
        )
        dpg.add_spacer(height=_sc(10))
        with dpg.group(horizontal=True):
            dpg.add_button(
                label=t("export"),
                width=_BW2,
                height=THUMB_H,
                callback=lambda s, a, u: _execute_palette_export(
                    dpg.get_value("pal_export_format_radio").split("(")[0].strip()
                ),
            )
            dpg.add_button(
                label=t("cancel"),
                width=_BW2,
                height=THUMB_H,
                callback=_hide_export_format,
            )

    dpg.add_spacer(height=2, parent="pal_edit_panel")
    if dpg.does_item_exist("pal_delete_btn"):
        dpg.bind_item_theme("pal_delete_btn", _close_btn_theme)

def _update_active_pal_swatch(idx, r, g, b):
    """Päivittää yhden swatch-suorakulmion täytön nopeasti (O(1))."""
    tag = f"pedit_sw_{idx}"
    if not dpg.does_item_exist(tag):
        return False
    children = dpg.get_item_children(tag, 2)
    if not children:
        return False
    rect_id = children[0]
    dpg.configure_item(rect_id, fill=[r, g, b, 255])
    return True

# ── Oikea nappi: värin poisto ─────────────────────────────────────────────────
def _pal_swatch_right_click(rmb_clicked):
    if not rmb_clicked:
        return
    if app._editing_pal and app._editing_pal in app.palettes:
        with _palettes_lock:
            ncolors = len(app.palettes.get(app._editing_pal, []))
        start, end, _ = _pal_edit_range(ncolors)
        for idx in range(start, end):
            stag = f"pedit_sw_{idx}"
            if dpg.does_item_exist(stag) and dpg.is_item_hovered(stag):
                _cancel_pal_drag()
                _push_pal_undo(app._editing_pal)
                with _palettes_lock:
                    nm = app._editing_pal
                    if nm and nm in app.palettes and idx < len(app.palettes[nm]):
                        app.palettes[nm].pop(idx)
                        # Viimeisenkin värin poisto jättää tyhjän paletin (undo toimii);
                        # koko paletin poisto tehdään vain Poista paletti -napilla.
                        if app._pal_selected_idx is not None:
                            if app._pal_selected_idx == idx:
                                app._pal_selected_idx = None
                                app._pal_pre_edit_snapshot = None
                                app._pal_edit_dirty = False
                            elif app._pal_selected_idx > idx:
                                app._pal_selected_idx -= 1
                        app._pal_pre_edit_snapshot = pack_colors(app.palettes[nm])
                        app._pal_edit_dirty = False
                app._pal_dirty = True
                request_palette_save()
                _rebuild_pal_rows()
                _refresh_pal_edit_panel()
                return

# ── Drag-and-drop järjestely ──────────────────────────────────────────────────
def _pal_swatch_drag(lmb_dn, lmb_clk, lmb_rel):
    """Hit-test only this page; drag destinations are final global indices."""
    name = app._editing_pal
    if not name or name not in app.palettes:
        return
    with _palettes_lock:
        start, end, _ = _pal_edit_range(len(app.palettes[name]))
    mouse = dpg.get_mouse_pos(local=False)
    if not mouse or len(mouse) < 2:
        return
    mx, my = mouse

    def hovered_slot():
        for index in range(start, end):
            tag = f'pedit_sw_{index}'
            if dpg.does_item_exist(tag) and dpg.is_item_hovered(tag):
                return index
        return None

    if lmb_clk:
        index = hovered_slot()
        if index is not None:
            if app._pal_edit_dirty:
                request_palette_save()
            app._pal_drag_idx = index
            app._pal_drag_insert = index
            app._pal_drag_active = False
            app._pal_drag_start = (mx, my)
            app._pal_edit_dirty = False
            if app._pal_selected_idx == index:
                app._pal_selected_idx = None
                app._pal_pre_edit_snapshot = None
            else:
                app._pal_selected_idx = index
                app._last_edit_color_sig = None
                with _palettes_lock:
                    app._pal_pre_edit_snapshot = pack_colors(app.palettes[name])
                    color = app.palettes[name][index]
                _force_update_color_picker([v / 255. for v in color[:3]] + [1.])
                from .sliders import _sync_all_modes
                _sync_all_modes()
                _update_wheel_contrast()
            _refresh_pal_edit_panel()

    if lmb_dn and app._pal_drag_idx is not None:
        dx, dy = mx - app._pal_drag_start[0], my - app._pal_drag_start[1]
        changed = False
        if not app._pal_drag_active and dx * dx + dy * dy > 16:
            app._pal_drag_active = True
            changed = True
        if app._pal_drag_active:
            hit = hovered_slot()
            if hit is not None and hit != app._pal_drag_insert:
                app._pal_drag_insert = hit
                changed = True
            if changed:
                _refresh_pal_edit_panel(app._pal_drag_idx, app._pal_drag_insert)

    if lmb_rel and app._pal_drag_idx is not None:
        source, destination, active = app._pal_drag_idx, app._pal_drag_insert, app._pal_drag_active
        _cancel_pal_drag()
        if active and destination is not None:
            _move_palette_color(source, destination)
        _refresh_pal_edit_panel()


# History selection
_hist_prev_selection: set = set()

def _update_history_selection_style():
    from .themes import _hist_selected_theme
    prev = _hist_prev_selection
    curr = app.selected_history_indices
    changed = prev.symmetric_difference(curr)
    for i in changed:
        if i < len(HIST_TAGS):
            tag = HIST_TAGS[i]
            if dpg.does_item_exist(tag):
                dpg.bind_item_theme(tag, _hist_selected_theme if i in curr else None)
    prev.clear(); prev.update(curr)

def _exit_palette_select_mode():
    app.palette_select_mode = False
    app.selected_history_indices.clear()
    app.selected_history_colors.clear()
    if dpg.does_item_exist("btn_new_palette"):
        dpg.set_item_label("btn_new_palette", t("new_palette"))
    dpg.bind_item_theme("btn_new_palette", None)
    if dpg.does_item_exist("btn_cancel_palette"):
        dpg.configure_item("btn_cancel_palette", show=False)
    _update_history_selection_style()

def save_selected_as_palette():
    if not app.selected_history_indices:
        return
    colors = [
        app.selected_history_colors[i]
        for i in sorted(app.selected_history_indices)
        if i in app.selected_history_colors
    ]
    if not colors:
        return
    if not can_add_palette():
        return
    with _palettes_lock:
        name = _auto_pal_name()
        app.palettes[name] = colors
        _register_palette(name)
    request_palette_save()
    _rebuild_pal_rows()
    _exit_palette_select_mode()

def toggle_palette_select_mode():
    if app.palette_select_mode:
        if app.selected_history_indices:
            save_selected_as_palette()
        else:
            _exit_palette_select_mode()
    else:
        from .themes import _new_palette_select_theme
        app.palette_select_mode = True
        app.selected_history_indices.clear()
        app.selected_history_colors.clear()
        dpg.set_item_label("btn_new_palette", t("save"))
        dpg.bind_item_theme("btn_new_palette", _new_palette_select_theme)
        if dpg.does_item_exist("btn_cancel_palette"):
            dpg.configure_item("btn_cancel_palette", show=True)
        _update_history_selection_style()

def history_click_cb(sender, app_data, user_data):
    i = user_data
    _do_style_update = False
    _apply_sel = False
    with _history_lock:
        if i >= len(app.history):
            return
        if app.palette_select_mode:
            if i in app.selected_history_indices:
                app.selected_history_indices.remove(i)
                app.selected_history_colors.pop(i, None)
            else:
                app.selected_history_indices.add(i)
                app.selected_history_colors[i] = list(app.history[i])
            _do_style_update = True
        else:
            col      = app.history[i]
            col_norm = [col[0] / 255., col[1] / 255., col[2] / 255., 1.]
            _force_update_color_picker(col_norm)
            from .sliders import _sync_all_modes
            _sync_all_modes()
            _update_wheel_contrast()
            _apply_sel = True
    if _do_style_update:
        _update_history_selection_style()
    if _apply_sel:
        _apply_external_color_to_pal()

def save_harmony_as_palette():
    """Tallentaa harmoniasävyt uutena paletina."""
    colors = []
    for tag in HARMONY_TAGS:
        if app.harmony_rgb.get(tag) is not None:
            r, g, b = app.harmony_rgb[tag]
            colors.append([r, g, b, 255])
    if not colors:
        return
    if not can_add_palette():
        return
    with _palettes_lock:
        name = _auto_pal_name()
        app.palettes[name] = colors
        _register_palette(name)
    request_palette_save()
    _rebuild_pal_rows()
