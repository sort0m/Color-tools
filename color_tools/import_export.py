"""
import_export.py — Palettien tuonti ja vienti (ASE, ACO, GPL, .swatches, .afpalette, kuvat, HTML).
"""

import os
import io
import zipfile
import colorsys
import traceback
import threading
import queue
import html as _html
import time as _time

import dearpygui.dearpygui as dpg

from .bootstrap import (
    _write_log, _win32_open_file, _win32_save_file,
    _PIL_AVAILABLE, _DPI_SCALE,
)
if _PIL_AVAILABLE:
    from PIL import Image as _PIL_Image
else:
    _PIL_Image = None

from .color_math import (
    rgb_to_cmyk, contrast_ratio, nearest_css, format_color,
)
from .i18n import t
from .state import app, _palettes_lock, _history_lock
from .storage import atomic_write
from .palette_formats import (
    parse_ase, parse_aco, parse_gpl, parse_swatches, parse_afpalette,
    _build_ase, _build_aco, _build_gpl, _build_swatches, _build_afpalette,
    MAX_COLORS as _MAX_IMPORT_COLORS, MAX_FILE_BYTES as _MAX_FILE_BYTES, MAX_ACO_COLORS,
)

# Tuetut kuvaformaatit
_IMAGE_EXTS = frozenset({
    ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tiff", ".tif",
})

# Tuonnin rajat (suojaa jäätymiseltä ja liian suurilta tiedostoilta)
_MAX_IMAGE_BYTES    = 100 * 1024 * 1024   # kuvatiedosto levyllä
_MAX_IMAGE_PIXELS   = 60_000_000          # purettavan kuvan pikselimäärä (JPEG: draft-pienennyksen jälkeen)

# ── Kuvan kvantisointi ────────────────────────────────────────────────────────
def _quantize_image(pil_img, n):
    """Palauttaa lajitellun listan [R,G,B,255]-väreistä PIL-kvantisoinnista."""
    try:
        method = _PIL_Image.Quantize.FASTOCTREE
    except AttributeError:
        method = getattr(_PIL_Image, 'FASTOCTREE', 2)
    quantized    = pil_img.quantize(colors=n, method=method)
    palette_data = quantized.getpalette() or []
    # Käytetään vain oikeasti käytetyt paletti-indeksit (ei mustia täytevärejä)
    try:
        used = sorted({idx for _cnt, idx in (quantized.getcolors(256) or [])})
    except Exception:
        used = []
    if not used:
        used = list(range(min(n, len(palette_data) // 3)))
    colors = [
        [palette_data[i*3], palette_data[i*3+1], palette_data[i*3+2], 255]
        for i in used
        if i * 3 + 2 < len(palette_data)
    ]
    quantized.close()
    colors.sort(key=lambda c: 0.2126*c[0] + 0.7152*c[1] + 0.0722*c[2])
    return colors

_IMG_PREVIEW_COLS = 10

def _commit_image_palette(colors):
    """Commit imported colors and reset history selection on the UI thread."""
    from .palette import _auto_pal_name, _register_palette, can_add_palette, _exit_palette_select_mode
    if not can_add_palette():
        return
    if app.palette_select_mode:
        _exit_palette_select_mode()
    with _history_lock:
        hist_set = {tuple(x) for x in app.history}
        for c in reversed(colors):
            key = tuple(c)
            if key not in hist_set:
                app.history.appendleft(c)
                hist_set.add(key)
        app._hist_dirty = True
    with _palettes_lock:
        name = _auto_pal_name()
        app.palettes[name] = list(colors)
        _register_palette(name)
    app._pending_save = True
    app._pal_dirty = True

def _rebuild_image_import_preview():
    """Piirtää väriswatch:t tuontipaneeliin."""
    if not dpg.does_item_exist("img_import_preview"):
        return
    dpg.delete_item("img_import_preview", children_only=True)
    from .palette import _nospc_theme
    PAL_SW = max(1, int(24 * _DPI_SCALE))
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
    # Source images have already been reduced to <= 200 x 200 pixels.
    # Quantize this small image synchronously in the UI callback so OK, close
    # and a new image can never overtake a result from an older session.
    if not app._image_modal_open or app._image_modal_pil is None:
        return
    try:
        app._image_modal_colors = _quantize_image(app._image_modal_pil, max(2, min(30, int(v))))
    except Exception:
        app._image_modal_colors = []
        _write_log('Image quantization failed:\n' + traceback.format_exc())
        _set_status(t('imp_failed'))
    if dpg.does_item_exist('btn_img_ok'):
        dpg.configure_item('btn_img_ok', enabled=bool(app._image_modal_colors))
    _rebuild_image_import_preview()

def _image_import_ok(s, a, u):
    if not app._image_modal_open or not app._image_modal_colors:
        return
    colors = list(app._image_modal_colors)
    _image_import_close()
    if colors:
        _commit_image_palette(colors)

def _image_import_cancel(s, a, u):
    _image_import_close()

def _image_import_close():
    app._image_modal_open   = False
    if app._image_modal_pil is not None:
        app._image_modal_pil.close()
    app._image_modal_pil    = None
    app._image_modal_colors = []
    if dpg.does_item_exist("img_import_win"):
        dpg.configure_item("img_import_win", show=False)
    if dpg.does_item_exist("pal_import_buttons"):
        dpg.configure_item("pal_import_buttons", show=True)
    if dpg.does_item_exist("pal_scroll"):
        dpg.configure_item("pal_scroll", show=not bool(app._editing_pal))
    if dpg.does_item_exist("pal_edit_win"):
        dpg.configure_item("pal_edit_win", show=bool(app._editing_pal))

def _open_image_import_panel(pil_img):
    """Näyttää värimäärän valintapaneelin. Vain pääsäikeestä."""
    if app._image_modal_open:
        pil_img.close()
        return
    app._image_modal_open   = True
    app._image_modal_pil    = pil_img
    _image_import_slider_cb(None, 8)
    if dpg.does_item_exist("pal_import_buttons"):
        dpg.configure_item("pal_import_buttons", show=False)
    if dpg.does_item_exist("pal_scroll"):
        dpg.configure_item("pal_scroll",   show=False)
    if dpg.does_item_exist("pal_edit_win"):
        dpg.configure_item("pal_edit_win", show=False)
    if dpg.does_item_exist("img_import_win"):
        dpg.configure_item("img_import_win", show=True)
    if dpg.does_item_exist("img_import_slider"):
        dpg.set_value("img_import_slider", 8)
    _rebuild_image_import_preview()

def _set_status(msg, secs=5.0, *, code=None):
    app._import_status = (msg, _time.perf_counter() + secs)
    app._import_status_code = code

_IMAGE_FORMATS = ('PNG', 'JPEG', 'BMP', 'GIF', 'WEBP', 'TIFF')
_import_active = threading.BoundedSemaphore(1)
_import_results = queue.SimpleQueue()
_dialog_active = threading.BoundedSemaphore(1)
_export_write_lock = threading.Lock()
_io_stopping = threading.Event()


def _read_image(path):
    # Select the parser by an explicit allowlist, not just a filename suffix.
    with _PIL_Image.open(path, formats=_IMAGE_FORMATS) as source:
        source.draft('RGB', (800, 800))
        if source.width * source.height > _MAX_IMAGE_PIXELS:
            raise ValueError('image exceeds pixel limit')
        source.thumbnail((1000, 1000))
        if (source.mode in ('RGBA', 'LA')
                or source.mode == 'P' and 'transparency' in source.info):
            with source.convert('RGBA') as rgba:
                with _PIL_Image.new('RGBA', rgba.size, (255, 255, 255, 255)) as bg:
                    with _PIL_Image.alpha_composite(bg, rgba) as composite:
                        image = composite.convert('RGB')
        else:
            image = source.convert('RGB')
        image.thumbnail((200, 200))
        return image


def _import_from_path(path):
    """At most one import, including its result waiting for the UI."""
    if (_io_stopping.is_set() or app._image_modal_open or app._drop_preview_open
            or not _import_active.acquire(blocking=False)):
        _set_status(t('imp_busy'))
        return
    ext = os.path.splitext(path)[1].lower()

    def _run():
        result = ('error', t('imp_failed'))
        try:
            size_limit = _MAX_IMAGE_BYTES if ext in _IMAGE_EXTS else _MAX_FILE_BYTES
            if os.path.getsize(path) > size_limit:
                result = ('error', t('imp_too_large'))
            elif ext in _IMAGE_EXTS:
                result = (('image', _read_image(path)) if _PIL_AVAILABLE
                          else ('error', t('imp_no_pillow')))
            elif ext in ('.ase', '.aco', '.gpl', '.swatches', '.afpalette'):
                # Bounded read also covers a file growing after the size check.
                with open(path, 'rb') as stream:
                    data = stream.read(_MAX_FILE_BYTES + 1)
                if len(data) > _MAX_FILE_BYTES:
                    raise ValueError('palette exceeds size limit')
                skipped = 0
                if ext == '.aco':
                    colors, capped, skipped = parse_aco(data)
                elif ext == '.gpl':
                    colors, capped = parse_gpl(io.StringIO(data.decode('utf-8-sig', errors='replace')))
                elif ext == '.swatches':
                    colors, capped = parse_swatches(io.BytesIO(data))
                else:
                    parser = {'.ase': parse_ase, '.afpalette': parse_afpalette}[ext]
                    colors, capped = parser(data)
                result = (('palette', colors, os.path.basename(path), capped, skipped) if colors
                          else ('error', t('imp_aco_none').format(n=skipped) if skipped
                                else t('imp_none').format(fmt=ext)))
            else:
                result = ('error', t('imp_unsupported').format(ext=ext))
        except Exception:
            _write_log('Import failed:\n' + traceback.format_exc())
        finally:
            _import_results.put(result)
    try:
        threading.Thread(target=_run, daemon=True).start()
    except Exception:
        _import_active.release()
        raise


def process_import_results():
    """Consume background results on the UI thread only."""
    try:
        result = _import_results.get_nowait()
    except queue.Empty:
        return
    try:
        if result[0] == 'image':
            _navigate_to_palettes()
            _open_image_import_panel(result[1])
        elif result[0] == 'palette':
            _, colors, name, capped, skipped = result
            app._drop_preview_colors_pending = colors
            _open_drop_preview(colors, name, capped=capped, skipped=skipped)
            if skipped or capped:
                _set_status(_import_warning(capped, skipped), secs=12.0)
        else:
            _set_status(result[1])
    finally:
        _import_active.release()


def _start_dialog(work):
    # Avoid a pile of unowned native dialogs and export threads.
    if _io_stopping.is_set() or not _dialog_active.acquire(blocking=False):
        _set_status(t('imp_busy'))
        return
    def run():
        try:
            work()
        except Exception:
            _write_log('File dialog failed:\n' + traceback.format_exc())
            _set_status(t('export_failed'), code='export_failed')
        finally:
            _dialog_active.release()
    try:
        threading.Thread(target=run, daemon=True).start()
    except Exception:
        _dialog_active.release()
        raise


def _write_export(path, data):
    with _export_write_lock:
        if _io_stopping.is_set():
            return False
        atomic_write(path, data)
    _set_status(t('export_saved'))
    return True


def shutdown_io():
    # Wait for an active atomic write. Dialogs still open are cancelled by exit;
    # their workers must not start a new write after this point.
    with _export_write_lock:
        _io_stopping.set()


def _import_open_dialog(s, a, u):
    """Avaa tiedostonvalintadialogin palettien tuontia varten."""
    def _run():
        path = _win32_open_file(
            t("dlg_import_title"),
            [
                (t("dlg_filter_all_supported"),
                 "*.ase;*.aco;*.gpl;*.swatches;*.afpalette;"
                 "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp;*.tiff;*.tif"),
                (t("dlg_filter_palette_files"), "*.ase;*.aco;*.gpl;*.swatches;*.afpalette"),
                (t("dlg_filter_affinity"), "*.afpalette"),
                (t("dlg_filter_images"),
                 "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp;*.tiff;*.tif"),
                (t("dlg_filter_all_files"), "*.*"),
            ],
        )
        if path:
            _import_from_path(path)
    _start_dialog(_run)

def _navigate_to_palettes():
    if dpg.does_item_exist("bottom_tab_bar") and dpg.does_item_exist("tab_palettes"):
        dpg.set_value("bottom_tab_bar", "tab_palettes")

def _import_warning(capped, skipped):
    messages = []
    if skipped:
        messages.append(t('imp_aco_skipped').format(n=skipped))
    if capped:
        messages.append(t('imp_capped').format(n=_MAX_IMPORT_COLORS))
    return '\n'.join(messages)


def refresh_import_warning():
    if dpg.does_item_exist('drop_preview_warning'):
        text = _import_warning(*app._drop_preview_warning)
        dpg.set_value('drop_preview_warning', text)
        dpg.configure_item('drop_preview_warning', show=bool(text))


def _open_drop_preview(colors, fname, *, capped=False, skipped=0):
    app._drop_preview_warning = (capped, skipped)
    refresh_import_warning()
    app._drop_preview_open = True
    if app._image_modal_open:       # sulje avoin kuvatuontipaneeli, ettei lippu jää päälle
        _image_import_close()
    _navigate_to_palettes()
    if not dpg.does_item_exist("drop_preview_win"):
        return
    from .palette import _nospc_theme
    PAL_SW = max(1, int(24 * _DPI_SCALE))
    if dpg.does_item_exist("drop_preview_colors"):
        dpg.delete_item("drop_preview_colors", children_only=True)
        for row_start in range(0, len(colors), _IMG_PREVIEW_COLS):
            grp = dpg.add_group(parent="drop_preview_colors", horizontal=True)
            dpg.bind_item_theme(grp, _nospc_theme)
            for c in colors[row_start: row_start + _IMG_PREVIEW_COLS]:
                dpg.add_color_button(
                    parent=grp, default_value=c,
                    width=PAL_SW, height=PAL_SW,
                    no_alpha=True, no_border=True,
                )
    if dpg.does_item_exist("drop_preview_label"):
        dpg.configure_item("drop_preview_label",
                           default_value=f"{fname}  ({len(colors)} colors)")
    for tag in ("pal_scroll", "pal_edit_win", "img_import_win"):
        if dpg.does_item_exist(tag):
            dpg.configure_item(tag, show=False)
    if dpg.does_item_exist("pal_import_buttons"):
        dpg.configure_item("pal_import_buttons", show=False)
    dpg.configure_item("drop_preview_win", show=True)

def _close_drop_preview():
    app._drop_preview_open = False
    if dpg.does_item_exist("drop_preview_win"):
        dpg.configure_item("drop_preview_win", show=False)
    if dpg.does_item_exist("pal_import_buttons"):
        dpg.configure_item("pal_import_buttons", show=True)
    editing = bool(app._editing_pal)
    if dpg.does_item_exist("pal_scroll"):
        dpg.configure_item("pal_scroll", show=not editing)
    if dpg.does_item_exist("pal_edit_win"):
        dpg.configure_item("pal_edit_win", show=editing)

# ── Vientiformaatin valinta ───────────────────────────────────────────────────
def _show_export_format(s=None, a=None, u=None):
    if dpg.does_item_exist("pal_export_format_panel"):
        dpg.configure_item("pal_export_format_panel", show=True)
    if dpg.does_item_exist("pal_edit_main"):
        dpg.configure_item("pal_edit_main", show=False)

def _hide_export_format(s=None, a=None, u=None):
    if dpg.does_item_exist("pal_export_format_panel"):
        dpg.configure_item("pal_export_format_panel", show=False)
    if dpg.does_item_exist("pal_edit_main"):
        dpg.configure_item("pal_edit_main", show=True)

def _execute_palette_export(fmt):
    if not app._editing_pal or app._editing_pal not in app.palettes:
        _hide_export_format()
        return
    with _palettes_lock:
        colors = list(app.palettes[app._editing_pal])
    if fmt == "HTML":
        h = _html_begin()
        for v in colors:
            h += _html_color_row(v[0], v[1], v[2])
        h += "</table></body></html>"
        _save_html_dialog(h, app._editing_pal)
    elif fmt == "ASE":
        _save_ase_dialog(colors, app._editing_pal)
    elif fmt == "ACO":
        _save_aco_dialog(colors, app._editing_pal)
    elif fmt == "GPL":
        _save_gpl_dialog(colors, app._editing_pal)
    elif fmt == "Procreate":
        _save_swatches_dialog(colors, app._editing_pal)
    elif fmt == "Affinity":
        _save_afpalette_dialog(colors, app._editing_pal)
    _hide_export_format()

# ── HTML-vienti ───────────────────────────────────────────────────────────────
_HCSS = (
    "<script>"
    "function cc(t,b){"
    "navigator.clipboard.writeText(t).then(()=>{"
    "var o=b.innerText;b.innerText='HTMLCOPIED_PLACEHOLDER';b.classList.add('copied');"
    "setTimeout(()=>{b.innerText=o;b.classList.remove('copied')},1200)})}"
    "function tg(b){"
    "var td=b.parentNode.parentNode;"
    "var s=td.querySelector('.th-s');"
    "var f=td.querySelector('.th-f');"
    "if(f.style.display==='none'){f.style.display='block';s.style.display='none';}"
    "else{f.style.display='none';s.style.display='block';}}"
    "</script>"
    "<style>"
    "HTMLFONTS_PLACEHOLDER"
    "HTMLTHEME_PLACEHOLDER"
    "body{font-family:var(--ui-font);font-size:var(--ui-size);background:var(--page);color:var(--text);padding:30px;margin:0}"
    "button{font-family:var(--ui-font);font-size:inherit}"
    "h2{margin-bottom:18px;font-weight:600;letter-spacing:.03em}"
    "table{border-collapse:collapse;background:var(--surface);white-space:nowrap}"
    "td,th{border:1px solid var(--border);padding:7px 10px;font-size:inherit;vertical-align:middle}"
    "th{background:var(--header);color:var(--text);font-weight:600;text-transform:uppercase;"
    "font-size:var(--ui-size);letter-spacing:.06em;white-space:nowrap}"
    "tr:hover td{background:var(--hover)}"
    ".sw{width:52px;height:30px;border-radius:5px;border:1px solid var(--border);display:block}"
    ".cb{cursor:pointer;padding:3px 10px;background:var(--button);color:var(--text);"
    "border:1px solid var(--border);border-radius:4px;font-size:var(--ui-size);display:block;width:100%;"
    "box-sizing:border-box;margin-top:2px}"
    ".cb:first-child{margin-top:0}"
    ".cb:hover{background:var(--button-hover)}"
    ".cb.copied{background:var(--button-active)}"
    "a{color:var(--link)}"
    ".color-value{font-family:var(--ui-font);font-size:var(--ui-size);font-weight:400;font-variant-numeric:tabular-nums}"
    ".wcag-aaa{color:var(--aaa);font-weight:700}"
    ".wcag-aa{color:var(--aa);font-weight:700}"
    ".wcag-aa-lg{color:var(--aa-large)}"
    ".wcag-fail{color:var(--fail)}"
    ".ctr-row{display:flex;align-items:center;gap:6px;margin-bottom:3px}"
    ".ctr-bar{display:inline-block;padding:1px 8px;border-radius:3px;"
    "font-size:var(--ui-size);font-weight:600;white-space:nowrap}"
    ".ctr-bar-w{background:#ffffff;border:1px solid #555}"
    ".ctr-bar-b{background:#000000;border:1px solid #555}"
    "</style>"
)

def _html_copy_btn(val):
    return f"<button class=\"cb\" onclick=\"cc('{val}',this)\">{t('html_copy')}</button>"

def _html_color_row(r, g, b):
    hx  = f"#{r:02x}{g:02x}{b:02x}".upper()
    rgb = f"rgb({r}, {g}, {b})"
    h, l, s   = colorsys.rgb_to_hls(r/255, g/255, b/255)
    hsl = f"hsl({h*360:.0f}, {s*100:.0f}%, {l*100:.0f}%)"
    h2, s2, v2 = colorsys.rgb_to_hsv(r/255, g/255, b/255)
    hsv = f"hsv({h2*360:.0f}, {s2*100:.0f}%, {v2*100:.0f}%)"
    c_, m_, y_, k_ = rgb_to_cmyk(r, g, b)
    cmyk = f"C{c_:.0f}% M{m_:.0f}% Y{y_:.0f}% K{k_:.0f}%"
    css = nearest_css(r, g, b)
    oklch = format_color((r, g, b), "OKLCH")
    cw = contrast_ratio(r, g, b, 255, 255, 255)
    cb = contrast_ratio(r, g, b, 0,   0,   0  )
    def _wcag_cls(ratio):
        if ratio >= 7:   return "wcag-aaa", "AAA"
        if ratio >= 4.5: return "wcag-aa",  "AA"
        if ratio >= 3:   return "wcag-aa-lg","AA*"
        return "wcag-fail", t("wcag_level_fail")
    cw_cls, cw_lbl = _wcag_cls(cw)
    cb_cls, cb_lbl = _wcag_cls(cb)
    from .harmony import color_theory_info
    info = color_theory_info((r, g, b))
    theory_html = _html.escape(info["full"], quote=False).replace("\n", "<br>")
    short_text  = _html.escape(info["association"], quote=False)

    theory_cell = (
        f"<div class=\"th-s\" style=\"display:block;\">{short_text} "
        f"<a href=\"javascript:void(0)\" onclick=\"tg(this)\" style=\"text-decoration:none;font-weight:bold;\">&gt;&gt;</a></div>"
        f"<div class=\"th-f\" style=\"display:none;\">{theory_html} <br>"
        f"<a href=\"javascript:void(0)\" onclick=\"tg(this)\" style=\"text-decoration:none;font-weight:bold;\">&lt;&lt;</a></div>"
    )

    return (
        f"<tr>"
        f"<td><span class=\"sw\" style=\"background:{hx}\"></span></td>"
        f"<td class=\"color-value\">{hx} {_html_copy_btn(hx)}</td>"
        f"<td class=\"color-value\">{rgb}</td>"
        f"<td class=\"color-value\">{hsl}</td>"
        f"<td class=\"color-value\">{hsv}</td>"
        f"<td class=\"color-value\">{cmyk}</td>"
        f"<td class=\"color-value\">{oklch}</td>"
        f"<td>{css}</td>"
        f"<td>"
        f"<div class=\"ctr-row\"><span class=\"{cw_cls}\">W &nbsp;{cw:.2f}:1 &nbsp;{cw_lbl}</span>"
        f"<span class=\"ctr-bar ctr-bar-w\" style=\"color:{hx}\">Text</span></div>"
        f"<div class=\"ctr-row\"><span class=\"{cb_cls}\">B &nbsp;{cb:.2f}:1 &nbsp;{cb_lbl}</span>"
        f"<span class=\"ctr-bar ctr-bar-b\" style=\"color:{hx}\">Text</span></div>"
        f"</td>"
        f"<td style=\"max-width:250px; white-space:normal; line-height:1.4;\">{theory_cell}</td>"
        f"</tr>"
    )

def _html_table_header():
    return (
        f"<tr><th>{t('html_th_preview')}</th><th>HEX</th><th>RGB</th>"
        f"<th>HSL</th><th>HSV</th><th>CMYK</th><th>OKLCH</th>"
        f"<th>{t('html_th_cssname')}</th><th>{t('html_th_contrast')}</th>"
        f"<th>{t('html_th_theory')}</th></tr>"
    )

def _html_theme_css():
    """Snapshot the selected application's actual, contrast-adjusted colors."""
    from .themes import _ACCESSIBLE_PALETTES
    from .accessibility import contrast_color
    fallback = {
        'Text': [245, 245, 245], 'WindowBg': [26, 26, 26],
        'ChildBg': [26, 26, 26], 'TableHeaderBg': [34, 34, 34],
        'HeaderHovered': [34, 34, 34], 'Button': [46, 46, 46],
        'ButtonHovered': [61, 61, 61], 'ButtonActive': [61, 61, 61],
        'FrameBg': [46, 46, 46],
    }
    palette = _ACCESSIBLE_PALETTES.get(app.theme_name, _ACCESSIBLE_PALETTES.get('Dark', fallback))
    colors = {name: palette[key] for name, key in {
        'text': 'Text', 'page': 'WindowBg', 'surface': 'ChildBg',
        'header': 'TableHeaderBg', 'hover': 'HeaderHovered',
        'button': 'Button', 'button-hover': 'ButtonHovered',
        'button-active': 'ButtonActive', 'border': 'FrameBg',
    }.items()}
    backgrounds = [colors['surface'], colors['hover']]
    for name, preferred in {
        'link': [128, 199, 255], 'aaa': [165, 224, 165],
        'aa': [139, 195, 74], 'aa-large': [205, 220, 57],
        'fail': [255, 170, 170],
    }.items():
        colors[name] = contrast_color(preferred, backgrounds)
    return ':root{' + ''.join(
        f'--{name}:#{int(round(rgb[0])):02x}{int(round(rgb[1])):02x}{int(round(rgb[2])):02x};'
        for name, rgb in colors.items()
    ) + '}'


def _html_begin():
    """HTML-dokumentin alku (otsikkopaikat täytetään tallennettaessa)."""
    from . import fonts
    font_css = (
        f":root{{--ui-font:'{fonts.UI_FONT_FAMILY}',Tahoma,Arial,sans-serif;"
        f"--ui-size:{fonts.HTML_UI_FONT_SIZE}px;}}"
    )
    css = (_HCSS.replace('HTMLCOPIED_PLACEHOLDER', t("html_copied"))
           .replace('HTMLFONTS_PLACEHOLDER', font_css)
           .replace('HTMLTHEME_PLACEHOLDER', _html_theme_css()))
    return (
        f"<!doctype html><html lang='{t('html_lang')}'><head><meta charset='utf-8'>HTMLTITLE_PLACEHOLDER{css}</head><body>"
        f"<h2>HTMLH2_PLACEHOLDER</h2>"
        f"<table>{_html_table_header()}"
    )

def _save_html_dialog(content, default_name):
    def _run():
        try:
            path = _win32_save_file(
                t("dlg_save_as"),
                [(t("dlg_html_file"), "*.html"), (t("dlg_filter_all_files"), "*.*")],
                default_ext=".html",
                initial_name=default_name,
            )
            if path:
                title = os.path.splitext(os.path.basename(path))[0]
                safe_title = _html.escape(title)
                final = content.replace(
                    'HTMLTITLE_PLACEHOLDER',
                    f'<title>{safe_title}</title>'
                ).replace('HTMLH2_PLACEHOLDER', safe_title)
                saved = _write_export(path, final)
                if saved and path.lower().endswith(('.html', '.htm')):   # älä aja muuta kuin HTML:ää
                    try:
                        os.startfile(path)
                    except OSError as e:
                        _write_log(f"HTML launch error: {e}")
                        _set_status(t("export_open_failed"), code="export_open_failed")
        except Exception as e:
            _write_log(f"Save error: {e}\n" + traceback.format_exc())
            _set_status(t('export_failed'), code='export_failed')
    _start_dialog(_run)

def export_palette():
    from .harmony import HARMONY_TAGS
    h = _html_begin()
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
    from .state import _history_lock
    with _history_lock:
        hist_snapshot = list(app.history)
    if not hist_snapshot:
        return
    h = _html_begin()
    for v in hist_snapshot:
        h += _html_color_row(v[0], v[1], v[2])
    h += "</table></body></html>"
    _save_html_dialog(h, "color_history")

# ── ASE-vienti ────────────────────────────────────────────────────────────────
def _save_ase_dialog(colors, default_name):
    def _run():
        try:
            path = _win32_save_file(t("dlg_export_title").format(fmt="ASE"),
                [("Adobe Swatch Exchange", "*.ase"), (t("dlg_filter_all_files"), "*.*")],
                default_ext=".ase", initial_name=default_name)
            if path:
                _write_export(path, _build_ase(colors))
        except Exception as e:
            _write_log(f"ASE export error: {e}\n" + traceback.format_exc())
            _set_status(t('export_failed'), code='export_failed')
    _start_dialog(_run)

# ── ACO-vienti ────────────────────────────────────────────────────────────────
def _save_aco_dialog(colors, default_name):
    if len(colors) > MAX_ACO_COLORS:
        _set_status(t("export_aco_limit"), code="export_aco_limit")
        return
    def _run():
        try:
            path = _win32_save_file(t("dlg_export_title").format(fmt="ACO"),
                [("Adobe Color", "*.aco"), (t("dlg_filter_all_files"), "*.*")],
                default_ext=".aco", initial_name=default_name)
            if path:
                _write_export(path, _build_aco(colors))
        except Exception as e:
            _write_log(f"ACO export error: {e}\n" + traceback.format_exc())
            _set_status(t('export_failed'), code='export_failed')
    _start_dialog(_run)

# ── GPL-vienti ────────────────────────────────────────────────────────────────
def _save_gpl_dialog(colors, default_name):
    def _run():
        try:
            path = _win32_save_file(t("dlg_export_title").format(fmt="GPL"),
                [("GIMP Palette", "*.gpl"), (t("dlg_filter_all_files"), "*.*")],
                default_ext=".gpl", initial_name=default_name)
            if path:
                pal_name = os.path.splitext(os.path.basename(path))[0]
                _write_export(path, _build_gpl(pal_name, colors))
        except Exception as e:
            _write_log(f"GPL export error: {e}\n" + traceback.format_exc())
            _set_status(t('export_failed'), code='export_failed')
    _start_dialog(_run)

# ── Procreate .swatches -vienti ───────────────────────────────────────────────
def _save_swatches_dialog(colors, default_name):
    def _run():
        try:
            path = _win32_save_file(t("dlg_export_title").format(fmt="Procreate"),
                [("Procreate Swatches", "*.swatches"), (t("dlg_filter_all_files"), "*.*")],
                default_ext=".swatches", initial_name=default_name)
            if path:
                pal_name = os.path.splitext(os.path.basename(path))[0]
                json_data = _build_swatches(pal_name, colors)
                buf = io.BytesIO()
                with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                    zf.writestr("Swatches.json", json_data)
                _write_export(path, buf.getvalue())
        except Exception as e:
            _write_log(f"Swatches export error: {e}\n" + traceback.format_exc())
            _set_status(t('export_failed'), code='export_failed')
    _start_dialog(_run)

# ── Affinity .afpalette -vienti ───────────────────────────────────────────────
def _save_afpalette_dialog(colors, default_name):

    def _run():
        try:
            path = _win32_save_file(
                t("dlg_export_title").format(fmt="Affinity"),
                [("Affinity Palette", "*.afpalette"), (t("dlg_filter_all_files"), "*.*")],
                default_ext=".afpalette",
                initial_name=default_name,
            )
            if path:
                pal_name = os.path.splitext(os.path.basename(path))[0]
                data = _build_afpalette(pal_name, colors)
                _write_export(path, data)
        except Exception as e:
            _write_log(f"Affinity palette export error: {e}\n" + traceback.format_exc())
            _set_status(t('export_failed'), code='export_failed')
    _start_dialog(_run)
