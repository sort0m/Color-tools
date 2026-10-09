"""
sliders.py — Liukusäätimien data, gradientti-laskelmat, moodi-konversiot.
"""

import colorsys
import time as _time

import dearpygui.dearpygui as dpg

from .bootstrap import _sc
from .color_math import (
    rgb_to_cmyk, cmyk_to_rgb, rgb_to_lab, lab_to_rgb, color_sig,
)
from .state import app

# ── Liukusäätimien asetukset ──────────────────────────────────────────────────
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

# Esilaketut DPG-tagit drawlist-kohteille
_DL_TAGS = {gname: f"dl_{gname}"
            for sliders in SLIDERS_BY_MODE.values()
            for _, gname, *_ in sliders}

# Esilaketut (smin, smax) alueet jokaiselle gname:lle
_SLIDER_RANGE = {gname: (smin, smax)
                 for sliders in SLIDERS_BY_MODE.values()
                 for (_, gname, _, smin, smax) in sliders}

# Esilaketut moodi-haut: mikä moodi omistaa gname:n
_GNAME_MODE = {gname: mode
               for mode, sliders in SLIDERS_BY_MODE.items()
               for (_, gname, *_) in sliders}

# Gradienttiliukusäätimien koot — importoidaan ui_main:ssa asetetuista layoutvakioista
# Nämä asetetaan myöhemmin ui_main.py:n toimesta kutsumalla init_layout_constants()
GRAD_W  = 0
GRAD_H  = 0
THUMB_H = 0
SEGS    = 48
_GRAD_SEG_X = []

def init_layout_constants(grad_w, grad_h, thumb_h):
    """Asetetaan ui_main.py:stä heti sen jälkeen kun layoutvakiot on laskettu."""
    global GRAD_W, GRAD_H, THUMB_H
    GRAD_W  = grad_w
    GRAD_H  = grad_h
    THUMB_H = thumb_h
    sw_unit = GRAD_W / SEGS
    _GRAD_SEG_X[:] = [
        (int(i * sw_unit), max(int(i * sw_unit) + 1, int((i + 1) * sw_unit)))
        for i in range(SEGS)
    ]

# Välimuisti: {gname: {'seg_ids': [...], 'indicator_id': int, 'y0': int, 'y1': int}}
_GRAD_ITEM_CACHE: dict = {}

# ── Gradienttilaskenta ────────────────────────────────────────────────────────
_grad_ctx_cache = [None, None]

def _grad_ctx():
    """Laskee perusväriarvot kerran per slider-piirto. Välimuistitettu."""
    vsig = color_sig(app.current_color)
    if _grad_ctx_cache[0] == vsig:
        return _grad_ctx_cache[1]
    r, g, b    = vsig
    rf, gf, bf = r / 255., g / 255., b / 255.
    h_, l_, s_ = colorsys.rgb_to_hls(rf, gf, bf)
    c0, m0, y0, k0 = rgb_to_cmyk(r, g, b)
    L0, a0, b0_    = rgb_to_lab(r, g, b)
    result = rf, gf, bf, h_, l_, s_, c0, m0, y0, k0, L0, a0, b0_
    _grad_ctx_cache[0] = vsig
    _grad_ctx_cache[1] = result
    return result

def _clamp_rgb(a, b2, c):
    return (
        max(0, min(255, int(round(a * 255)))),
        max(0, min(255, int(round(b2 * 255)))),
        max(0, min(255, int(round(c * 255)))),
    )

def _grad_color(name, t, ctx):
    """Palauttaa RGB-tuplan gradienttisegmentille t."""
    rf, gf, bf, h_, l_, s_, c0, m0, y0, k0, L0, a0, b0_ = ctx
    if   name == "R":    return _clamp_rgb(t, gf, bf)
    elif name == "G":    return _clamp_rgb(rf, t, bf)
    elif name == "B":    return _clamp_rgb(rf, gf, t)
    elif name == "H":
        return _clamp_rgb(*colorsys.hls_to_rgb(
            t,
            l_ if .05 < l_ < .95 else .5,
            s_ if s_ > .05 else 1.,
        ))
    elif name == "S":    return _clamp_rgb(*colorsys.hls_to_rgb(h_, l_, t))
    elif name == "L":    return _clamp_rgb(*colorsys.hls_to_rgb(h_, t, s_))
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

# ── Moodi-konversiot ──────────────────────────────────────────────────────────
def _precise_mode_values(r, g, b):
    """Kaikkien liukusäädinten tarkat (pyöristämättömät) arvot RGB-värille."""
    rf, gf, bf = r / 255., g / 255., b / 255.
    h, l, s = colorsys.rgb_to_hls(rf, gf, bf)
    c, m, y, k = rgb_to_cmyk(r, g, b)
    L, a, b_ = rgb_to_lab(r, g, b)
    return {
        "R": float(r), "G": float(g), "B": float(b),
        "H": h * 360., "S": s * 100., "L": l * 100.,
        "C": c, "M": m, "Yv": y, "K": k,
        "LA": L, "La": a, "Lb": b_,
        "Gray": .2126 * r + .7152 * g + .0722 * b,
    }

def _display_value(gname, value):
    """Pyöristää tarkan arvon liukusäätimen kokonaislukunäyttöä varten."""
    v = int(round(value))
    if gname == "H":
        return v % 360
    smin, smax = _SLIDER_RANGE[gname]
    return max(smin, min(smax, v))

def _fill_modes(r, g, b, skip=()):
    """Päivittää liukusäädinten näyttöarvot (app.sl_vals) värin r, g, b mukaan.
    skip: liukusäätimien nimet (gname), joiden arvoon ei kosketa."""
    sv = app.sl_vals
    for gname, value in _precise_mode_values(r, g, b).items():
        if gname not in skip:
            sv[gname] = _display_value(gname, value)

# Viimeisimmän moodisovelluksen "aikomus": tarkat liukuluvut, joista väri laskettiin.
# Ilman tätä 8-bittinen kvantisointi kertyisi koskemattomiin säätimiin (esim. H)
# aina kun toista säädintä liikutetaan. Ankkuri mitätöityy kun väri muuttuu muualta.
_mode_anchor = {"sig": None, "mode": None, "vals": {}}

def _apply_from_mode(mode):
    """Soveltaa liukusäätimien arvot nykyiseksi väriksi.

    sl_vals sisältää vain pyöristettyjä näyttöarvoja. Jotta pyöristys ei
    kerry väriin, säätimet joihin käyttäjä ei koskenut luetaan tarkkana
    nykyisestä värestä, ja vain muutettu säädin otetaan sl_vals:sta."""
    if mode not in SLIDERS_BY_MODE:
        return
    sv = app.sl_vals
    cr, cg, cb = color_sig(app.current_color)
    precise = _precise_mode_values(cr, cg, cb)
    if (_mode_anchor["sig"] == (cr, cg, cb) and _mode_anchor["mode"] == mode):
        precise.update(_mode_anchor["vals"])

    val = {}
    changed = set()
    for (_, gname, *_rest) in SLIDERS_BY_MODE[mode]:
        if sv[gname] == _display_value(gname, precise[gname]):
            val[gname] = precise[gname]      # koskematon -> tarkka arvo
        else:
            val[gname] = sv[gname]           # käyttäjän asettama arvo
            changed.add(gname)

    if mode == "RGB":
        r, g, b = _clamp_rgb(val["R"] / 255., val["G"] / 255., val["B"] / 255.)
    elif mode == "HSL":
        r, g, b = _clamp_rgb(*colorsys.hls_to_rgb(
            val["H"] / 360., val["L"] / 100., val["S"] / 100.))
    elif mode == "CMYK":
        r, g, b = cmyk_to_rgb(val["C"], val["M"], val["Yv"], val["K"])
    elif mode == "LAB":
        r, g, b = lab_to_rgb(val["LA"], val["La"], val["Lb"])
    else:  # Grayscale
        r = g = b = max(0, min(255, int(round(val["Gray"]))))

    col = [r / 255., g / 255., b / 255., 1.]
    app.current_color = col
    app.change_time   = _time.perf_counter()
    app._picker_processing = True
    try:
        dpg.set_value("picker_wheel", [r, g, b, 255])
    finally:
        app._picker_processing = False

    if mode == "LAB":
        # LAB:ssa arvot palautetaan todellisiksi (gamut-rajaus)
        _fill_modes(r, g, b)
        _mode_anchor["sig"] = None
    else:
        own = {s_[1] for s_ in SLIDERS_BY_MODE[mode]}
        _fill_modes(r, g, b, skip=own)   # muut moodit uuden värin mukaan
        for gname in own:                # oman moodin säätimet: käyttäjän arvo / ankkuri
            if gname not in changed:
                sv[gname] = _display_value(gname, val[gname])
        _mode_anchor.update(sig=color_sig(col), mode=mode, vals=dict(val))

def _sync_all_modes():
    """Synkronoi kaikki moodit nykyiseen väriin."""
    r, g, b = color_sig(app.current_color)
    _fill_modes(r, g, b)
    _mode_anchor["sig"] = None

# ── Gradientin piirtäminen ────────────────────────────────────────────────────
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
        seg_ids = []
        for x0s, x1s in _GRAD_SEG_X:
            rid = dpg.draw_rectangle(
                parent=tag_dl,
                pmin=(x0s, y0), pmax=(x1s, y1),
                fill=[128, 128, 128, 255], color=[0, 0, 0, 0],
            )
            seg_ids.append(rid)
        dpg.draw_rectangle(
            parent=tag_dl,
            pmin=(0, y0), pmax=(GRAD_W, y1),
            fill=[0, 0, 0, 0], color=[70, 70, 70, 180],
        )
        ind_col = [30, 30, 30, 240] if app.theme_name == "Light" else [255, 255, 255, 240]
        indicator_id = dpg.draw_line(
            parent=tag_dl,
            p1=(0, y0 - 2), p2=(0, y1 + 2),
            color=ind_col, thickness=2,
        )
        cache = {'seg_ids': seg_ids, 'indicator_id': indicator_id,
                 'y0': y0, 'y1': y1, 'ind_col': ind_col, 'last_theme': app.theme_name}
        _GRAD_ITEM_CACHE[gname] = cache

    seg_ids      = cache['seg_ids']
    indicator_id = cache['indicator_id']
    cy0          = cache['y0']
    cy1          = cache['y1']

    if cache.get('last_ctx') != ctx:
        cache['last_ctx'] = ctx
        for i, rid in enumerate(seg_ids):
            t = (i + 0.5) / SEGS
            r, g, b = _grad_color(gname, t, ctx)
            dpg.configure_item(rid, fill=[r, g, b, 255])

    t_val   = max(0., min(1., (val - smin) / (smax - smin))) if smax > smin else 0.
    tx      = int(t_val * (GRAD_W - 1))
    if cache['last_theme'] != app.theme_name:
        cache['ind_col']    = [30, 30, 30, 240] if app.theme_name == "Light" else [255, 255, 255, 240]
        cache['last_theme'] = app.theme_name
    ind_col = cache['ind_col']
    dpg.configure_item(indicator_id,
                       p1=(tx, cy0 - 2), p2=(tx, cy1 + 2), color=ind_col)
    dpg.set_value(f"slval_{gname}", val)

def _slval_input_cb(sender, val, user_data):
    """Callback: numeerinen syöte liukusäätimelle."""
    gname = user_data
    smin, smax = _SLIDER_RANGE[gname]
    app.sl_vals[gname] = max(smin, min(smax, val))
    mode = _GNAME_MODE.get(gname)
    if mode:
        _apply_from_mode(mode)
        from .harmony import _update_selected_pal_color, _update_wheel_contrast
        _update_selected_pal_color()
        _update_wheel_contrast()

def _make_grad_slider(label, label_color, gname):
    """Lisää yhden gradienttiliukusäätimen DPG:hen."""
    with dpg.group(horizontal=True):
        dpg.add_text(label, color=label_color)
        dpg.add_drawlist(width=GRAD_W, height=THUMB_H, tag=f"dl_{gname}")
        dpg.add_input_int(
            tag=f"slval_{gname}",
            default_value=0,
            width=_sc(52),
            step=0, step_fast=0,
            callback=_slval_input_cb,
            user_data=gname,
        )
