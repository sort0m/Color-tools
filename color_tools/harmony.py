"""
harmony.py — Harmonialogiikka, WCAG-kontrasti, värikortti-teoria, history-lisäys.

Sisältää myös update_wheel_contrast ja picker-callback, koska ne ovat
tiiviisti kytketty toisiinsa.
"""

import colorsys
import re
import time as _time

import dearpygui.dearpygui as dpg
from .accessibility import contrast_color

from .color_math import (
    contrast_ratio, wcag_level, wcag_color, nearest_css,
    color_sig, format_color,
)
from .i18n import t          # i18n ei tuo harmonya -> ei kehäriippuvuutta
from .state import app, _history_lock

# ── Harmonia-tagit ja -avaimet ────────────────────────────────────────────────
HARMONY_TAGS = ["main", "h1", "h2", "h3", "h4", "h5", "h6"]
HIST_TAGS    = [f"hist_{i}" for i in range(60)]

_HARMONY_ICON_DATA = {
    "Complementary":       ([180],          None),
    "Split Complementary": ([150, 210],     None),
    "Analogous":           ([30, -30],      None),
    "Triadic":             ([120, 240],     None),
    "Tetradic":            ([90, 180, 270], None),
    "Rectangle":           ([60, 180, 240], None),
    "Tints":               ([],             "tint"),
    "Shades":              ([],             "shade"),
    "Tones":               ([],             "tone"),
}
_HARMONY_ANGLES = {k: v[0] for k, v in _HARMONY_ICON_DATA.items() if v[0]}

HARMONY_KEYS = {tag: {
    "group":       f"group_{tag}",
    "rect":        f"rect_{tag}",
    "val":         f"val_{tag}",
    "copy_btn":    f"copy_btn_{tag}",
} for tag in HARMONY_TAGS}

# ── Harmonia-arvojen päivitys ─────────────────────────────────────────────────
def format_value(tag, fmt):
    """Muotoilee harmonia-tagin arvon valitussa formaatissa."""
    rgb = app.harmony_rgb.get(tag)
    return format_color(rgb, fmt) if rgb is not None else ""

def refresh_harmony_values():
    from .fonts import bind_value_font
    fmt = app.fmt_mode
    for tag in HARMONY_TAGS:
        keys = HARMONY_KEYS[tag]
        if not dpg.does_item_exist(keys["val"]):
            continue
        if not dpg.is_item_shown(keys["group"]):
            continue
        dpg.configure_item(keys["val"],      show=True)
        dpg.configure_item(keys["copy_btn"], show=True)
        dpg.set_value(f"val_{tag}", format_value(tag, fmt))
        bind_value_font(keys["val"])

# ── WCAG-kontrastin päivitys ──────────────────────────────────────────────────
_wc_themes_created = [False]
_wc_last_sig       = [None]

def set_current_as_sample_background():
    app.contrast_bg = [*color_sig(app.current_color), 255]
    _update_wheel_contrast()

def _wcag_text(ratio):
    """WCAG-taso käännettynä ("Fail" -> käännös)."""
    lvl = wcag_level(ratio)
    return t("wcag_level_fail") if lvl == "Fail" else lvl

def _update_wheel_contrast():
    """Päivittää WCAG-kontrastinäytön."""
    if not all(dpg.does_item_exist(tag) for tag in ("wheel_ctr_w", "wheel_ctr_b")):
        return

    r, g, b = color_sig(app.current_color)
    custom_bg = list(app.contrast_bg[:3])

    _sig = (r, g, b, tuple(custom_bg))
    if _sig == _wc_last_sig[0] and _wc_themes_created[0]:
        return
    _wc_last_sig[0] = _sig

    cw = contrast_ratio(r, g, b, 255, 255, 255)
    cb = contrast_ratio(r, g, b, 0,   0,   0)
    cc = contrast_ratio(r, g, b, *custom_bg)

    cw_color = contrast_color(wcag_color(cw), [[0, 0, 0]])
    cb_color = contrast_color(wcag_color(cb), [[0, 0, 0]])
    cc_color = contrast_color(wcag_color(cc), [[0, 0, 0]])

    # W/B/C -napit päivitetään
    _w_lbl = t("wcag_w")[0]   # "W" tai "V"
    _b_lbl = t("wcag_b")[0]   # "B" tai "M"
    _c_lbl = t("wcag_c")[0]   # "C" tai "O"
    dpg.configure_item("wheel_ctr_w",
        label=f"{_w_lbl}  {cw:.2f}:1")
    dpg.set_value("wheel_ctr_w_res", _wcag_text(cw))
    
    dpg.configure_item("wheel_ctr_b",
        label=f"{_b_lbl}  {cb:.2f}:1")
    dpg.set_value("wheel_ctr_b_res", _wcag_text(cb))
    
    if dpg.does_item_exist("wheel_ctr_custom"):
        dpg.configure_item("wheel_ctr_custom",
            label=f"{_c_lbl}  {cc:.2f}:1")
        dpg.set_value("wheel_ctr_custom_res", _wcag_text(cc))

    _BLK      = [0, 0, 0, 255]
    _rgb      = [r, g, b, 255]
    _cust_bg  = custom_bg + [255]

    # Simulaatio
    from .color_math import simulate_cvd
    prot_fg = simulate_cvd(r, g, b, "Protanopia") + [255]
    prot_bg = simulate_cvd(*custom_bg, "Protanopia") + [255]
    deut_fg = simulate_cvd(r, g, b, "Deuteranopia") + [255]
    deut_bg = simulate_cvd(*custom_bg, "Deuteranopia") + [255]
    trit_fg = simulate_cvd(r, g, b, "Tritanopia") + [255]
    trit_bg = simulate_cvd(*custom_bg, "Tritanopia") + [255]

    _specs = [
        # (teema-tagi, widget-tagi, taustaväri, tekstiväri)
        ("wheel_ctr_w_dynamic_theme",      "wheel_ctr_w",         _BLK, cw_color),
        ("wheel_ctr_b_dynamic_theme",      "wheel_ctr_b",         _BLK, cb_color),
        ("wheel_ctr_custom_dynamic_theme", "wheel_ctr_custom",    _BLK, cc_color),
        # Sample Text: nykyinen väri tekstinä, oma taustaväri taustana
        ("wheel_preview_w_dynamic_theme",  "wheel_preview_w_txt", _cust_bg, _rgb),
        # CVD Sample texts
        ("sim_prot_dynamic_theme", "btn_sim_prot", prot_bg, prot_fg),
        ("sim_deut_dynamic_theme", "btn_sim_deut", deut_bg, deut_fg),
        ("sim_trit_dynamic_theme", "btn_sim_trit", trit_bg, trit_fg),
        ("sim_norm_dynamic_theme", "btn_sim_norm", _cust_bg, _rgb),
    ]
    for th_tag, item_tag, bg, txt in _specs:
        if not dpg.does_item_exist(item_tag):
            continue
        
        c_bg_tag = f"{th_tag}_bg"
        c_hover_tag = f"{th_tag}_hover"
        c_active_tag = f"{th_tag}_active"
        c_txt_tag = f"{th_tag}_txt"

        if not dpg.does_item_exist(th_tag):
            with dpg.theme(tag=th_tag):
                with dpg.theme_component(dpg.mvButton):
                    dpg.add_theme_color(dpg.mvThemeCol_Button,        bg, tag=c_bg_tag)
                    dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, bg, tag=c_hover_tag)
                    dpg.add_theme_color(dpg.mvThemeCol_ButtonActive,  bg, tag=c_active_tag)
                    dpg.add_theme_color(dpg.mvThemeCol_Text,          txt, tag=c_txt_tag)
            dpg.bind_item_theme(item_tag, th_tag)
        else:
            dpg.set_value(c_bg_tag, bg)
            dpg.set_value(c_hover_tag, bg)
            dpg.set_value(c_active_tag, bg)
            dpg.set_value(c_txt_tag, txt)

    _wc_themes_created[0] = True

# ── History ───────────────────────────────────────────────────────────────────
def add_to_history(color):
    r, g, b = color_sig(color)
    e = [r, g, b, 255]
    with _history_lock:
        if app.history and app.history[0] == e:
            return
        app.history.appendleft(e)
        app._hist_dirty = True
        app._history_save_due = _time.perf_counter() + 2.0   # tallenna viiveellä
        if app.palette_select_mode:
            app._palette_select_cancel_pending = True

# ── Palette-edit värin päivitys ───────────────────────────────────────────────
def _update_selected_pal_color():
    from .state import _palettes_lock
    if app._editing_pal and app._pal_selected_idx is not None:
        with _palettes_lock:
            if app._editing_pal in app.palettes:
                colors = app.palettes[app._editing_pal]
                if app._pal_selected_idx < len(colors):
                    r, g, b = color_sig(app.current_color)
                    if not app._pal_edit_dirty:
                        snap = app._pal_pre_edit_snapshot
                        if snap is not None:
                            nm = app._editing_pal
                            from .palette_history import record_snapshot
                            record_snapshot(app._pal_undo_stack, nm, snap)
                            if dpg.does_item_exist("pal_undo_btn"):
                                dpg.configure_item("pal_undo_btn", enabled=any(
                                    entry[0] == nm for entry in app._pal_undo_stack))
                        app._pal_edit_dirty = True
                        app._pal_pre_edit_snapshot = None
                    colors[app._pal_selected_idx] = [r, g, b, 255]

def _apply_external_color_to_pal():
    """Pipetti, hex-kenttä ja historiaklikkaus: vie nykyinen väri valittuun
    palettiswatchiin (data + undo), eikä vain näytölle."""
    if app._editing_pal and app._pal_selected_idx is not None:
        _update_selected_pal_color()
        app.change_time = _time.perf_counter()   # laukaisee tallennuksen

# ── Picker-callback ja värin pakottaminen ─────────────────────────────────────
def _force_update_color_picker(col_norm):
    app.current_color = list(col_norm)
    app._last_picker_force_time = _time.perf_counter()
    app._pending_picker_color   = list(col_norm)

def picker_cb(s, v):
    if _time.perf_counter() - app._last_picker_force_time < 0.5:
        return
    if app._picker_processing:
        return
    app._picker_processing = True
    try:
        if any(x > 1.01 for x in v):
            v = [x / 255.0 for x in v]
        app.current_color = list(v)
        app.change_time   = _time.perf_counter()
        from .sliders import _sync_all_modes
        _sync_all_modes()
        _update_selected_pal_color()
        _update_wheel_contrast()
    finally:
        app._picker_processing = False

# ── Kopiointifunktiot ─────────────────────────────────────────────────────────
def copy_preview(s, v):
    from .bootstrap import safe_copy
    r, g, b = color_sig(app.current_color)
    safe_copy(f"#{r:02x}{g:02x}{b:02x}".upper())

def copy_harmony_val(s, a, u):
    from .bootstrap import safe_copy
    safe_copy(dpg.get_value(f"val_{u}"))

def hex_input_cb(s, v):
    raw = v.strip()
    if raw.startswith('#'):
        raw = raw[1:]
    if not re.fullmatch(r'[0-9a-fA-F]{3}|[0-9a-fA-F]{6}', raw):
        return
    if len(raw) == 3:
        raw = raw[0]*2 + raw[1]*2 + raw[2]*2
    r = int(raw[0:2], 16)
    g = int(raw[2:4], 16)
    b = int(raw[4:6], 16)
    col_norm = [r / 255., g / 255., b / 255., 1.]
    _force_update_color_picker(col_norm)
    from .sliders import _sync_all_modes
    _sync_all_modes()
    _update_wheel_contrast()
    add_to_history(col_norm)
    _apply_external_color_to_pal()

def toggle_on_top(s, v):
    from .state import save_config
    app.always_on_top = bool(v)
    dpg.set_viewport_always_top(v)
    save_config()

# ── Värikortti-teoria ─────────────────────────────────────────────────────────
# Valintalogiikka on kielestä riippumaton: funktiot palauttavat tekstien tunnisteet,
# ja varsinaiset tekstit ovat i18n.py:ssä (theory_a_N / r_N / h_N / m_N / t_N / s_N /
# tt_N / td_N / ui_N).

def _theory_ids(hue, l, s, delta):
    """Palauttaa (assosiaatio, tunnelma, kylläisyys, historia) -tunnisteet."""
    neutral     = s < 0.05 or delta < 8
    tinted_grey = not neutral and s < 0.2                        # taupe, liuskekivi, ...
    chromatic   = not neutral and not tinted_grey
    cream = chromatic and 20 <= hue < 65 and l >= 0.82           # kerma, norsunluu, persikka
    sand  = chromatic and 20 <= hue < 65 and 0.58 <= l < 0.82 and s < 0.6   # hiekka, beige, tan

    # -- assosiaatio --
    if neutral:
        if l > 0.8:
            a_id = 1
        elif l < 0.2:
            a_id = 2
        else:
            a_id = 3
    elif tinted_grey:
        a_id = 38 if (hue < 70 or hue >= 330) else 39            # lämmin / viileä harmaa
    elif cream:
        a_id = 37
    elif sand:
        a_id = 36
    elif 0 <= hue < 15 or 345 <= hue <= 360:
        if l < 0.35:
            a_id = 4
        elif s < 0.4:
            a_id = 5
        elif s > 0.8 and l > 0.4:
            a_id = 6
        elif l > 0.65:
            a_id = 7
        else:
            a_id = 8
    elif 15 <= hue < 40:
        if l < 0.4 or (l < 0.55 and s < 0.85):                   # ruskea
            a_id = 9
        elif l > 0.7 and s < 0.6:
            a_id = 10
        else:
            a_id = 11
    elif 40 <= hue < 65:
        if l < 0.3:                                              # oliivi, tumma sinappi
            a_id = 40
        elif l < 0.6 and s < 0.5:
            a_id = 12
        elif s < 0.75 and 0.4 < l < 0.6:
            a_id = 13
        elif s > 0.9 and l > 0.45:
            a_id = 14
        else:
            a_id = 15
    elif 65 <= hue < 150:
        if l < 0.25:
            a_id = 16
        elif s < 0.4 and l < 0.4:
            a_id = 17
        elif 70 <= hue < 100 and s < 0.6 and l < 0.6:
            a_id = 18
        elif s > 0.85 and l > 0.4:
            a_id = 19
        else:
            a_id = 20
    elif 150 <= hue < 195:
        if l < 0.35:
            a_id = 21
        elif s > 0.7 and l > 0.4:
            a_id = 22
        else:
            a_id = 23
    elif 195 <= hue < 250:
        if l < 0.3:
            a_id = 24
        elif s > 0.8 and l > 0.4:
            a_id = 25
        elif l > 0.7 and s < 0.5:
            a_id = 26
        else:
            a_id = 27
    elif 250 <= hue < 290:
        if l < 0.35:
            a_id = 28
        elif s > 0.7 and l > 0.4:
            a_id = 29
        elif l > 0.6 and s < 0.5:
            a_id = 30
        else:
            a_id = 31
    else:                                                        # 290-345 (magenta, pinkki)
        if s > 0.7 and l > 0.4:
            a_id = 32
        elif l > 0.7:
            a_id = 33
        else:
            a_id = 34

    # -- tunnelma (kirkkaus) --
    if l < 0.18:
        m_id = 1
    elif l < 0.3:
        m_id = 2
    elif l > 0.82:
        m_id = 3
    elif l > 0.68:
        m_id = 4
    elif l < 0.45:
        m_id = 5
    elif l > 0.55:
        m_id = 6
    else:
        m_id = 7

    # -- kylläisyys --
    if s < 0.12:
        s_id = 1
    elif s < 0.24:
        s_id = 2
    elif s < 0.42:
        s_id = 3
    elif s < 0.62:
        s_id = 4
    elif s < 0.8:
        s_id = 5
    else:
        s_id = 6

    # -- historia --
    if neutral:
        h_id = 1
    elif tinted_grey:
        h_id = 10
    elif cream or sand:
        h_id = 11
    elif 0 <= hue < 15 or 345 <= hue <= 360:
        h_id = 2
    elif 15 <= hue < 40:
        h_id = 3
    elif 40 <= hue < 65:
        h_id = 4
    elif 65 <= hue < 150:
        h_id = 5
    elif 150 <= hue < 195:
        h_id = 6
    elif 195 <= hue < 250:
        h_id = 7
    elif 250 <= hue < 290:
        h_id = 8
    else:
        h_id = 9
    return a_id, m_id, s_id, h_id


# Sävyn "lämpötila" (-1 viileä ... +1 lämmin) ankkuripisteinä. Käyrä ei ole symmetrinen:
# punainen-oranssi-keltainen on lämmin, syaani-sininen viilein, mutta vihreä ja magenta ovat
# "välimaastossa" (lämpötila riippuu alasävystä ja ympäröivistä väreistä).
_TEMP_CURVE = [(0, 0.85), (30, 1.0), (60, 0.85), (90, 0.45), (120, 0.0), (150, -0.5),
               (180, -0.85), (210, -1.0), (240, -0.85), (270, -0.45), (300, 0.0),
               (330, 0.55), (360, 0.85)]

def _hue_temperature(hue):
    hue %= 360
    for (h0, v0), (h1, v1) in zip(_TEMP_CURVE, _TEMP_CURVE[1:]):
        if h0 <= hue <= h1:
            return v0 + (v1 - v0) * (hue - h0) / (h1 - h0)
    return 0.0


def _temperature(hue, delta):
    """Värilämpötila: (taso 1-6, mittarin arvo 0..1; 0 = viileä, 1 = lämmin).

    Tasot: 1 viileä, 2 melko viileä, 3 neutraali (lähes harmaa väri), 4 melko lämmin,
    5 lämmin, 6 välimaastossa (selvä väri, jonka sävy on lämpimän ja viileän rajalla:
    vihreä, magenta). Vaikutus skaalataan värin värikkyydellä."""
    strength = min(1.0, (delta / 255.0) / 0.30)
    value = _hue_temperature(hue) * strength
    if strength < 0.25:
        level = 3                      # lähes harmaa: ei lämpötilaa
    elif abs(value) < 0.2:
        level = 6                      # selvä väri, mutta lämpötila epäselvä
    elif value >= 0.55:
        level = 5
    elif value > 0:
        level = 4
    elif value > -0.55:
        level = 2
    else:
        level = 1
    return level, (value + 1.0) / 2.0


def _ui_role_id(hue, l, s, delta):
    """Värin tyypillinen rooli käyttöliittymässä (tunniste ui_N)."""
    if s < 0.2 or delta < 8:                       # neutraalit ja sävyttyneet harmaat
        return 1 if l > 0.85 else 2 if l < 0.2 else 3
    if 20 <= hue < 65:                             # samat rajat kuin kermalla ja hiekalla (_theory_ids)
        if l >= 0.82:
            return 4                               # kerma, norsunluu: hento tausta
        if 0.58 <= l < 0.82 and s < 0.6:
            return 15                              # hiekka, beige: hillitty korostus
    if l > 0.85:
        return 4                                   # hento tausta / hover
    if l < 0.22:
        return 5                                   # tumma korostus
    if (15 <= hue < 40 and (l < 0.4 or (l < 0.55 and s < 0.85))) or (40 <= hue < 65 and l < 0.45):
        return 16                                  # ruskea, oliivi: maanläheinen (samat rajat kuin _theory_ids)
    if l > 0.72:
        return 14                                  # pehmeä korostus
    if s < 0.45:
        return 15                                  # hillitty korostus
    if hue < 15 or hue >= 345:
        return 6                                   # punainen: virhe
    if hue < 40:
        return 7                                   # oranssi: CTA / varoitus
    if hue < 65:
        return 8                                   # keltainen: ilmoitus (tumma teksti)
    if hue < 150:
        return 9                                   # vihreä: onnistuminen
    if hue < 195:
        return 10                                  # turkoosi: toissijainen
    if hue < 250:
        return 11                                  # sininen: pääväri / linkit
    if hue < 290:
        return 12                                  # violetti: premium
    return 13                                      # magenta/pinkki: leikkisä


def color_theory_info(rgb):
    """Värin teoriatiedot nykyisellä kielellä (dict):
    name, hex, association, reason, history, mood, saturation, tone, temperature,
    temp_value (0..1), temp_heading (otsikko mittarille), ui_role, tags (lista), tags_text,
    header (otsikkorivi), body (selitysteksti) ja full (header + tagit + body)."""
    r, g, b = [max(0, min(255, int(c))) for c in rgb]
    h, l, s = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
    hue = h * 360
    delta = max(r, g, b) - min(r, g, b)
    a_id, m_id, s_id, h_id = _theory_ids(hue, l, s, delta)
    temp_id, temp_value = _temperature(hue, delta)
    info = {
        "name":        nearest_css(r, g, b),
        "hex":         f"{r:02X}{g:02X}{b:02X}",
        "association": t(f"theory_a_{a_id}"),
        "reason":      t(f"theory_r_{a_id}"),
        "history":     t(f"theory_h_{h_id}"),
        "mood":        t(f"theory_m_{m_id}"),
        "tone":        t(f"theory_t_{m_id}"),
        "saturation":  t(f"theory_s_{s_id}"),
        "temperature": f'{t(f"theory_tt_{temp_id}")} - {t(f"theory_td_{temp_id}")}',   # ASCII: oletusfontissa ei ole ajatusviivaa
        "temp_value":  temp_value,
        "ui_role":     t(f"theory_ui_{_ui_role_id(hue, l, s, delta)}"),
    }
    keywords = [k.strip() for k in info["association"].split(",") if k.strip()]
    info["tags"] = keywords        # avainsanat näytetään vain kerran (tunnisteina); lämpötila on oma mittarinsa
    info["temp_heading"] = f'{t("theory_temp_title")}: {t(f"theory_tt_{temp_id}")}'
    info["tags_text"] = "  ·  ".join(info["tags"])
    info["association_heading"] = t("theory_association_title")
    info["header"] = t("theory_header").format(**info)
    info["body"]   = t("theory_body").format(**info)
    info["full"]   = info["header"] + "\n" + info["tags_text"] + "\n\n" + info["body"]
    return info
