"""
color_math.py — Kaikki värilaskelmat.

Ei DPG-riippuvuuksia. Voidaan importoida ja testata itsenäisesti.
"""

import colorsys
from functools import lru_cache

# ── CSS-värikirjasto ──────────────────────────────────────────────────────────
# Kaikki CSS Color Module Level 4 -nimetyt värit. Aliakset (Aqua = Cyan,
# Fuchsia = Magenta, Grey-kirjoitusasut) jätetty pois, koska niillä on sama
# RGB-arvo kuin jo mukana olevilla nimillä.
CSS_COLORS = {
    "AliceBlue": (240, 248, 255),       "AntiqueWhite": (250, 235, 215),
    "Aquamarine": (127, 255, 212),      "Azure": (240, 255, 255),
    "Beige": (245, 245, 220),           "Bisque": (255, 228, 196),
    "Black": (0, 0, 0),                 "BlanchedAlmond": (255, 235, 205),
    "Blue": (0, 0, 255),                "BlueViolet": (138, 43, 226),
    "Brown": (165, 42, 42),             "BurlyWood": (222, 184, 135),
    "CadetBlue": (95, 158, 160),        "Chartreuse": (127, 255, 0),
    "Chocolate": (210, 105, 30),        "Coral": (255, 127, 80),
    "CornflowerBlue": (100, 149, 237),  "Cornsilk": (255, 248, 220),
    "Crimson": (220, 20, 60),           "Cyan": (0, 255, 255),
    "DarkBlue": (0, 0, 139),            "DarkCyan": (0, 139, 139),
    "DarkGoldenrod": (184, 134, 11),    "DarkGray": (169, 169, 169),
    "DarkGreen": (0, 100, 0),           "DarkKhaki": (189, 183, 107),
    "DarkMagenta": (139, 0, 139),       "DarkOliveGreen": (85, 107, 47),
    "DarkOrange": (255, 140, 0),        "DarkOrchid": (153, 50, 204),
    "DarkRed": (139, 0, 0),             "DarkSalmon": (233, 150, 122),
    "DarkSeaGreen": (143, 188, 143),    "DarkSlateBlue": (72, 61, 139),
    "DarkSlateGray": (47, 79, 79),      "DarkTurquoise": (0, 206, 209),
    "DarkViolet": (148, 0, 211),        "DeepPink": (255, 20, 147),
    "DeepSkyBlue": (0, 191, 255),       "DimGray": (105, 105, 105),
    "DodgerBlue": (30, 144, 255),       "FireBrick": (178, 34, 34),
    "FloralWhite": (255, 250, 240),     "ForestGreen": (34, 139, 34),
    "Gainsboro": (220, 220, 220),       "GhostWhite": (248, 248, 255),
    "Gold": (255, 215, 0),              "Goldenrod": (218, 165, 32),
    "Gray": (128, 128, 128),            "Green": (0, 128, 0),
    "GreenYellow": (173, 255, 47),      "Honeydew": (240, 255, 240),
    "HotPink": (255, 105, 180),         "IndianRed": (205, 92, 92),
    "Indigo": (75, 0, 130),             "Ivory": (255, 255, 240),
    "Khaki": (240, 230, 140),           "Lavender": (230, 230, 250),
    "LavenderBlush": (255, 240, 245),   "LawnGreen": (124, 252, 0),
    "LemonChiffon": (255, 250, 205),    "LightBlue": (173, 216, 230),
    "LightCoral": (240, 128, 128),      "LightCyan": (224, 255, 255),
    "LightGoldenrodYellow": (250, 250, 210),
    "LightGray": (211, 211, 211),       "LightGreen": (144, 238, 144),
    "LightPink": (255, 182, 193),       "LightSalmon": (255, 160, 122),
    "LightSeaGreen": (32, 178, 170),    "LightSkyBlue": (135, 206, 250),
    "LightSlateGray": (119, 136, 153),  "LightSteelBlue": (176, 196, 222),
    "LightYellow": (255, 255, 224),     "Lime": (0, 255, 0),
    "LimeGreen": (50, 205, 50),         "Linen": (250, 240, 230),
    "Magenta": (255, 0, 255),           "Maroon": (128, 0, 0),
    "MediumAquamarine": (102, 205, 170),"MediumBlue": (0, 0, 205),
    "MediumOrchid": (186, 85, 211),     "MediumPurple": (147, 112, 219),
    "MediumSeaGreen": (60, 179, 113),   "MediumSlateBlue": (123, 104, 238),
    "MediumSpringGreen": (0, 250, 154), "MediumTurquoise": (72, 209, 204),
    "MediumVioletRed": (199, 21, 133),  "MidnightBlue": (25, 25, 112),
    "MintCream": (245, 255, 250),       "MistyRose": (255, 228, 225),
    "Moccasin": (255, 228, 181),        "NavajoWhite": (255, 222, 173),
    "Navy": (0, 0, 128),                "OldLace": (253, 245, 230),
    "Olive": (128, 128, 0),             "OliveDrab": (107, 142, 35),
    "Orange": (255, 165, 0),            "OrangeRed": (255, 69, 0),
    "Orchid": (218, 112, 214),          "PaleGoldenrod": (238, 232, 170),
    "PaleGreen": (152, 251, 152),       "PaleTurquoise": (175, 238, 238),
    "PaleVioletRed": (219, 112, 147),   "PapayaWhip": (255, 239, 213),
    "PeachPuff": (255, 218, 185),       "Peru": (205, 133, 63),
    "Pink": (255, 192, 203),            "Plum": (221, 160, 221),
    "PowderBlue": (176, 224, 230),      "Purple": (128, 0, 128),
    "RebeccaPurple": (102, 51, 153),    "Red": (255, 0, 0),
    "RosyBrown": (188, 143, 143),       "RoyalBlue": (65, 105, 225),
    "SaddleBrown": (139, 69, 19),       "Salmon": (250, 128, 114),
    "SandyBrown": (244, 164, 96),       "SeaGreen": (46, 139, 87),
    "SeaShell": (255, 245, 238),        "Sienna": (160, 82, 45),
    "Silver": (192, 192, 192),          "SkyBlue": (135, 206, 235),
    "SlateBlue": (106, 90, 205),        "SlateGray": (112, 128, 144),
    "Snow": (255, 250, 250),            "SpringGreen": (0, 255, 127),
    "SteelBlue": (70, 130, 180),        "Tan": (210, 180, 140),
    "Teal": (0, 128, 128),              "Thistle": (216, 191, 216),
    "Tomato": (255, 99, 71),            "Turquoise": (64, 224, 208),
    "Violet": (238, 130, 238),          "Wheat": (245, 222, 179),
    "White": (255, 255, 255),           "WhiteSmoke": (245, 245, 245),
    "Yellow": (255, 255, 0),            "YellowGreen": (154, 205, 50),
}

_CSS_COLORS_LAB = {}

@lru_cache(maxsize=512)
def nearest_css(r, g, b):
    """Palauttaa lähimmän CSS-värin nimen CIELAB-avaruuden (Delta-E) avulla."""
    if not _CSS_COLORS_LAB:
        for n, (cr, cg, cb_) in CSS_COLORS.items():
            _CSS_COLORS_LAB[n] = rgb_to_lab(cr, cg, cb_)
    
    L, a, b_ = rgb_to_lab(r, g, b)
    best, bd = "Unknown", float('inf')
    for n, (cL, ca, cb) in _CSS_COLORS_LAB.items():
        d = (L - cL) ** 2 + (a - ca) ** 2 + (b_ - cb) ** 2
        if d < bd:
            bd = d
            best = n
    return best

# ── Linearisaatio-LUT (IEC 61966-2-1) ────────────────────────────────────────
_LIN_RGB_LUT = [
    (i/255.)/12.92 if (i/255.) <= 0.04045 else (((i/255.)+0.055)/1.055)**2.4
    for i in range(256)
]

def _lin(c):
    # round() eikä int(): liukuluvut (esim. 127.9) eivät katkea alaspäin
    return _LIN_RGB_LUT[max(0, min(255, int(round(c))))]

# ── Luminanssi ja kontrasti ───────────────────────────────────────────────────
def luminance(r, g, b):
    return 0.2126 * _lin(r) + 0.7152 * _lin(g) + 0.0722 * _lin(b)

def contrast_ratio(r1, g1, b1, r2, g2, b2):
    l1 = luminance(r1, g1, b1) + 0.05
    l2 = luminance(r2, g2, b2) + 0.05
    return max(l1, l2) / min(l1, l2)

def wcag_level(r):
    return "AAA" if r >= 7 else "AA" if r >= 4.5 else "AA*" if r >= 3 else "Fail"

def wcag_color(r):
    """Palauttaa [R,G,B]-värin WCAG-tasolle."""
    if r >= 7:   return [80,  210,  80]
    if r >= 4.5: return [170, 220,  60]
    if r >= 3:   return [230, 160,  40]
    return           [200,  70,  70]

# ── CMYK-konversiot ───────────────────────────────────────────────────────────
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

# ── LAB-konversiot (tarkka laskenta) ──────────────────────────────────────────
# Aiempi 4096-alkioinen LUT tuotti kylläisillä väreillä jopa 5 tason virheen
# RGB -> LAB -> RGB -kierroksella (kuutiojuuri on nollan lähellä jyrkkä).
# Tarkka laskenta on riittävän nopea (muutama sata kutsua ruutua kohti).
_LAB_EPS = 216 / 24389     # CIE-kynnys (~0.008856)
_LAB_K   = 841 / 108       # CIE-lineaarikerroin (~7.787)

def _xyz_f(t):
    return t ** (1 / 3) if t > _LAB_EPS else _LAB_K * t + 16 / 116

def _fi_lab(t):
    t3 = t * t * t
    return t3 if t3 > _LAB_EPS else (t - 16 / 116) / _LAB_K

def _delin(c):
    """Lineaarinen -> sRGB-gamma, rajattu välille 0..1."""
    if c <= 0.0:
        return 0.0
    if c >= 1.0:
        return 1.0
    return 12.92 * c if c <= 0.0031308 else 1.055 * c ** (1 / 2.4) - 0.055

def rgb_to_lab(r, g, b):
    rl = _lin(r); gl = _lin(g); bl = _lin(b)
    X = rl * .4124 + gl * .3576 + bl * .1805
    Y = rl * .2126 + gl * .7152 + bl * .0722
    Z = rl * .0193 + gl * .1192 + bl * .9505
    X /= .95047; Z /= 1.08883
    fY = _xyz_f(Y)
    return (
        116 * fY - 16,
        500 * (_xyz_f(X) - fY),
        200 * (fY - _xyz_f(Z)),
    )

def lab_to_rgb(L, a, b_):
    fy = (L + 16) / 116
    fx = a / 500 + fy
    fz = fy - b_ / 200
    X  = _fi_lab(fx) * .95047; Y = _fi_lab(fy); Z = _fi_lab(fz) * 1.08883
    rl =  X * 3.2406 - Y * 1.5372 - Z * .4986
    gl = -X * .9689  + Y * 1.8758  + Z * .0415
    bl =  X * .0557  - Y * .2040   + Z * 1.0570
    return (
        max(0, min(255, int(round(_delin(rl) * 255)))),
        max(0, min(255, int(round(_delin(gl) * 255)))),
        max(0, min(255, int(round(_delin(bl) * 255)))),
    )

# ── Apufunktiot ───────────────────────────────────────────────────────────────
def clamp01(v):
    return max(0., min(1., v))

def to_int(v):
    return max(0, min(255, int(round(v * 255))))

def color_sig(c):
    """Muuntaa normalisoitun värin (0-1) kokonaisluku-RGB:ksi (0-255)."""
    return tuple(to_int(v) for v in c[:3])

# ── Värisokeussimulaatio (CVD) ────────────────────────────────────────────────
CVD_MATRICES = {
    "Protanopia": [
        [0.152286, 1.052583, -0.204868],
        [0.114503, 0.786281, 0.099216],
        [-0.003882, -0.048116, 1.051998]
    ],
    "Deuteranopia": [
        [0.367322, 0.860646, -0.227968],
        [0.280085, 0.672501, 0.047413],
        [-0.011820, 0.042940, 0.968881]
    ],
    "Tritanopia": [
        [1.255528, -0.076749, -0.178779],
        [-0.078411, 0.930809, 0.147602],
        [0.004733, 0.691367, 0.303900]
    ]
}

def simulate_cvd(r, g, b, cvd_type):
    if cvd_type not in CVD_MATRICES: return [r, g, b]
    mat = CVD_MATRICES[cvd_type]
    rl, gl, bl = _lin(r), _lin(g), _lin(b)
    s_rl = rl * mat[0][0] + gl * mat[0][1] + bl * mat[0][2]
    s_gl = rl * mat[1][0] + gl * mat[1][1] + bl * mat[1][2]
    s_bl = rl * mat[2][0] + gl * mat[2][1] + bl * mat[2][2]
    return [
        max(0, min(255, int(round(_delin(s_rl) * 255)))),
        max(0, min(255, int(round(_delin(s_gl) * 255)))),
        max(0, min(255, int(round(_delin(s_bl) * 255))))
    ]

# ── Formaatit ─────────────────────────────────────────────────────────────────
FORMAT_OPTIONS = ["HEX", "RGB", "HSL", "HSV", "CSS Name", "CMYK"]

def format_color(rgb, fmt):
    """Muotoilee RGB-tuplan annetussa formaatissa merkkijonoksi."""
    if rgb is None:
        return ""
    r, g, b = rgb
    if fmt == "HEX":
        return f"#{r:02x}{g:02x}{b:02x}".upper()
    if fmt == "RGB":
        return f"{r}, {g}, {b}"
    if fmt == "HSL":
        h, l, s = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
        return f"{h*360:.0f}  {s*100:.0f}%  {l*100:.0f}%"
    if fmt == "HSV":
        h, s, v = colorsys.rgb_to_hsv(r / 255, g / 255, b / 255)
        return f"{h*360:.0f}  {s*100:.0f}%  {v*100:.0f}%"
    if fmt == "CSS Name":
        return nearest_css(r, g, b)
    if fmt == "CMYK":
        c, m, y, k = rgb_to_cmyk(r, g, b)
        return f"C{c:.0f}% M{m:.0f}% Y{y:.0f}% K{k:.0f}%"
    return ""
