"""Bounded palette readers. No dialogs, threads or application state."""

import colorsys
import json
import math
import struct
import zipfile
import time as _time

from .color_math import lab_to_rgb

MAX_COLORS = 1000
MAX_ACO_COLORS = 65535
MAX_FILE_BYTES = 20 * 1024 * 1024
MAX_ZIP_JSON_BYTES = 5 * 1024 * 1024


def _clamp8(value):
    return max(0, min(255, round(value)))


def _unit(value):
    value = float(value)
    if not math.isfinite(value):
        raise ValueError('non-finite color')
    return max(0., min(1., value))


def _aco_to_rgb(space, w, x, y, z):
    if space == 0:
        return w >> 8, x >> 8, y >> 8
    if space == 1:
        return tuple(_clamp8(v * 255) for v in colorsys.hsv_to_rgb(w / 65535., x / 65535., y / 65535.))
    if space == 2:
        return tuple(_clamp8(255 * v / 65535. * z / 65535.) for v in (w, x, y))
    return None


def _ase_to_rgb(model, values):
    if not all(math.isfinite(v) for v in values):
        raise ValueError('non-finite color')
    if model == b'RGB ':
        return tuple(_clamp8(_unit(v) * 255) for v in values)
    if model == b'CMYK':
        c, m, y, k = map(_unit, values)
        return tuple(_clamp8(255 * (1 - v) * (1 - k)) for v in (c, m, y))
    if model == b'Gray':
        return (_clamp8(_unit(values[0]) * 255),) * 3
    if model == b'LAB ':
        return lab_to_rgb(_unit(values[0]) * 100,
                          max(-128., min(127., values[1])),
                          max(-128., min(127., values[2])))
    return None


def parse_ase(data, limit=MAX_COLORS):
    if len(data) < 12 or data[:4] != b'ASEF' or struct.unpack_from('>H', data, 4)[0] != 1:
        raise ValueError('invalid ASE header')
    count = struct.unpack_from('>I', data, 8)[0]
    pos, colors = 12, []
    for _ in range(count):
        if pos + 6 > len(data):
            raise ValueError('truncated ASE block header')
        kind, size = struct.unpack_from('>HI', data, pos)
        pos += 6
        end = pos + size
        if end > len(data):
            raise ValueError('truncated ASE block')
        if kind == 1:
            if pos + 2 > end:
                raise ValueError('missing ASE color name')
            chars = struct.unpack_from('>H', data, pos)[0]
            pos += 2
            if chars < 1 or pos + chars * 2 + 4 > end:
                raise ValueError('invalid ASE name length')
            if data[pos + (chars - 1) * 2:pos + chars * 2] != b'\0\0':
                raise ValueError('unterminated ASE name')
            pos += chars * 2
            model = data[pos:pos + 4]
            pos += 4
            components = {b'RGB ': 3, b'CMYK': 4, b'Gray': 1, b'LAB ': 3}.get(model)
            if components:
                if pos + components * 4 + 2 > end:
                    raise ValueError('ASE components outside block')
                rgb = _ase_to_rgb(model, struct.unpack_from('>' + 'f' * components, data, pos))
                colors.append([*rgb, 255])
        pos = end
        if len(colors) > limit:
            break
    return colors[:limit], len(colors) > limit


def parse_aco(data, limit=MAX_COLORS):
    """Return supported colors, limit flag and unsupported-color count."""
    if len(data) < 4:
        raise ValueError('invalid ACO header')
    version, count = struct.unpack_from('>HH', data)
    if version not in (1, 2):
        raise ValueError('unsupported ACO version')
    pos, colors = 4, []
    skipped = supported = 0
    for _ in range(count):
        if pos + 10 > len(data):
            raise ValueError('truncated ACO color')
        rgb = _aco_to_rgb(*struct.unpack_from('>HHHHH', data, pos))
        pos += 10
        if version == 2:
            if pos + 4 > len(data):
                raise ValueError('missing ACO name length')
            chars = struct.unpack_from('>I', data, pos)[0]
            pos += 4
            if chars < 1 or pos + chars * 2 > len(data):
                raise ValueError('invalid ACO name length')
            pos += chars * 2
        if rgb is None:
            skipped += 1
        else:
            supported += 1
            if len(colors) < limit:
                colors.append([*rgb, 255])
    return colors, supported > limit, skipped


def parse_gpl(stream, limit=MAX_COLORS):
    if stream.readline().strip() != 'GIMP Palette':
        raise ValueError('invalid GPL header')
    colors = []
    for line in stream:
        fields = line.split(maxsplit=3)
        if len(fields) < 3:
            continue
        try:
            rgb = [int(v) for v in fields[:3]]
        except ValueError:
            continue
        if all(0 <= v <= 255 for v in rgb):
            colors.append([*rgb, 255])
        if len(colors) > limit:
            break
    return colors[:limit], len(colors) > limit


def parse_swatches(stream, limit=MAX_COLORS):
    with zipfile.ZipFile(stream) as archive:
        names = archive.namelist()
        name = next((n for n in names if n.lower() == 'swatches.json'), None)
        if name is None:
            name = next((n for n in names if n.lower().endswith('.json')), None)
        if name is None or archive.getinfo(name).file_size > MAX_ZIP_JSON_BYTES:
            raise ValueError('invalid swatches JSON')
        with archive.open(name) as member:
            raw = member.read(MAX_ZIP_JSON_BYTES + 1)
        if len(raw) > MAX_ZIP_JSON_BYTES:
            raise ValueError('swatches JSON too large')
        doc = json.loads(raw)
    if not isinstance(doc, dict) or not isinstance(doc.get('swatches'), list):
        raise ValueError('invalid swatches structure')
    colors = []
    for swatch in doc['swatches']:
        if not isinstance(swatch, dict):
            continue
        try:
            rgb = colorsys.hsv_to_rgb(*(_unit(swatch.get(k, 0)) for k in ('hue', 'saturation', 'brightness')))
        except (TypeError, ValueError, OverflowError):
            continue
        colors.append([*(_clamp8(v * 255) for v in rgb), 255])
        if len(colors) > limit:
            break
    return colors[:limit], len(colors) > limit


def parse_afpalette(data, limit=MAX_COLORS):
    """Read the supported uncompressed Affinity v2 swatch layout.

    Length-prefixed names are skipped, never scanned for color markers.
    Unknown layouts are rejected rather than guessed from arbitrary bytes.
    """
    if len(data) < 80 or data[:4] != b'\x00\xffKA':
        raise ValueError('invalid Affinity header')
    end = struct.unpack_from('<Q', data, 16)[0]
    if not 80 <= end <= len(data):
        raise ValueError('invalid Affinity payload boundary')
    payload = memoryview(data)[80:end]
    pos = 0

    def take(size):
        nonlocal pos
        if pos + size > len(payload):
            raise ValueError('truncated Affinity palette')
        value = bytes(payload[pos:pos + size])
        pos += size
        return value

    def expect(value):
        if take(len(value)) != value:
            raise ValueError('unsupported Affinity palette layout')

    expect(b'\x02\x00VlaP\x01\x00')
    # Both layout words occur in real files: our writer emits 0x20,
    # Affinity's Colours.afpalette emits 0x21. Their color records share
    # the structure below. Do not infer support for other layout words.
    layout_word = struct.unpack('<I', take(4))[0]
    if layout_word not in (0x20, 0x21):
        raise ValueError('unsupported Affinity palette layout word')
    expect(b'+NClP')
    take(struct.unpack('<I', take(4))[0])
    expect(b'\xb1VlaP')
    count, version = struct.unpack('<II', take(8))
    if version != 1:
        raise ValueError('unsupported Affinity palette version')
    expect(b'\x00\x00')
    colors = []
    for index in range(count):
        if index == 0:
            expect(b'SliF\x01\x00\x00\x00lliF\x00\x00\x00\x021rloC\x01\x01\x00\x00\x00\x00')
            model = take(4)
            expect(b'\x01\x00\x00\x02Dloc_')
        else:
            expect(b'SliF1rloC\x01')
            take(4)  # object id
            expect(b'\x01')
            model = take(4)
            expect(b'Dloc_')
        values = struct.unpack('<ffff', take(16))
        if not all(math.isfinite(v) for v in values):
            raise ValueError('non-finite Affinity color')
        a, b, c, d = map(_unit, values)
        if model == b'ABGR':
            rgb = a, b, c
        elif model == b'ALSH':
            rgb = colorsys.hls_to_rgb(a, c, b)
        elif model == b'KYMC':
            rgb = tuple((1 - v) * (1 - d) for v in (a, b, c))
        else:
            raise ValueError('unsupported Affinity color space')
        colors.append([*(_clamp8(v * 255) for v in rgb), 255])
        if len(colors) > limit:
            break
        expect(b'\x00\x00')
        if index < count - 1:
            expect(b'\x01')
            take(4)  # next object id
            expect(b'\x01')
    return colors[:limit], len(colors) > limit


# Export serialization

def _build_ase(color_list):
    def _enc_name(n):
        return struct.pack('>H', len(n) + 1) + (n + '\x00').encode('utf-16-be')
    entries = []
    for c in color_list:
        rf = c[0] / 255.0; gf = c[1] / 255.0; bf = c[2] / 255.0
        label = f"#{c[0]:02X}{c[1]:02X}{c[2]:02X}"
        block_data = (
            _enc_name(label) + b'RGB '
            + struct.pack('>fff', rf, gf, bf)
            + struct.pack('>H', 2)
        )
        entries.append(struct.pack('>HI', 0x0001, len(block_data)) + block_data)
    header = b'ASEF' + struct.pack('>HH', 1, 0) + struct.pack('>I', len(entries))
    return header + b''.join(entries)


def _build_aco(color_list):
    n = len(color_list)
    if n > MAX_ACO_COLORS:
        raise ValueError(f'ACO supports at most {MAX_ACO_COLORS} colors')
    def _encode_color(c):
        r16 = (c[0] << 8) | c[0]; g16 = (c[1] << 8) | c[1]; b16 = (c[2] << 8) | c[2]
        return struct.pack('>HHHHH', 0, r16, g16, b16, 0)
    v1 = struct.pack('>HH', 1, n)
    for c in color_list: v1 += _encode_color(c)
    v2 = struct.pack('>HH', 2, n)
    for i, c in enumerate(color_list):
        name = f"Color {i + 1}"
        encoded = (name + '\x00').encode('utf-16-be')
        v2 += _encode_color(c)
        v2 += struct.pack('>HH', 0, len(name) + 1) + encoded
    return v1 + v2


def _build_gpl(palette_name, color_list):
    lines = ["GIMP Palette", f"Name: {palette_name}", "Columns: 16", "#"]
    for c in color_list:
        hex_val = f"{c[0]:02X}{c[1]:02X}{c[2]:02X}"
        lines.append(f"{c[0]:3d} {c[1]:3d} {c[2]:3d}     #{hex_val}")
    return "\n".join(lines).encode('utf-8')


def _build_swatches(palette_name, color_list):
    swatches = []
    for c in color_list:
        r, g, b = c[0] / 255.0, c[1] / 255.0, c[2] / 255.0
        hue, saturation, brightness = colorsys.rgb_to_hsv(r, g, b)
        swatches.append({
            "hue": hue, "saturation": saturation, "brightness": brightness,
            "alpha": 1.0, "colorSpace": 0,
            "red": r, "green": g, "blue": b,
        })
    while len(swatches) < 30:
        swatches.append(None)
    doc = {"name": palette_name, "swatches": swatches}
    return json.dumps(doc, separators=(',', ':')).encode('utf-8')


def _build_afpalette(palette_name, color_list):
    """Rakentaa Affinity v2 -yhteensopivan .afpalette -tiedoston."""
    n_colors = len(color_list)
    name_bytes = palette_name.encode('utf-8')

    name_sec = (
        b'\x02\x00VlaP\x01\x00\x20\x00\x00\x00'
        + b'+NClP'
        + struct.pack('<I', len(name_bytes))
        + name_bytes
        + b'\xb1VlaP'
        + struct.pack('<II', n_colors, 1)
        + b'\x00\x00'
    )

    color_blocks = []
    for idx, c in enumerate(color_list):
        r_f = c[0] / 255.0
        g_f = c[1] / 255.0
        b_f = c[2] / 255.0
        a_f = 1.0
        float_bytes = struct.pack('<ffff', r_f, g_f, b_f, a_f)

        if idx == 0:
            blk = (
                b'SliF\x01\x00\x00\x00lliF\x00\x00\x00\x021rloC\x01\x01\x00\x00\x00\x00ABGR\x01\x00\x00\x02Dloc_'
                + float_bytes
            )
            if n_colors > 1:
                blk += b'\x00\x00\x01\x02\x00\x00\x00\x01'
            else:
                blk += b'\x00\x00'
            color_blocks.append(blk)
        else:
            color_id = 2 * idx + 1
            blk = (
                b'SliF1rloC\x01'
                + struct.pack('<I', color_id)
                + b'\x01ABGRDloc_'
                + float_bytes
            )
            if idx < n_colors - 1:
                next_id = 2 * idx + 2
                blk += b'\x00\x00\x01' + struct.pack('<I', next_id) + b'\x01'
            else:
                blk += b'\x00\x00'
            color_blocks.append(blk)

    colors_data = b''.join(color_blocks)

    names_entries = []
    for c in color_list:
        hex_name = f'#{c[0]:02X}{c[1]:02X}{c[2]:02X}'.encode('ascii')
        names_entries.append(struct.pack('<I', len(hex_name)) + hex_name)
    names_payload = b''.join(names_entries)

    vnap_block = (
        b'\xabVNaP'
        + struct.pack('<II', len(names_payload), n_colors)
        + names_payload
        + b'\x00\xff\xff\xff\xff'
    )

    payload = name_sec + colors_data + vnap_block

    off1 = 80 + len(payload)
    fsize = off1 + 115
    off2 = off1 - 80

    timestamp = int(_time.time()) & 0xFFFFFFFF
    ts_bytes = struct.pack('<Q', timestamp)

    file_header = (
        b'\x00\xffKA\x0c\x00\x00\x00htwS#Inf'
        + struct.pack('<QQQ', off1, fsize, off2)
        + b'\x00' * 8
        + ts_bytes
        + b'\x02\x00\x00\x00\x02\x00\x00\x00'
        + b'Prot\x04\x00\x00\x00#Fil\x00\xffKS'
    )

    hash_val = (timestamp ^ (off2 & 0xFFFFFFFF)) & 0xFFFFFFFF
    h_bytes = struct.pack('<I', hash_val)

    footer = (
        b'#FT4'
        + b'\x00' * 8
        + ts_bytes
        + struct.pack('<QQ', fsize, off2)
        + b'\x00' * 8
        + struct.pack('<Q', 1)
        + b'\x38\x00\x00\x00\x00\x00\x00\x01'
        + b'\x00\x00\x00\x00'
        + b'\x48\x00\x00\x00\x00\x00\x00\x00'
        + struct.pack('<QQ', off2, off2)
        + h_bytes
        + b'\x00\x20\x00\x00\x00'
        + h_bytes
        + b'\x0c\x00Swatches.dat'
    )

    return file_header + payload + footer

