"""Minimal JBIG2 (ITU-T T.88) encoder for renderer test fixtures.

Encodes generic regions (MQ arithmetic with templates 0 to 3, typical prediction and adaptive pixels, or MMR via
Pillow's Group 4 TIFF writer), and arithmetic symbol dictionaries with text regions. The output is embedded in PDF
JBIG2Decode streams so an independent decoder (Poppler) can render reference images.
"""

import io
import struct

QE = [
    (0x5601, 1, 1, 1), (0x3401, 2, 6, 0), (0x1801, 3, 9, 0), (0x0AC1, 4, 12, 0), (0x0521, 5, 29, 0), (0x0221, 38, 33, 0),
    (0x5601, 7, 6, 1), (0x5401, 8, 14, 0), (0x4801, 9, 14, 0), (0x3801, 10, 14, 0), (0x3001, 11, 17, 0), (0x2401, 12, 18, 0),
    (0x1C01, 13, 20, 0), (0x1601, 29, 21, 0), (0x5601, 15, 14, 1), (0x5401, 16, 14, 0), (0x5101, 17, 15, 0), (0x4801, 18, 16, 0),
    (0x3801, 19, 17, 0), (0x3401, 20, 18, 0), (0x3001, 21, 19, 0), (0x2801, 22, 19, 0), (0x2401, 23, 20, 0), (0x2201, 24, 21, 0),
    (0x1C01, 25, 22, 0), (0x1801, 26, 23, 0), (0x1601, 27, 24, 0), (0x1401, 28, 25, 0), (0x1201, 29, 26, 0), (0x1101, 30, 27, 0),
    (0x0AC1, 31, 28, 0), (0x09C1, 32, 29, 0), (0x08A1, 33, 30, 0), (0x0521, 34, 31, 0), (0x0441, 35, 32, 0), (0x02A1, 36, 33, 0),
    (0x0221, 37, 34, 0), (0x0141, 38, 35, 0), (0x0111, 39, 36, 0), (0x0085, 40, 37, 0), (0x0049, 41, 38, 0), (0x0025, 42, 39, 0),
    (0x0015, 43, 40, 0), (0x0009, 44, 41, 0), (0x0005, 45, 42, 0), (0x0001, 45, 43, 0), (0x5601, 46, 46, 0),
]


class MQEncoder:
    """MQ encoder (T.88 E.2 / T.800 C.2)."""

    def __init__(self):
        self.a = 0x8000
        self.c = 0
        self.ct = 12
        self.out = bytearray()

    def _b(self):
        return self.out[-1] if self.out else 0

    def _byte_out(self):
        if self.out and self.out[-1] == 0xFF:
            self.out.append((self.c >> 20) & 0xFF)
            self.c &= 0xFFFFF
            self.ct = 7
        elif self.c < 0x8000000:
            self.out.append((self.c >> 19) & 0xFF)
            self.c &= 0x7FFFF
            self.ct = 8
        else:
            if self.out:
                self.out[-1] += 1
            if self.out and self.out[-1] == 0xFF:
                self.c &= 0x7FFFFFF
                self.out.append((self.c >> 20) & 0xFF)
                self.c &= 0xFFFFF
                self.ct = 7
            else:
                self.out.append((self.c >> 19) & 0xFF)
                self.c &= 0x7FFFF
                self.ct = 8

    def _renorm(self):
        while True:
            self.a = (self.a << 1) & 0xFFFF
            self.c = (self.c << 1) & 0xFFFFFFFF
            self.ct -= 1
            if self.ct == 0:
                self._byte_out()
            if self.a & 0x8000:
                break

    def encode(self, contexts, cx, bit):
        state = contexts.get(cx, 0)
        index, mps = state >> 1, state & 1
        qe, nmps, nlps, switch = QE[index]
        if bit == mps:
            self.a -= qe
            if (self.a & 0x8000) == 0:
                if self.a < qe:
                    self.a = qe
                else:
                    self.c += qe
                index = nmps
                self._renorm()
            else:
                self.c += qe
        else:
            self.a -= qe
            if self.a < qe:
                self.c += qe
            else:
                self.a = qe
            if switch:
                mps = 1 - mps
            index = nlps
            self._renorm()
        contexts[cx] = (index << 1) | mps

    def flush(self):
        temp = self.c + self.a
        self.c |= 0xFFFF
        if self.c >= temp:
            self.c -= 0x8000
        self.c = (self.c << self.ct) & 0xFFFFFFFF
        self._byte_out()
        self.c = (self.c << self.ct) & 0xFFFFFFFF
        self._byte_out()
        # The 0xFFAC marker terminates arithmetic data (E.2.9); decoders read ahead into it.
        return bytes(self.out) + b'\xff\xac'

    def encode_integer(self, contexts, value):
        prev = [1]

        def bits(v, n):
            for i in range(n - 1, -1, -1):
                b = (v >> i) & 1
                self.encode(contexts, prev[0], b)
                prev[0] = (prev[0] << 1) | b if prev[0] < 256 else ((((prev[0] << 1) | b) & 511) | 256)
        if value is None:
            bits(1, 1)
            bits(0, 1)
            bits(0, 2)
            return
        sign = 1 if value < 0 else 0
        m = abs(value)
        bits(sign, 1)
        for prefix_len, low, width in ((1, 0, 2), (2, 4, 4), (3, 20, 6), (4, 84, 8), (5, 340, 12), (5, 4436, 32)):
            if width == 32 or m < low + (1 << width):
                if prefix_len == 5 and width == 32:
                    bits(0b11111, 5)
                elif prefix_len == 5:
                    bits(0b11110, 5)
                else:
                    bits(((1 << prefix_len) - 2), prefix_len)
                bits(m - low, width)
                return

    def encode_id(self, contexts, value, length):
        prev = 1
        for i in range(length - 1, -1, -1):
            b = (value >> i) & 1
            self.encode(contexts, prev, b)
            prev = (prev << 1) | b


GENERIC_TEMPLATES = [
    [(-1, -2), (0, -2), (1, -2), (-2, -1), (-1, -1), (0, -1), (1, -1), (2, -1), (-4, 0), (-3, 0), (-2, 0), (-1, 0)],
    [(-1, -2), (0, -2), (1, -2), (2, -2), (-2, -1), (-1, -1), (0, -1), (1, -1), (2, -1), (-3, 0), (-2, 0), (-1, 0)],
    [(-1, -2), (0, -2), (1, -2), (-2, -1), (-1, -1), (0, -1), (1, -1), (-2, 0), (-1, 0)],
    [(-3, -1), (-2, -1), (-1, -1), (0, -1), (1, -1), (-4, 0), (-3, 0), (-2, 0), (-1, 0)],
]
DEFAULT_AT = [[(3, -1), (-3, -1), (2, -2), (-2, -2)], [(3, -1)], [(2, -1)], [(2, -1)]]
TPGDON_CONTEXT = [0x9B25, 0x0795, 0x00E5, 0x0195]


def context_offsets(template, at):
    pixels = [(x, y, -1) for x, y in GENERIC_TEMPLATES[template]]
    pixels += [(x, y, i) for i, (x, y) in enumerate(DEFAULT_AT[template])]
    pixels.sort(key=lambda p: (p[1], p[0]))
    return [at[p[2]] if p[2] >= 0 else (p[0], p[1]) for p in pixels]


def encode_generic(encoder, contexts, bitmap, template, tpgdon, at):
    height = len(bitmap)
    width = len(bitmap[0]) if height else 0
    offsets = context_offsets(template, at)

    def pixel(x, y):
        return bitmap[y][x] if 0 <= x < width and 0 <= y < height else 0
    ltp = 0
    for y in range(height):
        if tpgdon:
            typical = 1 if y > 0 and bitmap[y] == bitmap[y - 1] else 0
            encoder.encode(contexts, TPGDON_CONTEXT[template], typical ^ ltp)
            ltp = typical
            if typical:
                continue
        for x in range(width):
            cx = 0
            for dx, dy in offsets:
                cx = (cx << 1) | pixel(x + dx, y + dy)
            encoder.encode(contexts, cx, bitmap[y][x])


def segment(number, kind, data, referred=(), page=1):
    header = struct.pack('>IB', number, kind)
    header += bytes([(len(referred) << 5) & 0xE0])
    header += bytes(referred)
    header += bytes([page])
    header += struct.pack('>I', len(data))
    return header + data


def page_info(width, height):
    return struct.pack('>IIIIBH', width, height, 0, 0, 0, 0)


def region_info(width, height, x, y, operator=0):
    return struct.pack('>IIIIB', width, height, x, y, operator)


def generic_region(bitmap, x, y, template=0, tpgdon=False, at=None, operator=0):
    at = at or DEFAULT_AT[template]
    height = len(bitmap)
    width = len(bitmap[0])
    flags = (template << 1) | (8 if tpgdon else 0)
    data = region_info(width, height, x, y, operator) + bytes([flags])
    for ax, ay in at:
        data += struct.pack('>bb', ax, ay)
    encoder = MQEncoder()
    encode_generic(encoder, {}, bitmap, template, tpgdon, at)
    return data + encoder.flush()


def mmr_region(bitmap, x, y, operator=0):
    from PIL import Image
    height = len(bitmap)
    width = len(bitmap[0])
    image = Image.new('1', (width, height), 1)
    for row in range(height):
        for column in range(width):
            if bitmap[row][column]:
                image.putpixel((column, row), 0)
    buffer = io.BytesIO()
    image.save(buffer, 'TIFF', compression='group4')
    tiff = buffer.getvalue()
    endian = '<' if tiff[:2] == b'II' else '>'
    offset = struct.unpack(endian + 'I', tiff[4:8])[0]
    count = struct.unpack(endian + 'H', tiff[offset:offset + 2])[0]
    tags = {}
    for i in range(count):
        tag, kind, n, value = struct.unpack(endian + 'HHII', tiff[offset + 2 + i * 12: offset + 14 + i * 12])
        if kind == 3 and n == 1:
            value = value & 0xFFFF if endian == '<' else value >> 16
        tags[tag] = value
    if tags.get(278, height) < height:
        raise ValueError('MMR fixture must be a single strip')
    strip = tiff[tags[273]:tags[273] + tags[279]]
    data = region_info(width, height, x, y, operator) + bytes([1])
    return data + strip


def symbol_dictionary(symbols):
    """Encodes bitmaps as one arithmetic symbol dictionary exporting every symbol (order: by height then input order)."""
    order = sorted(range(len(symbols)), key=lambda i: (len(symbols[i]), i))
    encoder = MQEncoder()
    integers = {}
    generic = {}

    def ctx(name):
        return integers.setdefault(name, {})
    height = 0
    index = 0
    while index < len(order):
        class_height = len(symbols[order[index]])
        encoder.encode_integer(ctx('IADH'), class_height - height)
        height = class_height
        width = 0
        while index < len(order) and len(symbols[order[index]]) == class_height:
            bitmap = symbols[order[index]]
            encoder.encode_integer(ctx('IADW'), len(bitmap[0]) - width)
            width = len(bitmap[0])
            encode_generic(encoder, generic, bitmap, 0, False, DEFAULT_AT[0])
            index += 1
        encoder.encode_integer(ctx('IADW'), None)
    encoder.encode_integer(ctx('IAEX'), 0)
    encoder.encode_integer(ctx('IAEX'), len(symbols))
    flags = 0  # arithmetic, no refinement, template 0
    data = struct.pack('>H', flags)
    for ax, ay in DEFAULT_AT[0]:
        data += struct.pack('>bb', ax, ay)
    data += struct.pack('>II', len(symbols), len(symbols))
    return data + encoder.flush(), order


def text_region(width, height, x, y, placements, symbol_count, operator=0):
    """placements: (symbol index in dictionary export order, left, top); reference corner TOPLEFT, one strip per row."""
    encoder = MQEncoder()
    integers = {}

    def ctx(name):
        return integers.setdefault(name, {})
    code_length = 0
    while (1 << code_length) < symbol_count:
        code_length += 1
    id_contexts = {}
    encoder.encode_integer(ctx('IADT'), 0)
    strip_t = 0
    first_s = 0
    rows = {}
    for symbol, left, top, w in placements:
        rows.setdefault(top, []).append((left, symbol, w))
    for top in sorted(rows):
        encoder.encode_integer(ctx('IADT'), top - strip_t)
        strip_t = top
        items = sorted(rows[top])
        current = None
        for left, symbol, w in items:
            if current is None:
                encoder.encode_integer(ctx('IAFS'), left - first_s)
                first_s = left
            else:
                encoder.encode_integer(ctx('IADS'), left - current)
            encoder.encode_id(id_contexts, symbol, code_length)
            current = left + w - 1
        encoder.encode_integer(ctx('IADS'), None)
    flags = (1 << 4)  # arithmetic, no refinement, one strip, TOPLEFT, OR
    data = region_info(width, height, x, y, operator) + struct.pack('>H', flags) + struct.pack('>I', len(placements))
    return data + encoder.flush()
