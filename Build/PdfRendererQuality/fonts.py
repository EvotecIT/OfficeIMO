"""Synthetic, license-free test fonts for the PDF renderer quality corpus.

Glyph outlines are simple geometric shapes drawn here, so the resulting Type 1, CFF and TrueType
programs carry no third-party font license. The Type 1 writer exercises subroutines, flex,
hint replacement, the div operator and seac accents; the CFF and TrueType programs are built
with fontTools (a development-only dependency of this generator).
"""

import math

UNITS = 1000  # design grid of the glyph descriptions below

# Each glyph: (advance, contours). A contour is a list of commands:
#   ('M', x, y) move, ('L', x, y) line, ('C', x1, y1, x2, y2, x3, y3) cubic curve.
# Outer contours are clockwise in a y-up system for TrueType; CFF and Type 1 use the opposite
# direction, which nonzero filling tolerates because each glyph keeps one consistent orientation.


def _rect(x0, y0, x1, y1):
    return [('M', x0, y0), ('L', x0, y1), ('L', x1, y1), ('L', x1, y0)]


def _circle(cx, cy, r, clockwise=True):
    k = 0.5522847498 * r
    if clockwise:
        return [('M', cx, cy + r),
                ('C', cx + k, cy + r, cx + r, cy + k, cx + r, cy),
                ('C', cx + r, cy - k, cx + k, cy - r, cx, cy - r),
                ('C', cx - k, cy - r, cx - r, cy - k, cx - r, cy),
                ('C', cx - r, cy + k, cx - k, cy + r, cx, cy + r)]
    return [('M', cx, cy + r),
            ('C', cx - k, cy + r, cx - r, cy + k, cx - r, cy),
            ('C', cx - r, cy - k, cx - k, cy - r, cx, cy - r),
            ('C', cx + k, cy - r, cx + r, cy - k, cx + r, cy),
            ('C', cx + r, cy + k, cx + k, cy + r, cx, cy + r)]


GLYPHS = {
    'space': (280, []),
    'H': (700, [[('M', 80, 0), ('L', 80, 700), ('L', 190, 700), ('L', 190, 410), ('L', 510, 410), ('L', 510, 700),
                 ('L', 620, 700), ('L', 620, 0), ('L', 510, 0), ('L', 510, 300), ('L', 190, 300), ('L', 190, 0)]]),
    'O': (760, [_circle(380, 350, 340, True), _circle(380, 350, 230, False)]),
    'A': (720, [[('M', 20, 0), ('L', 300, 700), ('L', 420, 700), ('L', 700, 0), ('L', 580, 0), ('L', 510, 190),
                 ('L', 210, 190), ('L', 140, 0)],
                [('M', 250, 290), ('L', 470, 290), ('L', 360, 590)]]),
    'E': (620, [[('M', 80, 0), ('L', 80, 700), ('L', 570, 700), ('L', 570, 600), ('L', 190, 600), ('L', 190, 410),
                 ('L', 520, 410), ('L', 520, 310), ('L', 190, 310), ('L', 190, 100), ('L', 580, 100), ('L', 580, 0)]]),
    'T': (640, [[('M', 270, 0), ('L', 270, 600), ('L', 30, 600), ('L', 30, 700), ('L', 610, 700), ('L', 610, 600),
                 ('L', 370, 600), ('L', 370, 0)]]),
    'L': (580, [[('M', 80, 0), ('L', 80, 700), ('L', 190, 700), ('L', 190, 100), ('L', 560, 100), ('L', 560, 0)]]),
    'X': (680, [[('M', 20, 0), ('L', 270, 350), ('L', 30, 700), ('L', 160, 700), ('L', 340, 440), ('L', 520, 700),
                 ('L', 650, 700), ('L', 410, 350), ('L', 660, 0), ('L', 530, 0), ('L', 340, 260), ('L', 150, 0)]]),
    'D': (720, [[('M', 80, 0), ('L', 80, 700), ('L', 330, 700),
                 ('C', 530, 700, 660, 560, 660, 350), ('C', 660, 140, 530, 0, 330, 0)],
                [('M', 190, 100), ('L', 320, 100), ('C', 470, 100, 545, 200, 545, 350),
                 ('C', 545, 500, 470, 600, 320, 600), ('L', 190, 600)]]),
    'one': (560, [[('M', 230, 0), ('L', 230, 560), ('L', 110, 500), ('L', 110, 600), ('L', 270, 700), ('L', 340, 700),
                   ('L', 340, 0)]]),
    'two': (560, [[('M', 60, 0), ('L', 60, 90), ('L', 340, 380), ('C', 400, 440, 420, 490, 420, 540),
                   ('C', 420, 600, 370, 630, 300, 630), ('C', 230, 630, 170, 600, 120, 550), ('L', 70, 630),
                   ('C', 140, 690, 220, 720, 310, 720), ('C', 450, 720, 530, 640, 530, 540),
                   ('C', 530, 470, 500, 410, 420, 330), ('L', 200, 100), ('L', 540, 100), ('L', 540, 0)]]),
    'period': (280, [_rect(80, 0, 200, 120)]),
    'acute': (400, [[('M', 150, 760), ('L', 250, 900), ('L', 360, 900), ('L', 220, 760)]]),
}

# Glyph order shared by every font program; .notdef first.
GLYPH_ORDER = ['.notdef', 'space', 'H', 'O', 'A', 'E', 'T', 'L', 'X', 'D', 'one', 'two', 'period', 'acute', 'Aacute']
NOTDEF = (500, [_rect(50, 0, 450, 700), [('M', 100, 50), ('L', 400, 50), ('L', 400, 650), ('L', 100, 650)]])


def outlines(name):
    if name == '.notdef':
        return NOTDEF
    if name == 'Aacute':
        advance, base = GLYPHS['A']
        _, accent = GLYPHS['acute']
        dx = 160
        moved = [[(c[0],) + tuple(v + (dx if i % 2 == 0 else 0) for i, v in enumerate(c[1:])) for c in contour] for contour in accent]
        return advance, base + moved
    return GLYPHS[name]


# ---------------------------------------------------------------- Type 1 ---

T1_OPS = {
    'hstem': [1], 'vstem': [3], 'vmoveto': [4], 'rlineto': [5], 'hlineto': [6], 'vlineto': [7], 'rrcurveto': [8],
    'closepath': [9], 'callsubr': [10], 'return': [11], 'hsbw': [13], 'endchar': [14], 'rmoveto': [21],
    'hmoveto': [22], 'vhcurveto': [30], 'hvcurveto': [31], 'dotsection': [12, 0], 'vstem3': [12, 1],
    'hstem3': [12, 2], 'seac': [12, 6], 'sbw': [12, 7], 'div': [12, 12], 'callothersubr': [12, 16], 'pop': [12, 17],
    'setcurrentpoint': [12, 33],
}


def t1_number(value):
    value = int(value)
    if -107 <= value <= 107:
        return bytes([value + 139])
    if 108 <= value <= 1131:
        value -= 108
        return bytes([(value >> 8) + 247, value & 0xFF])
    if -1131 <= value <= -108:
        value = -value - 108
        return bytes([(value >> 8) + 251, value & 0xFF])
    return bytes([255]) + (value & 0xFFFFFFFF).to_bytes(4, 'big')


def t1_program(tokens):
    out = bytearray()
    for token in tokens:
        out += bytes(T1_OPS[token]) if isinstance(token, str) else t1_number(token)
    return bytes(out)


def encrypt(data, key, lead=4):
    r = key
    out = bytearray()
    for b in bytes(lead) + data:
        c = b ^ (r >> 8)
        r = ((c + r) * 52845 + 22719) & 0xFFFF
        out.append(c)
    return bytes(out)


def _t1_contours(contours, scale, start, tokens):
    x, y = start
    for contour in contours:
        for command in contour:
            pts = [int(round(v * scale)) for v in command[1:]]
            if command[0] == 'M':
                tokens += [pts[0] - x, pts[1] - y, 'rmoveto']
                x, y = pts
            elif command[0] == 'L':
                dx, dy = pts[0] - x, pts[1] - y
                if dy == 0:
                    tokens += [dx, 'hlineto']
                elif dx == 0:
                    tokens += [dy, 'vlineto']
                else:
                    tokens += [dx, dy, 'rlineto']
                x, y = pts
            else:
                tokens += [pts[0] - x, pts[1] - y, pts[2] - pts[0], pts[3] - pts[1], pts[4] - pts[2], pts[5] - pts[3], 'rrcurveto']
                x, y = pts[4], pts[5]
        tokens.append('closepath')
    return x, y


def build_type1(font_name, scale=2):
    """Returns (pfa_cleartext_bytes, eexec_binary_bytes) for a Type 1 font with 1000*scale units/em."""
    subrs = [
        t1_program([3, 0, 'callothersubr', 'pop', 'pop', 'setcurrentpoint', 'return']),
        t1_program([0, 1, 'callothersubr', 'return']),
        t1_program([0, 2, 'callothersubr', 'return']),
        t1_program(['return']),
        # 4: replacement hints
        t1_program([0, 100 * scale, 'hstem', 'return']),
        # 5: the period square, drawn relative to the current point (a shared subroutine)
        t1_program([120 * scale, 'hlineto', 120 * scale, 'vlineto', -120 * scale, 'hlineto', 'closepath', 'return']),
    ]
    charstrings = {}
    for name in GLYPH_ORDER:
        advance, contours = outlines(name)
        tokens = [0, advance * scale, 'hsbw']
        if name == 'Aacute':
            # seac: asb adx ady bchar achar (StandardEncoding codes of A and acute)
            tokens = [0, advance * scale, 'hsbw', 0, 160 * scale, 0, 65, 194, 'seac']
        elif name == 'period':
            tokens += [80 * scale, 0, 'rmoveto', 5, 'callsubr', 'endchar']
        elif name == 'D':
            # Flex the outer bowl: reference point and two curves through the six flex points.
            s = scale
            tokens += [4, 1, 3, 'callothersubr', 'pop', 'callsubr']  # hint replacement via subr 4
            tokens += [80 * s, 0, 'rmoveto', 700 * s, 'vlineto', 250 * s, 'hlineto']
            tokens += [1, 'callsubr']
            tokens += [330 * s, -350 * s, 'rmoveto', 2, 'callsubr']  # reference point (660, 350)
            # flex points relative to the previous point
            pts = [(530, 700), (660, 560), (660, 350), (660, 140), (530, 0), (330, 0)]
            prev = (660 * s, 350 * s)
            for px, py in pts:
                tokens += [(px * s) - prev[0], (py * s) - prev[1], 'rmoveto', 2, 'callsubr']
                prev = (px * s, py * s)
            tokens += [50, 330 * s, 0, 0, 'callsubr']
            tokens += ['closepath']
            x, y = 330 * s, 0
            _t1_contours([contours[1]], scale, (x, y), tokens)
            tokens.append('endchar')
        elif name == 'two':
            # Exercise div: (120 * scale) 2 div is the first move's x offset.
            first = contours[0][0]
            tokens += [first[1] * scale * 2, 2, 'div', 0, 'rmoveto']
            rest = [('L',) + command[1:] if command[0] == 'M' else command for command in contours[0][1:]]
            _t1_contours([rest], scale, (first[1] * scale, 0), tokens)
            tokens.append('endchar')
        else:
            _t1_contours(contours, scale, (0, 0), tokens)
            tokens.append('endchar')
        charstrings[name] = t1_program(tokens)

    encoding_lines = []
    custom_codes = {'H': 72, 'O': 79, 'A': 65, 'E': 69, 'T': 84, 'L': 76, 'X': 88, 'D': 68, 'one': 49, 'two': 50,
                    'period': 46, 'space': 32, 'acute': 194, 'Aacute': 193}
    for name, code in sorted(custom_codes.items(), key=lambda item: item[1]):
        encoding_lines.append('dup %d /%s put' % (code, name))
    m = 1.0 / (1000 * scale)
    clear = ('%!PS-AdobeFont-1.0: {name} 001.000\n'
             '12 dict begin\n'
             '/FontInfo 2 dict dup begin /FullName ({name}) readonly def end readonly def\n'
             '/FontName /{name} def\n'
             '/PaintType 0 def\n'
             '/FontType 1 def\n'
             '/FontMatrix [{m} 0 0 {m} 0 0] readonly def\n'
             '/FontBBox {{0 {bmin} {bmax} {btop}}} readonly def\n'
             '/Encoding 256 array\n0 1 255 {{1 index exch /.notdef put}} for\n{enc}\nreadonly def\n'
             'currentdict end\ncurrentfile eexec\n').format(name=font_name, m=repr(m), bmin=-50 * scale, bmax=760 * scale,
                                                           btop=900 * scale, enc='\n'.join(encoding_lines))
    priv = bytearray()
    priv += b'dup /Private 8 dict dup begin\n/RD{string currentfile exch readstring pop}executeonly def\n'
    priv += b'/ND{noaccess def}executeonly def\n/NP{noaccess put}executeonly def\n/lenIV 4 def\n'
    priv += b'/BlueValues [-10 0 700 710] def\n/MinFeature{16 16} def\n/password 5839 def\n'
    priv += ('/Subrs %d array\n' % len(subrs)).encode()
    for index, subr in enumerate(subrs):
        enc = encrypt(subr, 4330)
        priv += ('dup %d %d RD ' % (index, len(enc))).encode() + enc + b' NP\n'
    priv += b'ND\n2 index /CharStrings %d dict dup begin\n' % len(charstrings)
    for name, program in charstrings.items():
        enc = encrypt(program, 4330)
        priv += ('/%s %d RD ' % (name, len(enc))).encode() + enc + b' ND\n'
    priv += b'end\nend\nreadonly put\nnoaccess put\ndup /FontName get exch definefont pop\nmark currentfile closefile\n'
    binary = encrypt(bytes(priv), 55665)
    trailer = ('0' * 64 + '\n') * 8 + 'cleartomark\n'
    return clear.encode('latin-1'), binary, trailer.encode('latin-1')


# ---------------------------------------------------- CFF and TrueType ---

def _draw(pen, contours, flip=False):
    for contour in contours:
        commands = contour
        for command in commands:
            if command[0] == 'M':
                pen.moveTo(command[1:3])
            elif command[0] == 'L':
                pen.lineTo(command[1:3])
            else:
                pen.curveTo(command[1:3], command[3:5], command[5:7])
        pen.closePath()


def build_otf(font_name, upm=1000):
    from fontTools.fontBuilder import FontBuilder
    from fontTools.pens.t2CharStringPen import T2CharStringPen
    fb = FontBuilder(upm, isTTF=False)
    fb.setupGlyphOrder(GLYPH_ORDER)
    cmap = {ord('H'): 'H', ord('O'): 'O', ord('A'): 'A', ord('E'): 'E', ord('T'): 'T', ord('L'): 'L', ord('X'): 'X',
            ord('D'): 'D', ord('1'): 'one', ord('2'): 'two', ord('.'): 'period', ord(' '): 'space', 0xB4: 'acute', 0xC1: 'Aacute'}
    fb.setupCharacterMap(cmap)
    charstrings = {}
    metrics = {}
    for name in GLYPH_ORDER:
        advance, contours = outlines(name)
        pen = T2CharStringPen(advance, None)
        _draw(pen, contours)
        charstrings[name] = pen.getCharString()
        metrics[name] = (advance, 0)
    fb.setupCFF(font_name, {'FullName': font_name}, charstrings, {})
    fb.setupHorizontalMetrics(metrics)
    fb.setupHorizontalHeader(ascent=900, descent=-200)
    fb.setupNameTable({'familyName': font_name, 'styleName': 'Regular'})
    fb.setupOS2(sTypoAscender=900, usWinAscent=900, usWinDescent=200)
    fb.setupPost()
    return fb.font


def build_ttf(font_name):
    from fontTools.fontBuilder import FontBuilder
    from fontTools.pens.ttGlyphPen import TTGlyphPen
    from fontTools.pens.cu2quPen import Cu2QuPen
    fb = FontBuilder(1000, isTTF=True)
    fb.setupGlyphOrder(GLYPH_ORDER)
    cmap = {ord('H'): 'H', ord('O'): 'O', ord('A'): 'A', ord('E'): 'E', ord('T'): 'T', ord('L'): 'L', ord('X'): 'X',
            ord('D'): 'D', ord('1'): 'one', ord('2'): 'two', ord('.'): 'period', ord(' '): 'space', 0xB4: 'acute', 0xC1: 'Aacute'}
    fb.setupCharacterMap(cmap)
    glyphs = {}
    metrics = {}
    for name in GLYPH_ORDER:
        advance, contours = outlines(name)
        pen = TTGlyphPen(None)
        _draw(Cu2QuPen(pen, 1.0, reverse_direction=True), contours)
        glyphs[name] = pen.glyph()
        xs = [v for contour in contours for command in contour for i, v in enumerate(command[1:]) if i % 2 == 0]
        metrics[name] = (advance, int(min(xs)) if xs else 0)
    fb.setupGlyf(glyphs)
    fb.setupHorizontalMetrics(metrics)
    fb.setupHorizontalHeader(ascent=900, descent=-200)
    fb.setupNameTable({'familyName': font_name, 'styleName': 'Regular'})
    fb.setupOS2(sTypoAscender=900, usWinAscent=900, usWinDescent=200)
    fb.setupPost()
    return fb.font


def font_bytes(font):
    import io
    buffer = io.BytesIO()
    font.save(buffer)
    return buffer.getvalue()


def widths(names, scale=1.0):
    return [outlines(n)[0] * scale for n in names]
