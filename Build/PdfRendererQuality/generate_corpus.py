#!/usr/bin/env python3
"""Generates the PDF renderer quality corpus and its Poppler baselines.

Each page isolates one rendering feature with deterministic, embedded-only content, so OfficeIMO's
managed rasterizer can be compared with Poppler pdftoppm on any host. The script writes the PDFs,
the pdftoppm PNG baselines and manifest.json into OfficeIMO.Pdf.Tests/Pdf/RendererQuality.

Requirements (development only): Python 3.9+, fontTools, and Poppler's pdftoppm on PATH.

    python3 Build/PdfRendererQuality/generate_corpus.py

Regenerate after changing a page; then rerun the PdfRendererQualityGateTests and update the
thresholds in manifest.json only when a change is an intended improvement or an accepted trade-off.
"""

import hashlib
import json
import math
import os
import struct
import subprocess
import sys
import zlib

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fonts  # noqa: E402

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT = os.path.join(ROOT, 'OfficeIMO.Pdf.Tests', 'Pdf', 'RendererQuality')
DPI = 150


class Pdf:
    def __init__(self):
        self.objects = {}
        self.next = 1

    def reserve(self):
        number = self.next
        self.next += 1
        return number

    def add(self, body, number=None):
        number = number or self.reserve()
        self.objects[number] = body if isinstance(body, bytes) else body.encode('latin-1')
        return number

    def stream(self, data, extra='', compress=True, number=None):
        if isinstance(data, str):
            data = data.encode('latin-1')
        if compress:
            data = zlib.compress(data, 9)
            extra += ' /Filter /FlateDecode'
        return self.add(('<< /Length %d %s >>\nstream\n' % (len(data), extra)).encode('latin-1') + data + b'\nendstream',
                        number)

    def write(self, path, root):
        out = bytearray(b'%PDF-1.7\n%\xe2\xe3\xcf\xd3\n')
        offsets = {}
        for number in sorted(self.objects):
            offsets[number] = len(out)
            out += b'%d 0 obj\n' % number + self.objects[number] + b'\nendobj\n'
        xref = len(out)
        size = max(self.objects) + 1
        out += b'xref\n0 %d\n0000000000 65535 f \n' % size
        for number in range(1, size):
            out += (b'%010d 00000 n \n' % offsets[number]) if number in offsets else b'0000000000 65535 f \n'
        out += b'trailer\n<< /Size %d /Root %d 0 R >>\nstartxref\n%d\n%%%%EOF\n' % (size, root, xref)
        with open(path, 'wb') as handle:
            handle.write(bytes(out))


def single_page(pdf, width, height, content, resources, extra_catalog=''):
    content_ref = pdf.stream(content)
    pages = pdf.reserve()
    page = pdf.add('<< /Type /Page /Parent %d 0 R /MediaBox [0 0 %s %s] /Contents %d 0 R /Resources %s >>'
                   % (pages, width, height, content_ref, resources))
    pdf.add('<< /Type /Pages /Kids [%d 0 R] /Count 1 >>' % page, pages)
    return pdf.add('<< /Type /Catalog /Pages %d 0 R %s >>' % (pages, extra_catalog)), page


def pdf_string(text):
    return '(' + text.replace('\\', '\\\\').replace('(', '\\(').replace(')', '\\)') + ')'


CODES = {'H': 72, 'O': 79, 'A': 65, 'E': 69, 'T': 84, 'L': 76, 'X': 88, 'D': 68, 'one': 49, 'two': 50, 'period': 46,
         'space': 32, 'acute': 194, 'Aacute': 193}


def simple_widths(first, last, code_names):
    out = []
    for code in range(first, last + 1):
        name = code_names.get(code)
        out.append(str(int(fonts.outlines(name)[0])) if name else '0')
    return '[' + ' '.join(out) + ']'


def text_block(font, lines, x=24, y=0, top=None):
    ops = ['BT']
    for size, text, ty in lines:
        ops.append('/%s %s Tf 1 0 0 1 %s %s Tm %s Tj' % (font, size, x, ty, pdf_string(text)))
    ops.append('ET')
    return '\n'.join(ops)


def page_type1(path):
    pdf = Pdf()
    clear, binary, trailer = fonts.build_type1('OIMOTestType1')
    program = pdf.stream(clear + binary + trailer, '/Length1 %d /Length2 %d /Length3 %d' % (len(clear), len(binary), len(trailer)))
    descriptor = pdf.add('<< /Type /FontDescriptor /FontName /OIMOTestType1 /Flags 4 /FontBBox [0 -50 760 900] '
                         '/ItalicAngle 0 /Ascent 900 /Descent -200 /CapHeight 700 /StemV 100 /FontFile %d 0 R >>' % program)
    code_names = {code: name for name, code in CODES.items()}
    widths = simple_widths(32, 194, code_names)
    builtin = pdf.add('<< /Type /Font /Subtype /Type1 /BaseFont /OIMOTestType1 /FirstChar 32 /LastChar 194 '
                      '/Widths %s /FontDescriptor %d 0 R >>' % (widths, descriptor))
    # The same program with Differences remapping a and b to O and H.
    diff_names = dict(code_names)
    diff_names[97] = 'O'
    diff_names[98] = 'H'
    remapped = pdf.add('<< /Type /Font /Subtype /Type1 /BaseFont /OIMOTestType1 /FirstChar 32 /LastChar 194 '
                       '/Widths %s /Encoding << /Type /Encoding /Differences [97 /O /H] >> /FontDescriptor %d 0 R >>'
                       % (simple_widths(32, 194, diff_names), descriptor))
    content = '\n'.join([
        text_block('F1', [(9, 'HOTEL TAXE 12.', 250), (18, 'HOTEL TAXE', 220), (36, 'DOOH 21', 170), (60, '\xc1HO', 95)]),
        text_block('F2', [(30, 'abba HE', 30)]),
    ])
    root, _ = single_page(pdf, 320, 280, content, '<< /Font << /F1 %d 0 R /F2 %d 0 R >> >>' % (builtin, remapped))
    pdf.write(path, root)


def page_type1c(path):
    pdf = Pdf()
    otf = fonts.build_otf('OIMOTestCFF')
    cff = fonts.font_bytes_table(otf, 'CFF ')
    program = pdf.stream(cff, '/Subtype /Type1C')
    descriptor = pdf.add('<< /Type /FontDescriptor /FontName /OIMOTestCFF /Flags 32 /FontBBox [0 -50 760 900] '
                         '/ItalicAngle 0 /Ascent 900 /Descent -200 /CapHeight 700 /StemV 100 /FontFile3 %d 0 R >>' % program)
    win = {code: name for name, code in CODES.items() if name not in ('acute', 'Aacute')}
    win[180] = 'acute'
    win[193] = 'Aacute'
    winansi = pdf.add('<< /Type /Font /Subtype /Type1 /BaseFont /OIMOTestCFF /FirstChar 32 /LastChar 193 /Widths %s '
                      '/Encoding /WinAnsiEncoding /FontDescriptor %d 0 R >>' % (simple_widths(32, 193, win), descriptor))
    diff = {1: 'H', 2: 'O', 3: 'T', 4: 'E', 5: 'L'}
    custom = pdf.add('<< /Type /Font /Subtype /Type1 /BaseFont /OIMOTestCFF /FirstChar 1 /LastChar 5 /Widths %s '
                     '/Encoding << /Differences [1 /H /O /T /E /L] >> /FontDescriptor %d 0 R >>'
                     % (simple_widths(1, 5, diff), descriptor))
    # Name-keyed CFF used as a CIDFontType0 program: CIDs select glyph ids directly.
    cid_program = pdf.stream(cff, '/Subtype /CIDFontType0C')
    cid_descriptor = pdf.add('<< /Type /FontDescriptor /FontName /OIMOTestCID /Flags 32 /FontBBox [0 -50 760 900] '
                             '/ItalicAngle 0 /Ascent 900 /Descent -200 /CapHeight 700 /StemV 100 /FontFile3 %d 0 R >>' % cid_program)
    order = fonts.GLYPH_ORDER
    w = ' '.join('%d [%d]' % (gid, fonts.outlines(name)[0]) for gid, name in enumerate(order))
    cid_font = pdf.add('<< /Type /Font /Subtype /CIDFontType0 /BaseFont /OIMOTestCID /CIDSystemInfo << /Registry (Adobe) '
                       '/Ordering (Identity) /Supplement 0 >> /FontDescriptor %d 0 R /DW 500 /W [%s] >>' % (cid_descriptor, w))
    type0 = pdf.add('<< /Type /Font /Subtype /Type0 /BaseFont /OIMOTestCID /Encoding /Identity-H /DescendantFonts [%d 0 R] >>' % cid_font)
    gids = [order.index(n) for n in ['D', 'O', 'T', 'space', 'two', 'one', 'period', 'X']]
    hex_text = '<' + ''.join('%04X' % g for g in gids) + '>'
    content = '\n'.join([
        text_block('F1', [(9, 'HOTEL TAXE 12.', 250), (18, 'HOTEL TAXE', 220), (36, 'DOOH 21', 170), (60, '\xc1HO', 95)]),
        text_block('F2', [(24, '\x01\x02\x03\x04\x05', 55)]),
        'BT /F3 24 Tf 1 0 0 1 24 22 Tm %s Tj ET' % hex_text,
    ])
    root, _ = single_page(pdf, 320, 280, content, '<< /Font << /F1 %d 0 R /F2 %d 0 R /F3 %d 0 R >> >>' % (winansi, custom, type0))
    pdf.write(path, root)


def page_truetype(path):
    pdf = Pdf()
    ttf = fonts.font_bytes(fonts.build_ttf('OIMOTestTT'))
    program = pdf.stream(ttf, '/Length1 %d' % len(ttf))
    descriptor = pdf.add('<< /Type /FontDescriptor /FontName /OIMOTestTT /Flags 32 /FontBBox [0 -50 760 900] '
                         '/ItalicAngle 0 /Ascent 900 /Descent -200 /CapHeight 700 /StemV 100 /FontFile2 %d 0 R >>' % program)
    win = {code: name for name, code in CODES.items() if name not in ('acute', 'Aacute')}
    win[193] = 'Aacute'
    font = pdf.add('<< /Type /Font /Subtype /TrueType /BaseFont /OIMOTestTT /FirstChar 32 /LastChar 193 /Widths %s '
                   '/Encoding /WinAnsiEncoding /FontDescriptor %d 0 R >>' % (simple_widths(32, 193, win), descriptor))
    content = text_block('F1', [(9, 'HOTEL TAXE 12.', 250), (18, 'HOTEL TAXE', 220), (36, 'DOOH 21', 170), (60, '\xc1HO', 95)])
    root, _ = single_page(pdf, 320, 280, content, '<< /Font << /F1 %d 0 R >> >>' % font)
    pdf.write(path, root)


def page_strokes(path):
    pdf = Pdf()
    c = ['1 1 1 rg 0 0 360 300 re f', '0 0 0 RG']
    for i, w in enumerate([0, 0.1, 0.25, 0.5, 0.75, 1, 1.5, 2]):
        y = 285 - i * 9
        c.append('%s w 12 %s m 170 %s l S' % (w, y, y))           # horizontal rules
        c.append('%s w 12 %s m 170 %s l S' % (w, y - 4, y - 2.5))  # nearly horizontal
    for i, w in enumerate([0, 0.25, 0.5, 1, 2]):
        x = 190 + i * 10
        c.append('%s w %s 290 m %s 210 l S' % (w, x, x))
    c.append('0.4 w 250 210 100 80 re S 0.2 w 262 222 76 56 re S')
    c.append('0 0.3 0.8 rg 250 200 100 0.3 re f 250 196 100 0.6 re f 250 192 0.4 -60 re f')
    c.append('12 w 0 J 0 j 0.7 0 0 RG 20 120 m 60 160 l 100 120 l S')
    c.append('1 J 1 j 0 0.5 0 RG 120 120 m 160 160 l 200 120 l S')
    c.append('2 J 2 j 0 0 0.7 RG 220 120 m 260 160 l 300 120 l S')
    c.append('8 w 0 J 0 j 0.2 g 2 M 20 70 m 60 95 l 100 70 l S 20 M 120 70 m 160 95 l 200 70 l S')
    c.append('3 w [9 4 2 4] 2 d 0.6 0 0.6 RG 20 40 m 340 40 l S [] 0 d')
    c.append('q 0.8 0.3 -0.3 0.8 250 60 cm 1 w 0 0 0 RG 0 0 m 60 0 l S 0 10 m 60 10 l S Q')
    c.append('q 1 0 0.6 1 230 15 cm 2 w 0.8 0.4 0 RG 0 0 m 0 18 l S 40 0 m 80 0 l S Q')
    root, _ = single_page(pdf, 360, 300, '\n'.join(c), '<< >>')
    pdf.write(path, root)


def circle_path(cx, cy, r):
    k = 0.5522847498 * r
    return ('%s %s m %s %s %s %s %s %s c %s %s %s %s %s %s c %s %s %s %s %s %s c %s %s %s %s %s %s c h'
            % (cx + r, cy, cx + r, cy + k, cx + k, cy + r, cx, cy + r, cx - k, cy + r, cx - r, cy + k, cx - r, cy,
               cx - r, cy - k, cx - k, cy - r, cx, cy - r, cx + k, cy - r, cx + r, cy - k, cx + r, cy))


def page_clipping(path):
    pdf = Pdf()
    ttf = fonts.font_bytes(fonts.build_ttf('OIMOTestTT'))
    program = pdf.stream(ttf, '/Length1 %d' % len(ttf))
    descriptor = pdf.add('<< /Type /FontDescriptor /FontName /OIMOTestTT /Flags 32 /FontBBox [0 -50 760 900] '
                         '/ItalicAngle 0 /Ascent 900 /Descent -200 /CapHeight 700 /StemV 100 /FontFile2 %d 0 R >>' % program)
    win = {code: name for name, code in CODES.items() if name not in ('acute', 'Aacute')}
    font = pdf.add('<< /Type /Font /Subtype /TrueType /BaseFont /OIMOTestTT /FirstChar 32 /LastChar 90 /Widths %s '
                   '/Encoding /WinAnsiEncoding /FontDescriptor %d 0 R >>' % (simple_widths(32, 90, win), descriptor))
    stripes = bytes(((x // 4 + y // 4) % 2) * 255 for y in range(64) for x in range(64))
    image = pdf.stream(stripes, '/Type /XObject /Subtype /Image /Width 64 /Height 64 /ColorSpace /DeviceGray /BitsPerComponent 8')
    c = ['1 1 1 rg 0 0 360 260 re f']
    c.append('q %s W n 0.1 0.3 0.8 rg 0 130 180 130 re f Q' % circle_path(90, 195, 55))
    c.append('q %s W n q 120 0 0 120 190 135 cm /Im1 Do Q Q' % circle_path(250, 195, 55))
    c.append('q 20 20 m 160 20 l 90 125 l h 50 60 m 130 60 l 90 110 l h W* n 0.8 0.2 0.1 rg 0 0 180 130 re f Q')
    c.append('q 200 15 140 100 re W n %s W n 0 0.55 0.2 rg 180 0 180 130 re f Q' % circle_path(270, 65, 70))
    c.append('q %s W n BT /F1 40 Tf 1 0 0 1 190 50 Tm 0 0 0 rg (HOTEL) Tj ET Q' % circle_path(260, 65, 38))
    c.append('q 12.3 5.6 50.4 20.2 re W n 0.5 0 0.5 rg 0 0 360 260 re f Q')
    root, _ = single_page(pdf, 360, 260, '\n'.join(c), '<< /Font << /F1 %d 0 R >> /XObject << /Im1 %d 0 R >> >>' % (font, image))
    pdf.write(path, root)


def page_text_modes(path):
    pdf = Pdf()
    otf = fonts.build_otf('OIMOTestCFF')
    cff = fonts.font_bytes_table(otf, 'CFF ')
    program = pdf.stream(cff, '/Subtype /Type1C')
    descriptor = pdf.add('<< /Type /FontDescriptor /FontName /OIMOTestCFF /Flags 32 /FontBBox [0 -50 760 900] '
                         '/ItalicAngle 0 /Ascent 900 /Descent -200 /CapHeight 700 /StemV 100 /FontFile3 %d 0 R >>' % program)
    win = {code: name for name, code in CODES.items() if name not in ('acute', 'Aacute')}
    font = pdf.add('<< /Type /Font /Subtype /Type1 /BaseFont /OIMOTestCFF /FirstChar 32 /LastChar 90 /Widths %s '
                   '/Encoding /WinAnsiEncoding /FontDescriptor %d 0 R >>' % (simple_widths(32, 90, win), descriptor))
    c = ['1 1 1 rg 0 0 360 300 re f']
    c.append('BT /F1 34 Tf 0.1 0.2 0.7 rg 0.8 0 0 RG 1 0 0 1 16 250 Tm 0 Tr (HOTEL) Tj ET')
    c.append('BT /F1 34 Tf 1.5 w 0.8 0 0 RG 1 0 0 1 16 205 Tm 1 Tr (HOTEL) Tj ET')
    c.append('BT /F1 34 Tf 2 w 0.1 0.2 0.7 rg 0.9 0.6 0 RG 1 0 0 1 16 160 Tm 2 Tr (HOTEL) Tj ET')
    c.append('BT /F1 34 Tf 0.4 w 0 g 0 G 1 0 0 1 16 115 Tm 2 Tr (TAXE 12) Tj ET')
    c.append('BT /F1 34 Tf 0 w 0 G 1 0 0 1 16 70 Tm 1 Tr (DOOH) Tj ET')
    c.append('q 0.95 0.3 -0.3 0.95 200 40 cm BT /F1 30 Tf 1 w 0 0.5 0 RG 1 0 0 1 0 0 Tm 1 Tr (LOT) Tj ET Q')
    c.append('q 1 0 0 1 0 0 cm BT /F1 40 Tf 1 0 0 1 16 20 Tm 7 Tr (HOLD) Tj ET 0.9 0.1 0.4 rg 0 0 360 70 re f Q')
    root, _ = single_page(pdf, 360, 300, '\n'.join(c), '<< /Font << /F1 %d 0 R >> >>' % font)
    pdf.write(path, root)


def page_shadings(path):
    pdf = Pdf()
    lin = '<< /FunctionType 2 /Domain [0 1] /C0 [0.9 0.1 0.1] /C1 [0.1 0.2 0.9] /N 1 >>'
    stitch = ('<< /FunctionType 3 /Domain [0 1] /Bounds [0.5] /Encode [0 1 0 1] /Functions ['
              '<< /FunctionType 2 /Domain [0 1] /C0 [1 1 0] /C1 [0 0.6 0.3] /N 1 >> '
              '<< /FunctionType 2 /Domain [0 1] /C0 [0 0.6 0.3] /C1 [0.2 0 0.5] /N 2 >>] >>')
    sh = {}
    sh['A1'] = pdf.add('<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 250 160 250] /Function %s /Extend [true true] >>' % lin)
    sh['A2'] = pdf.add('<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [200 250 300 250] /Function %s /Extend [false false] >>' % stitch)
    sh['A3'] = pdf.add('<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 200 160 160] /Function %s /Extend [true false] /BBox [20 150 140 215] >>' % lin)
    sh['R1'] = pdf.add('<< /ShadingType 3 /ColorSpace /DeviceRGB /Coords [250 180 0 250 180 45] /Function %s /Extend [false true] >>' % stitch)
    sh['R2'] = pdf.add('<< /ShadingType 3 /ColorSpace /DeviceRGB /Coords [60 90 10 120 70 30] /Function %s /Extend [true true] >>' % lin)
    sh['R3'] = pdf.add('<< /ShadingType 3 /ColorSpace /DeviceGray /Coords [320 160 30 330 165 5] '
                       '/Function << /FunctionType 2 /Domain [0 1] /C0 [0] /C1 [1] /N 1 >> /Extend [true false] >>')
    fr = pdf.stream(b'{ mul }', '/FunctionType 4 /Domain [0 1 0 1] /Range [0 1]', False)
    fg = pdf.stream(b'{ pop }', '/FunctionType 4 /Domain [0 1 0 1] /Range [0 1]', False)
    fb = pdf.stream(b'{ exch pop 1 exch sub }', '/FunctionType 4 /Domain [0 1 0 1] /Range [0 1]', False)
    sh['F1'] = pdf.add('<< /ShadingType 1 /ColorSpace /DeviceRGB /Domain [0 1 0 1] /Matrix [60 0 0 60 160 20] '
                       '/Function [%d 0 R %d 0 R %d 0 R] >>' % (fr, fg, fb))
    tri = b''.join(bytes([0]) + struct.pack('>HH', x, y) + bytes(c) for x, y, c in
                   [(0, 0, (255, 0, 0)), (65535, 0, (0, 255, 0)), (32767, 65535, (0, 0, 255)),
                    (65535, 65535, (255, 255, 0))])
    tri = tri[:3 * 8] + bytes([1]) + tri[3 * 8 + 1:]
    sh['M4'] = pdf.stream(tri, '/ShadingType 4 /ColorSpace /DeviceRGB /BitsPerCoordinate 16 /BitsPerComponent 8 /BitsPerFlag 8 '
                          '/Decode [240 340 20 100 0 1 0 1 0 1]', False)
    lattice = b''.join(struct.pack('>HH', x, y) + bytes([t]) for y, row in enumerate([0, 32767, 65535])
                       for x, t in zip([0, 32767, 65535], [(row // 257 + i * 80) % 256 for i in range(3)]))
    sh['M5'] = pdf.stream(lattice, '/ShadingType 5 /ColorSpace /DeviceRGB /BitsPerCoordinate 16 /BitsPerComponent 8 '
                          '/VerticesPerRow 3 /Decode [20 120 110 150 0 1] /Function %s' % lin, False)
    pts = [(220, 120), (215, 135), (230, 150), (225, 165), (260, 175), (290, 160), (330, 168), (335, 150), (345, 130), (330, 115),
           (290, 105), (255, 112)]

    def q(v, lo, hi):
        return int((v - lo) / (hi - lo) * 65535)
    coons = bytes([0]) + b''.join(struct.pack('>HH', q(x, 0, 360), q(y, 0, 280)) for x, y in pts) + bytes([255, 0, 0, 0, 160, 0, 0, 0, 255, 255, 220, 0])
    coons += bytes([2]) + b''.join(struct.pack('>HH', q(x, 0, 360), q(y, 0, 280)) for x, y in
                                  [(345, 100), (350, 85), (330, 70), (300, 75), (270, 72), (240, 85), (232, 100), (225, 110)])
    coons += bytes([0, 200, 200, 120, 0, 120])
    sh['M6'] = pdf.stream(coons, '/ShadingType 6 /ColorSpace /DeviceRGB /BitsPerCoordinate 16 /BitsPerComponent 8 /BitsPerFlag 8 '
                          '/Decode [0 360 0 280 0 1 0 1 0 1]', False)
    tensor_pts = pts + [(240, 130), (240, 155), (310, 150), (310, 125)]
    tensor = bytes([0]) + b''.join(struct.pack('>HH', q(x - 200, 0, 360), q(y - 100, 0, 280)) for x, y in tensor_pts) + bytes([0, 0, 0, 255, 0, 255, 0, 255, 255, 255, 255, 255])
    sh['M7'] = pdf.stream(tensor, '/ShadingType 7 /ColorSpace /DeviceRGB /BitsPerCoordinate 16 /BitsPerComponent 8 /BitsPerFlag 8 '
                          '/Decode [0 360 0 280 0 1 0 1 0 1]', False)
    c = ['0.96 g 0 0 360 280 re f']
    c.append('q 20 225 140 50 re W n /A1 sh Q')
    c.append('q 180 225 160 50 re W n /A2 sh Q')
    c.append('q 10 140 160 80 re W n /A3 sh Q')
    c.append('q 190 130 120 90 re W n /R1 sh Q')
    c.append('q 10 20 140 110 re W n /R2 sh Q')
    c.append('q 290 125 65 75 re W n /R3 sh Q')
    c.append('q 160 20 60 60 re W n /F1 sh Q')
    c.append('/M4 sh /M5 sh /M6 sh /M7 sh')
    resources = '<< /Shading << %s >> >>' % ' '.join('/%s %d 0 R' % (name, ref) for name, ref in sh.items())
    root, _ = single_page(pdf, 360, 280, '\n'.join(c), resources)
    pdf.write(path, root)


PAGES = [
    ('fonts-type1', page_type1, 'Embedded Type 1 (FontFile): eexec, subrs, flex, hint replacement, div, seac, built-in encoding and Differences.'),
    ('fonts-type1c', page_type1c, 'Embedded CFF (FontFile3 Type1C) with WinAnsi and Differences, and as a CIDFontType0C program.'),
    ('fonts-truetype', page_truetype, 'Embedded TrueType (FontFile2) with WinAnsi encoding.'),
    ('clipping', page_clipping, 'Anti-aliased circular, even-odd, nested and rectangular clips over fills, images and text.'),
    ('text-modes', page_text_modes, 'Text rendering modes: fill, stroke, fill and stroke, hairline outline, rotated outline and text clipping.'),
    ('shadings', page_shadings, 'Shading types 1 to 7 with Extend, BBox, non-nested radial circles, Gouraud, lattice, Coons and tensor meshes.'),
    ('strokes', page_strokes, 'Hairlines, sub-pixel rules, caps, joins, miter limits, dashes, rectangles and skewed pens.'),
]


def font_bytes_table(font, tag):
    data = fonts.font_bytes(font)
    count = struct.unpack('>H', data[4:6])[0]
    for index in range(count):
        record = 12 + index * 16
        if data[record:record + 4] == tag.encode('latin-1'):
            offset, length = struct.unpack('>II', data[record + 8:record + 16])
            return data[offset:offset + length]
    raise KeyError(tag)


fonts.font_bytes_table = font_bytes_table


def poppler_version():
    result = subprocess.run(['pdftoppm', '-v'], capture_output=True, text=True)
    return (result.stderr or result.stdout).splitlines()[0].strip()


def main():
    os.makedirs(OUT, exist_ok=True)
    selected = set(sys.argv[1:])
    manifest_path = os.path.join(OUT, 'manifest.json')
    previous = {}
    if os.path.exists(manifest_path):
        with open(manifest_path) as handle:
            for page in json.load(handle).get('pages', []):
                previous[page['id']] = page
    pages = []
    for page_id, build, description in PAGES:
        pdf_path = os.path.join(OUT, page_id + '.pdf')
        png_prefix = os.path.join(OUT, page_id + '.poppler')
        if not selected or page_id in selected:
            build(pdf_path)
            subprocess.run(['pdftoppm', '-r', str(DPI), '-singlefile', '-png', pdf_path, png_prefix], check=True)
        with open(pdf_path, 'rb') as handle:
            pdf_hash = hashlib.sha256(handle.read()).hexdigest()
        with open(png_prefix + '.png', 'rb') as handle:
            png_hash = hashlib.sha256(handle.read()).hexdigest()
        entry = {
            'id': page_id,
            'description': description,
            'pdf': page_id + '.pdf',
            'pdfSha256': pdf_hash,
            'baseline': page_id + '.poppler.png',
            'baselineSha256': png_hash,
        }
        old = previous.get(page_id, {})
        for key in ('minimumPsnr', 'minimumSsim', 'measuredPsnr', 'measuredSsim'):
            if key in old and old.get('pdfSha256') == pdf_hash:
                entry[key] = old[key]
        pages.append(entry)
    manifest = {
        'version': 1,
        'purpose': 'Full-colour PSNR and SSIM gates for the managed PDF rasterizer against Poppler pdftoppm on a synthetic, embedded-only corpus.',
        'reference': {'rasterizer': poppler_version(), 'dpi': DPI, 'command': 'pdftoppm -r %d -singlefile -png' % DPI},
        'pages': pages,
    }
    with open(manifest_path, 'w') as handle:
        json.dump(manifest, handle, indent=2)
        handle.write('\n')


if __name__ == '__main__':
    main()
