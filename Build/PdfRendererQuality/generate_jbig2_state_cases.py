#!/usr/bin/env python3
"""Rebuild bounded JBIG2 state fixtures; expected packed pixels come from Apache JBIG2 ImageIO.

Development only: Python 3, Java 11+ and an explicitly supplied JBIG2 ImageIO 3.0.5 JAR.
Temporary streams and decoded PBMs stay in --scratch; no product runtime dependency.
T.88 7.4.2.2 (retained generic/refinement MQ state), 7.4.8.5 (page operators).
"""
import argparse
import base64
import copy
import hashlib
import json
from pathlib import Path
import struct
import subprocess
import jbig2 as j


def refinement(encoder, contexts, bitmap, reference):
    # T.88 Figure 12, refinement template 1, reference displacement zero.
    coding = [(-1, -1), (0, -1), (1, -1), (-1, 0)]
    ref = [(0, -1), (-1, 0), (0, 0), (1, 0), (0, 1), (1, 1)]
    def pixel(image, x, y):
        return image[y][x] if 0 <= y < len(image) and 0 <= x < len(image[0]) else 0
    for y, row in enumerate(bitmap):
        for x, bit in enumerate(row):
            context = 0
            for dx, dy in coding:
                context = (context << 1) | pixel(bitmap, x + dx, y + dy)
            for dx, dy in ref:
                context = (context << 1) | pixel(reference, x + dx, y + dy)
            encoder.encode(contexts, context, bit)


def dictionary(bitmap, imported=(), state=None, retain=True, refine=False):
    used = state is not None
    generic, refinement_state = copy.deepcopy(state) if used else ({}, {})
    encoder, integers, ids = j.MQEncoder(), {}, {}
    def integer(name, value):
        encoder.encode_integer(integers.setdefault(name, {}), value)
    integer('IADH', len(bitmap))
    integer('IADW', len(bitmap[0]))
    if refine:
        integer('IAAI', 1)
        code_length = len(imported).bit_length()  # input count + one new symbol
        encoder.encode_id(ids, 0, code_length)
        integer('IARDX', 0)
        integer('IARDY', 0)
        refinement(encoder, refinement_state, bitmap, imported[0])
    else:
        j.encode_generic(encoder, generic, bitmap, 0, False, j.DEFAULT_AT[0])
    integer('IADW', None)
    integer('IAEX', len(imported))  # only export the new symbol
    integer('IAEX', 1)
    flags = (0x100 if used else 0) | (0x200 if retain else 0) | (0x1002 if refine else 0)
    body = struct.pack('>H', flags) + b''.join(struct.pack('bb', *at) for at in j.DEFAULT_AT[0])
    body += struct.pack('>II', 1, 1) + encoder.flush()
    return body, (generic, refinement_state)


def page(width, height, default=0, override=False, initial=0):
    flags = (default << 3) | (0x40 if override else 0) | (initial << 2)
    return j.segment(1, 48, struct.pack('>IIIIBH', width, height, 0, 0, flags, 0))


def cases():
    a = [[int((x * 3 + y * 5) % 7 < 3) for x in range(8)] for y in range(8)]
    b = [[int((x + y * 2) % 5 < 2) for x in range(8)] for y in range(8)]
    for kind in ('generic', 'text'):
        for operator in range(4):
            for allowed in (False, True):
                data = page(8, 8, operator, allowed, 1)
                if kind == 'generic':
                    data += j.segment(2, 38, j.generic_region(a, 0, 0, operator=0 if allowed else operator))
                    data += j.segment(3, 38, j.generic_region(b, 0, 0, operator=0 if allowed else operator))
                else:
                    body, _ = j.symbol_dictionary([a, b])
                    data += j.segment(2, 0, body)
                    for number, symbol in ((3, 0), (4, 1)):
                        data += j.segment(number, 6, j.text_region(8, 8, 0, 0, [(symbol, 0, 0, 8)], 2, operator=0 if allowed else operator), referred=[2])
                yield '%s-op%d-override%d' % (kind, operator, allowed), data
    # Refinement against the page retains REPLACE even when direct-region overrides are disabled.
    e = j.MQEncoder()
    refinement(e, {}, b, [[1] * 8 for _ in range(8)])
    info = bytearray(page(8, 8, default=1, initial=1))
    info[27] |= 2  # page may contain refinement regions
    yield 'page-refinement-replace', bytes(info) + j.segment(2, 42, j.region_info(8, 8, 0, 0, 4) + bytes([1]) + e.flush())
    for refine in (False, True):
        data = page(24, 8)
        if refine:
            seed, _ = dictionary(a, retain=False)
            data += j.segment(2, 0, seed)
            parent, state = dictionary(b, [a], refine=True)
            data += j.segment(3, 0, parent, referred=[2])
            reference, parent_number, start = b, 3, 4
        else:
            parent, state = dictionary(a)
            data += j.segment(2, 0, parent)
            reference, parent_number, start = a, 2, 3
        # Both children reuse the same parent. Reusing/mutating its arrays in place corrupts child two.
        for number, bitmap in ((start, b), (start + 1, a)):
            body, _ = dictionary(bitmap, [reference], state, retain=False, refine=refine)
            data += j.segment(number, 0, body, referred=[parent_number])
        for index, symbol_dictionary in enumerate((parent_number, start, start + 1)):
            data += j.segment(start + 2 + index, 6, j.text_region(8, 8, index * 8, 0, [(0, 0, 0, 8)], 1), referred=[symbol_dictionary])
        yield 'retained-' + ('refinement' if refine else 'generic'), data


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--scratch', type=Path, required=True)
    parser.add_argument('--jbig2-imageio', type=Path, required=True, help='Test-only Apache JBIG2 ImageIO 3.0.5 JAR')
    parser.add_argument('--output', type=Path, default=Path(__file__).resolve().parents[2] / 'OfficeIMO.Pdf.Tests/Pdf/Fixtures/Rendering/jbig2-state.json')
    args = parser.parse_args()
    args.scratch.mkdir(parents=True, exist_ok=True)
    jar_hash = hashlib.sha256(args.jbig2_imageio.read_bytes()).hexdigest()
    if jar_hash != 'd426cc8b6ac211dd5ffd120a86a65e55ae65d808c818207c1261bc1880b42653':
        raise ValueError('Expected the Maven Central Apache JBIG2 ImageIO 3.0.5 JAR')
    result = []
    for name, encoded in cases():
        width, height = (24 if name.startswith('retained') else 8), 8
        prefix = args.scratch / name
        raw = args.scratch / (name + '.jb2')
        raw.write_bytes(encoded)
        run = subprocess.run(['java', '-cp', str(args.jbig2_imageio), str(Path(__file__).with_name('DecodeJbig2.java')), str(raw), str(prefix) + '-000.pbm'], capture_output=True, check=True)
        if run.stderr:
            (args.scratch / (name + '.stderr')).write_bytes(run.stderr)
        pbm = Path(str(prefix) + '-000.pbm').read_bytes()
        header, dims, pixels = pbm.split(b'\n', 2)
        assert header == b'P4' and dims == ('%d %d' % (width, height)).encode()
        decoded = bytes(value ^ 255 for value in pixels)  # PDF JBIG2Decode: zero is black.
        result.append(dict(name=name, width=width, height=height, encoded=base64.b64encode(encoded).decode(), decoded=base64.b64encode(decoded).decode(), encodedSha256=hashlib.sha256(encoded).hexdigest()))
    version = 'Apache JBIG2 ImageIO 3.0.5'
    args.output.write_text(json.dumps(dict(oracle=version, oracleJarSha256=jar_hash, specification='ITU-T T.88 7.4.2.2 and 7.4.8.5', cases=result), indent=2) + '\n')
    print('Generated %d independently decoded cases' % len(result))


if __name__ == '__main__':
    main()
