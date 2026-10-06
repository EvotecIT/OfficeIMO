"""Inventory slide background fills in pinned native Keynote packages.

Requires numbers-parser 4.19.0 for IWA framing. Field identities use the pinned
independent KN schema. Templates are included, but selection belongs to the reader.
"""
import argparse
import hashlib
import json
import struct
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile
from numbers_parser.iwafile import IWACompressedChunk, get_archive_info_and_remainder
from iwork_protobuf_evidence import fields

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('corpus', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
assert version('numbers-parser') == '4.19.0'
fixtures = {
    'iwork-converter/a.key': '929347827a7478c123dd3e3828e9751b5cf2ae977d2edd2a5f0774014fc4fefc',
    'nim-iwork/simple.key': 'ba95755df82ceb0ca834e1e03e2777c34fad906320d8336b4f3fefc6b48607eb',
    'keynotekit/tabledeck-v15.2.1.key': '384962b1fff18abc5a901b59dc5f8820c2a959977f18f90dc9cd10095bdd0a56',
    'keynotekit/imagedeck-v15.2.1.key': 'a9af589197588e04ee52388b0aa6c2dad1110e5d6db814b58afe543831cf2128',
    'native-exports/keynote-colors-v15.4.key': 'd9c7c5d0b1bb2bff80e44ea683239601dcc5c2e632205ad8611bb0110f3502f1'}
sources = []
for name, expected in fixtures.items():
    source = args.corpus / name
    assert hashlib.sha256(source.read_bytes()).hexdigest() == expected
    records = {}
    with ZipFile(source) as package:
        for entry in package.namelist():
            if not entry.endswith('.iwa'):
                continue
            data = b''.join(IWACompressedChunk._decompress_all(package.read(entry)))
            while data:
                header, payload = get_archive_info_and_remainder(data)
                position = 0
                for payload_index, info in enumerate(header.message_infos):
                    content = payload[position:position + info.length]
                    position += info.length
                    if payload_index == 0 and info.type in (5, 9):
                        assert header.identifier not in records
                        records[header.identifier] = (info.type, entry, fields(content))
                data = payload[position:]
    slides = []
    for identifier, (kind, entry, slide) in records.items():
        if kind != 5:
            continue
        style_id = fields(slide[1][0])[1][0]
        chain, seen = [], set()
        while style_id is not None:
            assert style_id not in seen
            seen.add(style_id)
            style_kind, style_entry, style = records[style_id]
            assert style_kind == 9
            chain.append((style_id, style))
            super_style = fields(style[1][0])
            style_id = fields(super_style[3][0])[1][0] if 3 in super_style else None
        result = {'identifier': identifier, 'entry': entry, 'styleChain': [item[0] for item in chain],
                  'backgroundKind': 'unspecified', 'rgb': None}
        unsupported = False
        for style_id, style in reversed(chain):
            properties = fields(style[11][0]) if 11 in style else {}
            if 1 not in properties:
                continue
            fill = fields(properties[1][0])
            if not fill:
                result.update(backgroundKind='none', rgb=None)
                continue
            assert set(fill) == {1} and len(fill[1]) == 1
            color = fields(fill[1][0])
            result['colorFields'] = sorted(color)
            assert color[1] == [1] and color[12] in ([1], [2])
            result['colorSpace'] = color[12][0]
            rgba = [struct.unpack('<f', color[field][0])[0] for field in (3, 4, 5, 6)]
            assert all(0 <= channel <= 1 for channel in rgba)
            result['rgba'] = rgba
            extra = color.get(13)
            neutral_extra = extra is None or (len(extra) == 1 and isinstance(extra[0], bytes)
                                              and len(extra[0]) == 4 and struct.unpack('<f', extra[0])[0] == 1)
            supported_fields = {1, 3, 4, 5, 6, 12} | ({13} if extra is not None else set())
            unsupported |= color[12] != [1] or set(color) != supported_fields or not neutral_extra or rgba[3] != 1
            result.update(backgroundKind='solid', rgb=''.join(f'{int(channel * 255 + .5):02X}' for channel in rgba[:3]))
        if unsupported:
            result.update(backgroundKind='unsupported', rgb=None)
        slides.append(result)
    sources.append({'path': name, 'sha256': expected, 'slidesIncludingTemplates': slides})
manifest = {'provider': 'numbers-parser', 'providerVersion': '4.19.0',
            'schema': {'repository': 'https://github.com/orcastor/iwork-converter',
                       'revision': 'be4828260466c9c022ec12f9dc9dbfeb15ab1dea',
                       'path': 'proto/KN/KNArchives.pb.go',
                       'fieldPath': 'SlideArchive.1 -> SlideStyleArchive.11 / SlideStylePropertiesArchive.1 / FillArchive.1'},
            'sources': sources,
            'limits': 'Opaque RGB/sRGB fills, including the fixed32 value 1 in field 13 observed in Keynote 15.4 controlled blue/red backgrounds and inherited white. The field meaning is unspecified; other values, Display P3 and other extra fields remain unqualified. Includes inactive templates. Native rendering qualification is limited to native-exports/keynote-colors-v15.4.json; this inventory does not qualify master layouts.'}
args.output.write_text(json.dumps(manifest, indent=2) + '\n', encoding='utf-8')
print('Extracted', sum(len(source['slidesIncludingTemplates']) for source in sources), 'slide background declarations')
