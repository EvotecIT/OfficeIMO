"""Extract no-effect transition evidence from pinned native Keynote packages.

Uses opt-in numbers-parser 4.19.0 and the independent KN schema cited in the output.
Includes template records for inventory only; it does not qualify active effects.
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
    'keynotekit/imagedeck-v15.2.1.key': 'a9af589197588e04ee52388b0aa6c2dad1110e5d6db814b58afe543831cf2128'}
sources = []
for name, expected in fixtures.items():
    source = args.corpus / name
    assert hashlib.sha256(source.read_bytes()).hexdigest() == expected
    slides = []
    with ZipFile(source) as package:
        for entry in package.namelist():
            if not entry.endswith('.iwa'):
                continue
            data = b''.join(IWACompressedChunk._decompress_all(package.read(entry)))
            while data:
                header, payload = get_archive_info_and_remainder(data)
                position = 0
                for info in header.message_infos:
                    content = payload[position:position + info.length]
                    position += info.length
                    if info.type != 5:
                        continue
                    slide = fields(content)
                    transition = fields(slide[4][0])
                    attributes = fields(transition[2][0])
                    animation = fields(attributes[8][0])
                    assert animation[1] == [b'Transition'] and animation[2] == [b'none']
                    assert animation[6] == [0]
                    slides.append({'identifier': header.identifier, 'entry': entry,
                                   'effect': animation[2][0].decode(), 'automatic': bool(animation[6][0]),
                                   'duration': struct.unpack('<d', animation[3][0])[0],
                                   'delay': struct.unpack('<d', animation[5][0])[0]})
                data = payload[position:]
    sources.append({'path': name, 'sha256': expected, 'slidesIncludingTemplates': slides})
manifest = {'provider': 'numbers-parser', 'providerVersion': '4.19.0',
            'schema': {'repository': 'https://github.com/orcastor/iwork-converter',
                       'revision': 'be4828260466c9c022ec12f9dc9dbfeb15ab1dea',
                       'path': 'proto/KN/KNArchives.pb.go',
                       'fieldPath': 'SlideArchive.4 / TransitionArchive.2 / TransitionAttributesArchive.8 / AnimationAttributesArchive'},
            'sources': sources,
            'limits': 'All observed effects are none with automatic advance disabled. Includes inactive templates. Nonzero dormant delay is not automatic advance. Active effects, timing, Apple exports and rendered transitions remain unqualified.'}
args.output.write_text(json.dumps(manifest, indent=2) + '\n', encoding='utf-8')
