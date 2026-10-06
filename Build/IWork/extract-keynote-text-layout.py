"""Extract selected text-frame/list facts from the pinned Keynote 14.5 fixture.

Requires the isolated numbers-parser 4.19.0 evidence environment. Optional
wrapped probes change only selected body text and keep-lines; they are derived
test inputs, not independently authored documents or an OfficeIMO writer.
"""
import argparse
import hashlib
import json
import struct
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile

import snappy
from google.protobuf.internal.encoder import _VarintBytes
from google.protobuf.json_format import MessageToDict
from numbers_parser.generated import TSWPArchives_pb2 as TSWP
from numbers_parser.iwafile import IWACompressedChunk, get_archive_info_and_remainder
from iwork_protobuf_evidence import fields

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('corpus', type=Path)
parser.add_argument('output', type=Path)
parser.add_argument('--wrapped-dir', type=Path)
args = parser.parse_args()
assert version('numbers-parser') == '4.19.0'
source = args.corpus / 'nim-iwork/simple.key'
expected = 'ba95755df82ceb0ca834e1e03e2777c34fad906320d8336b4f3fefc6b48607eb'
assert hashlib.sha256(source.read_bytes()).hexdigest() == expected
records = {}
with ZipFile(source) as package:
    entries = [(entry, package.read(entry)) for entry in package.infolist()]
for entry, content in entries:
    if not entry.filename.endswith('.iwa'):
        continue
    data = b''.join(IWACompressedChunk._decompress_all(content))
    while data:
        header, payload = get_archive_info_and_remainder(data)
        length = sum(info.length for info in header.message_infos)
        assert header.identifier not in records
        records[header.identifier] = (header.message_infos[0].type,
                                      payload[:header.message_infos[0].length])
        data = payload[length:]

frames = []
for slide_id in (1221831, 1225067):
    kind, content = records[slide_id]
    assert kind == 5
    slide = fields(content)
    for role in (5, 6):
        identifier = fields(slide[role][0])[1][0]
        kind, content = records[identifier]
        assert kind == 7
        info = TSWP.ShapeInfoArchive.FromString(fields(content)[1][0])
        chain = []
        style_id = info.super.style.identifier
        while style_id:
            assert style_id not in [item['identifier'] for item in chain]
            kind, content = records[style_id]
            assert kind == 2025
            style = TSWP.ShapeStyleArchive.FromString(content)
            chain.append({'identifier': style_id,
                          'properties': MessageToDict(style.shape_properties)})
            style_id = style.super.super.parent.identifier
        frames.append({'slide': slide_id, 'role': role, 'placeholder': identifier,
                       'geometry': MessageToDict(info.super.super.geometry),
                       'styleChain': chain})
kind, content = records[1220734]
assert kind == 2023
list_style = MessageToDict(TSWP.ListStyleArchive.FromString(content))
manifest = {
    'source': 'nim-iwork/simple.key', 'sourceSha256': expected,
    'provider': 'numbers-parser', 'providerVersion': '4.19.0',
    'placeholderSchema': {
        'revision': 'be4828260466c9c022ec12f9dc9dbfeb15ab1dea',
        'url': 'https://github.com/orcastor/iwork-converter/blob/be4828260466c9c022ec12f9dc9dbfeb15ab1dea/proto/KN/KNArchives.pb.go',
        'fieldPath': 'PlaceholderArchive.1 -> ShapeInfoArchive.1 -> ShapeArchive.1 -> DrawableArchive.1'},
    'selectedFrames': frames,
    'selectedList': {'identifier': 1220734, 'properties': list_style},
    'limits': 'Pinned selected frames and character-list declarations only. Apple export/render evidence is recorded separately.'}

if args.wrapped_dir:
    args.wrapped_dir.mkdir(parents=True, exist_ok=True)
    probes = []
    text = 'first bullet wraps across several lines and remains inside this fixed text frame'
    for keep in (True, False):
        path = args.wrapped_dir / ('keynote-fixed-frame-wrapped-' + ('keep' if keep else 'no-keep') + '.key')
        changed = []
        with ZipFile(path, 'w') as output:
            for entry, content in entries:
                if entry.filename.endswith('.iwa'):
                    data = b''.join(IWACompressedChunk._decompress_all(content))
                    segments = []
                    while data:
                        header, payload = get_archive_info_and_remainder(data)
                        length = sum(info.length for info in header.message_infos)
                        offset, messages = 0, []
                        for info in header.message_infos:
                            message = payload[offset:offset + info.length]
                            offset += info.length
                            if header.identifier == 1221867 and info.type == 2001:
                                storage = TSWP.StorageArchive.FromString(message)
                                assert list(storage.text) == ['first bullet']
                                assert all(item.character_index == 0 for field, value in storage.ListFields()
                                           if field.name.startswith('table_') for item in value.entries)
                                storage.text[0] = text
                                message = storage.SerializeToString()
                                changed.append(header.identifier)
                            elif header.identifier == 1220740 and info.type == 2022:
                                paragraph = TSWP.ParagraphStyleArchive.FromString(message)
                                paragraph.para_properties.keep_lines_together = keep
                                message = paragraph.SerializeToString()
                                changed.append(header.identifier)
                            info.length = len(message)
                            messages.append(message)
                        encoded_header = header.SerializeToString()
                        segments.append(_VarintBytes(len(encoded_header)) + encoded_header + b''.join(messages))
                        data = payload[length:]
                    raw = b''.join(segments)
                    content = b''.join(b'\x00' + struct.pack('<I', len(compressed))[:3] + compressed
                                       for compressed in (snappy.compress(raw[i:i + 65536])
                                                          for i in range(0, len(raw), 65536)))
                output.writestr(entry, content)
        assert sorted(changed) == [1220740, 1221867]
        probes.append({'path': path.name, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest(),
                       'keepLinesTogether': keep, 'changedRecordIdentifiers': sorted(changed)})
    manifest['wrappedProbes'] = {'text': text, 'inputs': probes,
                               'provenance': 'Test-only mutations of the pinned licensed native fixture; no independent-producer claim.'}
args.output.write_text(json.dumps(manifest, indent=2) + '\n', encoding='utf-8')
print('Extracted', len(frames), 'selected frames and one character-list style')
