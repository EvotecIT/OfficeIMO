"""Reproduce selected list styles and paragraph data from pinned native Pages fixtures.

Uses opt-in numbers-parser 4.19.0 schemas; it does not rewrite the packages or qualify Apple exports.
"""
import argparse
import hashlib
import json
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile
from numbers_parser.iwafile import IWACompressedChunk, get_archive_info_and_remainder
from numbers_parser.generated.TSWPArchives_pb2 import StorageArchive, ListStyleArchive

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('corpus', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
assert version('numbers-parser') == '4.19.0'
fixtures = {
    'picodocs/sample-v14.4.pages': '4714477138d0a4090fc2ee2ba2ebb6adcd0fb6ce20a28897a6247a8e17d1ddce',
}
sources = []
for name, expected_hash in fixtures.items():
    source = args.corpus / name
    source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    assert source_hash == expected_hash
    records = {}
    with ZipFile(source) as package:
        for entry in package.namelist():
            if not entry.endswith('.iwa'):
                continue
            data = b''.join(IWACompressedChunk._decompress_all(package.read(entry)))
            while data:
                header, payload = get_archive_info_and_remainder(data)
                offset = 0
                for payload_index, info in enumerate(header.message_infos):
                    content = payload[offset:offset + info.length]
                    offset += info.length
                    if payload_index == 0:
                        assert header.identifier not in records
                        records[header.identifier] = (info.type, content)
                data = payload[offset:]
    storages = []
    selected_styles = {}
    for storage_id, (kind, content) in records.items():
        if kind != 2001:
            continue
        storage = StorageArchive.FromString(content)
        if 'First ordered item' not in ''.join(storage.text):
            continue
        selections = []
        for entry in storage.table_list_style.entries:
            if not entry.HasField('object'):
                selections.append({'characterOffset': entry.character_index})
                continue
            identifier = entry.object.identifier
            selections.append({'characterOffset': entry.character_index,
                               'styleIdentifier': identifier})
            pending = [identifier]
            while pending:
                current = pending.pop()
                if current in selected_styles:
                    continue
                style_kind, payload = records[current]
                assert style_kind == 2023
                style = ListStyleArchive.FromString(payload)
                parent = style.super.parent.identifier if style.super.HasField('parent') else None
                selected_styles[current] = {
                    'identifier': current, 'parentIdentifier': parent,
                    'labelTypes': list(style.label_types),
                    'numberTypes': list(style.number_types),
                    'strings': list(style.strings), 'indents': list(style.indents),
                    'tieredNumbers': list(style.tiered_numbers),
                    'fontName': style.font_name if style.HasField('font_name') else None,
                    'fontNameCleared': style.font_name_null if style.HasField('font_name_null') else None}
                if parent is not None:
                    pending.append(parent)
        storages.append({'identifier': storage_id, 'listSelections': selections,
                         'paragraphData': [{'characterOffset': e.character_index,
                                            'first': e.first, 'second': e.second}
                                           for e in storage.table_para_data.entries]})
    sources.append({'path': name, 'sha256': source_hash, 'storages': storages,
                    'selectedStyles': sorted(selected_styles.values(), key=lambda s: s['identifier'])})
assert len(sources[0]['storages']) == 1
assert len(sources[0]['selectedStyles']) == 3
manifest = {'provider': 'numbers-parser', 'providerVersion': '4.19.0',
            'offsetUnit': 'UTF-16 code units before marker removal',
            'numberTypeNames': {str(v.number): v.name for v in
                                ListStyleArchive.DESCRIPTOR.enum_types_by_name['NumberType'].values},
            'sources': sources,
            'listLevelEvidence': {'repository': 'https://github.com/LibreOffice/libetonyek',
                                  'revision': '6d19ce365f64a6217791a9438c85f4b1851ca7e0',
                                  'source': 'src/lib/IWAParser.cpp',
                                  'contract': 'Storage field 6 entries select list levels from entry field 2 at field 1 character offsets.'},
            'limits': 'Selected native declarations only. Paragraph-data first is the explicit list level; second semantics, restarts, continuation, tiered numbering and destination rendering are not qualified by this extraction.'}
args.output.write_text(json.dumps(manifest, indent=2) + '\n', encoding='utf-8')
