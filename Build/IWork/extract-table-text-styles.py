"""Extract table text defaults and selected keys with opt-in numbers-parser 4.19.0 schemas.

The immutable native fixtures qualify fonts, emphasis and alignment; Apple exports and full appearance remain separate.
"""
import argparse
import hashlib
import json
import struct
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile
from numbers_parser.iwafile import IWACompressedChunk, get_archive_info_and_remainder
from numbers_parser.generated.TSTArchives_pb2 import TableModelArchive, TableDataList, Tile
from numbers_parser.generated.TSWPArchives_pb2 import ParagraphStyleArchive

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('corpus', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
assert version('numbers-parser') == '4.19.0'
fixtures = ['keynotekit/tabledeck-v15.2.1.key', 'picodocs/sample-v14.4.pages',
            'numbers-parser/test-10-formulas.numbers']
results = []
for name in fixtures:
    source = args.corpus / name
    records = {}
    with ZipFile(source) as package:
        for entry in package.namelist():
            if not entry.endswith('.iwa'):
                continue
            data = b''.join(IWACompressedChunk._decompress_all(package.read(entry)))
            while data:
                header, payload = get_archive_info_and_remainder(data)
                offset = 0
                for position, info in enumerate(header.message_infos):
                    content = payload[offset:offset + info.length]
                    offset += info.length
                    if position == 0:
                        assert header.identifier not in records
                        records[header.identifier] = (info.type, content)
                data = payload[offset:]

    def paragraph_style(identifier):
        chain = []
        current = identifier
        while current:
            assert current not in [key for key, _ in chain]
            kind, payload = records[current]
            assert kind == 2022
            style = ParagraphStyleArchive.FromString(payload)
            chain.append((current, style))
            current = style.super.parent.identifier if style.super.HasField('parent') else 0
        properties = {}
        for _, style in reversed(chain):
            char = style.char_properties
            for field in ('bold', 'italic', 'font_size', 'font_name', 'underline', 'strikethru'):
                if char.HasField(field):
                    properties[field] = getattr(char, field)
            if char.HasField('font_name_null') and char.font_name_null:
                properties.pop('font_name', None)
            if style.para_properties.HasField('alignment'):
                properties['alignment'] = style.para_properties.alignment
        return {'identifiers': [key for key, _ in chain], 'properties': properties}

    tables = []
    for identifier, (kind, payload) in sorted(records.items()):
        if kind != 6001:
            continue
        model = TableModelArchive.FromString(payload)
        roles = {}
        for field in ('body_text_style', 'header_row_text_style', 'header_column_text_style', 'footer_row_text_style'):
            if model.HasField(field):
                roles[field] = paragraph_style(getattr(model, field).identifier)
        selected = []
        entries = {}
        if model.base_data_store.HasField('styleTable'):
            catalog_kind, catalog_payload = records[model.base_data_store.styleTable.identifier]
            assert catalog_kind == 6005
            catalog = TableDataList.FromString(catalog_payload)
            assert catalog.listType == 4
            entries = {entry.key: entry.reference.identifier for entry in catalog.entries}
            assert len(entries) == len(catalog.entries)
        for tile_entry in model.base_data_store.tiles.tiles:
            tile_kind, tile_payload = records[tile_entry.tile.identifier]
            assert tile_kind == 6002
            tile = Tile.FromString(tile_payload)
            for row in tile.rowInfos:
                offsets = struct.unpack('<' + 'H' * (len(row.cell_offsets) // 2), row.cell_offsets)
                buffer = row.cell_storage_buffer
                for column, offset in enumerate(offsets):
                    if offset == 65535:
                        continue
                    if row.has_wide_offsets:
                        offset *= 4
                    assert buffer[offset] == 5
                    flags = struct.unpack_from('<I', buffer, offset + 8)[0]
                    if not flags & (1 << 6):
                        continue
                    selected_offset = offset + 12 + sum(16 if bit == 0 else 8 if bit in (1, 2) else 4
                        for bit in range(6) if flags & (1 << bit))
                    key = struct.unpack_from('<I', buffer, selected_offset)[0]
                    selected.append({'row': tile_entry.tileid * 256 + row.tile_row_index + 1,
                                     'column': column + 1, 'styleKey': key, **paragraph_style(entries[key])})
        tables.append({'modelIdentifier': identifier, 'name': model.table_name,
                       'rows': model.number_of_rows, 'columns': model.number_of_columns,
                       'headerRows': model.number_of_header_rows, 'headerColumns': model.number_of_header_columns,
                       'footerRows': model.number_of_footer_rows, 'roles': roles, 'selected': selected})
    results.append({'path': name, 'sha256': hashlib.sha256(source.read_bytes()).hexdigest(), 'tables': tables})
args.output.write_text(json.dumps({'provider': 'numbers-parser', 'providerVersion': '4.19.0',
    'sources': results, 'limits': 'Fonts, emphasis, horizontal alignment, role references and selected text-style keys only; full typography, Apple exports and appearance are not qualified.'}, indent=2) + '\n', encoding='utf-8')
