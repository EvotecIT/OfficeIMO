"""Extract unbanded Keynote table-role fills with independent pinned protobuf schemas.

Requires opt-in numbers-parser 4.19.0. Does not qualify Apple exports or complete appearance.
"""
import argparse
import hashlib
import json
import struct
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile
from numbers_parser.iwafile import IWACompressedChunk, get_archive_info_and_remainder
from numbers_parser.generated.TSTArchives_pb2 import TableModelArchive, TableStyleArchive, CellStyleArchive, Tile

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('corpus', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
assert version('numbers-parser') == '4.19.0'
name = 'keynotekit/tabledeck-v15.2.1.key'
source = args.corpus / name
source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
assert source_hash == '384962b1fff18abc5a901b59dc5f8820c2a959977f18f90dc9cd10095bdd0a56'
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
                if position == 0:
                    assert header.identifier not in records
                    records[header.identifier] = (info.type, payload[offset:offset + info.length])
                offset += info.length
            data = payload[offset:]

def chain(identifier, expected_type, schema):
    result = []
    while identifier:
        assert identifier not in [item[0] for item in result]
        kind, payload = records[identifier]
        assert kind == expected_type
        style = schema.FromString(payload)
        result.append((identifier, style))
        identifier = style.super.parent.identifier if style.super.HasField('parent') else 0
    return result

tables = []
for identifier, (kind, payload) in sorted(records.items()):
    if kind != 6001:
        continue
    model = TableModelArchive.FromString(payload)
    # Effective role-only expectations require absence of modern selected cell-style keys.
    for selected_tile in model.base_data_store.tiles.tiles:
        tile_kind, tile_payload = records[selected_tile.tile.identifier]
        assert tile_kind == 6002
        tile = Tile.FromString(tile_payload)
        for row in tile.rowInfos:
            offsets = struct.unpack('<' + 'H' * (len(row.cell_offsets) // 2), row.cell_offsets)
            for offset in offsets:
                if offset == 65535:
                    continue
                if row.has_wide_offsets:
                    offset *= 4
                assert row.cell_storage_buffer[offset] == 5
                assert not struct.unpack_from('<I', row.cell_storage_buffer, offset + 8)[0] & 0x20
    table_chain = chain(model.table_style.identifier, 6003, TableStyleArchive)
    banded = False
    for _, style in reversed(table_chain):
        if style.table_properties.HasField('banded_rows'):
            banded = style.table_properties.banded_rows
    assert not banded
    roles = {}
    for role in ('body_cell_style', 'header_row_style', 'header_column_style', 'footer_row_style'):
        role_chain = chain(getattr(model, role).identifier, 6004, CellStyleArchive)
        fill = None
        for _, style in reversed(role_chain):
            if style.cell_properties.HasField('cell_fill'):
                assert style.cell_properties.cell_fill.SerializeToString() == b''
                fill = {'kind': 'none'}
        assert fill is not None
        roles[role] = {'styleIdentifiers': [ident for ident, _ in role_chain], 'fill': fill}
    assert (model.number_of_rows, model.number_of_columns, model.number_of_header_rows,
            model.number_of_header_columns, model.number_of_footer_rows) == (3, 3, 1, 0, 0)
    tables.append({'modelIdentifier': identifier, 'name': model.table_name,
                   'tableStyleIdentifiers': [ident for ident, _ in table_chain], 'bandedRows': banded,
                   'roles': roles, 'cells': [{'row': row, 'column': column, 'fill': {'kind': 'none'}}
                                          for row in range(1, 4) for column in range(1, 4)]})
assert len(tables) == 1
args.output.write_text(json.dumps({'provider': 'numbers-parser', 'providerVersion': '4.19.0',
    'source': {'path': name, 'sha256': source_hash}, 'tables': tables,
    'limits': 'Unbanded native Keynote role no-fill declarations only. Role intersections, banding, Apple exports and complete appearance are not qualified.'}, indent=2) + '\n', encoding='utf-8')
