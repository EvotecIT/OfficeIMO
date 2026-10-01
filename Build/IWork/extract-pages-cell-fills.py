"""Extract selected Pages cell fills and layout through pinned independent protobuf schemas.

Requires opt-in numbers-parser 4.19.0. No native Apple export/appearance claim.
"""
import argparse
import hashlib
import json
import math
import struct
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile
from numbers_parser.iwafile import IWACompressedChunk, get_archive_info_and_remainder
from numbers_parser.generated.TSTArchives_pb2 import TableModelArchive, CellStyleArchive, Tile, TableDataList

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('corpus', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
assert version('numbers-parser') == '4.19.0'
name = 'picodocs/sample-v14.4.pages'
source = args.corpus / name
source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
assert source_hash == '4714477138d0a4090fc2ee2ba2ebb6adcd0fb6ce20a28897a6247a8e17d1ddce'
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

def read_fill(identifier):
    chain = []
    current = identifier
    while current:
        assert current not in [identifier for identifier, _ in chain]
        kind, payload = records[current]
        assert kind == 6004
        style = CellStyleArchive.FromString(payload)
        chain.append((current, style))
        current = style.super.parent.identifier if style.super.HasField('parent') else 0
    result = None
    padding = None
    vertical = None
    for _, style in reversed(chain):
        properties = style.cell_properties
        if properties.HasField('vertical_alignment'):
            assert properties.vertical_alignment in (0, 1, 2)
            vertical = ['top', 'middle', 'bottom'][properties.vertical_alignment]
        if properties.HasField('padding'):
            # A declared message replaces the whole property; omitted scalar sides default to zero.
            padding = {side: getattr(properties.padding, side) for side in ('left', 'top', 'right', 'bottom')}
            assert all(math.isfinite(value) and value >= 0 for value in padding.values())
        if not style.cell_properties.HasField('cell_fill'):
            continue
        fill = style.cell_properties.cell_fill
        assert not fill.HasField('gradient') and not fill.HasField('image')
        if not fill.HasField('color'):
            assert len(fill.SerializeToString()) == 0
            result = {'kind': 'none'}
        else:
            color = fill.color
            assert color.model == 1 and color.rgbspace == 1 and color.a == 1
            components = [color.r, color.g, color.b]
            assert all(math.isfinite(value) and 0 <= value <= 1 for value in components)
            result = {'kind': 'solid', 'rgbHex': ''.join(f'{math.floor(value * 255 + 0.5):02X}' for value in components)}
    assert result is not None
    return result, [identifier for identifier, _ in chain], padding, vertical

tables = []
for identifier, (kind, content) in sorted(records.items()):
    if kind != 6001:
        continue
    model = TableModelArchive.FromString(content)
    catalog_kind, payload = records[model.base_data_store.styleTable.identifier]
    assert catalog_kind == 6005
    catalog = TableDataList.FromString(payload)
    assert catalog.listType == 4
    entries = {entry.key: entry.reference.identifier for entry in catalog.entries}
    assert len(entries) == len(catalog.entries)
    cells = []
    for selected_tile in model.base_data_store.tiles.tiles:
        tile_kind, payload = records[selected_tile.tile.identifier]
        assert tile_kind == 6002
        tile = Tile.FromString(payload)
        for row in tile.rowInfos:
            buffer = row.cell_storage_buffer
            offsets = struct.unpack('<' + 'H' * (len(row.cell_offsets) // 2), row.cell_offsets)
            for column, offset in enumerate(offsets):
                if offset == 65535:
                    continue
                if row.has_wide_offsets:
                    offset *= 4
                assert buffer[offset] == 5
                flags = struct.unpack_from('<I', buffer, offset + 8)[0]
                if not flags & 0x20:
                    continue
                style_offset = offset + 12 + sum(
                    16 if bit == 0 else 8 if bit in (1, 2) else 4
                    for bit in range(5) if flags & (1 << bit))
                key = struct.unpack_from('<I', buffer, style_offset)[0]
                fill, chain, padding, vertical = read_fill(entries[key])
                cells.append({'row': selected_tile.tileid * 256 + row.tile_row_index + 1,
                              'column': column + 1, 'empty': buffer[offset + 1] == 0,
                              'styleKey': key, 'styleIdentifiers': chain, 'fill': fill,
                              'paddingPoints': padding, 'verticalAlignment': vertical})
    tables.append({'modelIdentifier': identifier, 'name': model.table_name, 'cells': cells})
assert len(tables) == 3 and sum(len(table['cells']) for table in tables) == 64
manifest = {'provider': 'numbers-parser', 'providerVersion': '4.19.0',
            'source': {'path': name, 'sha256': source_hash}, 'tables': tables,
            'limits': 'Selected modern cell fills, four-sided padding, vertical alignment and inheritance only; table-role defaults, banding, other styles, Apple exports and complete appearance are not qualified.'}
args.output.write_text(json.dumps(manifest, indent=2) + '\n', encoding='utf-8')
