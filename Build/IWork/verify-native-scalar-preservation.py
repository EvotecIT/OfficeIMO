"""Opt-in Numbers scalar-edit evidence using independent schemas and numbers-parser.

Compares every file entry, raw IWA header, sibling payload and non-value byte.
Coordinates are one-based table coordinates; no OfficeIMO assembly is loaded.
"""
import argparse
import hashlib
import importlib.util
import json
import struct
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile

from google.protobuf import descriptor_pb2, descriptor_pool, message_factory
from google.protobuf.internal.decoder import _DecodeVarint
from numbers_parser.cell import _unpack_decimal128


spec = importlib.util.spec_from_file_location('preservation', Path(__file__).with_name('verify-native-preservation.py'))
preservation = importlib.util.module_from_spec(spec)
spec.loader.exec_module(preservation)


def fields(data, start=0, end=None):
    end = len(data) if end is None else end
    while start < end:
        key, start = _DecodeVarint(data, start)
        field, wire, value_start = key >> 3, key & 7, start
        assert field > 0
        if wire == 0:
            _, start = _DecodeVarint(data, start)
        elif wire == 2:
            length, value_start = _DecodeVarint(data, start)
            start = value_start + length
        elif wire in (1, 5):
            start += 8 if wire == 1 else 4
        else:
            raise ValueError(f'Unsupported wire kind {wire}')
        assert start <= end
        yield field, wire, value_start, start


def locate(records, pool, table_id, row, column, kind):
    def decode(identifier, expected_type, name):
        archive_type, data = records[identifier]
        assert archive_type == expected_type
        message = message_factory.GetMessageClass(pool.FindMessageTypeByName(name)).FromString(data)
        assert not message.FindInitializationErrors()
        return message, data

    table, _ = decode(table_id, 6000, 'TST.TableInfoArchive')
    model, _ = decode(table.tableModel.identifier, 6001, 'TST.TableModelArchive')
    assert 1 <= row <= model.number_of_rows and 1 <= column <= model.number_of_columns
    selected = [item for item in model.base_data_store.tiles.tiles if item.tileid == (row - 1) // 256]
    assert len(selected) == 1
    tile_id = selected[0].tile.identifier
    tile, data = decode(tile_id, 6002, 'TST.Tile')
    assert tile.storage_version == 5 and tile.last_saved_in_BNC
    rows = [item for item in fields(data) if item[0:2] == (5, 2)]
    assert len(rows) == len(tile.rowInfos)
    selected_rows = [i for i, item in enumerate(tile.rowInfos) if item.tile_row_index == (row - 1) % 256]
    assert len(selected_rows) == 1
    index = selected_rows[0]
    row_info = tile.rowInfos[index]
    _, _, row_start, row_end = rows[index]
    buffers = [item for item in fields(data, row_start, row_end) if item[0:2] == (6, 2)]
    assert len(buffers) == 1
    _, _, buffer_start, buffer_end = buffers[0]
    encoded = struct.unpack_from('<H', row_info.cell_offsets, (column - 1) * 2)[0]
    assert encoded != 65535
    cell = buffer_start + encoded * (4 if row_info.has_wide_offsets else 1)
    assert cell + 12 <= buffer_end and data[cell] == 5
    flags = struct.unpack_from('<I', data, cell + 8)[0]
    assert flags & 3 in (1, 2) and not flags & (1 << 9)
    length = 16 if flags & 1 else 8
    assert cell + 12 + length <= buffer_end
    scalar = data[cell + 12:cell + 12 + length]
    value = _unpack_decimal128(scalar) if length == 16 else struct.unpack('<d', scalar)[0]
    if kind == 'boolean':
        assert data[cell + 1] == 6 and length == 8 and value in (0, 1)
        value = bool(value)
    else:
        assert data[cell + 1] in (2, 10)
    return tile_id, cell + 12, length, value


def verify(source, output, descriptor, cells):
    pool = descriptor_pool.DescriptorPool()
    for item in descriptor_pb2.FileDescriptorSet.FromString(descriptor.read_bytes()).file:
        pool.Add(item)
    with ZipFile(source) as original, ZipFile(output) as saved:
        names = {item.filename for item in original.infolist() if not item.is_dir()}
        assert len(original.namelist()) == len(set(original.namelist()))
        assert len(saved.namelist()) == len(set(saved.namelist())) and names == set(saved.namelist())
        before_graph, after_graph = preservation.archives(original), preservation.archives(saved)
        before_records = {identifier: (kind, raw[start:end]) for raw, rows in before_graph.values()
            for identifier, index, kind, start, end in rows if index == 0}
        after_records = {identifier: (kind, raw[start:end]) for raw, rows in after_graph.values()
            for identifier, index, kind, start, end in rows if index == 0}
        edits = {}
        for table_id, row, column, kind, expected in cells:
            old = locate(before_records, pool, table_id, row, column, kind)
            new = locate(after_records, pool, table_id, row, column, kind)
            assert old[:3] == new[:3] and new[3] == expected
            tile_id, start, length, _ = old
            spans = edits.setdefault(tile_id, [])
            assert all(start + length <= a or start >= b for a, b in spans)
            spans.append((start, start + length))
        record_count, changed_tiles = 0, set()
        assert before_graph.keys() == after_graph.keys()
        for path, (raw, rows) in before_graph.items():
            after, after_rows = after_graph[path]
            assert rows == after_rows and len(raw) == len(after)
            normalized = bytearray(after)
            record_count += len(rows)
            for identifier, index, kind, start, end in rows:
                if identifier not in edits or index != 0:
                    continue
                assert kind == 6002
                for a, b in edits[identifier]:
                    normalized[start + a:start + b] = raw[start + a:start + b]
                changed_tiles.add(identifier)
            assert normalized == raw, path
        assert changed_tiles == set(edits)
        changed = [name for name in sorted(names) if original.read(name) != saved.read(name)]
        expected_paths = {path for path, (_, rows) in before_graph.items() if any(row[0] in edits for row in rows)}
        assert set(changed) == expected_paths
    return {'source': source.name, 'output': output.name,
        'sourceSha256': hashlib.sha256(source.read_bytes()).hexdigest(),
        'outputSha256': hashlib.sha256(output.read_bytes()).hexdigest(),
        'fileEntries': len(names), 'recordsIncludingAuxiliary': record_count,
        'changedEntries': changed, 'changedTiles': sorted(changed_tiles), 'cells': cells,
        'onlyDeclaredScalarSpansDiffer': True, 'rawHeadersSiblingsAndOtherBytesIdentical': True,
        'numbersParser': version('numbers-parser'), 'protobuf': version('protobuf'),
        'descriptorSha256': hashlib.sha256(descriptor.read_bytes()).hexdigest()}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('output', type=Path)
    parser.add_argument('--descriptor', required=True, type=Path)
    parser.add_argument('--cell', action='append', required=True, nargs=5,
        metavar=('TABLE_ID', 'ROW', 'COLUMN', 'KIND', 'VALUE'))
    parser.add_argument('--out', required=True, type=Path)
    args = parser.parse_args()
    cells = []
    for table, row, column, kind, value in args.cell:
        assert kind in ('number', 'boolean')
        assert kind != 'boolean' or value.lower() in ('true', 'false')
        cells.append((int(table), int(row), int(column), kind, value.lower() == 'true' if kind == 'boolean' else float(value)))
    result = verify(args.source, args.output, args.descriptor, cells)
    args.out.write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(result, indent=2))
