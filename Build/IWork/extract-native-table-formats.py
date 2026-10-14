"""Extract selected numeric formats from native Apple table references.

Opt-in evidence only. Uses a pinned descriptor and numbers-parser's independent
IWA and Decimal128 decoders; never loads OfficeIMO assemblies.
"""
import argparse
import hashlib
import json
import struct
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile
from google.protobuf import descriptor_pb2, descriptor_pool, message_factory
from numbers_parser.cell import _unpack_decimal128
from iwork_protobuf_evidence import archive_messages


def extract(path, descriptor):
    pool = descriptor_pool.DescriptorPool()
    descriptor_bytes = descriptor.read_bytes()
    for file in descriptor_pb2.FileDescriptorSet.FromString(descriptor_bytes).file:
        pool.Add(file)

    def decode(name, payload):
        message = message_factory.GetMessageClass(pool.FindMessageTypeByName(name))()
        message.ParseFromString(payload)
        assert not message.FindInitializationErrors(), (name, message.FindInitializationErrors())
        return message

    with ZipFile(path) as package:
        graph = {header.identifier: (info.type, payload)
                 for _, header, info, payload in archive_messages(package, primary_only=True)}
    tables = []
    for identifier, (kind, payload) in graph.items():
        if kind != 6001:
            continue
        model = decode('TST.TableModelArchive', payload)
        store = model.base_data_store
        formats = decode('TST.TableDataList', graph[store.format_table.identifier][1])
        catalog = {entry.key: entry.format for entry in formats.entries if entry.HasField('format')}
        cells = []
        for tile_ref in store.tiles.tiles:
            tile = decode('TST.Tile', graph[tile_ref.tile.identifier][1])
            assert tile.storage_version == 5 and not tile.should_use_wide_rows
            for row in tile.rowInfos:
                assert len(row.cell_offsets) % 2 == 0
                offsets = struct.unpack(f'<{len(row.cell_offsets) // 2}H', row.cell_offsets)
                occupied = [(column, offset) for column, offset in enumerate(offsets) if offset != 65535]
                for index, (column, offset) in enumerate(occupied):
                    end = occupied[index + 1][1] if index + 1 < len(occupied) else len(row.cell_storage_buffer)
                    data = row.cell_storage_buffer[offset:end]
                    assert data[0] == 5
                    if data[1] != 2:
                        continue
                    flags = struct.unpack_from('<I', data, 8)[0]
                    position, selected = 12, {}
                    for bit in range(22):
                        if flags & (1 << bit):
                            length = 16 if bit == 0 else 8 if bit in (1, 2) else 4
                            selected[bit] = data[position:position + length]
                            assert len(selected[bit]) == length
                            position += length
                    assert flags >> 22 == 0 and position == len(data)
                    assert 0 in selected and 13 in selected
                    key = struct.unpack('<I', selected[13])[0]
                    item = catalog[key]
                    assert item.format_type in (256, 258, 259)
                    cells.append({'row': tile_ref.tileid * 256 + row.tile_row_index + 1,
                        'column': column + 1, 'value': _unpack_decimal128(selected[0]),
                        'activeNumericFormat': bool(struct.unpack_from('<H', data, 6)[0] & 1),
                        'suggestedType': struct.unpack('<I', selected[12])[0] if 12 in selected else None,
                        'formatKey': key,
                        'numberFormat': {'kind': {256: 'Number', 258: 'Percentage', 259: 'Scientific'}[item.format_type],
                            'decimalPlaces': item.decimal_places, 'negativeStyle': item.negative_style,
                            'thousandsSeparator': item.show_thousands_separator}})
        tables.append({'id': identifier, 'name': model.table_name, 'rows': model.number_of_rows,
            'columns': model.number_of_columns, 'numericCells': cells})
    return {'source': path.name, 'bytes': path.stat().st_size,
        'sha256': hashlib.sha256(path.read_bytes()).hexdigest(),
        'descriptorSha256': hashlib.sha256(descriptor_bytes).hexdigest(),
        'numbersParserVersion': version('numbers-parser'), 'tables': tables}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('package', type=Path)
    parser.add_argument('--descriptor', required=True, type=Path)
    parser.add_argument('--out', required=True, type=Path)
    args = parser.parse_args()
    assert version('numbers-parser') == '4.19.0'
    result = extract(args.package, args.descriptor)
    args.out.write_text(json.dumps(result, indent=2, ensure_ascii=False) + '\n')
    print({'source': result['source'], 'sha256': result['sha256'],
        'tables': len(result['tables']), 'numericCells': sum(len(table['numericCells']) for table in result['tables'])})
