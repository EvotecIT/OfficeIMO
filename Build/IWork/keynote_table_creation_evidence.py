"""Independent scalar-table checks for the owned native Keynote creation profile.

Opt-in validation only. Decodes native cell storage through numbers-parser's
Decimal128 implementation; never loads OfficeIMO or becomes a runtime provider.
"""
from collections import Counter
import struct
from numbers_parser.cell import _unpack_decimal128


def verify_tables(graph, decoded):
    tables = []
    for identifier, model in decoded.items():
        if graph[identifier][0] != 6001:
            continue
        assert 1 <= model.number_of_rows <= 256 and 1 <= model.number_of_columns <= 256
        assert model.number_of_header_rows == model.number_of_header_columns == model.number_of_footer_rows == 0
        assert not decoded[model.table_style.identifier].table_properties.auto_resize and not model.table_name_enabled
        store = model.base_data_store
        assert len(store.tiles.tiles) == 1
        tile = decoded[store.tiles.tiles[0].tile.identifier]
        assert tile.storage_version == 5 and tile.last_saved_in_BNC and not tile.should_use_wide_rows
        assert tile.numrows == len(tile.rowInfos) == model.number_of_rows
        strings = decoded[store.stringTable.identifier]
        values = {entry.key: entry.string for entry in strings.entries}
        assert len(values) == len(strings.entries)
        refcounts = Counter()
        cells = []
        for row_number, row in enumerate(tile.rowInfos):
            assert row.tile_row_index == row_number and row.cell_count == model.number_of_columns
            assert len(row.cell_offsets) == len(row.cell_offsets_pre_bnc) == 512
            offsets = struct.unpack('<256H', row.cell_offsets)
            legacy_offsets = struct.unpack('<256H', row.cell_offsets_pre_bnc)
            assert offsets[0] == 0 and all(offset == 65535 for offset in offsets[model.number_of_columns:])
            assert len(row.cell_storage_buffer_pre_bnc) == model.number_of_columns * 12
            for column in range(model.number_of_columns):
                assert legacy_offsets[column] == column * 12
                assert row.cell_storage_buffer_pre_bnc[column * 12:column * 12 + 12] == bytes([4]) + bytes(11)
                end = offsets[column + 1] if column + 1 < model.number_of_columns else len(row.cell_storage_buffer)
                assert offsets[column] < end <= len(row.cell_storage_buffer)
                data = row.cell_storage_buffer[offsets[column]:end]
                assert data[0] == 5
                kind = {0: 'Empty', 2: 'Number', 3: 'Text', 6: 'Boolean'}[data[1]]
                flags = struct.unpack_from('<I', data, 8)[0]
                value = None
                if kind == 'Number':
                    assert flags == 1 and len(data) == 28
                    value = _unpack_decimal128(data[12:28])
                elif kind == 'Text':
                    assert flags == 8 and len(data) == 16
                    key = struct.unpack_from('<I', data, 12)[0]
                    value = values[key]
                    refcounts[key] += 1
                elif kind == 'Boolean':
                    assert flags == 2 and len(data) == 20
                    number = struct.unpack_from('<d', data, 12)[0]
                    assert number in (0, 1)
                    value = bool(number)
                else:
                    assert flags == 0 and len(data) == 12
                cells.append({'row': row_number + 1, 'column': column + 1, 'kind': kind, 'value': value})
        assert tile.numCells == len(cells) == model.number_of_rows * model.number_of_columns
        assert all(entry.refcount == refcounts[entry.key] for entry in strings.entries)
        row_headers = decoded[store.rowHeaders.buckets[0].identifier].headers
        column_headers = decoded[store.columnHeaders.identifier].headers
        assert len(row_headers) == model.number_of_rows and len(column_headers) == model.number_of_columns
        assert all(header.index == i and header.size == model.default_row_height
                   and header.numberOfCells == model.number_of_columns for i, header in enumerate(row_headers))
        assert all(header.index == i and header.size == model.default_column_width
                   and header.numberOfCells == model.number_of_rows for i, header in enumerate(column_headers))
        preset = decoded[model.table_style_preset.identifier]
        network = decoded[preset.style_network.identifier]
        assert network.table_style.identifier == model.table_style.identifier
        assert network.body_cell_style.identifier == model.body_cell_style.identifier
        assert network.body_text_style.identifier == model.body_text_style.identifier
        tables.append({'id': identifier, 'name': model.table_name, 'rows': model.number_of_rows,
                       'columns': model.number_of_columns, 'cells': cells,
                       'row_offsets_and_counts_complete': True, 'style_preset_network_complete': True})
    return tables
