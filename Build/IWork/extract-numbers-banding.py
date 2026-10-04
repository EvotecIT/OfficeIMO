"""Extract the bounded native Numbers banding fixture and its independent XLSX fill oracle.

Requires opt-in numbers-parser 4.19.0. Run with the corpus directory and output manifest path.
The source and exports are produced by Apple Numbers, not by OfficeIMO.
"""
import argparse
import hashlib
import json
import struct
from importlib.metadata import version
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile
from numbers_parser.iwafile import IWACompressedChunk, get_archive_info_and_remainder
from numbers_parser.generated.TSTArchives_pb2 import TableModelArchive, TableStyleArchive, CellStyleArchive, Tile

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('corpus', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
assert version('numbers-parser') == '4.19.0'
base = 'native-exports/numbers-banding-v14.5'
source = args.corpus / (base + '.numbers')
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


def role_fill(identifier):
    kind, payload = records[identifier]
    assert kind == 6004
    style = CellStyleArchive.FromString(payload)
    assert not style.super.HasField('parent')
    assert style.cell_properties.HasField('cell_fill')
    fill = style.cell_properties.cell_fill
    if not fill.HasField('color'):
        assert fill.SerializeToString() == b''
        return None
    color = fill.color
    assert color.a == 1 and color.model == 1 and color.rgbspace == 1
    return ''.join(f'{round(component * 255):02X}' for component in (color.r, color.g, color.b))


ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with ZipFile(args.corpus / (base + '.xlsx')) as package:
    styles = ET.fromstring(package.read('xl/styles.xml'))
    palette = styles.find('s:colors/s:indexedColors', ns)
    fills = styles.find('s:fills', ns)
    formats = styles.find('s:cellXfs', ns)
    workbook = ET.fromstring(package.read('xl/workbook.xml'))
    exports = {}
    for index, sheet in enumerate(workbook.find('s:sheets', ns), 1):
        content = ET.fromstring(package.read(f'xl/worksheets/sheet{index}.xml'))
        cells = []
        for cell in content.findall('s:sheetData/s:row/s:c', ns):
            address = cell.get('r')
            row = int(address[1:]) - 1  # The native table title occupies row one.
            if row == 0:
                continue
            pattern = fills[int(formats[int(cell.get('s', '0'))].get('fillId', '0'))].find('s:patternFill', ns)
            color = None
            if pattern.get('patternType') == 'solid':
                foreground = pattern.find('s:fgColor', ns)
                assert 'indexed' in foreground.attrib
                color = palette[int(foreground.get('indexed'))].get('rgb')[2:].upper()
            else:
                assert pattern.get('patternType') == 'none'
            cells.append({'row': row, 'column': ord(address[0]) - ord('A') + 1, 'rgb': color})
        assert len(cells) == 24
        exports[sheet.get('name')] = {'worksheetIndex': index, 'cells': cells}


tables = []
for identifier, (kind, payload) in sorted(records.items()):
    if kind != 6001:
        continue
    model = TableModelArchive.FromString(payload)
    assert (model.number_of_rows, model.number_of_columns, model.number_of_header_columns,
            model.number_of_footer_rows) == (8, 3, 1, 1)
    for selected_tile in model.base_data_store.tiles.tiles:
        tile_kind, tile_payload = records[selected_tile.tile.identifier]
        assert tile_kind == 6002
        for row in Tile.FromString(tile_payload).rowInfos:
            offsets = struct.unpack('<' + 'H' * (len(row.cell_offsets) // 2), row.cell_offsets)
            for offset in offsets:
                if offset == 65535:
                    continue
                if row.has_wide_offsets:
                    offset *= 4
                assert row.cell_storage_buffer[offset] == 5
                assert not struct.unpack_from('<I', row.cell_storage_buffer, offset + 8)[0] & 0x20
    style_kind, style_payload = records[model.table_style.identifier]
    assert style_kind == 6003
    style = TableStyleArchive.FromString(style_payload)
    assert not style.super.HasField('parent') and style.table_properties.banded_rows
    band_color = style.table_properties.banded_fill.color
    assert band_color.a == 1 and band_color.model == 1 and band_color.rgbspace == 1
    band = ''.join(f'{round(component * 255):02X}' for component in (band_color.r, band_color.g, band_color.b))
    roles = {name: role_fill(getattr(model, name).identifier) for name in
             ('body_cell_style', 'header_row_style', 'header_column_style', 'footer_row_style')}
    name = f'Headers{model.number_of_header_rows}'
    export = exports[name]
    for cell in export['cells']:
        row, column = cell['row'], cell['column']
        expected = (roles['header_row_style'] if row <= model.number_of_header_rows else
                    roles['footer_row_style'] if row == 8 else
                    roles['header_column_style'] if column == 1 else
                    band if (row - model.number_of_header_rows) % 2 == 0 else roles['body_cell_style'])
        assert cell['rgb'] == expected
    tables.append({'modelIdentifier': identifier, 'sheet': name, 'headerRows': model.number_of_header_rows,
                   'bandedBodyRgb': band, 'roles': roles, **export})
assert {table['headerRows'] for table in tables} == {0, 1, 2}
args.output.write_text(json.dumps({
    'schemaVersion': 1, 'provider': 'numbers-parser', 'providerVersion': '4.19.0',
    'sourceFixture': base + '.numbers', 'sourceSha256': hashlib.sha256(source.read_bytes()).hexdigest(),
    'producer': {'application': 'Apple Numbers', 'version': '14.5', 'build': '7045.0.17',
                 'operatingSystem': 'macOS 27.0.1 26A434', 'exportedUtc': '2026-10-01T14:59:53Z'},
    'sourceLicense': 'OfficeIMO-authored fixture; repository MIT license',
    'exportOptions': {'xlsx': 'One worksheet per table; overview disabled; no password',
                      'pdf': 'Fit each sheet to one page; best image quality; no password',
                      'nativeTableTitleRowOffset': 1},
    'pdf': {'pageCount': 3, 'fontProvenance': 'Apple Numbers/Quartz embedded subsets; no standalone fonts copied'},
    'artifacts': [{'path': Path(base).name + suffix,
                   'sha256': hashlib.sha256((args.corpus / (base + suffix)).read_bytes()).hexdigest()}
                  for suffix in ('.xlsx', '.pdf')],
    'tables': tables,
    'limitations': 'Qualifies opaque sRGB role fills and alternating body-row fills for this Numbers 14.5 fixture. Selected overrides are tested separately. Other producers, gradients, transparency, complete appearance and pagination are not qualified.'
}, indent=2) + '\n', encoding='utf-8')
