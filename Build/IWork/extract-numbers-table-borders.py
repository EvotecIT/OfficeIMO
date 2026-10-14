"""Verify a native Numbers border reference and every paired XLSX cell edge.

Opt-in evidence only: numbers-parser 4.19.0 and the pinned KN/TST descriptor.
Uses native declarations and Apple's independent XLSX export; never loads OfficeIMO.
The controlled 1/2-point strokes qualify thin/medium XLSX mapping, not arbitrary widths.
"""
import argparse
import hashlib
import json
from importlib.metadata import version
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from google.protobuf import descriptor_pb2, descriptor_pool, message_factory
from numbers_parser import Document
from iwork_protobuf_evidence import archive_messages

SOURCE_SHA256 = 'f26c9e0ef5e2e385dee2d117a491a9f078ab29bdf206154be774a922185a0911'
XLSX_SHA256 = 'c54562907c40affd5aadd60263dd98340a39fb821ce328b65f8e0dbd095c4531'
DESCRIPTOR_SHA256 = '2408cf25920c39fe59d500a0b580bd64bf10a24630c81ca2d543e64322853917'
SIDES = {'left': (False, 0, 0), 'right': (False, 0, 1), 'top': (True, 0, 0), 'bottom': (True, 1, 0)}


def native_xlsx_cells(path):
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    with ZipFile(path) as package:
        assert package.testzip() is None
        styles = ET.fromstring(package.read('xl/styles.xml'))
        strings = ET.fromstring(package.read('xl/sharedStrings.xml'))
        values = [''.join(text.text or '' for text in node.findall('.//s:t', ns)) for node in strings.findall('s:si', ns)]
        palette = [node.attrib['rgb'][-6:].upper() for node in styles.findall('s:colors/s:indexedColors/s:rgbColor', ns)]
        formats = styles.findall('s:cellXfs/s:xf', ns)
        borders = styles.findall('s:borders/s:border', ns)
        sheet = ET.fromstring(package.read('xl/worksheets/sheet1.xml'))
    cells = {}
    for cell in sheet.findall('s:sheetData/s:row/s:c', ns):
        address = cell.attrib['r']
        value = cell.find('s:v', ns)
        assert value is None or cell.attrib.get('t') == 's'
        border = borders[int(formats[int(cell.attrib.get('s', '0'))].attrib['borderId'])]
        edges = {}
        for side in SIDES:
            node = border.find('s:' + side, ns)
            style = None if node is None else node.attrib.get('style')
            assert style in (None, 'thin', 'medium')
            color = None if node is None else node.find('s:color', ns)
            rgb = '000000' if style is None else palette[int(color.attrib['indexed'])]
            edges[side] = {'widthPoints': {None: 0, 'thin': 1, 'medium': 2}[style], 'rgb': rgb}
        cells[address] = {'value': None if value is None else values[int(value.text)], 'edges': edges}
    assert len(cells) == 16
    return cells


def extract(source, descriptor, xlsx):
    assert version('numbers-parser') == '4.19.0'
    assert hashlib.sha256(source.read_bytes()).hexdigest() == SOURCE_SHA256
    assert hashlib.sha256(xlsx.read_bytes()).hexdigest() == XLSX_SHA256
    raw = descriptor.read_bytes()
    assert hashlib.sha256(raw).hexdigest() == DESCRIPTOR_SHA256
    pool = descriptor_pool.DescriptorPool()
    for file in descriptor_pb2.FileDescriptorSet.FromString(raw).file:
        pool.Add(file)
    graph = {}
    with ZipFile(source) as package:
        assert package.testzip() is None
        for _, header, info, content in archive_messages(package):
            graph.setdefault((header.identifier, info.type), []).append(content)

    def decode(identifier, kind, name):
        payloads = graph[(identifier, kind)]
        assert len(payloads) == 1
        message = message_factory.GetMessageClass(pool.FindMessageTypeByName(name)).FromString(payloads[0])
        assert not message.FindInitializationErrors()
        return message

    models = [identifier for identifier, kind in graph if kind == 6001]
    assert models == [905280]
    model = decode(models[0], 6001, 'TST.TableModelArchive')
    assert model.table_name == 'Tabela 1'
    assert (model.number_of_rows, model.number_of_columns) == (4, 4)
    assert (model.number_of_header_rows, model.number_of_header_columns, model.number_of_footer_rows) == (1, 1, 1)
    style = decode(model.table_style.identifier, 6003, 'TST.TableStyleArchive')
    defaults = [(field.name, value) for field, value in style.table_properties.ListFields()]
    assert len(defaults) == 16
    outlines = {'table_body_horizontal_border_stroke', 'table_body_vertical_border_stroke'}
    assert all(stroke.width == 1 and stroke.pattern.type == (1 if name in outlines else 2)
               and stroke.color.a == (1 if name in outlines else 0) for name, stroke in defaults)
    sidecar = decode(model.stroke_sidecar.identifier, 6305, 'TST.StrokeSidecarArchive')
    assert (sidecar.max_order, sidecar.row_count, sidecar.column_count) == (3, 4, 4)
    edges, layers = {}, []
    for name in ('left_column_stroke_layers', 'right_column_stroke_layers', 'top_row_stroke_layers', 'bottom_row_stroke_layers'):
        horizontal = 'row' in name
        increment = 1 if name.startswith(('right', 'bottom')) else 0
        references = getattr(sidecar, name)
        assert len(references) == 4
        for reference in references:
            layer = decode(reference.identifier, 6306, 'TST.StrokeLayerArchive')
            line = layer.row_column_index + increment
            runs = []
            for run in layer.stroke_runs:
                stroke = run.stroke
                assert stroke.pattern.type == 1 and stroke.width in (1, 2, 3) and stroke.color.a == 1
                assert 0 <= run.origin < 4 and 0 < run.length <= 4 - run.origin and run.order <= sidecar.max_order
                rgb = ''.join(f'{round(component * 255):02X}' for component in (stroke.color.r, stroke.color.g, stroke.color.b))
                value = {'widthPoints': stroke.width, 'rgb': rgb}
                runs.append({'origin': run.origin, 'length': run.length, 'order': run.order, **value})
                for position in range(run.origin, run.origin + run.length):
                    key = (horizontal, line, position)
                    previous = edges.get(key)
                    if previous is None or previous[0] < run.order:
                        edges[key] = (run.order, value)
                    elif previous[0] == run.order:
                        assert previous[1] == value
            layers.append({'id': reference.identifier, 'side': name, 'line': line, 'runs': runs})
    assert edges[(True, 2, 1)] == (3, {'widthPoints': 2, 'rgb': 'FF3333'})
    assert (False, 2, 1) not in edges
    owner = decode(model.category_owner.identifier, 6372, 'TST.CategoryOwnerRefArchive')
    assert owner.group_by and all(not decode(ref.identifier, 6373, 'TST.GroupByArchive').is_enabled for ref in owner.group_by)
    document = Document(source)
    table = document.sheets[0].tables[0]
    exported = native_xlsx_cells(xlsx)
    cells = []
    for row in range(4):
        for column in range(4):
            address = chr(65 + column) + str(row + 1)
            sides = {}
            for side, (horizontal, row_delta, column_delta) in SIDES.items():
                line = row + row_delta if horizontal else column + column_delta
                position = column if horizontal else row
                sides[side] = edges.get((horizontal, line, position), (0, {'widthPoints': 0, 'rgb': '000000'}))[1]
            value = table.cell(row, column).value
            assert exported[address] == {'value': value, 'edges': sides}, (address, exported[address], value, sides)
            cells.append({'row': row + 1, 'column': column + 1, 'address': address, 'value': value, 'edges': sides})
    assert len(cells) == 16 and sum(len(cell['edges']) for cell in cells) == 64
    return {'sourceSha256': SOURCE_SHA256, 'nativeXlsxSha256': XLSX_SHA256, 'descriptorSha256': DESCRIPTOR_SHA256,
            'model': models[0], 'table': model.table_name, 'rows': 4, 'columns': 4,
            'headerRows': 1, 'headerColumns': 1, 'footerRows': 1,
            'emptyDefaultStrokes': [name for name, _ in defaults if name not in outlines],
            'shadowedOutlineStrokes': sorted(outlines), 'layers': layers, 'cells': cells}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('descriptor', type=Path)
    parser.add_argument('native_xlsx', type=Path)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    args.output.write_text(json.dumps(extract(args.source, args.descriptor, args.native_xlsx), indent=2) + '\n')
    print('Native Numbers declarations, later shared-edge edit and all 64 native XLSX cell edges verified.')
