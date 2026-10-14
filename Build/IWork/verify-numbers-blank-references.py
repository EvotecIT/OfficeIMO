"""Verify bounded native integer references using typed storage, raw AST and XLSX.

Opt-in evidence only. Requires numbers-parser4.19.0; uses raw tile row indexes
to retain leading unstored rows. Public presentation strings are not the
coordinate oracle. Does not load OfficeIMO or modify the input artifacts.
"""
import argparse
import hashlib
import json
import math
import uuid
from importlib.metadata import version
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from numbers_parser import Document

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def native_xlsx(path):
    with ZipFile(path) as archive:
        names = [s.get('name') for s in ET.fromstring(archive.read('xl/workbook.xml')).findall('s:sheets/s:sheet', NS)]
        assert names == ['Inputs - Operands', 'Checks - Results']
        return [{c.get('r'): {'formula': c.findtext('s:f', namespaces=NS),
                            'value': c.findtext('s:v', namespaces=NS), 'type': c.get('t')}
                 for c in ET.fromstring(archive.read(f'xl/worksheets/sheet{i}.xml')).findall('.//s:c', NS)}
                for i in (1, 2)]


def identifier(words):
    return str(uuid.UUID(bytes=b''.join(w.to_bytes(4, 'little') for w in
        (words.uuid_w0, words.uuid_w1, words.uuid_w2, words.uuid_w3))))


def formula_text(nodes, target):
    stack = []
    for node in nodes:
        kind = node.AST_node_type
        if kind == 36:
            assert identifier(node.AST_cross_table_reference_extra_info.table_id) == target
            assert node.AST_column.absolute and node.AST_row.absolute
            assert node.AST_column.column == 2 and node.AST_row.row in (2, 3, 4, 7)
            stack.append("'Inputs - Operands'!$C$" + str(node.AST_row.row + 1))
        elif kind == 17:
            assert node.AST_number_node_number == 6
            stack.append('6')
        elif kind == 29:
            last, first = stack.pop(), stack.pop()
            stack.append(first + ':' + last.rsplit('!', 1)[-1])
        elif kind == 16:
            count = node.AST_function_node_numArgs
            assert node.AST_function_node_index in (58, 75) and 1 <= count <= len(stack)
            args = stack[-count:]
            del stack[-count:]
            stack.append(('GCD' if node.AST_function_node_index == 58 else 'LCM') + '(' + ','.join(args) + ')')
        else:
            raise AssertionError('Unqualified AST node: ' + str(kind))
    assert len(stack) == 1
    return stack[0]


def verify(source, exported, filled):
    document = Document(source)
    assert [s.name for s in document.sheets] == ['Inputs - Operands', 'Checks - Results']
    inputs, checks = [s.tables[0] for s in document.sheets]
    assert all(len(s.tables) == 1 for s in document.sheets)
    assert [(t.num_rows, t.num_cols) for t in (inputs, checks)] == [(10, 5), (10, 5)]
    row_maps = {}
    for table in (inputs, checks):
        storage = document._model.objects[table._table_id].base_data_store.tiles
        assert storage.tile_size == 256 and len(storage.tiles) == 1 and storage.tiles[0].tileid == 0
        rows = document._model.objects[storage.tiles[0].tile.identifier].rowInfos
        positions = [row.tile_row_index for row in rows]
        assert positions == sorted(set(positions)) and positions[0] == 1
        row_maps[table._table_id] = {row: i for i, row in enumerate(positions)}

    def cell(table, row, column):
        position = row_maps[table._table_id][row - 1]
        return table.cell(position, column - 1)

    operand = cell(inputs, 4, 3)
    assert type(operand).__name__ == ('NumberCell' if filled else 'EmptyCell')
    assert operand.value == (18 if filled else None) and not operand.is_formula
    assert cell(inputs, 3, 3).value == 30 and cell(inputs, 8, 3).value == 0
    assert type(cell(inputs, 5, 3)).__name__ == 'TextCell' and cell(inputs, 5, 3).value == '12'
    target = str(uuid.UUID(document._model.objects[inputs._table_id].table_id))
    xlsx = native_xlsx(exported)
    assert xlsx[0].get('C4', {}).get('value') == ('18' if filled else None)
    values = [18, 6] if filled else [6]
    expected = {3: (math.gcd(*values), math.lcm(*values)),
                4: (math.gcd(30, 18 if filled else 0), math.lcm(30, 18 if filled else 0)),
                5: (6, 12), 6: (18, 18) if filled else (None, None),
                7: (18, 18) if filled else (None, None), 8: (12, 12),
                9: (18 if filled else 0, 0)}
    records = []
    for row, pair in expected.items():
        for column, value in zip((3, 4), pair):
            item = cell(checks, row, column)
            assert item.is_formula
            nodes = document._model.formula_ast(checks._table_id)[item._formula_id]
            formula = formula_text(nodes, target)
            address = chr(64 + column) + str(row)
            assert xlsx[1][address]['formula'] == formula
            if value is None:
                assert type(item).__name__ == 'ErrorCell' and item.value is None
                assert xlsx[1][address]['value'] is None
            else:
                assert type(item).__name__ == 'NumberCell' and item.value == value
                assert float(xlsx[1][address]['value']) == value
            records.append({'row': row, 'column': column, 'excelFormula': formula,
                'nativeValue': value, 'qualified': value is not None,
                'nativeKind': type(item).__name__, 'nodeTypes': [n.AST_node_type for n in nodes]})
    assert sum(c.is_formula for row in checks.rows() for c in row) == 14
    return {'source': source.name, 'nativeXlsx': exported.name, 'cells': records,
        'qualifiedFormulaCells': 14 if filled else 10,
        'referenceTableUuid': target,
        'nativeTileRows': {name: row_maps[t._table_id]
                           for name, t in [('inputs', inputs), ('checks', checks)]}}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('root', type=Path)
    parser.add_argument('--out', type=Path, required=True)
    args = parser.parse_args()
    assert version('numbers-parser') == '4.19.0'
    files = [args.root / ('numbers-blank-reference' + suffix + '-v14.5' + extension)
             for suffix in ('', '-filled') for extension in ('.numbers', '.xlsx')]
    cases = [verify(files[0], files[1], False), verify(files[2], files[3], True)]
    assert all(a['excelFormula'] == b['excelFormula'] for a, b in zip(cases[0]['cells'], cases[1]['cells']))
    result = {'schemaVersion': 1, 'producer': 'Numbers14.5 (7045.0.17)',
        'decoder': 'numbers-parser4.19.0, raw tile row indexes and AST; native XLSX XML', 'cases': cases,
        'inputs': [{'path': p.name, 'bytes': p.stat().st_size, 'sha256': hashlib.sha256(p.read_bytes()).hexdigest()} for p in files],
        'limits': ['Controlled Apple-imported/saved OfficeIMO-owned XLSX fixture; not arbitrary independent producers.',
                   'All-blank direct arguments retain native errors and remain partial.',
                   'No mixed text/Boolean ranges, formula dependencies, whole axes or appearance equivalence.']}
    args.out.write_text(json.dumps(result, indent=2) + '\n')
    print('PASS | ten blank/native-text calls, four all-blank error controls and fourteen filled calls agree with raw AST and native XLSX')
