"""Independently check the controlled Apple-edited integer reference fixture.

Requires numbers-parser 4.19.0. Native inputs remain unchanged; this opt-in
evidence extractor is not an OfficeIMO runtime dependency.
"""
import argparse
import hashlib
import json
import math
from importlib.metadata import version
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from numbers_parser import Document

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('source', type=Path)
parser.add_argument('native_xlsx', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
assert version('numbers-parser') == '4.19.0'
assert hashlib.sha256(args.source.read_bytes()).hexdigest() == 'a09e06af365cfbe210df7c1f5dc781630061df61c0e74ed2571a9e0c834b4970'
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with ZipFile(args.native_xlsx) as archive:
    root = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
    exported = {c.get('r'): {'kind': c.get('t', 'n'), 'formula': c.findtext('s:f', namespaces=ns),
                            'value': c.findtext('s:v', namespaces=ns)} for c in root.findall('.//s:c', ns)}

document = Document(str(args.source))
table = document.sheets[0].tables[0]
assert table.num_rows == 18 and table.num_cols == 8
assert table.cell(1, 1).value == 30
expected = {'G2': ('GCD(B2,C2,D2)', 6), 'H2': ('LCM($B$2,$C$2:$D$2,1)', 180)}
for row in (2, 4, 5, 6, 17, 18):
    cells = [table.cell(row - 1, column) for column in (1, 2, 3)]
    assert all(type(c).__name__ == 'NumberCell' and not c.is_formula for c in cells)
    values = [math.trunc(c.value) for c in cells]
    expected[f'E{row}'] = (f'GCD(B{row}:D{row})', math.gcd(*values))
    expected[f'F{row}'] = (f'LCM(B{row}:D{row})', math.lcm(*values))

selected = []
for address, (formula, value) in expected.items():
    row, column = int(address[1:]), ord(address[0]) - 64
    cell = table.cell(row - 1, column - 1)
    native = exported[address]
    assert cell.is_formula and cell.formula == formula, address
    assert cell.value == value and native['formula'] == formula and float(native['value']) == value, address
    nodes = document._model.formula_ast(table._table_id)[cell._formula_id]
    selected.append({'address': address, 'row': row, 'column': column, 'formula': formula,
                     'nativeValue': cell.value, 'independentIntegerValue': value, 'nativeXlsx': native,
                     'nodeTypes': [n.AST_node_type for n in nodes],
                     'functions': [{'index': n.AST_function_node_index, 'arguments': n.AST_function_node_numArgs}
                                   for n in nodes if n.AST_node_type == 16]})
assert len(selected) == 14
assert table.cell(9, 5).value == 0  # Numeric-text LCM remains outside qualification.
assert table.cell(9, 1).value == '24'
assert sum(c.is_formula for row in table.rows() for c in row) == 70
result = {'schemaVersion': 1, 'producer': 'Numbers 14.5 (7045.0.17)', 'decoder': 'numbers-parser 4.19.0',
          'source': args.source.name, 'nativeXlsx': args.native_xlsx.name, 'rows': 18, 'columns': 8,
          'formulaCells': 70, 'qualifiedReferenceCalls': 14, 'cells': selected,
          'inputs': [{'path': p.name, 'bytes': p.stat().st_size,
                      'sha256': hashlib.sha256(p.read_bytes()).hexdigest()} for p in (args.source, args.native_xlsx)],
          'limits': ['Controlled local reference profile, saved/reopened in the recorded Numbers version.',
                     'Numeric-text LCM remains cached zero and is not a qualified expression.',
                     'No coercion, blank, formula-dependency, merge, whole-axis or general appearance qualification.']}
args.output.write_text(json.dumps(result, indent=2) + '\n')
print('PASS | 14 native numeric-reference calls agree with Apple XLSX and independent integer arithmetic')
