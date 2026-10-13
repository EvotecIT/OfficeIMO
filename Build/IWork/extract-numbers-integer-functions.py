"""Decode the controlled Numbers GCD/LCM intake and fresh native Excel results.

Requires numbers-parser 4.19.0. All inputs remain unchanged. This is an opt-in
independent-producer evidence extractor, not an OfficeIMO runtime dependency.
"""
import argparse
import hashlib
import json
from importlib.metadata import version
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from numbers_parser import Document
from numbers_parser.generated.functionmap import FUNCTION_MAP

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('source', type=Path)
parser.add_argument('native_xlsx', type=Path)
parser.add_argument('excel_xlsx', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
assert version('numbers-parser') == '4.19.0'
assert hashlib.sha256(args.source.read_bytes()).hexdigest() == '99f2f0445b9985feac658fa876502c854452026e383f078d41d41260a9cfc0b1'
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def read_xlsx(path):
    with ZipFile(path) as archive:
        strings = [''.join(si.itertext()) for si in ET.fromstring(archive.read('xl/sharedStrings.xml'))]
        result = {}
        for cell in ET.fromstring(archive.read('xl/worksheets/sheet1.xml')).findall('.//s:c', ns):
            kind, value = cell.get('t', 'n'), cell.findtext('s:v', namespaces=ns)
            if kind == 's':
                value = strings[int(value)]
            result[cell.get('r')] = {'kind': kind, 'formula': cell.findtext('s:f', namespaces=ns), 'value': value}
        return result


document = Document(str(args.source))
table = document.sheets[0].tables[0]
assert table.num_rows == 18 and table.num_cols == 8
numbers_export, excel_export = read_xlsx(args.native_xlsx), read_xlsx(args.excel_xlsx)
cells = []
for row in range(1, 18):
    for column in range(1, 8):
        cell = table.cell(row, column)
        address = f'{chr(65 + column)}{row + 1}'
        empty = {'kind': 'n', 'formula': None, 'value': None}
        native = numbers_export.get(address, empty)
        fresh = excel_export.get(address, empty)
        provider_formula = cell.formula if cell.is_formula else None
        formula = native['formula']
        assert formula == fresh['formula'], address
        if row + 1 in (13, 14) and column + 1 in (7, 8):
            # The independent decoder prints these large literals with .0.
            assert formula.replace('9007199254740990', '9007199254740990.0') == provider_formula
        else:
            assert formula == provider_formula, address
        functions = []
        if formula:
            nodes = document._model.formula_ast(table._table_id)[cell._formula_id]
            functions = [{'index': n.AST_function_node_index, 'name': FUNCTION_MAP[n.AST_function_node_index],
                          'arguments': n.AST_function_node_numArgs} for n in nodes if n.AST_node_type == 16]
        cells.append({'row': row + 1, 'column': column + 1, 'address': address,
                      'case': table.cell(row, 0).value, 'nativeType': type(cell).__name__,
                      'nativeValue': cell.value, 'formula': formula, 'providerFormula': provider_formula, 'functions': functions,
                      'numbersXlsx': native, 'excelRecalculatedXlsx': fresh})

manifest = {'source': args.source.name, 'nativeXlsx': args.native_xlsx.name, 'excelXlsx': args.excel_xlsx.name,
            'producer': 'Numbers 14.5 (7045.0.17)', 'excelProducer': 'Microsoft Excel 16.113.3',
            'decoder': 'numbers-parser 4.19.0', 'rows': 18, 'columns': 8,
            'formulaCells': sum(c['formula'] is not None for c in cells), 'cells': cells,
            'inputs': [{'path': p.name, 'bytes': p.stat().st_size,
                        'sha256': hashlib.sha256(p.read_bytes()).hexdigest()}
                       for p in [args.source, args.native_xlsx, args.excel_xlsx]],
            'limits': ['Numbers source was created by importing the controlled XLSX and saved natively. B13 and B14 were then replaced natively by =2^53-1 and =2^53.',
                       'Excel results were captured after Calculate Now and saved to a separate local workbook.',
                       'Numbers rounds the B13/B14 power caches to 9007199254740990. G/H large literal operands are separately rounded import values.',
                       'Numeric-text range LCM and large results differ. Native Numbers error cells have no XLSX error code.',
                       'Native PDF and both app screenshots were inspected. This fixture does not qualify other producer versions or general native appearance.']}
assert manifest['formulaCells'] == 70
args.output.write_text(json.dumps(manifest, indent=2) + '\n')
print(f"Decoded {manifest['formulaCells']} formulas and both native XLSX exports")
