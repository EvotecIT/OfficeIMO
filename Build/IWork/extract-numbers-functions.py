"""Extract native Numbers function expressions, IDs and XLSX result oracles.

Requires opt-in numbers-parser 4.19.0. The intake files contain no formula caches;
Apple Numbers imports, evaluates, saves and exports them independently of OfficeIMO.
Run with the corpus directory and output manifest path.
"""
import argparse
import hashlib
import json
import math
from importlib.metadata import version
from xml.etree import ElementTree as ET
from zipfile import ZipFile
from numbers_parser import Document
from numbers_parser.generated.functionmap import FUNCTION_MAP

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('corpus')
parser.add_argument('output')
parser.add_argument('--profile', choices=['functions', 'fractional-random'], default='functions')
args = parser.parse_args()
from pathlib import Path
root = Path(args.corpus) / 'native-exports'
assert version('numbers-parser') == '4.19.0'
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
fixtures = []
cases = [('numbers-functions-v14.5', 'numbers-functions-intake.xlsx', 9, 4),
         ('numbers-function-edges-v14.5', 'numbers-function-edges-intake.xlsx', 17, 4)]
if args.profile == 'fractional-random':
    cases = [('numbers-fractional-bounds-v14.5', 'numbers-fractional-bounds-intake.xlsx', 16, 1)]
for stem, intake, count, column in cases:
    artifacts = [stem + '.numbers', stem + '.xlsx', intake]
    excel_cells = None
    if args.profile == 'fractional-random':
        excel = 'excel-fractional-bounds-v16.113.3.xlsx'
        artifacts += [excel, 'numbers-fractional-native-v14.5.png', 'excel-fractional-native-v16.113.3.png']
        with ZipFile(root / excel) as package:
            excel_cells = {c.get('r'): c for c in ET.fromstring(package.read('xl/worksheets/sheet1.xml')).findall('.//s:c', ns)}
    if (root / (stem + '.pdf')).exists():
        artifacts.append(stem + '.pdf')
    with ZipFile(root / intake) as package:
        intake_cells = {c.get('r'): c for c in ET.fromstring(package.read('xl/worksheets/sheet1.xml')).findall('.//s:c', ns)}
    with ZipFile(root / (stem + '.xlsx')) as package:
        export_cells = {c.get('r'): c for c in ET.fromstring(package.read('xl/worksheets/sheet1.xml')).findall('.//s:c', ns)}
    document = Document(str(root / (stem + '.numbers')))
    assert len(document.sheets) == 1 and len(document.sheets[0].tables) == 1
    table = document.sheets[0].tables[0]
    cells = []
    for row in range(1, count + 1):
        address = f'{chr(65 + column)}{row + 1}'
        source = table.cell(row, column)
        original = intake_cells[address]
        exported = export_cells[address]
        assert original.find('s:v', ns) is None
        assert source.formula == original.find('s:f', ns).text == exported.find('s:f', ns).text
        cache = exported.find('s:v', ns)
        value = float(cache.text) if cache is not None else None
        assert value == source.value
        nodes = document._model.formula_ast(table._table_id)[source._formula_id]
        functions = [{'index': n.AST_function_node_index, 'name': FUNCTION_MAP[n.AST_function_node_index],
                      'argumentCount': n.AST_function_node_numArgs} for n in nodes if n.AST_node_type == 16]
        cell = {'row': row + 1, 'column': column + 1, 'formula': source.formula,
                'nativeNumericCache': value, 'nativeXlsxHasCache': cache is not None,
                'functions': functions}
        if excel_cells is not None:
            excel_cell = excel_cells[address]
            assert source.formula == excel_cell.find('s:f', ns).text
            excel_value = excel_cell.find('s:v', ns).text
            cell['nativeExcelCache'] = excel_value if excel_cell.get('t') == 'e' else float(excel_value)
            lower, upper = (float(token) for token in source.formula[len('RANDBETWEEN('):-1].split(','))
            cell['localEvaluation'] = ('reversed-bounds' if lower > upper else
                                      'empty-integer-interval' if math.ceil(lower) > math.floor(upper) else
                                      'integer-interval')
        cells.append(cell)
    fixtures.append({'source': stem + '.numbers', 'sheet': document.sheets[0].name, 'table': table.name,
                     'rows': table.num_rows, 'columns': table.num_cols,
                     'artifacts': [{'path': path, 'sha256': hashlib.sha256((root / path).read_bytes()).hexdigest()}
                                   for path in artifacts], 'cells': cells})
manifest = {'producer': 'Apple Numbers 14.5 (7045.0.17), macOS 27.0.1 (26A434)',
            'provenance': 'OfficeIMO-authored cache-free OOXML intake built with Python standard-library ZIP/XML; imported, evaluated, saved and exported through native Numbers. No OfficeIMO writer or evaluator produced the oracles.',
            'license': 'MIT, OfficeIMO contributors', 'extractorVersion': 'numbers-parser 4.19.0',
            'qualification': 'Function IDs, argument counts, expressions and numeric caches match the independent native XLSX exports. Random caches are snapshots, not reproducible random sequences. Native error cells have no XLSX cache, so no native error-code equivalence is claimed. Fractional RANDBETWEEN returned zero in this producer import and remains unqualified for local evaluation. PDF is a native appearance reference only; full layout and dynamic array semantics are not qualified.',
            'fixtures': fixtures}
if args.profile == 'fractional-random':
    manifest['provenance'] = ('OfficeIMO-authored cache-free OOXML intake; imported, evaluated, saved and exported by Numbers. '
                              'A separate cache-free copy is evaluated and saved by Microsoft Excel 16.113.3. '
                              'Native caches come from those applications, not the OfficeIMO evaluator.')
    manifest['comparisonProducer'] = 'Microsoft Excel 16.113.3, macOS 27.0.1 (26A434)'
    manifest['qualification'] = ('Sixteen native expressions, IDs and caches agree with the Numbers XLSX export. '
                                 'Seven nonempty integer intervals agree with native Excel; empty integer intervals have '
                                 'different native results and retain producer caches without local evaluation. '
                                 'Reversed bounds have native Numbers errors without exported error caches; Excel has #NUM!. '
                                 'This evidence does not establish identical random sequences, full appearance, or other producer versions.')
Path(args.output).write_text(json.dumps(manifest, indent=2) + '\n')
print(f'Extracted {sum(len(f["cells"]) for f in fixtures)} native function cases')
