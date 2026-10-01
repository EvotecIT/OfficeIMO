"""Extract native Numbers function expressions, IDs and XLSX result oracles.

Requires opt-in numbers-parser 4.19.0. The intake files contain no formula caches;
Apple Numbers imports, evaluates, saves and exports them independently of OfficeIMO.
Run with the corpus directory and output manifest path.
"""
import argparse
import hashlib
import json
from importlib.metadata import version
from xml.etree import ElementTree as ET
from zipfile import ZipFile
from numbers_parser import Document
from numbers_parser.generated.functionmap import FUNCTION_MAP

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('corpus')
parser.add_argument('output')
args = parser.parse_args()
from pathlib import Path
root = Path(args.corpus) / 'native-exports'
assert version('numbers-parser') == '4.19.0'
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
fixtures = []
for stem, intake, count in [('numbers-functions-v14.5', 'numbers-functions-intake.xlsx', 9),
                            ('numbers-function-edges-v14.5', 'numbers-function-edges-intake.xlsx', 17)]:
    artifacts = [stem + '.numbers', stem + '.xlsx', intake]
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
        address = f'E{row + 1}'
        source = table.cell(row, 4)
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
        cells.append({'row': row + 1, 'column': 5, 'formula': source.formula,
                      'nativeNumericCache': value, 'nativeXlsxHasCache': cache is not None,
                      'functions': functions})
    fixtures.append({'source': stem + '.numbers', 'sheet': document.sheets[0].name, 'table': table.name,
                     'rows': table.num_rows, 'columns': table.num_cols,
                     'artifacts': [{'path': path, 'sha256': hashlib.sha256((root / path).read_bytes()).hexdigest()}
                                   for path in artifacts], 'cells': cells})
manifest = {'producer': 'Apple Numbers 14.5 (7045.0.17), macOS 27.0.1 (26A434)',
            'provenance': 'OfficeIMO-authored cache-free OOXML intake built with Python standard-library ZIP/XML; imported, evaluated, saved and exported through native Numbers. No OfficeIMO writer or evaluator produced the oracles.',
            'license': 'MIT, OfficeIMO contributors', 'extractorVersion': 'numbers-parser 4.19.0',
            'qualification': 'Function IDs, argument counts, expressions and numeric caches match the independent native XLSX exports. Random caches are snapshots, not reproducible random sequences. Native error cells have no XLSX cache, so no native error-code equivalence is claimed. Fractional RANDBETWEEN returned zero in this producer import and remains unqualified for local evaluation. PDF is a native appearance reference only; full layout and dynamic array semantics are not qualified.',
            'fixtures': fixtures}
Path(args.output).write_text(json.dumps(manifest, indent=2) + '\n')
print(f'Extracted {sum(len(f["cells"]) for f in fixtures)} native function cases')
