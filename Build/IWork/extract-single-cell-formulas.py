"""Qualify native single-cell and same-table endpoint references without rewriting the package.

Requires opt-in numbers-parser 4.19.0. Apple export and rendering equivalence remain open.
"""
import argparse
import hashlib
import json
import plistlib
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile
from numbers_parser import Document
from numbers_parser.generated.functionmap import FUNCTION_MAP
from numbers_parser.numbers_uuid import NumbersUUID
from numbers_parser.xrefs import xl_rowcol_to_cell

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('source', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
if version('numbers-parser') != '4.19.0':
    raise RuntimeError('The extractor requires numbers-parser 4.19.0.')
source_hash = hashlib.sha256(args.source.read_bytes()).hexdigest()
qualified = {
    '4ce0593faa61bf159bde0752ade1fa202afcd142317ba16fdec8eb71a7149170':
        ('test-all-formulas.numbers', [('Reference', 14), ('Reference', 83), ('Reference', 84), ('Math', 141), ('Statistical', 37)],
         [('Text', 9), ('Text', 10), ('Text', 21), ('Text', 43)]),
    '3deb8e3b868be60d7b8924336d2839fcd690db2b138e6c2170d6c22c4aca48dc':
        ('test-extra-formulas.numbers', [('Formulas', 29), ('Formulas', 79), ('Formulas', 102)],
         [('Formulas', 95), ('Formulas', 145)])
}
if source_hash not in qualified:
    raise RuntimeError('The source differs from the pinned independent fixture.')
document = Document(str(args.source))
model = document._model
targets = {t._table_id: (s, t) for s in document.sheets for t in s.tables}
cases = []
upstream_path, selections, scalar_selections = qualified[source_hash]
for sheet_name, row in selections:
    sheet = document.sheets[sheet_name]
    table = sheet.tables['Tests']
    cell = table.cell(row - 1, 1)
    nodes = model.formula_ast(table._table_id)[cell._formula_id]
    references = []
    target_ids = set()
    for node in nodes:
        if not node.HasField('AST_cross_table_reference_extra_info'):
            continue
        assert node.AST_node_type == 36 and node.HasField('AST_row') and node.HasField('AST_column')
        ref = model.node_to_ref(table._table_id, row - 1, 1, node)
        target_ids.add(ref.to_table_id)
        references.append({'address': xl_rowcol_to_cell(ref.row_start, ref.col_start,
                            row_abs=ref.row_start_is_abs, col_abs=ref.col_start_is_abs),
                           'row': ref.row_start + 1, 'column': ref.col_start + 1,
                           'rowAbsolute': ref.row_start_is_abs, 'columnAbsolute': ref.col_start_is_abs,
                           'targetUuid': NumbersUUID(node.AST_cross_table_reference_extra_info.table_id).hex})
    assert len(target_ids) == 1
    target_sheet, target = targets[target_ids.pop()]
    assert target.name == 'Data'
    if len(nodes) == 1:
        computed = target.cell(references[0]['row'] - 1, references[0]['column'] - 1).value
    elif sheet_name == 'Reference':
        computed = references[0]['column']
    elif sheet_name == 'Math':
        threshold = target.cell('C2').value
        computed = sum(target.cell(r, 1).value for r in range(1, 5) if target.cell(r, 0).value > threshold)
    elif sheet_name == 'Formulas':
        if row == 29:
            computed = sum(target.cell(r, 2).value is None for r in range(1, 18))
        elif row == 79:
            computed = max(target.cell(r, 2).value for r in range(1, 6))
        else:
            computed = target.cell('C3').value
    else:
        criterion = target.cell('A4').value.casefold()
        computed = sum(target.cell(r, 0).value.casefold() == criterion for r in range(1, 5))
    assert computed == cell.value, (sheet_name, row, computed, cell.value)
    cases.append({'sourceSheet': sheet_name, 'sourceTable': table.name, 'row': row, 'column': 2,
                  'sourceFormula': cell.formula, 'cachedValue': cell.value,
                  'computedCurrentValue': computed, 'targetSheet': target_sheet.name,
                  'targetTable': target.name, 'references': references})
scalar_cases = []
for sheet_name, row in scalar_selections:
    table = document.sheets[sheet_name].tables['Tests']
    cell = table.cell(row - 1, 1)
    nodes = model.formula_ast(table._table_id)[cell._formula_id]
    function = nodes[-1]
    assert function.AST_node_type == 16
    name = FUNCTION_MAP[function.AST_function_node_index]
    arguments = nodes[:-1]
    assert function.AST_function_node_numArgs == len(arguments)
    if name == 'NOT':
        assert len(arguments) == 1 and arguments[0].AST_node_type == 17
        computed = arguments[0].AST_number_node_number == 0
    else:
        assert all(n.AST_node_type == 19 for n in arguments)
        values = [n.AST_string_node_string for n in arguments]
        # These fixtures contain ASCII arguments; they do not qualify locale/Unicode casing.
        assert all(value.isascii() for value in values)
        if name == 'EXACT':
            assert len(values) == 2
            computed = values[0] == values[1]
        elif name == 'LOWER':
            assert len(values) == 1
            computed = values[0].lower()
        elif name == 'UPPER':
            assert len(values) == 1
            computed = values[0].upper()
        else:
            assert name == 'TRIM' and len(values) == 1
            computed = ' '.join(part for part in values[0].split(' ') if part)
    assert type(computed) is type(cell.value) and computed == cell.value
    scalar_cases.append({'sourceSheet': sheet_name, 'sourceTable': table.name, 'row': row, 'column': 2,
                         'sourceFormula': cell.formula, 'cachedValue': cell.value,
                         'computedCurrentValue': computed,
                         'functionIndex': function.AST_function_node_index,
                         'functionName': name, 'argumentCount': len(arguments)})
with ZipFile(args.source) as package:
    builds = plistlib.loads(package.read('Metadata/BuildVersionHistory.plist'))
manifest = {'upstream': 'https://github.com/masaccio/numbers-parser',
            'revision': 'd3836ebda1110b5c13b8722642ca61111fe8e865',
            'upstreamPath': 'tests/data/' + upstream_path, 'sourceSha256': source_hash,
            'extractorVersion': 'numbers-parser 4.19.0', 'license': 'MIT, copyright Jon Connell',
            'buildVersionHistory': builds,
            'qualification': 'Native node-36 target identities and mixed coordinates, plus selected scalar function nodes; independent computations agree with numeric/text/Boolean caches. Scalar text cases use ASCII literals. Build metadata is retained in this manifest. No Apple export/render oracle.',
            'cases': cases, 'scalarFunctionCases': scalar_cases}
args.output.write_text(json.dumps(manifest, indent=2) + '\n')
print(f'Extracted {len(cases)} reference and {len(scalar_cases)} scalar formulas from {source_hash}')
