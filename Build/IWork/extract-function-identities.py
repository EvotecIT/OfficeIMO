"""Reproduce function-ID evidence with opt-in numbers-parser 4.19.0.

The selected argument counts exercise reconstruction, not evaluation or Apple export fidelity.
"""
import argparse
import hashlib
import json
from importlib.metadata import version
from pathlib import Path
from numbers_parser import Document
from numbers_parser.generated.functionmap import FUNCTION_MAP
from numbers_parser.xrefs import xl_rowcol_to_cell

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('source', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
assert version('numbers-parser') == '4.19.0'
source_hash = hashlib.sha256(args.source.read_bytes()).hexdigest()
assert source_hash == '9371c5b1d6ee4dfa17569097f064eba9c67f804d88b48638efbbeeb459d07dd4'
# Counts follow the public function signatures; names/IDs come from the independent provider.
samples = {'ABS': 1, 'AND': 1, 'AVERAGE': 1, 'COLUMN': 0, 'COUNT': 1,
           'COUNTA': 1, 'COUNTBLANK': 1, 'COUNTIF': 2, 'DATE': 3, 'DAY': 1,
           'FALSE': 0, 'FIND': 2, 'HOUR': 1, 'HYPERLINK': 2, 'IF': 3,
           'INDEX': 2, 'LEFT': 1, 'LEN': 1, 'MAX': 1, 'MEDIAN': 1,
           'MID': 3, 'MIN': 1, 'MINA': 1, 'MINUTE': 1, 'NOW': 0, 'OR': 2, 'PI': 0,
           'POWER': 2, 'RIGHT': 1, 'ROUND': 2, 'SECOND': 1, 'SUM': 1, 'SUMIF': 2, 'TEXTJOIN': 3}
identities = [{'index': index, 'name': name, 'sampleArgumentCount': samples[name]}
              for index, name in FUNCTION_MAP.items() if name in samples]
assert len(identities) == len(samples)
document = Document(str(args.source))
native_cases = []
for sheet_name, table_name, row, column in [('Main Sheet', 'Formula Tests', 16, 0),
                                          ('Powers Sheet', 'Powers of Two', 0, 0)]:
    table = document.sheets[sheet_name].tables[table_name]
    cell = table.cell(row, column)
    nodes = document._model.formula_ast(table._table_id)[cell._formula_id]
    native_cases.append({'sourceSheet': sheet_name, 'sourceTable': table_name,
                         'row': row + 1, 'column': column + 1, 'formula': cell.formula,
                         'cachedValue': cell.value,
                         'functions': [{'index': n.AST_function_node_index,
                                        'name': FUNCTION_MAP[n.AST_function_node_index],
                                        'argumentCount': n.AST_function_node_numArgs}
                                       for n in nodes if n.AST_node_type == 16]})
table = document.sheets['Main Sheet'].tables['Formula Tests']
cell = table.cell(15, 0)
nodes = document._model.formula_ast(table._table_id)[cell._formula_id]
assert len(nodes) == 4 and nodes[-1].AST_function_node_index == 328
assert nodes[-1].AST_function_node_numArgs == 3
reference = document._model.node_to_ref(table._table_id, 15, 0, nodes[2])
target_sheet = next(s for s in document.sheets if any(t._table_id == reference.to_table_id for t in s.tables))
target = next(t for t in target_sheet.tables if t._table_id == reference.to_table_id)
first = xl_rowcol_to_cell(reference.row_start, reference.col_start, row_abs=reference.row_start_is_abs, col_abs=reference.col_start_is_abs)
last = xl_rowcol_to_cell(reference.row_end, reference.col_end, row_abs=reference.row_end_is_abs, col_abs=reference.col_end_is_abs)
assert cell.formula == 'TEXTJOIN(",",FALSE,Animal Table::B1:D1)'
computed = ','.join(str(target.cell(r, c).value) for r in range(reference.row_start, reference.row_end + 1)
                    for c in range(reference.col_start, reference.col_end + 1))
assert computed == cell.value
text_join = {'sourceSheet': 'Main Sheet', 'sourceTable': 'Formula Tests', 'row': 16, 'column': 1,
             'sourceFormula': cell.formula, 'functionIndex': 328, 'argumentCount': 3,
             'targetSheet': target_sheet.name, 'targetTable': target.name,
             'address': first + ':' + last, 'cachedValue': cell.value,
             'computedCurrentRangeValue': computed}
manifest = {'upstream': 'https://github.com/masaccio/numbers-parser',
            'revision': '1c6c5c3d2e29a9abb601596678089f0a6c85d64c',
            'upstreamPath': 'tests/data/create-formulas.numbers',
            'sourceSha256': source_hash, 'extractorVersion': 'numbers-parser 4.19.0',
            'license': 'MIT, copyright Jon Connell',
            'qualification': 'Independent function identities plus three unmodified native expressions and caches; synthetic argument samples do not qualify evaluation. No Apple export, recalculation or appearance oracle.',
            'identities': identities, 'nativeCases': native_cases, 'textJoinCase': text_join,
            'unqualifiedIdentities': [{'index': index, 'name': FUNCTION_MAP.get(index)}
                                     for index in [101, 112, 119, 169]]}
args.output.write_text(json.dumps(manifest, indent=2) + '\n')
print(f'Extracted {len(identities)} identities and {len(native_cases) + 1} native expressions')
