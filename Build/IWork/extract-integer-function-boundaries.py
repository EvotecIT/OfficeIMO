"""Record native integer-function results that differ from Excel's documented bounds.

Requires numbers-parser 4.19.0. Does not qualify conversion or native recalculation.
"""
import argparse
import hashlib
import json
import math
from importlib.metadata import version
from pathlib import Path
from numbers_parser import Document
from numbers_parser.generated.functionmap import FUNCTION_MAP

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('source', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
assert version('numbers-parser') == '4.19.0'
source_hash = hashlib.sha256(args.source.read_bytes()).hexdigest()
assert source_hash == '4ce0593faa61bf159bde0752ade1fa202afcd142317ba16fdec8eb71a7149170'
document = Document(str(args.source))
table = document.sheets['Math'].tables['Tests']
cases = []
for row, name in [(34, 'GCD'), (49, 'LCM')]:
    cell = table.cell(row - 1, 1)
    assert cell.formula == f'{name}(128,80,44,2^53)'
    nodes = document._model.formula_ast(table._table_id)[cell._formula_id]
    stack = []
    for node in nodes:
        if node.AST_node_type == 17:
            stack.append(node.AST_number_node_number)
        elif node.AST_node_type == 5:
            right, left = stack.pop(), stack.pop()
            stack.append(left ** right)
        else:
            assert node.AST_node_type == 16 and FUNCTION_MAP[node.AST_function_node_index] == name
            assert node.AST_function_node_numArgs == len(stack) == 4
    operands = [int(value) for value in stack]
    assert operands == [128, 80, 44, 2**53]
    exact = math.gcd(*operands) if name == 'GCD' else math.lcm(*operands)
    assert math.isclose(cell.value, float(exact), rel_tol=1e-14)
    cases.append({'row': row, 'column': 2, 'formula': cell.formula,
                  'functionIndex': nodes[-1].AST_function_node_index,
                  'operands': operands, 'exactIntegerResult': str(exact),
                  'nativeCachedValue': cell.value,
                  'excelDocumentedError': '#NUM!',
                  'excelContract': f'https://support.microsoft.com/en-us/excel/functions/{name.lower()}-function'})
manifest = {'source': 'numbers-parser/single-cell-formulas.numbers', 'sha256': source_hash,
            'provider': 'numbers-parser', 'providerVersion': '4.19.0',
            'sourceSheet': 'Math', 'sourceTable': 'Tests', 'cases': cases,
            'limits': 'Unchanged producer caches and independently computed integer results. Excel error is a documented contract, not a live Excel execution. No Apple export or native recalculation oracle.'}
args.output.write_text(json.dumps(manifest, indent=2) + '\n')
print(f'Extracted {len(cases)} native integer-function boundary cases')
