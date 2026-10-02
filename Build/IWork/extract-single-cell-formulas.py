"""Qualify native single-cell and same-table endpoint references without rewriting the package.

Requires opt-in numbers-parser 4.19.0. Apple export and rendering equivalence remain open.
"""
import argparse
import hashlib
import json
import math
from decimal import Decimal, ROUND_UP, ROUND_DOWN, ROUND_CEILING, ROUND_FLOOR, localcontext
import plistlib
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile
from numbers_parser import Document
from numbers_parser.cell import DECIMAL128_BIAS
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
        ('test-all-formulas.numbers', [('Reference', 18), ('Reference', 19), ('Reference', 49), ('Reference', 52), ('Reference', 53), ('Reference', 14), ('Reference', 83), ('Reference', 84), ('Math', 97), ('Math', 153), ('Math', 154), ('Math', 141), ('Statistical', 37)],
         [('Reference', 16), ('Reference', 17), ('Reference', 48), ('Reference', 50), ('Reference', 51), ('Text', 9), ('Text', 10), ('Text', 21), ('Text', 43)]),
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
        name = FUNCTION_MAP[nodes[-1].AST_function_node_index]
        if name in ('ROWS', 'COLUMNS'):
            axis = 'row' if name == 'ROWS' else 'column'
            computed = abs(references[-1][axis] - references[0][axis]) + 1
        else:
            assert name in ('ROW', 'COLUMN')
            computed = references[0]['row' if name == 'ROW' else 'column']
    elif sheet_name == 'Math' and row == 97:
        assert cell.formula == 'PRODUCT(Data::A1:E1)'
        computed = math.prod(target.cell(0, column).value for column in range(5))
    elif sheet_name == 'Math' and row in (153, 154):
        assert cell.formula == ('SUMSQ(3,4,Data::A1)' if row == 153 else 'SUMSQ(3,4,Data::A1,Data::A13)')
        assert target.cell('A13').value is None
        computed = 3**2 + 4**2 + target.cell('A1').value**2
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
    if name == 'ROW':
        assert len(arguments) == 0
        computed = float(row)
    elif name in ('ROWS', 'COLUMNS'):
        assert len(arguments) == 1 and arguments[0].AST_node_type in (17, 19)
        computed = 1.0
    elif name == 'NOT':
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
def stored_decimal128(cell):
    # Inspect the unchanged v5 buffer exposed by the independent reader. Its
    # float unpacker multiplies by a binary 10**exponent, which can add a rounding
    # step (the stored 0.24 is reported as 0.24000000000000002). Decimal arithmetic
    # retains the producer's coefficient/exponent before one conversion to float.
    assert cell._buffer[0] == 5 and cell._flags & 1
    data = bytes(cell._buffer[12:28])
    assert len(data) == 16 and data[15] & 0x78 != 0x78
    exponent = (((data[15] & 0x7f) << 7) | (data[14] >> 1)) - DECIMAL128_BIAS
    coefficient = int.from_bytes(data[:14], 'little') + ((data[14] & 1) << 112)
    assert coefficient < 10**34
    if data[15] & 0x80:
        coefficient = -coefficient
    with localcontext() as context:
        context.prec = 34
        value = Decimal(coefficient).scaleb(exponent)
    return value, {'bytes': data.hex(), 'value': str(value)}

numeric_cases = []
if upstream_path == 'test-all-formulas.numbers':
    table = document.sheets['Math'].tables['Tests']
    for row in [3, 4, 5, 6, 7, 13, 27, 28, 29, 30, 31, 46, 47, 48, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 94, 95, 101, 102, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 125, 126, 127, 128, 129, 152, 164, 165, 166]:
        cell = table.cell(row - 1, 1)
        nodes = model.formula_ast(table._table_id)[cell._formula_id]
        stack, functions = [], []
        for node in nodes:
            if node.AST_node_type == 17:
                stack.append(node.AST_number_node_number)
            elif node.AST_node_type == 19:
                # The selected PRODUCT expression has an explicit numeric text literal.
                stack.append(float(node.AST_string_node_string))
            elif node.AST_node_type == 22:
                stack.append(None)  # An explicit omitted operand; not an absent node.
            elif node.AST_node_type == 13:
                stack.append(-stack.pop())
            elif node.AST_node_type == 2:
                right, left = stack.pop(), stack.pop()
                stack.append(left - right)
            elif node.AST_node_type == 5:
                right, left = stack.pop(), stack.pop()
                stack.append(left ** right)
            else:
                assert node.AST_node_type == 16
                count = node.AST_function_node_numArgs
                name = FUNCTION_MAP[node.AST_function_node_index]
                values = stack[-count:]
                del stack[-count:]
                if name in ('CEILING', 'FLOOR'):
                    assert count == 2
                    number, factor = [Decimal(str(v)) if v is not None else Decimal(0) for v in values]
                    assert number == 0 or factor == 0 or (number > 0) == (factor > 0)
                    if number == 0 or (name == 'CEILING' and factor == 0):
                        computed = 0.0
                    else:
                        assert factor != 0
                        rounding = ROUND_CEILING if name == 'CEILING' else ROUND_FLOOR
                        computed = float((number / factor).to_integral_value(rounding=rounding) * factor)
                elif name == 'PRODUCT':
                    computed = math.prod(values)
                elif name == 'SUMSQ':
                    computed = sum(value * value for value in values)
                elif name == 'RADIANS':
                    assert count == 1
                    computed = math.radians(values[0])
                elif name == 'INT':
                    assert count == 1
                    computed = math.floor(values[0])
                elif name == 'MOD':
                    assert count == 2 and values[1] != 0
                    computed = values[0] - values[1] * math.floor(values[0] / values[1])
                elif name == 'SIGN':
                    assert count == 1
                    computed = (values[0] > 0) - (values[0] < 0)
                elif name in ('TRUNC', 'ROUNDUP', 'ROUNDDOWN'):
                    assert (count in (1, 2) if name == 'TRUNC' else count == 2)
                    digits = int(values[1]) if count == 2 else 0
                    assert count == 1 or values[1] == digits
                    scale = Decimal(10) ** digits
                    number = Decimal(str(values[0])) * scale
                    rounding = ROUND_UP if name == 'ROUNDUP' else ROUND_DOWN
                    computed = float(number.to_integral_value(rounding=rounding) / scale)
                elif name == 'SQRT':
                    assert count == 1 and values[0] >= 0
                    computed = math.sqrt(values[0])
                elif name == 'EXP':
                    assert count == 1
                    computed = math.exp(values[0])
                elif name in ('LN', 'LOG', 'LOG10'):
                    assert (count in (1, 2) if name == 'LOG' else count == 1) and values[0] > 0
                    base = values[1] if count == 2 else 10
                    assert base > 0 and base != 1
                    computed = math.log(values[0]) if name == 'LN' else math.log10(values[0]) if name == 'LOG10' or base == 10 else math.log(values[0], base)
                else:
                    assert name == 'ABS' and count == 1
                    computed = abs(values[0])
                stack.append(computed)
                functions.append({'index': node.AST_function_node_index, 'name': name, 'argumentCount': count})
        assert len(stack) == 1
        transcendental = any(f['name'] in ('EXP', 'LN', 'LOG', 'LOG10', 'RADIANS') for f in functions)
        multiple = any(f['name'] in ('CEILING', 'FLOOR') for f in functions)
        # The producer stores rounded transcendental caches. Preserve those exact
        # caches and record the tolerance used only for independent computation.
        if transcendental:
            assert math.isclose(stack[0], cell.value, rel_tol=1e-14, abs_tol=1e-15)
        elif multiple:
            decimal_cache, stored_cache = stored_decimal128(cell)
            assert stack[0] == float(decimal_cache)
        else:
            assert stack[0] == cell.value
        case = {'sourceSheet': 'Math', 'sourceTable': table.name, 'row': row, 'column': 2,
                              'sourceFormula': cell.formula, 'cachedValue': cell.value,
                              'computedCurrentValue': stack[0], 'nodeTypes': [n.AST_node_type for n in nodes],
                              'functions': functions}
        if transcendental:
            case['computationTolerance'] = {'relative': 1e-14, 'absolute': 1e-15}
        elif multiple:
            case['providerCachedValue'] = cell.value
            case['storedDecimal128'] = stored_cache
            case['cachedValue'] = float(decimal_cache)
        numeric_cases.append(case)
with ZipFile(args.source) as package:
    builds = plistlib.loads(package.read('Metadata/BuildVersionHistory.plist'))
manifest = {'upstream': 'https://github.com/masaccio/numbers-parser',
            'revision': 'd3836ebda1110b5c13b8722642ca61111fe8e865',
            'upstreamPath': 'tests/data/' + upstream_path, 'sourceSha256': source_hash,
            'extractorVersion': 'numbers-parser 4.19.0', 'license': 'MIT, copyright Jon Connell',
            'buildVersionHistory': builds,
            'qualification': 'Native node-36 target identities and mixed coordinates, plus selected scalar function nodes; independent computations agree with numeric/text/Boolean caches. Scalar text cases use ASCII literals. Build metadata is retained in this manifest. No Apple export/render oracle.',
            'cases': cases, 'scalarFunctionCases': scalar_cases}
if numeric_cases:
    manifest['numericFunctionCases'] = numeric_cases
    manifest['qualification'] += ' Numeric INT/MOD/SQRT/SIGN/TRUNC/ROUNDUP/ROUNDDOWN expressions include unary negatives, subtraction, optional and signed digit arguments and nested ABS; independent numeric/decimal computations agree with native caches.'
    manifest['qualification'] += ' EXP/LN/LOG/LOG10 cases include optional/explicit bases, nesting and exponentiation. Their exact producer caches are retained; independent transcendental computations use the recorded relative/absolute tolerance, not exact binary equality.'
    manifest['qualification'] += ' Ten CEILING/FLOOR expressions qualify signed multiples, zero significance and native omitted node-22 operands. Independent decimal computations agree exactly with stored Decimal128 coefficient/exponent values. Raw cache bytes and the independent provider float (which can add a binary rounding step) are retained separately.'
    manifest['qualification'] += ' Five PRODUCT/RADIANS/SUMSQ expressions qualify scalar numeric operands and one explicit numeric text literal; Three additional reference cases qualify PRODUCT over A1:E1 and SUMSQ with numeric and blank cross-table cells. Unions and arrays are not covered.'
args.output.write_text(json.dumps(manifest, indent=2) + '\n')
print(f'Extracted {len(cases)} reference, {len(scalar_cases)} scalar and {len(numeric_cases)} numeric formulas from {source_hash}')
