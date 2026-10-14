"""Decode controlled Numbers conditional aggregates and native Excel recalculation.

Requires isolated numbers-parser 4.19.0. Inputs remain unchanged. This opt-in
qualification tool is not a product runtime or ordinary build dependency.
"""
import argparse
import hashlib
import json
import math
import re
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
parser.add_argument('excel_edited_xlsx', type=Path)
parser.add_argument('output', type=Path)
parser.add_argument('--empty-average-input', type=Path, required=True)
parser.add_argument('--empty-average-xlsx', type=Path, required=True)
parser.add_argument('--typed-error-input', type=Path, required=True)
parser.add_argument('--typed-error-xlsx', type=Path, required=True)
parser.add_argument('--artifacts', nargs='*', type=Path, default=[])
args = parser.parse_args()
assert version('numbers-parser') == '4.19.0'
assert hashlib.sha256(args.source.read_bytes()).hexdigest() == 'a3ea9206083d1062d1d8e190d87350e7a60291aa2600254341bb433de9007bb1'
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def canonical_formula(formula):
    # Only function tokens in the fixed profile are normalized. Keep native raw
    # storage separately, including single-cell array declarations and literals.
    parts = re.split(r'("(?:[^"]|"")*")', formula)
    for index in range(0, len(parts), 2):
        parts[index] = re.sub(r'\b_xlfn\.(?=(?:SUMIFS|AVERAGEIFS|COUNTIFS|AVERAGEIF|IFERROR)\s*\()', '', parts[index])
    return ''.join(parts)


def read_xlsx(path):
    with ZipFile(path) as archive:
        strings = []
        if 'xl/sharedStrings.xml' in archive.namelist():
            strings = [''.join(si.itertext()) for si in ET.fromstring(archive.read('xl/sharedStrings.xml'))]
        result = {}
        shared = {}
        for cell in ET.fromstring(archive.read('xl/worksheets/sheet1.xml')).findall('.//s:c', ns):
            kind = cell.get('t', 'n')
            value = cell.findtext('s:v', namespaces=ns)
            if kind == 's':
                value = strings[int(value)]
            elif kind == 'inlineStr':
                value = ''.join(cell.find('s:is', ns).itertext())
            function = cell.find('s:f', ns)
            result[cell.get('r')] = {'kind': kind, 'formula': cell.findtext('s:f', namespaces=ns),
                                   'formulaType': None if function is None else function.get('t', 'normal'),
                                   'formulaRange': None if function is None else function.get('ref'), 'value': value}
            if function is not None and function.get('t') == 'shared':
                result[cell.get('r')]['rawFormula'] = function.text
                result[cell.get('r')]['sharedFormulaIndex'] = function.get('si')
                if function.text:
                    shared[function.get('si')] = (cell.get('r'), function.text, function.get('ref'))
        # Excel shares the three fixed dependent-handler columns. Qualify only
        # their same-row G reference; this is not a general formula translator.
        handlers = {'H': 'IFERROR(G{},99)', 'I': '_xlfn.IFNA(G{},99)', 'J': 'ISERROR(G{})'}
        for address, cell in result.items():
            if cell['formulaType'] != 'shared':
                continue
            anchor, formula, reference = shared[cell['sharedFormulaIndex']]
            match = re.fullmatch(r'([HIJ])(\d+)', anchor)
            assert match and formula == handlers[match[1]].format(match[2]), anchor
            bounds = re.fullmatch(r'([HIJ])(\d+):([HIJ])(\d+)', reference)
            actual = re.fullmatch(r'([HIJ])(\d+)', address)
            assert bounds and actual and actual[1] == bounds[1] == bounds[3] == match[1]
            assert int(bounds[2]) == int(match[2]) <= int(actual[2]) <= int(bounds[4])
            cell['formula'] = handlers[actual[1]].format(actual[2])
        return result


document = Document(str(args.source))
table = document.sheets[0].tables[0]
assert table.num_rows == 21 and table.num_cols == 7
native = read_xlsx(args.native_xlsx)
excel = read_xlsx(args.excel_xlsx)
edited = read_xlsx(args.excel_edited_xlsx)
assert float(native['C2']['value']) == float(excel['C2']['value']) == 10
assert float(edited['C2']['value']) == 40
cells = []
identities = set()
for row in range(1, 21):
    cell = table.cell(row, 6)
    address = f'G{row + 1}'
    assert cell.is_formula
    assert cell.formula == canonical_formula(native[address]['formula']) == canonical_formula(excel[address]['formula']) == canonical_formula(edited[address]['formula']), address
    assert native[address]['formulaType'] == 'array' and native[address]['formulaRange'] == address
    assert native[address]['kind'] == excel[address]['kind'] == edited[address]['kind'] == 'n', address
    assert math.isclose(cell.value, float(native[address]['value']), rel_tol=1e-12, abs_tol=1e-12), address
    assert math.isclose(cell.value, float(excel[address]['value']), rel_tol=1e-12, abs_tol=1e-12), address
    nodes = document._model.formula_ast(table._table_id)[cell._formula_id]
    functions = [{'index': n.AST_function_node_index, 'name': FUNCTION_MAP[n.AST_function_node_index],
                  'arguments': n.AST_function_node_numArgs} for n in nodes if n.AST_node_type == 16]
    identities.update((f['index'], f['name']) for f in functions)
    cells.append({'row': row + 1, 'column': 7, 'address': address,
                  'case': table.cell(row, 5).value, 'nativeType': type(cell).__name__,
                  'nativeValue': cell.value, 'formula': cell.formula, 'functions': functions,
                  'numbersXlsx': native[address], 'excelRecalculatedXlsx': excel[address],
                  'excelEditedXlsx': edited[address]})
assert identities == {(231, 'SUMIFS'), (232, 'AVERAGEIFS'), (233, 'COUNTIFS'), (234, 'AVERAGEIF'), (235, 'IFERROR')}
assert sum(c['formula'] is not None for c in native.values()) == len(cells) == 20
assert sum(c['formula'] is not None for c in excel.values()) == 20
assert sum(c['formula'] is not None for c in edited.values()) == 20
error_input = read_xlsx(args.empty_average_input)
error_reference = read_xlsx(args.empty_average_xlsx)
errors = []
for row in range(2, 6):
    for column, kind, value in [('E', 'e', '#DIV/0!'), ('F', 'n', '99'),
                                ('G', 'e', '#DIV/0!'), ('H', 'b', '1')]:
        address = f'{column}{row}'
        actual = error_reference[address]
        assert actual['formula'] == error_input[address]['formula'], address
        assert actual['kind'] == kind and actual['value'] == value, address
        assert error_input[address]['value'] is None, address
        errors.append({'address': address, 'row': row, 'column': ord(column) - ord('A') + 1,
                       **actual})
assert sum(c['formula'] is not None for c in error_reference.values()) == len(errors) == 16
typed_input = read_xlsx(args.typed_error_input)
typed_reference = read_xlsx(args.typed_error_xlsx)
typed_errors = []
for address, cell in typed_input.items():
    if cell['formula'] is None:
        continue
    actual = typed_reference[address]
    assert actual['formula'] == cell['formula'], address
    assert cell['value'] is None and actual['value'] is not None, address
    assert actual['kind'] in {'n', 'e', 'b'}, address
    assert actual['kind'] != 'e' or actual['value'] in {'#N/A', '#DIV/0!'}, address
    match = re.fullmatch(r'([A-Z]+)(\d+)', address)
    column = 0
    for letter in match[1]:
        column = column * 26 + ord(letter) - ord('A') + 1
    typed_errors.append({'address': address, 'row': int(match[2]), 'column': column, **actual})
assert len(typed_errors) == sum(c['formula'] is not None for c in typed_reference.values()) == 397
assert typed_reference['G2']['value'] == typed_reference['G4']['value'] == '#N/A'
assert typed_reference['G28']['value'] == '30' and typed_reference['G44']['value'] == '0'
assert typed_reference['G52']['value'] == '10' and typed_reference['G68']['value'] == '0'
assert [typed_reference[f'G{row}']['value'] for row in range(92, 100)] == ['30', '1', '30', '1', '0', '0', '0', '0']
inputs = [args.source, args.native_xlsx, args.excel_xlsx, args.excel_edited_xlsx,
          args.empty_average_input, args.empty_average_xlsx,
          args.typed_error_input, args.typed_error_xlsx, *args.artifacts]
assert len({p.name for p in inputs}) == len(inputs)
assert args.output.resolve() not in {p.resolve() for p in inputs}
manifest = {
    'source': args.source.name, 'nativeXlsx': args.native_xlsx.name,
    'excelXlsx': args.excel_xlsx.name, 'excelEditedXlsx': args.excel_edited_xlsx.name,
    'producer': 'Numbers 14.5 (7045.0.17)',
    'excelProducer': 'Microsoft Excel 16.113.3 (16.113.26092714)',
    'decoder': 'numbers-parser 4.19.0', 'rows': 21, 'columns': 7, 'formulaCells': 20,
    'edit': {'address': 'C2', 'row': 2, 'column': 3, 'original': 10, 'value': 40},
    'cells': cells,
    'emptyAverageInput': args.empty_average_input.name,
    'emptyAverageXlsx': args.empty_average_xlsx.name,
    'emptyAverageCells': errors,
    'typedErrorInput': args.typed_error_input.name,
    'typedErrorXlsx': args.typed_error_xlsx.name,
    'typedErrorCases': 98, 'typedErrorFormulaCells': 397,
    'typedErrorCells': typed_errors,
    'inputs': [{'path': p.name, 'bytes': p.stat().st_size, 'sha256': hashlib.sha256(p.read_bytes()).hexdigest()}
               for p in inputs],
    'limits': [
        'OfficeIMO-owned synthetic XLSX intake was imported, recalculated and saved natively in Numbers. The Numbers source and direct native XLSX export are independent producer references.',
        'Both fresh Excel workbooks came from saved OfficeIMO conversion, with Calculate Now followed by local Save As. The edited case changed C2 from 10 to 40 before native Excel calculation.',
        'The profile covers numeric aggregates, finite rectangular and single-cell references, literal and cell criteria, case-insensitive text, wildcards and escaped wildcards, zero and no matches, and IFERROR around an empty average.',
        'Native source PDF and receiving Excel screenshots are inspected for content and values. This does not qualify exact destination typography or print layout.',
        'Apple XLSX uses _xlfn aliases and single-cell array metadata. Comparisons normalize only the five qualified function-token aliases and retain raw native formula spelling separately.',
        'Native cell-padding omission still requires the explicit partial editable policy. Formula qualification does not establish complete document fidelity.',
        'The separate Excel intake has no cached formula results. Fresh native calculation qualifies typed #DIV/0! for empty AVERAGEIF/AVERAGEIFS, IFERROR fallback, IFNA propagation and ISERROR; IFNA uses the required _xlfn storage prefix.',
        'The typed-error intake has 98 conditional cases and 397 uncached formulas including five error operands. Native Excel qualifies selected and unmatched errors, mixed numeric/error operands, typed error criteria versus ordinary error text, not-equal and wildcard/ordering criteria, no-match MINIFS/MAXIFS, and dependent IFERROR/IFNA/ISERROR results. MINIFS/MAXIFS use explicit _xlfn storage spelling.',
        'Wider mixed-type coercion, Numbers native error mappings, unequal range shapes, whole axes, dynamic/named targets, localized criteria and other producer versions remain unqualified.'
    ]
}
args.output.write_text(json.dumps(manifest, indent=2) + '\n')
print('Decoded 20 conditional aggregates, both fresh Excel calculations, 16 empty-average and 397 typed-error results')
