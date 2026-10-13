"""Independently check controlled Numbers integer-text single-cell references.

Opt-in evidence only: numbers-parser 4.19.0, no OfficeIMO assemblies, no
mutation of the native Numbers or XLSX inputs. Native Excel offsets both grids
by one row and column; OfficeIMO retains table-relative coordinates.
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


def exported_cells(path):
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    with ZipFile(path) as archive:
        names = [s.get('name') for s in ET.fromstring(archive.read('xl/workbook.xml')).findall('s:sheets/s:sheet', ns)]
        assert names == ['Inputs - Operands', 'Checks - Results']
        strings = [''.join(t.text or '' for t in si.findall('.//s:t', ns))
                   for si in ET.fromstring(archive.read('xl/sharedStrings.xml')).findall('s:si', ns)]
        cells = {}
        for index, name in enumerate(names, 1):
            selected = {}
            for cell in ET.fromstring(archive.read(f'xl/worksheets/sheet{index}.xml')).findall('.//s:c', ns):
                value = cell.findtext('s:v', namespaces=ns)
                selected[cell.get('r')] = {'type': cell.get('t'),
                    'value': strings[int(value)] if cell.get('t') == 's' else value,
                    'formula': cell.findtext('s:f', namespaces=ns)}
            cells[name] = selected
        return cells


def verify(source, exported, text):
    document = Document(source)
    assert [s.name for s in document.sheets] == ['Inputs', 'Checks']
    tables = [t for s in document.sheets for t in s.tables]
    assert [(t.name, t.num_rows, t.num_cols) for t in tables] == [('Operands', 7, 3), ('Results', 8, 3)]
    operands, results = tables
    cell = operands.cell(3, 1)
    assert type(cell).__name__ == 'TextCell' and not cell.is_formula and cell.value == text
    integer = int(text)
    assert str(integer) == text and 0 <= integer <= 999999999999999
    native = exported_cells(exported)
    assert native['Inputs - Operands']['C5']['type'] == 's'
    assert native['Inputs - Operands']['C5']['value'] == text
    identifier = str(uuid.UUID(document._model.objects[operands._table_id].table_id))
    specs = [(4, 2, 'GCD(Operands::$B$4,6)', math.gcd(integer, 6)),
             (4, 3, 'LCM(Operands::$B$4,6)', math.lcm(integer, 6)),
             (7, 2, 'GCD(Operands::$B$4)', integer),
             (7, 3, 'LCM(Operands::$B$4)', integer)]
    selected = []
    for row, column, formula, expected in specs:
        cell = results.cell(row - 1, column - 1)
        assert cell.is_formula and cell.formula == formula and cell.value == expected
        nodes = document._model.formula_ast(results._table_id)[cell._formula_id]
        assert [n.AST_node_type for n in nodes] == ([36, 17, 16] if row == 4 else [36, 16])
        words = nodes[0].AST_cross_table_reference_extra_info.table_id
        reference = str(uuid.UUID(bytes=b''.join(word.to_bytes(4, 'little') for word in
            (words.uuid_w0, words.uuid_w1, words.uuid_w2, words.uuid_w3))))
        assert reference == identifier
        address = chr(65 + column) + str(row + 1)
        native_cell = native['Checks - Results'][address]
        assert float(native_cell['value']) == expected
        assert native_cell['formula'] == formula.replace('Operands::$B$4', "'Inputs - Operands'!$C$5")
        selected.append({'row': row, 'column': column, 'formula': formula,
            'excelFormula': formula.replace('Operands::', "'Inputs'!"),
            'nativeValue': expected, 'nativeXlsxAddress': address,
            'referenceTableUuid': reference, 'nodeTypes': [n.AST_node_type for n in nodes]})
    # Controls remain outside the admitted conversion domain. A blank single
    # argument and a blank within a range do not share LCM coercion behavior.
    assert results.cell(1, 2).value == 6 and results.cell(2, 2).value == 0
    assert results.cell(5, 1).value == 1 and results.cell(5, 2).value == 6
    assert type(results.cell(4, 1)).__name__ == type(results.cell(4, 2)).__name__ == 'ErrorCell'
    assert sum(c.is_formula for table in tables for row in table.rows() for c in row) == 14
    return selected


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ('source', 'native_xlsx', 'changed_source', 'changed_xlsx', 'output'):
        parser.add_argument(name, type=Path)
    args = parser.parse_args()
    assert version('numbers-parser') == '4.19.0'
    before = verify(args.source, args.native_xlsx, '12')
    after = verify(args.changed_source, args.changed_xlsx, '18')
    assert all(a['formula'] == b['formula'] and a['referenceTableUuid'] == b['referenceTableUuid']
               for a, b in zip(before, after))
    result = {'schemaVersion': 1, 'producer': 'Numbers 14.5 (7045.0.17)',
        'decoder': 'numbers-parser 4.19.0', 'source': args.source.name,
        'nativeXlsx': args.native_xlsx.name, 'changedSource': args.changed_source.name,
        'changedXlsx': args.changed_xlsx.name, 'cells': before, 'changedCells': after,
        'qualifiedFormulaCells': 4, 'nativeFormulaCells': 14,
        'inputs': [{'path': p.name, 'bytes': p.stat().st_size, 'sha256': hashlib.sha256(p.read_bytes()).hexdigest()}
                   for p in (args.source, args.native_xlsx, args.changed_source, args.changed_xlsx)],
        'limits': ['Canonical nonnegative integer text in direct scalar cell references only.',
                   'Entered one-cell ranges normalize to native single-cell AST nodes.',
                   'No wider text ranges, blanks, Booleans, localized text, formula dependencies or appearance equivalence.']}
    args.output.write_text(json.dumps(result, indent=2) + '\n')
    print('PASS | four native integer-text references, typed operands and changed caches agree with native XLSX')
