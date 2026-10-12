"""Check controlled Numbers cross-table integer calls and native recalculation.

Opt-in evidence only. Requires numbers-parser 4.19.0, uses no OfficeIMO
assemblies and never writes the supplied Numbers/XLSX packages.
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


def xlsx_cells(path):
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    with ZipFile(path) as archive:
        names=[sheet.get('name') for sheet in ET.fromstring(archive.read('xl/workbook.xml')).findall('s:sheets/s:sheet',ns)]
        assert names==['Budget - Data','Budget - Totals','Empty grid - Empty']
        return {names[index-1]:{c.get('r'): {'formula': c.findtext('s:f', namespaces=ns),
                              'value': c.findtext('s:v', namespaces=ns)}
                 for c in ET.fromstring(archive.read(f'xl/worksheets/sheet{index}.xml')).findall('.//s:c', ns)}
                for index in range(1, 4)}


def verify(source, exported, first_value):
    document = Document(source)
    assert [sheet.name for sheet in document.sheets] == ['Budget', 'Empty grid']
    objects=document._model.objects
    root=objects[objects.find_refs('DocumentArchive')[0]]
    by_record={table._table_id:table for sheet in document.sheets for table in sheet.tables}
    # numbers-parser's high-level collection can change order after a native
    # edit. Read declared TN sheet/drawable order directly from the archive.
    tables=[by_record[objects[drawable.identifier].tableModel.identifier] for sheet in root.sheets
            for drawable in objects[sheet.identifier].drawable_infos]
    by_name={table.name:table for table in tables}
    assert sorted((t.name, t.num_rows, t.num_cols) for t in tables) == [('Data', 4, 4), ('Empty', 2, 3), ('Totals', 2, 2)]
    identities = {str(uuid.UUID(document._model.objects[table._table_id].table_id)): (sheet.name, table.name)
                  for sheet in document.sheets for table in sheet.tables}
    inputs = [by_name['Data'].cell(row, 1) for row in (1, 2, 3)] + [by_name['Totals'].cell(0, 1)]
    assert all(type(cell).__name__ == 'NumberCell' and not cell.is_formula for cell in inputs)
    a, b, c, total = [math.trunc(cell.value) for cell in inputs]
    assert (a, b, c, total) == (first_value, 36, 60, 99)
    specs = [
        ('Data', 2, 4, 'GCD(Totals::B1,6)', math.gcd(total, 6)),
        ('Data', 3, 4, 'LCM(Totals::B1,6)', math.lcm(total, 6)),
        ('Data', 4, 4, '30+6', 36),
        ('Empty', 1, 1, 'GCD(Data::B2:B4)', math.gcd(a, b, c)),
        ('Empty', 1, 2, 'LCM(Data::B2:B4)', math.lcm(a, b, c)),
        ('Empty', 1, 3, 'GCD(Data::B2,Totals::B1)', math.gcd(a, total)),
        ('Empty', 2, 1, 'LCM(Data::$B$2,Totals::$B$1)', math.lcm(a, total)),
        ('Empty', 2, 2, 'GCD(Data::B2:B4,Totals::B1)', math.gcd(a, b, c, total)),
        ('Empty', 2, 3, 'LCM(Data::B2,Totals::B1,0)', math.lcm(a, total, 0)),
    ]
    native = xlsx_cells(exported)
    selected = []
    for table_name, row, column, formula, expected in specs:
        table = by_name[table_name];table_index=tables.index(table)
        cell = table.cell(row - 1, column - 1)
        assert cell.is_formula and cell.formula == formula and cell.value == expected
        # Apple offsets the small table to B2:D3. OfficeIMO keeps its table-relative A1:C2 grid.
        native_address = chr(64 + column + (table_name == 'Empty')) + str(row + (table_name == 'Empty'))
        native_sheet='Empty grid - Empty' if table_name=='Empty' else 'Budget - '+table_name
        exported_cell = native[native_sheet][native_address]
        assert float(exported_cell['value']) == expected
        excel_formula = formula.replace('Totals::', "'Budget - Totals'!").replace('Data::', "'Budget - Data'!")
        assert exported_cell['formula'] == excel_formula
        nodes = document._model.formula_ast(table._table_id)[cell._formula_id]
        references = []
        for node in nodes:
            if not node.HasField('AST_cross_table_reference_extra_info'):
                continue
            words = node.AST_cross_table_reference_extra_info.table_id
            identifier = str(uuid.UUID(bytes=b''.join(word.to_bytes(4, 'little') for word in
                (words.uuid_w0, words.uuid_w1, words.uuid_w2, words.uuid_w3))))
            assert identifier in identities
            sheet_name, table_name = identities[identifier]
            references.append({'uuid': identifier, 'sheet': sheet_name, 'table': table_name})
        assert len(references) == (0 if formula == '30+6' else 1 if formula in (
            'GCD(Totals::B1,6)', 'LCM(Totals::B1,6)', 'GCD(Data::B2:B4)', 'LCM(Data::B2:B4)') else 2)
        selected.append({'tableIndex': table_index, 'tableName':table.name, 'row': row, 'column': column,
            'formula': formula, 'excelFormula': excel_formula, 'nativeValue': cell.value,
            'independentIntegerValue': expected, 'nativeXlsxAddress': native_address,
            'nodeTypes': [node.AST_node_type for node in nodes], 'references': references})
    assert sum(cell.is_formula for table in tables for row in table.rows() for cell in row) == 9
    return selected,[table.name for table in tables]


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('native_xlsx', type=Path)
    parser.add_argument('changed_source', type=Path)
    parser.add_argument('changed_xlsx', type=Path)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    assert version('numbers-parser') == '4.19.0'
    baseline,source_tables = verify(args.source, args.native_xlsx, 30)
    changed,changed_tables = verify(args.changed_source, args.changed_xlsx, 24)
    assert source_tables==changed_tables
    assert all(before['references']==after['references'] and before['formula']==after['formula']
               for before,after in zip(baseline,changed))
    result = {'schemaVersion': 1, 'producer': 'Numbers 14.5 (7045.0.17)', 'decoder': 'numbers-parser 4.19.0',
        'source': args.source.name, 'nativeXlsx': args.native_xlsx.name,
        'changedSource': args.changed_source.name, 'changedXlsx': args.changed_xlsx.name,
        'formulaCells': 9, 'qualifiedCrossTableIntegerCalls': 8, 'cells': baseline, 'changedCells': changed,
        'sourceTables':source_tables,'changedSourceTables':changed_tables,
        'inputs': [{'path': path.name, 'bytes': path.stat().st_size,
                    'sha256': hashlib.sha256(path.read_bytes()).hexdigest()}
                   for path in (args.source, args.native_xlsx, args.changed_source, args.changed_xlsx)],
        'limits': ['Controlled numeric scalar references across two tables and two sheets in the recorded producer.',
                   'Native inputs, formula caches and typed XLSX exports agree with independent integer arithmetic.',
                   'Native XLSX offsets the small table by one row and column; OfficeIMO uses table-relative positions.',
                   'No broader coercion, formula dependency, whole-axis, exact appearance or complete PDF qualification.']}
    args.output.write_text(json.dumps(result, indent=2) + '\n')
    print('PASS | 8 native cross-table integer calls, arithmetic control and changed-operand caches agree with native XLSX')
