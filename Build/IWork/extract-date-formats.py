"""Extract selected date patterns, caches and display from unchanged Numbers files.

Requires opt-in numbers-parser 4.19.0. Native export differences are separate evidence.
"""
import argparse
import hashlib
import json
import locale
import platform
from datetime import datetime
import xml.etree.ElementTree as ET
import zipfile
from importlib.metadata import version
from pathlib import Path
from numbers_parser import Document
from numbers_parser.generated import TSKArchives_pb2

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('corpus', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
if version('numbers-parser') != '4.19.0':
    raise RuntimeError('The extractor requires numbers-parser 4.19.0.')
scalar = json.loads((args.corpus / 'numbers-parser/cell-format-selectors.json').read_text())
reference = json.loads((args.corpus / 'native-exports/numbers-formulas-v14.5.json').read_text())
sources = [(p['source'], p['sourceSha256']) for p in scalar['packages']]
sources.append((Path(reference['sourceFixture']).name, reference['sourceSha256']))
locale.setlocale(locale.LC_TIME, 'C')
packages = []
patterns = set()
for name, digest in sources:
    path = args.corpus / 'numbers-parser' / name
    if hashlib.sha256(path.read_bytes()).hexdigest() != digest:
        raise RuntimeError('The source differs from the pinned fixture: ' + name)
    document = Document(str(path))
    cases = []
    for sheet in document.sheets:
        for table in sheet.tables:
            for row in range(table.num_rows):
                for column in range(table.num_cols):
                    buffer = document._model.storage_buffer(table._table_id, row, column)
                    if not buffer or len(buffer) < 12 or buffer[0] != 5:
                        continue
                    flags = int.from_bytes(buffer[8:12], 'little')
                    if not flags & (1 << 15):
                        continue
                    offset = 12 + sum(16 if bit == 0 else 8 if bit in (1, 2) else 4
                                      for bit in range(15) if flags & (1 << bit))
                    if offset + 4 > len(buffer):
                        raise RuntimeError('Truncated selected source format.')
                    key = int.from_bytes(buffer[offset:offset + 4], 'little')
                    record = document._model.table_format(table._table_id, key)
                    settings = {f.name: value for f, value in record.ListFields()}
                    if record.format_type != 261 or not record.date_time_format:
                        raise RuntimeError('Expected a selected date format.')
                    cell = table.cell(row, column)
                    patterns.add(record.date_time_format)
                    cases.append({'sheet': sheet.name, 'table': table.name, 'row': row + 1, 'column': column + 1,
                                  'cellType': buffer[1], 'formatKey': key, 'formatHex': record.SerializeToString().hex(),
                                  'settings': settings, 'value': cell.value.isoformat(),
                                  'independentDisplayText': cell.formatted_value})
    packages.append({'source': 'numbers-parser/' + name, 'sourceSha256': digest, 'cells': cases})
native_path = args.corpus / 'native-exports/numbers-formulas-v14.5.xlsx'
native_hash = hashlib.sha256(native_path.read_bytes()).hexdigest()
if native_hash != next(a['sha256'] for a in reference['artifacts'] if a['path'] == native_path.name):
    raise RuntimeError('The native export differs from the pinned fixture.')
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
native_cells = []
with zipfile.ZipFile(native_path) as package:
    sheet = ET.fromstring(package.read('xl/worksheets/sheet1.xml'))
    styles = ET.fromstring(package.read('xl/styles.xml'))
    for address in ('C4', 'C5'):
        cell = sheet.find(f".//s:c[@r='{address}']", ns)
        format_id = list(styles.find('s:cellXfs', ns))[int(cell.attrib['s'])].attrib['numFmtId']
        code = styles.find(f"s:numFmts/s:numFmt[@numFmtId='{format_id}']", ns).attrib['formatCode']
        formula = cell.find('s:f', ns)
        native_cells.append({'cell': address, 'valueDays': float(cell.find('s:v', ns).text), 'formatCode': code,
                             'formula': None if formula is None else formula.text})
fields = {f.name: f.number for f in TSKArchives_pb2.FormatStructArchive.DESCRIPTOR.fields}
manifest = {'extractorVersion': 'numbers-parser 4.19.0',
            'displayEnvironment': {'operatingSystem': platform.system(), 'architecture': platform.machine(),
                                   'pythonVersion': platform.python_version(), 'timeLocale': locale.setlocale(locale.LC_TIME),
                                   'earlyYearDisplayProbe': datetime(9, 12, 31).strftime('%Y')},
            'qualification': 'Selected date pattern metadata, saved caches and independent invariant display in five unchanged native sources. The paired Numbers export localizes punctuation and refreshes volatile NOW caches; it does not prove source/native value or appearance equivalence. Locale, calendar, timezone, other patterns and suppression controls remain unqualified.',
            'formatFieldNumbers': {name: fields[name] for name in ('format_type', 'suppress_date_format', 'suppress_time_format', 'date_time_format')},
            'patterns': sorted(patterns), 'packages': packages,
            'nativeExport': {'path': 'native-exports/' + native_path.name, 'sha256': native_hash,
                             'producer': reference['producer'], 'cells': native_cells}}
args.output.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + '\n')
