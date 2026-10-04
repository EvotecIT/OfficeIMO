"""Qualify the fixed hour/minute duration format against an existing native export.

Requires opt-in numbers-parser 4.19.0. Source packages and Apple exports are unchanged.
"""
import argparse
import hashlib
import json
import xml.etree.ElementTree as ET
import zipfile
from importlib.metadata import version
from numbers_parser import Document
from numbers_parser.generated import TSKArchives_pb2

from pathlib import Path

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('corpus', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
if version('numbers-parser') != '4.19.0':
    raise RuntimeError('The extractor requires numbers-parser 4.19.0.')
source_path = args.corpus / 'numbers-parser/test-10-formulas.numbers'
reference_manifest = json.loads((args.corpus / 'native-exports/numbers-formulas-v14.5.json').read_text())
source_hash = hashlib.sha256(source_path.read_bytes()).hexdigest()
if source_hash != reference_manifest['sourceSha256']:
    raise RuntimeError('The source differs from the pinned native fixture.')
native_path = args.corpus / 'native-exports/numbers-formulas-v14.5.xlsx'
native_hash = hashlib.sha256(native_path.read_bytes()).hexdigest()
native = next(a for a in reference_manifest['artifacts'] if a['path'] == native_path.name)
if native_hash != native['sha256']:
    raise RuntimeError('The native export differs from the reference manifest.')
document = Document(str(source_path))
table = document.sheets[0].tables[0]
buffer = document._model.storage_buffer(table._table_id, 4, 2)
flags = int.from_bytes(buffer[8:12], 'little')
if buffer[0] != 5 or buffer[1] != 7 or not flags & (1 << 16):
    raise RuntimeError('Expected a selected modern duration cell.')
offset = 12 + sum(16 if bit == 0 else 8 if bit in (1, 2) else 4
                  for bit in range(16) if flags & (1 << bit))
key = int.from_bytes(buffer[offset:offset + 4], 'little')
format_record = document._model.table_format(table._table_id, key)
fields = {f.name: f.number for f in TSKArchives_pb2.FormatStructArchive.DESCRIPTOR.fields}
selected = {f.name: value for f, value in format_record.ListFields()}
if selected != {'format_type': 268, 'duration_style': 1, 'duration_unit_largest': 4,
                'duration_unit_smallest': 8, 'use_automatic_duration_units': False}:
    raise RuntimeError('Unexpected source duration format.')
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with zipfile.ZipFile(native_path) as package:
    cell = ET.fromstring(package.read('xl/worksheets/sheet1.xml')).find(".//s:c[@r='C6']", ns)
    styles = ET.fromstring(package.read('xl/styles.xml'))
    format_id = list(styles.find('s:cellXfs', ns))[int(cell.attrib['s'])].attrib['numFmtId']
    code = styles.find(f"s:numFmts/s:numFmt[@numFmtId='{format_id}']", ns).attrib['formatCode']
    serial = float(cell.find('s:v', ns).text)
    formula = cell.find('s:f', ns).text
if code != '[h]"h" m"m"' or serial * 86400 != table.cell(4, 2)._double:
    raise RuntimeError('The paired native duration differs from the source.')
manifest = {
    'extractorVersion': 'numbers-parser 4.19.0',
    'sourceFixture': reference_manifest['sourceFixture'], 'sourceSha256': source_hash,
    'nativeExport': 'native-exports/numbers-formulas-v14.5.xlsx', 'nativeExportSha256': native_hash,
    'producer': reference_manifest['producer'],
    'qualification': 'Selected fixed abbreviated hours/minutes and saved XLSX code/value in this paired native fixture. Other duration ranges/styles, automatic units, negative rounding, locale and appearance equivalence remain unqualified.',
    'formatFieldNumbers': {name: fields[name] for name in selected},
    'source': {'sheet': 'Sheet 1', 'table': 'Table 1', 'row': 5, 'column': 3, 'formatKey': key,
               'formatHex': format_record.SerializeToString().hex(), 'settings': selected,
               'seconds': table.cell(4, 2)._double, 'producerDisplayText': table.cell(4, 2).formatted_value},
    'native': {'sheet': 1, 'cell': 'C6', 'valueDays': serial, 'formatCode': code, 'formula': formula}
}
args.output.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + '\n')
