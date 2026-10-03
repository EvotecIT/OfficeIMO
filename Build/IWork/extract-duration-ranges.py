"""Extract selected durations from hash-pinned native sources using numbers-parser 4.19.0.

This qualifies declarations, caches and independent display, not Apple export appearance.
"""
import argparse
import hashlib
import json
from importlib.metadata import version
from pathlib import Path
from numbers_parser import Document

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('corpus', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
if version('numbers-parser') != '4.19.0':
    raise RuntimeError('The extractor requires numbers-parser 4.19.0.')
sources = json.loads((args.corpus / 'cell-format-selectors.json').read_text())
packages = []
for source in sources['packages']:
    path = args.corpus / source['source']
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    if digest != source['sourceSha256']:
        raise RuntimeError('The source differs from the pinned fixture: ' + path.name)
    document = Document(str(path))
    cells = []
    for sheet in document.sheets:
        for table in sheet.tables:
            for row in range(table.num_rows):
                for column in range(table.num_cols):
                    buffer = document._model.storage_buffer(table._table_id, row, column)
                    if not buffer or len(buffer) < 12 or buffer[0] != 5:
                        continue
                    flags = int.from_bytes(buffer[8:12], 'little')
                    if not flags & (1 << 16):
                        continue
                    offset = 12 + sum(16 if bit == 0 else 8 if bit in (1, 2) else 4
                                      for bit in range(16) if flags & (1 << bit))
                    if offset + 4 > len(buffer):
                        raise RuntimeError('Truncated selected source duration format.')
                    key = int.from_bytes(buffer[offset:offset + 4], 'little')
                    record = document._model.table_format(table._table_id, key)
                    cell = table.cell(row, column)
                    cells.append({'sheet': sheet.name, 'table': table.name, 'row': row + 1, 'column': column + 1,
                                  'formatHex': record.SerializeToString().hex(),
                                  'settings': {field.name: value for field, value in record.ListFields()},
                                  'seconds': cell._double, 'independentDisplayText': cell.formatted_value})
    packages.append({'source': path.name, 'sourceSha256': digest, 'cells': cells})
manifest = {'extractorVersion': 'numbers-parser 4.19.0',
            'qualification': 'Selected abbreviated fixed day-only and week/day/hour declarations and saved caches. Whole-day display has independent parser evidence including negative values. Fractional-day rounding, mixed-unit XLSX representation, native exports, locale and appearance remain unqualified.',
            'packages': packages}
args.output.write_text(json.dumps(manifest, indent=2) + '\n')
