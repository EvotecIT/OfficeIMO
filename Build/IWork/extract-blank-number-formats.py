"""Extract blank numeric format selections from unchanged independent Numbers files.

Requires numbers-parser 4.19.0. This qualifies source metadata, not Apple exports.
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
source_manifest = json.loads((args.corpus / 'cell-format-selectors.json').read_text())
packages = []
for source in source_manifest['packages']:
    path = args.corpus / source['source']
    if hashlib.sha256(path.read_bytes()).hexdigest() != source['sourceSha256']:
        raise RuntimeError('The source differs from the pinned fixture: ' + path.name)
    document = Document(str(path))
    cases = []
    for sheet in document.sheets:
        for table in sheet.tables:
            for row in range(table.num_rows):
                for column in range(table.num_cols):
                    buffer = document._model.storage_buffer(table._table_id, row, column)
                    if not buffer or len(buffer) < 12 or buffer[0] != 5 or buffer[1] != 0:
                        continue
                    flags = int.from_bytes(buffer[8:12], 'little')
                    if not flags & ((1 << 13) | (1 << 14)):
                        continue
                    bit = 14 if flags & (1 << 14) else 13
                    offset = 12 + sum(16 if i == 0 else 8 if i in (1, 2) else 4
                                      for i in range(bit) if flags & (1 << i))
                    if offset + 4 > len(buffer):
                        raise RuntimeError('Truncated selected format.')
                    key = int.from_bytes(buffer[offset:offset + 4], 'little')
                    record = document._model.table_format(table._table_id, key)
                    cases.append({'sheet': sheet.name, 'table': table.name,
                                  'row': row + 1, 'column': column + 1,
                                  'flags': flags, 'selectorBit': bit, 'formatKey': key,
                                  'formatType': record.format_type,
                                  'decimalPlaces': record.decimal_places,
                                  'thousandsSeparator': record.show_thousands_separator,
                                  'formatHex': record.SerializeToString().hex(),
                                  'hasOtherScalarSelector': bool(flags & sum(1 << i for i in range(15, 19)))})
    packages.append({'source': source['source'], 'sourceSha256': source['sourceSha256'], 'cells': cases})
args.output.write_text(json.dumps({'extractorVersion': 'numbers-parser 4.19.0',
    'qualification': 'Independent blank-cell format metadata; multiple scalar selector precedence, Apple export and appearance remain unqualified.',
    'packages': packages}, indent=2) + '\n')
