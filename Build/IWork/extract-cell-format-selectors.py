"""Qualify selected default/scalar format declarations in unchanged Numbers fixtures.

Requires opt-in numbers-parser 4.19.0; this does not qualify Apple exports or appearance.
"""
import argparse
import hashlib
import json
from importlib.metadata import version
from pathlib import Path
from numbers_parser import Document

SOURCES = {
    'cross-table-formulas': '9371c5b1d6ee4dfa17569097f064eba9c67f804d88b48638efbbeeb459d07dd4',
    'endpoint-formulas': '3deb8e3b868be60d7b8924336d2839fcd690db2b138e6c2170d6c22c4aca48dc',
    'single-cell-formulas': '4ce0593faa61bf159bde0752ade1fa202afcd142317ba16fdec8eb71a7149170',
    'issue-102-v15.1': '88a9fa7be095d03004478393a87a4a97602d7468f839d067ec9118c524c55176',
}
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('corpus', type=Path, help='The existing numbers-parser fixture directory.')
parser.add_argument('output', type=Path)
args = parser.parse_args()
if version('numbers-parser') != '4.19.0':
    raise RuntimeError('The extractor requires numbers-parser 4.19.0.')
packages = []
for name, digest in SOURCES.items():
    source = args.corpus / (name + '.numbers')
    if hashlib.sha256(source.read_bytes()).hexdigest() != digest:
        raise RuntimeError('The source differs from the pinned native fixture: ' + name)
    document = Document(str(source))
    model = document._model
    cases, seen = [], set()
    counts = {str(bit): 0 for bit in (15, 16, 17, 18)}
    for sheet in document.sheets:
        for table in sheet.tables:
            for row in range(table.num_rows):
                for column in range(table.num_cols):
                    buffer = model.storage_buffer(table._table_id, row, column)
                    if not buffer or len(buffer) < 12 or buffer[0] != 5:
                        continue
                    flags = int.from_bytes(buffer[8:12], 'little')
                    offset = 12
                    for bit in range(21):
                        if not flags & (1 << bit):
                            continue
                        size = 16 if bit == 0 else 8 if bit in (1, 2) else 4
                        if offset + size > len(buffer):
                            raise RuntimeError('Truncated selected native field.')
                        if bit in (15, 16, 17, 18):
                            counts[str(bit)] += 1
                            key = int.from_bytes(buffer[offset:offset + 4], 'little')
                            format_record = model.table_format(table._table_id, key)
                            default = (bit == 17 and format_record.format_type == 260
                                       or bit == 18 and format_record.format_type == 1)
                            default = default and len(format_record.ListFields()) == 1
                            signature = (bit, format_record.SerializeToString())
                            if signature not in seen:
                                seen.add(signature)
                                cases.append({'sheet': sheet.name, 'table': table.name,
                                              'row': row + 1, 'column': column + 1, 'selectorBit': bit,
                                              'cellType': buffer[1], 'formatKey': key,
                                              'formatType': format_record.format_type,
                                              'formatHex': signature[1].hex(), 'formatDeclaration': str(format_record),
                                              'isDefaultScalarFormat': default})
                        offset += size
    packages.append({'source': source.name, 'sourceSha256': digest,
                     'selectorCounts': counts, 'cases': cases})
manifest = {'extractorVersion': 'numbers-parser 4.19.0',
            'qualification': 'Independent selected format metadata in unchanged licensed native fixtures; no Apple export, rendering or recalculation qualification.',
            'packages': packages}
args.output.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + '\n')
