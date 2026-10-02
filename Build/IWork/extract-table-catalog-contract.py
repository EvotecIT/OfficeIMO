"""Qualify table catalog aliases and required kinds with opt-in numbers-parser 4.19.0.

Reads the checked-in corpus without rewriting packages. Does not qualify selection or appearance.
"""
import argparse
import hashlib
import io
import json
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile
from numbers_parser.iwafile import IWACompressedChunk, get_archive_info_and_remainder
from numbers_parser.generated.TSTArchives_pb2 import TableModelArchive, TableDataList
from numbers_parser.generated.mapping import ID_NAME_MAP

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('corpus', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
assert version('numbers-parser') == '4.19.0'
aliases = sorted(identifier for identifier, kind in ID_NAME_MAP.items()
                 if getattr(getattr(kind, 'DESCRIPTOR', None), 'full_name', None) == 'TST.TableDataList')
assert aliases == [6005, 6201]


def entries(package):
    for name in package.namelist():
        if name.endswith('.iwa'):
            yield package.read(name)
        elif name == 'Index.zip':
            with ZipFile(io.BytesIO(package.read(name))) as nested:
                yield from entries(nested)


results = []
for source in sorted(args.corpus.rglob('*')):
    if source.suffix not in ['.numbers', '.pages', '.key']:
        continue
    records = {}
    catalogs = []
    with ZipFile(source) as package:
        for framed in entries(package):
            data = b''.join(IWACompressedChunk._decompress_all(framed))
            while data:
                header, payload = get_archive_info_and_remainder(data)
                offset = 0
                for i, info in enumerate(header.message_infos):
                    content = payload[offset:offset + info.length]
                    offset += info.length
                    if i == 0:
                        assert header.identifier not in records
                        records[header.identifier] = (info.type, content)
                data = payload[offset:]
    for identifier, (kind, content) in sorted(records.items()):
        if kind != 6001:
            continue
        model = TableModelArchive.FromString(content)
        for name, expected in [('stringTable', 1), ('styleTable', 4), ('formula_table', 3),
                               ('rich_text_table', 8), ('commentStorageTable', 10), ('format_table', 2)]:
            if not model.base_data_store.HasField(name):
                continue
            target = getattr(model.base_data_store, name).identifier
            target_type, payload = records[target]
            assert target_type in aliases
            catalog = TableDataList.FromString(payload)
            assert catalog.HasField('listType') and catalog.listType == expected
            catalogs.append({'model': identifier, 'field': name, 'target': target,
                             'targetType': target_type, 'hasListType': catalog.HasField('listType'),
                             'listType': catalog.listType, 'expected': expected, 'entries': len(catalog.entries)})
    results.append({'path': str(source.relative_to(args.corpus)),
                    'sha256': hashlib.sha256(source.read_bytes()).hexdigest(), 'catalogs': catalogs})
manifest = {'provider': 'numbers-parser', 'providerVersion': '4.19.0',
            'registry': 'https://github.com/masaccio/numbers-parser/blob/d3836ebda1110b5c13b8722642ca61111fe8e865/src/numbers_parser/generated/mapping.py',
            'revision': 'd3836ebda1110b5c13b8722642ca61111fe8e865', 'catalogTypeAliases': aliases,
            'qualification': 'Declared catalog type and required list kind only. Model records may be inactive; no selected-content, Apple export or appearance claim.',
            'sources': results}
args.output.write_text(json.dumps(manifest, indent=2) + '\n')
print(f'Qualified {sum(len(source["catalogs"]) for source in results)} catalog links in {len(results)} packages')
