"""Extract native Pages table row sizing through pinned independent protobuf schemas.

Requires opt-in numbers-parser 4.19.0. Does not modify the package or qualify Apple exports.
"""
import argparse
import hashlib
import json
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile
from numbers_parser.iwafile import IWACompressedChunk, get_archive_info_and_remainder
from numbers_parser.generated.TSTArchives_pb2 import TableModelArchive, TableStyleArchive, HeaderStorageBucket

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('corpus', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
assert version('numbers-parser') == '4.19.0'
name = 'picodocs/sample-v14.4.pages'
source = args.corpus / name
source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
assert source_hash == '4714477138d0a4090fc2ee2ba2ebb6adcd0fb6ce20a28897a6247a8e17d1ddce'
records = {}
with ZipFile(source) as package:
    for entry in package.namelist():
        if not entry.endswith('.iwa'):
            continue
        data = b''.join(IWACompressedChunk._decompress_all(package.read(entry)))
        while data:
            header, payload = get_archive_info_and_remainder(data)
            offset = 0
            for payload_index, info in enumerate(header.message_infos):
                content = payload[offset:offset + info.length]
                offset += info.length
                if payload_index == 0:
                    assert header.identifier not in records
                    records[header.identifier] = (info.type, content)
            data = payload[offset:]

tables = []
for identifier, (kind, content) in sorted(records.items()):
    if kind != 6001:
        continue
    model = TableModelArchive.FromString(content)
    styles = []
    current = model.table_style.identifier
    while current:
        assert current not in [identifier for identifier, _ in styles]
        style_kind, payload = records[current]
        assert style_kind == 6003
        style = TableStyleArchive.FromString(payload)
        styles.append((current, style))
        current = style.super.parent.identifier if style.super.HasField('parent') else 0
    auto_resize = None
    for _, style in reversed(styles):
        if style.HasField('table_properties') and style.table_properties.HasField('auto_resize'):
            auto_resize = style.table_properties.auto_resize
    assert auto_resize is True
    rows = []
    for reference in model.base_data_store.rowHeaders.buckets:
        bucket_kind, payload = records[reference.identifier]
        assert bucket_kind == 6006
        for row in HeaderStorageBucket.FromString(payload).headers:
            assert row.hidingState == 0
            rows.append({'row': row.index + 1, 'heightPoints': row.size})
    tables.append({'modelIdentifier': identifier, 'name': model.table_name,
                   'styleIdentifiers': [identifier for identifier, _ in styles],
                   'autoResizeRows': auto_resize, 'rowHeights': sorted(rows, key=lambda row: row['row'])})
assert len(tables) == 3
manifest = {'provider': 'numbers-parser', 'providerVersion': '4.19.0',
            'source': {'path': name, 'sha256': source_hash}, 'tables': tables,
            'limits': 'Source row-height values and automatic-resize declarations only. Apple export, complete styling and pagination are not qualified.'}
args.output.write_text(json.dumps(manifest, indent=2) + '\n', encoding='utf-8')
