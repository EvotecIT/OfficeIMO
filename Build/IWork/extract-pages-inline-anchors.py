"""Reproduce inline attachment identities and UTF-16 offsets from pinned native Pages fixtures.

Uses opt-in numbers-parser 4.19.0 schemas; it does not rewrite the packages or qualify Apple exports.
"""
import argparse
import hashlib
import json
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile
from numbers_parser.iwafile import IWACompressedChunk, get_archive_info_and_remainder
from numbers_parser.generated.TSWPArchives_pb2 import StorageArchive, DrawableAttachmentArchive

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('corpus', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
assert version('numbers-parser') == '4.19.0'
fixtures = {
    'iwork-converter/a.pages': '8481e3071c8ea1cc9543354bcd1ff66f79e6c2c8686096fa3ae60ab96fd49ebf',
    'picodocs/sample-v14.4.pages': '4714477138d0a4090fc2ee2ba2ebb6adcd0fb6ce20a28897a6247a8e17d1ddce',
}
sources = []
for name, expected_hash in fixtures.items():
    source = args.corpus / name
    source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    assert source_hash == expected_hash
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
    anchors = []
    for storage_id, (kind, content) in records.items():
        if kind != 2001:
            continue
        storage = StorageArchive.FromString(content)
        if not storage.HasField('table_attachment'):
            continue
        text = ''.join(storage.text).encode('utf-16-le')
        for entry in storage.table_attachment.entries:
            attachment_type, content = records[entry.object.identifier]
            assert attachment_type == 2003
            attachment = DrawableAttachmentArchive.FromString(content)
            assert attachment.h_offset_type == attachment.v_offset_type == 0
            assert attachment.h_offset == attachment.v_offset == 0
            position = entry.character_index
            assert text[position * 2:position * 2 + 2] == '\ufffc'.encode('utf-16-le')
            drawable_type = records[attachment.drawable.identifier][0]
            assert drawable_type in (3005, 6000, 6007)
            anchors.append({'storageIdentifier': storage_id, 'characterOffset': position,
                            'attachmentIdentifier': entry.object.identifier,
                            'drawableIdentifier': attachment.drawable.identifier,
                            'drawableType': drawable_type})
    sources.append({'path': name, 'sha256': source_hash, 'anchors': anchors})
assert [len(source['anchors']) for source in sources] == [1, 4]
manifest = {'provider': 'numbers-parser', 'providerVersion': '4.19.0',
            'offsetUnit': 'UTF-16 code units before marker removal', 'sources': sources,
            'limits': 'Native source identities and zero-offset placements only; Apple export, pagination and rendered appearance are not qualified.'}
args.output.write_text(json.dumps(manifest, indent=2) + '\n', encoding='utf-8')
