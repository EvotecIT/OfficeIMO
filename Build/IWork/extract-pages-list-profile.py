"""Extract native Pages body/list declarations using pinned independent schemas.

Opt-in evidence tooling only. Retains sparse paragraph boundaries and selected
style inheritance; does not calculate counters or qualify native appearance.
Pair this output with the source package and Apple PDF/DOCX exports.
"""
import argparse
import hashlib
import json
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile

from google.protobuf import descriptor_pb2, descriptor_pool, json_format, message_factory
from iwork_protobuf_evidence import archive_messages

SCHEMAS = {10000: 'TP.DocumentArchive', 2001: 'TSWP.StorageArchive',
           2021: 'TSWP.CharacterStyleArchive', 2022: 'TSWP.ParagraphStyleArchive',
           2023: 'TSWP.ListStyleArchive'}


def extract(source, descriptor):
    assert version('numbers-parser') == '4.19.0'
    pool = descriptor_pool.DescriptorPool()
    descriptor_bytes = descriptor.read_bytes()
    for file in descriptor_pb2.FileDescriptorSet.FromString(descriptor_bytes).file:
        pool.Add(file)
    records = {}
    with ZipFile(source) as package:
        for entry, header, info, payload in archive_messages(package, primary_only=True):
            assert header.identifier not in records
            message = None
            if info.type in SCHEMAS:
                message = message_factory.GetMessageClass(pool.FindMessageTypeByName(SCHEMAS[info.type]))()
                message.ParseFromString(payload)
                assert not message.FindInitializationErrors()
            records[header.identifier] = (entry, info.type, message)
    roots = [record[2] for record in records.values() if record[1] == 10000]
    assert len(roots) == 1
    body_id = roots[0].body_storage.identifier
    body = records[body_id][2]
    assert body.kind == 0
    selected = {body_id}
    for table in (body.table_para_style, body.table_char_style, body.table_list_style):
        selected.update(entry.object.identifier for entry in table.entries if entry.HasField('object'))
    pending = list(selected - {body_id})
    while pending:
        identifier = pending.pop()
        message = records[identifier][2]
        assert message is not None
        parent = message.super.parent.identifier if message.super.HasField('parent') else None
        if parent is not None and parent not in selected:
            selected.add(parent)
            pending.append(parent)
    return {
        'source': source.name, 'sha256': hashlib.sha256(source.read_bytes()).hexdigest(),
        'provider': 'numbers-parser', 'providerVersion': '4.19.0',
        'descriptorSha256': hashlib.sha256(descriptor_bytes).hexdigest(),
        'offsetUnit': 'UTF-16 code units', 'bodyIdentifier': body_id,
        'primaryRecordCount': len(records),
        'records': [{'identifier': identifier, 'entry': records[identifier][0],
                     'type': records[identifier][1],
                     'declaration': json_format.MessageToDict(records[identifier][2], preserving_proto_field_name=True)}
                    for identifier in sorted(selected)],
        'scope': 'Selected native body and style declarations. Counter values and appearance require paired Apple exports.'}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('--descriptor', required=True, type=Path)
    parser.add_argument('--out', required=True, type=Path)
    args = parser.parse_args()
    result = extract(args.source, args.descriptor)
    args.out.write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
    print(f"Extracted {len(result['records'])} selected declarations: {result['sha256']}")
