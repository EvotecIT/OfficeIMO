"""Decode native Pages body-image frames and resources using pinned opt-in schemas.

This examines independently saved source records; it does not qualify arbitrary
Pages content or execute OfficeIMO assemblies.
"""
import argparse
import hashlib
import json
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile

from google.protobuf import descriptor_pb2, descriptor_pool, message_factory
from iwork_protobuf_evidence import archive_messages

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('source', type=Path)
parser.add_argument('descriptor', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
assert version('numbers-parser') == '4.19.0'
pool = descriptor_pool.DescriptorPool()
for file in descriptor_pb2.FileDescriptorSet.FromString(args.descriptor.read_bytes()).file:
    pool.Add(file)


def decode(schema, payload):
    message = message_factory.GetMessageClass(pool.FindMessageTypeByName(schema))()
    message.ParseFromString(payload)
    assert not message.FindInitializationErrors(), (schema, message.FindInitializationErrors())
    return message


with ZipFile(args.source) as package:
    records = {}
    for _, header, info, payload in archive_messages(package, primary_only=True):
        key = (header.identifier, info.type)
        assert key not in records
        records[key] = payload
    document = decode('TP.DocumentArchive', next(payload for (_, kind), payload in records.items() if kind == 10000))
    body = decode('TSWP.StorageArchive', records[document.body_storage.identifier, 2001])
    metadata = decode('TSP.PackageMetadata', next(payload for (_, kind), payload in records.items() if kind == 11006))
    resources = {data.identifier: data for data in metadata.datas}
    images, attachments = [], []
    for entry in body.table_attachment.entries:
        attachment = decode('TSWP.DrawableAttachmentArchive', records[entry.object.identifier, 2003])
        assert attachment.h_offset_type == attachment.v_offset_type == 0
        assert attachment.h_offset == attachment.v_offset == 0
        assert body.text[0].encode('utf-16-le')[entry.character_index * 2:entry.character_index * 2 + 2] == '\ufffc'.encode('utf-16-le')
        identifier = attachment.drawable.identifier
        kinds = [kind for (record_id, kind) in records if record_id == identifier]
        assert len(kinds) == 1 and kinds[0] in (3005, 6000)
        attachments.append({'offset': entry.character_index, 'attachment': entry.object.identifier,
            'drawable': identifier, 'type': kinds[0]})
        if kinds[0] != 3005:
            continue
        image = decode('TSD.ImageArchive', records[identifier, 3005])
        assert not image.HasField('mask')
        data = resources[image.data.identifier]
        assert data.file_name and '/' not in data.file_name and '\\' not in data.file_name
        path = 'Data/' + data.file_name
        payload = package.read(path)
        assert hashlib.sha1(payload).digest() == data.digest
        images.append({'id': identifier, 'path': path, 'sha256': hashlib.sha256(payload).hexdigest(),
            'frame': [image.super.geometry.size.width, image.super.geometry.size.height],
            'position': [image.super.geometry.position.x, image.super.geometry.position.y],
            'rotation': image.super.geometry.angle,
            'pixelSize': [image.naturalSize.width, image.naturalSize.height],
            'description': image.super.accessibility_description})

result = {'source': args.source.name, 'sha256': hashlib.sha256(args.source.read_bytes()).hexdigest(),
    'bytes': args.source.stat().st_size, 'provider': 'numbers-parser', 'providerVersion': version('numbers-parser'),
    'descriptorSha256': hashlib.sha256(args.descriptor.read_bytes()).hexdigest(),
    'body': list(body.text), 'attachments': attachments, 'images': images,
    'scope': 'Body image attachment identity, UTF-16 positions, frames, descriptions and resource digests; wider native layout is not assessed.'}
args.output.write_text(json.dumps(result, indent=2, ensure_ascii=False) + '\n')
print(f'Verified {len(images)} native images and {len(attachments)} body attachments: {result["sha256"]}')
