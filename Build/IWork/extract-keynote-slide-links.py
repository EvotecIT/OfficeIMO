"""Extract native UUID and next-slide links from the three-slide Apple reference.

Opt-in evidence only. Requires numbers-parser 4.19.0 and the pinned Keynote
descriptor set used by verify-keynote-creation.py. Loads no OfficeIMO assemblies.
It assesses selected slide nodes and text/frame links, not every native feature.
"""
import argparse
import hashlib
import json
import struct
import uuid
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile

from google.protobuf import descriptor_pb2, descriptor_pool, message_factory
from numbers_parser.iwafile import IWACompressedChunk, get_archive_info_and_remainder

DESCRIPTOR_SHA256 = '2c46d6ee3eaf4d8419d0cc8329b74776f69517c1e14de0571e0434e29c4fc216'


def extract(path, descriptor):
    assert version('numbers-parser') == '4.19.0'
    descriptor_bytes = descriptor.read_bytes()
    assert hashlib.sha256(descriptor_bytes).hexdigest() == DESCRIPTOR_SHA256
    pool = descriptor_pool.DescriptorPool()
    for file in descriptor_pb2.FileDescriptorSet.FromString(descriptor_bytes).file:
        pool.Add(file)

    graph = {}
    with ZipFile(path) as package:
        for name in package.namelist():
            if not name.endswith('.iwa'):
                continue
            data = b''.join(IWACompressedChunk._decompress_all(package.read(name)))
            while data:
                header, payload = get_archive_info_and_remainder(data)
                offset = 0
                for info in header.message_infos:
                    key = (header.identifier, info.type)
                    content = payload[offset:offset + info.length]
                    assert len(content) == info.length
                    graph.setdefault(key, []).append(content)
                    offset += info.length
                data = payload[offset:]

    def decode(identifier, kind, name):
        candidates = graph[(identifier, kind)]
        assert len(candidates) == 1, (identifier, kind)
        message = message_factory.GetMessageClass(pool.FindMessageTypeByName(name))()
        message.ParseFromString(candidates[0])
        return message

    roots = [key[0] for key in graph if key[1] == 1]
    assert len(roots) == 1
    document = decode(roots[0], 1, 'KN.DocumentArchive')
    show = decode(document.show.identifier, 2, 'KN.ShowArchive')
    nodes = []
    links = []
    for position, reference in enumerate(show.slideTree.slides, 1):
        node = decode(reference.identifier, 4, 'KN.SlideNodeArchive')
        nodes.append({'position': position, 'node': reference.identifier, 'slide': node.slide.identifier})
        slide = decode(node.slide.identifier, 5, 'KN.SlideArchive')
        for drawable in slide.owned_drawables:
            shape = decode(drawable.identifier, 7, 'KN.PlaceholderArchive').super \
                if (drawable.identifier, 7) in graph else decode(drawable.identifier, 2011, 'TSWP.ShapeInfoArchive')
            storage_id = shape.owned_storage.identifier if shape.HasField('owned_storage') \
                else shape.deprecated_storage.identifier
            storage = decode(storage_id, 2001, 'TSWP.StorageArchive')
            text = ''.join(storage.text).rstrip('\n')
            if shape.super.super.hyperlink_url:
                links.append({'sourceSlide': position, 'kind': 'frame', 'record': drawable.identifier,
                    'text': text, 'target': shape.super.super.hyperlink_url})
            for attribute in storage.table_smartfield.entries:
                field = decode(attribute.object.identifier, 2032, 'TSWP.HyperlinkFieldArchive')
                links.append({'sourceSlide': position, 'kind': 'text', 'record': attribute.object.identifier,
                    'storage': storage_id, 'offset': attribute.character_index,
                    'text': text, 'target': field.url_ref})
    assert len(nodes) == 3 and len(links) == 3

    metadata_ids = [key[0] for key in graph if key[1] == 11006]
    assert len(metadata_ids) == 1
    metadata = decode(metadata_ids[0], 11006, 'TSP.PackageMetadata')
    selected = {node['node']: node for node in nodes}
    identities = []
    for component in metadata.components:
        for entry in component.object_uuid_map_entries:
            if entry.identifier in selected:
                identifier = str(uuid.UUID(bytes=struct.pack('<QQ', entry.uuid.lower, entry.uuid.upper)))
                identities.append({**selected[entry.identifier], 'uuid': identifier,
                    'lower': entry.uuid.lower, 'upper': entry.uuid.upper})
    assert len(identities) == 3 and len({item['uuid'] for item in identities}) == 3
    return {'source': path.name, 'sourceSha256': hashlib.sha256(path.read_bytes()).hexdigest(),
        'provider': 'numbers-parser', 'providerVersion': version('numbers-parser'),
        'descriptorSha256': DESCRIPTOR_SHA256, 'metadataRecord': metadata_ids[0],
        'uuidPath': 'PackageMetadata.3 / ComponentInfo.11 / ObjectUUIDMapEntry.2',
        'uuidEncoding': 'consecutive little-endian lower and upper words',
        'selectedNodes': sorted(identities, key=lambda item: item['position']), 'selectedLinks': links,
        'limits': 'Three selected slide nodes and their text/frame links; export targets and appearance are separate evidence.'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('package', type=Path)
    parser.add_argument('descriptor', type=Path)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    result = extract(args.package, args.descriptor)
    args.output.write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
    print('Extracted', len(result['selectedNodes']), 'slide UUIDs and', len(result['selectedLinks']), 'native links')


if __name__ == '__main__':
    main()
