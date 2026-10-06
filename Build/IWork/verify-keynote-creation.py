"""Independently decode the OfficeIMO native Keynote creation profile.

Opt-in evidence tooling only: numbers-parser 4.19.0 and a descriptor set compiled
from the pinned iwork-converter proto schemas. No OfficeIMO assemblies are loaded.
This validates the graph and package profile, not native rendering or all Keynote.
"""
import argparse
import hashlib
import json
import plistlib
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile

from google.protobuf import descriptor_pb2, descriptor_pool, message_factory
from numbers_parser.iwafile import IWACompressedChunk, get_archive_info_and_remainder
from iwork_protobuf_evidence import fields

SCHEMAS = {1: 'KN.DocumentArchive', 11006: 'TSP.PackageMetadata', 2: 'KN.ShowArchive',
    10: 'KN.ThemeArchive', 401: 'TSS.StylesheetArchive', 2021: 'TSWP.CharacterStyleArchive',
    2022: 'TSWP.ParagraphStyleArchive', 2025: 'TSWP.ShapeStyleArchive', 9: 'KN.SlideStyleArchive',
    2023: 'TSWP.ListStyleArchive', 2050: 'TSWP.TextStylePresetArchive', 5: 'KN.SlideArchive',
    4: 'KN.SlideNodeArchive', 2011: 'TSWP.ShapeInfoArchive', 2001: 'TSWP.StorageArchive'}


def records(package):
    result = {}
    for name in package.namelist():
        if not name.endswith('.iwa'):
            continue
        data = b''.join(IWACompressedChunk._decompress_all(package.read(name)))
        while data:
            header, payload = get_archive_info_and_remainder(data)
            assert len(header.message_infos) == 1
            info = header.message_infos[0]
            assert header.identifier not in result
            result[header.identifier] = (info.type, payload[:info.length], header)
            data = payload[info.length:]
    return result


def references(message):
    if message.DESCRIPTOR.full_name == 'TSP.Reference':
        return {message.identifier}
    result = set()
    for field, value in message.ListFields():
        if field.message_type is None:
            continue
        for nested in value if field.is_repeated else [value]:
            result.update(references(nested))
    return result


def verify(path, descriptor):
    pool = descriptor_pool.DescriptorPool()
    descriptor_bytes = descriptor.read_bytes()
    for file in descriptor_pb2.FileDescriptorSet.FromString(descriptor_bytes).file:
        pool.Add(file)

    def decode(name, payload):
        message = message_factory.GetMessageClass(pool.FindMessageTypeByName(name))()
        message.ParseFromString(payload)
        assert not message.FindInitializationErrors(), (name, message.FindInitializationErrors())
        return message

    with ZipFile(path) as package:
        graph = records(package)
        objects, text_colors = [], []
        for identifier, (kind, payload, header) in graph.items():
            message = decode(SCHEMAS[kind], payload)
            refs = references(message)
            if kind == 10:
                theme = fields(fields(payload)[1][0])
                for number, schema in [(100, 'TSD.ThemePresetsArchive'), (110, 'TSWP.ThemePresetsArchive')]:
                    for value in theme.get(number, []):
                        refs.update(references(decode(schema, value)))
            assert set(header.message_infos[0].object_references) == refs, identifier
            assert refs.issubset(graph.keys()), (identifier, refs - graph.keys())
            objects.append({'id': identifier, 'type': kind, 'schema': SCHEMAS[kind], 'references': sorted(refs)})
            if kind in (2021, 2022):
                character = message.char_properties
                assert character.HasField('tsd_fill') and character.tsd_fill.HasField('color')
                assert character.font_color == character.tsd_fill.color
                color = character.tsd_fill.color
                text_colors.append({'id': identifier, 'font': character.font_name, 'size': character.font_size,
                    'rgb': [color.r, color.g, color.b]})
            if kind == 5:
                assert message.HasField('name') != message.HasField('template_slide')
            if kind == 2001:
                assert all(fields(payload).get(number) for number in (6, 14, 24))
        names = package.namelist()
        assert names == sorted(names) and len(names) == 5
        assert all(item.compress_type == 0 and item.date_time == (1980, 1, 1, 0, 0, 0)
                   for item in package.infolist())
        properties = plistlib.loads(package.read('Metadata/Properties.plist'))
        identifier = package.read('Metadata/DocumentIdentifier').decode('ascii')
        assert properties['documentUUID'] == identifier
        plistlib.loads(package.read('Metadata/BuildVersionHistory.plist'))
    return {'source': str(path), 'sha256': hashlib.sha256(path.read_bytes()).hexdigest(),
        'bytes': path.stat().st_size, 'descriptor_sha256': hashlib.sha256(descriptor_bytes).hexdigest(),
        'numbers_parser_version': version('numbers-parser'), 'required_fields_complete': True,
        'reference_graph_complete': True, 'deterministic_zip_profile': True,
        'records': len(objects), 'objects': objects, 'modern_text_colors': text_colors}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('package', type=Path)
    parser.add_argument('descriptor', type=Path)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    assert version('numbers-parser') == '4.19.0'
    result = verify(args.package, args.descriptor)
    args.output.write_text(json.dumps(result, indent=2) + '\n')
    print({key: value for key, value in result.items() if key not in ('objects', 'modern_text_colors')})


if __name__ == '__main__':
    main()
