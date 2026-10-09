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
from keynote_table_creation_evidence import verify_tables

SCHEMAS = {1: 'KN.DocumentArchive', 11006: 'TSP.PackageMetadata', 2: 'KN.ShowArchive',
    10: 'KN.ThemeArchive', 401: 'TSS.StylesheetArchive', 2021: 'TSWP.CharacterStyleArchive',
    2022: 'TSWP.ParagraphStyleArchive', 2025: 'TSWP.ShapeStyleArchive', 9: 'KN.SlideStyleArchive',
    2023: 'TSWP.ListStyleArchive', 2050: 'TSWP.TextStylePresetArchive', 5: 'KN.SlideArchive',
    4: 'KN.SlideNodeArchive', 2011: 'TSWP.ShapeInfoArchive', 2001: 'TSWP.StorageArchive',
    3005: 'TSD.ImageArchive', 3016: 'TSD.MediaStyleArchive',
    6000: 'TST.TableInfoArchive', 6001: 'TST.TableModelArchive', 6002: 'TST.Tile',
    6003: 'TST.TableStyleArchive', 6004: 'TST.CellStyleArchive', 6005: 'TST.TableDataList',
    6006: 'TST.HeaderStorageBucket', 6008: 'TST.TableStylePresetArchive', 6247: 'TST.TableStyleNetworkArchive'}


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


def references(message, name='TSP.Reference'):
    if message.DESCRIPTOR.full_name == name:
        return {message.identifier}
    result = set()
    for field, value in message.ListFields():
        if field.message_type is None:
            continue
        for nested in value if field.is_repeated else [value]:
            result.update(references(nested, name))
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
        # Keynote uses these identities for implicit document/support metadata.
        assert not ({71, 72} & graph.keys())
        objects, text_colors, decoded = [], [], {}
        data_references = {}
        for identifier, (kind, payload, header) in graph.items():
            message = decode(SCHEMAS[kind], payload)
            decoded[identifier] = message
            refs = references(message)
            if kind == 10:
                theme = fields(fields(payload)[1][0])
                for number, schema in [(100, 'TSD.ThemePresetsArchive'), (110, 'TSWP.ThemePresetsArchive'),
                                       (200, 'TST.ThemePresetsArchive')]:
                    for value in theme.get(number, []):
                        refs.update(references(decode(schema, value)))
            assert set(header.message_infos[0].object_references) == refs, identifier
            assert refs.issubset(graph.keys()), (identifier, refs - graph.keys())
            data_refs = references(message, 'TSP.DataReference')
            assert set(header.message_infos[0].data_references) == data_refs, identifier
            if data_refs:
                data_references[identifier] = data_refs
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
        text_runs = []
        for identifier, (kind, _, _) in graph.items():
            if kind != 2001:
                continue
            storage = decoded[identifier]
            text = storage.text if isinstance(storage.text, str) else ''.join(storage.text)
            utf16_length = len(text.encode('utf-16-le')) // 2
            boundaries, offset = {0}, 0
            for character in text:
                offset += len(character.encode('utf-16-le')) // 2
                boundaries.add(offset)
            ranges = list(storage.table_char_style.entries)
            assert ranges and ranges[0].character_index == 0
            positions = [entry.character_index for entry in ranges]
            assert positions == sorted(set(positions))
            assert all(position in boundaries and position < utf16_length for position in positions)
            encoded_text = text.encode('utf-16-le')
            for index, entry in enumerate(ranges):
                assert entry.HasField('object')
                style_id = entry.object.identifier
                assert graph[style_id][0] == 2021
                character = decoded[style_id].char_properties
                # Each rich range describes a complete formatting snapshot. Explicit
                # false values must survive so a later plain run clears decoration.
                if len(ranges) > 1:
                    assert all(character.HasField(name) for name in ('bold', 'italic', 'underline', 'strikethru'))
                end = positions[index + 1] if index + 1 < len(ranges) else utf16_length
                text_runs.append({'storage': identifier, 'start': entry.character_index, 'end': end,
                    'text': encoded_text[entry.character_index * 2:end * 2].decode('utf-16-le'),
                    'style': style_id, 'font': character.font_name, 'size': character.font_size,
                    'bold': character.bold, 'italic': character.italic,
                    'underline': int(character.underline), 'strikethrough': int(character.strikethru)})
        metadata = next(message for identifier, message in decoded.items() if graph[identifier][0] == 11006)
        datas = {data.identifier: data for data in metadata.datas}
        assert len(datas) == len(metadata.datas)
        referenced = set().union(*data_references.values()) if data_references else set()
        assert referenced == set(datas)
        expected_component_refs = {data: {identifier for identifier, refs in data_references.items() if data in refs}
            for data in referenced}
        component_refs = {}
        for component in metadata.components:
            for reference in component.data_references:
                assert reference.data_identifier not in component_refs
                assert all(object.count == 1 for object in reference.object_reference_list)
                component_refs[reference.data_identifier] = {object.object_identifier for object in reference.object_reference_list}
        assert component_refs == expected_component_refs
        images = []
        resource_paths = set()
        for identifier, message in decoded.items():
            if graph[identifier][0] != 3005:
                continue
            assert not message.HasField('mask') and message.HasField('data')
            data = datas[message.data.identifier]
            assert data.file_name and '/' not in data.file_name and '\\' not in data.file_name
            resource = 'Data/' + data.file_name
            resource_paths.add(resource)
            payload = package.read(resource)
            assert hashlib.sha1(payload).digest() == data.digest
            assert payload.startswith(b'\x89PNG\r\n\x1a\n') or payload.startswith(b'\xff\xd8')
            # The pinned schema omits DataAttributes extensions. Decode the native
            # image class marker explicitly, as confirmed in Apple-produced data.
            attributes = fields(data.attributes.SerializeToString())
            assert set(attributes) == {100} and len(attributes[100]) == 1
            image_attributes = decode('TSD.ImageDataAttributes', attributes[100][0])
            assert image_attributes.HasField('pixel_size')
            assert image_attributes.pixel_size == message.naturalSize
            assert image_attributes.HasField('should_be_interpreted_as_generic_if_untagged')
            assert not image_attributes.should_be_interpreted_as_generic_if_untagged
            images.append({'id': identifier, 'data_id': data.identifier, 'path': resource,
                'bytes': len(payload), 'sha256': hashlib.sha256(payload).hexdigest(),
                'position': [message.super.geometry.position.x, message.super.geometry.position.y],
                'frame': [message.super.geometry.size.width, message.super.geometry.size.height],
                'pixel_size': [message.naturalSize.width, message.naturalSize.height],
                'image_attributes_extension': 100,
                'description': message.super.accessibility_description})
        tables = verify_tables(graph, decoded)
        names = package.namelist()
        assert names == sorted(names) and len(names) == 5 + len(resource_paths)
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
        'character_ranges_complete': True, 'records': len(objects), 'objects': objects,
        'modern_text_colors': text_colors, 'native_text_runs': text_runs,
        'data_reference_graph_complete': True, 'image_resources': images, 'native_tables': tables}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('package', type=Path)
    parser.add_argument('descriptor', type=Path)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    assert version('numbers-parser') == '4.19.0'
    result = verify(args.package, args.descriptor)
    args.output.write_text(json.dumps(result, indent=2) + '\n')
    print({key: value for key, value in result.items() if key not in
           ('objects', 'modern_text_colors', 'native_text_runs', 'image_resources', 'native_tables')})


if __name__ == '__main__':
    main()
