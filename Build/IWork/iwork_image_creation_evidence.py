"""Shared independent image/resource checks for owned native creation fixtures."""
import hashlib

from iwork_protobuf_evidence import fields


def verify_images(graph, decoded, decode, package, data_references):
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
    images, resource_paths = [], set()
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
    return images, resource_paths
