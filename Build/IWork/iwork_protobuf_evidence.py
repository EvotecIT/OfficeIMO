"""Wire-field inspection for pinned, opt-in native iWork evidence extractors."""
from google.protobuf.internal.decoder import _DecodeVarint
from numbers_parser.iwafile import IWACompressedChunk, get_archive_info_and_remainder


def fields(data):
    position = 0
    result = {}
    while position < len(data):
        tag, position = _DecodeVarint(data, position)
        field, wire = tag >> 3, tag & 7
        if wire == 0:
            value, position = _DecodeVarint(data, position)
        elif wire == 2:
            length, position = _DecodeVarint(data, position)
            value = data[position:position + length]
            assert len(value) == length
            position += length
        elif wire in (1, 5):
            length = 8 if wire == 1 else 4
            value = data[position:position + length]
            assert len(value) == length
            position += length
        else:
            raise ValueError(f'Unsupported wire kind {wire}')
        result.setdefault(field, []).append(value)
    return result


def creation_records(package):
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


def creation_references(message, name='TSP.Reference'):
    if message.DESCRIPTOR.full_name == name:
        return {message.identifier}
    result = set()
    for field, value in message.ListFields():
        if field.message_type is None:
            continue
        for nested in value if field.is_repeated else [value]:
            result.update(creation_references(nested, name))
    return result
