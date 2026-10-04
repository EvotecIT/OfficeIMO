"""Wire-field inspection for pinned, opt-in native iWork evidence extractors."""
from google.protobuf.internal.decoder import _DecodeVarint


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
