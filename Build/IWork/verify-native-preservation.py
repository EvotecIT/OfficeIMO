"""Opt-in independent entry, IWA envelope and image-frame evidence.

Uses numbers-parser's Snappy decoder and pinned protobuf descriptors, never
OfficeIMO assemblies. File-entry payloads are compared independently of the ZIP
carrier. All raw IWA bytes must match except the declared position/size floats.
"""
import argparse
import hashlib
import json
import struct
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile

from google.protobuf import descriptor_pb2, descriptor_pool, message_factory
from google.protobuf.internal.decoder import _DecodeVarint
from numbers_parser.iwafile import IWACompressedChunk, get_archive_info_and_remainder


def geometry_spans(data, include_size=False):
    """Locate the requested fixed32 leaves independently in the raw message."""
    spans = []
    paths = ((1, 1, 1, 1), (1, 1, 1, 2))
    if include_size:
        paths += ((1, 1, 2, 1), (1, 1, 2, 2))
    for path in paths:
        start, end = 0, len(data)
        for level, selected_field in enumerate(path):
            found, offset = [], start
            while offset < end:
                tag, offset = _DecodeVarint(data, offset)
                field, wire, value_start = tag >> 3, tag & 7, offset
                if wire == 0:
                    _, offset = _DecodeVarint(data, offset)
                elif wire == 2:
                    length, value_start = _DecodeVarint(data, offset)
                    offset = value_start + length
                elif wire in (1, 5):
                    offset += 8 if wire == 1 else 4
                else:
                    raise ValueError(f'Unsupported wire kind {wire}')
                assert offset <= end
                if field == selected_field:
                    found.append((value_start, offset, wire))
            assert len(found) == 1
            start, end, wire = found[0]
            assert wire == (5 if level == len(path) - 1 else 2)
        spans.append((start, end))
    return spans


def archives(package):
    result = {}
    for name in package.namelist():
        if not name.endswith('.iwa'):
            continue
        raw = b''.join(IWACompressedChunk._decompress_all(package.read(name)))
        offset, rows = 0, []
        while offset < len(raw):
            length, start = _DecodeVarint(raw, offset)
            header, _ = get_archive_info_and_remainder(raw[offset:])
            payload_offset = start + length
            for index, info in enumerate(header.message_infos):
                end = payload_offset + info.length
                assert end <= len(raw)
                rows.append((header.identifier, index, info.type, payload_offset, end))
                payload_offset = end
            assert payload_offset > offset
            offset = payload_offset
        result[name] = (raw, rows)
    return result


def verify(source, output, descriptor=None, image=None, frames=None):
    expected_images = ([image] if image else []) + (frames or [])
    expected_by_id = {item[0]: item for item in expected_images}
    assert len(expected_by_id) == len(expected_images), 'Duplicate image expectations'
    with ZipFile(source) as original, ZipFile(output) as saved:
        assert len(original.namelist()) == len(set(original.namelist()))
        assert len(saved.namelist()) == len(set(saved.namelist()))
        names = {entry.filename for entry in original.infolist() if not entry.is_dir()}
        assert names == set(saved.namelist())
        changed = [name for name in sorted(names) if original.read(name) != saved.read(name)]
        original_archives, saved_archives = archives(original), archives(saved)
        assert original_archives.keys() == saved_archives.keys()
        edited, records = [], 0
        if expected_images:
            assert descriptor is not None
            pool = descriptor_pool.DescriptorPool()
            for file in descriptor_pb2.FileDescriptorSet.FromString(descriptor.read_bytes()).file:
                pool.Add(file)
            image_type = message_factory.GetMessageClass(pool.FindMessageTypeByName('TSD.ImageArchive'))
        for path, (raw, rows) in original_archives.items():
            saved_raw, saved_rows = saved_archives[path]
            assert rows == saved_rows and len(raw) == len(saved_raw), path
            records += len(rows)
            normalized = bytearray(saved_raw)
            for identifier, index, kind, start, end in rows:
                before, after = raw[start:end], saved_raw[start:end]
                if before == after:
                    continue
                selected = expected_by_id.get(identifier)
                assert selected and index == 0 and kind == 3005
                old, new = image_type.FromString(before), image_type.FromString(after)
                assert not old.FindInitializationErrors() and not new.FindInitializationErrors()
                assert new.super.geometry.position.x == selected[1] and new.super.geometry.position.y == selected[2]
                new.super.geometry.position.CopyFrom(old.super.geometry.position)
                if len(selected) == 5:
                    assert new.super.geometry.size.width == selected[3] and new.super.geometry.size.height == selected[4]
                    new.super.geometry.size.CopyFrom(old.super.geometry.size)
                assert old.SerializeToString(deterministic=True) == new.SerializeToString(deterministic=True)
                # Prove the entire payload differs only by the declared leaves.
                expected = bytearray(before)
                for (position_start, position_end), value in zip(geometry_spans(before, len(selected) == 5), selected[1:]):
                    expected[position_start:position_end] = struct.pack('<f', value)
                assert after == expected
                normalized[start:end] = before
                item = {'entry': path, 'id': identifier, 'position': list(selected[1:3])}
                if len(selected) == 5:
                    item['frame'] = list(selected[3:])
                edited.append(item)
            assert raw == normalized, path
        if expected_images:
            assert len(edited) == len(expected_images) and changed == sorted({item['entry'] for item in edited})
        else:
            assert not edited and not changed
    return {'source': source.name, 'output': output.name,
        'source_sha256': hashlib.sha256(source.read_bytes()).hexdigest(),
        'output_sha256': hashlib.sha256(output.read_bytes()).hexdigest(),
        'file_entries': len(names), 'records_including_auxiliary': records,
        'changed_entries': changed, 'edited_images': edited,
        'untouched_file_entries_identical': True, 'raw_headers_and_sibling_payloads_identical': True,
        'numbers_parser': version('numbers-parser'), 'protobuf': version('protobuf'),
        'descriptor_sha256': hashlib.sha256(descriptor.read_bytes()).hexdigest() if descriptor else None}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('output', type=Path)
    parser.add_argument('--descriptor', type=Path)
    parser.add_argument('--image', nargs=3, metavar=('ID', 'X', 'Y'))
    parser.add_argument('--image-frame', action='append', nargs=5, metavar=('ID', 'X', 'Y', 'WIDTH', 'HEIGHT'))
    parser.add_argument('--out', required=True, type=Path)
    args = parser.parse_args()
    image = (int(args.image[0]), float(args.image[1]), float(args.image[2])) if args.image else None
    frames = [(int(item[0]), *(float(value) for value in item[1:])) for item in args.image_frame or []]
    result = verify(args.source, args.output, args.descriptor, image, frames)
    args.out.write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(result, indent=2))
