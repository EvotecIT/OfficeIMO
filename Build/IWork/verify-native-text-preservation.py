"""Opt-in independent proof of fixed-length Pages/Keynote paragraph edits.

Uses numbers-parser's archive decoder and pinned Apple protobuf schemas. Compares
all file entries, raw envelopes, sibling payloads and every non-text-range byte.
OfficeIMO assemblies are not loaded. Paragraph expectations use one-based indices.
"""
import argparse
import hashlib
import importlib.util
import json
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile

from google.protobuf import descriptor_pb2, descriptor_pool, message_factory
from google.protobuf.internal.decoder import _DecodeVarint

spec = importlib.util.spec_from_file_location('preservation', Path(__file__).with_name('verify-native-preservation.py'))
preservation = importlib.util.module_from_spec(spec)
spec.loader.exec_module(preservation)


def text_field_span(data):
    offset, matches = 0, []
    while offset < len(data):
        tag, offset = _DecodeVarint(data, offset)
        field, wire, start = tag >> 3, tag & 7, offset
        assert field > 0
        if wire == 0:
            _, offset = _DecodeVarint(data, offset)
        elif wire == 2:
            length, start = _DecodeVarint(data, offset)
            offset = start + length
        elif wire in (1, 5):
            offset += 8 if wire == 1 else 4
        else:
            raise ValueError(f'Unsupported wire kind {wire}')
        assert offset <= len(data)
        if field == 3:
            assert wire == 2
            matches.append((start, offset))
    assert len(matches) == 1
    return matches[0]


def paragraph_patch(payload, storage, changes):
    assert len(storage.text) == 1
    text = storage.text[0]
    # These fixtures use LF delimiters. Other profiles need separate evidence.
    assert not any(character in text for character in '\r\u2029\u0004\u0005\u000c')
    paragraphs, expected, spans, seen = text.split('\n'), bytearray(payload), [], set()
    field_start, field_end = text_field_span(payload)
    assert payload[field_start:field_end].decode('utf-8') == text
    for item in changes:
        number, before, after = item['Paragraph'], item['Before'], item['After']
        assert 1 <= number <= len(paragraphs) and number not in seen
        assert paragraphs[number - 1] == before
        assert not any(character in after for character in '\n\r\u2028\u2029\ufffb\ufffc')
        assert len(before.encode('utf-16-le')) == len(after.encode('utf-16-le'))
        assert len(before.encode('utf-8')) == len(after.encode('utf-8'))
        prefix = '\n'.join(paragraphs[:number - 1]) + ('\n' if number > 1 else '')
        start = field_start + len(prefix.encode('utf-8'))
        end = start + len(before.encode('utf-8'))
        expected[start:end] = after.encode('utf-8')
        spans.append({'paragraph': number, 'payloadByteOffset': start, 'bytes': end - start,
            'before': before, 'after': after})
        seen.add(number)
    assert spans
    return expected, spans


def verify(source, output, descriptor, changes, kind='pages'):
    pool = descriptor_pool.DescriptorPool()
    for item in descriptor_pb2.FileDescriptorSet.FromString(descriptor.read_bytes()).file:
        pool.Add(item)

    def decode(name, payload):
        value = message_factory.GetMessageClass(pool.FindMessageTypeByName(name)).FromString(payload)
        assert not value.FindInitializationErrors(), value.FindInitializationErrors()
        return value

    with ZipFile(source) as original, ZipFile(output) as saved:
        names = {item.filename for item in original.infolist() if not item.is_dir()}
        assert len(original.namelist()) == len(set(original.namelist()))
        assert len(saved.namelist()) == len(set(saved.namelist())) and names == set(saved.namelist())
        before_graph, after_graph = preservation.archives(original), preservation.archives(saved)
        assert before_graph.keys() == after_graph.keys()
        records = {identifier: (kind, raw[start:end]) for raw, rows in before_graph.values()
            for identifier, index, kind, start, end in rows if index == 0}
        selected, context = {}, {}
        if kind == 'pages':
            roots = [payload for record_kind, payload in records.values() if record_kind == 10000]
            assert len(roots) == 1
            body_id = decode('TP.DocumentArchive', roots[0]).body_storage.identifier
            selected[body_id] = changes
            context['bodyRecordId'] = body_id
        else:
            assert kind == 'keynote'
            roots = [payload for record_kind, payload in records.values() if record_kind == 1]
            assert len(roots) == 1
            document = decode('KN.DocumentArchive', roots[0])
            show = decode('KN.ShowArchive', records[document.show.identifier][1])
            slides = {}
            for reference in show.slideTree.slides:
                node = decode('KN.SlideNodeArchive', records[reference.identifier][1])
                slide = decode('KN.SlideArchive', records[node.slide.identifier][1])
                slides[node.slide.identifier] = {ref.identifier for ref in slide.owned_drawables}
            boxes = []
            for item in changes:
                shape_id, storage_id = item['TextBoxId'], item['StorageId']
                assert records[shape_id][0] == 2011 and records[storage_id][0] == 2001
                shape = decode('TSWP.ShapeInfoArchive', records[shape_id][1])
                assert shape.is_text_box and not shape.super.super.locked
                assert shape.owned_storage.identifier == shape.deprecated_storage.identifier == storage_id
                slide_id = shape.super.super.parent.identifier
                assert shape_id in slides[slide_id]
                selected.setdefault(storage_id, []).append(item)
                boxes.append({'textBox': shape_id, 'storage': storage_id, 'slide': slide_id})
            context['textBoxes'] = boxes
        expected, spans, storages = {}, [], {}
        for storage_id, items in selected.items():
            record_kind, payload = records[storage_id]
            assert record_kind == 2001
            storage = decode('TSWP.StorageArchive', payload)
            assert kind != 'pages' or storage.kind == 0
            expected[storage_id], ranges = paragraph_patch(payload, storage, items)
            spans.extend(dict(item, storageRecordId=storage_id) for item in ranges)
            storages[storage_id] = storage
        changed, count, found = [], 0, set()
        for path, (raw, rows) in before_graph.items():
            after, after_rows = after_graph[path]
            assert rows == after_rows and len(raw) == len(after), path
            normalized = bytearray(after)
            count += len(rows)
            for identifier, index, kind, start, end in rows:
                if identifier not in selected or index != 0:
                    continue
                assert identifier not in found and kind == 2001 and after[start:end] == expected[identifier]
                new = decode('TSWP.StorageArchive', after[start:end])
                old = storages[identifier]
                new.text[0] = old.text[0]
                assert new.SerializeToString(deterministic=True) == old.SerializeToString(deterministic=True)
                normalized[start:end] = raw[start:end]
                changed.append(path)
                found.add(identifier)
            assert normalized == raw, path
        assert found == selected.keys()
        changed_entries = [name for name in sorted(names) if original.read(name) != saved.read(name)]
        assert changed_entries == sorted(set(changed))
    return {'source': source.name, 'output': output.name,
        'sourceSha256': hashlib.sha256(source.read_bytes()).hexdigest(),
        'outputSha256': hashlib.sha256(output.read_bytes()).hexdigest(),
        'fileEntries': len(names), 'recordsIncludingAuxiliary': count, **context,
        'paragraphs': spans, 'changedEntries': changed_entries,
        'onlyDeclaredUtf8RangesDiffer': True, 'rawHeadersSiblingsAndOtherBytesIdentical': True,
        'nativeStoragePropertiesIdenticalExceptText': True,
        'numbersParser': version('numbers-parser'), 'protobuf': version('protobuf'),
        'descriptorSha256': hashlib.sha256(descriptor.read_bytes()).hexdigest()}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('output', type=Path)
    parser.add_argument('--descriptor', required=True, type=Path)
    parser.add_argument('--changes', required=True, type=Path)
    parser.add_argument('--kind', choices=('pages', 'keynote'), default='pages')
    parser.add_argument('--out', required=True, type=Path)
    args = parser.parse_args()
    result = verify(args.source, args.output, args.descriptor, json.loads(args.changes.read_text(encoding='utf-8')), args.kind)
    args.out.write_text(json.dumps(result, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')
    print(json.dumps(result, indent=2, ensure_ascii=False))
