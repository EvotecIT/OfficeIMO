"""Extract selected paragraph layout declarations from the unchanged Pages corpus.

Requires opt-in numbers-parser 4.19.0. This is source evidence, not an Apple export oracle.
"""
import argparse
import hashlib
import json
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile
from numbers_parser.iwafile import IWACompressedChunk, get_archive_info_and_remainder
from numbers_parser.generated.TSWPArchives_pb2 import StorageArchive, ParagraphStyleArchive

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('corpus', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
assert version('numbers-parser') == '4.19.0'
name = 'picodocs/sample-v14.4.pages'
source = args.corpus / name
source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
assert source_hash == '4714477138d0a4090fc2ee2ba2ebb6adcd0fb6ce20a28897a6247a8e17d1ddce'
records = {}
with ZipFile(source) as package:
    for entry in package.namelist():
        if not entry.endswith('.iwa'):
            continue
        data = b''.join(IWACompressedChunk._decompress_all(package.read(entry)))
        while data:
            header, payload = get_archive_info_and_remainder(data)
            offset = 0
            for ordinal, info in enumerate(header.message_infos):
                content = payload[offset:offset + info.length]
                offset += info.length
                if ordinal == 0:
                    assert header.identifier not in records
                    records[header.identifier] = (info.type, content)
            data = payload[offset:]
selected, storages = set(), []
for identifier, (kind, content) in records.items():
    if kind != 2001:
        continue
    storage = StorageArchive.FromString(content)
    if 'First ordered item' not in ''.join(storage.text):
        continue
    storages.append(identifier)
    pending = [entry.object.identifier for entry in storage.table_para_style.entries if entry.HasField('object')]
    while pending:
        current = pending.pop()
        if current in selected:
            continue
        selected.add(current)
        style_kind, payload = records[current]
        assert style_kind == 2022
        style = ParagraphStyleArchive.FromString(payload)
        if style.super.HasField('parent'):
            pending.append(style.super.parent.identifier)
assert len(storages) == 1
styles = []
for identifier in sorted(selected):
    style = ParagraphStyleArchive.FromString(records[identifier][1])
    properties = style.para_properties
    declarations = []
    for field, flag in [('line_spacing', 'line_spacing_null'), ('tabs', 'tabs_null')]:
        if not properties.HasField(field):
            continue
        assert not properties.HasField(flag) or not getattr(properties, flag)
        message = getattr(properties, field)
        evidence = {'fieldPath': '12/' + str(properties.DESCRIPTOR.fields_by_name[field].number),
                    'property': field, 'decodedProperties': str(message).strip()}
        if field == 'line_spacing' and message.mode == 0 and message.HasField('amount') and not message.HasField('baselineRule'):
            evidence['relativeMultiplier'] = message.amount
        declarations.append(evidence)
    if declarations:
        styles.append({'recordIdentifier': identifier, 'declarations': declarations})
manifest = {'source': name, 'sourceSha256': source_hash, 'extractorVersion': 'numbers-parser 4.19.0',
            'license': 'The unchanged fixture retains its existing corpus provenance and license.',
            'qualification': 'Selected body paragraph styles and their parent chain; Explicit relative multipliers are decoded independently; remaining line spacing and tab declarations stay unassessed by the shared projection. Empty declarations are retained without inventing default semantics. No native export or rendered appearance qualification.',
            'storageIdentifiers': storages, 'styles': styles}
args.output.write_text(json.dumps(manifest, indent=2) + '\n')
print(f'Extracted {sum(len(s["declarations"]) for s in styles)} declarations from {len(styles)} selected styles')
