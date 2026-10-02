"""Extract native Pages section selection declarations using pinned numbers-parser 4.19.0.

Field names follow TP.SectionArchive at the schema revision recorded in the output.
This opt-in evidence extractor does not modify fixtures or qualify native export fidelity.
"""
import argparse
import hashlib
import json
from importlib.metadata import version
from zipfile import ZipFile
from pathlib import Path
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


parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('corpus', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
assert version('numbers-parser') == '4.19.0'
name = 'picodocs/sample-v14.4.pages'
source = args.corpus / name
source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
assert source_hash == '4714477138d0a4090fc2ee2ba2ebb6adcd0fb6ce20a28897a6247a8e17d1ddce'
sections = []
with ZipFile(source) as package:
    for entry in package.namelist():
        if not entry.endswith('.iwa'):
            continue
        data = b''.join(IWACompressedChunk._decompress_all(package.read(entry)))
        while data:
            header, payload = get_archive_info_and_remainder(data)
            position = 0
            for info in header.message_infos:
                content = payload[position:position + info.length]
                position += info.length
                if info.type != 10011:
                    continue
                declaration = fields(content)
                sections.append({
                    'identifier': header.identifier,
                    'entry': entry,
                    'flags': {str(k): declaration.get(k, []) for k in (17, 18, 19, 28)},
                    'templateReferences': {str(k): [fields(v)[1][0] for v in declaration.get(k, [])]
                                           for k in (23, 24, 25)}})
            data = payload[position:]
assert len(sections) == 2
manifest = {
    'provider': 'numbers-parser', 'providerVersion': '4.19.0',
    'schema': {
        'repository': 'https://github.com/orcastor/iwork-converter',
        'revision': 'be4828260466c9c022ec12f9dc9dbfeb15ab1dea',
        'path': 'proto/TP/TPArchives.pb.go', 'message': 'SectionArchive',
        'fields': {'17': 'inherit_previous_header_footer',
                   '18': 'section_template_first_page_different',
                   '19': 'section_template_even_odd_pages_different',
                   '28': 'section_template_first_page_hides_header_footer',
                   '23': 'first_section_template_page', '24': 'even_section_template_page',
                   '25': 'odd_section_template_page'}},
    'source': name, 'sourceSha256': source_hash, 'sections': sections,
    'limits': 'Native declarations only. Both sections disable alternate templates and inheritance. Enabled variants and hidden first pages require separate native qualification; synthetic saved-output tests cover their conversion.'}
args.output.write_text(json.dumps(manifest, indent=2) + '\n', encoding='utf-8')
