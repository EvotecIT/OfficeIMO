"""Independently decode OfficeIMO's owned Pages body-document creation profile.

Opt-in evidence tooling only. Uses the same pinned numbers-parser/protobuf schema
inputs as verify-keynote-creation.py; no OfficeIMO assemblies or native app calls.
"""
import argparse
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile

from google.protobuf import descriptor_pb2, descriptor_pool, message_factory
from iwork_protobuf_evidence import fields, creation_records, creation_references

SCHEMAS = {10000: 'TP.DocumentArchive', 11006: 'TSP.PackageMetadata', 10001: 'TP.ThemeArchive',
    401: 'TSS.StylesheetArchive', 10012: 'TP.SettingsArchive', 10011: 'TP.SectionArchive',
    10143: 'TP.SectionTemplateArchive', 2021: 'TSWP.CharacterStyleArchive', 2022: 'TSWP.ParagraphStyleArchive',
    2023: 'TSWP.ListStyleArchive', 2024: 'TSWP.ColumnStyleArchive', 2050: 'TSWP.TextStylePresetArchive',
    2001: 'TSWP.StorageArchive', 2025: 'TSWP.ShapeStyleArchive', 3016: 'TSD.MediaStyleArchive'}


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
        graph = creation_records(package)
        decoded, objects = {}, []
        assert not ({71, 72} & graph.keys())
        for identifier, (kind, payload, header) in graph.items():
            decoded[identifier] = decode(SCHEMAS[kind], payload)
        native_variations = 22 in decoded and graph[22][0] == 2022 and decoded[22].super.style_identifier == 'paragraph-style-default'
        for identifier, (kind, payload, header) in graph.items():
            message = decoded[identifier]
            refs = creation_references(message, omit_style_sheet_backrefs=native_variations)
            if kind == 10001:
                theme = fields(fields(payload)[1][0])
                for field, schema in [(100, 'TSD.ThemePresetsArchive'), (110, 'TSWP.ThemePresetsArchive')]:
                    for value in theme.get(field, []):
                        refs.update(creation_references(decode(schema, value), omit_style_sheet_backrefs=native_variations))
            assert set(header.message_infos[0].object_references) == refs, identifier
            assert refs.issubset(graph.keys()), (identifier, refs - graph.keys())
            assert not header.message_infos[0].data_references
            objects.append({'id': identifier, 'type': kind, 'references': sorted(refs)})
        roots = [identifier for identifier, record in graph.items() if record[0] == 10000]
        assert len(roots) == 1
        root = decoded[roots[0]]
        body = decoded[root.body_storage.identifier]
        settings = decoded[root.settings.identifier]
        assert body.kind == 0 and body.in_document and settings.body
        assert not settings.headers and not settings.footers
        text = ''.join(body.text)
        encoded = text.encode('utf-16-le')
        positions = [entry.character_index for entry in body.table_char_style.entries]
        assert positions and positions[0] == 0 and positions == sorted(set(positions))
        boundaries, offset = {0}, 0
        for character in text:
            offset += len(character.encode('utf-16-le')) // 2
            boundaries.add(offset)
        assert all(position in boundaries and position < offset for position in positions)
        runs = []
        for index, entry in enumerate(body.table_char_style.entries):
            style = decoded[entry.object.identifier].char_properties
            assert all(style.HasField(name) for name in ('bold', 'italic', 'underline', 'strikethru'))
            assert style.tsd_fill.color == style.font_color
            end = positions[index + 1] if index + 1 < len(positions) else offset
            runs.append({'start': entry.character_index, 'end': end,
                'text': encoded[entry.character_index * 2:end * 2].decode('utf-16-le'),
                'font': style.font_name, 'size': style.font_size, 'bold': style.bold, 'italic': style.italic,
                'underline': int(style.underline), 'strikethrough': int(style.strikethru)})
        para_entries = [entry for entry in body.table_para_style.entries if entry.HasField('object')]
        para_positions = [entry.character_index for entry in para_entries]
        expected_positions = [0]
        offset = 0
        for paragraph in text.splitlines(keepends=True):
            offset += len(paragraph.encode('utf-16-le')) // 2
            if offset < len(encoded) // 2:
                expected_positions.append(offset)
        assert para_positions == expected_positions
        if native_variations:
            terminator = body.table_para_style.entries[-1]
            assert not terminator.HasField('object') and terminator.character_index == len(encoded) // 2
        paragraph_layouts = []
        for entry in para_entries:
            style = decoded[entry.object.identifier].para_properties
            assert style.alignment == 0 and style.default_tab_stops == 36
            assert style.first_line_indent == style.left_indent == style.right_indent == 0
            assert style.line_spacing.mode == 0 and style.line_spacing.amount > 0
            assert all(not tab.leader for tab in style.tabs.tabs)
            paragraph_layouts.append({'start': entry.character_index, 'styleId': entry.object.identifier,
                'lineSpacingMultiplier': style.line_spacing.amount, 'spaceBeforePoints': style.space_before,
                'spaceAfterPoints': style.space_after, 'decimalTab': style.decimal_tab,
                'tabStops': [{'positionPoints': tab.position, 'alignment': tab.alignment} for tab in style.tabs.tabs]})
        assert len(body.table_section.entries) == 1
        section = decoded[body.table_section.entries[0].object.identifier]
        assert section.section_page_number_start == 1
        for reference in (section.first_section_template_page, section.even_section_template_page, section.odd_section_template_page):
            template = decoded[reference.identifier]
            assert len(template.headers) == len(template.footers) == 3
            for storage in list(template.headers) + list(template.footers):
                assert graph[storage.identifier][0] == 2001 and ''.join(decoded[storage.identifier].text) == '\n'
        return {'source': str(path), 'sha256': hashlib.sha256(path.read_bytes()).hexdigest(),
            'descriptorSha256': hashlib.sha256(descriptor_bytes).hexdigest(), 'recordCount': len(graph),
            'page': {'width': root.page_width, 'height': root.page_height,
                'margins': [root.left_margin, root.right_margin, root.top_margin, root.bottom_margin]},
            'paragraphCount': len(para_positions), 'text': text, 'runs': runs, 'paragraphLayouts': paragraph_layouts, 'objects': objects,
            'scope': 'Owned body paragraphs and graph declarations; native application acceptance is separate.'}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('--descriptor', required=True, type=Path)
    parser.add_argument('--out', required=True, type=Path)
    arguments = parser.parse_args()
    result = verify(arguments.source, arguments.descriptor)
    arguments.out.write_text(json.dumps(result, indent=2) + '\n')
    print(f"Verified {result['recordCount']} records and {result['paragraphCount']} paragraphs: {result['sha256']}")
