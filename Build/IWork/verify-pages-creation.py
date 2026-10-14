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
from iwork_table_creation_evidence import verify_tables
from iwork_image_creation_evidence import verify_images
from iwork_list_creation_evidence import verify_decimal_lists

SCHEMAS = {10000: 'TP.DocumentArchive', 11006: 'TSP.PackageMetadata', 10001: 'TP.ThemeArchive',
    401: 'TSS.StylesheetArchive', 10012: 'TP.SettingsArchive', 10011: 'TP.SectionArchive',
    10143: 'TP.SectionTemplateArchive', 2021: 'TSWP.CharacterStyleArchive', 2022: 'TSWP.ParagraphStyleArchive',
    2023: 'TSWP.ListStyleArchive', 2024: 'TSWP.ColumnStyleArchive', 2050: 'TSWP.TextStylePresetArchive',
    2001: 'TSWP.StorageArchive', 2003: 'TSWP.DrawableAttachmentArchive',
    2025: 'TSWP.ShapeStyleArchive', 3005: 'TSD.ImageArchive', 3016: 'TSD.MediaStyleArchive',
    6000: 'TST.TableInfoArchive', 6001: 'TST.TableModelArchive', 6002: 'TST.Tile',
    6003: 'TST.TableStyleArchive', 6004: 'TST.CellStyleArchive', 6005: 'TST.TableDataList',
    6006: 'TST.HeaderStorageBucket', 6008: 'TST.TableStylePresetArchive', 6247: 'TST.TableStyleNetworkArchive'}


def verify(path, descriptor, page_text=False):
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
        decoded, objects, data_references = {}, [], {}
        assert not ({71, 72} & graph.keys())
        for identifier, (kind, payload, header) in graph.items():
            decoded[identifier] = decode(SCHEMAS[kind], payload)
        native_variations = 22 in decoded and graph[22][0] == 2022 and decoded[22].super.style_identifier == 'paragraph-style-default'
        for identifier, (kind, payload, header) in graph.items():
            message = decoded[identifier]
            refs = creation_references(message, omit_style_sheet_backrefs=native_variations)
            if kind == 10001:
                theme = fields(fields(payload)[1][0])
                for field, schema in [(100, 'TSD.ThemePresetsArchive'), (110, 'TSWP.ThemePresetsArchive'), (200, 'TST.ThemePresetsArchive')]:
                    for value in theme.get(field, []):
                        refs.update(creation_references(decode(schema, value), omit_style_sheet_backrefs=native_variations))
            assert set(header.message_infos[0].object_references) == refs, identifier
            assert refs.issubset(graph.keys()), (identifier, refs - graph.keys())
            data_refs = creation_references(message, 'TSP.DataReference')
            assert set(header.message_infos[0].data_references) == data_refs, identifier
            if data_refs:
                data_references[identifier] = data_refs
            objects.append({'id': identifier, 'type': kind, 'references': sorted(refs)})
        roots = [identifier for identifier, record in graph.items() if record[0] == 10000]
        assert len(roots) == 1
        root = decoded[roots[0]]
        body = decoded[root.body_storage.identifier]
        settings = decoded[root.settings.identifier]
        assert body.kind == 0 and body.in_document and settings.body
        if not page_text:
            assert not settings.headers and not settings.footers
        else:
            assert settings.headers or settings.footers
            assert not root.lays_out_body_vertically and not root.change_tracking_enabled
        text = ''.join(body.text)
        encoded = text.encode('utf-16-le')
        decimal_items = verify_decimal_lists(graph, decoded, body, text)
        tables = verify_tables(graph, decoded)
        images, resource_paths = verify_images(graph, decoded, decode, package, data_references)
        assert package.namelist() == sorted(package.namelist())
        assert len(package.namelist()) == 5 + len(resource_paths)
        assert all(item.compress_type == 0 and item.date_time == (1980, 1, 1, 0, 0, 0)
                   for item in package.infolist())
        inline_tables, inline_images = [], []
        for entry in body.table_attachment.entries:
            attachment = decoded[entry.object.identifier]
            assert graph[entry.object.identifier][0] == 2003
            assert attachment.h_offset_type == attachment.v_offset_type == 0
            assert attachment.h_offset == attachment.v_offset == 0
            drawable = decoded[attachment.drawable.identifier]
            kind = graph[attachment.drawable.identifier][0]
            assert kind in (6000, 3005) and drawable.super.parent.identifier == root.body_storage.identifier
            wrap = drawable.super.exterior_text_wrap
            assert drawable.super.HasField('exterior_text_wrap') and wrap.type == 0 and wrap.direction == 2
            assert wrap.fit_type == 0 and wrap.margin == 12 and wrap.alpha_threshold == 0.5 and not wrap.is_html_wrap
            geometry = drawable.super.geometry
            assert geometry.position.x == geometry.position.y == geometry.angle == 0
            assert geometry.size.width <= root.page_width - root.left_margin - root.right_margin
            assert geometry.size.height <= root.page_height - root.top_margin - root.bottom_margin
            start = entry.character_index * 2
            assert encoded[start:start + 4].decode('utf-16-le') == '\ufffc\n'
            assert start == 0 or encoded[start - 2:start].decode('utf-16-le') == '\n'
            item = {'start': entry.character_index, 'attachmentId': entry.object.identifier,
                'drawableId': attachment.drawable.identifier,
                'widthPoints': geometry.size.width, 'heightPoints': geometry.size.height}
            if kind == 6000:
                assert drawable.formula_coord_space == 0
                inline_tables.append(dict(item, tableId=drawable.tableModel.identifier))
            else:
                inline_images.append(item)
        assert len(inline_tables) == len(tables)
        assert {table['id'] for table in tables} == {table['tableId'] for table in inline_tables}
        assert {image['id'] for image in images} == {item['drawableId'] for item in inline_images}
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
        assert not page_text or not (section.section_template_first_page_different
            or section.section_template_even_odd_pages_different or section.section_template_first_page_hides_header_footer)
        page_texts = []
        for reference in (section.first_section_template_page, section.even_section_template_page, section.odd_section_template_page):
            template = decoded[reference.identifier]
            assert len(template.headers) == len(template.footers) == 3
            for index, storage in enumerate(list(template.headers) + list(template.footers)):
                assert graph[storage.identifier][0] == 2001
                content = decoded[storage.identifier]
                if not page_text:
                    assert ''.join(content.text) == '\n'
                    continue
                assert content.kind == 1 and content.in_document
                value = ''.join(content.text)
                assert not any(character in value for character in '\n\r\t\u2028\u2029')
                enabled = settings.headers if index < 3 else settings.footers
                if index not in (0, 3) or not enabled:
                    assert value == ''
                    continue
                distance = root.header_margin if index == 0 else root.footer_margin
                margin = root.top_margin if index == 0 else root.bottom_margin
                assert 0 <= distance < margin
                attributes = content.table_char_style.entries
                assert len(attributes) == 1 and attributes[0].character_index == 0
                style = decoded[attributes[0].object.identifier].char_properties
                paragraph = decoded[content.table_para_style.entries[0].object.identifier].para_properties
                assert paragraph.alignment == 0 and paragraph.line_spacing.amount == 1
                if reference.identifier == section.odd_section_template_page.identifier and not any(
                        item['storage'] == storage.identifier for item in page_texts):
                    page_texts.append({'kind': 'header' if index == 0 else 'footer', 'storage': storage.identifier,
                        'text': value, 'distancePoints': distance, 'font': style.font_name, 'size': style.font_size,
                        'bold': style.bold, 'italic': style.italic, 'underline': int(style.underline),
                        'strikethrough': int(style.strikethru)})
        return {'source': str(path), 'sha256': hashlib.sha256(path.read_bytes()).hexdigest(),
            'descriptorSha256': hashlib.sha256(descriptor_bytes).hexdigest(), 'recordCount': len(graph),
            'page': {'width': root.page_width, 'height': root.page_height,
                'margins': [root.left_margin, root.right_margin, root.top_margin, root.bottom_margin]},
            'paragraphCount': len(para_positions), 'text': text, 'runs': runs, 'paragraphLayouts': paragraph_layouts, 'objects': objects,
            'nativeTables': tables, 'inlineTables': inline_tables,
            'imageResources': images, 'inlineImages': inline_images, 'decimalItems': decimal_items,
            **({'pageTexts': page_texts} if page_text else {}),
            'scope': 'Owned body paragraphs, inline PNG/JPEG images, scalar tables and graph declarations; native application acceptance is separate.'}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('--descriptor', required=True, type=Path)
    parser.add_argument('--out', required=True, type=Path)
    parser.add_argument('--page-text', action='store_true', help='Verify the bounded default header/footer extension.')
    arguments = parser.parse_args()
    result = verify(arguments.source, arguments.descriptor, arguments.page_text)
    arguments.out.write_text(json.dumps(result, indent=2) + '\n')
    print(f"Verified {result['recordCount']} records and {result['paragraphCount']} paragraphs: {result['sha256']}")
