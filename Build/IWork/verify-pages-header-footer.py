"""Verify Apple-saved default Pages header/footer text and its native DOCX export.

Opt-in independent evidence only: pinned numbers-parser and protobuf schemas;
no OfficeIMO assemblies. Appearance and the owned creation graph have separate proof.
"""
import argparse
import hashlib
import json
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from google.protobuf import descriptor_pb2, descriptor_pool, message_factory
from iwork_protobuf_evidence import archive_messages


def verify(source, docx, descriptor):
    assert version('numbers-parser') == '4.19.0'
    pool = descriptor_pool.DescriptorPool()
    for item in descriptor_pb2.FileDescriptorSet.FromString(descriptor.read_bytes()).file:
        pool.Add(item)

    def decode(name, data):
        result = message_factory.GetMessageClass(pool.FindMessageTypeByName(name)).FromString(data)
        assert not result.FindInitializationErrors(), result.FindInitializationErrors()
        return result

    graph, physical_count = {}, 0
    with ZipFile(source) as package:
        for _, header, info, payload in archive_messages(package):
            physical_count += 1
        for _, header, info, payload in archive_messages(package, primary_only=True):
            assert header.identifier not in graph
            graph[header.identifier] = (info.type, payload)

    def read(identifier, kind, name):
        actual, payload = graph[identifier]
        assert actual == kind
        return decode(name, payload)

    roots = [identifier for identifier, (kind, _) in graph.items() if kind == 10000]
    assert len(roots) == 1
    root = read(roots[0], 10000, 'TP.DocumentArchive')
    settings = read(root.settings.identifier, 10012, 'TP.SettingsArchive')
    assert settings.headers and settings.footers
    body = read(root.body_storage.identifier, 2001, 'TSWP.StorageArchive')
    assert len(body.table_section.entries) == 1
    section = read(body.table_section.entries[0].object.identifier, 10011, 'TP.SectionArchive')
    assert not (section.section_template_first_page_different or section.section_template_even_odd_pages_different
                or section.section_template_first_page_hides_header_footer)
    template = read(section.odd_section_template_page.identifier, 10143, 'TP.SectionTemplateArchive')
    assert len(template.headers) == len(template.footers) == 3
    page_text = []
    for index, reference in enumerate(list(template.headers) + list(template.footers)):
        storage = read(reference.identifier, 2001, 'TSWP.StorageArchive')
        text = ''.join(storage.text)
        if index not in (0, 3):
            assert text == ''
            continue
        assert text == ('OfficeIMO — Native header' if index == 0 else 'Zażółć café — native footer')
        attributes = storage.table_char_style.entries
        assert len(attributes) == 1 and attributes[0].character_index == 0
        style = read(attributes[0].object.identifier, 2021, 'TSWP.CharacterStyleArchive').char_properties
        assert style.font_name == 'Arial' and style.font_size == (14 if index == 0 else 12)
        assert style.bold == (index == 0) and style.italic == (index == 3)
        page_text.append({'kind': 'header' if index == 0 else 'footer', 'storage': reference.identifier,
            'text': text, 'font': style.font_name, 'size': style.font_size, 'bold': style.bold, 'italic': style.italic})
    paragraphs = ['Native header and footer creation'] + [f'Body line {index:02} — header/footer creation' for index in range(1, 73)]
    assert ''.join(body.text) == '\n'.join(paragraphs) + '\n'
    assert root.header_margin == 30 and root.footer_margin == 36
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    key = lambda name: '{' + ns['w'] + '}' + name
    text = lambda element: ''.join(item.text or '' for item in element.findall('.//w:t', ns))
    native_formatting = {}
    with ZipFile(docx) as package:
        document = ET.fromstring(package.read('word/document.xml'))
        assert [text(item) for item in document.findall('w:body/w:p', ns)] == paragraphs
        margin = document.find('.//w:sectPr/w:pgMar', ns)
        assert margin.get(key('header')) == '600' and margin.get(key('footer')) == '720'
        for kind, expected in [('header', page_text[0]['text']), ('footer', page_text[1]['text'])]:
            paths = [path for path in package.namelist() if path.startswith('word/' + kind) and path.endswith('.xml')]
            assert len(paths) == 1
            content = ET.fromstring(package.read(paths[0])); assert text(content) == expected
            native_formatting[kind] = {'fonts': sorted({item.get(key('ascii')) for item in content.findall('.//w:rFonts', ns)}),
                'boldDeclarations': sorted({item.get(key('val'), '1') for item in content.findall('.//w:b', ns)}),
                'italicDeclarations': sorted({item.get(key('val'), '1') for item in content.findall('.//w:i', ns)})}
    return {'source': source.name, 'sourceSha256': hashlib.sha256(source.read_bytes()).hexdigest(),
        'docx': docx.name, 'docxSha256': hashlib.sha256(docx.read_bytes()).hexdigest(),
        'descriptorSha256': hashlib.sha256(descriptor.read_bytes()).hexdigest(),
        'primaryRecords': len(graph), 'physicalRecords': physical_count, 'bodyParagraphs': len(paragraphs),
        'pageTexts': page_text, 'headerDistancePoints': 30, 'footerDistancePoints': 36,
        'nativeDocxFormatting': native_formatting,
        'limits': 'Native source and PDF retain Arial bold/italic; this Pages DOCX export maps fonts to Arial Unicode MS and clears explicit bold/italic. Text and distances match; native DOCX appearance is not qualified.'}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ('source', 'docx'):
        parser.add_argument(name, type=Path)
    parser.add_argument('--descriptor', type=Path, required=True)
    parser.add_argument('--out', dest='output', type=Path, required=True)
    args = parser.parse_args()
    result = verify(args.source, args.docx, args.descriptor)
    args.output.write_text(json.dumps(result, indent=2) + '\n')
    print(f"Verified {result['bodyParagraphs']} body paragraphs and both native page-text edits")
