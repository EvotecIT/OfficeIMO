"""Decode selected Apple-saved Pages body layout and compare its native Word export.

Opt-in evidence only: pinned iWork protobuf descriptors and numbers-parser 4.19.0.
This inspects primary body/style records, not every auxiliary native message.
"""
import argparse
import hashlib
import json
import xml.etree.ElementTree as ET
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile

from google.protobuf import descriptor_pb2, descriptor_pool, message_factory
from numbers_parser.iwafile import IWACompressedChunk, get_archive_info_and_remainder

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def word_paragraphs(path):
    with ZipFile(path) as package:
        body = ET.fromstring(package.read('word/document.xml')).find(W + 'body')
        styles = ET.fromstring(package.read('word/styles.xml'))
    catalog = {style.get(W + 'styleId'): style for style in styles.findall(W + 'style')}
    default = styles.find(f'{W}docDefaults/{W}pPrDefault/{W}pPr')

    def parents(identifier, seen=None):
        seen = set() if seen is None else seen
        assert identifier not in seen
        seen.add(identifier)
        style = catalog.get(identifier)
        if style is None:
            return []
        parent = style.find(W + 'basedOn')
        return (parents(parent.get(W + 'val'), seen) if parent is not None else []) + [style.find(W + 'pPr')]

    result = []
    for paragraph in body.findall(W + 'p'):
        identifier = paragraph.find(f'{W}pPr/{W}pStyle')
        chain = [default] + (parents(identifier.get(W + 'val')) if identifier is not None else []) + [paragraph.find(W + 'pPr')]
        spacing, tabs = {}, {}
        for properties in chain:
            if properties is None:
                continue
            declaration = properties.find(W + 'spacing')
            if declaration is not None:
                spacing.update(declaration.attrib)
            for tab in properties.findall(f'{W}tabs/{W}tab'):
                position = int(tab.get(W + 'pos'))
                if tab.get(W + 'val') == 'clear':
                    tabs.pop(position, None)
                else:
                    tabs[position] = tab.get(W + 'val')
        assert spacing.get(W + 'lineRule', 'auto') == 'auto'
        text = ''.join(node.text or '' if node.tag == W + 't' else '\t'
            for run in paragraph.iter(W + 'r') for node in run.iter() if node.tag in (W + 't', W + 'tab'))
        result.append({'text': text, 'lineSpacingMultiplier': int(spacing.get(W + 'line', '240')) / 240,
            'spaceBeforePoints': int(spacing.get(W + 'before', '0')) / 20,
            'spaceAfterPoints': int(spacing.get(W + 'after', '0')) / 20,
            'tabStops': [{'positionPoints': position / 20, 'alignment': ['left', 'center', 'right', 'decimal'].index(alignment)}
                for position, alignment in sorted(tabs.items())]})
    return result


def extract(source, descriptor, docx):
    assert version('numbers-parser') == '4.19.0'
    pool = descriptor_pool.DescriptorPool()
    for file in descriptor_pb2.FileDescriptorSet.FromString(descriptor.read_bytes()).file:
        pool.Add(file)

    def decode(name, content):
        value = message_factory.GetMessageClass(pool.FindMessageTypeByName(name))()
        value.ParseFromString(content)
        assert not value.FindInitializationErrors()
        return value

    records = {}
    with ZipFile(source) as package:
        for name in package.namelist():
            if not name.endswith('.iwa'):
                continue
            data = b''.join(IWACompressedChunk._decompress_all(package.read(name)))
            while data:
                header, payload = get_archive_info_and_remainder(data)
                offset = 0
                for ordinal, info in enumerate(header.message_infos):
                    if ordinal == 0:
                        assert header.identifier not in records
                        records[header.identifier] = (info.type, payload[:info.length])
                    offset += info.length
                data = payload[offset:]
    bodies = [(identifier, decode('TSWP.StorageArchive', payload)) for identifier, (kind, payload) in records.items() if kind == 2001]
    identifier, body = next((identifier, value) for identifier, value in bodies if value.kind == 0 and value.in_document)
    assert sum(value.kind == 0 and value.in_document for _, value in bodies) == 1
    text = ''.join(body.text)
    encoded = text.encode('utf-16-le')
    entries = [entry for entry in body.table_para_style.entries if entry.HasField('object')]
    paragraphs = []
    for ordinal, entry in enumerate(entries):
        chain, current = [], entry.object.identifier
        while current:
            assert current not in [identifier for identifier, _ in chain]
            kind, payload = records[current]
            assert kind == 2022
            style = decode('TSWP.ParagraphStyleArchive', payload)
            chain.append((current, style))
            current = style.super.parent.identifier if style.super.HasField('parent') else 0
        layout = {'lineSpacingMultiplier': 1, 'spaceBeforePoints': 0, 'spaceAfterPoints': 0, 'decimalTabSymbol': None, 'tabStops': []}
        for _, style in reversed(chain):
            properties = style.para_properties
            for field, key in [('space_before', 'spaceBeforePoints'), ('space_after', 'spaceAfterPoints'), ('decimal_tab', 'decimalTabSymbol')]:
                if properties.HasField(field):
                    layout[key] = getattr(properties, field)
            if properties.HasField('decimal_tab_null') and properties.decimal_tab_null:
                layout['decimalTabSymbol'] = None
            if properties.HasField('line_spacing_null') and properties.line_spacing_null:
                layout['lineSpacingMultiplier'] = 1
            if properties.HasField('line_spacing'):
                assert properties.line_spacing.mode == 0
                layout['lineSpacingMultiplier'] = properties.line_spacing.amount if properties.line_spacing.HasField('amount') else 1
                assert layout['lineSpacingMultiplier'] > 0
            if properties.HasField('tabs_null') and properties.tabs_null:
                layout['tabStops'] = []
            if properties.HasField('tabs'):
                assert all(not tab.leader for tab in properties.tabs.tabs)
                layout['tabStops'] = [{'positionPoints': tab.position, 'alignment': tab.alignment} for tab in properties.tabs.tabs]
        end = entries[ordinal + 1].character_index if ordinal + 1 < len(entries) else len(encoded) // 2
        layout['text'] = encoded[entry.character_index * 2:end * 2].decode('utf-16-le').removesuffix('\n')
        paragraphs.append({'start': entry.character_index, 'styleChain': [identifier for identifier, _ in chain], **layout})
    reference = word_paragraphs(docx)
    assert len(paragraphs) == len(reference)
    for actual, exported in zip(paragraphs, reference):
        assert {key: actual[key] for key in exported} == exported, (actual, exported)
    return {'source': source.name, 'sourceSha256': hashlib.sha256(source.read_bytes()).hexdigest(),
        'descriptorSha256': hashlib.sha256(descriptor.read_bytes()).hexdigest(),
        'nativeWordExport': docx.name, 'nativeWordExportSha256': hashlib.sha256(docx.read_bytes()).hexdigest(),
        'selectedStorageIdentifier': identifier, 'paragraphs': paragraphs,
        'scope': 'Selected body paragraph styles and parent chains; complete authored body text, relative spacing and tabs agree with the native Word export. Auxiliary messages and other native features are not assessed.'}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('--descriptor', type=Path, required=True)
    parser.add_argument('--docx', type=Path, required=True)
    parser.add_argument('--out', type=Path, required=True)
    arguments = parser.parse_args()
    result = extract(arguments.source, arguments.descriptor, arguments.docx)
    arguments.out.write_text(json.dumps(result, indent=2) + '\n')
    print(f"Decoded {len(result['paragraphs'])} native paragraphs matching Apple Word layout: {result['sourceSha256']}")
