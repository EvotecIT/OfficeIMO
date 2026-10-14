"""Extract native character ranges and Apple PPTX runs with numbers-parser 4.19.0.

Opt-in qualification only. No OfficeIMO assemblies are loaded. Native property
presence and PPTX inheritance are recorded separately; this does not qualify
unselected templates, all Keynote records, or visual equivalence.
"""
import argparse
import hashlib
import json
import xml.etree.ElementTree as ET
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

from google.protobuf.json_format import MessageToDict
from numbers_parser.generated import TSWPArchives_pb2 as TSWP
from numbers_parser.iwafile import IWACompressedChunk, get_archive_info_and_remainder

NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}


def native_ranges(path):
    graph = {}
    with ZipFile(path) as package:
        entries = {name: package.read(name) for name in package.namelist() if name.endswith('.iwa')}
        if 'Index.zip' in package.namelist():
            with ZipFile(BytesIO(package.read('Index.zip'))) as index:
                entries.update({name: index.read(name) for name in index.namelist() if name.endswith('.iwa')})
        for name, payload in entries.items():
            data = b''.join(IWACompressedChunk._decompress_all(payload))
            while data:
                header, remainder = get_archive_info_and_remainder(data)
                offset = 0
                for info in header.message_infos:
                    schema = {2001: TSWP.StorageArchive, 2021: TSWP.CharacterStyleArchive,
                              2022: TSWP.ParagraphStyleArchive}.get(info.type)
                    if schema:
                        message = schema()
                        message.ParseFromString(remainder[offset:offset + info.length])
                        assert not message.FindInitializationErrors()
                        graph[header.identifier] = (info.type, message)
                    offset += info.length
                data = remainder[offset:]
    result = []
    for identifier, (kind, message) in graph.items():
        if kind != 2001:
            continue
        text = ''.join(message.text)
        if not text.strip():
            continue
        ranges = []
        for entry in message.table_char_style.entries:
            style = entry.object.identifier if entry.HasField('object') else None
            assert style is None or graph[style][0] == 2021
            ranges.append({'start': entry.character_index, 'style': style,
                'explicit_properties': MessageToDict(graph[style][1].char_properties,
                    preserving_proto_field_name=True) if style is not None else {}})
        paragraphs = [{'start': entry.character_index, 'style': entry.object.identifier,
            'explicit_character_properties': MessageToDict(graph[entry.object.identifier][1].char_properties,
                preserving_proto_field_name=True)} for entry in message.table_para_style.entries
            if entry.HasField('object')]
        result.append({'storage': identifier, 'text': text, 'utf16_length': len(text.encode('utf-16-le')) // 2,
            'character_ranges': ranges, 'paragraph_defaults': paragraphs})
    return result


def properties(element):
    if element is None:
        return {}
    result = dict(element.attrib)
    for tag in ('latin', 'ea', 'cs'):
        child = element.find('a:' + tag, NS)
        if child is not None:
            result[tag] = child.attrib['typeface']
    color = element.find('a:solidFill/a:srgbClr', NS)
    if color is not None:
        result['color'] = color.attrib['val']
    return result


def pptx_runs(path):
    result = []
    with ZipFile(path) as package:
        theme_fonts = {}
        for name in sorted(package.namelist()):
            if name.startswith('ppt/theme/') and name.endswith('.xml'):
                theme = ET.fromstring(package.read(name))
                scheme = theme.find('a:themeElements/a:fontScheme', NS)
                if scheme is not None:
                    theme_fonts[name] = {
                        role.tag.rsplit('}', 1)[-1]: {
                            face.tag.rsplit('}', 1)[-1]: face.attrib.get('typeface', '')
                            for face in role if face.tag.rsplit('}', 1)[-1] in ('latin', 'ea', 'cs')}
                        for role in scheme}
        presentation = ET.fromstring(package.read('ppt/presentation.xml'))
        size = presentation.find('p:sldSz', NS)
        slide_count = len(presentation.find('p:sldIdLst', NS))
        for number in range(1, slide_count + 1):
            slide = ET.fromstring(package.read(f'ppt/slides/slide{number}.xml'))
            boxes = []
            for shape in slide.findall('.//p:sp', NS):
                body = shape.find('p:txBody', NS)
                if body is None:
                    continue
                paragraphs = []
                for paragraph in body.findall('a:p', NS):
                    defaults = properties(paragraph.find('a:pPr/a:defRPr', NS))
                    runs = []
                    for run in paragraph.findall('a:r', NS):
                        overrides = properties(run.find('a:rPr', NS))
                        effective = dict(defaults, **overrides)
                        runs.append({'text': run.find('a:t', NS).text or '',
                            'properties': effective, 'overrides': overrides})
                    paragraphs.append({'defaults': defaults, 'runs': runs})
                geometry = shape.find('p:spPr/a:xfrm', NS)
                boxes.append({'paragraphs': paragraphs, 'geometry':
                    {child.tag.rsplit('}', 1)[-1]: dict(child.attrib) for child in geometry}
                    if geometry is not None else None})
            result.append({'slide': number, 'text_boxes': boxes})
    return {'canvas_emu': dict(size.attrib), 'theme_font_schemes': theme_fonts, 'slides': result}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('native', type=Path)
    parser.add_argument('pptx', type=Path)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    assert version('numbers-parser') == '4.19.0'
    result = {'numbers_parser_version': version('numbers-parser'),
        'native_sha256': hashlib.sha256(args.native.read_bytes()).hexdigest(),
        'pptx_sha256': hashlib.sha256(args.pptx.read_bytes()).hexdigest(),
        'native_storage': native_ranges(args.native), 'pptx': pptx_runs(args.pptx)}
    args.output.write_text(json.dumps(result, indent=2, ensure_ascii=False) + '\n')
    print({'native_storages': len(result['native_storage']), 'pptx_slides': len(result['pptx']['slides'])})


if __name__ == '__main__':
    main()
