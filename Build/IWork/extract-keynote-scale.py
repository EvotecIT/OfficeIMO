"""Verify the bounded 512-slide Apple Keynote scale reference independently.

Opt-in evidence only: numbers-parser 4.19.0 and the pinned Keynote descriptor
decode selected native records. ElementTree reads Apple's PPTX; Poppler supplies
the PDF bbox XML. No OfficeIMO assemblies are loaded. Two repeated layouts do
not qualify varied large presentations, opaque records or complete appearance.
"""
import argparse
import colorsys
import hashlib
import json
import posixpath
import struct
import unicodedata
import xml.etree.ElementTree as ET
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile

from google.protobuf import descriptor_pb2, descriptor_pool, message_factory
from iwork_protobuf_evidence import archive_messages, fields

DESCRIPTOR_SHA256 = '2c46d6ee3eaf4d8419d0cc8329b74776f69517c1e14de0571e0434e29c4fc216'
TEXT = [['hello keynote', 'first bullet'],
        ['wrapped title keeps all of its lines in this fixed frame']]
COLORS = ['56C1FF', 'FF968D']
NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'h': 'http://www.w3.org/1999/xhtml'}


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def native_slides(path, descriptor):
    assert version('numbers-parser') == '4.19.0'
    assert sha(descriptor) == DESCRIPTOR_SHA256
    pool = descriptor_pool.DescriptorPool()
    for file in descriptor_pb2.FileDescriptorSet.FromString(descriptor.read_bytes()).file:
        pool.Add(file)
    graph = {}
    with ZipFile(path) as package:
        for _, header, info, payload in archive_messages(package):
            graph.setdefault((header.identifier, info.type), []).append(payload)

    def payload(identifier, kind):
        candidates = graph[(identifier, kind)]
        assert len(candidates) == 1, (identifier, kind)
        return candidates[0]

    def decode(identifier, kind, name):
        message = message_factory.GetMessageClass(pool.FindMessageTypeByName(name))()
        message.ParseFromString(payload(identifier, kind))
        assert not message.FindInitializationErrors(), (identifier, name)
        return message

    def background(identifier):
        chain = []
        while identifier:
            assert identifier not in [item[0] for item in chain]
            style = fields(payload(identifier, 9))
            chain.append((identifier, style))
            parent = fields(style[1][0]).get(3)
            identifier = fields(parent[0])[1][0] if parent else None
        rgb = None
        for _, style in reversed(chain):
            properties = fields(style[11][0]) if 11 in style else {}
            if 1 not in properties:
                continue
            fill = fields(properties[1][0])
            assert set(fill) == {1} and len(fill[1]) == 1
            color = fields(fill[1][0])
            assert set(color) in ({1, 3, 4, 5, 6, 12}, {1, 3, 4, 5, 6, 12, 13})
            assert color[1] == [1] and color[12] == [1]
            if 13 in color:
                assert len(color[13]) == 1 and struct.unpack('<f', color[13][0])[0] == 1
            rgba = [struct.unpack('<f', color[field][0])[0] for field in (3, 4, 5, 6)]
            assert rgba[3] == 1 and all(0 <= value <= 1 for value in rgba)
            rgb = ''.join(f'{int(channel * 255 + .5):02X}' for channel in rgba[:3])
        return rgb

    roots = [identifier for identifier, kind in graph if kind == 1]
    assert len(roots) == 1
    document = decode(roots[0], 1, 'KN.DocumentArchive')
    show = decode(document.show.identifier, 2, 'KN.ShowArchive')
    assert (show.size.width, show.size.height) == (1024, 768)
    rows = []
    for position, reference in enumerate(show.slideTree.slides, 1):
        node = decode(reference.identifier, 4, 'KN.SlideNodeArchive')
        slide = decode(node.slide.identifier, 5, 'KN.SlideArchive')
        assert not node.isSkipped and slide.inDocument
        assert [ref.identifier for ref in slide.owned_drawables] == [ref.identifier for ref in slide.drawables_z_order]
        boxes = []
        for drawable in slide.owned_drawables:
            shape = decode(drawable.identifier, 7, 'KN.PlaceholderArchive').super \
                if (drawable.identifier, 7) in graph else decode(drawable.identifier, 2011, 'TSWP.ShapeInfoArchive')
            storage_id = shape.owned_storage.identifier if shape.HasField('owned_storage') else shape.deprecated_storage.identifier
            storage = decode(storage_id, 2001, 'TSWP.StorageArchive')
            geometry = shape.super.super.geometry
            assert geometry.angle == 0 and not shape.super.pathsource.horizontalFlip and not shape.super.pathsource.verticalFlip
            boxes.append({'text': ''.join(storage.text), 'geometryPoints': [geometry.position.x,
                geometry.position.y, geometry.size.width, geometry.size.height]})
        note = decode(slide.note.identifier, 15, 'KN.NoteArchive')
        notes = ''.join(decode(note.containedStorage.identifier, 2001, 'TSWP.StorageArchive').text)
        profile = (position - 1) % 2
        assert [box['text'] for box in boxes] == TEXT[profile], position
        assert notes == ('note text here' if profile else ''), position
        rgb = background(slide.style.identifier)
        assert rgb == COLORS[profile], position
        rows.append({'position': position, 'node': reference.identifier, 'slide': node.slide.identifier,
                     'background': rgb, 'boxes': boxes, 'notes': notes})
    assert len(rows) == 512 and len({row['slide'] for row in rows}) == 512
    assert len({row['node'] for row in rows}) == 512
    return rows, sum(len(values) for values in graph.values())


def pptx_slides(path):
    with ZipFile(path) as package:
        def xml(part):
            return ET.fromstring(package.read(part))

        def relationships(part):
            rel_part = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
            if rel_part not in package.namelist():
                return {}
            result = {}
            for item in xml(rel_part):
                assert item.get('TargetMode') != 'External'
                result[item.get('Id')] = (item.get('Type').rsplit('/', 1)[-1],
                    posixpath.normpath(posixpath.join(posixpath.dirname(part), item.get('Target'))).lstrip('/'))
            return result

        def target(part, kind):
            values = [name for role, name in relationships(part).values() if role == kind]
            assert len(values) <= 1, (part, kind)
            return values[0] if values else None

        def text(shape):
            return '\n'.join(''.join(node.text or '' for node in paragraph.findall('.//a:t', NS))
                             for paragraph in shape.findall('p:txBody/a:p', NS))

        def geometry(shape, layout):
            transform = shape.find('p:spPr/a:xfrm', NS)
            if transform is None:
                placeholder = shape.find('p:nvSpPr/p:nvPr/p:ph', NS)
                assert placeholder is not None
                matches = [candidate for candidate in layout.findall('p:cSld/p:spTree/p:sp', NS)
                    if (ph := candidate.find('p:nvSpPr/p:nvPr/p:ph', NS)) is not None
                    and ph.get('idx', '0') == placeholder.get('idx', '0')]
                assert len(matches) == 1
                transform = matches[0].find('p:spPr/a:xfrm', NS)
            assert transform is not None and int(transform.get('rot', '0')) == 0
            assert transform.get('flipH', '0') in ('0', 'false') and transform.get('flipV', '0') in ('0', 'false')
            offset, extent = transform.find('a:off', NS), transform.find('a:ext', NS)
            return [int(offset.get('x')) / 12700, int(offset.get('y')) / 12700,
                    int(extent.get('cx')) / 12700, int(extent.get('cy')) / 12700]

        presentation = xml('ppt/presentation.xml')
        size = presentation.find('p:sldSz', NS)
        assert (int(size.get('cx')), int(size.get('cy'))) == (1024 * 12700, 768 * 12700)
        rels = relationships('ppt/presentation.xml')
        rows, parts = [], set()
        for position, ref in enumerate(presentation.findall('p:sldIdLst/p:sldId', NS), 1):
            role, part = rels[ref.get('{' + NS['r'] + '}id')]
            assert role == 'slide' and part not in parts
            parts.add(part)
            slide = xml(part)
            assert slide.get('show', '1') == '1'
            layout_part = target(part, 'slideLayout')
            layout = xml(layout_part)
            master_part = target(layout_part, 'slideMaster')
            master = xml(master_part)
            theme = xml(target(master_part, 'theme'))
            fill = slide.find('p:cSld/p:bg/p:bgPr/a:solidFill', NS)
            assert fill is not None and len(fill) == 1
            color = fill[0]
            transforms = list(color)
            if color.tag == '{' + NS['a'] + '}schemeClr':
                name = color.get('val')
                name = master.find('p:clrMap', NS).get(name, name)
                color = theme.find('a:themeElements/a:clrScheme/a:' + name, NS)[0]
            assert color.tag in ('{' + NS['a'] + '}srgbClr', '{' + NS['a'] + '}sysClr')
            rgb = color.get('val') if color.tag.endswith('srgbClr') else color.get('lastClr')
            h, lightness, saturation = colorsys.rgb_to_hls(*(int(rgb[i:i + 2], 16) / 255 for i in (0, 2, 4)))
            for transform in transforms:
                value = int(transform.get('val'))
                if transform.tag == '{' + NS['a'] + '}hueOff':
                    h = (h + value / 21600000) % 1
                elif transform.tag == '{' + NS['a'] + '}lumOff':
                    lightness = min(1, max(0, lightness + value / 100000))
                else:
                    raise AssertionError('Unqualified color transform: ' + transform.tag)
            rgb = ''.join(f'{int(value * 255 + .5):02X}' for value in colorsys.hls_to_rgb(h, lightness, saturation))
            assert all(item.tag in ('{' + NS['p'] + '}' + name for name in ('nvGrpSpPr', 'grpSpPr', 'sp'))
                       for item in slide.find('p:cSld/p:spTree', NS))
            shapes = slide.findall('p:cSld/p:spTree/p:sp', NS)
            boxes = [{'text': text(shape), 'geometryPoints': geometry(shape, layout)} for shape in shapes]
            notes_part = target(part, 'notesSlide')
            notes = []
            if notes_part:
                for shape in xml(notes_part).findall('p:cSld/p:spTree/p:sp', NS):
                    placeholder = shape.find('p:nvSpPr/p:nvPr/p:ph', NS)
                    if placeholder is not None and placeholder.get('type') == 'body':
                        notes.append(text(shape))
            rows.append({'position': position, 'part': part, 'background': rgb,
                         'boxes': boxes, 'notes': '\n'.join(notes)})
        assert len(rows) == 512
        return rows


def extract(args):
    native, records = native_slides(args.package, args.descriptor)
    exported = pptx_slides(args.pptx)
    destination_rows = pptx_slides(args.destination_pptx) if args.destination_pptx else None
    pages = ET.parse(args.pdf_bbox).findall('.//h:page', NS)
    pdf_text = json.loads(args.pdf_text.read_text(encoding='utf-8'))
    assert pdf_text['sourceSha256'] == sha(args.pdf)
    assert len(pages) == len(pdf_text['pages']) == len(native) == len(exported) == 512
    for source, destination, page, text_page in zip(native, exported, pages, pdf_text['pages']):
        assert source['background'] == destination['background'] and source['notes'] == destination['notes']
        assert len(source['boxes']) == len(destination['boxes'])
        for left, right in zip(source['boxes'], destination['boxes']):
            assert left['text'] == right['text']
            assert all(abs(a - b) < .001 for a, b in zip(left['geometryPoints'], right['geometryPoints']))
        assert (float(page.get('width')), float(page.get('height'))) == (1024, 768)
        assert text_page['position'] == source['position'] and text_page['sizePoints'] == [1024, 768]
        assert text_page['text'].count('•') == (1 if source['position'] % 2 else 0)
        text = unicodedata.normalize('NFKC', text_page['text']).replace('•', ' ')
        expected = ' '.join(TEXT[(source['position'] - 1) % 2])
        assert ' '.join(text.split()) == expected, source['position']
    if destination_rows is not None:
        for source, destination in zip(native, destination_rows):
            assert source['background'] == destination['background'] and source['notes'] == destination['notes']
            assert len(source['boxes']) == len(destination['boxes'])
            for left, right in zip(source['boxes'], destination['boxes']):
                # OfficeIMO preserves native paragraph terminators in saved text.
                assert left['text'] == right['text'].rstrip('\n')
                assert all(abs(a - b) < .001 for a, b in zip(left['geometryPoints'], right['geometryPoints']))
    identities = '\n'.join(str(row['slide']) for row in native).encode('ascii')
    result = {'provider': 'numbers-parser', 'providerVersion': version('numbers-parser'),
        'descriptorSha256': DESCRIPTOR_SHA256,
        'schemaRevision': 'be4828260466c9c022ec12f9dc9dbfeb15ab1dea',
        'files': {name: {'path': path.name, 'bytes': path.stat().st_size, 'sha256': sha(path)}
                  for name, path in [('source', args.package), ('applePptx', args.pptx),
                                     ('applePdf', args.pdf), ('pdfBbox', args.pdf_bbox), ('pdfText', args.pdf_text)]},
        'nativeMessageCount': records, 'verifiedSlides': len(native), 'uniqueNativeSlideIdentities': 512,
        'orderedNativeSlideIdentitiesSha256': hashlib.sha256(identities).hexdigest(),
        'orderedSelectedSemanticsSha256': hashlib.sha256(json.dumps(native, sort_keys=True, separators=(',', ':')).encode()).hexdigest(),
        'slideSizePoints': [1024, 768], 'geometryTolerancePoints': .001,
        'pdfTextProvider': pdf_text['provider'], 'pdfTextProviderVersion': pdf_text['providerVersion'],
        'profileOccurrences': [256, 256], 'profiles': [native[0], native[1]],
        'selectedSlides': [native[i - 1] for i in (1, 2, 255, 256, 511, 512)],
        'checks': ['All ordered native slide/node identities distinct and initialized',
            'All 512 native slides: exact text, notes, backgrounds, drawable order, no skipped slides',
            'All 512 Apple PPTX slides: relationship order, text, notes, inherited geometry and resolved background',
            'All 512 Apple PDF pages: canvas and complete ordered NFKC text, known bullet removed'],
        'limits': 'Two repeated text layouts in one Keynote 15.4 document. Selected native records only; auxiliary/opaque records are inventoried, not fully assessed. No broad master/theme, typography, varied-object or complete PDF appearance claim. Geometry tolerance covers EMU rounding. pypdf text ligatures normalize with NFKC; Poppler bbox text misplaces fi ligatures and is used only for canvas checks. pypdf warns about two wrong-pointing objects in this Apple PDF; no conformance claim.'}
    if destination_rows is not None:
        result['officeimoPptx'] = {'path': args.destination_pptx.name, 'sha256': sha(args.destination_pptx),
                                 'bytes': args.destination_pptx.stat().st_size, 'verifiedSlides': len(destination_rows),
                                 'checks': 'Every slide text, notes, background, frame geometry, canvas and visible status; no native rendering or appearance claim.'}
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('package', type=Path)
    parser.add_argument('--descriptor', type=Path, required=True)
    parser.add_argument('--pptx', type=Path, required=True)
    parser.add_argument('--pdf', type=Path, required=True)
    parser.add_argument('--pdf-bbox', type=Path, required=True)
    parser.add_argument('--pdf-text', type=Path, required=True)
    parser.add_argument('--destination-pptx', type=Path)
    parser.add_argument('--out', type=Path, required=True)
    args = parser.parse_args()
    result = extract(args)
    args.out.write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
    print('Verified all', result['verifiedSlides'], 'native slides, Apple PPTX slides and PDF pages')


if __name__ == '__main__':
    main()
