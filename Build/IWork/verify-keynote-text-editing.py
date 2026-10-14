"""Independent selected native/PPTX proof for the three-slide text-edit fixture.

Opt-in evidence: pinned Apple schemas and numbers-parser decode native storage,
ordered shapes, geometry and image resources. ElementTree reads the saved PPTX.
This does not assess opaque records or qualify general Keynote appearance.
"""
import argparse
import hashlib
import json
import posixpath
import xml.etree.ElementTree as ET
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile

from google.protobuf import descriptor_pb2, descriptor_pool, message_factory
from iwork_protobuf_evidence import archive_messages

DESCRIPTOR_SHA256 = '2c46d6ee3eaf4d8419d0cc8329b74776f69517c1e14de0571e0434e29c4fc216'
NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}


def sha(data):
    return hashlib.sha256(data).hexdigest()


def native(path, descriptor):
    assert sha(descriptor.read_bytes()) == DESCRIPTOR_SHA256
    pool = descriptor_pool.DescriptorPool()
    for file in descriptor_pb2.FileDescriptorSet.FromString(descriptor.read_bytes()).file:
        pool.Add(file)
    with ZipFile(path) as package:
        records, messages = {}, 0
        for _, header, info, payload in archive_messages(package):
            messages += 1
            records.setdefault((header.identifier, info.type), []).append(payload)

        def decode(identifier, kind, name):
            values = records[(identifier, kind)]
            assert len(values) == 1, (identifier, kind)
            result = message_factory.GetMessageClass(pool.FindMessageTypeByName(name)).FromString(values[0])
            assert not result.FindInitializationErrors()
            return result

        roots = [identifier for identifier, kind in records if kind == 1]
        assert len(roots) == 1
        document = decode(roots[0], 1, 'KN.DocumentArchive')
        show = decode(document.show.identifier, 2, 'KN.ShowArchive')
        assert (show.size.width, show.size.height) == (960, 540)
        metadata_ids = [identifier for identifier, kind in records if kind == 11006]
        assert len(metadata_ids) == 1
        metadata = decode(metadata_ids[0], 11006, 'TSP.PackageMetadata')
        datas = {item.identifier: item.file_name for item in metadata.datas}
        rows = []
        for reference in show.slideTree.slides:
            node = decode(reference.identifier, 4, 'KN.SlideNodeArchive')
            slide = decode(node.slide.identifier, 5, 'KN.SlideArchive')
            assert not node.isSkipped and slide.inDocument
            order = [item.identifier for item in slide.drawables_z_order]
            assert set(order) == {item.identifier for item in slide.owned_drawables}
            drawables = []
            for identifier in order:
                if (identifier, 2011) in records:
                    shape = decode(identifier, 2011, 'TSWP.ShapeInfoArchive')
                    assert shape.is_text_box and shape.owned_storage.identifier == shape.deprecated_storage.identifier
                    storage = decode(shape.owned_storage.identifier, 2001, 'TSWP.StorageArchive')
                    item = {'kind': 'text', 'text': ''.join(storage.text).rstrip('\n')}
                    drawable = shape.super.super
                else:
                    image = decode(identifier, 3005, 'TSD.ImageArchive')
                    assert not image.HasField('mask')
                    data = package.read('Data/' + datas[image.data.identifier])
                    item = {'kind': 'image', 'sha256': sha(data),
                            'description': image.super.accessibility_description}
                    drawable = image.super
                assert drawable.parent.identifier == node.slide.identifier and drawable.geometry.angle == 0
                geometry = drawable.geometry
                item['geometry'] = [geometry.position.x, geometry.position.y, geometry.size.width, geometry.size.height]
                drawables.append(item)
            rows.append({'slide': node.slide.identifier, 'drawables': drawables})
        assert len(rows) == 3 and [len(row['drawables']) for row in rows] == [4, 3, 3]
        return rows, messages


def pptx(path):
    with ZipFile(path) as package:
        def xml(part):
            return ET.fromstring(package.read(part))

        def rels(part):
            name = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
            values = {}
            for item in xml(name):
                assert item.get('TargetMode') != 'External'
                values[item.get('Id')] = posixpath.normpath(posixpath.join(posixpath.dirname(part), item.get('Target'))).lstrip('/')
            return values

        presentation = xml('ppt/presentation.xml')
        size = presentation.find('p:sldSz', NS)
        assert (int(size.get('cx')), int(size.get('cy'))) == (960 * 12700, 540 * 12700)
        relationships, rows = rels('ppt/presentation.xml'), []
        for reference in presentation.findall('p:sldIdLst/p:sldId', NS):
            part = relationships[reference.get('{' + NS['r'] + '}id')]
            slide, related, drawables = xml(part), rels(part), []
            for shape in slide.find('p:cSld/p:spTree', NS):
                tag = shape.tag.rsplit('}', 1)[-1]
                if tag in ('nvGrpSpPr', 'grpSpPr'):
                    continue
                assert tag in ('sp', 'pic'), tag
                transform = shape.find('p:spPr/a:xfrm', NS)
                assert transform is not None and int(transform.get('rot', '0')) == 0
                assert all(transform.get(field, '0') in ('0', 'false') for field in ('flipH', 'flipV'))
                offset, extent = transform.find('a:off', NS), transform.find('a:ext', NS)
                geometry = [int(offset.get('x')) / 12700, int(offset.get('y')) / 12700,
                            int(extent.get('cx')) / 12700, int(extent.get('cy')) / 12700]
                if tag == 'sp':
                    paragraphs = [''.join(node.text or '' for node in item.findall('.//a:t', NS))
                                  for item in shape.findall('p:txBody/a:p', NS)]
                    item = {'kind': 'text', 'text': '\n'.join(paragraphs).rstrip('\n')}
                else:
                    blip = shape.find('p:blipFill/a:blip', NS)
                    data = package.read(related[blip.get('{' + NS['r'] + '}embed')])
                    properties = shape.find('p:nvPicPr/p:cNvPr', NS)
                    item = {'kind': 'image', 'sha256': sha(data), 'description': properties.get('descr', '')}
                drawables.append(dict(item, geometry=geometry))
            rows.append({'drawables': drawables})
        assert len(rows) == 3
        return rows


def verify(args):
    assert version('numbers-parser') == '4.19.0'
    source, messages = native(args.package, args.descriptor)
    exported = pptx(args.pptx)
    expected = ['Edited PNG — image behind text', 'Edited alpha PNG',
                'Native checked JPEG' if args.native_edit else 'Edited JPEG']
    for index, (before, after) in enumerate(zip(source, exported)):
        assert before['drawables'][0]['text'] == expected[index]
        assert len(before['drawables']) == len(after['drawables'])
        for left, right in zip(before['drawables'], after['drawables']):
            for field in left.keys() - {'geometry'}:
                assert left[field] == right[field], (index + 1, field, left, right)
            assert all(abs(a - b) < .001 for a, b in zip(left['geometry'], right['geometry']))
    result = {'nativeSha256': sha(args.package.read_bytes()), 'pptxSha256': sha(args.pptx.read_bytes()),
        'descriptorSha256': DESCRIPTOR_SHA256, 'numbersParser': version('numbers-parser'),
        'protobuf': version('protobuf'), 'recordsIncludingAuxiliary': messages, 'slideSizePoints': [960, 540],
        'slides': source, 'geometryTolerancePoints': .001,
        'allOrderedTextGeometryImageBytesAndDescriptionsMatch': True,
        'limits': 'Selected records in the recorded three-slide profile. Auxiliary/opaque records and full native appearance are unassessed.'}
    args.out.write_text(json.dumps(result, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')
    print('PASS | all three slides: ordered text, geometry, image bytes and accessibility descriptions')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('package', type=Path)
    parser.add_argument('pptx', type=Path)
    parser.add_argument('--descriptor', required=True, type=Path)
    parser.add_argument('--native-edit', action='store_true')
    parser.add_argument('--out', required=True, type=Path)
    verify(parser.parse_args())
