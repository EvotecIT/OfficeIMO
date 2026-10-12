"""Verify the controlled native Keynote table-border reference independently.

Opt-in evidence only: numbers-parser 4.19.0 and the pinned KN/TST descriptor
used by verify-keynote-creation.py. No OfficeIMO assemblies are loaded.
This checks inherited native declarations and selected sparse edits; it is not
a general-purpose native border renderer or an appearance-equivalence oracle.
"""
import argparse
import hashlib
import json
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile

from google.protobuf import descriptor_pb2, descriptor_pool, message_factory, json_format
from iwork_protobuf_evidence import archive_messages

DESCRIPTOR_SHA256 = '2408cf25920c39fe59d500a0b580bd64bf10a24630c81ca2d543e64322853917'
SOURCE_SHA256 = '2032ce6cdee7f75da8de21d502b47d4e5c7f4eb4d152786a84dd81feda176d45'


def extract(path, descriptor):
    assert version('numbers-parser') == '4.19.0'
    raw = descriptor.read_bytes()
    assert hashlib.sha256(raw).hexdigest() == DESCRIPTOR_SHA256
    assert hashlib.sha256(path.read_bytes()).hexdigest() == SOURCE_SHA256
    pool = descriptor_pool.DescriptorPool()
    for file in descriptor_pb2.FileDescriptorSet.FromString(raw).file:
        pool.Add(file)
    graph = {}
    with ZipFile(path) as package:
        assert package.testzip() is None
        for _, header, info, content in archive_messages(package):
            graph.setdefault((header.identifier, info.type), []).append(content)

    def decode(identifier, kind, name):
        candidates = graph[(identifier, kind)]
        assert len(candidates) == 1, (identifier, kind)
        return message_factory.GetMessageClass(pool.FindMessageTypeByName(name)).FromString(candidates[0])

    models = {identifier: decode(identifier, kind, 'TST.TableModelArchive')
              for identifier, kind in graph if kind == 6001}
    assert len(models) == 3
    selected = [(identifier, model) for identifier, model in models.items() if model.table_name == 'Values']
    assert len(selected) == 1
    identifier, model = selected[0]
    assert (model.number_of_rows, model.number_of_columns) == (4, 4)
    assert (model.number_of_header_rows, model.number_of_header_columns, model.number_of_footer_rows) == (1, 1, 1)

    properties = {}
    chain = []
    current = model.table_style.identifier
    while current:
        assert current not in chain and len(chain) < 64
        chain.append(current)
        style = decode(current, 6003, 'TST.TableStyleArchive')
        current = style.super.parent.identifier if style.super.HasField('parent') else 0
    for current in reversed(chain):
        style = decode(current, 6003, 'TST.TableStyleArchive')
        for field, value in style.table_properties.ListFields():
            properties[field.name] = value
    assert not properties['h_strokes_visible'] and not properties['table_hr_divider_visible']
    for name in ('v_strokes_visible', 'hr_separator_visible', 'hc_separator_visible', 'footer_separator_visible',
                 'table_border_visible', 'table_header_border_visible', 'table_hc_divider_visible', 'table_footer_divider_visible'):
        if name in properties:
            assert properties[name]
    outlines = ('header_row_border_stroke', 'header_column_border_stroke', 'footer_row_border_stroke',
                'table_body_horizontal_border_stroke', 'table_body_vertical_border_stroke')
    for name in outlines:
        assert properties[name].width == 2 and properties[name].pattern.type == 1
    for name in ('header_row_separator_stroke', 'header_column_separator_stroke', 'footer_row_separator_stroke',
                 'header_column_horizontal_stroke', 'table_body_vertical_stroke'):
        assert properties[name].width == 1

    sidecar = decode(model.stroke_sidecar.identifier, 6305, 'TST.StrokeSidecarArchive')
    assert (sidecar.max_order, sidecar.column_count, sidecar.row_count) == (6, 4, 4)
    assert not sidecar.left_column_stroke_layers and not sidecar.top_row_stroke_layers
    assert len(sidecar.right_column_stroke_layers) == len(sidecar.bottom_row_stroke_layers) == 1
    layers = {}
    for name in ('right_column_stroke_layers', 'bottom_row_stroke_layers'):
        layer = decode(getattr(sidecar, name)[0].identifier, 6306, 'TST.StrokeLayerArchive')
        assert layer.row_column_index == 1 and len(layer.stroke_runs) == 1
        run = layer.stroke_runs[0]
        assert (run.origin, run.length) == (1, 1)
        assert list(run.stroke.pattern.pattern) == [0] * 6
        layers[name] = json_format.MessageToDict(layer, preserving_proto_field_name=True)
        if name.startswith('right'):
            assert run.order == 6 and run.stroke.width == 0 and run.stroke.pattern.type == 2
        else:
            assert run.order == 5 and run.stroke.width == 3 and run.stroke.pattern.type == 1
            assert [round(value, 6) for value in (run.stroke.color.r, run.stroke.color.g, run.stroke.color.b, run.stroke.color.a)] == [1, 0.2, 0.2, 1]

    disabled = []
    for table_id, table in models.items():
        owner = decode(table.category_owner.identifier, 6372, 'TST.CategoryOwnerRefArchive')
        for reference in owner.group_by:
            group = decode(reference.identifier, 6373, 'TST.GroupByArchive')
            assert group.HasField('is_enabled') and not group.is_enabled
            disabled.append({'tableModel': table_id, 'group': reference.identifier, 'enabled': False})
    return {'sourceSha256': SOURCE_SHA256, 'descriptorSha256': DESCRIPTOR_SHA256,
            'model': identifier, 'styleChain': chain, 'sidecar': model.stroke_sidecar.identifier,
            'outlinesPoints': {name: properties[name].width for name in outlines},
            'hiddenBodyHorizontalAndHeaderVertical': True, 'layers': layers, 'disabledCategoryGroups': disabled}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('descriptor', type=Path)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    args.output.write_text(json.dumps(extract(args.source, args.descriptor), indent=2) + '\n')
    print('Native table-border defaults, region visibility, ordered sparse edits and disabled category groups verified.')
