"""Extract empty native hidden-state extents and disabled filters with independent schemas.

Requires opt-in numbers-parser 4.19.0. Does not rewrite packages or qualify hidden-position reconstruction.
"""
import argparse
import hashlib
import json
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile
from numbers_parser.iwafile import IWACompressedChunk, get_archive_info_and_remainder
from numbers_parser.generated.TSTArchives_pb2 import TableModelArchive, FilterSetArchive

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('corpus', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
assert version('numbers-parser') == '4.19.0'
sources = []
for name, expected_count in [('nim-iwork/simple.numbers', 1), ('picodocs/sample-v14.4.pages', 3),
                             ('keynotekit/tabledeck-v15.2.1.key', 1)]:
    path = args.corpus / name
    records = {}
    with ZipFile(path) as package:
        for entry in package.namelist():
            if not entry.endswith('.iwa'):
                continue
            data = b''.join(IWACompressedChunk._decompress_all(package.read(entry)))
            while data:
                header, payload = get_archive_info_and_remainder(data)
                offset = 0
                for position, info in enumerate(header.message_infos):
                    if position == 0:
                        assert header.identifier not in records
                        records[header.identifier] = (info.type, payload[offset:offset + info.length])
                    offset += info.length
                data = payload[offset:]
    tables = []
    for identifier, (kind, content) in sorted(records.items()):
        if kind != 6001:
            continue
        model = TableModelArchive.FromString(content)
        assert model.IsInitialized() and model.HasField('hidden_states_owner')
        assert all(getattr(model, field.name) == 0 for field in model.DESCRIPTOR.fields
                   if field.number in (14, 15, 40, 41, 42))
        extents = []
        assert len(model.hidden_states_owner.hidden_states) == 1
        for states in model.hidden_states_owner.hidden_states:
            for direction, extent in enumerate((states.column_hidden_state_extent, states.row_hidden_state_extent)):
                assert extent.row_or_column_direction == direction
                assert len(extent.base_hidden_states) == len(extent.summary_hidden_states) == 0
                assert not extent.needs_to_update_filter_set_for_import
                assert not any(field.number in (5, 7, 9, 10, 11) for field, _ in extent.ListFields())
                assert extent.HasField('filter_set')
                filter_id = extent.filter_set.identifier
                filter_kind, filter_payload = records[filter_id]
                assert filter_kind == 6220
                filter_set = FilterSetArchive.FromString(filter_payload)
                assert filter_set.IsInitialized() and not filter_set.is_enabled
                assert len(filter_set.filter_rules) == len(filter_set.filter_rules_prepivot) == 0
                extents.append({'direction': direction, 'filterIdentifier': filter_id,
                                'filterEnabled': False, 'baseStateCount': 0, 'summaryStateCount': 0})
        tables.append({'modelIdentifier': identifier, 'extents': extents})
    assert len(tables) == expected_count
    sources.append({'path': name, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest(), 'tables': tables})
manifest = {'provider': 'numbers-parser', 'providerVersion': '4.19.0', 'sources': sources,
            'limits': 'Empty native extents and disabled filter sets only. Positive hidden states, positions, active filtering and Apple export equivalence remain unqualified.'}
args.output.write_text(json.dumps(manifest, indent=2) + '\n', encoding='utf-8')
