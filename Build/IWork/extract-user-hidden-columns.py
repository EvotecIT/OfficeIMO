"""Extract base user-hidden column positions from the pinned independent Numbers corpus.

Requires numbers-parser 4.19.0. The package is not rewritten; native exports remain unqualified.
"""
import argparse
import hashlib
import json
import plistlib
from importlib.metadata import version
from zipfile import ZipFile
from numbers_parser import Document

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('source')
parser.add_argument('output')
args = parser.parse_args()
assert version('numbers-parser') == '4.19.0'
with open(args.source, 'rb') as file:
    source_hash = hashlib.sha256(file.read()).hexdigest()
assert source_hash == '9371c5b1d6ee4dfa17569097f064eba9c67f804d88b48638efbbeeb459d07dd4'
document = Document(args.source)
objects = document._model.objects
tables = []
for sheet in document.sheets:
    for table in sheet.tables:
        model = objects[table._table_id]
        selected = []
        assert len(model.hidden_states_owner.hidden_states) == 1
        states = model.hidden_states_owner.hidden_states[0]
        assert not states.row_hidden_state_extent.base_hidden_states
        assert not states.column_hidden_state_extent.summary_hidden_states
        for state in states.column_hidden_state_extent.base_hidden_states:
            assert state.user_hidden and not state.filtered and not state.pivot_hidden
            selected.append((state.row_or_column_uid.lower, state.row_or_column_uid.upper))
        if not selected:
            continue
        mapping = objects[model.base_column_row_uids.identifier]
        assert len(mapping.sorted_column_uids) == len(mapping.column_index_for_uid) == table.num_cols
        assert sorted(mapping.column_index_for_uid) == list(range(table.num_cols))
        assert all(mapping.column_uid_for_index[index] == ordinal
                   for ordinal, index in enumerate(mapping.column_index_for_uid))
        positions = {(uid.lower, uid.upper): index + 1 for uid, index in
                     zip(mapping.sorted_column_uids, mapping.column_index_for_uid)}
        assert len(positions) == table.num_cols
        hidden = sorted(positions[uid] for uid in selected)
        assert len(hidden) == 2
        assert all(table.cell(row, column - 1).value is None for column in hidden for row in range(table.num_rows))
        assert all(getattr(model, field.name) == 0 for field in model.DESCRIPTOR.fields
                   if field.number in (14, 15, 40, 41, 42))
        tables.append({'modelIdentifier': table._table_id, 'mapIdentifier': model.base_column_row_uids.identifier,
                       'sheet': sheet.name, 'table': table.name, 'rowCount': table.num_rows,
                       'columnCount': table.num_cols, 'hiddenColumns': hidden,
                       'selectedUuids': [{'lower': str(lower), 'upper': str(upper)} for lower, upper in selected]})
assert len(tables) == 3
with ZipFile(args.source) as package:
    builds = plistlib.loads(package.read('Metadata/BuildVersionHistory.plist'))
manifest = {'provider': 'numbers-parser', 'providerVersion': '4.19.0',
            'source': {'path': 'numbers-parser/cross-table-formulas.numbers', 'sha256': source_hash,
                       'buildVersionHistory': builds}, 'tables': tables,
            'limits': 'Base user-hidden columns and zero legacy counts in an unchanged producer fixture. Hidden cells are empty. Positive rows use synthetic proof; filtering, groups and Apple export/render equivalence remain unqualified.'}
with open(args.output, 'w', encoding='utf-8') as file:
    json.dump(manifest, file, indent=2)
    file.write('\n')
