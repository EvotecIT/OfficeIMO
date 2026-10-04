"""Qualify selected native root comments; requires opt-in numbers-parser 4.19.0."""
import argparse
import hashlib
import json
import plistlib
from datetime import datetime, timedelta, timezone
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile
from numbers_parser import Document

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('source', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
if version('numbers-parser') != '4.19.0':
    raise RuntimeError('The extractor requires numbers-parser 4.19.0.')
source_hash = hashlib.sha256(args.source.read_bytes()).hexdigest()
EXPECTED_HASH = '81814eec7d90108595f3a6e41457b980c1fd935e16a74d77bad20ab11006dcaf'
if source_hash != EXPECTED_HASH:
    raise RuntimeError('The source differs from the pinned native fixture.')
document = Document(str(args.source))
model = document._model
cases = []
for sheet in document.sheets:
    for table in sheet.tables:
        for row in range(table.num_rows):
            for column in range(table.num_cols):
                buffer = model.storage_buffer(table._table_id, row, column)
                if not buffer or buffer[0] != 5 or len(buffer) < 12:
                    continue
                flags = int.from_bytes(buffer[8:12], 'little')
                if not flags & (1 << 19):
                    continue
                offset = 12 + sum(16 if bit == 0 else 8 if bit in (1, 2) else 4
                                  for bit in range(19) if flags & (1 << bit))
                key = int.from_bytes(buffer[offset:offset + 4], 'little')
                catalog_id = model.objects[table._table_id].base_data_store.commentStorageTable.identifier
                catalog = model.objects[catalog_id]
                assert catalog.listType == 10
                entry = next(entry for entry in catalog.entries if entry.key == key)
                comment_id = entry.comment_storage.identifier
                comment = model.objects[comment_id]
                assert comment.DESCRIPTOR.full_name == 'TSD.CommentStorageArchive' and not comment.replies
                author_id = comment.author.identifier
                author = model.objects[author_id]
                assert author.DESCRIPTOR.full_name == 'TSK.AnnotationAuthorArchive'
                timestamp = datetime(2001, 1, 1, tzinfo=timezone.utc) + timedelta(seconds=comment.creation_date.seconds)
                cases.append({'sheet': sheet.name, 'table': table.name, 'row': row + 1, 'column': column + 1,
                              'catalogId': catalog_id, 'selector': key, 'commentId': comment_id,
                              'authorId': author_id, 'text': comment.text, 'author': author.name,
                              'appleEpochSeconds': comment.creation_date.seconds,
                              'creationDateUtc': timestamp.isoformat(), 'replyCount': len(comment.replies)})
assert len(cases) == 3
with ZipFile(args.source) as package:
    builds = plistlib.loads(package.read('Metadata/BuildVersionHistory.plist'))
manifest = {'upstream': 'https://github.com/den-frie-vilje/cupertino-files',
            'revision': '6879e6ed49e0eed7e7f393ff4ee558dc2dbec561',
            'upstreamPath': 'fixtures/olekristensen-v26.3-demo06-formulas-round8.numbers',
            'sourceSha256': source_hash, 'extractorVersion': 'numbers-parser 4.19.0',
            'license': 'MIT, copyright 2026 Ole Kristensen', 'buildVersionHistory': builds,
            'qualification': 'Native selected root comments with plain text, display author and Apple-epoch creation time. No replies, export equivalence or visual rendering qualification.',
            'cases': cases}
args.output.write_text(json.dumps(manifest, indent=2) + '\n')
print(f'Extracted {len(cases)} native root comments from {source_hash}')
