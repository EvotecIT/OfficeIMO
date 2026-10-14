"""Verify the nine-slot decimal Pages family independently of OfficeIMO.

Opt-in evidence tooling uses pinned Apple schemas, numbers-parser and Poppler.
Creation, native save/edit and exact-range preservation are separate checks.
"""
import argparse
import copy
import hashlib
import json
from pathlib import Path
from runpy import run_path

profile = run_path(str(Path(__file__).with_name('verify-pages-third-level.py')))
preservation = run_path(str(Path(__file__).with_name('verify-native-text-preservation.py')))['verify']
LEVEL_NAMES = ['Root seven', 'Level one', 'Level two', 'Level three', 'Level four',
               'Level five', 'Level six', 'Level seven', 'Level eight']
REPLACEMENTS = {text: text.upper() for text in LEVEL_NAMES}
PLAIN = ('Nine decimal levels', 'Plain gap', 'End nine-level profile')


def expected_items():
    items = [(level, 7 if level == 0 else 1, text) for level, text in enumerate(LEVEL_NAMES)]
    items += [(8, 2, 'Deep continued two'), (8, 3, 'Deep after gap three'), (4, 4, 'Middle restart four')]
    items += [(level, 1, f'Reset level {level}') for level in range(5, 9)]
    items += [(8, 9, 'Deep restart nine'), (0, 3, 'Other root three'),
              (8, 10, 'Resumed deep ten'), (0, 8, 'Root eight')]
    items += [(level, 1, f'Root reset level {level}') for level in range(1, 9)]
    items += [(0, 4, 'Root restart four')]
    items += [(level, 1, f'Restart level {level}') for level in range(1, 9)]
    assert len(items) == 37
    return items


def verify(root, descriptor):
    original = profile['extract'](root / 'pages-nine-level-saved-v15.4.pages', descriptor)
    old_body = next(row for row in original['records'] if row['identifier'] == original['bodyIdentifier'])['declaration']
    changes = [{'Paragraph': number, 'Before': text, 'After': REPLACEMENTS[text]}
               for number, text in enumerate(''.join(old_body['text']).split('\n'), 1) if text in REPLACEMENTS]
    assert len(changes) == 9
    expected = expected_items()
    patched = [(level, number, REPLACEMENTS.get(text, text)) for level, number, text in expected]
    native_changes = {'ROOT SEVEN': 'Root novel', 'LEVEL FOUR': 'Level FOUR', 'LEVEL EIGHT': 'Level EIGHT'}
    edited = [(level, number, native_changes.get(text, text)) for level, number, text in patched]
    edited += [(8, 2, 'Native deep two')]
    rows, decoded = [], []
    cases = [('pages-nine-level-created', expected), ('pages-nine-level-saved-v15.4', expected),
             ('pages-nine-level-preserved', patched), ('pages-nine-level-edited-v15.4', edited)]
    for stem, selected in cases:
        source = root / (stem + '.pages')
        result = profile['extract'](source, descriptor)
        assert profile['source_items'](result, maximum_level=8) == selected
        if stem == 'pages-nine-level-preserved':
            before, after = copy.deepcopy(original['records']), copy.deepcopy(result['records'])
            old = next(row for row in before if row['identifier'] == original['bodyIdentifier'])['declaration']
            new = next(row for row in after if row['identifier'] == result['bodyIdentifier'])['declaration']
            text = ''.join(old['text'])
            for original_text, replacement in REPLACEMENTS.items():
                text = text.replace(original_text, replacement)
            assert ''.join(new['text']) == text
            old.pop('text'); new.pop('text')
            assert before == after
        profile['verify_docx'](root / (stem + '.docx'), selected, levels=range(9))
        profile['verify_docx'](root / (stem + '-officeimo.docx'), selected, levels=range(9))
        profile['verify_pdf'](root / (stem + '.pdf'), selected, plain=PLAIN)
        profile['verify_pdf'](root / (stem + '-officeimo.pdf'), selected, plain=PLAIN)
        rows.append({'source': source.name, 'sha256': result['sha256'], 'numberedItems': len(selected),
                     'levels': list(range(9)), 'nativeAndOfficeimoDocxVerified': True,
                     'nativeAndComposedPdfTextAndCountersVerified': True})
        decoded.append(result)
    profile['verify_pdf'](root / 'pages-nine-level-officeimo-docx-v15.4.pdf', patched, plain=PLAIN)
    return {'provider': 'numbers-parser4.19.0; pinned Apple schemas; Poppler',
            'descriptorSha256': original['descriptorSha256'], 'cases': rows,
            'exactPatch': preservation(root / 'pages-nine-level-saved-v15.4.pages',
                                       root / 'pages-nine-level-preserved.pages', descriptor, changes),
            'receivingApplicationPdfVerified': True,
            'limits': ['OfficeIMO-authored and Apple-saved source; no arbitrary independent-producer coverage.',
                       'Selected one-page Pages15.4 Arial12 decimal body profile; wider typography/pagination remain open.',
                       'Composed PDFs use explicit portable-font substitution; no native typography equivalence claim.',
                       'Native resave updates application metadata; exact preservation applies before resave.']}, decoded


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('root', type=Path)
    parser.add_argument('--descriptor', type=Path, required=True)
    parser.add_argument('--out', type=Path, required=True)
    args = parser.parse_args()
    receipt, decoded = verify(args.root, args.descriptor)
    args.out.write_text(json.dumps(receipt, indent=2) + '\n', encoding='utf-8')
    for item in decoded:
        args.out.with_name(Path(item['source']).stem + '-decode.json').write_text(json.dumps(item, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({'passed': len(receipt['cases']), 'exactPatch': receipt['exactPatch']}, indent=2))
