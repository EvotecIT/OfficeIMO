"""Verify fixed-length Pages decimal-item edits independently of OfficeIMO.

Opt-in evidence tooling uses pinned Apple schemas, numbers-parser and Poppler.
Every package entry and IWA message is compared before native resave. Native
resave is checked against the selected list declarations and Apple exports.
"""
import argparse
import copy
import hashlib
import json
from pathlib import Path
from runpy import run_path

profile = run_path(str(Path(__file__).with_name('verify-pages-third-level.py')))
preservation = run_path(str(Path(__file__).with_name('verify-native-text-preservation.py')))['verify']
REPLACEMENTS = {
    'Parent seven': 'Edited seven', 'Child one': 'Child two',
    'Grandchild one': 'Grandchild six', 'Resumed grandchild six': 'Updated grandchild six'}


def sha(data):
    return hashlib.sha256(data).hexdigest()


def verify(root, descriptor, original):
    extract = profile['extract']
    source = extract(original, descriptor)
    body = next(item for item in source['records'] if item['identifier'] == source['bodyIdentifier'])['declaration']
    changes = [{'Paragraph': number, 'Before': text, 'After': REPLACEMENTS[text]}
               for number, text in enumerate(''.join(body['text']).split('\n'), 1) if text in REPLACEMENTS]
    assert len(changes) == len(REPLACEMENTS)
    rows, decoded = [], []
    expected = [(level, number, REPLACEMENTS.get(text, text)) for level, number, text in profile['EXPECTED']]
    for suffix in ('', '-edited'):
        path = root / f'pages-preserved-decimal{suffix}-v15.4.pages'
        result = extract(path, descriptor)
        if not suffix:
            before, after = copy.deepcopy(source['records']), copy.deepcopy(result['records'])
            old_body = next(item for item in before if item['identifier'] == source['bodyIdentifier'])['declaration']
            new_body = next(item for item in after if item['identifier'] == result['bodyIdentifier'])['declaration']
            old_text = ''.join(old_body['text'])
            for old, new in REPLACEMENTS.items():
                old_text = old_text.replace(old, new)
            assert ''.join(new_body['text']) == old_text
            old_body.pop('text'); new_body.pop('text')
            assert before == after
        selected = [(level, number, 'Native seven' if suffix and text == 'Edited seven' else text)
                    for level, number, text in expected]
        assert profile['source_items'](result) == selected
        profile['verify_docx'](root / f'pages-preserved-decimal{suffix}-v15.4.docx', selected)
        profile['verify_pdf'](root / f'pages-preserved-decimal{suffix}-v15.4.pdf', selected)
        profile['verify_pdf'](root / f'pages-preserved-decimal{suffix}-v15.4-officeimo.pdf', selected)
        rows.append({'source': path.name, 'sha256': sha(path.read_bytes()),
                     'numberedItems': len(selected), 'levels': sorted({row[0] for row in selected}),
                     'nativeDocxAndPdfMatched': True, 'composedPdfTextAndCountersMatched': True})
        decoded.append(result)
    profile['verify_pdf'](root / 'pages-preserved-decimal-officeimo-docx-v15.4.pdf', expected)
    return {'provider': 'numbers-parser4.19.0; pinned Apple schemas; Poppler',
            'original': original.name, 'originalSha256': sha(original.read_bytes()),
            'descriptorSha256': source['descriptorSha256'],
            'exactPatch': preservation(original, root / 'pages-preserved-decimal-v15.4.pages', descriptor, changes),
            'cases': rows, 'receivingApplicationPdfMatched': True,
            'limits': ['Selected one-page Arial12 decimal body profile; no broader typography or pagination claim.',
                       'Composed PDFs use explicit portable-font substitution; text/counters do not prove native typography.',
                       'Native resave may update unrelated application metadata; exact preservation applies before resave.']}, decoded


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('root', type=Path)
    parser.add_argument('--original', type=Path)
    parser.add_argument('--descriptor', type=Path, required=True)
    parser.add_argument('--out', type=Path, required=True)
    args = parser.parse_args()
    receipt, decoded = verify(args.root, args.descriptor, args.original or args.root / 'pages-third-level-saved-v15.4.pages')
    args.out.write_text(json.dumps(receipt, indent=2) + '\n', encoding='utf-8')
    for item in decoded:
        args.out.with_name(Path(item['source']).stem + '-decode.json').write_text(json.dumps(item, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({'passed': len(receipt['cases']), 'exactPatch': receipt['exactPatch']}, indent=2))
