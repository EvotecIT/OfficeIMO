"""Requalify decimal-marker identity without widening native content support.

Opt-in evidence tooling compares every package payload with the preceding
qualified creation fixtures, then decodes fresh Apple saves with pinned schemas.
It checks native exports and strict saved DOCX separately from document identity.
"""
import argparse
import hashlib
import json
import plistlib
from pathlib import Path
from runpy import run_path
from uuid import UUID
from zipfile import ZipFile

sizes = run_path(str(Path(__file__).with_name('verify-pages-mixed-size-lists.py')))
fonts, forms = sizes['fonts'], sizes['forms']
profile, nine = fonts['profile'], forms['nine']


def identity_delta(before, after):
    def entries(path):
        with ZipFile(path) as package:
            return {name: package.read(name) for name in package.namelist()}

    old, new = entries(before), entries(after)
    assert old.keys() == new.keys()
    changed = sorted(name for name in old if old[name] != new[name])
    assert changed == ['Metadata/DocumentIdentifier', 'Metadata/Properties.plist'], changed
    old_id, new_id = (data['Metadata/DocumentIdentifier'].decode('ascii') for data in (old, new))
    assert old_id != new_id and str(UUID(new_id)) == new_id
    old_plist, new_plist = (plistlib.loads(data['Metadata/Properties.plist']) for data in (old, new))
    assert new_plist == {key: value.replace(old_id, new_id) if isinstance(value, str) else value
                         for key, value in old_plist.items()}
    for key in ('documentUUID', 'privateUUID', 'shareUUID', 'stableDocumentUUID', 'versionUUID'):
        assert new_plist[key] == new_id
    assert new_plist['revision'] == '0::' + new_id
    return {'previousSource': before.name,
            'previousSha256': hashlib.sha256(before.read_bytes()).hexdigest(),
            'identifier': new_id, 'changedEntries': changed,
            'unchangedEntries': len(old) - len(changed), 'allIwaArchivesByteIdentical': True,
            'propertiesChangeOnlyMatchingUuidSubstitution': True}


def expected_items(name, edited):
    rows = nine['expected_items']() if name in ('parentheses', 'closing') else fonts['expected_items']()
    return [(level, number, text.replace('Root seven', 'Root native') if edited else text)
            for level, number, text in rows]


def verify(root, descriptor, previous_root):
    cases, decoded, changes = [], [], []
    public_cases = []
    period_hashes = {
        'plain': 'b70d7651d395916396f79d638b1e3d65a7e1e775b105e2e7ab369c2a98a7db77',
        'mixed-font': 'c31f84ad2c73595132ff336d24cc89921e33513dfd5b0689552f6ec6c7cc04b6',
        'mixed-size': '557d306ca88e514ff5a3f25bd17b666cac01df55382db212fcdabffd98d35c3c'}
    for mode in period_hashes:
        identifiers = []
        for style in ('period', 'parentheses', 'closing-parenthesis'):
            path = root / f'pages-identity-{mode}-{style}.pages'
            digest = hashlib.sha256(path.read_bytes()).hexdigest()
            with ZipFile(path) as package:
                identifier = package.read('Metadata/DocumentIdentifier').decode('ascii')
                properties = plistlib.loads(package.read('Metadata/Properties.plist'))
            assert str(UUID(identifier)) == identifier
            assert all(properties[key] == identifier for key in
                       ('documentUUID', 'privateUUID', 'shareUUID', 'stableDocumentUUID', 'versionUUID'))
            assert properties['revision'] == '0::' + identifier
            if style == 'period':
                assert digest == period_hashes[mode]
            identifiers.append(identifier)
            public_cases.append({'source': path.name, 'mode': mode, 'style': style,
                                 'sha256': digest, 'identifier': identifier})
        assert len(set(identifiers)) == 3, (mode, identifiers)
    for name, previous in (
            ('parentheses', 'pages-decimal-parentheses-created.pages'),
            ('closing', 'pages-decimal-closing-created.pages'),
            ('mixed-font', 'pages-mixed-font-created.pages'),
            ('mixed-size', 'pages-mixed-size-created.pages')):
        prefix = 'pages-identity-' + name
        changes.append(identity_delta(previous_root / previous, root / (prefix + '-created.pages')))
        for suffix in ('created', 'edited-v15.4'):
            stem = prefix + '-' + suffix
            edited = suffix != 'created'
            expected = expected_items(name, edited)
            result = profile['extract'](root / (stem + '.pages'), descriptor)
            native_differences = []
            if name in ('parentheses', 'closing'):
                forms['verify_source'](result, expected, name)
                forms['verify_docx'](root / (stem + '-officeimo.docx'), expected, name)
                if edited:
                    forms['verify_docx'](root / (stem + '.docx'), expected, name)
                    forms['verify_pdf'](root / (stem + '.pdf'), expected, name)
            else:
                run_sizes = sizes['run_sizes'] if name == 'mixed-size' else None
                fonts['verify_source'](result, expected, run_sizes)
                fonts['verify_docx_fonts'](root / (stem + '-officeimo.docx'), expected, True, run_sizes)
                if edited:
                    native_differences = fonts['verify_docx_fonts'](root / (stem + '.docx'), expected, False, run_sizes)
                    title = 'Mixed size decimal items' if name == 'mixed-size' else 'Mixed family decimal items'
                    fonts['verify_pdf'](root / (stem + '.pdf'), expected, title)
                    if run_sizes is not None:
                        sizes['verify_marker_sizes'](root / (stem + '.pdf'), expected)
            cases.append({'source': result['source'], 'sha256': result['sha256'],
                          'numberedItems': len(expected), 'strictSavedDocxVerified': True,
                          'nativeExportsVerified': edited,
                          'nativeDocxTypographyDifferences': native_differences})
            decoded.append(result)
    return {'provider': 'numbers-parser 4.19.0; pinned Apple schemas; Poppler; Python plistlib',
            'descriptorSha256': decoded[0]['descriptorSha256'], 'identityChanges': changes,
            'publicIdentityCases': public_cases, 'periodBytesUnchanged': True,
            'cases': cases, 'nativePdfPages': 4,
            'limits': ['Identity requalification retains the existing native content profiles.',
                       'Fresh Pages 15.4 lifecycle edits Root seven to Root native; no insertion is claimed here.',
                       'Native DOCX typography substitutions remain separate from strict OfficeIMO DOCX.',
                       'Broader producers, containers, typography and loaded rich-run editing remain unqualified.']}, decoded


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('root', type=Path)
    parser.add_argument('--descriptor', type=Path, required=True)
    parser.add_argument('--previous-root', type=Path)
    parser.add_argument('--out', type=Path, required=True)
    args = parser.parse_args()
    receipt, decoded = verify(args.root, args.descriptor, args.previous_root or args.root)
    args.out.write_text(json.dumps(receipt, indent=2) + '\n', encoding='utf-8')
    for item in decoded:
        args.out.with_name(Path(item['source']).stem + '-decode.json').write_text(
            json.dumps(item, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({'passed': len(receipt['cases']), 'metadataOnlyCases': len(receipt['identityChanges']),
                      'nativePdfPages': receipt['nativePdfPages']}, indent=2))
