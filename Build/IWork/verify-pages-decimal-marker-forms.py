"""Verify the two parenthesis decimal marker forms independently of OfficeIMO.

Opt-in evidence tooling uses pinned Apple schemas, numbers-parser and Poppler.
Creation, Apple save/edit, exact text preservation and receiving-app layout are
separate checks. No OfficeIMO assembly is loaded.
"""
import argparse
import json
import subprocess
from pathlib import Path
from runpy import run_path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

nine = run_path(str(Path(__file__).with_name('verify-pages-nine-levels.py')))
profile = nine['profile']
W = profile['W']
VAL = '{' + W['w'] + '}val'


def marker(number, text, form):
    if text == 'Other root three':
        return str(number) + '.'
    return f'({number})' if form == 'parentheses' else str(number) + ')'


def verify_source(decoded, expected, form):
    number_type = 'kNumericDoubleParen' if form == 'parentheses' else 'kNumericRightParen'
    assert profile['source_items'](decoded, maximum_level=8,
                                  number_types=('kNumericDecimal', number_type)) == expected
    records = {str(row['identifier']): row['declaration'] for row in decoded['records']}
    body = records[str(decoded['bodyIdentifier'])]
    offset, selected = 0, []
    for raw in ''.join(body['text']).splitlines(keepends=True):
        style_id = profile['attribute_at'](body['table_list_style']['entries'], offset)['object']['identifier']
        style = records[style_id]
        level = profile['attribute_at'](body['table_para_data']['entries'], offset)['first']
        if style['label_types'][level] == 'kNumber':
            text = raw.removesuffix('\n')
            assert style['number_types'][level] == ('kNumericDecimal' if text == 'Other root three' else number_type)
            selected.append(text)
        offset += len(raw.encode('utf-16-le')) // 2
    assert selected == [text for _, _, text in expected]


def verify_docx(path, expected, form):
    verify_docx_markers(path, expected, lambda number, text: marker(number, text, form))


def verify_docx_markers(path, expected, marker_for):
    profile['verify_docx'](path, expected, levels=range(9))
    with ZipFile(path) as package:
        body = ET.fromstring(package.read('word/document.xml'))
        numbering = ET.fromstring(package.read('word/numbering.xml'))
        styles = ET.fromstring(package.read('word/styles.xml'))
    nums = {node.get('{' + W['w'] + '}numId'): node for node in numbering.findall('w:num', W)}
    abstracts = {node.get('{' + W['w'] + '}abstractNumId'): node
                 for node in numbering.findall('w:abstractNum', W)}
    style_nums = {node.get('{' + W['w'] + '}styleId'): node.find('w:pPr/w:numPr/w:numId', W).get(VAL)
                  for node in styles.findall('w:style', W) if node.find('w:pPr/w:numPr/w:numId', W) is not None}

    def resolve(num_id, level, seen=()):
        assert num_id not in seen
        num = nums[num_id]
        override = num.find(f'w:lvlOverride[@w:ilvl="{level}"]/w:lvl', W)
        if override is not None:
            return override
        abstract = abstracts[num.find('w:abstractNumId', W).get(VAL)]
        node = abstract.find(f'w:lvl[@w:ilvl="{level}"]', W)
        if node is not None:
            return node
        link = abstract.find('w:numStyleLink', W)
        assert link is not None
        return resolve(style_nums[link.get(VAL)], level, seen + (num_id,))

    paragraphs = [p for p in body.findall('.//w:p', W) if p.find('w:pPr/w:numPr', W) is not None]
    for paragraph, (level, _, text) in zip(paragraphs, expected):
        num_id = paragraph.find('w:pPr/w:numPr/w:numId', W).get(VAL)
        node = resolve(num_id, level)
        assert node.find('w:numFmt', W).get(VAL) == 'decimal'
        pattern = marker_for('%' + str(level + 1), text)
        assert node.find('w:lvlText', W).get(VAL) == pattern, (path.name, text, pattern)


def verify_pdf(path, expected, form):
    text = subprocess.check_output(['pdftotext', '-layout', str(path), '-'], text=True)
    assert text.count('\f') == 1, path.name
    normalized, previous = ''.join(text.split()), -1
    for _, number, value in expected:
        wanted = ''.join((marker(number, value, form) + value).split())
        position = normalized.find(wanted, previous + 1)
        assert position > previous, (path.name, wanted)
        previous = position
    for paragraph in nine['PLAIN']:
        assert ''.join(paragraph.split()) in normalized


def word_bounds(path):
    data = subprocess.check_output(['pdftotext', '-bbox', str(path), '-'])
    root = ET.fromstring(data)
    ns = {'x': 'http://www.w3.org/1999/xhtml'}
    pages = root.findall('.//x:page', ns)
    assert len(pages) == 1
    return [(word.text, [float(word.get(key)) for key in ('xMin', 'yMin', 'xMax', 'yMax')])
            for word in pages[0].findall('.//x:word', ns)]


def verify(root, descriptor):
    rows, decoded, geometry, patches = [], [], [], []
    original_items = nine['expected_items']()
    patched = [(level, number, nine['REPLACEMENTS'].get(text, text)) for level, number, text in original_items]
    native_changes = {'ROOT SEVEN': 'Root novel', 'LEVEL EIGHT': 'Level EIGHT'}
    edited = [(level, number, native_changes.get(text, text)) for level, number, text in patched]
    edited += [(8, 2, 'Native deep two')]
    for form in ('parentheses', 'closing'):
        prefix = 'pages-decimal-' + form
        cases = [('created', original_items), ('saved-v15.4', original_items),
                 ('preserved', patched), ('edited-v15.4', edited)]
        for suffix, expected in cases:
            stem = prefix + '-' + suffix
            result = profile['extract'](root / (stem + '.pages'), descriptor)
            verify_source(result, expected, form)
            verify_docx(root / (stem + '-officeimo.docx'), expected, form)
            verify_pdf(root / (stem + '-officeimo.pdf'), expected, form)
            if suffix != 'created':
                verify_docx(root / (stem + '.docx'), expected, form)
                verify_pdf(root / (stem + '.pdf'), expected, form)
            rows.append({'source': result['source'], 'sha256': result['sha256'], 'form': form,
                         'numberedItems': len(expected), 'levels': list(range(9)),
                         'nativeExportsVerified': suffix != 'created',
                         'savedOfficeimoDocxAndComposedPdfVerified': True})
            decoded.append(result)
        original = next(item for item in decoded if item['source'] == prefix + '-saved-v15.4.pages')
        body = next(row['declaration'] for row in original['records'] if row['identifier'] == original['bodyIdentifier'])
        changes = [{'Paragraph': number, 'Before': text, 'After': nine['REPLACEMENTS'][text]}
                   for number, text in enumerate(''.join(body['text']).split('\n'), 1) if text in nine['REPLACEMENTS']]
        assert len(changes) == 9
        patches.append(nine['preservation'](root / (prefix + '-saved-v15.4.pages'),
                                           root / (prefix + '-preserved.pages'), descriptor, changes))
        receiving = root / (prefix + '-officeimo-docx-v15.4.pdf')
        verify_pdf(receiving, patched, form)
        native_words = word_bounds(root / (prefix + '-preserved.pdf'))
        receiving_words = word_bounds(receiving)
        assert len(native_words) == len(receiving_words) and len(native_words) > 100
        assert [text for text, _ in native_words] == [text for text, _ in receiving_words]
        delta = max(abs(a - b) for (_, native), (_, received) in zip(native_words, receiving_words)
                    for a, b in zip(native, received))
        assert delta <= 0.01, (form, delta)
        geometry.append({'form': form, 'wordBounds': len(native_words), 'maximumDeltaPoints': delta,
                         'nativePdf': prefix + '-preserved.pdf', 'receivingPdf': receiving.name})
    return {'provider': 'numbers-parser4.19.0; pinned Apple schemas; Poppler',
            'descriptorSha256': decoded[0]['descriptorSha256'], 'cases': rows,
            'exactTextPatches': patches, 'receivingApplicationGeometry': geometry,
            'limits': ['OfficeIMO-authored and Apple-saved sources; no arbitrary independent-producer coverage.',
                       'Pages15.4 Arial12 single-page body profile; wider fonts, containers and pagination remain open.',
                       'Uniform non-tiered decimal marker form per list; tiered and mixed ancestor forms remain partial.',
                       'Portable composed PDFs substitute fonts explicitly; no native typography equivalence claim.',
                       'Exact byte preservation applies before Apple resave, which updates application metadata.']}, decoded


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
    print(json.dumps({'passed': len(receipt['cases']), 'patches': len(receipt['exactTextPatches']),
                      'geometry': receipt['receivingApplicationGeometry']}, indent=2))
