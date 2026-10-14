"""Verify the bounded third-level Pages reference with independent schemas and Apple exports.

Opt-in evidence tooling only. Requires the pinned numbers-parser descriptors and
Poppler's pdftotext; no OfficeIMO assembly is loaded.
"""
import argparse
import hashlib
import json
import subprocess
from pathlib import Path
from runpy import run_path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

extract = run_path(str(Path(__file__).with_name('extract-pages-list-profile.py')))['extract']


EXPECTED = [
    (0, 7, 'Parent seven'), (1, 1, 'Child one'), (2, 1, 'Grandchild one'),
    (2, 2, 'Grandchild two'), (2, 3, 'Grandchild continued three'),
    (1, 2, 'Child two resets grandchild'), (2, 1, 'Grandchild reset one'),
    (2, 5, 'Grandchild restart five'), (0, 3, 'Separate root three'),
    (2, 6, 'Resumed grandchild six'), (0, 8, 'Parent eight resets descendants'),
    (1, 1, 'Child reset one'), (2, 1, 'Grandchild reset after root one'),
    (0, 4, 'Parent restart four'), (1, 1, 'Child after restart one'),
    (2, 1, 'Grandchild after restart one')]
W = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def attribute_at(entries, offset):
    return next(entry for entry in reversed(entries) if entry['character_index'] <= offset)


def source_items(decoded, maximum_level=2, number_types=('kNumericDecimal',), marker_fonts=('Arial',)):
    records = {str(record['identifier']): record for record in decoded['records']}
    body = records[str(decoded['bodyIdentifier'])]['declaration']
    text = ''.join(body['text'])
    scope, counters, result, offset = None, {}, [], 0
    for raw in text.splitlines(keepends=True):
        selected = attribute_at(body['table_list_style']['entries'], offset)['object']['identifier']
        style = records[selected]['declaration']
        level_data = attribute_at(body['table_para_data']['entries'], offset)
        start_data = attribute_at(body['table_para_starts']['entries'], offset)
        level = level_data['first']
        if style['label_types'][level] == 'kNumber':
            assert 0 <= level <= maximum_level
            assert style['label_types'][:level + 1] == ['kNumber'] * (level + 1)
            number_type = style['number_types'][0]
            assert number_type in number_types
            assert style['number_types'][:level + 1] == [number_type] * (level + 1)
            assert not any(style['tiered_numbers'][:level + 1])
            assert style['indents'][level] == level * 18
            assert style['text_indents'][level] == 1.5 and style['font_name'] in marker_fonts
            start = start_data['first'] if start_data['character_index'] == offset else 0
            if selected != scope:
                scope, counters = selected, {}
                assert start > 0
            number = start or counters.get(level, 0) + 1
            assert 0 < number <= 2147483647
            counters = {depth: value for depth, value in counters.items() if depth < level}
            counters[level] = number
            result.append((level, number, raw.removesuffix('\n')))
        else:
            assert style['label_types'][level] == 'kNone'
        offset += len(raw.encode('utf-16-le')) // 2
    return result


def verify_docx(path, expected, levels=(0, 1, 2)):
    with ZipFile(path) as package:
        body = ET.fromstring(package.read('word/document.xml'))
        numbering = ET.fromstring(package.read('word/numbering.xml'))
    paragraphs = [p for p in body.findall('.//w:p', W) if p.find('w:pPr/w:numPr', W) is not None]
    assert len(paragraphs) == len(expected)
    for paragraph, (level, _, text) in zip(paragraphs, expected):
        assert ''.join(node.text or '' for node in paragraph.findall('.//w:t', W)) == text
        assert int(paragraph.find('w:pPr/w:numPr/w:ilvl', W).attrib[ET.QName(W['w'], 'val')]) == level
    for level in levels:
        choices = numbering.findall(f'w:abstractNum/w:lvl[@w:ilvl="{level}"]', W)
        assert any(indent is not None
                   and indent.attrib.get(ET.QName(W['w'], 'left')) == str((level + 1) * 360)
                   and indent.attrib.get(ET.QName(W['w'], 'hanging')) == '360'
                   for indent in (choice.find('w:pPr/w:ind', W) for choice in choices))


def verify_pdf(path, expected, plain=('Three decimal levels', 'Plain gap', 'End third-level profile')):
    text = subprocess.check_output(['pdftotext', '-layout', str(path), '-'], text=True)
    assert '\f' in text and text.count('\f') == 1
    normalized, previous = ''.join(text.split()), -1
    for _, number, value in expected:
        wanted = ''.join(f'{number}.{value}'.split())
        position = normalized.find(wanted, previous + 1)
        assert position > previous, (path.name, wanted)
        previous = position
    for paragraph in plain:
        assert ''.join(paragraph.split()) in normalized


def verify(root, descriptor):
    result = []
    for suffix, expected in [('', EXPECTED), ('-saved-v15.4', EXPECTED),
                             ('-edited-v15.4', EXPECTED + [(2, 2, 'Native grandchild two')])]:
        stem = 'pages-third-level-created' if not suffix else 'pages-third-level' + suffix
        source = root / (stem + '.pages')
        decoded = extract(source, descriptor)
        actual = source_items(decoded)
        assert actual == expected
        row = {'source': source.name, 'sha256': decoded['sha256'], 'items': actual}
        if suffix:
            verify_docx(root / (stem + '.docx'), expected)
            verify_pdf(root / (stem + '.pdf'), expected)
            row['nativeExportsVerified'] = ['docx', 'pdf']
        result.append(row)
    verify_pdf(root / 'pages-third-level-officeimo-docx-v15.4.pdf', EXPECTED)
    return {'provider': 'numbers-parser', 'providerVersion': '4.19.0',
            'descriptorSha256': hashlib.sha256(descriptor.read_bytes()).hexdigest(),
            'cases': result, 'destination': 'OfficeIMO DOCX reopened and rendered by Pages 15.4',
            'scope': 'Levels 0,1,2, counters and selected indents; no wider typography or producer claim.'}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('root', type=Path)
    parser.add_argument('--descriptor', required=True, type=Path)
    parser.add_argument('--out', required=True, type=Path)
    args = parser.parse_args()
    result = verify(args.root, args.descriptor)
    args.out.write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
    print('Verified three native graphs, two native DOCX/PDF exports and the OfficeIMO DOCX rendering.')
