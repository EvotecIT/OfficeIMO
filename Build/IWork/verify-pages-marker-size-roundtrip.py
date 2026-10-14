"""Verify the bounded Pages marker-size roundtrip independently of OfficeIMO.

Opt-in tooling uses pinned Apple schemas, numbers-parser and Poppler. It checks
native declarations, controlled DOCX variants and complete receiving PDF bounds.
"""
import argparse
import json
import re
import subprocess
from pathlib import Path
from runpy import run_path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

profile = run_path(str(Path(__file__).with_name('verify-pages-third-level.py')))
forms = run_path(str(Path(__file__).with_name('verify-pages-decimal-marker-forms.py')))
W = profile['W']
VAL = '{' + W['w'] + '}val'
NUMBERS = [None, 7, 8, None, 9, 1, 2, 10, 1, 5, 6, None, 7, 11, 3, 4]
LEVELS = [-1, 0, 0, -1, 0, 1, 1, 0, 1, 1, 1, -1, 1, 0, 0, 0]


def verify_source(decoded, texts, default_size):
    records = {str(row['identifier']): row['declaration'] for row in decoded['records']}
    body = records[str(decoded['bodyIdentifier'])]

    def properties(identifier, seen=()):
        assert identifier not in seen
        row = records[identifier]
        parent = row.get('super', {}).get('parent', {}).get('identifier')
        result = properties(parent, seen + (identifier,)) if parent is not None else {}
        for key, value in row.items():
            result[key] = dict(result.get(key, {}), **value) if isinstance(value, dict) else value
        return result

    def at(table, offset):
        return profile['attribute_at'](body[table]['entries'], offset)

    def selected(table, offset):
        return next(row['object']['identifier'] for row in reversed(body[table]['entries'])
                    if row['character_index'] <= offset and 'object' in row)

    lines = ''.join(body['text']).splitlines(keepends=True)
    assert [line.removesuffix('\n') for line in lines] == texts
    offset, scopes, items = 0, {}, []
    for index, raw in enumerate(lines):
        style_id = selected('table_list_style', offset)
        style = properties(style_id)
        level = at('table_para_data', offset)['first']
        if NUMBERS[index] is not None:
            assert level == LEVELS[index] and style['label_types'][level] == 'kNumber'
            assert style['number_types'][level] == 'kNumericDecimal'
            assert not style['tiered_numbers'][level]
            start = at('table_para_starts', offset)
            explicit = start['first'] if start['character_index'] == offset else 0
            counters = scopes.setdefault(style_id, {})
            number = explicit or counters.get(level, 0) + 1
            assert number == NUMBERS[index]
            scopes[style_id] = {depth: value for depth, value in counters.items() if depth < level}
            scopes[style_id][level] = number
            defaults = properties(selected('table_para_style', offset))['char_properties']
            character = dict(defaults, **properties(selected('table_char_style', offset)).get('char_properties', {}))
            assert defaults['font_size'] == default_size and defaults['font_name'] == 'Calibri'
            assert character['font_size'] == 12 and character['font_name'] == 'ArialMT'
            geometry = style['geometries'][level]
            assert abs(geometry['scale'] * 12 - 11) < 1e-5
            assert geometry['baseline_offset'] == 0 and geometry['scale_with_text']
            assert style['indents'][level] == level * 18
            assert style['text_indents'][level] * default_size == 16.5
            items.append({'text': texts[index], 'level': level, 'number': number,
                          'paragraphDefaultPoints': default_size, 'bodyPoints': 12,
                          'markerScale': geometry['scale'], 'markerPoints': 11,
                          'textOffsetPoints': 16.5})
        else:
            assert style['label_types'][level] == 'kNone'
        offset += len(raw.encode('utf-16-le')) // 2
    assert len(items) == 13
    return items


def verify_docx(path, texts, leading_size, mark_size):
    with ZipFile(path) as package:
        body = ET.fromstring(package.read('word/document.xml'))
        numbering = ET.fromstring(package.read('word/numbering.xml'))
    paragraphs = body.findall('w:body/w:p', W)
    assert [''.join(p.itertext()) for p in paragraphs] == texts
    nums = {row.get('{' + W['w'] + '}numId'): row for row in numbering.findall('w:num', W)}
    abstracts = {row.get('{' + W['w'] + '}abstractNumId'): row for row in numbering.findall('w:abstractNum', W)}
    for index, paragraph in enumerate(paragraphs):
        if NUMBERS[index] is None:
            continue
        selected = paragraph.find('w:pPr/w:numPr', W)
        level = selected.find('w:ilvl', W).get(VAL)
        assert int(level) == LEVELS[index]
        abstract = nums[selected.find('w:numId', W).get(VAL)].find('w:abstractNumId', W).get(VAL)
        marker = abstracts[abstract].find(f'w:lvl[@w:ilvl="{level}"]', W)
        assert marker.find('w:rPr/w:sz', W).get(VAL) == '22'
        assert marker.find('w:rPr/w:rFonts', W).get('{' + W['w'] + '}ascii') == 'Calibri'
        assert marker.find('w:lvlText', W).get(VAL) == '%' + str(int(level) + 1) + '.'
        assert paragraph.find('w:pPr/w:rPr/w:sz', W).get(VAL) == str(mark_size * 2)
        runs = paragraph.findall('w:r', W)
        assert runs[0].find('w:rPr/w:sz', W).get(VAL) == str(leading_size * 2)
        assert not ''.join(runs[0].itertext())
        assert all(run.find('w:rPr/w:sz', W).get(VAL) == '24' for run in runs if ''.join(run.itertext()))
        indent = paragraph.find('w:pPr/w:ind', W)
        assert indent.get('{' + W['w'] + '}left') == ('330' if level == '0' else '690')
        assert indent.get('{' + W['w'] + '}hanging') == '330'


def pdf_markers(path, size):
    xml = subprocess.check_output(['pdftohtml', '-xml', '-stdout', '-zoom', '1', '-hidden', str(path)])
    pages = ET.fromstring(xml).findall('page')
    assert len(pages) == 1
    fonts = {row.get('id'): row for row in pages[0].findall('fontspec')}
    markers = []
    for row in pages[0].findall('text'):
        if re.fullmatch(r'\d+\.', ''.join(row.itertext()).strip()):
            font = fonts[row.get('font')]
            assert float(font.get('size')) == size and font.get('family').endswith('+Calibri')
            markers.append(''.join(row.itertext()).strip())
    assert markers == [str(number) + '.' for number in NUMBERS if number is not None]
    return {'markerCount': len(markers), 'markerPoints': size}


def verify(root, descriptor):
    texts = json.loads((root / 'pages-list-continuation-v15.4.json').read_text())['inputRecipe']['paragraphs']
    decoded, sources = [], []
    for stem, default in [('pages-list-continuation-v15.4', 11), ('pages-list-marker-size-roundtrip-v15.4', 12)]:
        result = profile['extract'](root / (stem + '.pages'), descriptor)
        sources.append({'source': result['source'], 'sha256': result['sha256'],
                        'items': verify_source(result, texts, default)})
        decoded.append(result)
    for suffix, leading, mark in [('current', 11, 11), ('mark-only', 11, 12),
                                  ('mark-and-empty-run', 12, 12), ('final', 12, 12)]:
        verify_docx(root / ('pages-list-marker-size-' + suffix + '.docx'), texts, leading, mark)
    with ZipFile(root / 'pages-list-marker-size-current.docx') as old, ZipFile(root / 'pages-list-marker-size-final.docx') as new:
        before, after = ET.fromstring(old.read('word/document.xml')), ET.fromstring(new.read('word/document.xml'))
        changed = 0
        for paragraph in before.findall('w:body/w:p', W):
            if paragraph.find('w:pPr/w:numPr', W) is not None:
                paragraph.find('w:pPr/w:rPr/w:sz', W).set(VAL, '24')
                paragraph.find('w:r/w:rPr/w:sz', W).set(VAL, '24')
                changed += 1
        assert changed == 13
        assert [ET.tostring(p) for p in before.findall('w:body/w:p', W)] == [
            ET.tostring(p) for p in after.findall('w:body/w:p', W)]
    reference = forms['word_bounds'](root / 'pages-list-continuation-v15.4.pdf')
    pdfs = []
    for filename, size in [('pages-list-continuation-v15.4.pdf', 11),
                           ('pages-list-marker-size-current-native.pdf', 12),
                           ('pages-list-marker-size-mark-only-native.pdf', 12),
                           ('pages-list-marker-size-mark-and-empty-run-native.pdf', 11),
                           ('pages-list-marker-size-roundtrip-v15.4.pdf', 11)]:
        path = root / filename
        actual = forms['word_bounds'](path)
        assert [row[0] for row in actual] == [row[0] for row in reference] and len(actual) == 55
        delta = max(abs(a - b) for (_, left), (_, right) in zip(reference, actual) for a, b in zip(left, right))
        start_delta = max(abs(left[0] - right[0]) for (_, left), (_, right) in zip(reference, actual))
        if size == 11:
            assert delta <= 0.01
        else:
            assert delta > 1
        if filename == 'pages-list-marker-size-current-native.pdf':
            assert start_delta <= 0.01
        pdfs.append({'path': filename, **pdf_markers(path, size), 'wordBounds': len(actual),
                     'maximumNativeWordBoundsDeltaPoints': delta, 'maximumWordStartDeltaPoints': start_delta})
    return {'provider': 'numbers-parser4.19.0; pinned Apple schemas; Poppler',
            'descriptorSha256': decoded[0]['descriptorSha256'], 'sources': sources,
            'strictDocxVariants': 4, 'finalParagraphChanges': 'Only 13 leading-run sizes and 13 paragraph-mark sizes:11 to12 points. Generated section revision and relationship IDs are outside this comparison.',
            'pdfs': pdfs, 'limits': ['Two-level Pages15.4 body lists; Calibri11 markers and ArialMT12 body text.',
                                    'Other scales, fonts, containers and producer versions require separate native evidence.']}, decoded


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('root', type=Path)
    parser.add_argument('--descriptor', type=Path, required=True)
    parser.add_argument('--out', type=Path, required=True)
    args = parser.parse_args()
    receipt, decoded = verify(args.root, args.descriptor)
    args.out.write_text(json.dumps(receipt, indent=2) + '\n', encoding='utf-8')
    for result in decoded:
        args.out.with_name(Path(result['source']).stem + '-decode.json').write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({'sources': len(receipt['sources']), 'pdfs': len(receipt['pdfs']),
                      'finalGeometry': receipt['pdfs'][-1]}, indent=2))
