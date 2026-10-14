"""Verify mixed-family decimal items independently of OfficeIMO.

Opt-in evidence tooling uses pinned Apple schemas, numbers-parser and Poppler.
Source declarations, native export, reconstructed DOCX and PDF appearance have
separate evidence. No OfficeIMO assembly or font binary is redistributed here.
"""
import argparse
import json
import subprocess
from pathlib import Path
from runpy import run_path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

forms = run_path(str(Path(__file__).with_name('verify-pages-decimal-marker-forms.py')))
profile = forms['profile']
W, VAL = forms['W'], forms['VAL']
FAMILIES = ('Arial', 'Times New Roman', 'Courier New')


def expected_items(edited=False):
    rows = [(0, 7, 'Root native' if edited else 'Root seven')]
    rows += [(level, 1, 'Deep native' if edited and level == 8 else f'Depth {level}')
             for level in range(1, 9)]
    rows += [(8, 2, 'Deep two'), (0, 8, 'Root eight'), (1, 1, 'Child one'),
             (0, 3, 'Serif marker'), (1, 1, 'Serif child'),
             (0, 5, 'Mono marker'), (1, 1, 'Mono child'), (1, 2, 'Resumed child')]
    result = [(level, number, text + ' serif mono plain') for level, number, text in rows]
    if edited:
        result.append((1, 3, 'Native child plain'))
    return result


def marker(number, text):
    return f'({number})' if text.startswith('Serif ') else str(number) + (')' if text.startswith('Mono ') else '.')


def marker_font(text):
    return 'Times New Roman' if text.startswith('Serif ') else 'Courier New' if text.startswith('Mono ') else 'Arial'


def segments(text):
    if text == 'Native child plain':
        return [(text, 'Arial', False, False, '000000')]
    return [(text.removesuffix(' mono plain'), 'Times New Roman', True, False, '003366'),
            (' mono', 'Courier New', False, True, '008000'), (' plain', 'Arial', False, False, '000000')]


def verify_source(decoded, expected, run_sizes=None):
    assert profile['source_items'](decoded, maximum_level=8,
        number_types=('kNumericDecimal', 'kNumericDoubleParen', 'kNumericRightParen'),
        marker_fonts=FAMILIES) == expected
    records = {str(row['identifier']): row['declaration'] for row in decoded['records']}
    body = records[str(decoded['bodyIdentifier'])]

    def properties(identifier, seen=()):
        assert identifier not in seen
        row = records[identifier]
        parent = row.get('super', {}).get('parent', {}).get('identifier')
        inherited = properties(parent, seen + (identifier,)) if parent is not None else {}
        return dict(inherited, **row.get('char_properties', {}))

    def at(table, offset):
        return profile['attribute_at'](body[table]['entries'], offset)

    def object_at(table, offset):
        # Native missing-object entries retain the preceding selected style.
        return next(row['object']['identifier'] for row in reversed(body[table]['entries'])
                    if row['character_index'] <= offset and 'object' in row)

    def character_properties(offset):
        return properties(object_at('table_char_style', offset))

    offset, selected = 0, 0
    for raw in ''.join(body['text']).splitlines(keepends=True):
        text = raw.removesuffix('\n')
        style = records[at('table_list_style', offset)['object']['identifier']]
        level = at('table_para_data', offset)['first']
        if style['label_types'][level] == 'kNumber':
            assert text == expected[selected][2]
            number_type = 'kNumericDoubleParen' if text.startswith('Serif ') else 'kNumericRightParen' if text.startswith('Mono ') else 'kNumericDecimal'
            assert style['number_types'][level] == number_type
            assert style['font_name'] == marker_font(text)
            if run_sizes is not None:
                geometry = style['geometries'][level]
                assert geometry['scale'] == 1 and geometry['baseline_offset'] == 0 and geometry['scale_with_text']
            defaults = properties(object_at('table_para_style', offset))
            assert defaults['font_name'] == marker_font(text) and defaults['font_size'] == 12
            assert not defaults['bold'] and not defaults['italic']
            assert all(defaults['font_color'][component] == 0 for component in ('r', 'g', 'b'))
            cursor = offset
            sizes = run_sizes(text) if run_sizes is not None else (12,) * len(segments(text))
            for (value, family, bold, italic, color), size in zip(segments(text), sizes):
                end = cursor + len(value.encode('utf-16-le')) // 2
                boundaries = {cursor} | {row['character_index'] for row in body['table_char_style']['entries']
                                         if cursor < row['character_index'] < end}
                for boundary in boundaries:
                    char = character_properties(boundary)
                    assert char['font_name'] == family and char['font_size'] == size
                    assert char['bold'] == bold and char['italic'] == italic
                    actual = ''.join(f'{round(char["font_color"][component] * 255):02X}' for component in ('r', 'g', 'b'))
                    assert actual == color and char['font_color']['a'] == 1
                cursor = end
            assert cursor == offset + len(text.encode('utf-16-le')) // 2
            selected += 1
        offset += len(raw.encode('utf-16-le')) // 2
    assert selected == len(expected)


def verify_docx_fonts(path, expected, strict, run_sizes=None):
    forms['verify_docx_markers'](path, expected, marker)
    with ZipFile(path) as package:
        body = ET.fromstring(package.read('word/document.xml'))
        styles = ET.fromstring(package.read('word/styles.xml'))
    catalog = {node.get('{' + W['w'] + '}styleId'): node for node in styles.findall('w:style', W)}

    def overlays(node):
        if node is None:
            return {}
        fonts = node.find('w:rFonts', W)
        result = {}
        if fonts is not None:
            result['family'] = fonts.get('{' + W['w'] + '}ascii') or fonts.get('{' + W['w'] + '}hAnsi')
        size = node.find('w:sz', W)
        if size is not None:
            result['size'] = float(size.get(VAL)) / 2
        for tag in ('b', 'i'):
            value = node.find('w:' + tag, W)
            if value is not None:
                result[tag] = value.get(VAL, 'true') not in ('false', '0', 'off')
        return result

    def style_properties(identifier, seen=()):
        assert identifier not in seen
        node = catalog[identifier]
        parent = node.find('w:basedOn', W)
        inherited = style_properties(parent.get(VAL), seen + (identifier,)) if parent is not None else {}
        return dict(inherited, **overlays(node.find('w:rPr', W)))

    paragraphs = [p for p in body.findall('.//w:p', W) if p.find('w:pPr/w:numPr', W) is not None]
    differences = set()
    for paragraph, (_, _, text) in zip(paragraphs, expected):
        if strict and run_sizes is not None:
            mark_size = paragraph.find('w:pPr/w:rPr/w:sz', W)
            assert mark_size is not None and float(mark_size.get(VAL)) / 2 == run_sizes(text)[0], (
                path.name, text, 'paragraph mark would rescale the native marker')
        defaults = overlays(styles.find('w:docDefaults/w:rPrDefault/w:rPr', W))
        paragraph_style = paragraph.find('w:pPr/w:pStyle', W)
        if paragraph_style is not None:
            defaults.update(style_properties(paragraph_style.get(VAL)))
        actual = []
        for run in paragraph.findall('w:r', W):
            value = ''.join(node.text or '' for node in run.findall('w:t', W))
            inherited = dict(defaults)
            run_style = run.find('w:rPr/w:rStyle', W)
            if run_style is not None:
                inherited.update(style_properties(run_style.get(VAL)))
            inherited.update(overlays(run.find('w:rPr', W)))
            actual.extend([inherited] * len(value))
        assert len(actual) == len(text)
        cursor = 0
        sizes = run_sizes(text) if run_sizes is not None else (None,) * len(segments(text))
        for (value, family, bold, italic, _), size in zip(segments(text), sizes):
            for properties in actual[cursor:cursor + len(value)]:
                if size is not None:
                    assert properties.get('size') == size, (path.name, text, size, properties)
                if (properties.get('family') != family or properties.get('b', False) != bold
                        or properties.get('i', False) != italic):
                    assert not strict, (path.name, text, properties)
                    differences.add((text, family, properties.get('family'), bold,
                                     properties.get('b', False), italic, properties.get('i', False)))
            cursor += len(value)
    return [{'text': text, 'expectedFamily': expected_family, 'actualFamily': actual_family,
             'expectedBold': bold, 'actualBold': actual_bold, 'expectedItalic': italic, 'actualItalic': actual_italic}
            for text, expected_family, actual_family, bold, actual_bold, italic, actual_italic in sorted(differences)]


def verify_pdf(path, expected, title='Mixed family decimal items'):
    text = subprocess.check_output(['pdftotext', '-layout', str(path), '-'], text=True)
    assert text.count('\f') == 1, path.name
    normalized, previous = ''.join(text.split()), -1
    for _, number, value in expected:
        wanted = ''.join((marker(number, value) + value).split())
        position = normalized.find(wanted, previous + 1)
        assert position > previous, (path.name, wanted)
        previous = position
    for plain in (title, 'Plain gap retains counters'):
        assert ''.join(plain.split()) in normalized
    return subprocess.check_output(['pdffonts', str(path)], text=True).splitlines()[2:]


def verify(root, descriptor):
    rows, decoded, pdfs = [], [], []
    for suffix in ('created', 'saved-v15.4', 'edited-v15.4'):
        expected = expected_items(suffix == 'edited-v15.4')
        stem = 'pages-mixed-font-' + suffix
        result = profile['extract'](root / (stem + '.pages'), descriptor)
        verify_source(result, expected)
        verify_docx_fonts(root / (stem + '-officeimo.docx'), expected, strict=True)
        differences = []
        if suffix != 'created':
            differences = verify_docx_fonts(root / (stem + '.docx'), expected, strict=False)
            pdfs.append({'path': stem + '.pdf', 'fonts': verify_pdf(root / (stem + '.pdf'), expected)})
        for lane in ('officeimo', 'caller-fonts'):
            name = stem + '-' + lane + '.pdf'
            pdfs.append({'path': name, 'fonts': verify_pdf(root / name, expected)})
        rows.append({'source': result['source'], 'sha256': result['sha256'], 'numberedItems': len(expected),
                     'sourceFontFamilies': list(FAMILIES), 'strictSavedDocxFontsVerified': True,
                     'nativeDocxTypographyDifferences': differences})
        decoded.append(result)
    name = 'pages-mixed-font-officeimo-docx-v15.4.pdf'
    pdfs.append({'path': name, 'fonts': verify_pdf(root / name, expected_items())})
    native = forms['word_bounds'](root / 'pages-mixed-font-saved-v15.4.pdf')
    receiving = forms['word_bounds'](root / name)
    assert [row[0] for row in native] == [row[0] for row in receiving]
    maximum = max(abs(a - b) for (_, left), (_, right) in zip(native, receiving) for a, b in zip(left, right))
    assert maximum <= 0.01, ('Receiving DOCX geometry exceeds 0.01 points', maximum)
    return {'provider': 'numbers-parser4.19.0; pinned Apple schemas; Poppler',
            'descriptorSha256': decoded[0]['descriptorSha256'], 'cases': rows, 'pdfs': pdfs,
            'receivingDocxGeometry': {'wordBounds': len(native), 'maximumDeltaPoints': maximum},
            'limits': ['OfficeIMO-authored and Apple-saved sources; no arbitrary independent-producer coverage.',
                       'One-page body lists at twelve points; mixed sizes, further font families and containers remain unqualified.',
                       'Portable PDF substitutions and embedded caller-font evidence are separate lanes; standalone font files are not redistributed.',
                       'Native DOCX font-family and emphasis differences do not qualify typography equivalence of that export.',
                       'Loaded rich-run editing, native visual-cache regeneration and other runtime architectures retain separate gates.']}, decoded


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
    print(json.dumps({'passed': len(receipt['cases']), 'pdfs': len(receipt['pdfs']),
                      'receivingDocxGeometry': receipt['receivingDocxGeometry']}, indent=2))
