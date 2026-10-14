"""Qualify whole-point mixed-size decimal items against Apple schemas and exports.

Opt-in evidence tooling shares the existing decimal/font verification helpers.
Native marker sizes, character sizes, counters and receiving-DOCX geometry are
checked separately from portable and caller-font PDF conversions.
"""
import argparse
import json
import subprocess
from pathlib import Path
from runpy import run_path
from xml.etree import ElementTree as ET

fonts = run_path(str(Path(__file__).with_name('verify-pages-mixed-font-lists.py')))
profile, forms = fonts['profile'], fonts['forms']


def run_sizes(text):
    if text.startswith(('Root seven ', 'Root native ')):
        return (18, 8, 14)
    if text.startswith('Root eight '):
        return (8, 18, 14)
    if text.startswith('Child one '):
        return (8, 10, 14)
    if text.startswith('Resumed child '):
        return (12, 24, 8)
    if text == 'Native child plain':
        return (8,)
    return (12, 12, 12)


def verify_marker_sizes(path, expected):
    xml = subprocess.check_output(['pdftohtml', '-xml', '-stdout', '-zoom', '1', '-hidden', str(path)], text=True)
    pages = ET.fromstring(xml).findall('page')
    assert len(pages) == 1
    catalog = {node.get('id'): node for node in pages[0].findall('fontspec')}
    markers = []
    for node in pages[0].findall('text'):
        if len(markers) >= len(expected):
            break
        _, number, text = expected[len(markers)]
        marker = fonts['marker'](number, text)
        value = ''.join(node.itertext()).strip()
        if value == marker or value == marker + ' ' + text:
            font = catalog[node.get('font')]
            actual_size = float(font.get('size'))
            expected_size = run_sizes(text)[0]
            assert actual_size == expected_size, (path.name, text, actual_size, expected_size)
            assert font.get('color') == '#000000', (path.name, text, font.attrib)
            assert node.find('b') is None and node.find('i') is None
            markers.append({'marker': marker, 'text': text, 'sizePoints': actual_size,
                            'family': font.get('family')})
    assert len(markers) == len(expected), (path.name, len(markers), len(expected))
    return markers


def verify(root, descriptor):
    cases, decoded, pdfs = [], [], []
    for suffix in ('orders-v15.4', 'created', 'saved-v15.4', 'edited-v15.4'):
        expected = fonts['expected_items'](suffix == 'edited-v15.4')
        stem = 'pages-mixed-size-' + suffix
        title = 'Mixed family decimal items' if suffix == 'orders-v15.4' else 'Mixed size decimal items'
        result = profile['extract'](root / (stem + '.pages'), descriptor)
        fonts['verify_source'](result, expected, run_sizes)
        fonts['verify_docx_fonts'](root / (stem + '-officeimo.docx'), expected, strict=True, run_sizes=run_sizes)
        differences = []
        if suffix != 'created':
            differences = fonts['verify_docx_fonts'](root / (stem + '.docx'), expected, strict=False, run_sizes=run_sizes)
            path = root / (stem + '.pdf')
            pdfs.append({'path': path.name, 'fonts': fonts['verify_pdf'](path, expected, title),
                         'nativeMarkers': verify_marker_sizes(path, expected)})
        for lane in ('officeimo', 'caller-fonts'):
            path = root / (stem + '-' + lane + '.pdf')
            pdfs.append({'path': path.name, 'fonts': fonts['verify_pdf'](path, expected, title)})
        cases.append({'source': result['source'], 'sha256': result['sha256'], 'numberedItems': len(expected),
                      'firstRunMarkerSizes': [run_sizes(text)[0] for _, _, text in expected],
                      'strictSavedDocxSizesAndFontsVerified': True, 'nativeDocxTypographyDifferences': differences})
        decoded.append(result)
    receiving = root / 'pages-mixed-size-officeimo-docx-final-v15.4.pdf'
    expected = fonts['expected_items']()
    pdfs.append({'path': receiving.name, 'fonts': fonts['verify_pdf'](receiving, expected, 'Mixed size decimal items'),
                 'nativeMarkers': verify_marker_sizes(receiving, expected)})
    native = forms['word_bounds'](root / 'pages-mixed-size-saved-v15.4.pdf')
    converted = forms['word_bounds'](receiving)
    assert [row[0] for row in native] == [row[0] for row in converted]
    maximum = max(abs(a - b) for (_, left), (_, right) in zip(native, converted) for a, b in zip(left, right))
    assert maximum <= 0.01, ('Receiving DOCX geometry exceeds 0.01 points', maximum)
    return {'provider': 'numbers-parser4.19.0; pinned Apple schemas; Poppler',
            'descriptorSha256': decoded[0]['descriptorSha256'], 'cases': cases, 'pdfs': pdfs,
            'receivingDocxGeometry': {'wordBounds': len(native), 'maximumDeltaPoints': maximum},
            'limits': ['One-page body lists with whole-point run sizes at marker scale1, paragraph default12 and three font families.',
                       'Marker size follows the first run; paragraph-default size governs the em indentation.',
                       'Fractional mixed sizes, other marker scales, bullets, tiered numbering and additional containers remain unqualified.',
                       'Native DOCX typography differences, portable PDF substitution and caller-font embedding have separate evidence.',
                       'Loaded rich-run editing, visual-cache regeneration and other runtime architectures retain separate gates.']}, decoded


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
