"""Extract every Apple scale-reference PDF page with the isolated pypdf tool.

PDF text is a separate oracle because Poppler bbox extraction misplaces this
fixture's fi ligatures. This does not validate PDF conformance or appearance.
"""
import argparse
import hashlib
import json
from pathlib import Path
from importlib.metadata import version

from pypdf import PdfReader


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pdf', type=Path)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    reader = PdfReader(args.pdf)
    pages = [{'position': index, 'sizePoints': [float(page.mediabox.width), float(page.mediabox.height)],
              'text': page.extract_text()} for index, page in enumerate(reader.pages, 1)]
    result = {'provider': 'pypdf', 'providerVersion': version('pypdf'),
              'sourceSha256': hashlib.sha256(args.pdf.read_bytes()).hexdigest(), 'pages': pages,
              'limits': 'Extracted PDF text and media boxes only; parser warnings remain in the execution log.'}
    args.output.write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
    print('Extracted', len(pages), 'PDF pages')


if __name__ == '__main__':
    main()
