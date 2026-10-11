"""Check OfficeIMO's head stylesheet contract in generated, optimized HTML."""

from html.parser import HTMLParser
from pathlib import Path
import sys
from urllib.parse import urlsplit


class HeadStyles(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.in_head = False
        self.inert_depth = 0
        self.links = []

    def handle_starttag(self, tag, attrs):
        if tag == 'head':
            self.in_head = True
        if tag in ('noscript', 'template'):
            self.inert_depth += 1
        if tag == 'link' and self.in_head and not self.inert_depth:
            self.links.append(attrs)

    def handle_endtag(self, tag):
        if tag == 'head':
            self.in_head = False
        if tag in ('noscript', 'template'):
            self.inert_depth = max(0, self.inert_depth - 1)


def check_styles(html):
    parser = HeadStyles()
    parser.feed(html)
    layout_count = 0
    for attrs in parser.links:
        link = dict(attrs)
        path = urlsplit(link.get('href') or '').path
        if not path.startswith('/css/') or not path.endswith('.css'):
            continue
        if len(attrs) != len(link):
            raise ValueError('Layout stylesheet has duplicate attributes.')
        if (link.get('rel') or '').lower().split() != ['stylesheet']:
            raise ValueError('Layout CSS must use rel=stylesheet.')
        if 'onload' in link or 'disabled' in link:
            raise ValueError('Layout CSS must apply immediately without onload or disabled.')
        media = ' '.join((link.get('media') or '').lower().split())
        if media in ('', 'all', 'screen'):
            layout_count += 1
        elif (Path(path).name == 'prism-theme.css' and
              media in ('(prefers-color-scheme: light)', '(prefers-color-scheme: dark)')):
            # Syntax colors may follow the theme; the layout must always apply.
            continue
        else:
            raise ValueError(f'Layout CSS has unsupported first-paint media: {media!r}.')
    if not layout_count:
        raise ValueError('The document head must contain an applicable layout stylesheet.')


if __name__ == '__main__':
    for filename in sys.argv[1:]:
        try:
            check_styles(Path(filename).read_text(encoding='utf-8-sig'))
        except (OSError, ValueError) as error:
            sys.exit(f'{filename}: {error}')
    print(f'First-paint stylesheet contracts passed for {len(sys.argv) - 1} routes.')
