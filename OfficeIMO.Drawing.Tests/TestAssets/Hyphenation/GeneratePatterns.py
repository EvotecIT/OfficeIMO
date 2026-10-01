"""Derive bounded Core resources from the pinned, unmodified TeX pattern inputs.
Usage: python GeneratePatterns.py <source-directory>
"""
from pathlib import Path
import hashlib, json, re, sys
root = Path(__file__).resolve().parents[3]
source = Path(sys.argv[1])
output = root / 'OfficeIMO.Core/Typography/Hyphenation'
output.mkdir(parents=True, exist_ok=True)
commit = '5684c0f51c0b81133db2efbe60a408b4155a3ff5'
expected = {'en-us': (4938, 14), 'de-1996': (36709, 0)}
source_hashes = {'en-us': 'f4ffcd96c5cbc886bdad23f95dcae8edc3cd3620eae62f7946eceda97c4e68f8',
                 'de-1996': '374ad1ce3263f8a2791ec070ef9b516c532dae64456f3fb656e617e8ae53a9d3'}
rows = []
for language, counts in expected.items():
    path = source / f'hyph-{language}.tex'
    original = path.read_bytes()
    assert hashlib.sha256(original).hexdigest() == source_hashes[language]
    text = original.decode('utf-8')
    clean = re.sub(r'%[^\n]*', '', text)
    patterns = ' '.join(re.findall(r'\\patterns\s*\{([^}]*)\}', clean)).split()
    exceptions = ' '.join(re.findall(r'\\hyphenation\s*\{([^}]*)\}', clean)).split()
    assert (len(patterns), len(exceptions)) == counts
    assert all(all(c.isalpha() or c in '0123456789.' for c in p) for p in patterns)
    url = f'https://raw.githubusercontent.com/hyphenation/tex-hyphen/{commit}/hyph-utf8/tex/generic/hyph-utf8/patterns/tex/{path.name}'
    header = '\n'.join('#' + line[1:] for line in text.split('\\patterns', 1)[0].splitlines() if line.startswith('%'))
    content = f'# Derived resource; pattern tokens and exceptions unchanged.\n# Source: {url}\n{header}\n'
    content += '\n'.join(patterns) + '\n'
    content += ''.join('!' + word + '\n' for word in exceptions)
    target = output / f'{language}.patterns'
    target.write_text(content, encoding='utf-8')
    rows.append({'language': language, 'source': url, 'sourceSha256': hashlib.sha256(original).hexdigest(),
                 'resource': str(target.relative_to(root)), 'resourceSha256': hashlib.sha256(target.read_bytes()).hexdigest(),
                 'patterns': len(patterns), 'exceptions': len(exceptions)})
(output / 'manifest.json').write_text(json.dumps({'schemaVersion': 1, 'upstreamCommit': commit, 'resources': rows}, indent=2) + '\n')
print(json.dumps(rows, indent=2))
