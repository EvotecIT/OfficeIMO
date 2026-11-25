"""Generate lossless VP8L fixtures with libwebp and record libwebp's own decode as the reference.
Requires Pillow 11.3.0/libwebp 1.5.0. Run from this directory; committed fixtures are test inputs.
The cases cover prefix codes longer than eight bits, the color cache, palettes packed 8, 4, 2
and 1 index per pixel word, overlapping backreferences and the predictor and color transforms.
"""
from pathlib import Path
import hashlib
import json
import math
import random

from PIL import Image, features

root = Path(__file__).parent
rng = random.Random(20261124)


def skewed():
    # Geometric-like values give some symbols long prefix codes.
    value = 0
    while value < 255 and rng.random() < 0.82:
        value += 1
    return value


def noise(width, height):
    return [(skewed(), skewed(), (x * 5 + skewed()) % 256, 255 - skewed())
            for y in range(height) for x in range(width)]


def smooth(width, height):
    # Photo-like content: smooth surfaces with mild noise and texture.
    def channel(x, y, a, b, c):
        return max(0, min(255, int(128 + 90 * math.sin(x / a + y / b) * math.cos(y / c)) + rng.randrange(-4, 5)))
    return [(channel(x, y, 9.0, 13.0, 17.0), channel(x, y, 11.0, 7.0, 23.0),
             channel(x, y, 5.0, 19.0, 29.0), 255 if x < width // 2 else channel(x, y, 31.0, 37.0, 41.0))
            for y in range(height) for x in range(width)]


def palette(colors):
    table = [(rng.randrange(256), rng.randrange(256), rng.randrange(256), rng.choice([96, 255]))
             for _ in range(colors)]

    def make(width, height):
        return [table[rng.randrange(colors)] for y in range(height) for x in range(width)]
    return make


def repeat(width, height):
    period = [(200, 10, 30, 255), (20, 220, 40, 255), (30, 40, 210, 200)]
    return [period[x % 3] if y % 4 else (x * 9 % 256, y * 11 % 256, 77, 255)
            for y in range(height) for x in range(width)]


cases = [
    ('lossless-noise', 61, 47, noise, 6, 100),
    ('lossless-smooth', 128, 48, smooth, 6, 100),
    ('lossless-smooth-fast', 96, 64, smooth, 2, 50),
    ('lossless-palette-2', 53, 29, palette(2), 4, 75),
    ('lossless-palette-4', 53, 29, palette(4), 4, 75),
    ('lossless-palette-16', 53, 29, palette(16), 4, 75),
    ('lossless-palette-64', 53, 29, palette(64), 4, 75),
    ('lossless-repeat', 97, 31, repeat, 6, 100),
]
manifest = {'producer': {'pillow': Image.__version__, 'libwebp': features.version('webp'),
    'command': 'Pillow Image.save WEBP lossless=True exact=True', 'seed': 20261124}, 'cases': []}
for name, width, height, make, method, quality in cases:
    image = Image.new('RGBA', (width, height))
    pixels = make(width, height)
    image.putdata(pixels)
    path = root / (name + '.webp')
    image.save(path, 'WEBP', lossless=True, exact=True, method=method, quality=quality)
    data = path.read_bytes()
    assert data[12:16] == b'VP8L', name
    reference = Image.open(path).convert('RGBA').tobytes()
    assert reference == image.tobytes(), name
    (root / (name + '.rgba')).write_bytes(reference)
    manifest['cases'].append({'id': name, 'path': path.name, 'width': width, 'height': height,
        'method': method, 'quality': quality, 'sha256': hashlib.sha256(data).hexdigest(),
        'referenceRgbaSha256': hashlib.sha256(reference).hexdigest()})
(root / 'manifest.json').write_text(json.dumps(manifest, indent=2) + '\n')
