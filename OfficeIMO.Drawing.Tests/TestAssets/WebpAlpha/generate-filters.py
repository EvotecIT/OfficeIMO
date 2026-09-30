"""Generate ALPH filter fixtures with libwebp streams and verify the complete file with libwebp.
Requires Pillow 11.3.0/libwebp 1.5.0. Run from this directory; committed fixtures are test inputs.
"""
from io import BytesIO
from pathlib import Path
import hashlib
import json
import random
import struct
from PIL import Image, features

root = Path(__file__).parent
width, height = 49, 33
alpha = [(x * 83 + y * 137) % 256 for y in range(height) for x in range(width)]
rng = random.Random(43017)
manifest = {'producer': {'pillow': Image.__version__, 'libwebp': features.version('webp'),
    'command': 'Pillow Image.save WEBP quality=82 method=6 exact=True', 'seed': 43017}, 'cases': []}
for name, noise in [('alpha-gradient', False), ('alpha-noise', True), ('opaque-control', None)]:
    image = Image.new('RGBA', (width, height))
    image.putdata([(40 + x * 3, 50 + y * 4, 160,
        255 if noise is None else rng.randrange(256) if noise else (x * 7 + y * 5) % 256)
        for y in range(height) for x in range(width)])
    path = root / (name + '.webp')
    image.save(path, 'WEBP', quality=82, method=6, exact=True)
    data = path.read_bytes()
    reference = Image.open(path).convert('RGBA').tobytes()
    (root / (name + '.rgba')).write_bytes(reference)
    manifest['cases'].append({'id': name, 'path': path.name, 'width': width, 'height': height,
        'sha256': hashlib.sha256(data).hexdigest(),
        'referenceRgbaSha256': hashlib.sha256(reference).hexdigest()})
opaque = (root / 'opaque-control.webp').read_bytes()[12:]
manifest['filterProducer'] = {
    'pillow': Image.__version__, 'libwebp': features.version('webp'),
    'source': 'generate-filters.py',
    'reference': 'Complete recontainerized files decoded independently with Pillow/libwebp.'
}

def chunk(tag, payload):
    return tag + struct.pack('<I', len(payload)) + payload + b'\0' * (len(payload) % 2)

for compressed in (False, True):
    for filtering in range(4):
        residual = []
        for y in range(height):
            for x in range(width):
                index = y * width + x
                left = alpha[index - 1] if x else 0
                above = alpha[index - width] if y else 0
                corner = alpha[index - width - 1] if x and y else 0
                predictor = 0
                if filtering:
                    if not y:
                        predictor = left
                    elif not x:
                        predictor = above
                    elif filtering == 1:
                        predictor = left
                    elif filtering == 2:
                        predictor = above
                    else:
                        predictor = max(0, min(255, left + above - corner))
                residual.append((alpha[index] - predictor) % 256)
        payload = bytes(residual)
        if compressed:
            plane = Image.new('RGBA', (width, height))
            plane.putdata([(0, value, 0, 255) for value in residual])
            buffer = BytesIO()
            plane.save(buffer, 'WEBP', lossless=True, method=6, exact=True)
            encoded = buffer.getvalue()
            assert encoded[12:16] == b'VP8L'
            count = struct.unpack_from('<I', encoded, 16)[0]
            payload = encoded[25:20 + count]  # Headerless implicit-dimension image stream.
        control = int(compressed) | filtering << 2
        canvas = bytes([16, 0, 0, 0]) + (width - 1).to_bytes(3, 'little') + (height - 1).to_bytes(3, 'little')
        chunks = chunk(b'VP8X', canvas) + chunk(b'ALPH', bytes([control]) + payload) + opaque
        data = b'RIFF' + struct.pack('<I', len(chunks) + 4) + b'WEBP' + chunks
        reference = Image.open(BytesIO(data)).convert('RGBA').tobytes()
        assert list(reference[3::4]) == alpha
        name = ('compressed' if compressed else 'raw') + '-filter-' + str(filtering)
        (root / (name + '.webp')).write_bytes(data)
        (root / (name + '.rgba')).write_bytes(reference)
        manifest['cases'] = [case for case in manifest['cases'] if case['id'] != name]
        manifest['cases'].append({
            'id': name, 'path': name + '.webp', 'width': width, 'height': height,
            'sha256': hashlib.sha256(data).hexdigest(),
            'referenceRgbaSha256': hashlib.sha256(reference).hexdigest(),
            'compression': int(compressed), 'filter': filtering
        })
(root / 'manifest.json').write_text(json.dumps(manifest, indent=2) + '\n')
print('Verified all eight complete ALPH filter files against libwebp.')
