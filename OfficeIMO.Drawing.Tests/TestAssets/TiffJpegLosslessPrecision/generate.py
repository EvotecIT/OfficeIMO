"""Wrap existing independently encoded JPEG samples with test-only LibTIFF 4.7.2.
Build ../TiffJpegArithmeticLossless/wrap.c, then pass the executable as argv[1].
The JPEG payloads are unchanged; this is container qualification, not a new encoder.
"""
from pathlib import Path
import csv, hashlib, subprocess, sys
root = Path(__file__).resolve().parent
rows = []
for bits in range(2, 17):
    if bits in (8, 12, 16):
        continue  # Existing TIFF corpora qualify these precisions.
    maximum = (1 << bits) - 1
    for coding, folder, width in [('huffman', 'JpegLosslessPrecision', 17), ('arithmetic', 'JpegArithmeticLossless', 19)]:
        for channels in (1, 3):
            for point in (0, bits - 1):
                name = (f'p{bits}-c{channels}-d1-t{point}-s{int(point == 0)}.jpg' if coding == 'huffman'
                        else f'b{bits}-c{channels}-p1-t{point}-r19.jpg')
                source = root.parent / folder / name
                payload = source.read_bytes()
                if coding == 'huffman':
                    raw = Path(str(source) + '.raw').read_bytes()
                    samples = [int.from_bytes(raw[i:i+2], 'little') for i in range(0, len(raw), 2)]
                else:
                    samples = [(((x * 193 + y * 791 + c * 3191) ^ (x * y * 53)) & maximum) >> point << point
                               for y in range(11) for x in range(width) for c in range(channels)]
                rgba = bytearray()
                for pixel in range(width * 11):
                    for c in (0, 0 if channels == 1 else 1, 0 if channels == 1 else 2):
                        rgba.append((samples[pixel * channels + c] * 255 + maximum // 2) // maximum)
                    rgba.append(255)
                for endian, mode in [('le', 'wl'), ('be', 'wb')]:
                    output = f'{coding}-{name}.{endian}.tif'
                    subprocess.run([sys.argv[1], str(source), str(root / output), str(bits), str(channels), mode, str(width), '11'], check=True)
                    (root / (output + '.rgba')).write_bytes(rgba)
                    rows.append([output, bits, channels, point, width, 11, endian, folder + '/' + name, hashlib.sha256(payload).hexdigest()])
with (root / 'manifest.csv').open('w') as f:
    writer = csv.writer(f, lineterminator='\n')
    writer.writerow(['name', 'bits', 'components', 'point', 'width', 'height', 'byteOrder', 'source', 'sourceSha256'])
    writer.writerows(rows)
(root / 'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest() + '  ' + p.name + '\n'
    for p in sorted(root.iterdir()) if p.suffix in ('.tif', '.rgba')))
print(len(rows), 'TIFF containers')
