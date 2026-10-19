"""Regenerate the synthetic PNG inputs; Pillow is an optional fixture oracle, never a runtime dependency."""
from io import BytesIO
from pathlib import Path
import struct
import zlib
from PIL import Image

root = Path(__file__).resolve().parent
colors = [(240, 20, 20), (20, 230, 20), (20, 20, 235), (235, 220, 20)]
for mode in ["RGB", "RGBA"]:
    image = Image.new(mode, (12, 8))
    for y in range(8):
        for x in range(12):
            q = (y // 4) * 2 + x // 6
            image.putpixel((x, y), colors[q] + ((255, 128, 0, 64)[q],) if mode == "RGBA" else colors[q])
    image.save(root / (mode.lower() + "-pillow.png"), compress_level=9)

def chunk(kind, data=b""):
    return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))

def paeth(a, b, c):
    p = a + b - c
    distances = [abs(p - a), abs(p - b), abs(p - c)]
    return [a, b, c][distances.index(min(distances))]

rows = [bytes(component for x in range(4) for component in ((x * 67 + y * 29) % 256,
        (x * 13 + y * 83) % 256, (x * 113 + y * 7) % 256, (x * 51 + y * 37) % 256)) for y in range(5)]
filtered = bytearray()
for filter_type, row in enumerate(rows):
    previous = rows[filter_type - 1] if filter_type else bytes(len(row))
    filtered.append(filter_type)
    for i, value in enumerate(row):
        a, b, c = row[i - 4] if i >= 4 else 0, previous[i], previous[i - 4] if i >= 4 else 0
        predictor = [0, a, b, (a + b) // 2, paeth(a, b, c)][filter_type]
        filtered.append((value - predictor) % 256)
encoded = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 4, 5, 8, 6, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(filtered)) + chunk(b"IEND")
assert Image.open(BytesIO(encoded)).tobytes() == b"".join(rows)
(root / "rgba-filters.png").write_bytes(encoded)
