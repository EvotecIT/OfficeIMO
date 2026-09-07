"""Regenerate independent T.4/T.6 strips and packed pixel expectations with Pillow/libtiff.

Run this script explicitly when changing fixtures. Product tests read the checked-in
bytes and do not require Python, Pillow, or libtiff.
"""
from io import BytesIO
from pathlib import Path
from PIL import Image

root = Path(__file__).parent
width, height = 3001, 17
image = Image.new("1", (width, height), 1)
for y in range(height):
    for x in range(width):
        black = y == 1 or (y >= 2 and ((x + y % 4) // (1 + y % 7)) % 5 == 0)
        if black:
            image.putpixel((x, y), 0)
(root / "fax-pattern.pixels").write_bytes(image.tobytes())
for compression in ("group3", "group4"):
    buffer = BytesIO()
    image.save(buffer, format="TIFF", compression=compression)
    buffer.seek(0)
    tiff = Image.open(buffer)
    offsets, lengths = tiff.tag_v2[273], tiff.tag_v2[279]
    assert len(offsets) == 1
    payload = buffer.getvalue()[offsets[0]:offsets[0] + lengths[0]]
    (root / ("fax-pattern." + compression)).write_bytes(payload)

# Independent JPEG 2000 samples exercise the PDF boundary with and without opacity.
for mode, color in (("RGB", (255, 0, 0)), ("RGBA", (255, 0, 0, 0))):
    Image.new(mode, (1, 1), color).save(root / ("red-" + mode.lower() + ".jp2"), format="JPEG2000")
