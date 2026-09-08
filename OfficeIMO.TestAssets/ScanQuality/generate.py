"""Generate deterministic degraded scan variants using Pillow, independently of OfficeIMO.

Source TIFFs and transcripts are pinned upstream fixtures; see sources.json.
Run explicitly with Pillow 11 or later. Neither product code nor normal tests need Python.
"""
from pathlib import Path
from PIL import Image
import zlib

root = Path(__file__).parent
for name in ("phototest", "eurotext"):
    source = Image.open(root / (name + ".tif")).convert("RGB")
    source.save(root / (name + ".png"))
    # Pillow angles are counterclockwise. The planted clockwise skew is +3 degrees.
    source.rotate(-3, resample=Image.Resampling.BICUBIC, expand=True, fillcolor="white").save(root / (name + "-skew.png"))
    source.transpose(Image.Transpose.ROTATE_270).save(root / (name + "-clockwise90.png"))
    source.transpose(Image.Transpose.ROTATE_180).save(root / (name + "-upside-down.png"))
    shaded = source.copy()
    pixels = shaded.load()
    for y in range(shaded.height):
        for x in range(shaded.width):
            paper = 0.30 + 0.65 * x / max(1, shaded.width - 1)
            r, g, b = pixels[x, y]
            pixels[x, y] = tuple(round(value * paper) for value in (r, g, b))
    shaded.rotate(-3, resample=Image.Resampling.BICUBIC, expand=True, fillcolor=(245, 245, 245)).save(root / (name + "-shadow-skew.png"))

# Independent image-only PDFs at 300 DPI, with no OfficeIMO writer involved.
for path in sorted(root.glob("*.png")):
    image = Image.open(path).convert("RGB")
    width, height = image.width * 72 / 300, image.height * 72 / 300
    pixels = zlib.compress(image.tobytes())
    content = f"q {width} 0 0 {height} 0 0 cm /Scan Do Q".encode("ascii")
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
        f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {width} {height}] /Resources << /XObject << /Scan 5 0 R >> >> /Contents 4 0 R >>".encode("ascii"),
        f"<< /Length {len(content)} >>\nstream\n".encode("ascii") + content + b"\nendstream",
        f"<< /Type /XObject /Subtype /Image /Width {image.width} /Height {image.height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Length {len(pixels)} >>\nstream\n".encode("ascii") + pixels + b"\nendstream",
    ]
    document = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = [0]
    for index, obj in enumerate(objects, 1):
        offsets.append(len(document))
        document.extend(f"{index} 0 obj\n".encode("ascii") + obj + b"\nendobj\n")
    xref = len(document)
    document.extend(f"xref\n0 {len(offsets)}\n0000000000 65535 f \n".encode("ascii"))
    for offset in offsets[1:]:
        document.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
    document.extend(f"trailer << /Size {len(offsets)} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode("ascii"))
    path.with_suffix(".pdf").write_bytes(document)
