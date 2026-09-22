"""Regenerate reference-srgb.csv with Pillow and its independent LittleCMS engine.

Reference version: Pillow 12.3.0 / LittleCMS 2.19. Run in this directory.
"""

import csv
from pathlib import Path

from PIL import Image, ImageCms


ROOT = Path(__file__).parent
CASES = {
    "icc-dci-p3-matrix.icc": ("RGB", [(0, 0, 0), (255, 255, 255),
                                      (64, 128, 192), (128, 128, 128),
                                      (128, 64, 32), (32, 64, 128)]),
    "littlecms-rgb-matrix.icc": ("RGB", [(0, 0, 0), (255, 255, 255),
                                         (64, 128, 192), (255, 0, 0),
                                         (0, 255, 0), (0, 0, 255)]),
    "icc-rgb-lut-v4.icc": ("RGB", [(0, 0, 0), (255, 255, 255),
                                   (64, 128, 192), (255, 0, 0),
                                   (0, 255, 0), (0, 0, 255)]),
    "littlecms-cmyk-lut.icc": ("CMYK", [(0, 0, 0, 0), (255, 255, 255, 255),
                                        (64, 128, 192, 26), (255, 0, 0, 0),
                                        (0, 255, 0, 0), (0, 0, 255, 0)]),
}


with (ROOT / "reference-srgb.csv").open("w", newline="", encoding="utf-8") as stream:
    writer = csv.writer(stream, lineterminator="\n")
    writer.writerow(("profile", "input", "expected"))
    for name, (mode, samples) in CASES.items():
        source = ImageCms.getOpenProfile(str(ROOT / name))
        target = ImageCms.createProfile("sRGB")
        transform = ImageCms.buildTransformFromOpenProfiles(
            source, target, mode, "RGB", renderingIntent=1, flags=0)
        image = Image.new(mode, (len(samples), 1))
        image.putdata(samples)
        for sample, result in zip(samples, ImageCms.applyTransform(image, transform).getdata()):
            writer.writerow((name, ":".join(map(str, sample)), ":".join(map(str, result))))
