#!/usr/bin/env python3
"""Create deterministic OCR comparison fixtures and their expected text."""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path

import PIL
from PIL import Image, ImageDraw, ImageFont


EXPECTED_PILLOW_VERSION = "11.0.0"
EXPECTED_FONTS = {
    "DejaVuSans.ttf": "ae7b7855e115a5966d8b1b3f80f254ccc117ec86f9965e202ee2940453837280",
    "DejaVuSans-Bold.ttf": "5c1247acef7f2b8522a31742c76d6adcb5569bacc0be7ceaa4dc39dd252ce895",
}


CASES = (
    {
        "name": "clean-english",
        "language": "eng",
        "lines": (
            "OFFICEIMO OCR VALIDATION",
            "Invoice 1042",
            "Total due: 1,234.56 USD",
            "Searchable documents should preserve exact words.",
        ),
        "foreground": (12, 20, 32),
        "background": (255, 255, 255),
        "rotation": 0.0,
    },
    {
        "name": "dense-table",
        "language": "eng",
        "lines": (
            "PRODUCT             QTY      PRICE      TOTAL",
            "Widget Alpha         12       9.99     119.88",
            "Widget Beta           7      24.50     171.50",
            "Service Plan          1      89.00      89.00",
            "GRAND TOTAL                            380.38",
        ),
        "foreground": (20, 25, 32),
        "background": (250, 250, 248),
        "rotation": 0.0,
    },
    {
        "name": "low-contrast-skew",
        "language": "eng",
        "lines": (
            "ARCHIVE COPY 2026-09-01",
            "Reference: OCR-7391-A",
            "Low contrast scans still need searchable text.",
        ),
        "foreground": (115, 115, 112),
        "background": (229, 228, 222),
        "rotation": 1.8,
    },
)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def font_path(bold: bool) -> Path:
    name = "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf"
    candidate = Path("/usr/share/fonts/truetype/dejavu") / name
    if not candidate.is_file():
        raise FileNotFoundError(f"Pinned fixture font was not found: {candidate}")
    actual = sha256(candidate)
    if actual != EXPECTED_FONTS[name]:
        raise RuntimeError(f"Pinned fixture font failed SHA-256 verification: {candidate}")
    return candidate


def verify_renderer() -> None:
    if PIL.__version__ != EXPECTED_PILLOW_VERSION:
        raise RuntimeError(
            f"Fixture generation requires Pillow {EXPECTED_PILLOW_VERSION}, found {PIL.__version__}."
        )
    font_path(False)
    font_path(True)


def render(case: dict, output: Path) -> None:
    image = Image.new("RGB", (1800, 1040), case["background"])
    draw = ImageDraw.Draw(image)
    title_font = ImageFont.truetype(str(font_path(True)), 58)
    body_font = ImageFont.truetype(str(font_path(False)), 46)
    y = 115
    for index, line in enumerate(case["lines"]):
        draw.text((115, y), line, font=title_font if index == 0 else body_font, fill=case["foreground"])
        y += 155 if index == 0 else 130
    if case["rotation"]:
        image = image.rotate(case["rotation"], resample=Image.Resampling.BICUBIC, expand=False, fillcolor=case["background"])
    image.save(output, format="PNG", optimize=False)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    verify_renderer()
    args.output.mkdir(parents=True, exist_ok=True)
    manifest = []
    for case in CASES:
        image_path = args.output / f"{case['name']}.png"
        render(case, image_path)
        manifest.append({
            "name": case["name"],
            "language": case["language"],
            "image": image_path.name,
            "expected": "\n".join(case["lines"]),
            "sha256": sha256(image_path),
            "byteCount": image_path.stat().st_size,
        })
    (args.output / "cases.json").write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")


if __name__ == "__main__":
    main()
