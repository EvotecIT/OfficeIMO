"""Opt-in extraction and 72 dpi geometry gate for native vertical PDF drawing."""

import argparse
import hashlib
import os
from pathlib import Path
import re
import subprocess
import tempfile
from urllib.request import urlopen
import xml.etree.ElementTree as ET

from PIL import Image


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
REFERENCE_URL = (
    "https://raw.githubusercontent.com/mozilla/pdf.js/"
    "d54c193bd4dd6c34759cb88f1a3b78db66f0962c/test/pdfs/vertical.pdf"
)
REFERENCE_SHA256 = "514511143db12309893fb69cdf98c76f2361c20da394f79fdff2c119ed7a4393"
CHROME_SHA256 = "c10e749fa6601a6c4074e20a4b43fd6c4d993ae7c87d52a0a5002e4ef8c7d182"
TEXT = "日本語（例）。"


def run(*command, env=None):
    result = subprocess.run(command, cwd=ROOT, env=env, capture_output=True, check=True)
    return result.stdout.decode("utf-8", errors="replace")


def check_hash(path, expected):
    actual = hashlib.sha256(path.read_bytes()).hexdigest()
    if actual != expected:
        raise AssertionError(f"{path.name}: SHA-256 {actual}, expected {expected}")


def ink(path):
    image = Image.open(path).convert("L")
    return image.size, [value < 200 for value in image.tobytes()]


def main(output):
    output.mkdir(parents=True, exist_ok=True)
    chrome = HERE / "chrome-identity-h.pdf"
    check_hash(chrome, CHROME_SHA256)
    dvipdfmx = output / "dvipdfmx-identity-v.pdf"
    if not dvipdfmx.exists():
        with urlopen(REFERENCE_URL, timeout=20) as response:
            payload = response.read(6906)
        if len(payload) != 6905:
            raise AssertionError("The pinned Identity-V fixture changed size or exceeded its bound")
        dvipdfmx.write_bytes(payload)
    check_hash(dvipdfmx, REFERENCE_SHA256)

    env = os.environ.copy()
    generated = output / "officeimo-vertical.pdf"
    env["OFFICEIMO_VERTICAL_PDF_EVIDENCE"] = str(generated)
    run("dotnet", "test", str(ROOT / "OfficeIMO.Drawing.HarfBuzz.Tests" /
                            "OfficeIMO.Drawing.HarfBuzz.Tests.csproj"),
        "--framework", "net10.0", "--filter",
        "FullyQualifiedName~JapanesePdfPositionsSubstitutedVerticalGlyphsAtProviderYAdvances",
        "--verbosity", "quiet", env=env)
    if not generated.exists():
        raise AssertionError("The native PDF integration test produced no artifact")

    font_v = run("pdffonts", str(dvipdfmx))
    font_h = run("pdffonts", str(chrome))
    font_out = run("pdffonts", str(generated))
    if not re.search(r"CID TrueType\s+Identity-V\s+yes\s+yes\s+no", font_v):
        raise AssertionError("The dvipdfmx fixture is not the expected embedded Identity-V font")
    for name, fonts in (("Chrome", font_h), ("OfficeIMO", font_out)):
        if not re.search(r"CID TrueType\s+Identity-H\s+yes\s+(?:yes|no)\s+yes", fonts):
            raise AssertionError(f"{name} does not have an embedded, searchable Identity-H font")

    for name, path, expected in (("dvipdfmx", dvipdfmx, "あいうえお日本語"),
                                 ("Chrome", chrome, TEXT), ("OfficeIMO", generated, TEXT)):
        extracted = run("pdftotext", "-enc", "UTF-8", "-layout", str(path), "-")
        compact = "".join(extracted.split())
        if compact != expected:
            raise AssertionError(f"{name} extraction: {compact!r}, expected {expected!r}")

    boxes = run("pdftotext", "-bbox", "-f", "1", "-l", "1", str(dvipdfmx), "-")
    root = ET.fromstring(boxes)
    words = {"".join(element.itertext()): element for element in root.iter()
             if element.tag.endswith("word")}
    kana, kanji = words["あいうえお"], words["日本語"]
    for word in (kana, kanji):
        width = float(word.attrib["xMax"]) - float(word.attrib["xMin"])
        height = float(word.attrib["yMax"]) - float(word.attrib["yMin"])
        if height <= width * 2:
            raise AssertionError("Identity-V reference is not a vertical glyph column")
    if float(kana.attrib["xMin"]) <= float(kanji.attrib["xMin"]):
        raise AssertionError("Identity-V reference columns are not ordered right to left")

    for name, path in (("chrome", chrome), ("officeimo", generated)):
        run("pdftoppm", "-f", "1", "-l", "1", "-r", "72", "-png",
            "-singlefile", str(path), str(output / name))
    size_h, mask_h = ink(output / "chrome.png")
    size_out, mask_out = ink(output / "officeimo.png")
    if size_h != (120, 180) or size_out != size_h:
        raise AssertionError("The comparable rendered pages must be 120 x 180 pixels")
    intersection = sum(left and right for left, right in zip(mask_h, mask_out))
    union = sum(left or right for left, right in zip(mask_h, mask_out))
    overlap = intersection / union if union else 0.0
    if overlap < 0.90:
        raise AssertionError(f"Vertical glyph ink overlap {overlap:.3f} is below 0.90")
    print(f"PASS: Identity-V columns and extraction; Identity-H extraction; "
          f"OfficeIMO/Chrome 72 dpi ink overlap {overlap:.3f} ({intersection}/{union}).")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, help="Retain generated PDF and PNG evidence here")
    arguments = parser.parse_args()
    if arguments.output:
        main(arguments.output.resolve())
    else:
        with tempfile.TemporaryDirectory(prefix="officeimo-vertical-reference-") as temporary:
            main(Path(temporary))
