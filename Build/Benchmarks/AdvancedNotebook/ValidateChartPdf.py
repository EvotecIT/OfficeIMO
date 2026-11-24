"""Require the expected text and visible colored chart in a rasterized PDF."""
import argparse
from pathlib import Path
import subprocess
import tempfile

parser = argparse.ArgumentParser()
parser.add_argument("pdf", type=Path)
parser.add_argument("witness")
parser.add_argument("--preview", type=Path)
parser.add_argument("--require-vectors", action="store_true")
args = parser.parse_args()
text = subprocess.check_output(["pdftotext", str(args.pdf), "-"], text=True)
if args.witness not in " ".join(text.split()):
    parser.error(f"Expected selectable text witness {args.witness!r} is absent.")
if "Saved figure" in text or "fig.show()" in text:
    parser.error("Hidden input leaked into the rendered PDF.")
if args.require_vectors:
    image_list = subprocess.check_output(["pdfimages", "-list", str(args.pdf)], text=True)
    if len(image_list.splitlines()) != 2:
        parser.error("Expected native vector geometry without raster images in this chart PDF.")

with tempfile.TemporaryDirectory() as tmp:
    prefix = Path(tmp) / "page"
    subprocess.run(["pdftoppm", "-r", "72", str(args.pdf), str(prefix)], check=True)
    pages = sorted(Path(tmp).glob("page-*.ppm"))
    most_color = 0
    chart_page = None
    for page in pages:
        magic, dims, rgb_max, pixels = page.read_bytes().split(b"\n", 3)
        width, height = map(int, dims.split())
        assert magic == b"P6" and rgb_max == b"255" and len(pixels) == 3 * width * height
        r, g, b = pixels[::3], pixels[1::3], pixels[2::3]
        colorful = sum(1 for rr, gg, bb in zip(r, g, b) if max(rr, gg, bb) - min(rr, gg, bb) >= 35)
        if colorful > most_color:
            most_color = colorful
            chart_page = int(page.stem.rsplit("-", 1)[1])
    if most_color < 300 or chart_page is None:
        parser.error("Expected a visible colored plot, with no verified image in this artifact.")
    if args.preview:
        subprocess.run([
            "pdftoppm", "-f", str(chart_page), "-l", str(chart_page), "-singlefile",
            "-r", "110", "-png", str(args.pdf), str(args.preview.with_suffix("")),
        ], check=True)
    print(f"{args.pdf.name}: {len(pages)} page(s), plot page {chart_page}, {most_color} color pixels")
