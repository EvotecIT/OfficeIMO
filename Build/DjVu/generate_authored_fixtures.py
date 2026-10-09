#!/usr/bin/env python3
"""Opt-in native fixture production. Requires explicitly supplied validation tools."""

import argparse
import hashlib
import json
from pathlib import Path
import struct
import subprocess


def chunk(name, data):
    return name + struct.pack(">I", len(data)) + data + (b"\0" if len(data) % 2 else b"")


def page_chunks(data):
    at = 16
    while at < len(data):
        size = int.from_bytes(data[at + 4:at + 8], "big")
        yield data[at:at + 4], data[at + 8:at + 8 + size]
        at += 8 + size + (size % 2)


def append(source, destination, tag, payload):
    data = source.read_bytes()
    body = data[12:12 + int.from_bytes(data[8:12], "big")]
    body += b"\0" if len(body) % 2 else b""
    destination.write_bytes(b"AT&T" + chunk(b"FORM", body + chunk(tag, payload)))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--djvulibre-bin", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--source-fixtures", type=Path, default=Path(__file__).resolve().parents[2] / "OfficeIMO.DjVu.Tests/Fixtures")
    parser.add_argument("--minidjvu", type=Path)
    parser.add_argument("--cjpeg", type=Path)
    parser.add_argument("--jb2-work-producer", type=Path)
    args = parser.parse_args()
    output = args.output.resolve()
    if output.exists():
        parser.error("--output must name a new task-owned directory")
    source = args.source_fixtures.resolve()
    output.mkdir(parents=True)

    def run(tool, *values):
        return subprocess.run([str((args.djvulibre_bin / tool).resolve()), *map(str, values)], check=True, stdout=subprocess.PIPE).stdout

    def reference(name):
        run("ddjvu", "-format=ppm", output / (name + ".djvu"), output / (name + "-reference.ppm"))

    # High-frequency input distinguishes full/half chroma, grayscale and edge lifting.
    for name, width, height, options in [
        ("noise-small", 13, 9, []), ("noise-full", 129, 97, ["-crcbfull"]),
        ("noise-normal", 129, 97, []), ("gray", 129, 97, ["-crcbnone"]),
    ]:
        pixels = bytes(((x * 73 + y * 41 + c * 97) ^ (x * y * 19)) & 255
                       for y in range(height) for x in range(width) for c in range(3))
        ppm = output / (name + "-source.ppm")
        ppm.write_bytes(f"P6\n{width} {height}\n255\n".encode() + pixels)
        run("c44", "-slice", "100", *options, ppm, output / (name + ".djvu"))
        reference(name)

    # Increase INFO dimensions only: the background must be interpolated independently.
    data = bytearray((output / "noise-full.djvu").read_bytes())
    info = data.index(b"INFO") + 8
    data[info:info + 4] = struct.pack(">HH", 387, 291)
    (output / "sampling.djvu").write_bytes(data)
    reference("sampling")

    width, height = 128, 96
    pixels = bytearray([255] * (width * height * 3))
    for band, color in enumerate([(0, 0, 0), (200, 20, 40), (10, 120, 220), (40, 180, 30)]):
        for line in range(3):
            for column in range(4):
                for y in range(4 + band * 23 + line * 6, 8 + band * 23 + line * 6):
                    for x in range(6 + column * 30, 18 + column * 30):
                        pixels[(y * width + x) * 3:(y * width + x) * 3 + 3] = bytes(color)
    ppm = output / "palette-source.ppm"
    ppm.write_bytes(f"P6\n{width} {height}\n255\n".encode() + pixels)
    run("cpaldjvu", "-colors", "5", "-bgwhite", "-dpi", "300", ppm, output / "palette.djvu")
    reference("palette")
    # DjVu v3 permits a single FG44 chunk. Encode the complete foreground in
    # that chunk and retain the native JB2 stencil; ddjvu qualifies composition.
    run("c44", "-slice", "100", ppm, output / "foreground-color.djvu")
    body = b"DJVU" + b"".join(chunk(tag, payload) for tag, payload in page_chunks((output / "palette.djvu").read_bytes()) if tag != b"FGbz")
    body += b"".join(chunk(b"FG44", payload) for tag, payload in page_chunks((output / "foreground-color.djvu").read_bytes()) if tag == b"BG44")
    (output / "iw44-foreground.djvu").write_bytes(b"AT&T" + chunk(b"FORM", body))
    reference("iw44-foreground")
    data = bytearray((output / "palette.djvu").read_bytes())
    data[data.index(b"INFO") + 17] = 5
    (output / "palette-rotated.djvu").write_bytes(data)
    reference("palette-rotated")

    text = output / "unicode-input.sexp"
    text.write_text('(page 0 0 128 96 (line 4 4 120 22 (word 4 4 54 22 "Zażółć") (word 60 4 110 22 "😀")))\n', encoding="utf-8")
    (output / "unicode.djvu").write_bytes(data)
    run("djvused", output / "unicode.djvu", "-s", "-e", f"select 1; set-txt {json.dumps(str(text))}")
    for command, suffix in [("print-pure-txt", "txt"), ("print-txt", "zones")]:
        (output / ("unicode." + suffix)).write_bytes(run("djvused", output / "unicode.djvu", "-e", f"select 1; {command}"))
    append(output / "palette.djvu", output / "empty.djvu", b"TXTa", b"\0\0\0\1")
    append(output / "palette.djvu", output / "corrupt.djvu", b"TXTa", b"\0\0\x64\1broken")
    run("djvm", "-c", output / "reader-book.djvu", *[output / (name + ".djvu") for name in ["palette", "unicode", "empty", "corrupt"]])
    outline = output / "outline-input.sexp"
    outline.write_text('(bookmarks ("Group 😀" "" ("Stored text" "#unicode.djvu") ("Image" "palette.djvu")) ("External" "https://example.org/archive") ("Other page" "#3"))\n', encoding="utf-8")
    run("djvused", output / "reader-book.djvu", "-s", "-e", f"set-outline {json.dumps(str(outline))}")
    (output / "reader-book.outline").write_bytes(run("djvused", output / "reader-book.djvu", "-e", "print-outline"))

    for name in ["short", "binary", "multiblock"]:
        run("bzz", "-e10", source / "Bzz" / (name + ".bin"), output / (name + ".bzz"))

    if args.cjpeg:
        jpeg = output / "background.jpg"
        subprocess.run([str(args.cjpeg.resolve()), "-quality", "90", "-sample", "1x1", "-outfile", str(jpeg), str(output / "noise-full-source.ppm")], check=True)
        run("djvumake", output / "jpeg-background.djvu", "INFO=129,97,100", f"BGjp={jpeg}")
        reference("jpeg-background")

    if args.minidjvu:
        shared = output / "Shared"
        shared.mkdir()
        data = (source / "Authored/symbols-reference.pbm").read_bytes()
        header = data.index(b"\n", data.index(b"\n") + 1) + 1
        raster = data[header:]
        first, second = shared / "shared-1.pbm", shared / "shared-2.pbm"
        first.write_bytes(data)
        second.write_bytes(data[:header] + raster[128:] + raster[:128])
        subprocess.run([str(args.minidjvu.resolve()), str(first), str(second), str(shared / "shared.djvu")], check=True)
        for page in [1, 2]:
            run("ddjvu", "-format=ppm", f"-page={page}", shared / "shared.djvu", shared / f"page-{page}-reference.ppm")
        indirect = shared / "Indirect"
        indirect.mkdir()
        subprocess.run([str(args.minidjvu.resolve()), "-i", str(first), str(second), "index.djvu"], cwd=indirect, check=True)

    if args.jb2_work_producer:
        subprocess.run([str(args.jb2_work_producer.resolve()), str(output)], check=True)
        info = next(payload for tag, payload in page_chunks((output / "palette.djvu").read_bytes()) if tag == b"INFO")
        for name in ["comments-page", "zero-area", "clipped-masks"]:
            width, height = (2, 65535) if name == "clipped-masks" else (128, 96)
            body = b"DJVU" + chunk(b"INFO", struct.pack(">HH", width, height) + info[4:])
            body += chunk(b"Sjbz", (output / (name + ".jb2")).read_bytes())
            (output / (name + ".djvu")).write_bytes(b"AT&T" + chunk(b"FORM", body))
            reference(name)

    manifest = [{"path": str(path.relative_to(output)), "bytes": path.stat().st_size,
                 "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}
                for path in sorted(output.rglob("*")) if path.is_file()]
    (output / "production-manifest.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
