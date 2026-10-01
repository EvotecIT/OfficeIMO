"""Opt-in primary grayscale AVIF fixtures from Pillow 12.3.0 / libavif 1.4.2.

Normal builds use the checked-in files; neither native producer is a runtime dependency.
The full/limited-range items encode the same odd-sized 8-bit source using unpatched AOM.
"""
import argparse
import hashlib
import json
from pathlib import Path

import PIL
from PIL import Image, _avif


def sha(data):
    return hashlib.sha256(data).hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--work-dir", type=Path, required=True)
    args = parser.parse_args()
    if PIL.__version__ != "12.3.0" or _avif.libavif_version != "1.4.2":
        raise RuntimeError("Use the declared independent producer versions")
    if not _avif.encoder_codec_available("aom"):
        raise RuntimeError("The declared AOM encoder is unavailable")
    work = args.work_dir.resolve()
    work.mkdir(parents=True, exist_ok=True)
    root = Path(__file__).resolve().parents[3]
    assets = root / "OfficeIMO.TestAssets/Documents/Html/Qualification/AvifMonochrome"
    assets.mkdir(parents=True, exist_ok=True)
    width, height = 49, 33
    source = bytes((x * 17 + y * 29 + (x // 7) * 41) % 256
                   for y in range(height) for x in range(width))
    opaque = Image.frombytes("L", (width, height), source)
    alpha_source = bytes((x * 13 + y * 7) % 256 for y in range(height) for x in range(width))
    translucent = opaque.convert("RGBA")
    translucent.putalpha(Image.frombytes("L", (width, height), alpha_source))
    options = dict(codec="aom", quality=85, speed=6, max_threads=1,
                   subsampling="4:0:0", autotiling=False, tile_rows=0, tile_cols=0)
    cases = []
    for range_name, has_alpha in (("full", False), ("limited", False), ("full", True), ("limited", True)):
        name = "avif-monochrome-" + range_name + ("-alpha" if has_alpha else "")
        image = translucent if has_alpha else opaque
        encoded = assets / (name + ".avif")
        image.save(encoded, format="AVIF", range=range_name, **options)
        with Image.open(encoded) as decoded:
            if decoded.size != (width, height) or decoded.n_frames != 1:
                raise RuntimeError("Unexpected independent decoded selection")
            rgba = decoded.convert("RGBA").tobytes()
        if any(rgba[p] != rgba[p + 1] or rgba[p] != rgba[p + 2]
               or (not has_alpha and rgba[p + 3] != 255) for p in range(0, len(rgba), 4)):
            raise RuntimeError("Independent decoder did not produce grayscale")
        (assets / (name + ".rgba")).write_bytes(rgba)
        data = encoded.read_bytes()
        config = data.index(b"av1C") + 4
        pixi = data.index(b"pixi") + 8
        nclx = data.index(b"nclx") + 4
        if data[config + 2] != 0x1c or data[pixi] != 1 or data[pixi + 1] != 8:
            raise RuntimeError("Fixture is not Main-8 single-plane AVIF")
        cases.append(dict(name=name, encodedBytes=len(data), encodedSha256=sha(data),
                          rgbaSha256=sha(rgba), av1C=data[config:config + 4].hex(),
                          pixelChannels=data[pixi], bitDepth=data[pixi + 1],
                          nclx=data[nclx:nclx + 7].hex(), range=range_name, hasAlpha=has_alpha))
    receipt = dict(pillow=PIL.__version__, libavif=_avif.libavif_version,
                   nativeLibrarySha256=sha(Path(_avif.__file__).read_bytes()),
                   generatorSha256=sha(Path(__file__).read_bytes()),
                   sourceModes=["L", "RGBA"], width=width, height=height, sourceSha256=sha(source),
                   alphaSourceSha256=sha(alpha_source),
                   encoderOptions=options, cases=cases)
    (work / "monochrome-oracle-receipt.json").write_text(json.dumps(receipt, indent=2) + "\n")
    print(json.dumps(receipt))


if __name__ == "__main__":
    main()
