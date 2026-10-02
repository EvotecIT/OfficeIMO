"""Opt-in Main 8/10-bit AVIF fixtures using libavif 1.4.2, AOM and independent dav1d.

Normal builds use immutable assets; the native library is never a runtime dependency.
Decoded YUV planes are little-endian 16-bit samples in Y/U/V/alpha order, omitting
absent planes. The accompanying RGBA reference uses libavif's 8-bit converter.
For --full-headers, use a private libavif build whose codec_aom.c sets
cfg->full_still_picture_hdr = 1 immediately before aom_codec_enc_init. This only
selects AOM's native header mode; leave AOM arithmetic and dav1d unchanged.
"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import urllib.request

COMMIT = "c5240fc79fe5c2407e10afd35f5505ef6333ea49"
HEADER_SHA = "43ec2563fcc739a4d5ca18e22f74bc15e76a11f8c9a82bd7e1ee572bff7afc04"


def sha(data):
    return hashlib.sha256(data).hexdigest()


def sequence_flags(data):
    """Inspect every native sequence OBU in mdat, including an auxiliary alpha item."""
    p = data.index(b"mdat") + 4
    flags = []
    while p < len(data):
        header = data[p]
        p += 1
        if header & 0x81 or not header & 2:
            raise RuntimeError("Unexpected producer OBU header")
        if header & 4:
            if data[p] != 0:
                raise RuntimeError("Layered producer OBU")
            p += 1
        size = 0
        for i in range(8):
            b = data[p]
            p += 1
            size |= (b & 127) << (7 * i)
            if not b & 128:
                break
        else:
            raise RuntimeError("Unterminated producer OBU size")
        if p + size > len(data):
            raise RuntimeError("Truncated producer OBU")
        if (header >> 3) & 15 == 1:
            if not size:
                raise RuntimeError("Empty producer sequence")
            flags.append(data[p] & 24)
        p += size
    return flags


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--work-dir", type=Path, required=True)
    parser.add_argument("--library", type=Path, required=True)
    parser.add_argument("--full-headers", action="store_true", help="Generate both depths with native non-reduced still headers")
    args = parser.parse_args()
    work, library = args.work_dir.resolve(), args.library.resolve(strict=True)
    work.mkdir(parents=True, exist_ok=True)
    include = work / "include/avif"
    include.mkdir(parents=True, exist_ok=True)
    header = urllib.request.urlopen(
        f"https://raw.githubusercontent.com/AOMediaCodec/libavif/{COMMIT}/include/avif/avif.h").read()
    if sha(header) != HEADER_SHA:
        raise RuntimeError("Pinned libavif header changed")
    (include / "avif.h").write_bytes(header)
    driver = Path(__file__).with_name("EncodeMain10Fixtures.c")
    exe = work / "encode-main10-fixtures"
    subprocess.run(["cc", "-O2", "-Wall", "-Wextra", "-I", str(work / "include"), str(driver),
                    str(library), "-Wl,-rpath," + str(library.parent), "-o", str(exe)], check=True)
    environment = os.environ.copy()
    if sys.platform == "darwin":
        environment["DYLD_LIBRARY_PATH"] = str(library.parent)
    generated = work / "generated"
    generated.mkdir(exist_ok=True)
    receipt = None
    for depth in ([8, 10] if args.full_headers else [10]):
        command = [str(exe), str(generated)]
        if args.full_headers:
            command += [str(depth), "-fullheader"]
        produced = json.loads(subprocess.check_output(command, env=environment))
        if receipt is None:
            receipt = produced
        else:
            if receipt["codecs"] != produced["codecs"]:
                raise RuntimeError("Producer codec version changed")
            receipt["cases"].extend(produced["cases"])
    if receipt["libavif"] != "1.4.2" or len(receipt["cases"]) != (16 if args.full_headers else 8):
        raise RuntimeError("Unexpected independent producer corpus")
    assets = Path(__file__).resolve().parents[3] / "OfficeIMO.TestAssets/Documents/Html/Qualification" / ("AvifFullHeaders" if args.full_headers else "AvifMain10")
    assets.mkdir(parents=True, exist_ok=True)
    for case in receipt["cases"]:
        name = case["name"]
        data = (generated / (name + ".avif")).read_bytes()
        config = data.index(b"av1C") + 4
        pixi = data.index(b"pixi") + 8
        channels = 1 if case["monochrome"] else 3
        depth = case["bitDepth"]
        if data[config + 1] >> 5 != 0 or data[config + 2] & 0x60 != (0x40 if depth == 10 else 0) or data[pixi] != channels:
            raise RuntimeError("Fixture is not the requested AV1 Main profile")
        if any(data[pixi + i + 1] != depth for i in range(channels)):
            raise RuntimeError("Fixture does not declare the requested channel depth")
        flags = sequence_flags(data)
        if flags != [16 if args.full_headers else 24] * (2 if case["hasAlpha"] else 1):
            raise RuntimeError("Native producer did not select the requested still-header mode")
        case["sequenceFlagBits"] = flags
        expected_samples = 49 * 33 + (0 if case["monochrome"] else 2 * 25 * 17) + (49 * 33 if case["hasAlpha"] else 0)
        case["av1C"] = data[config:config + 4].hex()
        case["files"] = {}
        for extension in ("avif", "rgba", "yuv16", "source-yuv16"):
            content = (generated / (name + "." + extension)).read_bytes()
            if extension == "rgba" and len(content) != 49 * 33 * 4:
                raise RuntimeError("Independent RGBA length mismatch")
            if extension.endswith("yuv16") and len(content) != expected_samples * 2:
                raise RuntimeError("Independent plane length mismatch")
            case["files"][extension] = dict(bytes=len(content), sha256=sha(content))
            if extension != "source-yuv16":
                (assets / (name + "." + extension)).write_bytes(content)
    receipt.update(nativeSourceCommit=COMMIT, headerSha256=HEADER_SHA,
                   nativeLibrarySha256=sha(library.read_bytes()), driverSha256=sha(driver.read_bytes()),
                   generatorSha256=sha(Path(__file__).read_bytes()),
                   encoderOptions=dict(codec="aom", quality=85, qualityAlpha=100, speed=6, maxThreads=1),
                   decoderOptions=dict(codec="dav1d", rgbDepth=8, avoidLibYUV=True, chromaUpsampling="bilinear"))
    (work / ("fullheader-oracle-receipt.json" if args.full_headers else "main10-oracle-receipt.json")).write_text(json.dumps(receipt, indent=2) + "\n")
    print(json.dumps(receipt))


if __name__ == "__main__":
    main()
