"""Opt-in libavif 1.3.0 RGBA oracle. Requires a matching native library and C compiler.

Normal OfficeIMO builds consume the checked-in fixture and do not link libavif.
The driver invokes the actual native converter with bilinear chroma and libyuv disabled.
"""
import argparse, gzip, hashlib, json, os, subprocess, sys, urllib.request
from pathlib import Path

ASSETS = Path(__file__).resolve().parent
COMMIT = '1aadfad932c98c069a1204261b1856f81f3bc199'
HEADER_SHA = 'ece1a0ab723ae006b72a191bc9bb3ae55e90d245e61356dac847185d85635d17'

def sha(data):
    return hashlib.sha256(data).hexdigest()

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--work-dir', type=Path, required=True)
    parser.add_argument('--library', type=Path, required=True)
    args = parser.parse_args()
    work, library = args.work_dir.resolve(), args.library.resolve(strict=True)
    work.mkdir(parents=True, exist_ok=True)
    include = work / 'include' / 'avif'; include.mkdir(parents=True, exist_ok=True)
    header = urllib.request.urlopen(f'https://raw.githubusercontent.com/AOMediaCodec/libavif/{COMMIT}/include/avif/avif.h').read()
    if sha(header) != HEADER_SHA:
        raise RuntimeError('Pinned native header changed')
    (include / 'avif.h').write_bytes(header)
    driver = ASSETS / 'ReadColorReference.c'
    exe = work / 'read-color-reference'
    subprocess.run(['cc', '-O2', '-I', str(work / 'include'), str(driver), str(library),
                    '-Wl,-rpath,' + str(library.parent), '-o', str(exe)], check=True)
    environment = os.environ.copy()
    # Wheel libraries may retain their build-time install name; lookup is scoped to this oracle process.
    if sys.platform == 'darwin':
        environment['DYLD_LIBRARY_PATH'] = str(library.parent)
    raw = subprocess.check_output([str(exe)], env=environment)
    fixture = json.loads(raw)
    if fixture['version'] != '1.3.0' or len(fixture['cases']) != 168:
        raise RuntimeError('Unexpected native version or corpus')
    canonical = json.dumps(fixture, separators=(',', ':'), sort_keys=True).encode()
    output = gzip.compress(canonical, mtime=0)
    (ASSETS / 'color-reference.json.gz').write_bytes(output)
    receipt = {'nativeVersion': fixture['version'], 'nativeSourceCommit': COMMIT,
               'headerSha256': sha(header), 'librarySha256': sha(library.read_bytes()),
               'driverSha256': sha(driver.read_bytes()), 'generatorSha256': sha(Path(__file__).read_bytes()),
               'fixtureSha256': sha(output), 'canonicalJsonSha256': sha(canonical), 'cases': len(fixture['cases']),
               'rgbTolerance': 1, 'alphaTolerance': 0, 'avoidLibYUV': True, 'chromaUpsampling': 'bilinear'}
    (work / 'color-oracle-receipt.json').write_text(json.dumps(receipt, indent=2) + '\n')
    print(json.dumps(receipt))

if __name__ == '__main__':
    main()
