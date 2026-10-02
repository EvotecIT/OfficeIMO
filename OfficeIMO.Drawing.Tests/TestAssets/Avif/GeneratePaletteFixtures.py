#!/usr/bin/env python3
"""Opt-in palette/filter-intra oracle using pinned AOM entropy, defaults and native
color-index context. Require task-owned --work-dir; no complete AV1 pixels are decoded.
"""
import argparse
import json
import pathlib
import re
import shutil
import subprocess
import sys
sys.dont_write_bytecode = True
from GenerateEntropyFixtures import sha256, write_json


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--work-dir', required=True, type=pathlib.Path)
    parser.add_argument('--output', type=pathlib.Path)
    parser.add_argument('--cc', default='clang')
    parser.add_argument('--bit-depth', type=int, choices=(8, 10), default=8)
    parser.add_argument('--native-source', type=pathlib.Path)
    parser.add_argument('--native-build', type=pathlib.Path)
    args = parser.parse_args()
    here = pathlib.Path(__file__).resolve().parent
    work = args.work_dir.resolve()
    if bool(args.native_source) != bool(args.native_build):
        raise ValueError('Supply both native source and build')
    if args.bit_depth == 10 and not args.native_source:
        raise ValueError('Ten-bit fixtures require the actual native decoder source and build')
    native = args.native_source.resolve() if args.native_source else work / 'native-aom'
    if args.native_source:
        from GeneratePaletteNativeReference import prepare
        prepare(here, work, native, args.native_build.resolve())
    else:
        subprocess.run([sys.executable, str(here / 'GenerateModeFixtures.py'), '--work-dir', str(work), '--cc', args.cc], check=True)
    text = (work / 'entropymode.c').read_text()
    tables = [('default_palette_y_mode_cdf', '[7][3][3]'), ('default_palette_uv_mode_cdf', '[2][3]'),
              ('default_palette_y_size_cdf', '[7][8]'), ('default_palette_uv_size_cdf', '[7][8]'),
              ('default_palette_y_color_index_cdf', '[7][5][9]'), ('default_palette_uv_color_index_cdf', '[7][5][9]'),
              ('default_filter_intra_cdfs', '[22][3]'), ('default_filter_intra_mode_cdf', '[6]')]
    declarations = []
    for name, shape in tables:
        match = re.search(r'static const aom_cdf_prob\s+' + name + r'.*?=\s*(\{.*?\});', text, re.S)
        if not match:
            raise ValueError('Missing pinned table: ' + name)
        declarations.append('static const aom_cdf_prob ' + name + shape + ' = ' + match[1] + ';')
    context = text[text.index('const int av1_palette_color_index_context_lookup'):text.index('void av1_init_mode_probs')]
    constants = '#define PALETTE_MAX_SIZE 8\n#define NUM_PALETTE_NEIGHBORS 3\n#define MAX_COLOR_CONTEXT_HASH 8\n#define PALETTE_COLOR_INDEX_CONTEXTS 5\n'
    (work / 'palette-defaults.h').write_text(text[:text.index('*/') + 2] + '\n' + '\n'.join(declarations) + '\n' + constants + context)
    compiler = shutil.which(args.cc)
    if not compiler:
        raise FileNotFoundError(args.cc)
    harness = here / 'GeneratePaletteFixtures.c'
    executable = work / 'generate-palettes'
    command = [compiler, '-std=c11', '-Wall', '-Wextra', '-Werror', '-O2', '-I', str(native), '-I', str(work), str(harness)]
    if args.native_source:
        command += ['-I', str(args.native_build.resolve())]
    command += [str(native / 'aom_dsp' / name) for name in ('entenc.c', 'entdec.c', 'entcode.c')]
    subprocess.run(command + ['-o', str(executable)], check=True)
    raw = work / 'native-palettes.json'
    subprocess.run([str(executable), str(raw), str(args.bit_depth)], check=True)
    vectors = json.loads(raw.read_text())
    contexts = [sum(c['mapContexts'][i] for c in vectors['cases']) for i in range(70)]
    # Context 1 requires three distinct neighbors and cannot occur for a two-color palette.
    if any(value == 0 for index, value in enumerate(contexts) if index not in (1, 36)):
        raise ValueError('Missing palette size/map context coverage: ' + str(contexts))
    entropy = json.loads((here / 'entropy-oracle.json').read_text())
    geometry = json.loads((here / 'partition-oracle.json').read_text())['prefixGeometry']
    preludes = json.loads(((here if args.native_source else work) / 'prelude-reference.json').read_text())['framePrefixes']
    prefixes = []
    for frame in entropy['framePrefixes'] if args.bit_depth == 8 else []:
        g = geometry[frame['name']]
        b = next(v for v in preludes if v['name'] == frame['name'])
        # Independent FFmpeg header traces: filter-intra=1 on all items; screen-content only on the color item.
        screen = frame['name'] == 'avif-opaque'
        output = work / (frame['name'] + '-palettes.json')
        subprocess.run([str(executable), str(here.parents[2] / frame['path']), str(frame['offset']), str(frame['length']),
                        str(g['miRows']), str(g['miCols']), str(b['baseQ']), str(int(b['cdef'])), str(int(b['deltaQ'])),
                        str(output), str(int(frame['alpha'])), str(int(screen))], check=True)
        prefix = {key: frame[key] for key in ('name', 'alpha', 'offset', 'length')}
        prefix['screen'] = screen
        prefix.update(json.loads(output.read_text()))
        prefixes.append(prefix)
    if args.bit_depth == 8:
        vectors['framePrefixes'] = prefixes
    else:
        vectors['bitDepth'] = args.bit_depth
    actual = None
    if args.native_source:
        from GeneratePaletteNativeReference import verify
        actual = verify(here, work, native, args.native_build.resolve(), compiler, vectors, args.bit_depth)
    output = args.output.resolve() if args.output else work / 'palette-reference.json'
    # Keep one leaf per line; its colors/maps are numeric fixture data, not prose.
    encoded = json.dumps(vectors, indent=2)
    encoded = re.sub(r'\{\s*"block":.*?"mapUv":\s*"[^"]*"\s*\}',
                     lambda match: json.dumps(json.loads(match[0]), separators=(',', ':')), encoded, flags=re.S)
    encoded = re.sub(r'\[\s*(-?\d+(?:,\s*-?\d+)+)\s*\]',
                     lambda match: '[' + re.sub(r'\s+', '', match[1]) + ']', encoded)
    output.write_text(encoded + '\n')
    receipt = {'generatorSha256': sha256(pathlib.Path(__file__).read_bytes()), 'harnessSha256': sha256(harness.read_bytes()),
               'sharedModeHarnessSha256': sha256((here / 'GenerateModeFixtures.c').read_bytes()),
               'tableSourceSha256': sha256((work / 'entropymode.c').read_bytes()),
               'fixtureSha256': sha256(output.read_bytes()), 'cases': len(vectors['cases']),
               'nativeSelfCheck': vectors['nativeSelfCheck'], 'prefixes': len(prefixes)}
    if actual:
        receipt['nativeDecoder'] = actual
    else:
        receipt['modeReceiptSha256'] = sha256((work / 'mode-oracle-receipt.json').read_bytes())
    write_json(work / 'palette-oracle-receipt.json', receipt)
    print('Generated ' + str(output) + ' SHA256 ' + receipt['fixtureSha256'])


if __name__ == '__main__':
    main()
