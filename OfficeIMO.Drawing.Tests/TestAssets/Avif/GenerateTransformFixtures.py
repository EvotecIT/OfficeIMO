#!/usr/bin/env python3
"""Opt-in transform-size/traversal oracle over pinned AOM entropy/defaults.
Requires task-owned scratch. Synthetic copy-mode inputs omit motion grammar;
real frozen prefixes stop before coefficient/type syntax and never finish a tile.
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
    args = parser.parse_args()
    here = pathlib.Path(__file__).resolve().parent
    work = args.work_dir.resolve()
    subprocess.run([sys.executable, str(here / 'GeneratePaletteFixtures.py'), '--work-dir', str(work), '--cc', args.cc], check=True)
    text = (work / 'entropymode.c').read_text()
    declarations = []
    for name, shape in [('default_tx_size_cdf', '[4][3][4]'), ('default_txfm_partition_cdf', '[21][3]')]:
        match = re.search(r'static const aom_cdf_prob\s+' + name + r'.*?=\s*(\{.*?\});', text, re.S)
        if not match:
            raise ValueError('Missing pinned table: ' + name)
        declarations.append('static const aom_cdf_prob ' + name + shape + ' = ' + match[1] + ';')
    (work / 'transform-defaults.h').write_text(text[:text.index('*/') + 2] + '\n' + '\n'.join(declarations) + '\n')
    compiler = shutil.which(args.cc)
    if not compiler:
        raise FileNotFoundError(args.cc)
    native = work / 'native-aom'
    harness = here / 'GenerateTransformFixtures.c'
    executable = work / 'generate-transforms'
    command = [compiler, '-std=c11', '-Wall', '-Wextra', '-Werror', '-O2', '-I', str(native), '-I', str(work), str(harness)]
    command += [str(native / 'aom_dsp' / name) for name in ('entenc.c', 'entdec.c', 'entcode.c')]
    subprocess.run(command + ['-o', str(executable)], check=True)
    raw = work / 'native-transforms.json'
    subprocess.run([str(executable), str(raw)], check=True)
    vectors = json.loads(raw.read_text())
    depth = [sum(c['depthContexts'][i] for c in vectors['cases']) for i in range(12)]
    split = [sum(c['splitContexts'][i] for c in vectors['cases']) for i in range(21)]
    if not all(depth) or not all(split):
        raise ValueError('Missing transform contexts: ' + str((depth, split)))
    entropy = json.loads((here / 'entropy-oracle.json').read_text())
    geometry = json.loads((here / 'partition-oracle.json').read_text())['prefixGeometry']
    preludes = json.loads((work / 'prelude-reference.json').read_text())['framePrefixes']
    prefixes = []
    for frame in entropy['framePrefixes']:
        g = geometry[frame['name']]
        b = next(v for v in preludes if v['name'] == frame['name'])
        assert b['skip'] == 0 and b['baseQ'] > 0
        # Independent retained FFmpeg frame-header traces: largest mode for color/alpha, SELECT for multi-tile.
        mode = 2 if frame['name'] == 'multitile' else 1
        screen = frame['name'] == 'avif-opaque'
        output = work / (frame['name'] + '-transforms.json')
        subprocess.run([str(executable), str(here.parents[2] / frame['path']), str(frame['offset']), str(frame['length']),
                        str(g['miRows']), str(g['miCols']), str(b['baseQ']), str(int(b['cdef'])), str(int(b['deltaQ'])),
                        str(output), str(int(frame['alpha'])), str(int(screen)), str(mode)], check=True)
        prefix = {key: frame[key] for key in ('name', 'alpha', 'offset', 'length')}
        prefix['mode'] = mode
        prefix.update(json.loads(output.read_text()))
        prefixes.append(prefix)
    vectors['framePrefixes'] = prefixes
    output = args.output.resolve() if args.output else work / 'transform-reference.json'
    encoded = json.dumps(vectors, indent=2)
    encoded = re.sub(r'\{\s*"block":.*?"residuals":\s*\[.*?\]\s*\}',
                     lambda match: json.dumps(json.loads(match[0]), separators=(',', ':')), encoded, flags=re.S)
    encoded = re.sub(r'\[\s*(-?\d+(?:,\s*-?\d+)+)\s*\]',
                     lambda match: '[' + re.sub(r'\s+', '', match[1]) + ']', encoded)
    output.write_text(encoded + '\n')
    receipt = {'generatorSha256': sha256(pathlib.Path(__file__).read_bytes()), 'harnessSha256': sha256(harness.read_bytes()),
               'sharedPaletteHarnessSha256': sha256((here / 'GeneratePaletteFixtures.c').read_bytes()),
               'tableSourceSha256': sha256((work / 'entropymode.c').read_bytes()),
               'paletteReceiptSha256': sha256((work / 'palette-oracle-receipt.json').read_bytes()),
               'fixtureSha256': sha256(output.read_bytes()), 'cases': len(vectors['cases']),
               'leaves': sum(len(c['states']) for c in vectors['cases']), 'depthContexts': depth, 'splitContexts': split,
               'nativeSelfCheck': vectors['nativeSelfCheck'], 'prefixes': len(prefixes)}
    write_json(work / 'transform-oracle-receipt.json', receipt)
    print('Generated ' + str(output) + ' SHA256 ' + receipt['fixtureSha256'])


if __name__ == '__main__':
    main()
