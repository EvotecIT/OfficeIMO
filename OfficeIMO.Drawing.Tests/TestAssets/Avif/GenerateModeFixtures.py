#!/usr/bin/env python3
"""Opt-in intra-mode component oracle using pinned AOM data/entropy. Require
--work-dir for task-owned scratch. This is not a complete AV1 pixel decoder.
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
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--work-dir', required=True, type=pathlib.Path)
    p.add_argument('--output', type=pathlib.Path)
    p.add_argument('--cc', default='clang')
    args = p.parse_args()
    here = pathlib.Path(__file__).resolve().parent
    work = args.work_dir.resolve()
    subprocess.run([sys.executable, str(here / 'GeneratePreludeFixtures.py'), '--work-dir', str(work), '--cc', args.cc], check=True)
    text = (work / 'entropymode.c').read_text(encoding='utf-8')
    tables = [('default_kf_y_mode_cdf', '[5][5][14]'), ('default_uv_mode_cdf', '[2][13][15]'),
              ('default_angle_delta_cdf', '[8][8]'), ('default_cfl_sign_cdf', '[9]'),
              ('default_cfl_alpha_cdf', '[6][17]'), ('default_intrabc_cdf', '[3]')]
    declarations = []
    for name, shape in tables:
        m = re.search(r'static const aom_cdf_prob\s+' + name + r'.*?=\s*(\{.*?\});', text, re.S)
        if not m:
            raise ValueError('Missing pinned AOM table: ' + name)
        declarations.append('static const aom_cdf_prob ' + name + shape + ' = ' + m[1] + ';')
    (work / 'mode-defaults.h').write_text(text[:text.index('*/') + 2] + '\n' + '\n'.join(declarations) + '\n', encoding='utf-8')
    compiler = shutil.which(args.cc)
    if not compiler:
        raise FileNotFoundError(args.cc)
    native = work / 'native-aom'
    harness = here / 'GenerateModeFixtures.c'
    executable = work / 'generate-modes'
    cmd = [compiler, '-std=c11', '-Wall', '-Wextra', '-Werror', '-O2', '-I', str(native), '-I', str(work), str(harness)]
    cmd += [str(native / 'aom_dsp' / name) for name in ('entenc.c', 'entdec.c', 'entcode.c')]
    subprocess.run(cmd + ['-o', str(executable)], check=True)
    raw = work / 'native-modes.json'
    subprocess.run([str(executable), str(raw)], check=True)
    vectors = json.loads(raw.read_text())
    visits = [sum(c['lumaContextVisits'][i] for c in vectors['cases']) for i in range(25)]
    if min(visits) == 0:
        raise ValueError('Component streams do not cover all luma contexts')
    entropy = json.loads((here / 'entropy-oracle.json').read_text())
    geometry = json.loads((here / 'partition-oracle.json').read_text())['prefixGeometry']
    preludes = json.loads((work / 'prelude-reference.json').read_text())['framePrefixes']
    prefixes = []
    for f in entropy['framePrefixes']:
        g = geometry[f['name']]
        b = next(v for v in preludes if v['name'] == f['name'])
        # allow_intrabc is zero in all three independently traced frame headers.
        result = work / (f['name'] + '-modes.json')
        subprocess.run([str(executable), str(here.parents[2] / f['path']), str(f['offset']), str(f['length']),
                        str(g['miRows']), str(g['miCols']), str(b['baseQ']), str(int(b['cdef'])), str(int(b['deltaQ'])),
                        str(result), str(int(f['alpha']))], check=True)
        prefix = {key: f[key] for key in ('name', 'alpha', 'offset', 'length')}
        prefix.update(json.loads(result.read_text()))
        prefixes.append(prefix)
    vectors['framePrefixes'] = prefixes
    output = args.output.resolve() if args.output else work / 'mode-reference.json'
    # Keep each state row on one line; rows are reproducible numeric data, not prose.
    encoded = json.dumps(vectors, indent=2)
    encoded = re.sub(r'\[\s*(-?\d+(?:,\s*-?\d+)+)\s*\]', lambda m: '[' + re.sub(r'\s+', '', m[1]) + ']', encoded)
    output.write_text(encoded + '\n', encoding='utf-8')
    receipt = {'generatorSha256': sha256(pathlib.Path(__file__).read_bytes()), 'harnessSha256': sha256(harness.read_bytes()),
               'sharedPreludeHarnessSha256': sha256((here / 'GeneratePreludeFixtures.c').read_bytes()),
               'tableSourceSha256': sha256((work / 'entropymode.c').read_bytes()),
               'preludeReceiptSha256': sha256((work / 'prelude-oracle-receipt.json').read_bytes()),
               'fixtureSha256': sha256(output.read_bytes()), 'cases': len(vectors['cases']),
               'nativeSelfCheck': vectors['nativeSelfCheck'], 'prefixes': len(prefixes)}
    write_json(work / 'mode-oracle-receipt.json', receipt)
    print('Generated ' + str(output) + ' SHA256 ' + receipt['fixtureSha256'])


if __name__ == '__main__':
    main()
