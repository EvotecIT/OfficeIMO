#!/usr/bin/env python3
"""Opt-in AV1 leaf-prelude reference streams. Uses pinned AOM tables/entropy,
not a full AV1 decoder. Require task-owned --work-dir; never run during builds.
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
    subprocess.run([sys.executable, str(here / 'GeneratePartitionFixtures.py'), '--work-dir', str(work), '--cc', args.cc], check=True)
    text = (work / 'entropymode.c').read_text(encoding='utf-8')
    tables = [('default_skip_txfm_cdfs', '[3][3]'), ('default_spatial_pred_seg_tree_cdf', '[3][9]'),
              ('default_delta_q_cdf', '[5]'), ('default_delta_lf_cdf', '[5]'), ('default_delta_lf_multi_cdf', '[4][5]')]
    declarations = []
    for name, shape in tables:
        m = re.search(r'static const aom_cdf_prob\s+' + name + r'.*?=\s*(\{.*?\});', text, re.S)
        if not m:
            raise ValueError('Missing pinned AOM table: ' + name)
        declarations.append('static const aom_cdf_prob ' + name + shape + ' = ' + m[1] + ';')
    (work / 'prelude-defaults.h').write_text(text[:text.index('*/') + 2] + '\n' + '\n'.join(declarations) + '\n', encoding='utf-8')
    compiler = shutil.which(args.cc)
    if not compiler:
        raise FileNotFoundError(args.cc)
    native = work / 'native-aom'
    harness = here / 'GeneratePreludeFixtures.c'
    executable = work / 'generate-preludes'
    cmd = [compiler, '-std=c11', '-Wall', '-Wextra', '-Werror', '-O2', '-I', str(native), '-I', str(work), str(harness)]
    cmd += [str(native / 'aom_dsp' / name) for name in ('entenc.c', 'entdec.c', 'entcode.c')]
    subprocess.run(cmd + ['-o', str(executable)], check=True)
    raw = work / 'native-preludes.json'
    subprocess.run([str(executable), str(raw)], check=True)
    vectors = json.loads(raw.read_text())
    manifest = json.loads((here / 'entropy-oracle.json').read_text())
    geometry = json.loads((here / 'partition-oracle.json').read_text())['prefixGeometry']
    prefixes = []
    # Frame flags independently traced by FFmpeg trace_headers; zero CDEF bits in the only enabled fixture.
    settings = {'avif-opaque': (32, 1, 1), 'avif-alpha': (24, 0, 0), 'multitile': (112, 0, 0)}
    for f in manifest['framePrefixes']:
        g = geometry[f['name']]
        q, cdef, delta = settings[f['name']]
        result = work / (f['name'] + '-prelude.json')
        subprocess.run([str(executable), str(here.parents[2] / f['path']), str(f['offset']), str(f['length']),
                        str(g['miRows']), str(g['miCols']), str(q), str(cdef), str(delta), str(result)], check=True)
        prefix = {key: f[key] for key in ('name', 'alpha', 'offset', 'length')}
        prefix.update({'baseQ': q, 'cdef': bool(cdef), 'deltaQ': bool(delta)})
        prefix.update(json.loads(result.read_text()))
        prefixes.append(prefix)
    vectors['framePrefixes'] = prefixes
    output = args.output.resolve() if args.output else work / 'prelude-reference.json'
    write_json(output, vectors)
    receipt = {'generatorSha256': sha256(pathlib.Path(__file__).read_bytes()), 'harnessSha256': sha256(harness.read_bytes()),
               'tableSourceSha256': sha256((work / 'entropymode.c').read_bytes()),
               'partitionReceiptSha256': sha256((work / 'partition-oracle-receipt.json').read_bytes()),
               'fixtureSha256': sha256(output.read_bytes()), 'cases': len(vectors['cases']),
               'nativeSelfCheck': vectors['nativeSelfCheck'], 'prefixes': len(prefixes)}
    write_json(work / 'prelude-oracle-receipt.json', receipt)
    print('Generated ' + str(output) + ' SHA256 ' + receipt['fixtureSha256'])


if __name__ == '__main__':
    main()
