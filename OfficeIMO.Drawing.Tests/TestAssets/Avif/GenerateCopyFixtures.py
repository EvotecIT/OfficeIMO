#!/usr/bin/env python3
"""Opt-in still-copy grammar oracle; pinned AOM entropy/defaults, no runtime dependency."""
import argparse
import base64
import gzip
import json
import pathlib
import re
import shutil
import subprocess
import sys
import urllib.request
sys.dont_write_bytecode = True
from GenerateEntropyFixtures import sha256, write_json

SOURCE_HASH = '3472e909bbee628a2c39d91f73d1e2502c7c4d47dc488c633b0ec3718a990f35'


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--work-dir', required=True, type=pathlib.Path)
    p.add_argument('--output', type=pathlib.Path)
    p.add_argument('--cc', default='clang')
    args = p.parse_args()
    here = pathlib.Path(__file__).resolve().parent
    work = args.work_dir.resolve()
    subprocess.run([sys.executable, str(here / 'GenerateEntropyFixtures.py'), '--work-dir', str(work), '--cc', args.cc], check=True)
    source = work / 'entropymv.c'
    if not source.exists():
        url = 'https://aomedia.googlesource.com/aom/+/refs/tags/v3.13.1/av1/common/entropymv.c?format=TEXT'
        source.write_bytes(base64.b64decode(urllib.request.urlopen(url, timeout=60).read(), validate=True))
    if sha256(source.read_bytes()) != SOURCE_HASH:
        raise ValueError('Motion defaults source hash mismatch')
    text = source.read_text()
    declaration = re.search(r'static const nmv_context default_nmv_context = (\{.*?\n\});', text, re.S)
    if not declaration:
        raise ValueError('Missing pinned motion defaults')
    types = ('typedef struct {aom_cdf_prob classes_cdf[12], class0_fp_cdf[2][5], fp_cdf[5], '
             'sign_cdf[3], class0_hp_cdf[3], hp_cdf[3], class0_cdf[3], bits_cdf[10][3];} nmv_component;\n'
             'typedef struct {aom_cdf_prob joints_cdf[5];nmv_component components[2];} nmv_context;\n')
    (work / 'copy-defaults.h').write_text(text[:text.index('*/')+2]+'\n'+types+'static const nmv_context default_nmv_context = '+declaration[1]+';\n')
    native = work / 'native-aom'
    compiler = shutil.which(args.cc)
    if not compiler:
        raise FileNotFoundError(args.cc)
    harness = here / 'GenerateCopyFixtures.c'
    executable = work / 'generate-copy'
    command = [compiler, '-std=c11', '-Wall', '-Wextra', '-Werror', '-O2', '-I', str(native), '-I', str(work), str(harness)]
    command += [str(native / 'aom_dsp' / f) for f in ('entenc.c', 'entdec.c', 'entcode.c')]
    subprocess.run(command+['-o', str(executable)], check=True)
    raw = work / 'copy-reference.json'
    subprocess.run([str(executable), str(raw)], check=True)
    values = json.loads(raw.read_bytes())
    canonical = (json.dumps(values, separators=(',', ':'))+'\n').encode()
    output = args.output.resolve() if args.output else work / 'copy-reference.json.gz'
    output.write_bytes(gzip.compress(canonical, compresslevel=9, mtime=0))
    receipt = {'tableSourceSha256': SOURCE_HASH, 'harnessSha256': sha256(harness.read_bytes()),
               'generatorSha256': sha256(pathlib.Path(__file__).read_bytes()), 'fixtureSha256': sha256(output.read_bytes()),
               'canonicalJsonSha256': sha256(canonical), 'cases': len(values['cases']),
               'leaves': sum(len(c['states']) for c in values['cases']),
               'copyLeaves': sum(s[4] for c in values['cases'] for s in c['states']),
               'rejectionCases': len(values['rejections']), 'sourceProbes': len(values['sourceProbes']),
               'nativeSelfCheck': values['nativeSelfCheck']}
    write_json(work / 'copy-oracle-receipt.json', receipt)
    print('Generated '+str(output)+' SHA256 '+receipt['fixtureSha256'])


if __name__ == '__main__':
    main()
