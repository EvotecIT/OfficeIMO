"""Access-only observer for the actual pinned AOM palette decoder.

The original component harness writes deterministic palette streams. This module
checks every returned color and padded map against separate native decoder paths.
It never runs during normal builds or product execution.
"""
import json
import pathlib
import re
import shutil
import struct
import subprocess
from GenerateEntropyFixtures import sha256

SOURCES = {
    'av1/decoder/decodemv.c': '86fb66757d52ea259eaf4bb654cdee0973daf4797601b892591d55754b8f34ed',
    'av1/decoder/detokenize.c': 'ac5139ecd526d22763bf72d4d6f26b7e0442573d7f7bd6874c560d22855a528c',
    'av1/common/pred_common.c': '027a67e686ca5f4340bd11b6a55007d64993c1f6c68b5857927e52746c486636',
    'av1/common/entropymode.c': '13c47672c1e00d77de9b47b5cf7915cf2569a62369bba35dec3ddb5f43f85e24',
}


def prepare(here, work, source, build):
    expected = json.loads((here / 'entropy-oracle.json').read_text())['sourceSha256']
    for name, digest in {**expected, **SOURCES}.items():
        if sha256((source / name).read_bytes()) != digest:
            raise ValueError('Unexpected native source: ' + name)
    if '#define CONFIG_AV1_HIGHBITDEPTH 1' not in (build / 'config/aom_config.h').read_text():
        raise ValueError('Native build lacks high-bit-depth support')
    work.mkdir(parents=True, exist_ok=True)
    for name in ('LICENSE', 'PATENTS'):
        shutil.copyfile(source / name, work / name)
    text = (source / 'av1/common/entropymode.c').read_text()
    # Shape declarations for the existing minimal component encoder includes.
    groups = {
        'partition': [('default_partition_cdf', '[20][11]')],
        'prelude': [('default_skip_txfm_cdfs', '[3][3]'), ('default_spatial_pred_seg_tree_cdf', '[3][9]'),
                    ('default_delta_q_cdf', '[5]'), ('default_delta_lf_cdf', '[5]'), ('default_delta_lf_multi_cdf', '[4][5]')],
        'mode': [('default_kf_y_mode_cdf', '[5][5][14]'), ('default_uv_mode_cdf', '[2][13][15]'),
                 ('default_angle_delta_cdf', '[8][8]'), ('default_cfl_sign_cdf', '[9]'),
                 ('default_cfl_alpha_cdf', '[6][17]'), ('default_intrabc_cdf', '[3]')],
    }
    for group, tables in groups.items():
        declarations = []
        for name, shape in tables:
            match = re.search(r'static const aom_cdf_prob\s+' + name + r'.*?=\s*(\{.*?\});', text, re.S)
            if not match:
                raise ValueError('Missing pinned native table: ' + name)
            declarations.append('static const aom_cdf_prob ' + name + shape + ' = ' + match[1] + ';')
        (work / (group + '-defaults.h')).write_text(text[:text.index('*/') + 2] + '\n' + '\n'.join(declarations) + '\n')
    (work / 'entropymode.c').write_text(text)


def verify(here, work, source, build, compiler, vectors, depth):
    probe = work / 'decodemv-probe.c'
    probe.write_bytes((source / 'av1/decoder/decodemv.c').read_bytes() + b'\n#include "OfficePaletteProbe.inc"\n')
    obj = work / 'decodemv-probe.o'
    subprocess.run([compiler, '-std=c99', '-O2', '-DNDEBUG', '-I', str(source), '-I', str(build),
                    '-I', str(here), '-c', str(probe), '-o', str(obj)], check=True)
    executable = work / 'read-palette-reference'
    subprocess.run([compiler, '-std=c11', '-Wall', '-Wextra', '-Werror', '-O2', '-I', str(source),
                    '-I', str(build), str(here / 'ReadPaletteReference.c'), str(obj), str(build / 'libaom.a'),
                    '-lm', '-o', str(executable)], check=True)
    binary = work / 'native-palette-cases.bin'
    with binary.open('wb') as out:
        out.write(struct.pack('<i', len(vectors['cases'])))
        for case in vectors['cases']:
            s = case['scenario']
            width, height = case['states'][0]['block'][2:]
            extent = 18 if s % 20 == 18 else 32 if width == 128 or width == 64 or height == 64 or s % 20 == 19 else 16
            encoded = bytes.fromhex(case['hex'])
            out.write(struct.pack('<8i', depth, extent, case['updates'], s != 21, s != 22, s == 20,
                                  len(encoded), len(case['states'])))
            out.write(encoded)
            for state in case['states']:
                out.write(struct.pack('<13i', *state['block'], *state['modes']))
    observed = subprocess.run([str(executable), str(binary)], check=True, capture_output=True, text=True).stdout
    (work / 'actual-native-palettes.jsonl').write_text(observed)
    records = [json.loads(line) for line in observed.splitlines()]
    if len(records) != len(vectors['cases']):
        raise ValueError('Native palette case count differs')
    for case_index, (case, states) in enumerate(zip(vectors['cases'], records)):
        if len(states) != len(case['states']):
            raise ValueError('Native palette leaf count differs')
        for state_index, (expected, actual) in enumerate(zip(case['states'], states)):
            for key in ('filter', 'y', 'u', 'v', 'mapY', 'mapUv'):
                value = expected[key]
                if key in ('y', 'u', 'v') and isinstance(value, str):
                    value = list(bytes.fromhex(value))
                if value != actual[key]:
                    raise ValueError(f'Actual decoder differs at case {case_index}, leaf {state_index}, {key}')
    return {
        'actualDecoder': True, 'bitDepth': depth, 'sourceSha256': SOURCES,
        'configurationSha256': sha256((build / 'config/aom_config.h').read_bytes()),
        'librarySha256': sha256((build / 'libaom.a').read_bytes()),
        'observationSha256': sha256(observed.encode()), 'cases': len(records),
        'leaves': sum(len(states) for states in records),
        'assetsSha256': {name: sha256((here / name).read_bytes()) for name in
                        ('OfficePaletteProbe.inc', 'ReadPaletteReference.c', 'GeneratePaletteNativeReference.py')},
    }
