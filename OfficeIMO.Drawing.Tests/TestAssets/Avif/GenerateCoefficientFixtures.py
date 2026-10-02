#!/usr/bin/env python3
"""Opt-in bounded coefficient oracle. Native AOM entropy/defaults/scans remain test tooling.
Use --work-dir for scratch; --core-tables-dir deliberately regenerates normative CDF facts.
"""
import argparse
import ast
import base64
import json
import gzip
import pathlib
import re
import shutil
import subprocess
import sys
import urllib.request
sys.dont_write_bytecode = True
from GenerateEntropyFixtures import sha256, write_json

SOURCES = {
    'token_cdfs.h': '027f8b28511c0be811d2fd0eabccc876cab4c8ead5820093cecad5f56b92fc9d',
    'scan.c': '1296162f04e01e208cfbf499d2279b88bb8655beb03b04584a17f0697f1e0e20',
}
FAMILIES = [('Dc', 'av1_default_dc_sign_cdfs', 6), ('Skip', 'av1_default_txb_skip_cdfs', 65),
            ('Extra', 'av1_default_eob_extra_cdfs', 90), ('Br', 'av1_default_coeff_lps_multi_cdfs', 210),
            ('Base', 'av1_default_coeff_base_multi_cdfs', 420), ('Last', 'av1_default_coeff_base_eob_multi_cdfs', 40)]
FAMILIES += [('Eob' + str(i), 'av1_default_eob_multi' + str(16 << i) + '_cdfs', 4) for i in range(7)]


def number(text):
    node = ast.parse(text.strip(), mode='eval').body
    def integer(n):
        if isinstance(n, ast.Constant) and type(n.value) is int:
            return n.value
        if isinstance(n, ast.BinOp) and isinstance(n.op, ast.Mult):
            return integer(n.left) * integer(n.right)
        raise ValueError('Unsupported probability expression')
    return integer(node)


def table(text, name):
    match = re.search(r'static const aom_cdf_prob\s+' + name + r'.*?=\s*(\{.*?\});', text, re.S)
    if not match:
        raise ValueError('Missing table ' + name)
    leaves = re.findall(r'\{\s*(?:AOM_CDF(\d+)\(([^)]*)\)|(0))\s*\}', match[1])
    result = []
    for size, values, zero in leaves:
        if zero:
            result.append([])
        else:
            row = [number(x) for x in values.split(',')]
            if len(row) != int(size)-1 or row != sorted(row) or not 0 < row[0] <= row[-1] < 32768:
                raise ValueError('Invalid probability row ' + name)
            result.append(row)
    return result


def write_tables(work, core):
    text = (work / 'token_cdfs.h').read_text()
    families = {short: table(text, name) for short, name, count in FAMILIES}
    for short, name, count in FAMILIES:
        assert len(families[short]) == count * 4, (name, len(families[short]))
    text = (work / 'entropymode.c').read_text()
    families['Intra'] = table(text, 'default_intra_ext_tx_cdf')
    families['Inter'] = table(text, 'default_inter_ext_tx_cdf')
    assert len(families['Intra']) == 3*4*13 and len(families['Inter']) == 4*4
    native = []
    for short, rows in families.items():
        padded = 17 if short in ('Intra','Inter') else max(len(row)+2 for row in rows)
        native.append('static const aom_cdf_prob coeff_' + short + '[' + str(len(rows)) + '][' + str(padded) + '] = {\n' +
                      ',\n'.join('{ AOM_CDF' + str(len(r)+1) + '(' + ','.join(map(str,r)) + ') }' if r else '{0}' for r in rows) + '\n};')
    notice = (work / 'token_cdfs.h').read_text().split('*/')[0] + '*/\n'
    (work / 'coefficient-defaults.h').write_text(notice + '\n'.join(native) + '\n')
    if core:
        core.mkdir(parents=True, exist_ok=True)
        for short, rows in families.items():
            parts = [rows[i*len(rows)//4:(i+1)*len(rows)//4] for i in range(4)] if short in ('Base','Br') else [rows]
            for q, part in enumerate(parts):
                suffix = short + (str(q) if len(parts)==4 else '')
                code = '// Generated normative AV1 default probability facts; regenerate with GenerateCoefficientFixtures.py.\n'
                code += '// Numeric data verified against hash-pinned AOM v3.13.1 token_cdfs.h / entropymode.c.\n'
                code += 'namespace OfficeIMO.Drawing;\n\ninternal sealed partial class OfficeAv1CoefficientReader {\n'
                code += '    private static int[][] Create' + suffix + '() => new int[][] {\n'
                code += ''.join('        new[] { '+','.join(map(str,r+[32768,0]))+' },\n' if r else '        System.Array.Empty<int>(),\n' for r in part)
                code += '    };\n}\n'
                (core / ('OfficeAv1CoefficientReader.Cdfs.' + suffix + '.cs')).write_text(code)
    scans = {}
    scan_text = (work / 'scan.c').read_text()
    for kind in ('default','mrow','mcol'):
        for w,h in ((4,4),(8,8),(16,16),(32,32),(4,8),(8,4),(8,16),(16,8),(16,32),(32,16),(4,16),(16,4),(8,32),(32,8)):
            name = kind + '_scan_' + str(w) + 'x' + str(h)
            m = re.search(r'\b' + name + r'\[.*?=\s*\{(.*?)\}', scan_text,re.S)
            assert m, name
            raw = [int(x) for x in re.findall(r'\d+',m[1])]
            assert len(raw) == w*h and sorted(raw) == list(range(w*h))
            scans[name] = [(v%h)*w+v//h for v in raw]
    write_json(work / 'coefficient-scans.json', scans)
    # Normalized native arrays are data, not generated from OfficeIMO's scan algorithm.
    (work / 'coefficient-scans.h').write_text(notice + '\n'.join('static const int ' + name + '[' + str(len(v)) + ']={' + ','.join(map(str,v)) + '};' for name,v in scans.items()) + '\n')
    return families, scans


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--work-dir', required=True, type=pathlib.Path)
    parser.add_argument('--core-tables-dir', type=pathlib.Path)
    parser.add_argument('--tables-only', action='store_true')
    parser.add_argument('--output', type=pathlib.Path)
    parser.add_argument('--cc',default='clang')
    args=parser.parse_args();here=pathlib.Path(__file__).resolve().parent;work=args.work_dir.resolve();work.mkdir(parents=True,exist_ok=True)
    subprocess.run([sys.executable,str(here/'GenerateTransformFixtures.py'),'--work-dir',str(work),'--cc',args.cc],check=True)
    for name,digest in SOURCES.items():
        path=work/name
        if not path.is_file():
            url='https://aomedia.googlesource.com/aom/+/refs/tags/v3.13.1/av1/common/'+name+'?format=TEXT'
            with urllib.request.urlopen(url,timeout=60) as response: data=base64.b64decode(response.read(),validate=True)
            if sha256(data)!=digest: raise ValueError('Source hash mismatch')
            path.write_bytes(data)
        if sha256(path.read_bytes())!=digest: raise ValueError('Cached source hash mismatch')
    families,scans=write_tables(work,args.core_tables_dir)
    if args.tables_only: return
    compiler=shutil.which(args.cc);native=work/'native-aom';harness=here/'GenerateCoefficientFixtures.c';executable=work/'generate-coefficients'
    command=[compiler,'-std=c11','-Wall','-Wextra','-Werror','-O2','-I',str(native),'-I',str(work),str(harness)]
    command += [str(native/'aom_dsp'/name) for name in ('entenc.c','entdec.c','entcode.c')]
    subprocess.run(command+['-o',str(executable)],check=True)
    raw=work/'native-coefficients.json';subprocess.run([str(executable),str(raw)],check=True);vectors=json.loads(raw.read_text());vectors['scans']=scans
    prefixes=[]
    entropy=json.loads((here/'entropy-oracle.json').read_text());geometry=json.loads((here/'partition-oracle.json').read_text())['prefixGeometry']
    preludes=json.loads((work/'prelude-reference.json').read_text())['framePrefixes']
    for frame in entropy['framePrefixes']:
        g=geometry[frame['name']];b=next(v for v in preludes if v['name']==frame['name']);output=work/(frame['name']+'-coefficients.json')
        mode=2 if frame['name']=='multitile' else 1
        subprocess.run([str(executable),str(here.parents[2]/frame['path']),str(frame['offset']),str(frame['length']),str(g['miRows']),str(g['miCols']),str(b['baseQ']),str(int(b['cdef'])),str(int(b['deltaQ'])),str(output),str(int(frame['alpha'])),str(int(frame['name']=='avif-opaque')),str(mode)],check=True)
        p={key:frame[key] for key in ('name','alpha','offset','length')};p.update(json.loads(output.read_text()));prefixes.append(p)
    vectors['framePrefixes']=prefixes
    output=args.output.resolve() if args.output else work/'coefficient-reference.json.gz'
    encoded=json.dumps(vectors,indent=2)
    encoded=re.sub(r'\{\s*"block":\s*\[[^]]*\],\s*"type":.*?"values":\s*\[.*?\]\s*\}',lambda m: json.dumps(json.loads(m[0]),separators=(',',':')),encoded,flags=re.S)
    encoded=re.sub(r'\[\s*(-?\d+(?:,\s*-?\d+)+)\s*\]',lambda m: '['+re.sub(r'\s+','',m[1])+']',encoded)
    canonical=(encoded+'\n').encode('utf-8')
    output.write_bytes(gzip.compress(canonical,mtime=0) if output.suffix=='.gz' else canonical)
    write_json(work/'coefficient-oracle-receipt.json',{'generatorSha256':sha256(pathlib.Path(__file__).read_bytes()),'harnessSha256':sha256(harness.read_bytes()),'sourceSha256':SOURCES,'fixtureSha256':sha256(output.read_bytes()),'canonicalJsonSha256':sha256(canonical),'cases':len(vectors['cases']),'prefixes':len(prefixes),'nativeSelfCheck':vectors['nativeSelfCheck']})
    print('Generated '+str(output)+' SHA256 '+sha256(output.read_bytes()))

if __name__=='__main__': main()
