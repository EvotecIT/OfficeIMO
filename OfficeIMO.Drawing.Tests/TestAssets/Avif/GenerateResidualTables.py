#!/usr/bin/env python3
"""Extract normative numeric AV1 facts from the pinned specification text, never reference implementation code."""
import argparse
import hashlib
import pathlib
import re

SPEC_SHA='879df3a6935502599b0f7dfddf83c8bf6b4af62ba88c5ee61d0dabeede528cf3'

def tables(spec,output):
    raw=spec.read_bytes()
    if hashlib.sha256(raw).hexdigest()!=SPEC_SHA: raise ValueError('Unexpected specification text')
    text='\n'.join(line for line in raw.decode().splitlines() if 'Section:' not in line and 'AV1 Bitstream & Decoding Process Specification' not in line)
    text=re.sub(r'/\*.*?\*/','',text,flags=re.S)
    def block(name):
        match=re.search(r'\b'+name+r'\s*\[[^=]+?=\s*\{',text);start=match.end()-1;depth=0
        for end in range(start,len(text)):
            depth+=(text[end]=='{')-(text[end]=='}')
            if depth==0: return [int(x) for x in re.findall(r'\b\d+\b',text[start:end+1])]
        raise ValueError('Incomplete table: '+name)
    dc=block('Dc_Qlookup');ac=block('Ac_Qlookup');matrices=block('Quantizer_Matrix')
    if len(dc)!=768 or len(ac)!=768 or len(matrices)!=100320 or any(not 0<x<256 for x in matrices):
        raise ValueError('Unexpected quantization fact dimensions')
    output.mkdir(parents=True,exist_ok=True)
    code='// Generated numeric facts: AV1 1.0.0 Errata 1, sections 7.12.2 and 9.5.3.\n// Regenerate with GenerateResidualTables.py and the pinned specification text.\nnamespace OfficeIMO.Drawing;\n\ninternal static partial class OfficeAv1QuantizationTables {\n'
    for name,values in [('Dc',dc[:256]),('Ac',ac[:256]),('Dc10',dc[256:512]),('Ac10',ac[256:512])]:
        code+='    internal static readonly short[] '+name+'={\n'
        code+=''.join('        '+','.join(map(str,values[i:i+16]))+',\n' for i in range(0,len(values),16))
        code+='    };\n'
    (output/'OfficeAv1QuantizationTables.Generated.cs').write_text(code+'}\n')
    (output/'OfficeAv1QuantizerMatrices.bin').write_bytes(bytes(matrices))
    return {'specTextSha256':SPEC_SHA,'matrixSha256':hashlib.sha256(bytes(matrices)).hexdigest(),'matrixBytes':len(matrices)}

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--spec-text',required=True,type=pathlib.Path);p.add_argument('--output-dir',required=True,type=pathlib.Path)
    a=p.parse_args();print(tables(a.spec_text,a.output_dir))
