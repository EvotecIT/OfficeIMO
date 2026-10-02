#!/usr/bin/env python3
"""Extract AV1 prediction numeric facts from the pinned normative text."""
import argparse
import hashlib
import pathlib
import re

SPEC_SHA='879df3a6935502599b0f7dfddf83c8bf6b4af62ba88c5ee61d0dabeede528cf3'
def tables(spec,output):
    raw=spec.read_bytes()
    if hashlib.sha256(raw).hexdigest()!=SPEC_SHA: raise ValueError('Unexpected specification text')
    text='\n'.join(x for x in raw.decode().splitlines() if 'Section:' not in x and 'AV1 Bitstream & Decoding Process Specification' not in x)
    def values(name,count):
        match=re.search(r'\b'+name+r'\s*\[[^=]+?=\s*\{',text)
        if not match: raise ValueError('Missing fact: '+name)
        start=match.end()-1;depth=0
        for end in range(start,len(text)):
            depth+=(text[end]=='{')-(text[end]=='}')
            if depth==0:
                result=[int(x) for x in re.findall(r'-?\b\d+\b',text[start:end+1])]
                if len(result)!=count: raise ValueError('Unexpected fact length: '+name)
                return result
        raise ValueError('Incomplete fact: '+name)
    facts={'ModeAngles':values('Mode_To_Angle',13),'Derivative':values('Dr_Intra_Derivative',90),
           'SmoothWeights':sum((values('Sm_Weights_Tx_'+str(n)+'x'+str(n),n) for n in (4,8,16,32,64)),[]),
           'FilterTaps':values('Intra_Filter_Taps',280)}
    code='// Generated numeric facts: AV1 1.0.0 Errata 1, section 9.5.3.\n// Regenerate with GeneratePredictionTables.py and the pinned specification text.\nnamespace OfficeIMO.Drawing;\n\ninternal sealed partial class OfficeAv1IntraPredictor {\n'
    for name,data in facts.items():
        code+='    private static readonly short[] '+name+'={\n'
        code+=''.join('        '+','.join(map(str,data[i:i+16]))+',\n' for i in range(0,len(data),16));code+='    };\n'
    output.mkdir(parents=True,exist_ok=True);(output/'OfficeAv1IntraPredictor.Tables.Generated.cs').write_text(code+'}\n')
    return {'specTextSha256':SPEC_SHA,'facts':facts}
if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--spec-text',required=True,type=pathlib.Path);p.add_argument('--output-dir',required=True,type=pathlib.Path)
    a=p.parse_args();print(tables(a.spec_text,a.output_dir))
