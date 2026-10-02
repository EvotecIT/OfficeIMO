#!/usr/bin/env python3
"""Opt-in native AOM inverse-transform reference, including independently normalized quantizer facts."""
import argparse
import gzip
import hashlib
import json
import pathlib
import shutil
import struct
import subprocess
import sys
sys.dont_write_bytecode=True
from GenerateResidualTables import tables

COMMIT='d772e334cc724105040382a977ebb10dfd393293'
SOURCES={
    'av1/common/av1_inv_txfm2d.c':'4e27a497fe384185826a6de08d1b63a6754c7dd2106844a77fdc1e0640fe99c5',
    'av1/common/av1_inv_txfm1d.c':'d16765d227aebed749398bd649e17bbdf7df4f818a2761eabc92e0e74ca13ae1',
    'av1/common/quant_common.c':'dd419b0adfd6178da6b9df1af1b1592feb65ac22811bc7954624fb6caa624a66',
    'av1/decoder/decodetxb.c':'cd87de06a3f215926b58453c3efda90b43c0f9ac3242ad87309b1703863c8243',
}
WIDTHS=[4,8,16,32,64,4,8,8,16,16,32,32,64,4,16,8,32,16,64]
HEIGHTS=[4,8,16,32,64,8,4,16,8,32,16,64,32,16,4,32,8,64,16]
ROWS=[0,0,1,1,0,1,1,1,1,2,2,0,2,1,2,1]
COLS=[0,1,0,1,1,0,1,1,1,2,0,2,1,2,1,2]
def digest(data): return hashlib.sha256(data).hexdigest()
def run(args,**kwargs): return subprocess.run([str(x) for x in args],check=True,**kwargs)
def cases(bit_depth=8):
    result=[]
    for size,(w,h) in enumerate(zip(WIDTHS,HEIGHTS)):
        for kind in range(16):
            if ((w==64 or h==64) and kind!=0) or (ROWS[kind]==1 and w>16) or (COLS[kind]==1 and h>16): continue
            for pattern in range(4):
                seed=len(result)+1;tw=min(32,w);th=min(32,h);coeff=[0]*(tw*th)
                q=seed*37%256 if pattern in (0,3) else 32+seed%33
                current=seed*13%256 if pattern in (0,3) else 32+seed%33
                delta=(seed//3)%2;seg=(seed//5)%8;segmentation=(seed//7)%3!=0;has_alt=(seed//11)%2!=0
                alt=(seed%511)-255 if pattern in (0,3) else seed%65-32
                d=[(seed*(i+3))%128-64 if pattern in (0,3) else seed%(i+9)-4 for i in range(5)]
                if pattern==0: coeff[0]=(-1 if seed%2 else 1)*(1+seed%63)
                elif pattern==1:
                    for i in (0,tw-1,(th-1)*tw,len(coeff)-1): coeff[i]=(-1 if (seed+i)%2 else 1)*(1+(seed+i)%4)
                elif pattern==2:
                    coeff=[((i*17+seed*19)%3)-1 for i in range(len(coeff))]
                else: coeff[(seed*11)%len(coeff)]=(-1 if seed%2 else 1)*0xfffff
                eob=len(coeff);lossless=max(0,min(255,q+(alt if segmentation and has_alt else 0)))==0 and not any(d)
                params=[size,kind,seed%3,q,current,delta,seg,int(segmentation),int(has_alt),alt,(seed//13)%2,(seed*3)%16,(seed*7)%16,(seed*11)%16,*d,0,int(lossless),eob,len(coeff)]
                assert len(params)==23
                result.append({'parameters':params,'coefficients':coeff})
    for pattern in range(8):
        coeff=[((i*7+pattern*3)%7-3)*(64 if bit_depth==10 else 1) for i in range(16)];params=[0,0,pattern%3,0,0,0,pattern%8,0,0,0,1,pattern%15,pattern%15,pattern%15,0,0,0,0,0,0,1,16,16]
        result.append({'parameters':params,'coefficients':coeff})
    # All-zero and segment/delta clamp cases use independent native reconstruction too.
    result.append({'parameters':[0,0,0,120,48,1,7,1,1,-255,1,0,0,0,0,0,0,0,0,1,1,0,16],'coefficients':[0]*16})
    # Dense ADST4 input can satisfy b7 while violating the required s/x stage precision.
    result.append({'parameters':[0,2,0,255,255,0,0,0,0,0,0,15,15,15,0,0,0,0,0,0,0,4,16],'coefficients':[64]*4+[0]*12})
    return result

def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--work-dir',required=True,type=pathlib.Path);p.add_argument('--spec-text',required=True,type=pathlib.Path);p.add_argument('--output',type=pathlib.Path);p.add_argument('--core-tables-dir',type=pathlib.Path);p.add_argument('--cc',default='clang');p.add_argument('--bit-depth',type=int,choices=(8,10),default=8);p.add_argument('--native-source',type=pathlib.Path);p.add_argument('--native-build',type=pathlib.Path);a=p.parse_args()
    here=pathlib.Path(__file__).resolve().parent;work=a.work_dir.resolve();work.mkdir(parents=True,exist_ok=True);source=work/'native-aom';build=work/'native-build';patch=here/'TraceInverseTransform.patch'
    if bool(a.native_source)!=bool(a.native_build): raise ValueError('Supply both native source and build')
    if a.native_source:
        # Observe three real native translation units without changing the retained producer source/build.
        source=a.native_source.resolve();build=a.native_build.resolve()
        for name,expected in SOURCES.items():
            if digest((source/name).read_bytes())!=expected: raise ValueError('Unexpected native source: '+name)
        if '#define CONFIG_AV1_HIGHBITDEPTH 1' not in (build/'config/aom_config.h').read_text():
            raise ValueError('Native build lacks high-bit-depth support')
        probes=work/'native-probes'
        if probes.exists(): raise ValueError('Use a new work directory for the observation source')
        for name in ('av1/common/av1_inv_txfm2d.c','av1/common/av1_inv_txfm1d.c','av1/decoder/decodetxb.c'):
            target=probes/name;target.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(source/name,target)
        run(['patch','--batch','--forward','-p1','-i',patch],cwd=probes)
    else:
        if not source.exists(): run(['git','clone','--depth','1','--branch','v3.13.1','https://aomedia.googlesource.com/aom',source])
        if subprocess.check_output(['git','-C',str(source),'rev-parse','HEAD'],text=True).strip()!=COMMIT: raise ValueError('Unexpected native revision')
        for name,expected in SOURCES.items():
            if digest(subprocess.check_output(['git','-C',str(source),'show','HEAD:'+name]))!=expected: raise ValueError('Unexpected native source: '+name)
        diff=subprocess.check_output(['git','-C',str(source),'diff','HEAD','--no-ext-diff','--unified=0'])
        if diff and diff!=patch.read_bytes(): raise ValueError('Unexpected native edits')
        if not diff: run(['git','-C',source,'apply','--unidiff-zero',patch])
        with (work/'native-configure.log').open('w') as log:
            run(['cmake','-S',source,'-B',build,'-DCMAKE_BUILD_TYPE=Release','-DENABLE_DOCS=0','-DENABLE_TESTS=0','-DENABLE_EXAMPLES=0','-DENABLE_TOOLS=0','-DCONFIG_AV1_ENCODER=0','-DCONFIG_MULTITHREAD=0','-DCONFIG_RUNTIME_CPU_DETECT=0','-DCONFIG_AV1_HIGHBITDEPTH=1','-DCMAKE_C_FLAGS=-I'+str(here)],stdout=log,stderr=subprocess.STDOUT)
        with (work/'native-build.log').open('w') as log: run(['cmake','--build',build,'-j','4'],stdout=log,stderr=subprocess.STDOUT)
    facts=tables(a.spec_text,work/'numeric-tables')
    if a.core_tables_dir: tables(a.spec_text,a.core_tables_dir)
    executable=work/'read-residual-reference';driver=here/'ReadResidualReference.c';objects=[]
    if a.native_source:
        for name in ('av1/common/av1_inv_txfm2d.c','av1/common/av1_inv_txfm1d.c','av1/decoder/decodetxb.c'):
            obj=work/(pathlib.Path(name).stem+'.o');objects.append(obj)
            run([shutil.which(a.cc),'-std=c99','-O2','-DNDEBUG','-I',source,'-I',build,'-I',here,'-c',probes/name,'-o',obj])
    run([shutil.which(a.cc),'-std=c11','-Wall','-Wextra','-Werror','-O2','-I',source,'-I',build,driver,*objects,build/'libaom.a','-lm','-o',executable])
    data=cases(a.bit_depth);input_path=work/'cases.bin';flat=[len(data)]+[x for c in data for x in c['parameters']+c['coefficients']];input_path.write_bytes(struct.pack('<'+'i'*len(flat),*flat))
    matrix=work/'native-matrices.bin';native=run([executable,input_path,matrix,a.bit_depth],stdout=subprocess.PIPE,text=True).stdout
    (work/'native-results.jsonl').write_text(native);records=[json.loads(line) for line in native.splitlines()];numeric=records.pop(0)
    if digest(matrix.read_bytes())!=facts['matrixSha256']: raise ValueError('Specification/native quantizer matrices differ')
    assert len(records)==len(data)
    for i,(case,result) in enumerate(zip(data,records)):
        assert result['scenario']==i;case['conforming']=result['conforming'];case['residual']=result['residual'];case['pixels']=result['pixels']
    fixture={'bitDepth':a.bit_depth,'reference':'AOM v3.13.1 actual C inverse transforms and native quantization facts','nativeCommit':COMMIT,'numericFacts':numeric,'matrixSha256':facts['matrixSha256'],'cases':data}
    canonical=(json.dumps(fixture,separators=(',',':'))+'\n').encode();output=a.output.resolve() if a.output else work/'residual-reference.json.gz';output.write_bytes(gzip.compress(canonical,mtime=0))
    assets=['ReadResidualReference.c','OfficeInverseProbe.h','OfficeRangeProbe.h','OfficeInverseProbe.inc','OfficeDequantProbe.inc','TraceInverseTransform.patch','GenerateResidualFixtures.py','GenerateResidualTables.py']
    receipt={'nativeCommit':COMMIT,'nativeBuild':{'configurationSha256':digest((build/'config/aom_config.h').read_bytes()),'librarySha256':digest((build/'libaom.a').read_bytes()),'reused':bool(a.native_source)},'sourceSha256':SOURCES,'assetsSha256':{n:digest((here/n).read_bytes()) for n in assets},'numericFacts':facts,'fixtureSha256':digest(output.read_bytes()),'canonicalJsonSha256':digest(canonical),'bitDepth':a.bit_depth,'nonconformingCases':sum(not c['conforming'] for c in data),'cases':len(data),'residualSamples':sum(len(c['residual']) for c in data),'transformSizes':sorted(set(c['parameters'][0] for c in data)),'transformTypes':sorted(set(c['parameters'][1] for c in data)),'losslessCases':sum(c['parameters'][20] for c in data)}
    (work/'residual-oracle-receipt.json').write_text(json.dumps(receipt,indent=2)+'\n');print(json.dumps(receipt,indent=2))
if __name__=='__main__': main()
