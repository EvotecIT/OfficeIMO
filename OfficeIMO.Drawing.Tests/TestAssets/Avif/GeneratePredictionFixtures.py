#!/usr/bin/env python3
"""Opt-in actual native AOM intra, CfL and palette prediction pixels."""
import argparse
import gzip
import json
import pathlib
import shutil
import struct
import subprocess
import sys
sys.dont_write_bytecode=True
from GenerateResidualFixtures import COMMIT,run,digest,WIDTHS,HEIGHTS
from GeneratePredictionTables import tables

SOURCES={'av1/common/reconintra.c':'710299b97e4bede074bd098ad2c0154d21c8734f0a0cb344de15d5d03dd62fe9',
         'av1/common/cfl.c':'2e5abfe4d008a2dda0b1c7f7bf6664bf7ca497b53fe653333941910cd37f792a',
         'aom_dsp/intrapred.c':'a10c6bc818c606962543c61aa6b293a485a22524cdac367cd4dc4b45ca0715f3'}
def samples(n,seed,multiplier,bit_depth):
    maximum=(1<<bit_depth)-1;scale=1<<(bit_depth-8)
    return [33*scale if seed%11==0 else (maximum if (i+seed)%2 else 0) if seed%7==0 else (i*multiplier+seed*37)%(maximum+1) for i in range(n)]
def cases(bit_depth):
    result=[];maximum=(1<<bit_depth)-1;middle=1<<(bit_depth-1)
    for size,(w,h) in enumerate(zip(WIDTHS,HEIGHTS)):
        modes=[(m,a,-1) for m in range(13) for a in (range(-3,4) if 1<=m<=8 else (0,))]
        if w<=32 and h<=32: modes += [(0,0,f) for f in range(5)]
        for mode,angle,filt in modes:
            for state in range(8):
                seed=len(result)+1
                top,left,tr,bl=w,h,h,w
                if state==1: tr=bl=0
                elif state==3: top=tr=0
                elif state==4: left=bl=0
                elif state==5: top=left=tr=bl=0
                elif state==6: top=w//2;left=h//2;tr=bl=0
                elif state==7: tr=h//2;bl=w//2
                above=samples(top+tr,seed,59,bit_depth);lefts=samples(left+bl,seed+1,71,bit_depth)
                p=[size,mode,angle,int(state!=0),int(state in (2,7)),filt,top,left,tr,bl,seed*11%(maximum+1),len(above),len(lefts),0,0,0]
                result.append({'kind':0,'parameters':p,'above':above,'left':lefts})
    for size,(w,h) in enumerate(zip(WIDTHS,HEIGHTS)):
        if w>32 or h>32: continue
        for subx,suby in ((0,0),(1,0),(1,1)):
            for alpha in range(-16,17):
                for clipped in (0,1):
                    seed=len(result)+1;sw=(w-clipped)<<subx;sh=(h-clipped)<<suby;stride=sw+int(sw<64)
                    luma=samples(stride*sh,seed,53,bit_depth);p=[size,alpha,(0,middle,maximum)[seed%3],subx,suby,sw,sh,stride,len(luma),0,0,0,0,0,0,0]
                    result.append({'kind':1,'parameters':p,'luma':luma})
    for size,(w,h) in enumerate(zip(WIDTHS,HEIGHTS)):
        for plane in range(3):
            limit=64 if plane==0 else 32
            if w>limit or h>limit: continue
            mw=min(limit,w*2);mh=min(limit,h*2)
            for colors in range(2,9):
                for offset in (0,1):
                    seed=len(result)+1;palette=samples(colors,seed,31,bit_depth);index=[(i*7+i//mw+seed)%colors for i in range(mw*mh)]
                    p=[size,plane,mw,mh,offset*(mw-w),offset*(mh-h),colors,0,0,0,0,0,0,0,0,0]
                    result.append({'kind':2,'parameters':p,'colors':palette,'map':index})
    return result
def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--work-dir',required=True,type=pathlib.Path);p.add_argument('--spec-text',required=True,type=pathlib.Path);p.add_argument('--output',type=pathlib.Path);p.add_argument('--core-tables-dir',type=pathlib.Path);p.add_argument('--cc',default='clang');p.add_argument('--bit-depth',type=int,choices=(8,10),default=8);p.add_argument('--native-source',type=pathlib.Path);p.add_argument('--native-build',type=pathlib.Path);a=p.parse_args()
    here=pathlib.Path(__file__).resolve().parent;work=a.work_dir.resolve();work.mkdir(parents=True,exist_ok=True);source=work/'native-aom';build=work/'native-build';patch=here/'TracePrediction.patch'
    if bool(a.native_source)!=bool(a.native_build): raise ValueError('Supply both native source and build')
    if a.native_source:
        source=a.native_source.resolve();build=a.native_build.resolve()
        for name,expected in SOURCES.items():
            if digest((source/name).read_bytes())!=expected: raise ValueError('Unexpected native source: '+name)
        if '#define CONFIG_AV1_HIGHBITDEPTH 1' not in (build/'config/aom_config.h').read_text():
            raise ValueError('Native build lacks high-bit-depth support')
        probes=work/'native-probes'
        if probes.exists(): raise ValueError('Use a new work directory for the observation source')
        for name in ('av1/common/reconintra.c','av1/common/cfl.c'):
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
    facts=tables(a.spec_text,work/'numeric-tables')
    if a.core_tables_dir: tables(a.spec_text,a.core_tables_dir)
    for name in ('LICENSE','PATENTS'): shutil.copyfile(source/name,work/name)
    if not a.native_source:
        with (work/'native-configure.log').open('w') as log: run(['cmake','-S',source,'-B',build,'-DCMAKE_BUILD_TYPE=Release','-DENABLE_DOCS=0','-DENABLE_TESTS=0','-DENABLE_EXAMPLES=0','-DENABLE_TOOLS=0','-DCONFIG_AV1_ENCODER=0','-DCONFIG_MULTITHREAD=0','-DCONFIG_RUNTIME_CPU_DETECT=0','-DCONFIG_AV1_HIGHBITDEPTH=1','-DCMAKE_C_FLAGS=-I'+str(here)],stdout=log,stderr=subprocess.STDOUT)
        with (work/'native-build.log').open('w') as log: run(['cmake','--build',build,'-j','4'],stdout=log,stderr=subprocess.STDOUT)
    executable=work/'read-prediction-reference';objects=[]
    if a.native_source:
        for name in ('av1/common/reconintra.c','av1/common/cfl.c'):
            obj=work/(pathlib.Path(name).stem+'.o');objects.append(obj)
            run([shutil.which(a.cc),'-std=c99','-O2','-DNDEBUG','-I',source,'-I',build,'-I',here,'-c',probes/name,'-o',obj])
    run([shutil.which(a.cc),'-std=c11','-Wall','-Wextra','-Werror','-O2','-I',source,'-I',build,here/'ReadPredictionReference.c',*objects,build/'libaom.a','-lm','-o',executable])
    data=cases(a.bit_depth);input_path=work/'cases.bin'
    with input_path.open('wb') as out:
        out.write(struct.pack('<i',len(data)))
        for case in data:
            out.write(struct.pack('<17i',case['kind'],*case['parameters']))
            for key in ('above','left','luma','colors','map'):
                if key in case: out.write(bytes(case[key]) if key=='map' or a.bit_depth==8 else struct.pack('<'+'H'*len(case[key]),*case[key]))
    native=run([executable,input_path,a.bit_depth],stdout=subprocess.PIPE,text=True).stdout;(work/'native-results.jsonl').write_text(native);records=[json.loads(x) for x in native.splitlines()];assert len(records)==len(data)
    for i,(case,result) in enumerate(zip(data,records)):
        assert result['scenario']==i;case['pixels']=result['pixels']
    fixture={'bitDepth':a.bit_depth,'reference':'AOM v3.13.1 actual native prediction builders and kernels','nativeCommit':COMMIT,'cases':data};canonical=(json.dumps(fixture,separators=(',',':'))+'\n').encode();output=a.output.resolve() if a.output else work/'prediction-reference.json.gz';output.write_bytes(gzip.compress(canonical,mtime=0))
    assets=['OfficePredictionProbe.inc','OfficeCflProbe.inc','ReadPredictionReference.c','TracePrediction.patch','GeneratePredictionFixtures.py','GeneratePredictionTables.py']
    receipt={'bitDepth':a.bit_depth,'nativeCommit':COMMIT,'nativeBuild':{'configurationSha256':digest((build/'config/aom_config.h').read_bytes()),'librarySha256':digest((build/'libaom.a').read_bytes()),'reused':bool(a.native_source)},'sourceSha256':SOURCES,'assetsSha256':{n:digest((here/n).read_bytes()) for n in assets},'specTextSha256':facts['specTextSha256'],'fixtureSha256':digest(output.read_bytes()),'canonicalJsonSha256':digest(canonical),'cases':len(data),'predictionSamples':sum(len(c['pixels']) for c in data),'byKind':{str(k):sum(c['kind']==k for c in data) for k in range(3)}}
    (work/'prediction-oracle-receipt.json').write_text(json.dumps(receipt,indent=2)+'\n');print(json.dumps(receipt,indent=2))
if __name__=='__main__': main()
