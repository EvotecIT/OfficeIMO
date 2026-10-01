#!/usr/bin/env python3
"""Opt-in actual AOM frame pixels before loop filters; no decoder algorithm or runtime dependency changes."""
import argparse,base64,gzip,hashlib,json,pathlib,shutil,subprocess,sys
sys.dont_write_bytecode=True
from GenerateTileFixtures import COMMIT,INPUTS,run,digest

SOURCE='653fe4bc6556f48064e20058057f231902bd9740c66a8930c33a059f4273ac51'

def main():
    ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('--work-dir',required=True,type=pathlib.Path);ap.add_argument('--output',type=pathlib.Path);args=ap.parse_args()
    here=pathlib.Path(__file__).resolve().parent;repo=here.parents[2];work=args.work_dir.resolve();work.mkdir(parents=True,exist_ok=True)
    source=work/'native-aom';build=work/'native-build';patch=here/'TraceReconstruction.patch'
    if not source.exists():run(['git','clone','--depth','1','--branch','v3.13.1','https://aomedia.googlesource.com/aom',source])
    assert subprocess.check_output(['git','-C',source,'rev-parse','HEAD'],text=True).strip()==COMMIT
    assert digest(subprocess.check_output(['git','-C',source,'show','HEAD:av1/decoder/decodeframe.c']))==SOURCE
    diff=subprocess.check_output(['git','-C',source,'diff','--no-ext-diff','--unified=0'])
    if diff and diff!=patch.read_bytes():raise ValueError('Unexpected native edits')
    if not diff:run(['git','-C',source,'apply','--unidiff-zero',patch])
    assert subprocess.check_output(['git','-C',source,'diff','--no-ext-diff','--unified=0'])==patch.read_bytes()
    shutil.copyfile(here/'OfficeReconstructionProbe.inc',source/'av1/decoder/OfficeReconstructionProbe.inc')
    for name in ['LICENSE','PATENTS']:shutil.copyfile(source/name,work/name)
    with (work/'native-configure.log').open('w') as log:
        run(['cmake','-S',source,'-B',build,'-DCMAKE_BUILD_TYPE=Release','-DAOM_TARGET_CPU=generic','-DENABLE_DOCS=0','-DENABLE_TESTS=0','-DENABLE_EXAMPLES=0','-DENABLE_TOOLS=0','-DCONFIG_AV1_ENCODER=1','-DCONFIG_MULTITHREAD=0','-DCONFIG_RUNTIME_CPU_DETECT=0'],stdout=log,stderr=subprocess.STDOUT)
    with (work/'native-build.log').open('w') as log:run(['cmake','--build',build,'-j','4'],stdout=log,stderr=subprocess.STDOUT)
    executable=work/'read-reconstructed-frame';run([shutil.which('clang'),'-std=c11','-Wall','-Wextra','-Werror','-O2','-I',source,here/'ReadTileTrace.c',build/'libaom.a','-lm','-o',executable])
    encoder=work/'encode-reconstruction-controls';run([shutil.which('clang'),'-std=c11','-Wall','-Wextra','-Werror','-O2','-I',source,here/'EncodeReconstructionControls.c',build/'libaom.a','-lm','-o',encoder])
    cases=[]
    for name,alpha,path,expected,offset,length in INPUTS:
        encoded=(repo/path).read_bytes();assert digest(encoded)==expected
        stem=name+('-alpha' if alpha else '-color');obu=work/(stem+'.obu');obu.write_bytes(encoded[offset:offset+length]);pixels=work/(stem+'.yuv')
        result=run([executable,obu,pixels],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True)
        final=json.loads(result.stdout);frame=json.loads(result.stderr);raw=[bytes.fromhex(v) for v in frame['planes']]
        assert all(len(v)==(frame['miCols']*4>>(p>0))*(frame['miRows']*4>>(p>0)) for p,v in enumerate(raw))
        frame['planes']=[base64.b64encode(v).decode() for v in raw];frame['planeSha256']=[digest(v) for v in raw]
        cases.append({'name':name,'alpha':alpha,'inputSha256':expected,'itemOffset':offset,'itemLength':length,'unfiltered':frame,'finalImage':final,'finalYuvSha256':digest(pixels.read_bytes())})
    for name in ['lossless-copy']:
        obu=work/(name+'.obu');pixels=work/(name+'.yuv');run([encoder,obu])
        result=run([executable,obu,pixels],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True);final=json.loads(result.stdout);frame=json.loads(result.stderr)
        raw=[bytes.fromhex(v) for v in frame['planes']];frame['planes']=[base64.b64encode(v).decode() for v in raw];frame['planeSha256']=[digest(v) for v in raw]
        assert frame['skipLeaves']>0
        assert frame['copyLeaves']>0 and frame['fractionalCopyLeaves']>0
        cases.append({'name':name,'alpha':False,'inputBase64':base64.b64encode(obu.read_bytes()).decode(),'inputSha256':digest(obu.read_bytes()),'unfiltered':frame,'finalImage':final,'finalYuvSha256':digest(pixels.read_bytes())})
    fixture={'nativeCommit':COMMIT,'reference':'Actual AOM v3.13.1 single-threaded decoder before frame filtering','cases':cases}
    canonical=(json.dumps(fixture,separators=(',',':'))+'\n').encode();output=args.output.resolve() if args.output else work/'reconstruction-reference.json.gz';output.write_bytes(gzip.compress(canonical,mtime=0))
    assets=['GenerateReconstructionFixtures.py','GenerateTileFixtures.py','TraceReconstruction.patch','OfficeReconstructionProbe.inc','ReadTileTrace.c','EncodeReconstructionControls.c']
    receipt={'nativeCommit':COMMIT,'decodeframeSourceSha256':SOURCE,'assetsSha256':{n:digest((here/n).read_bytes()) for n in assets},'fixtureSha256':digest(output.read_bytes()),'canonicalJsonSha256':digest(canonical),'cases':len(cases),'samples':sum(len(base64.b64decode(p)) for c in cases for p in c['unfiltered']['planes'])}
    (work/'reconstruction-oracle-receipt.json').write_text(json.dumps(receipt,indent=2)+'\n');print(json.dumps(receipt,indent=2))

if __name__=='__main__':main()
