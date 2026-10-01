#!/usr/bin/env python3
"""Opt-in independently decoded pre/post-deblock pixels and actual native narrow/wide kernels."""
import argparse,base64,gzip,json,pathlib,shutil,subprocess,sys
sys.dont_write_bytecode=True
from GenerateTileFixtures import COMMIT,INPUTS,run,digest
from GenerateReconstructionFixtures import SOURCE

def main():
    ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('--work-dir',required=True,type=pathlib.Path);ap.add_argument('--output',type=pathlib.Path);args=ap.parse_args()
    here=pathlib.Path(__file__).resolve().parent;repo=here.parents[2];work=args.work_dir.resolve();work.mkdir(parents=True,exist_ok=True)
    source=work/'native-aom';build=work/'native-build';patch=here/'TraceDeblocking.patch'
    if not source.exists():run(['git','clone','--depth','1','--branch','v3.13.1','https://aomedia.googlesource.com/aom',source])
    assert subprocess.check_output(['git','-C',source,'rev-parse','HEAD'],text=True).strip()==COMMIT
    assert digest(subprocess.check_output(['git','-C',source,'show','HEAD:av1/decoder/decodeframe.c']))==SOURCE
    diff=subprocess.check_output(['git','-C',source,'diff','--no-ext-diff','--unified=0'])
    if diff and diff!=patch.read_bytes():raise ValueError('Unexpected native edits')
    if not diff:run(['git','-C',source,'apply','--unidiff-zero',patch])
    assert subprocess.check_output(['git','-C',source,'diff','--no-ext-diff','--unified=0'])==patch.read_bytes()
    shutil.copyfile(here/'OfficeReconstructionProbe.inc',source/'av1/decoder/OfficeReconstructionProbe.inc')
    for name in ['LICENSE','PATENTS']:shutil.copyfile(source/name,work/name)
    with (work/'native-configure.log').open('w') as log:run(['cmake','-S',source,'-B',build,'-DCMAKE_BUILD_TYPE=Release','-DAOM_TARGET_CPU=generic','-DENABLE_DOCS=0','-DENABLE_TESTS=0','-DENABLE_EXAMPLES=0','-DENABLE_TOOLS=0','-DCONFIG_AV1_ENCODER=0','-DCONFIG_MULTITHREAD=0','-DCONFIG_RUNTIME_CPU_DETECT=0'],stdout=log,stderr=subprocess.STDOUT)
    with (work/'native-build.log').open('w') as log:run(['cmake','--build',build,'-j','4'],stdout=log,stderr=subprocess.STDOUT)
    common=[shutil.which('clang'),'-std=c11','-Wall','-Wextra','-Werror','-O2','-I',source,'-I',build]
    driver=work/'read-deblocked-frame';run(common+[here/'ReadTileTrace.c',build/'libaom.a','-lm','-o',driver])
    kernels=work/'read-deblock-kernels';run(common+[here/'ReadDeblockKernels.c',build/'libaom.a','-lm','-o',kernels])
    cases=[]
    for name,alpha,path,expected,offset,length in INPUTS:
        encoded=(repo/path).read_bytes();assert digest(encoded)==expected
        stem=name+('-alpha' if alpha else '-color');obu=work/(stem+'.obu');obu.write_bytes(encoded[offset:offset+length]);pixels=work/(stem+'.yuv')
        result=run([driver,obu,pixels],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True);stages=[json.loads(v) for v in result.stderr.splitlines()];assert len(stages)==2
        for frame in stages:
            raw=[bytes.fromhex(v) for v in frame['planes']];frame['planes']=[base64.b64encode(v).decode() for v in raw];frame['planeSha256']=[digest(v) for v in raw]
        cases.append({'name':name,'alpha':alpha,'inputSha256':expected,'unfiltered':stages[0],'deblocked':stages[1],'finalImage':json.loads(result.stdout)})
    controls=json.loads(gzip.decompress((here/'reconstruction-reference.json.gz').read_bytes()))
    c=next(v for v in controls['cases'] if v['name']=='lossless-copy');encoded=base64.b64decode(c['inputBase64']);assert digest(encoded)==c['inputSha256']
    obu=work/'lossless-copy.obu';pixels=work/'lossless-copy.yuv';obu.write_bytes(encoded)
    result=run([driver,obu,pixels],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True);stages=[json.loads(v) for v in result.stderr.splitlines()];assert len(stages)==2
    for frame in stages:
        raw=[bytes.fromhex(v) for v in frame['planes']];frame['planes']=[base64.b64encode(v).decode() for v in raw];frame['planeSha256']=[digest(v) for v in raw]
    assert stages[0]['planes']==stages[1]['planes'] and stages[0]['copyLeaves']>0 and stages[0]['skipLeaves']>0
    cases.append({'name':'lossless-copy','alpha':False,'inputBase64':c['inputBase64'],'inputSha256':c['inputSha256'],'unfiltered':stages[0],'deblocked':stages[1],'finalImage':json.loads(result.stdout)})
    numeric=[json.loads(v) for v in run([kernels],stdout=subprocess.PIPE,text=True).stdout.splitlines()]
    fixture={'nativeCommit':COMMIT,'reference':'Actual AOM single-threaded pre/post-deblock planes and native kernels/threshold initialization','cases':cases,'kernels':numeric}
    canonical=(json.dumps(fixture,separators=(',',':'))+'\n').encode();output=args.output.resolve() if args.output else work/'deblock-reference.json.gz';output.write_bytes(gzip.compress(canonical,mtime=0))
    assets=['GenerateDeblockFixtures.py','GenerateTileFixtures.py','GenerateReconstructionFixtures.py','TraceDeblocking.patch','OfficeReconstructionProbe.inc','ReadTileTrace.c','ReadDeblockKernels.c','reconstruction-reference.json.gz']
    receipt={'nativeCommit':COMMIT,'decodeframeSourceSha256':SOURCE,'assetsSha256':{n:digest((here/n).read_bytes()) for n in assets},'fixtureSha256':digest(output.read_bytes()),'canonicalJsonSha256':digest(canonical),'cases':len(cases),'kernels':len(numeric),'changedKernelCases':sum(v['input']!=v['output'] for v in numeric),'changedFrameSamples':sum(sum(a!=b for a,b in zip(base64.b64decode(u),base64.b64decode(d))) for c in cases for u,d in zip(c['unfiltered']['planes'],c['deblocked']['planes']))}
    (work/'deblock-oracle-receipt.json').write_text(json.dumps(receipt,indent=2)+'\n');print(json.dumps(receipt,indent=2))

if __name__=='__main__':main()
