#!/usr/bin/env python3
"""Opt-in actual native frame CDEF pixels and native direction/filter kernels."""
import argparse,base64,gzip,json,pathlib,shutil,subprocess,sys
sys.dont_write_bytecode=True
from GenerateTileFixtures import COMMIT,INPUTS,run,digest
from GenerateReconstructionFixtures import SOURCE

def main():
    ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('--work-dir',required=True,type=pathlib.Path);ap.add_argument('--output',type=pathlib.Path);args=ap.parse_args()
    here=pathlib.Path(__file__).resolve().parent;repo=here.parents[2];work=args.work_dir.resolve();work.mkdir(parents=True,exist_ok=True)
    source=work/'native-aom';build=work/'native-build';patch=here/'TraceCdef.patch'
    if not source.exists():run(['git','clone','--depth','1','--branch','v3.13.1','https://aomedia.googlesource.com/aom',source])
    assert subprocess.check_output(['git','-C',source,'rev-parse','HEAD'],text=True).strip()==COMMIT
    assert digest(subprocess.check_output(['git','-C',source,'show','HEAD:av1/decoder/decodeframe.c']))==SOURCE
    diff=subprocess.check_output(['git','-C',source,'diff','HEAD','--no-ext-diff','--unified=0'])
    if diff and diff!=patch.read_bytes():raise ValueError('Unexpected native edits')
    if not diff:run(['git','-C',source,'apply','--unidiff-zero',patch])
    assert subprocess.check_output(['git','-C',source,'diff','HEAD','--no-ext-diff','--unified=0'])==patch.read_bytes()
    shutil.copyfile(here/'OfficeReconstructionProbe.inc',source/'av1/decoder/OfficeReconstructionProbe.inc')
    for name in ['LICENSE','PATENTS']:shutil.copyfile(source/name,work/name)
    with (work/'native-configure.log').open('w') as log:run(['cmake','-S',source,'-B',build,'-DCMAKE_BUILD_TYPE=Release','-DAOM_TARGET_CPU=generic','-DENABLE_DOCS=0','-DENABLE_TESTS=0','-DENABLE_EXAMPLES=0','-DENABLE_TOOLS=0','-DCONFIG_AV1_ENCODER=1','-DCONFIG_MULTITHREAD=0','-DCONFIG_RUNTIME_CPU_DETECT=0'],stdout=log,stderr=subprocess.STDOUT)
    with (work/'native-build.log').open('w') as log:run(['cmake','--build',build,'-j','4'],stdout=log,stderr=subprocess.STDOUT)
    common=[shutil.which('clang'),'-std=c11','-Wall','-Wextra','-Werror','-O2','-I',source,'-I',build]
    driver=work/'read-cdef-frame';run(common+[here/'ReadTileTrace.c',build/'libaom.a','-lm','-o',driver])
    kernels=work/'read-cdef-kernels';run(common+[here/'ReadCdefKernels.c',build/'libaom.a','-lm','-o',kernels])
    encoder=work/'encode-filter-controls';run(common+[here/'EncodeFilterControls.c',build/'libaom.a','-lm','-o',encoder])
    cases=[]
    def decode(name,alpha,encoded,input_hash,raw=False,mono=False):
        stem=name+('-alpha' if alpha else '-color');obu=work/(stem+'.obu');pixels=work/(stem+'.yuv');obu.write_bytes(encoded)
        result=run([driver,obu,pixels],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True);stages=[json.loads(v) for v in result.stderr.splitlines()];assert len(stages)==3
        for frame in stages:
            planes=[bytes.fromhex(v) for v in frame['planes']];frame['planes']=[base64.b64encode(v).decode() for v in planes];frame['planeSha256']=[digest(v) for v in planes]
        c={'name':name,'alpha':alpha,'monochrome':mono,'inputSha256':input_hash,'unfiltered':stages[0],'deblocked':stages[1],'cdef':stages[2],'finalImage':json.loads(result.stdout)}
        if raw:c['inputBase64']=base64.b64encode(encoded).decode()
        cases.append(c);return c
    for name,alpha,path,expected,offset,length in INPUTS:
        encoded=(repo/path).read_bytes();assert digest(encoded)==expected;decode(name,alpha,encoded[offset:offset+length],expected,mono=alpha)
    controls=json.loads(gzip.decompress((here/'reconstruction-reference.json.gz').read_bytes()))
    c=next(v for v in controls['cases'] if v['name']=='lossless-copy');encoded=base64.b64decode(c['inputBase64']);assert digest(encoded)==c['inputSha256'];decode('lossless-copy',False,encoded,c['inputSha256'],True)
    for name,mode in [('odd-color',0),('odd-mono',1),('odd-tiled-color',2)]:
        obu=work/(name+'.obu');run([encoder,obu,str(mode)]);encoded=obu.read_bytes();c=decode(name,False,encoded,digest(encoded),True,mode==1)
        assert c['deblocked']['planes']!=c['unfiltered']['planes'] and c['cdef']['planes']!=c['deblocked']['planes'],'Control must demonstrate active deblocking and CDEF'
        if mode==2:
            f=c['unfiltered'];assert f['tileCols']*f['tileRows']>1,'Control must contain multiple tiles'
    encoded=(here/'cdef-svt-mixed-skip.obu').read_bytes()
    assert digest(encoded)=='a7df0ead410eeec2758c51d0764f73463aacbccced234f18dcd97af5ab74c46f'
    c=decode('svt-tiled-mixed-skip',False,encoded,digest(encoded),True)
    f=c['unfiltered'];assert f['tileCols']*f['tileRows']>1 and 0<f['skipLeaves']<f['leaves']
    assert c['cdef']['planes']!=c['deblocked']['planes'],'Mixed-skip control must demonstrate active CDEF'
    numeric=[json.loads(v) for v in run([kernels],stdout=subprocess.PIPE,text=True).stdout.splitlines()]
    fixture={'nativeCommit':COMMIT,'reference':'Actual AOM single-threaded pre/post-CDEF planes and unmodified direction/filter kernels','cases':cases,'directions':[v for v in numeric if v['kind']=='direction'],'kernels':[v for v in numeric if v['kind']=='kernel']}
    canonical=(json.dumps(fixture,separators=(',',':'))+'\n').encode();output=args.output.resolve() if args.output else work/'cdef-reference.json.gz';output.write_bytes(gzip.compress(canonical,mtime=0))
    assets=['GenerateCdefFixtures.py','GenerateSvtCdefControl.py','cdef-svt-mixed-skip.obu','GenerateTileFixtures.py','GenerateReconstructionFixtures.py','TraceCdef.patch','OfficeReconstructionProbe.inc','ReadTileTrace.c','EncodeFilterControls.c','ReadCdefKernels.c','reconstruction-reference.json.gz']
    receipt={'nativeCommit':COMMIT,'decodeframeSourceSha256':SOURCE,'assetsSha256':{n:digest((here/n).read_bytes()) for n in assets},'fixtureSha256':digest(output.read_bytes()),'canonicalJsonSha256':digest(canonical),'cases':len(cases),'directions':len(fixture['directions']),'kernels':len(fixture['kernels']),'changedKernelCases':sum(v['input']!=v['output'] for v in fixture['kernels']),'changedCdefSamples':sum(sum(a!=b for a,b in zip(base64.b64decode(u),base64.b64decode(d))) for c in cases for u,d in zip(c['deblocked']['planes'],c['cdef']['planes']))}
    (work/'cdef-oracle-receipt.json').write_text(json.dumps(receipt,indent=2)+'\n');print(json.dumps(receipt,indent=2))

if __name__=='__main__':main()
