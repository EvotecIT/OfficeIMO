#!/usr/bin/env python3
"""Opt-in actual AOM pixels before/after restoration, including native-selected filter units."""
import argparse,base64,gzip,json,pathlib,shutil,subprocess,sys
sys.dont_write_bytecode=True
from GenerateTileFixtures import COMMIT,INPUTS,run,digest
from GenerateReconstructionFixtures import SOURCE

def main():
    ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('--work-dir',required=True,type=pathlib.Path);ap.add_argument('--output',type=pathlib.Path);args=ap.parse_args()
    here=pathlib.Path(__file__).resolve().parent;repo=here.parents[2];work=args.work_dir.resolve();work.mkdir(parents=True,exist_ok=True)
    source=work/'native-aom';build=work/'native-build';patch=here/'TraceSuperres.patch'
    if not source.exists():run(['git','clone','--depth','1','--branch','v3.13.1','https://aomedia.googlesource.com/aom',source])
    assert subprocess.check_output(['git','-C',source,'rev-parse','HEAD'],text=True).strip()==COMMIT
    assert digest(subprocess.check_output(['git','-C',source,'show','HEAD:av1/decoder/decodeframe.c']))==SOURCE
    diff=subprocess.check_output(['git','-C',source,'diff','HEAD','--no-ext-diff','--unified=0'])
    if diff and diff!=patch.read_bytes():raise ValueError('Unexpected native edits')
    if not diff:run(['git','-C',source,'apply','--unidiff-zero',patch])
    assert subprocess.check_output(['git','-C',source,'diff','HEAD','--no-ext-diff','--unified=0'])==patch.read_bytes()
    for n in ['OfficeReconstructionProbe.inc','OfficeRestorationProbe.inc']:shutil.copyfile(here/n,source/'av1/decoder'/n)
    for n in ['LICENSE','PATENTS']:shutil.copyfile(source/n,work/n)
    with (work/'native-configure.log').open('w') as log:run(['cmake','-S',source,'-B',build,'-DCMAKE_BUILD_TYPE=Release','-DAOM_TARGET_CPU=generic','-DENABLE_DOCS=0','-DENABLE_TESTS=0','-DENABLE_EXAMPLES=0','-DENABLE_TOOLS=0','-DCONFIG_AV1_ENCODER=1','-DCONFIG_MULTITHREAD=0','-DCONFIG_RUNTIME_CPU_DETECT=0'],stdout=log,stderr=subprocess.STDOUT)
    with (work/'native-build.log').open('w') as log:run(['cmake','--build',build,'-j','4'],stdout=log,stderr=subprocess.STDOUT)
    common=[shutil.which('clang'),'-std=c11','-Wall','-Wextra','-Werror','-O2','-I',source,'-I',build]
    driver=work/'read-restored-frame';run(common+[here/'ReadTileTrace.c',build/'libaom.a','-lm','-o',driver])
    encoder=work/'encode-restoration-controls';run(common+[here/'EncodeRestorationControls.c',build/'libaom.a','-lm','-o',encoder])
    kernel=work/'read-restoration-kernels';run(common+[here/'ReadRestorationKernels.c',build/'libaom.a','-lm','-o',kernel])
    upscale=work/'read-upscale-kernels';run(common+[here/'ReadUpscaleKernels.c',build/'libaom.a','-lm','-o',upscale])
    cases=[]
    def decode(name,alpha,encoded,input_hash,raw=False,mono=False):
        stem=name+('-alpha' if alpha else '-color');obu=work/(stem+'.obu');pixels=work/(stem+'.yuv');obu.write_bytes(encoded)
        result=run([driver,obu,pixels],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True);stages=[json.loads(v) for v in result.stderr.splitlines()];assert len(stages)==5
        for f in stages:
            planes=[bytes.fromhex(v) for v in f['planes']];f['planes']=[base64.b64encode(v).decode() for v in planes];f['planeSha256']=[digest(v) for v in planes]
        # Compare the final observation with the actual decoder output, separately written through aom_image_t.
        assert pixels.read_bytes()==b''.join(base64.b64decode(v) for v in stages[4]['planes'])
        c={'name':name,'alpha':alpha,'monochrome':mono,'inputSha256':input_hash,'cdef':stages[2],'upscaled':stages[3],'restored':stages[4],'finalImage':json.loads(result.stdout)}
        if raw:c['inputBase64']=base64.b64encode(encoded).decode()
        cases.append(c);return c
    for name,alpha,path,expected,offset,length in INPUTS:
        encoded=(repo/path).read_bytes();assert digest(encoded)==expected;decode(name,alpha,encoded[offset:offset+length],expected,mono=alpha)
    controls=json.loads(gzip.decompress((here/'cdef-reference.json.gz').read_bytes()))
    for c in controls['cases']:
        if 'inputBase64' not in c:continue
        encoded=base64.b64decode(c['inputBase64']);assert digest(encoded)==c['inputSha256'];decode(c['name'],False,encoded,c['inputSha256'],True,c['monochrome'])
    for mode in range(16):
        name='restore-control-'+str(mode);obu=work/(name+'.obu');run([encoder,obu,str(mode)]);encoded=obu.read_bytes()
        decode(name,False,encoded,digest(encoded),True,mode in (2,14))
    numeric=[json.loads(v) for v in run([kernel],stdout=subprocess.PIPE,text=True).stdout.splitlines()]
    upscaling=[json.loads(v) for v in run([upscale],stdout=subprocess.PIPE,text=True).stdout.splitlines()]
    fixture={'nativeCommit':COMMIT,'reference':'Actual AOM post-CDEF/final restored samples and unmodified filter kernels; final capture equals separately retrieved decoder output','cases':cases,'kernels':numeric,'upscaleKernels':upscaling}
    canonical=(json.dumps(fixture,separators=(',',':'))+'\n').encode();output=args.output.resolve() if args.output else work/'restored-reference.json.gz';output.write_bytes(gzip.compress(canonical,mtime=0))
    assets=['GenerateRestoredFrames.py','GenerateTileFixtures.py','GenerateReconstructionFixtures.py','TraceSuperres.patch','OfficeReconstructionProbe.inc','OfficeRestorationProbe.inc','ReadTileTrace.c','ReadRestorationKernels.c','ReadUpscaleKernels.c','EncodeRestorationControls.c','cdef-reference.json.gz']
    types=sorted({u['type'] for c in cases for u in c['restored']['restoration']})
    assert 1 in types and 2 in types,'Native controls must exercise both Wiener and self-guided restoration'
    receipt={'nativeCommit':COMMIT,'decodeframeSourceSha256':SOURCE,'assetsSha256':{n:digest((here/n).read_bytes()) for n in assets},'fixtureSha256':digest(output.read_bytes()),'canonicalJsonSha256':digest(canonical),'cases':len(cases),'nativeUnitTypes':types,'upscaleKernels':len(upscaling),'kernels':len(numeric),'changedKernelCases':sum(v['input']!=v['output'] for v in numeric)}
    (work/'restoration-oracle-receipt.json').write_text(json.dumps(receipt,indent=2)+'\n');print(json.dumps(receipt,indent=2))

if __name__=='__main__':main()
