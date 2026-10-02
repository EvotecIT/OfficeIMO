#!/usr/bin/env python3
"""Opt-in actual AOM post-superresolution frames and unchanged normative row kernels."""
import argparse,base64,gzip,json,pathlib,shutil,subprocess,sys
sys.dont_write_bytecode=True
from GenerateTileFixtures import COMMIT,run,digest
from GenerateReconstructionFixtures import SOURCE,prepare_decoder

def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--work-dir',required=True,type=pathlib.Path);ap.add_argument('--output',type=pathlib.Path)
    ap.add_argument('--native-source',type=pathlib.Path);ap.add_argument('--native-build',type=pathlib.Path)
    args=ap.parse_args()
    if bool(args.native_source)!=bool(args.native_build):ap.error('Native source/build must be supplied together')
    here=pathlib.Path(__file__).resolve().parent;repo=here.parents[2];work=args.work_dir.resolve();work.mkdir(parents=True,exist_ok=True)
    source,build,driver=prepare_decoder(here,work,10,args.native_source,args.native_build,'TraceUpscale.patch','read-upscaled-frame')
    sources={'av1/common/resize.c':'86432c307a87b05ce213fd72bdfa7343e1b28af6360650e4c1585fb8b18639b9',
        'av1/common/convolve.c':'979df52cfcbe8046d9057aecc350b62128f8801c01a810e16988ca0e31e511dd'}
    for n,h in sources.items():assert digest((source/n).read_bytes())==h
    common=[shutil.which('clang'),'-std=c11','-Wall','-Wextra','-Werror','-O2','-I',source,'-I',build]
    encoder=work/'encode-upscale-controls';kernels=work/'read-upscale-kernels'
    run(common+[here/'EncodeUpscaleControls.c',build/'libaom.a','-lm','-o',encoder])
    run(common+[here/'ReadUpscaleKernels.c',build/'libaom.a','-lm','-o',kernels])
    cases=[]
    def decode(name,alpha,mono,encoded,input_hash,raw=False):
        stem=name+('-alpha' if alpha else '-color');obu=work/(stem+'.obu');yuv=work/(stem+'.yuv');obu.write_bytes(encoded)
        r=run([driver,obu,yuv],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True)
        stages=[json.loads(v) for v in r.stderr.splitlines()];assert len(stages)==5
        for stage in stages:
            planes=[bytes.fromhex(v) for v in stage['planes']]
            stage['planes']=[base64.b64encode(v).decode() for v in planes];stage['planeSha256']=[digest(v) for v in planes]
        final=json.loads(r.stdout);assert final['depth']==10
        c={'name':name,'alpha':alpha,'monochrome':mono,'inputSha256':input_hash,'cdef':stages[2],'upscaled':stages[3],
            'finalImage':final,'finalYuvSha256':digest(yuv.read_bytes())}
        assert all(len(base64.b64decode(v))==2*((c['upscaled']['width']+(p>0))>>(p>0))*((c['upscaled']['height']+(p>0))>>(p>0)) for p,v in enumerate(c['upscaled']['planes']))
        if raw:c['inputBase64']=base64.b64encode(encoded).decode()
        cases.append(c);return c
    prior=json.loads(gzip.decompress((here/'cdef-main10-reference.json.gz').read_bytes()))
    inputs=json.loads((here/'reconstruction-main10-inputs.json').read_text())['inputs']
    for old in prior['cases']:
        raw='inputBase64' in old
        if raw:encoded=base64.b64decode(old['inputBase64']);assert digest(encoded)==old['inputSha256']
        else:
            v=next(v for v in inputs if v['name']==old['name'] and v['alpha']==old['alpha'])
            container=(repo/v['path']).read_bytes();assert digest(container)==v['sha256']==old['inputSha256']
            encoded=container[v['offset']:v['offset']+v['length']];assert len(encoded)==v['length']
        c=decode(old['name'],old['alpha'],old['monochrome'],encoded,old['inputSha256'],raw)
        assert c['cdef']==old['cdef'] and c['finalImage']==old['finalImage'] and c['finalYuvSha256']==old['finalYuvSha256']
    for mode in range(10):
        name='upscale-control-'+str(mode);obu=work/(name+'.obu');run([encoder,obu,str(mode),'10'])
        encoded=obu.read_bytes();c=decode(name,False,mode==8,encoded,digest(encoded),True)
        assert c['cdef']['width']<c['upscaled']['width'] and c['cdef']['tileCols']>1
        assert b''.join(base64.b64decode(v) for v in c['upscaled']['planes'])==(work/(name+'-color.yuv')).read_bytes(), 'Restoration-disabled native final output must equal upscale observation'
    numeric=[json.loads(v) for v in run([kernels,'10'],stdout=subprocess.PIPE,text=True).stdout.splitlines()]
    eight=[json.loads(v) for v in run([kernels],stdout=subprocess.PIPE,text=True).stdout.splitlines()]
    original=json.loads(gzip.decompress((here/'restored-reference.json.gz').read_bytes()))
    assert eight==original['upscaleKernels'],'Every original eight-bit native row must remain unchanged'
    fixture={'nativeCommit':COMMIT,'reference':'Actual unmodified AOM row upscaler and access-only post-superresolution frame observations','cases':cases,'kernels':numeric}
    canonical=(json.dumps(fixture,separators=(',',':'))+'\n').encode();output=args.output.resolve() if args.output else work/'upscaled-main10-reference.json.gz'
    output.write_bytes(gzip.compress(canonical,mtime=0))
    assets=['GenerateUpscaleFixtures.py','GenerateReconstructionFixtures.py','TraceUpscale.patch','OfficeUpscaleProbe.inc','OfficeReconstructionProbe.inc','ReadTileTrace.c','ReadUpscaleKernels.c','EncodeUpscaleControls.c','cdef-main10-reference.json.gz','restored-reference.json.gz']
    receipt={'nativeCommit':COMMIT,'decodeframeSourceSha256':SOURCE,'upscaleSourceSha256':sources,
        'configurationSha256':digest((build/'config/aom_config.h').read_bytes()),'librarySha256':digest((build/'libaom.a').read_bytes()),
        'assetsSha256':{n:digest((here/n).read_bytes()) for n in assets},'fixtureSha256':digest(output.read_bytes()),'canonicalJsonSha256':digest(canonical),
        'cases':len(cases),'activeUpscaleFrames':10,'kernels':len(numeric),'lowBitPhases':[0,1,2,3],
        'eightBitRowsIdentical':len(eight),'frameSamples':sum(len(base64.b64decode(v))//2 for c in cases for v in c['upscaled']['planes'])}
    (work/'upscale-oracle-receipt.json').write_text(json.dumps(receipt,indent=2)+'\n');print(json.dumps(receipt,indent=2))
if __name__=='__main__':main()
