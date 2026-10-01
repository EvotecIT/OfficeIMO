#!/usr/bin/env python3
"""Opt-in independently decoded pre/post-deblock pixels and actual native narrow/wide kernels."""
import argparse,base64,gzip,json,pathlib,shutil,subprocess,sys
sys.dont_write_bytecode=True
from GenerateTileFixtures import COMMIT,INPUTS,run,digest
from GenerateReconstructionFixtures import SOURCE,prepare_decoder

def samples(encoded,depth):
    raw=base64.b64decode(encoded)
    return raw if depth==8 else [int.from_bytes(raw[i:i+2],'little') for i in range(0,len(raw),2)]

def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--work-dir',required=True,type=pathlib.Path);ap.add_argument('--output',type=pathlib.Path)
    ap.add_argument('--bit-depth',type=int,choices=(8,10),default=8)
    ap.add_argument('--native-source',type=pathlib.Path);ap.add_argument('--native-build',type=pathlib.Path)
    args=ap.parse_args()
    if bool(args.native_source)!=bool(args.native_build):ap.error('Native source and build must be supplied together')
    here=pathlib.Path(__file__).resolve().parent;repo=here.parents[2];work=args.work_dir.resolve();work.mkdir(parents=True,exist_ok=True)
    source,build,driver=prepare_decoder(here,work,args.bit_depth,args.native_source,args.native_build,
        'TraceDeblocking.patch','read-deblocked-frame')
    sources={'aom_dsp/loopfilter.c':'bb9d6de53a4ad448af145892fc59dd72c767224d02532409305505d9d330fdb5',
        'av1/common/av1_loopfilter.c':'385077f5637f33947d4a8ef3dca02931cbce4d4764a19124be4860d3494d71bd'}
    for name,expected in sources.items():assert digest((source/name).read_bytes())==expected
    common=[shutil.which('clang'),'-std=c11','-Wall','-Wextra','-Werror','-O2','-I',source,'-I',build]
    kernels=work/'read-deblock-kernels';run(common+[here/'ReadDeblockKernels.c',build/'libaom.a','-lm','-o',kernels])
    inputs=INPUTS if args.bit_depth==8 else [(v['name'],v['alpha'],v['path'],v['sha256'],v['offset'],v['length'])
        for v in json.loads((here/'reconstruction-main10-inputs.json').read_text())['inputs']]
    reconstruction='reconstruction-main10-reference.json.gz' if args.bit_depth==10 else 'reconstruction-reference.json.gz'
    controls=json.loads(gzip.decompress((here/reconstruction).read_bytes()))
    cases=[]
    for name,alpha,path,expected,offset,length in inputs:
        encoded=(repo/path).read_bytes();assert digest(encoded)==expected
        assert 0<=offset and 0<length and offset+length<=len(encoded)
        stem=name+('-alpha' if alpha else '-color');obu=work/(stem+'.obu');obu.write_bytes(encoded[offset:offset+length]);pixels=work/(stem+'.yuv')
        result=run([driver,obu,pixels],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True);stages=[json.loads(v) for v in result.stderr.splitlines()];assert len(stages)==2
        for frame in stages:
            raw=[bytes.fromhex(v) for v in frame['planes']];frame['planes']=[base64.b64encode(v).decode() for v in raw];frame['planeSha256']=[digest(v) for v in raw]
        cases.append({'name':name,'alpha':alpha,'inputSha256':expected,'unfiltered':stages[0],'deblocked':stages[1],'finalImage':json.loads(result.stdout),'finalYuvSha256':digest(pixels.read_bytes())})
    c=next(v for v in controls['cases'] if v['name']=='lossless-copy');encoded=base64.b64decode(c['inputBase64']);assert digest(encoded)==c['inputSha256']
    obu=work/'lossless-copy.obu';pixels=work/'lossless-copy.yuv';obu.write_bytes(encoded)
    result=run([driver,obu,pixels],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True);stages=[json.loads(v) for v in result.stderr.splitlines()];assert len(stages)==2
    for frame in stages:
        raw=[bytes.fromhex(v) for v in frame['planes']];frame['planes']=[base64.b64encode(v).decode() for v in raw];frame['planeSha256']=[digest(v) for v in raw]
    assert stages[0]['planes']==stages[1]['planes'] and stages[0]['copyLeaves']>0 and stages[0]['skipLeaves']>0
    cases.append({'name':'lossless-copy','alpha':False,'inputBase64':c['inputBase64'],'inputSha256':c['inputSha256'],'unfiltered':stages[0],'deblocked':stages[1],'finalImage':json.loads(result.stdout),'finalYuvSha256':digest(pixels.read_bytes())})
    for case in cases:
        old=next(v for v in controls['cases'] if v['name']==case['name'] and v['alpha']==case['alpha'])
        assert case['unfiltered']['planes']==old['unfiltered']['planes']
        assert case['finalYuvSha256']==old['finalYuvSha256']
        assert case['finalImage']['depth']==args.bit_depth
        for stage in ('unfiltered','deblocked'):
            f=case[stage]
            assert all(len(samples(p,args.bit_depth))==(f['miCols']*4>>(i>0))*(f['miRows']*4>>(i>0)) for i,p in enumerate(f['planes']))
    if args.bit_depth==10:
        encoder=work/'encode-filter-controls';run(common+[here/'EncodeFilterControls.c',build/'libaom.a','-lm','-o',encoder])
        for name,mode in [('odd-color',0),('odd-mono',1),('odd-tiled-color',2)]:
            obu=work/(name+'.obu');pixels=work/(name+'.yuv');run([encoder,obu,str(mode),str(args.bit_depth)])
            encoded=obu.read_bytes();result=run([driver,obu,pixels],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True)
            stages=[json.loads(v) for v in result.stderr.splitlines()];assert len(stages)==2
            for f in stages:
                raw=[bytes.fromhex(v) for v in f['planes']]
                f['planes']=[base64.b64encode(v).decode() for v in raw];f['planeSha256']=[digest(v) for v in raw]
                assert all(len(samples(p,args.bit_depth))==(f['miCols']*4>>(i>0))*(f['miRows']*4>>(i>0)) for i,p in enumerate(f['planes']))
            final=json.loads(result.stdout);assert final['depth']==10
            assert stages[0]['planes']!=stages[1]['planes'],'Control must demonstrate active deblocking'
            if mode==2:assert stages[0]['tileCols']*stages[0]['tileRows']>1
            cases.append({'name':name,'alpha':False,'inputBase64':base64.b64encode(encoded).decode(),'inputSha256':digest(encoded),
                'unfiltered':stages[0],'deblocked':stages[1],'finalImage':final,'finalYuvSha256':digest(pixels.read_bytes())})
    numeric=[json.loads(v) for v in run([kernels,str(args.bit_depth)],stdout=subprocess.PIPE,text=True).stdout.splitlines()]
    fixture={'nativeCommit':COMMIT,'reference':'Actual AOM single-threaded pre/post-deblock planes and native kernels/threshold initialization','cases':cases,'kernels':numeric}
    canonical=(json.dumps(fixture,separators=(',',':'))+'\n').encode();output=args.output.resolve() if args.output else work/('deblock-main10-reference.json.gz' if args.bit_depth==10 else 'deblock-reference.json.gz');output.write_bytes(gzip.compress(canonical,mtime=0))
    assets=['GenerateDeblockFixtures.py','GenerateTileFixtures.py','GenerateReconstructionFixtures.py','TraceDeblocking.patch','OfficeReconstructionProbe.inc','ReadTileTrace.c','ReadDeblockKernels.c','EncodeFilterControls.c',reconstruction,'reconstruction-main10-inputs.json']
    receipt={'nativeCommit':COMMIT,'decodeframeSourceSha256':SOURCE,'loopFilterSourceSha256':sources,'bitDepth':args.bit_depth,'configurationSha256':digest((build/'config/aom_config.h').read_bytes()),'librarySha256':digest((build/'libaom.a').read_bytes()),'assetsSha256':{n:digest((here/n).read_bytes()) for n in assets},'fixtureSha256':digest(output.read_bytes()),'canonicalJsonSha256':digest(canonical),'cases':len(cases),'kernels':len(numeric),'changedKernelCases':sum(v['input']!=v['output'] for v in numeric),'changedFrameSamples':sum(sum(a!=b for a,b in zip(samples(u,args.bit_depth),samples(d,args.bit_depth))) for c in cases for u,d in zip(c['unfiltered']['planes'],c['deblocked']['planes']))}
    (work/'deblock-oracle-receipt.json').write_text(json.dumps(receipt,indent=2)+'\n');print(json.dumps(receipt,indent=2))

if __name__=='__main__':main()
