#!/usr/bin/env python3
"""Opt-in actual native frame CDEF pixels and native direction/filter kernels."""
import argparse,base64,gzip,json,pathlib,shutil,subprocess,sys
sys.dont_write_bytecode=True
from GenerateTileFixtures import COMMIT,INPUTS,run,digest
from GenerateReconstructionFixtures import SOURCE,prepare_decoder
from GenerateDeblockFixtures import samples

def main():
    ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('--work-dir',required=True,type=pathlib.Path);ap.add_argument('--output',type=pathlib.Path)
    ap.add_argument('--bit-depth',type=int,choices=(8,10),default=8)
    ap.add_argument('--native-source',type=pathlib.Path);ap.add_argument('--native-build',type=pathlib.Path);args=ap.parse_args()
    if bool(args.native_source)!=bool(args.native_build):ap.error('Native source and build must be supplied together')
    here=pathlib.Path(__file__).resolve().parent;repo=here.parents[2];work=args.work_dir.resolve();work.mkdir(parents=True,exist_ok=True)
    source,build,driver=prepare_decoder(here,work,args.bit_depth,args.native_source,args.native_build,'TraceCdef.patch','read-cdef-frame')
    sources={'av1/common/cdef.c':'a2832163c6041efb3085b8322b55b3d531a69a4641db14bdb4a6e01189c2fb3f',
        'av1/common/cdef_block.c':'1a332f8523bb1daf8a21c6b67aaa9d53db760a18cf1fed8260c92cee017168ac',
        'av1/common/cdef_block.h':'f41e9021c4869741c1d9711ed03d50ac3b1f132f04a1c394a11066d9cd2cb0b5'}
    for name,expected in sources.items():assert digest((source/name).read_bytes())==expected
    common=[shutil.which('clang'),'-std=c11','-Wall','-Wextra','-Werror','-O2','-I',source,'-I',build]
    kernels=work/'read-cdef-kernels';run(common+[here/'ReadCdefKernels.c',build/'libaom.a','-lm','-o',kernels])
    cases=[]
    def decode(name,alpha,encoded,input_hash,raw=False,mono=False):
        stem=name+('-alpha' if alpha else '-color');obu=work/(stem+'.obu');pixels=work/(stem+'.yuv');obu.write_bytes(encoded)
        result=run([driver,obu,pixels],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True);stages=[json.loads(v) for v in result.stderr.splitlines()];assert len(stages)==3
        for frame in stages:
            planes=[bytes.fromhex(v) for v in frame['planes']];frame['planes']=[base64.b64encode(v).decode() for v in planes];frame['planeSha256']=[digest(v) for v in planes]
            assert all(len(samples(p,args.bit_depth))==(frame['miCols']*4>>(i>0))*(frame['miRows']*4>>(i>0)) for i,p in enumerate(frame['planes']))
        final=json.loads(result.stdout);assert final['depth']==args.bit_depth
        c={'name':name,'alpha':alpha,'monochrome':mono,'inputSha256':input_hash,'unfiltered':stages[0],'deblocked':stages[1],'cdef':stages[2],'finalImage':final,'finalYuvSha256':digest(pixels.read_bytes())}
        if raw:c['inputBase64']=base64.b64encode(encoded).decode()
        cases.append(c);return c
    if args.bit_depth==10:
        controls=json.loads(gzip.decompress((here/'deblock-main10-reference.json.gz').read_bytes()))
        inputs=json.loads((here/'reconstruction-main10-inputs.json').read_text())['inputs']
        for old in controls['cases']:
            raw='inputBase64' in old
            if raw:
                encoded=base64.b64decode(old['inputBase64']);assert digest(encoded)==old['inputSha256']
            else:
                v=next(v for v in inputs if v['name']==old['name'] and v['alpha']==old['alpha'])
                container=(repo/v['path']).read_bytes();assert digest(container)==v['sha256']==old['inputSha256']
                encoded=container[v['offset']:v['offset']+v['length']];assert len(encoded)==v['length']
            c=decode(old['name'],old['alpha'],encoded,old['inputSha256'],raw,len(old['unfiltered']['planes'])==1)
            assert c['unfiltered']==old['unfiltered'] and c['deblocked']==old['deblocked'] and c['finalYuvSha256']==old['finalYuvSha256']
            if old['name'].startswith('odd-'):assert c['cdef']['planes']!=c['deblocked']['planes'],'Control must demonstrate active CDEF'
    else:
        for name,alpha,path,expected,offset,length in INPUTS:
            encoded=(repo/path).read_bytes();assert digest(encoded)==expected;decode(name,alpha,encoded[offset:offset+length],expected,mono=alpha)
        controls=json.loads(gzip.decompress((here/'reconstruction-reference.json.gz').read_bytes()))
        c=next(v for v in controls['cases'] if v['name']=='lossless-copy');encoded=base64.b64decode(c['inputBase64']);assert digest(encoded)==c['inputSha256'];decode('lossless-copy',False,encoded,c['inputSha256'],True)
        encoder=work/'encode-filter-controls';run(common+[here/'EncodeFilterControls.c',build/'libaom.a','-lm','-o',encoder])
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
    numeric=[json.loads(v) for v in run([kernels,str(args.bit_depth)],stdout=subprocess.PIPE,text=True).stdout.splitlines()]
    fixture={'nativeCommit':COMMIT,'reference':'Actual AOM single-threaded pre/post-CDEF planes and unmodified direction/filter kernels','cases':cases,'directions':[v for v in numeric if v['kind']=='direction'],'kernels':[v for v in numeric if v['kind']=='kernel']}
    canonical=(json.dumps(fixture,separators=(',',':'))+'\n').encode();output=args.output.resolve() if args.output else work/('cdef-main10-reference.json.gz' if args.bit_depth==10 else 'cdef-reference.json.gz');output.write_bytes(gzip.compress(canonical,mtime=0))
    assets=['GenerateCdefFixtures.py','GenerateDeblockFixtures.py','GenerateSvtCdefControl.py','cdef-svt-mixed-skip.obu','GenerateTileFixtures.py','GenerateReconstructionFixtures.py','TraceCdef.patch','OfficeReconstructionProbe.inc','ReadTileTrace.c','EncodeFilterControls.c','ReadCdefKernels.c','reconstruction-reference.json.gz','deblock-main10-reference.json.gz','reconstruction-main10-inputs.json']
    receipt={'nativeCommit':COMMIT,'decodeframeSourceSha256':SOURCE,'cdefSourceSha256':sources,'bitDepth':args.bit_depth,'configurationSha256':digest((build/'config/aom_config.h').read_bytes()),'librarySha256':digest((build/'libaom.a').read_bytes()),'assetsSha256':{n:digest((here/n).read_bytes()) for n in assets},'fixtureSha256':digest(output.read_bytes()),'canonicalJsonSha256':digest(canonical),'cases':len(cases),'directions':len(fixture['directions']),'kernels':len(fixture['kernels']),'changedKernelCases':sum(v['input']!=v['output'] for v in fixture['kernels']),'changedCdefSamples':sum(sum(a!=b for a,b in zip(samples(u,args.bit_depth),samples(d,args.bit_depth))) for c in cases for u,d in zip(c['deblocked']['planes'],c['cdef']['planes']))}
    (work/'cdef-oracle-receipt.json').write_text(json.dumps(receipt,indent=2)+'\n');print(json.dumps(receipt,indent=2))

if __name__=='__main__':main()
