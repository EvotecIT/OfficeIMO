#!/usr/bin/env python3
"""Opt-in native Main10 restoration; observations and encoder unit-size controls only."""
import argparse,base64,gzip,json,os,pathlib,shutil,subprocess,sys
sys.dont_write_bytecode=True
from GenerateTileFixtures import COMMIT,run,digest
from GenerateReconstructionFixtures import SOURCE,prepare_decoder

def main():
    ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('--work-dir',required=True,type=pathlib.Path)
    ap.add_argument('--output',type=pathlib.Path);ap.add_argument('--native-source',type=pathlib.Path);ap.add_argument('--native-build',type=pathlib.Path)
    args=ap.parse_args()
    if bool(args.native_source)!=bool(args.native_build):ap.error('Native source/build must be supplied together')
    here=pathlib.Path(__file__).resolve().parent;repo=here.parents[2];work=args.work_dir.resolve();work.mkdir(parents=True,exist_ok=True)
    source,build,driver=prepare_decoder(here,work,10,args.native_source,args.native_build,'TraceSuperres.patch','read-restored-main10-frame')
    sources={'av1/common/restoration.c':'5a547e46f94b52e8117cc9e9f47e60c9677e6ea3e82f3f1a49047fd57611c92d',
        'av1/common/convolve.c':'979df52cfcbe8046d9057aecc350b62128f8801c01a810e16988ca0e31e511dd'}
    for name,expected in sources.items():assert digest((source/name).read_bytes())==expected
    # The pinned source/build is never edited. Only a copied encoder policy unit is patched.
    encoder_source=source/'av1/encoder/pickrst.c';assert digest(encoder_source.read_bytes())=='098561c7dd795d9605625c4d15c997959b9cf275f52eba9bc2826a6e61809b21'
    policy=work/'producer-policy';unit=policy/'av1/encoder/pickrst.c';unit.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(encoder_source,unit)
    run(['patch','--batch','--forward','-p1','-i',here/'ForceRestorationUnits.patch'],cwd=policy)
    common=[shutil.which('clang'),'-std=c11','-Wall','-Wextra','-Werror','-O2','-I',source,'-I',build]
    obj=work/'pickrst-policy.o';run([common[0],'-std=c99','-O2','-DNDEBUG','-I',source,'-I',build,'-c',unit,'-o',obj])
    encoder=work/'encode-restoration-controls';kernel=work/'read-restoration-kernels'
    run(common+[here/'EncodeRestorationControls.c',obj,build/'libaom.a','-lm','-o',encoder]);run(common+[here/'ReadRestorationKernels.c',build/'libaom.a','-lm','-o',kernel])
    cases=[]
    def decode(name,alpha,mono,encoded,input_hash,depth,raw=False):
        stem=name+('-alpha' if alpha else '-color');obu=work/(stem+'.obu');yuv=work/(stem+'.yuv');obu.write_bytes(encoded)
        r=run([driver,obu,yuv],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True);stages=[json.loads(v) for v in r.stderr.splitlines()];assert len(stages)==5
        for stage in stages:
            planes=[bytes.fromhex(v) for v in stage['planes']];stage['planes']=[base64.b64encode(v).decode() for v in planes];stage['planeSha256']=[digest(v) for v in planes]
        final=json.loads(r.stdout);assert final['depth']==depth
        assert b''.join(base64.b64decode(v) for v in stages[4]['planes'])==yuv.read_bytes()
        c={'name':name,'alpha':alpha,'monochrome':mono,'inputSha256':input_hash,'cdef':stages[2],'upscaled':stages[3],'restored':stages[4],'finalImage':final,'finalYuvSha256':digest(yuv.read_bytes())}
        if raw:c['inputBase64']=base64.b64encode(encoded).decode()
        return c
    inputs=json.loads((here/'reconstruction-main10-inputs.json').read_text())['inputs']
    def encoded(old,depth):
        if 'inputBase64' in old:
            data=base64.b64decode(old['inputBase64']);assert digest(data)==old['inputSha256'];return data,True
        if depth==10:
            item=next(v for v in inputs if v['name']==old['name'] and v['alpha']==old['alpha']);data=(repo/item['path']).read_bytes();assert digest(data)==item['sha256']==old['inputSha256'];return data[item['offset']:item['offset']+item['length']],False
        from GenerateTileFixtures import INPUTS
        _,_,path,expected,offset,length=next(v for v in INPUTS if v[0]==old['name'] and v[1]==old['alpha']);data=(repo/path).read_bytes();assert digest(data)==expected==old['inputSha256'];return data[offset:offset+length],False
    prior=json.loads(gzip.decompress((here/'upscaled-main10-reference.json.gz').read_bytes()))
    for old in prior['cases']:
        data,raw=encoded(old,10);c=decode(old['name'],old['alpha'],old['monochrome'],data,old['inputSha256'],10,raw)
        assert c['cdef']==old['cdef'] and c['finalImage']==old['finalImage'] and c['finalYuvSha256']==old['finalYuvSha256']
        assert all(c['upscaled'][k]==v for k,v in old['upscaled'].items());cases.append(c)
    for mode in range(20):
        name='restore-control-'+str(mode);obu=work/(name+'.obu');env=dict(os.environ)
        env.pop('OFFICEIMO_AV1_RESTORATION_UNIT',None);env.pop('OFFICEIMO_AV1_CHROMA_HALF',None)
        if mode>=17:env['OFFICEIMO_AV1_RESTORATION_UNIT']='128' if mode in (17,19) else '64'
        if mode==18:env['OFFICEIMO_AV1_CHROMA_HALF']='1'
        run([encoder,obu,str(mode),'10'],env=env);data=obu.read_bytes();cases.append(decode(name,False,mode in (2,14,19),data,digest(data),10,True))
    numeric=[json.loads(v) for v in run([kernel,'10'],stdout=subprocess.PIPE,text=True).stdout.splitlines()]
    original=json.loads(gzip.decompress((here/'restored-reference.json.gz').read_bytes()));eight=[json.loads(v) for v in run([kernel],stdout=subprocess.PIPE,text=True).stdout.splitlines()]
    assert eight==original['kernels']
    for old in original['cases']:
        data,raw=encoded(old,8);c=decode(old['name'],old['alpha'],old['monochrome'],data,old['inputSha256'],8,raw)
        assert all(c[k]==v for k,v in old.items()),old['name']
    fixture={'nativeCommit':COMMIT,'reference':'Actual unchanged native decoder/filter output, with access-only observations and encoder-only unit-size search controls','cases':cases,'kernels':numeric}
    canonical=(json.dumps(fixture,separators=(',',':'))+'\n').encode();output=args.output.resolve() if args.output else work/'restored-main10-reference.json.gz';output.write_bytes(gzip.compress(canonical,mtime=0))
    assets=['GenerateRestorationMain10Fixtures.py','GenerateReconstructionFixtures.py','TraceSuperres.patch','OfficeRestorationProbe.inc','OfficeReconstructionProbe.inc','ReadTileTrace.c','ReadRestorationKernels.c','EncodeRestorationControls.c','ForceRestorationUnits.patch','upscaled-main10-reference.json.gz','restored-reference.json.gz']
    types=sorted({u['type'] for c in cases for u in c['restored']['restoration']});assert 1 in types and 2 in types
    receipt={'nativeCommit':COMMIT,'decodeframeSourceSha256':SOURCE,'restorationSourcesSha256':{n:digest((source/n).read_bytes()) for n in sources},'encoderPolicySourceSha256':digest(encoder_source.read_bytes()),'configurationSha256':digest((build/'config/aom_config.h').read_bytes()),'librarySha256':digest((build/'libaom.a').read_bytes()),'assetsSha256':{n:digest((here/n).read_bytes()) for n in assets},'fixtureSha256':digest(output.read_bytes()),'canonicalJsonSha256':digest(canonical),'cases':len(cases),'kernels':len(numeric),'changedKernelCases':sum(v['input']!=v['output'] for v in numeric),'nativeUnitTypes':types,'activeUnitSizes':sorted({u['size'] for c in cases for u in c['restored']['restoration'] if u['type']!=0}),'eightBitFramesIdentical':len(original['cases']),'eightBitKernelsIdentical':len(eight)}
    (work/'restoration-oracle-receipt.json').write_text(json.dumps(receipt,indent=2)+'\n');print(json.dumps(receipt,indent=2))
if __name__=='__main__':main()
