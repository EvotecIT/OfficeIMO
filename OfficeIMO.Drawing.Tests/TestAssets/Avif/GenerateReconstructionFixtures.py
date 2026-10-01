#!/usr/bin/env python3
"""Opt-in actual AOM frame pixels before loop filters; no decoder algorithm or runtime dependency changes."""
import argparse,base64,gzip,hashlib,json,pathlib,shutil,subprocess,sys
sys.dont_write_bytecode=True
from GenerateTileFixtures import COMMIT,INPUTS,run,digest

SOURCE='653fe4bc6556f48064e20058057f231902bd9740c66a8930c33a059f4273ac51'

def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--work-dir',required=True,type=pathlib.Path)
    ap.add_argument('--output',type=pathlib.Path)
    ap.add_argument('--bit-depth',type=int,choices=(8,10),default=8)
    ap.add_argument('--native-source',type=pathlib.Path)
    ap.add_argument('--native-build',type=pathlib.Path)
    args=ap.parse_args()
    if bool(args.native_source)!=bool(args.native_build):ap.error('Native source and build must be supplied together')
    here=pathlib.Path(__file__).resolve().parent;repo=here.parents[2];work=args.work_dir.resolve();work.mkdir(parents=True,exist_ok=True)
    source=args.native_source.resolve() if args.native_source else work/'native-aom'
    build=args.native_build.resolve() if args.native_build else work/'native-build'
    patch=here/'TraceReconstruction.patch'
    if not args.native_source:
        if not source.exists():run(['git','clone','--depth','1','--branch','v3.13.1','https://aomedia.googlesource.com/aom',source])
        assert subprocess.check_output(['git','-C',source,'rev-parse','HEAD'],text=True).strip()==COMMIT
        with (work/'native-configure.log').open('w') as log:
            run(['cmake','-S',source,'-B',build,'-DCMAKE_BUILD_TYPE=Release','-DAOM_TARGET_CPU=generic','-DENABLE_DOCS=0','-DENABLE_TESTS=0','-DENABLE_EXAMPLES=0','-DENABLE_TOOLS=0','-DCONFIG_AV1_ENCODER=1','-DCONFIG_AV1_HIGHBITDEPTH=1','-DCONFIG_MULTITHREAD=0','-DCONFIG_RUNTIME_CPU_DETECT=0'],stdout=log,stderr=subprocess.STDOUT)
        with (work/'native-build.log').open('w') as log:run(['cmake','--build',build,'-j','4'],stdout=log,stderr=subprocess.STDOUT)
    assert digest((source/'av1/decoder/decodeframe.c').read_bytes())==SOURCE
    if args.bit_depth==10 and '#define CONFIG_AV1_HIGHBITDEPTH 1' not in (build/'config/aom_config.h').read_text():
        raise ValueError('Native build lacks high-bit-depth support')
    for name in ['LICENSE','PATENTS']:shutil.copyfile(source/name,work/name)
    # Access-only observation in a copied decoder unit; never patch the reusable producer/build.
    probe=work/'observation';unit=probe/'av1/decoder/decodeframe.c';unit.parent.mkdir(parents=True,exist_ok=True)
    shutil.copyfile(source/'av1/decoder/decodeframe.c',unit)
    run(['patch','--batch','--forward','-p1','-i',patch],cwd=probe)
    obj=work/'decodeframe-probe.o';compiler=shutil.which('clang')
    run([compiler,'-std=c99','-O2','-DNDEBUG','-I',source,'-I',build,'-I',here,'-c',unit,'-o',obj])
    executable=work/'read-reconstructed-frame'
    run([compiler,'-std=c11','-Wall','-Wextra','-Werror','-O2','-I',source,here/'ReadTileTrace.c',obj,build/'libaom.a','-lm','-o',executable])
    encoder=work/'encode-reconstruction-controls'
    run([compiler,'-std=c11','-Wall','-Wextra','-Werror','-O2','-I',source,here/'EncodeReconstructionControls.c',build/'libaom.a','-lm','-o',encoder])
    inputs=INPUTS if args.bit_depth==8 else [
        (v['name'],v['alpha'],v['path'],v['sha256'],v['offset'],v['length'])
        for v in json.loads((here/'reconstruction-main10-inputs.json').read_text())['inputs']]
    cases=[]
    for name,alpha,path,expected,offset,length in inputs:
        encoded=(repo/path).read_bytes();assert digest(encoded)==expected
        assert 0<=offset and 0<length and offset+length<=len(encoded)
        stem=name+('-alpha' if alpha else '-color');obu=work/(stem+'.obu');obu.write_bytes(encoded[offset:offset+length]);pixels=work/(stem+'.yuv')
        result=run([executable,obu,pixels],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True)
        final=json.loads(result.stdout);frame=json.loads(result.stderr);raw=[bytes.fromhex(v) for v in frame['planes']]
        assert final['depth']==args.bit_depth
        assert all(len(v)==(frame['miCols']*4>>(p>0))*(frame['miRows']*4>>(p>0))*(2 if args.bit_depth==10 else 1) for p,v in enumerate(raw))
        frame['planes']=[base64.b64encode(v).decode() for v in raw];frame['planeSha256']=[digest(v) for v in raw]
        cases.append({'name':name,'alpha':alpha,'inputSha256':expected,'itemOffset':offset,'itemLength':length,'unfiltered':frame,'finalImage':final,'finalYuvSha256':digest(pixels.read_bytes())})
    for name in ['lossless-copy']:
        obu=work/(name+'.obu');pixels=work/(name+'.yuv');run([encoder,obu,str(args.bit_depth)])
        result=run([executable,obu,pixels],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True);final=json.loads(result.stdout);frame=json.loads(result.stderr)
        assert final['depth']==args.bit_depth
        raw=[bytes.fromhex(v) for v in frame['planes']];frame['planes']=[base64.b64encode(v).decode() for v in raw];frame['planeSha256']=[digest(v) for v in raw]
        assert frame['skipLeaves']>0
        assert frame['copyLeaves']>0 and frame['fractionalCopyLeaves']>0
        cases.append({'name':name,'alpha':False,'inputBase64':base64.b64encode(obu.read_bytes()).decode(),'inputSha256':digest(obu.read_bytes()),'unfiltered':frame,'finalImage':final,'finalYuvSha256':digest(pixels.read_bytes())})
    fixture={'nativeCommit':COMMIT,'reference':'Actual AOM v3.13.1 single-threaded decoder before frame filtering','cases':cases}
    canonical=(json.dumps(fixture,separators=(',',':'))+'\n').encode();output=args.output.resolve() if args.output else work/('reconstruction-reference.json.gz' if args.bit_depth==8 else 'reconstruction-main10-reference.json.gz');output.write_bytes(gzip.compress(canonical,mtime=0))
    assets=['GenerateReconstructionFixtures.py','GenerateTileFixtures.py','TraceReconstruction.patch','OfficeReconstructionProbe.inc','ReadTileTrace.c','EncodeReconstructionControls.c','reconstruction-main10-inputs.json']
    receipt={'nativeCommit':COMMIT,'decodeframeSourceSha256':SOURCE,'bitDepth':args.bit_depth,'configurationSha256':digest((build/'config/aom_config.h').read_bytes()),'librarySha256':digest((build/'libaom.a').read_bytes()),'assetsSha256':{n:digest((here/n).read_bytes()) for n in assets},'fixtureSha256':digest(output.read_bytes()),'canonicalJsonSha256':digest(canonical),'cases':len(cases),'samples':sum(len(base64.b64decode(p))//(2 if args.bit_depth==10 else 1) for c in cases for p in c['unfiltered']['planes'])}
    (work/'reconstruction-oracle-receipt.json').write_text(json.dumps(receipt,indent=2)+'\n');print(json.dumps(receipt,indent=2))

if __name__=='__main__':main()
