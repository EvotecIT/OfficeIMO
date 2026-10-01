#!/usr/bin/env python3
"""Opt-in pinned AOM full-image decoder trace; native dependencies remain outside normal build/restore."""
import argparse
import gzip
import hashlib
import json
import pathlib
import shutil
import subprocess

COMMIT='d772e334cc724105040382a977ebb10dfd393293'
SOURCES={
    'PATENTS':'661fb8e504744e95587b556b94a58343448300606a41bea8c7a9b97125696e61',
    'LICENSE':'4764a286d8b2faeaf42f4418e7d7a28d58fc8fd4d00a3d0a7f44b0a4099de7f2',
    'aom_dsp/binary_codes_writer.c':'0dd2d7d157be8df162c64ca91b609a11b1717540b75966a44892d0fd951e2304',
    'av1/common/entropymode.c':'13c47672c1e00d77de9b47b5cf7915cf2569a62369bba35dec3ddb5f43f85e24',
    'av1/common/restoration.c':'5a547e46f94b52e8117cc9e9f47e60c9677e6ea3e82f3f1a49047fd57611c92d',
    'av1/decoder/decodeframe.c':'653fe4bc6556f48064e20058057f231902bd9740c66a8930c33a059f4273ac51',
    'av1/decoder/decodetxb.c':'cd87de06a3f215926b58453c3efda90b43c0f9ac3242ad87309b1703863c8243',
}
INPUTS=[
    ('avif-opaque',False,'OfficeIMO.TestAssets/Documents/Html/Qualification/StaticPdfGaps/avif-opaque.avif','1980b75f79092e557b0fbf28774f87ad39cfc6b13dbecbb3bf6eb60af27c7ad3',275,58),
    ('avif-alpha',False,'OfficeIMO.TestAssets/Documents/Html/Qualification/StaticPdfGaps/avif-alpha.avif','1ac1730fae1498d1d0662bd917e63099393692030139c688f2b66b010b29f3c6',490,58),
    ('avif-alpha',True,'OfficeIMO.TestAssets/Documents/Html/Qualification/StaticPdfGaps/avif-alpha.avif','1ac1730fae1498d1d0662bd917e63099393692030139c688f2b66b010b29f3c6',430,60),
    ('multitile',False,'OfficeIMO.Drawing.Tests/TestAssets/Avif/multitile.avif','465ea0b0086c23b60cb6234eba16efdceecc4775cf37854b314020cfb103b25a',275,4831),
]

def digest(data): return hashlib.sha256(data).hexdigest()
def run(args,**kwargs): return subprocess.run([str(a) for a in args],check=True,**kwargs)

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--work-dir',required=True,type=pathlib.Path)
    parser.add_argument('--output',type=pathlib.Path)
    parser.add_argument('--cc',default='clang')
    parser.add_argument('--restoration-output',type=pathlib.Path)
    args=parser.parse_args();here=pathlib.Path(__file__).resolve().parent;repo=here.parents[2]
    work=args.work_dir.resolve();work.mkdir(parents=True,exist_ok=True)
    source=work/'native-aom';build=work/'native-build';patch=here/'TraceTileDecoder.patch'
    if not source.exists():
        run(['git','clone','--depth','1','--branch','v3.13.1','https://aomedia.googlesource.com/aom',source])
    head=subprocess.check_output(['git','-C',str(source),'rev-parse','HEAD'],text=True).strip()
    if head!=COMMIT: raise ValueError('Unexpected native reference revision')
    for name,expected in SOURCES.items():
        data=subprocess.check_output(['git','-C',str(source),'show','HEAD:'+name])
        if digest(data)!=expected: raise ValueError('Native baseline hash mismatch: '+name)
    diff=subprocess.check_output(['git','-C',str(source),'diff','--no-ext-diff','--unified=0'])
    if diff and diff!=patch.read_bytes(): raise ValueError('Unexpected native source edits')
    if not diff: run(['git','-C',source,'apply','--unidiff-zero',patch])
    if subprocess.check_output(['git','-C',str(source),'diff','--no-ext-diff','--unified=0'])!=patch.read_bytes():
        raise ValueError('Native trace patch mismatch')
    with (work/'native-configure.log').open('w') as log:
        run(['cmake','-S',source,'-B',build,'-DCMAKE_BUILD_TYPE=Release','-DENABLE_DOCS=0','-DENABLE_TESTS=0','-DENABLE_EXAMPLES=0','-DENABLE_TOOLS=0','-DCONFIG_AV1_ENCODER=0','-DCONFIG_MULTITHREAD=0','-DCONFIG_RUNTIME_CPU_DETECT=0'],stdout=log,stderr=subprocess.STDOUT)
    with (work/'native-build.log').open('w') as log:
        run(['cmake','--build',build,'-j','4'],stdout=log,stderr=subprocess.STDOUT)
    driver=here/'ReadTileTrace.c';executable=work/'read-tile-trace'
    run([shutil.which(args.cc),'-std=c11','-Wall','-Wextra','-Werror','-O2','-I',source,driver,build/'libaom.a','-lm','-o',executable])
    cases=[]
    for name,alpha,path,expected,offset,length in INPUTS:
        encoded=(repo/path).read_bytes()
        if digest(encoded)!=expected: raise ValueError('Frozen image hash mismatch')
        stem=name+('-alpha' if alpha else '-color');obu=work/(stem+'.obu');obu.write_bytes(encoded[offset:offset+length]);pixels=work/(stem+'.yuv')
        with (work/(stem+'-trace.jsonl')).open('w') as trace:
            result=run([executable,obu,pixels],stdout=subprocess.PIPE,stderr=trace,text=True)
        image=json.loads(result.stdout);events=[json.loads(line) for line in (work/(stem+'-trace.jsonl')).read_text().splitlines()]
        raw=pixels.read_bytes();planeHashes=[];start=0
        for size in image['planeBytes']: planeHashes.append(digest(raw[start:start+size]));start+=size
        assert start==len(raw)
        tiles=[]
        for event in events:
            if event['event']=='tile': tiles.append({'extent':event['extent'],'events':[]})
            else:
                if event['event']=='restoration' and event['type']!=0: event['type']+=1
                tiles[-1]['events'].append(event)
        assert tiles and all(t['events'] for t in tiles)
        cases.append({'name':name,'alpha':alpha,'sha256':expected,'itemOffset':offset,'itemLength':length,'image':image,'planeSha256':planeHashes,'tiles':tiles})
    restorationHarness=here/'GenerateRestorationFixtures.c';restorationExecutable=work/'generate-restoration'
    command=[shutil.which(args.cc),'-std=c11','-Wall','-Wextra','-Werror','-O2','-I',build,'-I',source,restorationHarness]
    command += [source/'aom_dsp'/name for name in ('bitwriter.c','binary_codes_writer.c','entenc.c')]
    run(command+[build/'libaom.a','-lm','-o',restorationExecutable])
    with (work/'restoration-cases.jsonl').open('w') as out, (work/'restoration-trace.jsonl').open('w') as trace:
        run([restorationExecutable],stdout=out,stderr=trace)
    restorationCases=[json.loads(line) for line in (work/'restoration-cases.jsonl').read_text().splitlines()]
    current=None
    for line in (work/'restoration-trace.jsonl').read_text().splitlines():
        event=json.loads(line)
        if event['event']=='case':
            current=restorationCases[event['scenario']];current['events']=[]
        else:
            assert current is not None and event['event']=='restoration'
            if event['type']!=0: event['type']+=1
            current['events'].append(event)
    restorationFixture={'reference':'AOM v3.13.1 native restoration reader, writer and superblock geometry','nativeCommit':COMMIT,'nativeSelfCheck':True,'cases':restorationCases}
    restorationCanonical=(json.dumps(restorationFixture,separators=(',',':'))+'\n').encode()
    restorationOutput=args.restoration_output.resolve() if args.restoration_output else work/'restoration-reference.json.gz'
    restorationOutput.write_bytes(gzip.compress(restorationCanonical,mtime=0))
    fixture={'reference':'AOM v3.13.1 complete native decoder, single threaded','nativeCommit':COMMIT,'cases':cases}
    canonical=(json.dumps(fixture,separators=(',',':'))+'\n').encode();output=args.output.resolve() if args.output else work/'tile-reference.json.gz'
    output.write_bytes(gzip.compress(canonical,mtime=0))
    receipt={'nativeCommit':COMMIT,'sourceSha256':SOURCES,'patchSha256':digest(patch.read_bytes()),'driverSha256':digest(driver.read_bytes()),'generatorSha256':digest(pathlib.Path(__file__).read_bytes()),'fixtureSha256':digest(output.read_bytes()),'canonicalJsonSha256':digest(canonical),'restorationHarnessSha256':digest(restorationHarness.read_bytes()),'restorationFixtureSha256':digest(restorationOutput.read_bytes()),'restorationCases':len(restorationCases),'restorationUnits':sum(len(c['events']) for c in restorationCases),'restorationNativeSelfCheck':True,'cases':len(cases),'tiles':sum(len(c['tiles']) for c in cases),'leaves':sum(e['event']=='leaf' for c in cases for t in c['tiles'] for e in t['events']),'residuals':sum(e['event']=='residual' for c in cases for t in c['tiles'] for e in t['events'])}
    (work/'tile-oracle-receipt.json').write_text(json.dumps(receipt,indent=2)+'\n')
    print(json.dumps(receipt,indent=2))

if __name__=='__main__': main()
