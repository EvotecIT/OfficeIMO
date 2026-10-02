"""Opt-in actual libavif1.4.2 conversion of ten-bit YUV/alpha to eight-bit straight RGBA."""
import argparse,gzip,hashlib,json,os,pathlib,shutil,subprocess,sys
sys.dont_write_bytecode=True
from GenerateMain10Fixtures import COMMIT,HEADER_SHA

def sha(data):return hashlib.sha256(data).hexdigest()
def main():
    ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('--work-dir',required=True,type=pathlib.Path)
    ap.add_argument('--library',required=True,type=pathlib.Path);ap.add_argument('--header',required=True,type=pathlib.Path)
    ap.add_argument('--output',type=pathlib.Path);args=ap.parse_args()
    here=pathlib.Path(__file__).resolve().parent;work=args.work_dir.resolve();work.mkdir(parents=True,exist_ok=True)
    library=args.library.resolve(strict=True);header=args.header.resolve(strict=True);assert sha(header.read_bytes())==HEADER_SHA
    include=work/'include/avif';include.mkdir(parents=True,exist_ok=True);shutil.copyfile(header,include/'avif.h')
    driver=here/'ReadColorReference.c';exe=work/'read-main10-color-reference'
    subprocess.run(['cc','-O2','-Wall','-Wextra','-Werror','-I',str(work/'include'),str(driver),str(library),'-Wl,-rpath,'+str(library.parent),'-o',str(exe)],check=True)
    env=dict(os.environ)
    if sys.platform=='darwin':env['DYLD_LIBRARY_PATH']=str(library.parent)
    def capture(depth):return json.loads(subprocess.check_output([str(exe),str(depth)],env=env))
    fixture=capture(10);assert fixture['version']=='1.4.2' and len(fixture['cases'])==768
    assert capture(10)==fixture
    original=json.loads(gzip.decompress((here/'color-reference.json.gz').read_bytes()));eight=capture(8)
    assert eight['cases']==original['cases'] and capture(8)==eight
    canonical=json.dumps(fixture,separators=(',',':'),sort_keys=True).encode();output=args.output.resolve() if args.output else work/'color-main10-reference.json.gz'
    output.write_bytes(gzip.compress(canonical,mtime=0))
    receipt={'nativeVersion':fixture['version'],'nativeSourceCommit':COMMIT,'headerSha256':sha(header.read_bytes()),'librarySha256':sha(library.read_bytes()),'driverSha256':sha(driver.read_bytes()),'generatorSha256':sha(pathlib.Path(__file__).read_bytes()),'fixtureSha256':sha(output.read_bytes()),'canonicalJsonSha256':sha(canonical),'cases':768,'eightBitCasesUnchanged':len(eight['cases']),'nativeRepeatIdentical':True,'rgbTolerance':1,'alphaTolerance':0,'rgbDepth':8,'avoidLibYUV':True,'chromaUpsampling':'bilinear','sourceDepth':10,'lowBitPhases':4}
    (work/'color-main10-oracle-receipt.json').write_text(json.dumps(receipt,indent=2)+'\n');print(json.dumps(receipt))
if __name__=='__main__':main()
