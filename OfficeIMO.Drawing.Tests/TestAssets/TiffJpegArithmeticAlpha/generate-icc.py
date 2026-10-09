"""Test-only LittleCMS references from independently decoded TIFF CMYK samples."""
from pathlib import Path
import ctypes as C,ctypes.util,os,json,hashlib,struct
root=Path(__file__).resolve().parent
library=os.environ.get('LCMS_LIBRARY') or C.util.find_library('lcms2')
if not library:raise RuntimeError('Set LCMS_LIBRARY to the existing test-only LittleCMS library.')
lib=C.CDLL(library);P=C.c_void_p;U=C.c_uint32

def api(name,result,*args):
 f=getattr(lib,name);f.restype=result;f.argtypes=args;return f
profile=root.parent/'IccColorCorpus/littlecms-cmyk-lut.icc'
source=api('cmsOpenProfileFromFile',P,C.c_char_p,C.c_char_p)(str(profile).encode(),b'r')
target=api('cmsCreate_sRGBProfile',P)()
transform=api('cmsCreateTransform',P,P,U,P,U,U,U)(source,(1<<22)|(6<<16)|(4<<3),target,(4<<16)|(3<<3)|1,1,0x0100)
assert source and target and transform
run=api('cmsDoTransform',None,P,P,P,U)
for corpus in (root,root.parent/'TiffJpegArithmeticLowAlpha',root.parent/'TiffJpegArithmetic12',root.parent/'TiffJpegArithmeticLosslessColor'):
 maximum=4095 if corpus.name=='TiffJpegArithmetic12' else 255
 for p in sorted(corpus.glob('p5-*.tif')):
  if corpus.name=='TiffJpegArithmeticLosslessColor':maximum=(1<<int(p.stem.split('-b',1)[1].split('-',1)[0]))-1
  extra=int(p.stem.rsplit('-e',1)[1]);raw=Path(str(p)+'.raw').read_bytes();values=[];alpha=[]
  if maximum==4095 or corpus.name=='TiffJpegArithmeticLosslessColor':raw=struct.unpack('<'+'H'*(len(raw)//2),raw)
  stride=4 if extra<0 else 5
  for i in range(0,len(raw),stride):
   a=raw[i+4] if extra>0 else maximum;alpha.append(int(a*255/maximum+.5))
   for c in raw[i:i+4]:values.append((min(1,c/a) if a else 0)*100 if extra==1 else c*100/maximum)
  inp=(C.c_double*len(values))(*values);out=(C.c_ubyte*(35*19*3))();run(transform,inp,out,35*19)
  Path(str(p)+'.icc-rgba').write_bytes(bytes(v for i,a in enumerate(alpha) for v in (*out[i*3:i*3+3],a)))
 files=sorted(corpus.glob('*.icc-rgba'))
 (corpus/'icc-reference.json').write_text(json.dumps(dict(littleCmsVersion=api('cmsGetEncodedCMMversion',C.c_int)(),profileSHA256=hashlib.sha256(profile.read_bytes()).hexdigest(),intent='relative colorimetric',input='native TIFF CMYK; associated colorants divided by native alpha before ICC; no Adobe inversion',sha256={p.name:hashlib.sha256(p.read_bytes()).hexdigest()for p in files}),indent=2)+'\n')
 (corpus/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n'for p in sorted(corpus.glob('*.tif*'))))
 print(corpus.name,len(files),'ICC references')
api('cmsDeleteTransform',None,P)(transform);api('cmsCloseProfile',C.c_int,P)(source);api('cmsCloseProfile',C.c_int,P)(target)
