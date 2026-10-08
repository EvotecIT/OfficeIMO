"""LittleCMS reference from native words, retaining color and alpha precision."""
from pathlib import Path
import csv,ctypes as C,ctypes.util,hashlib,json,os,struct
root=Path(__file__).resolve().parent
library=os.environ.get('LCMS_LIBRARY') or ctypes.util.find_library('lcms2')
if not library:raise RuntimeError('Set LCMS_LIBRARY to the test-only LittleCMS library.')
lib=C.CDLL(library);P=C.c_void_p;U=C.c_uint32

def api(name,result,*args):
 f=getattr(lib,name);f.restype=result;f.argtypes=args;return f
profile=root.parent/'IccColorCorpus/icc-dci-p3-matrix.icc'
source=api('cmsOpenProfileFromFile',P,C.c_char_p,C.c_char_p)(str(profile).encode(),b'r')
target=api('cmsCreate_sRGBProfile',P)()
transform=api('cmsCreateTransform',P,P,U,P,U,U,U)(source,(1<<22)|(4<<16)|(3<<3),target,(4<<16)|(3<<3)|1,1,0x0140)
assert source and target and transform
run=api('cmsDoTransform',None,P,P,P,U)
for row in csv.DictReader((root/'manifest.csv').open()):
 photo=int(row['photometric'])
 if photo not in(2,6):continue
 bits=int(row['bits']);maximum=(1<<bits)-1;midpoint=1<<(bits-1);kind=int(row['alphaKind'])
 raw=(root/(row['name']+'.raw')).read_bytes();words=struct.unpack('<'+'H'*(len(raw)//2),raw);values=[];alphas=[]
 for at in range(0,len(words),4):
  r,g,b,a=words[at:at+4];alphas.append(int(a*255/maximum+.5))
  if photo==6:
   y=r;cb=g-midpoint;cr=b-midpoint;r=y+1.402*cr;b=y+1.772*cb;g=(y-.299*r-.114*b)/.587
  values.extend(0 if kind==1 and a==0 else max(0,min(1,c/(a if kind==1 else maximum)))for c in(r,g,b))
 inp=(C.c_double*len(values))(*values);out=(C.c_ubyte*(35*19*3))();run(transform,inp,out,35*19)
 (root/(row['name']+'.icc-rgba')).write_bytes(bytes(v for i,a in enumerate(alphas)for v in(*out[i*3:i*3+3],a)))
(root/'icc-reference.json').write_text(json.dumps(dict(littleCmsVersion=api('cmsGetEncodedCMMversion',C.c_int)(),profile=profile.name,profileSHA256=hashlib.sha256(profile.read_bytes()).hexdigest(),intent='relative colorimetric',input='native RGB or fractional YCbCr-derived RGB unassociated with native alpha before ICC'),indent=2)+'\n')
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n'for p in sorted(root.glob('*.tif*'))))
api('cmsDeleteTransform',None,P)(transform);api('cmsCloseProfile',C.c_int,P)(source);api('cmsCloseProfile',C.c_int,P)(target)
print(len(list(root.glob('*.icc-rgba'))),'ICC references')
