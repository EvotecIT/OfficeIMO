"""Test-only LittleCMS references from native YCbCr words before RGB quantization."""
from pathlib import Path
import csv, ctypes as C, ctypes.util, hashlib, json, os, struct
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
 bits=int(row['bits']);maximum=(1<<bits)-1;midpoint=1<<(bits-1);width=int(row['width']);height=int(row['height'])
 jpeg=root.parent/row['source'];assert hashlib.sha256(jpeg.read_bytes()).hexdigest()==row['sourceSha256']
 if row['name'].startswith('huffman'):
  raw=Path(str(jpeg)+'.raw').read_bytes();words=struct.unpack('<'+'H'*(len(raw)//2),raw)
 else:words=[((x*193+y*791+c*3191)^(x*y*53))&maximum for y in range(height)for x in range(width)for c in range(3)]
 values=[]
 for at in range(0,len(words),3):
  y,cb,cr=words[at:at+3];cb-=midpoint;cr-=midpoint
  r=y+1.402*cr;b=y+1.772*cb;g=(y-.299*r-.114*b)/.587
  values.extend(max(0,min(1,c/maximum))for c in(r,g,b))
 inp=(C.c_double*len(values))(*values);out=(C.c_ubyte*(width*height*3))();run(transform,inp,out,width*height)
 (root/(row['name']+'.icc-rgba')).write_bytes(bytes(v for i in range(width*height)for v in(*out[i*3:i*3+3],255)))
(root/'icc-reference.json').write_text(json.dumps(dict(littleCmsVersion=api('cmsGetEncodedCMMversion',C.c_int)(),profileSHA256=hashlib.sha256(profile.read_bytes()).hexdigest(),intent='relative colorimetric',input='native YCbCr converted to normalized RGB without intermediate sample rounding'),indent=2)+'\n')
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n'for p in sorted(root.glob('*.tif*'))))
api('cmsDeleteTransform',None,P)(transform);api('cmsCloseProfile',C.c_int,P)(source);api('cmsCloseProfile',C.c_int,P)(target)
print(len(list(root.glob('*.icc-rgba'))),'ICC references')
