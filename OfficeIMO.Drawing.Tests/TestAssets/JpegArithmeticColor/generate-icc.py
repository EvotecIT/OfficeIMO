"""Native ICC reference from full-precision canonical CMYK (test-only LittleCMS)."""
from pathlib import Path
import ctypes as C, ctypes.util, hashlib, json, os, struct
root=Path(__file__).resolve().parent
library=os.environ.get('LCMS_LIBRARY') or C.util.find_library('lcms2')
if not library: raise RuntimeError('Set LCMS_LIBRARY to the existing test-only LittleCMS shared library.')
lib=C.CDLL(library)
P=C.c_void_p;U=C.c_uint32

def api(name,result,*args):
 f=getattr(lib,name);f.restype=result;f.argtypes=args;return f
profile=root.parent/'IccColorCorpus/littlecms-cmyk-lut.icc'
source=api('cmsOpenProfileFromFile',P,C.c_char_p,C.c_char_p)(str(profile).encode(),b'r')
target=api('cmsCreate_sRGBProfile',P)()
transform=api('cmsCreateTransform',P,P,U,P,U,U,U)(source,(1<<22)|(6<<16)|(4<<3),target,(4<<16)|(3<<3)|1,1,0x0100)
absolute=api('cmsCreateTransform',P,P,U,P,U,U,U)(source,(1<<22)|(6<<16)|(4<<3),target,(4<<16)|(3<<3)|1,3,0x0100)
assert source and target and transform and absolute
run=api('cmsDoTransform',None,P,P,P,U)
for p in sorted(root.glob('*.jpg')):
 bits=int(p.name.split('-')[0][1:]);maximum=(1<<bits)-1
 raw=Path(str(p)+'.nearest.cmyk16').read_bytes();words=struct.unpack('<'+'H'*(len(raw)//2),raw)
 values=(C.c_double*len(words))(*((maximum-v)*100/maximum for v in words));result=(C.c_ubyte*(35*19*3))()
 run(transform,values,result,35*19);Path(str(p)+'.srgb').write_bytes(bytes(result))
 if bits==8:
  for inverted in (False,True):
   pdf_values=(C.c_double*len(words))(*((maximum-v if inverted else v)*100/maximum for v in words))
   run(absolute,pdf_values,result,35*19)
   Path(str(p)+('.pdf-inverted.srgb' if inverted else '.pdf-normal.srgb')).write_bytes(bytes(result))
api('cmsDeleteTransform',None,P)(absolute);api('cmsDeleteTransform',None,P)(transform);api('cmsCloseProfile',C.c_int,P)(source);api('cmsCloseProfile',C.c_int,P)(target)
files=sorted(root.glob('*.srgb'))
(root/'icc-reference.json').write_text(json.dumps({'littleCmsVersion':api('cmsGetEncodedCMMversion',C.c_int)(),'profileSHA256':hashlib.sha256(profile.read_bytes()).hexdigest(),'intent':{'image':'relative colorimetric','pdf':'absolute colorimetric'},'input':'native CMYK words with nearest chroma; image and pdf-inverted references invert all four Adobe channels; pdf-normal preserves raw polarity; double percentages','sha256':{p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in files}},indent=2)+'\n')
print(len(files),'native ICC references')

artifacts=sorted(p for p in root.iterdir() if p.suffix in ('.jpg','.cmyk16','.srgb'))
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n' for p in artifacts))
