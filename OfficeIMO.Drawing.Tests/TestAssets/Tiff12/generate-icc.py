"""Independent twelve-bit CMYK to sRGB references; test-only LittleCMS C API.

Set LCMS_LIBRARY if discovery fails. Input doubles retain all twelve-bit samples.
"""
from pathlib import Path
import ctypes as C, ctypes.util, os, struct, json, hashlib
root=Path(__file__).resolve().parent
lib=C.CDLL(os.environ.get('LCMS_LIBRARY') or C.util.find_library('lcms2'))
P=C.c_void_p;U=C.c_uint32

def api(name,result,*args):
 f=getattr(lib,name);f.restype=result;f.argtypes=args;return f
open_profile=api('cmsOpenProfileFromFile',P,C.c_char_p,C.c_char_p)
srgb=api('cmsCreate_sRGBProfile',P)()
source=open_profile(str(root.parent/'IccColorCorpus/littlecms-cmyk-lut.icc').encode(),b'r')
# TYPE_CMYK_DBL, TYPE_RGB_8, relative colorimetric, no optimization.
transform=api('cmsCreateTransform',P,P,U,P,U,U,U)(source,(1<<22)|(6<<16)|(4<<3),srgb,(4<<16)|(3<<3)|1,1,0x0100)
assert source and srgb and transform
run=api('cmsDoTransform',None,P,P,P,U)
for file in sorted(root.glob('k6-*.tif')):
 raw=Path(str(file)+'.raw').read_bytes();words=struct.unpack('<'+'H'*(len(raw)//2),raw)
 values=(C.c_double*len(words))(*(v*100/4095 for v in words));result=(C.c_ubyte*(35*19*3))()
 run(transform,values,result,35*19);Path(str(file)+'.srgb').write_bytes(bytes(result))
api('cmsDeleteTransform',None,P)(transform);api('cmsCloseProfile',C.c_int,P)(source);api('cmsCloseProfile',C.c_int,P)(srgb)
files=sorted(root.glob('*.srgb'))
(root/'icc-reference.json').write_text(json.dumps(dict(littleCmsVersion=api('cmsGetEncodedCMMversion',C.c_int)(),intent='relative colorimetric',input='CMYK double percentage',sha256={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in files}),indent=2)+'\n')
print(len(files),'native ICC references')
