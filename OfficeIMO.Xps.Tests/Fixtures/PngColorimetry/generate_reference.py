from pathlib import Path
import argparse, ctypes as C, ctypes.util, struct,zlib,csv
from PIL import Image,ImageCms
import PIL
class xyY(C.Structure): _fields_=[('x',C.c_double),('y',C.c_double),('Y',C.c_double)]
class Triple(C.Structure): _fields_=[('Red',xyY),('Green',xyY),('Blue',xyY)]
parser=argparse.ArgumentParser();parser.add_argument('--lcms-library');args=parser.parse_args()
lib=args.lcms_library or next((Path(PIL.__file__).parent/'.dylibs').glob('*lcms*'),None) or ctypes.util.find_library('lcms2')
if not lib:raise SystemExit('Supply --lcms-library for the test-only LittleCMS runtime.')
cms=C.CDLL(str(lib))
cms.cmsBuildGamma.argtypes=[C.c_void_p,C.c_double];cms.cmsBuildGamma.restype=C.c_void_p
cms.cmsCreateRGBProfile.argtypes=[C.POINTER(xyY),C.POINTER(Triple),C.POINTER(C.c_void_p)];cms.cmsCreateRGBProfile.restype=C.c_void_p
cms.cmsSaveProfileToMem.argtypes=[C.c_void_p,C.c_void_p,C.POINTER(C.c_uint32)];cms.cmsSaveProfileToMem.restype=C.c_int
cms.cmsCloseProfile.argtypes=[C.c_void_p];cms.cmsFreeToneCurve.argtypes=[C.c_void_p]
def profile(coords,gamma):
 w=xyY(coords[0]/1e5,coords[1]/1e5,1);p=Triple(*[xyY(coords[i]/1e5,coords[i+1]/1e5,1) for i in (2,4,6)])
 curve=cms.cmsBuildGamma(None,1e5/gamma); curves=(C.c_void_p*3)(curve,curve,curve);h=cms.cmsCreateRGBProfile(C.byref(w),C.byref(p),curves);assert h
 length=C.c_uint32();assert cms.cmsSaveProfileToMem(h,None,C.byref(length));data=C.create_string_buffer(length.value);assert cms.cmsSaveProfileToMem(h,data,C.byref(length));cms.cmsCloseProfile(h);cms.cmsFreeToneCurve(curve)
 return ImageCms.ImageCmsProfile(__import__('io').BytesIO(data.raw))
srgb=[31270,32900,64000,33000,30000,60000,15000,6000];p3=[31270,32900,68000,32000,26500,69000,15000,6000];adobe=[31270,32900,64000,33000,21000,71000,15000,6000];d50=[34570,35850,64000,33000,30000,60000,15000,6000]
cases=[('gamma-linear',100000,None,'RGBA',(128,64,32,128)),('gamma-half',50000,None,'RGBA',(64,128,192,255)),('gamma-standard-power',45455,srgb,'RGBA',(128,64,32,255)),('p3',45455,p3,'RGBA',(128,64,32,255)),('adobe',55556,adobe,'RGBA',(32,128,192,128)),('d50',50000,d50,'RGBA',(128,64,32,255)),('gray',100000,srgb,'LA',(128,128)),('indexed',100000,p3,'P',(128,64,32,128))]
out=Path(__file__).parent;out.mkdir(exist_ok=True)
def chunk(kind,data):return struct.pack('>I',len(data))+kind+data+struct.pack('>I',zlib.crc32(kind+data))
rows=[]
for name,gamma,chrm,mode,sample in cases:
 image=Image.new(mode,(16,16),sample if mode!='P' else 0)
 if mode=='P':image.putpalette(list(sample[:3])+[0]*765);image.info['transparency']=bytes([sample[3]])
 path=out/(name+'.png');image.save(path);data=path.read_bytes();added=chunk(b'gAMA',struct.pack('>I',gamma))
 if chrm:added+=chunk(b'cHRM',struct.pack('>8I',*chrm))
 path.write_bytes(data[:33]+added+data[33:])
 rgb=image.convert('RGB');trans=ImageCms.buildTransformFromOpenProfiles(profile(chrm or srgb,gamma),ImageCms.createProfile('sRGB'),'RGB','RGB',renderingIntent=1,flags=0)
 expected=ImageCms.applyTransform(rgb,trans).getpixel((8,8));alpha=image.convert('RGBA').getpixel((8,8))[3]
 rows.append([path.name,gamma,mode,*rgb.getpixel((8,8)),alpha,*expected]);print(rows[-1])
# Canonical sRGB cICP overrides gAMA in PNG 3. No synthesized ICC expectation.
data=(out/'gamma-linear.png').read_bytes()
(out/'canonical-cicp.png').write_bytes(data[:33]+chunk(b'cICP',bytes([1,13,0,1]))+data[33:])
rows.append(['canonical-cicp.png',100000,'RGBA',128,64,32,128,'','',''])
with (out/'expected.csv').open('w',newline='') as f:
 writer=csv.writer(f,lineterminator='\n');writer.writerow(['file','gamma','mode','r','g','b','a','calibrated_r','calibrated_g','calibrated_b']);writer.writerows(rows)
print('Reference engine:',PIL.__version__,ImageCms.core.littlecms_version)
