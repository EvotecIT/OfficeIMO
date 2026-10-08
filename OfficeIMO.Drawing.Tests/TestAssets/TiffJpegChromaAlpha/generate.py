"""Construct subsampled alpha TIFFs, verify native words, and compare ICC with LittleCMS.
Usage: python3 generate.py <compiled ../TiffJpegChroma16/decode.c> <scratch-directory>
"""
from pathlib import Path
import csv,ctypes as C,ctypes.util,hashlib,json,os,shutil,struct,subprocess,sys
root=Path(__file__).resolve().parent;tool=Path(sys.argv[1]).resolve();scratch=Path(sys.argv[2]).resolve()
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
run=api('cmsDoTransform',None,P,P,P,U);rows=[];effect=0
for bits in range(2,17):
 for kind in(1,2):
  output=scratch/f'b{bits}-a{kind}'
  subprocess.run([sys.executable,str(root.parent/'TiffJpegChroma16/generate.py'),str(tool),str(bits),str(output),'fractional',str(kind)],check=True)
  for row in csv.DictReader((output/'manifest.csv').open()):
   name=row['name'];width=int(row['width']);height=int(row['height']);count=width*height
   for suffix in('.tif','.rgba'):shutil.copy2(output/(name+suffix),root/(name+'.tif'+('' if suffix=='.tif' else suffix)))
   raw=(output/(name+'.rgb-f64')).read_bytes();values=struct.unpack('<'+'d'*(len(raw)//8),raw)
   inp=(C.c_double*len(values))(*values);out=(C.c_ubyte*(count*3))();run(transform,inp,out,count)
   device=(output/(name+'.rgba')).read_bytes()
   rgba=bytes(v for i in range(count)for v in(*out[i*3:i*3+3],device[i*4+3]))
   effect=max(effect,max(abs(device[i]-rgba[i])for i in range(len(rgba))))
   (root/(name+'.tif.icc-rgba')).write_bytes(rgba)
   rows.append([name+'.tif',6,bits,row['layout'],kind,width,height,row['horizontal'],row['vertical'],row['position']])
with(root/'manifest.csv').open('w')as f:
 w=csv.writer(f,lineterminator='\n');w.writerow(['name','photometric','bits','layout','alphaKind','width','height','horizontal','vertical','position']);w.writerows(rows)
(root/'icc-reference.json').write_text(json.dumps(dict(littleCmsVersion=api('cmsGetEncodedCMMversion',C.c_int)(),profile=profile.name,profileSHA256=hashlib.sha256(profile.read_bytes()).hexdigest(),intent='relative colorimetric',input='Pillow fractional chroma converted to RGB then unassociated with native alpha before ICC',maximumProfileEffect=effect),indent=2)+'\n')
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n'for p in sorted(root.glob('*.tif*'))))
api('cmsDeleteTransform',None,P)(transform);api('cmsCloseProfile',C.c_int,P)(source);api('cmsCloseProfile',C.c_int,P)(target)
print(len(rows),'TIFFs with device and ICC references; profile effect',effect)
