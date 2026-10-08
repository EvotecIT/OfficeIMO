"""Native arithmetic alpha TIFFs at additional precisions; test-only tools.
Usage: generate.py <prepared-jpeg> <wrap> <decode> <scratch-directory>
"""
from pathlib import Path
import csv,ctypes as C,ctypes.util,hashlib,json,os,shutil,struct,subprocess,sys
root=Path(__file__).resolve().parent;scratch=Path(sys.argv[4]).resolve();rows=[]
precisions='2,3,4,5,6,7,9,10,11,13,14,15'
for kind,script,args in [('full','TiffJpegArithmeticLosslessColor',[sys.argv[1],sys.argv[2]]),('chroma','TiffJpegArithmeticLosslessChroma',sys.argv[1:4])]:
 output=scratch/kind
 subprocess.run([sys.executable,str(root.parent/script/'generate.py'),*args,str(scratch/(kind+'-scratch')),precisions,str(output),'alpha'],check=True)
 for row in csv.DictReader((output/'manifest.csv').open()):
  name=row['file'];photo=int(row['photometric']);bits=int(row['bits']);alpha=int(row['extra'])
  for suffix in ('','.rgba'):shutil.copy2(output/(name+suffix),root/(name+suffix))
  rows.append([name,photo,bits,f"be{row['bigEndian']}-t{row['tiled']}-pl{row['planar']}",alpha,35,19,kind,row['predictor'],row['point']])
library=os.environ.get('LCMS_LIBRARY') or ctypes.util.find_library('lcms2')
if not library:raise RuntimeError('Set LCMS_LIBRARY to the test-only LittleCMS library.')
lib=C.CDLL(library);P=C.c_void_p;U=C.c_uint32

def api(name,result,*args):
 f=getattr(lib,name);f.restype=result;f.argtypes=args;return f
run=api('cmsDoTransform',None,P,P,P,U);target=api('cmsCreate_sRGBProfile',P)();transforms={};profiles={}
for photo,profile_name in [(2,'icc-dci-p3-matrix.icc'),(5,'littlecms-cmyk-lut.icc')]:
 profile=root.parent/'IccColorCorpus'/profile_name;handle=api('cmsOpenProfileFromFile',P,C.c_char_p,C.c_char_p)(str(profile).encode(),b'r');channels=4 if photo==5 else 3
 transform=api('cmsCreateTransform',P,P,U,P,U,U,U)(handle,(1<<22)|((6 if photo==5 else 4)<<16)|(channels<<3),target,(4<<16)|(3<<3)|1,1,0x0100 if photo==5 else 0x0140)
 assert handle and transform;transforms[photo]=(handle,transform);profiles[profile_name]=hashlib.sha256(profile.read_bytes()).hexdigest()
for name,photo,bits,layout,alpha,width,height,kind,predictor,point in rows:
 if photo not in(2,5,6):continue
 maximum=(1<<bits)-1;middle=1<<(bits-1);count=width*height
 if kind=='chroma':
  raw=(scratch/kind/(name+'.rgb-f64')).read_bytes();values=struct.unpack('<'+'d'*(len(raw)//8),raw)
 else:
  raw=(scratch/kind/(name+'.raw')).read_bytes();words=struct.unpack('<'+'H'*(len(raw)//2),raw);channels=4 if photo==5 else 3;values=[]
  for at in range(0,len(words),channels+1):
   color=list(words[at:at+channels]);a=words[at+channels]
   if photo==6:
    y,cb,cr=color;cb-=middle;cr-=middle;r=y+1.402*cr;b=y+1.772*cb;color=[r,(y-.299*r-.114*b)/.587,b]
   values.extend((0 if alpha==1 and a==0 else max(0,min(1,c/(a if alpha==1 else maximum))))*(100 if photo==5 else 1)for c in color)
 inp=(C.c_double*len(values))(*values);out=(C.c_ubyte*(count*3))();run(transforms[5 if photo==5 else 2][1],inp,out,count)
 device=(root/(name+'.rgba')).read_bytes();(root/(name+'.icc-rgba')).write_bytes(bytes(v for i in range(count)for v in(*out[i*3:i*3+3],device[i*4+3])))
with(root/'manifest.csv').open('w')as f:
 w=csv.writer(f,lineterminator='\n');w.writerow(['name','photometric','bits','layout','alphaKind','width','height','sampling','predictor','point']);w.writerows(rows)
(root/'icc-reference.json').write_text(json.dumps(dict(littleCmsVersion=api('cmsGetEncodedCMMversion',C.c_int)(),profiles=profiles,intent='relative colorimetric',input='native color and alpha; fractional YCbCr interpolation and RGB conversion retained before unassociation and ICC'),indent=2)+'\n')
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n'for p in sorted(root.glob('*.tif*'))))
for handle,transform in transforms.values():api('cmsDeleteTransform',None,P)(transform);api('cmsCloseProfile',C.c_int,P)(handle)
api('cmsCloseProfile',C.c_int,P)(target)
print(len(rows),'native arithmetic TIFFs;',len(list(root.glob('*.icc-rgba'))),'ICC references')
