"""Regenerate multichannel test profiles and swatches using LittleCMS.

Set LCMS_LIBRARY when the native library is not discoverable. Test-only tool;
not invoked by builds, tests, or runtime code. Reference version: LittleCMS 2.19.
"""
from pathlib import Path
import ctypes as C,itertools,csv,os
from ctypes.util import find_library
out=Path(__file__).resolve().parent
l=C.CDLL(os.environ.get('LCMS_LIBRARY') or find_library('lcms2'));P=C.c_void_p;U=C.c_uint32;D=C.c_double;S=C.c_char_p
def api(n,r,*a):
 f=getattr(l,n);f.restype=r;f.argtypes=a;return f
sig=lambda x:int.from_bytes(x.encode(),'big')
create=api('cmsCreateProfilePlaceholder',P,P);version=api('cmsSetProfileVersion',None,P,D)
cls=api('cmsSetDeviceClass',None,P,U);space=api('cmsSetColorSpace',None,P,U);pcs=api('cmsSetPCS',None,P,U)
write=api('cmsWriteTag',C.c_int,P,U,P);save=api('cmsSaveProfileToFile',C.c_int,P,S)
pipe=api('cmsPipelineAlloc',P,P,U,U);insert=api('cmsPipelineInsertStage',C.c_int,P,C.c_int,P)
curves=api('cmsStageAllocToneCurves',P,P,U,P)
clut=api('cmsStageAllocCLut16bitGranular',P,P,C.POINTER(U),U,U,C.POINTER(C.c_uint16))
freepipe=api('cmsPipelineFree',None,P);close=api('cmsCloseProfile',C.c_int,P)
srgb=api('cmsCreate_sRGBProfile',P)();openfile=api('cmsOpenProfileFromFile',P,S,S)
transform=api('cmsCreateTransform',P,P,U,P,U,U,U);run=api('cmsDoTransform',None,P,P,P,U);delete=api('cmsDeleteTransform',None,P)
rows=[]
for n,kind in itertools.product(range(3,9),['lut8','lut16','mab']):
 p=create(None);version(p,2.1 if kind!='mab' else 4.3);cls(p,sig('scnr'));space(p,sig(f'{n}CLR'));pcs(p,sig('Lab '))
 white=(D*3)(.9642,1,.8249);assert write(p,sig('wtpt'),white)
 grids=[3]*n if kind!='mab' else [2+i%2 for i in range(n)]
 values=[]
 for coords in itertools.product(*(range(g) for g in grids)):
  x=[v/(g-1) for v,g in zip(coords,grids)];mean=sum((i+1)*v for i,v in enumerate(x))/(n*(n+1)/2)
  values += [round((.25+.5*mean)*65535), round((.5+.1*(x[0]-x[-1]))*65535),round((.5+.1*(x[1]-mean))*65535)]
 tabcurve=api('cmsBuildTabulatedToneCurve16',P,P,U,C.POINTER(C.c_uint16));freecurve=api('cmsFreeToneCurve',None,P)
 entries=256 if kind=='lut8' else 4096
 inputcurves=[tabcurve(None,entries,(C.c_uint16*entries)(*(round((v/(entries-1))**(1.1+i*.15)*65535) for v in range(entries)))) for i in range(n)]
 pl=pipe(None,n,3);assert insert(pl,1,curves(None,n,(P*n)(*inputcurves)));
 for curve in inputcurves:freecurve(curve)
 assert insert(pl,1,clut(None,(U*n)(*grids),n,3,(C.c_uint16*len(values))(*values)));assert insert(pl,1,curves(None,3,None));api('cmsPipelineSetSaveAs8bitsFlag',C.c_int,P,C.c_int)(pl,kind=='lut8');assert write(p,sig('A2B0'),pl)
 name=f'littlecms-{n}clr-{kind}.icc';path=out/name;assert save(p,str(path).encode());freepipe(pl);close(p)
 p=openfile(str(path).encode(),b'r');t=transform(p,((n+14)<<16)|(n<<3)|2,srgb,(4<<16)|(3<<3)|1,1,0x0100);assert t
 samples=[[0]*n,[65535]*n,[round(65535*(i+1)/(n+1)) for i in range(n)]]+[[65535 if i==c else 0 for i in range(n)] for c in range(n)]
 for sample in samples:
  result=(C.c_ubyte*3)();run(t,(C.c_uint16*n)(*sample),result,1);rows.append([name,':'.join(map(str,sample)),':'.join(map(str,result))])
 delete(t);close(p)
with (out/'reference-nchannel.csv').open('w') as f:
 w=csv.writer(f,lineterminator='\n');w.writerow(['profile','input16','expected']);w.writerows(rows)
close(srgb);print(len(rows),'native reference swatches')
