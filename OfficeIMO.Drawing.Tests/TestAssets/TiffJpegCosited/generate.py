from pathlib import Path
from PIL import Image
import hashlib,struct,subprocess,sys
root=Path(__file__).resolve().parent;exe=Path(sys.argv[1]).resolve()
rows=['file,bigEndian,tiled,tables,horizontal,vertical,planar,width,height,positioning']
for h,v in [(1,1),(2,1),(2,2),(4,1),(4,2),(4,4)]:
 for planar in ((2,) if h*v>8 else (1,2)):
  for big in (0,1):
   for tiled in (0,1):
    for tables in (0,3):
     width,height=(68,36) if tables else (67,35)
     for positioning in ((2,1) if h in (2,4) and v==2 and planar==1 and big==0 and tiled==1 and tables==3 else (2,)):
      name=f'h{h}-v{v}-pl{planar}-be{big}-t{tiled}-q{tables}'+('-centered' if positioning==1 else '')+'.tif';file=root/name
      subprocess.run([str(exe),str(file),str(big),str(tiled),str(tables),str(h),str(v),str(planar),str(positioning)],check=True)
      data=file.with_suffix('.tif.planes').read_bytes();offset=0;planes=[Image.new('L',(width,height)) for _ in range(3)]
      while offset<len(data):
       plane,x,y,w,hh=struct.unpack_from('<5I',data,offset);offset+=20
       image=Image.frombytes('L',(w,hh),data[offset:offset+w*hh]);offset+=w*hh
       if plane:
        w=min(w,(width-x+h-1)//h);hh=min(hh,(height-y+v-1)//v)
        image=image.crop((0,0,w,hh))
        # Extend the last sample before affine interpolation so boundaries clamp.
        padded=Image.new('L',(w+1,hh+1));padded.paste(image)
        padded.paste(image.crop((w-1,0,w,hh)),(w,0));padded.paste(image.crop((0,hh-1,w,hh)),(0,hh));padded.putpixel((w,hh),image.getpixel((w-1,hh-1)))
        image=padded.transform((w*h,hh*v),Image.Transform.AFFINE,(1/h,0,(.5-.5/h if positioning==2 else 0),0,1/v,(.5-.5/v if positioning==2 else 0)),Image.Resampling.BILINEAR)
       planes[plane].paste(image.crop((0,0,min(image.width,width-x),min(image.height,height-y))),(x,y))
      (root/(name+'.rgb')).write_bytes(Image.merge('YCbCr',planes).convert('RGB').tobytes())
      rows.append(f'{name},{big},{tiled},{tables},{h},{v},{planar},{width},{height},{positioning}')
(root/'manifest.csv').write_text('\n'.join(rows)+'\n')
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n' for p in sorted(root.glob('*.tif*'))))
print(len(rows)-1,'independent cosited fixtures')
