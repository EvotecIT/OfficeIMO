from pathlib import Path
from PIL import Image
import hashlib,struct,subprocess,sys
low='--low-alpha' in sys.argv[2:]
root=Path(__file__).resolve().parent
if low:root=root.parent/'TiffJpegLowAlpha'
root.mkdir(exist_ok=True)
exe=Path(sys.argv[1]).resolve();rows=['file,photometric,bigEndian,tiled,shared,planar,subsampling,extra']
for photo in (0,1,2,5,6):
 for big in (0,1):
  for tile in (0,1):
   for shared in (0,1):
    for planar in (1,2):
     for sub in ((1,2) if photo==6 else (1,)):
      for extra in ((1,2) if low else (0,1,2)):
       name=f'p{photo}-be{big}-t{tile}-q{shared}-pl{planar}-s{sub}-e{extra}.tif';file=root/name
       subprocess.run([str(exe),str(file),str(photo),str(big),str(tile),str(shared),str(planar),str(sub),str(extra),str(int(low))],check=True)
       base=4 if photo==5 else 3 if photo in (2,6) else 1;n=base+1
       if planar==2:
        data=file.with_suffix('.tif.planes').read_bytes();offset=0;planes=[Image.new('L',(35,19)) for _ in range(n)]
        while offset<len(data):
         plane,x,y,w,h=struct.unpack_from('<5I',data,offset);offset+=20;im=Image.frombytes('L',(w,h),data[offset:offset+w*h]);offset+=w*h
         if photo==6 and plane in (1,2):
          w=min(w,(35-x+sub-1)//sub);h=min(h,(19-y+sub-1)//sub);im=im.crop((0,0,w,h)).resize((w*sub,h*sub),Image.Resampling.BILINEAR)
         planes[plane].paste(im.crop((0,0,min(im.width,35-x),min(im.height,19-y))),(x,y))
        raw=bytes(c for pix in zip(*(p.tobytes() for p in planes)) for c in pix);file.with_suffix('.tif.raw').write_bytes(raw)
       raw=file.with_suffix('.tif.raw').read_bytes();rgba=bytearray()
       if photo==6:
        ycc=bytes(raw[i+c] for i in range(0,len(raw),n) for c in range(3));rgb=Image.frombytes('YCbCr',(35,19),ycc).convert('RGB').tobytes()
       for i in range(35*19):
        values=list(raw[i*n:i*n+base]);alpha=raw[i*n+base] if extra else 255
        if photo==6:values=list(rgb[i*3:i*3+3])
        if extra==1:values=[min(255,round(v*255/alpha)) if alpha else 0 for v in values]
        if photo==5:values=[255-min(255,values[c]+values[3]) for c in range(3)]
        elif photo in (0,1):values=[255-values[0] if photo==0 else values[0]]*3
        rgba.extend(values+[alpha])
       file.with_suffix('.tif.rgba').write_bytes(rgba);rows.append(f'{name},{photo},{big},{tile},{shared},{planar},{sub},{extra}')
(root/'manifest.csv').write_text('\n'.join(rows)+'\n')
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n' for p in sorted(root.glob('*.tif*'))))
print(len(rows)-1,'alpha fixtures')
