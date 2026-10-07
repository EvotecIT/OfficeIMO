"""Independent twelve-bit arithmetic TIFF samples and declared color/alpha projections."""
from pathlib import Path
from PIL import Image
import os,sys,subprocess,struct,math,hashlib
root=Path(__file__).resolve().parent;environment={**os.environ,'TIFF_JPEG_ARITHMETIC':'1','TIFF_JPEG_PRECISION':'12'}
rows=['file,photometric,bigEndian,tiled,shared,planar,subsampling,extra']
maximum=4095;midpoint=2048
for photo in (0,1,2,5,6):
 for big in (0,1):
  for tile in (0,1):
   for shared in (0,1):
    for planar in (1,2):
     for sub in ((1,2) if photo==6 else (1,)):
      for extra in (-1,0,1,2):
       name=f'p{photo}-be{big}-t{tile}-q{shared}-pl{planar}-s{sub}-e{extra}.tif';file=root/name
       environment.pop('TIFF_JPEG_OPAQUE',None)
       if extra<0:environment['TIFF_JPEG_OPAQUE']='1'
       subprocess.run([sys.argv[1],str(file),str(photo),str(big),str(tile),str(shared),str(planar),str(sub),str(extra),str(int(extra>0))],check=True,env=environment)
       base=4 if photo==5 else 3 if photo in (2,6) else 1;n=base+(extra>=0)
       if planar==2:
        data=Path(str(file)+'.planes').read_bytes();offset=0;planes=[[0]*(35*19)for _ in range(n)]
        while offset<len(data):
         plane,x,y,w,h=struct.unpack_from('<5I',data,offset);offset+=20
         words=struct.unpack_from('<'+'H'*(w*h),data,offset);offset+=w*h*2
         im=Image.new('F',(w,h));im.putdata(words)
         if photo==6 and plane in (1,2):
          w=min(w,(35-x+sub-1)//sub);h=min(h,(19-y+sub-1)//sub);im=im.crop((0,0,w,h)).resize((w*sub,h*sub),Image.Resampling.BILINEAR)
         for yy in range(min(im.height,19-y)):
          for xx in range(min(im.width,35-x)):planes[plane][(y+yy)*35+x+xx]=int(math.floor(im.getpixel((xx,yy))+.5))
        words=[v for pixel in zip(*planes)for v in pixel];Path(str(file)+'.raw').write_bytes(struct.pack('<'+'H'*len(words),*words))
       else:
        raw=Path(str(file)+'.raw').read_bytes();words=struct.unpack('<'+'H'*(len(raw)//2),raw)
        assert Path(str(file)+'.planes').stat().st_size==0;Path(str(file)+'.planes').unlink()
       rgba=bytearray();q=lambda v:min(255,max(0,math.floor(v*255+.5)))
       for i in range(35*19):
        values=list(words[i*n:i*n+base]);a=words[i*n+base] if extra>0 else maximum
        if photo==6:
         y=values[0];cb=values[1]-midpoint;cr=values[2]-midpoint;r=y+cr*1.402;b=y+cb*1.772;g=(y-.299*r-.114*b)/.587
         values=[min(maximum,max(0,round(v)))for v in (r,g,b)]
        color=[min(1,v/a)if a else 0 for v in values]if extra==1 else [v/maximum for v in values]
        if photo==5:color=[255-min(255,q(color[c])+q(color[3]))for c in range(3)]
        elif photo in (0,1):color=[q(1-color[0]if photo==0 else color[0])]*3
        else:color=[q(v)for v in color]
        rgba.extend(color+[q(a/maximum)])
       Path(str(file)+'.rgba').write_bytes(rgba);rows.append(f'{name},{photo},{big},{tile},{shared},{planar},{sub},{extra}')
(root/'manifest.csv').write_text('\n'.join(rows)+'\n')
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n'for p in sorted(root.glob('*.tif*'))))
print(len(rows)-1,'twelve-bit TIFFs')
