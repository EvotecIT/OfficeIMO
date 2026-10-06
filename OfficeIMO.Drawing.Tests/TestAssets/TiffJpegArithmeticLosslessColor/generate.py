"""Native SOF11 JPEG segments in LibTIFF strips/tiles; no product codec is used.

Usage: generate.py /path/to/prepared/jpeg /path/to/wrap /task/scratch
Prepare the pinned JPEG oracle with ../JpegArithmeticLossless/prepare_oracle.py.
"""
from pathlib import Path
import csv, hashlib, math, os, struct, subprocess, sys

root=Path(__file__).resolve().parent
oracle,wrapper=map(lambda p:str(Path(p).resolve()),sys.argv[1:3])
scratch=Path(sys.argv[3]).resolve();scratch.mkdir(parents=True,exist_ok=True)
rows=[];segment_count=0

def run(args,env=None):
    result=subprocess.run(args,env=env,capture_output=True,text=True)
    if result.returncode or result.stderr:raise RuntimeError(str(args)+'\n'+result.stderr)

def q(value):return min(255,max(0,math.floor(value*255+.5)))

for bits in (8,12,16):
 maximum=(1<<bits)-1;middle=1<<(bits-1);size=1 if bits==8 else 2
 for photo in (0,1,2,5,6):
  base=4 if photo==5 else 3 if photo in (2,6) else 1
  for extra in (-1,1,2):
   n=base+(extra>=0)
   # Both endian orders, strips/tiles and contiguous/separate samples. The
   # native encoder supports at most four components; do not fake a fifth.
   for big,tiled,planar in ((0,0,1),(1,1,1),(1,0,2),(0,1,2)):
    if n>4 and planar==1:continue
    predictor=1+(len(rows)%7);point=2 if len(rows)%3==1 else 0
    name=f'p{photo}-b{bits}-be{big}-t{tiled}-pl{planar}-e{extra}.tif'
    samples=[]
    for y in range(19):
     for x in range(35):
      alpha=[0,1,2,3,4,8,16,32,64,maximum//2,maximum-1,maximum][(x+y*3)%12]
      values=[((x*193+y*791+c*3191)^(x*y*53))&maximum for c in range(base)]
      if extra==1:
       if photo==6:
        # Premultiply RGB then convert into the TIFF YCbCr coding domain.
        r,g,b=[v*alpha/maximum for v in values]
        yy=.299*r+.587*g+.114*b
        values=[round(yy),round(middle+(b-yy)/1.772),round(middle+(r-yy)/1.402)]
       else:values=[v*alpha//maximum for v in values]
      if extra>=0:values.append(alpha)
      samples.extend(min(maximum,max(0,v))>>point<<point for v in values)
    segments=[];sw=16 if tiled else 35;sh=16 if tiled else 7
    for plane in range(n if planar==2 else 1):
     for top in range(0,19,sh):
      for left in range(0,35,sw):
       height=sh if tiled else min(sh,19-top);depth=1 if planar==2 else n
       words=[]
       for y in range(height):
        for x in range(sw):
         at=(min(18,top+y)*35+min(34,left+x))*n
         words.extend(samples[at+plane:at+plane+1] if planar==2 else samples[at:at+n])
       stem=scratch/'segment';source=stem.with_suffix('.pnm');jpg=stem.with_suffix('.jpg')
       source.write_bytes(f'P{5 if depth==1 else 6}\n{sw} {height}\n{maximum}\n'.encode()+b''.join(v.to_bytes(size,'big')for v in words))
       env=dict(os.environ,OFFICEIMO_TEST_DEPTH=str(depth),OFFICEIMO_TEST_PREDICTOR=str(predictor),OFFICEIMO_TEST_POINT=str(point))
       run([oracle,'-a','-p','-c','-z',str(sw*3),str(source),str(jpg)],env)
       output=stem.with_suffix('.decoded');run([oracle,'-c',str(jpg),str(output)])
       if depth in (1,3):
        data=output.read_bytes().split(b'\n',3)[3]
        actual=[int.from_bytes(data[i:i+size],'big')for i in range(0,len(data),size)]
       else:
        planes=[Path(str(output)+f'_{c}.raw').read_bytes()for c in range(depth)]
        actual=[int.from_bytes(planes[c][i*size:(i+1)*size],'big')for i in range(sw*height)for c in range(depth)]
       assert actual==words,(name,plane,left,top,'native decode mismatch')
       segment=scratch/f'part-{len(segments)}.jpg';segment.write_bytes(jpg.read_bytes());segments.append(str(segment));segment_count+=1
    run([wrapper,str(root/name),str(photo),str(bits),str(n),str(planar),str(tiled),str(extra),'wb'if big else'wl',*segments])
    rgba=bytearray()
    for i in range(35*19):
     values=samples[i*n:i*n+base];a=samples[i*n+base]if extra>0 else maximum
     if photo==6:
      yy,cb,cr=values;cb-=middle;cr-=middle;r=yy+cr*1.402;b=yy+cb*1.772;g=(yy-.299*r-.114*b)/.587
      values=[min(maximum,max(0,round(v)))for v in (r,g,b)]
     colors=[min(1,v/a)if a else 0 for v in values]if extra==1 else[v/maximum for v in values]
     if photo==5:colors=[255-min(255,q(colors[c])+q(colors[3]))for c in range(3)]
     elif photo in (0,1):colors=[q(1-colors[0]if photo==0 else colors[0])]*3
     else:colors=[q(v)for v in colors]
     rgba.extend(colors+[q(a/maximum)])
    Path(str(root/name)+'.rgba').write_bytes(rgba)
    Path(str(root/name)+'.raw').write_bytes(struct.pack('<'+'H'*len(samples),*samples))
    rows.append([name,photo,bits,big,tiled,planar,extra,predictor,point])
with (root/'manifest.csv').open('w')as f:
 writer=csv.writer(f,lineterminator='\n');writer.writerow(['file','photometric','bits','bigEndian','tiled','planar','extra','predictor','point']);writer.writerows(rows)
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n'for p in sorted(root.glob('*.tif*'))))
print(len(rows),'TIFF files;',segment_count,'native SOF11 segments, all exactly decoded')
