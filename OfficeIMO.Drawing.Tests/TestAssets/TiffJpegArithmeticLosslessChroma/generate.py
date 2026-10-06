"""Native subsampled SOF11 segments; independent Pillow TIFF chroma projection.
Usage: generate.py /prepared/jpeg /compiled/wrap /compiled/decode /task/scratch
"""
from pathlib import Path
import csv,hashlib,os,struct,subprocess,sys
sys.path.insert(0,str(Path(__file__).resolve().parent.parent))
from tiff_chroma_reference import interpolate
root=Path(__file__).resolve().parent
oracle,wrapper,turbo=map(lambda p:str(Path(p).resolve()),sys.argv[1:4])
scratch=Path(sys.argv[4]).resolve();scratch.mkdir(parents=True,exist_ok=True)
rows=[];segments_total=0

def run(args,env=None):
 r=subprocess.run(args,env=env,capture_output=True,text=True)
 if r.returncode or r.stderr:raise RuntimeError(str(args)+'\n'+r.stderr)

def encode(values,w,height,n,bits,h,v,predictor,point,stem):
 global segments_total
 maximum=(1<<bits)-1;size=1 if bits==8 else 2
 source=scratch/(stem+'.pnm');jpg=scratch/(stem+'.jpg')
 source.write_bytes(f'P{5 if n==1 else 6}\n{w} {height}\n{maximum}\n'.encode()+b''.join(value.to_bytes(size,'big')for value in values))
 env=dict(os.environ,OFFICEIMO_TEST_DEPTH=str(n),OFFICEIMO_TEST_PREDICTOR=str(predictor),OFFICEIMO_TEST_POINT=str(point))
 sampling=','.join(f'{h}x{v}'if c in (1,2)else'1x1'for c in range(n))
 options=['-p','-c','-z',str(((w+h-1)//h)*2),'-s',sampling]
 run([oracle,'-a',*options,str(source),str(jpg)],env)
 # A companion Huffman stream uses identical source samples/prediction. Decode
 # it with libjpeg-turbo; the producer's decoder corrupts odd-height edge rows.
 companion=scratch/(stem+'-huffman.jpg');raw=scratch/(stem+'.raw')
 run([oracle,*options,str(source),str(companion)],env)
 run([turbo,str(companion),str(raw),'0'])
 data=raw.read_bytes();assert len(data)==w*height*n*2
 words=struct.unpack('<'+'H'*(len(data)//2),data)
 planes=[]
 for c in range(n):
  hs=h if c in (1,2)else 1;vs=v if c in (1,2)else 1
  plane=[[words[(y*w+x)*n+c]for x in range(0,w,hs)]for y in range(0,height,vs)]
  assert all(words[(y*w+x)*n+c]==plane[y//vs][x//hs]for y in range(height)for x in range(w)),(stem,c,'nearest sampling')
  planes.append(plane)
 segments_total+=1
 return jpg.read_bytes(),planes

for bits in (8,12,16):
 maximum=(1<<bits)-1;middle=1<<(bits-1)
 for h,v in ((2,1),(2,2),(4,2)):
  for position in (1,2):
   for extra in (-1,1,2):
    n=3+(extra>=0)
    for big,tiled,planar in ((0,0,1),(1,1,1),(1,0,2),(0,1,2)):
     if h==4 and n==4 and planar==1:continue  # 18 sampling units exceed the interleaved-scan limit.
     name=f'b{bits}-h{h}-v{v}-pos{position}-be{big}-t{tiled}-pl{planar}-e{extra}.tif'
     sw=16 if tiled else 35;sh=16 if tiled else 8 if v>1 else 7
     predictor=1+len(rows)%7;point=2 if len(rows)%3==1 else 0
     references=[[0]*(35*19)for _ in range(n)];blocks=[]
     for top in range(0,19,sh):
      for left in range(0,35,sw):
       dh=sh if tiled else min(sh,19-top);vw=min(sw,35-left);vh=min(dh,19-top)
       values=[]
       for y in range(dh):
        for x in range(sw):
         xx=min(34,left+x);yy=min(18,top+y)
         a=[0,1,2,4,16,64,maximum//2,maximum-1,maximum][(xx//4+yy//4)%9]
         rgb=[((xx*193+yy*791+c*3191)^(xx*yy*53))&maximum for c in range(3)]
         if extra==1:rgb=[value*a/maximum for value in rgb]
         r,g,b=rgb;luma=.299*r+.587*g+.114*b
         samples=[round(luma),round(middle+(b-luma)/1.772),round(middle+(r-luma)/1.402)]
         if extra>=0:samples.append(a)
         values.extend(min(maximum,max(0,value))for value in samples)
       if h==4 and n==4:
        colors=[values[i*n+c]for i in range(sw*dh)for c in range(3)]
        jpeg,planes=encode(colors,sw,dh,3,bits,h,v,predictor,point,'source')
        _,alpha=encode(values[3::4],sw,dh,1,bits,1,1,predictor,point,'alpha')
        planes.extend(alpha)
       else:jpeg,planes=encode(values,sw,dh,n,bits,h,v,predictor,point,'source')
       # Full-resolution luma and alpha provide exact point-transform checks.
       for c in (0,3) if n==4 else (0,):
        assert sum(planes[c],[])==[values[i*n+c]>>point<<point for i in range(sw*dh)],name
       if planar==2:
        encoded=[]
        for c,plane in enumerate(planes):
         data,check=encode(sum(plane,[]),len(plane[0]),len(plane),1,bits,1,1,predictor,0,f'plane{c}')
         assert check==[plane],name;encoded.append(data)
       else:encoded=[jpeg]
       blocks.append(encoded)
       for c,plane in enumerate(planes):
        if c in (1,2):
         visible=[row[:(vw+h-1)//h]for row in plane[:(vh+v-1)//v]]
         grid=interpolate(visible,vw,vh,h,v,position)
        else:grid=sum([row[:vw]for row in plane[:vh]],[])
        for y in range(vh):
         for x in range(vw):references[c][(top+y)*35+left+x]=grid[y*vw+x]
     payloads=[block[c]for c in range(n)for block in blocks]if planar==2 else[block[0]for block in blocks]
     paths=[]
     for i,data in enumerate(payloads):
      path=scratch/f'part-{i}.jpg';path.write_bytes(data);paths.append(str(path))
     env=dict(os.environ,TIFF_SAMPLE_H=str(h),TIFF_SAMPLE_V=str(v),TIFF_SAMPLE_POSITION=str(position))
     run([wrapper,str(root/name),'6',str(bits),str(n),str(planar),str(tiled),str(extra),'wb'if big else'wl',*paths],env)
     rgba=bytearray()
     for i in range(35*19):
      yy,cb,cr=[references[c][i]for c in range(3)];cb-=middle;cr-=middle
      r=yy+cr*1.402;b=yy+cb*1.772;g=(yy-.299*r-.114*b)/.587
      colors=[min(maximum,max(0,round(value)))for value in (r,g,b)];a=references[3][i]if n==4 else maximum
      if extra==1:colors=[min(255,round(value*255/a))if a else 0 for value in colors]
      else:colors=[(value*255+maximum//2)//maximum for value in colors]
      rgba.extend(colors+[(a*255+maximum//2)//maximum])
     Path(str(root/name)+'.rgba').write_bytes(rgba)
     rows.append([name,6,bits,h,v,position,big,tiled,planar,extra,predictor,point])
with(root/'manifest.csv').open('w')as f:
 writer=csv.writer(f,lineterminator='\n');writer.writerow(['file','photometric','bits','horizontal','vertical','position','bigEndian','tiled','planar','extra','predictor','point']);writer.writerows(rows)
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n'for p in sorted(root.glob('*.tif*'))))
print(len(rows),'TIFFs;',segments_total,'native JPEG encode/decode operations')
