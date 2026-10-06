"""T.81-authored streams, libjpeg-turbo word references, Pillow interpolation."""
from pathlib import Path
import csv,hashlib,math,struct,subprocess,sys
precision=int(sys.argv[2]) if len(sys.argv)>2 else 16
assert 2<=precision<=16
fractional=len(sys.argv)>4 and sys.argv[4]=='fractional'
alpha_kind=int(sys.argv[5]) if len(sys.argv)>5 else 0
assert alpha_kind in (0,1,2) and (not alpha_kind or fractional)
channels=4 if alpha_kind else 3
maximum=(1<<precision)-1;midpoint=1<<(precision-1)
root=Path(sys.argv[3]).resolve() if len(sys.argv)>3 else Path(__file__).resolve().parent
root.mkdir(parents=True,exist_ok=True)
sys.path.insert(0,str(Path(__file__).resolve().parent.parent))
from tiff_chroma_reference import interpolate

def alpha(x,y):
 return [0,1,min(128,maximum),maximum//2,maximum][(x+2*y)%5]

def value(c,x,y,h=1,v=1):
 if c==3:return alpha(x,y)
 raw=(10000+x*1777+y*3331+(x*y)*71) % (maximum+1) if c==0 else (24500+x*4441-y*2333+(x*y)*137+c*5101)%(maximum+1)
 if alpha_kind!=1:return raw
 # Associate luma at full resolution. Chroma represents a block-constant
 # source chroma field associated per pixel, then averaged over its group.
 if c==0:return int(math.floor(raw*alpha(x,y)/maximum+.5))
 coverage=sum(alpha(x*h+xx,y*v+yy)for yy in range(v)for xx in range(h))/(h*v)
 return int(math.floor(midpoint+(raw-midpoint)*coverage/maximum+.5))

def jpeg(planes,w,height,h,v,separate):
 b=bytearray(b'\xff\xd8');n=len(planes)
 def segment(m,p):b.extend(bytes([255,m])+(len(p)+2).to_bytes(2,'big')+p)
 factors=[(h,v) if c in (0,3) else (1,1) for c in range(n)]
 segment(195,bytes([precision])+struct.pack('>HHB',height,w,n)+b''.join(bytes([i+1,(a<<4)|z,0]) for i,(a,z) in enumerate(factors)))
 segment(196,bytes([0]+[0]*4+[17]+[0]*11+list(range(17))))
 def scan(ids):
  segment(218,bytes([len(ids)])+b''.join(bytes([i+1,0])for i in ids)+bytes([1,0,0]));bits='';single=len(ids)==1
  cols=len(planes[ids[0]][0]) if single else (w+h-1)//h;rows=len(planes[ids[0]]) if single else (height+v-1)//v
  for my in range(rows):
   for mx in range(cols):
    for c in ids:
     ah,av=(1,1) if single else factors[c];p=planes[c]
     for yy in range(av):
      for xx in range(ah):
       x=mx*ah+xx;y=my*av+yy
       sample=lambda x,y:p[min(y,len(p)-1)][min(x,len(p[0])-1)]
       current=sample(x,y);pred=sample(x-1,y) if x else sample(0,y-1) if y else midpoint
       diff=(current-pred+32768)%65536-32768 if precision==16 else current-pred;cat=abs(diff).bit_length();bits+=f'{cat:05b}'
       if 0<cat<16:bits+=format(diff if diff>0 else diff+(1<<cat)-1,f'0{cat}b')
  bits+='1'*((-len(bits))%8)
  for i in range(0,len(bits),8):
   q=int(bits[i:i+8],2);b.append(q)
   if q==255:b.append(0)
 if separate:
  for c in range(n):scan([c])
 else:scan(list(range(n)))
 b.extend(b'\xff\xd9');return bytes(b)

def decode(encoded,stem):
 path=root/(stem+'.jpg');path.write_bytes(encoded);raw=root/(stem+'.raw')
 subprocess.run([sys.argv[1],str(path),str(raw),'0'],check=True)
 result=list(struct.unpack('<'+'H'*(raw.stat().st_size//2),raw.read_bytes()))
 path.unlink();raw.unlink();return result


def tiff(width,height,h,v,layout,payloads):
 tile=layout in (2,3,6,7);planar=2 if layout>=4 else 1;position=2 if layout in (1,3,5,6) else 1;endian='>' if layout%2 else '<';sw=16 if tile else width;sh=16 if tile else 8
 entries=[(256,4,1,width),(257,4,1,height),(258,3,channels,[precision]*channels),(259,3,1,7),(262,3,1,6),(277,3,1,channels),(284,3,1,planar),(530,3,2,[h,v]),(531,3,1,position)]
 if alpha_kind:entries.append((338,3,1,alpha_kind))
 if fractional:entries.append((532,5,6,[0,1,maximum,1,midpoint,1,maximum,1,midpoint,1,maximum,1]))
 entries+=([(322,4,1,sw),(323,4,1,sh)] if tile else [(278,4,1,sh)])
 offsetTag,countTag=(324,325) if tile else (273,279)
 entries.extend([(offsetTag,4,len(payloads),[0]*len(payloads)),(countTag,4,len(payloads),[len(p) for p in payloads])]);entries.sort();count=len(entries)
 def header(offsets):
  extra=bytearray();directory=bytearray();start=8+2+count*12+4
  for tag,kind,n,values in entries:
   if tag==offsetTag:values=offsets
   if isinstance(values,int):values=[values]
   raw=struct.pack(endian+('H' if kind==3 else 'I')*(n*2 if kind==5 else n),*values)
   if len(raw)<=4:field=raw+bytes(4-len(raw))
   else:field=struct.pack(endian+'I',start+len(extra));extra.extend(raw)
   directory.extend(struct.pack(endian+'HHI',tag,kind,n)+field)
  return (b'MM\0*' if endian=='>' else b'II*\0')+struct.pack(endian+'IH',8,count)+directory+bytes(4)+extra
 head=header([0]*len(payloads));offsets=[];at=len(head)
 for p in payloads:offsets.append(at);at+=len(p)
 return header(offsets)+b''.join(payloads)

rows=[]
for h,v in ([(2,1),(2,2),(4,2),(4,4)] if fractional else [(2,1),(2,2),(4,2)]):
 for width,height in [(5,3),(17,11)]:
  for layout in range(8):
   if (2 if alpha_kind else 1)*h*v+2>10 and layout in (0,3):continue  # An interleaved JPEG MCU has at most ten samples.
   name=(f'a{alpha_kind}-' if alpha_kind else '')+(f'b{precision}-' if fractional else '')+f'h{h}-v{v}-w{width}-h{height}-l{layout}';tile=layout in (2,3,6,7);planar=layout>=4;separate=layout in (1,2);position=2 if layout in (1,3,5,6) else 1;sw=16 if tile else width;sh=16 if tile else 8
   segments=[];reference=[0]*(width*height*channels)
   for top in range(0,height,sh):
    for left in range(0,width,sw):
     dw=sw;dh=sh if tile else min(sh,height-top);vw=min(sw,width-left);vh=min(sh,height-top)
     planes=[[[value(c,left//(h if c in (1,2) else 1)+x,top//(v if c in (1,2) else 1)+y,h,v) for x in range((dw+h-1)//h if c in (1,2) else dw)]for y in range((dh+v-1)//v if c in (1,2) else dh)]for c in range(channels)]
     encoded=[]
     if planar:
      for c,p in enumerate(planes):
       stream=jpeg([p],len(p[0]),len(p),1,1,False);words=decode(stream,name+'-temp');assert words==sum(p,[]);encoded.append(stream)
     else:
      stream=jpeg(planes,dw,dh,h,v,separate);words=decode(stream,name+'-temp');expected=[planes[c][y//(v if c in (1,2) else 1)][x//(h if c in (1,2) else 1)]for y in range(dh)for x in range(dw)for c in range(channels)];assert words==expected,(name,'independent words');encoded=[stream]
      if top==0 and left==0 and not fractional:
       (root/(name+'.jpg')).write_bytes(stream);(root/(name+'.nearest.raw')).write_bytes(struct.pack('<'+'H'*len(words),*words))
       interpolated=[sum(planes[0],[])]+[interpolate(p,dw,dh,h,v,1)for p in planes[1:3]]
       interleaved=[interpolated[c][i]for i in range(dw*dh)for c in range(channels)];(root/(name+'.bilinear.raw')).write_bytes(struct.pack('<'+'H'*len(interleaved),*interleaved))
     segments.append(encoded)
     grids=[sum([row[:vw] for row in planes[0][:vh]],[])]+[interpolate([row[:(vw+h-1)//h]for row in p[:(vh+v-1)//v]],vw,vh,h,v,position,round_samples=False)for p in planes[1:3]]
     if alpha_kind:grids.append(sum([row[:vw] for row in planes[3][:vh]],[]))
     for y in range(vh):
      for x in range(vw):
       for c in range(channels):reference[((top+y)*width+left+x)*channels+c]=grids[c][y*vw+x]
   payloads=[segment[c]for c in range(channels)for segment in segments]if planar else [segment[0]for segment in segments]
   (root/(name+'.tif')).write_bytes(tiff(width,height,h,v,layout,payloads));rgba=bytearray();normalized=[]
   for i in range(width*height):
    y,cb,cr=reference[i*channels:i*channels+3];cb-=midpoint;cr-=midpoint;r=y+1.402*cr;b=y+1.772*cb;g=(y-.299*r-.114*b)/.587
    a=reference[i*channels+3] if alpha_kind else maximum
    rgb=[0 if alpha_kind==1 and a==0 else max(0,min(1,q/(a if alpha_kind==1 else maximum)))for q in(r,g,b)]
    normalized.extend(rgb)
    if alpha_kind:rgba.extend([int(math.floor(q*255+.5)) for q in rgb]+[int(math.floor(a*255/maximum+.5))])
    else:rgba.extend([int(math.floor(min(maximum,max(0,q))*255/maximum+.5))for q in(r,g,b)]+[255])
   if alpha_kind:(root/(name+'.rgb-f64')).write_bytes(struct.pack('<'+'d'*len(normalized),*normalized))
   (root/(name+'.rgba')).write_bytes(rgba);rows.append([name,width,height,h,v,layout,position])
with (root/'manifest.csv').open('w')as f:
 writer=csv.writer(f);writer.writerow(['name','width','height','horizontal','vertical','layout','position']);writer.writerows(rows)
files=sorted(p for p in root.iterdir()if p.suffix in ('.jpg','.raw','.tif','.rgba'))
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n'for p in files));print(len(rows),'TIFF cases; varying samples verified by libjpeg-turbo, interpolation by Pillow')
