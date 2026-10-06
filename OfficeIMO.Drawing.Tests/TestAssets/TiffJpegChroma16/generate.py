"""T.81-authored streams, libjpeg-turbo word references, Pillow interpolation."""
from pathlib import Path
import csv,hashlib,struct,subprocess,sys
root=Path(__file__).resolve().parent
sys.path.insert(0,str(root.parent))
from tiff_chroma_reference import interpolate

def value(c,x,y):
 return (10000+x*1777+y*3331+(x*y)*71) % 65536 if c==0 else (24500+x*4441-y*2333+(x*y)*137+c*5101)%65536

def jpeg(planes,w,height,h,v,separate):
 b=bytearray(b'\xff\xd8');n=len(planes)
 def segment(m,p):b.extend(bytes([255,m])+(len(p)+2).to_bytes(2,'big')+p)
 factors=[(h,v)]+[(1,1)]*(n-1)
 segment(195,bytes([16])+struct.pack('>HHB',height,w,n)+b''.join(bytes([i+1,(a<<4)|z,0]) for i,(a,z) in enumerate(factors)))
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
       current=sample(x,y);pred=sample(x-1,y) if x else sample(0,y-1) if y else 32768
       diff=(current-pred+32768)%65536-32768;cat=abs(diff).bit_length();bits+=f'{cat:05b}'
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
 entries=[(256,4,1,width),(257,4,1,height),(258,3,3,[16]*3),(259,3,1,7),(262,3,1,6),(277,3,1,3),(284,3,1,planar),(530,3,2,[h,v]),(531,3,1,position)]
 entries+=([(322,4,1,sw),(323,4,1,sh)] if tile else [(278,4,1,sh)])
 offsetTag,countTag=(324,325) if tile else (273,279)
 entries.extend([(offsetTag,4,len(payloads),[0]*len(payloads)),(countTag,4,len(payloads),[len(p) for p in payloads])]);entries.sort();count=len(entries)
 def header(offsets):
  extra=bytearray();directory=bytearray();start=8+2+count*12+4
  for tag,kind,n,values in entries:
   if tag==offsetTag:values=offsets
   if isinstance(values,int):values=[values]
   raw=struct.pack(endian+('H' if kind==3 else 'I')*n,*values)
   if len(raw)<=4:field=raw+bytes(4-len(raw))
   else:field=struct.pack(endian+'I',start+len(extra));extra.extend(raw)
   directory.extend(struct.pack(endian+'HHI',tag,kind,n)+field)
  return (b'MM\0*' if endian=='>' else b'II*\0')+struct.pack(endian+'IH',8,count)+directory+bytes(4)+extra
 head=header([0]*len(payloads));offsets=[];at=len(head)
 for p in payloads:offsets.append(at);at+=len(p)
 return header(offsets)+b''.join(payloads)

rows=[]
for h,v in [(2,1),(2,2),(4,2)]:
 for width,height in [(5,3),(17,11)]:
  for layout in range(8):
   name=f'h{h}-v{v}-w{width}-h{height}-l{layout}';tile=layout in (2,3,6,7);planar=layout>=4;separate=layout in (1,2);position=2 if layout in (1,3,5,6) else 1;sw=16 if tile else width;sh=16 if tile else 8
   segments=[];reference=[0]*(width*height*3)
   for top in range(0,height,sh):
    for left in range(0,width,sw):
     dw=sw;dh=sh if tile else min(sh,height-top);vw=min(sw,width-left);vh=min(sh,height-top)
     planes=[[[value(c,left//(h if c else 1)+x,top//(v if c else 1)+y) for x in range((dw+h-1)//h if c else dw)]for y in range((dh+v-1)//v if c else dh)]for c in range(3)]
     encoded=[]
     if planar:
      for c,p in enumerate(planes):
       stream=jpeg([p],len(p[0]),len(p),1,1,False);words=decode(stream,name+'-temp');assert words==sum(p,[]);encoded.append(stream)
     else:
      stream=jpeg(planes,dw,dh,h,v,separate);words=decode(stream,name+'-temp');expected=[planes[c][y//(v if c else 1)][x//(h if c else 1)]for y in range(dh)for x in range(dw)for c in range(3)];assert words==expected,(name,'independent words');encoded=[stream]
      if top==0 and left==0:
       (root/(name+'.jpg')).write_bytes(stream);(root/(name+'.nearest.raw')).write_bytes(struct.pack('<'+'H'*len(words),*words))
       interpolated=[sum(planes[0],[])]+[interpolate(p,dw,dh,h,v,1)for p in planes[1:]]
       interleaved=[interpolated[c][i]for i in range(dw*dh)for c in range(3)];(root/(name+'.bilinear.raw')).write_bytes(struct.pack('<'+'H'*len(interleaved),*interleaved))
     segments.append(encoded)
     grids=[sum([row[:vw] for row in planes[0][:vh]],[])]+[interpolate([row[:(vw+h-1)//h]for row in p[:(vh+v-1)//v]],vw,vh,h,v,position)for p in planes[1:]]
     for y in range(vh):
      for x in range(vw):
       for c in range(3):reference[((top+y)*width+left+x)*3+c]=grids[c][y*vw+x]
   payloads=[segment[c]for c in range(3)for segment in segments]if planar else [segment[0]for segment in segments]
   (root/(name+'.tif')).write_bytes(tiff(width,height,h,v,layout,payloads));rgba=bytearray()
   for i in range(width*height):
    y,cb,cr=reference[i*3:i*3+3];cb-=32768;cr-=32768;r=y+1.402*cr;b=y+1.772*cb;g=(y-.299*r-.114*b)/.587
    rgb=[min(65535,max(0,round(q)))for q in (r,g,b)];rgba.extend([(q*255+32767)//65535 for q in rgb]+[255])
   (root/(name+'.rgba')).write_bytes(rgba);rows.append([name,width,height,h,v,layout,position])
with (root/'manifest.csv').open('w')as f:
 writer=csv.writer(f);writer.writerow(['name','width','height','horizontal','vertical','layout','position']);writer.writerows(rows)
files=sorted(p for p in root.iterdir()if p.suffix in ('.jpg','.raw','.tif','.rgba'))
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n'for p in files));print(len(rows),'TIFF cases; varying samples verified by libjpeg-turbo, interpolation by Pillow')
