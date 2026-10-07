from pathlib import Path
import struct,json
from PIL import Image,ImageChops
root=Path(__file__).resolve().parent;results=[]
for path in sorted(root.glob('*.tiff')):
 if path.stem=='ojpeg_single_strip_no_rowsperstrip':continue
 b=path.read_bytes();e='<'if b[:2]==b'II'else'>';at=struct.unpack_from(e+'I',b,4)[0];n=struct.unpack_from(e+'H',b,at)[0];fields={};tags={}
 for i in range(n):
  pos=at+2+12*i;tag,t,count=struct.unpack_from(e+'HHI',b,pos)
  if t not in(1,2,3,4,5,7):continue
  size={1:1,2:1,3:2,4:4,5:8,7:1}[t];off=pos+8 if count*size<=4 else struct.unpack_from(e+'I',b,pos+8)[0];data=b[off:off+count*size];fields[tag]=(t,count,data)
  if t in(3,4):tags[tag]=list(struct.unpack(e+('H'if t==3 else'I')*count,data))
 w=tags[256][0];h=tags[257][0];n=tags[277][0];hs,vs=tags.get(530,[2,2]);tile=324 in tags;sw=tags[322][0]if tile else w;sh=tags[323][0]if tile else tags.get(278,[h])[0];offsets=tags[324 if tile else 273];counts=tags[325 if tile else 279];payloads=[]
 def marker(m,data):return bytes([255,m])+struct.pack('>H',len(data)+2)+data
 tables=b''
 for c in range(n):
  tables+=marker(219,bytes([c])+b[tags[519][c]:tags[519][c]+64])
  for kind,tag in [(0,520),(1,521)]:
   offset=tags[tag][c];length=16+sum(b[offset:offset+16]);tables+=marker(196,bytes([(kind<<4)|c])+b[offset:offset+length])
 for index,(off,length)in enumerate(zip(offsets,counts)):
  rows=sh if tile else min(sh,h-index*sh)
  sof=bytes([8])+struct.pack('>HHB',rows,sw,n)+b''.join(bytes([c+1,(hs<<4)|vs if c==0 else 17,c])for c in range(n))
  sos=bytes([n])+b''.join(bytes([c+1,(c<<4)|c])for c in range(n))+bytes([0,63,0])
  entropy=b[off:off+length]
  if entropy.startswith(b'\xff\xda'):entropy=entropy[2+int.from_bytes(entropy[2:4],'big'):]
  if entropy.endswith(b'\xff\xd9'):entropy=entropy[:-2]
  j=b'\xff\xd8'+tables+marker(192,sof)+marker(218,sos)+entropy+b'\xff\xd9';payloads.append(j)
 # Retain metadata and TIFF geometry; replace segment payloads and legacy tag fields.
 for tag in list(fields):
  if 512<=tag<=521:del fields[tag]
 fields[259]=(3,1,struct.pack(e+'H',7))
 if 532 in tags:fields[532]=(5,len(tags[532]),b''.join(struct.pack(e+'II',x,1)for x in tags[532]))
 if not tile and 278 not in fields:fields[278]=(4,1,struct.pack(e+'I',h))
 offsetTag=324 if tile else 273;countTag=325 if tile else 279;fields[offsetTag]=(4,len(payloads),bytes(4*len(payloads)));fields[countTag]=(4,len(payloads),struct.pack(e+'I'*len(payloads),*[len(p)for p in payloads]))
 def header(offsets):
  fields[offsetTag]=(4,len(offsets),struct.pack(e+'I'*len(offsets),*offsets));count=len(fields);extra=bytearray();directory=bytearray();start=8+2+count*12+4
  for tag,(t,n,data)in sorted(fields.items()):
   if len(data)<=4:value=data+bytes(4-len(data))
   else:value=struct.pack(e+'I',start+len(extra));extra.extend(data)
   directory.extend(struct.pack(e+'HHI',tag,t,n)+value)
  return b[:4]+struct.pack(e+'IH',8,count)+directory+bytes(4)+extra
 head=header([0]*len(payloads));offsets=[];at=len(head)
 for p in payloads:offsets.append(at);at+=len(p)
 dest=root/(path.stem+'.reference.tif');dest.write_bytes(header(offsets)+b''.join(payloads));im=Image.open(dest).convert('RGB');reference=Image.open(root/(path.stem+'.png')).convert('RGB');extrema=ImageChops.difference(im,reference).getextrema();print(path.name,extrema);results.append(dict(file=path.name,maxima=extrema))
print(json.dumps(results,indent=2))
