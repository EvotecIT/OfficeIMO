"""Test-only TIFF directory and native JPEG multi-scan assembly helpers."""
import struct
sizes={1:1,2:1,3:2,4:4,5:8,7:1,11:4,12:8}

def entries(data):
 endian='<'if data[:2]==b'II'else'>';at=struct.unpack_from(endian+'I',data,4)[0]
 result={}
 for i in range(struct.unpack_from(endian+'H',data,at)[0]):
  p=at+2+i*12;tag,kind,count=struct.unpack_from(endian+'HHI',data,p);length=sizes[kind]*count
  start=p+8 if length<=4 else struct.unpack_from(endian+'I',data,p+8)[0]
  result[tag]=(kind,count,data[start:start+length])
 return endian,result

def values(tags,tag,e):
 kind,count,data=tags[tag];return struct.unpack(e+('H'if kind==3 else'I')*count,data)

def combine(parts,h,v,reverse):
 frames=[];scans=[]
 for c,part in enumerate(parts):
  assert part[:2]==b'\xff\xd8'and part[-2:]==b'\xff\xd9'
  at=2;controls=bytearray(b'\xff\xdd\0\4\0\0')
  while True:
   assert part[at]==255;marker=part[at+1];length=int.from_bytes(part[at+2:at+4],'big');end=at+2+length
   if marker==203:
    precision,height,width,n=struct.unpack('>BHHB',part[at+4:at+10]);assert n==1 and part[at+11]==17
    frames.append((precision,width,height))
   elif marker==218:
    scan=bytearray(part[at:end]);assert scan[4]==1;scan[5]=c+1
    scans.append(bytes(controls)+scan+part[end:-2]);break
   elif marker in(204,221):controls.extend(part[at:end])
   else:assert marker==254 or 224<=marker<=239
   at=end
 assert len(frames)==len(parts)
 precision,width,height=frames[0];n=len(parts)
 for c,(p,w,hh)in enumerate(frames):
  assert p==precision and w==((width+h-1)//h if c in(1,2)else width)and hh==((height+v-1)//v if c in(1,2)else height)
 payload=struct.pack('>BHHB',precision,height,width,n)+b''.join(bytes([c+1,17 if c in(1,2)else(h<<4)|v,0])for c in range(n))
 order=range(n-1,-1,-1)if reverse else range(n)
 return b'\xff\xd8\xff\xcb'+struct.pack('>H',len(payload)+2)+payload+b''.join(scans[c]for c in order)+b'\xff\xd9'

def write_tiff(tags,e,offset_tag,count_tag,streams,planar=1):
 tags=dict(tags);tags[284]=(3,1,struct.pack(e+'H',planar));count=len(streams)
 tags[offset_tag]=(4,count,bytes(count*4));tags[count_tag]=(4,count,struct.pack(e+'I'*count,*map(len,streams)))
 start=8+2+12*len(tags)+4
 def header(offsets):
  fields=bytearray();extra=bytearray()
  for tag,(kind,n,data)in sorted(tags.items()):
   if tag==offset_tag:data=struct.pack(e+'I'*count,*offsets)
   if len(data)<=4:value=data+bytes(4-len(data))
   else:
    if len(extra)%2:extra.append(0)
    value=struct.pack(e+'I',start+len(extra));extra.extend(data)
   fields.extend(struct.pack(e+'HHI',tag,kind,n)+value)
  return (b'II*\0'if e=='<'else b'MM\0*')+struct.pack(e+'IH',8,len(tags))+fields+bytes(4)+extra
 data=header([0]*count);offsets=[];at=len(data)
 for stream in streams:offsets.append(at);at+=len(stream)
 return header(offsets)+b''.join(streams)
