"""Combine qualified native planar JPEG streams into constructed multi-scan TIFF.
No entropy bytes are rewritten; source TIFF color/ICC references remain valid.
"""
from pathlib import Path
import csv,hashlib,struct
root=Path(__file__).resolve().parent
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

def write_tiff(tags,e,offset_tag,count_tag,streams):
 tags=dict(tags);tags[284]=(3,1,struct.pack(e+'H',1));count=len(streams)
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

rows=[];stream_count=0
for folder in ('TiffJpegArithmeticLosslessColor','TiffJpegArithmeticLosslessChroma','TiffJpegArithmeticExtra'):
 source=root.parent/folder;hashes={name:digest for digest,name in(line.split('  ',1)for line in(source/'SHA256SUMS').read_text().splitlines())}
 for row in csv.DictReader((source/'manifest.csv').open()):
  name=row['file'];original=(source/name).read_bytes();e,tags=entries(original)
  photo=values(tags,262,e)[0];n=values(tags,277,e)[0]
  if values(tags,284,e)[0]!=2:continue
  h,v=values(tags,530,e)if photo==6 else(1,1)
  if not(photo==5 and n==5 or photo==6 and n==4 and(h,v)==(4,2)):continue
  digest=hashlib.sha256(original).hexdigest();assert digest==hashes[name]
  offset_tag,count_tag=(324,325)if 324 in tags else(273,279)
  offsets=values(tags,offset_tag,e);lengths=values(tags,count_tag,e);assert len(offsets)%n==0
  per_plane=len(offsets)//n;parts=[original[at:at+length]for at,length in zip(offsets,lengths)]
  reverse=e=='>';streams=[combine([parts[c*per_plane+i]for c in range(n)],h,v,reverse)for i in range(per_plane)]
  output=name.replace('-pl2-','-pl1-');assert output!=name and not any(r[0]==output for r in rows)
  (root/output).write_bytes(write_tiff(tags,e,offset_tag,count_tag,streams));stream_count+=len(streams)
  for suffix in('.rgba','.icc-rgba'):
   p=source/(name+suffix)
   if p.exists():assert hashlib.sha256(p.read_bytes()).hexdigest()==hashes[p.name];(root/(output+suffix)).write_bytes(p.read_bytes())
  rows.append([output,photo,row['bits'],folder,name,digest,'reverse'if reverse else'forward'])
with(root/'manifest.csv').open('w')as f:
 writer=csv.writer(f,lineterminator='\n');writer.writerow(['file','photometric','bits','sourceCorpus','sourceFile','sourceSHA256','scanOrder']);writer.writerows(rows)
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n'for p in sorted(root.glob('*.tif*'))))
print(len(rows),'TIFFs;',stream_count,'combined JPEG streams;',len(list(root.glob('*.icc-rgba'))),'CMYK references')
