"""Recontainer independently encoded JPEG TIFF samples using legacy TIFF tags."""
from pathlib import Path
import csv,struct,hashlib
root=Path(__file__).resolve().parent

def read_tiff(path):
 data=path.read_bytes();e='<'if data[:2]==b'II'else'>';at=struct.unpack_from(e+'I',data,4)[0];n=struct.unpack_from(e+'H',data,at)[0];fields={}
 for i in range(n):
  pos=at+2+i*12;tag,t,count=struct.unpack_from(e+'HHI',data,pos);size={1:1,2:1,3:2,4:4,5:8,7:1}.get(t)
  if not size:continue
  off=pos+8 if count*size<=4 else struct.unpack_from(e+'I',data,pos+8)[0];fields[tag]=(t,count,data[off:off+size*count])
 def values(tag,default=None):
  if tag not in fields:return default
  t,n,b=fields[tag];return list(struct.unpack(e+('H'if t==3 else'I')*n,b))
 return data,e,fields,values

def numeric(fields,e,tag,values,t=4):fields[tag]=(t,len(values),struct.pack(e+('H'if t==3 else'I')*len(values),*values))

def write_tiff(e,fields,payloads,tables=None,interchange=False):
 fields=dict(fields);tile=324 in fields;offsetTag=324 if tile else 273;countTag=325 if tile else 279;tables=tables or {}
 if interchange:
  for tag in (273,278,279,322,323,324,325):fields.pop(tag,None)
  numeric(fields,e,513,[0]);numeric(fields,e,514,[len(payloads[0])])
 else:numeric(fields,e,offsetTag,[0]*len(payloads));numeric(fields,e,countTag,[len(p)for p in payloads])
 for tag,items in tables.items():numeric(fields,e,tag,[0]*len(items))
 def header():
  count=len(fields);start=8+2+count*12+4;extra=bytearray();directory=bytearray()
  for tag,(t,n,data)in sorted(fields.items()):
   if len(data)<=4:value=data+bytes(4-len(data))
   else:value=struct.pack(e+'I',start+len(extra));extra.extend(data)
   directory.extend(struct.pack(e+'HHI',tag,t,n)+value)
  return (b'II*\0'if e=='<'else b'MM\0*')+struct.pack(e+'IH',8,count)+directory+bytes(4)+extra
 at=len(header());tableBytes=bytearray()
 for tag,items in tables.items():
  offsets=[]
  for item in items:offsets.append(at+len(tableBytes));tableBytes.extend(item)
  numeric(fields,e,tag,offsets)
 at+=len(tableBytes);offsets=[]
 for p in payloads:offsets.append(at);at+=len(p)
 numeric(fields,e,513 if interchange else offsetTag,offsets)
 return header()+tableBytes+b''.join(payloads)

def parse_jpeg(data,inherited=None):
 quant=dict((inherited or {}).get('quant',{}));huff=dict((inherited or {}).get('huff',{}));frame=None;scans=[];restart=0;x=2
 while x<len(data):
  assert data[x]==255
  marker=data[x+1];x+=2
  if marker==217:break
  length=int.from_bytes(data[x:x+2],'big');part=data[x+2:x+length];x+=length
  if marker==219:
   p=0
   while p<len(part):
    info=part[p];assert info<4;quant[info]=part[p+1:p+65];p+=65
  elif marker==196:
   p=0
   while p<len(part):
    info=part[p];count=sum(part[p+1:p+17]);huff[info]=part[p+1:p+17+count];p+=17+count
  elif marker in(192,193,195):
   precision,height,width,n=struct.unpack('>BHHB',part[:6]);components=[tuple(part[6+i*3:9+i*3])for i in range(n)];frame=(marker,precision,width,height,components)
  elif marker==221:restart=int.from_bytes(part,'big')
  elif marker==218:
   begin=x
   while x<len(data)-1:
    if data[x]!=255:x+=1;continue
    if data[x+1]==0 or 208<=data[x+1]<=215:x+=2;continue
    break
   scans.append((part,data[begin:x]))
 return dict(quant=quant,huff=huff,frame=frame,scans=scans,restart=restart)

def build(source,mode,name,crop=False):
 data,e,fields,val=read_tiff(source);tile=324 in fields;offsetTag=324 if tile else 273;countTag=325 if tile else 279
 offsets=val(offsetTag);counts=val(countTag);samples=val(277)[0];planar=val(284,[1])[0];perPlane=len(offsets)//(samples if planar==2 else 1)
 shared=fields.get(347,(0,0,b''))[2];inherited=parse_jpeg(shared)if shared else None
 streams=[data[o:o+n]for o,n in zip(offsets,counts)];parsed=[parse_jpeg(s,inherited)for s in streams]
 if crop:
  indices=[p*perPlane for p in range(samples if planar==2 else 1)]
  streams=[streams[i]for i in indices];parsed=[parsed[i]for i in indices];perPlane=1
  if tile:
   numeric(fields,e,256,[min(val(256)[0],parsed[0]['frame'][2]-3)]);numeric(fields,e,257,[min(val(257)[0],parsed[0]['frame'][3]-5)])
  else:numeric(fields,e,257,[parsed[0]['frame'][3]]);numeric(fields,e,278,[parsed[0]['frame'][3]])
 # A complete interchange image carries all channels in one JPEG frame.
 if mode=='interchange':
  assert planar==1 and len(streams)==1
  stream=streams[0]if not shared else b'\xff\xd8'+shared[2:-2]+streams[0][2:]
  payloads=[stream]
 elif mode=='segments':payloads=[s if not shared else b'\xff\xd8'+shared[2:-2]+s[2:]for s in streams]
 else:payloads=[]
 reference=write_tiff(e,fields,streams);(root/(name+'.reference.tif')).write_bytes(reference)
 tables={};predictors=[];points=[];canonical=[]
 if mode=='raw':
  tables={520:[]}
  if parsed[0]['frame'][0]!=195:tables.update({519:[],521:[]})
  for plane in range(samples if planar==2 else 1):
   initial=parsed[plane*perPlane];frame=initial['frame'];assert len(initial['scans'])==1
   scan,_=initial['scans'][0];n=len(frame[4]);assert scan[0]==n
   signature=[]
   for c,(identity,sampling,qid)in enumerate(frame[4]):
    assert scan[1+c*2]==identity
    selectors=scan[2+c*2];dc=initial['huff'][selectors>>4];tables[520].append(dc)
    q=initial['quant'].get(qid);ac=initial['huff'].get(16+(selectors&15));signature.append((q,dc,ac))
    if frame[0]!=195:tables[519].append(q);tables[521].append(ac)
    predictors.append(scan[-3]);points.append(scan[-1])
   canonical.append(signature)
  for i,item in enumerate(parsed):
   assert len(item['scans'])==1
   frame=item['frame'];scan,entropy=item['scans'][0];signature=[]
   for c,(_,_,qid)in enumerate(frame[4]):
    selector=scan[2+c*2];signature.append((item['quant'].get(qid),item['huff'][selector>>4],item['huff'].get(16+(selector&15))))
   assert signature==canonical[i//perPlane if planar==2 else 0],(name,'changing tables')
   assert item['restart']==parsed[0]['restart'];payloads.append(entropy)
 for tag in list(fields):
  if tag==347 or 512<=tag<=521:fields.pop(tag)
 numeric(fields,e,259,[6],3);process=14 if parsed[0]['frame'][0]==195 else 1;numeric(fields,e,512,[process],3)
 if mode=='raw':
  numeric(fields,e,515,[parsed[0]['restart']],3)
  if process==14:numeric(fields,e,517,predictors,3);numeric(fields,e,518,points,3)
 (root/(name+'.tif')).write_bytes(write_tiff(e,fields,payloads,tables,mode=='interchange'))
 width=struct.unpack(e+'I',fields[256][2])[0]if fields[256][0]==4 else struct.unpack(e+'H',fields[256][2])[0]
 height=struct.unpack(e+'I',fields[257][2])[0]if fields[257][0]==4 else struct.unpack(e+'H',fields[257][2])[0]
 return [name+'.tif',width,height,val(262)[0],samples,process,mode,source.parent.name,source.name]

rows=[]
for record in csv.DictReader((root.parent/'TiffJpeg/manifest.csv').open()):
 if record['tables']not in('0','3'):continue
 source=root.parent/'TiffJpeg'/record['file'];rows.append(build(source,'raw','raw-'+source.stem,record['tables']=='0'))
# Lossless table definitions can change per segment. Retain one strip per plane.
for corpus in ['TiffJpegLossless','TiffJpegLossless16']:
 for photo in(0,1,2,5):
  for layout in(0,4):
   paths=sorted((root.parent/corpus).glob(f'p{photo}-d2-l{layout}-*.tif'))
   source=next(p for p in paths if '.reference.'not in p.name)
   # Predictor2/layout0 uses an interleaved scan; planar data has one component per scan.
   rows.append(build(source,'raw',corpus+'-raw-'+source.stem,True))
 for photo in(1,2,5):
  source=next(p for p in sorted((root.parent/corpus).glob(f'p{photo}-d2-l0-*.tif'))if '.reference.'not in p.name)
  rows.append(build(source,'interchange',corpus+'-interchange-'+source.stem,True))
# Self-contained compatibility striles retain their own tables and scans.
for record in list(csv.DictReader((root.parent/'TiffJpeg/manifest.csv').open()))[::20]:
 source=root.parent/'TiffJpeg'/record['file'];rows.append(build(source,'segments','segments-'+source.stem))
# One zero-difference Huffman bit followed by seven one padding bits is a
# complete 1x1 lossless scan. This guards against a baseline-only byte minimum.
for precision in (8,16):
 for e,label in [('<','le'),('>','be')]:
  fields={}
  for tag,value,t in [(256,1,4),(257,1,4),(258,precision,3),(259,6,3),(262,1,3),(277,1,3),(278,1,4),(512,14,3),(517,1,3),(518,0,3)]:numeric(fields,e,tag,[value],t)
  dc=bytes([1])+bytes(15)+bytes([0])
  name=f'tiny-lossless-{precision}-{label}'
  (root/(name+'.tif')).write_bytes(write_tiff(e,fields,[bytes([127])],{520:[dc]}))
  def marker(code,payload):return bytes([255,code])+struct.pack('>H',len(payload)+2)+payload
  jpeg=b'\xff\xd8'+marker(196,bytes([0])+dc)+marker(195,bytes([precision])+struct.pack('>HHB',1,1,1)+bytes([1,17,0]))+marker(218,bytes([1,1,0,1,0,0]))+bytes([127,255,217])
  for tag in (512,517,518):fields.pop(tag)
  numeric(fields,e,259,[7],3)
  (root/(name+'.reference.tif')).write_bytes(write_tiff(e,fields,[jpeg]))
  rows.append([name+'.tif',1,1,1,1,14,'raw','T.81-authored','one zero difference'])
with(root/'manifest.csv').open('w')as f:
 w=csv.writer(f);w.writerow(['file','width','height','photometric','samples','process','mode','sourceCorpus','source']);w.writerows(rows)
files=sorted(p for p in root.rglob('*') if p.suffix in ('.tif','.tiff','.png'));(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.relative_to(root).as_posix()+'\n'for p in files));print(len(rows),'legacy cases generated')
