"""Re-encode qualified component planes with the native arithmetic producer.
Usage: generate.py <prepared-jpeg> <compiled-decode> <scratch-directory>
"""
from pathlib import Path
import csv,hashlib,os,struct,subprocess,sys
root=Path(__file__).resolve().parent;assets=root.parent;sys.path.insert(0,str(assets))
from tiff_jpeg_multiscan_reference import entries,values,combine,write_tiff
oracle,turbo=map(lambda p:str(Path(p).resolve()),sys.argv[1:3]);scratch=Path(sys.argv[3]).resolve();scratch.mkdir(parents=True,exist_ok=True)
rows=[];native_count=0;stored_count=0

def run(args,env=None):
 r=subprocess.run(args,env=env,capture_output=True,text=True)
 if r.returncode or r.stderr:raise RuntimeError(str(args)+'\n'+r.stderr)

def checked(path,hashes):
 data=path.read_bytes();assert hashlib.sha256(data).hexdigest()==hashes[path.name];return data

def parts(data):
 e,tags=entries(data);ot,ct=(324,325)if 324 in tags else(273,279)
 return e,tags,ot,ct,[data[a:a+n]for a,n in zip(values(tags,ot,e),values(tags,ct,e))]

def transcode(part,bits,predictor):
 global native_count
 at=2
 while part[at+1]!=195:at+=2+int.from_bytes(part[at+2:at+4],'big')
 precision,height,width,n=struct.unpack('>BHHB',part[at+4:at+10]);assert precision==bits and n==1
 source=scratch/'source.jpg';raw=scratch/'source.raw';source.write_bytes(part);run([turbo,str(source),str(raw),'0'])
 data=raw.read_bytes();words=struct.unpack('<'+'H'*(len(data)//2),data);assert len(words)==width*height
 size=1 if bits<=8 else 2;pnm=scratch/'input.pnm';pnm.write_bytes(f'P5\n{width} {height}\n{(1<<bits)-1}\n'.encode()+b''.join(v.to_bytes(size,'big')for v in words))
 output=scratch/'encoded.jpg';decoded=scratch/'decoded.pnm';env=dict(os.environ,OFFICEIMO_TEST_DEPTH='1',OFFICEIMO_TEST_POINT='0',OFFICEIMO_TEST_PREDICTOR=str(predictor))
 run([oracle,'-a','-p','-c','-z',str(width*2),str(pnm),str(output)],env);run([oracle,'-c',str(output),str(decoded)])
 data=decoded.read_bytes().split(b'\n',3)[3];actual=tuple(int.from_bytes(data[i:i+size],'big')for i in range(0,len(data),size));assert actual==words
 native_count+=1;return output.read_bytes()

for folder in ['TiffJpegChromaPrecision','TiffJpegChromaAlpha']:
 source=assets/folder;hashes={n:h for h,n in(line.split('  ',1)for line in(source/'SHA256SUMS').read_text().splitlines())}
 for row in csv.DictReader((source/'manifest.csv').open()):
  if row['horizontal']!='4' or row['vertical']!='4' or int(row['layout'])<4:continue
  name=row['name']+('.tif'if folder=='TiffJpegChromaPrecision'else'');original=(source/name).read_bytes();assert hashlib.sha256(original).hexdigest()==hashes[name]
  e,tags,ot,ct,planes=parts(original);n=values(tags,277,e)[0];per_plane=len(planes)//n;bits=values(tags,258,e)[0];width=values(tags,256,e)[0];height=values(tags,257,e)[0];alpha=values(tags,338,e)[0]if 338 in tags else 0
  predictor=1+(len(rows)//2)%7;encoded=[transcode(part,bits,predictor)for part in planes]
  device=checked(source/(name+'.rgba'if folder=='TiffJpegChromaAlpha'else Path(name).stem+'.rgba'),hashes)
  if folder=='TiffJpegChromaAlpha':profiled=checked(source/(name+'.icc-rgba'),hashes)
  else:
   companion=assets/'TiffJpegChromaAlpha'/('a2-'+name);companion_hashes={n:h for h,n in(line.split('  ',1)for line in(companion.parent/'SHA256SUMS').read_text().splitlines())}
   assert planes==parts(checked(companion,companion_hashes))[4][:len(planes)]
   profiled=bytearray(checked(Path(str(companion)+'.icc-rgba'),companion_hashes));profiled[3::4]=bytes([255])*(width*height)
  for planar in (2,1):
   streams=encoded if planar==2 else[combine([encoded[c*per_plane+i]for c in range(n)],4,4,e=='>')for i in range(per_plane)]
   output=Path(name).stem+('-planar'if planar==2 else'-multiscan')+'.tif'
   (root/output).write_bytes(write_tiff(tags,e,ot,ct,streams,planar));(root/(output+'.rgba')).write_bytes(device);(root/(output+'.icc-rgba')).write_bytes(profiled);stored_count+=len(streams)
   rows.append([output,6,bits,'planar'if planar==2 else'multiscan',alpha,width,height,folder,name,hashlib.sha256(original).hexdigest(),predictor,'reverse'if e=='>'else'forward'])
with(root/'manifest.csv').open('w')as f:
 w=csv.writer(f,lineterminator='\n');w.writerow(['name','photometric','bits','layout','alphaKind','width','height','sourceCorpus','sourceFile','sourceSHA256','predictor','scanOrder']);w.writerows(rows)
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n'for p in sorted(root.glob('*.tif*'))))
print(len(rows),'TIFFs;',native_count,'native re-encoded and exactly decoded component streams;',stored_count,'stored streams')
