"""Combine qualified native planar JPEG streams into constructed multi-scan TIFF.
No entropy bytes are rewritten; source TIFF color/ICC references remain valid.
"""
from pathlib import Path
import csv,hashlib,sys
root=Path(__file__).resolve().parent
sys.path.insert(0,str(root.parent))
from tiff_jpeg_multiscan_reference import entries,values,combine,write_tiff

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
