"""Declare existing independently referenced straight-alpha samples unspecified.
Only TIFF ExtraSamples changes; JPEG bytes and referenced color samples do not.
"""
from pathlib import Path
import csv,hashlib,struct
root=Path(__file__).resolve().parent;rows=[]
for folder in ('TiffJpegArithmeticLosslessColor','TiffJpegArithmeticLosslessChroma'):
 source=root.parent/folder
 hashes={name:digest for digest,name in (line.split('  ',1)for line in(source/'SHA256SUMS').read_text().splitlines())}
 for row in csv.DictReader((source/'manifest.csv').open()):
  if row['extra']!='2':continue
  name=row['file'];original=(source/name).read_bytes();digest=hashlib.sha256(original).hexdigest();assert digest==hashes[name]
  data=bytearray(original);endian='>'if data[:2]==b'MM'else'<';offset=struct.unpack_from(endian+'I',data,4)[0]
  count=struct.unpack_from(endian+'H',data,offset)[0];found=False
  for i in range(count):
   at=offset+2+i*12;tag,kind,n=struct.unpack_from(endian+'HHI',data,at)
   if tag!=338:continue
   assert kind==3 and n==1 and struct.unpack_from(endian+'H',data,at+8)[0]==2
   struct.pack_into(endian+'H',data,at+8,0);found=True
  assert found and sum(a!=b for a,b in zip(data,original))==1
  output=name.replace('-e2.tif','-e0.tif');assert not any(r[0]==output for r in rows)
  (root/output).write_bytes(data)
  for suffix in ('.rgba','.icc-rgba'):
   p=source/(name+suffix)
   if not p.exists():continue
   assert hashlib.sha256(p.read_bytes()).hexdigest()==hashes[p.name]
   reference=bytearray(p.read_bytes());assert len(reference)==35*19*4
   # Straight color remains unchanged even where the ignored channel is zero.
   reference[3::4]=bytes([255])*(35*19)
   (root/(output+suffix)).write_bytes(reference)
  rows.append([output,row['photometric'],row['bits'],folder,name,digest])
with(root/'manifest.csv').open('w')as f:
 writer=csv.writer(f,lineterminator='\n');writer.writerow(['file','photometric','bits','sourceCorpus','sourceFile','sourceSHA256']);writer.writerows(rows)
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n'for p in sorted(root.glob('*.tif*'))))
print(len(rows),'derived TIFFs;',len(list(root.glob('*.icc-rgba'))),'profiled CMYK references')
