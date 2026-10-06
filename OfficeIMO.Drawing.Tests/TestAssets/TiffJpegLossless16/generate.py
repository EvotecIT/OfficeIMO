from pathlib import Path
import csv,hashlib,math,struct,subprocess,sys
precision=int(sys.argv[2]) if len(sys.argv)>2 else 16
assert precision in (12,16)
maximum=(1<<precision)-1;midpoint=1<<(precision-1)
root=Path(sys.argv[3]).resolve() if len(sys.argv)>3 else Path(__file__).resolve().parent;root.mkdir(parents=True,exist_ok=True);rows=[]
from references import write_references
for pred in range(1,8):
 for photo in (0,1,2,5,6):
  for layout in range(8):
   point=(0,1,8 if precision==16 else 6,precision-1)[(pred+layout)%4];separate=(pred+layout)%2;restart=2 if layout&2 else 0;kind=1 if layout%2==0 else 2
   if photo==5:kind=0
   name=f'p{photo}-d{pred}-l{layout}-t{point}-s{separate}-r{restart}-a{kind}.tif'
   subprocess.run([sys.argv[1],str(root/name),str(photo),str(pred),str(point),str(layout),str(separate),str(restart),str(kind),str(precision)],check=True)
   if photo==6:
    data=bytearray((root/name).read_bytes());endian='>' if data[:2]==b'MM' else '<';offset=struct.unpack_from(endian+'I',data,4)[0];count=struct.unpack_from(endian+'H',data,offset)[0]
    for i in range(count):
     at=offset+2+i*12
     if struct.unpack_from(endian+'H',data,at)[0]==262:struct.pack_into(endian+'H',data,at+8,6)
    (root/name).write_bytes(data)
   write_references(root,name,photo,kind,precision)
   base=4 if photo==5 else 3 if photo in (2,6) else 1;n=4 if photo==5 else base+1
   rows.append([name,photo,n,base if kind else -1,kind,35,19,0,16 if layout&1 else 35,16,1 if layout&4 else n,pred,point,separate,restart])
with (root/'manifest.csv').open('w') as f:
 w=csv.writer(f);w.writerow(['file','photometric','samples','alphaIndex','alphaKind','width','height','tolerance','jpegWidth','jpegHeight','jpegComponents','predictor','pointTransform','separateScans','restartRows']);w.writerows(rows)
files=sorted(p for p in root.iterdir() if p.suffix in ('.tif','.raw','.rgba','.jpg'))
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n' for p in files))
print(f'{len(rows)} fixtures generated, {len(files)} fixture/reference files')
