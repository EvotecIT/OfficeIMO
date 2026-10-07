"""Native libjpeg-turbo/LibTIFF alpha cases using the shared lossless generator."""
from pathlib import Path
import csv,hashlib,struct,subprocess,sys
root=Path(__file__).resolve().parent
sys.path.insert(0,str(root.parent/'TiffJpegLossless16'))
from references import write_references
rows=[]
for bits in range(2,17):
 if bits in(8,12,16):continue
 for photo in(0,1,2,6):
  for layout in(0,3,4,7):
   kind=1 if layout%2==0 else 2;point=bits-1 if photo==0 else 1 if layout==7 else 0
   predictor=1 if layout<4 else 7;separate=int(layout%2==1);restart=2
   name=f'p{photo}-b{bits}-l{layout}-a{kind}.tif'
   subprocess.run([sys.argv[1],str(root/name),str(photo),str(predictor),str(point),str(layout),str(separate),str(restart),str(kind),str(bits)],check=True)
   if photo==6:
    data=bytearray((root/name).read_bytes());e='>'if data[:2]==b'MM'else'<';at=struct.unpack_from(e+'I',data,4)[0]
    for i in range(struct.unpack_from(e+'H',data,at)[0]):
     p=at+2+i*12
     if struct.unpack_from(e+'H',data,p)[0]==262:struct.pack_into(e+'H',data,p+8,6)
    (root/name).write_bytes(data)
   write_references(root,name,photo,kind,bits)
   # Native words and direct RGBA/ICC references own this corpus; a second
   # uncompressed TIFF representation is unnecessary here.
   (root/(name+'.reference.tif')).unlink()
   (root/(name+'.jpg')).unlink()
   (root/(name+'.jpg.raw')).unlink()
   rows.append([name,photo,bits,layout,kind,predictor,point,separate,restart])
with(root/'manifest.csv').open('w')as f:
 w=csv.writer(f,lineterminator='\n');w.writerow(['name','photometric','bits','layout','alphaKind','predictor','point','separate','restart']);w.writerows(rows)
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n'for p in sorted(root.glob('*.tif*'))))
print(len(rows),'native lossless alpha containers')
