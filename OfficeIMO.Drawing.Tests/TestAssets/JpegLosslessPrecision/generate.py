"""Generate precision/predictor boundary cases with the independent C helper."""
from pathlib import Path
import csv,sys,subprocess,hashlib,struct,math
root=Path(__file__).resolve().parent;tool=Path(sys.argv[1]).resolve();rows=[]
for precision in range(2,17):
 for channels in (1,3):
  for predictor in ((1,2,3,4,5,6,7) if precision==12 else (1,7)):
   for point in (0,precision-1):
    separate=int(predictor%2==1 and point==0)
    name=f'p{precision}-c{channels}-d{predictor}-t{point}-s{separate}.jpg'
    subprocess.run([str(tool),str(root/name),str(precision),str(channels),str(predictor),str(point),str(separate)],check=True)
    rows.append([name,precision,channels,predictor,point,separate,17,11,"gray" if channels==1 else "rgb"])
for precision in range(2,17):
 name=f'p{precision}-c3-ycc.jpg'
 subprocess.run([str(tool),str(root/name),str(precision),'3','1','0','0','ycc'],check=True)
 words=struct.unpack('<'+'H'*(17*11*3),(root/(name+'.raw')).read_bytes())
 maximum=(1<<precision)-1;midpoint=1<<(precision-1);rgba=bytearray()
 for at in range(0,len(words),3):
  y,cb,cr=words[at:at+3];cb-=midpoint;cr-=midpoint
  rgb=(y+1.402*cr,y-.344136286*cb-.714136286*cr,y+1.772*cb)
  rgba.extend(max(0,min(255,math.floor(v*255/maximum+.5)))for v in rgb);rgba.append(255)
 (root/(name+'.rgba')).write_bytes(rgba)
 rows.append([name,precision,3,1,0,0,17,11,'ycc'])
with(root/'manifest.csv').open('w')as f:
 writer=csv.writer(f,lineterminator='\n');writer.writerow(['file','precision','channels','predictor','point','separate','width','height','color']);writer.writerows(rows)
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n'for p in sorted(root.iterdir())if p.suffix in ('.jpg','.raw','.rgba')))
print(len(rows),'independently encoded/decoded JPEGs')
