"""Independent sequential arithmetic JPEG corpus using libjpeg-turbo 3.2.0."""
from pathlib import Path
import subprocess,sys,csv,hashlib
root=Path(__file__).resolve().parent;rows=[]
for precision in (8,12):
 for color in (0,1,2):
  for sampling in ((0,1,2) if color==2 else (0,)):
   for quality in (1,75,100):
    for restart in (0,2,3):
     name=f'b{precision}-c{color}-q{quality}-s{sampling}-r{restart}.jpg'
     subprocess.run([sys.argv[1],str(root/name),str(color),str(quality),'0',str(sampling),str(restart),str(precision)],check=True)
     rows.append([name,precision,color,quality,sampling,restart])
with (root/'manifest.csv').open('w')as f:
 w=csv.writer(f,lineterminator='\n');w.writerow(['file','precision','color','quality','sampling','restartMcus']);w.writerows(rows)
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256((root/(r[0]+suffix)).read_bytes()).hexdigest()+'  '+r[0]+suffix+'\n' for r in rows for suffix in ('','.nearest.rgba','.bilinear.rgba','.tif')))
print(len(rows),'independent arithmetic JPEGs')
