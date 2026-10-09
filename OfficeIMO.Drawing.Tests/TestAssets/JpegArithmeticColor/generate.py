"""Independent CMYK/YCCK fixtures using libjpeg-turbo 3.2.0.
Build generate.c against the test-only native library, then pass its executable.
"""
from pathlib import Path
import csv,subprocess,sys,hashlib
root=Path(__file__).resolve().parent;rows=[]
for bits in (8,12):
 for color in (0,1):
  for progressive in (0,1):
   for sampling in ((0,1,2) if color else (0,)):
    for arithmetic in (0,1):
     for quality in (30,90):
      name=f'b{bits}-c{color}-p{progressive}-s{sampling}-a{arithmetic}-q{quality}.jpg'
      subprocess.run([sys.argv[1],str(root/name),str(bits),str(color),str(progressive),str(sampling),str(arithmetic),str(quality)],check=True)
      rows.append([name,bits,color,progressive,sampling,arithmetic,quality])
with (root/'manifest.csv').open('w') as f:
 w=csv.writer(f,lineterminator='\n');w.writerow(['name','bits','ycck','progressive','sampling','arithmetic','quality']);w.writerows(rows)
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256((root/(r[0]+suffix)).read_bytes()).hexdigest()+'  '+r[0]+suffix+'\n' for r in rows for suffix in ('','.nearest.cmyk16','.fancy.cmyk16')))
print(len(rows),'native JPEGs')
