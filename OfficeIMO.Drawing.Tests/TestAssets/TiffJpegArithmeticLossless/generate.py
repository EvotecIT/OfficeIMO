"""Wrap qualified JPEG samples using the external, test-only LibTIFF writer.
Build wrap.c against LibTIFF 4.7.2, then run: python3 generate.py /path/to/wrap
"""
from pathlib import Path
import csv,hashlib,subprocess,sys
root=Path(__file__).resolve().parent; jpeg=root.parent/'JpegArithmeticLossless';rows=[]
for row in csv.DictReader((jpeg/'manifest.csv').open()):
 if int(row['bits']) not in (8,12,16):continue
 for endian,mode in [('le','wl'),('be','wb')]:
  name=row['name']+'.'+endian+'.tif'
  subprocess.run([sys.argv[1],str(jpeg/row['name']),str(root/name),row['bits'],row['components'],mode],check=True)
  rows.append([name,row['bits'],row['components'],row['predictor'],row['point'],row['restart'],endian,row['name']])
with (root/'manifest.csv').open('w') as f:
 w=csv.writer(f,lineterminator='\n');w.writerow(['name','bits','components','predictor','point','restart','byteOrder','jpeg']);w.writerows(rows)
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256((root/r[0]).read_bytes()).hexdigest()+'  '+r[0]+'\n' for r in rows))
print(len(rows),'LibTIFF arithmetic lossless containers')
