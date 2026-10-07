"""Generate all precision cases with native decoding and fractional Pillow references.
Pass the compiled ../TiffJpegChroma16/decode.c executable and a scratch directory.
"""
from pathlib import Path
import csv,hashlib,shutil,subprocess,sys
root=Path(__file__).resolve().parent;tool=Path(sys.argv[1]).resolve();scratch=Path(sys.argv[2]).resolve();rows=[]
for precision in range(2,17):
 output=scratch/f'b{precision}'
 subprocess.run([sys.executable,str(root.parent/'TiffJpegChroma16/generate.py'),str(tool),str(precision),str(output),'fractional'],check=True)
 for p in output.iterdir():
  if p.suffix in('.tif','.rgba'):shutil.copy2(p,root/p.name)
 with(output/'manifest.csv').open()as f:rows.extend(list(csv.reader(f))[1:])
with(root/'manifest.csv').open('w')as f:
 w=csv.writer(f,lineterminator='\n');w.writerow(['name','width','height','horizontal','vertical','layout','position']);w.writerows(rows)
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n'for p in sorted(root.iterdir())if p.suffix in('.tif','.rgba')))
print(len(rows),'TIFF cases')
