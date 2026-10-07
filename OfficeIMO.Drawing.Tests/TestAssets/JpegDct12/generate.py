from pathlib import Path
import csv,subprocess,hashlib,sys
root=Path(__file__).resolve().parent;tool=Path(sys.argv[1]).resolve();rows=[]
for color in range(3):
 for quality in (1,75,100):
  for progressive in (0,1):
   for sampling in (range(3) if color==2 else [0]):
    for restart in (0,2):
     name=f'c{color}-q{quality}-p{progressive}-s{sampling}-r{restart}.jpg'
     subprocess.run([str(tool),str(root/name),str(color),str(quality),str(progressive),str(sampling),str(restart)],check=True)
     rows.append([name,color,quality,progressive,sampling,restart,35,19])
with(root/'manifest.csv').open('w')as f:
 w=csv.writer(f,lineterminator='\n');w.writerow(['file','color','quality','progressive','sampling','restart','width','height']);w.writerows(rows)
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n'for p in sorted(root.iterdir())if p.suffix in ('.jpg','.rgba')))
print(len(rows),'independent twelve-bit DCT cases')
