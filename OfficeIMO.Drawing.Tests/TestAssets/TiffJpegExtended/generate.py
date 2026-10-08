from pathlib import Path
import csv,hashlib,subprocess,sys
root=Path(__file__).resolve().parent
exe=Path(sys.argv[1]).resolve()
manifest=(root.parent/'TiffJpeg/manifest.csv').read_text()
for row in list(csv.reader(manifest.splitlines()))[1:]:
 subprocess.run([str(exe),str(root/row[0]),*row[1:]],check=True)
(root/'manifest.csv').write_text(manifest)
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n' for p in sorted(root.glob('*.tif*'))))
