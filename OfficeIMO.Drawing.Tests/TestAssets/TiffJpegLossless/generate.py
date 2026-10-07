from pathlib import Path
import csv,hashlib,subprocess,sys,tempfile
root=Path(__file__).resolve().parent;rows=[]
for pred in range(1,8):
 for photo in (0,1,2,5):
  for layout in range(8):
   point=(0,1,3,7)[(pred+layout)%4];separate=(pred+layout)%2;restart=2 if layout&2 else 0
   name=f'p{photo}-d{pred}-l{layout}-t{point}-s{separate}-r{restart}.tif'
   subprocess.run([sys.argv[1],str(root/name),str(photo),str(pred),str(point),str(layout),str(separate),str(restart)],check=True)
   with tempfile.TemporaryDirectory() as tmp:
    decoded=Path(tmp)/'decoded.raw'
    subprocess.run([sys.argv[2],str(root/name),str(decoded)],check=True)
    assert decoded.read_bytes()==(root/(name+'.raw')).read_bytes(),name
   base=4 if photo==5 else 3 if photo==2 else 1;n=4 if photo==5 else base+1
   raw=(root/(name+'.raw')).read_bytes();rgba=bytearray()
   for i in range(35*19):
    p=raw[i*n:(i+1)*n]
    if photo==5:color=[255-min(255,p[c]+p[3]) for c in range(3)]
    elif photo==2:color=list(p[:3])
    else:color=[255-p[0] if photo==0 else p[0]]*3
    rgba.extend(color+[255 if photo==5 else p[-1]])
   (root/(name+'.rgba')).write_bytes(rgba)
   rows.append([name,photo,n,base if n>base else -1,2 if n>base else 0,35,19,0,16 if layout&1 else 35,16,1 if layout&4 else n,pred,point,separate,restart])
with (root/'manifest.csv').open('w') as f:
 w=csv.writer(f);w.writerow(['file','photometric','samples','alphaIndex','alphaKind','width','height','tolerance','jpegWidth','jpegHeight','jpegComponents','predictor','pointTransform','separateScans','restartRows']);w.writerows(rows)
files=sorted(p for p in root.iterdir() if p.suffix in ('.tif','.raw','.rgba','.jpg'))
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n' for p in files))
print(f'{len(rows)} fixtures generated, {len(files)} fixture/reference files')
