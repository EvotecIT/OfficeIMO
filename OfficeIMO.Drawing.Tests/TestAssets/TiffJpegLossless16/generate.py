from pathlib import Path
import csv,hashlib,math,struct,subprocess,sys
precision=int(sys.argv[2]) if len(sys.argv)>2 else 16
assert precision in (12,16)
maximum=(1<<precision)-1;midpoint=1<<(precision-1)
root=Path(sys.argv[3]).resolve() if len(sys.argv)>3 else Path(__file__).resolve().parent;root.mkdir(parents=True,exist_ok=True);rows=[]
def reference_tiff(words,photo,n,kind):
 count=11 if kind else 10;bitsat=8+2+count*12+4;pixelat=bitsat+n*2
 entries=[(256,4,1,35),(257,4,1,19),(258,3,n,(16|(16<<16)) if n==2 else bitsat),(259,3,1,1),(262,3,1,2 if photo==6 else photo),(273,4,1,pixelat),(277,3,1,n),(278,4,1,19),(279,4,1,len(words)*2),(284,3,1,1)]
 if kind:entries.append((338,3,1,kind))
 return b'II*\0'+struct.pack('<I',8)+struct.pack('<H',count)+b''.join(struct.pack('<HHII',*e) for e in entries)+bytes(4)+struct.pack('<'+'H'*n,*([16]*n))+struct.pack('<'+'H'*len(words),*[round(v*65535/maximum) for v in words])
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
   base=4 if photo==5 else 3 if photo in (2,6) else 1;n=4 if photo==5 else base+1
   raw=(root/(name+'.raw')).read_bytes();words=list(struct.unpack('<'+'H'*(len(raw)//2),raw));rgbwords=words.copy();rgba=bytearray()
   for i in range(35*19):
    p=words[i*n:(i+1)*n];a=p[-1] if kind else maximum
    if photo==6:
     y=p[0];cb=p[1]-midpoint;cr=p[2]-midpoint;r=y+cr*1.402;b=y+cb*1.772;g=(y-.299*r-.114*b)/.587
     p[:3]=[min(maximum,max(0,round(v))) for v in (r,g,b)];rgbwords[i*n:i*n+3]=p[:3]
    color=[min(1,p[c]/a) if a else 0 for c in range(base)] if kind==1 else [p[c]/maximum for c in range(base)]
    q=lambda v:min(255,max(0,math.floor(v*255+0.5)))
    if photo==5:color=[255-min(255,q(color[c])+q(color[3])) for c in range(3)]
    elif photo in (0,1):color=[q(1-color[0] if photo==0 else color[0])]*3
    else:color=[q(v) for v in color]
    rgba.extend(color+[q(a/maximum)])
   (root/(name+'.rgba')).write_bytes(rgba)
   (root/(name+'.reference.tif')).write_bytes(reference_tiff(rgbwords,photo,n,kind))
   rows.append([name,photo,n,base if kind else -1,kind,35,19,0,16 if layout&1 else 35,16,1 if layout&4 else n,pred,point,separate,restart])
with (root/'manifest.csv').open('w') as f:
 w=csv.writer(f);w.writerow(['file','photometric','samples','alphaIndex','alphaKind','width','height','tolerance','jpegWidth','jpegHeight','jpegComponents','predictor','pointTransform','separateScans','restartRows']);w.writerows(rows)
files=sorted(p for p in root.iterdir() if p.suffix in ('.tif','.raw','.rgba','.jpg'))
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n' for p in files))
print(f'{len(rows)} fixtures generated, {len(files)} fixture/reference files')
