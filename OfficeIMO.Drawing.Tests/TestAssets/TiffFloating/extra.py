from pathlib import Path
import itertools,subprocess,struct,json,hashlib,math
root=Path(__file__).resolve().parent;out=root/'corpus';manifest=json.loads((out/'manifest.json').read_text());cases=manifest['cases']
q=lambda v:math.floor(max(0,min(1,v))*255+.5)
for bits,big,photo,alpha in itertools.product([16,32,64],[0,1],[0,1,5],[1,2]):
 planar=big;tile=alpha==1;compression=8 if big else 5
 ident=f'f{bits}-'+('be' if big else 'le')+f'-photo{photo}-a{alpha}';p=out/ident;n=4 if photo==5 else 1
 subprocess.run([str(root/'generate'),str(p.with_suffix('.tif')),str(bits),str(big),str(int(planar)),str(int(tile)),str(compression),str(alpha),str(photo)],check=True,capture_output=True)
 subprocess.run([str(root/'decode'),str(p.with_suffix('.tif')),str(p.with_suffix('.raw'))],check=True,capture_output=True)
 values=[];rgba=[]
 for y in range(17):
  for x in range(19):
   a=((x+y)%5)/4;ch=[((x*3+y*5+c*7)%17)/16 for c in range(n)];values.extend([v*a if alpha==1 else v for v in ch]+[a]);ch=[0 if alpha==1 and a==0 else v for v in ch]
   if photo==5:rgb=[255-min(255,q(ch[c])+q(ch[3])) for c in range(3)]
   else:rgb=[q(1-ch[0] if photo==0 else ch[0])]*3
   rgba.extend(rgb+[q(a)])
 raw=struct.pack('<'+{16:'e',32:'f',64:'d'}[bits]*len(values),*values);assert p.with_suffix('.raw').read_bytes()==raw,ident;p.with_suffix('.rgba').write_bytes(bytes(rgba))
 cases.append(dict(id=ident,width=19,height=17,bits=bits,bigEndian=bool(big),planar=bool(planar),tiled=bool(tile),compression=compression,alpha=alpha,photometric=photo,tiffSha256=hashlib.sha256(p.with_suffix('.tif').read_bytes()).hexdigest(),rgbaSha256=hashlib.sha256(bytes(rgba)).hexdigest()))
(out/'manifest.json').write_text(json.dumps(manifest,indent=2)+'\n');print('Total',len(cases))
