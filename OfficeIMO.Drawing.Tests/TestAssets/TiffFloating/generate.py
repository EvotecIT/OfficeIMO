from pathlib import Path
import itertools,subprocess,struct,json,hashlib,math
root=Path(__file__).resolve().parent;out=root/'corpus';out.mkdir(exist_ok=True);cases=[]
for bits,big,planar,tile,compression,alpha in itertools.product([16,32,64],[0,1],[0,1],[0,1],[1,5,8,32773],[1,2]):
 ident=f'f{bits}-'+('be' if big else 'le')+f'-planar{planar}-tile{tile}-c{compression}-a{alpha}';p=out/ident
 subprocess.run([str(root/'generate'),str(p.with_suffix('.tif')),str(bits),str(big),str(planar),str(tile),str(compression),str(alpha)],check=True,capture_output=True)
 subprocess.run([str(root/'decode'),str(p.with_suffix('.tif')),str(p.with_suffix('.raw'))],check=True,capture_output=True)
 values=[];rgba=[]
 for y in range(17):
  for x in range(19):
   a=((x+y)%5)/4;channels=[((x*3+y*5+c*7)%17)/16 for c in range(3)]
   values.extend([v*a if alpha==1 else v for v in channels]+[a])
   rgba.extend([math.floor((0 if alpha==1 and a==0 else v)*255+.5) for v in channels]+[math.floor(a*255+.5)])
 raw=struct.pack('<'+{16:'e',32:'f',64:'d'}[bits]*len(values),*values);assert p.with_suffix('.raw').read_bytes()==raw,ident
 p.with_suffix('.rgba').write_bytes(bytes(rgba))
 cases.append(dict(id=ident,width=19,height=17,bits=bits,bigEndian=bool(big),planar=bool(planar),tiled=bool(tile),compression=compression,alpha=alpha,tiffSha256=hashlib.sha256(p.with_suffix('.tif').read_bytes()).hexdigest(),rgbaSha256=hashlib.sha256(bytes(rgba)).hexdigest()))
(out/'manifest.json').write_text(json.dumps({'producer':'libtiff','cases':cases},indent=2)+'\n');print('Independent encoded/decoded cases:',len(cases))
