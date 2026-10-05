from pathlib import Path
from PIL import Image
import hashlib,struct,subprocess,sys
root=Path(__file__).resolve().parent
exe=Path(sys.argv[1]).resolve()
rows=['file,bigEndian,tiled,tables,horizontal,vertical']
for h,v in [(1,1),(2,1),(2,2),(4,1),(4,2),(4,4)]:
 for big in (0,1):
  for tiled in (0,1):
   for tables in (0,3):
    name=f'h{h}-v{v}-be{big}-t{tiled}-q{tables}.tif';file=root/name
    subprocess.run([str(exe),str(file),str(big),str(tiled),str(tables),str(h),str(v)],check=True)
    data=file.with_suffix('.tif.planes').read_bytes();offset=0;planes=[Image.new('L',(67,35)) for _ in range(3)]
    while offset<len(data):
     plane,x,y,w,height=struct.unpack_from('<5I',data,offset);offset+=20
     image=Image.frombytes('L',(w,height),data[offset:offset+w*height]);offset+=w*height
     if plane:image=image.resize((w*h,height*v),Image.Resampling.BILINEAR)
     planes[plane].paste(image.crop((0,0,min(image.width,67-x),min(image.height,35-y))),(x,y))
    (root/(name+'.rgb')).write_bytes(Image.merge('YCbCr',planes).convert('RGB').tobytes())
    rows.append(f'{name},{big},{tiled},{tables},{h},{v}')
(root/'manifest.csv').write_text('\n'.join(rows)+'\n')
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n' for p in sorted(root.glob('*.tif*'))))
print(len(rows)-1,'independently encoded/decoded planar TIFF fixtures')
