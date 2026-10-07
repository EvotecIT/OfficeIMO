"""Wrap native full-resolution JPEG samples as TIFF YCbCr; use LibTIFF wrap.c.
This constructs TIFF color interpretation over unchanged independently encoded
JPEG components. It is not independent TIFF producer/decoder acceptance.
"""
from pathlib import Path
import csv, hashlib, math, struct, subprocess, sys
root=Path(__file__).resolve().parent;rows=[]
for bits in range(2,17):
 maximum=(1<<bits)-1;midpoint=1<<(bits-1)
 for coding,folder,width in [('huffman','JpegLosslessPrecision',17),('arithmetic','JpegArithmeticLossless',19)]:
  name=f'p{bits}-c3-ycc.jpg' if coding=='huffman' else f'b{bits}-c3-p1-t0-r19.jpg'
  source=root.parent/folder/name;payload=source.read_bytes()
  if coding=='huffman':
   raw=Path(str(source)+'.raw').read_bytes();words=struct.unpack('<'+'H'*(len(raw)//2),raw)
  else:words=[((x*193+y*791+c*3191)^(x*y*53))&maximum for y in range(11)for x in range(width)for c in range(3)]
  rgba=bytearray()
  for at in range(0,len(words),3):
   y,cb,cr=words[at:at+3];cb-=midpoint;cr-=midpoint
   red=y+1.402*cr;blue=y+1.772*cb;green=(y-.299*red-.114*blue)/.587
   rgba.extend(max(0,min(255,math.floor(c*255/maximum+.5)))for c in (red,green,blue));rgba.append(255)
  for endian,mode in [('le','wl'),('be','wb')]:
   output=f'{coding}-b{bits}-{endian}.tif'
   subprocess.run([sys.argv[1],str(source),str(root/output),str(bits),'3',mode,str(width),'11','6'],check=True)
   (root/(output+'.rgba')).write_bytes(rgba)
   rows.append([output,bits,width,11,folder+'/'+name,hashlib.sha256(payload).hexdigest()])
with(root/'manifest.csv').open('w')as f:
 w=csv.writer(f,lineterminator='\n');w.writerow(['name','bits','width','height','source','sourceSha256']);w.writerows(rows)
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n' for p in sorted(root.iterdir())if p.suffix in('.tif','.rgba')))
print(len(rows),'YCbCr TIFFs')
