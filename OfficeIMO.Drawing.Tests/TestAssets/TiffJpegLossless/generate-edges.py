# T.81-authored constant samples; independently decoded by libjpeg-turbo djpeg.
from pathlib import Path
import csv,hashlib,struct,subprocess,sys
root=Path(__file__).resolve().parent/'edges';root.mkdir(exist_ok=True);rows=[]
for h,v in [(2,1),(1,2),(2,2),(4,2)]:
 for w,height in [(4,4),(5,3)]:
  for separate in (False,True):
   b=bytearray(b'\xff\xd8')
   def segment(m,p): b.extend(bytes([255,m])+(len(p)+2).to_bytes(2,'big')+p)
   segment(195,bytes([8,0,height,0,w,3,1,(h<<4)|v,0,2,17,0,3,17,0]))
   segment(196,bytes([0,1]+[0]*15+[0]))
   def scan(ids,n):
    segment(218,bytes([len(ids)])+b''.join(bytes([i,0]) for i in ids)+bytes([1,0,0]))
    bits='0'*n;bits+='1'*((-len(bits))%8)
    for i in range(0,len(bits),8):
     q=int(bits[i:i+8],2);b.append(q)
     if q==255:b.append(0)
   if separate:
    for c in (1,2,3): scan([c],w*height if c==1 else ((w+h-1)//h)*((height+v-1)//v))
   else: scan([1,2,3],((w+h-1)//h)*((height+v-1)//v)*(h*v+2))
   b.extend(b'\xff\xd9');name=f'h{h}-v{v}-w{w}-h{height}-s{int(separate)}';(root/(name+'.jpg')).write_bytes(b)
   # Minimal classic TIFF, centered YCbCr; JPEG retains all entropy bytes.
   count=12;bitsat=8+2+count*12+4;jpegat=bitsat+6
   entries=[(256,4,1,w),(257,4,1,height),(258,3,3,bitsat),(259,3,1,7),(262,3,1,6),(273,4,1,jpegat),(277,3,1,3),(278,4,1,height),(279,4,1,len(b)),(284,3,1,1),(530,3,2,h|(v<<16)),(531,3,1,1)]
   t=b'II*\0'+struct.pack('<I',8)+struct.pack('<H',count)+b''.join(struct.pack('<HHII',*e) for e in entries)+bytes(4)+struct.pack('<HHH',8,8,8)+b
   if v<=h: (root/(name+'.tif')).write_bytes(t)
   subprocess.run([sys.argv[1],'-outfile',str(root/(name+'.ppm')),str(root/(name+'.jpg'))],check=True)
   ppm=(root/(name+'.ppm')).read_bytes();assert ppm==f'P6\n{w} {height}\n255\n'.encode()+bytes([128]*(w*height*3)),name
   rows.append([name,w,height,h,v,int(separate)])
with (root/'manifest.csv').open('w') as f:
 cw=csv.writer(f);cw.writerow(['name','width','height','horizontal','vertical','separateScans']);cw.writerows(rows)
files=sorted(p for p in root.iterdir() if p.suffix in ('.jpg','.tif','.ppm'))
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n' for p in files))
print(len(rows),'independently decoded subsampled edge cases')
