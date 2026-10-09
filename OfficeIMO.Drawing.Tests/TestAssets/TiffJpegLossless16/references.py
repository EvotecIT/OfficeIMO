"""Reference TIFF/RGBA from independently decoded native JPEG words.
YCbCr uses floating RGB to avoid introducing an intermediate quantization.
"""
import math,struct
def reference_tiff(words,photo,n,kind,maximum):
 floating=photo==6;bits=64 if floating else 16
 count=(11 if kind else 10)+(1 if floating else 0);bitsat=8+2+count*12+4;pixelat=bitsat+n*2
 entries=[(256,4,1,35),(257,4,1,19),(258,3,n,(bits|(bits<<16)) if n==2 else bitsat),(259,3,1,1),(262,3,1,2 if photo==6 else photo),(273,4,1,pixelat),(277,3,1,n),(278,4,1,19),(279,4,1,len(words)*(8 if floating else 2)),(284,3,1,1)]
 if floating:entries.append((339,3,n,(3|(3<<16)) if n==2 else pixelat));pixelat+=n*2;entries=[(tag,kind_,count_,pixelat if tag==273 else value)for tag,kind_,count_,value in entries]
 if kind:entries.append((338,3,1,kind))
 return b'II*\0'+struct.pack('<I',8)+struct.pack('<H',count)+b''.join(struct.pack('<HHII',*e) for e in sorted(entries))+bytes(4)+struct.pack('<'+'H'*n,*([bits]*n))+(struct.pack('<'+'H'*n,*([3]*n)) if floating else b'')+struct.pack('<'+('d' if floating else 'H')*len(words),*[v/maximum if floating else round(v*65535/maximum) for v in words])
def write_references(root,name,photo,kind,precision):
 maximum=(1<<precision)-1;midpoint=1<<(precision-1)
 base=4 if photo==5 else 3 if photo in (2,6) else 1;n=4 if photo==5 else base+1
 raw=(root/(name+'.raw')).read_bytes();words=list(struct.unpack('<'+'H'*(len(raw)//2),raw));rgbwords=words.copy();rgba=bytearray()
 for i in range(35*19):
  p=words[i*n:(i+1)*n];a=p[-1] if kind else maximum
  if photo==6:
   y=p[0];cb=p[1]-midpoint;cr=p[2]-midpoint;r=y+cr*1.402;b=y+cb*1.772;g=(y-.299*r-.114*b)/.587
   p[:3]=[min(maximum,max(0,v)) for v in (r,g,b)];rgbwords[i*n:i*n+3]=p[:3]
  color=[min(1,p[c]/a) if a else 0 for c in range(base)] if kind==1 else [p[c]/maximum for c in range(base)]
  q=lambda v:min(255,max(0,math.floor(v*255+0.5)))
  if photo==5:color=[255-min(255,q(color[c])+q(color[3])) for c in range(3)]
  elif photo in (0,1):color=[q(1-color[0] if photo==0 else color[0])]*3
  else:color=[q(v) for v in color]
  rgba.extend(color+[q(a/maximum)])
 (root/(name+'.rgba')).write_bytes(rgba)
 (root/(name+'.reference.tif')).write_bytes(reference_tiff(rgbwords,photo,n,kind,maximum))
