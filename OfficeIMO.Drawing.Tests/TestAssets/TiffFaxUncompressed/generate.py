"""Specification-authored T.4/T.6 fixtures; no OfficeIMO encoder is used."""
from pathlib import Path
import struct, hashlib
root=Path(__file__).resolve().parent
# T.4 terminating codes for lengths 0..8, sufficient for the resumed runs.
white=['00110101','000111','0111','1000','1011','1100','1110','1111','10011']
black=['0000110111','010','11','10','011','0011','0010','00011','000101']
def literal(values, tail, color):
 s='';zeros=0
 for v in values:
  s+=str(v);zeros=0 if v else zeros+1
  if zeros==5:s+='1';zeros=0
 return s+'0'*(6+tail)+'1'+str(color)
def row(y, two):
 entry='0000001111' if two else '000000001111'
 tail=y%5;color=(y//5)%2
 prefix=([0]*5+[1,0,1]) if y%2==0 else [1,0,1,0,0,0,0,1]
 # Every exit length and color, followed by ordinary coded runs.
 samples=prefix+[0]*tail+[color]*(8-tail)+[1-color]*8+[color]*8
 encoded=entry+literal(prefix,tail,color)
 runs=[8-tail,8,8,0]
 if two:encoded+='001'+(black if color else white)[runs[0]]+(white if color else black)[8]+'001'+(black if color else white)[8]+(white if color else black)[0]
 else:encoded+=''.join((black if (color+i)%2 else white)[n] for i,n in enumerate(runs[:3]))
 if y==10: # Fully literal row; exit must still be consumed at the row boundary.
  samples=[0]*10+[1]*7+[0]*4+[1,0]*5+[0]
  encoded=entry+literal(samples,0,0)
 return encoded,samples
rows=['file,compression,options,photometric,bigEndian,tiled,fillOrder']
for compression,options in [(3,2),(3,3),(3,6),(3,7),(4,2)]:
 for big in [0,1]:
  for photo in [0,1]:
   for tiled in [0,1]:
    for fill in [1,2]:
     chunks=[];pixels=[]
     for start in range(0,12,16 if tiled else 6):
      bits=''
      for local in range(16 if tiled else min(6,12-start)):
       y=start+local;two=compression==4 or bool(options&1 and local%2)
       line,samples=row(y%11,two)
       if compression==3:
        if options&4:bits+='0'*((-len(bits)-12)%8)
        bits+='000000000001'
        if options&1:bits+=str(int(not two))
       bits+=line
       if y<12:pixels+=samples
      bits+=('000000000001'*2 if compression==4 else ('000000000001'+('1' if options&1 else ''))*6)
      bits+='0'*(-len(bits)%8)
      data=bytes(int(bits[i:i+8],2) for i in range(0,len(bits),8))
      if fill==2:data=bytes(int(f'{b:08b}'[::-1],2) for b in data)
      chunks.append(data)
     endian='>' if big else '<';pack=lambda fmt,*v:struct.pack(endian+fmt,*v)
     tags={256:(4,[32]),257:(4,[12]),258:(3,[1]),259:(3,[compression]),262:(3,[photo]),266:(3,[fill]),277:(3,[1]),284:(3,[1]),292 if compression==3 else 293:(4,[options])}
     offsetsTag,countsTag=(324,325) if tiled else (273,279)
     if tiled:tags.update({322:(4,[32]),323:(4,[16])})
     else:tags[278]=(4,[6])
     tags[offsetsTag]=(4,[0]*len(chunks));tags[countsTag]=(4,list(map(len,chunks)))
     base=8+2+12*len(tags)+4
     extraSize=sum(4*len(v) for t,v in tags.values() if len(v)>1)
     off=base+extraSize;offsets=[]
     for c in chunks:offsets.append(off);off+=len(c)
     tags[offsetsTag]=(4,offsets);entries=b'';extra=b''
     for tag,(typ,values) in sorted(tags.items()):
      v=pack(('H' if typ==3 else 'I')*len(values),*values)
      if len(v)>4:field=pack('I',base+len(extra));extra+=v
      else:field=v.ljust(4,b'\0')
      entries+=pack('HHI',tag,typ,len(values))+field
     name=f'c{compression}-o{options}-p{photo}-be{big}-tile{tiled}-f{fill}.tif'
     (root/name).write_bytes((b'MM' if big else b'II')+pack('HI',42,8)+pack('H',len(tags))+entries+pack('I',0)+extra+b''.join(chunks))
     (root/(name+'.rgb')).write_bytes(bytes((255*(1-v if photo==0 else v)) for v in pixels for _ in range(3)))
     rows.append(f'{name},{compression},{options},{photo},{big},{tiled},{fill}')
(root/'manifest.csv').write_text('\n'.join(rows)+'\n')
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n' for p in sorted(root.glob('*.tif*'))))
print(len(rows)-1)
