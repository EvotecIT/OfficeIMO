from pathlib import Path
from PIL import Image
import hashlib, math, struct, subprocess, sys
root=Path(__file__).resolve().parent
native,jpeg,decoder=map(lambda x:Path(x).resolve(),sys.argv[1:4])
rows=['file,photometric,samples,alphaIndex,alphaKind,width,height,tolerance']
def project(name,raw,photo,n,ai,kind,w,h,values):
 rgba=bytearray()
 for i in range(w*h):
  pixel=values[i*n:(i+1)*n];a=pixel[ai] if ai>=0 else 1.0;color=pixel[:4 if photo==5 else 3 if photo in (2,6) else 1]
  if kind==1:color=[min(1,max(0,v/a)) if a>0 else 0 for v in color]
  q=lambda v:min(255,max(0,math.floor(v*255+0.5)))
  if photo==5:color=[255-min(255,q(color[c])+q(color[3])) for c in range(3)]
  elif photo in (0,1):color=[q(1-color[0] if photo==0 else color[0])]*3
  else:color=list(map(q,color))
  rgba.extend(color+[q(a)])
 (root/(name+'.rgba')).write_bytes(rgba)
 rows.append(f'{name},{photo},{n},{ai},{kind},{w},{h},{6 if name.startswith("j-") and kind==1 else 3 if name.startswith("j-") else 1}')
for mode,(bits,floating) in enumerate([(8,0),(16,0),(16,1),(24,1),(32,1),(64,1)]):
 for layout in range(8):
  big=layout&1;planar=1+((layout>>1)&1);tile=(layout>>2)&1
  for ci,extras in enumerate(['100','020','001','000']):
   photo=[0,1,2,5][(mode+ci)%4];base=4 if photo==5 else 3 if photo==2 else 1;n=base+3
   ai=next((base+i for i,c in enumerate(extras) if c!='0'),-1);kind=int(extras[ai-base]) if ai>=0 else 0
   compression=[1,5,8,32773][(layout+ci)%4]
   name=f'n-{mode}-l{layout}-e{extras}-p{photo}.tif';f=root/name
   subprocess.run([str(native),str(f),str(bits),str(floating),str(big),str(planar),str(tile),str(compression),str(photo),extras],check=True)
   subprocess.run([str(decoder),str(f),str(f)+'.raw'],check=True);raw=(root/(name+'.raw')).read_bytes()
   if not floating:values=[v/(255 if bits==8 else 65535) for v in (raw if bits==8 else struct.unpack('<'+'H'*(len(raw)//2),raw))]
   elif bits!=24:values=list(struct.unpack('<'+{16:'e',32:'f',64:'d'}[bits]*(len(raw)//(bits//8)),raw))
   else:
    values=[]
    for i in range(0,len(raw),3):
     v=int.from_bytes(raw[i:i+3],'little');e=(v>>16)&127;m=v&65535
     values.append(float('nan') if e==127 else (m*2**-78 if e==0 else (1+m/65536)*2**(e-63))*(-1 if v&0x800000 else 1))
   project(name,raw,photo,n,ai,kind,19,17,values)
for photo in [0,1,2,5,6]:
 for sub in ([1,2] if photo==6 else [1]):
  for layout in range(4):
   big=layout&1;planar=1+((layout>>1)&1);tile=layout&1;shared=(layout>>1)&1
   for extras in ['100','020','001','000']:
    base=4 if photo==5 else 3 if photo in (2,6) else 1;n=base+3
    ai=next((base+i for i,c in enumerate(extras) if c!='0'),-1);kind=int(extras[ai-base]) if ai>=0 else 0
    name=f'j-p{photo}-s{sub}-l{layout}-e{extras}.tif';f=root/name
    subprocess.run([str(jpeg),str(f),str(photo),str(big),str(tile),str(shared),str(planar),str(sub),'0','0',extras],check=True)
    data=(root/(name+'.planes')).read_bytes()
    if data:
     planes=[Image.new('L',(35,19)) for _ in range(n)];offset=0
     while offset<len(data):
      c,x,y,w,h=struct.unpack_from('<5I',data,offset);offset+=20;im=Image.frombytes('L',(w,h),data[offset:offset+w*h]);offset+=w*h
      if photo==6 and c in (1,2):
       w=min(w,(35-x+sub-1)//sub);h=min(h,(19-y+sub-1)//sub);im=im.crop((0,0,w,h)).resize((w*sub,h*sub),Image.Resampling.BILINEAR)
      planes[c].paste(im.crop((0,0,min(im.width,35-x),min(im.height,19-y))),(x,y))
     raw=bytes(c for px in zip(*(p.tobytes() for p in planes)) for c in px);(root/(name+'.raw')).write_bytes(raw)
    else:raw=(root/(name+'.raw')).read_bytes()
    if photo==6:
     rgb=Image.frombytes('YCbCr',(35,19),bytes(raw[i+c] for i in range(0,len(raw),n) for c in range(3))).convert('RGB').tobytes()
     raw=bytes(v for i in range(35*19) for v in (rgb[i*3:i*3+3]+raw[i*n+3:(i+1)*n]))
    project(name,raw,photo,n,ai,kind,35,19,[v/255 for v in raw])
(root/'manifest.csv').write_text('\n'.join(rows)+'\n')
(root/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n' for p in sorted(root.glob('*.tif*'))))
print(len(rows)-1,'fixtures')
