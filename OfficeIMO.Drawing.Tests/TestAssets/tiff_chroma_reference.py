"""Independent Pillow interpolation of native TIFF chroma sample planes."""
from PIL import Image

def interpolate(plane,w,height,h,v,position):
 # Pad one sample on each edge so Pillow's affine interpolator clamps the grid.
 ph=len(plane);pw=len(plane[0]);image=Image.new('F',(pw+2,ph+2))
 image.putdata([plane[min(ph-1,max(0,y-1))][min(pw-1,max(0,x-1))] for y in range(ph+2) for x in range(pw+2)])
 ox=1 if position==1 else 1.5-.5/h;oy=1 if position==1 else 1.5-.5/v
 return [round(x) for x in image.transform((w,height),Image.Transform.AFFINE,(1/h,0,ox,0,1/v,oy),Image.Resampling.BILINEAR).getdata()]
