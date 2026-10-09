#include <tiffio.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
static int value(int x,int y){return y%7==0?0:y%7==1?1:y%7==2?(x&1):((x/13+y/2)&1);}
int main(int argc,char**argv){
 if(argc!=8&&argc!=9)return 2;
 int bits=1,photo=atoi(argv[2]),big=atoi(argv[3]),tile=atoi(argv[4]),comp=atoi(argv[5]),opts=atoi(argv[6]),fill=atoi(argv[7]);
 int w=83,h=19,sw=tile?16:w,sh=tile?16:(argc==9?19:5),stride=(sw*bits+7)/8,max=(1<<bits)-1;
 TIFF*t=TIFFOpen(argv[1],big?"wb":"wl");if(!t)return 3;
 TIFFSetField(t,256,w);TIFFSetField(t,257,h);TIFFSetField(t,258,bits);TIFFSetField(t,277,1);TIFFSetField(t,262,photo);TIFFSetField(t,284,1);TIFFSetField(t,259,comp);TIFFSetField(t,266,fill);if(comp==3)TIFFSetField(t,292,opts);if(comp==4)TIFFSetField(t,293,opts);
 uint16_t r[16],g[16],b[16];for(int i=0;i<=max;i++){r[i]=i*65535/max;g[i]=(max-i)*65535/max;b[i]=(i*7%(max+1))*65535/max;}
 if(photo==3)TIFFSetField(t,320,r,g,b);
 if(tile){TIFFSetField(t,322,16);TIFFSetField(t,323,16);}else TIFFSetField(t,278,argc==9?0xFFFFFFFFU:(unsigned)sh);
 unsigned char buf[256];
 for(int y=0;y<h;y+=sh)for(int x=0;x<w;x+=sw){
  memset(buf,255,sizeof(buf));
  for(int yy=0;yy<sh&&y+yy<h;yy++)for(int xx=0;xx<sw&&x+xx<w;xx++){
   int v=value(x+xx,y+yy),shift=8-bits-xx*bits%8;
   buf[yy*stride+xx*bits/8]=(buf[yy*stride+xx*bits/8]&~(max<<shift))|(v<<shift);
  }
  int size=stride*(tile?sh:(h-y<sh?h-y:sh));
  if((tile?TIFFWriteEncodedTile(t,TIFFComputeTile(t,x,y,0,0),buf,size):TIFFWriteEncodedStrip(t,TIFFComputeStrip(t,y,0),buf,size))<0)return 4;
 }
 TIFFClose(t);t=TIFFOpen(argv[1],"r");if(!t)return 5;
 for(int y=0;y<h;y+=sh)for(int x=0;x<w;x+=sw){
  if((tile?TIFFReadEncodedTile(t,TIFFComputeTile(t,x,y,0,0),buf,sizeof(buf)):TIFFReadEncodedStrip(t,TIFFComputeStrip(t,y,0),buf,sizeof(buf)))<0)return 6;
  for(int yy=0;yy<sh&&y+yy<h;yy++)for(int xx=0;xx<sw&&x+xx<w;xx++)if(((buf[yy*stride+xx*bits/8]>>(8-bits-xx*bits%8))&max)!=(value(x+xx,y+yy)))return 7;
 }
 TIFFClose(t);return 0;
}
