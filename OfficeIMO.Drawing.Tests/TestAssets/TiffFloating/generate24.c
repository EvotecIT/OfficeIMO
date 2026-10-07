#include <tiffio.h>
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>
#include <fenv.h>
#include "imcd.h"
static int baseSamples=3;
static double value(int x,int y,int c,int alpha) {
 if(c==baseSamples) return ((x+y)%5)/4.0;
 double v=((x*3+y*5+c*7)%17)/16.0;
 return alpha==1?v*((x+y)%5)/4.0:v;
}
static void store(unsigned char*p,double v,int bits) {
 if(bits==24){float f=(float)v;if(imcd_float24_encode((const uint8_t*)&f,4,p,1,FE_TONEAREST)!=4)abort();}else if(bits==16){_Float16 h=(_Float16)v;memcpy(p,&h,2);}else if(bits==32){float f=(float)v;memcpy(p,&f,4);}else memcpy(p,&v,8);
}
int main(int argc,char**argv){
 if(argc!=8 && argc!=9)return 2;
 int bits=atoi(argv[2]),big=atoi(argv[3]),planar=atoi(argv[4]),tile=atoi(argv[5]),compression=atoi(argv[6]),alpha=atoi(argv[7]);
 int photo=argc==9?atoi(argv[8]):2;baseSamples=photo==2?3:photo==5?4:1;
 int w=19,h=17,n=baseSamples+1,size=bits/8,predictor=(compression==5||compression==8)?3:1;
 TIFF*t=TIFFOpen(argv[1],big?"wb":"wl");if(!t)return 3;
 TIFFSetField(t,TIFFTAG_IMAGEWIDTH,w);TIFFSetField(t,TIFFTAG_IMAGELENGTH,h);
 TIFFSetField(t,TIFFTAG_SAMPLESPERPIXEL,n);TIFFSetField(t,TIFFTAG_BITSPERSAMPLE,bits);TIFFSetField(t,TIFFTAG_SAMPLEFORMAT,SAMPLEFORMAT_IEEEFP);
 TIFFSetField(t,TIFFTAG_PHOTOMETRIC,photo);TIFFSetField(t,TIFFTAG_PLANARCONFIG,planar?2:1);
 TIFFSetField(t,TIFFTAG_COMPRESSION,compression);if(predictor==3)TIFFSetField(t,TIFFTAG_PREDICTOR,3);
 uint16_t extra=alpha;TIFFSetField(t,TIFFTAG_EXTRASAMPLES,1,&extra);
 if(tile){TIFFSetField(t,TIFFTAG_TILEWIDTH,16);TIFFSetField(t,TIFFTAG_TILELENGTH,16);}else TIFFSetField(t,TIFFTAG_ROWSPERSTRIP,5);
 int sw=tile?16:w,sh=tile?16:5,sn=planar?1:n;unsigned char*buf=calloc(sw*sh*sn,size);
 for(int p=0;p<(planar?n:1);p++)for(int y=0;y<h;y+=sh)for(int x=0;x<w;x+=sw){
  memset(buf,0,sw*sh*sn*size);
  for(int yy=0;yy<sh&&y+yy<h;yy++)for(int xx=0;xx<sw&&x+xx<w;xx++)for(int c=0;c<sn;c++)store(buf+((yy*sw+xx)*sn+c)*size,value(x+xx,y+yy,planar?p:c,alpha),bits);
  tmsize_t count=tile?TIFFWriteEncodedTile(t,TIFFComputeTile(t,x,y,0,p),buf,sw*sh*sn*size):TIFFWriteEncodedStrip(t,TIFFComputeStrip(t,y,p),buf,sw*(y+sh>h?h-y:sh)*sn*size);
  if(count<0)return 4;
 }
 free(buf);TIFFClose(t);return 0;
}
