#include <tiffio.h>
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>
/* Independent LibTIFF container/compression producer. Three extra channels. */
static void store(unsigned char*p,double v,int bits,int floating) {
 if(!floating){if(bits==8)*p=(unsigned char)round(v*255);else{uint16_t n=(uint16_t)round(v*65535);memcpy(p,&n,2);}return;}
 if(bits==16){_Float16 h=(_Float16)v;memcpy(p,&h,2);}
 else if(bits==24){uint32_t n;if(isnan(v))n=0x7f0001;else if(v==0)n=0;else{int e;double f=frexp(v,&e);n=((e+62)<<16)|(uint32_t)round((f*2-1)*65536);}p[0]=n;p[1]=n>>8;p[2]=n>>16;}
 else if(bits==32){float f=(float)v;memcpy(p,&f,4);}else memcpy(p,&v,8);
}
int main(int argc,char**argv){
 if(argc!=10)return 2;
 int bits=atoi(argv[2]),floating=atoi(argv[3]),big=atoi(argv[4]),planar=atoi(argv[5]),tile=atoi(argv[6]),compression=atoi(argv[7]),photo=atoi(argv[8]);
 int base=photo==2?3:photo==5?4:1,n=base+3,w=19,h=17,size=bits/8,alphaIndex=-1,alphaKind=0;
 uint16_t extras[3];for(int i=0;i<3;i++){extras[i]=argv[9][i]-'0';if(extras[i]){alphaIndex=base+i;alphaKind=extras[i];}}
 TIFF*t=TIFFOpen(argv[1],big?"wb":"wl");if(!t)return 3;
 TIFFSetField(t,256,w);TIFFSetField(t,257,h);TIFFSetField(t,277,n);TIFFSetField(t,258,bits);TIFFSetField(t,339,floating?3:1);
 TIFFSetField(t,262,photo);TIFFSetField(t,284,planar);TIFFSetField(t,259,compression);TIFFSetField(t,338,3,extras);
 if(compression==5||compression==8)TIFFSetField(t,317,floating?3:2);
 if(tile){TIFFSetField(t,322,16);TIFFSetField(t,323,16);}else TIFFSetField(t,278,5);
 int sw=tile?16:w,sh=tile?16:5,sn=planar==2?1:n;unsigned char*buf=calloc(sw*sh*sn,size);
 for(int p=0;p<(planar==2?n:1);p++)for(int y=0;y<h;y+=sh)for(int x=0;x<w;x+=sw){
  memset(buf,0,sw*sh*sn*size);
  for(int yy=0;yy<sh&&y+yy<h;yy++)for(int xx=0;xx<sw&&x+xx<w;xx++)for(int c=0;c<sn;c++){
   int ch=planar==2?p:c;double a=((x+xx+y+yy)%5)/4.0,v=((x+xx)*3+(y+yy)*5+ch*7)%17/16.0;
   if(ch==alphaIndex)v=a;else if(ch>=base)v=floating?NAN:0.875;else if(alphaKind==1)v*=a;
   store(buf+((yy*sw+xx)*sn+c)*size,v,bits,floating);
  }
  if((tile?TIFFWriteEncodedTile(t,TIFFComputeTile(t,x,y,0,p),buf,sw*sh*sn*size):TIFFWriteEncodedStrip(t,TIFFComputeStrip(t,y,p),buf,sw*(y+sh>h?h-y:sh)*sn*size))<0)return 4;
 }
 free(buf);TIFFClose(t);return 0;
}
