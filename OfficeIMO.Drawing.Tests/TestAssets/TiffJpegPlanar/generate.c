#include <tiffio.h>
#include <stdio.h>
#include <stdlib.h>
/* Raw YCbCr plane producer and independent JPEG sample decoder. */
static void word(FILE*f,unsigned n){for(int i=0;i<4;i++)fputc((n>>(i*8))&255,f);}
int main(int argc,char**argv){
 if(argc!=7)return 2;
 int big=atoi(argv[2]),tile=atoi(argv[3]),tables=atoi(argv[4]),hs=atoi(argv[5]),vs=atoi(argv[6]);
 int w=67,h=35,sw=tile?32:w,sh=32;
 TIFF*t=TIFFOpen(argv[1],big?"wb":"wl");if(!t)return 3;
 TIFFSetField(t,256,w);TIFFSetField(t,257,h);TIFFSetField(t,258,8);TIFFSetField(t,277,3);TIFFSetField(t,262,6);TIFFSetField(t,284,2);TIFFSetField(t,259,7);TIFFSetField(t,530,hs,vs);TIFFSetField(t,531,1);
 float ref[6]={0,255,128,255,128,255};TIFFSetField(t,532,ref);
 TIFFSetField(t,TIFFTAG_JPEGQUALITY,95);TIFFSetField(t,TIFFTAG_JPEGTABLESMODE,tables);TIFFSetField(t,TIFFTAG_JPEGCOLORMODE,JPEGCOLORMODE_RAW);
 if(tile){TIFFSetField(t,322,sw);TIFFSetField(t,323,sh);}else TIFFSetField(t,278,sh);
 unsigned char buf[67*32];
 for(int p=0;p<3;p++)for(int y=0;y<h;y+=sh)for(int x=0;x<w;x+=sw){
  int hh=p?hs:1,vv=p?vs:1,cw=(sw+hh-1)/hh,ch=((tile?sh:(h-y<sh?h-y:sh))+vv-1)/vv;
  for(int yy=0;yy<ch;yy++)for(int xx=0;xx<cw;xx++)buf[yy*sw+xx]=p?80+(x/hh+xx+2*(y/vv+yy)+p*17)%90:32+(x+xx+2*(y+yy))%185;
  if((tile?TIFFWriteEncodedTile(t,TIFFComputeTile(t,x,y,0,p),buf,sw*ch):TIFFWriteEncodedStrip(t,TIFFComputeStrip(t,y,p),buf,sw*ch))<0)return 4;
 }
 TIFFClose(t);t=TIFFOpen(argv[1],"r");if(!t)return 5;
 TIFFSetField(t,TIFFTAG_JPEGCOLORMODE,JPEGCOLORMODE_RAW);
 char path[4096];snprintf(path,sizeof(path),"%s.planes",argv[1]);FILE*f=fopen(path,"wb");if(!f)return 6;
 for(int p=0;p<3;p++)for(int y=0;y<h;y+=sh)for(int x=0;x<w;x+=sw){
  int hh=p?hs:1,vv=p?vs:1,cw=(sw+hh-1)/hh,ch=((tile?sh:(h-y<sh?h-y:sh))+vv-1)/vv;
  if((tile?TIFFReadEncodedTile(t,TIFFComputeTile(t,x,y,0,p),buf,sw*ch):TIFFReadEncodedStrip(t,TIFFComputeStrip(t,y,p),buf,sw*ch))!=sw*ch)return 7;
  word(f,p);word(f,x);word(f,y);word(f,cw);word(f,ch);for(int yy=0;yy<ch;yy++)fwrite(buf+yy*sw,1,cw,f);
 }
 fclose(f);TIFFClose(t);return 0;
}
