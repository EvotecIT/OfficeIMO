#include <tiffio.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
/* LibTIFF/libjpeg independently encode the TIFF and supply decoded reference samples. */
int main(int argc,char**argv){
 if(argc!=8)return 2;
 int photo=atoi(argv[2]),big=atoi(argv[3]),tile=atoi(argv[4]),tables=atoi(argv[5]),planar=atoi(argv[6]),sub=atoi(argv[7]);
 int samples=photo==5?4:photo==2||photo==6?3:1,w=35,h=19,sw=tile?16:w,sh=16;
 TIFF*t=TIFFOpen(argv[1],big?"wb":"wl");if(!t)return 3;
 TIFFSetField(t,256,w);TIFFSetField(t,257,h);TIFFSetField(t,258,8);TIFFSetField(t,277,samples);TIFFSetField(t,262,photo);TIFFSetField(t,284,planar);TIFFSetField(t,259,7);
 TIFFSetField(t,TIFFTAG_JPEGQUALITY,1);TIFFSetField(t,TIFFTAG_JPEGTABLESMODE,tables);
 if(photo==6){TIFFSetField(t,530,sub,sub);TIFFSetField(t,TIFFTAG_JPEGCOLORMODE,JPEGCOLORMODE_RGB);}
 if(tile){TIFFSetField(t,322,sw);TIFFSetField(t,323,sh);}else TIFFSetField(t,278,sh);
 unsigned char buf[35*16*4];
 for(int p=0;p<(planar==2?samples:1);p++)for(int y=0;y<h;y+=sh)for(int x=0;x<w;x+=sw){
  int n=planar==2?1:samples;
  for(int yy=0;yy<sh;yy++)for(int xx=0;xx<sw;xx++)for(int c=0;c<n;c++){
   int gx=x+xx<w?x+xx:w-1,gy=y+yy<h?y+yy:h-1,ch=planar==2?p:c;
   buf[(yy*sw+xx)*n+c]=(unsigned char)(20+(gx*3+gy*2+ch*49)%210);
  }
  int size=sw*(tile?sh:(h-y<sh?h-y:sh))*n;
  if((tile?TIFFWriteEncodedTile(t,TIFFComputeTile(t,x,y,0,p),buf,size):TIFFWriteEncodedStrip(t,TIFFComputeStrip(t,y,p),buf,size))<0)return 4;
 }
 TIFFClose(t);t=TIFFOpen(argv[1],"r");if(!t)return 5;
 if(photo==6)TIFFSetField(t,TIFFTAG_JPEGCOLORMODE,JPEGCOLORMODE_RGB);
 unsigned char out[35*19*4]={0};
 for(int p=0;p<(planar==2?samples:1);p++)for(int y=0;y<h;y+=sh)for(int x=0;x<w;x+=sw){
  if((tile?TIFFReadEncodedTile(t,TIFFComputeTile(t,x,y,0,p),buf,sizeof(buf)):TIFFReadEncodedStrip(t,TIFFComputeStrip(t,y,p),buf,sizeof(buf)))<0)return 6;
  int n=planar==2?1:samples;
  for(int yy=0;yy<sh&&y+yy<h;yy++)for(int xx=0;xx<sw&&x+xx<w;xx++)for(int c=0;c<n;c++)out[((y+yy)*w+x+xx)*samples+(planar==2?p:c)]=buf[(yy*sw+xx)*n+c];
 }
 TIFFClose(t);char path[4096];snprintf(path,sizeof(path),"%s.raw",argv[1]);FILE*f=fopen(path,"wb");if(!f)return 7;fwrite(out,1,w*h*samples,f);fclose(f);return 0;
}
