/* Test-only full-file LibTIFF sample oracle. Retains packed-sample discrepancies. */
#include <tiffio.h>
#include <stdio.h>
#include <stdlib.h>
int main(int argc,char**argv){
 if(argc!=3)return 2;TIFF*t=TIFFOpen(argv[1],"r");if(!t)return 3;
 uint32_t w,h,sw,sh;uint16_t samples,planar,bits;
 TIFFGetField(t,256,&w);TIFFGetField(t,257,&h);TIFFGetField(t,258,&bits);TIFFGetField(t,277,&samples);TIFFGetField(t,284,&planar);
 if(bits!=12||w!=35||h!=19||samples>4)return 4;
 int tile=TIFFIsTiled(t);sw=w;
 if(tile){TIFFGetField(t,322,&sw);TIFFGetField(t,323,&sh);}else TIFFGetField(t,278,&sh);
 int channels=planar==2?1:samples;size_t rb=(sw*channels*12+7)/8;
 unsigned char*buffer=calloc(sh,rb);unsigned short*out=calloc(w*h*samples,2);if(!buffer||!out)return 5;
 for(int p=0;p<(planar==2?samples:1);p++)for(uint32_t y=0;y<h;y+=sh)for(uint32_t x=0;x<w;x+=sw){
  size_t rows=tile?sh:(h-y<sh?h-y:sh);
  tmsize_t n=tile?TIFFReadEncodedTile(t,TIFFComputeTile(t,x,y,0,p),buffer,rows*rb):TIFFReadEncodedStrip(t,TIFFComputeStrip(t,y,p),buffer,rows*rb);
  if(n<0)return 6;
  for(uint32_t yy=0;yy<rows&&y+yy<h;yy++)for(uint32_t xx=0;xx<sw&&x+xx<w;xx++)for(int c=0;c<channels;c++){
   int bit=(xx*channels+c)*12,at=yy*rb+bit/8;int v=bit%8?((buffer[at]&15)<<8)|buffer[at+1]:(buffer[at]<<4)|(buffer[at+1]>>4);
   out[((y+yy)*w+x+xx)*samples+(planar==2?p:c)]=v;
  }
 }
 FILE*f=fopen(argv[2],"wb");if(!f)return 7;for(unsigned i=0;i<w*h*samples;i++){fputc(out[i]&255,f);fputc(out[i]>>8,f);}fclose(f);free(out);free(buffer);TIFFClose(t);return 0;
}
