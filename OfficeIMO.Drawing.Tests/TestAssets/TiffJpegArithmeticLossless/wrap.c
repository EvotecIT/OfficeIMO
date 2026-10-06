#include <tiffio.h>
#include <stdio.h>
#include <stdlib.h>
int main(int argc,char**argv) {
 if(argc!=6 && argc!=8 && argc!=9)return 1;
 FILE*f=fopen(argv[1],"rb");if(!f)return 2;fseek(f,0,SEEK_END);long n=ftell(f);rewind(f);
 unsigned char*p=malloc(n);if(!p||fread(p,1,n,f)!=(size_t)n)return 3;fclose(f);
 TIFF*t=TIFFOpen(argv[2],argv[5]);if(!t)return 4;int bits=atoi(argv[3]),channels=atoi(argv[4]);
 int width=argc>=8?atoi(argv[6]):19,height=argc>=8?atoi(argv[7]):11;
 TIFFSetField(t,TIFFTAG_IMAGEWIDTH,width);TIFFSetField(t,TIFFTAG_IMAGELENGTH,height);
 TIFFSetField(t,TIFFTAG_BITSPERSAMPLE,bits);TIFFSetField(t,TIFFTAG_SAMPLESPERPIXEL,channels);
 int photo=argc==9?atoi(argv[8]):channels==1?PHOTOMETRIC_MINISBLACK:PHOTOMETRIC_RGB;
 TIFFSetField(t,TIFFTAG_PHOTOMETRIC,photo);
 if(photo==PHOTOMETRIC_YCBCR){
  float maximum=(float)((1<<bits)-1),midpoint=(float)(1<<(bits-1));
  float reference[6]={0,maximum,midpoint,maximum,midpoint,maximum};
  TIFFSetField(t,TIFFTAG_YCBCRSUBSAMPLING,1,1);TIFFSetField(t,TIFFTAG_REFERENCEBLACKWHITE,reference);
 }
 TIFFSetField(t,TIFFTAG_PLANARCONFIG,PLANARCONFIG_CONTIG);TIFFSetField(t,TIFFTAG_COMPRESSION,COMPRESSION_JPEG);
 TIFFSetField(t,TIFFTAG_JPEGTABLESMODE,0);TIFFSetField(t,TIFFTAG_ROWSPERSTRIP,height);
 int ok=TIFFWriteRawStrip(t,0,p,n)==n;TIFFClose(t);free(p);return ok?0:5;
}
