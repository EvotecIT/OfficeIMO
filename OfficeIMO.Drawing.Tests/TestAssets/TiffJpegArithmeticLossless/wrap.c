#include <tiffio.h>
#include <stdio.h>
#include <stdlib.h>
int main(int argc,char**argv) {
 if(argc!=6)return 1;
 FILE*f=fopen(argv[1],"rb");if(!f)return 2;fseek(f,0,SEEK_END);long n=ftell(f);rewind(f);
 unsigned char*p=malloc(n);if(!p||fread(p,1,n,f)!=(size_t)n)return 3;fclose(f);
 TIFF*t=TIFFOpen(argv[2],argv[5]);if(!t)return 4;int bits=atoi(argv[3]),channels=atoi(argv[4]);
 TIFFSetField(t,TIFFTAG_IMAGEWIDTH,19);TIFFSetField(t,TIFFTAG_IMAGELENGTH,11);
 TIFFSetField(t,TIFFTAG_BITSPERSAMPLE,bits);TIFFSetField(t,TIFFTAG_SAMPLESPERPIXEL,channels);
 TIFFSetField(t,TIFFTAG_PHOTOMETRIC,channels==1?PHOTOMETRIC_MINISBLACK:PHOTOMETRIC_RGB);
 TIFFSetField(t,TIFFTAG_PLANARCONFIG,PLANARCONFIG_CONTIG);TIFFSetField(t,TIFFTAG_COMPRESSION,COMPRESSION_JPEG);
 TIFFSetField(t,TIFFTAG_JPEGTABLESMODE,0);TIFFSetField(t,TIFFTAG_ROWSPERSTRIP,11);
 int ok=TIFFWriteRawStrip(t,0,p,n)==n;TIFFClose(t);free(p);return ok?0:5;
}
