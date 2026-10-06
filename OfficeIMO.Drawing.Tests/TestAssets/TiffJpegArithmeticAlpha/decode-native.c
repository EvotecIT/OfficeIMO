/* Test-only full-file TIFF component oracle. No color or alpha projection. */
#include <tiffio.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
int main(int argc,char**argv) {
 if(argc!=3)return 1;
 TIFF*t=TIFFOpen(argv[1],"r");if(!t)return 2;
 uint32_t w=0,h=0,sw=0,sh=0;uint16_t n=0,bits=0,planar=0;
 TIFFGetField(t,TIFFTAG_IMAGEWIDTH,&w);TIFFGetField(t,TIFFTAG_IMAGELENGTH,&h);
 TIFFGetField(t,TIFFTAG_SAMPLESPERPIXEL,&n);TIFFGetField(t,TIFFTAG_BITSPERSAMPLE,&bits);
 TIFFGetFieldDefaulted(t,TIFFTAG_PLANARCONFIG,&planar);
 if(w!=35||h!=19||bits!=8||n<2||n>5)return 3;
 int tiled=TIFFIsTiled(t);sw=w;
 if(tiled){TIFFGetField(t,TIFFTAG_TILEWIDTH,&sw);TIFFGetField(t,TIFFTAG_TILELENGTH,&sh);}
 else TIFFGetField(t,TIFFTAG_ROWSPERSTRIP,&sh);
 if(sw<1||sh<1)return 4;
 tmsize_t capacity=tiled?TIFFTileSize(t):TIFFStripSize(t);
 unsigned char*buffer=malloc(capacity),*output=calloc((size_t)w*h*n,1);if(!buffer||!output)return 5;
 int channels=planar==2?1:n;
 for(int plane=0;plane<(planar==2?n:1);plane++)for(uint32_t y=0;y<h;y+=sh)for(uint32_t x=0;x<w;x+=sw){
  uint32_t rows=tiled?sh:(h-y<sh?h-y:sh);
  tmsize_t read=tiled?TIFFReadEncodedTile(t,TIFFComputeTile(t,x,y,0,plane),buffer,capacity):TIFFReadEncodedStrip(t,TIFFComputeStrip(t,y,plane),buffer,capacity);
  if(read<(tmsize_t)((size_t)sw*rows*channels))return 6;
  for(uint32_t yy=0;yy<rows&&y+yy<h;yy++)for(uint32_t xx=0;xx<sw&&x+xx<w;xx++)for(int c=0;c<channels;c++)
   output[((y+yy)*w+x+xx)*n+(planar==2?plane:c)]=buffer[(yy*sw+xx)*channels+c];
 }
 FILE*f=fopen(argv[2],"wb");if(!f)return 7;int ok=fwrite(output,1,(size_t)w*h*n,f)==(size_t)w*h*n;
 fclose(f);TIFFClose(t);free(buffer);free(output);return ok?0:8;
}
