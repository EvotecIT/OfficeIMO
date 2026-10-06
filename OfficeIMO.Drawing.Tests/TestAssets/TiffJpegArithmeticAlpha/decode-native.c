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
 TIFFGetFieldDefaulted(t,TIFFTAG_SAMPLESPERPIXEL,&n);TIFFGetField(t,TIFFTAG_BITSPERSAMPLE,&bits);
 TIFFGetFieldDefaulted(t,TIFFTAG_PLANARCONFIG,&planar);
 if(w!=35||h!=19||(bits!=8&&bits!=12)||n<1||n>5)return 3;
 int tiled=TIFFIsTiled(t);sw=w;
 if(tiled){TIFFGetField(t,TIFFTAG_TILEWIDTH,&sw);TIFFGetField(t,TIFFTAG_TILELENGTH,&sh);}
 else TIFFGetField(t,TIFFTAG_ROWSPERSTRIP,&sh);
 if(sw<1||sh<1)return 4;
 uint16_t photo=0,hs=1,vs=1;TIFFGetField(t,TIFFTAG_PHOTOMETRIC,&photo);
 if(photo==6&&n==3){TIFFGetFieldDefaulted(t,TIFFTAG_YCBCRSUBSAMPLING,&hs,&vs);if(hs!=1||vs!=1){fprintf(stderr,"Oracle does not reconstruct packed subsampled YCbCr units.\n");TIFFClose(t);return 9;}}
 tmsize_t capacity=tiled?TIFFTileSize(t):TIFFStripSize(t);
 unsigned char*buffer=calloc(capacity,1);uint16_t*output=calloc((size_t)w*h*n,sizeof(uint16_t));if(!buffer||!output)return 5;
 int channels=planar==2?1:n;
 tmsize_t rowbytes=tiled?TIFFTileRowSize(t):TIFFScanlineSize(t);
 for(int plane=0;plane<(planar==2?n:1);plane++)for(uint32_t y=0;y<h;y+=sh)for(uint32_t x=0;x<w;x+=sw){
  uint32_t rows=tiled?sh:(h-y<sh?h-y:sh);
  memset(buffer,0,capacity);
  tmsize_t read=tiled?TIFFReadEncodedTile(t,TIFFComputeTile(t,x,y,0,plane),buffer,capacity):TIFFReadEncodedStrip(t,TIFFComputeStrip(t,y,plane),buffer,capacity);
  if(read<(tmsize_t)(rowbytes*rows))return 6;
  for(uint32_t yy=0;yy<rows&&y+yy<h;yy++)for(uint32_t xx=0;xx<sw&&x+xx<w;xx++)for(int c=0;c<channels;c++)
  {
   int sample=xx*channels+c;const unsigned char*row=buffer+yy*rowbytes;int at=sample*12/8;
   output[((y+yy)*w+x+xx)*n+(planar==2?plane:c)]=bits==8?row[sample]:sample%2?((row[at]&15)<<8)|row[at+1]:(row[at]<<4)|(row[at+1]>>4);
  }
 }
 FILE*f=fopen(argv[2],"wb");if(!f)return 7;int ok=1;for(size_t i=0;i<(size_t)w*h*n;i++){if(fputc(output[i]&255,f)==EOF)ok=0;if(bits==12&&fputc(output[i]>>8,f)==EOF)ok=0;}
 fclose(f);TIFFClose(t);free(buffer);free(output);return ok?0:8;
}
