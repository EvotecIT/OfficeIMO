#include <tiffio.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
int main(int argc,char**argv){
 if(argc!=3)return 2;TIFF*t=TIFFOpen(argv[1],"r");if(!t)return 3;
 uint32_t w,h,tw,th;uint16_t bits,n,planar;
 TIFFGetField(t,TIFFTAG_IMAGEWIDTH,&w);TIFFGetField(t,TIFFTAG_IMAGELENGTH,&h);TIFFGetField(t,TIFFTAG_BITSPERSAMPLE,&bits);TIFFGetField(t,TIFFTAG_SAMPLESPERPIXEL,&n);TIFFGetField(t,TIFFTAG_PLANARCONFIG,&planar);
 int size=bits/8,sn=planar==2?1:n;unsigned char*out=calloc(w*h*n,size);
 if(TIFFIsTiled(t)){TIFFGetField(t,TIFFTAG_TILEWIDTH,&tw);TIFFGetField(t,TIFFTAG_TILELENGTH,&th);unsigned char*b=malloc(TIFFTileSize(t));
 for(int p=0;p<(planar==2?n:1);p++)for(int y=0;y<h;y+=th)for(int x=0;x<w;x+=tw){if(TIFFReadTile(t,b,x,y,0,p)<0)return 4;
 for(int yy=0;yy<th&&y+yy<h;yy++)for(int xx=0;xx<tw&&x+xx<w;xx++)for(int c=0;c<sn;c++)memcpy(out+(((y+yy)*w+x+xx)*n+(planar==2?p:c))*size,b+((yy*tw+xx)*sn+c)*size,size);}
 free(b);
 }else{unsigned char*b=malloc(TIFFScanlineSize(t));for(int p=0;p<(planar==2?n:1);p++)for(int y=0;y<h;y++){if(TIFFReadScanline(t,b,y,p)<0)return 5;for(int x=0;x<w;x++)for(int c=0;c<sn;c++)memcpy(out+((y*w+x)*n+(planar==2?p:c))*size,b+(x*sn+c)*size,size);}free(b);}
 FILE*f=fopen(argv[2],"wb");fwrite(out,size,w*h*n,f);fclose(f);free(out);TIFFClose(t);return 0;
}
