/* Test-only independent TIFF producer/decoder: LibTIFF 4.7.2 + libjpeg-turbo 3.2.0.
   Build with cc generate.c -I<libtiff>/include -L<libtiff>/lib -ltiff -o generate. */
#include <tiffio.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <jpeglib.h>
static int sample(int x,int y,int c,int kind){
 if(x>34)x=34;if(y>18)y=18;
 int v=(x*137+y*211+c*1103+((x/5+y/3)%2)*897)&4095;
 if(x==0)v=0;if(x==34)v=4095;
 if(kind==3||kind==4){int a=(x+y)%7==0?1:((x*71+y*157)&4095);if(c==3)return a;if(kind==3)v=v*a/4095;}
 return v;
}
static void put(unsigned char*b,int n,int v){for(int k=0;k<12;k++)if(v&(1<<(11-k)))b[(n*12+k)/8]|=128>>((n*12+k)%8);}
static int get(const unsigned char*b,int n){int bit=n*12,at=bit/8;return bit%8?((b[at]&15)<<8)|b[at+1]:(b[at]<<4)|(b[at+1]>>4);}
int main(int argc,char**argv){
 if(argc!=8)return 2;
 int kind=atoi(argv[2]),comp=atoi(argv[3]),be=atoi(argv[4]),planar=atoi(argv[5]),tile=atoi(argv[6]),tables=atoi(argv[7]);
 int photo=kind==0?0:kind==1?1:kind==5?6:kind==6?5:2,channels=kind<2?1:(kind==3||kind==4||kind==6)?4:3;
 int w=35,h=19,sw=tile?16:w,sh=tile||kind==5?16:8,sc=planar==2?1:channels;
 TIFF*t=TIFFOpen(argv[1],be?"wb":"wl");if(!t)return 3;
 TIFFSetField(t,256,w);TIFFSetField(t,257,h);TIFFSetField(t,258,12);TIFFSetField(t,277,channels);TIFFSetField(t,262,photo);TIFFSetField(t,284,planar);TIFFSetField(t,259,comp);
 if(kind==3||kind==4){uint16_t extra=kind==3?1:2;TIFFSetField(t,338,1,&extra);}
 if(tile){TIFFSetField(t,322,sw);TIFFSetField(t,323,sh);}else TIFFSetField(t,278,sh);
 if(comp==7){TIFFSetField(t,TIFFTAG_JPEGQUALITY,90);TIFFSetField(t,TIFFTAG_JPEGTABLESMODE,tables?1:0);if(photo==6){TIFFSetField(t,530,2,2);TIFFSetField(t,TIFFTAG_JPEGCOLORMODE,JPEGCOLORMODE_RGB);}}
 int rowbytes=(sw*sc*12+7)/8;unsigned char*buffer=calloc(sh,rowbytes);if(!buffer)return 4;
 for(int p=0;p<(planar==2?channels:1);p++)for(int y=0;y<h;y+=sh)for(int x=0;x<w;x+=sw){
  memset(buffer,0,sh*rowbytes);
  for(int yy=0;yy<sh;yy++)for(int xx=0;xx<sw;xx++)for(int c=0;c<sc;c++)put(buffer+yy*rowbytes,xx*sc+c,sample(x+xx,y+yy,planar==2?p:c,kind));
  tmsize_t result=tile?TIFFWriteEncodedTile(t,TIFFComputeTile(t,x,y,0,p),buffer,rowbytes*sh):TIFFWriteEncodedStrip(t,TIFFComputeStrip(t,y,p),buffer,rowbytes*(h-y<sh?h-y:sh));
  if(result<0)return 5;
 }
 TIFFClose(t);free(buffer);t=TIFFOpen(argv[1],"r");if(!t)return 6;
 if(photo==6)TIFFSetField(t,TIFFTAG_JPEGCOLORMODE,JPEGCOLORMODE_RGB);
 rowbytes=tile?(int)TIFFTileRowSize(t):(int)TIFFScanlineSize(t);buffer=calloc(sh,rowbytes);if(!buffer)return 7;
 unsigned short*values=calloc(w*h*channels,sizeof(unsigned short));
 for(int p=0;p<(planar==2?channels:1);p++)for(int y=0;y<h;y+=sh)for(int x=0;x<w;x+=sw){
  tmsize_t result=tile?TIFFReadEncodedTile(t,TIFFComputeTile(t,x,y,0,p),buffer,rowbytes*sh):TIFFReadEncodedStrip(t,TIFFComputeStrip(t,y,p),buffer,rowbytes*(h-y<sh?h-y:sh));
  if(result<0)return 8;
  for(int yy=0;yy<sh&&y+yy<h;yy++)for(int xx=0;xx<sw&&x+xx<w;xx++)for(int c=0;c<sc;c++)values[((y+yy)*w+x+xx)*channels+(planar==2?p:c)]=get(buffer+yy*rowbytes,xx*sc+c);
 }
 /* Preserve full LibTIFF output separately before cross-checking JPEG samples. */
 char nativepath[4096];snprintf(nativepath,sizeof(nativepath),"%s.libtiff.raw",argv[1]);
 FILE*nf=fopen(nativepath,"wb");if(!nf)return 9;
 for(int i=0;i<w*h*channels;i++){fputc(values[i]&255,nf);fputc(values[i]>>8,nf);}fclose(nf);
 if(comp==7){
  for(int p=0;p<(planar==2?channels:1);p++)for(int y=0;y<h;y+=sh)for(int x=0;x<w;x+=sw){
   uint32_t index=tile?TIFFComputeTile(t,x,y,0,p):TIFFComputeStrip(t,y,p);
   uint64_t*counts;TIFFGetField(t,tile?TIFFTAG_TILEBYTECOUNTS:TIFFTAG_STRIPBYTECOUNTS,&counts);
   unsigned char*encoded=malloc(counts[index]);
   tmsize_t size=tile?TIFFReadRawTile(t,index,encoded,counts[index]):TIFFReadRawStrip(t,index,encoded,counts[index]);if(size<0)return 10;
   struct jpeg_decompress_struct d;struct jpeg_error_mgr err;d.err=jpeg_std_error(&err);jpeg_create_decompress(&d);
   uint32_t ts=0;void*tb=NULL;if(TIFFGetField(t,TIFFTAG_JPEGTABLES,&ts,&tb)){jpeg_mem_src(&d,tb,ts);jpeg_read_header(&d,FALSE);}
   jpeg_mem_src(&d,encoded,size);jpeg_read_header(&d,TRUE);d.dct_method=JDCT_ISLOW;
   d.out_color_space=planar==2||channels==1?JCS_GRAYSCALE:photo==5?JCS_CMYK:JCS_RGB;
   jpeg_start_decompress(&d);J12SAMPLE*line=calloc(d.output_width*d.output_components,sizeof(J12SAMPLE));
   while(d.output_scanline<d.output_height){int yy=d.output_scanline;J12SAMPROW row=line;jpeg12_read_scanlines(&d,&row,1);
    if(y+yy<h)for(int xx=0;xx<sw&&x+xx<w;xx++)for(int c=0;c<sc;c++)values[((y+yy)*w+x+xx)*channels+(planar==2?p:c)]=line[xx*sc+c];
   }
   free(line);jpeg_finish_decompress(&d);jpeg_destroy_decompress(&d);free(encoded);
  }
 }
 char path[4096];snprintf(path,sizeof(path),"%s.raw",argv[1]);FILE*f=fopen(path,"wb");if(!f)return 9;
 for(int i=0;i<w*h*channels;i++){fputc(values[i]&255,f);fputc(values[i]>>8,f);}fclose(f);free(values);free(buffer);TIFFClose(t);return 0;
}
