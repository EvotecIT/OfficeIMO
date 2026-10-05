#include <tiffio.h>
#include <stdio.h>
#include <stdlib.h>
#include <jpeglib.h>
#include <string.h>
static void word(FILE*f,unsigned n){for(int i=0;i<4;i++)fputc((n>>(i*8))&255,f);}
static int imageWidth,imageHeight,horizontal,vertical;
static unsigned char value(int p,int x,int y){
 if(p && (x >= (imageWidth+horizontal-1)/horizontal || y >= (imageHeight+vertical-1)/vertical))return p==1?0:255;
 return p?70+(x+2*y+p*17)%110:32+(x+2*y)%185;
}
int main(int argc,char**argv){
 if(argc!=9)return 2;
 int big=atoi(argv[2]),tile=atoi(argv[3]),tables=atoi(argv[4]),hs=atoi(argv[5]),vs=atoi(argv[6]),planar=atoi(argv[7]);
 int w=tables?68:67,h=tables?36:35,sw=tile?32:w,sh=32;
 imageWidth=w;imageHeight=h;horizontal=hs;vertical=vs;
 TIFF*t=TIFFOpen(argv[1],big?"wb":"wl");if(!t)return 3;
 TIFFSetField(t,256,w);TIFFSetField(t,257,h);TIFFSetField(t,258,8);TIFFSetField(t,277,3);TIFFSetField(t,262,6);TIFFSetField(t,284,planar);TIFFSetField(t,259,7);TIFFSetField(t,530,hs,vs);TIFFSetField(t,531,atoi(argv[8]));
 float ref[6]={0,255,128,255,128,255};TIFFSetField(t,532,ref);
 TIFFSetField(t,TIFFTAG_JPEGQUALITY,95);TIFFSetField(t,TIFFTAG_JPEGTABLESMODE,tables);TIFFSetField(t,TIFFTAG_JPEGCOLORMODE,JPEGCOLORMODE_RAW);
 if(tile){TIFFSetField(t,322,sw);TIFFSetField(t,323,sh);}else TIFFSetField(t,278,sh);
 unsigned char buf[16384];
 for(int p=0;p<(planar==2?3:1);p++)for(int y=0;y<h;y+=sh)for(int x=0;x<w;x+=sw){
  int rows=tile?sh:(h-y<sh?h-y:sh),cw=(sw+hs-1)/hs,ch=(rows+vs-1)/vs,size;
  if(planar==2){int pw=p?cw:sw,ph=p?ch:rows;size=sw*ph;
   for(int yy=0;yy<ph;yy++)for(int xx=0;xx<pw;xx++)buf[yy*sw+xx]=value(p,x/(p?hs:1)+xx,y/(p?vs:1)+yy);
  }else{size=0;
   for(int yy=0;yy<ch;yy++)for(int xx=0;xx<cw;xx++){
    for(int dy=0;dy<vs;dy++)for(int dx=0;dx<hs;dx++)buf[size++]=value(0,x+xx*hs+dx,y+yy*vs+dy);
    buf[size++]=value(1,x/hs+xx,y/vs+yy);buf[size++]=value(2,x/hs+xx,y/vs+yy);
   }
  }
  if((tile?TIFFWriteEncodedTile(t,TIFFComputeTile(t,x,y,0,p),buf,size):TIFFWriteEncodedStrip(t,TIFFComputeStrip(t,y,p),buf,size))<0)return 4;
 }
 TIFFClose(t);t=TIFFOpen(argv[1],"r");if(!t)return 5;TIFFSetField(t,TIFFTAG_JPEGCOLORMODE,JPEGCOLORMODE_RAW);
 char path[4096];snprintf(path,sizeof(path),"%s.planes",argv[1]);FILE*f=fopen(path,"wb");if(!f)return 6;
 for(int p=0;p<(planar==2?3:1);p++)for(int y=0;y<h;y+=sh)for(int x=0;x<w;x+=sw){
  unsigned strip=tile?TIFFComputeTile(t,x,y,0,p):TIFFComputeStrip(t,y,p);
  uint64_t length=TIFFGetStrileByteCount(t,strip);unsigned char*encoded=malloc(length);
  if((tile?TIFFReadRawTile(t,strip,encoded,length):TIFFReadRawStrip(t,strip,encoded,length))!=(tmsize_t)length)return 7;
  struct jpeg_decompress_struct d;struct jpeg_error_mgr err;d.err=jpeg_std_error(&err);jpeg_create_decompress(&d);
  unsigned count=0;void*table=NULL;
  if(TIFFGetField(t,347,&count,&table)){jpeg_mem_src(&d,table,count);if(jpeg_read_header(&d,FALSE)!=JPEG_HEADER_TABLES_ONLY)return 8;}
  jpeg_mem_src(&d,encoded,length);if(jpeg_read_header(&d,TRUE)!=JPEG_HEADER_OK)return 9;
  d.raw_data_out=TRUE;jpeg_start_decompress(&d);
  JSAMPARRAY groups[3];unsigned char*planes[3];
  for(int c=0;c<d.num_components;c++){
   jpeg_component_info*ci=&d.comp_info[c];groups[c]=(*d.mem->alloc_sarray)((j_common_ptr)&d,JPOOL_IMAGE,ci->width_in_blocks*8,ci->v_samp_factor*8);
   planes[c]=calloc(ci->downsampled_width,ci->downsampled_height);
  }
  unsigned block=0;
  while(d.output_scanline<d.output_height){
   if(jpeg_read_raw_data(&d,groups,d.max_v_samp_factor*8)==0)return 10;
   for(int c=0;c<d.num_components;c++){
    jpeg_component_info*ci=&d.comp_info[c];
    for(unsigned row=0;row<(unsigned)ci->v_samp_factor*8&&block*ci->v_samp_factor*8+row<ci->downsampled_height;row++)
     memcpy(planes[c]+(block*ci->v_samp_factor*8+row)*ci->downsampled_width,groups[c][row],ci->downsampled_width);
   }
   block++;
  }
  for(int c=0;c<d.num_components;c++){
   jpeg_component_info*ci=&d.comp_info[c];word(f,planar==2?p:c);word(f,x);word(f,y);word(f,ci->downsampled_width);word(f,ci->downsampled_height);
   fwrite(planes[c],1,ci->downsampled_width*ci->downsampled_height,f);free(planes[c]);
  }
  jpeg_finish_decompress(&d);jpeg_destroy_decompress(&d);free(encoded);
 }
 fclose(f);TIFFClose(t);return 0;
}
