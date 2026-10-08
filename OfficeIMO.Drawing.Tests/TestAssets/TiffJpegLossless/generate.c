#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <tiffio.h>
#include <jpeglib.h>
/* Independent eight-bit SOF3 producer and decoder; no OfficeIMO code is used. */
int main(int argc,char **argv) {
 if(argc!=8)return 2;
 int photo=atoi(argv[2]),predictor=atoi(argv[3]),point=atoi(argv[4]),layout=atoi(argv[5]),separate=atoi(argv[6]),restart=atoi(argv[7]);
 int w=35,h=19,base=photo==5?4:photo==2?3:1,n=photo==5?4:base+1;
 int tile=layout&1,big=layout&2,planar=layout&4?2:1,sw=tile?16:w,sh=16;
 TIFF*t=TIFFOpen(argv[1],big?"wb":"wl");if(!t)return 3;
 TIFFSetField(t,256,w);TIFFSetField(t,257,h);TIFFSetField(t,258,8);TIFFSetField(t,277,n);
 TIFFSetField(t,262,photo);TIFFSetField(t,284,planar);TIFFSetField(t,259,7);
 if(n>base){uint16_t extra=2;TIFFSetField(t,338,1,&extra);}
 if(tile){TIFFSetField(t,322,sw);TIFFSetField(t,323,sh);}else TIFFSetField(t,278,sh);
 unsigned char output[35*19*4]={0};
 for(int plane=0;plane<(planar==2?n:1);plane++)for(int y=0;y<h;y+=sh)for(int x=0;x<w;x+=sw){
  int rows=tile?sh:(h-y<sh?h-y:sh),channels=planar==2?1:n;
  unsigned char pixels[35*16*4],decoded[35*16*4];
  for(int yy=0;yy<rows;yy++)for(int xx=0;xx<sw;xx++)for(int c=0;c<channels;c++){
   int ch=planar==2?plane:c,gx=x+xx,gy=y+yy;
   if(gx>=w)gx=w-1;if(gy>=h)gy=h-1;
   pixels[(yy*sw+xx)*channels+c]=(unsigned char)((gx*73+gy*151+ch*59+(gx^gy)*17)&255);
  }
  struct jpeg_compress_struct c;struct jpeg_error_mgr err;c.err=jpeg_std_error(&err);jpeg_create_compress(&c);
  c.image_width=sw;c.image_height=rows;c.input_components=channels;c.in_color_space=JCS_UNKNOWN;jpeg_set_defaults(&c);
  jpeg_enable_lossless(&c,predictor,point);c.restart_in_rows=restart;
  jpeg_scan_info scans[4];memset(scans,0,sizeof(scans));
  if(separate&&channels>1){for(int i=0;i<channels;i++){scans[i].comps_in_scan=1;scans[i].component_index[0]=i;scans[i].Ss=predictor;scans[i].Al=point;}c.scan_info=scans;c.num_scans=channels;}
  unsigned char*encoded=NULL;unsigned long length=0;jpeg_mem_dest(&c,&encoded,&length);jpeg_start_compress(&c,TRUE);
  while(c.next_scanline<c.image_height){JSAMPROW row=pixels+c.next_scanline*sw*channels;jpeg_write_scanlines(&c,&row,1);}
  jpeg_finish_compress(&c);jpeg_destroy_compress(&c);
  if((tile?TIFFWriteRawTile(t,TIFFComputeTile(t,x,y,0,plane),encoded,length):TIFFWriteRawStrip(t,TIFFComputeStrip(t,y,plane),encoded,length))<0)return 4;
  struct jpeg_decompress_struct d;d.err=jpeg_std_error(&err);jpeg_create_decompress(&d);jpeg_mem_src(&d,encoded,length);jpeg_read_header(&d,TRUE);
  d.jpeg_color_space=JCS_UNKNOWN;d.out_color_space=JCS_UNKNOWN;jpeg_start_decompress(&d);
  while(d.output_scanline<d.output_height){JSAMPROW row=decoded+d.output_scanline*sw*channels;jpeg_read_scanlines(&d,&row,1);}
  jpeg_finish_decompress(&d);jpeg_destroy_decompress(&d);
  for(int i=0;i<sw*rows*channels;i++)if(decoded[i]!=(pixels[i]&~((1<<point)-1)))return 5;
  if(x==0&&y==0&&plane==0){char path[4096];snprintf(path,sizeof(path),"%s.jpg",argv[1]);FILE*f=fopen(path,"wb");if(!f)return 6;fwrite(encoded,1,length,f);fclose(f);
   snprintf(path,sizeof(path),"%s.jpg.raw",argv[1]);f=fopen(path,"wb");if(!f)return 6;fwrite(decoded,1,sw*rows*channels,f);fclose(f);}
  for(int yy=0;yy<rows&&y+yy<h;yy++)for(int xx=0;xx<sw&&x+xx<w;xx++)for(int cc=0;cc<channels;cc++)output[((y+yy)*w+x+xx)*n+(planar==2?plane:cc)]=decoded[(yy*sw+xx)*channels+cc];
  free(encoded);
 }
 TIFFClose(t);char path[4096];snprintf(path,sizeof(path),"%s.raw",argv[1]);FILE*f=fopen(path,"wb");if(!f)return 6;fwrite(output,1,w*h*n,f);fclose(f);return 0;
}
