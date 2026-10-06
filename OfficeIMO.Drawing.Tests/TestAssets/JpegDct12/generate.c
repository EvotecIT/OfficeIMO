/* Independent twelve-bit DCT producer and pixel oracle: libjpeg-turbo 3.2.0. */
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <jpeglib.h>
int main(int argc,char**argv){
 if(argc!=7)return 2;
 int color=atoi(argv[2]),quality=atoi(argv[3]),progressive=atoi(argv[4]),sampling=atoi(argv[5]),restart=atoi(argv[6]);
 const int w=35,h=19;int channels=color==0?1:3;
 J12SAMPLE input[35*19*3];
 for(int y=0;y<h;y++)for(int x=0;x<w;x++)for(int c=0;c<channels;c++){
  int value=(x*137+y*211+c*1103+((x/5+y/3)%2)*897)&4095;
  if(x==0)value=0;if(x==w-1)value=4095;input[(y*w+x)*channels+c]=value;
 }
 struct jpeg_compress_struct c;struct jpeg_error_mgr err;c.err=jpeg_std_error(&err);jpeg_create_compress(&c);
 c.image_width=w;c.image_height=h;c.input_components=channels;c.in_color_space=channels==1?JCS_GRAYSCALE:JCS_RGB;jpeg_set_defaults(&c);
 c.data_precision=12;if(color==1)jpeg_set_colorspace(&c,JCS_RGB);jpeg_set_quality(&c,quality,FALSE);c.optimize_coding=TRUE;c.restart_in_rows=restart;
 if(color==2){c.comp_info[0].h_samp_factor=sampling==0?1:2;c.comp_info[0].v_samp_factor=sampling==2?2:1;}
 jpeg_scan_info scans[3];memset(scans,0,sizeof(scans));
 if(progressive)jpeg_simple_progression(&c);
 else if(restart&&channels==3){for(int i=0;i<3;i++){scans[i].comps_in_scan=1;scans[i].component_index[0]=i;scans[i].Se=63;}c.scan_info=scans;c.num_scans=3;}
 unsigned char*encoded=NULL;unsigned long length=0;jpeg_mem_dest(&c,&encoded,&length);jpeg_start_compress(&c,TRUE);
 while(c.next_scanline<c.image_height){J12SAMPROW row=input+c.next_scanline*w*channels;jpeg12_write_scanlines(&c,&row,1);}jpeg_finish_compress(&c);jpeg_destroy_compress(&c);
 FILE*f=fopen(argv[1],"wb");if(!f)return 3;fwrite(encoded,1,length,f);fclose(f);
 for(int fancy=0;fancy<=1;fancy++){
  struct jpeg_decompress_struct d;d.err=jpeg_std_error(&err);jpeg_create_decompress(&d);jpeg_mem_src(&d,encoded,length);jpeg_read_header(&d,TRUE);
  d.out_color_space=channels==1?JCS_GRAYSCALE:JCS_RGB;d.do_fancy_upsampling=fancy;d.dct_method=JDCT_ISLOW;jpeg_start_decompress(&d);
  char path[4096];snprintf(path,sizeof(path),"%s.%s.rgba",argv[1],fancy?"bilinear":"nearest");f=fopen(path,"wb");if(!f)return 3;
  J12SAMPLE buffer[35*3];
  while(d.output_scanline<d.output_height){J12SAMPROW row=buffer;jpeg12_read_scanlines(&d,&row,1);
   for(int x=0;x<w;x++){for(int ch=0;ch<3;ch++){int v=buffer[x*channels+(channels==1?0:ch)];fputc((v*255+2047)/4095,f);}fputc(255,f);}
  }
  fclose(f);jpeg_finish_decompress(&d);jpeg_destroy_decompress(&d);
 }
 free(encoded);return 0;
}
