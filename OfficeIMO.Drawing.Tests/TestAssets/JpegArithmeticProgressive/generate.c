/* Independent progressive arithmetic DCT producer and pixel oracle: libjpeg-turbo 3.2.0. */
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <jpeglib.h>
int main(int argc,char**argv){
 if(argc!=8)return 2;
 int precision=atoi(argv[7]),maximum=(1<<precision)-1;
 int color=atoi(argv[2]),quality=atoi(argv[3]),progressive=atoi(argv[4]),sampling=atoi(argv[5]),restart=atoi(argv[6]);
 const int w=35,h=19;int channels=color==0?1:3;
 J12SAMPLE input[35*19*3];
 for(int y=0;y<h;y++)for(int x=0;x<w;x++)for(int c=0;c<channels;c++){
  int value=(x*137+y*211+c*1103+((x/5+y/3)%2)*897)&maximum;
  if(x==0)value=0;if(x==w-1)value=maximum;input[(y*w+x)*channels+c]=value;
 }
 struct jpeg_compress_struct c;struct jpeg_error_mgr err;c.err=jpeg_std_error(&err);jpeg_create_compress(&c);
 c.image_width=w;c.image_height=h;c.input_components=channels;c.in_color_space=channels==1?JCS_GRAYSCALE:JCS_RGB;jpeg_set_defaults(&c);
 c.data_precision=precision;c.arith_code=TRUE;if(color==1)jpeg_set_colorspace(&c,JCS_RGB);jpeg_set_quality(&c,quality,FALSE);c.optimize_coding=FALSE;c.restart_interval=restart;
 if(color==2){c.comp_info[0].h_samp_factor=sampling==0?1:2;c.comp_info[0].v_samp_factor=sampling==2?2:1;}
 if(restart){c.comp_info[0].dc_tbl_no=15;c.comp_info[0].ac_tbl_no=14;for(int i=0;i<NUM_ARITH_TBLS;i++){c.arith_dc_L[i]=2;c.arith_dc_U[i]=5;c.arith_ac_K[i]=12;}}
 jpeg_scan_info scans[100];memset(scans,0,sizeof(scans));
 if(progressive==1)jpeg_simple_progression(&c);
 else if(progressive>=2){
  int n=0,depth=progressive>=3?3:0;
  for(int level=depth;level>=0;level--){
   int high=level==depth?0:level+1;
   for(int component=0;component<channels;component++){
    scans[n].comps_in_scan=1;scans[n].component_index[0]=component;scans[n].Ah=high;scans[n++].Al=level;
   }
   if(progressive==4)break;
   for(int component=0;component<channels;component++)for(int band=0;band<2;band++){
    scans[n].comps_in_scan=1;scans[n].component_index[0]=component;scans[n].Ss=band?10:1;scans[n].Se=band?63:9;scans[n].Ah=high;scans[n++].Al=level;
   }
  }
  c.scan_info=scans;c.num_scans=n;
 }
 else if(restart==2&&channels==3){for(int i=0;i<3;i++){scans[i].comps_in_scan=1;scans[i].component_index[0]=i;scans[i].Se=63;}c.scan_info=scans;c.num_scans=3;}
 unsigned char*encoded=NULL;unsigned long length=0;jpeg_mem_dest(&c,&encoded,&length);jpeg_start_compress(&c,TRUE);
 while(c.next_scanline<c.image_height){J12SAMPROW row=input+c.next_scanline*w*channels;if(precision==12)jpeg12_write_scanlines(&c,&row,1);else{JSAMPLE b[35*3];for(int j=0;j<w*channels;j++)b[j]=(JSAMPLE)row[j];JSAMPROW r=b;jpeg_write_scanlines(&c,&r,1);}}jpeg_finish_compress(&c);jpeg_destroy_compress(&c);
 FILE*f=fopen(argv[1],"wb");if(!f)return 3;fwrite(encoded,1,length,f);fclose(f);
 for(int fancy=0;fancy<=1;fancy++){
  struct jpeg_decompress_struct d;d.err=jpeg_std_error(&err);jpeg_create_decompress(&d);jpeg_mem_src(&d,encoded,length);jpeg_read_header(&d,TRUE);
  d.out_color_space=channels==1?JCS_GRAYSCALE:JCS_RGB;d.do_fancy_upsampling=fancy;d.dct_method=JDCT_ISLOW;d.do_block_smoothing=FALSE;jpeg_start_decompress(&d);
  char path[4096];snprintf(path,sizeof(path),"%s.%s.rgba",argv[1],fancy?"bilinear":"nearest");f=fopen(path,"wb");if(!f)return 3;
  J12SAMPLE buffer[35*3];
  while(d.output_scanline<d.output_height){J12SAMPROW row=buffer;if(precision==12)jpeg12_read_scanlines(&d,&row,1);else{JSAMPLE b[35*3];JSAMPROW r=b;jpeg_read_scanlines(&d,&r,1);for(int j=0;j<w*channels;j++)row[j]=b[j];}
   for(int x=0;x<w;x++){for(int ch=0;ch<3;ch++){int v=buffer[x*channels+(channels==1?0:ch)];fputc((v*255+maximum/2)/maximum,f);}fputc(255,f);}
  }
  fclose(f);jpeg_finish_decompress(&d);jpeg_destroy_decompress(&d);
 }
 free(encoded);return 0;
}
