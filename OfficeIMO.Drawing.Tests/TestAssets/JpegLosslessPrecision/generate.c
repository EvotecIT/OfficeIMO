/* Independent libjpeg-turbo lossless producer and native-sample decoder.
   Compile against libjpeg-turbo 3.2.0; no OfficeIMO code is used. */
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <jpeglib.h>
int main(int argc,char **argv) {
 if(argc!=7&&argc!=8)return 2;
 int ycc=argc==8;
 int precision=atoi(argv[2]),channels=atoi(argv[3]),predictor=atoi(argv[4]),point=atoi(argv[5]),separate=atoi(argv[6]);
 const int w=17,h=11;int maximum=(1<<precision)-1,count=w*h*channels;
 unsigned short words[17*11*3],decoded[17*11*3];unsigned char bytes[17*11*3];
 for(int y=0;y<h;y++)for(int x=0;x<w;x++)for(int c=0;c<channels;c++){
  int at=(y*w+x)*channels+c,value=(x*7919+y*19531+c*17659+(x^y)*2311)&maximum;
  if(x==0)value=0;if(x==w-1)value=maximum;
  if(ycc&&c>0&&x%3==0)value=1<<(precision-1);
  words[at]=value;bytes[at]=value;
 }
 struct jpeg_compress_struct c;struct jpeg_error_mgr error;c.err=jpeg_std_error(&error);jpeg_create_compress(&c);
 c.image_width=w;c.image_height=h;c.input_components=channels;c.in_color_space=channels==1?JCS_GRAYSCALE:ycc?JCS_UNKNOWN:JCS_RGB;jpeg_set_defaults(&c);
 if(ycc)for(int i=0;i<3;i++)c.comp_info[i].component_id=i+1;
 c.data_precision=precision;jpeg_enable_lossless(&c,predictor,point);c.restart_in_rows=2;
 for(int i=0;i<channels;i++){c.comp_info[i].h_samp_factor=1;c.comp_info[i].v_samp_factor=1;c.comp_info[i].quant_tbl_no=0;c.comp_info[i].dc_tbl_no=0;c.comp_info[i].ac_tbl_no=0;}
 jpeg_scan_info scans[3];memset(scans,0,sizeof(scans));
 if(separate&&channels==3){for(int i=0;i<channels;i++){scans[i].comps_in_scan=1;scans[i].component_index[0]=i;scans[i].Ss=predictor;scans[i].Al=point;}c.scan_info=scans;c.num_scans=channels;}
 unsigned char*encoded=NULL;unsigned long length=0;jpeg_mem_dest(&c,&encoded,&length);jpeg_start_compress(&c,TRUE);
 while(c.next_scanline<c.image_height){int at=c.next_scanline*w*channels;
  if(precision<=8){JSAMPROW row=bytes+at;jpeg_write_scanlines(&c,&row,1);}
  else if(precision<=12){J12SAMPROW row=(J12SAMPROW)(words+at);jpeg12_write_scanlines(&c,&row,1);}
  else{J16SAMPROW row=words+at;jpeg16_write_scanlines(&c,&row,1);}
 }
 jpeg_finish_compress(&c);jpeg_destroy_compress(&c);
 struct jpeg_decompress_struct d;d.err=jpeg_std_error(&error);jpeg_create_decompress(&d);jpeg_mem_src(&d,encoded,length);jpeg_read_header(&d,TRUE);if(ycc){d.jpeg_color_space=JCS_UNKNOWN;d.out_color_space=JCS_UNKNOWN;}jpeg_start_decompress(&d);
 while(d.output_scanline<d.output_height){int at=d.output_scanline*w*channels;
  if(precision<=8){JSAMPROW row=bytes+at;jpeg_read_scanlines(&d,&row,1);for(int i=0;i<w*channels;i++)decoded[at+i]=bytes[at+i];}
  else if(precision<=12){J12SAMPROW row=(J12SAMPROW)(decoded+at);jpeg12_read_scanlines(&d,&row,1);}
  else{J16SAMPROW row=decoded+at;jpeg16_read_scanlines(&d,&row,1);}
 }
 jpeg_finish_decompress(&d);jpeg_destroy_decompress(&d);
 for(int i=0;i<count;i++)if(decoded[i]!=(words[i]&~((1<<point)-1)))return 3;
 FILE*f=fopen(argv[1],"wb");if(!f)return 4;fwrite(encoded,1,length,f);fclose(f);free(encoded);
 char path[4096];snprintf(path,sizeof(path),"%s.raw",argv[1]);f=fopen(path,"wb");if(!f)return 4;
 for(int i=0;i<count;i++){fputc(decoded[i]&255,f);fputc(decoded[i]>>8,f);}fclose(f);return 0;
}
