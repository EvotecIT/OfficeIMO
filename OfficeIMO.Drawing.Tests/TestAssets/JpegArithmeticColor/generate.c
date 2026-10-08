#include <stdio.h>
#include <stdlib.h>
#include <jpeglib.h>
int main(int argc,char**argv){
 if(argc!=8)return 1;int bits=atoi(argv[2]),ycck=atoi(argv[3]),progressive=atoi(argv[4]),sampling=atoi(argv[5]);
 int arithmetic=atoi(argv[6]),quality=atoi(argv[7]);
 int max=(1<<bits)-1,w=35,h=19;J12SAMPLE input[35*19*4];
 for(int y=0;y<h;y++)for(int x=0;x<w;x++)for(int ch=0;ch<4;ch++)input[(y*w+x)*4+ch]=(x*137+y*211+ch*1103+((x/5+y/3)%2)*897)&max;
 struct jpeg_compress_struct c;struct jpeg_error_mgr err;c.err=jpeg_std_error(&err);jpeg_create_compress(&c);
 c.image_width=w;c.image_height=h;c.input_components=4;c.in_color_space=JCS_CMYK;jpeg_set_defaults(&c);c.data_precision=bits;c.arith_code=arithmetic;jpeg_set_colorspace(&c,ycck?JCS_YCCK:JCS_CMYK);jpeg_set_quality(&c,quality,FALSE);c.restart_interval=3;
 if(ycck){c.comp_info[0].h_samp_factor=c.comp_info[3].h_samp_factor=sampling?2:1;c.comp_info[0].v_samp_factor=c.comp_info[3].v_samp_factor=sampling==2?2:1;}
 if(progressive)jpeg_simple_progression(&c);
 unsigned char*encoded=NULL;unsigned long length=0;jpeg_mem_dest(&c,&encoded,&length);jpeg_start_compress(&c,TRUE);
 while(c.next_scanline<c.image_height){J12SAMPROW row=input+c.next_scanline*w*4;if(bits==12)jpeg12_write_scanlines(&c,&row,1);else{JSAMPLE b[35*4];for(int i=0;i<w*4;i++)b[i]=row[i];JSAMPROW p=b;jpeg_write_scanlines(&c,&p,1);}}
 jpeg_finish_compress(&c);jpeg_destroy_compress(&c);FILE*f=fopen(argv[1],"wb");fwrite(encoded,1,length,f);fclose(f);
 for(int fancy=0;fancy<2;fancy++){
 struct jpeg_decompress_struct d;d.err=jpeg_std_error(&err);jpeg_create_decompress(&d);jpeg_mem_src(&d,encoded,length);jpeg_read_header(&d,TRUE);d.out_color_space=JCS_CMYK;d.do_fancy_upsampling=fancy;d.do_block_smoothing=FALSE;d.dct_method=JDCT_ISLOW;jpeg_start_decompress(&d);
 char name[4096];snprintf(name,sizeof(name),"%s.%s.cmyk16",argv[1],fancy?"fancy":"nearest");f=fopen(name,"wb");
 while(d.output_scanline<d.output_height){J12SAMPLE row[35*4];if(bits==12){J12SAMPROW p=row;jpeg12_read_scanlines(&d,&p,1);}else{JSAMPLE b[35*4];JSAMPROW p=b;jpeg_read_scanlines(&d,&p,1);for(int i=0;i<w*4;i++)row[i]=b[i];}for(int i=0;i<w*4;i++){fputc(row[i]&255,f);fputc(row[i]>>8,f);}}
 fclose(f);jpeg_finish_decompress(&d);jpeg_destroy_decompress(&d);}
 free(encoded);return 0;
}
