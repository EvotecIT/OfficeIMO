#include <stdio.h>
#include <stdlib.h>
#include <jpeglib.h>
int main(int argc,char**argv){
 if(argc!=3)return 2;
 struct jpeg_compress_struct c;struct jpeg_error_mgr e;c.err=jpeg_std_error(&e);jpeg_create_compress(&c);
 FILE*f=fopen(argv[1],"wb");if(!f)return 3;jpeg_stdio_dest(&c,f);
 c.image_width=8;c.image_height=8;c.input_components=1;c.in_color_space=JCS_GRAYSCALE;jpeg_set_defaults(&c);
 for(int i=0;i<64;i++)c.quant_tbl_ptrs[0]->quantval[i]=65535;
 jvirt_barray_ptr a=(*c.mem->request_virt_barray)((j_common_ptr)&c,JPOOL_IMAGE,TRUE,1,1,1);
 jpeg_calc_jpeg_dimensions(&c);
 jpeg_write_coefficients(&c,&a);
 JBLOCKARRAY block=(*c.mem->access_virt_barray)((j_common_ptr)&c,a,0,1,TRUE);
 block[0][0][atoi(argv[2])]=1;
 jpeg_finish_compress(&c);jpeg_destroy_compress(&c);fclose(f);return 0;
}
