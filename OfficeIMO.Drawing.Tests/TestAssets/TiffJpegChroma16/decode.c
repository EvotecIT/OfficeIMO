/* Test-only libjpeg-turbo component decoder; output is little-endian words. */
#include <stdio.h>
#include <stdlib.h>
#include <jpeglib.h>
int main(int argc,char **argv) {
    if(argc!=4)return 2;
    FILE *f=fopen(argv[1],"rb");if(!f)return 3;
    struct jpeg_decompress_struct d;struct jpeg_error_mgr e;
    d.err=jpeg_std_error(&e);jpeg_create_decompress(&d);jpeg_stdio_src(&d,f);
    jpeg_read_header(&d,TRUE);d.jpeg_color_space=JCS_UNKNOWN;d.out_color_space=JCS_UNKNOWN;
    d.do_fancy_upsampling=atoi(argv[3]);jpeg_start_decompress(&d);
    FILE *out=fopen(argv[2],"wb");if(!out)return 4;
    size_t count=(size_t)d.output_width*d.output_components;
    void *buffer=malloc(count*sizeof(J16SAMPLE));if(!buffer)return 5;
    while(d.output_scanline<d.output_height) {
        if(d.data_precision<=8) {JSAMPROW row=buffer;if(jpeg_read_scanlines(&d,&row,1)!=1)return 6;}
        else if(d.data_precision<=12) {J12SAMPROW row=buffer;if(jpeg12_read_scanlines(&d,&row,1)!=1)return 6;}
        else {J16SAMPROW row=buffer;if(jpeg16_read_scanlines(&d,&row,1)!=1)return 6;}
        for(size_t i=0;i<count;i++) {
            unsigned value=d.data_precision<=8?((JSAMPLE*)buffer)[i]:d.data_precision<=12?((J12SAMPLE*)buffer)[i]:((J16SAMPLE*)buffer)[i];
            if(fputc(value&255,out)==EOF||fputc(value>>8,out)==EOF)return 7;
        }
    }
    jpeg_finish_decompress(&d);jpeg_destroy_decompress(&d);free(buffer);
    fclose(f);return fclose(out)==0?0:8;
}
