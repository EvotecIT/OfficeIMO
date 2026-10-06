#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <tiffio.h>
#include <jpeglib.h>
/* Independent twelve/sixteen-bit SOF3 component producer and decoder. References are little-endian words. */
int main(int argc,char **argv) {
 if(argc!=9&&argc!=10)return 2;
 int precision=argc==10?atoi(argv[9]):16;if(precision!=12&&precision!=16)return 2;
 int maximum=(1<<precision)-1,midpoint=1<<(precision-1);
 int alphaKind=atoi(argv[8]);
 int photo=atoi(argv[2]),predictor=atoi(argv[3]),point=atoi(argv[4]),layout=atoi(argv[5]),separate=atoi(argv[6]),restart=atoi(argv[7]);
 int w=35,h=19,base=photo==5?4:(photo==2||photo==6)?3:1,n=photo==5?4:base+1;
 int tile=layout&1,big=layout&2,planar=layout&4?2:1,sw=tile?16:w,sh=16;
 TIFF*t=TIFFOpen(argv[1],big?"wb":"wl");if(!t)return 3;
 TIFFSetField(t,256,w);TIFFSetField(t,257,h);TIFFSetField(t,258,precision);TIFFSetField(t,277,n);
 TIFFSetField(t,262,photo==6?2:photo);TIFFSetField(t,284,planar);TIFFSetField(t,259,7);
 if(photo==6){TIFFSetField(t,530,1,1);float ref[6]={0,maximum,midpoint,maximum,midpoint,maximum};TIFFSetField(t,532,ref);}
 if(n>base){uint16_t extra=alphaKind;TIFFSetField(t,338,1,&extra);}
 if(tile){TIFFSetField(t,322,sw);TIFFSetField(t,323,sh);}else TIFFSetField(t,278,sh);
 uint16_t output[35*19*4]={0};
 for(int plane=0;plane<(planar==2?n:1);plane++)for(int y=0;y<h;y+=sh)for(int x=0;x<w;x+=sw){
  int rows=tile?sh:(h-y<sh?h-y:sh),channels=planar==2?1:n;
  uint16_t pixels[35*16*4],decoded[35*16*4];
  for(int yy=0;yy<rows;yy++)for(int xx=0;xx<sw;xx++)for(int c=0;c<channels;c++){
   int ch=planar==2?plane:c,gx=x+xx,gy=y+yy;
   if(gx>=w)gx=w-1;if(gy>=h)gy=h-1;
   const int alphas[]={0,1,2,3,4,8,16,128,256,1024,midpoint,maximum};
   int a=alphas[(gx+gy*3)%12],value=(gx*7919+gy*19531+ch*17659+(gx^gy)*2311)&maximum;
   if(n>base&&ch==base)value=a;
   else if(n>base&&alphaKind==1){
    if(photo==6&&ch>0)value=midpoint+(int)(((long long)value-midpoint)*a/maximum);
    else value=(int)((long long)value*a/maximum);
   }
   pixels[(yy*sw+xx)*channels+c]=(uint16_t)value;
  }
  struct jpeg_compress_struct c;struct jpeg_error_mgr err;c.err=jpeg_std_error(&err);jpeg_create_compress(&c);
  c.image_width=sw;c.image_height=rows;c.input_components=channels;c.in_color_space=JCS_UNKNOWN;jpeg_set_defaults(&c);c.data_precision=precision;
  jpeg_enable_lossless(&c,predictor,point);c.restart_in_rows=restart;
  jpeg_scan_info scans[4];memset(scans,0,sizeof(scans));
  if(separate&&channels>1){for(int i=0;i<channels;i++){scans[i].comps_in_scan=1;scans[i].component_index[0]=i;scans[i].Ss=predictor;scans[i].Al=point;}c.scan_info=scans;c.num_scans=channels;}
  unsigned char*encoded=NULL;unsigned long length=0;jpeg_mem_dest(&c,&encoded,&length);jpeg_start_compress(&c,TRUE);
  while(c.next_scanline<c.image_height){J16SAMPROW row=pixels+c.next_scanline*sw*channels;if(precision==16)jpeg16_write_scanlines(&c,&row,1);else jpeg12_write_scanlines(&c,(J12SAMPARRAY)&row,1);}
  jpeg_finish_compress(&c);jpeg_destroy_compress(&c);
  if((tile?TIFFWriteRawTile(t,TIFFComputeTile(t,x,y,0,plane),encoded,length):TIFFWriteRawStrip(t,TIFFComputeStrip(t,y,plane),encoded,length))<0)return 4;
  struct jpeg_decompress_struct d;d.err=jpeg_std_error(&err);jpeg_create_decompress(&d);jpeg_mem_src(&d,encoded,length);jpeg_read_header(&d,TRUE);
  d.jpeg_color_space=JCS_UNKNOWN;d.out_color_space=JCS_UNKNOWN;jpeg_start_decompress(&d);
  while(d.output_scanline<d.output_height){J16SAMPROW row=decoded+d.output_scanline*sw*channels;if(precision==16)jpeg16_read_scanlines(&d,&row,1);else jpeg12_read_scanlines(&d,(J12SAMPARRAY)&row,1);}
  jpeg_finish_decompress(&d);jpeg_destroy_decompress(&d);
  for(int i=0;i<sw*rows*channels;i++)if(decoded[i]!=(pixels[i]&~((1<<point)-1)))return 5;
  if(x==0&&y==0&&plane==0){char path[4096];snprintf(path,sizeof(path),"%s.jpg",argv[1]);FILE*f=fopen(path,"wb");if(!f)return 6;fwrite(encoded,1,length,f);fclose(f);
   snprintf(path,sizeof(path),"%s.jpg.raw",argv[1]);f=fopen(path,"wb");if(!f)return 6;for(int i=0;i<sw*rows*channels;i++){fputc(decoded[i]&255,f);fputc(decoded[i]>>8,f);}fclose(f);}
  for(int yy=0;yy<rows&&y+yy<h;yy++)for(int xx=0;xx<sw&&x+xx<w;xx++)for(int cc=0;cc<channels;cc++)output[((y+yy)*w+x+xx)*n+(planar==2?plane:cc)]=decoded[(yy*sw+xx)*channels+cc];
  free(encoded);
 }
 TIFFClose(t);char path[4096];snprintf(path,sizeof(path),"%s.raw",argv[1]);FILE*f=fopen(path,"wb");if(!f)return 6;for(int i=0;i<w*h*n;i++){fputc(output[i]&255,f);fputc(output[i]>>8,f);}fclose(f);return 0;
}
