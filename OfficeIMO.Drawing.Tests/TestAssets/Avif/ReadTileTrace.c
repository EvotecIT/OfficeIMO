/* Opt-in AOM reference decoder driver. No native code is linked to OfficeIMO runtime projects. */
#include <stdio.h>
#include <stdlib.h>
#include <stdint.h>
#include "aom/aom_decoder.h"
#include "aom/aomdx.h"

int main(int argc,char **argv) {
  if(argc!=3) return 2;
  FILE *input=fopen(argv[1],"rb"); if(!input) return 3;
  if(fseek(input,0,SEEK_END)) return 3;
  long size=ftell(input); if(size<=0 || size>128*1024*1024) return 3;
  rewind(input); uint8_t *bytes=malloc((size_t)size);if(!bytes) return 4;
  if(fread(bytes,1,(size_t)size,input)!=(size_t)size) return 3;
  fclose(input);
  aom_codec_ctx_t decoder;
  aom_codec_dec_cfg_t cfg={1,0,0,1};
  if(aom_codec_dec_init(&decoder,aom_codec_av1_dx(),&cfg,0)) return 5;
  if(aom_codec_decode(&decoder,bytes,(size_t)size,NULL)) {
    fprintf(stderr,"Native AV1 error: %s (%s)\n",aom_codec_error(&decoder),aom_codec_error_detail(&decoder));return 6;
  }
  aom_codec_iter_t iter=NULL;
  aom_image_t *image=aom_codec_get_frame(&decoder,&iter);
  if(!image || (image->bit_depth!=8 && image->bit_depth!=10) ||
     ((image->fmt & AOM_IMG_FMT_HIGHBITDEPTH)!=0)!=(image->bit_depth==10)) return 7;
  FILE *output=fopen(argv[2],"wb");if(!output) return 8;
  int planes=image->monochrome?1:3;
  printf("{\"width\":%u,\"height\":%u,\"planes\":%d,\"depth\":%u,\"planeBytes\":[",image->d_w,image->d_h,planes,image->bit_depth);
  for(int p=0;p<planes;p++) {
    int w=(image->d_w+(p?1:0))>>(p?1:0),h=(image->d_h+(p?1:0))>>(p?1:0);
    printf("%s%d",p?",":"",w*h*(image->bit_depth==10?2:1));
    for(int y=0;y<h;y++) {
      if(image->bit_depth==8) {
        if(fwrite(image->planes[p]+y*image->stride[p],1,(size_t)w,output)!=(size_t)w) return 8;
      } else {
        const uint16_t *row=(const uint16_t *)(image->planes[p]+y*image->stride[p]);
        for(int x=0;x<w;x++) {uint8_t le[2]={(uint8_t)row[x],(uint8_t)(row[x]>>8)};
          if(fwrite(le,1,2,output)!=2) return 8;}
      }
    }
  }
  printf("]}\n"); if(fclose(output)) return 8;
  if(aom_codec_get_frame(&decoder,&iter)!=NULL) return 9;
  aom_codec_destroy(&decoder);free(bytes);return 0;
}
