/* Opt-in native producer: lossy odd-sized color and monochrome still frames. */
#include <stdio.h>
#include <stdlib.h>
#include "aom/aom_encoder.h"
#include "aom/aomcx.h"
int main(int argc,char **argv) {
  if(argc!=3)return 2;const int w=97,h=65,mono=atoi(argv[2]);
  if(mono!=0 && mono!=1)return 2;
  aom_codec_enc_cfg_t cfg;
  if(aom_codec_enc_config_default(aom_codec_av1_cx(),&cfg,AOM_USAGE_GOOD_QUALITY))return 3;
  cfg.g_w=w;cfg.g_h=h;cfg.g_threads=1;cfg.g_timebase.num=1;cfg.g_timebase.den=1;
  cfg.g_limit=1;cfg.g_lag_in_frames=0;cfg.monochrome=mono;
  cfg.rc_end_usage=AOM_Q;cfg.rc_min_quantizer=cfg.rc_max_quantizer=36;
  aom_codec_ctx_t enc;if(aom_codec_enc_init(&enc,aom_codec_av1_cx(),&cfg,0))return 4;
  if(aom_codec_control(&enc,AOME_SET_CPUUSED,2) || aom_codec_control(&enc,AV1E_SET_ENABLE_INTRABC,0) ||
     aom_codec_control(&enc,AV1E_SET_ENABLE_PALETTE,0) || aom_codec_control(&enc,AV1E_SET_ENABLE_RESTORATION,0) ||
     aom_codec_control(&enc,AV1E_SET_COLOR_RANGE,mono?1:0))return 5;
  aom_image_t *image=aom_img_alloc(NULL,AOM_IMG_FMT_I420,w,h,1);if(!image)return 6;
  for(int p=0;p<3;p++) {
    int width=(w+(p?1:0))>>(p?1:0),height=(h+(p?1:0))>>(p?1:0);
    for(int y=0;y<height;y++)for(int x=0;x<width;x++)
      image->planes[p][y*image->stride[p]+x]=(unsigned char)(48+((x*3+y*2+p*17)%127)+((x/13+y/11)%2)*31);
  }
  FILE *out=fopen(argv[1],"wb");if(!out)return 7;
  for(int pass=0;pass<2;pass++) {
    if(aom_codec_encode(&enc,pass?NULL:image,pass,1,pass?0:AOM_EFLAG_FORCE_KF))return 8;
    aom_codec_iter_t iter=NULL;const aom_codec_cx_pkt_t *pkt;
    while((pkt=aom_codec_get_cx_data(&enc,&iter))!=NULL)
      if(pkt->kind==AOM_CODEC_CX_FRAME_PKT && fwrite(pkt->data.frame.buf,1,pkt->data.frame.sz,out)!=pkt->data.frame.sz)return 9;
  }
  if(fclose(out))return 10;aom_img_free(image);aom_codec_destroy(&enc);return 0;
}
