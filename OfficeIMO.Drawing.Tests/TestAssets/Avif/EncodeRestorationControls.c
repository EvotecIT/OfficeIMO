/* Opt-in native producer: odd-sized still frames with texture/noise and native-selected restoration. */
#include <stdio.h>
#include <stdlib.h>
#include "aom/aom_encoder.h"
#include "aom/aomcx.h"
int main(int argc,char **argv) {
  if(argc!=3)return 2;const int mode=atoi(argv[2]),w=mode>=3?513:193,h=mode==5?513:mode>=3?257:137,mono=mode==2;
  if(mode<0 || mode>5)return 2;
  aom_codec_enc_cfg_t cfg;
  if(aom_codec_enc_config_default(aom_codec_av1_cx(),&cfg,AOM_USAGE_GOOD_QUALITY))return 3;
  cfg.g_w=w;cfg.g_h=h;cfg.g_threads=1;cfg.g_timebase.num=1;cfg.g_timebase.den=1;
  cfg.g_limit=1;cfg.g_lag_in_frames=0;cfg.monochrome=mono;cfg.rc_end_usage=AOM_Q;
  cfg.rc_min_quantizer=cfg.rc_max_quantizer=36;
  aom_codec_ctx_t enc;if(aom_codec_enc_init(&enc,aom_codec_av1_cx(),&cfg,0))return 4;
  if(aom_codec_control(&enc,AOME_SET_CPUUSED,2) || aom_codec_control(&enc,AV1E_SET_ENABLE_INTRABC,0) ||
     aom_codec_control(&enc,AV1E_SET_ENABLE_PALETTE,0) || aom_codec_control(&enc,AV1E_SET_ENABLE_RESTORATION,1) ||
     aom_codec_control(&enc,AV1E_SET_SUPERBLOCK_SIZE,AOM_SUPERBLOCK_SIZE_64X64) ||
     aom_codec_control(&enc,AV1E_SET_TILE_COLUMNS,1) || aom_codec_control(&enc,AV1E_SET_COLOR_RANGE,mono?1:0))return 5;
  aom_image_t *image=aom_img_alloc(NULL,AOM_IMG_FMT_I420,w,h,1);if(!image)return 6;
  unsigned int random=7;
  for(int p=0;p<3;p++) {
    int width=(w+(p?1:0))>>(p?1:0),height=(h+(p?1:0))>>(p?1:0);
    for(int y=0;y<height;y++)for(int x=0;x<width;x++) {
      random=random*1664525u+1013904223u;
      int base=mode==4?64+(x/16%2)*96:mode>=3?96+(x*17+y*23+p*11)%31:
          mode==1?48+((x*3+y*2+p*17)%127):64+((x/12+y/9+p)%3)*45;
      image->planes[p][y*image->stride[p]+x]=(unsigned char)(base+(int)(random>>(mode>=3?26:28)));
    }
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
