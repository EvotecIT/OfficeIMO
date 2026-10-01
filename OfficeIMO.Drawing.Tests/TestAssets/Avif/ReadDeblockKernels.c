/* Opt-in actual native Main-8 loop-filter kernels and threshold initialization. */
#include <stdint.h>
#include <stdio.h>
#include <string.h>
#include "config/aom_dsp_rtcd.h"
#include "av1/common/av1_loopfilter.h"
#include "av1/common/av1_common_int.h"
typedef void (*Filter)(uint8_t *,int,const uint8_t *,const uint8_t *,const uint8_t *);
static int sample(int distance,int line,int pattern) {
  if(pattern>=9)return 255-sample(distance,line,pattern-9);
  int side=distance<0?-1:1,k=distance<0?-distance-1:distance;
  switch(pattern) {
    case 0:return 128;
    case 1:return 128+side*4;
    case 2:return 128+side*6+((k+line)%3==0?1:0);
    case 3:return 128+side*(4+k*4);
    case 4:return distance<0?0:255;
    case 5:return 128+side*(k<4?4:12);
    case 6:return 128+side*(k==1?10:4);
    case 7:return distance<0?20:95;
    default:return 128+side*8+((k*19+line*7)%17)-8;
  }
}
static void hex(const uint8_t *p) {for(int i=0;i<256;i++)printf("%02x",p[i]);}
int main(void) {
  const int levels[]={1,3,16,31,32,48,63},sharpness[]={0,3,7},sizes[]={4,8,16,8};
  const Filter vertical[]={aom_lpf_vertical_4_c,aom_lpf_vertical_8_c,aom_lpf_vertical_14_c,aom_lpf_vertical_6_c};
  const Filter horizontal[]={aom_lpf_horizontal_4_c,aom_lpf_horizontal_8_c,aom_lpf_horizontal_14_c,aom_lpf_horizontal_6_c};
  for(int sharp=0;sharp<3;sharp++) {
    AV1_COMMON cm;memset(&cm,0,sizeof(cm));cm.lf.sharpness_level=sharpness[sharp];av1_loop_filter_init(&cm);
    for(int l=0;l<7;l++) for(int kind=0;kind<4;kind++) for(int pass=0;pass<2;pass++) for(int pattern=0;pattern<18;pattern++) {
      uint8_t in[256],out[256];
      for(int y=0;y<16;y++)for(int x=0;x<16;x++)in[y*16+x]=(uint8_t)sample((pass?y:x)-8,pass?x:y,pattern);
      memcpy(out,in,sizeof(in));const loop_filter_thresh *t=&cm.lf_info.lfthr[levels[l]];
      (pass?horizontal:vertical)[kind](out+8*16+8,16,t->mblim,t->lim,t->hev_thr);
      printf("{\"size\":%d,\"chroma\":%s,\"pass\":%d,\"level\":%d,\"sharpness\":%d,\"pattern\":%d,\"limit\":%d,\"blimit\":%d,\"threshold\":%d,\"input\":\"",
          sizes[kind],kind==3?"true":"false",pass,levels[l],sharpness[sharp],pattern,t->lim[0],t->mblim[0],t->hev_thr[0]);
      hex(in);printf("\",\"output\":\"");hex(out);printf("\"}\n");
    }
  }
  return 0;
}
