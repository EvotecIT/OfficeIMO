/* Opt-in actual unmodified native row upscaler; tile, odd chroma and all legal denominator controls. */
#include <stdio.h>
#include <stdint.h>
#include <string.h>
#include "av1/common/av1_common_int.h"
#include "av1/common/resize.h"
static uint8_t input[3*512],saved[3*512],output[3*512];
static void hex(const uint8_t *p,int count){for(int i=0;i<count;i++)printf("%02x",p[i]);}
int main(void) {
  const int widths[5]={17,33,65,129,257};
  for(int denom=9;denom<=16;denom++)for(int size=0;size<5;size++)for(int plane=0;plane<2;plane++)
  for(int pattern=0;pattern<2;pattern++)for(int tiled=0;tiled<2;tiled++) {
    AV1_COMMON cm;SequenceHeader seq;memset(&cm,0,sizeof(cm));memset(&seq,0,sizeof(seq));
    int upscaled=widths[size],width=(upscaled*8+denom/2)/denom,height=3;
    cm.width=width;cm.height=height;cm.superres_upscaled_width=upscaled;cm.superres_scale_denominator=denom;
    cm.mi_params.mi_cols=2*((width+7)/8);cm.mi_params.mi_rows=2*((height+7)/8);
    seq.subsampling_x=seq.subsampling_y=1;seq.bit_depth=AOM_BITS_8;seq.mib_size_log2=4;cm.seq_params=&seq;
    int blocks=(cm.mi_params.mi_cols+15)/16;
    cm.tiles.cols=tiled && blocks>1?2:1;cm.tiles.col_start_sb[0]=0;
    cm.tiles.col_start_sb[1]=cm.tiles.cols==2?blocks/2:blocks;cm.tiles.col_start_sb[2]=blocks;
    int rows=(height+(plane?1:0))>>(plane?1:0),outwidth=(upscaled+(plane?1:0))>>(plane?1:0);
    memset(input,0,sizeof(input));memset(output,0,sizeof(output));
    for(int y=0;y<rows;y++)for(int x=0;x<(cm.mi_params.mi_cols*4>>(plane?1:0));x++)
      input[y*512+16+x]=(uint8_t)(pattern?((x+y)%2?255:0):(x*17+y*29+x/7)%256);
    memcpy(saved,input,sizeof(input));
    av1_upscale_normative_rows(&cm,input+16,512,output,512,plane,rows);
    if(memcmp(saved,input,sizeof(input)))return 2;
    printf("{\"width\":%d,\"height\":%d,\"codedWidth\":%d,\"denominator\":%d,\"plane\":%d,\"tiles\":%d,\"inputStride\":%d,\"input\":\"",upscaled,height,width,denom,plane,cm.tiles.cols,496);
    for(int y=0;y<rows;y++)hex(input+y*512+16,512-16);
    printf("\",\"output\":\"");for(int y=0;y<rows;y++)hex(output+y*512,outwidth);printf("\"}\n");
  }
  return 0;
}
