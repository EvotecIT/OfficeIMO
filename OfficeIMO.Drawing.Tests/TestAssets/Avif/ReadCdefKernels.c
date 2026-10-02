/* Opt-in actual native Main-8/Main10 direction search and CDEF kernels, with unavailable edges. */
#include <stdint.h>
#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include "config/av1_rtcd.h"
#include "av1/common/cdef_block.h"
static void hex(const uint8_t *p) {for(int i=0;i<144;i++)printf("%02x",p[i]);}
static void hex16(const uint16_t *p) {for(int i=0;i<144;i++)printf("%02x%02x",p[i]&255,p[i]>>8);}
static int line(int x,int y,int d) {
  switch(d) {case 0:return y+x;case 1:return y+x/2;case 2:return y;case 3:return 3+y-x/2;
    case 4:return 7+y-x;case 5:return 3-y/2+x;case 6:return x;default:return y/2+x;}
}
int main(int argc,char **argv) {
  if(argc!=1 && argc!=2)return 2;
  const int depth=argc==2?atoi(argv[1]):8,shift=depth-8;
  if(depth!=8 && depth!=10)return 2;
  uint16_t in[CDEF_BSTRIDE*16],wide[144],wide_out[144];uint8_t raw[144],out[144];
  for(int pattern=0;pattern<17;pattern++)for(int phase=0;phase<(depth==10?4:1);phase++) {
    for(int y=0;y<12;y++)for(int x=0;x<12;x++) {
      int value=pattern==0?128:128+(line(x,y,(pattern-1)%8)*7%21)-10;
      if(pattern>8)value=255-value;raw[y*12+x]=(uint8_t)value;
      wide[y*12+x]=(uint16_t)((value<<shift)+(depth==10?((x+3*y+phase)&3):0));in[y*CDEF_BSTRIDE+x]=wide[y*12+x];
    }
    int32_t variance;int dir=cdef_find_dir_c(in,CDEF_BSTRIDE,&variance,shift);
    printf("{\"kind\":\"direction\",\"direction\":%d,\"variance\":%d,\"input\":\"",dir,variance);
    if(depth==8) {hex(raw);printf("\"}\n");} else {hex16(wide);printf("\",\"phase\":%d}\n",phase);}
  }
  const int primary[]={0,1,6,15},secondary[]={0,2,4},damping[]={2,3,6};
  const cdef_filter_block_func filters[]={cdef_filter_8_0_c,cdef_filter_8_1_c,cdef_filter_8_2_c,cdef_filter_8_3_c};
  const cdef_filter_block_func high_filters[]={cdef_filter_16_0_c,cdef_filter_16_1_c,cdef_filter_16_2_c,cdef_filter_16_3_c};
  for(int size=4;size<=8;size+=4)for(int dir=0;dir<8;dir++)for(int p=0;p<4;p++)for(int s=0;s<3;s++)
    for(int d=0;d<3;d++)for(int pattern=0;pattern<4;pattern++)for(int phase=0;phase<(depth==10?4:1);phase++) {
      int edge=pattern==3,x0=edge?0:2,y0=edge?0:2,width=edge?size+1:12,height=edge?size+3:12;
      for(int y=0;y<12;y++)for(int x=0;x<12;x++) {
        int value=pattern==0?128:128+(x*3-y*2)%17;if(pattern==2)value=255-value;
        if(edge)value=90+(x*7+y*13)%31;raw[y*12+x]=(uint8_t)value;
        wide[y*12+x]=(uint16_t)((value<<shift)+(depth==10?((x+3*y+phase)&3):0));
      }
      for(int i=0;i<CDEF_BSTRIDE*16;i++)in[i]=CDEF_VERY_LARGE;
      for(int y=0;y<height;y++)for(int x=0;x<width;x++)in[(y+2)*CDEF_BSTRIDE+x+2]=wide[y*12+x];
      memcpy(out,raw,sizeof(out));int index=(secondary[s]==0)|((primary[p]==0)<<1);
      if(depth==8) filters[index](out+y0*12+x0,12,in+(y0+2)*CDEF_BSTRIDE+x0+2,primary[p],secondary[s],dir,damping[d],damping[d],0,size,size);
      else {
        memcpy(wide_out,wide,sizeof(wide));
        high_filters[index](wide_out+y0*12+x0,12,in+(y0+2)*CDEF_BSTRIDE+x0+2,primary[p]<<shift,secondary[s]<<shift,dir,damping[d]+shift,damping[d]+shift,shift,size,size);
      }
      printf("{\"kind\":\"kernel\",\"size\":%d,\"direction\":%d,\"primary\":%d,\"secondary\":%d,\"damping\":%d,\"x\":%d,\"y\":%d,\"width\":%d,\"height\":%d,\"input\":\"",
        size,dir,primary[p],secondary[s],damping[d],x0,y0,width,height);
      if(depth==8) {hex(raw);printf("\",\"output\":\"");hex(out);printf("\"}\n");}
      else {hex16(wide);printf("\",\"output\":\"");hex16(wide_out);printf("\",\"phase\":%d}\n",phase);}
    }
  return 0;
}
