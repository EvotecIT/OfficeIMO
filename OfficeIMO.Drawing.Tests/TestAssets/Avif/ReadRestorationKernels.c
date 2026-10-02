/* Opt-in unmodified native Wiener/self-guided kernels, including signed extremes and full scratch footprints. */
#include <stdio.h>
#include <stdint.h>
#include <string.h>
#include <stdlib.h>
#include "aom_ports/mem.h"
#include "config/av1_rtcd.h"
#include "av1/common/restoration.h"
#include "av1/common/convolve.h"
static uint8_t input[76*76],output[76*76];
static uint16_t wide[76*76],wide_output[76*76];
static int32_t scratch[SGRPROJ_TMPBUF_SIZE];
static void hex(const uint8_t *p) {for(int i=0;i<76*76;i++)printf("%02x",p[i]);}
static void hex16(const uint16_t *p){for(int i=0;i<76*76;i++)printf("%02x%02x",p[i]&255,p[i]>>8);}
int main(int argc,char **argv) {
  if(argc!=1 && argc!=2)return 2;const int depth=argc==2?atoi(argv[1]):8;if(depth!=8 && depth!=10)return 2;
  for(int type=2;type<=3;type++)for(int set=0;set<(type==2?4:16);set++)
  for(int variant=0;variant<2;variant++)for(int pattern=0;pattern<2;pattern++)for(int extent=0;extent<2;extent++)for(int phase=0;phase<(depth==10?4:1);phase++) {
    int w=extent?64:17,h=extent?64:9,x0=variant?31:-96,x1=variant?95:-32;
    int taps[2][3]={{3,-7,15},{3,-7,15}};
    if(type==2) {
      if(set==1){taps[0][0]=-5;taps[0][1]=-23;taps[0][2]=-17;taps[1][0]=10;taps[1][1]=8;taps[1][2]=46;}
      if(set==2){taps[0][0]=10;taps[0][1]=8;taps[0][2]=46;taps[1][0]=-5;taps[1][1]=-23;taps[1][2]=-17;}
      if(set==3){memset(taps,0,sizeof(taps));}
      if(variant) {int temp[3];memcpy(temp,taps[0],sizeof(temp));memcpy(taps[0],taps[1],sizeof(temp));memcpy(taps[1],temp,sizeof(temp));}
    } else {if(set>=10 && set<=13)x0=0;if(set>=14){x1=128-x0;if(x1>95)x1=95;if(x1<-32)x1=-32;}}
    for(int y=0;y<76;y++)for(int x=0;x<76;x++)input[y*76+x]=(uint8_t)(pattern?((x+y)%2?255:0):(x*7+y*11+(x/13)*9)%256);
    memcpy(output,input,sizeof(input));
    for(int i=0;i<76*76;i++)wide[i]=(uint16_t)((input[i]<<2)+((i+phase)&3));
    memcpy(wide_output,wide,sizeof(wide));
    if(type==2) {
      _Alignas(256) int16_t v[8]={0},hfilter[8]={0};
      for(int t=0;t<3;t++){v[t]=v[6-t]=(int16_t)taps[0][t];hfilter[t]=hfilter[6-t]=(int16_t)taps[1][t];v[3]-=2*v[t];hfilter[3]-=2*hfilter[t];}
      WienerConvolveParams params=get_conv_params_wiener(depth);
      if(depth==8)av1_wiener_convolve_add_src_c(input+6*76+6,76,output+6*76+6,76,hfilter,16,v,16,w,h,&params);
      else av1_highbd_wiener_convolve_add_src_c(CONVERT_TO_BYTEPTR(wide+6*76+6),76,CONVERT_TO_BYTEPTR(wide_output+6*76+6),76,hfilter,16,v,16,w,h,&params,depth);
    } else {
      int xqd[2]={x0,x1};
      if(depth==8) {if(av1_apply_selfguided_restoration_c(input+6*76+6,w,h,76,set,xqd,output+6*76+6,76,scratch,8,0))return 2;}
      else if(av1_apply_selfguided_restoration_c(CONVERT_TO_BYTEPTR(wide+6*76+6),w,h,76,set,xqd,CONVERT_TO_BYTEPTR(wide_output+6*76+6),76,scratch,depth,1))return 2;
    }
    printf("{\"type\":%d,\"set\":%d,\"x0\":%d,\"x1\":%d,\"width\":%d,\"height\":%d,\"taps\":[[%d,%d,%d],[%d,%d,%d]],\"input\":\"",type,set,x0,x1,w,h,taps[0][0],taps[0][1],taps[0][2],taps[1][0],taps[1][1],taps[1][2]);
    if(depth==8)hex(input);else hex16(wide);printf("\",\"output\":\"");if(depth==8)hex(output);else hex16(wide_output);
    if(depth==8)printf("\"}\n");else printf("\",\"phase\":%d}\n",phase);
  }
  return 0;
}
