/* Opt-in native AOM restoration writer/decoder matrix. Uses native geometry, CDFs and filter readers. */
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "aom_dsp/bitwriter.h"
#include "aom_dsp/binary_codes_writer.h"
#include "aom_dsp/bitreader.h"
#include "av1/common/av1_common_int.h"
#include "av1/common/entropy.h"
#include "av1/common/restoration.h"
extern void office_trace_restoration_probe(const AV1_COMMON *,MACROBLOCKD *,aom_reader *,int,int);

static int frame_type(int value) { return value==1?RESTORE_SWITCHABLE:value==2?RESTORE_WIENER:value==3?RESTORE_SGRPROJ:RESTORE_NONE; }
static void write_unit(aom_writer *w,FRAME_CONTEXT *fc,int mode,int p,int seed,
                         int taps[3][2][3],int xqd[3][2],RestorationUnitInfo *expected) {
  int type=mode==RESTORE_SWITCHABLE?seed%3:(seed%7==0?RESTORE_NONE:mode);
  expected->restoration_type=(RestorationType)type;
  if(mode==RESTORE_SWITCHABLE) aom_write_symbol(w,type,fc->switchable_restore_cdf,3);
  else aom_write_symbol(w,type!=RESTORE_NONE,mode==RESTORE_WIENER?fc->wiener_restore_cdf:fc->sgrproj_restore_cdf,2);
  if(type==RESTORE_WIENER) {
    for(int pass=0;pass<2;pass++) for(int j=p?1:0;j<3;j++) {
      int low=j==0?-5:j==1?-23:-17,count=j==0?16:j==1?32:64;
      int value=low+((seed*13+pass*7+j*3)%count);
      aom_write_primitive_refsubexpfin(w,count,j+1,taps[p][pass][j]-low,value-low);taps[p][pass][j]=value;
      if(pass==0) expected->wiener_info.vfilter[j]=value;else expected->wiener_info.hfilter[j]=value;
    }
  } else if(type==RESTORE_SGRPROJ) {
    int set=seed%16;expected->sgrproj_info.ep=set;aom_write_literal(w,set,4);
    for(int j=0;j<2;j++) {
      int low=j==0?-96:-32,value=low+(seed*17+j*11)%128;
      if((j==0 && set>=10 && set<=13) || (j==1 && set>=14)) {
        value=j==0?0:clamp(128-xqd[p][0],-32,95);
      } else aom_write_primitive_refsubexpfin(w,128,4,xqd[p][j]-low,value-low);
      expected->sgrproj_info.xqd[j]=value;xqd[p][j]=value;
    }
  }
}
static void check(RestorationUnitInfo *a,RestorationUnitInfo *b) {
  if(a->restoration_type!=b->restoration_type) abort();
  if(a->restoration_type==RESTORE_WIENER) for(int j=0;j<3;j++) {
    if(a->wiener_info.vfilter[j]!=b->wiener_info.vfilter[j] || a->wiener_info.hfilter[j]!=b->wiener_info.hfilter[j]) abort();
  }
  if(a->restoration_type==RESTORE_SGRPROJ && (a->sgrproj_info.ep!=b->sgrproj_info.ep || a->sgrproj_info.xqd[0]!=b->sgrproj_info.xqd[0] || a->sgrproj_info.xqd[1]!=b->sgrproj_info.xqd[1])) abort();
}
int main(void) {
  for(int scenario=0;scenario<192;scenario++) {
    int large=scenario%2,mono=scenario/2%2,updates=scenario/4%2,denom=8+scenario/8%9;
    int upwidth=511+scenario%3,height=265+scenario%5,width=(upwidth*8+denom/2)/denom;
    int units=large?32:16,mirows=2*((height+7)/8),micols=2*((width+7)/8);
    int rowStart=scenario/9%2?units:0,colStart=scenario/11%2?units:0;
    SequenceHeader seq={0};seq.sb_size=large?BLOCK_128X128:BLOCK_64X64;seq.subsampling_x=seq.subsampling_y=1;
    AV1_COMMON cm={0};cm.seq_params=&seq;cm.width=width;cm.height=height;cm.superres_upscaled_width=upwidth;cm.superres_upscaled_height=height;cm.superres_scale_denominator=denom;
    int size=large?(scenario/16%2?256:128):(64<<(scenario/16%3));
    int sizes[3]={size,size>>(scenario/7%2),size>>(scenario/7%2)},modes[3];
    RestorationUnitInfo expected[3][256]={0},decoded[3][256]={0};
    for(int p=0;p<3;p++) {
      modes[p]=p>=1 && mono?0:1+((scenario/3+p)%3);
      if(scenario%13==0 && p==1) modes[p]=0;
      RestorationInfo *info=&cm.rst_info[p];info->frame_restoration_type=frame_type(modes[p]);info->restoration_unit_size=sizes[p];
      info->horz_units=av1_lr_count_units(sizes[p],(upwidth+(p?1:0))>>(p?1:0));
      info->vert_units=av1_lr_count_units(sizes[p],(height+(p?1:0))>>(p?1:0));info->unit_info=decoded[p];
      if(info->horz_units*info->vert_units>256) abort();
    }
    FRAME_CONTEXT fc={0};av1_init_mode_probs(&fc);
    int taps[3][2][3],xqd[3][2];
    for(int p=0;p<3;p++) for(int pass=0;pass<2;pass++) { taps[p][pass][0]=3;taps[p][pass][1]=-7;taps[p][pass][2]=15;xqd[p][pass]=pass?31:-32; }
    uint8_t buffer[1048576];aom_writer writer;writer.allow_update_cdf=updates;aom_start_encode(&writer,buffer);
    int ordinal=0;
    for(int r=rowStart;r<mirows;r+=units) for(int c=colStart;c<micols;c+=units) for(int p=0;p<(mono?1:3);p++) if(modes[p]) {
      int c0,c1,r0,r1;
      if(av1_loop_restoration_corners_in_sb(&cm,p,r,c,seq.sb_size,&c0,&c1,&r0,&r1))
        for(int y=r0;y<r1;y++) for(int x=c0;x<c1;x++) write_unit(&writer,&fc,frame_type(modes[p]),p,scenario*19+ordinal++,taps,xqd,&expected[p][y*cm.rst_info[p].horz_units+x]);
    }
    aom_stop_encode(&writer);
    printf("{\"scenario\":%d,\"large\":%d,\"mono\":%d,\"updates\":%d,\"width\":%d,\"upwidth\":%d,\"height\":%d,\"denom\":%d,\"extent\":[%d,%d,%d,%d],\"types\":[%d,%d,%d],\"sizes\":[%d,%d,%d],\"hex\":\"",scenario,large,mono,updates,width,upwidth,height,denom,rowStart,mirows,colStart,micols,modes[0],modes[1],modes[2],sizes[0],sizes[1],sizes[2]);
    for(unsigned i=0;i<writer.pos;i++) printf("%02x",buffer[i]);printf("\"}\n");
    fprintf(stderr,"{\"event\":\"case\",\"scenario\":%d}\n",scenario);
    av1_init_mode_probs(&fc);MACROBLOCKD xd={0};xd.tile_ctx=&fc;av1_reset_loop_restoration(&xd,mono?1:3);
    aom_reader reader;if(aom_reader_init(&reader,buffer,writer.pos)) abort();reader.allow_update_cdf=updates;
    for(int r=rowStart;r<mirows;r+=units) for(int c=colStart;c<micols;c+=units) for(int p=0;p<(mono?1:3);p++) if(modes[p]) {
      int c0,c1,r0,r1;
      if(av1_loop_restoration_corners_in_sb(&cm,p,r,c,seq.sb_size,&c0,&c1,&r0,&r1)) for(int y=r0;y<r1;y++) for(int x=c0;x<c1;x++) {
        int index=y*cm.rst_info[p].horz_units+x;office_trace_restoration_probe(&cm,&xd,&reader,p,index);check(&expected[p][index],&decoded[p][index]);
      }
    }
    if(aom_reader_has_overflowed(&reader)) abort();
  }
  return 0;
}
