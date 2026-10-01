/* Opt-in observation of AOM quantization facts and actual integer inverse transforms. */
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "av1/common/av1_common_int.h"
#include "av1/common/quant_common.h"
extern void office_av1_inverse(const int32_t *,int32_t *,uint16_t *,int,int,int);
extern int office_av1_dequant(int,int,int,int,int,int);
static int read_int(FILE *f) {
  uint8_t b[4];if(fread(b,1,4,f)!=4) abort();
  uint32_t v=(uint32_t)b[0]|((uint32_t)b[1]<<8)|((uint32_t)b[2]<<16)|((uint32_t)b[3]<<24);
  return (int)((int64_t)v-((v&0x80000000)?INT64_C(4294967296):0));
}
static int clip(int v) {return v<0?0:v>255?255:v;}
static void numbers(const int32_t *v,int n) {putchar('[');for(int i=0;i<n;i++) printf("%s%d",i?",":"",v[i]);putchar(']');}
int main(int argc,char **argv) {
  if(argc!=3) return 2;
  FILE *input=fopen(argv[1],"rb"),*matrix=fopen(argv[2],"wb");if(!input || !matrix) return 3;
  CommonQuantParams qp;memset(&qp,0,sizeof(qp));av1_qm_init(&qp,3);
  const int offsets[]={0,16,80,336,336,1360,1392,1424,1552,1680,2192,336,336,2704,2768,2832,3088,1680,2192};
  for(int level=0;level<15;level++) for(int plane=0;plane<2;plane++) {
    uint8_t facts[3344]={0},seen[3344]={0};
    for(int size=0;size<19;size++) {
      int w=tx_size_wide[size]<32?tx_size_wide[size]:32,h=tx_size_high[size]<32?tx_size_high[size]:32;
      const qm_val_t *weights=qp.giqmatrix[level][plane][size];
      for(int r=0;r<h;r++) for(int c=0;c<w;c++) {
        int index=offsets[size]+r*w+c;uint8_t value=weights[c*h+r];
        if(seen[index] && facts[index]!=value) abort();facts[index]=value;seen[index]=1;
      }
    }
    for(int i=0;i<3344;i++) if(!seen[i]) abort();
    if(fwrite(facts,1,sizeof(facts),matrix)!=sizeof(facts)) abort();
  }
  fclose(matrix);
  int32_t lookup[256];printf("{\"event\":\"facts\",\"dc\":");
  for(int q=0;q<256;q++) lookup[q]=av1_dc_quant_QTX(q,0,AOM_BITS_8);numbers(lookup,256);
  printf(",\"ac\":");for(int q=0;q<256;q++) lookup[q]=av1_ac_quant_QTX(q,0,AOM_BITS_8);numbers(lookup,256);puts("}");
  int count=read_int(input);
  for(int scenario=0;scenario<count;scenario++) {
    int p[23];for(int i=0;i<23;i++) p[i]=read_int(input);
    int size=p[0],type=p[1],plane=p[2],w=tx_size_wide[size],h=tx_size_high[size],tw=w<32?w:32,th=h<32?h:32;
    if(p[22]!=tw*th) abort();
    int q=clip((p[5]?p[4]:p[3])+(p[7] && p[8]?p[9]:0));
    const int dcs[]={p[14],p[15],p[17]},acs[]={0,p[16],p[18]};
    int dc=av1_dc_quant_QTX(q,dcs[plane],AOM_BITS_8),ac=av1_ac_quant_QTX(q,acs[plane],AOM_BITS_8);
    int level=p[10] && type<9 && !p[20]?p[11+plane]:15;
    const qm_val_t *weights=level<15?qp.giqmatrix[level][plane][size]:NULL;
    int32_t dequant[1024]={0},residual[4096];uint16_t pixels[4096];
    for(int i=0;i<w*h;i++) {residual[i]=1234567;pixels[i]=128;}
    for(int r=0;r<th;r++) for(int c=0;c<tw;c++) {
      int pos=c*th+r,value=read_int(input);
      dequant[pos]=office_av1_dequant(value,dc,ac,pos,weights?weights[pos]:0,size);
    }
    office_av1_inverse(dequant,residual,pixels,size,type,p[20]);
    printf("{\"event\":\"case\",\"scenario\":%d,\"residual\":",scenario);numbers(residual,w*h);printf(",\"pixels\":[");
    for(int i=0;i<w*h;i++) {if(residual[i]==1234567) abort();printf("%s%u",i?",":"",pixels[i]);}
    puts("]}");
  }
  if(fgetc(input)!=EOF) abort();fclose(input);return 0;
}
