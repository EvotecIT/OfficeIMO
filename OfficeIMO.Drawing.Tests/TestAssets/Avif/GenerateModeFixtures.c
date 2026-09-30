/* Original AV1 mode-syntax harness. Share the pinned entropy/prelude setup; normal builds never run this. */
#define main prelude_component_main
#include "GeneratePreludeFixtures.c"
#undef main
#include "mode-defaults.h"

typedef struct {
    codec entropy;
    aom_cdf_prob y[5][5][14], uv[2][13][15], angles[8][8], signs[9], alphas[6][17], copy[3];
    unsigned char modes[32][32];
    int contexts[25], cfl_index;
} mode_codec;
static void mode_init(mode_codec *m,int decode,int updates) {
    memset(m,0,sizeof(*m)); init(&m->entropy,decode,updates);
    memcpy(m->y,default_kf_y_mode_cdf,sizeof(m->y)); memcpy(m->uv,default_uv_mode_cdf,sizeof(m->uv));
    memcpy(m->angles,default_angle_delta_cdf,sizeof(m->angles)); memcpy(m->signs,default_cfl_sign_cdf,sizeof(m->signs));
    memcpy(m->alphas,default_cfl_alpha_cdf,sizeof(m->alphas)); memcpy(m->copy,default_intrabc_cdf,sizeof(m->copy));
}
static int mode_symbol(mode_codec *m,aom_cdf_prob *cdf,int n,int expected) {
    int actual=expected;
    codec *c=&m->entropy;
    if(c->decode) { actual=od_ec_decode_cdf_q15(&c->dec,cdf,n); if(expected>=0) assert(actual==expected); }
    else { assert(expected>=0); od_ec_encode_cdf_q15(&c->enc,expected,cdf,n); }
    if(c->updates) update_cdf(cdf,actual,n);
    return actual;
}
static int context_for(int mode) { const int map[13]={0,1,2,3,4,4,4,4,3,0,1,2,0};assert(mode>=0&&mode<13);return map[mode]; }
static int y_at(int index) {
    unsigned h=(unsigned)index+1; h^=h>>16; h*=0x7feb352dU; h^=h>>15; h*=0x846ca68bU; h^=h>>16;
    return (int)(h%13);
}
static void modes(mode_codec *m,int r,int c,int w,int h,int mono,int lossless,int allow_copy,int index,int *result) {
    int known=index>=0, chroma=!mono && !(h==4&&!(r&1)) && !(w==4&&!(c&1));
    int copy=allow_copy?mode_symbol(m,m->copy,2,known?index%9==0:-1):0;
    int y=0,uv=0,ay=0,av=0,au=0,bv=0,allowed=0;
    if(!copy) {
        int above=r?context_for(m->modes[r-1][c]):0, left=c?context_for(m->modes[r][c-1]):0;
        m->contexts[above*5+left]++;
        y=mode_symbol(m,m->y[above][left],13,known?y_at(index):-1);
        if(w*h>=64 && y>=1 && y<=8) ay=mode_symbol(m,m->angles[y-1],7,known?(index*3+y)%7:-1)-3;
        if(chroma) {
            allowed=lossless?w<=8&&h<=8:w<=32&&h<=32;
            int expected=known?(allowed&&index%3==0?13:(index*11+3)%13):-1;
            uv=mode_symbol(m,m->uv[allowed][y],allowed?14:13,expected);
            if(uv==13) {
                int signs=mode_symbol(m,m->signs,8,known?m->cfl_index%8:-1); m->cfl_index++;
                int su=(signs+1)/3,sv=(signs+1)%3;
                if(su) {au=1+mode_symbol(m,m->alphas[(su-1)*3+sv],16,known?(index*5+signs)%16:-1);if(su==1)au=-au;}
                if(sv) {bv=1+mode_symbol(m,m->alphas[(sv-1)*3+su],16,known?(index*7+signs)%16:-1);if(sv==1)bv=-bv;}
            }
            if(w*h>=64 && uv>=1 && uv<=8) av=mode_symbol(m,m->angles[uv-1],7,known?(index*3+uv+1)%7:-1)-3;
        }
    }
    int values[9]={copy,chroma,allowed,y,uv,ay,av,au,bv};memcpy(result,values,sizeof(values));
    for(int yy=r;yy<r+h/4;yy++) for(int x=c;x<c+w/4;x++) m->modes[yy][x]=(unsigned char)y;
}
static void run_modes(mode_codec *m,FILE *out,int scenario) {
    int sizes[12][2]={{8,8},{8,8},{8,8},{16,16},{4,4},{4,16},{4,8},{64,64},{4,4},{128,128},{8,4},{16,4}};
    int w=sizes[scenario][0],h=sizes[scenario][1],extent=scenario==9||scenario==3?32:16,first=1,index=0;
    for(int r=0;r<extent;r+=h/4) for(int c=0;c<extent;c+=w/4,index++) {
        int state[9];modes(m,r,c,w,h,scenario==1,scenario==2||scenario==3||scenario==10,scenario==8,index,state);
        if(out) {fprintf(out,"%s[%d,%d,%d,%d",first?"":",",r,c,w,h);for(int i=0;i<9;i++)fprintf(out,",%d",state[i]);fprintf(out,"]");}
        first=0;
    }
}
static void make_modes(FILE *out,int scenario,int updates,int *first) {
    mode_codec enc,dec;mode_init(&enc,0,updates);od_ec_enc_init(&enc.entropy.enc,512);
    run_modes(&enc,NULL,scenario);uint32_t size;unsigned char *bytes=od_ec_enc_done(&enc.entropy.enc,&size);assert(bytes&&!enc.entropy.enc.error);
    mode_init(&dec,1,updates);od_ec_dec_init(&dec.entropy.dec,bytes,size);
    fprintf(out,"%s{\"scenario\":%d,\"updates\":%s,\"hex\":\"",*first?"":",",scenario,updates?"true":"false");*first=0;
    for(uint32_t i=0;i<size;i++)fprintf(out,"%02x",bytes[i]);fprintf(out,"\",\"states\":[");run_modes(&dec,out,scenario);fprintf(out,"],\"lumaContextVisits\":[");
    for(int i=0;i<25;i++)fprintf(out,"%s%d",i?",":"",dec.contexts[i]);fprintf(out,"]}");
    assert(memcmp(enc.y,dec.y,sizeof(enc.y))==0&&memcmp(enc.uv,dec.uv,sizeof(enc.uv))==0&&memcmp(enc.angles,dec.angles,sizeof(enc.angles))==0&&memcmp(enc.alphas,dec.alphas,sizeof(enc.alphas))==0);
    od_ec_enc_clear(&enc.entropy.enc);
}
static void mode_prefix(char **argv) {
    mode_codec m;mode_init(&m,1,1);unsigned char *bytes;int skip,ci,q;
    int pixels=first_leaf(argv,&m.entropy,&bytes,&skip,&ci,&q),values[9];
    modes(&m,0,0,pixels,pixels,atoi(argv[10]),0,0,-1,values);
    FILE *out=fopen(argv[9],"wb");assert(out);fprintf(out,"{\"pixels\":%d,\"preludeQ\":%d,\"modes\":[",pixels,q);
    for(int i=0;i<9;i++)fprintf(out,"%s%d",i?",":"",values[i]);fprintf(out,"]}\n");fclose(out);free(bytes);
}
int main(int argc,char **argv) {
    if(argc==11) {mode_prefix(argv);return 0;}assert(argc==2);
    FILE *out=fopen(argv[1],"wb");assert(out);fprintf(out,"{\"producer\":\"AOM v3.13.1 tables/entropy, original intra-mode syntax harness\",\"nativeSelfCheck\":true,\"cases\":[");
    int first=1;for(int s=0;s<12;s++)for(int u=0;u<=1;u++)make_modes(out,s,u,&first);
    fprintf(out,"]}\n");fclose(out);return 0;
}
