/* Original component-syntax harness over pinned official AOM entropy and CDF data.
 * Does not decode prediction, residuals or pixels and is not a complete AV1 oracle. */
#include <assert.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "aom_dsp/entenc.h"
#include "aom_dsp/entdec.h"
#include "aom_dsp/prob.h"
#include "partition-defaults.h"
#include "prelude-defaults.h"

typedef struct {
    od_ec_enc enc; od_ec_dec dec;
    int decode, updates, q, lf[4];
    aom_cdf_prob skips[3][3], segments[3][9], dq[5], dlf[5], multi[4][5];
    unsigned char skip_map[64][64], segment_map[64][64];
} codec;

static void init(codec *c, int decode, int updates) {
    memset(c, 0, sizeof(*c)); c->decode = decode; c->updates = updates; c->q = 120;
    memcpy(c->skips, default_skip_txfm_cdfs, sizeof(c->skips));
    memcpy(c->segments, default_spatial_pred_seg_tree_cdf, sizeof(c->segments));
    memcpy(c->dq, default_delta_q_cdf, sizeof(c->dq));
    memcpy(c->dlf, default_delta_lf_cdf, sizeof(c->dlf));
    memcpy(c->multi, default_delta_lf_multi_cdf, sizeof(c->multi));
}
static int symbol(codec *c, aom_cdf_prob *cdf, int n, int expected) {
    int value = expected;
    if (c->decode) { value = od_ec_decode_cdf_q15(&c->dec, cdf, n); assert(value == expected); }
    else od_ec_encode_cdf_q15(&c->enc, value, cdf, n);
    if (c->updates) update_cdf(cdf, value, n);
    return value;
}
static int literal(codec *c, int bits, int expected) {
    int value = 0;
    for (int i = bits - 1; i >= 0; i--) {
        int bit = (expected >> i) & 1;
        if (c->decode) { int read = od_ec_decode_bool_q15(&c->dec, 16384); assert(read == bit); bit = read; }
        else od_ec_encode_bool_q15(&c->enc, bit, 16384);
        value = (value << 1) | bit;
    }
    return value;
}
static int deinterleave(int diff, int ref) {
    if (!ref) return diff;
    if (ref == 7) return 7 - diff;
    if (2 * ref < 8) {
        if (diff > 2 * ref) return diff;
    } else if (diff > 2 * (7 - ref)) return 7 - diff;
    return diff & 1 ? ref + ((diff + 1) / 2) : ref - diff / 2;
}
static int segment(codec *c, int row, int col, int skip, int i) {
    int u = row ? c->segment_map[row-1][col] : -1, l = col ? c->segment_map[row][col-1] : -1;
    int ul = row && col ? c->segment_map[row-1][col-1] : -1;
    int pred = u < 0 ? l < 0 ? 0 : l : l < 0 ? u : ul == u ? u : l;
    int ctx = ul < 0 ? 0 : ul == u && ul == l ? 2 : ul == u || ul == l || u == l ? 1 : 0;
    return skip ? pred : deinterleave(symbol(c, c->segments[ctx], 8, (i * 3 + i / 4) % 8), pred);
}
static int delta(codec *c, aom_cdf_prob *cdf, int index) {
    int abs = symbol(c, cdf, 4, index % 4);
    if (abs == 3) {
        int bits = literal(c, 3, (index / 4) % 8) + 1;
        abs = literal(c, bits, ((1 << bits) - 1) / 2) + (1 << bits) + 1;
    }
    return abs && literal(c, 1, (index / 3) % 2) ? -abs : abs;
}
static void run(codec *c, FILE *out, int sb, int scenario) {
    int units = sb / 4, pre = scenario == 4, seg = scenario == 3 || pre;
    int cdef = scenario != 5 && scenario != 6, full = scenario == 7 || scenario == 9;
    int width = full ? units : 8, height = full ? units : scenario == 8 ? 4 : 8;
    int first = 1, i = 0;
    if (scenario == 5) c->q = 0;
    // Four superblocks, raster order. Rectangular leaf rows are covered exactly once.
    for (int sr = 0; sr < 2*units; sr += units) for (int sc = 0; sc < 2*units; sc += units) {
        int ci[2][2] = {{-1,-1},{-1,-1}}, first_leaf = 1;
        for (int r = sr; r < sr+units; r += height) for (int col = sc; col < sc+units; col += width, i++) {
            int s = seg && pre ? segment(c,r,col,0,i) : 0;
            int ctx = (r ? c->skip_map[r-1][col] : 0) + (col ? c->skip_map[r][col-1] : 0);
            int skip = pre && s == 1 ? 1 : symbol(c,c->skips[ctx],2,scenario==7 ? 1 : scenario==9 ? 0 : i%5==0 || i%7==0);
            if (seg && !pre) s = segment(c,r,col,skip,i);
            int rr=(r-sr)/16, cc=(col-sc)/16;
            if (cdef && !skip && ci[rr][cc]<0) {
                int value=literal(c,2,(i/3)%4);
                for(int y=rr;y<rr+(height+15)/16;y++) for(int x=cc;x<cc+(width+15)/16;x++) ci[y][x]=value;
            }
            if(first_leaf && scenario != 5 && !(full && skip)) {
                int q_step=(sr/units*2+sc/units)*5+scenario;
                c->q=clamp(c->q+delta(c,c->dq,q_step)*4,1,255);
                int count=scenario==6?0:scenario==2?1:scenario==1?2:4;
                for(int k=0;k<count;k++) c->lf[k]=clamp(c->lf[k]+delta(c,scenario==2?c->dlf:c->multi[k],i+k+1)*2,-63,63);
            }
            first_leaf=0;
            if(out) fprintf(out,"%s[%d,%d,%d,%d,%d,%d,%d,%d,%d,%d,%d,%d,%d]",first?"":",",r,col,width*4,height*4,skip,s,scenario==5 || (seg&&s==2),ci[rr][cc],c->q,c->lf[0],c->lf[1],c->lf[2],c->lf[3]);
            first=0;
            for(int y=r;y<r+height;y++) for(int x=col;x<col+width;x++) { c->skip_map[y][x]=(unsigned char)skip; c->segment_map[y][x]=(unsigned char)s; }
        }
    }
}
static void make_case(FILE *out,int sb,int scenario,int updates,int *first) {
    codec enc,dec; init(&enc,0,updates); od_ec_enc_init(&enc.enc,256);
    run(&enc,NULL,sb,scenario); uint32_t size; unsigned char *bytes=od_ec_enc_done(&enc.enc,&size); assert(bytes&&!enc.enc.error);
    init(&dec,1,updates); od_ec_dec_init(&dec.dec,bytes,size);
    fprintf(out,"%s{\"superblock\":%d,\"scenario\":%d,\"updates\":%s,\"hex\":\"",*first?"":",",sb,scenario,updates?"true":"false"); *first=0;
    for(uint32_t i=0;i<size;i++) fprintf(out,"%02x",bytes[i]);
    fprintf(out,"\",\"states\":["); run(&dec,out,sb,scenario); fprintf(out,"]}");
    assert(enc.q==dec.q && memcmp(enc.lf,dec.lf,sizeof(enc.lf))==0 && memcmp(enc.skips,dec.skips,sizeof(enc.skips))==0 && memcmp(enc.segments,dec.segments,sizeof(enc.segments))==0);
    od_ec_enc_clear(&enc.enc);
}
/* Shared frozen first-leaf setup; payload remains alive until the caller frees it. */
static int first_leaf(char **argv, codec *state, unsigned char **payload, int *skip_out, int *ci_out, int *q_out) {
    FILE *in=fopen(argv[1],"rb"); assert(in); int offset=atoi(argv[2]),size=atoi(argv[3]); assert(offset>=0&&size>0);
    unsigned char *bytes=malloc((size_t)size); assert(bytes);
    assert(fseek(in,offset,SEEK_SET)==0 && fread(bytes,1,(size_t)size,in)==(size_t)size); fclose(in);
    int rows=atoi(argv[4]),cols=atoi(argv[5]),q=atoi(argv[6]),cdef=atoi(argv[7]),dq=atoi(argv[8]);
    codec c; init(&c,1,1); od_ec_dec_init(&c.dec,bytes,(uint32_t)size);
    int pixels=64;
    for(;;) {
        int half=pixels/8,hr=half<rows,hc=half<cols,kind=0;
        if(pixels>=8) {
            if(!hr&&!hc) kind=3;
            else {
                int index=0;for(int p=pixels;p>8;p>>=1) index++;
                int n=index==0?4:index==4?8:10;
                aom_cdf_prob cdf[11]; memcpy(cdf,default_partition_cdf[index*4],sizeof(cdf));
                // Frozen prefix descents are either full partitions or forced splits.
                assert(hr&&hc); kind=od_ec_decode_cdf_q15(&c.dec,cdf,n);
            }
        }
        if(kind!=3) { assert(kind==0); break; } pixels>>=1;
    }
    int skip=od_ec_decode_cdf_q15(&c.dec,c.skips[0],2),ci=!skip&&cdef?0:-1;
    if(dq && !(pixels==64&&skip)) {
        int abs=od_ec_decode_cdf_q15(&c.dec,c.dq,4);
        if(abs==3) { int bits=0;for(int i=0;i<3;i++) bits=(bits<<1)|od_ec_decode_bool_q15(&c.dec,16384);bits++;
            abs=0;for(int i=0;i<bits;i++) abs=(abs<<1)|od_ec_decode_bool_q15(&c.dec,16384);abs+=(1<<bits)+1; }
        if(abs && od_ec_decode_bool_q15(&c.dec,16384)) abs=-abs;
        q=clamp(q+abs,1,255);
    }
    *state=c; *payload=bytes; *skip_out=skip; *ci_out=ci; *q_out=q; return pixels;
}
static void prefix(char **argv) {
    codec c; unsigned char *bytes; int skip,ci,q;
    int pixels=first_leaf(argv,&c,&bytes,&skip,&ci,&q);
    FILE *out=fopen(argv[9],"wb");assert(out);
    fprintf(out,"{\"pixels\":%d,\"skip\":%d,\"segment\":0,\"cdefIndex\":%d,\"q\":%d}\n",pixels,skip,ci,q);
    fclose(out);free(bytes);
}
int main(int argc,char **argv) {
    if(argc==10) { prefix(argv);return 0; } assert(argc==2);
    FILE *out=fopen(argv[1],"wb");assert(out);
    fprintf(out,"{\"producer\":\"AOM v3.13.1 tables and entropy; original prelude-syntax harness\",\"nativeSelfCheck\":true,\"cases\":[");
    int first=1;for(int scenario=0;scenario<=9;scenario++) for(int updates=0;updates<=1;updates++) make_case(out,64,scenario,updates,&first);
    for(int scenario=0;scenario<=9;scenario+=9) for(int updates=0;updates<=1;updates++) make_case(out,128,scenario,updates,&first);
    fprintf(out,"]}\n");fclose(out);return 0;
}
