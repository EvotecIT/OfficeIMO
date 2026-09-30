/* Original normative coefficient component harness. Native entropy/defaults/scan data are pinned AOM.
 * Full spatial grids deliberately differ from the managed border-overlay representation. */
#define OFFICEIMO_AV1_COEFFICIENT_INCLUDE
#include "GenerateTransformFixtures.c"
#include "coefficient-defaults.h"
#include "coefficient-scans.h"

typedef struct {
    transform_codec tx;
    aom_cdf_prob skip[65][3],dc[6][3],extra[90][3],base[420][5],br[210][5],last[40][4];
    aom_cdf_prob eob[7][4][12],intra[156][17],inter[16][17];
    unsigned char level[3][64][64],sign[3][64][64],types[64][64];
    int q,segment_q,reduced,copy,lossless,chroma,ymode,uvmode,filter,seed,serial;
    int type_hits[16],size_hits[19],zero_hits[2],eob_hits[11],level_hits[5],sign_hits[3],q_context;
} coefficient_codec;
typedef struct {int type,eob,w,h,values[1024];} coefficient_result;
static int txclass(int type) {return type==10||type==12||type==14?2:type==11||type==13||type==15?1:0;}
static int square(int size) {return log2n((txw[size]<txh[size]?txw[size]:txh[size])/4);}
static int txset(coefficient_codec *c,int size) {
    int up=log2n((txw[size]>txh[size]?txw[size]:txh[size])/4),sq=square(size);
    if(up>3)return 0;
    if(c->copy)return c->reduced||up==3?3:sq==2?2:1;
    return up==3?0:c->reduced||sq==2?2:1;
}
static const int inverse_intra[3][7]={{0},{9,0,10,11,3,1,2},{9,0,3,1,2}};
static const int inverse_inter[4][16]={{0},{9,10,11,12,13,14,15,0,1,2,4,5,3,6,7,8},{9,10,11,0,1,2,4,5,3,6,7,8},{9,0}};
static const unsigned short intra_sets[3]={1,0xe0f,0x20f},inter_sets[4]={1,0xffff,0xfff,0x201};
static unsigned random_value(coefficient_codec *c) {
    unsigned v=(unsigned)(++c->serial + c->seed*7919);v^=v>>16;v*=0x7feb352dU;v^=v>>15;v*=0x846ca68bU;v^=v>>16;return v;
}
static int csym(coefficient_codec *c,aom_cdf_prob *cdf,int n,int wanted) {
    return mode_symbol(&c->tx.entropy,cdf,n,c->tx.entropy.entropy.decode?-1:wanted);
}
static int cbit(coefficient_codec *c,int wanted) {
    codec *e=&c->tx.entropy.entropy;
    if(e->decode)return od_ec_decode_bool_q15(&e->dec,16384);
    od_ec_encode_bool_q15(&e->enc,wanted,16384);return wanted;
}
static void coeff_init(coefficient_codec *c,int decode,int updates,int rows,int cols,int q) {
    memset(c,0,sizeof(*c));tx_init(&c->tx,decode,updates,rows,cols);c->q=q;c->segment_q=q;
    int group=q<=20?0:q<=60?1:q<=120?2:3;c->q_context=group;
    memcpy(c->skip,coeff_Skip+group*65,sizeof(c->skip));memcpy(c->dc,coeff_Dc+group*6,sizeof(c->dc));
    memcpy(c->extra,coeff_Extra+group*90,sizeof(c->extra));memcpy(c->base,coeff_Base+group*420,sizeof(c->base));
    memcpy(c->br,coeff_Br+group*210,sizeof(c->br));memcpy(c->last,coeff_Last+group*40,sizeof(c->last));
    memcpy(c->intra,coeff_Intra,sizeof(c->intra));memcpy(c->inter,coeff_Inter,sizeof(c->inter));
    const aom_cdf_prob *tables[7]={&coeff_Eob0[group*4][0],&coeff_Eob1[group*4][0],&coeff_Eob2[group*4][0],
        &coeff_Eob3[group*4][0],&coeff_Eob4[group*4][0],&coeff_Eob5[group*4][0],&coeff_Eob6[group*4][0]};
    for(int i=0;i<7;i++)for(int j=0;j<4;j++)memcpy(c->eob[i][j],tables[i]+j*(i+6),(size_t)(i+6)*sizeof(aom_cdf_prob));
}
static const int *scan_for(int size,int cls) {
    int w=txw[size]>32?32:txw[size],h=txh[size]>32?32:txh[size];
    int adjusted=tx_find(w,h);
#define SCAN(id,shape) case id: return cls==2?mrow_scan_##shape:cls==1?mcol_scan_##shape:default_scan_##shape;
    switch(adjusted) {
        SCAN(0,4x4) SCAN(1,8x8) SCAN(2,16x16) SCAN(3,32x32)
        SCAN(5,4x8) SCAN(6,8x4) SCAN(7,8x16) SCAN(8,16x8) SCAN(9,16x32) SCAN(10,32x16)
        SCAN(13,4x16) SCAN(14,16x4) SCAN(15,8x32) SCAN(16,32x8)
        default:assert(0);return NULL;
    }
#undef SCAN
}
static int read_type(coefficient_codec *c,residual_block b) {
    int set=txset(c,b.size),type=0,sq=square(b.size);
    if(set && c->segment_q>0) {
        if(c->copy) {
            int n=set==1?16:set==2?12:2;
            type=inverse_inter[set][csym(c,c->inter[set*4+sq],n,(int)(random_value(c)%(unsigned)n))];
        } else {
            const int dirs[5]={0,1,2,6,0};int dir=c->filter>=0?dirs[c->filter]:c->ymode,n=set==1?7:5;
            type=inverse_intra[set][csym(c,c->intra[(set*4+sq)*13+dir],n,(int)(random_value(c)%(unsigned)n))];
        }
    }
    for(int y=b.y/4;y<(b.y+txh[b.size])/4;y++)for(int x=b.x/4;x<(b.x+txw[b.size])/4;x++)c->types[y][x]=(unsigned char)type;
    return type;
}
static int plane_type(coefficient_codec *c,residual_block b) {
    if(c->lossless||txw[b.size]>32||txh[b.size]>32)return 0;
    if(!b.plane)return c->types[b.y/4][b.x/4];
    const int map[14]={0,1,2,0,3,1,2,2,1,3,1,2,3,0};
    int type=c->copy?c->types[b.y/2>c->tx.r?b.y/2:c->tx.r][b.x/2>c->tx.c?b.x/2:c->tx.c]:map[c->uvmode];
    return ((c->copy?inter_sets:intra_sets)[txset(c,b.size)]&(1u<<type))?type:0;
}
static int border_value(coefficient_codec *c,residual_block b,int above,int index,int dc) {
    int sub=b.plane>0,x=b.x/4,y=b.y/4;
    if(above) {if(!y || x+index>=(c->tx.cols>>sub))return 0;y--;x+=index;}
    else {if(!x || y+index>=(c->tx.rows>>sub))return 0;x--;y+=index;}
    return dc?c->sign[b.plane][y][x]:c->level[b.plane][y][x];
}
static int skip_context(coefficient_codec *c,residual_block b) {
    int top=0,left=0;
    for(int i=0;i<txw[b.size]/4;i++) {int v=border_value(c,b,1,i,0);top=b.plane?top|v|border_value(c,b,1,i,1):(v>top?v:top);}
    for(int i=0;i<txh[b.size]/4;i++) {int v=border_value(c,b,0,i,0);left=b.plane?left|v|border_value(c,b,0,i,1):(v>left?v:left);}
    if(!b.plane) {
        if(bw[c->tx.id]==txw[b.size]&&bh[c->tx.id]==txh[b.size])return 0;
        if(!top&&!left)return 1;if(!top||!left)return 2+(top>3||left>3);
        if(top<=3&&left<=3)return 4;if(top<=3||left<=3)return 5;return 6;
    }
    int w=bw[c->tx.id]/2,h=bh[c->tx.id]/2;if(w<4)w=4;if(h<4)h=4;
    return 7+(top!=0)+(left!=0)+3*(w*h>txw[b.size]*txh[b.size]);
}
static const int offsets[3][5][5]={
    {{0,1,6,6,21},{1,6,6,21,21},{6,6,21,21,21},{6,21,21,21,21},{21,21,21,21,21}},
    {{0,11,11,11,11},{11,11,11,11,11},{6,6,21,21,21},{6,21,21,21,21},{21,21,21,21,21}},
    {{0,16,6,6,21},{16,16,6,21,21},{16,16,21,21,21},{16,16,21,21,21},{16,16,21,21,21}}};
static int base_context(coefficient_result *r,int size,int pos,int cls,int br) {
    const int a[3][5][2]={{{0,1},{1,0},{1,1},{0,2},{2,0}},{{0,1},{1,0},{0,2},{0,3},{0,4}},{{0,1},{1,0},{2,0},{3,0},{4,0}}};
    const int b[3][3][2]={{{0,1},{1,0},{1,1}},{{0,1},{1,0},{0,2}},{{0,1},{1,0},{2,0}}};
    int row=pos/r->w,col=pos%r->w,mag=0;
    for(int i=0;i<(br?3:5);i++) {
        int y=row+(br?b[cls][i][0]:a[cls][i][0]),x=col+(br?b[cls][i][1]:a[cls][i][1]);
        if(y<r->h&&x<r->w) {int v=r->values[y*r->w+x],cap=br?15:3;mag+=v<cap?v:cap;}
    }
    mag=(mag+1)/2;if(mag>(br?6:4))mag=br?6:4;
    if(br)return mag+(!pos?0:(cls==0?row<2&&col<2:cls==1?col==0:row==0)?7:14);
    if(cls)return mag+26+5*((cls==2?row:col)>2?2:cls==2?row:col);
    if(!pos)return 0;
    int shape=txw[size]<txh[size]?1:txw[size]>txh[size]?2:0;
    return mag+offsets[shape][row>4?4:row][col>4?4:col];
}
static int sign_context(coefficient_codec *c,residual_block b) {
    int sum=0;
    for(int a=0;a<2;a++)for(int i=0;i<(a?txw[b.size]:txh[b.size])/4;i++) {
        int v=border_value(c,b,a,i,1);sum+=v==1?-1:v==2?1:0;
    }
    return sum<0?1:sum>0?2:0;
}
static int choose_eob(coefficient_codec *c,int area) {
    unsigned v=random_value(c);int point=(int)(v%(unsigned)(log2n(area)+1));
    if(point==0)return 1;if(point==1)return 2;
    int minimum=(1<<(point-1))+1,maximum=1<<point;
    return minimum+(int)(random_value(c)%(unsigned)(maximum-minimum+1));
}
static void coefficients(coefficient_codec *c,residual_block b,int skip,coefficient_result *r) {
    memset(r,0,sizeof(*r));r->w=txw[b.size]>32?32:txw[b.size];r->h=txh[b.size]>32?32:txh[b.size];c->size_hits[b.size]++;
    if(skip)return;
    int up=log2n((txw[b.size]>txh[b.size]?txw[b.size]:txh[b.size])/4),tc=(square(b.size)+up+1)/2,pt=b.plane>0;
    int zero=csym(c,c->skip[tc*13+skip_context(c,b)],2,random_value(c)%7==0);c->zero_hits[zero]++;
    int total=0,dc=0;
    if(!zero) {
        if(!b.plane)read_type(c,b);r->type=plane_type(c,b);c->type_hits[r->type]++;
        int cls=txclass(r->type),area=r->w*r->h,multi=log2n(area)-4,eob=choose_eob(c,area);
        int point=eob==1?1:log2n(eob-1)+2;
        point=csym(c,c->eob[multi][pt*2+(cls!=0)],multi+5,point-1)+1;c->eob_hits[point-1]++;
        int start=point<2?point:(1<<(point-2))+1;r->eob=start;
        if(point>=3) {
            int extra=eob-start;
            if(csym(c,c->extra[(tc*2+pt)*9+point-3],2,(extra>>(point-3))&1))r->eob+=1<<(point-3);
            for(int s=point-4;s>=0;s--)if(cbit(c,(extra>>s)&1))r->eob+=1<<s;
        }
        const int *scan=scan_for(b.size,cls);
        for(int i=r->eob-1;i>=0;i--) {
            int pos=scan[i],level;
            if(i==r->eob-1) {
                int ctx=i==0?0:i<=area/8?1:i<=area/4?2:3;
                level=csym(c,c->last[(tc*2+pt)*4+ctx],3,(int)(random_value(c)%3))+1;
            } else level=csym(c,c->base[(tc*2+pt)*42+base_context(r,b.size,pos,cls,0)],4,(int)(random_value(c)%4));
            if(level>2)for(int j=0;j<4;j++) {
                int increment=csym(c,c->br[((tc>3?3:tc)*2+pt)*21+base_context(r,b.size,pos,cls,1)],4,(int)(random_value(c)%4));
                level+=increment;if(increment<3)break;
            }
            r->values[pos]=level;
        }
        for(int i=0;i<r->eob;i++) {
            int pos=scan[i],level=r->values[pos],negative=0;
            if(level)negative=i?cbit(c,(int)(random_value(c)%2)):csym(c,c->dc[pt*3+sign_context(c,b)],2,(int)(random_value(c)%2));
            if(level>14) {
                int length=1+(int)(random_value(c)%20);
                if(c->tx.entropy.entropy.decode) {length=1;while(!cbit(c,0)) {assert(length<20);length++;}}
                else {for(int j=1;j<length;j++)cbit(c,0);cbit(c,1);}
                int v=1;for(int j=length-2;j>=0;j--)v=(v<<1)|cbit(c,(int)(random_value(c)%2));level=v+14;
            }
            if(!pos&&level)dc=negative?1:2;
            int mask=level&0xfffff;total+=mask;r->values[pos]=negative?-mask:mask;
            c->level_hits[level==0?0:level==1?1:level==2?2:level<=14?3:4]++;
        }
    } else if(!b.plane)for(int y=b.y/4;y<(b.y+txh[b.size])/4;y++)for(int x=b.x/4;x<(b.x+txw[b.size])/4;x++)c->types[y][x]=0;
    if(total>63)total=63;c->sign_hits[dc]++;
    for(int y=b.y/4;y<(b.y+txh[b.size])/4;y++)for(int x=b.x/4;x<(b.x+txw[b.size])/4;x++) {
        c->level[b.plane][y][x]=(unsigned char)total;c->sign[b.plane][y][x]=(unsigned char)dc;
    }
}
static void print_coeff(FILE *out,residual_block b,coefficient_result *r) {
    fprintf(out,"{\"block\":[%d,%d,%d,%d],\"type\":%d,\"eob\":%d,\"values\":[",b.plane,b.x,b.y,b.size,r->type,r->eob);
    int first=1;for(int i=0;i<r->w*r->h;i++)if(r->values[i]) {fprintf(out,"%s[%d,%d]",first?"":",",i,r->values[i]);first=0;}fprintf(out,"]}");
}
static void run_coefficients(coefficient_codec *c,FILE *out,int scenario) {
    int id=scenario%22,index=0,first=1;int mode=scenario<22?1:2;
    for(int y=0;y<c->tx.rows;y+=bh[id]/4)for(int x=0;x<c->tx.cols;x+=bw[id]/4,index++) {
        c->copy=scenario/22%3==1;c->reduced=scenario/22%3==2;
        int segment=index%7==2?1:index%7==4?2:0;c->segment_q=segment==1?0:segment==2?255:c->q;c->lossless=c->segment_q==0;
        c->chroma=!(bh[id]==4&&!(y&1))&&!(bw[id]==4&&!(x&1));
        c->ymode=(index*7+scenario)%13;c->uvmode=(index*11+scenario)%14;c->filter=!c->copy && c->ymode==0 && index%8<5?index%8:-1;
        int skip=index%9==3;
        tx_sizes(&c->tx,y,x,id,mode,c->lossless,c->copy,skip);residual_layout layout;residuals(&c->tx,c->lossless,c->copy,c->chroma,&layout);
        if(skip)for(int p=0;p<1+2*c->chroma;p++) {
            int sub=p>0;for(int yy=y>>sub;yy<(y+bh[id]/4)>>sub;yy++)for(int xx=x>>sub;xx<(x+bw[id]/4)>>sub;xx++)c->level[p][yy][xx]=c->sign[p][yy][xx]=0;
        }
        if(out)fprintf(out,"%s{\"block\":[%d,%d,%d,%d],\"segment\":%d,\"skip\":%s,\"lossless\":%s,\"copy\":%s,\"reduced\":%s,\"chroma\":%s,\"y\":%d,\"uv\":%d,\"filter\":%d,\"coefficients\":[",first?"":",",y,x,bw[id],bh[id],segment,skip?"true":"false",c->lossless?"true":"false",c->copy?"true":"false",c->reduced?"true":"false",c->chroma?"true":"false",c->ymode,c->uvmode,c->filter);
        first=0;
        for(int i=0;i<layout.count;i++) {
            coefficient_result r;coefficients(c,layout.blocks[i],skip,&r);
            if(out) {if(i)fprintf(out,",");print_coeff(out,layout.blocks[i],&r);}
        }
        if(out)fprintf(out,"]}");
    }
}
static void make_coefficients(FILE *out,int scenario,int q,int updates,int *first) {
    int id=scenario%22,extent=bw[id]>=64||bh[id]>=64?32:8;if(scenario/22==2)extent=18;
    coefficient_codec *enc=malloc(sizeof(*enc)),*dec=malloc(sizeof(*dec));assert(enc&&dec);
    coeff_init(enc,0,updates,extent,extent,q);enc->seed=scenario;enc->tx.index=scenario*97;
    od_ec_enc_init(&enc->tx.entropy.entropy.enc,4096);run_coefficients(enc,NULL,scenario);
    uint32_t n;unsigned char *bytes=od_ec_enc_done(&enc->tx.entropy.entropy.enc,&n);assert(bytes&&n&&!enc->tx.entropy.entropy.enc.error);
    coeff_init(dec,1,updates,extent,extent,q);dec->seed=scenario;dec->tx.index=scenario*97;
    od_ec_dec_init(&dec->tx.entropy.entropy.dec,bytes,n);
    fprintf(out,"%s{\"scenario\":%d,\"q\":%d,\"mode\":%d,\"extent\":%d,\"updates\":%s,\"hex\":",*first?"":",",scenario,q,scenario<22?1:2,extent,updates?"true":"false");*first=0;hex(out,bytes,n);
    fprintf(out,",\"states\":[");run_coefficients(dec,out,scenario);fprintf(out,"],\"typeHits\":[");
    for(int i=0;i<16;i++)fprintf(out,"%s%d",i?",":"",dec->type_hits[i]);fprintf(out,"],\"sizeHits\":[");
    for(int i=0;i<19;i++)fprintf(out,"%s%d",i?",":"",dec->size_hits[i]);fprintf(out,"],\"levelHits\":[");
    for(int i=0;i<5;i++)fprintf(out,"%s%d",i?",":"",dec->level_hits[i]);fprintf(out,"],\"eobHits\":[");
    for(int i=0;i<11;i++)fprintf(out,"%s%d",i?",":"",dec->eob_hits[i]);fprintf(out,"],\"zeroHits\":[%d,%d],\"signHits\":[%d,%d,%d]}",dec->zero_hits[0],dec->zero_hits[1],dec->sign_hits[0],dec->sign_hits[1],dec->sign_hits[2]);
    assert(memcmp(enc->base,dec->base,sizeof(enc->base))==0&&memcmp(enc->level,dec->level,sizeof(enc->level))==0&&memcmp(enc->sign,dec->sign,sizeof(enc->sign))==0&&memcmp(enc->types,dec->types,sizeof(enc->types))==0);
    od_ec_enc_clear(&enc->tx.entropy.entropy.enc);free(enc);free(dec);
}
static void coefficient_prefix(char **argv) {
    palette_codec p;unsigned char *bytes;int q,modes[9];palette_result palette;
    int pixels=palette_first_leaf(argv,&p,&bytes,&q,modes,&palette);
    coefficient_codec *c=malloc(sizeof(*c));assert(c);coeff_init(c,1,1,atoi(argv[4]),atoi(argv[5]),atoi(argv[6]));
    c->tx.entropy.entropy.dec=p.mode.entropy.dec;c->chroma=modes[1];c->ymode=modes[3];c->uvmode=modes[4];c->filter=palette.f;
    tx_sizes(&c->tx,0,0,block_id(pixels,pixels),atoi(argv[12]),0,0,0);
    residual_layout layout;residuals(&c->tx,0,0,c->chroma,&layout);coefficient_result result;coefficients(c,layout.blocks[0],0,&result);
    FILE *out=fopen(argv[9],"wb");assert(out);fprintf(out,"{\"pixels\":%d,\"preludeQ\":%d,\"coefficient\":",pixels,q);print_coeff(out,layout.blocks[0],&result);fprintf(out,"}\n");fclose(out);free(bytes);free(c);
}
static void golomb_boundary(FILE *out,int malformed) {
    coefficient_codec *c=malloc(sizeof(*c));assert(c);coeff_init(c,0,0,2,2,20);
    od_ec_enc_init(&c->tx.entropy.entropy.enc,128);
    // One nonzero 4x4 luma transform, DC-only eob, DCT type and four saturated base-range symbols.
    csym(c,c->skip[0],2,0);csym(c,c->intra[52],7,1);csym(c,c->eob[0][0],5,0);
    csym(c,c->last[0],3,2);for(int i=0;i<4;i++)csym(c,c->br[0],4,3);csym(c,c->dc[0],2,0);
    if(malformed) {for(int i=0;i<20;i++)cbit(c,0);}
    else {
        for(int i=0;i<19;i++)cbit(c,0);cbit(c,1);
        // q = (2^20 - 14) + 14 wraps to zero; DC category must still be positive before masking.
        int value=(1<<20)-14;for(int i=18;i>=0;i--)cbit(c,(value>>i)&1);
        // The adjacent DC coefficient uses positive left DC category even though its masked level was zero.
        csym(c,c->skip[0],2,0);csym(c,c->intra[52],7,1);csym(c,c->eob[0][0],5,0);csym(c,c->last[0],3,0);csym(c,c->dc[2],2,1);
    }
    uint32_t n;unsigned char *bytes=od_ec_enc_done(&c->tx.entropy.entropy.enc,&n);assert(bytes&&n);
    fprintf(out,"%s",malformed?",\"malformedGolombHex\":":" ,\"maskedDcHex\":");hex(out,bytes,n);
    od_ec_enc_clear(&c->tx.entropy.entropy.enc);free(c);
}
int main(int argc,char **argv) {
    if(argc==13){coefficient_prefix(argv);return 0;}assert(argc==2);FILE *out=fopen(argv[1],"wb");assert(out);
    fprintf(out,"{\"producer\":\"Pinned AOM v3.13.1 entropy/defaults/scans, original coefficient grammar and full spatial grids\",\"nativeSelfCheck\":true,\"cases\":[");
    int first=1;const int qs[5]={0,20,60,120,255};
    for(int scenario=0;scenario<66;scenario++)for(int q=0;q<5;q++)for(int u=0;u<2;u++)make_coefficients(out,scenario,qs[q],u,&first);
    fprintf(out,"]");golomb_boundary(out,0);golomb_boundary(out,1);fprintf(out,"}\n");fclose(out);return 0;
}
