/* Original normative AV1 transform-size/traversal harness over pinned AOM entropy and CDF defaults. */
#define OFFICEIMO_AV1_TRANSFORM_INCLUDE
#include "GeneratePaletteFixtures.c"
#include "transform-defaults.h"

static const int txw[19]={4,8,16,32,64,4,8,8,16,16,32,32,64,4,16,8,32,16,64};
static const int txh[19]={4,8,16,32,64,8,4,16,8,32,16,64,32,16,4,32,8,64,16};
static const int split_tx[19]={0,0,1,2,3,0,0,1,1,2,2,3,3,5,6,7,8,9,10};
static const int bw[22]={4,4,8,8,8,16,16,16,32,32,32,64,64,64,128,128,4,16,8,32,16,64};
static const int bh[22]={4,8,4,8,16,8,16,32,16,32,64,32,64,128,64,128,16,4,32,8,64,16};
static const int max_rect[22]={0,5,6,1,7,8,2,9,10,3,11,12,4,4,4,4,13,14,15,16,17,18};
static const int max_depth[22]={0,1,1,1,2,2,2,3,3,3,4,4,4,4,4,4,2,2,3,3,4,4};
typedef struct {
    mode_codec entropy;
    aom_cdf_prob depth[4][3][4],split[21][3];
    unsigned char grid[64][64],inter[64][64],skip[64][64],block[64][64];
    int depth_contexts[12],split_contexts[21],rows,cols,r,c,id,last,index;
} transform_codec;
typedef struct {int plane,x,y,size;} residual_block;
typedef struct {residual_block blocks[1536];int count;} residual_layout;
static void tx_init(transform_codec *t,int decode,int updates,int rows,int cols) {
    memset(t,0,sizeof(*t));mode_init(&t->entropy,decode,updates);t->rows=rows;t->cols=cols;
    memcpy(t->depth,default_tx_size_cdf,sizeof(t->depth));memcpy(t->split,default_txfm_partition_cdf,sizeof(t->split));
}
static int tx_find(int w,int h) {for(int i=0;i<19;i++)if(txw[i]==w&&txh[i]==h)return i;assert(0);return 0;}
static int tx_symbol(transform_codec *t,aom_cdf_prob *cdf,int n) {
    int expected=-1;
    if(!t->entropy.entropy.decode) {
        unsigned v=(unsigned)(t->index++ + 1);v^=v>>16;v*=0x7feb352dU;v^=v>>15;v*=0x846ca68bU;v^=v>>16;
        expected=(int)(v%(unsigned)n);
    }
    return mode_symbol(&t->entropy,cdf,n,expected);
}
static int above_width(transform_codec *t,int r,int c) {
    if(r==t->r) {if(!r)return 64;if(t->skip[r-1][c]&&t->inter[r-1][c])return bw[t->block[r-1][c]];}
    return txw[t->grid[r-1][c]];
}
static int left_height(transform_codec *t,int r,int c) {
    if(c==t->c) {if(!c)return 64;if(t->skip[r][c-1]&&t->inter[r][c-1])return bh[t->block[r][c-1]];}
    return txh[t->grid[r][c-1]];
}
static void variable_size(transform_codec *t,int r,int c,int size,int depth) {
    if(r>=t->rows||c>=t->cols)return;
    int split=0;
    if(size&&depth<2) {
        int maximum=log2n((bw[t->id]>bh[t->id]?bw[t->id]:bh[t->id])>64?16:(bw[t->id]>bh[t->id]?bw[t->id]:bh[t->id])/4);
        int square_up=log2n((txw[size]>txh[size]?txw[size]:txh[size])/4);
        int context=(square_up!=maximum)*3+(4-maximum)*6+(above_width(t,r,c)<txw[size])+(left_height(t,r,c)<txh[size]);
        t->split_contexts[context]++;
        // Force 64x64 roots to split so later roots have two smaller completed neighbors.
        split=size==4&&depth==0&&!t->entropy.entropy.decode?mode_symbol(&t->entropy,t->split[context],2,1):tx_symbol(t,t->split[context],2);
    }
    if(split) {
        int child=split_tx[size];
        for(int y=0;y<txh[size]/4;y+=txh[child]/4)for(int x=0;x<txw[size]/4;x+=txw[child]/4)variable_size(t,r+y,c+x,child,depth+1);
    } else {
        for(int y=0;y<txh[size]/4;y++)for(int x=0;x<txw[size]/4;x++)t->grid[r+y][c+x]=(unsigned char)size;
        t->last=size;
    }
}
static void tx_sizes(transform_codec *t,int r,int c,int id,int mode,int lossless,int inter,int skip) {
    t->r=r;t->c=c;t->id=id;t->last=max_rect[id];
    if(mode==2&&id&&inter&&!skip&&!lossless) {
        int maximum=max_rect[id];
        for(int y=r;y<r+bh[id]/4;y+=txh[maximum]/4)for(int x=c;x<c+bw[id]/4;x+=txw[maximum]/4)variable_size(t,y,x,maximum,0);
    } else {
        if(lossless)t->last=0;
        else if(id&&mode==2&&(!skip||!inter)) {
            int a=r?(t->inter[r-1][c]?bw[t->block[r-1][c]]:above_width(t,r,c)):0;
            int l=c?(t->inter[r][c-1]?bh[t->block[r][c-1]]:left_height(t,r,c)):0;
            int ctx=(a>=txw[t->last])+(l>=txh[t->last]),category=max_depth[id]-1;
            t->depth_contexts[category*3+ctx]++;
            int depth=tx_symbol(t,t->depth[category][ctx],category?3:2);
            while(depth--)t->last=split_tx[t->last];
        }
        for(int y=r;y<r+bh[id]/4;y++)for(int x=c;x<c+bw[id]/4;x++)t->grid[y][x]=(unsigned char)t->last;
    }
    for(int y=r;y<r+bh[id]/4;y++)for(int x=c;x<c+bw[id]/4;x++) {
        t->inter[y][x]=(unsigned char)inter;t->skip[y][x]=(unsigned char)skip;t->block[y][x]=(unsigned char)id;
    }
}
static void tx_add(transform_codec *t,residual_layout *out,int plane,int x,int y,int size) {
    int sub=plane>0;if(x>=(t->cols*4>>sub)||y>=(t->rows*4>>sub))return;
    assert(out->count<1536);out->blocks[out->count++]=(residual_block){plane,x,y,size};
}
static void residual_tree(transform_codec *t,residual_layout *out,int x,int y,int w,int h) {
    if(x>=t->cols*4||y>=t->rows*4)return;
    int size=t->grid[y/4][x/4];
    if(w<=txw[size]&&h<=txh[size])tx_add(t,out,0,x,y,tx_find(w,h));
    else if(w>h) {residual_tree(t,out,x,y,w/2,h);residual_tree(t,out,x+w/2,y,w/2,h);}
    else if(w<h) {residual_tree(t,out,x,y,w,h/2);residual_tree(t,out,x,y+h/2,w,h/2);}
    else {residual_tree(t,out,x,y,w/2,h/2);residual_tree(t,out,x+w/2,y,w/2,h/2);
        residual_tree(t,out,x,y+h/2,w/2,h/2);residual_tree(t,out,x+w/2,y+h/2,w/2,h/2);}
}
static void residuals(transform_codec *t,int lossless,int inter,int chroma,residual_layout *out) {
    memset(out,0,sizeof(*out));int w=bw[t->id],h=bh[t->id],wc=w>64?w/64:1,hc=h>64?h/64:1;
    for(int cy=0;cy<hc;cy++)for(int cx=0;cx<wc;cx++)for(int p=0;p<1+2*chroma;p++) {
        int sub=p>0,chunk_w=wc>1||hc>1?64:w,chunk_h=wc>1||hc>1?64:h;
        int pw=chunk_w>>sub,ph=chunk_h>>sub;if(pw<4)pw=4;if(ph<4)ph=4;
        int x=((t->c+cx*16)>>sub)*4,y=((t->r+cy*16)>>sub)*4;
        if(inter&&!lossless&&!p)residual_tree(t,out,x,y,pw,ph);
        else {
            int size=lossless?0:t->last;
            if(p&&!lossless) {
                int uw=w/2,uh=h/2;if(uw<4)uw=4;if(uh<4)uh=4;if(uw>64)uw=64;if(uh>64)uh=64;
                size=tx_find(uw,uh);
                if(txw[size]==64||txh[size]==64)size=txw[size]==16?9:txh[size]==16?10:3;
            }
            for(int yy=0;yy<ph;yy+=txh[size])for(int xx=0;xx<pw;xx+=txw[size])tx_add(t,out,p,x+xx,y+yy,size);
        }
    }
}
static void print_layout(FILE *out,transform_codec *t,residual_layout *layout) {
    fprintf(out,"\"last\":%d,\"grid\":\"",t->last);
    for(int y=t->r;y<t->r+bh[t->id]/4;y++)for(int x=t->c;x<t->c+bw[t->id]/4;x++)fprintf(out,"%02x",t->grid[y][x]);
    fprintf(out,"\",\"residuals\":[");
    for(int i=0;i<layout->count;i++) {residual_block *b=&layout->blocks[i];fprintf(out,"%s[%d,%d,%d,%d]",i?",":"",b->plane,b->x,b->y,b->size);}
    fprintf(out,"]");
}
static void one_tx_leaf(transform_codec *t,FILE *out,int r,int c,int id,int mode,int lossless,int inter,int skip,int *first) {
    int chroma=!(bh[id]==4&&!(r&1))&&!(bw[id]==4&&!(c&1));
    tx_sizes(t,r,c,id,mode,lossless,inter,skip);residual_layout layout;residuals(t,lossless,inter,chroma,&layout);
    if(out) {
        fprintf(out,"%s{\"block\":[%d,%d,%d,%d],\"lossless\":%s,\"inter\":%s,\"skip\":%s,\"chroma\":%s,",*first?"":",",r,c,bw[id],bh[id],lossless?"true":"false",inter?"true":"false",skip?"true":"false",chroma?"true":"false");
        print_layout(out,t,&layout);fprintf(out,"}");*first=0;
    }
}
static void mixed_leaves(transform_codec *t,FILE *out,int r,int c,int pixels,int depth,int *index,int *first) {
    int split=pixels>4 && (depth==0?(c==0):(r+c+depth)%3!=0);
    if(split) {
        int half=pixels/2,units=half/4;
        mixed_leaves(t,out,r,c,half,depth+1,index,first);mixed_leaves(t,out,r,c+units,half,depth+1,index,first);
        mixed_leaves(t,out,r+units,c,half,depth+1,index,first);mixed_leaves(t,out,r+units,c+units,half,depth+1,index,first);
    } else {
        int i=(*index)++;one_tx_leaf(t,out,r,c,block_id(pixels,pixels),2,i%17==5,i%3!=0,i%7==3,first);
    }
}
static void run_tx(transform_codec *t,FILE *out,int scenario) {
    int id=scenario%22,variant=scenario/22,mode=variant==3?1:variant==4?0:2,first=1,index=0;
    if(scenario==110) {
        for(int r=0;r<t->rows;r+=16)for(int c=0;c<t->cols;c+=16)mixed_leaves(t,out,r,c,64,0,&index,&first);
        return;
    }
    for(int r=0;r<t->rows;r+=bh[id]/4)for(int c=0;c<t->cols;c+=bw[id]/4,index++) {
        int inter=variant==1 || (variant==2&&index%3!=0),skip=index%7==3,lossless=variant==4;
        one_tx_leaf(t,out,r,c,id,mode,lossless,inter,skip,&first);
    }
}
static void make_tx(FILE *out,int scenario,int updates,int *first) {
    int id=scenario%22,extent=bw[id]>=64||bh[id]>=64?32:16;if(scenario/22==2)extent=18;if(scenario==110)extent=32;
    transform_codec enc,dec;tx_init(&enc,0,updates,extent,extent);enc.index=scenario*97;od_ec_enc_init(&enc.entropy.entropy.enc,1024);run_tx(&enc,NULL,scenario);
    uint32_t n;unsigned char *bytes=od_ec_enc_done(&enc.entropy.entropy.enc,&n);assert(bytes&&n>0&&!enc.entropy.entropy.enc.error);
    tx_init(&dec,1,updates,extent,extent);od_ec_dec_init(&dec.entropy.entropy.dec,bytes,n);
    fprintf(out,"%s{\"scenario\":%d,\"mode\":%d,\"extent\":%d,\"updates\":%s,\"hex\":",*first?"":",",scenario,scenario/22==3?1:scenario/22==4?0:2,extent,updates?"true":"false");*first=0;hex(out,bytes,n);
    fprintf(out,",\"states\":[");run_tx(&dec,out,scenario);fprintf(out,"],\"depthContexts\":[");
    for(int i=0;i<12;i++)fprintf(out,"%s%d",i?",":"",dec.depth_contexts[i]);fprintf(out,"],\"splitContexts\":[");
    for(int i=0;i<21;i++)fprintf(out,"%s%d",i?",":"",dec.split_contexts[i]);fprintf(out,"]}");
    assert(memcmp(enc.depth,dec.depth,sizeof(enc.depth))==0&&memcmp(enc.split,dec.split,sizeof(enc.split))==0&&memcmp(enc.grid,dec.grid,sizeof(enc.grid))==0);
    od_ec_enc_clear(&enc.entropy.entropy.enc);
}
static void tx_prefix(char **argv) {
    palette_codec p;unsigned char *bytes;int q,modes[9];palette_result result;
    int pixels=palette_first_leaf(argv,&p,&bytes,&q,modes,&result);
    transform_codec t;tx_init(&t,1,1,atoi(argv[4]),atoi(argv[5]));t.entropy.entropy.dec=p.mode.entropy.dec;
    tx_sizes(&t,0,0,block_id(pixels,pixels),atoi(argv[12]),0,0,0);
    residual_layout layout;residuals(&t,0,0,modes[1],&layout);
    FILE *out=fopen(argv[9],"wb");assert(out);fprintf(out,"{\"pixels\":%d,\"preludeQ\":%d,",pixels,q);print_layout(out,&t,&layout);fprintf(out,"}\n");fclose(out);free(bytes);
}
#ifdef OFFICEIMO_AV1_COEFFICIENT_INCLUDE
#define main transform_component_main
#endif
int main(int argc,char **argv) {
    if(argc==13){tx_prefix(argv);return 0;}assert(argc==2);FILE *out=fopen(argv[1],"wb");assert(out);
    fprintf(out,"{\"producer\":\"AOM v3.13.1 entropy/defaults with original normative transform syntax and full-grid traversal harness\",\"nativeSelfCheck\":true,\"cases\":[");
    int first=1;for(int s=0;s<66;s++)for(int u=0;u<2;u++)make_tx(out,s,u,&first);
    const int other[]={66,69,78,81,88,91,100,103,110};
    for(unsigned i=0;i<sizeof(other)/sizeof(other[0]);i++)for(int u=0;u<2;u++)make_tx(out,other[i],u,&first);
    fprintf(out,"]}\n");fclose(out);return 0;
}

#ifdef OFFICEIMO_AV1_COEFFICIENT_INCLUDE
#undef main
#endif
