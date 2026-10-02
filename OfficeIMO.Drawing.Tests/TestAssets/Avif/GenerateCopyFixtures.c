/* Original reduced-still copy-motion grammar. Native entropy/defaults are opt-in reference code only. */
#define main entropy_component_main
#include "GenerateEntropyFixtures.c"
#undef main
#include "copy-defaults.h"

typedef struct { int row,col; } vector;
typedef struct { int width,height,done,copy;vector motion; } neighbor;
typedef struct {
    od_ec_enc enc;od_ec_dec dec;nmv_context probabilities;
    neighbor *grid;int rows,cols,row_start,col_start,sb,mono,updates,decode;
    vector stack[8];int weights[8],count;
} copy_context;
static int symbol(copy_context *c,aom_cdf_prob *cdf,int n,int expected) {
    int actual=expected;
    if(c->decode) {actual=od_ec_decode_cdf_q15(&c->dec,cdf,n);assert(actual==expected);}
    else od_ec_encode_cdf_q15(&c->enc,expected,cdf,n);
    if(c->updates)update_cdf(cdf,actual,n);return actual;
}
static int component(copy_context *c,int axis,int value) {
    int sign=value<0,mag=abs(value),kind=0;
    assert(mag>=8&&mag<=32768&&(mag&7)==0);
    while(kind<10 && mag>(16<<kind))kind++;
    nmv_component *p=&c->probabilities.components[axis];
    symbol(c,p->sign_cdf,2,sign);symbol(c,p->classes_cdf,11,kind);
    int result;
    if(kind==0)result=(symbol(c,p->class0_cdf,2,mag/8-1)+1)*8;
    else {
        int code=(mag-(8<<kind))/8-1,bits=0;
        for(int i=0;i<kind;i++)bits|=symbol(c,p->bits_cdf[i],2,(code>>i)&1)<<i;
        result=(8<<kind)+(bits+1)*8;
    }
    return sign?-result:result;
}
static vector difference(copy_context *c,vector delta) {
    int joint=(delta.row?2:0)|(delta.col?1:0);
    joint=symbol(c,c->probabilities.joints_cdf,4,joint);
    vector result={0,0};
    if(joint&2)result.row=component(c,0,delta.row);
    if(joint&1)result.col=component(c,1,delta.col);
    return result;
}
static int inside(copy_context *c,int r,int x) {return r>=c->row_start&&r<c->rows&&x>=c->col_start&&x<c->cols;}
static neighbor *cell(copy_context *c,int r,int x) {assert(inside(c,r,x));return &c->grid[r*c->cols+x];}
static void add(copy_context *c,neighbor *n,int weight) {
    if(!n->done||!n->copy)return;
    for(int i=0;i<c->count;i++)if(c->stack[i].row==n->motion.row&&c->stack[i].col==n->motion.col) {c->weights[i]+=weight;return;}
    if(c->count<8) {c->stack[c->count]=n->motion;c->weights[c->count++]=weight;}
}
static void scan(copy_context *c,int r,int x,int w,int h,int delta,int horizontal) {
    int bw=w/4,bh=h/4,dr=horizontal?delta:0,dc=horizontal?0:delta;
    if(abs(delta)>1) {if(horizontal){dr+=r&1;dc=1-(x&1);}else{dr=1-(r&1);dc+=x&1;}}
    int extent=horizontal?bw:bh,end=extent;
    if(end>(horizontal?c->cols-x:c->rows-r))end=horizontal?c->cols-x:c->rows-r;
    if(end>16)end=16;
    for(int i=0;i<end;) {
        int rr=r+dr+(horizontal?0:i),xx=x+dc+(horizontal?i:0);if(!inside(c,rr,xx))break;
        neighbor *n=cell(c,rr,xx);int length=horizontal?n->width/4:n->height/4;
        if(length<1)length=1;if(length>extent)length=extent;
        if(abs(horizontal?dr:dc)>1&&length<2)length=2;if(extent>=16&&length<4)length=4;
        add(c,n,2*length);i+=length;
    }
}
static void point(copy_context *c,int r,int x) {if(inside(c,r,x))add(c,cell(c,r,x),4);}
static void sort(copy_context *c,int begin,int end) {
    while(end>begin) {
        int last=begin;
        for(int i=begin+1;i<end;i++)if(c->weights[i-1]<c->weights[i]) {
            int weight=c->weights[i-1];c->weights[i-1]=c->weights[i];c->weights[i]=weight;
            vector v=c->stack[i-1];c->stack[i-1]=c->stack[i];c->stack[i]=v;last=i;
        }
        end=last;
    }
}
static vector predictor(copy_context *c,int r,int x,int w,int h) {
    c->count=0;scan(c,r,x,w,h,-1,1);scan(c,r,x,w,h,-1,0);
    if(w<=64&&h<=64)point(c,r-1,x+w/4);
    int near=c->count;for(int i=0;i<near;i++)c->weights[i]+=640;
    point(c,r-1,x-1);scan(c,r,x,w,h,-3,1);scan(c,r,x,w,h,-3,0);
    if(h>4)scan(c,r,x,w,h,-5,1);if(w>4)scan(c,r,x,w,h,-5,0);
    sort(c,0,near);sort(c,near,c->count);
    for(int i=c->count;i<2;i++)c->stack[i]=(vector){0,0};
    for(int i=0;i<c->count;i++) {
        c->stack[i].row=clamp(c->stack[i].row,-r*32-128-h*8,(c->rows-h/4-r)*32+128+h*8);
        c->stack[i].col=clamp(c->stack[i].col,-x*32-128-w*8,(c->cols-w/4-x)*32+128+w*8);
    }
    vector p=c->stack[0];if(!p.row&&!p.col)p=c->stack[1];
    if(!p.row&&!p.col)p=r-c->sb/4<c->row_start?(vector){0,-(c->sb+256)*8}:(vector){-c->sb*8,0};
    return p;
}
static int valid(copy_context *c,int r,int x,int w,int h,vector mv,int chroma) {
    if(abs(mv.row)>=16384||abs(mv.col)>=16384||(mv.row&7)||(mv.col&7))return 0;
    int top=r*4+mv.row/8,left=x*4+mv.col/8,bottom=top+h,right=left+w;
    if(chroma) {if(w<8)left-=4;if(h<8)top-=4;}
    if(top<c->row_start*4||left<c->col_start*4||bottom>c->rows*4||right>c->cols*4)return 0;
    int active_row=r*4/c->sb,active_col=x/16,source_row=(bottom-1)/c->sb,source_col=(right-1)/64;
    int per_row=(c->cols-c->col_start+15)/16;
    if(source_row*per_row+source_col>=active_row*per_row+active_col-4)return 0;
    return source_row<=active_row&&source_col<active_col-4+(5+(c->sb==128))*(active_row-source_row);
}
static void publish(copy_context *c,int r,int x,int w,int h,int copy,vector mv) {
    for(int y=r;y<r+h/4&&y<c->rows;y++)for(int xx=x;xx<x+w/4&&xx<c->cols;xx++)
        *cell(c,y,xx)=(neighbor){w,h,1,copy,mv};
}
static vector chosen(copy_context *c,int r,int x,int w,int h,int chroma,int index,int *copy) {
    vector mv={0,0};*copy=0;
    if(index%7==0)return mv;
    int sx[6]={c->col_start*4+8,c->col_start*4+24,c->col_start*4+64,x*4-320,x*4-384,c->col_start*4};
    int sy[6]={c->row_start*4+8,c->row_start*4,c->row_start*4+16,r*4,r*4-128,r*4-64};
    for(int i=0;i<36;i++) {
        int choice=(i+index*5)%36;
        mv=(vector){(sy[choice/6]-r*4)*8,(sx[choice%6]-x*4)*8};
        if(valid(c,r,x,w,h,mv,chroma)) {*copy=1;return mv;}
    }
    return (vector){0,0};
}
static void leaf(copy_context *c,FILE *out,int r,int x,int w,int h,int *index,int *first) {
    int chroma=!c->mono&&!(h==4&&!(r&1))&&!(w==4&&!(x&1)),copy;
    vector mv=chosen(c,r,x,w,h,chroma,(*index)++,&copy),p={0,0};
    if(copy) {
        p=predictor(c,r,x,w,h);vector delta={mv.row-p.row,mv.col-p.col};
        delta=difference(c,delta);assert(mv.row==p.row+delta.row&&mv.col==p.col+delta.col);
        assert(valid(c,r,x,w,h,mv,chroma));
    }
    publish(c,r,x,w,h,copy,mv);
    if(out)fprintf(out,"%s[%d,%d,%d,%d,%d,%d,%d,%d,%d,%d]",*first?"":",",r,x,w,h,copy,chroma,p.row,p.col,mv.row,mv.col);
    *first=0;
}
static void square(copy_context *c,FILE *out,int r,int x,int pixels,int w,int h,int *index,int *first) {
    if(r>=c->rows||x>=c->cols)return;
    int target=w>h?w:h;
    if(pixels>target) {
        int half=pixels/2;
        for(int i=0;i<4;i++)square(c,out,r+i/2*(half/4),x+i%2*(half/4),half,w,h,index,first);
    } else {
        for(int rr=r;rr<r+pixels/4&&rr<c->rows;rr+=h/4)for(int xx=x;xx<x+pixels/4&&xx<c->cols;xx+=w/4)
            leaf(c,out,rr,xx,w,h,index,first);
    }
}
static void run(copy_context *c,FILE *out,int width,int height) {
    int index=0,first=1;
    for(int sr=c->row_start;sr<c->rows;sr+=c->sb/4)for(int sc=c->col_start;sc<c->cols;sc+=c->sb/4) {
        const int mixed[6][2]={{8,8},{8,16},{16,8},{4,4},{32,16},{16,32}};
        int choice=((sr-c->row_start)/(c->sb/4)*3+(sc-c->col_start)/(c->sb/4))%6;
        int w=width?width:mixed[choice][0],h=height?height:mixed[choice][1];
        square(c,out,sr,sc,c->sb,w,h,&index,&first);
    }
}
static copy_context *create(int sb,int variant,int updates,int decode) {
    copy_context *c=calloc(1,sizeof(*c));assert(c);c->sb=sb;c->mono=variant==2;c->updates=updates;c->decode=decode;
    c->row_start=variant==1?sb/4:0;c->col_start=c->row_start;
    c->rows=c->row_start+32+(variant==2?2:0);c->cols=c->col_start+96+(variant==2?2:0);
    c->grid=calloc((size_t)c->rows*c->cols,sizeof(neighbor));assert(c->grid);c->probabilities=default_nmv_context;return c;
}
static void copy_case(FILE *out,int w,int h,int sb,int variant,int updates,int *first) {
    copy_context *enc=create(sb,variant,updates,0),*dec=create(sb,variant,updates,1);
    od_ec_enc_init(&enc->enc,4096);run(enc,NULL,w,h);
    uint32_t size;unsigned char *bytes=od_ec_enc_done(&enc->enc,&size);assert(bytes&&!enc->enc.error);
    od_ec_dec_init(&dec->dec,bytes,size);
    fprintf(out,"%s{\"sb\":%d,\"variant\":%d,\"updates\":%s,\"rows\":%d,\"cols\":%d,\"start\":%d,\"hex\":\"",*first?"":",",sb,variant,updates?"true":"false",enc->rows,enc->cols,enc->row_start);*first=0;
    for(uint32_t i=0;i<size;i++)fprintf(out,"%02x",bytes[i]);fprintf(out,"\",\"states\":[");run(dec,out,w,h);fprintf(out,"]}");
    assert(memcmp(&enc->probabilities,&dec->probabilities,sizeof(nmv_context))==0);
    assert(memcmp(enc->grid,dec->grid,(size_t)enc->rows*enc->cols*sizeof(neighbor))==0);
    od_ec_enc_clear(&enc->enc);free(enc->grid);free(dec->grid);free(enc);free(dec);
}
static void rejection(FILE *out,int axis,int kind,int sign,int maximum,int *first) {
    copy_context *c=create(64,0,1,0);od_ec_enc_init(&c->enc,128);
    // At tile origin there is no conforming copy source, irrespective of decoded difference.
    int value=(maximum?(16<<kind):(kind?((8<<kind)+8):8))*(sign?-1:1);vector delta={axis?0:value,axis?value:0};
    difference(c,delta);uint32_t size;unsigned char *bytes=od_ec_enc_done(&c->enc,&size);assert(bytes&&!c->enc.error);
    fprintf(out,"%s{\"axis\":%d,\"class\":%d,\"sign\":%d,\"maximum\":%s,\"hex\":\"",*first?"":",",axis,kind,sign,maximum?"true":"false");*first=0;
    for(uint32_t i=0;i<size;i++)fprintf(out,"%02x",bytes[i]);fprintf(out,"\"}");od_ec_enc_clear(&c->enc);free(c->grid);free(c);
}
static int prefix_square(copy_context *c,FILE *out,int r,int x,int pixels,int w,int h,int target_r,int target_x,int *first) {
    if(r>=c->rows||x>=c->cols)return 0;
    int target=w>h?w:h;
    if(pixels>target) {
        int half=pixels/2;
        for(int i=0;i<4;i++)if(prefix_square(c,out,r+i/2*(half/4),x+i%2*(half/4),half,w,h,target_r,target_x,first))return 1;
    } else {
        for(int rr=r;rr<r+pixels/4&&rr<c->rows;rr+=h/4)for(int xx=x;xx<x+pixels/4&&xx<c->cols;xx+=w/4) {
            if(rr==target_r&&xx==target_x)return 1;
            publish(c,rr,xx,w,h,0,(vector){0,0});fprintf(out,"%s[%d,%d,%d,%d]",*first?"":",",rr,xx,w,h);*first=0;
        }
    }
    return 0;
}
static void source_probe(FILE *out,const char *name,int cols,int r,int x,int w,int h,int mono,vector mv,int expected,int *first) {
    copy_context *c=create(64,0,1,0);free(c->grid);c->cols=cols;c->mono=mono;c->grid=calloc((size_t)c->rows*c->cols,sizeof(neighbor));assert(c->grid);
    fprintf(out,"%s{\"name\":\"%s\",\"rows\":%d,\"cols\":%d,\"mono\":%s,\"valid\":%s,\"block\":[%d,%d,%d,%d],\"motion\":[%d,%d],\"prefix\":[",*first?"":",",name,c->rows,cols,mono?"true":"false",expected?"true":"false",r,x,w,h,mv.row,mv.col);*first=0;
    int done=0,prefix_first=1;
    for(int sr=0;sr<c->rows&&!done;sr+=16)for(int sc=0;sc<cols&&!done;sc+=16)
        done=prefix_square(c,out,sr,sc,64,w,h,r,x,&prefix_first);
    assert(done);vector p=predictor(c,r,x,w,h),delta={mv.row-p.row,mv.col-p.col};
    od_ec_enc_init(&c->enc,128);difference(c,delta);uint32_t size;unsigned char *bytes=od_ec_enc_done(&c->enc,&size);assert(bytes&&!c->enc.error);
    int chroma=!mono&&!(h==4&&!(r&1))&&!(w==4&&!(x&1));assert(valid(c,r,x,w,h,mv,chroma)==expected);
    c->decode=1;c->probabilities=default_nmv_context;od_ec_dec_init(&c->dec,bytes,size);vector actual=difference(c,delta);
    assert(actual.row==delta.row&&actual.col==delta.col);
    fprintf(out,"],\"hex\":\"");for(uint32_t i=0;i<size;i++)fprintf(out,"%02x",bytes[i]);fprintf(out,"\"}");
    od_ec_enc_clear(&c->enc);free(c->grid);free(c);
}
int main(int argc,char **argv) {
    assert(argc==2);FILE *out=fopen(argv[1],"wb");assert(out);
    int sizes[22][2]={{4,4},{4,8},{8,4},{8,8},{8,16},{16,8},{16,16},{16,32},{32,16},{32,32},{32,64},{64,32},{64,64},{64,128},{128,64},{128,128},{4,16},{16,4},{8,32},{32,8},{16,64},{64,16}};
    fprintf(out,"{\"producer\":\"AOM v3.13.1 entropy/defaults and original still-copy grammar\",\"nativeSelfCheck\":true,\"cases\":[");int first=1;
    for(int size=0;size<22;size++)for(int sb=64;sb<=128;sb*=2)if(sizes[size][0]<=sb&&sizes[size][1]<=sb)
        for(int variant=0;variant<3;variant++)for(int updates=0;updates<2;updates++)copy_case(out,sizes[size][0],sizes[size][1],sb,variant,updates,&first);
    for(int sb=64;sb<=128;sb*=2)for(int variant=0;variant<3;variant++)for(int updates=0;updates<2;updates++)copy_case(out,0,0,sb,variant,updates,&first);
    fprintf(out,"],\"rejections\":[");first=1;
    for(int axis=0;axis<2;axis++)for(int kind=0;kind<11;kind++)for(int sign=0;sign<2;sign++)for(int maximum=0;maximum<2;maximum++)rejection(out,axis,kind,sign,maximum,&first);
    fprintf(out,"],\"sourceProbes\":[");first=1;
    source_probe(out,"wavefront-delay",96,16,0,8,8,0,(vector){-512,512},0,&first);
    source_probe(out,"linear-delay",64,16,16,8,8,0,(vector){-512,0},0,&first);
    source_probe(out,"tile-top",96,16,0,8,8,0,(vector){-520,0},0,&first);
    source_probe(out,"tile-left",96,16,0,8,8,0,(vector){-512,-8},0,&first);
    source_probe(out,"shared-chroma-footprint",96,17,1,4,4,0,(vector){-544,-32},0,&first);
    source_probe(out,"monochrome-footprint-control",96,17,1,4,4,1,(vector){-544,-32},1,&first);
    source_probe(out,"horizontal-delay-control",96,0,80,8,8,0,(vector){0,-2560},1,&first);
    source_probe(out,"previous-row-control",96,16,0,8,8,0,(vector){-512,0},1,&first);
    for(int sign=0;sign<2;sign++)for(int bit=0;bit<2;bit++) {
        int diff=(8+bit*8)*(sign?-1:1);
        source_probe(out,"class0-row",96,18,0,8,8,0,(vector){-512+diff,0},1,&first);
        source_probe(out,"class0-col",96,0,82,8,8,0,(vector){0,-2560+diff},1,&first);
    }
    source_probe(out,"class9-exact-control",1024,0,320,8,8,0,(vector){0,-10240},1,&first);
    source_probe(out,"class10-exact-control",1024,0,448,8,8,0,(vector){0,-14336},1,&first);
    fprintf(out,"]}\n");fclose(out);return 0;
}
