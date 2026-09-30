/* Original AV1 palette/filter syntax harness; native AOM supplies entropy, CDFs and color-index context. */
#define OFFICEIMO_AV1_COMPONENT_INCLUDE
#include "GenerateModeFixtures.c"
#include "palette-defaults.h"

typedef struct {
    mode_codec mode;
    aom_cdf_prob has_y[7][3][3],has_uv[2][3],size_y[7][8],size_uv[7][8];
    aom_cdf_prob map_y[7][5][9],map_uv[7][5][9],filter[22][3],filter_mode[6];
    unsigned char sizes[2][64][64],colors[2][64][64][8];
    int contexts[2][35],cache_bits,cache_hits,filter_hits;
} palette_codec;
typedef struct { int ny,nu,f,w,h;unsigned char y[8],u[8],v[8],map_y[4096],map_uv[1024]; } palette_result;
static void palette_init(palette_codec *p,int decode,int updates) {
    memset(p,0,sizeof(*p));mode_init(&p->mode,decode,updates);
    memcpy(p->has_y,default_palette_y_mode_cdf,sizeof(p->has_y));memcpy(p->has_uv,default_palette_uv_mode_cdf,sizeof(p->has_uv));
    memcpy(p->size_y,default_palette_y_size_cdf,sizeof(p->size_y));memcpy(p->size_uv,default_palette_uv_size_cdf,sizeof(p->size_uv));
    memcpy(p->map_y,default_palette_y_color_index_cdf,sizeof(p->map_y));memcpy(p->map_uv,default_palette_uv_color_index_cdf,sizeof(p->map_uv));
    memcpy(p->filter,default_filter_intra_cdfs,sizeof(p->filter));memcpy(p->filter_mode,default_filter_intra_mode_cdf,sizeof(p->filter_mode));
}
static int psym(palette_codec *p,aom_cdf_prob *cdf,int n,int expected) { return mode_symbol(&p->mode,cdf,n,expected); }
static int praw(palette_codec *p,int bits,int expected) {
    int actual=0;
    for(int b=bits-1;b>=0;b--) {
        int value=p->mode.entropy.decode?od_ec_decode_bool_q15(&p->mode.entropy.dec,16384):(expected>>b)&1;
        if(!p->mode.entropy.decode) od_ec_encode_bool_q15(&p->mode.entropy.enc,value,16384);
        actual=(actual<<1)|value;
    }
    if(expected>=0) assert(actual==expected);return actual;
}
static int log2n(int x) {int n=0;while((x>>=1))n++;return n;}
static int cmp_byte(const void *a,const void *b) {return *(const unsigned char*)a-*(const unsigned char*)b;}
static int cache(palette_codec *p,int r,int c,int plane,unsigned char *out) {
    int n=0;if(r&&r%16) {int k=p->sizes[plane][r-1][c];memcpy(out,p->colors[plane][r-1][c],k);n+=k;}
    if(c) {int k=p->sizes[plane][r][c-1];memcpy(out+n,p->colors[plane][r][c-1],k);n+=k;}
    qsort(out,n,1,cmp_byte);int unique=0;
    for(int i=0;i<n;i++)if(!unique||out[i]!=out[unique-1])out[unique++]=out[i];return unique;
}
static void palette_colors(palette_codec *p,int r,int c,int plane,int n,int index,unsigned char *colors) {
    unsigned char cached[16];int available=cache(p,r,c,plane,cached),i=0,known=index>=0;
    for(int j=0;j<available&&i<n;j++) {
        int use=praw(p,1,known?(index%6==0||(index+j)%3!=2):-1);p->cache_bits++;p->cache_hits+=use;
        if(use)colors[i++]=cached[j];
    }
    if(i<n) colors[i++]=(unsigned char)praw(p,8,known?(index%5==0?242:(index*17+plane*23)%32):-1);
    int bits=i<n?5+praw(p,2,known?index%4:-1):0;
    while(i<n) {
        int delta=praw(p,bits,known?(index+i*19)% (1<<bits):-1)+(plane==0);
        int value=colors[i-1]+delta;if(value>255)value=255;colors[i++]=(unsigned char)value;
        int range=256-value-(plane==0),limit=range<=1?0:log2n(range-1)+1;if(bits>limit)bits=limit;
    }
    qsort(colors,n,1,cmp_byte);
}
static void vcolors(palette_codec *p,int n,int index,unsigned char *colors) {
    int known=index>=0,delta=praw(p,1,known?index%2:-1);
    if(!delta)for(int i=0;i<n;i++)colors[i]=(unsigned char)praw(p,8,known?(index*13+i*31)%256:-1);
    else {
        int bits=4+praw(p,2,known?index%4:-1);colors[0]=(unsigned char)praw(p,8,known?index%3==0?250:4:-1);
        for(int i=1;i<n;i++) {
            int d=praw(p,bits,known?(index+i*7)%(1<<bits):-1);
            if(d&&praw(p,1,known?(index+i)%2:-1))d=-d;colors[i]=(unsigned char)((colors[i-1]+d+256)&255);
        }
    }
}
static int ns(palette_codec *p,int n,int expected) {
    int bits=log2n(n)+1,m=(1<<bits)-n;
    int value=praw(p,bits-1,expected<0?-1:expected<m?expected:(expected+m)/2);
    return value<m?value:2*value-m+praw(p,1,expected<0?-1:(expected+m)&1);
}
static int color_at(int index,int row,int col,int width,int count) {
    unsigned value=(unsigned)(index*4096+row*width+col+1);
    value^=value>>16;value*=0x7feb352dU;value^=value>>15;value*=0x846ca68bU;value^=value>>16;
    return (row/4+col/4+index)%3==0?index%count:(int)(value%(unsigned)count);
}
static void palette_map(palette_codec *p,int n,int w,int h,int ow,int oh,int plane,int index,unsigned char *map) {
    int known=index>=0;map[0]=(unsigned char)ns(p,n,known?index%n:-1);
    for(int d=1;d<ow+oh-1;d++)for(int c=d<ow?d:ow-1;c>=(d-oh+1>0?d-oh+1:0);c--) {
        int r=d-c,rank=-1;uint8_t order[8];
        int expected=known?color_at(index,r,c,w,n):-1;
        if(known)map[r*w+c]=(unsigned char)expected;
        int context=av1_get_palette_color_index_context(map,w,r,c,n,order,known?&rank:NULL);
        int symbol=psym(p,plane?p->map_uv[n-2][context]:p->map_y[n-2][context],n,rank);
        map[r*w+c]=order[symbol];p->contexts[plane][(n-2)*5+context]++;
    }
    for(int r=0;r<oh;r++)for(int c=ow;c<w;c++)map[r*w+c]=map[r*w+ow-1];
    for(int r=oh;r<h;r++)memcpy(map+r*w,map+(oh-1)*w,w);
}
static int block_id(int w,int h) {
    const int widths[22]={4,4,8,8,8,16,16,16,32,32,32,64,64,64,128,128,4,16,8,32,16,64};
    const int heights[22]={4,8,4,8,16,8,16,32,16,32,64,32,64,128,64,128,16,4,32,8,64,16};
    for(int i=0;i<22;i++)if(widths[i]==w&&heights[i]==h)return i;assert(0);return 0;
}
static void palette(palette_codec *p,int r,int c,int w,int h,int rows,int cols,int screen,int filter,int *modes,int index,palette_result *out) {
    memset(out,0,sizeof(*out));out->w=w;out->h=h;out->f=-1;int known=index>=0;
    if(!modes[0]) {
        if(screen&&w*h>=64&&w<=64&&h<=64) {
            int bsize=log2n(w/4)+log2n(h/4)-2;
            if(!modes[3]) {
                int ctx=(r&&p->sizes[0][r-1][c]>0)+(c&&p->sizes[0][r][c-1]>0);
                if(psym(p,p->has_y[bsize][ctx],2,known?index%4!=3:-1)) {
                    out->ny=psym(p,p->size_y[bsize],7,known?(index*3+index/7)%7:-1)+2;
                    palette_colors(p,r,c,0,out->ny,index,out->y);
                }
            }
            if(modes[1]&&!modes[4]&&psym(p,p->has_uv[out->ny>0],2,known?index%3!=2:-1)) {
                out->nu=psym(p,p->size_uv[bsize],7,known?(index*5+index/3)%7:-1)+2;
                palette_colors(p,r,c,1,out->nu,index,out->u);vcolors(p,out->nu,index,out->v);
            }
        }
        if(filter&&!modes[3]&&!out->ny&&w<=32&&h<=32&&psym(p,p->filter[block_id(w,h)],2,known?index%3!=1:-1)) {
            out->f=psym(p,p->filter_mode,5,known?index%5:-1);p->filter_hits++;
        }
    }
    int ow=w<(cols-c)*4?w:(cols-c)*4,oh=h<(rows-r)*4?h:(rows-r)*4;
    if(out->ny)palette_map(p,out->ny,w,h,ow,oh,0,index,out->map_y);
    if(out->nu) {int uw=w/2,uh=h/2,uow=ow/2,uoh=oh/2;if(uw<4){uw+=2;uow+=2;}if(uh<4){uh+=2;uoh+=2;}
        palette_map(p,out->nu,uw,uh,uow,uoh,1,index,out->map_uv);}
    for(int yy=r;yy<r+h/4&&yy<rows;yy++)for(int x=c;x<c+w/4&&x<cols;x++) {
        p->sizes[0][yy][x]=(unsigned char)out->ny;p->sizes[1][yy][x]=(unsigned char)out->nu;
        memcpy(p->colors[0][yy][x],out->y,out->ny);memcpy(p->colors[1][yy][x],out->u,out->nu);
    }
}
static void hex(FILE *out,const unsigned char *values,int n) {fprintf(out,"\"");for(int i=0;i<n;i++)fprintf(out,"%02x",values[i]);fprintf(out,"\"");}
static void print_result(FILE *out,palette_result *p) {
    fprintf(out,"\"filter\":%d,\"y\":",p->f);hex(out,p->y,p->ny);fprintf(out,",\"u\":");hex(out,p->u,p->nu);fprintf(out,",\"v\":");hex(out,p->v,p->nu);
    fprintf(out,",\"mapY\":");hex(out,p->map_y,p->ny?p->w*p->h:0);
    fprintf(out,",\"mapUv\":");hex(out,p->map_uv,p->nu?(p->w<8?4:p->w/2)*(p->h<8?4:p->h/2):0);
}
static void run_palette(palette_codec *p,FILE *out,int scenario) {
    const int sizes[20][2]={{8,8},{16,16},{32,32},{64,64},{4,16},{16,4},{8,32},{32,8},{16,64},{64,16},
        {4,4},{4,8},{8,4},{8,16},{16,8},{16,32},{32,16},{128,128},{16,16},{8,8}};
    int w=sizes[scenario%20][0],h=sizes[scenario%20][1];
    int extent=scenario%20==18?18:w==128?32:w==64||h==64?32:16;
    if(scenario%20==19)extent=32;
    int index=0,first=1;
    for(int r=0;r<extent;r+=h/4)for(int c=0;c<extent;c+=w/4,index++) {
        int mono=scenario==20,screen=scenario!=21,filter=scenario!=22;
        int mode[9]={scenario==23&&index%4==0,!mono&&!(h==4&&!(r&1))&&!(w==4&&!(c&1)),0,index%11==10?1:0,index%7==6?1:0,0,0,0,0};
        palette_result result;palette(p,r,c,w,h,extent,extent,screen,filter,mode,index,&result);
        if(out) {
            fprintf(out,"%s{\"block\":[%d,%d,%d,%d],\"modes\":[",first?"":",",r,c,w,h);first=0;
            for(int i=0;i<9;i++)fprintf(out,"%s%d",i?",":"",mode[i]);fprintf(out,"],");print_result(out,&result);fprintf(out,"}");
        }
    }
}
static void make_palette(FILE *out,int scenario,int updates,int *first) {
    palette_codec enc,dec;palette_init(&enc,0,updates);od_ec_enc_init(&enc.mode.entropy.enc,4096);run_palette(&enc,NULL,scenario);
    uint32_t n;unsigned char *bytes=od_ec_enc_done(&enc.mode.entropy.enc,&n);assert(bytes&&!enc.mode.entropy.enc.error&&n>0);
    palette_init(&dec,1,updates);od_ec_dec_init(&dec.mode.entropy.dec,bytes,n);
    fprintf(out,"%s{\"scenario\":%d,\"updates\":%s,\"hex\":",*first?"":",",scenario,updates?"true":"false");*first=0;hex(out,bytes,n);
    fprintf(out,",\"states\":[");run_palette(&dec,out,scenario);fprintf(out,"],\"mapContexts\":[");
    for(int i=0;i<70;i++)fprintf(out,"%s%d",i?",":"",dec.contexts[i/35][i%35]);
    fprintf(out,"],\"cacheBits\":%d,\"cacheHits\":%d,\"filterHits\":%d}",dec.cache_bits,dec.cache_hits,dec.filter_hits);
    assert(memcmp(enc.has_y,dec.has_y,sizeof(enc.has_y))==0&&memcmp(enc.map_y,dec.map_y,sizeof(enc.map_y))==0&&memcmp(enc.map_uv,dec.map_uv,sizeof(enc.map_uv))==0&&memcmp(enc.filter,dec.filter,sizeof(enc.filter))==0);
    od_ec_enc_clear(&enc.mode.entropy.enc);
}
static void palette_prefix(char **argv) {
    palette_codec p;palette_init(&p,1,1);unsigned char *bytes;int q,mode[9];
    int pixels=mode_first_leaf(argv,&p.mode,&bytes,&q,mode);palette_result result;
    palette(&p,0,0,pixels,pixels,64,64,atoi(argv[11]),1,mode,-1,&result);
    FILE *out=fopen(argv[9],"wb");assert(out);fprintf(out,"{\"pixels\":%d,\"preludeQ\":%d,",pixels,q);print_result(out,&result);fprintf(out,"}\n");fclose(out);free(bytes);
}
int main(int argc,char **argv) {
    if(argc==12){palette_prefix(argv);return 0;}assert(argc==2);
    FILE *out=fopen(argv[1],"wb");assert(out);fprintf(out,"{\"producer\":\"AOM v3.13.1 entropy, tables and native palette context; original remaining syntax harness\",\"nativeSelfCheck\":true,\"cases\":[");
    int first=1;for(int s=0;s<24;s++)for(int u=0;u<2;u++)make_palette(out,s,u,&first);
    fprintf(out,"]}\n");fclose(out);return 0;
}
