/* Test-only independent TIFF producer and sample round-trip oracle.
 * Compile: cc generate_libtiff.c $(pkg-config --cflags --libs libtiff-4 lcms2) -lm -o producer
 * Run: producer OUTPUT_DIRECTORY RGB_PROFILE CMYK_PROFILE GRAY_PROFILE
 * LibTIFF and LittleCMS are optional validation tools, not OfficeIMO dependencies.
 */
#include <tiffio.h>
#include <lcms2.h>
#include <stdint.h>
#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include <math.h>
#define W 19
#define H 13
static int page_index=0;
static uint16_t sample(int x,int y,int c,int base,int extra) {
    static const uint16_t alphas[]={0,1,129,257,16384,32768,65535};
    uint16_t a=alphas[(x+2*y+page_index)%7];
    if(c==base) return a;
    uint16_t v=(uint16_t)((x*4133+y*7907+c*12347+193+page_index*20011)%65536);
    return extra==1 ? (uint16_t)(((uint64_t)v*a+32767)/65535) : v;
}
static unsigned char q(double x) { return (unsigned char)floor(fmax(0,fmin(1,x))*255+0.5); }
static unsigned char *file_bytes(const char *name,uint32_t *n) {
    FILE*f=fopen(name,"rb"); if(!f)exit(2); fseek(f,0,SEEK_END);*n=(uint32_t)ftell(f); rewind(f);
    unsigned char*b=malloc(*n);if(fread(b,1,*n,f)!=*n)exit(2);fclose(f);return b;
}
static void directory(TIFF*t,int photo,int extra,int planar,int tiled,int compression,int predictor,
                      const unsigned char*profile,uint32_t profile_n) {
    int base=photo==2?3:photo==5?4:1, samples=base+(extra>=0);
    TIFFSetField(t,TIFFTAG_IMAGEWIDTH,W); TIFFSetField(t,TIFFTAG_IMAGELENGTH,H);
    TIFFSetField(t,TIFFTAG_BITSPERSAMPLE,16); TIFFSetField(t,TIFFTAG_SAMPLESPERPIXEL,samples);
    TIFFSetField(t,TIFFTAG_SAMPLEFORMAT,SAMPLEFORMAT_UINT); TIFFSetField(t,TIFFTAG_PHOTOMETRIC,photo);
    TIFFSetField(t,TIFFTAG_FILLORDER,FILLORDER_MSB2LSB);
    TIFFSetField(t,TIFFTAG_PLANARCONFIG,planar); TIFFSetField(t,TIFFTAG_COMPRESSION,compression);
    if(predictor==2)TIFFSetField(t,TIFFTAG_PREDICTOR,2);
    if(photo==5)TIFFSetField(t,TIFFTAG_INKSET,INKSET_CMYK);
    if(extra>=0){uint16_t e=(uint16_t)extra;TIFFSetField(t,TIFFTAG_EXTRASAMPLES,1,&e);}
    if(profile)TIFFSetField(t,TIFFTAG_ICCPROFILE,profile_n,profile);
    if(tiled){TIFFSetField(t,TIFFTAG_TILEWIDTH,16);TIFFSetField(t,TIFFTAG_TILELENGTH,16);}
    else TIFFSetField(t,TIFFTAG_ROWSPERSTRIP,4);
    int planes=planar==2?samples:1, channels=planar==2?1:samples;
    int tw=tiled?16:W,th=tiled?16:4;
    for(int plane=0;plane<planes;plane++) for(int y=0;y<H;y+=th) for(int x=0;x<W;x+=tw) {
        int rows=tiled?th:(H-y<th?H-y:th);
        size_t n=(size_t)tw*rows*channels;
        uint16_t*b=calloc(n,sizeof(uint16_t));
        for(int dy=0;dy<rows;dy++)for(int dx=0;dx<tw;dx++)for(int c=0;c<channels;c++)
            if(x+dx<W&&y+dy<H)b[(dy*tw+dx)*channels+c]=sample(x+dx,y+dy,planar==2?plane:c,base,extra);
        tmsize_t result=tiled?TIFFWriteEncodedTile(t,TIFFComputeTile(t,x,y,0,plane),b,n*2):
                              TIFFWriteEncodedStrip(t,TIFFComputeStrip(t,y,plane),b,n*2);
        if(result<0)exit(3);free(b);
    }
}
static void verify_samples(const char*path,int photo,int extra,int planar,int tiled,int pages) {
    TIFF*t=TIFFOpen(path,"r");if(!t)exit(4);
    int base=photo==2?3:photo==5?4:1,samples=base+(extra>=0);
    for(int page=0;page<pages;page++) {
        page_index=page;
        int planes=planar==2?samples:1,channels=planar==2?1:samples;
        int tw=tiled?16:W,th=tiled?16:4;
        for(int plane=0;plane<planes;plane++)for(int y=0;y<H;y+=th)for(int x=0;x<W;x+=tw){
            int rows=tiled?th:(H-y<th?H-y:th);size_t n=(size_t)tw*rows*channels;
            uint16_t*b=calloc(n,2);
            tmsize_t got=tiled?TIFFReadEncodedTile(t,TIFFComputeTile(t,x,y,0,plane),b,n*2):
                               TIFFReadEncodedStrip(t,TIFFComputeStrip(t,y,plane),b,n*2);
            if(got!=(tmsize_t)n*2)exit(5);
            for(int dy=0;dy<rows&&y+dy<H;dy++)for(int dx=0;dx<tw&&x+dx<W;dx++)for(int c=0;c<channels;c++)
                if(b[(dy*tw+dx)*channels+c]!=sample(x+dx,y+dy,planar==2?plane:c,base,extra))exit(6);
            free(b);
        }
        if(page+1<pages&&!TIFFReadDirectory(t))exit(7);
    }TIFFClose(t);
}
static void produce(const char*out,FILE*manifest,const char*name,int big,int photo,int extra,int planar,
                    int tiled,int compression,int predictor,const char*profile_file,int pages) {
    char path[2048];snprintf(path,sizeof(path),"%s/%s.tif",out,name);
    uint32_t pn=0;unsigned char*profile=profile_file?file_bytes(profile_file,&pn):NULL;
    TIFF*t=TIFFOpen(path,big?"wb":"wl");if(!t)exit(2);
    for(int p=0;p<pages;p++){page_index=p;directory(t,photo,extra,planar,tiled,compression,predictor,profile,pn);if(!TIFFWriteDirectory(t))exit(3);}TIFFClose(t);
    verify_samples(path,photo,extra,planar,tiled,pages);
    cmsHTRANSFORM transform=NULL;cmsHPROFILE input=NULL,target=NULL;
    if(profile){input=cmsOpenProfileFromMem(profile,pn);target=cmsCreate_sRGBProfile();cmsSetProfileVersion(target,2.1);
        uint32_t format=photo==5?TYPE_CMYK_DBL:photo==2?TYPE_RGB_DBL:TYPE_GRAY_DBL;
        transform=cmsCreateTransform(input,format,target,TYPE_RGB_DBL,INTENT_RELATIVE_COLORIMETRIC,cmsFLAGS_NOOPTIMIZE|cmsFLAGS_NOCACHE);
        if(!transform)exit(8);}
    for(int p=0;p<pages;p++) {
    page_index=p;
    if(p==0)snprintf(path,sizeof(path),"%s/%s.rgba",out,name);
    else snprintf(path,sizeof(path),"%s/%s.page%d.rgba",out,name,p);
    FILE*f=fopen(path,"wb");if(!f)exit(2);
    int base=photo==2?3:photo==5?4:1;
    for(int y=0;y<H;y++)for(int x=0;x<W;x++){
        double v[4]={0},rgb[3];uint16_t a=extra>0?sample(x,y,base,base,extra):65535;
        for(int c=0;c<base;c++)v[c]=extra==1?(a?fmin(1,sample(x,y,c,base,extra)/(double)a):0):sample(x,y,c,base,extra)/65535.0;
        if(photo==0)v[0]=1-v[0];
        if(transform){if(photo==5)for(int c=0;c<4;c++)v[c]*=100;cmsDoTransform(transform,v,rgb,1);}
        else if(photo==2){for(int c=0;c<3;c++)rgb[c]=v[c];}
        else if(photo==5){for(int c=0;c<3;c++)rgb[c]=(255-fmin(255,q(v[c])+q(v[3])))/255.0;}
        else rgb[0]=rgb[1]=rgb[2]=v[0];
        unsigned char rgba[]={q(rgb[0]),q(rgb[1]),q(rgb[2]),q(a/65535.0)};fwrite(rgba,1,4,f);
    }fclose(f);
    }
    if(transform){cmsDeleteTransform(transform);cmsCloseProfile(input);cmsCloseProfile(target);}free(profile);
    fprintf(manifest,"%s.tif,%d,%d,%d,%d,%d,%d\n",name,W,H,photo,extra,profile_file!=NULL,pages);
}
int main(int argc,char**argv){
    if(argc!=5)return 1;char path[2048],name[200];snprintf(path,sizeof(path),"%s/manifest.csv",argv[1]);FILE*m=fopen(path,"w");if(!m)return 2;
    fprintf(m,"file,width,height,photometric,extra,icc,pages\n");
    const int comps[]={COMPRESSION_NONE,COMPRESSION_LZW,COMPRESSION_PACKBITS,COMPRESSION_ADOBE_DEFLATE};
    const char*cn[]={"none","lzw","packbits","deflate"};int index=0;
    for(int big=0;big<2;big++)for(int planar=1;planar<=2;planar++)for(int tiled=0;tiled<2;tiled++)for(int c=0;c<4;c++){
        int extra=(index++%4)-1;
        snprintf(name,sizeof(name),"rgb-extra%d-%s-%s-%s-%s",extra,cn[c],tiled?"tile":"strip",planar==1?"chunky":"planar",big?"be":"le");
        produce(argv[1],m,name,big,2,extra,planar,tiled,comps[c],c==1||c==3?2:1,NULL,1);
    }
    for(int big=0;big<2;big++){
        snprintf(name,sizeof(name),"gray-white-%s",big?"be":"le");produce(argv[1],m,name,big,0,-1,2,1,8,2,NULL,1);
        snprintf(name,sizeof(name),"gray-alpha-%s",big?"be":"le");produce(argv[1],m,name,big,1,1,1,0,5,2,NULL,1);
        snprintf(name,sizeof(name),"gray-icc-alpha-%s",big?"be":"le");produce(argv[1],m,name,big,1,1,2,0,8,2,argv[4],1);
        snprintf(name,sizeof(name),"rgb-icc-alpha-%s",big?"be":"le");produce(argv[1],m,name,big,2,1,2,1,8,2,argv[2],1);
        snprintf(name,sizeof(name),"cmyk-icc-%s",big?"be":"le");produce(argv[1],m,name,big,5,-1,1,0,5,2,argv[3],1);
        snprintf(name,sizeof(name),"cmyk-icc-alpha-%s",big?"be":"le");produce(argv[1],m,name,big,5,1,2,1,8,2,argv[3],1);
        snprintf(name,sizeof(name),"cmyk-device-%s",big?"be":"le");produce(argv[1],m,name,big,5,2,2,0,32773,1,NULL,1);
    }
    produce(argv[1],m,"rgb-multipage-be",1,2,2,2,0,8,2,NULL,2);
    fclose(m);printf("Verified 47 independently encoded TIFFs against their original 16-bit samples.\n");return 0;
}
