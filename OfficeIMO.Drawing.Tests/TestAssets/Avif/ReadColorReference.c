/* Opt-in independent libavif color oracle. No conversion arithmetic is implemented here. */
#include "avif/avif.h"
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

static void hex(const uint8_t *p, int pitch, int width, int height, int depth) {
    putchar('"');
    for(int y=0;y<height;y++)for(int x=0;x<width;x++) {
        if(depth==8)printf("%02x",p[y*pitch+x]);
        else {const uint16_t *row=(const uint16_t *)(p+y*pitch);printf("%02x%02x",row[x]&255,row[x]>>8);}
    }
    putchar('"');
}
int main(int argc,char **argv) {
    const char *version=avifVersion();
    if(strcmp(version,"1.3.0") && strcmp(version,"1.4.2"))return 1;
    int depth=argc==2?atoi(argv[1]):8;if((depth!=8 && depth!=10) || argc>2)return 1;
    int matrices[]={1,2,4,5,6,7,9},widths[]={1,3,17,49},heights[]={1,5,9,33},first=1;
    uint8_t values[]={0,1,15,16,17,127,128,235,236,254,255};
    printf("{\"version\":\"%s\",\"cases\":[",version);
    for(int mono=0;mono<(depth==10?2:1);mono++)for(int m=0;m<7;m++)for(int full=0;full<2;full++)
    for(int a=0;a<3;a++)for(int s=0;s<4;s++)for(int phase=0;phase<(depth==10?4:1);phase++) {
        if(mono && matrices[m]!=6)continue;
        int w=widths[s],h=heights[s];
        avifImage *im=avifImageCreate(w,h,depth,mono?AVIF_PIXEL_FORMAT_YUV400:AVIF_PIXEL_FORMAT_YUV420);if(!im)return 2;
        im->matrixCoefficients=matrices[m];im->yuvRange=full?AVIF_RANGE_FULL:AVIF_RANGE_LIMITED;
        if(avifImageAllocatePlanes(im,a?AVIF_PLANES_ALL:AVIF_PLANES_YUV)!=AVIF_RESULT_OK)return 3;
        for(int p=0;p<(mono?1:3);p++)for(int y=0;y<(p?(h+1)/2:h);y++)for(int x=0;x<(p?(w+1)/2:w);x++) {
            int value=values[(x*3+y*7+p*5+s)%11];
            if(depth==8)im->yuvPlanes[p][y*im->yuvRowBytes[p]+x]=(uint8_t)value;
            else ((uint16_t *)(im->yuvPlanes[p]+y*im->yuvRowBytes[p]))[x]=(uint16_t)((value<<2)+((x+3*y+p+phase)&3));
        }
        if(a)for(int y=0;y<h;y++)for(int x=0;x<w;x++) {
            if(depth==8)im->alphaPlane[y*im->alphaRowBytes+x]=a==2?0:(uint8_t)((x*47+y*31)%256);
            else ((uint16_t *)(im->alphaPlane+y*im->alphaRowBytes))[x]=(uint16_t)(a==2?0:((x*181+y*127+phase)&1023));
        }
        avifRGBImage rgb;avifRGBImageSetDefaults(&rgb,im);rgb.format=AVIF_RGB_FORMAT_RGBA;rgb.depth=8;
        rgb.avoidLibYUV=AVIF_TRUE;rgb.chromaUpsampling=AVIF_CHROMA_UPSAMPLING_BILINEAR;
        if(avifRGBImageAllocatePixels(&rgb)!=AVIF_RESULT_OK || avifImageYUVToRGB(im,&rgb)!=AVIF_RESULT_OK)return 4;
        if(!first)putchar(',');
        first=0;
        printf("{\"width\":%d,\"height\":%d,\"matrix\":%d,\"fullRange\":%s,\"alphaMode\":%d,",w,h,matrices[m],full?"true":"false",a);
        if(depth==10)printf("\"bitDepth\":10,\"monochrome\":%s,\"lowBitPhase\":%d,",mono?"true":"false",phase);
        printf("\"planes\":[");
        for(int p=0;p<(mono?1:3);p++){if(p)putchar(',');hex(im->yuvPlanes[p],im->yuvRowBytes[p],p?(w+1)/2:w,p?(h+1)/2:h,depth);}
        printf("],\"alpha\":");if(a)hex(im->alphaPlane,im->alphaRowBytes,w,h,depth);else printf("null");
        printf(",\"rgba\":");hex(rgb.pixels,rgb.rowBytes,w*4,h,8);putchar('}');
        avifRGBImageFreePixels(&rgb);avifImageDestroy(im);
    }
    printf("]}\n");return 0;
}
