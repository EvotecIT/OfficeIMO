/* Opt-in independent libavif 1.3.0 color oracle. No conversion arithmetic is implemented here. */
#include "avif/avif.h"
#include <stdio.h>
#include <string.h>

static void hex(const uint8_t *p, int pitch, int width, int height) {
    putchar('"'); for(int y=0;y<height;y++)for(int x=0;x<width;x++)printf("%02x",p[y*pitch+x]); putchar('"');
}
int main(void) {
    if(strcmp(avifVersion(),"1.3.0"))return 1;
    int matrices[]={1,2,4,5,6,7,9},widths[]={1,3,17,49},heights[]={1,5,9,33},first=1;
    uint8_t values[]={0,1,15,16,17,127,128,235,236,254,255};
    printf("{\"version\":\"%s\",\"cases\":[",avifVersion());
    for(int m=0;m<7;m++)for(int full=0;full<2;full++)for(int a=0;a<3;a++)for(int s=0;s<4;s++) {
        int w=widths[s],h=heights[s];
        avifImage *im=avifImageCreate(w,h,8,AVIF_PIXEL_FORMAT_YUV420);if(!im)return 2;
        im->matrixCoefficients=matrices[m];im->yuvRange=full?AVIF_RANGE_FULL:AVIF_RANGE_LIMITED;
        if(avifImageAllocatePlanes(im,a?AVIF_PLANES_ALL:AVIF_PLANES_YUV)!=AVIF_RESULT_OK)return 3;
        for(int p=0;p<3;p++)for(int y=0;y<(p?(h+1)/2:h);y++)for(int x=0;x<(p?(w+1)/2:w);x++)
            im->yuvPlanes[p][y*im->yuvRowBytes[p]+x]=values[(x*3+y*7+p*5+s)%11];
        if(a)for(int y=0;y<h;y++)for(int x=0;x<w;x++)im->alphaPlane[y*im->alphaRowBytes+x]=a==2?0:(uint8_t)((x*47+y*31)%256);
        avifRGBImage rgb;avifRGBImageSetDefaults(&rgb,im);rgb.format=AVIF_RGB_FORMAT_RGBA;
        rgb.avoidLibYUV=AVIF_TRUE;rgb.chromaUpsampling=AVIF_CHROMA_UPSAMPLING_BILINEAR;
        if(avifRGBImageAllocatePixels(&rgb)!=AVIF_RESULT_OK || avifImageYUVToRGB(im,&rgb)!=AVIF_RESULT_OK)return 4;
        if(!first)putchar(',');first=0;
        printf("{\"width\":%d,\"height\":%d,\"matrix\":%d,\"fullRange\":%s,\"alphaMode\":%d,\"planes\":[",w,h,matrices[m],full?"true":"false",a);
        for(int p=0;p<3;p++){if(p)putchar(',');hex(im->yuvPlanes[p],im->yuvRowBytes[p],p?(w+1)/2:w,p?(h+1)/2:h);}
        printf("],\"alpha\":");if(a)hex(im->alphaPlane,im->alphaRowBytes,w,h);else printf("null");
        printf(",\"rgba\":");hex(rgb.pixels,rgb.rowBytes,w*4,h);putchar('}');
        avifRGBImageFreePixels(&rgb);avifImageDestroy(im);
    }
    printf("]}\n");return 0;
}
