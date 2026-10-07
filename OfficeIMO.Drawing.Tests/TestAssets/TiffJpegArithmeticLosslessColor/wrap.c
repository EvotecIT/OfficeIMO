/* Test-only LibTIFF container writer. JPEG segments are independently encoded. */
#include <tiffio.h>
#include <stdio.h>
#include <stdlib.h>

int main(int argc, char **argv) {
    if (argc < 10) return 1;
    int photo=atoi(argv[2]), bits=atoi(argv[3]), channels=atoi(argv[4]);
    int planar=atoi(argv[5]), tiled=atoi(argv[6]), extra=atoi(argv[7]);
    int horizontal=getenv("TIFF_SAMPLE_H")?atoi(getenv("TIFF_SAMPLE_H")):1;
    int vertical=getenv("TIFF_SAMPLE_V")?atoi(getenv("TIFF_SAMPLE_V")):1;
    int position=getenv("TIFF_SAMPLE_POSITION")?atoi(getenv("TIFF_SAMPLE_POSITION")):1;
    TIFF *t=TIFFOpen(argv[1],argv[8]);
    if (!t) return 2;
    TIFFSetField(t,TIFFTAG_IMAGEWIDTH,35); TIFFSetField(t,TIFFTAG_IMAGELENGTH,19);
    TIFFSetField(t,TIFFTAG_BITSPERSAMPLE,bits); TIFFSetField(t,TIFFTAG_SAMPLESPERPIXEL,channels);
    /* LibTIFF's JPEG writer bookkeeping rejects YCbCr extras. Only raw segments
       are written; restore the declared photometric tag after closing. */
    TIFFSetField(t,TIFFTAG_PHOTOMETRIC,photo==6?2:photo);
    TIFFSetField(t,TIFFTAG_PLANARCONFIG,planar); TIFFSetField(t,TIFFTAG_COMPRESSION,7);
    TIFFSetField(t,TIFFTAG_JPEGTABLESMODE,0);
    if (extra>=0) { uint16_t value=(uint16_t)extra; TIFFSetField(t,TIFFTAG_EXTRASAMPLES,1,&value); }
    if (photo==6) {
        float maximum=(float)((1U<<bits)-1), middle=(float)(1U<<(bits-1));
        float reference[6]={0,maximum,middle,maximum,middle,maximum};
        TIFFSetField(t,TIFFTAG_YCBCRSUBSAMPLING,horizontal,vertical);
        if (horizontal>1 || vertical>1) TIFFSetField(t,TIFFTAG_YCBCRPOSITIONING,position);
        TIFFSetField(t,TIFFTAG_REFERENCEBLACKWHITE,reference);
    }
    if (tiled) { TIFFSetField(t,TIFFTAG_TILEWIDTH,16); TIFFSetField(t,TIFFTAG_TILELENGTH,16); }
    else TIFFSetField(t,TIFFTAG_ROWSPERSTRIP,vertical>1?8:7);
    for (int i=9;i<argc;i++) {
        FILE *f=fopen(argv[i],"rb"); if(!f)return 3;
        fseek(f,0,SEEK_END); long size=ftell(f); rewind(f);
        unsigned char *data=malloc(size); if(!data||fread(data,1,size,f)!=(size_t)size)return 4;
        fclose(f);
        tmsize_t written=tiled?TIFFWriteRawTile(t,i-9,data,size):TIFFWriteRawStrip(t,i-9,data,size);
        free(data); if(written!=size)return 5;
    }
    TIFFClose(t);
    if(photo==6) {
        int big=argv[8][1]=='b',found=0; FILE *f=fopen(argv[1],"r+b"); unsigned char h[8],e[12],c[2];
        if(!f||fread(h,1,8,f)!=8)return 6;
        unsigned int offset=big?((unsigned)h[4]<<24)|((unsigned)h[5]<<16)|(h[6]<<8)|h[7]:h[4]|(h[5]<<8)|((unsigned)h[6]<<16)|((unsigned)h[7]<<24);
        fseek(f,offset,SEEK_SET); if(fread(c,1,2,f)!=2)return 7;
        int count=big?(c[0]<<8)|c[1]:c[0]|(c[1]<<8);
        for(int i=0;i<count;i++) {
            long at=offset+2+12*i; fseek(f,at,SEEK_SET); if(fread(e,1,12,f)!=12)return 8;
            int tag=big?(e[0]<<8)|e[1]:e[0]|(e[1]<<8);
            if(tag==262) { fseek(f,at+8,SEEK_SET); fputc(big?0:6,f); fputc(big?6:0,f); found=1; break; }
        }
        fclose(f); if(!found)return 9;
    }
    return 0;
}
