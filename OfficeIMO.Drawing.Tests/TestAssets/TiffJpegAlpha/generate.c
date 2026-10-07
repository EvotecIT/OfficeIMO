#include <tiffio.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <jpeglib.h>
/* Independent component encoding/decoding; TIFF owns alpha and color meaning. */
static void samples(FILE*f,const J12SAMPLE*p,int count,int bits){for(int i=0;i<count;i++){fputc(p[i]&255,f);if(bits==12)fputc(p[i]>>8,f);}}
static void word(FILE*f,unsigned n){for(int i=0;i<4;i++)fputc((n>>(8*i))&255,f);}
int main(int argc,char**argv){
 if(argc<9||argc>11)return 2;
 int photo=atoi(argv[2]),big=atoi(argv[3]),tile=atoi(argv[4]),shared=atoi(argv[5]),planar=atoi(argv[6]),sub=atoi(argv[7]),extra=atoi(argv[8]);
 int bits=getenv("TIFF_JPEG_PRECISION")?atoi(getenv("TIFF_JPEG_PRECISION")):8;if(bits!=8&&bits!=12)return 2;int maximum=(1<<bits)-1,midpoint=1<<(bits-1);
 int lowAlpha=argc>=10?atoi(argv[9]):0;
 int base=photo==5?4:photo==2||photo==6?3:1,n=base+1,w=35,h=19,sw=tile?16:w,sh=16;
 uint16_t extras[4]={extra,0,0,0};int extraCount=1,alphaChannel=base;
 if(argc==11){extraCount=(int)strlen(argv[10]);if(extraCount<1||extraCount>4)return 2;alphaChannel=-1;extra=0;
  for(int i=0;i<extraCount;i++){extras[i]=argv[10][i]-'0';if(extras[i]>2)return 2;if(extras[i]){if(alphaChannel>=0)return 2;alphaChannel=base+i;extra=extras[i];}}n=base+extraCount;
 }
 if(getenv("TIFF_JPEG_OPAQUE")){n=base;extraCount=0;alphaChannel=-1;extra=0;}
 TIFF*t=TIFFOpen(argv[1],big?"wb":"wl");if(!t)return 3;
 TIFFSetField(t,256,w);TIFFSetField(t,257,h);TIFFSetField(t,258,bits);TIFFSetField(t,277,n);TIFFSetField(t,262,photo==6?2:photo);TIFFSetField(t,284,planar);TIFFSetField(t,259,7);
 if(extraCount)TIFFSetField(t,338,extraCount,extras);
 if(photo==6){TIFFSetField(t,530,sub,sub);float ref[6]={0,maximum,midpoint,maximum,midpoint,maximum};TIFFSetField(t,532,ref);}
 if(tile){TIFFSetField(t,322,sw);TIFFSetField(t,323,sh);}else TIFFSetField(t,278,sh);
 char planePath[4096];snprintf(planePath,sizeof(planePath),"%s.planes",argv[1]);FILE*planes=fopen(planePath,"wb");if(!planes)return 9;
 J12SAMPLE output[35*19*8]={0};
 for(int plane=0;plane<(planar==2?n:1);plane++)for(int y=0;y<h;y+=sh)for(int x=0;x<w;x+=sw){
  int rows=tile?sh:(h-y<sh?h-y:sh),channels=planar==2?1:n;
  int reduced=photo==6&&planar==2&&(plane==1||plane==2),dw=reduced?(sw+sub-1)/sub:sw,dh=reduced?(rows+sub-1)/sub:rows;
  J12SAMPLE pixels[35*16*8];
  for(int yy=0;yy<dh;yy++)for(int xx=0;xx<dw;xx++)for(int cc=0;cc<channels;cc++){
   int ch=planar==2?plane:cc,gx=x+xx*(reduced?sub:1),gy=y+yy*(reduced?sub:1);if(gx>=w)gx=w-1;if(gy>=h)gy=h-1;
   static const int alphaLevels[]={0,1,2,3,4,8,16,32,64,128,192,254,255};
   int alpha=lowAlpha?alphaLevels[(gx/3+gy/3)%13]:96+(gx+gy*2)%144,value=32+(gx*2+gy*3+ch*41)%160;
   if(bits==12){static const int levels12[]={0,1,2,3,4,8,16,32,64,2048,3072,4094,4095};alpha=lowAlpha?levels12[(gx/3+gy/3)%13]:alpha*4095/255;value=256+(gx*37+gy*53+ch*641)%3072;}
   if(ch==alphaChannel)value=alpha;
   else if(extra==1&&ch<base){if(photo==6&&ch>0)value=midpoint+(value-midpoint)*alpha/maximum;else value=value*alpha/maximum;}
   pixels[(yy*dw+xx)*channels+cc]=value;
  }
  struct jpeg_compress_struct c;struct jpeg_error_mgr err;c.err=jpeg_std_error(&err);jpeg_create_compress(&c);
  c.image_width=dw;c.image_height=dh;c.input_components=channels;c.in_color_space=JCS_UNKNOWN;jpeg_set_defaults(&c);c.data_precision=bits;jpeg_set_quality(&c,95,TRUE);
  if(getenv("TIFF_JPEG_ARITHMETIC")){c.arith_code=TRUE;c.restart_interval=3;}
  if(photo==6&&planar==1)for(int cc=0;cc<channels;cc++){c.comp_info[cc].h_samp_factor=(cc==0||cc>=3)?sub:1;c.comp_info[cc].v_samp_factor=(cc==0||cc>=3)?sub:1;}
  jpeg_scan_info scans[8];memset(scans,0,sizeof(scans));
  if(channels>4){for(int cc=0;cc<channels;cc++){scans[cc].comps_in_scan=1;scans[cc].component_index[0]=cc;scans[cc].Se=63;}c.scan_info=scans;c.num_scans=channels;}
  unsigned char*tables=NULL,*encoded=NULL;unsigned long tablesLength=0,encodedLength=0;
  if(shared){jpeg_mem_dest(&c,&tables,&tablesLength);jpeg_write_tables(&c);if(x==0&&y==0&&plane==0)TIFFSetField(t,347,(uint32_t)tablesLength,tables);}
  jpeg_mem_dest(&c,&encoded,&encodedLength);jpeg_start_compress(&c,!shared);
  while(c.next_scanline<c.image_height){J12SAMPROW row=pixels+c.next_scanline*dw*channels;if(bits==12)jpeg12_write_scanlines(&c,&row,1);else{JSAMPLE b[35*8];for(int i=0;i<dw*channels;i++)b[i]=row[i];JSAMPROW p=b;jpeg_write_scanlines(&c,&p,1);}}
  jpeg_finish_compress(&c);jpeg_destroy_compress(&c);
  if((tile?TIFFWriteRawTile(t,TIFFComputeTile(t,x,y,0,plane),encoded,encodedLength):TIFFWriteRawStrip(t,TIFFComputeStrip(t,y,plane),encoded,encodedLength))<0)return 4;
  J12SAMPLE decoded[35*16*8];
  for(int channel=0;channel<(channels>4?channels:1);channel++){
  int sampleWidth=dw,sampleHeight=dh;
  if(channels>4&&photo==6&&(channel==1||channel==2)){sampleWidth=(dw+sub-1)/sub;sampleHeight=(dh+sub-1)/sub;}
  unsigned char*decodeBytes=encoded;unsigned long decodeLength=encodedLength;unsigned char*split=NULL;
  if(channels>4){
   unsigned sof=0,firstScan=0,scanStart=0,scanEnd=0,pos=2;int scanIndex=0;
   while(pos+3<encodedLength){
    int marker=encoded[pos+1];if(marker==217)break;
    unsigned len=(encoded[pos+2]<<8)|encoded[pos+3];
    if(marker==192||marker==193||marker==201)sof=pos;
    if(marker==218){
     if(!firstScan)firstScan=pos;unsigned end=pos+2+len;
     while(end+1<encodedLength){if(encoded[end]!=255){end++;continue;}if(encoded[end+1]==0||(encoded[end+1]>=208&&encoded[end+1]<=215)){end+=2;continue;}break;}
     if(scanIndex++==channel){scanStart=pos;scanEnd=end;break;}pos=end;
    }else pos+=2+len;
   }
   if(!sof||!firstScan||!scanStart)return 8;
   split=malloc(encodedLength+2);unsigned at=0;memcpy(split,encoded,sof);at=sof;
   memcpy(split+at,encoded+sof,10);split[at+2]=0;split[at+3]=11;split[at+9]=1;at+=10;
   memcpy(split+at,encoded+sof+10+channel*3,3);split[at+1]=17;at+=3;
   split[sof+5]=sampleHeight>>8;split[sof+6]=sampleHeight;split[sof+7]=sampleWidth>>8;split[sof+8]=sampleWidth;
   unsigned after=sof+2+((encoded[sof+2]<<8)|encoded[sof+3]);memcpy(split+at,encoded+after,firstScan-after);at+=firstScan-after;
   memcpy(split+at,encoded+scanStart,scanEnd-scanStart);at+=scanEnd-scanStart;split[at++]=255;split[at++]=217;decodeBytes=split;decodeLength=at;
  }
  struct jpeg_decompress_struct d;d.err=jpeg_std_error(&err);jpeg_create_decompress(&d);
  if(shared){jpeg_mem_src(&d,tables,tablesLength);if(jpeg_read_header(&d,FALSE)!=JPEG_HEADER_TABLES_ONLY)return 5;}
  jpeg_mem_src(&d,decodeBytes,decodeLength);jpeg_read_header(&d,TRUE);d.jpeg_color_space=JCS_UNKNOWN;d.out_color_space=JCS_UNKNOWN;jpeg_start_decompress(&d);
  while(d.output_scanline<d.output_height){int n=sampleWidth*(channels>4?1:channels);J12SAMPROW row=decoded+d.output_scanline*n;if(bits==12)jpeg12_read_scanlines(&d,&row,1);else{JSAMPLE b[35*8];JSAMPROW p=b;jpeg_read_scanlines(&d,&p,1);for(int i=0;i<n;i++)row[i]=b[i];}}
  jpeg_finish_decompress(&d);jpeg_destroy_decompress(&d);
  if(argc==11&&channels>4){word(planes,channel);word(planes,x);word(planes,y);word(planes,sampleWidth);word(planes,sampleHeight);samples(planes,decoded,sampleWidth*sampleHeight,bits);}
  if(channels>4){for(int i=0;i<sampleWidth*sampleHeight;i++)pixels[i*channels+channel]=decoded[i];}else memcpy(pixels,decoded,dw*dh*channels*sizeof(J12SAMPLE));
  free(split);
  }
  /* Subsampled planar references are stored separately for external interpolation. */
  if(planar==2){word(planes,plane);word(planes,x);word(planes,y);word(planes,dw);word(planes,dh);samples(planes,pixels,dw*dh,bits);}
  if(reduced){free(encoded);free(tables);continue;}
  for(int yy=0;yy<dh&&y+yy<h;yy++)for(int xx=0;xx<dw&&x+xx<w;xx++)for(int cc=0;cc<channels;cc++)output[((y+yy)*w+x+xx)*n+(planar==2?plane:cc)]=pixels[(yy*dw+xx)*channels+cc];
  free(encoded);free(tables);
 }
 fclose(planes);TIFFClose(t);
 if(photo==6){
  FILE*patch=fopen(argv[1],"r+b");unsigned char header[8];fread(header,1,8,patch);
  unsigned ifd=big?(header[4]<<24)|(header[5]<<16)|(header[6]<<8)|header[7]:header[4]|(header[5]<<8)|(header[6]<<16)|(header[7]<<24);
  fseek(patch,ifd,SEEK_SET);unsigned char count[2];fread(count,1,2,patch);int entries=big?(count[0]<<8)|count[1]:count[0]|(count[1]<<8);
  for(int i=0;i<entries;i++){unsigned char entry[12];long at=ifd+2+12*i;fseek(patch,at,SEEK_SET);fread(entry,1,12,patch);int tag=big?(entry[0]<<8)|entry[1]:entry[0]|(entry[1]<<8);if(tag==262){fseek(patch,at+8,SEEK_SET);fputc(big?0:6,patch);fputc(big?6:0,patch);break;}}
  fclose(patch);
 }
 char path[4096];snprintf(path,sizeof(path),"%s.raw",argv[1]);FILE*f=fopen(path,"wb");if(!f)return 7;samples(f,output,w*h*n,bits);fclose(f);return 0;
}
