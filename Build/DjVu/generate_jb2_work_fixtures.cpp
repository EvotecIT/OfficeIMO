// OfficeIMO-authored opt-in test input producer. Native ZP arithmetic coding;
// integer decisions and record sequence follow the DjVu v3 appendix 2 grammar.
// This executable, its native library and headers are validation-only.
#include "ZPCodec.h"
#include "ByteStream.h"
#include "GURL.h"
#include "JB2Image.h"
#include "GException.h"
#include <array>
#include <vector>
#include <string>
#include <cstdio>
using namespace DJVU;
struct Node { unsigned char context=0; int children[2]={-1,-1}; };
struct Writer {
 GP<ZPCodec> zp; std::vector<Node> nodes; std::array<int,16> roots; unsigned char direct[1024]={0}, eventual=0, offset=0;
 Writer(const char* path):zp(ZPCodec::create(ByteStream::create(GURL::Filename::UTF8(path),"wb"),true,true)) {
  nodes.reserve(100000); for(auto &r: roots){r=nodes.size();nodes.push_back(Node());}
 }
 int child(int at,int bit) {if(nodes[at].children[bit]<0){int next=nodes.size();nodes[at].children[bit]=next;nodes.push_back(Node());}return nodes[at].children[bit];}
 int bit(int at,int value,bool inferred) {if(!inferred)zp->encoder(value,nodes[at].context);return child(at,value);}
 void number(int context,int low,int high,int value) {
  int n=roots[context],positive=value>=0; n=bit(n,positive,low>=0||high<0);
  if(!positive){int l=low;low=-high-1;high=-l-1;value=-value-1;}if(low<0)low=0;
  int first=0,width=1;
  for(;;){int last=first+width-1,beyond=value>last;n=bit(n,beyond,low>last||high<=last);if(!beyond)break;first+=width;width*=2;}
  while(width>1){width/=2;int middle=first+width,upper=value>=middle;n=bit(n,upper,low>=middle||high<middle);if(upper)first=middle;}
 }
 void record(int type){number(0,0,11,type);}
 void start(int w,int h){record(0);number(1,0,262142,w);number(1,0,262142,h);zp->encoder(0,eventual);}
 void comment(int length){record(10);number(13,0,262142,length);for(int i=0;i<length;i++)number(14,0,255,'A');}
 void empty(int type,int height){record(type);number(3,0,262142,0);number(4,0,262142,height);if(type==8){number(7,1,128,1);number(8,1,96,1);}}
 void solid(int height){record(1);number(3,0,262142,1);number(4,0,262142,height);
  for(int y=height-1;y>=0;y--){int ctx=(y+2<height?256:0)|(y+1<height?16:0);zp->encoder(1,direct[ctx]);}
  zp->encoder(1,offset);number(11,-262143,262142,1);number(12,-262143,262142,0);
 }
 void repeat(){record(7);number(2,0,0,0);zp->encoder(1,offset);number(11,-262143,262142,0);number(12,-262143,262142,0);}
 void finish(){record(11);zp=0;}
};
int main(int argc,char** argv){
 if(argc!=2)return 2;std::string out=argv[1];
 {Writer w((out+"/comments-page.jb2").c_str());w.start(128,96);w.comment(4096);w.comment(4096);w.finish();}
 {Writer w((out+"/comments-dictionary.jb2").c_str());w.start(0,0);w.comment(4096);w.finish();}
 {Writer w((out+"/zero-area.jb2").c_str());w.start(128,96);for(int i=0;i<1000;i++)w.empty(2,65535);w.empty(8,65535);w.finish();}
 {Writer w((out+"/clipped-masks.jb2").c_str());w.start(2,65535);w.solid(65535);for(int i=0;i<4999;i++)w.repeat();w.finish();}
 try{
  for(const char* name:{"comments-page.jb2","zero-area.jb2","clipped-masks.jb2"}){
   auto image=JB2Image::create();image->decode(ByteStream::create(GURL::Filename::UTF8((out+"/"+name).c_str()),"rb"));
   std::printf("Native accepted %s: %dx%d, shapes=%d, placements=%d, comment=%d\n",name,image->get_width(),image->get_height(),image->get_shape_count(),image->get_blit_count(),image->comment.length());
  }
 }catch(const GException& e){std::fprintf(stderr,"%s\n",e.get_cause());return 1;}
}
