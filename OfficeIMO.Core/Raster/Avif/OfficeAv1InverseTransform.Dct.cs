using System;

namespace OfficeIMO.Drawing;

internal static partial class OfficeAv1InverseTransform {
    // AV1 7.13.2.3: the same staged butterfly network serves 4/8/16/32/64-point DCTs.
    private static void Dct(int[] t,int[] copy,int n,int range) {
        int count=1<<n;Array.Copy(t,copy,count);
        for(int i=0;i<count;i++) t[i]=copy[Reverse(n,i)];
        if(n==6) for(int i=0;i<16;i++) Rotate(t,32+i,63-i,63-4*Reverse(4,i),false,range);
        if(n>=5) for(int i=0;i<8;i++) Rotate(t,16+i,31-i,6+(Reverse(3,7-i)<<3),false,range);
        if(n==6) for(int i=0;i<16;i++) Sum(t,32+2*i,33+2*i,(i&1)!=0,range);
        if(n>=4) for(int i=0;i<4;i++) Rotate(t,8+i,15-i,12+(Reverse(2,3-i)<<4),false,range);
        if(n>=5) for(int i=0;i<8;i++) Sum(t,16+2*i,17+2*i,(i&1)!=0,range);
        if(n==6) for(int i=0;i<4;i++) for(int j=0;j<2;j++) Rotate(t,62-4*i-j,33+4*i+j,60-16*Reverse(2,i)+64*j,true,range);
        if(n>=3) for(int i=0;i<2;i++) Rotate(t,4+i,7-i,56-32*i,false,range);
        if(n>=4) for(int i=0;i<4;i++) Sum(t,8+2*i,9+2*i,(i&1)!=0,range);
        if(n>=5) for(int i=0;i<2;i++) for(int j=0;j<2;j++) Rotate(t,30-4*i-j,17+4*i+j,24+(j<<6)+((1-i)<<5),true,range);
        if(n==6) for(int i=0;i<8;i++) for(int j=0;j<2;j++) Sum(t,32+4*i+j,35+4*i-j,(i&1)!=0,range);
        for(int i=0;i<2;i++) Rotate(t,2*i,2*i+1,32+16*i,i==0,range);
        if(n>=3) for(int i=0;i<2;i++) Sum(t,4+2*i,5+2*i,i!=0,range);
        if(n>=4) for(int i=0;i<2;i++) Rotate(t,14-i,9+i,48+64*i,true,range);
        if(n>=5) for(int i=0;i<4;i++) for(int j=0;j<2;j++) Sum(t,16+4*i+j,19+4*i-j,(i&1)!=0,range);
        if(n==6) for(int i=0;i<2;i++) for(int j=0;j<4;j++) Rotate(t,61-8*i-j,34+8*i+j,56-32*i+(j>>1)*64,true,range);
        for(int i=0;i<2;i++) Sum(t,i,3-i,false,range);
        if(n>=3) Rotate(t,6,5,32,true,range);
        if(n>=4) for(int i=0;i<2;i++) for(int j=0;j<2;j++) Sum(t,8+4*i+j,11+4*i-j,i!=0,range);
        if(n>=5) for(int i=0;i<4;i++) Rotate(t,29-i,18+i,48+(i>>1)*64,true,range);
        if(n==6) for(int i=0;i<4;i++) for(int j=0;j<4;j++) Sum(t,32+8*i+j,39+8*i-j,(i&1)!=0,range);
        if(n>=3) for(int i=0;i<4;i++) Sum(t,i,7-i,false,range);
        if(n>=4) for(int i=0;i<2;i++) Rotate(t,13-i,10+i,32,true,range);
        if(n>=5) for(int i=0;i<2;i++) for(int j=0;j<4;j++) Sum(t,16+8*i+j,23+8*i-j,i!=0,range);
        if(n==6) for(int i=0;i<8;i++) Rotate(t,59-i,36+i,i<4?48:112,true,range);
        if(n>=4) for(int i=0;i<8;i++) Sum(t,i,15-i,false,range);
        if(n>=5) for(int i=0;i<4;i++) Rotate(t,27-i,20+i,32,true,range);
        if(n==6) for(int i=0;i<8;i++) {Sum(t,32+i,47-i,false,range);Sum(t,48+i,63-i,true,range);}
        if(n>=5) for(int i=0;i<16;i++) Sum(t,i,31-i,false,range);
        if(n==6) for(int i=0;i<8;i++) Rotate(t,55-i,40+i,32,true,range);
        if(n==6) for(int i=0;i<32;i++) Sum(t,i,63-i,false,range);
    }
}
