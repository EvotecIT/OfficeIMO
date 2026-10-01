using System;

namespace OfficeIMO.Drawing;

internal static partial class OfficeAv1InverseTransform {
    private static void Adst(int[] t,int[] copy,int n) {
        if(n==2) {Adst4(t);return;}
        int count=1<<n;Array.Copy(t,copy,count);
        for(int i=0;i<count;i++) t[i]=copy[(i&1)!=0?i-1:count-i-1];
        if(n==4) {
            for(int i=0;i<8;i++) Rotate(t,2*i,2*i+1,62-8*i,true);
            for(int i=0;i<8;i++) Sum(t,i,8+i,false);
            for(int i=0;i<2;i++) {Rotate(t,8+2*i,9+2*i,56-32*i,true);Rotate(t,13+2*i,12+2*i,8+32*i,true);}
            for(int j=0;j<2;j++) for(int i=0;i<4;i++) Sum(t,8*j+i,4+8*j+i,false);
            for(int j=0;j<2;j++) for(int i=0;i<2;i++) Rotate(t,4+8*j+3*i,5+8*j+i,48-32*i,true);
            for(int j=0;j<4;j++) for(int i=0;i<2;i++) Sum(t,4*j+i,2+4*j+i,false);
            for(int i=0;i<4;i++) Rotate(t,2+4*i,3+4*i,32,true);
        } else {
            for(int i=0;i<4;i++) Rotate(t,2*i,2*i+1,60-16*i,true);
            for(int i=0;i<4;i++) Sum(t,i,4+i,false);
            for(int i=0;i<2;i++) Rotate(t,4+3*i,5+i,48-32*i,true);
            for(int j=0;j<2;j++) for(int i=0;i<2;i++) Sum(t,4*j+i,2+4*j+i,false);
            for(int i=0;i<2;i++) Rotate(t,2+4*i,3+4*i,32,true);
        }
        Array.Copy(t,copy,count);
        for(int i=0;i<count;i++) {
            int a=(i>>3)&1,b=((i>>2)^(i>>3))&1,c=((i>>1)^(i>>2))&1,d=(i^(i>>1))&1;
            int index=((d<<3)|(c<<2)|(b<<1)|a)>>(4-n);t[i]=(i&1)!=0?-copy[index]:copy[index];
        }
    }
    private static void Adst4(int[] t) {
        long s0=1321L*t[0]+3803L*t[2]+2482L*t[3];
        long s1=2482L*t[0]-1321L*t[2]-3803L*t[3],s3=3344L*t[1];
        int a7=t[0]-t[2],b7=a7+t[3];
        if(b7< -32768 || b7>32767) throw new FormatException("AV1 ADST intermediate exceeds its range.");
        t[0]=Round(s0+s3,12);t[1]=Round(s1+s3,12);t[2]=Round(3344L*b7,12);t[3]=Round(s0+s1-s3,12);
    }
}
