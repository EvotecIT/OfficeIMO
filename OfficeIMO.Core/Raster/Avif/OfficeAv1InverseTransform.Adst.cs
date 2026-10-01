using System;

namespace OfficeIMO.Drawing;

internal static partial class OfficeAv1InverseTransform {
    private static void Adst(int[] t,int[] copy,int n,int range) {
        if(n==2) {Adst4(t,range);return;}
        int count=1<<n;Array.Copy(t,copy,count);
        for(int i=0;i<count;i++) t[i]=copy[(i&1)!=0?i-1:count-i-1];
        if(n==4) {
            for(int i=0;i<8;i++) Rotate(t,2*i,2*i+1,62-8*i,true,range);
            for(int i=0;i<8;i++) Sum(t,i,8+i,false,range);
            for(int i=0;i<2;i++) {Rotate(t,8+2*i,9+2*i,56-32*i,true,range);Rotate(t,13+2*i,12+2*i,8+32*i,true,range);}
            for(int j=0;j<2;j++) for(int i=0;i<4;i++) Sum(t,8*j+i,4+8*j+i,false,range);
            for(int j=0;j<2;j++) for(int i=0;i<2;i++) Rotate(t,4+8*j+3*i,5+8*j+i,48-32*i,true,range);
            for(int j=0;j<4;j++) for(int i=0;i<2;i++) Sum(t,4*j+i,2+4*j+i,false,range);
            for(int i=0;i<4;i++) Rotate(t,2+4*i,3+4*i,32,true,range);
        } else {
            for(int i=0;i<4;i++) Rotate(t,2*i,2*i+1,60-16*i,true,range);
            for(int i=0;i<4;i++) Sum(t,i,4+i,false,range);
            for(int i=0;i<2;i++) Rotate(t,4+3*i,5+i,48-32*i,true,range);
            for(int j=0;j<2;j++) for(int i=0;i<2;i++) Sum(t,4*j+i,2+4*j+i,false,range);
            for(int i=0;i<2;i++) Rotate(t,2+4*i,3+4*i,32,true,range);
        }
        Array.Copy(t,copy,count);
        for(int i=0;i<count;i++) {
            int a=(i>>3)&1,b=((i>>2)^(i>>3))&1,c=((i>>1)^(i>>2))&1,d=(i^(i>>1))&1;
            int index=((d<<3)|(c<<2)|(b<<1)|a)>>(4-n);t[i]=(i&1)!=0?-copy[index]:copy[index];
        }
    }
    private static void Adst4(int[] t,int range) {
        // AV1 7.13.2.6 bounds every stored s/x value before later transform clipping.
        int bits=range+12;
        long s0=AdstPrecision(1321L*t[0],bits),s1=AdstPrecision(2482L*t[0],bits);
        long s2=AdstPrecision(3344L*t[1],bits),s3=AdstPrecision(3803L*t[2],bits);
        long s4=AdstPrecision(1321L*t[2],bits),s5=AdstPrecision(2482L*t[3],bits),s6=AdstPrecision(3803L*t[3],bits);
        long a7=AdstPrecision((long)t[0]-t[2],range+1),b7=AdstPrecision(a7+t[3],range);
        s0=AdstPrecision(s0+s3,bits);s1=AdstPrecision(s1-s4,bits);
        s3=s2;s2=AdstPrecision(3344L*b7,bits);
        s0=AdstPrecision(s0+s5,bits);s1=AdstPrecision(s1-s6,bits);
        long x0=AdstPrecision(s0+s3,bits),x1=AdstPrecision(s1+s3,bits);
        long x3=AdstPrecision(s0+s1,bits);x3=AdstPrecision(x3-s3,bits);
        t[0]=Round(x0,12);t[1]=Round(x1,12);t[2]=Round(s2,12);t[3]=Round(x3,12);
    }
    private static long AdstPrecision(long value,int bits) {
        long limit=1L<<(bits-1);
        if(value< -limit || value>=limit) throw new FormatException("AV1 ADST intermediate exceeds its range.");
        return value;
    }
}
