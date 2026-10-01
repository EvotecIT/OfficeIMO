using System;

namespace OfficeIMO.Drawing;

/// <summary>Main-8 directional deringing arithmetic from AV1 7.15.2–3.</summary>
internal static class OfficeAv1CdefFilter {
    private static readonly int[] Divisors={0,840,420,280,210,168,140,120,105};
    private static readonly int[,,] Directions={
        {{-1,1},{-2,2}},{{0,1},{-1,2}},{{0,1},{0,2}},{{0,1},{1,2}},
        {{1,1},{2,2}},{{1,0},{2,1}},{{1,0},{2,0}},{{1,0},{2,-1}}};

    /// <summary>Finds direction and variance using the owner's reusable bounded scratch.</summary>
    internal static int FindDirection(byte[] pixels,int stride,int offset,int[] partial,int[] cost,out int variance) {
        if(stride<8 || offset<0 || offset+(long)7*stride+7>=pixels.Length || partial.Length<120 || cost.Length<8)
            throw new FormatException("Invalid AV1 CDEF direction footprint.");
        Array.Clear(partial,0,120);Array.Clear(cost,0,8);
        for(int i=0;i<8;i++) for(int j=0;j<8;j++) {
            int x=pixels[offset+i*stride+j]-128;
            partial[i+j]+=x;partial[15+i+j/2]+=x;partial[30+i]+=x;
            partial[45+3+i-j/2]+=x;partial[60+7+i-j]+=x;
            partial[75+3-i/2+j]+=x;partial[90+j]+=x;partial[105+i/2+j]+=x;
        }
        for(int i=0;i<8;i++) {cost[2]+=Square(partial[30+i]);cost[6]+=Square(partial[90+i]);}
        cost[2]*=Divisors[8];cost[6]*=Divisors[8];
        for(int i=0;i<7;i++) {
            cost[0]+=(Square(partial[i])+Square(partial[14-i]))*Divisors[i+1];
            cost[4]+=(Square(partial[60+i])+Square(partial[74-i]))*Divisors[i+1];
        }
        cost[0]+=Square(partial[7])*Divisors[8];cost[4]+=Square(partial[67])*Divisors[8];
        for(int i=1;i<8;i+=2) {
            for(int j=0;j<5;j++) cost[i]+=Square(partial[i*15+3+j]);
            cost[i]*=Divisors[8];
            for(int j=0;j<3;j++) cost[i]+=(Square(partial[i*15+j])+Square(partial[i*15+10-j]))*Divisors[2*j+2];
        }
        int best=0,direction=0;
        for(int i=0;i<8;i++) if(cost[i]>best) {best=cost[i];direction=i;}
        variance=(best-cost[(direction+4)&7])>>10;return direction;
    }

    /// <summary>Filters one 8x8 or subsampled 4x4 block, reading only the immutable input plane.</summary>
    internal static void Apply(byte[] input,byte[] output,int stride,int x0,int y0,int size,
        int width,int height,int primary,int secondary,int damping,int direction) {
        if(ReferenceEquals(input,output) || (size!=4 && size!=8) || primary<0 || primary>15 ||
           (secondary!=0 && secondary!=1 && secondary!=2 && secondary!=4) || damping<2 || damping>6 ||
           direction<0 || direction>7 || width<1 || width>stride || height<1 || x0<0 || y0<0 ||
           x0+size>width || y0+size>height || (long)stride*height>input.Length || (long)stride*height>output.Length)
            throw new FormatException("Invalid AV1 CDEF filter footprint or parameters.");
        int primaryShift=primary==0?0:Math.Max(0,damping-Log2(primary));
        int secondaryShift=secondary==0?0:Math.Max(0,damping-Log2(secondary));
        for(int y=y0;y<y0+size;y++) for(int x=x0;x<x0+size;x++) {
            int value=input[y*stride+x],sum=0,min=value,max=value;
            for(int k=0;k<2;k++) for(int sign=-1;sign<=1;sign+=2) {
                int primaryTap=(primary&1)==0?(k==0?4:2):3;
                Add(input,stride,x,y,width,height,direction,k,sign,value,primary,primaryTap,primaryShift,ref sum,ref min,ref max);
                for(int d=-2;d<=2;d+=4)
                    Add(input,stride,x,y,width,height,(direction+d)&7,k,sign,value,secondary,k==0?2:1,secondaryShift,ref sum,ref min,ref max);
            }
            output[y*stride+x]=(byte)Math.Max(min,Math.Min(max,value+((8+sum-(sum<0?1:0))>>4)));
        }
    }
    private static void Add(byte[] input,int stride,int x,int y,int width,int height,int direction,int k,int sign,
        int value,int strength,int tap,int shift,ref int sum,ref int min,ref int max) {
        int ny=y+sign*Directions[direction,k,0],nx=x+sign*Directions[direction,k,1];
        if((uint)ny>=(uint)height || (uint)nx>=(uint)width) return;
        int sample=input[ny*stride+nx],diff=sample-value,absolute=Math.Abs(diff);
        if(strength!=0) {
            sum+=tap*(diff<0?-1:1)*Math.Max(0,Math.Min(absolute,strength-(absolute>>shift)));
        }
        min=Math.Min(min,sample);max=Math.Max(max,sample);
    }
    internal static int AdjustStrength(int strength,int variance)=>variance==0?0:
        (strength*(4+((variance>>6)==0?0:Math.Min(Log2(variance>>6),12)))+8)>>4;
    private static int Square(int value)=>value*value;
    private static int Log2(int value) {int n=0;while((value>>=1)!=0)n++;return n;}
}
