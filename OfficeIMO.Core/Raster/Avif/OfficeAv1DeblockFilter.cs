using System;

namespace OfficeIMO.Drawing;

/// <summary>AV1 7.14.6 Main-8 narrow/wide filtering across one sample boundary.</summary>
/// <remarks>The owning frame validates geometry and retains this fixed scratch. Reads are captured
/// before writes so wide taps never consume their own modified samples.</remarks>
internal static class OfficeAv1DeblockFilter {
    internal static void Apply(byte[] pixels,int offset,int step,int size,bool chroma,
        int limit,int blimit,int threshold,int[] scratch) {
        if(size!=4 && size!=8 && size!=16 || chroma && size==16 || step<1 || scratch.Length<32)
            throw new ArgumentOutOfRangeException(nameof(size));
        int count=size==4?2:chroma?3:size==8?4:7;
        if(offset-(long)count*step<0 || offset+(long)(count-1)*step>=pixels.Length)
            throw new FormatException("AV1 deblocking footprint exceeds its plane.");
        for(int i=-count;i<count;i++) scratch[i+7]=pixels[offset+i*step];
        int p0=scratch[6],p1=scratch[5],q0=scratch[7],q1=scratch[8];
        if(Math.Abs(p1-p0)>limit || Math.Abs(q1-q0)>limit ||
           2*Math.Abs(p0-q0)+Math.Abs(p1-q1)/2>blimit) return;
        for(int i=2;i<Math.Min(count,4);i++)
            if(Math.Abs(scratch[6-i]-scratch[7-i])>limit || Math.Abs(scratch[7+i]-scratch[6+i])>limit) return;
        bool flat=size>4;
        for(int i=1;i<Math.Min(count,4);i++)
            flat&=Math.Abs(scratch[6-i]-p0)<=1 && Math.Abs(scratch[7+i]-q0)<=1;
        bool flat2=flat && size==16;
        if(flat2) for(int i=4;i<7;i++)
            flat2&=Math.Abs(scratch[6-i]-p0)<=1 && Math.Abs(scratch[7+i]-q0)<=1;
        if(flat) {
            int n=flat2?6:chroma?2:3,bits=flat2?4:3,n2=flat2 || chroma?1:0;
            for(int i=-n;i<n;i++) {
                int sum=0;
                for(int j=-n;j<=n;j++) sum+=scratch[7+Math.Max(-n-1,Math.Min(n,i+j))]*(Math.Abs(j)<=n2?2:1);
                scratch[16+i+n]=(sum+(1<<(bits-1)))>>bits;
            }
            for(int i=-n;i<n;i++) pixels[offset+i*step]=(byte)scratch[16+i+n];
            return;
        }
        bool hev=Math.Abs(p1-p0)>threshold || Math.Abs(q1-q0)>threshold;
        int f=Clamp((hev?Clamp(p1-q1):0)+3*(q0-p0));
        int a=Clamp(f+4)>>3,b=Clamp(f+3)>>3;
        pixels[offset]=(byte)(Clamp(q0-128-a)+128);
        pixels[offset-step]=(byte)(Clamp(p0-128+b)+128);
        if(!hev) {
            int outer=(a+1)>>1;
            pixels[offset+step]=(byte)(Clamp(q1-128-outer)+128);
            pixels[offset-2*step]=(byte)(Clamp(p1-128+outer)+128);
        }
    }
    private static int Clamp(int value)=>Math.Max(-128,Math.Min(127,value));
}
