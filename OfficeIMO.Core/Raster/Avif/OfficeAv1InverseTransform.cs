using System;

namespace OfficeIMO.Drawing;

/// <summary>Integer Main-8/Main-10 inverse transforms; output samples are signed residuals in paint order.</summary>
/// <remarks>AV1 sections 7.12–7.13. No floating-point approximation or native codec is used.</remarks>
internal static partial class OfficeAv1InverseTransform {
    private static readonly byte[] RowKinds={0,0,1,1,0,1,1,1,1,2,2,0,2,1,2,1};
    private static readonly byte[] ColKinds={0,1,0,1,1,0,1,1,1,2,0,2,1,2,1,2};
    private static readonly byte[] RowShifts={0,1,2,2,2,0,0,1,1,1,1,1,1,1,1,2,2,2,2};
    private static readonly short[] Cosine={4096,4095,4091,4085,4076,4065,4052,4036,4017,3996,3973,3948,3920,3889,3857,3822,3784,3745,3703,3659,3612,3564,3513,3461,3406,3349,3290,3229,3166,3102,3035,2967,2896,2824,2751,2675,2598,2520,2440,2359,2276,2191,2106,2019,1931,1842,1751,1660,1567,1474,1380,1285,1189,1092,995,897,799,700,601,501,401,301,201,101,0};

    internal static void Validate(int size,int type,bool lossless) {
        if((uint)size>=19 || (uint)type>=16) throw new FormatException("Invalid AV1 inverse transform identifier.");
        int w=OfficeAv1TransformSize.Width(size),h=OfficeAv1TransformSize.Height(size);
        if((lossless && (size!=0 || type!=0)) ||
           ((w==64 || h==64) && type!=0) ||
           (RowKinds[type]==1 && w>16) || (ColKinds[type]==1 && h>16))
            throw new FormatException("Invalid AV1 inverse transform size or type.");
    }

    // The caller supplies a full-sized dequantized/output buffer and two bounded line buffers.
    internal static void Apply(int[] values,int size,int type,bool lossless,int bitDepth,int[] line,int[] copy,System.Threading.CancellationToken cancellation) {
        Validate(size,type,lossless);
        if(bitDepth!=8 && bitDepth!=10) throw new FormatException("Unsupported AV1 inverse transform bit depth.");
        int w=OfficeAv1TransformSize.Width(size),h=OfficeAv1TransformSize.Height(size);
        int rowRange=bitDepth+8,colRange=Math.Max(bitDepth+6,16);
        int rowShift=lossless?0:RowShifts[size],colShift=lossless?0:4;
        bool rectangular=Math.Abs(Log2(w)-Log2(h))==1;
        for(int row=0;row<h;row++) {
            cancellation.ThrowIfCancellationRequested();
            for(int col=0;col<w;col++) line[col]=rectangular?Round((long)values[row*w+col]*2896,12):values[row*w+col];
            Transform(line,copy,w,RowKinds[type],lossless,2,rowRange);
            for(int col=0;col<w;col++) values[row*w+col]=Clamp(Round(line[col],rowShift),colRange);
        }
        bool flipX=type==5 || type==6 || type==7 || type==15;
        bool flipY=type==4 || type==6 || type==8 || type==14;
        for(int col=0;col<w;col++) {
            cancellation.ThrowIfCancellationRequested();
            for(int row=0;row<h;row++) line[row]=values[row*w+col];
            Transform(line,copy,h,ColKinds[type],lossless,0,colRange);
            for(int row=0;row<h;row++) {
                int value=Round(line[row],colShift);
                if(lossless && (value< -(1<<bitDepth) || value>=(1<<bitDepth))) throw new FormatException("AV1 lossless residual exceeds the sample range.");
                values[(flipY?h-1-row:row)*w+col]=value;
            }
        }
        if(flipX) for(int row=0;row<h;row++) for(int col=0;col<w/2;col++) {
            int a=row*w+col,b=row*w+w-1-col,v=values[a];values[a]=values[b];values[b]=v;
        }
    }

    private static void Transform(int[] t,int[] copy,int count,int kind,bool lossless,int whtShift,int range) {
        if(lossless) {
            int a=t[0]>>whtShift,c=t[1]>>whtShift,d=t[2]>>whtShift,b=t[3]>>whtShift;
            a+=c;d-=b;int e=(a-d)>>1;b=e-b;c=e-c;a-=b;d+=c;
            t[0]=a;t[1]=b;t[2]=c;t[3]=d;
        } else if(kind==0) Dct(t,copy,Log2(count),range);
        else if(kind==1) Adst(t,copy,Log2(count),range);
        else for(int i=0;i<count;i++) t[i]=count==4?Round((long)t[i]*5793,12):count==8?t[i]*2:count==16?Round((long)t[i]*11586,12):t[i]*4;
    }
    private static int Log2(int value) {int n=0;while((value>>=1)!=0)n++;return n;}
    private static int Reverse(int n,int x) {int result=0;for(int i=0;i<n;i++) result=(result<<1)|((x>>i)&1);return result;}
    private static int Round(long value,int shift) => (int)(shift==0?value:(value+(1L<<(shift-1)))>>shift);
    private static int Clamp(int value,int bits) => Math.Max(-(1<<(bits-1)),Math.Min((1<<(bits-1))-1,value));
    private static int Cos(int angle) {angle&=255;return angle<=64?Cosine[angle]:angle<=128?-Cosine[128-angle]:angle<=192?-Cosine[angle-128]:Cosine[256-angle];}
    private static void Rotate(int[] t,int a,int b,int angle,bool flip,int range) {
        int c=Cos(angle),s=Cos(angle-64),x=Round((long)t[a]*c-(long)t[b]*s,12),y=Round((long)t[a]*s+(long)t[b]*c,12);
        // AV1 7.13.2.1 requires butterfly outputs to fit the current row/column precision.
        int limit=1<<(range-1);
        if(x< -limit || x>=limit || y< -limit || y>=limit) throw new FormatException("AV1 butterfly exceeds its intermediate range.");
        t[a]=flip?y:x;t[b]=flip?x:y;
    }
    private static void Sum(int[] t,int a,int b,bool flip,int range) {
        if(flip) {int temp=a;a=b;b=temp;}
        int x=t[a],y=t[b];t[a]=Clamp(x+y,range);t[b]=Clamp(x-y,range);
    }
}
