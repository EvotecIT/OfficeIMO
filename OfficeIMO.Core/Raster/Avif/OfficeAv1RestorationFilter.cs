using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Main-8 Wiener and self-guided restoration, with immutable stripe-aware sample sources.</summary>
/// <remarks>Normative AV1 sections 7.17.2–6. Scratch is reused for at most 64x64 output samples.</remarks>
internal sealed class OfficeAv1RestorationFilter {
    private static readonly int[,] SgrParameters={
        {2,12,1,4},{2,15,1,6},{2,18,1,8},{2,21,1,9},{2,24,1,10},{2,29,1,11},
        {2,36,1,12},{2,45,1,13},{2,56,1,14},{2,68,1,15},{0,0,1,5},{0,0,1,8},
        {0,0,1,11},{0,0,1,14},{2,30,0,0},{2,75,0,0}};
    private readonly int[] _a=new int[66*66],_b=new int[66*66],_sum=new int[71*71],_squares=new int[71*71];
    private readonly int[] _flt0=new int[64*64],_flt1=new int[64*64],_intermediate=new int[70*64];
    private readonly int[] _horizontal=new int[7],_vertical=new int[7];
    private readonly CancellationToken _cancellation;
    private ushort[] _deblocked=Array.Empty<ushort>(),_cdef=Array.Empty<ushort>();
    private int _stride,_planeWidth,_planeHeight,_stripeStart,_stripeEnd;
    internal OfficeAv1RestorationFilter(CancellationToken cancellation) { _cancellation=cancellation; }

    /// <summary>Reads deblocked samples outside the stripe and CDEF samples inside, without output feedback.</summary>
    internal void Apply(ushort[] deblocked,ushort[] cdef,ushort[] output,int stride,int planeWidth,int planeHeight,
        int stripeStart,int stripeEnd,int x,int y,int width,int height,OfficeAv1RestorationUnit unit) {
        _cancellation.ThrowIfCancellationRequested();
        if(ReferenceEquals(output,deblocked) || ReferenceEquals(output,cdef) || stride<planeWidth || planeWidth<1 || planeHeight<1 ||
           (long)stride*planeHeight>deblocked.Length || (long)stride*planeHeight>cdef.Length || (long)stride*planeHeight>output.Length ||
           width<1 || width>64 || height<1 || height>64 || x<0 || y<0 || x+(long)width>planeWidth || y+(long)height>planeHeight ||
           stripeStart< -8 || stripeEnd>(long)planeHeight+63 || stripeStart>y || stripeEnd<y+height-1 ||
           stripeStart>stripeEnd || (uint)unit.Plane>2 || (unit.Type!=2 && unit.Type!=3))
            throw new FormatException("Invalid AV1 restoration footprint or sample sources.");
        _deblocked=deblocked;_cdef=cdef;_stride=stride;_planeWidth=planeWidth;_planeHeight=planeHeight;
        _stripeStart=stripeStart;_stripeEnd=stripeEnd;
        if(unit.Type==2) Wiener(output,x,y,width,height,unit);
        else SelfGuided(output,x,y,width,height,unit);
    }
    private int Sample(int x,int y) {
        x=Math.Max(0,Math.Min(_planeWidth-1,x));y=Math.Max(0,Math.Min(_planeHeight-1,y));
        if(y<_stripeStart) return _deblocked[Math.Max(_stripeStart-2,y)*_stride+x];
        if(y>_stripeEnd) return _deblocked[Math.Min(_stripeEnd+2,y)*_stride+x];
        return _cdef[y*_stride+x];
    }
    private void Wiener(ushort[] output,int x,int y,int width,int height,OfficeAv1RestorationUnit unit) {
        for(int pass=0;pass<2;pass++) {
            int[] filter=pass==0?_vertical:_horizontal;filter[3]=128;
            for(int t=0;t<3;t++) {
                int v=unit.WienerTap(pass,t),low=t==0?-5:t==1?-23:-17,high=t==0?10:t==1?8:46;
                if(v<low || v>high || (unit.Plane>0 && t==0 && v!=0)) throw new FormatException("Invalid AV1 Wiener coefficients.");
                filter[t]=filter[6-t]=v;filter[3]-=2*v;
            }
        }
        for(int r=0;r<height+6;r++) {
            _cancellation.ThrowIfCancellationRequested();
            for(int c=0;c<width;c++) {
                int sum=0;for(int t=0;t<7;t++) sum+=_horizontal[t]*Sample(x+c+t-3,y+r-3);
                _intermediate[r*width+c]=Math.Max(-2048,Math.Min(6143,(sum+4)>>3));
            }
        }
        for(int r=0;r<height;r++) {
            _cancellation.ThrowIfCancellationRequested();
            for(int c=0;c<width;c++) {
                int sum=0;for(int t=0;t<7;t++) sum+=_vertical[t]*_intermediate[(r+t)*width+c];
                output[(y+r)*_stride+x+c]=Clip((sum+1024)>>11);
            }
        }
    }
    private void SelfGuided(ushort[] output,int x,int y,int width,int height,OfficeAv1RestorationUnit unit) {
        int set=unit.SgrSet;
        if((uint)set>=16 || unit.X0<-96 || unit.X0>31 || unit.X1<-32 || unit.X1>95 ||
           (set>=10 && set<=13 && unit.X0!=0)) throw new FormatException("Invalid AV1 self-guided coefficients.");
        int pitch=width+7,patchWidth=width+6,patchHeight=height+6;
        Array.Clear(_sum,0,pitch);Array.Clear(_squares,0,pitch);
        for(int r=0;r<patchHeight;r++) {
            _cancellation.ThrowIfCancellationRequested();int sum=0,square=0;
            _sum[(r+1)*pitch]=_squares[(r+1)*pitch]=0;
            for(int c=0;c<patchWidth;c++) {
                int v=Sample(x+c-3,y+r-3);sum+=v;square+=v*v;
                _sum[(r+1)*pitch+c+1]=_sum[r*pitch+c+1]+sum;
                _squares[(r+1)*pitch+c+1]=_squares[r*pitch+c+1]+square;
            }
        }
        if(SgrParameters[set,0]!=0) Box(width,height,set,0,pitch,x,y,_flt0);
        if(SgrParameters[set,2]!=0) Box(width,height,set,1,pitch,x,y,_flt1);
        for(int r=0;r<height;r++) {
            _cancellation.ThrowIfCancellationRequested();
            for(int c=0;c<width;c++) {
                int u=_cdef[(y+r)*_stride+x+c]<<4,index=r*width+c;
                int f0=SgrParameters[set,0]==0?u:_flt0[index],f1=SgrParameters[set,2]==0?u:_flt1[index];
                int value=unit.X1*u+unit.X0*f0+(128-unit.X0-unit.X1)*f1;
                output[(y+r)*_stride+x+c]=Clip((value+1024)>>11);
            }
        }
    }
    private void Box(int width,int height,int set,int pass,int pitch,int x,int y,int[] filtered) {
        int radius=SgrParameters[set,pass*2],eps=SgrParameters[set,pass*2+1],n=(2*radius+1)*(2*radius+1);
        int scale=((1<<20)+n*n*eps/2)/(n*n*eps),reciprocal=((1<<12)+n/2)/n,abPitch=width+2;
        for(int r=-1;r<=height;r++) {
            _cancellation.ThrowIfCancellationRequested();
            for(int c=-1;c<=width;c++) {
                int top=r+3-radius,left=c+3-radius,bottom=r+4+radius,right=c+4+radius;
                int sum=Area(_sum,pitch,top,left,bottom,right),square=Area(_squares,pitch,top,left,bottom,right);
                long variance=Math.Max(0,(long)square*n-(long)sum*sum),z=(variance*scale+(1<<19))>>20;
                int a=z>=255?256:z==0?1:(int)((z*256+z/2)/(z+1));
                int index=(r+1)*abPitch+c+1;_a[index]=a;
                _b[index]=(int)(((long)(256-a)*sum*reciprocal+(1<<11))>>12);
            }
        }
        for(int r=0;r<height;r++) {
            _cancellation.ThrowIfCancellationRequested();int shift=pass==0 && (r&1)!=0?4:5;
            for(int c=0;c<width;c++) {
                int a=0,b=0;
                for(int dy=-1;dy<=1;dy++) for(int dx=-1;dx<=1;dx++) {
                    int weight=pass==0?(((r+dy)&1)!=0?(dx==0?6:5):0):(dx==0 || dy==0?4:3);
                    int index=(r+1+dy)*abPitch+c+1+dx;a+=weight*_a[index];b+=weight*_b[index];
                }
                int value=a*_cdef[(y+r)*_stride+x+c]+b,bits=8+shift-4;
                filtered[r*width+c]=(value+(1<<(bits-1)))>>bits;
            }
        }
    }
    private static int Area(int[] integral,int pitch,int top,int left,int bottom,int right)=>
        integral[bottom*pitch+right]-integral[top*pitch+right]-integral[bottom*pitch+left]+integral[top*pitch+left];
    private static ushort Clip(int value)=>(ushort)Math.Max(0,Math.Min(255,value));
}
