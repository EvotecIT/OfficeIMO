using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeAv1IntraPredictor {
    private void Directional(byte[] output,int w,int h,int angle,OfficeAv1PredictionEdges edges,bool filter,bool smooth) {
        if((angle<=90 && edges.AboveCount==0) || (angle>=180 && edges.LeftCount==0)) {
            byte value=angle<=90?(edges.LeftCount>0?edges.Left[0]:(byte)127):(edges.AboveCount>0?edges.Above[0]:(byte)129);
            for(int y=0;y<h;y++) {_cancellation.ThrowIfCancellationRequested();for(int x=0;x<w;x++) output[y*w+x]=value;}
            return;
        }
        int upTop=0,upLeft=0;
        if(filter && angle!=90 && angle!=180) {
            if(angle>90 && angle<180 && w+h>=24) {
                int corner=(_left[Origin]*5+_above[Origin-1]*6+_above[Origin]*5+8)>>4;
                _above[Origin-1]=_left[Origin-1]=corner;
            }
            if(edges.AboveCount>0 && angle<180) FilterEdge(_above,edges.AboveCount+(angle<90?h:0)+1,Strength(w+h,angle-90,smooth));
            if(edges.LeftCount>0 && angle>90) FilterEdge(_left,edges.LeftCount+(angle>180?w:0)+1,Strength(w+h,angle-180,smooth));
            upTop=Upsample(w+h,angle-90,smooth)?1:0;upLeft=Upsample(w+h,angle-180,smooth)?1:0;
            if(upTop!=0) UpsampleEdge(_above,w+(angle<90?h:0));
            if(upLeft!=0) UpsampleEdge(_left,h+(angle>180?w:0));
        }
        int dx=angle<90?Derivative[angle]:angle<180 && angle>90?Derivative[180-angle]:0;
        int dy=angle>180?Derivative[270-angle]:angle>90 && angle<180?Derivative[angle-90]:0;
        for(int y=0;y<h;y++) {
            _cancellation.ThrowIfCancellationRequested();
            for(int x=0;x<w;x++) {
                int v;
                if(angle==90) v=_above[Origin+x];
                else if(angle==180) v=_left[Origin+y];
                else if(angle<90) {
                    int idx=(y+1)*dx,b=(idx>>(6-upTop))+(x<<upTop),max=(w+h-1)<<upTop;
                    v=b<max?Interpolate(_above,b,((idx<<upTop)>>1)&31):_above[Origin+max];
                } else if(angle<180) {
                    int idx=(x<<6)-(y+1)*dx,b=idx>>(6-upTop);
                    if(b>=-(1<<upTop)) v=Interpolate(_above,b,((idx<<upTop)>>1)&31);
                    else {idx=(y<<6)-(x+1)*dy;b=idx>>(6-upLeft);v=Interpolate(_left,b,((idx<<upLeft)>>1)&31);}
                } else {
                    int idx=(x+1)*dy,b=(idx>>(6-upLeft))+(y<<upLeft),max=(w+h-1)<<upLeft;
                    v=b<max?Interpolate(_left,b,((idx<<upLeft)>>1)&31):_left[Origin+max];
                }
                output[y*w+x]=(byte)v;
            }
        }
    }
    private static int Interpolate(int[] edge,int index,int shift)=>(edge[Origin+index]*(32-shift)+edge[Origin+index+1]*shift+16)>>5;
    private static bool Upsample(int sum,int delta,bool smooth)=>Math.Abs(delta)>0 && Math.Abs(delta)<40 && sum<=(smooth?8:16);
    private static int Strength(int sum,int delta,bool smooth) {
        int d=Math.Abs(delta);
        if(smooth) return sum<=8?d>=64?2:d>=40?1:0:sum<=16?d>=48?2:d>=20?1:0:sum<=24?d>=4?3:0:3;
        return sum<=8?d>=56?1:0:sum<=16?d>=40?1:0:sum<=24?d>=32?3:d>=16?2:d>=8?1:0:sum<=32?d>=32?3:d>=4?2:1:3;
    }
    private void FilterEdge(int[] edge,int count,int strength) {
        if(strength==0) return;
        Array.Copy(edge,Origin-1,_copy,0,count);
        for(int i=1;i<count;i++) {
            int sum=0;for(int j=0;j<5;j++) sum+=EdgeKernels[(strength-1)*5+j]*_copy[Math.Max(0,Math.Min(count-1,i-2+j))];
            edge[Origin+i-1]=(sum+8)>>4;
        }
    }
    private void UpsampleEdge(int[] edge,int count) {
        _copy[0]=edge[Origin-1];for(int i=-1;i<count;i++) _copy[i+2]=edge[Origin+i];_copy[count+2]=edge[Origin+count-1];
        edge[Origin-2]=_copy[0];
        for(int i=0;i<count;i++) {edge[Origin+2*i-1]=Clip((-_copy[i]+9*_copy[i+1]+9*_copy[i+2]-_copy[i+3]+8)>>4);edge[Origin+2*i]=_copy[i+2];}
    }
    private static readonly byte[] EdgeKernels={0,4,8,4,0,0,5,6,5,0,2,4,4,4,2};
}
