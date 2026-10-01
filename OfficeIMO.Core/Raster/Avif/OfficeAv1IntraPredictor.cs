using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Sequential Main-8 prediction owner over explicitly available reconstructed edges.</summary>
/// <remarks>The reconstruction consumer derives availability from tile and decode order, and budgets
/// retained outputs separately. Input edges are copied before filtering; results never expose scratch.</remarks>
internal sealed partial class OfficeAv1IntraPredictor {
    internal const long ContextBytes=12288;
    private const int Origin=2;
    private readonly int[] _above,_left,_copy,_neighbors,_luma;
    private readonly long _retained;
    private readonly long _maximumPixels;
    private readonly CancellationToken _cancellation;

    internal OfficeAv1IntraPredictor(OfficeRasterDecodeOptions options) {
        if(options==null) throw new ArgumentNullException(nameof(options));
        options.Validate();options.CancellationToken.ThrowIfCancellationRequested();
        const long context=ContextBytes; // reusable edges, CfL samples, numeric facts and object/array overhead
        if(options.RetainedManagedBytes>OfficeRasterGuards.MaximumDecodedBytes-context)
            throw new FormatException("AV1 prediction contexts exceed the retained-memory limit.");
        _retained=options.RetainedManagedBytes+context;_maximumPixels=options.MaximumDecodedPixels;
        _cancellation=options.CancellationToken;
        _above=new int[258];_left=new int[258];_copy=new int[132];_neighbors=new int[7];_luma=new int[1024];
    }

    /// <summary>Predicts one transform using raw available edge samples and their explicit counts.</summary>
    /// <param name="size">Normative transform size identifier.</param>
    /// <param name="mode">Ordinary intra mode; CfL and palettes have separate entry points.</param>
    /// <param name="angleDelta">Leaf mode reader's angle delta; eligibility is based on the leaf, not its transform.</param>
    /// <param name="edges">Available edge prefixes and continuations, before filtering.</param>
    /// <param name="enableEdgeFilter">Sequence-level directional edge filter enable flag.</param>
    /// <param name="smoothNeighbors">Whether an available neighboring leaf uses a smooth prediction mode.</param>
    /// <param name="filterMode">-1 for ordinary prediction, or the decoded luma DC recursive filter mode 0..4.</param>
    internal OfficeAv1Prediction Predict(int size,OfficeAv1IntraMode mode,int angleDelta,
        OfficeAv1PredictionEdges edges,bool enableEdgeFilter,bool smoothNeighbors,int filterMode=-1) {
        var dimensions=Dimensions(size);int w=dimensions.Width,h=dimensions.Height;
        if((uint)mode>12 || angleDelta< -3 || angleDelta>3 || filterMode< -1 || filterMode>4 ||
           (filterMode>=0 && (mode!=OfficeAv1IntraMode.Dc || w>32 || h>32)) ||
           (angleDelta!=0 && (mode<OfficeAv1IntraMode.Vertical || mode>OfficeAv1IntraMode.Diagonal67)))
            throw new FormatException("Invalid AV1 intra prediction mode or angle.");
        Prepare(edges,w,h);
        var result=new byte[w*h];
        if(filterMode>=0) Recursive(result,w,h,filterMode);
        else if(mode>=OfficeAv1IntraMode.Vertical && mode<=OfficeAv1IntraMode.Diagonal67)
            Directional(result,w,h,ModeAngles[(int)mode]+3*angleDelta,edges,enableEdgeFilter,smoothNeighbors);
        else Basic(result,w,h,mode,edges.AboveCount>0,edges.LeftCount>0);
        _cancellation.ThrowIfCancellationRequested();return new OfficeAv1Prediction(w,h,result);
    }

    private (int Width,int Height) Dimensions(int size) {
        _cancellation.ThrowIfCancellationRequested();
        if((uint)size>=19) throw new FormatException("Invalid AV1 prediction transform size.");
        int w=OfficeAv1TransformSize.Width(size),h=OfficeAv1TransformSize.Height(size);
        if(w*h>_maximumPixels || w*h+24>OfficeRasterGuards.MaximumDecodedBytes-_retained)
            throw new FormatException("AV1 prediction output exceeds its pixel or retained-memory limit.");
        return (w,h);
    }

    private void Prepare(OfficeAv1PredictionEdges e,int w,int h) {
        if(e.Above==null || e.Left==null || e.Above.Length>128 || e.Left.Length>128 ||
           (uint)e.AboveCount>(uint)w || (uint)e.LeftCount>(uint)h ||
           (uint)e.AboveRightCount>(uint)h || (uint)e.BelowLeftCount>(uint)w ||
           (e.AboveRightCount>0 && e.AboveCount!=w) || (e.BelowLeftCount>0 && e.LeftCount!=h) ||
           e.Above.Length<e.AboveCount+e.AboveRightCount || e.Left.Length<e.LeftCount+e.BelowLeftCount)
            throw new FormatException("Invalid AV1 prediction edge availability or storage.");
        int top=e.AboveCount+e.AboveRightCount,left=e.LeftCount+e.BelowLeftCount;
        int corner=e.AboveCount>0 && e.LeftCount>0?e.Corner:e.AboveCount>0?e.Above[0]:e.LeftCount>0?e.Left[0]:128;
        _above[Origin-1]=_left[Origin-1]=corner;
        for(int i=0;i<w+h;i++) {
            _above[Origin+i]=top>0?e.Above[Math.Min(i,top-1)]:left>0?e.Left[0]:127;
            _left[Origin+i]=left>0?e.Left[Math.Min(i,left-1)]:top>0?e.Above[0]:129;
        }
    }

    private void Basic(byte[] output,int w,int h,OfficeAv1IntraMode mode,bool top,bool left) {
        int dc=128;
        if(mode==OfficeAv1IntraMode.Dc) {
            int count=0,sum=0;
            if(top) {for(int x=0;x<w;x++) sum+=_above[Origin+x];count+=w;}
            if(left) {for(int y=0;y<h;y++) sum+=_left[Origin+y];count+=h;}
            if(count!=0) dc=(sum+count/2)/count;
        }
        for(int y=0;y<h;y++) {
            _cancellation.ThrowIfCancellationRequested();
            for(int x=0;x<w;x++) {
                int a=_above[Origin+x],l=_left[Origin+y],v=dc;
                if(mode==OfficeAv1IntraMode.Paeth) {
                    int c=_above[Origin-1],b=a+l-c,pl=Math.Abs(b-l),pt=Math.Abs(b-a),pc=Math.Abs(b-c);
                    v=pl<=pt && pl<=pc?l:pt<=pc?a:c;
                } else if(mode>=OfficeAv1IntraMode.Smooth && mode<=OfficeAv1IntraMode.SmoothHorizontal) {
                    int wx=SmoothWeights[w-4+x],wy=SmoothWeights[h-4+y];
                    int vertical=wy*a+(256-wy)*_left[Origin+h-1],horizontal=wx*l+(256-wx)*_above[Origin+w-1];
                    v=mode==OfficeAv1IntraMode.Smooth?(vertical+horizontal+256)>>9:
                        mode==OfficeAv1IntraMode.SmoothVertical?(vertical+128)>>8:(horizontal+128)>>8;
                }
                output[y*w+x]=(byte)v;
            }
        }
    }

    private void Recursive(byte[] output,int w,int h,int mode) {
        for(int y=0;y<h;y+=2) {
            _cancellation.ThrowIfCancellationRequested();
            for(int x=0;x<w;x+=4) {
                for(int i=0;i<5;i++) _neighbors[i]=y==0?_above[Origin+x+i-1]:x==0 && i==0?_left[Origin+y-1]:output[(y-1)*w+x+i-1];
                for(int i=5;i<7;i++) _neighbors[i]=x==0?_left[Origin+y+i-5]:output[(y+i-5)*w+x-1];
                for(int dy=0;dy<2;dy++) for(int dx=0;dx<4;dx++) {
                    int sum=0,offset=(mode*8+dy*4+dx)*7;
                    for(int i=0;i<7;i++) sum+=FilterTaps[offset+i]*_neighbors[i];
                    output[(y+dy)*w+x+dx]=Clip((sum+8)>>4);
                }
            }
        }
    }
    private static byte Clip(int value)=>(byte)Math.Max(0,Math.Min(255,value));
}

/// <summary>Raw top/left samples, their available prefixes and optional continuations; no ownership transfer.</summary>
internal readonly struct OfficeAv1PredictionEdges {
    internal OfficeAv1PredictionEdges(byte[] above,byte[] left,byte corner,int aboveCount,int leftCount,int aboveRightCount=0,int belowLeftCount=0) {
        Above=above;Left=left;Corner=corner;AboveCount=aboveCount;LeftCount=leftCount;AboveRightCount=aboveRightCount;BelowLeftCount=belowLeftCount;
    }
    internal byte[] Above {get;}
    internal byte[] Left {get;}
    internal byte Corner {get;}
    internal int AboveCount {get;}
    internal int LeftCount {get;}
    internal int AboveRightCount {get;}
    internal int BelowLeftCount {get;}
}

/// <summary>Immutable row-major prediction samples, before residual addition and final clipping.</summary>
internal sealed class OfficeAv1Prediction {
    private readonly byte[] _values;
    internal OfficeAv1Prediction(int width,int height,byte[] values) {Width=width;Height=height;_values=values;}
    internal int Width {get;}
    internal int Height {get;}
    internal int Count=>_values.Length;
    internal byte Value(int index)=>_values[index];
}
