using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Retains only the Main-8 intra metadata needed for the normative frame deblocking pass.</summary>
internal sealed class OfficeAv1Deblocker {
    private readonly OfficeAv1StillFrame _frame;
    private readonly byte[] _levels;
    private readonly byte[][] _sizes;
    private readonly int[] _scratch=new int[32];
    private readonly CancellationToken _cancellation;
    internal static long ContextBytes(OfficeAv1StillFrame frame,bool monochrome)=>
        (long)frame.MiRows*frame.MiCols*(monochrome?10:11)/2+1024;
    internal OfficeAv1Deblocker(OfficeAv1StillFrame frame,bool monochrome,OfficeRasterDecodeOptions options) {
        options.CancellationToken.ThrowIfCancellationRequested();
        if(frame.BitDepth!=8) throw new FormatException("AV1 high-bit-depth deblocking is not qualified.");
        if(options.RetainedManagedBytes>OfficeRasterGuards.MaximumDecodedBytes-ContextBytes(frame,monochrome))
            throw new FormatException("AV1 deblocking contexts exceed retained memory.");
        if(frame.LoopFilterSharpness<0 || frame.LoopFilterSharpness>7) throw new FormatException("Invalid AV1 filter sharpness.");
        foreach(int v in frame.LoopFilterLevels) if(v<0 || v>63) throw new FormatException("Invalid AV1 filter level.");
        _frame=frame;_cancellation=options.CancellationToken;
        _levels=new byte[checked(frame.MiRows*frame.MiCols*4)];_sizes=new byte[monochrome?1:3][];
        for(int p=0;p<_sizes.Length;p++) _sizes[p]=new byte[(frame.MiRows>>(p>0?1:0))*(frame.MiCols>>(p>0?1:0))];
    }
    internal void Transform(OfficeAv1TransformBlock block) {
        int sub=block.Plane==0?0:1,rows=_frame.MiRows>>sub,cols=_frame.MiCols>>sub;
        for(int y=block.Y/4;y<Math.Min(rows,(block.Y+block.Height)/4);y++) {
            _cancellation.ThrowIfCancellationRequested();
            for(int x=block.X/4;x<Math.Min(cols,(block.X+block.Width)/4);x++) _sizes[block.Plane][y*cols+x]=(byte)block.Size;
        }
    }
    internal void Leaf(OfficeAv1TileBlock block) {
        var prelude=block.Prelude;var b=block.Region;
        for(int i=0;i<4;i++) {
            int delta=!_frame.DeltaLoopFilterMulti?prelude.Filter0:i==0?prelude.Filter0:i==1?prelude.Filter1:i==2?prelude.Filter2:prelude.Filter3;
            int level=Clip(delta+_frame.LoopFilterLevels[i]);
            if(_frame.SegmentFeatures[prelude.SegmentId,i+1]) level=Clip(level+_frame.SegmentData[prelude.SegmentId,i+1]);
            if(_frame.LoopFilterDeltaEnabled) level=Clip(level+(_frame.LoopFilterReferenceDeltas[0]<<(level>>5)));
            for(int y=b.MiRow;y<Math.Min(_frame.MiRows,b.MiRow+b.Height/4);y++) {
                _cancellation.ThrowIfCancellationRequested();
                for(int x=b.MiCol;x<Math.Min(_frame.MiCols,b.MiCol+b.Width/4);x++) _levels[(y*_frame.MiCols+x)*4+i]=(byte)level;
            }
        }
    }
    /// <summary>Filters vertical boundaries before horizontal boundaries; no partial result is published.</summary>
    internal void Apply(ushort[][] planes,int stride) {
        _cancellation.ThrowIfCancellationRequested();
        if(_frame.AllowIntraBlockCopy || (_frame.LoopFilterLevels[0]==0 && _frame.LoopFilterLevels[1]==0)) return;
        for(int p=0;p<planes.Length;p++) {
            if(p>0 && _frame.LoopFilterLevels[p+1]==0) continue;
            int sub=p==0?0:1,step=1<<sub,cols=_frame.MiCols>>sub,pitch=stride>>sub;
            for(int pass=0;pass<2;pass++) for(int r=0;r<_frame.MiRows;r+=step) {
                _cancellation.ThrowIfCancellationRequested();
                for(int c=0;c<_frame.MiCols;c+=step) {
                    int x=c*4,y=r*4;
                    if(x>=_frame.Width || y>=_frame.Height || (pass==0?x==0:y==0)) continue;
                    int row=r|sub,col=c|sub,pr=row-(pass==1?step:0),pc=col-(pass==0?step:0);
                    int size=_sizes[p][(row>>sub)*cols+(col>>sub)],previous=_sizes[p][(pr>>sub)*cols+(pc>>sub)];
                    int dimension=pass==0?OfficeAv1TransformSize.Width(size):OfficeAv1TransformSize.Height(size);
                    int coordinate=(pass==0?x:y)>>sub;
                    if(coordinate%dimension!=0) continue; // All reduced-still leaves are intra, including skipped intra.
                    int neighbor=pass==0?OfficeAv1TransformSize.Width(previous):OfficeAv1TransformSize.Height(previous);
                    int width=Math.Min(p==0?16:8,Math.Min(dimension,neighbor)),index=p==0?pass:p+1;
                    int level=_levels[(row*_frame.MiCols+col)*4+index];
                    if(level==0) level=_levels[(pr*_frame.MiCols+pc)*4+index];
                    if(level==0) continue;
                    int sharp=_frame.LoopFilterSharpness,shift=sharp>4?2:sharp>0?1:0;
                    int limit=Math.Max(1,sharp>0?Math.Min(9-sharp,level>>shift):level);
                    for(int i=0;i<4;i++) {
                        int offset=((y>>sub)+(pass==0?i:0))*pitch+(x>>sub)+(pass==1?i:0);
                        OfficeAv1DeblockFilter.Apply(planes[p],offset,pass==0?1:pitch,width,p>0,limit,2*(level+2)+limit,level>>4,_scratch);
                    }
                }
            }
        }
    }
    private static int Clip(int value)=>Math.Max(0,Math.Min(63,value));
}
