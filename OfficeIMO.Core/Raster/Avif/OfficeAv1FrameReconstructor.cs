using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>One-shot Main-8/Main10 reconstruction into owned padded planes through the selected pipeline stage.</summary>
/// <remarks>Only a completely terminated set of tiles publishes a result. All partial pixels stay private.
/// Color/alpha composition and public raster integration are separate owners.</remarks>
internal sealed partial class OfficeAv1FrameReconstructor : IOfficeAv1TileConsumer {
    private readonly OfficeAv1StillFrame _frame;
    private readonly OfficeAv1StillSequence _sequence;
    private readonly OfficeRasterDecodeOptions _tileOptions;
    private readonly CancellationToken _cancellation;
    private readonly ushort[][] _pixels;
    private readonly byte[] _yModes,_uvModes;
    private readonly ushort[] _above,_left,_luma;
    private readonly bool[][] _decoded;
    private readonly OfficeAv1IntraPredictor _predictor;
    private readonly OfficeAv1ResidualTransform _residual;
    private readonly OfficeAv1Deblocker? _deblocker;
    private readonly OfficeAv1Cdef? _cdef;
    private readonly OfficeAv1Restorer? _restorer;
    private readonly OfficeAv1Upscaler? _upscaler;
    private readonly int _stride,_rows,_sb,_maximumSample;
    private OfficeAv1Tile _tile;
    private OfficeAv1TileBlock _block;
    private int _sbRow=-1,_sbCol=-1,_index,_maxLumaX,_maxLumaY;
    private bool _active,_tileComplete;

    private OfficeAv1FrameReconstructor(byte[] bytes,OfficeAv1StillSequence sequence,
        OfficeAv1StillFrame frame,OfficeRasterDecodeOptions options,OfficeAv1ReconstructionStage stage) {
        if(bytes==null) throw new ArgumentNullException(nameof(bytes));
        if(sequence==null) throw new ArgumentNullException(nameof(sequence));
        if(frame==null) throw new ArgumentNullException(nameof(frame));
        if(options==null) throw new ArgumentNullException(nameof(options));
        options.Validate();options.CancellationToken.ThrowIfCancellationRequested();
        if((sequence.BitDepth!=8 && sequence.BitDepth!=10) || frame.BitDepth!=sequence.BitDepth)
            throw new FormatException("Invalid or inconsistent AV1 reconstruction bit depth.");
        if(stage<OfficeAv1ReconstructionStage.Unfiltered || stage>OfficeAv1ReconstructionStage.Restored)
            throw new ArgumentOutOfRangeException(nameof(stage));
        if(frame.Width<1 || frame.Width>65536 || frame.UpscaledWidth<frame.Width || frame.UpscaledWidth>65536 || frame.Height<1 || frame.Height>65536 ||
           frame.MiCols!=((frame.Width+7)/8)*2 || frame.MiRows!=((frame.Height+7)/8)*2 ||
           (long)frame.UpscaledWidth*frame.Height>Math.Min(options.MaximumDecodedPixels,options.MaximumInspectionWorkPixels) || bytes.Length>options.MaximumEncodedBytes)
            throw new FormatException("AV1 reconstruction exceeds frame or input limits.");
        _frame=frame;_sequence=sequence;_cancellation=options.CancellationToken;_sb=sequence.Use128Superblock?128:64;
        _maximumSample=(1<<sequence.BitDepth)-1;
        _stride=(frame.Width+_sb-1)/_sb*_sb;_rows=(frame.Height+_sb-1)/_sb*_sb;
        long tileContext=ValidateTiles(frame,_sb),storage=(long)_stride*_rows*(sequence.Monochrome?2:3)+
            (long)frame.MiRows*frame.MiCols*2;
        // Reserve edges/maps/CfL and the live prediction/residual results, including array/object overhead.
        const long scratch=131072;
        long deblock=stage>=OfficeAv1ReconstructionStage.Deblocked?OfficeAv1Deblocker.ContextBytes(frame,sequence.Monochrome):0;
        long cdef=stage>=OfficeAv1ReconstructionStage.Cdef && sequence.Cdef && !frame.CodedLossless && !frame.AllowIntraBlockCopy?
            OfficeAv1Cdef.ContextBytes(frame,sequence.Monochrome,_sb):0;
        long restoration=stage>=OfficeAv1ReconstructionStage.Restored &&
            (frame.RestorationTypes[0]!=0 || frame.RestorationTypes[1]!=0 || frame.RestorationTypes[2]!=0)?
            OfficeAv1Restorer.ContextBytes(frame,sequence.Monochrome,_sb):0;
        long upscale=stage>=OfficeAv1ReconstructionStage.Upscaled && frame.Width!=frame.UpscaledWidth?
            OfficeAv1Upscaler.ContextBytes(frame,sequence.Monochrome,_sb):0;
        long filter=deblock+cdef+restoration+upscale;
        long owned=storage+scratch+OfficeAv1IntraPredictor.ContextBytes+OfficeAv1ResidualTransform.ContextBytes+filter;
        long all=owned+tileContext+bytes.LongLength;
        if(options.RetainedManagedBytes>OfficeRasterGuards.MaximumDecodedBytes-all)
            throw new FormatException("AV1 reconstruction exceeds aggregate retained memory.");
        _tileOptions=options.WithAdditionalRetainedManagedBytes(owned);
        var child=options.WithAdditionalRetainedManagedBytes(storage+scratch+tileContext+bytes.LongLength+filter);
        _predictor=new OfficeAv1IntraPredictor(sequence.BitDepth,child.WithAdditionalRetainedManagedBytes(OfficeAv1ResidualTransform.ContextBytes));
        _residual=new OfficeAv1ResidualTransform(frame,child.WithAdditionalRetainedManagedBytes(OfficeAv1IntraPredictor.ContextBytes));
        if(deblock!=0) _deblocker=new OfficeAv1Deblocker(frame,sequence.Monochrome,
            options.WithAdditionalRetainedManagedBytes(owned-deblock+tileContext+bytes.LongLength));
        if(cdef!=0) _cdef=new OfficeAv1Cdef(frame,sequence.Monochrome,_sb,
            options.WithAdditionalRetainedManagedBytes(owned-cdef+tileContext+bytes.LongLength));
        if(upscale!=0) _upscaler=new OfficeAv1Upscaler(frame,sequence,_sb,
            options.WithAdditionalRetainedManagedBytes(owned-upscale+tileContext+bytes.LongLength));
        if(restoration!=0) _restorer=new OfficeAv1Restorer(frame,sequence,_sb,
            options.WithAdditionalRetainedManagedBytes(owned-restoration+tileContext+bytes.LongLength));
        _pixels=new ushort[sequence.Monochrome?1:3][];
        for(int p=0;p<_pixels.Length;p++) _pixels[p]=new ushort[checked((_stride>>(p>0?1:0))*(_rows>>(p>0?1:0)))];
        _yModes=new byte[frame.MiRows*frame.MiCols];_uvModes=new byte[_yModes.Length];
        _decoded=new[] {new bool[34*34],new bool[34*34],new bool[34*34]};
        _above=new ushort[128];_left=new ushort[128];_luma=new ushort[4096];
    }

    /// <summary>Decodes every declared tile. Throws without exposing partial output on any failure.</summary>
    internal static OfficeAv1ReconstructedFrame Decode(byte[] bytes,OfficeAv1StillSequence sequence,
        OfficeAv1StillFrame frame,OfficeRasterDecodeOptions options,
        OfficeAv1ReconstructionStage stage=OfficeAv1ReconstructionStage.Unfiltered) {
        var owner=new OfficeAv1FrameReconstructor(bytes,sequence,frame,options,stage);
        foreach(var tile in frame.Tiles) {
            owner._tile=tile;owner._tileComplete=false;owner._sbRow=owner._sbCol=-1;
            new OfficeAv1TileReader(frame,sequence,tile,owner._tileOptions).Read(bytes,owner);
            if(!owner._tileComplete) throw new FormatException("AV1 reconstruction tile did not terminate.");
        }
        owner._deblocker?.Apply(owner._pixels,owner._stride);
        owner._restorer?.CaptureDeblocked(owner._pixels,owner._stride);
        owner._cdef?.Apply(owner._pixels,owner._stride);
        ushort[][] pixels=owner._pixels;int stride=owner._stride,rows=owner._rows,width=frame.Width;
        if(owner._upscaler!=null) {
            owner._restorer?.UpscaleDeblocked(owner._upscaler,owner._stride);
            pixels=owner._upscaler.Apply(pixels,owner._stride);
            stride=owner._upscaler.Stride;rows=owner._upscaler.Rows;width=frame.UpscaledWidth;
        }
        owner._restorer?.Apply(pixels,stride);
        owner._cancellation.ThrowIfCancellationRequested();
        return new OfficeAv1ReconstructedFrame(width,frame.Height,stride,rows,sequence.BitDepth,pixels);
    }

    private static long ValidateTiles(OfficeAv1StillFrame f,int sb) {
        if(f.Tiles==null || f.Tiles.Length<1 || f.Tiles.Length>4096) throw new FormatException("Invalid AV1 reconstruction tile set.");
        int row=0,col=0,end=0;long maximum=0;
        foreach(var t in f.Tiles) {
            _=new OfficeAv1TileGeometry(f,t,sb);
            if(t.MiRowStart!=row || t.MiColStart!=col || (col!=0 && t.MiRowEnd!=end))
                throw new FormatException("AV1 reconstruction tiles must cover the frame in row order.");
            end=t.MiRowEnd;col=t.MiColEnd;
            if(col==f.MiCols) {row=end;col=0;}
            maximum=Math.Max(maximum,OfficeAv1TileReader.ContextBytes(f,t));
        }
        if(row!=f.MiRows || col!=0) throw new FormatException("Incomplete AV1 reconstruction tile coverage.");
        return maximum;
    }

    void IOfficeAv1TileConsumer.Restoration(OfficeAv1RestorationUnit unit) { _cancellation.ThrowIfCancellationRequested();_restorer?.Unit(unit); }
    void IOfficeAv1TileConsumer.BeginBlock(OfficeAv1TileBlock block) {
        if(_active) throw new InvalidOperationException("AV1 reconstruction leaf is already pending.");
        _block=block;_index=0;_active=true;var b=block.Region;
        int r=b.MiRow/(_sb/4)*(_sb/4),c=b.MiCol/(_sb/4)*(_sb/4);
        if(r!=_sbRow || c!=_sbCol) { _sbRow=r;_sbCol=c;ResetDecoded(); }
    }
    void IOfficeAv1TileConsumer.Residual(OfficeAv1Coefficients coefficients) {
        _cancellation.ThrowIfCancellationRequested();
        if(!_active || _index>=_block.Transforms.Count) throw new InvalidOperationException("Unexpected AV1 reconstruction residual.");
        var b=coefficients.Block;var expected=_block.Transforms.Block(_index++);
        if(b.Plane!=expected.Plane || b.X!=expected.X || b.Y!=expected.Y || b.Size!=expected.Size)
            throw new FormatException("AV1 reconstruction residual order mismatch.");
        OfficeAv1Prediction? prediction=_block.Modes.UseIntraBlockCopy?null:Predict(b);
        var residual=coefficients.EndOfBlock==0?null:_residual.Reconstruct(coefficients,_block.Prelude);
        int sub=b.Plane==0?0:1,stride=_stride>>sub;
        for(int y=0;y<b.Height;y++) {
            _cancellation.ThrowIfCancellationRequested();
            for(int x=0;x<b.Width;x++) {
                int v=prediction==null?CopySample(b,x,y):prediction.Value(y*b.Width+x);
                if(residual!=null) v+=residual.Value(y*b.Width+x);
                _pixels[b.Plane][(b.Y+y)*stride+b.X+x]=(ushort)Math.Max(0,Math.Min(_maximumSample,v));
            }
        }
        if(b.Plane==0) {_maxLumaX=b.X+b.Width;_maxLumaY=b.Y+b.Height;}
        MarkDecoded(b);
        _deblocker?.Transform(b);
    }
    void IOfficeAv1TileConsumer.EndBlock() {
        if(!_active || _index!=_block.Transforms.Count) throw new FormatException("Incomplete AV1 reconstruction leaf.");
        var b=_block.Region;
        for(int r=b.MiRow;r<Math.Min(_tile.MiRowEnd,b.MiRow+b.Height/4);r++)
            for(int c=b.MiCol;c<Math.Min(_tile.MiColEnd,b.MiCol+b.Width/4);c++) {
                int i=r*_frame.MiCols+c;_yModes[i]=(byte)_block.Modes.YMode;
                _uvModes[i]=_block.Modes.UseIntraBlockCopy?(byte)255:(byte)_block.Modes.UvMode;
            }
        _deblocker?.Leaf(_block);_cdef?.Leaf(_block);_active=false;
    }
    void IOfficeAv1TileConsumer.CompleteTile() {
        if(_active) throw new FormatException("AV1 reconstruction has a pending leaf.");
        _cancellation.ThrowIfCancellationRequested();_tileComplete=true;
    }
}

/// <summary>Internal pipeline boundary; later frame filtering remains a separate stage.</summary>
internal enum OfficeAv1ReconstructionStage { Unfiltered, Deblocked, Cdef, Upscaled, Restored }

/// <summary>Immutable owned reconstructed planes. Padded samples are retained for subsequent filter owners.</summary>
internal sealed class OfficeAv1ReconstructedFrame {
    private readonly ushort[][] _planes;
    private readonly int _stride,_rows;
    internal OfficeAv1ReconstructedFrame(int width,int height,int stride,int rows,int bitDepth,ushort[][] planes) {
        Width=width;Height=height;BitDepth=bitDepth;_stride=stride;_rows=rows;_planes=planes;
    }
    internal int Width {get;}
    internal int Height {get;}
    internal int BitDepth {get;}
    internal int PlaneCount=>_planes.Length;
    /// <summary>Owned plane storage charged while another decode/composition owner is live.</summary>
    internal long StorageBytes {
        get { long bytes=0;foreach(ushort[] plane in _planes)bytes+=plane.LongLength*2;return bytes; }
    }
    internal ushort Value(int plane,int x,int y) {
        if((uint)plane>=(uint)_planes.Length) throw new ArgumentOutOfRangeException(nameof(plane));
        int sub=plane==0?0:1;
        if((uint)x>=(uint)(_stride>>sub) || (uint)y>=(uint)(_rows>>sub)) throw new ArgumentOutOfRangeException();
        return _planes[plane][y*(_stride>>sub)+x];
    }
}
